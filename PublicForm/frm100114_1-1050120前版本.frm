VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100114_1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人案件查詢"
   ClientHeight    =   5655
   ClientLeft      =   1695
   ClientTop       =   3105
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9300
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印對造資料(&O)"
      Enabled         =   0   'False
      Height          =   300
      Index           =   8
      Left            =   7720
      Style           =   1  '圖片外觀
      TabIndex        =   45
      Top             =   795
      Width           =   1515
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含投資法務開拓資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5370
      TabIndex        =   43
      Top             =   1440
      Width           =   1770
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   6855
      TabIndex        =   8
      Top             =   1140
      Width           =   1600
   End
   Begin VB.OptionButton Option1 
      Caption         =   "E-Mail："
      Height          =   180
      Index           =   3
      Left            =   5940
      TabIndex        =   41
      Top             =   1200
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "聯絡人(&T)"
      Height          =   345
      Index           =   7
      Left            =   5895
      Style           =   1  '圖片外觀
      TabIndex        =   40
      Top             =   390
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "往來紀錄(&N)"
      Height          =   345
      Index           =   6
      Left            =   6885
      Style           =   1  '圖片外觀
      TabIndex        =   39
      Top             =   390
      Width           =   1170
   End
   Begin VB.CheckBox ChkPCT 
      Caption         =   "是否顯示PCT 案"
      Height          =   225
      Left            =   3390
      TabIndex        =   11
      Top             =   1770
      Width           =   1635
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "法務進度(&L)"
      Height          =   345
      Index           =   5
      Left            =   8070
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   375
      Width           =   1170
   End
   Begin VB.CheckBox chk 
      Caption         =   "所有系統類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   90
      TabIndex        =   38
      Top             =   90
      Width           =   2355
   End
   Begin VB.Frame Frame2 
      Height          =   360
      Left            =   5190
      TabIndex        =   37
      Top             =   705
      Width           =   2475
      Begin VB.OptionButton Option3 
         Caption         =   "字首比對"
         Height          =   180
         Index           =   0
         Left            =   96
         TabIndex        =   5
         Top             =   144
         Width           =   1125
      End
      Begin VB.OptionButton Option3 
         Caption         =   "模糊比對"
         Height          =   180
         Index           =   1
         Left            =   1230
         TabIndex        =   6
         Top             =   144
         Value           =   -1  'True
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Height          =   360
      Left            =   2910
      TabIndex        =   36
      Top             =   2460
      Visible         =   0   'False
      Width           =   3195
      Begin VB.OptionButton Option2 
         Caption         =   "日文"
         Height          =   204
         Index           =   2
         Left            =   2265
         TabIndex        =   4
         Top             =   120
         Width           =   732
      End
      Begin VB.OptionButton Option2 
         Caption         =   "英文"
         Height          =   204
         Index           =   1
         Left            =   1185
         TabIndex        =   3
         Top             =   120
         Width           =   732
      End
      Begin VB.OptionButton Option2 
         Caption         =   "中文"
         Height          =   204
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   120
         Value           =   -1  'True
         Width           =   732
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   2685
      Left            =   0
      TabIndex        =   25
      Top             =   2940
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   4736
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin VB.OptionButton Option1 
      Caption         =   "代理人國籍："
      Height          =   204
      Index           =   2
      Left            =   72
      TabIndex        =   17
      Top             =   2655
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "代理人名稱："
      Height          =   204
      Index           =   1
      Left            =   72
      TabIndex        =   1
      Top             =   825
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "代理人編號："
      Height          =   204
      Index           =   0
      Left            =   72
      TabIndex        =   27
      Top             =   450
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   5544
      MaxLength       =   4
      TabIndex        =   15
      Top             =   2070
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1515
      MaxLength       =   4
      TabIndex        =   18
      Top             =   2640
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1104
      MaxLength       =   4
      TabIndex        =   16
      Top             =   2355
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1104
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1755
      Width           =   492
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   0
      Left            =   1590
      MaxLength       =   9
      TabIndex        =   0
      Top             =   420
      Width           =   1932
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1590
      MaxLength       =   80
      TabIndex        =   7
      Top             =   795
      Width           =   3550
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   4344
      MaxLength       =   4
      TabIndex        =   14
      Top             =   2070
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1104
      MaxLength       =   7
      TabIndex        =   12
      Top             =   2070
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1104
      TabIndex        =   9
      Top             =   1440
      Width           =   2772
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2304
      MaxLength       =   7
      TabIndex        =   13
      Top             =   2070
      Width           =   852
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   7416
      Style           =   1  '圖片外觀
      TabIndex        =   26
      Top             =   3735
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   345
      Left            =   3960
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "關係企業(&R)"
      Height          =   345
      Index           =   2
      Left            =   7260
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件資料(&B)"
      Height          =   345
      Index           =   1
      Left            =   6060
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "代理人資料(&O)"
      Height          =   345
      Index           =   0
      Left            =   4755
      Style           =   1  '圖片外觀
      TabIndex        =   20
      Top             =   10
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   4
      Left            =   8460
      Style           =   1  '圖片外觀
      TabIndex        =   24
      Top             =   10
      Width           =   800
   End
   Begin VB.Label Label1 
      Caption         =   "＊：舊的名稱　＄：有呆帳　 ●：特殊客戶　 ⊕：不得代理"
      ForeColor       =   &H000000C0&
      Height          =   720
      Left            =   7560
      TabIndex        =   47
      Top             =   1830
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "輸入名稱之特取部分, 不要取國家,省份,城市,例：不可輸美商..,廣東..,廣州.."
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   46
      Top             =   1200
      Width           =   5805
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "註：紅色資料不可承接案件"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   6210
      TabIndex        =   44
      Top             =   2730
      Width           =   3030
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "模糊比對"
      Height          =   180
      Left            =   8520
      TabIndex        =   42
      Top             =   1200
      Width           =   720
   End
   Begin VB.Line Line2 
      X1              =   5310
      X2              =   5430
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line Line1 
      X1              =   2070
      X2              =   2190
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Label lbl1 
      Height          =   210
      Index           =   1
      Left            =   2640
      TabIndex        =   35
      Top             =   2670
      Width           =   3270
   End
   Begin VB.Label lbl1 
      Height          =   210
      Index           =   0
      Left            =   2250
      TabIndex        =   34
      Top             =   2370
      Width           =   3735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   150
      TabIndex        =   33
      Top             =   2370
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "（1.收文  2.發文）"
      Height          =   180
      Left            =   1710
      TabIndex        =   32
      Top             =   1755
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "查詢別："
      Height          =   180
      Left            =   150
      TabIndex        =   31
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                               (ALL：全部)"
      Height          =   180
      Left            =   150
      TabIndex        =   30
      Top             =   1470
      Width           =   4725
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   180
      Left            =   150
      TabIndex        =   29
      Top             =   2070
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   3390
      TabIndex        =   28
      Top             =   2070
      Width           =   900
   End
End
Attribute VB_Name = "frm100114_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2013/12/04 拿掉查無資料查對造功能(已加查對照) 拿掉中、英、日查詢選項
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
'重整 2005/10/05 nickc
Option Explicit

Dim s As Long, i As Long, j As Long, strSql As String
Dim StrToGrid As String
'92.04.16 nick 紀錄作用按鍵
Public CmdState As Integer
'Add by Amy 2013/12/02
Dim StrToPrint As String '記錄編號 for 對造列印
Dim strTp(3) As String
Dim PLeft() As Integer
Dim ColName() As String
Dim intCounter As Integer
Dim intRecord As Integer
Dim intPage As Integer
Dim kk As Integer
Dim bolPrint As Boolean '是否有對造
'end 2013/12/04

Private Sub SetDataListWidth()
   grdDataList.row = 0
   grdDataList.col = 0: grdDataList.Text = "V"
   grdDataList.ColWidth(0) = 200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 1: grdDataList.Text = "編號"
   'grdDataList.ColWidth(1) = 1000
   grdDataList.ColWidth(1) = 1200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 2: grdDataList.Text = "名稱"
   grdDataList.ColWidth(2) = 4000 'Modify by Amy 2014/06/05 改與申請人查詢同 原:4600
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 3: grdDataList.Text = "國籍"
   grdDataList.ColWidth(3) = 1200 'Modify by Amy 2014/06/05 改與申請人查詢同 原:1500
   grdDataList.CellAlignment = flexAlignCenterCenter
   'Modify by Amy 2013/12/04 插入智權人員欄
   grdDataList.col = 4: grdDataList.Text = "智權人員"
   grdDataList.ColWidth(4) = 800
   grdDataList.CellAlignment = flexAlignCenterCenter
   
   grdDataList.col = 5: grdDataList.Text = "狀態"
   grdDataList.ColWidth(5) = 1000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 6: grdDataList.Text = "備註"
   grdDataList.ColWidth(6) = 2000
   grdDataList.CellAlignment = flexAlignLeftCenter
    'Add by Amy 2013/12/04
   '因查詢服務對造資料需依sp09抓不智權人員資料,故加申請國家
   grdDataList.col = 7: grdDataList.Text = "申請國家"
   grdDataList.ColWidth(7) = 0
   '抓取對造欄位 for 列印
   grdDataList.col = 8: grdDataList.Text = "總收文號"
   grdDataList.ColWidth(8) = 0
   grdDataList.col = 9: grdDataList.Text = "案件性質"
   grdDataList.ColWidth(9) = 0
   grdDataList.col = 10: grdDataList.Text = "收文日"
   grdDataList.ColWidth(10) = 0
   'end 2013/12/04
End Sub

Private Sub chk_Click()
   'Add By Cheng 2003/04/01
   '若勾選所有系統類別
   If Me.chk.Value = vbChecked Then
       Me.txt1(2).Text = "ALL"
   '若取消勾選所有系統類別
   Else
       Me.txt1(2).Text = Systemkind_g
   End If
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
'Add by Amy 2014/05/07
Dim strTmp As String
Dim strCaseNo As String '本所案號(for 對造)

   Select Case CmdState
      Case 0
           Me.Enabled = False
           For i = 1 To grdDataList.Rows - 1
           grdDataList.col = 0
           grdDataList.row = i
           If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               'Add By Sindy 2012/3/21
               grdDataList.col = 1
               'Modify by Amy 2013/12/10 +判斷對造
               If Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, 5) = "對造" Then
                  For j = 0 To grdDataList.Cols - 1
                     grdDataList.col = j
                     grdDataList.CellBackColor = &H8080FF
                  Next j
               Else
               '2012/3/21 End
                  For j = 0 To grdDataList.Cols - 1
                     'add by nickc 2005/12/14
                     If j <> 1 Then
                         grdDataList.col = j
                         grdDataList.CellBackColor = QBColor(15)
                     End If
                  Next j
               End If
               grdDataList.col = 1
               If Not IsNull(grdDataList.Text) Then
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                  End If
                  Screen.MousePointer = vbHourglass
                  'Modify by Morgan 2007/12/21 加判斷第一碼切不同畫面
                  'frm100101_10.Show
                  'frm100101_10.Tag = Pub_RplStr(GrdDataList.Text)
                  'frm100101_10.StrMenu
                  strExc(1) = Pub_RplStr(grdDataList.Text)
                  Select Case Left(strExc(1), 1)
                     Case "X"
                        'Add by Morgan 2008/8/11
                        If Mid(strExc(1), 10, 1) = "-" Then
                           strExc(1) = Left(strExc(1), 9)
                        End If
                        frm100101_11.Show
                        frm100101_11.Tag = strExc(1)
                        frm100101_11.StrMenu
                        
                     Case "Y"
                        'Add by Morgan 2008/8/11
                        If Mid(strExc(1), 10, 1) = "-" Then
                           strExc(1) = Left(strExc(1), 9)
                        End If
                        frm100101_10.Show
                        frm100101_10.Tag = strExc(1)
                        frm100101_10.StrMenu
      
                     Case "R"
                        'Modify By Sindy 2009/06/24 判斷是國外或是國內潛在客戶
                        strExc(0) = "select * from potcustomer where pcu01(+)='" & Left(strExc(1), 8) & "' and pcu02(+)='" & Mid(strExc(1), 9, 1) & "' "
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        strExc(2) = ""
                        If intI = 1 Then
                           strExc(2) = "" & RsTemp.Fields(0)
                        End If
                        If strExc(2) <> "" Then '國外
                           frm100101_14.Show
                           frm100101_14.Tag = strExc(1)
                           frm100101_14.StrMenu
                        Else '國內
                           frm100101_21.Show
                           frm100101_21.Tag = strExc(1)
                           frm100101_21.StrMenu
                        End If
                     'Add by Amy 2015/03/27 +客戶端平台帳號
                     Case "平"
                        'Modify by Amy 2015/04/15 改以平台編號抓權限
                        If PUB_ChkCustWebLimit(grdDataList.TextMatrix(grdDataList.row, 10), strUserNum) = True Then
                           frm100101_27.Show
                           frm100101_27.Tag = Trim(grdDataList.TextMatrix(grdDataList.row, 10))
                           frm100101_27.StrMenu
                        Else
                           Me.Show
                           MsgBox "您無權限查詢此客戶端平台帳號！", vbInformation
                        End If
                     'Add By Sindy 2009/07/22
                     Case Else
                        'Modify By Sindy 2012/3/21 +不得代理案件之客戶或代理人
                        If InStr(strExc(1), "-") = 0 Then
                           frm100101_25.Show
                           frm100101_25.Tag = strExc(1)
                           frm100101_25.StrMenu
                        Else
                        '2012/3/21 End
                           frm100101_22.Show
                           frm100101_22.Tag = strExc(1)
                           frm100101_22.StrMenu
                        End If
                     '2009/07/22 End
                  End Select
                  'end 2007/12/21
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
           End If
           Next i
           Me.Enabled = True
      Case 1
           Me.Enabled = False
           For i = 1 To grdDataList.Rows - 1
           grdDataList.col = 0
           grdDataList.row = i
           If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               'Add By Sindy 2012/3/21
               grdDataList.col = 1
               'Modify by Amy 2013/12/10 +判斷對造
               If Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, 5) = "對造" Then
                  For j = 0 To grdDataList.Cols - 1
                     grdDataList.col = j
                     grdDataList.CellBackColor = &H8080FF
                  Next j
               Else
               '2012/3/21 End
                  For j = 0 To grdDataList.Cols - 1
                      'add by nickc 2005/12/14
                      If j <> 1 Then
                          grdDataList.col = j
                          grdDataList.CellBackColor = QBColor(15)
                      End If
                  Next j
               End If
               grdDataList.col = 1
               If Not IsNull(grdDataList.Text) Then
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                  End If
                  
                  'Modify by Amy 2014/05/07 +以本所案號抓案件資料
                  If grdDataList.TextMatrix(grdDataList.row, 5) = "對造" Or grdDataList.TextMatrix(grdDataList.row, 5) = "其他相關人" Then
                    strCaseNo = Pub_RplStr(grdDataList.Text)
                    strTmp = GetPrjPeopleNum1(strCaseNo)
                  Else
                    strTmp = Pub_RplStr(grdDataList.Text)
                  End If
                  'end 2014/05/07
                  
                  'Modify by Amy 2014/05/07
                  Select Case UCase(Left(strTmp, 1)) '2014/05/07 原:UCase(Mid(GrdDataList.Text, 1, 1))
                  Case "X" '申請人
                     Screen.MousePointer = vbHourglass
                     With frm100102_2
                        .Show
                        .Tag = strTmp '2014/05/07 原:Pub_RplStr(GrdDataList.Text)
                        'add b nickc 2007/12/21
                        .ChkPCT = Me.ChkPCT
                        
                        If strCaseNo <> "" Then
                            .m_CaseNo = strCaseNo
                            .StrMenu4
                        Else
                            'Modify by Morgan 2008/11/27
                            '為使查詢案件畫面共用條件改參數方式傳遞
                            '.StrMenu2
                            .m_Sys = txt1(2)
                            .m_Type = txt1(3)
                            .m_Date1 = txt1(4)
                            .m_Date2 = txt1(5)
                            .m_Pty1 = txt1(6)
                            .m_Pty2 = txt1(7)
                            .m_Cty1 = txt1(8)
                            .m_Cty2 = txt1(8)
                            .StrMenu
                            'end 2008/11/27
                        End If
                     End With
                     'end 2014/05/07
                     Screen.MousePointer = vbDefault
                  
                  Case "Y" '代理人
                      Screen.MousePointer = vbHourglass
                      With frm100114_2
                        .Show
                        .Tag = strTmp '2014/05/07 原:Pub_RplStr(GrdDataList.Text)
                        'add b nickc 2007/12/21
                        .ChkPCT = Me.ChkPCT
                        'Modify by Morgan 2008/11/21
                        '為使查詢案件畫面共用條件改參數方式傳遞
                        .m_Sys = txt1(2)
                        .m_Type = txt1(3)
                        .m_Date1 = txt1(4)
                        .m_Date2 = txt1(5)
                        .m_Pty1 = txt1(6)
                        .m_Pty2 = txt1(7)
                        .m_Cty1 = txt1(8)
                        .m_Cty2 = txt1(8)
                        'end 2008/11/21
                        .StrMenu
                      End With
                      Screen.MousePointer = vbDefault
                  Case "R"
                     Me.Show
                     MsgBox "該編號為潛在客戶不會有案件資料！", vbInformation
                  Case Else
                     Me.Show
                  End Select
                  Me.Enabled = True
                  Exit Sub
               End If
           End If
           Next i
           Me.Enabled = True
      Case 2
            Me.Enabled = False
            cnnConnection.Execute "delete from r100114 where id='" & strUserNum & "' "
            For i = 1 To grdDataList.Rows - 1
              grdDataList.col = 0
              grdDataList.row = i
              If Trim(grdDataList.Text) = "V" Then
                  grdDataList.col = 1
                  Screen.MousePointer = vbHourglass
                  Call StrMenu(Pub_RplStr(grdDataList.Text))
                  cmdOK(2).Enabled = False
                  Screen.MousePointer = vbDefault
              End If
            Next i
            Call StrMenu1
            Me.Enabled = True
      Case 3
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 4
           fnCloseAllFrm100
      'add by nickc 2005/10/05 法務進度
      Case 5
           Me.Enabled = False
           For i = 1 To grdDataList.Rows - 1
           grdDataList.col = 0
           grdDataList.row = i
           If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               'Add By Sindy 2012/3/21
               grdDataList.col = 1
               'Modify by Amy 2013/12/10 +判斷對造
               If Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, 5) = "對造" Then
                  For j = 0 To grdDataList.Cols - 1
                     grdDataList.col = j
                     grdDataList.CellBackColor = &H8080FF
                  Next j
               Else
               '2012/3/21 End
                  For j = 0 To grdDataList.Cols - 1
                      'add by nickc 2005/12/14
                      If j <> 1 Then
                          grdDataList.col = j
                          grdDataList.CellBackColor = QBColor(15)
                      End If
                  Next j
               End If
               grdDataList.col = 1
               If Not IsNull(grdDataList.Text) Then
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                  End If
                  If UCase(Mid(grdDataList.Text, 1, 1)) = "X" Then
                  '申請人
                     Screen.MousePointer = vbHourglass
                     With frm100102_2
                     .Show
                     .Tag = Pub_RplStr(grdDataList.Text)
                     'add b nickc 2007/12/21
                     .ChkPCT = Me.ChkPCT
                     .bolIsL = True
                     'Modify by Morgan 2008/11/27
                     '為使查詢案件畫面共用條件改參數方式傳遞
                     '.StrMenu2
                     .m_Sys = txt1(2)
                     .m_Type = txt1(3)
                     .m_Date1 = txt1(4)
                     .m_Date2 = txt1(5)
                     .m_Pty1 = txt1(6)
                     .m_Pty2 = txt1(7)
                     .m_Cty1 = txt1(8)
                     .m_Cty2 = txt1(8)
                     .StrMenu
                     'end 2008/11/27
                     End With
                     Screen.MousePointer = vbDefault
                     
                  Else
                  '代理人
                      Screen.MousePointer = vbHourglass
                      With frm100114_2
                      .Show
                      .Tag = Pub_RplStr(grdDataList.Text)
                      'add b nickc 2007/12/21
                      .ChkPCT = Me.ChkPCT
                      .bolIsL = True
                      'Add by Morgan 2008/11/21
                      .m_Sys = txt1(2)
                      .m_Type = txt1(3)
                      .m_Date1 = txt1(4)
                      .m_Date2 = txt1(5)
                      .m_Pty1 = txt1(6)
                      .m_Pty2 = txt1(7)
                      .m_Cty1 = txt1(8)
                      .m_Cty2 = txt1(8)
                      'end 2008/11/21
                      .StrMenu
                      End With
                      Screen.MousePointer = vbDefault
                  End If
                  Me.Enabled = True
                  Exit Sub
               End If
           End If
           Next i
           Me.Enabled = True
      'Add by Morgan 2007/12/18
      Case 6 '往來紀錄
            Me.Enabled = False
            For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               'Add By Sindy 2012/3/21
               grdDataList.col = 1
               'Modify by Amy 2013/12/10 +判斷對造
               If Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, 5) = "對造" Then
                  For j = 0 To grdDataList.Cols - 1
                     grdDataList.col = j
                     grdDataList.CellBackColor = &H8080FF
                  Next j
               Else
               '2012/3/21 End
                  For j = 0 To grdDataList.Cols - 1
                     If j <> 1 Then
                         grdDataList.col = j
                         grdDataList.CellBackColor = QBColor(15)
                     End If
                  Next j
               End If
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               grdDataList.col = 1
               Screen.MousePointer = vbHourglass
               strExc(1) = Pub_RplStr(grdDataList.Text)
               
               'Modify By Sindy 2010/02/23 判斷是國外或是國內潛在客戶
               '客戶檔
               strExc(3) = "select cu12,cu13 from customer where cu01(+)='" & Left(strExc(1), 8) & "' and cu02(+)='" & Mid(strExc(1), 9, 1) & "' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(3))
               strExc(4) = ""
               If intI = 1 Then
                  strExc(4) = "" & RsTemp.Fields("cu12")
               End If
               '潛在客戶檔
               strExc(0) = "select * from potcustomer where pcu01(+)='" & Left(strExc(1), 8) & "' and pcu02(+)='" & Mid(strExc(1), 9, 1) & "' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               strExc(2) = ""
               If intI = 1 Then
                  strExc(2) = "" & RsTemp.Fields(0)
               End If
               If strExc(2) <> "" Or Left(Trim(strExc(1)), 1) = "Y" Or Left(Trim(strExc(4)), 1) = "F" Then '國外
                  frm100101_15.Show
                  frm100101_15.Tag = strExc(1)
                  frm100101_15.StrMenu
               Else '國內
                  frm100101_20.Show
                  frm100101_20.Tag = strExc(1)
                  frm100101_20.StrMenu
               End If
               
               Screen.MousePointer = vbDefault
               grdDataList.col = 0
               grdDataList.Text = ""
               'Add By Sindy 2012/3/21
               grdDataList.col = 1
               'Modify by Amy 2013/12/10 +判斷對造
               If Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, 5) = "對造" Then
                  For j = 0 To grdDataList.Cols - 1
                     grdDataList.col = j
                     grdDataList.CellBackColor = &H8080FF
                  Next j
               Else
               '2012/3/21 End
                  For j = 0 To grdDataList.Cols - 1
                     If j <> 1 Then
                         grdDataList.col = j
                         grdDataList.CellBackColor = QBColor(15)
                     End If
                  Next j
               End If
               Me.Enabled = True
               Exit Sub
            End If
            Next i
            Me.Enabled = True
      'Add by Morgan 2008/7/23
      Case 7 '聯絡人
            Me.Enabled = False
            For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
               grdDataList.col = 0
               grdDataList.Text = ""
               'Add By Sindy 2012/3/21
               grdDataList.col = 1
               'Modify by Amy 2013/12/10 +判斷對造
               If Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, 5) = "對造" Then
                  For j = 0 To grdDataList.Cols - 1
                     grdDataList.col = j
                     grdDataList.CellBackColor = &H8080FF
                  Next j
               Else
               '2012/3/21 End
                  For j = 0 To grdDataList.Cols - 1
                     If j <> 1 Then
                         grdDataList.col = j
                         grdDataList.CellBackColor = QBColor(15)
                     End If
                  Next j
               End If
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               grdDataList.col = 1
               Screen.MousePointer = vbHourglass
               strExc(1) = Pub_RplStr(grdDataList.Text)
               'Modify by Morgan 2008/8/5 國內外客戶跑不同畫面
               Select Case Left(strExc(1), 1)
                  'Add by Morgan 2008/9/1 潛在客戶跑申請人資料畫面
                  Case "R"
                     frm100101_14.Show
                     frm100101_14.Tag = strExc(1)
                     frm100101_14.StrMenu
                     
                  Case Else
                     strExc(2) = "F"
                     If Left(strExc(1), 1) = "X" Then
                        strExc(0) = "select st03 from customer,staff where cu01(+)='" & Left(strExc(1), 8) & "' and cu02(+)='" & Mid(strExc(1), 9, 1) & "' and st01(+)=cu13"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           strExc(2) = "" & RsTemp.Fields(0)
                        End If
                     End If
                     If Left(strExc(2), 1) = "F" Then
                        frm100101_17.Show
                        frm100101_17.Tag = strExc(1)
                        frm100101_17.StrMenu
                     Else
                        frm100101_18.Show
                        frm100101_18.Tag = strExc(1)
                        frm100101_18.StrMenu
                     End If
               End Select
               'end 2008/8/5
               
               Screen.MousePointer = vbDefault
               grdDataList.col = 0
               grdDataList.Text = ""
               'Add By Sindy 2012/3/21
               grdDataList.col = 1
               'Modify by Amy 2013/12/10 +判斷對造
               If Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, 5) = "對造" Then
                  For j = 0 To grdDataList.Cols - 1
                     grdDataList.col = j
                     grdDataList.CellBackColor = &H8080FF
                  Next j
               Else
               '2012/3/21 End
                  For j = 0 To grdDataList.Cols - 1
                     If j <> 1 Then
                         grdDataList.col = j
                         grdDataList.CellBackColor = QBColor(15)
                     End If
                  Next j
               End If
               Me.Enabled = True
               Exit Sub
            End If
            Next i
            Me.Enabled = True
      'Add by Amy 2013/12/04
      Case 8 '列印對造資料
            'Modify by Amy 2014/02/25 改印暫存資料
            'PrintDataA4
            PrintDataA4_Temp
            'end 2014/02/25
      Case Else
   End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
   'add by nickc 2007/01/12
   If Len(Trim(Me.txt1(2).Text)) = 0 Then
       Me.txt1(2).Text = "ALL"
   End If
   '92.04.16 nick 紀錄作用按鍵
   CmdState = Index
   PubShowNextData
   Exit Sub
End Sub

Sub StrMenu(StrToGrid)
   'Modify By Cheng 2004/03/02
   'strSQL = "SELECT FA01||FA02,DECODE(FA10,'013',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)),NA03 FROM FAGENT,NATION WHERE FA01>='" & Left(StrToGrid, 6) & "00' AND FA01<='" & Left(StrToGrid, 6) & "zz' AND fa10=NA01(+) "
   'strSQL = strSQL & " union all SELECT cu01||cu02,DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)),NA03 FROM customer,NATION WHERE cu01>='" & Left(StrToGrid, 6) & "00' AND cu01<='" & Left(StrToGrid, 6) & "zz' AND cu10=NA01(+) "
   strSql = "SELECT FA01||FA02||Decode(FA02,'0','','＊'),SUBSTR(DECODE(FA10,'013',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)),1,80),NA03 FROM FAGENT,NATION WHERE FA01>='" & Left(StrToGrid, 6) & "00' AND FA01<='" & Left(StrToGrid, 6) & "zz' AND fa10=NA01(+) "
   strSql = strSql & " union all SELECT cu01||cu02||Decode(CU02,'0','','＊'),SUBSTR(DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)),1,80),NA03 FROM customer,NATION WHERE cu01>='" & Left(StrToGrid, 6) & "00' AND cu01<='" & Left(StrToGrid, 6) & "zz' AND cu10=NA01(+) "
   'End
   'Add By Sindy 98/03/20
   strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),NA03 FROM PotCustomer,Nation WHERE PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz'   AND NA01(+)=PCU09"
   strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','＊'),POC03,NA03 FROM PotCustomer1,Nation WHERE POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz'   AND NA01(+)=POC04"
   '傳入R1時找出相關的X
   strSql = strSql & " union  SELECT cu01||cu02||Decode(CU02,'0','','＊'),SUBSTR(DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)),1,80),NA03 " & _
                                                    "From CUSTOMER, PotCustomer1, Nation " & _
                                               "WHERE CU10=NA01(+) " & _
                                                    "AND POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND CU01>=(substr(POC16,1,6)||'00') AND CU01<=(substr(POC16,1,6)||'zz') " & _
                                                    "AND POC16 is not null "
   '找出R1的關係企業
   strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','＊'),POC03,NA03 " & _
                                                    "From PotCustomer1, Nation " & _
                                                "WHERE NA01(+)=POC04 " & _
                                                     "AND POC16>='" & Left(StrToGrid, 6) & "00' AND POC16<='" & Left(StrToGrid, 6) & "zz' " & _
                                                     "AND POC16 is not null "
   '傳入R1時找出相關的R
   strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),NA03 " & _
                                                    "From PotCustomer, Nation, PotCustomer1 " & _
                                               "WHERE NA01(+)=PCU09 " & _
                                                    "AND POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND PCU47>=(substr(POC16,1,6)||'00') AND PCU47<=(substr(POC16,1,6)||'zz') " & _
                                                    "AND POC16 is not null AND PCU47 is not null "
   '98/03/19 End
   'Add By Sindy 2009/06/24
   '傳入R時找出相關的X
   strSql = strSql & " union  SELECT cu01||cu02||Decode(CU02,'0','','＊'),SUBSTR(DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)),1,80),NA03 " & _
                                                    "From CUSTOMER, PotCustomer, Nation " & _
                                               "WHERE CU10=NA01(+) " & _
                                                    "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND CU01>=(substr(PCU47,1,6)||'00') AND CU01<=(substr(PCU47,1,6)||'zz') " & _
                                                    "AND PCU47 is not null "
   '傳入R時找出相關的Y
   strSql = strSql & " union  SELECT FA01||FA02||Decode(FA02,'0','','＊'),NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),NA03 " & _
                                                    "From Fagent, PotCustomer, Nation " & _
                                                "WHERE NA01(+)=FA10 " & _
                                                     "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                     "AND FA01>=(substr(PCU47,1,6)||'00') AND FA01<=(substr(PCU47,1,6)||'zz') " & _
                                                     "AND PCU47 is not null "
   '找出R的關係企業
   strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)),NA03 " & _
                                                    "From PotCustomer, Nation " & _
                                               "WHERE NA01(+)=PCU09 " & _
                                                    "AND PCU47>='" & Left(StrToGrid, 6) & "00' AND PCU47<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND PCU47 is not null "
   '傳入R時找出相關的R1
   strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','＊'),POC03,NA03 " & _
                                                    "From PotCustomer1, Nation, PotCustomer " & _
                                               "WHERE NA01(+)=POC04 " & _
                                                    "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                    "AND POC16>=(substr(PCU47,1,6)||'00') AND POC16<=(substr(PCU47,1,6)||'zz') " & _
                                                    "AND PCU47 is not null AND POC16 is not null "
   '2009/06/24 End
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 Then
       adoRecordset.MoveFirst
       'cnnConnection.Execute "delete from r100114 where id='" & strUserNum & "' " Mark by Amy 2015/03/27
       Do While adoRecordset.EOF = False
       
       strSql = "INSERT INTO R100114 values ('"
       If Not IsNull(adoRecordset.Fields(0)) Then
           strSql = strSql + ChgSQL(adoRecordset.Fields(0)) + "','"
       Else
           strSql = strSql + "','"
       End If
       If Not IsNull(adoRecordset.Fields(1)) Then
           strSql = strSql & ChgSQL(adoRecordset.Fields(1)) + "','"
       Else
           strSql = strSql + "','"
       End If
       If Not IsNull(adoRecordset.Fields(2)) Then
           strSql = strSql + ChgSQL(adoRecordset.Fields(2)) + "','" & strUserNum & "')"
       Else
           strSql = strSql + "','" & strUserNum & "')"
       End If
       cnnConnection.Execute strSql
       adoRecordset.MoveNext
       Loop
   Else
   End If
   CheckOC
End Sub

Sub StrMenu1()
   Screen.MousePointer = vbHourglass
   '92.10.15 MODIFY BY SONIA
   'strSQL = "SELECT '' AS V,R07001 AS 編號,R07002 AS 名稱,R07003 AS 國籍 FROM R100114 ORDER BY 編號"
   'edit by nickc 2005/12/06
   ' strSQL = "SELECT '' AS V,R07001 AS 編號,R07002 AS 名稱,R07003 AS 國籍,CU80 AS 狀態,CU79 AS 備註 FROM R100114,CUSTOMER where SUBSTR(R07001,1,1)='X' AND SUBSTR(R07001,1,8)=CU01(+) AND SUBSTR(R07001,9,1)=CU02(+)"
   'strSQL = strSQL & "UNION ALL SELECT '' AS V,R07001 AS 編號,R07002 AS 名稱,R07003 AS 國籍,FA69 AS 狀態,FA29 AS 備註 FROM R100114,FAGENT where SUBSTR(R07001,1,1)='Y' AND SUBSTR(R07001,1,8)=FA01(+) AND SUBSTR(R07001,9,1)=FA02(+) ORDER BY 編號"
   'Modify by Amy 2013/12/10 +智權人/申請國家/總收文號/案件性質/收文日
   'Modify by Amy 2015/05/12 智權人 抓ST02
    strSql = "SELECT '' AS V,R07001||decode(cu111,'Y','$','') AS 編號,R07002 AS 名稱,R07003 AS 國籍,ST02 as 智權人員,CU80 AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM R100114,CUSTOMER,Staff where id='" & strUserNum & "' And SUBSTR(R07001,1,1)='X' AND SUBSTR(R07001,1,8)=CU01(+) AND SUBSTR(R07001,9,1)=CU02(+) AND CU13=ST01(+)"
   strSql = strSql & "UNION ALL SELECT '' AS V,R07001||decode(fa77,'Y','$','') AS 編號,R07002 AS 名稱,R07003 AS 國籍,'' as 智權人員,FA69 AS 狀態,FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM R100114,FAGENT where id='" & strUserNum & "' And SUBSTR(R07001,1,1)='Y' AND SUBSTR(R07001,1,8)=FA01(+) AND SUBSTR(R07001,9,1)=FA02(+)"
   '92.10.15 END
   'Add By Sindy 98/03/20
   'Modify by Amy 2015/05/12 智權人 抓ST02
   strSql = strSql & "UNION ALL SELECT '' AS V,R07001 AS 編號,R07002 AS 名稱,R07003 AS 國籍,ST02 as 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM R100114,POTCUSTOMER,Staff where id='" & strUserNum & "' AND SUBSTR(R07001,1,1)='R' AND SUBSTR(R07001,1,8)=PCU01 AND SUBSTR(R07001,9,1)=PCU02 and substr(LTrim(PCU38),1,5)=ST01(+) "
   strSql = strSql & "UNION ALL SELECT '' AS V,R07001 AS 編號,R07002 AS 名稱,R07003 AS 國籍,ST02 as 智權人員,POC14 AS 狀態,POC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM R100114,POTCUSTOMER1,Staff where id='" & strUserNum & "' AND SUBSTR(R07001,1,1)='R' AND SUBSTR(R07001,1,8)=POC01 AND SUBSTR(R07001,9,1)=POC02 and POC13=ST01(+) ORDER BY 編號"
   'end 2015/05/12
   '98/03/20 End
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   '911029 nick edit
   If adoRecordset.RecordCount <> 0 Then
       Set grdDataList.Recordset = adoRecordset
   End If
   CheckOC
   SetDataListWidth
   If Me.grdDataList.Rows = 2 Then
      '911029 nick add
      grdDataList.row = 1
      grdDataList.col = 1
      If grdDataList.Text <> "" Then
      '911029 nick end
           grdDataList.Visible = False
           grdDataList.row = 1
           grdDataList.col = 0
           grdDataList.Text = "V"
           For i = 0 To grdDataList.Cols - 1
               grdDataList.col = i
               grdDataList.CellBackColor = &HFFC0C0
           Next i
           grdDataList.Visible = True
      '911029 nick add
      End If
      '911029 nick end
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
'Add By Cheng 2002/07/09
Dim StrSQLa As String
'910801 nick
Dim StrSqlB As String
Dim strSQLc As String
Dim strSQLD As String 'Add By Sindy 2011/10/11
Dim strCheckWay As String
Dim strSQLE As String 'Add By Sindy 2012/3/21
'Add by Amy 2013/12/04
Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim strSwhSQL1 As String, strSwhSQL2 As String
Dim strSubSQL1 As String, strSubSQL2 As String
Dim strNo As String, Str01 As String
bolPrint = False '先設定無對造
StrToPrint = ""
'end 2013/12/04

   
   'Modify By Cheng 2002/03/14
   ''Add By Cheng 2002/01/07
   'txt1_LostFocus 2
   If Option1(0).Value = True Then
       If Len(Trim(txt1(0))) = 0 Then
           s = MsgBox("編號不可空白", , "USER 輸入資料錯誤")
           txt1(0).SetFocus
           Exit Sub
       End If
   Else
       If Option1(1).Value = True Then
           If Len(Trim(txt1(1))) = 0 Then
               s = MsgBox("名稱不可空白", , "USER 輸入資料錯誤")
               txt1(1).SetFocus
               Exit Sub
           End If
       Else
           If Option1(2).Value = True Then
               If Len(Trim(txt1(9))) = 0 Then
                   s = MsgBox("國籍不可空白", , "USER 輸入資料錯誤")
                   txt1(9).SetFocus
                   Exit Sub
               End If
           End If
       End If
   End If
   
   'add by Toni 2008/12/03
   If Option1(3).Value = True Then
       If Len(Trim(txt1(10))) = 0 Then
           s = MsgBox("條件不可空白", , "輸入條件錯誤")
           txt1(10).SetFocus
           Exit Sub
       End If
   End If
   
   'If Len(Trim(txt1(3))) = 0 Then
   '    S = MsgBox("查詢別不可空白", , "USER 輸入資料錯誤")
   '    txt1(3).SetFocus
   '    Exit Sub
   'End If
   'If Len(Trim(txt1(4))) = 0 Or Len(Trim(txt1(5))) = 0 Then
   '    S = MsgBox("日期區間不可空白", , "USER 輸入資料錯誤")
   '    If Len(Trim(txt1(5))) = 0 Then txt1(5).SetFocus
   '    If Len(Trim(txt1(4))) = 0 Then txt1(4).SetFocus
   '    Exit Sub
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/4 清除查詢印表記錄檔欄位
   Screen.MousePointer = vbHourglass
   'End If
   grdDataList.Clear
   grdDataList.Rows = 2
   SetDataListWidth
   'Add By Cheng 2002/07/09
   '若國籍為"013"或"020"則名稱抓中-->英-->日, 否則抓英-->中-->日
   StrSQLa = "DECODE(FA10,'013',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 名稱,"
   StrSqlB = "DECODE(cu10,'013',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),'020',NVL(cu04,DECODE(cu05,NULL,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)),DECODE(cu05,NULL,NVL(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90)) as 名稱,"
   'Add by Morgan 2007/12/14
   strSQLc = "DECODE(instr('013,020',pcu09),0,decode(pcu03,NULL,nvl(pcu08,pcu07),rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06)),NVL(pcu08,DECODE(pcu03,NULL,pcu07,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06)))) as 名稱,"
   'Add By Sindy 2011/10/11
   strSQLD = "DECODE(instr('013,020',poc04),0,decode(poc23,NULL,nvl(poc03,poc27),rtrim(poc23||' '||poc24||' '||poc25||' '||poc26)),NVL(poc03,DECODE(poc23,NULL,poc27,rtrim(poc23||' '||poc24||' '||poc25||' '||poc26)))) as 名稱,"
   'Add By Sindy 2012/3/21
   strSQLE = "DECODE(instr('013,020',nt08),0,decode(nt03,NULL,nvl(nt02,nt07),rtrim(nt03||' '||nt04||' '||nt05||' '||nt06)),NVL(nt02,DECODE(nt03,NULL,nt07,rtrim(nt03||' '||nt04||' '||nt05||' '||nt06)))) as 名稱,"
   
   'Modify by Amy 2013/10/30 讀取Fagent及Customer的狀態欄時，先檢查FA103或CU142，有值顯示 處理情形的內容，無值才抓原狀態欄位
   'Modify by Amy 2013/09/30 trim掉空白檢查:編號,名稱,E-Mail
   'Modify by Morgan 2007/12/14 程式邏輯整理
   '以編號查詢
   If Option1(0).Value = True Then
       'Modify by Amy 2013/12/04 +智權人/申請國家/總收文號/案件性質/收文日
      'Modify by Morgan 2007/12/14 加可查潛在客戶
      If UCase(Left(Trim(txt1(0)), 1)) = "R" Then
         strSql = "SELECT ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號," & strSQLc & "NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM POTCUSTOMER,NATION,STAFF WHERE PCU09=NA01(+) AND PCU01='" & Left(GetNewFagent(Trim(txt1(0))), 8) & "' and substr(LTrim(PCU38),1,5)=ST01(+) "
         'Add By Sindy 2011/10/11
         strSql = strSql & " union all SELECT ' ' AS V ,PoC01||PoC02||Decode(PoC02,'0','','＊') AS 編號," & strSQLD & "NA03 AS 國籍,ST02 AS 智權人員,PoC14 AS 狀態,PoC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM POTCUSTOMER1,NATION,STAFF WHERE PoC04=NA01(+) AND PoC01='" & Left(GetNewFagent(Trim(txt1(0))), 8) & "' and poc13=ST01(+) "
      Else
         'edit by nickc 2005/12/06
         'strSQL = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊') AS 編號," & StrSQLa & "NA03 AS 國籍,FA69 AS 狀態, FA29 AS 備註 FROM FAGENT,NATION WHERE FA01='" & Left(GetNewFagent(txt1(0)), 8) & "' AND fa10=na01(+) "
         strSql = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號," & StrSQLa & "NA03 AS 國籍,' ' as 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM FAGENT,NATION WHERE FA01='" & Left(GetNewFagent(Trim(txt1(0))), 8) & "' AND fa10=na01(+) "
         'edit by nickc 2005/12/06
         'strSQL = strSQL & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊') AS 編號," & StrSqlB & "NA03 AS 國籍,CU80 AS 狀態,CU79 AS 備註 FROM customer,NATION WHERE cu01='" & Left(GetNewFagent(txt1(0)), 8) & "' AND cu10=na01(+) "
         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號," & StrSqlB & "NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM customer,NATION,STAFF WHERE cu01='" & Left(GetNewFagent(Trim(txt1(0))), 8) & "' AND cu10=na01(+) AND CU13=ST01(+) "
         'Add By Sindy 2012/3/21
         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號," & strSQLE & "NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from notagent,nation,STAFF where nt08=na01(+) and nt01='" & IIf(Len(Trim(txt1(0))) >= 3, Trim(txt1(0)), Right("000" & Trim(txt1(0)), 3)) & "' AND nt18=ST01(+) "
      End If
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Trim(txt1(0)) 'Add By Sindy 2010/11/4
   '以名稱查詢
   ElseIf Option1(1).Value = True Then
      pub_QL05 = pub_QL05 & ";" & Option1(1).Caption 'Add By Sindy 2010/11/4
      '模糊比對
      If Option3(0).Value = False Then
         strCheckWay = ">0"
         pub_QL05 = pub_QL05 & ";" & Option3(1).Caption 'Add By Sindy 2010/11/4
      '字首比對
      Else
         strCheckWay = "=1"
         pub_QL05 = pub_QL05 & ";" & Option3(0).Caption 'Add By Sindy 2010/11/4
      End If
      'Add by Amy 2013/12/04
      strTp(3) = ChgSQL(UCase(Trim(txt1(1))))
      '對造
      strSQL1 = " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
      strSQL2 = " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
      StrSQL3 = " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
      StrSQL4 = " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
      strSQL5 = " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
      'end 2013/12/04
'Modify by Amy 2013/12/04 拿掉中英日 +查對造 +智權人,申請國家,總收文號,案件性質,收文日 欄位
'      '以中文名稱查詢
'      If Option2(0).Value = True Then
'         'edit by nickc 2005/12/06
'         'strSQL = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊') AS 編號,FA04 AS 名稱,NA03 AS 國籍,FA69 AS 狀態, FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA04,'" & ChgSQL(txt1(1)) & "')>0 ) A WHERE FA01=A.A1 AND fa10=na01(+) "
'         strSql = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA04 AS 名稱,NA03 AS 國籍,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA04,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=na01(+) "
'         'End
'         'Add by Morgan 2007/12/14
'         strSql = strSql & " union all SELECT '' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,pcu08 AS 名稱,NA03 AS 國籍,pcu39 AS 狀態, pcu40 AS 備註 FROM POTCUSTOMER,NATION, (Select Distinct pcu01 As A1 From POTCUSTOMER Where instr(pcu08,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE pcu01=A.A1 AND pcu09=na01(+) "
'         'end 2007/12/14
'         'edit by nickc 2005/12/06
'         'strSQL = strSQL & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊') AS 編號,cu04 AS 名稱,NA03 AS 國籍,CU80 AS 狀態,CU79 AS 備註 FROM customer,NATION, (Select Distinct CU01 As A1 From Customer Where instr(cu04,'" & ChgSQL(txt1(1)) & "')>0 ) A WHERE CU01=A.A1 And cu10=na01(+)"
'         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu04 AS 名稱,NA03 AS 國籍,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,CU79 AS 備註 FROM customer,NATION, (Select Distinct CU01 As A1 From Customer Where instr(cu04,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 And cu10=na01(+)"
'         'End
'
'         'Add By Sindy 98/03/20
'         strSql = strSql & " union all SELECT '' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,poc03 AS 名稱,NA03 AS 國籍,poc14 AS 狀態, poc15 AS 備註 FROM POTCUSTOMER1,NATION, (Select Distinct poc01 As A1 From POTCUSTOMER1 Where instr(poc03,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE poc01=A.A1 AND poc04=na01(+) "
'         '98/03/20 End
'
'         'Add by Morgan 2007/12/21 加可查聯絡人
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,CUSTOMER,NATION WHERE CU10=NA01(+) AND CU01(+)=PCC01 AND CU02='0' "
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,PCU39 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,potcustomer,nation where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' "
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, PCC13 AS 備註 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,POC14 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,potcustomer1,nation where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' "
'         'end 2007/12/21
'         'Add By Sindy 2012/3/21
'         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,nt02 AS 名稱,NA03 AS 國籍,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註 from notagent,nation, (Select Distinct nt01 As A1 From notagent Where instr(nt02,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A where nt08=na01(+) and nt01=A.A1"
'         pub_QL05 = pub_QL05 & ";" & Option2(0).Caption & ";" & Trim(txt1(1)) 'Add By Sindy 2010/11/4
'
'      '以英文名稱查詢
'      ElseIf Option2(1).Value = True Then
'         'edit by nickc 2005/12/06
'         'strSQL = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊') AS 編號,FA05||' '||FA63||' '||FA64||' '||FA65 AS 名稱,NA03 AS 國籍,FA69 AS 狀態, FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(upper(FA05||' '||FA63||' '||FA64||' '||FA65),'" & UCase(ChgSQL(txt1(1))) & "')>0 ) A WHERE FA01=A.A1 AND fa10=NA01(+) "
'         'strSQL = strSQL & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊') AS 編號,cu05||' '||cu88||' '||cu89||' '||cu90 AS 名稱,NA03 AS 國籍,CU80 AS 狀態,CU79 AS 備註 FROM customer,NATION, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(txt1(1))) & "')>0 ) A WHERE CU01=A.A1 AND cu10=na01(+)"
'         strSql = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA05||' '||FA63||' '||FA64||' '||FA65 AS 名稱,NA03 AS 國籍,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(upper(FA05||' '||FA63||' '||FA64||' '||FA65),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=NA01(+) "
'         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu05||' '||cu88||' '||cu89||' '||cu90 AS 名稱,NA03 AS 國籍,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,CU79 AS 備註 FROM customer,NATION, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 AND cu10=na01(+)"
'         'End
'         'Add by Morgan 2007/12/14
'         strSql = strSql & " union all SELECT '' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06) AS 名稱,NA03 AS 國籍,pcu39 AS 狀態, pcu40 AS 備註 FROM POTCUSTOMER,NATION, (Select Distinct pcu01 As A1 From POTCUSTOMER Where instr(upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & " ) A WHERE pcu01=A.A1 AND pcu09=na01(+) "
'         'end 2007/12/14
'         'Add By Sindy 2010/02/12
'         strSql = strSql & " union all SELECT '' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,rtrim(poc23||' '||poc24||' '||poc25||' '||poc26) AS 名稱,NA03 AS 國籍,poc14 AS 狀態, poc15 AS 備註 FROM POTCUSTOMER1,NATION, (Select Distinct poc01 As A1 From POTCUSTOMER1 Where instr(upper(poc23||' '||poc24||' '||poc25||' '||poc26),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & " ) A WHERE poc01=A.A1 AND poc04=na01(+) "
'         '2010/02/12 End
'         'Add by Morgan 2007/12/21 加可查聯絡人
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & ") A,CUSTOMER,NATION WHERE CU10=NA01(+) AND CU01(+)=PCC01 AND CU02='0' "
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,PCU39 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & ") A,potcustomer,nation where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' "
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, PCC13 AS 備註 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,POC14 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & ") A,potcustomer1,nation where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' "
'         'end 2007/12/21
'         'Add By Sindy 2012/3/21
'         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT03||' '||NT04||' '||NT05||' '||NT06 AS 名稱,NA03 AS 國籍,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註 from notagent,nation, (Select Distinct NT01 As A1 From notagent Where instr(upper(NT03||' '||NT04||' '||NT05||' '||NT06),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & " ) A where nt08=na01(+) and nt01=A.A1"
'         pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & ";" & Trim(txt1(1)) 'Add By Sindy 2010/11/4
'
'      '以日文名稱查詢
'      Else
'         'edit by nickc 2005/12/06
'         'strSQL = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊') AS 編號,FA06 AS 名稱,NA03 AS 國籍,FA69 AS 狀態, FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA06,'" & ChgSQL(txt1(1)) & "')>0 ) A WHERE FA01=A.A1 AND fa10=na01(+) "
'         strSql = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA06 AS 名稱,NA03 AS 國籍,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA06,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=na01(+) "
'         'End
'         'Add by Morgan 2007/12/14
'         strSql = strSql & " union all SELECT '' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,pcu07 AS 名稱,NA03 AS 國籍,pcu39 AS 狀態, pcu40 AS 備註 FROM POTCUSTOMER,NATION, (Select Distinct pcu01 As A1 From POTCUSTOMER Where instr(pcu07,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE pcu01=A.A1 AND pcu09=na01(+) "
'         'end 2007/12/14
'
'         'Add By Sindy 2010/02/12
'         strSql = strSql & " union all SELECT '' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,poc27 AS 名稱,NA03 AS 國籍,poc14 AS 狀態, poc15 AS 備註 FROM POTCUSTOMER1,NATION, (Select Distinct poc01 As A1 From POTCUSTOMER1 Where instr(poc27,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE poc01=A.A1 AND poc04=na01(+) "
'         '2010/02/12 End
'
'         'edit by nickc 2005/12/06
'         'strSQL = strSQL & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊') AS 編號,cu06 AS 名稱,NA03 AS 國籍,CU80 AS 狀態,CU79 AS 備註 FROM customer,NATION, (Select Distinct CU01 As A1 From Customer Where instr(cu06,'" & ChgSQL(txt1(1)) & "')>0 ) A WHERE CU01=A.A1 AND cu10=na01(+)"
'         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu06 AS 名稱,NA03 AS 國籍,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,CU79 AS 備註 FROM customer,NATION, (Select Distinct CU01 As A1 From Customer Where instr(cu06,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 AND cu10=na01(+)"
'         'End
'         'Add by Morgan 2007/12/21 加可查聯絡人
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,CUSTOMER,NATION WHERE CU10=NA01(+) AND CU01(+)=PCC01 AND CU02='0' "
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,PCU39 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,potcustomer,nation where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' "
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, PCC13 AS 備註 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,POC14 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,potcustomer1,nation where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' "
'         'end 2007/12/21
'         'Add By Sindy 2012/3/21
'         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT07 AS 名稱,NA03 AS 國籍,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註 from notagent,nation, (Select Distinct NT01 As A1 From Notagent Where instr(NT07,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A where nt08=na01(+) and nt01=A.A1"
'         pub_QL05 = pub_QL05 & ";" & Option2(2).Caption & ";" & Trim(txt1(1)) 'Add By Sindy 2010/11/4
'      End If

      'Modify by Amy 2014/02/25 對造由下搬上來改語法存至暫存檔
            cnnConnection.Execute "Delete From R100102_1 Where ID='" & strUserNum & "' "
            '對造(中)
            strSubSQL1 = " And InStr(Upper(CP40),'" & strTp(3) & "') " & strCheckWay
            strSubSQL2 = " And InStr(Upper(CP50),'" & strTp(3) & "') " & strCheckWay
            strSwhSQL1 = " CP40>' ' "
            strSwhSQL2 = " CP50>' ' "
            
      'Modify by Amy 2015/03/27 拿掉對造案件編號符號,+客戶端平台帳號資料
            '商標
            strSql = "Insert Into R100102_1 (r021001,r021002,r021003,r021004,r021005,r021006,r021007,r021008,r021009,r021010,r021011,r021012,r021013,r021014,r021015,r021016,r021017,r021018,ID) " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP40 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap ,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL1
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP50 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL2
            '專利
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP40 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL1
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP50 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL2
            '法務
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP40 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & StrSQL3 & strSubSQL1
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP50 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家 ,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & StrSQL3 & strSubSQL2
            '顧問
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP40 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & StrSQL4 & strSubSQL1
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP50 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & StrSQL4 & strSubSQL2
            '服務
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP40 as 名稱,' ' as 智權人,'1' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & strSQL5 & strSubSQL1
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP50 as 名稱,' ' as 智權人,'2' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & strSQL5 & strSubSQL2
            
            '對造(英)
            strSubSQL1 = " And InStr(Upper(CP41),'" & strTp(3) & "') " & strCheckWay
            strSubSQL2 = " And InStr(Upper(CP51),'" & strTp(3) & "') " & strCheckWay
            strSwhSQL1 = " CP41>' ' "
            strSwhSQL2 = " CP51>' ' "
            '商標
            strSql = strSql & " Union " & _
                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP41 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL1
            strSql = strSql & " Union " & _
                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP51 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL2
            '專利
            strSql = strSql & " Union " & _
                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP41 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL1
            strSql = strSql & " Union " & _
                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP51 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL2
            '法務
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP41 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & StrSQL3 & strSubSQL1
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP51 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & StrSQL3 & strSubSQL2
            '顧問
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP41 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & StrSQL4 & strSubSQL1
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP51 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & StrSQL4 & strSubSQL2
            '服務
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP41 as 名稱,' ' as 智權人,'1' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & strSQL5 & strSubSQL1
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP51 as 名稱,' ' as 智權人,'2' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & strSQL5 & strSubSQL2
            
            '對造(日)
            strSubSQL1 = " And InStr(Upper(CP42),'" & strTp(3) & "') " & strCheckWay
            strSubSQL2 = " And InStr(Upper(CP52),'" & strTp(3) & "') " & strCheckWay
            strSwhSQL1 = " CP42>' ' "
            strSwhSQL2 = " CP52>' ' "
            '商標
            strSql = strSql & " Union " & _
                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP42 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL1
            strSql = strSql & " Union " & _
                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP52 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & strSQL1 & strSubSQL2
            '專利
            strSql = strSql & " Union " & _
                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP42 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL1
            strSql = strSql & " Union " & _
                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP52 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & strSQL2 & strSubSQL2
            '法務
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP42 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & StrSQL3 & strSubSQL1
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP52 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & StrSQL3 & strSubSQL2
            '顧問
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP42 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & StrSQL4 & strSubSQL1
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP52 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & StrSQL4 & strSubSQL2
            '服務
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP42 as 名稱,' ' as 智權人,'1' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & strSQL5 & strSubSQL1
            strSql = strSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP52 as 名稱,' ' as 智權人,'2' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & strSQL5 & strSubSQL2
            cnnConnection.Execute strSql
           
           '刪除對造與申請人相同資料
           strSql = "Delete From R100102_1 Where ID='" & strUserNum & "' And (ltrim(rtrim(R021002))=ltrim(rtrim(R021008)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021009)) " & _
                       "Or ltrim(rtrim(R021002))=ltrim(rtrim(R021010)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021011)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021012))) "
           cnnConnection.Execute strSql
      'end 2014/02/25
      
      'Add by Amy 2014/03/17 將所有商標案InStr(R021014,'T')且案件性質為1202(核准通知)者狀態改為 其他相關人
      'Modify by Amy 2015/12/03 增加商標案(CFC/S) 案件性質202(申請意見書)及303(延期)者 狀態改為 其他相關人
      strSql = "Update R100102_1 Set R021004='2' Where (InStr(R021014,'T')>0 or R021014='CFC' or R021014='S') And (R021018='1202' or R021018='202' or R021018='303')"
      cnnConnection.Execute strSql
      'end 2014/03/17
      'Add by Amy 2015/12/03 所有專利案件性質404(延期) 者狀態改為 其他相關人
      strSql = "Update R100102_1 Set R021004='2' Where (InStr(R021014,'P')>0 or R021014='FG') And R021018='404' "
      cnnConnection.Execute strSql
      'end 2015/12/03
       
      '查Fagent 代理人 檔
         strSql = "SELECT '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA04 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA04,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=na01(+) "
         strSql = strSql & " union all Select '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA05||' '||FA63||' '||FA64||' '||FA65 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(upper(FA05||' '||FA63||' '||FA64||' '||FA65),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=NA01(+) "
         strSql = strSql & " union all Select '' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,FA06 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where instr(FA06,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE FA01=A.A1 AND fa10=na01(+) "

      '查customer 客戶 檔
         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM customer,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(cu04,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 And cu10=na01(+) AND CU13=ST01(+) "
         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu05||' '||cu88||' '||cu89||' '||cu90 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM customer,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 AND cu10=na01(+) AND CU13=ST01(+) "
         strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,cu06 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM customer,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where instr(cu06,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE CU01=A.A1 AND cu10=na01(+) AND CU13=ST01(+) "

      'Modify by Amy 2015/04/15 客戶端平台帳號資料
      strSql = strSql & " union all Select ' ' as V,'平台'||CW01 AS 編號, CW12 AS 名稱,'平台' AS 國籍,' ' AS 智權人員,Nvl(CW19,'') AS 狀態,'' AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,CW01 as 收文日 From CustWeb Where InStr(Upper(CW12),'" & strTp(3) & "') " & strCheckWay
 
      '查potcustomer 國外潛在客戶 檔
         strSql = strSql & " union all SELECT '' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,pcu08 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,pcu39 AS 狀態, pcu40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM POTCUSTOMER,NATION,STAFF, (Select Distinct pcu01 As A1 From POTCUSTOMER Where instr(pcu08,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE pcu01=A.A1 AND pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
         strSql = strSql & " union all SELECT '' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,pcu39 AS 狀態, pcu40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM POTCUSTOMER,NATION,STAFF, (Select Distinct pcu01 As A1 From POTCUSTOMER Where instr(upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & " ) A WHERE pcu01=A.A1 AND pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
         strSql = strSql & " union all SELECT '' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,pcu07 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,pcu39 AS 狀態, pcu40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM POTCUSTOMER,NATION,STAFF, (Select Distinct pcu01 As A1 From POTCUSTOMER Where instr(pcu07,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE pcu01=A.A1 AND pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
      
      '查potcustomer1 國內潛在客戶 檔
         strSql = strSql & " union all SELECT '' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,poc03 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,poc14 AS 狀態, poc15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM POTCUSTOMER1,NATION,STAFF, (Select Distinct poc01 As A1 From POTCUSTOMER1 Where instr(poc03,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE poc01=A.A1 AND poc04=na01(+) and poc13=ST01(+) "
         strSql = strSql & " union all SELECT '' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,rtrim(poc23||' '||poc24||' '||poc25||' '||poc26) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,poc14 AS 狀態, poc15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM POTCUSTOMER1,NATION,STAFF, (Select Distinct poc01 As A1 From POTCUSTOMER1 Where instr(upper(poc23||' '||poc24||' '||poc25||' '||poc26),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & " ) A WHERE poc01=A.A1 AND poc04=na01(+) and poc13=ST01(+) "
         strSql = strSql & " union all SELECT '' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,poc27 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,poc14 AS 狀態, poc15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM POTCUSTOMER1,NATION,STAFF, (Select Distinct poc01 As A1 From POTCUSTOMER1 Where instr(poc27,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A WHERE poc01=A.A1 AND poc04=na01(+) and poc13=ST01(+) "

      '查NotAgent 不得代理案件之客戶或代理人 檔
         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,nt02 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from notagent,nation,staff, (Select Distinct nt01 As A1 From notagent Where instr(nt02,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A where nt08=na01(+) and nt01=A.A1 AND nt18=ST01(+) "
         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT03||' '||NT04||' '||NT05||' '||NT06 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from notagent,nation,staff, (Select Distinct NT01 As A1 From notagent Where instr(upper(NT03||' '||NT04||' '||NT05||' '||NT06),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & " ) A where nt08=na01(+) and nt01=A.A1 AND nt18=ST01(+) "
         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT07 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from notagent,nation,staff, (Select Distinct NT01 As A1 From Notagent Where instr(NT07,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & " ) A where nt08=na01(+) and nt01=A.A1 AND nt18=ST01(+) "
      
      '查聯絡人(中文)
         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU01(+)=PCC01 AND CU02='0' AND CU13=ST01(+) "
         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,potcustomer,nation,STAFF where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' AND substr(LTrim(PCU38),1,5)=ST01(+) "
         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,''AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,POC14 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,potcustomer1,nation,STAFF where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' AND poc13=ST01(+) "

      '查聯絡人(英文)
         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU01(+)=PCC01 AND CU02='0' AND CU13=ST01(+) "
         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & ") A,potcustomer,nation,STAFF where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' AND substr(LTrim(PCU38),1,5)=ST01(+) "
         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,' ' as 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,POC14 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(txt1(1)))) & "')" & strCheckWay & ") A,potcustomer1,nation,STAFF where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' AND poc13=ST01(+) "

      '查聯絡人(日文)
         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+) AND CU01(+)=PCC01 AND CU02='0' AND CU13=ST01(+) "
         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,potcustomer,nation,STAFF where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' AND substr(LTrim(PCU38),1,5)=ST01(+) "
         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,' ' as 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' "
         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,POC14 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(txt1(1))) & "')" & strCheckWay & ") A,potcustomer1,nation,STAFF where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' AND poc13=ST01(+) "
       '抓暫存檔對造 Add by Amy 2014/02/25
        strSql = strSql & " union select ' ' as V,R021001 AS 編號,R021002 AS 名稱,'' AS 國籍,'' AS 智權人員,Decode(R021004,'1','對造','其他相關人') AS 狀態,'' AS 備註,'' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 From R100102_1 Where ID='" & strUserNum & "' And R021004<3 "
      'end 2015/03/27
      
      'Mark by Amy 2014/02/25 往上搬
'      '對造(中)
'            strSubSQL1 = " And InStr(Upper(CP40),'" & strTp(3) & "') " & strCheckWay
'            strSubSQL2 = " And InStr(Upper(CP50),'" & strTp(3) & "') " & strCheckWay
'            strSwhSQL1 = " CP40>' ' "
'            strSwhSQL2 = " CP50>' ' "
'            '商標
'            strSql = strSql & " Union " & _
'                        "Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP40 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
'                        " Union  Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP50 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
'            '專利
'            strSql = strSql & " Union " & _
'                        "Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP40 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
'                        " Union  Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP50 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
'            '法務
'            strSql = strSql & " Union " & _
'                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP40 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1 & _
'                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP50 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家 ,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
'            '顧問
'            strSql = strSql & " Union " & _
'                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP40 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1 & _
'                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP50 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
'            '服務
'            strSql = strSql & " Union " & _
'                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP40 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1 & _
'                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP50 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
'
'            '對造(英)
'            strSubSQL1 = " And InStr(Upper(CP41),'" & strTp(3) & "') " & strCheckWay
'            strSubSQL2 = " And InStr(Upper(CP51),'" & strTp(3) & "') " & strCheckWay
'            strSwhSQL1 = " CP41>' ' "
'            strSwhSQL2 = " CP51>' ' "
'            '商標
'            strSql = strSql & " Union " & _
'                         "Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP41 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
'                         " Union Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP51 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
'            '專利
'            strSql = strSql & " Union " & _
'                         "Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP41 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
'                         " Union Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP51 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
'            '法務
'            strSql = strSql & " Union " & _
'                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP41 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1 & _
'                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP51 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
'            '顧問
'            strSql = strSql & " Union " & _
'                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP41 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1 & _
'                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP51 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
'            '服務
'            strSql = strSql & " Union " & _
'                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP41 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1 & _
'                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP51 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
'
'            '對造(日)
'            strSubSQL1 = " And InStr(Upper(CP42),'" & strTp(3) & "') " & strCheckWay
'            strSubSQL2 = " And InStr(Upper(CP52),'" & strTp(3) & "') " & strCheckWay
'            strSwhSQL1 = " CP42>' ' "
'            strSwhSQL2 = " CP52>' ' "
'            '商標
'            strSql = strSql & " Union " & _
'                         "Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP42 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
'                         " Union Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP52 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
'            '專利
'            strSql = strSql & " Union " & _
'                         "Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP42 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
'                         " Union Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP52 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
'            '法務
'            strSql = strSql & " Union " & _
'                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP42 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1 & _
'                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP52 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
'            '顧問
'            strSql = strSql & " Union " & _
'                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP42 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1 & _
'                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP52 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
'            '服務
'            strSql = strSql & " Union " & _
'                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP42 as 名稱,' ' as 國籍,' ' as 智權人,'對造' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1 & _
'                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP52 as 名稱,' ' as 國籍,' ' as 智權人,'其他相關人' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
'                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
      'end 2014/02/25
'end 2013/12/04
      
      ' Add By Sindy 98/02/13 開拓客戶
      If Check1.Value = 1 Then
         'Modify by Amy 2013/12/04 +智權人員/申請國家/總收文號/案件性質/收文日
         'Modify by Amy 2013/09/30 原只檢查ecd11,ecd12卻顯示ecd03,ecd04
         'strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,NVL(ecd03,'')||NVL(ecd04,'') AS 名稱,NA03 AS 國籍,ecd15 AS 狀態,ecd16 AS 備註 from expandcusdetail,expandcusattr,nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') as A1 From expandcusdetail Where instr(ecd11,'" & ChgSQL(Trim(txt1(1))) & "') " & strCheckWay & " or instr(ecd12,'" & ChgSQL(Trim(txt1(1))) & "') " & strCheckWay & ") A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
         strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,NVL(ecd03,'')||NVL(ecd04,'') AS 名稱,NA03 AS 國籍,' ' AS 智權人員,'投法開拓'||decode(ecd15,null,null,'-'||ecd15) AS 狀態,ecd16 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from expandcusdetail,expandcusattr,nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') as A1 From expandcusdetail Where instr(Upper(ecd03),'" & ChgSQL(UCase(Trim(txt1(1)))) & "') " & strCheckWay & " or instr(Upper(ecd04),'" & ChgSQL(UCase(Trim(txt1(1)))) & "') " & strCheckWay & ") A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
         strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,NVL(ecd11,'')||NVL(ecd12,'') AS 名稱,NA03 AS 國籍,' ' AS 智權人員,'投法開拓'||decode(ecd15,null,null,'-'||ecd15) AS 狀態,ecd16 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from expandcusdetail,expandcusattr,nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') as A1 From expandcusdetail Where instr(Upper(ecd11),'" & ChgSQL(UCase(Trim(txt1(1)))) & "') " & strCheckWay & " or instr(Upper(ecd12),'" & ChgSQL(UCase(Trim(txt1(1)))) & "') " & strCheckWay & ") A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
      End If
      ' 98/02/13 End
      
      
   '以國籍查詢
   ElseIf Option1(2).Value = True Then
      'Modify by Amy 2013/12/04 +智權人員/申請國家/總收文號/案件性質/收文日
      'edit by nickc 2005/12/06
      'strSQL = "SELECT ''AS V,FA01||FA02||Decode(FA02,'0','','＊') AS 編號," & StrSQLa & "NA03 AS 國籍,FA69 AS 狀態, FA29 AS 備註 FROM FAGENT,NATION WHERE INSTR(FA10, '" & txt1(9) & "') = 1 AND fa10=NA01(+) "
      strSql = "SELECT ''AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號," & StrSQLa & "NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM FAGENT,NATION WHERE INSTR(FA10, '" & txt1(9) & "') = 1 AND fa10=NA01(+) "
      'Add by Morgan 2007/12/14
      strSql = strSql & " union all SELECT ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號," & strSQLc & "NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM POTCUSTOMER,NATION,Staff WHERE INSTR(pcu09, '" & txt1(9) & "') = 1 and PCU09=NA01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
      'end 2007/12/14
      
      'Add By Sindy 2011/10/11
      strSql = strSql & " union all SELECT ' ' AS V ,PoC01||PoC02||Decode(PoC02,'0','','＊') AS 編號," & strSQLD & "NA03 AS 國籍,ST02 AS 智權人員,PoC14 AS 狀態,PoC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM POTCUSTOMER1,NATION,Staff WHERE INSTR(poc04, '" & txt1(9) & "') = 1 and PoC04=NA01(+) and poc13=ST01(+) "
      
      'edit by nickc 2005/12/06
      'strSQL = strSQL & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊') AS 編號," & StrSqlB & "NA03 AS 國籍,CU80 AS 狀態,CU79 AS 備註 FROM customer,NATION WHERE INSTR(CU10, '" & txt1(9) & "') = 1 AND cu10=na01(+)"
      strSql = strSql & " union all SELECT '' AS V,cu01||cu02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號," & StrSqlB & "NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM customer,NATION,Staff WHERE INSTR(CU10, '" & txt1(9) & "') = 1 AND cu10=na01(+) AND CU13=ST01(+) "
      'Add By Sindy 2012/3/21
      strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號," & strSQLE & "NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from notagent,nation,Staff where nt08=na01(+) and INSTR(nt08, '" & txt1(9) & "') = 1 AND nt18=ST01(+) "
      pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & txt1(9) 'Add By Sindy 2010/11/4
      
   'E-Mail  add by Toni 2008/12/03
   ElseIf Option1(3).Value = True Then
        strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●','') AS 編號,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 FROM CUSTOMER,NATION,Staff  Where (instr(NLS_Upper(CU20),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0 or  instr(NLS_Upper(CU115),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0 or instr(NLS_Upper(CU116),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0  or instr(NLS_Upper(CU117),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')>0 or instr(NLS_Upper(CU118),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 )  and CU10=NA01(+) AND CU13=ST01(+) "
                                                                                                                                                               
        strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from potcustomer,nation,Staff  Where (instr(NLS_Upper(pcu18),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') >0 ) and pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
        'Add By Sindy 98/03/20
        strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號," & strSQLD & "NA03 AS 國籍,ST02 AS 智權人員,POC14 AS 狀態,POC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from potcustomer1,nation,Staff  Where (instr(NLS_Upper(poc09),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') >0 ) and poc04=na01(+) and poc13=ST01(+) "
        '98/03/20 End
        strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from fagent,nation Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(txt1(10)))) & "')> 0 or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 )   and  fa10=na01(+)   "
   
        strSql = strSql & " union all SELECT ' ' AS V,PCC01||'0-'||PCC02 AS 編號,Decode(PCC05,'',PCC03,'',PCC04,PCC05) AS 名稱,' ' AS 國籍,' ' as 智權人員,' ' AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from PotCustCont Where (instr(NLS_Upper(PCC08),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 )  "
        
        ' Add By Sindy 98/02/13
        If Check1.Value = 1 Then
         'Modify by Amy 2013/12/04 +智權人員/申請國家/總收文號/案件性質/收文日
         'Modify by Amy 2013/09/30 原:ecd15 AS 狀態
         strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,NVL(ecd03,'')||NVL(ecd04,'') AS 名稱,NA03 AS 國籍,' ' as 智權人員,'投法開拓'||decode(ecd15,null,null,'-'||ecd15) AS 狀態,ecd16 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日 from expandcusdetail,expandcusattr,nation Where (instr(NLS_Upper(ecd13),'" & UCase(ChgSQL(Trim(txt1(10)))) & "') > 0 ) and ecd10=na01(+) and ecd02=eca01(+) "
        End If
        ' 98/02/13 End
        pub_QL05 = pub_QL05 & ";" & Option1(3).Caption & Trim(txt1(10)) 'Add By Sindy 2010/11/4
   End If
   
   If Check1.Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Check1.Caption 'Add By Sindy 2010/11/4
   End If
   
   CheckOC
   'modify by nickc 2005/06/03
   'strSQL = strSQL & " order by 名稱 "
   '2008/12/3 modify by sonia
   'strSQL = "select * from (" & strSQL & ") X order by upper(名稱) "
   If Option1(1).Value = True Then
      'Modify by Amy 2014/01/15 +編號排
      strSql = "select * from (" & strSql & ") X order by upper(名稱),編號 "
   Else
      strSql = "select * from (" & strSql & ") X order by 編號 "
   End If
   '2008/12/3 end
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/4
       cmdOK(0).Enabled = True
       cmdOK(1).Enabled = True
       cmdOK(2).Enabled = True
       'add by nickc 2005/10/24
       cmdOK(5).Enabled = True
       '911029 nick move from down
       Set grdDataList.Recordset = adoRecordset
   Else
       InsertQueryLog (0) 'Add By Sindy 2010/11/4
 'Modify by Amy 2013/12/04 Mark Option1(1).Value = True And Trim(txt1(1)) <> "" Then 掉不需再找對造
'       'Add By Sindy 2010/02/05
'       If Option1(1).Value = True And Trim(txt1(1)) <> "" Then
'          MsgBox "非本所客戶或代理人，系統會再搜尋案件對造資料，請注意是否有雙方代理情形！", vbInformation
'          Me.Enabled = False
'          If fnSaveParentForm(Me) = False Then
'             Me.Enabled = True
'             Exit Sub
'          End If
'          Screen.MousePointer = vbHourglass
'          frm100110_1.Option1(1).Value = True
'          frm100110_1.txt1(1) = Trim(txt1(1))
'          frm100110_3.StrMenu
'          Unload frm100110_1
'          Screen.MousePointer = vbDefault
'          Me.Enabled = True
'          Exit Sub
'       '2010/02/05 End
'       Else
          'Modify by Amy 2013/12/04 +畫面訊息開放可列印
          Pub_Can_Copy_Pic = True
          ShowNoData
          Pub_Can_Copy_Pic = False
          'end 2013/12/04
          cmdOK(0).Enabled = False
          cmdOK(1).Enabled = False
          cmdOK(2).Enabled = False
          'add by nickc 2005/10/24
          cmdOK(5).Enabled = False
          Screen.MousePointer = vbDefault
          Exit Sub
'       End If
   End If
   Me.grdDataList.Visible = False 'Add by Amy 2013/12/04
   CheckOC
   '911029 nick move to up
   'Set GrdDataList.Recordset = adoRecordset
   SetDataListWidth
   
   'add by nickc 2005/12/14 變色
   With Me.grdDataList
        If .Rows > 0 Then 'Add by Amy 2013/12/04
            For i = 0 To .Rows - 1
                .row = i
                .col = 1
                If Right(.Text, 1) = "$" Then
                    .CellBackColor = &HFF&
                    'Add By Sindy 2012/3/21
                ElseIf Right(.Text, 1) = "⊕" Or .TextMatrix(i, 5) = "對造" Or .TextMatrix(i, 5) = "其他相關人" Then
                    'Modify by Amy 2013/12/04 對造重抓智權人資料
                    If Me.grdDataList.TextMatrix(i, 5) = "對造" Or .TextMatrix(i, 5) = "其他相關人" Then
                        bolPrint = True '有對造資料
                        strNo = Pub_RplStr(.TextMatrix(i, 1))
                        StrToPrint = strNo & ","
                        Str01 = SystemNumber(strNo, 1)
                        Select Case Str01
                            Case "FCP", "FG"
                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCPSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
                            Case "FCL", "LIN"
                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCLSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
                            Case "FCT"
                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
                            Case "S"
                                If .TextMatrix(i, 7) = "000" Then
                                    .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
                                Else
                                    .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
                                End If
                            Case Else
                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
                        End Select
                        .TextMatrix(i, 9) = .TextMatrix(i, 9) & PUB_GetRelateCasePropertyName(.TextMatrix(i, 8), "1")
                        'Add by Amy 2014/02/25 更新智權人員至暫存檔
                        strExc(0) = "Update R100102_1 Set R021003='" & .TextMatrix(i, 4) & "' Where R021014='" & Str01 & "' And R021015='" & SystemNumber(strNo, 2) & "' And R021016='" & SystemNumber(strNo, 3) & "' And R021017='" & SystemNumber(strNo, 4) & "' "
                        cnnConnection.Execute strExc(0)
                        'end 2014/02/25
                    End If
                    'end 2013/12/04
                    If Right(.Text, 1) = "⊕" Or .TextMatrix(i, 5) = "對造" Then
                        For j = 0 To .Cols - 1
                            .col = j
                            .CellBackColor = &H8080FF
                        Next j
                    End If
                    '2012/3/21 End
                End If
            Next i
        End If 'end 2013/12/04
   End With
   
   'Add By Cheng 2001/12/26
   '若查詢結果僅有一筆資料, 則直接勾選
   If Me.grdDataList.Rows = 2 Then
      '911029 nick add
      grdDataList.col = 1
      grdDataList.row = 1
      If grdDataList.Text <> "" Then
      '911029 nick end
           grdDataList.Visible = False
           grdDataList.row = 1
           grdDataList.col = 0
           grdDataList.Text = "V"
           For i = 0 To grdDataList.Cols - 1
               'add by nickc 2005/12/14
               'Modify By Sindy 2012/3/21 old:If i <> 1 Then
               If i <> 1 And (i = 2 And Right(grdDataList.TextMatrix(1, 1), 1) = "⊕") = False Then
                   grdDataList.col = i
                   grdDataList.CellBackColor = &HFFC0C0
               End If
           Next i
           grdDataList.Visible = True
       '911029 nick add
       End If
       '911029 nick end
   End If
   'Add by Amy 2013/12/04
   Me.grdDataList.Visible = True
   If bolPrint Then
        cmdOK(8).Enabled = True
   Else
        cmdOK(8).Enabled = False
   End If
   'end 2013/12/04
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   If bolFNation = False Then
       s = MsgBox("國內人員不可查詢代理人案件", , "違規.....")
       Unload Me
       Exit Sub
   End If
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   '2011/12/6 modify by sonia
   'txt1(2) = Systemkind_g
   Me.chk.Value = vbChecked
   txt1(2) = "ALL"
   '2011/12/6 end
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txt1(1).IMEMode = 1
   
   Option1(1).Value = False
   Option2(0).Enabled = False
   Option2(1).Enabled = False
   Option2(2).Enabled = False
   'txt1(1).Enabled = False
   Option1(2).Value = False
   'txt1(9).Enabled = False
   cmdOK(0).Enabled = False
   cmdOK(1).Enabled = False
   cmdOK(2).Enabled = False
   'add by nickc 2005/10/24
   cmdOK(5).Enabled = False
   '92.04.16 nick
   CmdState = -1
   
   ' Add By Sindy 98/02/16
   'MODIFY BY SONIA 2015/5/20 因P31及F31人員併入L02,但內外法不開放權限,故改用員工等級控制
   'If Pub_StrUserSt03 = "F31" Or Pub_StrUserSt03 = "F41" Then
   If Pub_strUserST05 >= "51" And Pub_strUserST05 <= "55" Then
      Check1.Value = 1
   Else
      Check1.Value = 0
   End If
   ' 98/02/16 End
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100114_1 = Nothing
End Sub

'Add by Amy 2014/04/25 原寫於grdDataList_SelChange
Private Sub GrdDataList_Click()
   Dim strCopyTxt As String ' Add by Amy 2014/04/25 複製編號文字
    
   grdDataList.row = grdDataList.MouseRow
    
   'Modify by Amy 2014/04/25 +選到編號欄=複製
   grdDataList.col = grdDataList.MouseCol
   If grdDataList.col = 1 Then
        strCopyTxt = grdDataList.TextMatrix(grdDataList.row, grdDataList.col)
        If strCopyTxt <> "" Then
            '複製編號至剪貼簿
            Clipboard.SetText strCopyTxt
            grdDataList.CellBackColor = QBColor(7)
            MsgBox "編號已複製", , MsgText(21)
        
            '設回原本顏色
            If Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, 5) = "對造" Then
                grdDataList.CellBackColor = &H8080FF
            Else
                grdDataList.CellBackColor = QBColor(15)
            End If
        End If
        Exit Sub
   End If
   'end 2014/04/25
   
   grdDataList.Visible = False
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
        If grdDataList.Text = "V" Then
            grdDataList.Text = ""
            'Add By Sindy 2012/3/21
            grdDataList.col = 1
            If Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, 5) = "對造" Then
                For i = 0 To grdDataList.Cols - 1
                    grdDataList.col = i
                    grdDataList.CellBackColor = &H8080FF
                Next i
            Else
            '2012/3/21 End
                For i = 0 To grdDataList.Cols - 1
                     'add by nickc 2005/12/14
                    If i <> 1 Then
                      grdDataList.col = i
                      grdDataList.CellBackColor = QBColor(15)
                    End If
                Next i
            End If
        Else
            grdDataList.Text = "V"
            For i = 0 To grdDataList.Cols - 1
                'add by nickc 2005/12/14
                'Modify By Sindy 2012/3/21 old:If i <> 1 Then
                If i <> 1 And (i = 2 And Right(grdDataList.TextMatrix(grdDataList.MouseRow, 1), 1) = "⊕") = False Then
                   grdDataList.col = i
                    grdDataList.CellBackColor = &HFFC0C0
                End If
            Next i
        End If
   End If
   grdDataList.Visible = True
End Sub

Private Sub Option1_Click(Index As Integer)
      
   Select Case Index
      Case 0
           If Option1(0).Value = True Then
              'txt1(0).Enabled = True
              Option1(1).Value = False
              Option1(2).Value = False
              Option2(0).Enabled = False
              Option2(1).Enabled = False
              Option2(2).Enabled = False
              Option3(0).Enabled = False
              Option3(1).Enabled = False
              txt1(0).SetFocus
              txt1_GotFocus (0)
              'txt1(1).Enabled = False
              'txt1(9).Enabled = False
           End If
      Case 1
           If Option1(1).Value = True Then
              Option2(0).Enabled = True
              Option2(1).Enabled = True
              Option2(2).Enabled = True
              'txt1(1).Enabled = True
              
              Option1(0).Value = False
              Option1(2).Value = False
              'txt1(0).Enabled = False
              'txt1(9).Enabled = False
              'txt1(1).SetFocus
              'txt1_GotFocus (1)
              Option3(0).Enabled = True
              Option3(1).Enabled = True
              Option3(1).Value = True    '2012/3/28 ADD BY SONIA
              txt1(1).SetFocus
           End If
      Case 2
           If Option1(2).Value = True Then
              'txt1(9).Enabled = True
              Option1(0).Value = False
              Option1(1).Value = False
              'txt1(0).Enabled = False
              Option2(0).Enabled = False
              Option2(1).Enabled = False
              Option2(2).Enabled = False
              Option3(0).Enabled = False
              Option3(1).Enabled = False
              'txt1(1).Enabled = False
              txt1(9).SetFocus
              txt1_GotFocus (9)
            End If
      'E-MAIL ADD BY Toni 2008/12/03
      Case 3
        If Option1(3).Value = True Then
           Option1(0).Value = False
           Option1(1).Value = False
           Option1(2).Value = False
           
           Option2(0).Enabled = False
           Option2(1).Enabled = False
           Option2(2).Enabled = False
           Option3(0).Enabled = False
           Option3(1).Enabled = False
           
           txt1(10).SetFocus
           txt1_GotFocus (10)
         End If
      Case Else
   End Select
End Sub

Private Sub Option2_Click(Index As Integer)

   Select Case Index
      Case 0
            If Option2(0).Value = True Then
               Option2(1).Value = False
               Option2(2).Value = False
               'Option3(1).Value = True   '2012/3/28 CANCEL BY SONIA 改在Option1(1)
               'edit by nickc 2007/06/06 切換輸入法改用API
               'txt1(1).IMEMode = 1
               OpenIme
            End If
      Case 1
            If Option2(1).Value = True Then
               Option2(0).Value = False
               Option2(2).Value = False
               'Option3(1).Value = True   '2012/3/28 CANCEL BY SONIA 改在Option1(1)
               'edit by nickc 2007/06/06 切換輸入法改用API
               'txt1(1).IMEMode = 0
               CloseIme
            End If
      Case 2
            If Option2(2).Value = True Then
               Option2(0).Value = False
               Option2(1).Value = False
               'Option3(1).Value = True   '2012/3/28 CANCEL BY SONIA 改在Option1(1)
               'edit by nickc 2007/06/06 切換輸入法改用API
               'txt1(1).IMEMode = 1
               OpenIme
            End If
      Case Else
   End Select
   txt1(1).SetFocus
   txt1_GotFocus (1)
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   If Index = 1 Then
      'If Option2(1).Value = True Then  'Modify by Amy 2013/12/10 改判斷部門
      If Left(Pub_StrUserSt03, 1) = "F" Then
         'edit by nickc 2007/06/06 切換輸入法改用API
         'txt1(1).IMEMode = 2
         CloseIme
      Else
         'edit by nickc 2007/06/06 切換輸入法改用API
         'txt1(1).IMEMode = 1
         OpenIme
      End If
   'add by sonia 2014/10/29
   Else
      CloseIme
   'end 2014/10/29
   End If
   
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   '2008/12/3 modify by sonia
   'If Index <> 1 Then
   If Index <> 1 And Index <> 10 Then
   '2008/12/3 end
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txt1_LostFocus(Index As Integer)

   Select Case Index
      Case 0
      Case 1
      Case 2 '系統類別
            'Modify By Cheng 2002/03/14
      '      'Add By Cheng 2002/01/07
      '      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
      Case 3
            If InStr(1, "12", txt1(3)) = 0 Then
                s = MsgBox("查詢別只可 1 或 2 !!", , "USER 輸入錯誤")
                txt1(3).SetFocus
                txt1(3).SelStart = 0
                txt1(3).SelLength = Len(txt1(3))
                Exit Sub
            End If
      Case 4, 5
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
         If Index = 5 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
                txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
            End If
          End If
      Case 6
      Case 7
            If RunNick(txt1(Index - 1), txt1(Index)) Then
                txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
            End If
      Case 8
            If Len(txt1(8)) <> 0 Then
                strSql = "SELECT NA03 FROM NATION WHERE NA01='" & txt1(8) & "'"
                CheckOC
                adoRecordset.CursorLocation = adUseClient
                adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                      If Not IsNull(adoRecordset.Fields(0)) Then
                          lbl1(0).Caption = adoRecordset.Fields(0)
                      Else
                          lbl1(0).Caption = ""
                      End If
                Else
                    lbl1(0).Caption = ""
                    s = MsgBox("國家輸入錯誤！", , "錯誤！")
                    txt1(Index).SetFocus
                    txt1_GotFocus (Index)
                    Exit Sub
                End If
                CheckOC
            Else
                lbl1(0).Caption = ""
            End If
      Case 9
            If Len(txt1(9)) <> 0 Then
                strSql = "SELECT NA03 FROM NATION WHERE NA01='" & txt1(9) & "'"
                CheckOC
                adoRecordset.CursorLocation = adUseClient
                adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                      If Not IsNull(adoRecordset.Fields(0)) Then
                          lbl1(1).Caption = adoRecordset.Fields(0)
                      Else
                          lbl1(1).Caption = ""
                      End If
                Else
                    lbl1(1).Caption = ""
                    s = MsgBox("國家輸入錯誤！", , "錯誤！")
                    txt1(Index).SetFocus
                    txt1_GotFocus (Index)
                    Exit Sub
                End If
                CheckOC
            Else
                lbl1(1).Caption = ""
            End If
      Case Else
   End Select
End Sub

Private Sub txt1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   Select Case Index
      Case 0
          Option1(0).Value = True
      Case 1
          Option1(1).Value = True
      Case 9
          Option1(2).Value = True
      Case 10
          Option1(3).Value = True
      Case Else
   End Select
End Sub

'Add by Amy 2014/02/25
Private Sub PrintDataA4_Temp()
    Dim rsPrint As New ADODB.Recordset
    Dim strPrint As String
    Dim ii As Integer, jj As Integer
On Error GoTo Checking
    intCounter = 1: intRecord = 1: intPage = 1
    
    Screen.MousePointer = vbHourglass
    Printer.PaperSize = PUB_GetPaperSize(9) '設定紙張 A4
    Printer.Orientation = 1 '直印
    PrintHeadA4
    
    Printer.FontBold = False
    strPrint = "Select R021001,R021002,R021003,Decode(R021004,'1','對造','其他相關人'),R021006,R021007,Nvl(To_Char(R021008-19110000),'') " & _
                 "From R100102_1 Where ID='" & strUserNum & "' Order by R021002,R021001"
    intI = 1
    Set rsPrint = ClsLawReadRstMsg(intI, strPrint)
    If intI = 1 Then
        rsPrint.MoveFirst
        For ii = 0 To rsPrint.RecordCount - 1
            If intRecord > 45 Then
                intPage = intPage + 1
                intRecord = 1
                Printer.NewPage
                intCounter = 1
                PrintHeadA4
                Printer.FontBold = False
            End If
            For jj = 0 To rsPrint.Fields.Count - 1
                If jj = rsPrint.Fields.Count - 1 Then
                    Printer.CurrentX = PLeft(jj + 1) - 300 - Printer.TextWidth(rsPrint.Fields(jj).Value) '最右邊
                Else
                    Printer.CurrentX = PLeft(jj)
                End If
                Printer.CurrentY = 300 + intCounter * 300

                Select Case jj
                    Case 0 '本所案號
                        Printer.Print Pub_RplStr(rsPrint.Fields(jj).Value)
                    Case 1 '名稱
                        Printer.Print StrToStr(rsPrint.Fields(jj).Value, 10)
                    Case 2, 3, 4 '智權人員,狀態,總收文號
                        Printer.Print rsPrint.Fields(jj).Value
                    Case 5 '案件性質
                        Printer.Print StrToStr(rsPrint.Fields(jj).Value, 6)
                    Case 6  '收文日
                        Printer.Print ChangeTStringToTDateString(rsPrint.Fields(jj).Value)
                    Case Else
                End Select
            Next jj
            intCounter = intCounter + 1
            intRecord = intRecord + 1
            rsPrint.MoveNext
        Next ii
    End If
    Printer.EndDoc
    Screen.MousePointer = vbDefault
    
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   Screen.MousePointer = vbDefault
End Sub
'end 2014/02/25

'2014/02/25未使用 Add by Amy 2013/12/04
Private Sub PrintDataA4()
    Dim ii As Integer, jj As Integer
On Error GoTo Checking
    intCounter = 1: intRecord = 1: intPage = 1
    
    Screen.MousePointer = vbHourglass
    Printer.PaperSize = PUB_GetPaperSize(9) '設定紙張 A4
    Printer.Orientation = 1 '直印
    PrintHeadA4
           
    Printer.FontBold = False
    With Me.grdDataList
        For ii = 1 To .Rows - 1
            If intRecord > 45 Then
                intPage = intPage + 1
                intRecord = 1
                Printer.NewPage
                intCounter = 1
                PrintHeadA4
                Printer.FontBold = False
            End If
            If Left(.TextMatrix(ii, 1), 1) <> "X" And Left(.TextMatrix(ii, 1), 1) <> "Y" And Left(.TextMatrix(ii, 1), 1) <> "R" Then
                For jj = 1 To .Cols - 1
                    If jj <= 2 Or jj = 4 Or jj = 5 Or (jj >= 8 And jj <= 10) Then
                        Select Case jj
                            Case 1 '本所案號
                                Printer.CurrentX = PLeft(jj - 1)
                                Printer.CurrentY = 300 + intCounter * 300
                                Printer.Print Pub_RplStr(.TextMatrix(ii, jj))
                            Case 2 '名稱
                                Printer.CurrentX = PLeft(jj - 1)
                                Printer.CurrentY = 300 + intCounter * 300
                                Printer.Print Left(.TextMatrix(ii, jj), 10)
                            Case 4, 5 '智權人員,狀態
                                Printer.CurrentX = PLeft(jj - 2)
                                Printer.CurrentY = 300 + intCounter * 300
                                Printer.Print .TextMatrix(ii, jj)
                            Case 8 '總收文號
                                Printer.CurrentX = PLeft(jj - 4)
                                Printer.CurrentY = 300 + intCounter * 300
                                Printer.Print .TextMatrix(ii, jj)
                            Case 9 '案件性質
                                Printer.CurrentX = PLeft(jj - 4)
                                Printer.CurrentY = 300 + intCounter * 300
                                Printer.Print Left(.TextMatrix(ii, jj), 6)
                            Case 10  '收文日
                                Printer.CurrentX = PLeft(jj - 3) - 300 - Printer.TextWidth(.TextMatrix(ii, jj))
                                Printer.CurrentY = 300 + intCounter * 300
                                Printer.Print ChangeTStringToTDateString(.TextMatrix(ii, jj))
                            Case Else
                        End Select
                    End If
                Next jj
                intCounter = intCounter + 1
                intRecord = intRecord + 1
            End If
        Next ii
    End With
    Printer.EndDoc
    Screen.MousePointer = vbDefault
    
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   Screen.MousePointer = vbDefault
End Sub

Private Sub PrintHeadA4()
  
   If intPage = 1 Then
        GetPleft
        strTp(0) = "以申請人查詢"
        strTp(1) = ""
       
        If Option3(0).Value = True Then
            strTp(1) = strTp(1) & "(字首比對)"
        ElseIf Option3(1).Value = True Then
            strTp(1) = strTp(1) & "(模糊比對)"
        End If
   End If
   strTp(2) = "名稱：" & strTp(3) & Space(6) & strTp(1)
   
   Printer.FontSize = 17
   Printer.FontBold = True
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTp(0)) / 2)
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print strTp(0)
   
   Printer.FontSize = 12
   intCounter = intCounter + 2
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTp(2)) / 2)
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print strTp(2)
   'Printer.Line (Printer.ScaleWidth / 2 - ((Printer.TextWidth(strTp(2)) - Printer.TextWidth("名稱：")) / 2) + 300, Printer.CurrentY + 30)-(Printer.ScaleWidth / 2 + Printer.TextWidth(strTp(2)) / 2, Printer.CurrentY + 30)
     
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "操作人員：" & StaffQuery(strUserNum)
   Printer.CurrentX = 8800
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "查詢日期：" & CFDate(ACDate(ServerDate))
'   intCounter = intCounter + 1
'   Printer.CurrentX = 12000
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "頁次: " & intPage
    intCounter = intCounter + 1
    For kk = 1 To UBound(PLeft)
        Printer.CurrentX = PLeft(kk - 1) + (PLeft(kk) - PLeft(kk - 1) - Printer.TextWidth(ColName(kk)) - 100) / 2
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print ColName(kk)
        Printer.Line (PLeft(kk - 1), Printer.CurrentY)-(PLeft(kk) - 100, Printer.CurrentY)
    Next kk
    
    intCounter = intCounter + 1
End Sub

Private Sub GetPleft()
   ReDim PLeft(0 To 7)
   ReDim ColName(1 To 7)
   PLeft(0) = 100
   PLeft(1) = PLeft(0) + 2000: ColName(1) = "本所案號"
   PLeft(2) = PLeft(1) + 2700: ColName(2) = "    名       稱    "
   PLeft(3) = PLeft(2) + 1200: ColName(3) = "智權人員"
   PLeft(4) = PLeft(3) + 1500: ColName(4) = " 狀  態 "
   PLeft(5) = PLeft(4) + 1300: ColName(5) = "總收文號"
   PLeft(6) = PLeft(5) + 1800: ColName(6) = "案件性質"
   PLeft(7) = PLeft(6) + 1200: ColName(7) = "收文日"
End Sub

'end 2013/12/04
