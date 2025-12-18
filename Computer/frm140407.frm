VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140407 
   BorderStyle     =   1  '單線固定
   Caption         =   "潛在客戶資料查詢"
   ClientHeight    =   5900
   ClientLeft      =   3780
   ClientTop       =   3700
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5900
   ScaleWidth      =   9310
   Begin VB.CheckBox Check3 
      Caption         =   "顯示有無案件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3720
      TabIndex        =   43
      Top             =   60
      Width           =   1665
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   3180
      Left            =   30
      TabIndex        =   23
      Top             =   2670
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   5609
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   18
   End
   Begin VB.ComboBox cboSort 
      Height          =   300
      ItemData        =   "frm140407.frx":0000
      Left            =   1080
      List            =   "frm140407.frx":0002
      TabIndex        =   13
      Text            =   "cboSort"
      Top             =   2340
      Width           =   5550
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印對造資料(&O)"
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   7770
      Style           =   1  '圖片外觀
      TabIndex        =   38
      Top             =   700
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      MaxLength       =   9
      TabIndex        =   1
      Top             =   30
      Width           =   1965
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含投資法務開拓資料"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4560
      TabIndex        =   34
      Top             =   1050
      Width           =   2000
   End
   Begin VB.TextBox Text10 
      Height          =   264
      Left            =   90
      MaxLength       =   200
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2670
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "潛在客戶資料(&A)"
      Height          =   345
      Index           =   2
      Left            =   6630
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   15
      Width           =   1450
   End
   Begin VB.OptionButton Option2 
      Caption         =   "開發日期："
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   8
      Top             =   1770
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.OptionButton Option2 
      Caption         =   "開發者："
      Height          =   180
      Index           =   3
      Left            =   60
      TabIndex        =   6
      Top             =   1410
      Width           =   1200
   End
   Begin VB.ListBox lstSort 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      ItemData        =   "frm140407.frx":0004
      Left            =   1095
      List            =   "frm140407.frx":000B
      MultiSelect     =   1  '簡易多重選取
      Sorted          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2490
      Visible         =   0   'False
      Width           =   5550
   End
   Begin VB.CommandButton cmdAddSort 
      Caption         =   "新增↑"
      Height          =   285
      Left            =   6690
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdRemSort 
      Caption         =   "移除↓"
      Height          =   285
      Left            =   6690
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "聯絡人(&T)"
      Height          =   345
      Index           =   0
      Left            =   5640
      Style           =   1  '圖片外觀
      TabIndex        =   20
      Top             =   360
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "往來記錄(&N)"
      Height          =   345
      Index           =   4
      Left            =   6630
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   345
      Width           =   1450
   End
   Begin VB.TextBox Text9 
      Height          =   300
      Left            =   1305
      TabIndex        =   5
      Top             =   1050
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "E-Mail："
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   1110
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      Height          =   350
      Left            =   5310
      TabIndex        =   28
      Top             =   630
      Width           =   2400
      Begin VB.OptionButton Option3 
         Caption         =   "模糊比對"
         Height          =   180
         Index           =   1
         Left            =   1260
         TabIndex        =   30
         Top             =   144
         Value           =   -1  'True
         Width           =   1100
      End
      Begin VB.OptionButton Option3 
         Caption         =   "字首比對"
         Height          =   180
         Index           =   0
         Left            =   72
         TabIndex        =   29
         Top             =   144
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Height          =   350
      Left            =   4470
      TabIndex        =   24
      Top             =   1305
      Visible         =   0   'False
      Width           =   2436
      Begin VB.OptionButton Option1 
         Caption         =   "日文"
         Height          =   180
         Index           =   2
         Left            =   1656
         TabIndex        =   27
         Top             =   135
         Width           =   732
      End
      Begin VB.OptionButton Option1 
         Caption         =   "英文"
         Height          =   180
         Index           =   1
         Left            =   900
         TabIndex        =   26
         Top             =   135
         Width           =   732
      End
      Begin VB.OptionButton Option1 
         Caption         =   "中文"
         Height          =   180
         Index           =   0
         Left            =   72
         TabIndex        =   25
         Top             =   135
         Value           =   -1  'True
         Width           =   732
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "潛在客戶/聯絡人名稱："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   750
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "潛在客戶編號："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   1095
      MaxLength       =   7
      TabIndex        =   11
      Top             =   2010
      Width           =   852
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   1305
      MaxLength       =   7
      TabIndex        =   9
      Top             =   1695
      Width           =   852
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1305
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1350
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Height          =   300
      Left            =   2295
      MaxLength       =   7
      TabIndex        =   12
      Top             =   2010
      Width           =   852
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   2505
      MaxLength       =   7
      TabIndex        =   10
      Top             =   1695
      Width           =   852
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   345
      Left            =   5640
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   15
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "關係企業(&R)"
      Height          =   345
      Index           =   1
      Left            =   8085
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   3
      Left            =   8085
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   345
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2(4)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   6936
      TabIndex        =   44
      Top             =   2448
      Width           =   672
   End
   Begin MSForms.TextBox Text2 
      Height          =   330
      Left            =   2250
      TabIndex        =   3
      Top             =   675
      Width           =   3015
      VariousPropertyBits=   671105051
      BackColor       =   16777215
      Size            =   "5318;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label10 
      Height          =   300
      Left            =   2190
      TabIndex        =   42
      Top             =   1380
      Width           =   1290
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2275;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "＊：舊的名稱　＄：有呆帳　     ●：特殊客戶　 ⊕：不得代理　 ▼：無案件"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   912
      Left            =   7980
      TabIndex        =   41
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "        黃底為待活化客戶"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   3
      Left            =   6936
      TabIndex        =   40
      Top             =   2256
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "註：紅色不可承接案件"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   6936
      TabIndex        =   35
      Top             =   2076
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "輸入名稱之特取部分, 不要取國家,省份,城市,例：不可輸美商..,廣東..,廣州.."
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   0
      TabIndex        =   39
      Top             =   360
      Width           =   5844
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "往來日期：                                                    (民國)"
      Height          =   180
      Left            =   90
      TabIndex        =   37
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "往來類別："
      Height          =   180
      Left            =   90
      TabIndex        =   36
      Top             =   2370
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(民國)"
      Height          =   180
      Index           =   1
      Left            =   3540
      TabIndex        =   32
      Top             =   1725
      Width           =   480
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2070
      X2              =   2190
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "模糊比對"
      Height          =   180
      Left            =   3300
      TabIndex        =   31
      Top             =   1080
      Width           =   720
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2295
      X2              =   2415
      Y1              =   1845
      Y2              =   1845
   End
End
Attribute VB_Name = "frm140407"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/12 改成Form2.0 ; GrdDataList改字型=新細明體-ExtB、Label2、Text2
'Memo by Amy 2013/12/04 +查對造功能 拿掉中、英、日查詢選項
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'2008/11/06 add by Toni
Option Explicit

Dim i As Long, j As Long
Dim StrTag As String, StrToGrid As String
Dim strSql As String, lngCounter As Long, lngCounterI As Long
Public cmdState As Integer
Dim m_dbl_LeftMargin As Double
Dim m_dbl_TopMargin  As Double
Dim SeekPrintL As Integer
Dim SeekPrint As Integer
Dim m_bolPrintRight As Boolean
Private Const CB_SHOWDROPDOWN = &H14F
'Add by Amy 2013/12/04
Dim StrToPrint As String '記錄編號 for 對造列印
Dim strTp(3) As String, ColName() As String
Dim intCounter As Integer, intRecord As Integer, intPage As Integer, kk As Integer, PLeft() As Integer
Dim bolPrint As Boolean '是否有對造
'end 2013/12/04
Dim m_blnColOrderAsc As Boolean 'Add by Amy 2020/09/04 欄位資料由小到大排序
Dim strField() As String 'Add by Amy 2023/03/08
Dim m_pub_QL05 As String 'Add By Sindy 2025/8/13 只記錄於此Form


'Modify by Amy 2023/08/30 +IsRelation
Private Sub SetDataListWidth(Optional ByVal IsRelation As Boolean = False)
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "編號"
grdDataList.ColWidth(1) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "名稱"
grdDataList.ColWidth(2) = 4000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "國籍"
grdDataList.ColWidth(3) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
'Modify by Amy 2013/12/04 插入智權人員欄
grdDataList.col = 4: grdDataList.Text = "智權人員"
grdDataList.ColWidth(4) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
   
grdDataList.col = 5: grdDataList.Text = "類別"
grdDataList.ColWidth(5) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "狀態"
grdDataList.ColWidth(6) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "備註"
grdDataList.ColWidth(7) = 2000
grdDataList.CellAlignment = flexAlignLeftCenter
'Add by Amy 2013/12/04
'因查詢服務對造資料需依sp09抓不智權人員資料,故加申請國家
grdDataList.col = 8: grdDataList.Text = "申請國家"
grdDataList.ColWidth(8) = 0
'抓取對造欄位 for 列印
grdDataList.col = 9: grdDataList.Text = "總收文號"
grdDataList.ColWidth(9) = 0
grdDataList.col = 10: grdDataList.Text = "案件性質"
grdDataList.ColWidth(10) = 0
grdDataList.col = 11: grdDataList.Text = "收文日"
grdDataList.ColWidth(11) = 0
'end 2013/12/04

'Added by Lydia 2017/02/14 關聯企業
'Modify by Amy 2019/09/17 改為日期判斷 原:欄位數
If strSrvDate(1) < 國外部關聯企業啟用日 Then 'Added by Lydia 2017/12/28
    grdDataList.col = 12: grdDataList.Text = "關聯編號"
    grdDataList.ColWidth(12) = 0
    grdDataList.col = 13: grdDataList.Text = "關聯名稱"
    grdDataList.ColWidth(13) = 0
    grdDataList.col = 14: grdDataList.Text = "關聯關係"
    grdDataList.ColWidth(14) = 0
    grdDataList.col = 15: grdDataList.Text = "關聯說明"
    grdDataList.ColWidth(15) = 0
    grdDataList.FixedCols = 0
End If 'Added by Lydia 2017/12/28
'end 2017/02/14
 'Modify by Amy 2022/08/19 +ORGN
 grdDataList.col = 16: grdDataList.Text = "ORGN"
 grdDataList.ColWidth(16) = 0
 grdDataList.FixedCols = 0
 'Add by Amy 2019/09/17 +待活化客戶
 grdDataList.col = 17: grdDataList.Text = "待活化客戶"
 grdDataList.ColWidth(17) = 0
 grdDataList.FixedCols = 0
 'end 2022/08/19
 
'Modify by Amy 2023/08/30 避免沒改到,從strMenu1搬過來
'關聯企業
If IsRelation = True Then
   'Modified by Lydia 2018/05/24 改由啟用日控制
   If strSrvDate(1) >= 國外部關聯企業啟用日 Then
      'Added by Lydia 2017/02/14 欄寬調整
      grdDataList.FixedCols = 3 '固定編號和名稱
      Call PUB_SetMSFGridColor(Me.grdDataList, "15") '底色設定為空白
      grdDataList.ColWidth(2) = 1200 '名稱
      grdDataList.ColWidth(3) = 800 '國籍
      grdDataList.ColWidth(7) = 1200 '備註
      grdDataList.ColWidth(12) = 1000 '關聯編號
      grdDataList.ColWidth(13) = 1200 '關聯名稱
      grdDataList.ColWidth(14) = 1200 '關聯關係
      grdDataList.ColWidth(15) = 1200 '關聯說明
      'end 2017/02/14
   End If
End If
End Sub

'Add by Amy 2023/03/08 變動的欄位
Private Sub GetField()
    ReDim strField(grdDataList.Cols - 1)
    For j = 0 To grdDataList.Cols - 1
        strField(j) = grdDataList.TextMatrix(0, j)
    Next j
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(strField)
        If UCase(strField(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function
'end 2023/03/08

'Add by Amy 2023/08/30 整理
Public Sub PubShowNextData()
   Dim strRepCon As String
   '列印對造資料
   If cmdState = 5 Then
      strRepCon = Text2
      If Option3(0).Value = True Then
         strRepCon = strRepCon & " (字首比對)"
      ElseIf Option3(1).Value = True Then
         strRepCon = strRepCon & " (模糊比對)"
      End If
      cmdok(cmdState).Enabled = False
   End If
   Call PubShowNextForm(cmdState, Me, grdDataList, strField, _
      IIf(Check3.Value = vbChecked, True, False), , , Text6, Text7, , , , , _
      , , m_bolPrintRight, cboSort, strRepCon)
   If cmdState = 5 Then cmdok(cmdState).Enabled = True
End Sub

'Mark by Amy 2023/08/30 改抓共用
Public Sub PubShowNextData_Old()
'Dim blnPrintAdd As Boolean
'Dim ii As Integer
'Dim j As Integer
'Dim strTmp As String
'
''Modify by Amy 2023/06/08 欄位改動態
'Select Case cmdState
'Case 0 '聯絡人
'      Me.Enabled = False
'      For i = 1 To grdDataList.Rows - 1
'      grdDataList.col = 0
'      grdDataList.row = i
'      If Trim(grdDataList.Text) = "V" Then
'         grdDataList.col = 0
'         grdDataList.Text = ""
'         'Add By Sindy 2012/3/21
'         grdDataList.col = 1
'         'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'         If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'            For j = 0 To grdDataList.Cols - 1
'                '呆帳
'                If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                    grdDataList.CellBackColor = &HFF& '紅色
'                '活化客戶
'                Else
'                    grdDataList.col = j
'                    grdDataList.CellBackColor = vbYellow
'                End If
'            Next
'         'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'         'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'         ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'            And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'            Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                For j = 0 To grdDataList.Cols - 1
'                    grdDataList.col = j
'                    grdDataList.CellBackColor = &H0 '黑色
'                    grdDataList.CellForeColor = &HFF00FF '粉紅色
'                Next j
'         'Modify by Amy 2013/12/10 +判斷對造
'         ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'            For j = 0 To grdDataList.Cols - 1
'               grdDataList.col = j
'               grdDataList.CellBackColor = &H8080FF
'            Next j
'         Else
'         '2012/3/21 End
'            For j = 0 To grdDataList.Cols - 1
'               If j <> 1 Then
'                   grdDataList.col = j
'                   grdDataList.CellBackColor = QBColor(15)
'               End If
'            Next j
'         End If
'         If fnSaveParentForm(Me) = False Then
'             Me.Enabled = True
'             Exit Sub
'         End If
'         grdDataList.col = 1
'         Screen.MousePointer = vbHourglass
'         strTmp = Pub_RplStr(grdDataList.Text)
'         Select Case Left(strTmp, 1)
'            Case "R"
'               frm100101_14.Show
'               frm100101_14.Tag = strTmp
'               frm100101_14.StrMenu
'            Case Else
'               strExc(2) = "F"
'               If Left(strTmp, 1) = "X" Then
'                  strExc(0) = "select st03 from customer,staff where cu01(+)='" & Left(strTmp, 8) & "' and cu02(+)='" & Mid(strTmp, 9, 1) & "' and st01(+)=cu13"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     strExc(2) = "" & RsTemp.Fields(0)
'                  End If
'               End If
'               If Left(strExc(2), 1) = "F" Then
'                  frm100101_17.Show
'                  frm100101_17.Tag = strTmp
'                  frm100101_17.StrMenu
'               Else
'                  frm100101_18.Show
'                  frm100101_18.Tag = strTmp
'                  frm100101_18.CmdOk1(0).Enabled = m_bolPrintRight
'                  frm100101_18.StrMenu
'               End If
'         End Select
'         Screen.MousePointer = vbDefault
'         grdDataList.col = 0
'         grdDataList.Text = ""
'         'Add By Sindy 2012/3/21
'         grdDataList.col = 1
'         'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'         If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'            For j = 0 To grdDataList.Cols - 1
'                '呆帳
'                If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                    grdDataList.CellBackColor = &HFF& '紅色
'                '活化客戶
'                Else
'                    grdDataList.col = j
'                    grdDataList.CellBackColor = vbYellow
'                End If
'            Next
'         'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'         'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'         ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'            And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'            Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                For j = 0 To grdDataList.Cols - 1
'                    grdDataList.col = j
'                    grdDataList.CellBackColor = &H0 '黑色
'                    grdDataList.CellForeColor = &HFF00FF '粉紅色
'                Next j
'         'Modify by Amy 2013/12/10 +判斷對造
'         ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'            For j = 0 To grdDataList.Cols - 1
'               grdDataList.col = j
'               grdDataList.CellBackColor = &H8080FF
'            Next j
'         Else
'         '2012/3/21 End
'            For j = 0 To grdDataList.Cols - 1
'               If j <> 1 Then
'                   grdDataList.col = j
'                   grdDataList.CellBackColor = QBColor(15)
'               End If
'            Next j
'         End If
'         Me.Enabled = True
'         Exit Sub
'      End If
'      Next i
'      Me.Enabled = True
'Case 1 '關係企業
'      Me.Enabled = False
'      strExc(9) = "" 'Added by Lydia 2017/08/18 勾選清單
'      'Modified by Lydia 2018/05/24 改由啟用日控制
'      If strSrvDate(1) < 國外部關聯企業啟用日 Then
'           cnnConnection.Execute "DELETE FROM R140407 where ID='" & strUserNum & "' "
'      End If
'      'end 2018/05/24
'      For i = 1 To grdDataList.Rows - 1
'        grdDataList.col = 0
'        grdDataList.row = i
'        If Trim(grdDataList.Text) = "V" Then
'            grdDataList.col = 0
'            grdDataList.Text = ""
'            'Add By Sindy 2012/3/21
'            grdDataList.col = 1
'            'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'            If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                For j = 0 To grdDataList.Cols - 1
'                    '呆帳
'                    If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                        grdDataList.CellBackColor = &HFF& '紅色
'                    '活化客戶
'                    Else
'                        grdDataList.col = j
'                        grdDataList.CellBackColor = vbYellow
'                    End If
'                Next
'            'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'            'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'            ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'                And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'                Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                    For j = 0 To grdDataList.Cols - 1
'                       grdDataList.col = j
'                       grdDataList.CellBackColor = &H0 '黑色
'                       grdDataList.CellForeColor = &HFF00FF '粉紅色
'                    Next j
'            'Modify by Amy 2013/12/10 +判斷對造
'            ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'               For j = 0 To grdDataList.Cols - 1
'                  grdDataList.col = j
'                  grdDataList.CellBackColor = &H8080FF
'               Next j
'            Else
'            '2012/3/21 End
'               For j = 0 To grdDataList.Cols - 1
'                   If j <> 1 Then
'                       grdDataList.col = j
'                       grdDataList.CellBackColor = QBColor(15)
'                   End If
'               Next j
'            End If
'            grdDataList.col = 1
'            'Add By Sindy 2011/01/03 檢查國內外權限
'            If CheckSR12(Pub_RplStr(grdDataList.Text)) = True Then
'            '2011/01/03 End
'               Screen.MousePointer = vbHourglass
'                'Modified by Lydia 2018/05/24 改由啟用日控制
'                If strSrvDate(1) < 國外部關聯企業啟用日 Then
'                    Call StrMenu(Pub_RplStr(grdDataList.Text))
'                Else
'                    'Added by Lydia 2017/02/14 抓關聯企業改成模組,暫存R100114_1
'                    'Modified by Lydia 2017/08/18 是否清除先前記錄
'                    'j = PUB_GetR100114_1(Me.Name, Pub_RplStr(GrdDataList.Text))
'                     j = PUB_GetR100114_1(IIf(strExc(9) = "", True, False), Me.Name, Pub_RplStr(grdDataList.Text))
'                    strExc(9) = strExc(9) & IIf(strExc(9) <> "", ",", "") & Pub_RplStr(grdDataList.Text)
'                    'end 2017/08/18
'                End If
'
'               cmdOK(1).Enabled = False
'               Screen.MousePointer = vbDefault
'            End If
'        End If
'      Next i
'      'Modified by Lydia 2018/05/24 改由啟用日控制
'      If strSrvDate(1) < 國外部關聯企業啟用日 Then
'           Call StrMenu1
'      Else
'           'Added by Lydia 2017/02/14 抓關聯企業改成模組,暫存R100114_1
'           If j > 1 Then Call StrMenu1
'      End If
'      'end 2018/05/24
'
'      Me.Enabled = True
''潛在客戶資料
'Case 2
'   Me.Enabled = False
'      For i = 1 To grdDataList.Rows - 1
'      grdDataList.col = 0
'      grdDataList.row = i
'      If Trim(grdDataList.Text) = "V" Then
'         grdDataList.col = 0
'         grdDataList.Text = ""
'         'Add By Sindy 2012/3/21
'         grdDataList.col = 1
'         'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'         If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'            For j = 0 To grdDataList.Cols - 1
'                '呆帳
'                If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                    grdDataList.CellBackColor = &HFF& '紅色
'                '活化客戶
'                Else
'                    grdDataList.col = j
'                    grdDataList.CellBackColor = vbYellow
'                End If
'            Next
'         'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'         'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'        ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'            And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'            Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                For j = 0 To grdDataList.Cols - 1
'                    grdDataList.col = j
'                    grdDataList.CellBackColor = &H0 '黑色
'                    grdDataList.CellForeColor = &HFF00FF '粉紅色
'                Next j
'         'Modify by Amy 2013/12/10 +判斷對造
'         ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'            For j = 0 To grdDataList.Cols - 1
'               grdDataList.col = j
'               grdDataList.CellBackColor = &H8080FF
'            Next j
'         Else
'         '2012/3/21 End
'            For j = 0 To grdDataList.Cols - 1
'               If j <> 1 Then
'                   grdDataList.col = j
'                   grdDataList.CellBackColor = QBColor(15)
'               End If
'            Next j
'         End If
'        If fnSaveParentForm(Me) = False Then
'            Me.Enabled = True
'            Exit Sub
'        End If
'        grdDataList.col = 1
'        Screen.MousePointer = vbHourglass
'
'         strTmp = Pub_RplStr(grdDataList.Text)
'         Select Case Left(strTmp, 1)
'            Case "X"
'               If Mid(strTmp, 10, 1) = "-" Then
'                  strTmp = Left(strTmp, 9)
'               End If
'               frm100101_11.Show
'               frm100101_11.Tag = strTmp
'               frm100101_11.StrMenu
'            Case "Y"
'               If Mid(strTmp, 10, 1) = "-" Then
'                  strTmp = Left(strTmp, 9)
'               End If
'               frm100101_10.Show
'               frm100101_10.Tag = strTmp
'               frm100101_10.StrMenu
'            Case "R"
'               'Modify By Sindy 2009/06/24 判斷是國外或是國內潛在客戶
'               strExc(0) = "select * from potcustomer where pcu01(+)='" & Left(strTmp, 8) & "' and pcu02(+)='" & Mid(strTmp, 9, 1) & "' "
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               strExc(2) = ""
'               If intI = 1 Then
'                  strExc(2) = "" & RsTemp.Fields(0)
'               End If
'               If strExc(2) <> "" Then '國外
'                  frm100101_14.Show
'                  frm100101_14.Tag = strTmp
'                  frm100101_14.StrMenu
'               Else '國內
'                  frm100101_21.Show
'                  frm100101_21.Tag = strTmp
'                  frm100101_21.StrMenu
'               End If
'            'Add by Amy 2015/03/27 +客戶端平台帳號
'            Case "平"
'                'Modify by Amy 2015/04/15 改以平台編號抓權限
'                If PUB_ChkCustWebLimit(grdDataList.TextMatrix(grdDataList.row, GetValue("收文日")), strUserNum) = True Then
'                    frm100101_27.Show
'                    frm100101_27.Tag = Trim(grdDataList.TextMatrix(grdDataList.row, GetValue("收文日")))
'                    frm100101_27.StrMenu
'                Else
'                    Me.Show
'                    MsgBox "您無權限查詢此客戶端平台帳號！", vbInformation
'                End If
'            'Add By Sindy 2009/07/22
'            Case Else
'               'Modify By Sindy 2012/3/21 +不得代理案件之客戶或代理人
'               If InStr(strTmp, "-") = 0 Then
'                  frm100101_25.Show
'                  frm100101_25.Tag = strTmp
'                  frm100101_25.StrMenu
'               Else
'               '2012/3/21 End
'                  frm100101_22.Show
'                  frm100101_22.Tag = strTmp
'                  frm100101_22.StrMenu
'               End If
'            '2009/07/22 End
'         End Select
'
'        Screen.MousePointer = vbDefault
'         grdDataList.col = 0
'         grdDataList.Text = ""
'         'Add By Sindy 2012/3/21
'         grdDataList.col = 1
'         'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'         If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'            For j = 0 To grdDataList.Cols - 1
'                '呆帳
'                If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                    grdDataList.CellBackColor = &HFF& '紅色
'                '活化客戶
'                Else
'                    grdDataList.col = j
'                    grdDataList.CellBackColor = vbYellow
'                End If
'            Next
'         'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'         'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'         ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'            And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'            Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                For j = 0 To grdDataList.Cols - 1
'                    grdDataList.col = j
'                    grdDataList.CellBackColor = &H0 '黑色
'                    grdDataList.CellForeColor = &HFF00FF '粉紅色
'                Next j
'         'Modify by Amy 2013/12/10 +判斷對造
'         ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'            For j = 0 To grdDataList.Cols - 1
'               grdDataList.col = j
'               grdDataList.CellBackColor = &H8080FF
'            Next j
'         Else
'         '2012/3/21 End
'            For j = 0 To grdDataList.Cols - 1
'               If j <> 1 Then
'                   grdDataList.col = j
'                   grdDataList.CellBackColor = QBColor(15)
'               End If
'            Next j
'         End If
'         Me.Enabled = True
'         Exit Sub
'      End If
'      Next i
'      Me.Enabled = True
'Case 3 '結束
'   Unload frm140407
'   Set frm140407 = Nothing
'
'Case 4 '往來記錄
'      Me.Enabled = False
'      For i = 1 To grdDataList.Rows - 1
'      grdDataList.col = 0
'      grdDataList.row = i
'      If Trim(grdDataList.Text) = "V" Then
'         grdDataList.col = 0
'         grdDataList.Text = ""
'         'Add By Sindy 2012/3/21
'         grdDataList.col = 1
'         'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'         If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'            For j = 0 To grdDataList.Cols - 1
'                '呆帳
'                If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                    grdDataList.CellBackColor = &HFF& '紅色
'                '活化客戶
'                Else
'                    grdDataList.col = j
'                    grdDataList.CellBackColor = vbYellow
'                End If
'            Next
'         'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'         'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'         ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'            And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'            Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                For j = 0 To grdDataList.Cols - 1
'                    grdDataList.col = j
'                    grdDataList.CellBackColor = &H0 '黑色
'                    grdDataList.CellForeColor = &HFF00FF '粉紅色
'                Next j
'         'Modify by Amy 2013/12/10 +判斷對造
'         ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'            For j = 0 To grdDataList.Cols - 1
'               grdDataList.col = j
'               grdDataList.CellBackColor = &H8080FF
'            Next j
'         Else
'         '2012/3/21 End
'            For j = 0 To grdDataList.Cols - 1
'               If j <> 1 Then
'                   grdDataList.col = j
'                   grdDataList.CellBackColor = QBColor(15)
'               End If
'            Next j
'         End If
'         If fnSaveParentForm(Me) = False Then
'             Me.Enabled = True
'             Exit Sub
'         End If
'         grdDataList.col = 1
'         Screen.MousePointer = vbHourglass
'         strTmp = Pub_RplStr(grdDataList.Text)
'
'         'Modify By Sindy 2009/06/24 判斷是國外或是國內潛在客戶
'         '客戶檔
'         strExc(3) = "select cu12,cu13 from customer where cu01(+)='" & Left(strTmp, 8) & "' and cu02(+)='" & Mid(strTmp, 9, 1) & "' "
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(3))
'         strExc(4) = ""
'         If intI = 1 Then
'            strExc(4) = "" & RsTemp.Fields("cu12")
'         End If
'         '潛在客戶檔
'         strExc(0) = "select * from potcustomer where pcu01(+)='" & Left(strTmp, 8) & "' and pcu02(+)='" & Mid(strTmp, 9, 1) & "' "
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         strExc(2) = ""
'         If intI = 1 Then
'            strExc(2) = "" & RsTemp.Fields(0)
'         End If
''         If strExc(2) <> "" Or Left(Trim(strTmp), 1) = "Y" Or Left(Trim(strExc(4)), 1) = "F" Then '國外
'            frm100101_15.Show
'            '2008/11/19 ADD BY SONIA
'            frm100101_15.CRdateF = Text6
'            frm100101_15.CRdateT = IIf(Len(Text7) <= 0, "", IIf(Len(Text7) <= 0, (ServerDate - 19110000), Text7))
'            frm100101_15.CRtype = cboSort.Text
'            '2008/11/19 END
'            frm100101_15.Tag = strTmp
'            'Modify By Sindy 2020/5/19
'            'Modify By Sindy 2021/3/25 + Or Left(Trim(strTmp), 1) = "平"
'            'If strExc(2) <> "" Or Left(Trim(strTmp), 1) = "Y" Or Left(Trim(strExc(4)), 1) = "F" Or Left(Trim(strTmp), 1) = "平" Then  '國外
'            If strExc(2) <> "" Or _
'                     (Left(Trim(strTmp), 1) = "Y" And Left(Pub_StrUserSt03, 1) = "F") Or _
'                     Left(Trim(strExc(4)), 1) = "F" Or _
'                     Pub_StrUserSt03 = "M51" Or _
'                     Left(Trim(strTmp), 1) = "平" Then '國外
'               frm100101_15.m_quyDataKind = 0
'               frm100101_15.StrMenu
'            Else
'               frm100101_15.m_quyDataKind = 1
'               frm100101_15.StrMenu2
'            End If
'            '2020/5/19 END
''         Else '國內
'            'Modify By Sindy 2020/5/19
''            frm100101_20.Show
''            frm100101_20.Tag = strTmp
''            frm100101_20.StrMenu
''         End If
'
'         Screen.MousePointer = vbDefault
'         grdDataList.col = 0
'         grdDataList.Text = ""
'         'Add By Sindy 2012/3/21
'         grdDataList.col = 1
'         'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'         If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'            For j = 0 To grdDataList.Cols - 1
'                '呆帳
'                If Right(grdDataList.Text, 1) = "$" And j = 1 Then
'                    grdDataList.CellBackColor = &HFF& '紅色
'                '活化客戶
'                Else
'                    grdDataList.col = j
'                    grdDataList.CellBackColor = vbYellow
'                End If
'            Next
'         'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'         'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'         ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'            And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'            Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                For j = 0 To grdDataList.Cols - 1
'                    grdDataList.col = j
'                    grdDataList.CellBackColor = &H0 '黑色
'                    grdDataList.CellForeColor = &HFF00FF '粉紅色
'                Next j
'         'Modify by Amy 2013/12/10 +判斷對造
'         ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'            For j = 0 To grdDataList.Cols - 1
'               grdDataList.col = j
'               grdDataList.CellBackColor = &H8080FF
'            Next j
'         Else
'         '2012/3/21 End
'            For j = 0 To grdDataList.Cols - 1
'               If j <> 1 Then
'                   grdDataList.col = j
'                   grdDataList.CellBackColor = QBColor(15)
'               End If
'            Next j
'         End If
'         Me.Enabled = True
'         Exit Sub
'      End If
'      Next i
'      Me.Enabled = True
''Add by Amy 2013/12/04
'Case 5 '列印對造資料
'    'Modify by Amy 2014/02/25 改印暫存資料
'    'PrintDataA4
'    PrintDataA4_Temp
'    'end 2014/02/25
'Case Else
'End Select
End Sub

Private Sub cboSort_Click()
   Dim iPos As Integer
   iPos = InStr(cboSort.Text, Chr(1))
   If iPos > 0 Then
      cboSort.Text = Left(cboSort.Text, iPos - 1)
   End If
End Sub

Private Sub cboSort_GotFocus()
   If cboSort.Locked = False Then
      CloseIme
      SendMessage cboSort.hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

Private Sub cmdAddSort_Click()
   If AddList(lstSort, cboSort) = True Then
      Text10 = ComposeList(lstSort)
      cboSort = ""
   End If
   cboSort.SetFocus
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Private Sub cmdRemSort_Click()
   If RemoveList(lstSort) = True Then
      Text10 = ComposeList(lstSort)
      cboSort.SetFocus
   End If
End Sub

'Modify by Amy 2022/08/08 名稱查詢語法改至共用Function,並整理程式
Private Sub cmdSearch_Click()
    Dim strCheckWay As String
    Dim stSQL1 As String, stSQL2 As String, stSQL3 As String, stSQL4 As String, stSQL5 As String
    Dim strSwhSQL1 As String, strSwhSQL2 As String, strSubSQL1 As String, strSubSQL2 As String
    Dim strNo As String, Str01 As String, strFields As String
    Dim s As Integer, IsDevelop As Boolean, bCancel As Boolean
    Dim strRtnVal As String 'Add by Amy 2023/10/03
On Error GoTo ErrHnd

    bolPrint = False '先設定無對造
    StrToPrint = ""
    lngCounterI = 0
    
    '潛在客戶編號
    If Option2(0).Value = True Then
       If Len(Trim(Text1)) = 0 Then
          s = MsgBox("條件不可空白", , "輸入條件錯誤")
          Text1.SetFocus
          Exit Sub
       End If
    End If
    '潛在客戶/聯絡人名稱
    If Option2(1).Value = True Then
        If Len(Trim(Text2)) = 0 Then
            s = MsgBox("條件不可空白", , "輸入條件錯誤")
            Text2.SetFocus
            Exit Sub
        End If
    Else
        '若輸開發者直接按Enter,人員名稱不會重抓
        Call Text3_Validate(bCancel)
    End If
    'E-Mail
    If Option2(2).Value = True Then
        If Len(Trim(Text9)) = 0 Then
            s = MsgBox("條件不可空白", , "輸入條件錯誤")
            Text9.SetFocus
            Exit Sub
        End If
    End If
    
    '開發者
    If Option2(3).Value = True Then
        If Len(Trim(Text3)) = 0 Then
            s = MsgBox("條件不可空白", , "輸入條件錯誤")
            Text3.SetFocus
        End If
    End If
    
    '開發日期
    If Option2(4).Value = True Then
        If Len(Trim(Text4)) = 0 And Len(Trim(Text5)) = 0 Then
            s = MsgBox("條件不可空白", , "輸入條件錯誤")
            Text4.SetFocus
        End If
    End If
    
   'Add by Amy 2023/10/03 屬於查詢置換字彈訊息
   If Option2(1).Value = True Then
      If ChkQuryChgTxt(Text2, strRtnVal) = True Then
         frm100137_1.Caption = "訊息"
         frm100137_1.txtOrg = Text2
         frm100137_1.txtChg = strRtnVal
         frm100137_1.Show vbModal
      End If
   End If
    
    ClearQueryLog (Me.Name) '清除查詢印表記錄檔欄位
    Screen.MousePointer = vbHourglass
    grdDataList.Clear
    grdDataList.Rows = 2
    SetDataListWidth
 
    'Modify by Amy 2022/08/19 +OrgN
    strFields = ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明,'' AS OrgN "
    
    If Option2(0).Value = True Then
'*** 編號 ***
        pub_QL05 = pub_QL05 & ";" & Option2(0).Caption & Trim(Text1)
        '潛在客戶
        If UCase(Left(Trim(Text1), 1)) = "R" Then
            strSql = "Select ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,Nvl(PCU08,Decode(PCU03,null,PCU07,RTrim(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,Decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCU40 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From PotCustomer,Nation,Staff Where PCU09=NA01(+) And PCU01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' And SubStr(LTrim(PCU38),1,5)=ST01(+) "
            strSql = strSql & " Union All Select ' ' AS V ,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,Nvl(POC03,Decode(POC23,null,POC27,RTrim(POC23||' '||POC24||' '||POC25||' '||POC26))) AS 名稱,NA03 AS 國籍,POC13 AS 智權人員,' ' AS 類別,POC14 AS 狀態,POC15 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From PotCustomer1,Nation,Staff Where POC04=NA01(+) And POC01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' And POC13=ST01(+) "
        Else
            strSql = "Select ' ' AS V ,CU01||CU02||Decode(CU02,'0','','＊')||Decode(CU111,'Y','$','')||Decode(CU121,'Y','●','') AS 編號,Nvl(CU04,Decode(CU05,null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From Customer,Nation,Staff Where CU10=NA01(+) And CU01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' And CU13=ST01(+) "
            strSql = strSql & " Union All Select ' ' AS V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(FA77,'Y','$','') AS 編號,Nvl(FA04,Decode(FA05,null,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) AS 名稱,NA03 AS 國籍,'' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From Fagent,Nation Where FA10=NA01(+) And FA01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' "
            strSql = strSql & " Union All Select ' ' AS V,NT01||Decode(NT21,null,'⊕','') AS 編號,Nvl(NT02,Decode(NT03,null,NT07,NT03||' '||NT04||' '||NT05||' '||NT06)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From NotAgent,Nation,Staff Where NT08=NA01(+) And NT01='" & IIf(Len(Trim(Text1)) >= 3, Trim(Text1), Right("000" & Trim(Text1), 3)) & "' And NT18=ST01(+) "
            'Add by Amy 2023/12/11 +風險檢查對象
            strSql = strSql & " Union All " & GetSearchRiskChkSql(1, Me.Name, Trim(Text1))
        End If
    ElseIf Option2(1).Value = True Then
'*** 名稱 ***
        pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & Trim(Text2)
        '模糊比對
        If Option3(0).Value = False Then
            pub_QL05 = pub_QL05 & ";" & Option3(1).Caption
            strCheckWay = ">0"
        '字首比對
        Else
            pub_QL05 = pub_QL05 & ";" & Option3(0).Caption
            strCheckWay = "=1"
        End If
        '對造
        stSQL1 = " And CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
        stSQL2 = " And CP01 IN (" & SQLGrpStr("", 1) & ") "
        stSQL3 = " And CP01 IN (" & SQLGrpStr("", 3) & ") "
        stSQL4 = " And CP01 IN (" & SQLGrpStr("", 4) & ") "
        stSQL5 = " And CP01 IN (" & SQLGrpStr("", 5) & ") "
        '含投資法務開拓
        If Check1.Value = 1 Then IsDevelop = True
        '刪除對造暫存檔資料
        cnnConnection.Execute "Delete From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' "
        strSql = GetSearchNameSql(Me.Name, Text2, strCheckWay, IsDevelop, True, stSQL1, stSQL2, stSQL3, stSQL4, stSQL5)
        pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & Trim(Text2)
    ElseIf Option2(2).Value = True Then
'*** E-Mail ***
        pub_QL05 = pub_QL05 & ";" & Option2(2).Caption & Trim(Text9)
        'Modified by Lydia 2024/09/18 +財務副本信箱CU200
        strSql = "Select ' ' AS V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(CU111,'Y','$','')||Decode(CU121,'Y','●','') AS 編號,Nvl(CU04,Decode(CU05,null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From Customer,Nation,PotCUstomer,Staff  Where (InStr(NLS_Upper(CU20),'" & UCase(ChgSQL(Trim(Text9))) & "')>0 or  InStr(NLS_Upper(CU115),'" & UCase(ChgSQL(Trim(Text9))) & "')>0 or InStr(NLS_Upper(CU116),'" & UCase(ChgSQL(Trim(Text9))) & "')>0  or InStr(NLS_Upper(CU117),'" & UCase(ChgSQL(Trim(Text9))) & "')>0 or InStr(NLS_Upper(CU118),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 or InStr(NLS_Upper(CU200),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0)  And CU10=NA01(+)  And PCU01(+)=CU01 And CU13=ST01(+) "
        strSql = strSql & " Union All Select ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,Nvl(PCU08,Decode(PCU03,null,PCU07,RTrim(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS 名稱,NA03 AS 國籍,PCU38 AS 智權人員,Decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCU40 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From PotCustomer,Nation,Staff  Where (InStr(NLS_Upper(PCU18),'" & UCase(ChgSQL(Trim(Text9))) & "') >0 ) And PCU09=NA01(+) And SubStr(LTrim(PCU38),1,5)=ST01(+) "
        strSql = strSql & " Union All Select ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,Nvl(POC03,Decode(POC23,null,POC27,RTrim(POC23||' '||POC24||' '||POC25||' '||POC26))) AS 名稱,NA03 AS 國籍,POC13 AS 智權人員,' ' AS 類別,POC14 AS 狀態,POC15 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From PotCustomer1,Nation,Staff  Where (InStr(NLS_Upper(POC09),'" & UCase(ChgSQL(Trim(Text9))) & "') >0 ) And POC04=NA01(+) And POC13=ST01(+) "
        'Modified by Lydia 2024/09/18 +財務副本信箱FA134
        strSql = strSql & " Union All Select ' ' AS V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(FA77,'Y','$','') AS 編號,Nvl(FA04,Decode(FA05,null,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) AS 名稱,NA03 AS 國籍,' ' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From Fagent,Nation Where (InStr(NLS_Upper(FA16),'" & UCase(ChgSQL(Trim(Text9))) & "')> 0 or InStr(NLS_Upper(FA79),'" & UCase(ChgSQL(Trim(Text9))) & "')> 0 or InStr(NLS_Upper(FA105),'" & UCase(ChgSQL(Trim(Text9))) & "')> 0 or InStr(NLS_Upper(FA80),'" & UCase(ChgSQL(Trim(Text9))) & "')> 0 or InStr(NLS_Upper(FA81),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 Or InStr(NLS_Upper(FA82),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 Or InStr(NLS_Upper(FA134),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0) And FA10=NA01(+)"
        strSql = strSql & " Union All Select ' ' AS V,PCC01||'0-'||PCC02 AS 編號,Nvl(PCC05,Nvl(PCC03,PCC04)) AS 名稱,NA03 AS 國籍,PCU38 AS 智權人員,Decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,' ' AS 狀態,PCC13 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From PotCustCont,PotCustomer,Nation,Staff Where (InStr(NLS_Upper(PCC08),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 ) And PCC01=PCU01(+) And PCU09=NA01(+) And SubStr(LTrim(PCU38),1,5)=ST01(+) "
        '開拓客戶資料
        If Check1.Value = 1 Then
            strSql = strSql & " Union All Select ' ' AS V,ECD02||'-'||LPAD(ECD01,6,'0') AS 編號,ECD03||' '||ECD04 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,eca02 AS 類別,'投法開拓'||Decode(ECD15,null,null,'-'||ECD15) AS 狀態,ECD16 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From expandcusdetail, expandcusattr ,Nation Where (InStr(NLS_Upper(ECD13),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 ) And ECD10=NA01(+) And ECD02=ECA01(+) "
        End If
        'Add By Sindy 2023/8/21 + 電子報特殊名單
        strSql = strSql & " Union All " & _
                    "Select ' ' as V,'電子報特殊名單-'||TBNP09 as 編號,TBNP01 as 名稱,'' as 國籍,'' as 智權人員,'' AS 類別,TBNP10 as 狀態,'' as 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & _
                    " From TMBulletinNp Where (instr(NLS_Upper(TBNP01),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 ) And TBNP08='M' "
        '2023/8/21 END
    ElseIf Option2(3).Value = True Then
'*** 開發者員編 ***
        strSql = "Select ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,Nvl(PCU08,Decode(PCU03,null,PCU07,RTrim(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,Decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCU40 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From PotCustomer,Nation,Staff Where (InStr(PCU38,'" & ChgSQL(Text3) & "') > 0)   And  PCU09=NA01(+) And SubStr(LTrim(PCU38),1,5)=ST01(+) "
        strSql = strSql & " Union All Select ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,Nvl(POC03,Decode(POC23,null,POC27,RTrim(POC23||' '||POC24||' '||POC25||' '||POC26))) AS 名稱,NA03 AS 國籍,POC13 AS 智權人員,' ' AS 類別,POC14 AS 狀態,POC15 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From PotCustomer1,Nation,Staff Where (InStr(POC13,'" & ChgSQL(Text3) & "') > 0)   And  POC04=NA01(+) And POC13=ST01(+) "
        strSql = strSql & " Union All Select ' ' AS V,PCC01||'0-'||PCC02 AS 編號,Nvl(PCC05,Nvl(PCC03,PCC04)) AS 名稱,' ' AS 國籍,' ' AS 智權人員,' ' AS 類別,' ' AS 狀態,PCC13 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From PotCustCont Where InStr(PCC12,'" & ChgSQL(strTp(1)) & "') > 0  And SubStr(PCC01,1,1)='R' "
        strSql = strSql & " Union All Select ' ' AS V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(CU111,'Y','$','')||Decode(CU121,'Y','●','') AS 編號,Nvl(CU04,Decode(CU05,null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,'  ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From Customer,Nation,Staff  Where (InStr(CU129,'" & ChgSQL(Text3) & "')>0 ) And CU10=NA01(+) And CU13=ST01(+) "
        strSql = strSql & " Union All Select ' ' AS V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(FA77,'Y','$','') AS 編號,Nvl(FA04,Decode(FA05,null,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) AS 名稱,NA03 AS 國籍,' ' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From Fagent,Nation Where (InStr(FA94,'" & ChgSQL(Text3) & "')> 0  ) And FA10=NA01(+)  "
        pub_QL05 = pub_QL05 & ";" & Option2(3).Caption & Text3 & Label10
    ElseIf Option2(4).Value = True Then
'*** 開發日 ***
        pub_QL05 = pub_QL05 & ";" & Option2(4).Caption & Text4 & "-" & Text5
        strSql = "Select ' ' AS V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號,Nvl(CU04,Decode(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From Customer,Nation,Staff Where cu14 >='" & ChangeTStringToWString(Text4) & "' And cu14 <='" & ChangeTStringToWString(Text5) & "' And CU10=NA01(+) And CU13=ST01(+) "
        strSql = strSql & " Union All Select ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,Nvl(PCU08,Decode(PCU03,null,PCU07,RTrim(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,Decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCU40 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From PotCustomer,Nation,Staff  Where PCU37 >='" & ChangeTStringToWString(Text4) & "' And PCU37 <='" & ChangeTStringToWString(Text5) & "' And pcu09=NA01(+) And SubStr(LTrim(PCU38),1,5)=ST01(+) "
        strSql = strSql & " Union All Select ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,Nvl(POC03,Decode(POC23,null,POC27,RTrim(POC23||' '||POC24||' '||POC25||' '||POC26))) AS 名稱,NA03 AS 國籍,POC13 AS 智權人員,' ' AS 類別,POC14 AS 狀態,POC15 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From PotCustomer1,Nation,Staff  Where POC12 >='" & ChangeTStringToWString(Text4) & "' And POC12 <='" & ChangeTStringToWString(Text5) & "' And POC04=NA01(+) And POC13=ST01(+) "
        strSql = strSql & " Union All Select ' ' AS V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(FA77,'Y','$','') AS 編號,Nvl(FA04,Decode(FA05,null,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) AS 名稱,NA03 AS 國籍,' ' AS 智權人員,NULL AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From Fagent,Nation Where FA11 >='" & ChangeTStringToWString(Text4) & "' And  FA11 <='" & ChangeTStringToWString(Text5) & "'  And FA10=NA01(+) "
        strSql = strSql & " Union All Select ' ' AS V,PCC01||'0-'||PCC02 AS 編號,Nvl(PCC05,Nvl(PCC03,PCC04)) AS 名稱,NA03 AS 國籍,PCU38 AS 智權人員,Decode(PCU11," & 國外潛在客戶類別 & ",PCU11)  AS 類別,' ' AS 狀態,PCC13 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & strFields & " From PotCustCont,PotCustomer,Nation,Staff Where PCC11 >='" & ChangeTStringToWString(Text4) & "' And PCC11 <='" & ChangeTStringToWString(Text5) & "' And PCC01=PCU01(+) And PCU09=NA01(+) And SubStr(LTrim(PCU38),1,5)=ST01(+) "
    End If
    '名稱
    If Option2(1).Value = True Then
        'Modify by Amy 2022/08/19 因名稱前加找到之中 or 英 or 日欄位,導致同編號無法排於一起 原:Order by Upper(名稱),編號
        'ex: 查 SONN & PARTNER 2筆(Y45656000/1)及投法981-000001,2筆Y編號無法排一起
        strSql = "Select X.*,Decode(Ocu01,null, '',Nvl(Ocu03,0)) AS Ocu03 From (" & strSql & ") X, OldCustomer Where SubStr(編號,1,8)= ocu01(+) order by upper(OrgN) "
    Else
        strSql = "Select X.*,Decode(Ocu01,null, '',Nvl(Ocu03,0)) AS Ocu03 From (" & strSql & ") X, OldCustomer Where SubStr(編號,1,8)= ocu01(+) order by 編號 "
    End If
    '含投資法務開拓
    If Check1.Value = 1 Then
        pub_QL05 = pub_QL05 & ";" & Check1.Caption
    End If
    '往來日期
    If Trim(Text6) <> "" Or Trim(Text7) <> "" Then
        pub_QL05 = pub_QL05 & ";" & Left(Label5, 5) & Text6 & "-" & Text7
    End If
    '往來類別
    If Trim(cboSort.Text) <> "" Then
        pub_QL05 = pub_QL05 & ";" & Label7 & cboSort.Text
    End If
    
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/13 記錄此Form的查詢條件
    If adoRecordset.RecordCount <> 0 Then
        InsertQueryLog (adoRecordset.RecordCount)
        If Not cmdok(0).Enabled Then cmdok(0).Enabled = True
        If Not cmdok(1).Enabled Then cmdok(1).Enabled = True
        If Not cmdok(4).Enabled Then cmdok(4).Enabled = True
        Set grdDataList.Recordset = adoRecordset
    Else
        InsertQueryLog (0)
        '畫面訊息開放可Copy
        Pub_Can_Copy_Pic = True
        ShowNoData
        Pub_Can_Copy_Pic = False
        cmdok(0).Enabled = False
        cmdok(1).Enabled = False
        cmdok(4).Enabled = False
        grdDataList.Clear
    End If
    
    Me.grdDataList.Visible = False
    SetDataListWidth
    CheckOC
    
    'Modify by Amy 2023/03/08 欄位改動態
    With Me.grdDataList
        If .Rows > 0 Then
            For i = 1 To .Rows - 1
                .row = i
                .col = 1
                .CellForeColor = &H0 '字黑色
                'Modify by Amy 2023/08/30 變色改為共用函數
'                'Add by Amy 2023/01/18 +X 或 Y 編號若無案件顯示▼
'                If Check3.Value = vbChecked And (Left(.Text, 1) = 客戶編號 Or Left(.Text, 1) = 代理人編號) Then
'                    If ChkXYCase(Left(.Text, 9)) = False Then
'                        .Text = .Text & "▼"
'                    End If
'                End If
'                'end 2023/01/18
'                '活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'                If .TextMatrix(.row, GetValue("待活化客戶")) = "0" And Right(.TextMatrix(.row, GetValue("編號")), 1) <> "＊" Then
'                    For j = 0 To .Cols - 1
'                        '呆帳
'                        If Right(.Text, 1) = "$" And j = 1 Then
'                            GrdDataList.CellBackColor = &HFF& '紅色
'                        '活化客戶
'                        Else
'                            .col = j
'                            .CellBackColor = vbYellow
'                        End If
'                    Next j
'                ElseIf Right(.Text, 1) = "$" Then
'                    .CellBackColor = &HFF&
'                '客戶狀態為 解散/廢止/撤銷/死亡 顯示黑底粉字
'                ElseIf (Left(.Text, 1) = "Y" Or Left(.Text, 1) = "X" Or Left(.Text, 1) = "R") _
'                    And (.TextMatrix(i, GetValue("狀態")) = "解散" Or .TextMatrix(i, GetValue("狀態")) = "廢止" Or .TextMatrix(i, GetValue("狀態")) = "撤銷" Or .TextMatrix(i, GetValue("狀態")) = "死亡") Then
'                        For j = 0 To .Cols - 1
'                            .col = j
'                            .CellBackColor = &H0 '黑色
'                            .CellForeColor = &HFF00FF '粉紅色
'                        Next j
'                ElseIf Right(.Text, 1) = "⊕" Or .TextMatrix(i, GetValue("狀態")) = "對造" Or .TextMatrix(i, GetValue("狀態")) = "其他相關人" Then
                    'Modify by Amy 2023/09/26 依狀態更新智權人員改為共用函數
                    '對造重抓智權人資料
                    If Me.grdDataList.TextMatrix(i, GetValue("狀態")) = "對造" Or .TextMatrix(i, GetValue("狀態")) = "其他相關人" Then
                        bolPrint = True '有對造資料
'                        strNo = Pub_RplStr(.TextMatrix(i, GetValue("編號")))
'                        StrToPrint = strNo & ","
'                        Str01 = SystemNumber(strNo, 1)
'                        Select Case Str01
'                            Case "FCP", "FG"
'                                .TextMatrix(i, GetValue("智權人員")) = GetPrjSalesNM(PUB_GetFCPSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "FCL", "LIN"
'                                .TextMatrix(i, GetValue("智權人員")) = GetPrjSalesNM(PUB_GetFCLSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "FCT"
'                                .TextMatrix(i, GetValue("智權人員")) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "S"
'                                If .TextMatrix(i, GetValue("申請國家")) = "000" Then
'                                    .TextMatrix(i, GetValue("智權人員")) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                                Else
'                                    .TextMatrix(i, GetValue("智權人員")) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                                End If
'                            Case Else
'                                .TextMatrix(i, GetValue("智權人員")) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                        End Select
'                        .TextMatrix(i, GetValue("案件性質")) = .TextMatrix(i, GetValue("案件性質")) & PUB_GetRelateCasePropertyName(.TextMatrix(i, GetValue("總收文號")), "1")
'                        '更新智權人員至暫存檔
'                        strExc(0) = "Update R100102_1 Set R021003='" & .TextMatrix(i, GetValue("智權人員")) & "' Where R021014='" & Str01 & "' And R021015='" & SystemNumber(strNo, 2) & "' And R021016='" & SystemNumber(strNo, 3) & "' And R021017='" & SystemNumber(strNo, 4) & "' "
'                        cnnConnection.Execute strExc(0)
                    End If
'                    If Right(.Text, 1) = "⊕" Or .TextMatrix(i, GetValue("狀態")) = "對造" Then
'                        For j = 0 To .Cols - 1
'                            .col = j
'                            .CellBackColor = &H8080FF
'                        Next j
'                    End If
'                'CW03=7.媒介平台,顯示橘色
'                ElseIf Left(.TextMatrix(i, GetValue("編號")), 1) = "平" And .TextMatrix(i, GetValue("案件性質")) = "7" Then
'                    .CellBackColor = &H80FF& '橘色
'                End If
'                '國內外潛在客戶 智權人員欄需重抓資料(可能多筆)
'                If Left(.Text, 1) = "R" Then
'                    .TextMatrix(i, GetValue("智權人員")) = GetDevelopP(.TextMatrix(i, GetValue("智權人員")))
'                End If
                Call UpdQuerySales(Me.Name, grdDataList, strField)
                'end 2023/09/26
                Call SetMSGridColorQCus(0, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
                'end 2023/08/30
            Next i
        End If
    End With
    
    '若只有一筆資料, 則直接設定為點選此筆資料
    'Modify by Amy 2023/08/30 原程式寫至共用SetGridOneData,避免有沒改到的
    cmdok(4).BackColor = &H8000000F
    Call SetGridOneData
    'end 2023/08/30
    Me.grdDataList.Visible = True
    If bolPrint Then
        cmdok(5).Enabled = True
    Else
        cmdok(5).Enabled = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
        
ErrHnd:
    If Err.Number = -2147217900 Then
        MsgBox "輸入的文字無法查詢,請電腦中心協助！"
    Else
        MsgBox Err.Description
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearchOLD_Click()
'Dim StrSQLa As String
'Dim strCheckWay As String
''Add by Amy 2013/12/04
'Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
'Dim strSwhSQL1 As String, strSwhSQL2 As String
'Dim strSubSQL1 As String, strSubSQL2 As String
'Dim strNo As String, Str01 As String
'
'bolPrint = False '先設定無對造
'StrToPrint = ""
''end 2013/12/04
'
'lngCounterI = 0
'Dim s As Integer
'Dim strFields As String 'Added by Lydia 2017/02/14 設定關聯代號欄位
'
''潛在客戶編號
'If Option2(0).Value = True Then
'   If Len(Trim(Text1)) = 0 Then
'      s = MsgBox("條件不可空白", , "輸入條件錯誤")
'      Text1.SetFocus
'      Exit Sub
'   End If
'End If
''潛在客戶/聯絡人名稱
'If Option2(1).Value = True Then
'   If Len(Trim(Text2)) = 0 Then
'      s = MsgBox("條件不可空白", , "輸入條件錯誤")
'      Text2.SetFocus
'      Exit Sub
'   End If
'End If
''E-Mail
'If Option2(2).Value = True Then
'   If Len(Trim(Text9)) = 0 Then
'      s = MsgBox("條件不可空白", , "輸入條件錯誤")
'      Text9.SetFocus
'      Exit Sub
'   End If
'End If
'
''開發者
'If Option2(3).Value = True Then
'   If Len(Trim(Text3)) = 0 Then
'      s = MsgBox("條件不可空白", , "輸入條件錯誤")
'      Text3.SetFocus
'   End If
'End If
'
''開發日期
'If Option2(4).Value = True Then
'   If Len(Trim(Text4)) = 0 And Len(Trim(Text5)) = 0 Then
'      s = MsgBox("條件不可空白", , "輸入條件錯誤")
'      Text4.SetFocus
'   End If
'End If
'
'ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/24 清除查詢印表記錄檔欄位
'Screen.MousePointer = vbHourglass
'GrdDataList.Clear
'GrdDataList.Rows = 2
'SetDataListWidth
'StrSQLa = ""
'
'strFields = ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明 " 'Added by Lydia 2017/02/14
'
''若為國內智權人員或國內工程師, 不可查代理人資料
''Modify By Sindy 2011/01/04 取消
''If bolFNation = False Then
''   StrSQLa = " And FA01<'Y' "
''End If
'
''潛在客戶編號
''Modify by Amy 2013/10/30 讀取Fagent及Customer的狀態欄時，先檢查FA103或CU142，有值顯示 處理情形的內容，無值才抓原狀態欄位
''Modify by Amy 2013/09/30 trim掉空白去檢查:編號,名稱,E-Mail
'If Option2(0).Value = True Then
'   pub_QL05 = pub_QL05 & ";" & Option2(0).Caption & Trim(Text1) 'Add By Sindy 2010/12/24
'   'Modify by Amy 2013/12/04 +智權人/申請國家/總收文號/案件性質/收文日
'   If UCase(Left(Trim(Text1), 1)) = "R" Then
'      'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'      'Modified by Lydia 2017/02/14 + strfields
'      'Modify By Sindy 2021/6/28 + '1','廠商','2','事務所' => 國外潛在客戶類別
'      strSql = "SELECT ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,NVL(PCU08,DECODE(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM POTCUSTOMER,NATION,Staff WHERE PCU09=NA01(+) AND PCU01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' and substr(LTrim(PCU38),1,5)=ST01(+) "
'      'Add By Sindy 2011/10/11
'      'Modified by Lydia 2017/02/14 + strfields
'      strSql = strSql & " union all SELECT ' ' AS V ,PoC01||PoC02||Decode(PoC02,'0','','＊') AS 編號,NVL(PoC03,DECODE(PoC23,null,PoC27,RTRIM(PoC23||' '||PoC24||' '||PoC25||' '||PoC26))) AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,' ' AS 類別,PoC14 AS 狀態,PoC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM POTCUSTOMER1,NATION,Staff WHERE PoC04=NA01(+) AND PoC01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' and poc13=ST01(+) "
'      'end 2020/03/16
'   Else
'      'Modified by Lydia 2017/02/14 + strfields
'      strSql = "SELECT ' ' AS V ,CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●','') AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM CUSTOMER,NATION,Staff WHERE CU10=NA01(+) AND CU01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' AND CU13=ST01(+) "
'      'Modified by Lydia 2017/02/14 + strfields
'      strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as 名稱,NA03 AS 國籍,'' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM fagent,nation where fa10=na01(+) and fa01='" & Left(GetNewFagent(Trim(Text1)), 8) & "' " & StrSQLa
'      'Add By Sindy 2012/3/21
'      'Modified by Lydia 2017/02/14 + strfields
'      strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NVL(NT02,DECODE(NT03,null,NT07,NT03||' '||NT04||' '||NT05||' '||NT06)) as 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM notagent,nation,Staff where nt08=na01(+) and nt01='" & IIf(Len(Trim(Text1)) >= 3, Trim(Text1), Right("000" & Trim(Text1), 3)) & "' AND nt18=ST01(+) "
'   End If
''潛在客戶/聯絡人名稱
'ElseIf Option2(1).Value = True Then
'     pub_QL05 = pub_QL05 & ";" & Option2(1).Caption & Trim(Text2) 'Add By Sindy 2010/12/24
'     '以編號或名稱
'     '模糊比對
'     If Option3(0).Value = False Then
'        pub_QL05 = pub_QL05 & ";" & Option3(1).Caption 'Add By Sindy 2010/12/24
'        strCheckWay = ">0"
'     '字首比對
'     Else
'        pub_QL05 = pub_QL05 & ";" & Option3(0).Caption 'Add By Sindy 2010/12/24
'        strCheckWay = "=1"
'     End If
'     'Add by Amy 2013/12/04
'     strTp(3) = ChgSQL(UCase(Trim(Text2)))
'     '對造
'     strSQL1 = " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
'     strSQL2 = " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
'     StrSQL3 = " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
'     StrSQL4 = " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
'     strSQL5 = " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
'     'end 2013/12/04
''Modify by Amy 2013/12/04 拿掉中英日 +查對造 +智權人,申請國家,總收文號,案件性質,收文日 欄位
''     '中文
''     If Option1(0).Value = True Then
''         pub_QL05 = pub_QL05 & ";" & "以" & Option1(0).Caption & "查詢" 'Add By Sindy 2010/12/24
''         strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●','') AS 編號,CU04 AS 名稱,NA03 AS 國籍,' ' AS 類別,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,CU79 AS 備註 FROM CUSTOMER,NATION, (Select Distinct CU01 As A1 From Customer Where instr(CU04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 "
''         strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,fa04 as 名稱,NA03 AS 國籍,' ' AS 類別,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, FA29 AS 備註 from fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(fa04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
''         '查潛在客戶
''         strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,PCU08 AS 名稱,NA03 AS 國籍,decode(PCU11,'1','廠商','2','事務所','') AS 類別,PCU39 AS 狀態,PCU40 AS 備註 from potcustomer,nation, (Select Distinct pcu01 As A1 From potcustomer Where instr(pcu08,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A where pcu09=na01(+) and pcu01=A.A1"
''         'Add By Sindy 98/03/19
''         '查國內潛在客戶
''         strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,POC03 AS 名稱,NA03 AS 國籍,' ' AS 類別,POC14 AS 狀態,POC15 AS 備註 from potcustomer1,nation, (Select Distinct poc01 As A1 From potcustomer1 Where instr(poc03,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A where poc04=na01(+) and poc01=A.A1"
''         '98/03/19 End
''         '查聯絡人
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,' ' AS 類別,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,CUSTOMER,NATION WHERE CU10=NA01(+) AND  CU01(+)=PCC01 AND CU02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,decode(PCU11,'1','廠商','2','事務所','') AS 類別,PCU39 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,potcustomer,nation where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,' ' AS 類別,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, PCC13 AS 備註 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' " & StrSQLa
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,' ' AS 類別,POC14 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,potcustomer1,nation where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' "
''         'Add By Sindy 2012/3/21
''         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT02 as 名稱,NA03 AS 國籍,' ' AS 類別,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註 from notagent,nation, (Select Distinct NT01 As A1 From notagent Where instr(nt02,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where nt08=na01(+) and NT01=A.A1"
''
''     '英文
''     ElseIf Option1(1).Value = True Then
''         pub_QL05 = pub_QL05 & ";" & "以" & Option1(1).Caption & "查詢" 'Add By Sindy 2010/12/24
''         strSql = "SELECT ' ' AS V,CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●','') AS 編號,cu05||' '||cu88||' '||cu89||' '||cu90 AS 名稱,NA03 AS 國籍,' ' AS 類別,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,CU79 AS 備註 FROM CUSTOMER,NATION, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 "
''         strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,fa05||' '||fa63||' '||fa64||' '||fa65 as 名稱,NA03 AS 國籍,' ' AS 類別,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, FA29 AS 備註 from fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(upper(fa05||' '||fa63||' '||fa64||' '||fa65),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
''         '查潛在客戶
''         strSql = strSql & " union all SELECT ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06) AS 名稱,NA03 AS 國籍,' ' AS 類別,PCU39 AS 狀態,PCU40 AS 備註 from potcustomer,nation, (Select Distinct pcu01 As A1 From potcustomer Where instr(upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A where pcu09=na01(+) and pcu01=A.A1"
''         'Add By Sindy 2010/02/12
''         strSql = strSql & " union all SELECT ' ' AS V ,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,RTRIM(POC23||' '||POC24||' '||POC25||' '||POC26) AS 名稱,NA03 AS 國籍,' ' AS 類別,POC14 AS 狀態,POC15 AS 備註 from potcustomer1,nation, (Select Distinct poc01 As A1 From potcustomer1 Where instr(upper(poc23||' '||poc24||' '||poc25||' '||poc26),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A where poc04=na01(+) and poc01=A.A1"
''         '2010/02/12 End
''         '查聯絡人
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,' ' AS 類別,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & ") A,CUSTOMER,NATION WHERE CU10=NA01(+)  AND CU01(+)=PCC01 AND CU02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,decode(PCU11,'1','廠商','2','事務所','') AS 類別,PCU39 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & ") A,potcustomer,nation where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,' ' AS 類別,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, PCC13 AS 備註 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' " & StrSQLa
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,' ' AS 類別,POC14 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & ") A,potcustomer1,nation where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' "
''         'Add By Sindy 2012/3/21
''         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,nt03||' '||nt04||' '||nt05||' '||nt06 as 名稱,NA03 AS 國籍,' ' AS 類別,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註 from notagent,nation, (Select Distinct NT01 As A1 From notagent Where instr(upper(nt03||' '||nt04||' '||nt05||' '||nt06),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A where nt08=na01(+) and NT01=A.A1"
''
''     '日文
''     ElseIf Option1(2).Value = True Then
''         pub_QL05 = pub_QL05 & ";" & "以" & Option1(2).Caption & "查詢" 'Add By Sindy 2010/12/24
''         strSql = "SELECT ' ' AS V,CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●','') AS 編號,CU06 AS 名稱,NA03 AS 國籍,' ' AS 類別,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,CU79 AS 備註 FROM CUSTOMER,NATION, (Select Distinct CU01 As A1 From Customer Where instr(CU06,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1"
''         strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,fa06 as 名稱,NA03 AS 國籍,' ' AS 類別,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, FA29 AS 備註 from fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(fa06,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
''         '查潛在客戶
''         strSql = strSql & " union all SELECT ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,PCU07 AS 名稱,NA03 AS 國籍,decode(PCU11,'1','廠商','2','事務所','') AS 類別,PCU39 AS 狀態,PCU40 AS 備註 from potcustomer,nation, (Select Distinct pcu01 As A1 From potcustomer Where instr(pCU07,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where pcu09=na01(+) and pcu01=A.A1"
''         'Add By Sindy 2010/02/12
''         strSql = strSql & " union all SELECT ' ' AS V ,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,poc27 AS 名稱,NA03 AS 國籍,' ' AS 類別,POC14 AS 狀態,POC15 AS 備註 from potcustomer1,nation, (Select Distinct poc01 As A1 From potcustomer1 Where instr(poc27,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where poc04=na01(+) and poc01=A.A1"
''         '2010/02/12 End
''         '查聯絡人
''         strSql = strSql & " union all select ' ' as V,CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','') AS 編號,CU06 AS 名稱,NA03 AS 國籍,' ' AS 類別,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,CU79 AS 備註 FROM CUSTOMER,NATION,STAFF, (Select Distinct pcc01 As A1 From potcustcont Where instr(pcc04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 "
''         strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,PCU07 AS 名稱,NA03 AS 國籍,decode(PCU11,'1','廠商','2','事務所','') AS 類別,PCU39 AS 狀態,PCU40 AS 備註 from potcustomer,nation, (Select Distinct pcc01 As A1 From potcustcont Where instr(pcc04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A where pcu09=na01(+) and pcu01=A.A1"
''         strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,fa06 as 名稱,NA03 AS 國籍,' ' AS 類別,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, FA29 AS 備註 from fagent,nation, (Select Distinct pcc01 As A1 From potcustcont Where instr(pcc04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,' ' AS 類別,Decode(CU142,null,CU80,Decode(CU142,'A','同意抵帳中',Decode(CU142,'B','宣告破產','帳款處理中'))) AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+)  AND CU01(+)=PCC01 AND CU02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,decode(PCU11,'1','廠商','2','事務所','') AS 類別,PCU39 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,potcustomer,nation where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' "
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,' ' AS 類別,Decode(FA103,null,FA69,Decode(FA103,'A','同意抵帳中',Decode(FA103,'B','宣告破產','帳款處理中'))) AS 狀態, PCC13 AS 備註 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' " & StrSQLa
''         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,' ' AS 類別,POC14 AS 狀態,PCC13 AS 備註 from (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,potcustomer1,nation where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' "
''         'Add By Sindy 2012/3/21
''         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,nt07 as 名稱,NA03 AS 國籍,' ' AS 類別,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註 from notagent,nation, (Select Distinct NT01 As A1 From notagent Where instr(nt07,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where nt08=na01(+) and NT01=A.A1"
''     End If
'
'      'Modify by Amy 2014/02/25 對造由下搬上來改語法存至暫存檔
'            'Modified by Lydia 2019/12/26
'            'cnnConnection.Execute "Delete From R100102_1 Where ID='" & strUserNum & "' "
'            cnnConnection.Execute "Delete From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' "
'
''Modified by Lydia 2019/12/26 改成共用模組Pub_ProcR100102_1
''           '對造(中)
''            strSubSQL1 = " And InStr(Upper(CP40),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP50),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP40>' ' "
''            strSwhSQL2 = " CP50>' ' "
''
''      'Modify by Amy 2015/03/27 拿掉對造案件編號符號,+客戶端平台帳號資料
''            '商標
''            strSql = "Insert Into R100102_1 (r021001,r021002,r021003,r021004,r021005,r021006,r021007,r021008,r021009,r021010,r021011,r021012,r021013,r021014,r021015,r021016,r021017,r021018,ID) " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP40 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+)" & strSQL1 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP50 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+)" & strSQL1 & strSubSQL2
''            '專利
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP40 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+)" & strSQL2 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP50 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+)" & strSQL2 & strSubSQL2
''            '法務
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP40 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+)" & StrSQL3 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP50 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家 ,CP09 as 總收文號,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+)" & StrSQL3 & strSubSQL2
''            '顧問
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP40 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+)" & StrSQL4 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP50 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+)" & StrSQL4 & strSubSQL2
''            '服務
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP40 as 名稱,' ' as 智權人,'1' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+)" & strSQL5 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP50 as 名稱,' ' as 智權人,'2' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+)" & strSQL5 & strSubSQL2
''
''            '對造(英)
''            strSubSQL1 = " And InStr(Upper(CP41),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP51),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP41>' ' "
''            strSwhSQL2 = " CP51>' ' "
''            '商標
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP41 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+)" & strSQL1 & strSubSQL1
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP51 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+)" & strSQL1 & strSubSQL2
''            '專利
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP41 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+)" & strSQL2 & strSubSQL1
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP51 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+)" & strSQL2 & strSubSQL2
''            '法務
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP41 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+)" & StrSQL3 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP51 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+)" & StrSQL3 & strSubSQL2
''            '顧問
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP41 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+)" & StrSQL4 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP51 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+)" & StrSQL4 & strSubSQL2
''            '服務
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP41 as 名稱,' ' as 智權人,'1' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+)" & strSQL5 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP51 as 名稱,' ' as 智權人,'2' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,NVL(C1.CU04,C1.CU06)) AS 申請人1,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,NVL(C2.CU04,C2.CU06)) AS 申請人2,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,NVL(C3.CU04,C3.CU06)) AS 申請人3," & _
''                          "NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,NVL(C4.CU04,C4.CU06)) AS 申請人4,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,NVL(C5.CU04,C5.CU06)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+)" & strSQL5 & strSubSQL2
''
''            '對造(日)
''            strSubSQL1 = " And InStr(Upper(CP42),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP52),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP42>' ' "
''            strSwhSQL2 = " CP52>' ' "
''            '商標
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP42 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+)" & strSQL1 & strSubSQL1
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, CP52 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+)" & _
''                          " and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+)" & _
''                          " and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+)" & strSQL1 & strSubSQL2
''            '專利
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP42 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+)" & strSQL2 & strSubSQL1
''            strSql = strSql & " Union " & _
''                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP52 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
''                          " and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
''                          " and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+)" & strSQL2 & strSubSQL2
''            '法務
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP42 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+)" & StrSQL3 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP52 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
''                          " and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
''                          " and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+)" & StrSQL3 & strSubSQL2
''            '顧問
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP42 as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+)" & StrSQL4 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP52 as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
''                          " and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
''                          " and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+)" & StrSQL4 & strSubSQL2
''            '服務
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP42 as 名稱,' ' as 智權人,'1' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+)" & strSQL5 & strSubSQL1
''            strSql = strSql & " Union " & _
''                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號,CP52 as 名稱,' ' as 智權人,'2' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,cp05 as 收文日," & _
''                          "NVL(C1.CU06,NVL(C1.CU04,C1.CU05||C1.CU88||C1.CU89||C1.CU90)) AS 申請人1,NVL(C2.CU06,NVL(C2.CU04,C2.CU05||C2.CU88||C2.CU89||C2.CU90)) AS 申請人2,NVL(C3.CU06,NVL(C3.CU04,C3.CU05||C3.CU88||C3.CU89||C3.CU90)) AS 申請人3," & _
''                          "NVL(C4.CU06,NVL(C4.CU04,C4.CU05||C4.CU88||C4.CU89||C4.CU90)) AS 申請人4,NVL(C5.CU06,NVL(C5.CU04,C5.CU05||C5.CU88||C5.CU89||C5.CU90)) AS 申請人5,CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & strUserNum & "' " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
''                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
''                          " and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
''                          " and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+)" & strSQL5 & strSubSQL2
''      cnnConnection.Execute strSql
''
''      '刪除對造與申請人相同資料
''      strSql = "Delete From R100102_1 Where ID='" & strUserNum & "' And (ltrim(rtrim(R021002))=ltrim(rtrim(R021008)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021009)) " & _
''                  "Or ltrim(rtrim(R021002))=ltrim(rtrim(R021010)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021011)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021012))) "
''      cnnConnection.Execute strSql
''      'end 2014/02/25
''
''      'Add by Amy 2014/03/17 將所有商標案InStr(R021014,'T')且案件性質為1202(核准通知)者狀態改為 其他相關人
''      'Modify by Amy 2015/12/03 增加商標案(CFC/S) 案件性質202(申請意見書)及303(延期)者 狀態改為 其他相關人
''      strSql = "Update R100102_1 Set R021004='2' Where (InStr(R021014,'T')>0 or R021014='CFC' or R021014='S') And (R021018='1202' or R021018='202' or R021018='303')"
''      cnnConnection.Execute strSql
''      'end 2014/03/17
''      'Add by Amy 2015/12/03 所有專利案件性質404(延期) 者狀態改為 其他相關人
''      strSql = "Update R100102_1 Set R021004='2' Where (InStr(R021014,'P')>0 or R021014='FG') And R021018='404' "
''      cnnConnection.Execute strSql
''      'end 2015/12/03
'       Call Pub_ProcR100102_1(strUserNum & "@" & Me.Name, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5, strTp(3), strCheckWay)
''end 2019/12/26
'
'      '查customer 客戶 檔
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號,CU04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM CUSTOMER,NATION,Staff, (Select Distinct CU01 As A1 From Customer Where instr(CU04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號,cu05||' '||cu88||' '||cu89||' '||cu90 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM CUSTOMER,NATION,Staff, (Select Distinct CU01 As A1 From Customer Where instr(upper(cu05||' '||cu88||' '||cu89||' '||cu90),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●','') AS 編號,CU06 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM CUSTOMER,NATION,Staff, (Select Distinct CU01 As A1 From Customer Where instr(CU06,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+) "
'
'      '查Fagent 代理人 檔
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(fa77,'Y','$','') AS 編號,fa04 as 名稱,NA03 AS 國籍,' ' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(fa04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(fa77,'Y','$','') AS 編號,fa05||' '||fa63||' '||fa64||' '||fa65 as 名稱,NA03 AS 國籍,' ' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(upper(fa05||' '||fa63||' '||fa64||' '||fa65),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(fa77,'Y','$','') AS 編號,fa06 as 名稱,NA03 AS 國籍,' ' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM fagent,nation, (Select Distinct FA01 As A1 From Fagent Where instr(fa06,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where fa10=na01(+) AND FA01=A.A1 " & StrSQLa
'
'       'Modify by Amy 2015/04/15 客戶端平台帳號資料
'       'Modified by Lydia 2017/02/14 + strfields
'       'Modify By Sindy 2021/3/25 '' as 案件性質, => CW03 as 案件性質,
'       strSql = strSql & " union all Select ' ' as V,'平台'||CW01 AS 編號, CW12 AS 名稱,'平台' AS 國籍,' ' AS 智權人員,' ' AS 類別,Nvl(CW19,'') AS 狀態,'' AS 備註,' ' as 申請國家,'' as 總收文號,CW03 as 案件性質,CW01 as 收文日" & strFields & " From CustWeb Where InStr(Upper(CW12),'" & strTp(3) & "') " & strCheckWay
'
'      '查potcustomer 國外潛在客戶 檔
'         'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'         'Modified by Lydia 2017/02/14 + strfields
'         'Modify By Sindy 2021/6/28 + '1','廠商','2','事務所' => 國外潛在客戶類別
'         strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,PCU08 AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,Decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM potcustomer,nation,Staff, (Select Distinct pcu01 As A1 From potcustomer Where instr(pcu08,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06) AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,' ' AS 類別,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM potcustomer,nation,Staff, (Select Distinct pcu01 As A1 From potcustomer Where instr(upper(pcu03||' '||pcu04||' '||pcu05||' '||pcu06),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V ,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,PCU07 AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,Decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM potcustomer,nation,Staff, (Select Distinct pcu01 As A1 From potcustomer Where instr(pCU07,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where pcu09=na01(+) and pcu01=A.A1 and substr(LTrim(PCU38),1,5)=ST01(+) "
'
'      '查potcustomer1 國內潛在客戶 檔
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,POC03 AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,' ' AS 類別,POC14 AS 狀態,POC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM potcustomer1,nation,Staff, (Select Distinct poc01 As A1 From potcustomer1 Where instr(poc03,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V ,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,RTRIM(POC23||' '||POC24||' '||POC25||' '||POC26) AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,' ' AS 類別,POC14 AS 狀態,POC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM potcustomer1,nation,Staff, (Select Distinct poc01 As A1 From potcustomer1 Where instr(upper(poc23||' '||poc24||' '||poc25||' '||poc26),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V ,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,poc27 AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,' ' AS 類別,POC14 AS 狀態,POC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM potcustomer1,nation,Staff, (Select Distinct poc01 As A1 From potcustomer1 Where instr(poc27,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where poc04=na01(+) and poc01=A.A1 and poc13=ST01(+) "
'         'end 2020/03/16
'
'      '查NotAgent 不得代理案件之客戶或代理人 檔
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,NT02 as 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM notagent,nation,Staff, (Select Distinct NT01 As A1 From notagent Where instr(nt02,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where nt08=na01(+) and NT01=A.A1 AND nt18=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,nt03||' '||nt04||' '||nt05||' '||nt06 as 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM notagent,nation,Staff, (Select Distinct NT01 As A1 From notagent Where instr(upper(nt03||' '||nt04||' '||nt05||' '||nt06),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & " ) A where nt08=na01(+) and NT01=A.A1 AND nt18=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,NT01||Decode(NT21,null,'⊕','') AS 編號,nt07 as 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM notagent,nation,Staff, (Select Distinct NT01 As A1 From notagent Where instr(nt07,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & " ) A where nt08=na01(+) and NT01=A.A1 AND nt18=ST01(+) "
'
'      '查聯絡人(中文)
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,CUSTOMER,NATION,Staff WHERE CU10=NA01(+) AND  CU01(+)=PCC01 AND CU02='0' AND CU13=ST01(+) "
'         'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'         'Modified by Lydia 2017/02/14 + strfields
'         'Modify By Sindy 2021/6/28 + '1','廠商','2','事務所' => 國外潛在客戶類別
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,Decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,potcustomer,nation,Staff where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,' ' AS 類別,POC14 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,potcustomer1,nation,Staff where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
'         'end 2020/03/16
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC05 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(pcc05,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' " & StrSQLa
'
'      '查聯絡人(英文)
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & ") A,CUSTOMER,NATION,Staff WHERE CU10=NA01(+)  AND CU01(+)=PCC01 AND CU02='0' AND CU13=ST01(+) "
'         'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'         'Modified by Lydia 2017/02/14 + strfields
'         'Modify By Sindy 2021/6/28 + '1','廠商','2','事務所' => 國外潛在客戶類別
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,Decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & ") A,potcustomer,nation,Staff where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,' ' AS 類別,POC14 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & ") A,potcustomer1,nation,Staff where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
'         'end 2020/03/16
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC03 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(upper(pcc03),'" & UCase(ChgSQL(Trim(Text2))) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' " & StrSQLa
'
'      '查聯絡人(日文)
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,CUSTOMER,NATION,STAFF WHERE CU10=NA01(+)  AND CU01(+)=PCC01 AND CU02='0' AND CU13=ST01(+) "
'         'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'         'Modified by Lydia 2017/02/14 + strfields
'         'Modify By Sindy 2021/6/28 + '1','廠商','2','事務所' => 國外潛在客戶類別
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,Decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,potcustomer,nation,Staff where pcu09=na01(+)  AND PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,' ' AS 類別,POC14 AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,potcustomer1,nation,Staff where poc04=na01(+)  AND POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
'         'end 2020/03/16
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all select ' ' as V,PCC01||'0-'||PCC02 AS 編號,PCC04 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM (Select * From potcustcont Where instr(PCC04,'" & ChgSQL(Trim(Text2)) & "')" & strCheckWay & ") A,fagent,nation where fa10=na01(+) AND FA01(+)=PCC01 AND FA02='0' " & StrSQLa
'
'      '抓暫存檔對造 Add by Amy 2014/02/25
'         'Modified by Lydia 2017/02/14 + strfields
'         'Modified by Lydia 2019/12/26 +@+Me.name
'         'Modify by Amy 2020/09/04 +all 因查 金杜 應出現2筆,中/日文都有輸
'         strSql = strSql & " union all select ' ' as V,R021001 AS 編號,R021002 AS 名稱,'' AS 國籍,'' AS 智權人員,' ' AS 類別,Decode(R021004,'1','對造','2','其他相關人','') AS 狀態,'' AS 備註,'' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "'  And R021004<3 "
'
'      'end 2015/03/27
'
'      'Mark by Amy 2014/02/25 往上搬
''           '對造(中)
''            strSubSQL1 = " And InStr(Upper(CP40),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP50),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP40>' ' "
''            strSwhSQL2 = " CP50>' ' "
''            '商標
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP40 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
''                        " Union  Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP50 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
''            '專利
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP40 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
''                        " Union  Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP50 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
''            '法務
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP40 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP50 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家 ,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
''            '顧問
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP40 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP50 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
''            '服務
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP40 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP50 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
''
''            '對造(英)
''            strSubSQL1 = " And InStr(Upper(CP41),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP51),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP41>' ' "
''            strSwhSQL2 = " CP51>' ' "
''            '商標
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP41 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
''                         " Union Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP51 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
''            '專利
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP41 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
''                         " Union Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP51 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
''            '法務
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP41 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP51 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
''            '顧問
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP41 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP51 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
''            '服務
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP41 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP51 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
''
''            '對造(日)
''            strSubSQL1 = " And InStr(Upper(CP42),'" & strTp(3) & "') " & strCheckWay
''            strSubSQL2 = " And InStr(Upper(CP52),'" & strTp(3) & "') " & strCheckWay
''            strSwhSQL1 = " CP42>' ' "
''            strSwhSQL2 = " CP52>' ' "
''            '商標
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP42 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL1 & _
''                         " Union Select ' ' as V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as 編號, CP52 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),TradeMark,CasePropertyMap Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & strSubSQL2
''            '專利
''            strSql = strSql & " Union " & _
''                         "Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP42 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL1 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL1 & _
''                         " Union Select ' ' as V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as 編號,CP52 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                         "From (Select * From CaseProgress Where " & strSwhSQL2 & "),Patent,CasePropertyMap Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & strSubSQL2
''            '法務
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP42 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(LC34,'')),null,'','●') AS 編號,CP52 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(LC15,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),LawCase,CasePropertyMap Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & strSubSQL2
''            '顧問
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP42 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(HC19,'')),null,'','●') as 編號,CP52 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,' ' as 申請國家,CP09 as 總收文號,NVL(decode(CPM03,null,CPM04,CPM03),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),HireCase,CasePropertyMap Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & strSubSQL2
''            '服務
''            strSql = strSql & " Union " & _
''                        "Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP42 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'對造' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL1 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL1 & _
''                        " Union Select ' ' as V,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(SP61,'')),null,'','●') as 編號,CP52 as 名稱,' ' as 國籍,' ' as 智權人,' ' AS 類別,'其他相關人' as 狀態,' ' as 備註,SP09 as申請國家,CP09 as 總收文號,NVL(decode(SP09,'000',CPM03,CPM04),CP10) as 案件性質,Nvl(To_Char(cp05-19110000),'') as 收文日 " & _
''                        "From (Select * From CaseProgress Where " & strSwhSQL2 & "),ServicePractice,CasePropertyMap Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & strSubSQL2
'      'end 2014/02/25
'
'     ' Add By Sindy 98/02/13
'     '開拓客戶資料
'     If Check1.Value = 1 Then
'        'Modify by Amy 2013/12/04 +智權人員/申請國家/總收文號/案件性質/收文日
'        'Modify by Amy 2013/09/30 原只檢查ecd11,ecd12卻顯示ecd03,ecd04
'        'strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,ecd03||' '||ecd04 AS 名稱,NA03 AS 國籍,eca02 AS 類別,ecd15 AS 狀態,ecd16 AS 備註 from expandcusdetail,expandcusattr,nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') As A1 From expandcusdetail Where instr(ecd11,'" & ChgSQL(Trim(Text2)) & "') " & strCheckWay & " or instr(ecd12,'" & ChgSQL(Trim(Text2)) & "') " & strCheckWay & ") A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,ecd03||' '||ecd04 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,eca02 AS 類別,'投法開拓'||decode(ecd15,null,null,'-'||ecd15) AS 狀態,ecd16 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " from expandcusdetail,expandcusattr,nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') As A1 From expandcusdetail Where instr(Upper(ecd03),'" & ChgSQL(UCase(Trim(Text2))) & "') " & strCheckWay & " or instr(Upper(ecd04),'" & ChgSQL(UCase(Trim(Text2))) & "') " & strCheckWay & ") A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
'         'Modified by Lydia 2017/02/14 + strfields
'         strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,ecd11||' '||ecd12 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,eca02 AS 類別,'投法開拓'||decode(ecd15,null,null,'-'||ecd15) AS 狀態,ecd16 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " from expandcusdetail,expandcusattr,nation,(Select Distinct nvl(ecd01,'')||nvl(ecd02,'') As A1 From expandcusdetail Where instr(Upper(ecd11),'" & ChgSQL(UCase(Trim(Text2))) & "') " & strCheckWay & " or instr(Upper(ecd12),'" & ChgSQL(UCase(Trim(Text2))) & "') " & strCheckWay & ") A Where ecd10=na01(+) and ecd02=eca01(+) and nvl(ecd01,'')||nvl(ecd02,'')=A.A1 "
'     End If
'     ' 98/02/13 End
'
''E-Mail
'ElseIf Option2(2).Value = True Then
'     pub_QL05 = pub_QL05 & ";" & Option2(2).Caption & Trim(Text9) 'Add By Sindy 2010/12/24
'     'Modify by Amy 2013/12/04 +智權人員/申請國家/總收文號/案件性質/收文日
'     'Modified by Lydia 2017/02/14 + strfields
'     strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●','') AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM CUSTOMER,NATION,PotCustomer,Staff  Where (instr(NLS_Upper(CU20),'" & UCase(ChgSQL(Trim(Text9))) & "')>0 or  instr(NLS_Upper(CU115),'" & UCase(ChgSQL(Trim(Text9))) & "')>0 or instr(NLS_Upper(CU116),'" & UCase(ChgSQL(Trim(Text9))) & "')>0  or instr(NLS_Upper(CU117),'" & UCase(ChgSQL(Trim(Text9))) & "')>0 or instr(NLS_Upper(CU118),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 )  and CU10=NA01(+)  and pcu01(+)=cu01 AND CU13=ST01(+) "
'     'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'     'Modify By Sindy 2021/6/28 + '1','廠商','2','事務所' => 國外潛在客戶類別
'     strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,NVL(PCU08,DECODE(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM potcustomer,nation,Staff  Where (instr(NLS_Upper(pcu18),'" & UCase(ChgSQL(Trim(Text9))) & "') >0 ) and pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
'     'Add By Sindy 98/03/19
'     '查國內潛在客戶
'     'Modified by Lydia 2017/02/14 + strfields
'     strSql = strSql & " union all SELECT ' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,NVL(PoC03,DECODE(PoC23,null,PoC27,RTRIM(PoC23||' '||PoC24||' '||PoC25||' '||PoC26))) AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,' ' AS 類別,POC14 AS 狀態,POC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM potcustomer1,nation,Staff  Where (instr(NLS_Upper(poc09),'" & UCase(ChgSQL(Trim(Text9))) & "') >0 ) and poc04=na01(+) and poc13=ST01(+) "
'     '98/03/19 End
'     'end 2020/03/16
'
'     'Modified by Lydia 2017/02/14 + strfields
'     'Modified by Lydia 2018/07/20 +FA105 財務信箱(CF)
'     'strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as 名稱,NA03 AS 國籍,' ' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM fagent,nation Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(Text9))) & "')> 0 or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(Text9))) & "')> 0 or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(Text9))) & "')> 0 or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 )   and  fa10=na01(+)" & StrSQLa
'     strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as 名稱,NA03 AS 國籍,' ' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM fagent,nation " & _
'                              " Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(Trim(Text9))) & "')> 0 or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(Trim(Text9))) & "')> 0 or instr(NLS_Upper(fa105),'" & UCase(ChgSQL(Trim(Text9))) & "')> 0 or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(Trim(Text9))) & "')> 0 or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 )   and  fa10=na01(+)" & StrSQLa
'     'Modified by Lydia 2017/02/14 + strfields
'     'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'     'Modify By Sindy 2021/6/28 + '1','廠商','2','事務所' => 國外潛在客戶類別
'     strSql = strSql & " union all SELECT ' ' AS V,PCC01||'0-'||PCC02 AS 編號,NVL(PCC05,NVL(PCC03,PCC04)) AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,' ' AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM PotCustCont,Potcustomer,Nation,Staff Where (instr(NLS_Upper(PCC08),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 ) and pcc01=pcu01(+) and pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
'
'     ' Add By Sindy 98/02/13
'     '開拓客戶資料
'     If Check1.Value = 1 Then
'      'Modify by Amy 2013/12/04 +智權人員/申請國家/總收文號/案件性質/收文日
'      'Modify by Amy 2013/09/30 原:ecd15 AS 狀態
'      'Modified by Lydia 2017/02/14 + strfields
'      strSql = strSql & " union all SELECT ' ' AS V,ecd02||'-'||LPAD(ecd01,6,'0') AS 編號,ecd03||' '||ecd04 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,eca02 AS 類別,'投法開拓'||decode(ecd15,null,null,'-'||ecd15) AS 狀態,ecd16 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM expandcusdetail, expandcusattr ,nation Where (instr(NLS_Upper(ecd13),'" & UCase(ChgSQL(Trim(Text9))) & "') > 0 ) and ecd10=na01(+) and ecd02=eca01(+) "
'     End If
'     ' 98/02/13 End
''開發者
'ElseIf Option2(3).Value = True Then
'     pub_QL05 = pub_QL05 & ";" & Option2(3).Caption & Text3 & Label10 'Add By Sindy 2010/12/24
'     'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'     'Modified by Lydia 2017/02/14 + strfields
'     'Modify By Sindy 2021/6/28 + '1','廠商','2','事務所' => 國外潛在客戶類別
'     strSql = "SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,NVL(PCU08,DECODE(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM PotCustomer,Nation,Staff Where (instr(PCU38,'" & ChgSQL(Text3) & "') > 0)   AND  PCU09=NA01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
'     'Add By Sindy 2011/10/11
'     'Modified by Lydia 2017/02/14 + strfields
'     strSql = strSql & " union all SELECT ' ' AS V,PoC01||PoC02||Decode(PoC02,'0','','＊') AS 編號,NVL(PoC03,DECODE(PoC23,null,PoC27,RTRIM(PoC23||' '||PoC24||' '||PoC25||' '||PoC26))) AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,' ' AS 類別,PoC14 AS 狀態,PoC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM PotCustomer1,Nation,Staff Where (instr(PoC13,'" & ChgSQL(Text3) & "') > 0)   AND  PoC04=NA01(+) and poc13=ST01(+) "
'     'end 2020/03/16
'     'Modified by Lydia 2017/02/14 + strfields
'     strSql = strSql & " union all SELECT ' ' AS V,PCC01||'0-'||PCC02 AS 編號,NVL(PCC05,NVL(PCC03,PCC04)) AS 名稱,' ' AS 國籍,' ' AS 智權人員,' ' AS 類別,' ' AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM PotCustCont Where instr(PCC12,'" & ChgSQL(Text3) & "') > 0  AND SUBSTR(PCC01,1,1)='R' "
'     'Modified by Lydia 2017/02/14 + strfields
'     strSql = strSql & "union all SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●','') AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,'  ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM CUSTOMER,NATION,Staff  Where (instr(CU129,'" & ChgSQL(Text3) & "')>0 )  and CU10=NA01(+) AND CU13=ST01(+) "
'     'Modified by Lydia 2017/02/14 + strfields
'     'Modified by Lydia 2022/01/22 debug  fagent,nation,PotCustomer=> fagent,nation
'     strSql = strSql & " union all select ' ' as V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as 名稱,NA03 AS 國籍,' ' AS 智權人員,' ' AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM fagent,nation Where (instr(fa94,'" & ChgSQL(Text3) & "')> 0  )   and  fa10=na01(+)   " & StrSQLa
'
''開發日期
'ElseIf Option2(4).Value = True Then
'   pub_QL05 = pub_QL05 & ";" & Option2(4).Caption & Text4 & "-" & Text5 'Add By Sindy 2010/12/24
'   'Modified by Lydia 2017/02/14 + strfields
'   strSql = "SELECT ' ' as V,CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●','') AS 編號,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,' ' AS 類別,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM CUSTOMER,NATION,Staff Where cu14 >='" & ChangeTStringToWString(Text4) & "' and cu14 <='" & ChangeTStringToWString(Text5) & "' and CU10=NA01(+) AND CU13=ST01(+) "
'   'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'   'Modified by Lydia 2017/02/14 + strfields
'   'Modify By Sindy 2021/6/28 + '1','廠商','2','事務所' => 國外潛在客戶類別
'   strSql = strSql & " union all SELECT ' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,NVL(PCU08,DECODE(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))) AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,decode(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM PotCustomer,nation,Staff  Where PCU37 >='" & ChangeTStringToWString(Text4) & "' and PCU37 <='" & ChangeTStringToWString(Text5) & "' and pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
'   'Add By Sindy 2011/10/11
'   'Modified by Lydia 2017/02/14 + strfields
'   strSql = strSql & " union all SELECT ' ' AS V,PoC01||PoC02||Decode(PoC02,'0','','＊') AS 編號,NVL(PoC03,DECODE(PoC23,null,PoC27,RTRIM(PoC23||' '||PoC24||' '||PoC25||' '||PoC26))) AS 名稱,NA03 AS 國籍,poc13 AS 智權人員,' ' AS 類別,PoC14 AS 狀態,PoC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM PotCustomer1,nation,Staff  Where PoC12 >='" & ChangeTStringToWString(Text4) & "' and PoC12 <='" & ChangeTStringToWString(Text5) & "' and poc04=na01(+) and poc13=ST01(+) "
'   'end 2020/03/16
'   'Modified by Lydia 2017/02/14 + strfields
'   strSql = strSql & " union all SELECT ' ' AS V,FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$','') AS 編號,NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)) as 名稱,NA03 AS 國籍,' ' AS 智權人員,NULL AS 類別,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM fagent,nation Where fa11 >='" & ChangeTStringToWString(Text4) & "' and  fa11 <='" & ChangeTStringToWString(Text5) & "'  and fa10=na01(+)" & StrSQLa
'   'Modified by Lydia 2017/02/14 + strfields
'   'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
'   'Modify By Sindy 2021/6/28 + '1','廠商','2','事務所' => 國外潛在客戶類別
'   strSql = strSql & " union all SELECT ' ' AS V,PCC01||'0-'||PCC02 AS 編號,NVL(PCC05,NVL(PCC03,PCC04)) AS 名稱,NA03 AS 國籍,pcu38 AS 智權人員,decode(PCU11," & 國外潛在客戶類別 & ",PCU11)  AS 類別,' ' AS 狀態,PCC13 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日" & strFields & " FROM PotCustCont,PotCustomer,Nation,Staff Where PCC11 >='" & ChangeTStringToWString(Text4) & "' and  PCC11 <='" & ChangeTStringToWString(Text5) & "'  and PCC01=PCU01(+) and pcu09=na01(+) and substr(LTrim(PCU38),1,5)=ST01(+) "
'End If
''2008/12/4 ADD BY SONIA
''Modify by Amy 2019/09/17 加待活化客戶
'If Option2(1).Value = True Then
'   'Modify by Amy 2014/01/15 +編號排
'   strSql = "select X.*,Decode(Ocu01,null, '',NVL(Ocu03,0)) as Ocu03 from (" & strSql & ") X, OldCustomer Where substr(編號,1,8)= ocu01(+) order by upper(名稱),編號 "
'Else
'   strSql = "select X.*,Decode(Ocu01,null, '',NVL(Ocu03,0)) as Ocu03 from (" & strSql & ") X, OldCustomer Where substr(編號,1,8)= ocu01(+) order by 編號 "
'End If
''end 2019/09/17
''2008/12/4 end
'
'If Check1.Value = 1 Then
'   pub_QL05 = pub_QL05 & ";" & Check1.Caption 'Add By Sindy 2010/12/24
'End If
'If Trim(Text6) <> "" Or Trim(Text7) <> "" Then
'   pub_QL05 = pub_QL05 & ";" & Left(Label5, 5) & Text6 & "-" & Text7 'Add By Sindy 2010/12/24
'End If
''If Trim(Text10) <> "" Then
''   pub_QL05 = pub_QL05 & ";" & Label7 & Text10 'Add By Sindy 2010/12/24
''End If
'If Trim(cboSort.Text) <> "" Then
'   pub_QL05 = pub_QL05 & ";" & Label7 & cboSort.Text 'Add By Sindy 2019/3/13
'End If
'
'CheckOC
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'
'If adoRecordset.RecordCount <> 0 Then
'    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/24
'    If Not cmdOK(0).Enabled Then cmdOK(0).Enabled = True
'    If Not cmdOK(1).Enabled Then cmdOK(1).Enabled = True
'    'If Not cmdOK(3).Enabled Then cmdOK(3).Enabled = True '2008/12/9 CANCEL BY SONIA
'    If Not cmdOK(4).Enabled Then cmdOK(4).Enabled = True
'    Set GrdDataList.Recordset = adoRecordset
'Else
'    InsertQueryLog (0) 'Add By Sindy 2010/12/24
'    'Add by Amy 2013/12/04 +畫面訊息開放可列印
'    Pub_Can_Copy_Pic = True
'    ShowNoData
'    Pub_Can_Copy_Pic = False
'    'end 2013/12/04
'    cmdOK(0).Enabled = False
'    cmdOK(1).Enabled = False
'    'cmdOK(3).Enabled = True '2008/12/9 CANCEL BY SONIA
'    cmdOK(4).Enabled = False
'    GrdDataList.Clear
'End If
'Me.GrdDataList.Visible = False 'Add by Amy 2013/12/04
'SetDataListWidth
'CheckOC
'
'With Me.GrdDataList
'    If .Rows > 0 Then 'Add by Amy 2013/12/04 +if
'        For i = 1 To .Rows - 1
'            .row = i
'            .col = 1
'            .CellForeColor = &H0 'Add by Amy 2019/08/29 字黑色
'            'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'            If .TextMatrix(.row, 16) = "0" And Right(.TextMatrix(.row, 1), 1) <> "＊" Then
'                For j = 0 To .Cols - 1
'                    '呆帳
'                    If Right(.Text, 1) = "$" And j = 1 Then
'                        GrdDataList.CellBackColor = &HFF& '紅色
'                    '活化客戶
'                    Else
'                        .col = j
'                        .CellBackColor = vbYellow
'                    End If
'                Next
'            ElseIf Right(.Text, 1) = "$" Then
'                .CellBackColor = &HFF&
'            'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底粉字
'            'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'            ElseIf (Left(.Text, 1) = "Y" Or Left(.Text, 1) = "X" Or Left(.Text, 1) = "R") _
'                And (.TextMatrix(i, 6) = "解散" Or .TextMatrix(i, 6) = "廢止" Or .TextMatrix(i, 6) = "撤銷" Or .TextMatrix(i, 6) = "死亡") Then
'                    For j = 0 To .Cols - 1
'                        .col = j
'                        .CellBackColor = &H0 '黑色
'                        .CellForeColor = &HFF00FF '粉紅色  'Modfiy by Amy 2019/08/29 原:ForeColor
'                    Next j
'            'Add By Sindy 2012/3/21
'            ElseIf Right(.Text, 1) = "⊕" Or .TextMatrix(i, 6) = "對造" Or .TextMatrix(i, 6) = "其他相關人" Then
'                    'Modify by Amy 2013/12/04 對造重抓智權人資料
'                    If Me.GrdDataList.TextMatrix(i, 6) = "對造" Or .TextMatrix(i, 6) = "其他相關人" Then
'                        bolPrint = True '有對造資料
'                        strNo = Pub_RplStr(.TextMatrix(i, 1))
'                        StrToPrint = strNo & ","
'                        Str01 = SystemNumber(strNo, 1)
'                        Select Case Str01
'                            Case "FCP", "FG"
'                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCPSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "FCL", "LIN"
'                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCLSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "FCT"
'                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                            Case "S"
'                                If .TextMatrix(i, 8) = "000" Then
'                                    .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetFCTSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                                Else
'                                    .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                                End If
'                            Case Else
'                                .TextMatrix(i, 4) = GetPrjSalesNM(PUB_GetAKindSalesNo(Str01, SystemNumber(strNo, 2), SystemNumber(strNo, 3), SystemNumber(strNo, 4)))
'                        End Select
'                        .TextMatrix(i, 10) = .TextMatrix(i, 10) & PUB_GetRelateCasePropertyName(.TextMatrix(i, 9), "1")
'                        'Add by Amy 2014/02/25 更新智權人員至暫存檔
'                        strExc(0) = "Update R100102_1 Set R021003='" & .TextMatrix(i, 4) & "' Where R021014='" & Str01 & "' And R021015='" & SystemNumber(strNo, 2) & "' And R021016='" & SystemNumber(strNo, 3) & "' And R021017='" & SystemNumber(strNo, 4) & "' "
'                        cnnConnection.Execute strExc(0)
'                        'end 2014/02/25
'                    End If
'                    'end 2013/12/04
'                    If Right(.Text, 1) = "⊕" Or .TextMatrix(i, 6) = "對造" Then
'                        For j = 0 To .Cols - 1
'                            .col = j
'                            .CellBackColor = &H8080FF
'                        Next j
'                    End If
'            '2012/3/21 End
'
'            'Add By Sindy 2021/3/25 針對CW03=7.媒介平台，在查詢系統顯示結果為橘色
'            ElseIf Left(.TextMatrix(i, 1), 1) = "平" And .TextMatrix(i, 10) = "7" Then
'                .CellBackColor = &H80FF& '橘色
'            '2021/3/25 END
'            End If
'
'            'Add by Amy 2020/03/16 國內外潛在客戶 智權人員欄需重抓資料(可能多筆)
'            If Left(.Text, 1) = "R" Then
'                .TextMatrix(i, 4) = GetDevelopP(.TextMatrix(i, 4))
'            End If
'        Next i
'   End If 'end 2013/12/04
'End With
'
''若只有一筆資料, 則直接設定為點選此筆資料
'With Me.GrdDataList
'   If .Rows = 2 Then
'      .row = 1
'      .col = 1
'      If .Text <> "" Then
'        .Visible = False
'        .row = 1
'        .col = 0
'        .Text = "V"
'        For i = 0 To .Cols - 1
'            'Modify By Sindy 2012/3/21 old:If i <> 1 Then
'            If i <> 1 And (i = 2 And Right(.TextMatrix(1, 1), 1) = "⊕") = False Then
'                .col = i
'                .CellBackColor = &HFFC0C0
'            End If
'        Next i
'        'Add by Amy 2020/10/15 勾選時判斷有往來記錄,往來記錄鈕變色
'        Call ChkContactRecordBT(.TextMatrix(1, 0), Left(.TextMatrix(1, 1), 8))
'        .Visible = True
'      End If
'   End If
'End With
''Add by Amy 2013/12/04
'Me.GrdDataList.Visible = True
'If bolPrint Then
'    cmdOK(5).Enabled = True
'Else
'    cmdOK(5).Enabled = False
'End If
''end 2013/12/04
'Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/13 還原此Form的查詢條件記錄
   If Text1.Visible = True Then
      Text1.SetFocus 'Add By Sindy 2012/4/11
   End If
End Sub

Private Sub Form_Load()
      
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   GetField 'Add by Amy 2023/03/08
   Label2(4).Caption = "　　紫底為風險警示" 'Modify by Amy 2024/01/31 +風險檢查對象,拿掉風險警示啟用日
   'end 2023/12/13
   cmdok(0).Enabled = True
   cmdok(1).Enabled = False
   
   'cmdOK(3).Enabled = True  '2008/12/9 CANCEL BY SONIA
   cmdok(4).Enabled = True
   
   Option2(0).Value = True
   Option1(0).Enabled = False
   Option1(1).Enabled = False
   Option1(2).Enabled = False
   
   ' Add By Sindy 98/02/16
   'MODIFY BY SONIA 2015/5/20 因P31及F31人員併入L02,但內外法不開放權限,故改用員工等級控制
   'If Pub_StrUserSt03 = "F31" Or Pub_StrUserSt03 = "F41" Then
   If Pub_strUserST05 >= "51" And Pub_strUserST05 <= "55" Then
      Check1.Value = 1
   Else
      Check1.Value = 0
   End If
   ' 98/02/16 End
   
   bolToEndByNick = False
   m_bolPrintRight = IsUserHasRightOfFunction("frm140407", strPrint, False)
   Me.cmdok(4).Enabled = m_bolPrintRight
   cmdState = -1
   
   'Added by Lydia 2018/05/24 改由啟用日控制
   If strSrvDate(1) >= 國外部關聯企業啟用日 Then cmdok(1).Caption = "關聯企業(&R)"
   
   m_blnColOrderAsc = True 'Add by Amy 2020/09/04
   'AddCombo cboSort
   'lstSort.Clear
   SetcboSort 'Add By Sindy 2019/3/11
   
   Label10.Caption = "" 'Added by Lydia 2022/01/12
   
End Sub

'Add By Sindy 2019/3/11 往來類別
Private Sub SetcboSort()
   cboSort.Clear
   strSql = "select ac02,ac03 from allcode where ac01='11'" & _
            " order by ac02 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   cboSort.AddItem ""
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         cboSort.AddItem RsTemp.Fields("ac02") & " " & RsTemp.Fields("ac03")
         RsTemp.MoveNext
      Loop
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/9/12 還原此Form的查詢條件記錄 (多筆查詢有影響)
   Set frm140407 = Nothing
End Sub

Sub StrMenu(StrToGrid)
'已潛在客戶查詢之資料庫
'Modify By Sindy 98/03/19
'strSQL = "SELECT CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●',''),NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),NA03,NULL,CU80,CU79 FROM CUSTOMER,NATION WHERE CU10=NA01(+) AND CU01>='" & Left(StrToGrid, 6) & "00' AND CU01<='" & Left(StrToGrid, 6) & "zz' "
'strSQL = strSQL & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,NVL(PCU03||' '||PCU04||' '||PCU05||' '||PCU06,PCU07)),NA03,PCU11,PCU39,PCU40 FROM PotCustomer,Nation WHERE PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz'   AND NA01(+)=PCU09"
'strSQL = strSQL & " union  SELECT FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$',''),DECODE(FA10,'013',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)),NA03,NULL,FA29,FA63 FROM FAGENT,NATION WHERE FA01>='" & Left(StrToGrid, 6) & "00' AND FA01<='" & Left(StrToGrid, 6) & "zz' AND fa10=NA01(+)  "
strSql = "SELECT CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●',''),NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),NA03,NULL,CU80,CU79 FROM CUSTOMER,NATION WHERE CU10=NA01(+) AND CU01>='" & Left(StrToGrid, 6) & "00' AND CU01<='" & Left(StrToGrid, 6) & "zz' "
strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,DECODE(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))),NA03,PCU11,PCU39,PCU40 FROM PotCustomer,Nation WHERE PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz'   AND NA01(+)=PCU09"
strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','＊'),POC03,NA03,'',POC14,POC15 FROM PotCustomer1,Nation WHERE POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz'   AND NA01(+)=POC04"
strSql = strSql & " union  SELECT FA01||FA02||Decode(FA02,'0','','＊')||decode(fa77,'Y','$',''),DECODE(FA10,'013',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)),NA03,NULL,FA29,FA63 FROM FAGENT,NATION WHERE FA01>='" & Left(StrToGrid, 6) & "00' AND FA01<='" & Left(StrToGrid, 6) & "zz' AND fa10=NA01(+)  "
'傳入R1時找出相關的X
strSql = strSql & " union  SELECT CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●',''),NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),NA03,NULL,CU80,CU79 " & _
                                                 "From CUSTOMER, PotCustomer1, Nation " & _
                                            "WHERE CU10=NA01(+) " & _
                                                 "AND POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                 "AND CU01>=(substr(POC16,1,6)||'00') AND CU01<=(substr(POC16,1,6)||'zz') " & _
                                                 "AND POC16 is not null "
'找出R1的關係企業
strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','＊'),POC03,NA03,'',POC14,POC15 " & _
                                                 "From PotCustomer1, Nation " & _
                                             "WHERE NA01(+)=POC04 " & _
                                                  "AND POC16>='" & Left(StrToGrid, 6) & "00' AND POC16<='" & Left(StrToGrid, 6) & "zz' " & _
                                                  "AND POC16 is not null "
'傳入R1時找出相關的R
strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,DECODE(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))),NA03,PCU11,PCU39,PCU40 " & _
                                                 "From PotCustomer, Nation, PotCustomer1 " & _
                                            "WHERE NA01(+)=PCU09 " & _
                                                 "AND POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                 "AND PCU47>=(substr(POC16,1,6)||'00') AND PCU47<=(substr(POC16,1,6)||'zz') " & _
                                                 "AND POC16 is not null AND PCU47 is not null "
'98/03/19 End
'Add By Sindy 2009/06/24
'傳入R時找出相關的X
strSql = strSql & " union  SELECT CU01||CU02||Decode(CU02,'0','','＊')||decode(cu111,'Y','$','')||decode(cu121,'Y','●',''),NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),NA03,NULL,CU80,CU79 " & _
                                                 "From CUSTOMER, PotCustomer, Nation " & _
                                            "WHERE CU10=NA01(+) " & _
                                                 "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                 "AND CU01>=(substr(PCU47,1,6)||'00') AND CU01<=(substr(PCU47,1,6)||'zz') " & _
                                                 "AND PCU47 is not null "
'傳入R時找出相關的Y
strSql = strSql & " union  SELECT FA01||FA02||Decode(FA02,'0','','＊'),NVL(fa04,DECODE(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),NA03,'',FA69,FA29 " & _
                                                 "From Fagent, PotCustomer, Nation " & _
                                             "WHERE NA01(+)=FA10 " & _
                                                  "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                                                  "AND FA01>=(substr(PCU47,1,6)||'00') AND FA01<=(substr(PCU47,1,6)||'zz') " & _
                                                  "AND PCU47 is not null "
'找出R的關係企業
strSql = strSql & " union  SELECT PCU01||PCU02||Decode(PCU02,'0','','＊'),NVL(PCU08,DECODE(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))),NA03,PCU11,PCU39,PCU40 " & _
                                                 "From PotCustomer, Nation " & _
                                            "WHERE NA01(+)=PCU09 " & _
                                                 "AND PCU47>='" & Left(StrToGrid, 6) & "00' AND PCU47<='" & Left(StrToGrid, 6) & "zz' " & _
                                                 "AND PCU47 is not null "
'傳入R時找出相關的R1
strSql = strSql & " union  SELECT POC01||POC02||Decode(POC02,'0','','＊'),POC03,NA03,'',POC14,POC15 " & _
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
    Do While adoRecordset.EOF = False
    strSql = "INSERT INTO R140407 values ('"
    If Not IsNull(adoRecordset.Fields(0)) Then
        strSql = strSql + ChgSQL(CheckStr(adoRecordset.Fields(0))) + "','"
    Else
        strSql = strSql + "','"
    End If
    If Not IsNull(adoRecordset.Fields(1)) Then
        strSql = strSql + ChgSQL(CheckStr(adoRecordset.Fields(1))) + "','"
    Else
        strSql = strSql + "','"
    End If
    If Not IsNull(adoRecordset.Fields(2)) Then
        strSql = strSql + ChgSQL(CheckStr(adoRecordset.Fields(2))) + "','"
    Else
        strSql = strSql + "','"
    End If
    
    If Not IsNull(adoRecordset.Fields(3)) Then
        strSql = strSql + ChgSQL(CheckStr(adoRecordset.Fields(3))) + "','"
    Else
        strSql = strSql + "','"
    End If
    
   strSql = strSql & "" & strUserNum & "')"
    
    cnnConnection.Execute strSql
    adoRecordset.MoveNext
    Loop
Else
    ShowNoData
    Screen.MousePointer = vbDefault
    Exit Sub
End If
CheckOC
End Sub

Sub StrMenu1()
    Dim k As Integer 'Add by Amy 2019/10/05
    
'Modfiy by Amy 2015/05/12 智權人員抓st02
'Modify by Amy 2013/12/10 +智權人員/申請國家/總收文號/案件性質/收文日
'Added by Lydia 2018/05/24 改由啟用日控制
If strSrvDate(1) < 國外部關聯企業啟用日 Then
    'Modify by Amy +4個''->關聯編號/名稱/關係/說明 避免加欄位困難
    strSql = "SELECT '' AS V,R06001 AS 編號,R06002 AS 名稱,R06003 AS 國籍,ST02 as 智權人員,decode(R06004,'1','廠商','')||decode(R06004,'2','事務所','') AS 類別,CU80 AS 狀態,CU79 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日,'' as 關聯編號,'' as 關聯名稱,'' as 關聯關係,'' as 關聯說明  FROM R140407,CUSTOMER,Staff where id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' AND  SUBSTR(R06001,1,8)=CU01(+) AND SUBSTR(R06001,9,1)=CU02(+) And CU13=ST01(+) "
    'Modify by Amy 2020/03/16 智權人員 原:st02 ,因開發人員可能多人
    'Modify by Amy 2019/10/05 原:Union All 把All  拿掉 ex:X29973 有兩筆(一筆為更名)->兩筆勾選->按「關係企業」->不應出現四筆
    strSql = strSql & " UNION SELECT '' AS V,R06001 AS 編號,R06002 AS 名稱,R06003 AS 國籍,pcu38 as 智權人員,decode(R06004,'1','廠商','')||decode(R06004,'2','事務所','') AS 類別,PCU39 AS 狀態,PCU40 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日,'' as 關聯編號,'' as 關聯名稱,'' as 關聯關係,'' as 關聯說明  FROM R140407,POTCUSTOMER,Staff where id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='R' AND SUBSTR(R06001,1,8)=PCU01 AND SUBSTR(R06001,9,1)=PCU02 and substr(LTrim(PCU38),1,5)=ST01(+) "
    'Add By Sindy 98/03/19
    strSql = strSql & " UNION SELECT '' AS V,R06001 AS 編號,R06002 AS 名稱,R06003 AS 國籍,poc13 as 智權人員,' 'AS 類別,POC14 AS 狀態,POC15 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日,'' as 關聯編號,'' as 關聯名稱,'' as 關聯關係,'' as 關聯說明  FROM R140407,POTCUSTOMER1,Staff where id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='R' AND SUBSTR(R06001,1,8)=POC01 AND SUBSTR(R06001,9,1)=POC02 and POC13=ST01(+) "
    'end 2015/05/12
    'end 2020/03/16
    strSql = strSql & " UNION SELECT '' AS V,R06001 AS 編號,R06002 AS 名稱,R06003 AS 國籍,'' as 智權人員,decode(R06004,'1','廠商','')||decode(R06004,'2','事務所','')AS 類別,FA69 AS 狀態,FA29 AS 備註,' ' as 申請國家,'' as 總收文號,'' as 案件性質,'' as 收文日,'' as 關聯編號,'' as 關聯名稱,'' as 關聯關係,'' as 關聯說明  FROM R140407,FAGENT where id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='Y' AND SUBSTR(R06001,1,8)=FA01(+) AND SUBSTR(R06001,9,1)=FA02(+) "
Else
    'Modified by Lydia 2017/02/14 抓關聯企業改成模組,暫存R100114_1
    strSql = "SELECT '' AS V,R11402 AS 編號,R11403 AS 名稱,NVL(NA03,R11405) AS 國籍 ,ST02 AS 智權人員,R11404 AS 類別,R11407 AS 狀態,R11408 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日," & _
             "R11409 AS 關聯編號,DECODE(SUBSTR(R11409,1,1),'X',DECODE(SIGN(INSTR('000,001,002,003,004,005,006,007,008,009,013,020',C1.CU10)),0,DECODE(C1.CU05,NULL,NVL(C1.CU04,C1.CU06),C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90),NVL(C1.CU04,DECODE(C1.CU05,NULL,C1.CU06,C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)))," & _
             "'Y',DECODE(SIGN(INSTR('000,001,002,003,004,005,006,007,008,009,013,020',F1.FA10)),0,DECODE(F1.FA05,NULL,NVL(F1.FA04,F1.FA06),F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65),NVL(F1.FA04,DECODE(F1.FA05,NULL,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65))),R11409) AS 關聯名稱," & _
             "R11410 AS 關聯關係, R11411 AS 關聯說明 FROM R100114_1,STAFF,NATION,CUSTOMER C1,FAGENT F1 " & _
             "WHERE ID='" & strUserNum & "' AND FORMID='" & UCase(Me.Name) & "' AND R11406=ST01(+) AND R11405=NA01(+) " & _
             "AND SUBSTR(R11409,1,8)=C1.CU01(+) AND '0'=C1.CU02(+) AND SUBSTR(R11409,1,8)=F1.FA01(+) AND '0'=F1.FA02(+) "
    'end 2017/02/14
End If
'Modify by Amy 2023/08/23 更名OCU03=>待活化客戶; 增加ORGN欄位
strSql = "Select X.*,'' as ORGN, Decode(Ocu01,null, '',NVL(Ocu03,0)) as 待活化客戶 from (" & strSql & ") X, OldCustomer Where substr(編號,1,8)= ocu01(+) "
If strSrvDate(1) < 國外部關聯企業啟用日 Then
    strSql = strSql & " ORDER BY 編號"
Else
    strSql = strSql & " ORDER BY 編號, 關聯編號 "
End If
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then
    Set grdDataList.Recordset = adoRecordset
    'Modify by Amy 2023/08/30 原程式搬至SetDataListWidth
    SetDataListWidth (True)
End If
CheckOC

'Add by Amy 2019/10/05 +所有顏色顯示
   grdDataList.Visible = False
   If grdDataList.Rows > 0 Then
        For i = 1 To grdDataList.Rows - 1
            grdDataList.row = i
            grdDataList.col = 1
            grdDataList.CellForeColor = &H0   '字黑色 ex:查儀大會整個變黑
            'Modify by Amy 2023/08/30 變色改共用函數
            'Modify by Amy 2023/09/26 依狀態更新智權人員改為共用函數
            Call UpdQuerySales(Me.Name, grdDataList, strField)
            'end 2023/09/26
            Call SetMSGridColorQCus(0, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
            'end 2023/08/30
        Next i
   End If
   
   '若只有一筆資料 , 則直接設定為點選此筆資料
   'Modify by Amy 2023/08/30 原程式改為共用SetGridOneData,避免有沒改到的
   cmdok(4).BackColor = &H8000000F
   Call SetGridOneData
   'end 2023/08/30
   grdDataList.Visible = True
   'end 2019/10/05
End Sub

Private Sub GrdDataList_Click()
Dim strCopyTxt As String ' Add by Amy 2014/04/25 複製編號文字

grdDataList.row = grdDataList.MouseRow
'Modify by Amy 2014/04/25 +選到編號欄=複製
grdDataList.col = grdDataList.MouseCol
If grdDataList.col = 1 Then
    grdDataList.CellForeColor = &H0 '黑色
    'Modify by Amy 2020/09/04 不小心按到欄位名稱也會copy
    If grdDataList.row > 0 Then
        strCopyTxt = grdDataList.TextMatrix(grdDataList.row, grdDataList.col)
    End If
    If strCopyTxt <> "" Then
        '複製編號至剪貼簿
        Clipboard.SetText strCopyTxt
        grdDataList.CellBackColor = QBColor(7)
        MsgBox "編號已複製", , MsgText(21)
        
        '設回原本顏色
        'Modify by Amy 2023/08/30 改寫至共用函數
'        'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'         If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'            '呆帳
'            If Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) = "$" Then
'                grdDataList.CellBackColor = &HFF& '紅色
'            '活化客戶
'            Else
'                grdDataList.CellBackColor = vbYellow
'            End If
'        'Add by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'        'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'        ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'              And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'               Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                grdDataList.CellBackColor = &H0 '黑色
'                grdDataList.CellForeColor = &HFF00FF '粉紅色
'        ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'            grdDataList.CellBackColor = &H8080FF
'        Else
'            grdDataList.CellBackColor = QBColor(15)
'        End If
         Call SetMSGridColorQCus(2, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
         'end 2023/08/30
    End If
    Exit Sub
End If
'end 2014/04/25
   
grdDataList.Visible = False
grdDataList.col = 0
If grdDataList.row <> 0 Then
    If grdDataList.Text = "V" Then
         grdDataList.Text = ""
         'Modify by Amy 2023/08/30 改抓共用函數
'         'Add By Sindy 2012/3/21
'         grdDataList.col = 1
'         'Add by Amy 2019/09/17 活化客戶顯示整列黃,若有呆帳編號底為紅其他欄為黃
'         If grdDataList.TextMatrix(grdDataList.row, GetValue("待活化客戶")) = "0" And Right(grdDataList.TextMatrix(grdDataList.row, GetValue("編號")), 1) <> "＊" Then
'                 For i = 0 To grdDataList.Cols - 1
'                    '呆帳
'                    If Right(grdDataList.Text, 1) = "$" And i = 1 Then
'                        grdDataList.CellBackColor = &HFF& '紅色
'                    '活化客戶
'                    Else
'                        grdDataList.col = i
'                        grdDataList.CellBackColor = vbYellow
'                    End If
'                Next
'         'Modify by Amy 2019/08/28 +客戶狀態為 遷移不明/解散/廢止/撤銷/停業/死亡 顯示黑底
'         'Modify by Amy 2022/05/23 拿掉 遷移不明 及 停業
'         ElseIf (Left(grdDataList.Text, 1) = "Y" Or Left(grdDataList.Text, 1) = "X" Or Left(grdDataList.Text, 1) = "R") _
'            And (grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "解散" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "廢止" _
'            Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "撤銷" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "死亡") Then
'                For i = 0 To grdDataList.Cols - 1
'                    grdDataList.col = i
'                    grdDataList.CellBackColor = &H0 '黑色
'                    grdDataList.CellForeColor = &HFF00FF '粉紅色
'                Next i
'         ElseIf Right(grdDataList.Text, 1) = "⊕" Or grdDataList.TextMatrix(grdDataList.row, GetValue("狀態")) = "對造" Then
'            For i = 0 To grdDataList.Cols - 1
'               grdDataList.col = i
'               grdDataList.CellBackColor = &H8080FF
'            Next i
'         Else
'         '2012/3/21 End
'            For i = 0 To grdDataList.Cols - 1
'                If i <> 1 Then
'                    grdDataList.col = i
'                    grdDataList.CellBackColor = QBColor(15)
'                End If
'            Next i
'         End If
         Call SetMSGridColorQCus(0, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
    '勾選
    Else
         grdDataList.Text = "V"
         'Modify by Amy 2023/08/30 改抓共用函數
'         For i = 0 To grdDataList.Cols - 1
'            'Modify By Sindy 2012/3/21 old:If i <> 1 Then
'            If i <> 1 And (i = 2 And Right(grdDataList.TextMatrix(grdDataList.MouseRow, GetValue("編號")), 1) = "⊕") = False Then
'                grdDataList.col = i
'                grdDataList.CellBackColor = &HFFC0C0
'            End If
'         Next i
         Call SetMSGridColorQCus(1, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
    End If
    'Add by Amy 2020/10/15 勾選時判斷有往來記錄,往來記錄鈕變色
    'Modify by Amy 2023/08/30 bug-聯絡人也會有往來記錄,故拿掉編號只取8碼
    strExc(10) = grdDataList.TextMatrix(grdDataList.row, GetValue("編號"))
    If Left(strExc(10), 1) = "X" Or Left(strExc(10), 1) = "Y" Or Left(strExc(10), 1) = "R" Or Left(strExc(10), 2) = "平台" Then
      Call ChkContactRecordBT(grdDataList.TextMatrix(grdDataList.row, GetValue("V")), strExc(10))
    End If
End If
grdDataList.Visible = True

End Sub

'Add by Amy 2020/09/04 +排序
Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If grdDataList.MouseCol < 0 Or grdDataList.MouseRow < 0 Then Exit Sub
    
    grdDataList.col = grdDataList.MouseCol
    grdDataList.row = grdDataList.MouseRow
    If grdDataList.col = 2 Then grdDataList.col = 16 'Modify by Amy 2022/08/19 名稱以OrgN排
    If Me.grdDataList.row < 1 And Me.grdDataList.Text <> "V" Then
        If m_blnColOrderAsc = True Then
            Me.grdDataList.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
        Else
            Me.grdDataList.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
        End If
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 1 Then
    CloseIme
    Text2.SetFocus
Else
    OpenIme
    Text2.SetFocus
End If
End Sub

Private Sub Option2_Click(Index As Integer)
Select Case Index
   '潛在客戶編號
   Case 0
        If Option2(0).Value = True Then
           Option2(1).Value = False
           Option2(2).Value = False
           
           Option1(0).Enabled = False
           Option1(1).Enabled = False
           Option1(2).Enabled = False
           Option3(0).Enabled = False
           Option3(1).Enabled = False
           Text1.SetFocus 'Add By Sindy 2012/4/11
        End If
   '潛在客戶/聯絡人名稱
   Case 1
        If Option2(1).Value = True Then
           Option1(0).Enabled = True
           Option2(2).Value = False
           
           Option1(0).Value = True
           Option1(1).Enabled = True
           Option1(2).Enabled = True
           Option2(0).Value = False
           Option3(0).Enabled = True
           Option3(1).Enabled = True
           Option3(1).Value = True     '2012/3/28 ADD BY SONIA
           Text2.SetFocus 'Add By Sindy 2012/4/11
        End If
   'E-Mail
   Case 2
        If Option2(2).Value = True Then
           Option2(0).Value = False
           Option2(1).Value = False
           
           Option1(0).Enabled = False
           Option1(1).Enabled = False
           Option1(2).Enabled = False
           Option3(0).Enabled = False
           Option3(1).Enabled = False
           Text9.SetFocus 'Add By Sindy 2012/4/11
        End If
   '開發者
   Case 3
        If Option2(3).Value = True Then
           Option2(0).Value = False
           Option2(1).Value = False
           Option2(2).Value = False
           Option2(4).Value = False
           
           Option1(0).Enabled = False
           Option1(1).Enabled = False
           Option1(2).Enabled = False
           Option3(0).Enabled = False
           Option3(1).Enabled = False
           Text3.SetFocus 'Add By Sindy 2012/4/11
        End If
   '開發日期
   Case 4
        If Option2(4).Value = True Then
           Option2(0).Value = False
           Option2(1).Value = False
           Option2(2).Value = False
           Option2(3).Value = False
           
           Option1(0).Enabled = False
           Option1(1).Enabled = False
           Option1(2).Enabled = False
           Option3(0).Enabled = False
           Option3(1).Enabled = False
           Text4.SetFocus 'Add By Sindy 2012/4/11
        End If
   Case Else
End Select
End Sub
Private Sub Text1_GotFocus()
   
   Me.Option2(0).Value = True
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1)
   
   CloseIme
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(0).Value = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   Me.Option2(1).Value = True
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2)
   '英文
   'If Option1(1).Value = True Then  'Modify by Amy 2013/12/10 改判斷部門
   If Left(Pub_StrUserSt03, 1) = "F" Then
      CloseIme
   Else
      OpenIme
   End If
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(1).Value = True
End Sub
Private Sub Text2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   'Modify by Amy 2013/12/04 Mark掉
'   '英文
'   If Option1(1).Value = True Then
'      CloseIme
'   Else
'      OpenIme
'   End If
End Sub

'Mark by Amy 2022/08/11
'Private Sub Text3_Click()
'   If Text3 <> "" And GetStaffName(Text3) = "" Then
'      Label10.Caption = ""
'      MsgBox "開發人員不存在,請查核", vbCritical
'      TextInverse Text3
'      Exit Sub
'   Else
'      Label10.Caption = GetStaffName(Text3)
'   End If
'End Sub

Private Sub Text3_GotFocus()
   Me.Option2(3).Value = True 'Add By Sindy 2019/3/13
   Text3.SelStart = 0
   Text3.SelLength = Len(Text3)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.Option2(3).Value = True 'Add By Sindy 2019/3/13
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   Dim strSTN As String 'Add by Amy 2022/08/11
   
   Cancel = False
   'Modify by Amy 2022/08/11 原:GetStaffName(Text3) ex:68005  已離職會彈,但又查的到資料
   strSTN = GetPrjSalesNM(Text3)
   If Text3 <> "" And strSTN = "" Then
      Cancel = True
      Label10.Caption = ""
      MsgBox "開發人員不存在,請查核", vbCritical
      TextInverse Text3
      Exit Sub
   Else
      'Label10.Caption = GetStaffName(Text3)
      Label10.Caption = strSTN
   End If
   'end 2022/08/11
End Sub

Private Sub Text4_GotFocus()
Me.Option2(4).Value = True 'Add By Sindy 2019/3/13
Text4.SelStart = 0
Text4.SelLength = Len(Text4)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_LostFocus()
   If PUB_CheckKeyInDate(Me.Text4) = -1 Then
      Me.Text4.SetFocus
      Text4_GotFocus
      Exit Sub
   End If
End Sub

Private Sub Text4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.Option2(4).Value = True 'Add By Sindy 2019/3/13
End Sub

Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_LostFocus()
   If PUB_CheckKeyInDate(Me.Text5) = -1 Then
      Me.Text5.SetFocus
      TextInverse Text5
      Exit Sub
   End If
   
   If Not nickChgRan(Text4, Text5, "開發日期") Then
      Text4.SetFocus
      TextInverse Text4
      Exit Sub
   End If
End Sub

Private Sub Text6_GotFocus()
Text6.SelStart = 0
Text6.SelLength = Len(Text6)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_LostFocus()
   If PUB_CheckKeyInDate(Me.Text6) = -1 Then
      Me.Text6.SetFocus
      TextInverse Text6
      Exit Sub
   End If
End Sub

Private Sub Text7_GotFocus()
   Text7.SelStart = 0
   Text7.SelLength = Len(Text7)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_LostFocus()
   If PUB_CheckKeyInDate(Me.Text7) = -1 Then
      Me.Text7.SetFocus
      TextInverse Text7
      Exit Sub
   End If
   
   If Not nickChgRan(Text6, Text7, "往來日期") Then
      Text6.SetFocus
      TextInverse Text6
      Exit Sub
   End If
End Sub

Private Sub Text9_GotFocus()
   Me.Option2(2).Value = True
   Text9.SelStart = 0
   Text9.SelLength = Len(Text9)
End Sub

Private Sub Text9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option2(2).Value = True
End Sub

Private Function ChgPotCustomer(ByVal strTemp As String) As String
 On Error GoTo ErrHand
   If strTemp = "" Then GoTo ErrHand
   
   If Len(strTemp) = 9 Then
      ChgPotCustomer = "PCU01='" & Left(strTemp, 8) & "' AND PCU02='" & Right(strTemp, 1) & "'"
   Else
      ChgPotCustomer = "PCU01='" & strTemp & String(8 - Len(strTemp), "0") & "' AND PCU02='0'"
   End If
   Exit Function
ErrHand:
   ChgPotCustomer = "PCU01 IS NULL AND PCU02 IS NULL"
End Function

Private Function ComposeList(oList As ListBox, Optional p_iOpt As Integer = 0) As String
   Dim iPos As Integer, stItem As String
   strExc(1) = ""
   If oList.ListCount > 0 Then
      For intI = 0 To oList.ListCount - 1
         If p_iOpt = 0 Then
            iPos = InStr(oList.List(intI), Chr(1))
            If iPos > 0 Then
               stItem = Left(oList.List(intI), iPos - 1)
            Else
               stItem = oList.List(intI)
            End If
         Else
            stItem = Format(oList.ITEMDATA(intI), "00")
         End If
         If intI = 0 Then
            strExc(1) = stItem
         Else
            strExc(1) = strExc(1) & "," & stItem
         End If
      Next
   End If
   ComposeList = strExc(1)
End Function
Private Function AddList(oList As ListBox, oCombo As ComboBox, Optional p_iOpt As Integer = 0) As Boolean
   Dim idx As Integer, bFound As Boolean, stNewItem As String, iNewItemData As Integer
   Dim stSort As String, iPos As Integer
   
   If oCombo.Text = "" Then
      Exit Function
   End If
   
   If p_iOpt = 1 Then
      If oCombo.ListIndex = -1 Then
         MsgBox "聯絡人資料不存在！"
         Exit Function
      End If
   End If
      
   '若有控制字元時後面為說明文字不抓
   iPos = InStr(oCombo, Chr(1))
   If iPos > 0 Then
      stNewItem = Left(oCombo, iPos - 1)
   Else
      stNewItem = oCombo
   End If
   iNewItemData = oCombo.ITEMDATA(oCombo.ListIndex)
      
   If InStr(stNewItem, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請改用其他符號！", vbExclamation
      oCombo.SetFocus
      Exit Function
   End If

   If stNewItem <> "" Then
      For idx = 0 To oList.ListCount - 1
         If oList.List(idx) = stNewItem And oList.ITEMDATA(idx) = iNewItemData Then
            MsgBox "資料已存在！"
            AddList = False
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         oList.AddItem stNewItem, 0
         If p_iOpt <> 0 Then
            oList.ITEMDATA(0) = oCombo.ITEMDATA(oCombo.ListIndex)
         End If
         AddList = True
      End If
   End If
End Function

Private Function RemoveList(oList As ListBox) As Boolean
   Dim ii As Integer
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            RemoveList = True
            oList.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
   End If
End Function

'Private Sub AddCombo(oCombo As ComboBox)
'   With oCombo
'      .Clear
'      .AddItem "IP諮詢"
'      .AddItem "非IP法律諮詢"
'      .AddItem "詢價"
'      .AddItem "申請所需文件"
'      .AddItem "利益衝突"
'      .AddItem "互惠"
'      .AddItem "詢問IP侵害"
'      .AddItem "訪談" & Chr(1) & "(包括來所訪問、出國拜訪、國際會議)"
'      .AddItem "客戶特別指示" & Chr(1) & "(譬如:不要寄confirmation copy, 聯絡方式只限fax or e-mail, 付款方式只限credit card or cheque…)"
'   End With
'End Sub

'Mark by Amy 2023/09/20 改為共用函數
'Add by Amy 2014/02/25 '原Add by Amy 2013/12/04 PrintDataA4刪除不使用
Private Sub PrintDataA4_Temp()
'    Dim rsPrint As New ADODB.Recordset
'    Dim strPrint As String
'    Dim ii As Integer, jj As Integer
'On Error GoTo Checking
'    intCounter = 1: intRecord = 1: intPage = 1
'
'    Screen.MousePointer = vbHourglass
'    Printer.PaperSize = PUB_GetPaperSize(9) '設定紙張 A4
'    Printer.Orientation = 1 '直印
'    PrintHeadA4
'
'    Printer.FontBold = False
'    'Modify by Amy 2020/09/08 ID+表單
'    strPrint = "Select R021001,R021002,R021003,Decode(R021004,'1','對造','2','其他相關人',''),R021006,R021007,Nvl(To_Char(R021008-19110000),'') " & _
'                 "From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' Order by R021002,R021001"
'    intI = 1
'    Set rsPrint = ClsLawReadRstMsg(intI, strPrint)
'    If intI = 1 Then
'        rsPrint.MoveFirst
'        For ii = 0 To rsPrint.RecordCount - 1
'            If intRecord > 45 Then
'                intPage = intPage + 1
'                intRecord = 1
'                Printer.NewPage
'                intCounter = 1
'                PrintHeadA4
'                Printer.FontBold = False
'            End If
'            For jj = 0 To rsPrint.Fields.Count - 1
'                If jj = rsPrint.Fields.Count - 1 Then
'                    Printer.CurrentX = PLeft(jj + 1) - 300 - Printer.TextWidth(rsPrint.Fields(jj).Value) '最右邊
'                Else
'                    Printer.CurrentX = PLeft(jj)
'                End If
'                Printer.CurrentY = 300 + intCounter * 300
'
'                Select Case jj
'                    Case 0 '本所案號
'                        Printer.Print Pub_RplStr(rsPrint.Fields(jj).Value)
'                    Case 1 '名稱
'                        Printer.Print StrToStr(rsPrint.Fields(jj).Value, 10)
'                    Case 2, 3, 4 '智權人員,狀態,總收文號
'                        Printer.Print rsPrint.Fields(jj).Value
'                    Case 5 '案件性質
'                        Printer.Print StrToStr(rsPrint.Fields(jj).Value, 6)
'                    Case 6  '收文日
'                        Printer.Print ChangeTStringToTDateString(rsPrint.Fields(jj).Value)
'                    Case Else
'                End Select
'            Next jj
'            intCounter = intCounter + 1
'            intRecord = intRecord + 1
'            rsPrint.MoveNext
'        Next ii
'    End If
'    Printer.EndDoc
'    Screen.MousePointer = vbDefault
'
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'   Screen.MousePointer = vbDefault
End Sub
'end 2014/02/25

Private Sub PrintHeadA4()
'   If intPage = 1 Then
'        GetPleft
'        strTp(0) = Me.Caption
'        strTp(1) = ""
'
'        If Option3(0).Value = True Then
'            strTp(1) = strTp(1) & "(字首比對)"
'        ElseIf Option3(1).Value = True Then
'            strTp(1) = strTp(1) & "(模糊比對)"
'        End If
'   End If
'   strTp(2) = "名稱：" & strTp(3) & Space(6) & strTp(1)
'
'   Printer.FontSize = 17
'   Printer.FontBold = True
'   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTp(0)) / 2)
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print strTp(0)
'
'   Printer.FontSize = 12
'   intCounter = intCounter + 2
'   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTp(2)) / 2)
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print strTp(2)
'   'Printer.Line (Printer.ScaleWidth / 2 - ((Printer.TextWidth(strTp(2)) - Printer.TextWidth("名稱：")) / 2) + 300, Printer.CurrentY + 30)-(Printer.ScaleWidth / 2 + Printer.TextWidth(strTp(2)) / 2, Printer.CurrentY + 30)
'
'   intCounter = intCounter + 1
'   Printer.CurrentX = 0
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "操作人員：" & StaffQuery(strUserNum)
'   Printer.CurrentX = 8800
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "查詢日期：" & CFDate(ACDate(ServerDate))
''   intCounter = intCounter + 1
''   Printer.CurrentX = 12000
''   Printer.CurrentY = 300 + intCounter * 300
''   Printer.Print "頁次: " & intPage
'    intCounter = intCounter + 1
'    For kk = 1 To UBound(PLeft)
'        Printer.CurrentX = PLeft(kk - 1) + (PLeft(kk) - PLeft(kk - 1) - Printer.TextWidth(ColName(kk)) - 100) / 2
'        Printer.CurrentY = 300 + intCounter * 300
'        Printer.Print ColName(kk)
'        Printer.Line (PLeft(kk - 1), Printer.CurrentY)-(PLeft(kk) - 100, Printer.CurrentY)
'    Next kk
'    intCounter = intCounter + 1
End Sub

Private Sub GetPleft()
'   ReDim PLeft(0 To 7)
'   ReDim ColName(1 To 7)
'   PLeft(0) = 100
'   PLeft(1) = PLeft(0) + 2000: ColName(1) = "本所案號"
'   PLeft(2) = PLeft(1) + 2700: ColName(2) = "    名       稱    "
'   PLeft(3) = PLeft(2) + 1200: ColName(3) = "智權人員"
'   PLeft(4) = PLeft(3) + 1500: ColName(4) = " 狀  態 "
'   PLeft(5) = PLeft(4) + 1300: ColName(5) = "總收文號"
'   PLeft(6) = PLeft(5) + 1800: ColName(6) = "案件性質"
'   PLeft(7) = PLeft(6) + 1200: ColName(7) = "收文日"
End Sub
'end 2013/12/04
'end 2023/09/20 不使用

'Add by Amy 2020/10/15 勾選時判斷有往來記錄,往來記錄鈕變色
Private Sub ChkContactRecordBT(ByVal stChk As String, ByVal stKey As String)
    'Memo by Amy 2023/09/27  原2023/08/30 將按鈕鎖住,有資料才可按,User 按此鈕新增,故不鎖
    cmdok(4).BackColor = &H8000000F
    If stChk = "V" And PUB_ChkContactRecord(stKey) = True Then
        cmdok(4).BackColor = vbYellow
    End If
End Sub

'Add by Amy 2023/08/30 查詢只有一筆資料Grid顏色設定
Private Sub SetGridOneData()
     grdDataList.Visible = False
     With Me.grdDataList
        If .Rows = 2 Then
            .row = 1
            .col = 1
            If .Text <> "" Then
                .row = 1
                .col = 0
                .Text = "V"
                Call SetMSGridColorQCus(1, Me.Name, grdDataList, strField, IIf(Check3.Value = vbChecked, True, False))
                '勾選時判斷有往來記錄,往來記錄鈕變色
                strExc(10) = grdDataList.TextMatrix(grdDataList.row, GetValue("編號"))
                If Left(strExc(10), 1) = "X" Or Left(strExc(10), 1) = "Y" Or Left(strExc(10), 1) = "R" Or Left(strExc(10), 2) = "平台" Then
                  Call ChkContactRecordBT(grdDataList.TextMatrix(grdDataList.row, GetValue("V")), strExc(10))
                End If
            End If
        End If
   End With
   grdDataList.Visible = True
End Sub

