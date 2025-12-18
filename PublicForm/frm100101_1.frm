VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100101_1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "案件資料及進度查詢"
   ClientHeight    =   5890
   ClientLeft      =   110
   ClientTop       =   1000
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5890
   ScaleWidth      =   9360
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   1740
      MaxLength       =   20
      TabIndex        =   41
      Top             =   2130
      Width           =   3165
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Client Matter ID："
      Height          =   180
      Index           =   5
      Left            =   60
      TabIndex        =   40
      Top             =   2190
      Width           =   1635
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "IDS清單"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   8
      Left            =   8460
      Style           =   1  '圖片外觀
      TabIndex        =   37
      Top             =   420
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "相似案"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   7
      Left            =   7680
      Style           =   1  '圖片外觀
      TabIndex        =   36
      Top             =   420
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "原始檔"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   345
      Index           =   14
      Left            =   5400
      Style           =   1  '圖片外觀
      TabIndex        =   32
      Top             =   420
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "國內A4名條"
      Height          =   345
      Index           =   6
      Left            =   6180
      Style           =   1  '圖片外觀
      TabIndex        =   33
      Top             =   420
      Width           =   1470
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   345
      Index           =   13
      Left            =   4620
      Style           =   1  '圖片外觀
      TabIndex        =   31
      Top             =   420
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "顯示代表圖(&P)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   345
      Index           =   5
      Left            =   7185
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   45
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "專利相關案件(&R)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   345
      Index           =   4
      Left            =   5715
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   45
      Width           =   1470
   End
   Begin VB.CheckBox chk 
      Caption         =   "所有系統類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   60
      TabIndex        =   25
      Top             =   120
      Width           =   1740
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   1605
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1140
      Width           =   2292
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1485
      Width           =   2532
   End
   Begin VB.OptionButton Option1 
      Caption         =   "彼所案號："
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   20
      Top             =   1830
      Width           =   1245
   End
   Begin VB.OptionButton Option1 
      Caption         =   "審定號數/證書號數："
      Height          =   180
      Index           =   3
      Left            =   60
      TabIndex        =   19
      Top             =   1515
      Width           =   2025
   End
   Begin VB.OptionButton Option1 
      Caption         =   "客戶案件案號："
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   18
      Top             =   1170
      Width           =   1680
   End
   Begin VB.OptionButton Option1 
      Caption         =   "申請案號："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   17
      Top             =   855
      Width           =   1305
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   16
      Top             =   540
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   0
      Left            =   2340
      MaxLength       =   6
      TabIndex        =   1
      Top             =   480
      Width           =   1212
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   1
      Left            =   3570
      MaxLength       =   1
      TabIndex        =   2
      Top             =   480
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Index           =   2
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   3
      Top             =   480
      Width           =   492
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3075
      Left            =   60
      TabIndex        =   22
      Top             =   2790
      Width           =   9270
      _ExtentX        =   16334
      _ExtentY        =   5415
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   16
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1605
      MaxLength       =   30
      TabIndex        =   4
      Top             =   810
      Width           =   2292
   End
   Begin VB.TextBox txtSystem 
      Height          =   300
      Left            =   1605
      MaxLength       =   3
      TabIndex        =   0
      Top             =   480
      Width           =   732
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1605
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1800
      Width           =   2325
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   1605
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2460
      Width           =   6276
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   8520
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料(&B)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   3135
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   45
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   4215
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   45
      Width           =   735
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   2340
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "子案(&S)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   345
      Index           =   3
      Left            =   4950
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   45
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "( 字首比對且大小寫相同 )"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   8
      Left            =   4950
      TabIndex        =   42
      Top             =   2220
      Width           =   2040
   End
   Begin VB.Label Label6 
      Caption         =   "e台灣電子證書"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   6810
      TabIndex        =   39
      Top             =   1920
      Width           =   2505
   End
   Begin VB.Label Label4 
      Caption         =   "△對造號數◇非本案目前號數"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   6810
      TabIndex        =   38
      Top             =   1710
      Width           =   2505
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " 大小寫需完全"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   7
      Left            =   8130
      TabIndex        =   35
      Top             =   855
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "相同！)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   5
      Left            =   8700
      TabIndex        =   34
      Top             =   1050
      Width           =   600
   End
   Begin VB.Label Label3 
      Caption         =   "符號：●代表銷卷＊代表閉卷"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   6810
      TabIndex        =   30
      Top             =   1500
      Width           =   2505
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "注意：大小寫需完全相同！"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   6
      Left            =   5280
      TabIndex        =   29
      Top             =   1200
      Width           =   2160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(字首比對查詢 )"
      Height          =   180
      Index           =   0
      Left            =   3960
      TabIndex        =   28
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(依主管機關來函碼數, 對造號數以模糊比對方式查詢, "
      Height          =   180
      Index           =   3
      Left            =   3960
      TabIndex        =   27
      Top             =   855
      Width           =   4200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(依主管機關來函碼數查詢)"
      Height          =   180
      Index           =   4
      Left            =   4620
      TabIndex        =   26
      Top             =   1515
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "( ALL：全部 )"
      Height          =   180
      Index           =   2
      Left            =   7935
      TabIndex        =   24
      Top             =   2505
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "( 字首比對且大小寫相同 )"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   1
      Left            =   3960
      TabIndex        =   23
      Top             =   1830
      Width           =   2040
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   270
      TabIndex        =   21
      Top             =   2490
      Width           =   900
   End
End
Attribute VB_Name = "frm100101_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2024/03/13 拿掉A4名條印表機Combo1的物件和程式
'Memo by Lydia 2021/12/29 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
'2006/08/24 nickc 整理 刪除95年以前的註記
Option Explicit

Dim StrTag As String, intTemp As Boolean, strTemp As String, strSQL1 As String, strSQL2 As String, StrSQL6 As String
Dim i As Integer, j As Integer, intK As Double, strSys As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, bolSelData As Boolean
Public cmdState As Integer
Dim SeekPrintL As Integer
Dim SeekPrint As Integer
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件
Dim strSQLE(1) As String, strSqlEW(1) As String 'Add by Amy 2022/12/15
Dim strEField(1) As String 'Add by Amy 2023/03/06

Private Sub SetDataListWidth()
Dim intX1 As Integer 'Added by Lydia 2019/11/01

   grdDataList.row = 0
   grdDataList.col = 0: grdDataList.Text = "V"
   grdDataList.ColWidth(0) = 200
   grdDataList.col = 1: grdDataList.Text = "本所案號"
   grdDataList.ColWidth(1) = 1620
   grdDataList.CellAlignment = flexAlignLeftCenter
   Dim iDep As String
   iDep = PUB_GetST06(strUserNum)
   grdDataList.col = 2: grdDataList.Text = "分所號"
   '電腦中心，跟分所才秀
   If PUB_GetST03(strUserNum) <> "M51" And iDep = "1" Then
       grdDataList.ColWidth(2) = 0
   Else
       grdDataList.ColWidth(2) = 900
   End If
   grdDataList.CellAlignment = flexAlignLeftCenter
   grdDataList.col = 3: grdDataList.Text = "案件名稱"
   grdDataList.ColWidth(3) = 2000
   grdDataList.CellAlignment = flexAlignLeftCenter
   grdDataList.col = 4: grdDataList.Text = "申請國家"
   grdDataList.ColWidth(4) = 1000
   grdDataList.CellAlignment = flexAlignLeftCenter
   grdDataList.col = 5: grdDataList.Text = "商品類別"
   grdDataList.ColWidth(5) = 1650
   grdDataList.CellAlignment = flexAlignLeftCenter
   grdDataList.col = 6: grdDataList.Text = "申請人"
   grdDataList.ColWidth(6) = 1250
   grdDataList.CellAlignment = flexAlignLeftCenter
   grdDataList.col = 7: grdDataList.Text = "專用期止日"
   grdDataList.ColWidth(7) = 1000
   grdDataList.CellAlignment = flexAlignLeftCenter
   grdDataList.col = 8: grdDataList.Text = "相關人"
   grdDataList.ColWidth(8) = 1250
   grdDataList.CellAlignment = flexAlignLeftCenter
   grdDataList.col = 9: grdDataList.Text = "對造號數"
   grdDataList.ColWidth(9) = 1250
   grdDataList.CellAlignment = flexAlignLeftCenter
   
   'Added by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
   If grdDataList.Cols > 9 Then
      For intX1 = 10 To grdDataList.Cols - 1
          grdDataList.col = intX1
          grdDataList.ColWidth(intX1) = 0
      Next intX1
   End If
End Sub

Private Sub chk_Click()
   '若勾選所有系統類別
   If Me.chk.Value = vbChecked Then
       Me.Text4.Text = "ALL"
   '若取消勾選所有系統類別
   Else
       Me.Text4.Text = Systemkind_g
   End If
End Sub

Public Sub PubShowNextData()
Dim i As Integer, j As Integer
Dim blnPrintAdd As Boolean
Dim ii As Integer
Dim strTmp As String
Dim bA4Print As Variant  'Added by Lydia 2016/11/10 是否列印A4名條選項
Dim nFrm As Form 'Added by Morgan 2020/12/30

   Select Case cmdState
      Case 0 '案件基本資料
            Me.Enabled = False
            For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            If Trim(grdDataList.Text) = "V" Then
              Dim Str01 As String
              grdDataList.col = 0
              grdDataList.Text = ""
              For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  grdDataList.CellBackColor = QBColor(15)
              Next j
              grdDataList.col = 1
              Str01 = SystemNumber(grdDataList, 1)
              'Modify by Amy 2019/01/10 查申請案號 104112419 △NP-121446會無資料
'              If Mid(UCase(Str01), 1, 1) = "N" Then
'                  Str01 = Mid(Str01, 2, 3)
'              End If
              Str01 = Pub_RplStr(Str01)
              'end 2019/01/10
              If Not IsNull(grdDataList.Text) Then
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                  End If
                  Select Case Pub_RplStr(Str01)
                      Case "CFP", "FCP", "P"   '專利
                            Screen.MousePointer = vbHourglass
                            frm100101_3.Show
                            frm100101_3.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_3.StrMenu
                            Screen.MousePointer = vbDefault
                      Case "CFT", "FCT", "T", "TF"   '商標
                            Screen.MousePointer = vbHourglass
                            frm100101_4.Show
                            frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_4.StrMenu
                            Screen.MousePointer = vbDefault
                      'Modify By Sindy 2009/07/24 增加LIN系統類別
                      'modify by sonia 2019/7/29 +ACS系統類別
                      Case "CFL", "FCL", "L", "LIN", "ACS"    '法務
                            Screen.MousePointer = vbHourglass
                            frm100101_5.Show
                            frm100101_5.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_5.StrMenu
                            Screen.MousePointer = vbDefault
                      Case "LA"            '顧問
                            Screen.MousePointer = vbHourglass
                            frm100101_6.Show
                            frm100101_6.Tag = Pub_RplStr(grdDataList.Text)
                            frm100101_6.StrMenu
                            Screen.MousePointer = vbDefault
                      Case Else                  '服務
                           Select Case Pub_RplStr(Str01)
                               Case "TB"    '條碼
                                  Screen.MousePointer = vbHourglass
                                  frm100101_7.Show
                                  frm100101_7.Tag = Pub_RplStr(grdDataList.Text)
                                  frm100101_7.StrMenu
                                  Screen.MousePointer = vbDefault
                               Case "TM"
                                  Screen.MousePointer = vbHourglass
                                  frm100101_8.Show
                                  frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
                                  frm100101_8.StrMenu
                                  Screen.MousePointer = vbDefault
                               Case "TD"
                                  Screen.MousePointer = vbHourglass
                                  frm100101_9.Show
                                  frm100101_9.Tag = Pub_RplStr(grdDataList.Text)
                                  frm100101_9.StrMenu
                                  Screen.MousePointer = vbDefault
                               Case "TC", "CFC"
                                  Screen.MousePointer = vbHourglass
                                  frm100101_A.Show
                                  frm100101_A.Tag = Pub_RplStr(grdDataList.Text)
                                  frm100101_A.StrMenu
                                  Screen.MousePointer = vbDefault
                               Case Else
                                  Screen.MousePointer = vbHourglass
                                  frm100101_B.Show
                                  frm100101_B.Tag = Pub_RplStr(grdDataList.Text)
                                  frm100101_B.StrMenu
                                  Screen.MousePointer = vbDefault
                            End Select
                  End Select
              End If
              Me.Enabled = True
              Exit Sub
           End If
           Next i
           Me.Enabled = True
      Case 1 '案件進度
           Me.Enabled = False
           StrTag = ""
           For i = 1 To grdDataList.Rows - 1
           grdDataList.col = 0
           grdDataList.row = i
           If Trim(grdDataList.Text) = "V" Then
              grdDataList.col = 0
              grdDataList.Text = ""
              For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  grdDataList.CellBackColor = QBColor(15)
              Next j
               grdDataList.col = 1
               If Not IsNull(grdDataList.Text) Then
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                  End If
                  Screen.MousePointer = vbHourglass
                  frm100101_2.Show
                  frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_2.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
           End If
           Next i
           Me.Enabled = True
      Case 2 '結束
         'Added by Lydia 2016/11/04 結束時跑列印A4名條清單
          If PUB_AddAddressA4List("", strExc(0)) Then
          End If
          If Val(strExc(0)) > 0 Then
             'Midified by Lydia 2016/11/10 增加放棄=刪除記錄
             'If MsgBox("尚有" & strExc(0) & "張A4名條未列印，現在是否要印？ ", vbInformation + vbYesNo) = vbYes Then
             'Modified by Lydia 2017/11/22 +國內
             bA4Print = MsgBox("尚有" & strExc(0) & "張國內A4名條未列印，現在是否要印？ (是:列印，否:下次列印，取消:刪除A4名條)", vbInformation + vbYesNoCancel)
             If bA4Print = 6 Then  '列印
                'Modified by Lydia 2017/11/03 改成操作介面
'                Load frm083014
'                frm083014.Hide
'                frm083014.Opt1(4).Value = True
'                frm083014.Text1(0).Text = strExc(0)
'                frm083014.Text1(3).Text = "1"
'                frm083014.Text1(4).Text = "1"
'                frm083014.SetPrinter Combo1
'                frm083014.cmdPrint_Click
'                Set Printer = Printers(SeekPrint)
'                Printer.Orientation = SeekPrintL
'                Unload frm083014
                frm083014.iStiu = 1
                frm083014.Show
                Me.Hide
                'end 2017/11/03
             'Added by Lydia 2016/11/10
             ElseIf bA4Print = 2 Then '取消
                cnnConnection.Execute "delete from AddressA4List where aal01='" & strUserNum & "' "
             End If
          End If
          'end 2016/11/04
          
           fnCloseAllFrm100
      '加查詢子案資料
      Case 3 '子案資料
           Me.Enabled = False
           StrTag = ""
           For i = 1 To grdDataList.Rows - 1
           grdDataList.col = 0
           grdDataList.row = i
           If Trim(grdDataList.Text) = "V" Then
              grdDataList.col = 0
              grdDataList.Text = ""
              For j = 0 To grdDataList.Cols - 1
                 grdDataList.col = j
                 grdDataList.CellBackColor = QBColor(15)
              Next j
               grdDataList.col = 1
               If Not IsNull(grdDataList.Text) Then
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                  End If
                  Screen.MousePointer = vbHourglass
                  frm100102_1.Hide
                  frm100102_3.Show
                  frm100102_3.Tag = Pub_RplStr(grdDataList.Text)
                  frm100102_3.Label3.Caption = PUB_GetCustNo(Replace(frm100102_3.Tag, "-", ""))
                  frm100102_3.Label7.Caption = PUB_GetCustName(Replace(frm100102_3.Tag, "-", ""))
                  frm100102_3.StrMenu
                  frm100102_3.Caption = Me.Caption
                  Unload frm100102_1
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
           End If
           Next i
           Me.Enabled = True
      'add by nick 2005/01/31
      'Modified by Lydia 2019/09/24 專利相關案件(4), 相似案(7)
      'Case 4
      Case 4, 7
           Me.Enabled = False
           StrTag = ""
           For i = 1 To grdDataList.Rows - 1
           grdDataList.col = 0
           grdDataList.row = i
           If Trim(grdDataList.Text) = "V" Then
              grdDataList.col = 0
              grdDataList.Text = ""
              For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  grdDataList.CellBackColor = QBColor(15)
              Next j
               grdDataList.col = 1
               If Not IsNull(grdDataList.Text) Then
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                  End If
                  Screen.MousePointer = vbHourglass
                  frm100101_h.Show
                  frm100101_h.KeyString = Pub_RplStr(grdDataList.Text)
                  'Added by Lydia 2019/09/24 區分相似案
                  If cmdState = 7 Then
                      frm100101_h.SearchKind = "相似案"
                  Else
                  'end 2019/09/24
                      frm100101_h.SearchKind = "本所案號"
                  End If 'end 2019/09/24
                  frm100101_h.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
           End If
           Next i
           Me.Enabled = True
      'add by nickc 2007/08/24 薛說要加入代表圖按鈕
      Case 5
           Me.Enabled = False
           StrTag = ""
           For i = 1 To grdDataList.Rows - 1
           grdDataList.col = 0
           grdDataList.row = i
           If Trim(grdDataList.Text) = "V" Then
              grdDataList.col = 0
              grdDataList.Text = ""
              For j = 0 To grdDataList.Cols - 1
                  grdDataList.col = j
                  grdDataList.CellBackColor = QBColor(15)
              Next j
               grdDataList.col = 1
               If Not IsNull(grdDataList.Text) Then
                  Me.Hide
                  Screen.MousePointer = vbHourglass
                  frmPic001.oCP01 = SystemNumber(Pub_RplStr(grdDataList.Text), 1)
                  frmPic001.oCP02 = SystemNumber(Pub_RplStr(grdDataList.Text), 2)
                  frmPic001.oCP03 = SystemNumber(Pub_RplStr(grdDataList.Text), 3)
                  frmPic001.oCP04 = SystemNumber(Pub_RplStr(grdDataList.Text), 4)
                  'Add By Sindy 2009/05/06
                  frmPic001.optColor(0).Enabled = False
                  frmPic001.optColor(1).Enabled = False
                  '2009/05/06 End
                  frmPic001.StrMenu
                  frmPic001.cmdOK(0).Visible = False
                  frmPic001.cmdOK(1).Visible = False
                  frmPic001.cmdOK(2).Visible = False
                  frmPic001.cmdOK(2).Enabled = False 'Add by Amy 2018/07/16
                  frmPic001.cmdOK(4).Visible = False
                  frmPic001.cmdOK(5).Visible = False
                  frmPic001.cmdOK(6).Visible = False
                  frmPic001.Label12.Visible = False
                  frmPic001.SetSeekCmdok 'Add by Amy 2018/07/16
                  frmPic001.Show vbModal
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Me.Show
               End If
           End If
           Next i
           Me.Enabled = True
      'Add By Sindy 2015/8/4
      Case 6 '地址條
          Screen.MousePointer = vbHourglass
          blnPrintAdd = False
          'Modified by Morgan 2021/6/23
          'Set Printer = Printers(Combo1.ListIndex)
          'PUB_RestorePrinter Combo1 'Mark by Lydia 2024/03/13
          'end 2021/6/23
          For ii = 1 To Me.grdDataList.Rows - 1
              If Me.grdDataList.TextMatrix(ii, 0) = "V" Then
                  strTmp = Pub_RplStr(Me.grdDataList.TextMatrix(ii, 1))
                  'Modified by Lydia 2016/11/4 改存在A4名條清單,結束時跑列印
                  'blnPrintAdd = True
                  'Load frm083014
                  'frm083014.Hide
                  'frm083014.Opt1(0).Value = True
                  'frm083014.Text1(9).Text = SystemNumber(strTmp, 1)
                  'frm083014.Text1(10).Text = SystemNumber(strTmp, 2)
                  'frm083014.Text1(11).Text = SystemNumber(strTmp, 3)
                  'frm083014.Text1(12).Text = SystemNumber(strTmp, 4)
                  'Call frm083014.Text1_LostFocus(12)
                  'frm083014.Text1(4).Text = "1"
                  'frm083014.SetPrinter Printer.DeviceName
                  'frm083014.cmdPrint_Click
                  'Unload frm083014
                  If PUB_AddAddressA4List(Replace(Pub_RplStr(strTmp), "-", ""), strExc(0)) Then
                     blnPrintAdd = True
                  End If
                  'Modified by Lydia 2017/11/22 +國內
                  If Val(strExc(0)) > 0 Then cmdOK(6).Caption = "國內A4名條 (" & Val(strExc(0)) & ")"
                  'end 2016/11/4
              End If
          Next ii
          Screen.MousePointer = vbDefault
          If blnPrintAdd = False Then
              'Modified by Lydia 2016/11/04 地址條=>A4名條
              MsgBox "請勾選欲列印A4名條的資料!!!", vbExclamation + vbOKOnly
          Else
              'Remove by Lydia 2016/11/04
              'ShowPrintOk
          End If
          '印完預設回預設印表機
          'Modified by Morgan 2021/6/23
          'Set Printer = Printers(SeekPrint)
          ''Mark by Lydia 2024/03/13
          'PUB_RestorePrinter Combo1.List(SeekPrint)
          ''end 2021/6/23
          'Printer.Orientation = SeekPrintL
          'end 2024/03/13
          
      'Added by Morgan 2020/12/30
      Case 8 'IDS清單
         Me.Enabled = False
         StrTag = ""
         For i = 1 To grdDataList.Rows - 1
         grdDataList.col = 0
         grdDataList.row = i
         If Trim(grdDataList.Text) = "V" Then
            grdDataList.col = 0
            grdDataList.Text = ""
            For j = 0 To grdDataList.Cols - 1
                grdDataList.col = j
                grdDataList.CellBackColor = QBColor(15)
            Next j
             grdDataList.col = 1
             If Not IsNull(grdDataList.Text) Then
                Me.Hide
                Set nFrm = Forms(0).GetForm("frm090401_1")
                If Not nFrm Is Nothing Then
                  grdDataList.col = 1
                  nFrm.m_CaseNo = Pub_RplStr(grdDataList.Text)
                  nFrm.Show vbModal
                End If
                Screen.MousePointer = vbDefault
                Me.Enabled = True
                Me.Show
             End If
         End If
         Next i
         Me.Enabled = True
         
      '2015/8/4 END
      'Add By Sindy 2013/7/1
      '卷宗區
      Case 13
            Me.Enabled = False
            StrTag = ""
            For i = 1 To grdDataList.Rows - 1
               grdDataList.col = 0
               grdDataList.row = i
               If Trim(grdDataList.Text) = "V" Then
                  grdDataList.col = 0
                  grdDataList.Text = ""
                  For j = 0 To grdDataList.Cols - 1
                      grdDataList.col = j
                      grdDataList.CellBackColor = QBColor(15)
                  Next j
                  grdDataList.col = 1
                  If Not IsNull(grdDataList.Text) Then
                     'Modify By Sindy 2025/6/9 原本mark,應該要放開才會記錄要回到那一作業
                     If fnSaveParentForm(Me) = False Then
                         Me.Enabled = True
                         Exit Sub
                     End If
                     '2025/6/9 END
                     Screen.MousePointer = vbHourglass
                     frm100101_L.m_strKey = Pub_RplStr(grdDataList.Text)
                     'frm100101_L.Hide
                     frm100101_L.SetParent Me
                     If frm100101_L.QueryData = True Then
                        frm100101_L.Show
                        Me.Hide
                     End If
                     Screen.MousePointer = vbDefault
                     Me.Enabled = True
                     Exit Sub
                  End If
               End If
            Next i
            Me.Enabled = True
      '2013/7/1 End
      'Add By Sindy 2018/1/17
      '原始檔
      Case 14
            Me.Enabled = False
            StrTag = ""
            For i = 1 To grdDataList.Rows - 1
               grdDataList.col = 0
               grdDataList.row = i
               If Trim(grdDataList.Text) = "V" Then
                  grdDataList.col = 0
                  grdDataList.Text = ""
                  For j = 0 To grdDataList.Cols - 1
                      grdDataList.col = j
                      grdDataList.CellBackColor = QBColor(15)
                  Next j
                  grdDataList.col = 1
                  If Not IsNull(grdDataList.Text) Then
                     'Modify By Sindy 2025/6/9 原本mark,應該要放開才會記錄要回到那一作業
                     If fnSaveParentForm(Me) = False Then
                         Me.Enabled = True
                         Exit Sub
                     End If
                     '2025/6/9 END
                     Screen.MousePointer = vbHourglass
                     frm100101_M.m_strKey = Pub_RplStr(grdDataList.Text)
                     'frm100101_M.Hide
                     frm100101_M.SetParent Me
                     If frm100101_M.QueryData = True Then
                        frm100101_M.Show
                        Me.Hide
                     End If
                     Screen.MousePointer = vbDefault
                     Me.Enabled = True
                     Exit Sub
                  End If
               End If
            Next i
            Me.Enabled = True
      '2018/1/17 End
      Case Else
   End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
   'add by nickc 2007/01/12
   If Len(Trim(Me.Text4.Text)) = 0 Then
       Me.Text4.Text = "ALL"
   End If
   cmdState = Index
   PubShowNextData
   Exit Sub
End Sub

'Modify By Sindy 2015/10/5
'Private Sub cmdSearch_Click()
Public Sub cmdSearch_Click()
'2015/10/5 END
Dim s As Integer
'Added by Lydia 2019/11/01 利益衝突案件：於對造號數後面，增加欄位
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String
Dim dblRow As Double 'Add By Sindy 2025/9/3

   bolSelData = False
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/22 清除查詢印表記錄檔欄位
   
   'add by nickc 2007/01/12
   If Len(Trim(Me.Text4.Text)) = 0 Then
      Me.Text4.Text = "ALL"
   Else
      pub_QL05 = pub_QL05 & ";" & Label5 & Text4 'Add By Sindy 2010/10/22
   End If
   strTemp = IIf(Me.Text4.Text <> "ALL", Me.Text4.Text, GetAllSysKind(Me.Text4))
   grdDataList.Rows = 2
   grdDataList.Clear
   SetDataListWidth
   If Option1(0).Value = True Then
       If Len(Trim(txtSystem)) = 0 Or Len(Trim(txtCode(0))) = 0 Then
           s = MsgBox("輸入條件不可空白", , "User 輸入錯誤")
           If Len(Trim(txtSystem)) = 0 Then txtSystem.SetFocus
           Exit Sub
       End If
   Else
       If Option1(1).Value = True Then
           If Len(Trim(Text1)) = 0 Then
               s = MsgBox("輸入條件不可空白", , "User 輸入錯誤")
               If Len(Trim(Text1)) = 0 Then Text1.SetFocus
               Exit Sub
           End If
       Else
           If Option1(2).Value = True Then
               If Len(Trim(Text5)) = 0 Then
                   s = MsgBox("輸入條件不可空白", , "User 輸入錯誤")
                   If Len(Trim(Text5)) = 0 Then Text5.SetFocus
                   Exit Sub
               End If
           Else
               If Option1(3).Value = True Then
                   If Len(Trim(Text2)) = 0 Then
                       s = MsgBox("輸入條件不可空白", , "User 輸入錯誤")
                       If Len(Trim(Text2)) = 0 Then Text2.SetFocus
                       Exit Sub
                   End If
               Else
                   If Option1(4).Value = True Then
                       If Len(Trim(Text3)) = 0 Then
                           s = MsgBox("輸入條件不可空白", , "User 輸入錯誤")
                           If Len(Trim(Text3)) = 0 Then Text3.SetFocus
                           Exit Sub
                       End If
                   'Added by Morgan 2023/2/4
                   ElseIf Option1(5).Value = True Then
                        If Len(Trim(Text6)) = 0 Then
                            s = MsgBox("輸入條件不可空白", , "User 輸入錯誤")
                            If Len(Trim(Text6)) = 0 Then Text6.SetFocus
                            Exit Sub
                        End If
                   'end 2023/2/4
                   End If
               End If
           End If
       End If
   End If
   Screen.MousePointer = vbHourglass
   strSQL1 = ""
   strSQL2 = ""
   StrSQL3 = ""
   StrSQL4 = ""
   strSQL5 = ""
   'Added by Lydia 2019/11/01 利益衝突案件：於對造號數後面，增加欄位
   SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
   SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
   SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
   SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
   SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
   intCufaCnt = 0
   'end 2019/11/01
   'Add by Amy 2022/12/15 +e符號
   'Modify by Amy 2023/03/06 bug-FCP-064401 該案無專用期且無領證發文,不應加e,避免未來資料多跑的慢優化語法
   'T、FCT已有專用期間者,或無專用期但進度檔商標之註冊費717已發文
   strSQLE(0) = ",CaseProgress E1"
   strSqlEW(0) = Replace(Replace(Replace(專利進度註冊費已發文語法, "cp", "E1.cp"), "601", "717"), "pa", "tm")
   strEField(0) = "Decode(tm136,'1',Decode(tm21,null,Decode(E1.cp10,'717','e'),'e'))"
   'P、FCP已有專用期間者,或無專用期但進度檔專利之領證601已發文
   strSQLE(1) = ",CaseProgress E1"
   strSqlEW(1) = Replace(專利進度註冊費已發文語法, "cp", "E1.cp")
   strEField(1) = "Decode(pa178,'1',Decode(pa24,null,Decode(E1.cp10,'601','e'),'e'))"
   'end 2023/03/06
   'end 2022/12/15
   
   If Option1(0).Value = True Then
      strSys = strTemp & IIf(Len(Me.txtSystem.Text) > 0, "," & txtSystem, "")
   Else
      strSys = strTemp
   End If
   StrSQL6 = ""
   Me.Enabled = False
   Dim strSql As String, lngCounter As Long, lngCounterI As Long
   Dim strText As String
   strText = ""
   lngCounterI = 0
   '5種基本資料庫分2種類
   '1.沒有輸入客戶案號,彼所案號,證書號數,申請案號         5個基本資料都讀
   '2.沒有輸入證書號數,申請案號                  4個基本資料 (商標,專利,法務, 服務)
   
   '查詢商標基本檔
   'edit by nickc 2006/08/24 加入銷卷
   'strSQL = "select ' ' AS V,decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號,NVL(NVL(TM05,TM06),TM07) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日, '' AS 相關人,'' AS 對造號數 from trademark,nation,CUSTOMER WHERE SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND "
   'Modified by Lydia 2019/11/01 利益衝突案件：增加欄位SeColTM
   'Modify by Amy 2022/12/15 +e符號 Nvl(EState,'')
   'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(1),並優化e符號語法
   strSql = "select ' ' AS V,decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||" & strEField(0) & " AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(NVL(TM05,TM06),TM07) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日, '' AS 相關人,'' AS 對造號數" & SeColTM & " from trademark,nation,CUSTOMER" & strSQLE(0) & " " & _
                "WHERE SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) " & strSqlEW(0) & "AND "
   '以本所案號查詢
   If Option1(0).Value = True Then
       If Len(txtSystem) <> 0 Then
           strSql = strSql + " TM01='" + txtSystem + "' AND "
       End If
       If Len(txtCode(0)) <> 0 Then
           strSql = strSql + " TM02='" + txtCode(0) + "' AND "
       End If
       If Len(txtCode(1)) <> 0 Then
           strSql = strSql + " TM03='" + txtCode(1) + "' AND "
       Else
           strSql = strSql + " TM03='0' AND "
       End If
       If Len(txtCode(2)) <> 0 Then
           strSql = strSql + " TM04='" + txtCode(2) + "' AND "
       Else
           strSql = strSql + " TM04='00' AND "
       End If
       pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txtSystem & "-" & txtCode(0) & "-" & IIf(Len(txtCode(1)) <> 0, txtCode(1), "0") & "-" & IIf(Len(txtCode(2)) <> 0, txtCode(2), "00") 'Add By Sindy 2010/10/22
   Else
      '以申請案號查詢
       If Option1(1).Value = True Then
           If Len(Text1) <> 0 Then
               'Modified by Lydia 2020/12/01 CFT-007984申請案號為074717/'88
               'strSql = strSql + " TM12='" + Text1 + "' AND "
               'pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Text1 'Add By Sindy 2010/10/22
               strSql = strSql + " TM12='" + ChgSQL(Text1) + "' AND "
               pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & ChgSQL(Text1)
               'end 2020/12/01
           End If
       Else
           '以客戶案件案號查詢
           If Option1(2).Value = True Then
               If Len(Text5) <> 0 Then
                   strSql = strSql + " TM35 LIKE '" + Text5 + "%' AND "
                   pub_QL05 = pub_QL05 & ";" & Option1(2).Caption & Text5 'Add By Sindy 2010/10/22
               End If
           Else
               '以審定號數/證書號數查詢
               If Option1(3).Value = True Then
                   If Len(Text2) <> 0 Then
                       strSql = strSql + " TM15='" + Text2 + "' AND "
                       pub_QL05 = pub_QL05 & ";" & Option1(3).Caption & Text2 'Add By Sindy 2010/10/22
                   End If
               Else
                   '以彼所案號查詢
                   If Option1(4).Value = True Then
                       If Len(Text3) <> 0 Then
                           '2010/3/15 MODIFY BY SONIA 改為字首比對且大小寫相同
                           'strSql = strSql + " TM45 LIKE '%" + Text3 + "%' AND "
                           strSql = strSql + " TM45 LIKE '" + Text3 + "%' AND "
                           pub_QL05 = pub_QL05 & ";" & Option1(4).Caption & Text3 'Add By Sindy 2010/10/22
                       End If
                   'Added by Morgan 2023/2/4
                   '以Client Matter ID查詢
                   ElseIf Option1(5).Value = True Then
                       If Len(Text6) <> 0 Then
                           strSql = strSql + " TM127 LIKE '" + Text6 + "%' AND "
                           pub_QL05 = pub_QL05 & ";" & Option1(5).Caption & Text6
                       End If
                   'end 2023/2/4
                   End If
               End If
           End If
       End If
   End If
   strSql = strSql + " TM10=NA01(+) and tm01 in (" & SQLGrpStr(strSys, 2) & ") "
   
   'edit by nickc 2006/08/24 加入銷卷
   'strSQL = strSQL + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01 ||'-'|| PA02 ||'-'|| PA03 ||'-'|| PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("PA25", False) & " AS 專用期止日,'' AS 相關人,'' AS 對造號數 from patent,nation,CUSTOMER WHERE SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND "
   'Modified by Lydia 2019/11/01 利益衝突案件：增加欄位SeColPA
   'Modify by Amy 2022/12/15 +e符號 Nvl(EState,'')
   'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(1),並優化e符號語法
   strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01 ||'-'|| PA02 ||'-'|| PA03 ||'-'|| PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||" & strEField(1) & " AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("PA25", False) & " AS 專用期止日,'' AS 相關人,'' AS 對造號數" & SeColPA & " from patent,nation,CUSTOMER" & strSQLE(1) & " " & _
                            "WHERE SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & strSqlEW(1) & "AND "
   If Option1(0).Value = True Then
   If Len(txtSystem) <> 0 Then
       strSql = strSql + " PA01='" + txtSystem + "' AND "
   End If
   If Len(txtCode(0)) <> 0 Then
       strSql = strSql + " PA02='" + txtCode(0) + "' AND "
   End If
   If Len(txtCode(1)) <> 0 Then
       strSql = strSql + " PA03='" + txtCode(1) + "' AND "
   Else
       strSql = strSql + " PA03='0' AND "
   End If
   If Len(txtCode(2)) <> 0 Then
       strSql = strSql + " PA04='" + txtCode(2) + "' AND "
   Else
       strSql = strSql + " PA04='00' AND "
   End If
   End If
   If Option1(1).Value = True Then
      If Len(Text1) <> 0 Then
          'Modified by Lydia 2020/12/01 CFT-007984申請案號為074717/'88
          'strSql = strSql + " PA11='" + Text1 + "' AND "
          strSql = strSql + " PA11='" + ChgSQL(Text1) + "' AND "
      End If
   End If
   If Option1(2).Value = True Then
      If Len(Text5) <> 0 Then
          strSql = strSql + " PA48 LIKE '" + Text5 + "%' AND "
      End If
   End If
   If Option1(3).Value = True Then
      If Len(Text2) <> 0 Then
          strSql = strSql + " PA22='" + Text2 + "' AND "
      End If
   End If
   If Option1(4).Value = True Then
      If Len(Text3) <> 0 Then
          '2010/3/15 MODIFY BY SONIA 改為字首比對且大小寫相同
          'strSql = strSql + " PA77 LIKE '%" + Text3 + "%' AND "
          strSql = strSql + " PA77 LIKE '" + Text3 + "%' AND "
      End If
   'Added by Morgan 2023/2/4
   ElseIf Option1(5).Value = True Then
      If Len(Text6) <> 0 Then
          strSql = strSql + " PA159 LIKE '" + Text6 + "%' AND "
      End If
   'end 2023/2/4
   End If
   strSql = strSql + " PA09=NA01(+) and pa01 in (" & SQLGrpStr(strSys, 1) & ") "
   
   'edit by nickc 2006/08/24 加入銷卷
   'strSQL = strSQL + " union all select ' ' AS V,SP01 ||'-'|| SP02 ||'-'|| SP03 ||'-'|| SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號,NVL(NVL(SP05,SP06),SP07) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("SP21", False) & " AS 專用期止日,'' AS 相關人,'' AS 對造號數 from servicepractice,nation,CUSTOMER WHERE SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND "
   'Modified by Lydia 2019/11/01 利益衝突案件：增加欄位SeColSP
   'Modify by Amy 2020/02/05 +SP73 商品類別
   strSql = strSql + " union all select ' ' AS V,SP01 ||'-'|| SP02 ||'-'|| SP03 ||'-'|| SP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(NVL(SP05,SP06),SP07) AS 案件名稱,NA03 AS 申請國家,NVL(SP73,'') AS商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("SP21", False) & " AS 專用期止日,'' AS 相關人,'' AS 對造號數" & SeColSP & " from servicepractice,nation,CUSTOMER WHERE SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) AND "
   If Option1(0).Value = True Then
      If Len(txtSystem) <> 0 Then
          strSql = strSql + " SP01='" + txtSystem + "' AND "
      End If
      If Len(txtCode(0)) <> 0 Then
          strSql = strSql + " SP02='" + txtCode(0) + "' AND "
      End If
      If Len(txtCode(1)) <> 0 Then
          strSql = strSql + " SP03='" + txtCode(1) + "' AND "
      Else
          strSql = strSql + " SP03='0' AND "
      End If
      If Len(txtCode(2)) <> 0 Then
          strSql = strSql + " SP04='" + txtCode(2) + "' AND "
      Else
          strSql = strSql + " SP04='00' AND "
      End If
   End If
   If Option1(1).Value = True Then
   
      If Len(Text1) <> 0 Then
         'Modified by Lydia 2020/12/01 CFT-007984申請案號為074717/'88
         'strSql = strSql + " SP11='" + Text1 + "' AND "
         strSql = strSql + " SP11='" + ChgSQL(Text1) + "' AND "
      End If
   End If
   If Option1(2).Value = True Then
      If Len(Text5) <> 0 Then
          strSql = strSql + " SP29 LIKE '" + Text5 + "%' AND "
      End If
   End If
   If Option1(3).Value = True Then
      If Len(Text2) <> 0 Then
          strSql = strSql + " (SP14='" + Text2 + "' Or SP32='" & Me.Text2.Text & "' ) AND "
      End If
   End If
   If Option1(4).Value = True Then
      If Len(Text3) <> 0 Then
          '2010/3/15 MODIFY BY SONIA 改為字首比對且大小寫相同
          'strSql = strSql + " SP27 LIKE '%" + Text3 + "%' AND "
          strSql = strSql + " SP27 LIKE '" + Text3 + "%' AND "
      End If
   'Added by Morgan 2023/2/4
   ElseIf Option1(5).Value = True Then
      If Len(Text6) <> 0 Then
          strSql = strSql + " SP84 LIKE '" + Text6 + "%' AND "
      End If
   'end 2023/2/4
   End If
   strSql = strSql + " SP09=NA01(+)  and sp01 in (" & SQLGrpStr(strSys, 5) & ") "
   
   '查詢法務案件
   '2008/6/26 modify by sonia 法務案若畫面申請案號欄有輸入條件但以本所案號查詢時會查不到案件
   'If Len(Text1) = 0 And Len(Text2) = 0 Then
   'Modified by Morgan 2023/2/4
   'If Option1(0).Value = True Then
   If Option1(0).Value = True Or Option1(2).Value = True Or Option1(4).Value = True Or Option1(5).Value = True Then
   'end 2023/2/4
   '2008/6/26 end
       'edit by nickc 2006/08/24 加入銷卷
       'strSQL = strSQL + " union all select ' ' AS V,LC01 ||'-'|| LC02 ||'-'|| LC03 ||'-'|| LC04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號,NVL(NVL(LC05,LC06),LC07) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,'' AS 相關人,'' AS 對造號數 from lawcase,nation,CUSTOMER  WHERE SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND "
       'Modified by Lydia 2019/11/01 利益衝突案件：增加欄位SeColLC
       strSql = strSql + " union all select ' ' AS V,LC01 ||'-'|| LC02 ||'-'|| LC03 ||'-'|| LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,'' AS 專用期止日,'' AS 相關人,'' AS 對造號數" & SeColLC & " from lawcase,nation,CUSTOMER  WHERE SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND "
       If Option1(0).Value = True Then
           If Len(txtSystem) <> 0 Then
               strSql = strSql + " LC01='" + txtSystem + "' AND "
           End If
           If Len(txtCode(0)) <> 0 Then
               strSql = strSql + " LC02='" + txtCode(0) + "' AND "
           End If
           If Len(txtCode(1)) <> 0 Then
               strSql = strSql + " LC03='" + txtCode(1) + "' AND "
           Else
               strSql = strSql + " LC03='0' AND "
           End If
           If Len(txtCode(2)) <> 0 Then
               strSql = strSql + " LC04='" + txtCode(2) + "' AND "
           Else
               strSql = strSql + " LC04='00' AND "
           End If
       End If
       If Option1(2).Value = True Then
           If Len(Text5) <> 0 Then
               strSql = strSql + " LC17 LIKE '" + Text5 + "%' AND "
           End If
       End If
       If Option1(4).Value = True Then
           If Len(Text3) <> 0 Then
               '2010/3/15 MODIFY BY SONIA 改為字首比對且大小寫相同
               'strSql = strSql + " LC23 LIKE '%" + Text3 + "%' AND "
               strSql = strSql + " LC23 LIKE '" + Text3 + "%' AND "
           End If
       'Added by Morgan 2023/2/4
       ElseIf Option1(5).Value = True Then
           If Len(Text6) <> 0 Then
               strSql = strSql + " LC51 LIKE '" + Text6 + "%' AND "
           End If
       'end 2023/2/4
       End If
       strSql = strSql + " LC15=NA01(+)  and lc01 in (" & SQLGrpStr(strSys, 3) & ") "
   End If
   
   '查詢顧問案件基本資料
   '2008/6/26 modify by sonia 法務案若畫面申請案號欄有輸入條件但以本所案號查詢時會查不到案件
   'If Len(Text3) = 0 And Len(Text5) = 0 And Len(Text1) = 0 And Len(Text2) = 0 Then
   If Option1(0).Value = True Then
   '2008/6/26 end
       'edit by nickc 2006/08/24 加入銷卷
       'strSQL = strSQL + " union all select ' ' AS V,hc01 ||'-'|| hc02 ||'-'|| hc03 ||'-'|| hc04||DECODE(HC08,'Y','＊','') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,'' AS 相關人,'' AS 對造號數 from hirecase,CUSTOMER WHERE SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) " & StrSQL4 & " and hc01 in (" & SQLGrpStr(strSys, 4) & ")  AND "
       'Modified by Lydia 2019/11/01 利益衝突案件：增加欄位SeColHC
       strSql = strSql + " union all select ' ' AS V,hc01 ||'-'|| hc02 ||'-'|| hc03 ||'-'|| hc04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,'' AS 專用期止日,'' AS 相關人,'' AS 對造號數" & SeColHC & " from hirecase,CUSTOMER WHERE SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) " & StrSQL4 & " and hc01 in (" & SQLGrpStr(strSys, 4) & ")  AND "
       If Option1(0).Value = True Then
       If Len(txtSystem) <> 0 Then
           strSql = strSql + " hc01='" + txtSystem + "' AND "
       End If
       If Len(txtCode(0)) <> 0 Then
           strSql = strSql + " hc02='" + txtCode(0) + "' AND "
       End If
       If Len(txtCode(1)) <> 0 Then
           strSql = strSql + " hc03='" + txtCode(1) + "' AND "
       Else
           strSql = strSql + " hc03='0' AND "
       End If
       If Len(txtCode(2)) <> 0 Then
           strSql = strSql + " hc04='" + txtCode(2) + "' AND "
       Else
           strSql = strSql + " hc04='00' AND "
       End If
       End If
   End If
   If Right(strSql, 4) = "AND " Then
       strSql = Left(strSql, Len(strSql) - 4)
   End If
   If Right(strSql, 6) = "WHERE " Then
       strSql = Left(strSql, Len(strSql) - 6)
   End If
   '檢查是否使用申請案號查詢
   If Option1(1).Value = True Then
       If Len(Trim(Text1)) <> 0 Then
   'edit by nickc 2006/07/24 秀玲說，案件名稱改成抓基本檔，只有對造號數抓對造名稱
   '        strSQL = strSQL + " union all select ' ' AS V,decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,trademark,nation,CUSTOMER WHERE CP30='" & Text1 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 2) & ") " & strSQL2
   '        strSQL = strSQL + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01 ||'-'|| PA02 ||'-'|| PA03 ||'-'|| PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("PA25", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,patent,nation,CUSTOMER WHERE CP30='" & Text1 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA09=NA01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 1) & ") " & strSQL1
   '        strSQL = strSQL + " union all select ' ' AS V,SP01 ||'-'|| SP02 ||'-'|| SP03 ||'-'|| SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("SP21", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,servicepractice,nation,CUSTOMER WHERE CP30='" & Text1 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 5) & ") " & strSQL5
   '        strSQL = strSQL + " union all select ' ' AS V,LC01 ||'-'|| LC02 ||'-'|| LC03 ||'-'|| LC04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,lawcase,nation,CUSTOMER WHERE CP30='" & Text1 & "' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC15=NA01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 3) & ") " & StrSQL3
   '        strSQL = strSQL + " union all select ' ' AS V,HC01 ||'-'|| HC02 ||'-'|| HC03 ||'-'|| HC04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,' ' AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,hirecase,CUSTOMER WHERE CP30='" & Text1 & "' AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 4) & ") " & StrSQL4 & " "
   'edit by nickc 2006/08/24 加入銷卷
   '        strSQL = strSQL + " union all select ' ' AS V,decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號,NVL(NVL(TM05,TM06),TM07) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,trademark,nation,CUSTOMER WHERE CP30='" & Text1 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 2) & ") and not (cp01='CFT' and cp10='304' and substr(cp09,1,1)='B' ) " & strSQL2
   '        strSQL = strSQL + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01 ||'-'|| PA02 ||'-'|| PA03 ||'-'|| PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("PA25", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,patent,nation,CUSTOMER WHERE CP30='" & Text1 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA09=NA01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 1) & ") " & strSQL1
   '        strSQL = strSQL + " union all select ' ' AS V,SP01 ||'-'|| SP02 ||'-'|| SP03 ||'-'|| SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號,NVL(NVL(SP05,SP06),SP07) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("SP21", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,servicepractice,nation,CUSTOMER WHERE CP30='" & Text1 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 5) & ") " & strSQL5
   '        strSQL = strSQL + " union all select ' ' AS V,LC01 ||'-'|| LC02 ||'-'|| LC03 ||'-'|| LC04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號,NVL(NVL(LC05,LC06),LC07) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,lawcase,nation,CUSTOMER WHERE CP30='" & Text1 & "' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC15=NA01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 3) & ") " & StrSQL3
   '        strSQL = strSQL + " union all select ' ' AS V,HC01 ||'-'|| HC02 ||'-'|| HC03 ||'-'|| HC04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號,hc07 AS 案件名稱,' ' AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,hirecase,CUSTOMER WHERE CP30='" & Text1 & "' AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 4) & ") " & StrSQL4 & " "
   'MODIFY BY Sindy 2012/5/31 抓出的資料來源為CP30者為非本所目前號數,本所案號前加◇
   'Modify By Sindy 2012/6/21 商標案件若CP30=TM12者不抓CP資料,例 10409446 故+and CP30<>TM12
   '                          專利案件若CP30=PA11者不抓CP資料,例 100203667 故+and CP30<>PA11
   '                          服務業務案件T*者,若CP30=SP11者不抓CP資料,例 2011Z11S008668 故+and CP30<>SP11 (檢查系統資料中TC有相同)
   'MODIFY BY SONIA 2016/3/2 6350700065會查不到T-193942故CP30<>TM12改為(TM12 IS NULL OR CP30<>TM12),CP30<>PA11改為(PA11 IS NULL OR CP30<>PA11),CP30<>SP11改為(SP11 IS NULL OR CP30<>SP11)
           'Modified by Lydia 2019/11/01 利益衝突案件：增加欄位SeColTM,SeColPA,SeColSP,SeColLC,SeColHC
           'Modify by Amy 2020/02/05 +SP73 商品類別
           'Modified by Lydia 2020/12/01 CFT-007984申請案號為074717/'88;  + chgsql
           'Modify by Amy 2022/12/15 +e符號 Nvl(EState,'')
           'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(x),並優化e符號語法
           strSql = strSql + " union all select ' ' AS V,'◇'||decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||" & strEField(0) & " AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(NVL(TM05,TM06),TM07) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP.CP40,CP.CP41),CP.CP42) AS 相關人,CP.CP36 AS 對造號數" & SeColTM & " from CASEPROGRESS CP,trademark,nation,CUSTOMER" & strSQLE(0) & " " & _
                                    "WHERE CP.CP30='" & ChgSQL(Text1) & "' AND CP.CP01=TM01(+) AND CP.CP02=TM02(+) AND CP.CP03=TM03(+) AND CP.CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and CP.cp01 in (" & SQLGrpStr(strSys, 2) & ") and not (CP.cp01='CFT' and CP.cp10='304' and substr(CP.cp09,1,1)='B') and (TM12 IS NULL OR CP.CP30<>TM12) " & strSqlEW(0) & strSQL2
           strSql = strSql + " union all select ' ' AS V,'◇'||decode(pa23,'1','','N')||PA01 ||'-'|| PA02 ||'-'|| PA03 ||'-'|| PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||" & strEField(1) & " AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("PA25", False) & " AS 專用期止日,NVL(NVL(CP.CP40,CP.CP41),CP.CP42) AS 相關人,CP.CP36 AS 對造號數" & SeColPA & " from CASEPROGRESS CP,patent,nation,CUSTOMER" & strSQLE(1) & " " & _
                                    "WHERE CP.CP30='" & ChgSQL(Text1) & "' AND CP.CP01=PA01(+) AND CP.CP02=PA02(+) AND CP.CP03=PA03(+) AND CP.CP04=PA04(+) AND PA09=NA01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and CP.cp01 in (" & SQLGrpStr(strSys, 1) & ") and (PA11 IS NULL OR CP.CP30<>PA11) " & strSqlEW(1) & strSQL1
           'end 2022/12/15
           strSql = strSql + " union all select ' ' AS V,'◇'||SP01 ||'-'|| SP02 ||'-'|| SP03 ||'-'|| SP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(NVL(SP05,SP06),SP07) AS 案件名稱,NA03 AS 申請國家,NVL(SP73,'') AS商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("SP21", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數" & SeColSP & " from CASEPROGRESS,servicepractice,nation,CUSTOMER WHERE CP30='" & ChgSQL(Text1) & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 5) & ") and (SP11 IS NULL OR CP30<>SP11) " & strSQL5
           strSql = strSql + " union all select ' ' AS V,'◇'||LC01 ||'-'|| LC02 ||'-'|| LC03 ||'-'|| LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數" & SeColLC & " from CASEPROGRESS,lawcase,nation,CUSTOMER WHERE CP30='" & ChgSQL(Text1) & "' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC15=NA01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 3) & ") " & StrSQL3
           strSql = strSql + " union all select ' ' AS V,'◇'||HC01 ||'-'|| HC02 ||'-'|| HC03 ||'-'|| HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,hc07 AS 案件名稱,' ' AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數" & SeColHC & " from CASEPROGRESS,hirecase,CUSTOMER WHERE CP30='" & ChgSQL(Text1) & "' AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 4) & ") " & StrSQL4 & " "
           
   'edit by nickc 2006/08/24 加入銷卷
   '        strSQL = strSQL + " union all select ' ' AS V,decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,trademark,nation,CUSTOMER WHERE CP36 like '" & Text1 & "%' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 2) & ") " & strSQL2
   '        strSQL = strSQL + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01 ||'-'|| PA02 ||'-'|| PA03 ||'-'|| PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("PA25", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,patent,nation,CUSTOMER WHERE CP36 like '" & Text1 & "%' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA09=NA01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 1) & ") " & strSQL1
   '        strSQL = strSQL + " union all select ' ' AS V,SP01 ||'-'|| SP02 ||'-'|| SP03 ||'-'|| SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("SP21", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,servicepractice,nation,CUSTOMER WHERE CP36 like '" & Text1 & "%' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 5) & ") " & strSQL5
   '        strSQL = strSQL + " union all select ' ' AS V,LC01 ||'-'|| LC02 ||'-'|| LC03 ||'-'|| LC04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,lawcase,nation,CUSTOMER WHERE CP36 like '" & Text1 & "%' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC15=NA01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 3) & ") " & StrSQL3
   '        strSQL = strSQL + " union all select ' ' AS V,HC01 ||'-'|| HC02 ||'-'|| HC03 ||'-'|| HC04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,' ' AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,hirecase,CUSTOMER WHERE CP36 like '" & Text1 & "%' AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 4) & ") " & StrSQL4 & " ORDER BY 本所案號"
   '2009/8/4 MODIFY BY SONIA 對造號數,本所案號前加△
           'Modified by Lydia 2019/11/01 利益衝突案件：增加欄位SeColTM,SeColPA,SeColSP,SeColLC,SeColHC
           'Modify by Amy 2020/02/05 +SP73 商品類別
           'Modified by Lydia 2020/12/01 CFT-007984申請案號為074717/'88;  + chgsql
           'Modify by Amy 2022/12/15 +e符號 Nvl(EState,'')
           'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(x),並優化e符號語法
           strSql = strSql + " union all select ' ' AS V,'△'||decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||" & strEField(0) & " AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(NVL(CP.CP37,CP.CP38),CP.CP39) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP.CP40,CP.CP41),CP.CP42) AS 相關人,CP.CP36 AS 對造號數" & SeColTM & " from CASEPROGRESS CP,trademark,nation,CUSTOMER" & strSQLE(0) & " " & _
                        "WHERE CP.CP36 like '" & ChgSQL(Text1) & "%' AND CP.CP01=TM01(+) AND CP.CP02=TM02(+) AND CP.CP03=TM03(+) AND CP.CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and CP.cp01 in (" & SQLGrpStr(strSys, 2) & ") " & strSqlEW(0) & strSQL2
           strSql = strSql + " union all select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01 ||'-'|| PA02 ||'-'|| PA03 ||'-'|| PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||" & strEField(1) & " AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(NVL(CP.CP37,CP.CP38),CP.CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("PA25", False) & " AS 專用期止日,NVL(NVL(CP.CP40,CP.CP41),CP.CP42) AS 相關人,CP.CP36 AS 對造號數" & SeColPA & " from CASEPROGRESS CP,patent,nation,CUSTOMER" & strSQLE(1) & " " & _
                        "WHERE CP.CP36 like '" & ChgSQL(Text1) & "%' AND CP.CP01=PA01(+) AND CP.CP02=PA02(+) AND CP.CP03=PA03(+) AND CP.CP04=PA04(+) AND PA09=NA01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and CP.cp01 in (" & SQLGrpStr(strSys, 1) & ") " & strSqlEW(1) & strSQL1
           'end 2022/12/15
           strSql = strSql + " union all select ' ' AS V,'△'||SP01 ||'-'|| SP02 ||'-'|| SP03 ||'-'|| SP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,NVL(SP73,'') AS商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("SP21", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數" & SeColSP & " from CASEPROGRESS,servicepractice,nation,CUSTOMER WHERE CP36 like '" & ChgSQL(Text1) & "%' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 5) & ") " & strSQL5
           strSql = strSql + " union all select ' ' AS V,'△'||LC01 ||'-'|| LC02 ||'-'|| LC03 ||'-'|| LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數" & SeColLC & " from CASEPROGRESS,lawcase,nation,CUSTOMER WHERE CP36 like '" & ChgSQL(Text1) & "%' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC15=NA01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 3) & ") " & StrSQL3
           strSql = strSql + " union all select ' ' AS V,'△'||HC01 ||'-'|| HC02 ||'-'|| HC03 ||'-'|| HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,' ' AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數" & SeColHC & " from CASEPROGRESS,hirecase,CUSTOMER WHERE CP36 like '" & ChgSQL(Text1) & "%' AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 4) & ") " & StrSQL4 & " ORDER BY 本所案號"
   
       End If
   End If
   '檢查是否使用審定號數/證書號數
   If Option1(3).Value = True Then
       If Len(Trim(Text2)) <> 0 Then
           'edit by nickc 2006/08/24 加入銷卷
   '        strSQL = strSQL + " union all select ' ' AS V,decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,trademark,nation,CUSTOMER WHERE CP36='" & Text2 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 2) & ") " & strSQL2
   '        strSQL = strSQL + " union all select ' ' AS V,decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號,NVL(NVL(TM05,TM06),TM07) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,trademark,nation,CUSTOMER WHERE CP30='" & Text2 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 2) & ") and cp01='CFT' and cp10='304' and substr(cp09,1,1)='B' " & strSQL2
   '        strSQL = strSQL + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01 ||'-'|| PA02 ||'-'|| PA03 ||'-'|| PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("PA25", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,patent,nation,CUSTOMER WHERE CP36='" & Text2 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA09=NA01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 1) & ") " & strSQL1
   '        strSQL = strSQL + " union all select ' ' AS V,SP01 ||'-'|| SP02 ||'-'|| SP03 ||'-'|| SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("SP21", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,servicepractice,nation,CUSTOMER WHERE CP36='" & Text2 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 5) & ") " & strSQL5
   '        strSQL = strSQL + " union all select ' ' AS V,LC01 ||'-'|| LC02 ||'-'|| LC03 ||'-'|| LC04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,lawcase,nation,CUSTOMER WHERE CP36='" & Text2 & "' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC15=NA01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 3) & ") " & StrSQL3
   '        strSQL = strSQL + " union all select ' ' AS V,HC01 ||'-'|| HC02 ||'-'|| HC03 ||'-'|| HC04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,' ' AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,hirecase,CUSTOMER WHERE CP36='" & Text2 & "' AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 4) & ") " & StrSQL4 & " ORDER BY 本所案號"
   '2009/8/4 MODIFY BY SONIA 對造號數,本所案號前加△
   'MODIFY BY Sindy 2012/5/31 抓出的資料來源為CP30者為非本所目前號數,本所案號前加◇
   'Modify By Sindy 2012/6/21 商標案件若CP30=TM12者不抓CP資料,例 10409446 故+and CP30<>TM12
           'Modified by Lydia 2019/11/01 利益衝突案件：增加欄位SeColTM
           'Modify by Amy 2022/12/15 +e符號 Nvl(EState,'')
           'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(x),並優化e符號語法
           strSql = strSql + " union all select ' ' AS V,'△'||decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||" & strEField(0) & " AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(NVL(CP.CP37,CP.CP38),CP.CP39) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP.CP40,CP.CP41),CP.CP42) AS 相關人,CP.CP36 AS 對造號數" & SeColTM & " from CASEPROGRESS CP,trademark,nation,CUSTOMER" & strSQLE(0) & " " & _
                        "WHERE CP.CP36='" & Text2 & "' AND CP.CP01=TM01(+) AND CP.CP02=TM02(+) AND CP.CP03=TM03(+) AND CP.CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and CP.cp01 in (" & SQLGrpStr(strSys, 2) & ") " & strSqlEW(0) & strSQL2
           '2013/9/10 modify by sonia 改cp10='304'為cp10 in ('304','102'),取消substr(cp09,1,1)='B',但加入and CP30<>TM15 (CFT延展證書,延展那一道存原註冊號數)
           'strSql = strSql + " union all select ' ' AS V,decode(tm28,'1','','N')||'◇'||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(NVL(TM05,TM06),TM07) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,trademark,nation,CUSTOMER WHERE CP30='" & Text2 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 2) & ") and cp01='CFT' and cp10='304' and substr(cp09,1,1)='B' and CP30<>TM12 " & strSQL2
           'CFT案之304,102
           'Modified by Lydia 2019/11/01 利益衝突案件：增加欄位SeColTM,SeColPA,SeColSP,SeColLC,SeColHC
           'Modify by Amy 2020/02/05 +SP73 商品類別
           strSql = strSql + " union all select ' ' AS V,decode(tm28,'1','','N')||'◇'||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||" & strEField(0) & " AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(NVL(TM05,TM06),TM07) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP.CP40,CP.CP41),CP.CP42) AS 相關人,CP.CP36 AS 對造號數" & SeColTM & " from CASEPROGRESS CP,trademark,nation,CUSTOMER" & strSQLE(0) & " " & _
                        "WHERE CP.CP30='" & Text2 & "' AND CP.CP01=TM01(+) AND CP.CP02=TM02(+) AND CP.CP03=TM03(+) AND CP.CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and CP.cp01 in (" & SQLGrpStr(strSys, 2) & ") and CP.cp01='CFT' and CP.cp10 in ('304','102') and (CP.CP30<>TM12 OR TM12 IS NULL) and (CP.CP30<>TM15 OR TM15 IS NULL) " & strSqlEW(0) & strSQL2
           strSql = strSql + " union all select ' ' AS V,'△'||decode(pa23,'1','','N')||PA01 ||'-'|| PA02 ||'-'|| PA03 ||'-'|| PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||" & strEField(1) & " AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(NVL(CP.CP37,CP.CP38),CP.CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("PA25", False) & " AS 專用期止日,NVL(NVL(CP.CP40,CP.CP41),CP.CP42) AS 相關人,CP.CP36 AS 對造號數" & SeColPA & " from CASEPROGRESS CP,patent,nation,CUSTOMER" & strSQLE(1) & " " & _
                        "WHERE CP.CP36='" & Text2 & "' AND CP.CP01=PA01(+) AND CP.CP02=PA02(+) AND CP.CP03=PA03(+) AND CP.CP04=PA04(+) AND PA09=NA01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and CP.cp01 in (" & SQLGrpStr(strSys, 1) & ") " & strSqlEW(1) & strSQL1
           'end 2022/12/15
           strSql = strSql + " union all select ' ' AS V,'△'||SP01 ||'-'|| SP02 ||'-'|| SP03 ||'-'|| SP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,NVL(SP73,'') AS商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("SP21", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數" & SeColSP & " from CASEPROGRESS,servicepractice,nation,CUSTOMER WHERE CP36='" & Text2 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 5) & ") " & strSQL5
           strSql = strSql + " union all select ' ' AS V,'△'||LC01 ||'-'|| LC02 ||'-'|| LC03 ||'-'|| LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數" & SeColLC & " from CASEPROGRESS,lawcase,nation,CUSTOMER WHERE CP36='" & Text2 & "' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC15=NA01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 3) & ") " & StrSQL3
           strSql = strSql + " union all select ' ' AS V,'△'||HC01 ||'-'|| HC02 ||'-'|| HC03 ||'-'|| HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,' ' AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數" & SeColHC & " from CASEPROGRESS,hirecase,CUSTOMER WHERE CP36='" & Text2 & "' AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 4) & ") " & StrSQL4 & " ORDER BY 本所案號"
       End If
   End If
   
   '檢查是否使用彼所案號
   If Option1(4).Value = True Then
      If Len(Text3) <> 0 Then
           'edit by nickc 2006/08/24 加入銷卷
   '        strSQL = strSQL + " union all select ' ' AS V,decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,trademark,nation,CUSTOMER WHERE CP45 LIKE '%" & Text3 & "%' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 2) & ") " & strSQL2
   '        strSQL = strSQL + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01 ||'-'|| PA02 ||'-'|| PA03 ||'-'|| PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("PA25", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,patent,nation,CUSTOMER WHERE CP45 LIKE '%" & Text3 & "%' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA09=NA01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 1) & ") " & strSQL1
   '        strSQL = strSQL + " union all select ' ' AS V,SP01 ||'-'|| SP02 ||'-'|| SP03 ||'-'|| SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("SP21", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,servicepractice,nation,CUSTOMER WHERE CP45 LIKE '%" & Text3 & "%' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 5) & ") " & strSQL5
   '        strSQL = strSQL + " union all select ' ' AS V,LC01 ||'-'|| LC02 ||'-'|| LC03 ||'-'|| LC04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,lawcase,nation,CUSTOMER WHERE CP45 LIKE '%" & Text3 & "%' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC15=NA01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 3) & ") " & StrSQL3
   '        strSQL = strSQL + " union all select ' ' AS V,HC01 ||'-'|| HC02 ||'-'|| HC03 ||'-'|| HC04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 AS 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,' ' AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,hirecase,CUSTOMER WHERE CP45 LIKE '%" & Text3 & "%' AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 4) & ") " & StrSQL4 & " ORDER BY 本所案號"
           '2010/3/15 MODIFY BY SONIA 改為字首比對且大小寫相同
           'strSql = strSql + " union all select ' ' AS V,decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,trademark,nation,CUSTOMER WHERE CP45 LIKE '%" & Text3 & "%' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 2) & ") " & strSQL2
           'strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01 ||'-'|| PA02 ||'-'|| PA03 ||'-'|| PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("PA25", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,patent,nation,CUSTOMER WHERE CP45 LIKE '%" & Text3 & "%' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA09=NA01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 1) & ") " & strSQL1
           'strSql = strSql + " union all select ' ' AS V,SP01 ||'-'|| SP02 ||'-'|| SP03 ||'-'|| SP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人," & SQLDate("SP21", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,servicepractice,nation,CUSTOMER WHERE CP45 LIKE '%" & Text3 & "%' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 5) & ") " & strSQL5
           'strSql = strSql + " union all select ' ' AS V,LC01 ||'-'|| LC02 ||'-'|| LC03 ||'-'|| LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,lawcase,nation,CUSTOMER WHERE CP45 LIKE '%" & Text3 & "%' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC15=NA01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 3) & ") " & StrSQL3
           'strSql = strSql + " union all select ' ' AS V,HC01 ||'-'|| HC02 ||'-'|| HC03 ||'-'|| HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,NVL(NVL(CP37,CP38),CP39) AS 案件名稱,' ' AS 申請國家,' ' AS 商品類別,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數 from CASEPROGRESS,hirecase,CUSTOMER WHERE CP45 LIKE '%" & Text3 & "%' AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 4) & ") " & StrSQL4 & " ORDER BY 本所案號"
           'MODIFY BY SONIA 2014/6/5 案件名稱改抓基本檔,原為NVL(NVL(CP37,CP38),CP39) CMT-1F-140270(T-192069)
           'Modified by Lydia 2019/11/01 +增加欄位SeColTM,SeColPA,SeColSP,SeColLC,SeColHC
           'Modify by Amy 2020/02/05 +SP73 商品類別
           'Modify by Amy 2022/12/15 +e符號 Nvl(EState,'')
           'Modify by Amy 2023/03/06 Nvl(EState,'')改為strEField(x),並優化e符號語法
           strSql = strSql + " union all select ' ' AS V,decode(tm28,'1','','N')||TM01 ||'-'|| TM02 ||'-'|| TM03 ||'-'|| TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||" & strEField(0) & " AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(NVL(TM05,TM06),TM07) AS 案件名稱,NA03 AS 申請國家,TM09 AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("TM22", False) & " AS 專用期止日,NVL(NVL(CP.CP40,CP.CP41),CP.CP42) AS 相關人,CP.CP36 AS 對造號數" & SeColTM & " from CASEPROGRESS CP,trademark,nation,CUSTOMER" & strSQLE(0) & " " & _
                        "WHERE CP.CP45 LIKE '" & Text3 & "%' AND CP.CP01=TM01(+) AND CP.CP02=TM02(+) AND CP.CP03=TM03(+) AND CP.CP04=TM04(+) AND TM10=NA01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and CP.cp01 in (" & SQLGrpStr(strSys, 2) & ") " & strSqlEW(0) & strSQL2
           strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01 ||'-'|| PA02 ||'-'|| PA03 ||'-'|| PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')||" & strEField(1) & " AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("PA25", False) & " AS 專用期止日,NVL(NVL(CP.CP40,CP.CP41),CP.CP42) AS 相關人,CP.CP36 AS 對造號數" & SeColPA & " from CASEPROGRESS CP,patent,nation,CUSTOMER" & strSQLE(1) & " " & _
                        "WHERE CP.CP45 LIKE '" & Text3 & "%' AND CP.CP01=PA01(+) AND CP.CP02=PA02(+) AND CP.CP03=PA03(+) AND CP.CP04=PA04(+) AND PA09=NA01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and CP.cp01 in (" & SQLGrpStr(strSys, 1) & ") " & strSqlEW(1) & strSQL1
           'end 2022/12/15
           strSql = strSql + " union all select ' ' AS V,SP01 ||'-'|| SP02 ||'-'|| SP03 ||'-'|| SP04||DECODE(SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(NVL(SP05,SP06),SP07) AS 案件名稱,NA03 AS 申請國家,NVL(SP73,'') AS商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人," & SQLDate("SP21", False) & " AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數" & SeColSP & " from CASEPROGRESS,servicepractice,nation,CUSTOMER WHERE CP45 LIKE '" & Text3 & "%' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 5) & ") " & strSQL5
           strSql = strSql + " union all select ' ' AS V,LC01 ||'-'|| LC02 ||'-'|| LC03 ||'-'|| LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(NVL(LC05,LC06),LC07) AS 案件名稱,NA03 AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數" & SeColLC & " from CASEPROGRESS,lawcase,nation,CUSTOMER WHERE CP45 LIKE '" & Text3 & "%' AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC15=NA01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 3) & ") " & StrSQL3
           strSql = strSql + " union all select ' ' AS V,HC01 ||'-'|| HC02 ||'-'|| HC03 ||'-'|| HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,hc07 AS 案件名稱,' ' AS 申請國家,' ' AS 商品類別,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,'' AS 專用期止日,NVL(NVL(CP40,CP41),CP42) AS 相關人,CP36 AS 對造號數" & SeColHC & " from CASEPROGRESS,hirecase,CUSTOMER WHERE CP45 LIKE '" & Text3 & "%' AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) and cp01 in (" & SQLGrpStr(strSys, 4) & ") " & StrSQL4 & " ORDER BY 本所案號"
       End If
   End If
   CheckOC
   Screen.MousePointer = vbHourglass
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   If adoRecordset.RecordCount <> 0 Then
       dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
       grdDataList.Rows = adoRecordset.RecordCount + 1
       If Not cmdOK(0).Enabled Then cmdOK(0).Enabled = True
       If Not cmdOK(1).Enabled Then cmdOK(1).Enabled = True
       'Add By Cheng 2003/04/09
       If Not cmdOK(3).Enabled Then cmdOK(3).Enabled = True
       'add by nick 2005/01/31
       If Not cmdOK(4).Enabled Then cmdOK(4).Enabled = True
       'Added by Lydia 2019/09/24 相似案: 外專開放給P,FCP案
       If ("" & adoRecordset.Fields("本所案號") <> "" And (Left("" & adoRecordset.Fields("本所案號"), 2) = "P-" Or Left("" & adoRecordset.Fields("本所案號"), 4) = "FCP-")) _
            And (Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 2) = "F2") Then
            cmdOK(7).Visible = True
       Else
            cmdOK(7).Visible = False
       End If
       'end 2019/09/24
       
       'Added by Morgan 2020/12/30
       cmdOK(8).Visible = False
       If (Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 2) = "P1") And Left(Pub_RplStr("" & adoRecordset.Fields("本所案號")), 4) = "CFP-" And "" & adoRecordset.Fields("申請國家") = "美國" Then
            PUB_GetIDSList Pub_RplStr("" & adoRecordset.Fields("本所案號")), intI
            If intI = 1 Then
               'Modified by Morgan 2021/1/18 +判斷有Form才顯示
               If Not Forms(0).GetForm("frm090401_1") Is Nothing Then
                  cmdOK(8).Visible = True
               End If
            End If
       End If
       'end 2020/12/30
       
       'add by nickc 2007/08/24
       If Not cmdOK(5).Enabled Then cmdOK(5).Enabled = True
       'Add By Sindy 2013/7/1
'       If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
         'Modify By Sindy 2015/3/9 Mark:不分系統別都可以查看
'         If txtSystem = "P" Or txtSystem = "CFP" Then
'            cmdOK(13).Visible = True
'            cmdOK(13).Enabled = True
'         Else
'            cmdOK(13).Visible = False
'            cmdOK(13).Enabled = False
'         End If
         If Not cmdOK(13).Enabled Then cmdOK(13).Enabled = True
         If Not cmdOK(14).Enabled Then cmdOK(14).Enabled = True
'       Else
'         cmdOK(13).Visible = False
'         cmdOK(13).Enabled = False
'       End If
       '2013/7/1 End
   Else
       InsertQueryLog (0) 'Add By Sindy 2010/10/22
       cmdOK(0).Enabled = False
       cmdOK(1).Enabled = False
       'Add By Cheng 2003/04/09
       cmdOK(3).Enabled = False
       'add by nick 2005/01/31
       cmdOK(4).Enabled = False
       'add by nickc 2007/08/24
       cmdOK(5).Enabled = False
       'Add by Sindy 2013/7/1
'       cmdOK(13).Visible = False
       cmdOK(13).Enabled = False
       '2013/7/1 END
       cmdOK(14).Enabled = False 'Add By Sindy 2018/1/17
       CheckOC
       ShowNoData
       Screen.MousePointer = vbDefault
       Me.Enabled = True
       Exit Sub
   End If
   
   '檢查是否重複
   Dim StrTest As String, StrTest2 As String, StrTest3 As Variant, StrTest4 As String
   'Modify By Cheng 2002/03/14
   'If Len(Trim(Text4)) <> 0 Then
   '   StrTest3 = Split(Text4, ",")
   'End If
   If Len(Trim(strTemp)) <> 0 Then
      StrTest3 = Split(strTemp, ",")
   End If
   adoRecordset.MoveFirst
   Do While adoRecordset.EOF = False
     StrTest2 = adoRecordset.Fields(1)
     If Left(StrTest2, 1) = "N" Then
        StrTest2 = Right(StrTest2, Len(StrTest2) - 1)
     End If
     If StrTest <> StrTest2 Then
         'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷
         If PUB_ChkCufaByCase(Me.Name, m_AllSys, StrTest2, "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
             intCufaCnt = intCufaCnt + 1
             adoRecordset.Delete
         Else
         'end 2019/11/01
             StrTest = StrTest2
         End If
     Else
        adoRecordset.Delete
     End If
     adoRecordset.MoveNext
   Loop
   
   'Added by Lydia 2019/11/01 利益衝突案件：限閱案件
   If intCufaCnt > 0 Then
       pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
       MsgBox "限閱案件", vbInformation, MsgText(1110)
   End If
   InsertQueryLog (dblRow) 'Add By Sindy 2010/10/22
   '把資料放進   GRID
   '911029 nick edit
   If adoRecordset.RecordCount <> 0 Then
       Set grdDataList.Recordset = adoRecordset
       intK = adoRecordset.RecordCount
   End If
   CheckOC
   
   'Add By Cheng 2001/12/26
   '若查詢結果只有一筆資料
   If Me.grdDataList.Rows = 2 Then
       '911029 nick add
       grdDataList.row = 1
       grdDataList.col = 1
       If grdDataList.Text <> "" Then
       '911029 nick end
           '直接選定
           bolSelData = True
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
       '912029 nick end
   End If
   'Add by Morgan 2004/3/11
   If Option1(0).Value = True Then
       Call txtCode_GotFocus(0)
   ElseIf Option1(1).Value = True Then
       Call Text1_GotFocus
   ElseIf Option1(2).Value = True Then
       Call Text5_GotFocus
   ElseIf Option1(3).Value = True Then
       Call Text2_GotFocus
   ElseIf Option1(4).Value = True Then
       Call Text3_GotFocus
   'Added by Morgan 2023/2/4
   ElseIf Option1(5).Value = True Then
       Call Text6_GotFocus
   'end 2023/2/4
   End If
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   'Modify By Cheng 2002/03/14
   'Text4 = strTemp
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   '主檔    不需傳入資料
   grdDataList.Clear
   SetDataListWidth
   cmdOK(0).Enabled = False
   cmdOK(1).Enabled = False
   'Add By Cheng 2003/04/09
   Me.cmdOK(3).Enabled = False
   'add by nick 2005/01/31
   cmdOK(4).Enabled = False
   cmdOK(7).Visible = False  'Added by Lydia 2019/09/24 相似案
   'add by nickc 2007/08/24
   cmdOK(5).Enabled = False
   'Text1.Enabled = False
   'Text2.Enabled = False
   'Text3.Enabled = False
   'Text5.Enabled = False
   '2011/12/6 modify by sonia
   'Text4 = Systemkind_g
   Me.chk.Value = vbChecked
   Text4 = "ALL"
   m_AllSys = GetAllSysKind(, Text4)  'Added by Lydia 2019/11/01
   '2011/12/6 end
   bolToEndByNick = False
   intTemp = False
   'txtSystem.SetFocus
   bolSelData = False
   '92.04.16 nick
   cmdState = -1
   'Add by Morgan 2010/8/23
'   If bolNewAppNoFormat Then
'      Label1(5).Caption = "注意：申請案號大小寫需完全相同！"
'      'Modified by Lydia 2016/11/04
'      'Label1(5).Move 6390, 540
'      'Remove by Lydia 2017/11/22
'      'Label1(5).Move 6560, 540
'      Label1(3).Caption = "(依主管機關來函碼數查詢,對造號數以模糊比對方式查詢)"
'      Label1(4).Caption = "(依主管機關來函碼數查詢)"
'   End If
   
   'Added by Lydia 2016/11/04 顯示未列印的A4名條數量
    If PUB_AddAddressA4List("", strExc(0)) Then
    End If
    'Modified by Lydia 2017/11/22 +國內
    If Val(strExc(0)) > 0 Then cmdOK(6).Caption = "國內A4名條 (" & Val(strExc(0)) & ")"
   'end 2016/11/04
   
   'Add By Sindy 2015/8/4
   SeekPrintL = Printer.Orientation
   'Mark by Lydia 2024/03/13
   'PUB_SetPrinter Me.Name, Me.Combo1, , , SeekPrint, , , True  'Modified by Moragn 2021/6/23 +只顯示有效的印表機參數
   '2015/8/4 END
   
   'Modify By Sindy 2017/1/23 智權同仁暫不能查看原始檔
   If Left(Pub_StrUserSt03, 1) = "S" Then
      cmdOK(14).Visible = False
   Else
      cmdOK(14).Visible = True '原始檔
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2015/8/4
   '若印表機或偏移值有變動, 則更新列印設定
   'Mark by Lydia 2024/03/13
   'If Me.Combo1.Text <> Me.Combo1.Tag Then
   '    PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, 0, 0, Me.Combo1.Text
   'End If
   'end 2024/03/13
   
   'Modified by Morgan 2021/6/23
   'Set Printer = Printers(SeekPrint)
   'Mark by Lydia 2024/03/13
   'PUB_RestorePrinter Combo1.List(SeekPrint)
   'end 2021/6/23
   'If SeekPrintL <> 0 Then
   '    Printer.Orientation = SeekPrintL
   'End If
   ''2015/8/4 END
   'end 2024/03/13
   
   Set frm100101_1 = Nothing
End Sub

Private Sub grdDataList_SelChange()
   bolSelData = True
   grdDataList.Visible = False
   grdDataList.row = grdDataList.MouseRow
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
   If grdDataList.Text = "V" Then
        grdDataList.Text = ""
        For i = 0 To grdDataList.Cols - 1
             grdDataList.col = i
             grdDataList.CellBackColor = QBColor(15)
       Next i
      'Add By Cheng 2002/03/15
       bolSelData = False
   Else
        grdDataList.Text = "V"
        For i = 0 To grdDataList.Cols - 1
            grdDataList.col = i
            grdDataList.CellBackColor = &HFFC0C0
        Next i
   
   End If
   End If
   grdDataList.Visible = True
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
           If Option1(0).Value = True Then
              Option1(0).Value = True
              Option1(1).Value = False
              Option1(2).Value = False
              Option1(3).Value = False
              Option1(4).Value = False
              Option1(5).Value = False 'Added by Morgan 2023/2/4
              If intTemp = False Then
                  txtSystem.SetFocus
              End If
              intTemp = False
           End If
      Case 1
           If Option1(1).Value = True Then
              Option1(1).Value = True
              Option1(0).Value = False
              Option1(2).Value = False
              Option1(3).Value = False
              Option1(4).Value = False
              Option1(5).Value = False 'Added by Morgan 2023/2/4
              Text1.SetFocus
           End If
      Case 2
           If Option1(2).Value = True Then
              Option1(2).Value = True
              Option1(1).Value = False
              Option1(0).Value = False
              Option1(3).Value = False
              Option1(4).Value = False
              Option1(5).Value = False 'Added by Morgan 2023/2/4
              Text5.SetFocus
           End If
      Case 3
           If Option1(3).Value = True Then
              Option1(3).Value = True
              Option1(1).Value = False
              Option1(2).Value = False
              Option1(0).Value = False
              Option1(4).Value = False
              Option1(5).Value = False 'Added by Morgan 2023/2/4
             Text2.SetFocus
           End If
      Case 4
           If Option1(4).Value = True Then
              Option1(4).Value = True
              Option1(1).Value = False
              Option1(2).Value = False
              Option1(3).Value = False
              Option1(0).Value = False
              Option1(5).Value = False 'Added by Morgan 2023/2/4
              Text3.SetFocus
           End If
      
      'Added by Morgan 2023/2/4
      Case 5
         If Option1(5).Value = True Then
              Option1(5).Value = True
              Option1(1).Value = False
              Option1(2).Value = False
              Option1(3).Value = False
              Option1(0).Value = False
              Option1(4).Value = False
              Text6.SetFocus
           End If
      'end 2023/2/4
      Case Else
   End Select
End Sub

Private Sub Text1_GotFocus()
   'Add by Morgan 2006/3/15
   '加輸入法控制
   'edit by nickc 2007/06/06
   'Text1.IMEMode = 2
   CloseIme
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   'Modify by Morgan 2004/4/15
   '需可輸入小寫
   'KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option1(1).Value = True
End Sub

Private Sub Text2_GotFocus()
   'Add by Morgan 2006/3/15
   '加輸入法控制
   'edit by nickc 2007/06/06
   'Text2.IMEMode = 2
   CloseIme
      
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2)
   
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option1(3).Value = True
End Sub

Private Sub Text3_GotFocus()
   'Add by Morgan 2006/3/15
   '加輸入法控制
   'edit by nickc 2007/06/06
   'Text3.IMEMode = 2
   CloseIme
   
   Text3.SelStart = 0
   Text3.SelLength = Len(Text3)
   
End Sub

'2010/3/15 cancel by sonia
'Private Sub Text3_KeyPress(KeyAscii As Integer)
'KeyAscii = UpperCase(KeyAscii)
'End Sub
'2010/3/15 end

Private Sub Text3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option1(4).Value = True
End Sub

Private Sub Text4_GotFocus()
   'Add by Morgan 2006/3/15
   '加輸入法控制
   'edit by nickc 2007/06/06
   'Text4.IMEMode = 2
   CloseIme
   
   Text4.SelStart = 0
   Text4.SelLength = Len(Text4)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_LostFocus()
   'Add By Cheng 2002/01/07
   'Modify By Cheng 2002/03/14
   'Me.Text4.Text = GetAllSysKind(Me.Text4)
End Sub

Private Sub Text5_GotFocus()
   'Add by Morgan 2006/3/15
   '加輸入法控制
   'edit by nickc 2007/06/06
   'Text5.IMEMode = 2
   CloseIme
   
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5)
   
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   'edit by nickc 2005/03/10 不鎖大小寫
   'KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option1(2).Value = True
End Sub


Private Sub Text6_GotFocus()
   CloseIme
   TextInverse Text6
End Sub

Private Sub Text6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option1(5).Value = True
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   'Add by Morgan 2006/3/15
   '加輸入法控制
   'edit by nickc 2007/06/06
   'txtCode(Index).IMEMode = 2
   CloseIme
   
   intTemp = True
   txtCode(Index).SelStart = 0
   txtCode(Index).SelLength = Len(txtCode(Index))
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   'If index = 0 Then KeyAscii = UpperCase(KeyAscii)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Option1(0).Value = True
End Sub

Private Sub txtSystem_GotFocus()
   'Add by Morgan 2006/3/15
   '加輸入法控制
   'edit by nickc 2007/06/06
   'txtSystem.IMEMode = 2
   CloseIme
   
'Modify By Cheng 2002/03/15
'If bolSelData = False Then
'   Option1(0).Value = True
'   txtSystem.SelStart = 0
'   txtSystem.SelLength = Len(txtSystem)
'Else


   If Option1(0).Value = True Then
      'Modified by Morgan 2012/10/17
      If txtSystem.Enabled = True Then
         txtSystem.SetFocus
         'Add By Cheng 2002/03/15
         txtSystem.SelStart = 0
         txtSystem.SelLength = Len(txtSystem)
      End If
   End If
   If Option1(1).Value = True Then
      Text1.SetFocus
      'Add By Cheng 2002/03/15
      Text1.SelStart = 0
      Text1.SelLength = Len(Text1)
   End If
   If Option1(2).Value = True Then
      Text5.SetFocus
      'Add By Cheng 2002/03/15
      Text5.SelStart = 0
      Text5.SelLength = Len(Text5)
   End If
   If Option1(3).Value = True Then
      Text2.SetFocus
      'Add By Cheng 2002/03/15
      Text2.SelStart = 0
      Text2.SelLength = Len(Text2)
   End If
   If Option1(4).Value = True Then
      Text3.SetFocus
      'Add By Cheng 2002/03/15
      Text3.SelStart = 0
      Text3.SelLength = Len(Text3)
   End If
   
   'Added by Morgan 2023/2/4
   If Option1(5).Value = True Then
      Text6.SetFocus
      Text6_GotFocus
   End If
   'end 2023/2/4

'End If
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

