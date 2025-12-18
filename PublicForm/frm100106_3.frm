VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100106_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "以期限管制日查詢"
   ClientHeight    =   5712
   ClientLeft      =   1908
   ClientTop       =   3108
   ClientWidth     =   9504
   ControlBox      =   0   'False
   FillColor       =   &H80000000&
   LinkTopic       =   "Form16"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5712
   ScaleWidth      =   9504
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度"
      Height          =   400
      Index           =   3
      Left            =   5850
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料"
      Height          =   400
      Index           =   2
      Left            =   4470
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   0
      Width           =   1365
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發E-Mail"
      Height          =   400
      Index           =   1
      Left            =   2625
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   5
      Left            =   8730
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   0
      Width           =   715
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      Default         =   -1  'True
      Height          =   400
      Index           =   4
      Left            =   7725
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印聯絡單"
      Height          =   400
      Index           =   0
      Left            =   1515
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   0
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "管制表"
      Height          =   400
      Index           =   6
      Left            =   780
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   0
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "電子檔"
      Height          =   400
      Index           =   7
      Left            =   60
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   0
      Width           =   720
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "重新查詢"
      Height          =   400
      Index           =   8
      Left            =   6765
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   0
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "管制備註"
      Height          =   400
      Index           =   9
      Left            =   3555
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   0
      Width           =   915
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00000000&
      Height          =   300
      ItemData        =   "frm100106_3.frx":0000
      Left            =   6000
      List            =   "frm100106_3.frx":001C
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   390
      Width           =   3195
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4785
      Left            =   60
      TabIndex        =   5
      Top             =   690
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   8446
      _Version        =   393216
      Cols            =   21
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   21
   End
   Begin VB.Label Label6 
      Caption         =   "以期限管制日查詢by本所期限"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4050
      TabIndex        =   9
      Top             =   5490
      Width           =   2505
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "P.S. 已排除核對已准專利和客戶提供文件"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   90
      TabIndex        =   8
      Top             =   5520
      Width           =   3195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "顏色符號說明："
      Height          =   180
      Left            =   4710
      TabIndex        =   7
      Top             =   450
      Width           =   1260
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   1
      Left            =   3168
      TabIndex        =   4
      Top             =   444
      Width           =   1092
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "迄"
      Height          =   180
      Left            =   2805
      TabIndex        =   3
      Top             =   450
      Width           =   180
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   0
      Left            =   1488
      TabIndex        =   2
      Top             =   444
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "起"
      Height          =   180
      Left            =   1128
      TabIndex        =   1
      Top             =   444
      Width           =   252
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Width           =   900
   End
End
Attribute VB_Name = "frm100106_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/05/17 Form2.0已修改: grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/13 日期欄已修改
'重整 by Morgan 2006/2/7
'Modified by Morgan 2024/10/8 ServerDate->strSrvDate(1)
Option Explicit

Dim strSql As String, i As Integer, j As Integer, s As Integer, k As Integer, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim strTemp As String, strTemp1 As String, strTemp2 As String
Dim StrS As String, StrSTemp As Variant, strTemp3 As String, StrTemp4 As String
Dim StrTag As String, StrTempCP27 As String, intK As Integer, StrTemp8 As Variant
Dim StrToMail(6) As String, StrChineseName As String, StrEnglishName As String, StrJanpenName As String
Dim StrTest4 As String
Dim StrR03001 As String
Dim StrR03002 As String
Dim StrR03003 As String
Dim StrR03004 As String
Dim StrR03005 As String
Dim StrR03006 As String
Dim StrR03007 As String
Dim StrR03008 As String
Dim StrR03009 As String
Dim StrR03010 As String
Dim StrR03011 As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Add By Cheng 2003/04/30
Public m_strFACUData  As String '判斷是否要列印代理人及客戶資料 'Modify by Amy 2016/07/18 改public
Dim strTempA(0 To 23) As String, PLeft(0 To 23) As Integer 'Modify by Amy 2014/07/11 +(23) 存客戶案件案號
Dim iPrint As Integer, Page As Integer
Dim SeekTmp(0 To 1) As String
'Add By Cheng 2003/05/02
Dim m_blnExportFile As Boolean '是否產生電子檔 'Modify by Amy 2016/07/18 改public
'Add By Cheng 2003/05/23
Public m_blnSales As Boolean '使用者為智權人員等級
Dim m_strSalesNo As String '智權人員代號
Public m_NpCon As String '下一程序控制條件
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序 Add By Sindy 2014/4/2
Dim strOldApply As String '申請人編號 Add by Amy 2016/07/18
Private Const m_Bottom As Integer = 10300 'Added by Lydia 2018/02/14 預設明細列最大Y位置
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件
'Added by Lydia 2021/05/19 Grid欄位的位置
Dim colCp09 As Integer, colCP10 As Integer, colCp10Name As Integer, colCaseNo As Integer, colCaseName As Integer
Dim colCP13 As Integer, colCP01 As Integer, colCP66 As Integer, colCP57 As Integer
Dim colDDate1 As Integer, colDDate2 As Integer, colDDate3 As Integer  'DDate1=所限/承辦期限、DDate2=法限/所限、DDate3=延期日
Dim colDDate4 As Integer 'Added by Morgan 2024/10/8 指定日期
Dim colSalesName As Integer, colCP27 As Integer
Dim colCP05 As Integer, colPA10name As Integer, colPA11 As Integer, colCP14name As Integer, colCP14 As Integer


Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   Me.Combo1.ListIndex = 0
   '92.04.16 nick
   cmdState = -1
   '設定操作功能權限
   Me.cmdOK(6).Enabled = IsUserHasRightOfFunction("frm100106_3", strPrint, False)
   Me.cmdOK(7).Enabled = IsUserHasRightOfFunction("frm100106_1", strPrint, False)
   
   'Add By Sindy 2013/8/13
   If frm100106_1.opt1(3).Value = True Then
      Label1.Caption = "承辦期限："
      LBL1(0).Caption = frm100106_1.txt6(0)
      LBL1(1).Caption = IIf(Len(frm100106_1.txt6(1).Text) > 0, frm100106_1.txt6(1).Text, strSrvDate(1) - 19110000)
   Else
      Label1.Caption = "本所期限："
      'Modify by Amy 2016/07/18 有點選才秀,否則查了本所期限再用本所案號查也會顯示
      If frm100106_1.opt1(0).Value = True Then
        LBL1(0).Caption = frm100106_1.txt1(0)
        LBL1(1).Caption = IIf(Len(frm100106_1.txt1(1).Text) > 0, frm100106_1.txt1(1).Text, strSrvDate(1) - 19110000)
      End If
   End If
   '2013/8/13 END
   m_blnColOrderAsc = True 'Add By Sindy 2014/4/2
End Sub

'Modified by Lydia 2021/05/19 +Optional ByVal bolReset As Boolean = True
Private Sub SetDataListWidth(Optional ByVal bolReset As Boolean = True)
   
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer, iCol As Integer
   
   'Modify by Amy 2016/07/18 +申請人no(ApplyNo)
   'Modified by Lydia 2019/11/01 +申請人1~5(cust01~cust05),FC代理人;
   'arrGridHeadText = Array("V", "本所期限", "本所案號", "分所號", "案件名稱" _
                     , "案件性質", "承辦人", "智權人員", "收文日", "法定期限", "進度備註" _
                     , "申請人", "是否出名", "延期日", "申請國家", "申請人國籍" _
                     , "發文日", "代理人", "彼所案號", "申請案號", "承辦人備註" _
                     , "取消收文日", "", "", "", "", "", "", "ApplyNo")
   'Modified by Lydia 2021/05/19 依前一畫面選項，改欄位標題
   'arrGridHeadText = Array("V", "本所期限", "本所案號", "分所號", "案件名稱", "案件性質" _
                     , "承辦人", "智權人員", "收文日", "法定期限", "進度備註" _
                     , "申請人", "是否出名", "延期日", "申請國家", "申請人國籍" _
                     , "發文日", "代理人", "彼所案號", "申請案號", "承辦人備註" _
                     , "取消收文日", "NP10", "NP01", "NP22", "CP14" _
                     , "NP02", "NP07", "ApplyNo", "CUST01", "CUST02" _
                     , "CUST03", "CUST04", "CUST05", "FCNO")
   'Added by Morgan 2024/1/8 +指定日期
   'Modified by Lydia 2024/01/17 只有已收文未發文增加欄位 +And frm100106_1.OPT2(1).Value = True
   If strSrvDate(1) >= 指定日期啟用日 And frm100106_1.opt2(1).Value = True Then
      arrGridHeadText = Split("V," & IIf(frm100106_1.opt1(3).Value = True, "承辦期限", "本所期限") & IIf(frm100106_1.opt1(0).Value = True, ",指定日期", "") & ",本所案號,分所號,案件名稱,案件性質" & _
                     ",承辦人,智權人員,收文日," & IIf(frm100106_1.opt1(3).Value = True, "本所期限", "法定期限") & "," & IIf(frm100106_1.opt2(0).Value = True, "備註", "進度備註") & _
                     ",申請人,是否出名,延期日,申請國家,申請人國籍,發文日,代理人,彼所案號,申請案號,承辦人備註,取消收文日,NP10,NP01,NP22,CP14,NP02,NP07,ApplyNo,CUST01,CUST02,CUST03,CUST04,CUST05,FCNO", ",")
   Else
   'end 2024/1/8
      arrGridHeadText = Array("V", IIf(frm100106_1.opt1(3).Value = True, "承辦期限", "本所期限"), "本所案號", "分所號", "案件名稱", "案件性質" _
                     , "承辦人", "智權人員", "收文日", IIf(frm100106_1.opt1(3).Value = True, "本所期限", "法定期限"), IIf(frm100106_1.opt2(0).Value = True, "備註", "進度備註") _
                     , "申請人", "是否出名", "延期日", "申請國家", "申請人國籍" _
                     , "發文日", "代理人", "彼所案號", "申請案號", "承辦人備註" _
                     , "取消收文日", "NP10", "NP01", "NP22", "CP14" _
                     , "NP02", "NP07", "ApplyNo", "CUST01", "CUST02" _
                     , "CUST03", "CUST04", "CUST05", "FCNO")
   End If
   'Memo by Lydia 2021/05/19   ApplyNo = 顯示的申請人編號
   If bolFNation = False Then
      'Modified by Lydia 2019/11/01 +申請人1~5(cust01~cust05),FC代理人
      'Added by Morgan 2024/1/8 +指定日期
      'Modified by Morgan 2024/4/19 +已收文未發文條件
      If strSrvDate(1) >= 指定日期啟用日 And frm100106_1.opt2(1).Value = True Then
         'Modified by Morgan 2025/10/20 欄位數要與arrGridHeadText相同，否則下面跑迴圈會當
         'arrGridHeadWidth = Split("200, 810" & IIf(frm100106_1.opt1(0).Value = True, ", 810", "") & ", 1300, 0, 800, 800, 750, 750, 810, 810, 810, 1000, 800, 810, 800, 1000, 810, 0, 800, 800, 1000, 1035, 0, 0, 0, 0, 0, 0, 0, 0", ",")
         arrGridHeadWidth = Split("200, 810" & IIf(frm100106_1.opt1(0).Value = True, ", 810", "") & ", 1300, 0, 800, 800, 750, 750, 810, 810, 810, 1000, 800, 810, 800, 1000, 810,0, 800, 800, 1000, 1035, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0", ",")
         'end 2025/10/20
      Else
      'end 2024/1/8
         arrGridHeadWidth = Array(200, 810, 1300, 0, 800, 800 _
                        , 750, 750, 810, 810, 810 _
                        , 1000, 800, 810, 800, 1000 _
                        , 810, 0, 800, 800, 1000 _
                        , 1035, 0, 0, 0, 0 _
                        , 0, 0, 0, 0, 0 _
                        , 0, 0, 0, 0)
      End If
      
   Else   'Memo by Lydia 2021/05/19 顯示代理人欄位
      'Modified by Lydia 2019/11/01 +申請人1~5(cust01~cust05),FC代理人
      'Added by Morgan 2024/1/8 +指定日期
      'Modified by Morgan 2024/4/19 +已收文未發文條件
      If strSrvDate(1) >= 指定日期啟用日 And frm100106_1.opt2(1).Value = True Then
         arrGridHeadWidth = Split("200, 810" & IIf(frm100106_1.opt1(0).Value = True, ", 810", "") & ", 1300, 0, 800, 800, 750, 750, 810, 810, 810, 1000, 800, 810, 800, 1000, 810,810, 800, 800, 1000, 1035, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0", ",")
      Else
      'end 2024/1/8
         arrGridHeadWidth = Array(200, 810, 1300, 0, 800, 800 _
                        , 750, 750, 810, 810, 810 _
                        , 1000, 800, 810, 800, 1000 _
                        , 810, 810, 800, 800, 1000 _
                        , 1035, 0, 0, 0, 0 _
                        , 0, 0, 0, 0, 0 _
                        , 0, 0, 0, 0)
      End If
      
   End If
   'end 2016/07/18
   'Added by Lydia 2021/05/19
   If bolReset = True Then
         grdDataList.Clear
         grdDataList.Rows = 2
   End If
   'end 2021/05/19
   
   grdDataList.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grdDataList.Cols - 1
      'edit by nick 2004/07/09
      grdDataList.row = 0
      grdDataList.col = iRow
      grdDataList.Text = arrGridHeadText(iRow)
      grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grdDataList.CellAlignment = flexAlignCenterCenter
   Next
    Dim iDep As String
    iDep = PUB_GetST06(strUserNum)
    '電腦中心，跟分所才秀
    'Modified by Morgan 2024/1/8 欄位調整,改以名稱設定
    'iCol = 3
    iCol = PUB_MGridGetId("分所號", grdDataList)
    'end 2024/1/8
    If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
        grdDataList.ColWidth(iCol) = 0
    Else
        grdDataList.ColWidth(iCol) = 620
    End If
   
   'Added by Lydia 2021/05/19 取得Grid欄位的位置
   If colCp09 = 0 Then
        colCp09 = PUB_MGridGetId("NP01", grdDataList) '收文號 CP09 / NP01
        colCP66 = PUB_MGridGetId("NP22", grdDataList) 'CP66 / NP22
        colCP10 = PUB_MGridGetId("NP07", grdDataList) 'CP10 / NP07
        colCp10Name = PUB_MGridGetId("案件性質", grdDataList)
        colCaseNo = PUB_MGridGetId("本所案號", grdDataList)
        colCaseName = PUB_MGridGetId("案件名稱", grdDataList)
        colCP13 = PUB_MGridGetId("NP10", grdDataList)  'CP13 / NP10
        colCP01 = PUB_MGridGetId("NP02", grdDataList) '系統別
        
        If frm100106_1.opt1(3).Value = True Then
            colDDate1 = PUB_MGridGetId("承辦期限", grdDataList)
            colDDate2 = PUB_MGridGetId("本所期限", grdDataList)
        Else
            colDDate1 = PUB_MGridGetId("本所期限", grdDataList)
            colDDate2 = PUB_MGridGetId("法定期限", grdDataList)
            colDDate4 = PUB_MGridGetId("指定日期", grdDataList) 'Added by Morgan 2024/10/8
        End If
        colDDate3 = PUB_MGridGetId("延期日", grdDataList)
        colSalesName = PUB_MGridGetId("智權人員", grdDataList)
        colCP05 = PUB_MGridGetId("收文日", grdDataList)
        colCP27 = PUB_MGridGetId("發文日", grdDataList)
        colCP57 = PUB_MGridGetId("取消收文日", grdDataList)
        colPA10name = PUB_MGridGetId("申請國家", grdDataList)
        colPA11 = PUB_MGridGetId("申請案號", grdDataList)
        colCP14 = PUB_MGridGetId("CP14", grdDataList)
        colCP14name = PUB_MGridGetId("承辦人", grdDataList)
   End If
   'end 2021/05/19
End Sub

Public Sub PubShowNextData()
   Dim strFileName As String '電子檔檔名
   
   On Error GoTo ErrorHandler
   Select Case cmdState
   Case 0 '列印聯絡單
        Me.Enabled = False
        StrTag = ""
        For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
           grdDataList.col = 0
           grdDataList.Text = ""
           grdDataList.col = 0
           grdDataList.CellBackColor = QBColor(15)
           'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseno
            grdDataList.col = colCaseNo
            StrTag = grdDataList.Text
            If Len(Trim(StrTag)) <> 0 Then
               'Modify by Morgan 2004/8/16
               'If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Then
               If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Or Left(StrTag, 1) = "!" Or Left(StrTag, 1) = "&" Or Left(StrTag, 1) = "N" Or Left(StrTag, 1) = "x" Then
                  StrTag = Right(StrTag, Len(StrTag) - 1)
                  StrToMail(1) = StrTag
               End If
               StrR03004 = StrTag
               Call StrMenu2
               If Not IsNull(StrTag) Then
                   If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                   End If
                  'Modified by Lydia 2021/05/19 改用變數取得: 8=>colCP05 (收文日)
                  grdDataList.col = colCp09
                  StrR03008 = grdDataList.Text
                  'Modified by Lydia 2021/05/19 改用變數取得: 5=>colCp10Name (案件性質)
                  grdDataList.col = colCp10Name
                  StrR03009 = grdDataList.Text
                  'Modified by Lydia 2021/05/19 改用變數取得: 1=>colDDate2
                  grdDataList.col = colDDate2
                  StrR03011 = grdDataList.Text
                  'Modified by Lydia 2021/05/19 改用變數取得: 9=>colDDate1
                  grdDataList.col = colDDate1
                  StrR03010 = grdDataList.Text
                  'Modified by Lydia 2021/05/19 改用變數取得: 22=>colCP13
                  grdDataList.col = colCP13
                  StrR03001 = grdDataList.Text
                  'Modified by Lydia 2021/05/19 改用變數取得: 23=>colCP09
                  grdDataList.col = colCp09
                  StrR03003 = grdDataList.Text
                  StrR03005 = StrChineseName
                  StrR03006 = StrEnglishName
                  StrR03007 = StrJanpenName
                  Call StrMenu3
                  Screen.MousePointer = vbHourglass
                  frm100106_5.Show
                  frm100106_5.Tag = StrTag
                  frm100106_5.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
            End If
        End If
        Next i
        Me.Enabled = True
   Case 1 '發E-Mail
        Me.Enabled = False
        StrTag = ""
        For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
           grdDataList.col = 0
           grdDataList.Text = ""
           grdDataList.col = 0
           grdDataList.CellBackColor = QBColor(15)
            
            'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseNo
            grdDataList.col = colCaseNo
            StrTag = grdDataList.Text
            If Len(Trim(StrTag)) <> 0 Then
               StrToMail(1) = StrTag
               'edit by nick 2004/08/12
               'If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Then
               'Modify by Morgan 2004/8/16
               'If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Or Left(StrTag, 1) = "!" Then
               If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Or Left(StrTag, 1) = "!" Or Left(StrTag, 1) = "&" Or Left(StrTag, 1) = "N" Or Left(StrTag, 1) = "x" Then
                  StrTag = Right(StrTag, Len(StrTag) - 1)
                  StrToMail(1) = StrTag
               End If
               Call StrMenu2
               If Not IsNull(StrTag) Then
                   If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                   End If
                  'Modify by Morgan 2007/1/12 欄位位置有變
                  'Modified by Lydia 2021/05/19 改用變數取得
                  'StrToMail(2) = grdDataList.TextMatrix(i, 4)
                  'StrToMail(3) = grdDataList.TextMatrix(i, 8)
                  'StrToMail(4) = grdDataList.TextMatrix(i, 5)
                  'StrToMail(5) = grdDataList.TextMatrix(i, 9)
                  'StrToMail(6) = grdDataList.TextMatrix(i, 1)
                  StrToMail(2) = grdDataList.TextMatrix(i, colCaseName)
                  StrToMail(3) = grdDataList.TextMatrix(i, colCP05)
                  StrToMail(4) = grdDataList.TextMatrix(i, colCp10Name)
                  StrToMail(5) = grdDataList.TextMatrix(i, colDDate2)
                  StrToMail(6) = grdDataList.TextMatrix(i, colDDate1)
                  'end 2021/05/19
                  'Modify by Morgan 2005/9/7 本所與法定抓反了
                  '..."           本所期限：" + StrToMail(5) + Space(30 - Len(StrToMail(5))) + "法定期限：" + StrToMail(6) + vbCrLf + vbCrLf
                  strExc(1) = "           本所案號：" + StrToMail(1) + vbCrLf + vbCrLf + _
                        "           案件名稱：" + StrToMail(2) + vbCrLf + vbCrLf
                  'Modify by Morgan 2008/12/24 未收文的不印收文日費用等資料
                  If frm100106_1.opt2(0).Value = True Then
                     strExc(1) = strExc(1) & _
                        "           案件性質：" + StrToMail(4) + vbCrLf + vbCrLf
                  Else
                     strExc(1) = strExc(1) & "           收 文 日：" + StrToMail(3) + Space(30 - Len(StrToMail(3))) + "案件性質：" + StrToMail(4) + vbCrLf + vbCrLf
                     'Remove by Morgan 2008/12/24 既然是空的就不必印
                     '"           費    用：" + "                     " + vbTab + vbTab + vbTab + "規費：" + "                     " + vbTab + vbTab + vbTab + "點數：" + "" + vbCrLf + vbCrLf
                  End If
                  'Add By Sindy 2013/8/13
                  If frm100106_1.opt1(3).Value = True Then
                     strExc(1) = strExc(1) & _
                           "           承辦期限：" + StrToMail(6) + Space(30 - Len(StrToMail(6))) + "本所期限：" + StrToMail(5) + vbCrLf + vbCrLf
                  Else
                  '2013/8/13 END
                     strExc(1) = strExc(1) & _
                           "           本所期限：" + StrToMail(6) + Space(30 - Len(StrToMail(6))) + "法定期限：" + StrToMail(5) + vbCrLf + vbCrLf
                  End If
                  frm100106_4.txt1(1) = strExc(1)
                  '2005/9/7 end
                  Screen.MousePointer = vbHourglass
                  frm100106_4.Show
                  frm100106_4.strCaseNo = StrToMail(1) 'Added by Lydia 2020/05/18 傳入本所案號
                  frm100106_4.Tag = StrTag
                  'Add by Morgan 2007/1/8
                  'Modified by Lydia 2021/05/19 改用變數取得: 23=> colCP09
                  frm100106_4.strCP09 = grdDataList.TextMatrix(i, colCp09)
                  frm100106_4.StrMenu
                  'end 2007/1/8
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
            End If
        End If
        Next i
        Me.Enabled = True
   Case 2 '案件基本資料
         Me.Enabled = False
         For i = 1 To grdDataList.Rows - 1
         grdDataList.col = 0
         grdDataList.row = i
         If Trim(grdDataList.Text) = "V" Then
           grdDataList.col = 0
           grdDataList.Text = ""
           grdDataList.col = 0
           grdDataList.CellBackColor = QBColor(15)
           Dim Str01 As String
           'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseNo
           grdDataList.col = colCaseNo
           StrTag = grdDataList.Text
           'edit by nick 2004/08/12
           'If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Then
           'Modify by Morgan 2004/8/16
           'If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Or Left(StrTag, 1) = "!" Then
           If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Or Left(StrTag, 1) = "!" Or Left(StrTag, 1) = "&" Or Left(StrTag, 1) = "N" Or Left(StrTag, 1) = "x" Then
              StrTag = Right(StrTag, Len(StrTag) - 1)
           End If
           Str01 = SystemNumber(StrTag, 1)
           If Mid(UCase(Str01), 1, 1) = "N" Then
               Str01 = Mid(Str01, 2, 3)
           End If
           If Not IsNull(grdDataList.Text) Then
               If fnSaveParentForm(Me) = False Then
                  Me.Enabled = True
                  Exit Sub
               End If
               Select Case Pub_RplStr(Str01)
               Case "CFP", "FCP", "P"   '專利
                     Screen.MousePointer = vbHourglass
                     frm100101_3.Show
                     frm100101_3.Tag = StrTag
                     frm100101_3.StrMenu
                     Screen.MousePointer = vbDefault
               Case "CFT", "FCT", "T", "TF"   '商標
                     Screen.MousePointer = vbHourglass
                     frm100101_4.Show
                     frm100101_4.Tag = StrTag
                     frm100101_4.StrMenu
                     Screen.MousePointer = vbDefault
               'Modify By Sindy 2009/07/24 增加LIN系統類別
               'modify by sonia 2019/7/29 +ACS系統類別
               Case "CFL", "FCL", "L", "LIN", "ACS" '法務
                     Screen.MousePointer = vbHourglass
                     frm100101_5.Show
                     frm100101_5.Hide
                     frm100101_5.Tag = StrTag
                     frm100101_5.StrMenu
                     Screen.MousePointer = vbDefault
               Case "LA"            '顧問
                     Screen.MousePointer = vbHourglass
                     frm100101_6.Show
                     frm100101_6.Tag = StrTag
                     frm100101_6.StrMenu
                     Screen.MousePointer = vbDefault
               Case Else                  '服務
                    Select Case Pub_RplStr(Str01)
                        Case "TB"    '條碼
                            Screen.MousePointer = vbHourglass
                            frm100101_7.Show
                            frm100101_7.Tag = StrTag
                            frm100101_7.StrMenu
                            Screen.MousePointer = vbDefault
                        Case "TM"
                            Screen.MousePointer = vbHourglass
                            frm100101_8.Show
                            frm100101_8.Tag = StrTag
                            frm100101_8.StrMenu
                            Screen.MousePointer = vbDefault
                        Case "TD"
                            Screen.MousePointer = vbHourglass
                            frm100101_9.Show
                            frm100101_9.Tag = StrTag
                            frm100101_9.StrMenu
                            Screen.MousePointer = vbDefault
                        Case "TC", "CFC"
                            Screen.MousePointer = vbHourglass
                            frm100101_A.Show
                            frm100101_A.Tag = StrTag
                            frm100101_A.StrMenu
                            Screen.MousePointer = vbDefault
                        Case Else
                            Screen.MousePointer = vbHourglass
                            frm100101_B.Show
                            frm100101_B.Tag = StrTag
                            frm100101_B.StrMenu
                            Screen.MousePointer = vbDefault
                     End Select
               End Select
               Me.Enabled = True
               Exit Sub
           End If
        End If
        Next i
        Me.Enabled = True
   Case 3 '案件進度
        Me.Enabled = False
        StrTag = ""
        For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
           grdDataList.col = 0
           grdDataList.Text = ""
           grdDataList.col = 0
           grdDataList.CellBackColor = QBColor(15)
            'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseNo
            grdDataList.col = colCaseNo
            StrTag = grdDataList.Text
            'edit by nick 2004/08/12
            'If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Then
            'Modify by Morgan 2004/8/16
            'If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Or Left(StrTag, 1) = "!" Then
            If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Or Left(StrTag, 1) = "!" Or Left(StrTag, 1) = "&" Or Left(StrTag, 1) = "N" Or Left(StrTag, 1) = "x" Then
               StrTag = Right(StrTag, Len(StrTag) - 1)
            End If
            If Not IsNull(StrTag) Then
               If fnSaveParentForm(Me) = False Then
                  Me.Enabled = True
                  Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               frm100101_2.Show
               frm100101_2.Tag = StrTag
               frm100101_2.StrMenu
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
        End If
        Next i
        Me.Enabled = True
   Case 4 '回前畫面
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Case 5 '結束
        fnCloseAllFrm100
   Case 6 '列印管制表
       If Me.grdDataList.Rows > 1 Then
           m_strFACUData = MsgBox("是否列印代理人及客戶資料???", vbExclamation + vbYesNo + vbDefaultButton2)
           m_blnSales = IIf(UCase(Left(GetST05(strUserNum), 1)) = "S", True, False)
           PrintData
       Else
           MsgBox "無資料可列印!!!", vbExclamation + vbOKOnly
       End If
       m_blnSales = False
   Case 7 '產生管制表電子檔
       m_blnExportFile = False
       If Me.grdDataList.Rows > 1 Then
           strFileName = InputBox("請輸入欲匯出的檔案名稱(不可指定路徑)？" & vbCrLf & vbCrLf & "檔案存放位置為==> My Documents 或 我的文件夾 " & vbCrLf & vbCrLf & "(註：檔案名稱若重複，會被覆蓋！)")
           If Trim("" & strFileName) <> "" Then
               m_blnExportFile = True
               Screen.MousePointer = vbHourglass
               Open GetMyDocPath & "\" & strFileName & ".doc" For Append As #1
               m_strFACUData = vbYes
               m_blnSales = IIf(UCase(Left(GetST05(strUserNum), 1)) = "S", True, False)
               PrintData
               Close #1
               Screen.MousePointer = vbDefault
               MsgBox "檔案 " & strFileName & ".doc 匯出成功!!!", vbExclamation + vbOKOnly
           Else
               MsgBox "您已取消作業或未輸入欲匯出的檔案名稱!!!", vbExclamation + vbOKOnly
           End If
       Else
           MsgBox "無資料可產生電子檔!!!", vbExclamation + vbOKOnly
       End If
       m_blnExportFile = False
       m_blnSales = False
   'Add by Amy 2018/09/19 重新查詢
   Case 8
      StrMenu
   'Added by Lydia 2021/05/19
   Case 9 '管制備註
         Me.Enabled = False
         StrTag = ""
         strExc(5) = "" '暫時記錄無權限的本所案號
         strExc(6) = "": strExc(7) = "": strExc(8) = ""
         For i = 1 To grdDataList.Rows - 1
            grdDataList.col = 0
            grdDataList.row = i
            'Memo by Lydia 2021/05/28 保留單筆的寫法
'            If Trim(grdDataList.Text) = "V" Then
'                grdDataList.col = 0
'                grdDataList.Text = ""
'                grdDataList.col = 0
'                grdDataList.CellBackColor = QBColor(15)
'                grdDataList.col = colCaseNo
'                StrTag = grdDataList.Text
'                If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Or Left(StrTag, 1) = "!" Or Left(StrTag, 1) = "&" Or Left(StrTag, 1) = "N" Or Left(StrTag, 1) = "x" Then
'                   StrTag = Right(StrTag, Len(StrTag) - 1)
'                End If
'                If Not IsNull(StrTag) Then
'                  Call ChgCaseNo(Replace(StrTag, "-", ""), strExc)
'                  '控制共同查詢是否有跨部門查詢案件明細權限
'                  If Len(strExc(2)) <> 6 Then
'                        Me.Enabled = True
'                        Exit Sub
'                  End If
'                  If CheckSR09(strUserNum, strExc(1), "Y", , strExc(1), strExc(2), strExc(3), strExc(4)) = False Then
'                        Me.Enabled = True
'                        Exit Sub
'                  End If
'                   Screen.MousePointer = vbHourglass
'                   Call frm100123_3.SetParent(Me, grdDataList.TextMatrix(i, colCP09), grdDataList.TextMatrix(i, colCP66))
'                   frm100123_3.Show
'                   Me.Hide
'                   Screen.MousePointer = vbDefault
'                   Me.Enabled = True
'                   Exit Sub
'                End If
'            End If
            'end 2021/05/28
            If Trim(grdDataList.Text) = "V" Then
                 grdDataList.col = colCaseNo
                 StrTag = grdDataList.Text
                 If Left(StrTag, 1) = "*" Or Left(StrTag, 1) = "#" Or LCase(Left(StrTag, 1)) = "v" Or Left(StrTag, 1) = "!" Or Left(StrTag, 1) = "&" Or Left(StrTag, 1) = "N" Or Left(StrTag, 1) = "x" Then
                    StrTag = Right(StrTag, Len(StrTag) - 1)
                 End If
                 If Not IsNull(StrTag) Then
                    Call ChgCaseNo(Replace(StrTag, "-", ""), strExc)
                   '控制共同查詢是否有跨部門查詢案件明細權限
                    If Len(strExc(2)) <> 6 Then
                          Me.Enabled = True
                          Exit Sub
                    End If
                    If CheckSR09(strUserNum, strExc(1), "Y", False, strExc(1), strExc(2), strExc(3), strExc(4)) = False Then
                          strExc(5) = strExc(5) & vbCrLf & strExc(1) & "-" & strExc(2) & "-" & strExc(3) & "-" & strExc(4)
                    Else
                          strExc(6) = strExc(6) & "," & Format(i, "000")
                          strExc(7) = strExc(7) & "," & Trim("" & grdDataList.TextMatrix(i, colCp09)) '總收文號
                          strExc(8) = strExc(8) & "," & grdDataList.TextMatrix(i, colCP66)
                    End If
                 End If
            End If
         Next i
         If strExc(5) <> "" Then
             MsgBox "您沒有查詢下列案件明細的權限：" & vbCrLf & strExc(5)
         End If
         If strExc(6) <> "" Then
             strExc(6) = Mid(strExc(6), 2)
             strExc(7) = Mid(strExc(7), 2)
             strExc(8) = Mid(strExc(8), 2)
             Call frm100123_3.SetParent(Me, strExc(6), strExc(7), strExc(8))
             frm100123_3.Show
             Me.Hide
             Screen.MousePointer = vbDefault
         End If
         Me.Enabled = True
   'end 2021/05/19
   End Select
   
ErrorHandler:
   If Err.Number <> 0 Then
      m_blnExportFile = False
      m_blnSales = False
      MsgBox "(" & Err.Number & ")" & Err.Description
   End If
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdok_Click(Index As Integer)
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
End Sub

'Modify by Amy 2016/07/18 +參數
Public Function StrMenu(Optional ByVal bolMsg As Boolean = True) As Boolean
Dim StrSQLa As String
Dim strTmp As String
Dim strTmp1 As String
Dim stPA(1 To 4) As String, iPos As Integer, stCaseNo As String
'Add by Amy 2016/07/18 for 申請人1~5用
Dim strSQL10(1 To 4) As String, StrSql20(1 To 4) As String, StrSql30(1 To 4) As String, StrSql40(1 To 4) As String, StrSql50(1 To 4) As String
'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String
Dim dblRow As Double 'Add By Sindy 2025/9/3

   StrMenu = True
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/01/22 清除查詢印表記錄檔欄位
   StrS = ""
   '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
   StrSQLa = "DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人,"
   '#############################################    分三支程式作
   If frm100106_1.opt2(0).Value = True Then     '未收文
       pub_QL05 = pub_QL05 & ";" & frm100106_1.Frame1.Caption & frm100106_1.opt2(0).Caption 'Add By Sindy 2010/01/22
       strSQL1 = ""
       strSQL2 = ""
       StrSQL3 = ""
       StrSQL4 = ""
       strSQL5 = ""
        'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
        SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
        SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
        SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
        SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
        SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
        m_AllSys = IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(, frm100106_1.txt5(0).Text))
        'Added by Lydia 2020/03/11 外專人員勾選含FMP外專管制期限時,要另外增加P,PS
        If frm100106_1.Check1.Value = 1 And Left(Pub_StrUserSt03, 2) = "F2" Then
            If InStr(m_AllSys & ",", ",P,") = 0 Then m_AllSys = m_AllSys & ",P,"
            If InStr(m_AllSys & ",", ",PS,") = 0 Then m_AllSys = m_AllSys & ",PS,"
            m_AllSys = Replace(m_AllSys, ",,", ",")
        End If
        'end 2020/03/11
        intCufaCnt = 0
        'end 2019/11/01
        
      '以本所期限查詢
      If frm100106_1.opt1(0).Value Then
         If Len(Trim(frm100106_1.txt1(0))) <> 0 Then
            strSQL1 = strSQL1 + " AND NP08>=" & Val(ChangeTStringToWString(frm100106_1.txt1(0))) & " "
         End If
         pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(0).Caption & frm100106_1.txt1(0) 'Add By Sindy 2010/01/22
         If Len(Trim(frm100106_1.txt1(1))) <> 0 Then
            strSQL1 = strSQL1 + " AND NP08<=" & Val(ChangeTStringToWString(frm100106_1.txt1(1))) & " "
            pub_QL05 = pub_QL05 & "-" & frm100106_1.txt1(1) 'Add By Sindy 2010/01/22
         Else
            strSQL1 = strSQL1 + " AND NP08<=" & Val(ChangeTStringToWString(strSrvDate(1) - 19110000)) & " "
            pub_QL05 = pub_QL05 & "-" & (strSrvDate(1) - 19110000) 'Add By Sindy 2010/01/22
         End If
      End If
       '以本所案號查詢
       If frm100106_1.opt1(2).Value Then
         If Len(Trim(frm100106_1.txt3(0))) <> 0 Then
             strSQL1 = strSQL1 + " AND NP02='" & frm100106_1.txt3(0) & "' "
         End If
         pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(2).Caption & frm100106_1.txt3(0) 'Add By Sindy 2010/01/22
         If Len(Trim(frm100106_1.txt3(1))) <> 0 Then
             strSQL1 = strSQL1 + " AND NP03='" & frm100106_1.txt3(1) & "' "
         End If
         pub_QL05 = pub_QL05 & "-" & frm100106_1.txt3(1) 'Add By Sindy 2010/01/22
         If Len(Trim(frm100106_1.txt3(2))) <> 0 Then
             strSQL1 = strSQL1 + " AND NP04='" & frm100106_1.txt3(2) & "' "
             pub_QL05 = pub_QL05 & "-" & frm100106_1.txt3(2) 'Add By Sindy 2010/01/22
           Else
             strSQL1 = strSQL1 + " AND NP04='0' "
             pub_QL05 = pub_QL05 & "-" & "0" 'Add By Sindy 2010/01/22
         End If
         If Len(Trim(frm100106_1.txt3(3))) <> 0 Then
             strSQL1 = strSQL1 + " AND NP05='" & frm100106_1.txt3(3) & "' "
             pub_QL05 = pub_QL05 & "-" & frm100106_1.txt3(3) 'Add By Sindy 2010/01/22
           Else
             strSQL1 = strSQL1 + " AND NP05='00' "
             pub_QL05 = pub_QL05 & "-" & "00" 'Add By Sindy 2010/01/22
         End If
       End If
       If Len(Trim(frm100106_1.txt5(1))) <> 0 Then
            'Modify by Morgan 2007/9/27 也要抓外譯編號
            'strSQL1 = strSQL1 & " AND CP14='" & frm100106_1.txt5(1) & "' "
            strExc(1) = PUB_GetMapID(frm100106_1.txt5(1), 0)
            If strExc(1) <> "" Then
               strSQL1 = strSQL1 & " AND CP14 in ('" & frm100106_1.txt5(1) & "','" & strExc(1) & "')"
            Else
               strSQL1 = strSQL1 & " AND CP14='" & frm100106_1.txt5(1) & "' "
            End If
            'end 2007/9/27
            pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(2) & frm100106_1.txt5(1) & frm100106_1.LBL1(0) 'Add By Sindy 2010/01/22
       End If
       'Add by Morgan 2007/9/21 加承辦人組別
       If Len(Trim(frm100106_1.txt5(14))) <> 0 Then
            strSQL1 = strSQL1 & " AND exists(select * from staff SX where SX.ST01=NVL(SIM01,CP14) and SX.ST16='" & frm100106_1.txt5(14) & "')"
            pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(10) & frm100106_1.txt5(14)  'Add By Sindy 2010/01/22
       End If
       'end 2007/9/21
       If Len(Trim(frm100106_1.txt5(2))) <> 0 Then
            strSQL1 = strSQL1 & " AND S2.ST15>='" & frm100106_1.txt5(2) & "' "
       End If
       If Len(Trim(frm100106_1.txt5(3))) <> 0 Then
            strSQL1 = strSQL1 & " AND S2.ST15<='" & frm100106_1.txt5(3) & "' "
       End If
       
       
       'Added by Morgan 2012/5/23 FMP案改可選擇
       If frm100106_1.Check1.Value = 1 Then
            pub_QL05 = pub_QL05 & ";" & frm100106_1.Check1.Caption
       Else
            strSQL1 = strSQL1 & " AND NOT (NP02 IN ('P','PS','CFP','CPS') AND SUBSTR(S2.ST15,1,1)='F') "
       End If
       'end 2012/5/23
       
       If Len(Trim(frm100106_1.txt5(2))) <> 0 Or Len(Trim(frm100106_1.txt5(3))) <> 0 Then
            pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(3) & frm100106_1.txt5(2) & "-" & frm100106_1.txt5(3) 'Add By Sindy 2010/01/22
       End If
       If Len(Trim(frm100106_1.txt5(4))) <> 0 Then
          '2008/3/31 MODIFY BY SONIA 加控制查詢87027陳淑芳同時查20001台中所
          'strSQL1 = strSQL1 + " AND NP10='" & frm100106_1.txt5(4) & "' "
          If frm100106_1.txt5(4) = "87027" Then
             strSQL1 = strSQL1 + " AND NP10 IN ('87027','20001') "
          Else
             strSQL1 = strSQL1 + " AND NP10='" & frm100106_1.txt5(4) & "' "
          End If
          '2008/3/31 END
          pub_QL05 = pub_QL05 & ";" & frm100106_1.Label2(0) & frm100106_1.txt5(4) & frm100106_1.LBL1(1) 'Add By Sindy 2010/01/22
       End If
       '申請人國籍
       If Len(Trim(frm100106_1.txt5(9))) <> 0 Then
           strSQL1 = strSQL1 + " AND CU10>='" & frm100106_1.txt5(9) & "' "
       End If
       If Len(Trim(frm100106_1.txt5(10))) <> 0 Then
           strSQL1 = strSQL1 + " AND CU10<='" & frm100106_1.txt5(10) & "z' "
       End If
       If Len(Trim(frm100106_1.txt5(9))) <> 0 Or Len(Trim(frm100106_1.txt5(10))) <> 0 Then
            pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(1) & frm100106_1.txt5(9) & "-" & frm100106_1.txt5(10) 'Add By Sindy 2010/01/22
       End If
       strSQL1 = strSQL1 & m_NpCon 'Add by Morgan 2008/10/2
       strSQL1 = strSQL1 + " and NP06 IS NULL "
       strSQL2 = strSQL1
       StrSQL3 = strSQL1
       StrSQL4 = strSQL1
       strSQL5 = strSQL1
       If Len(Trim(frm100106_1.txt5(0))) <> 0 Then
         'Added by Morgan 2015/10/20
         '外專人員勾選含FMP外專管制期限時
         If frm100106_1.Check1.Value = 1 And Left(Pub_StrUserSt03, 2) = "F2" Then
            strSQL1 = strSQL1 & " AND (NP02 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 1) & ") or (NP02 IN ('P','CFP') AND SUBSTR(S2.ST15,1,1)='F')) "
            strSQL5 = strSQL5 & " AND (NP02 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 5) & ") or (NP02 IN ('PS','CPS') AND SUBSTR(S2.ST15,1,1)='F')) "
         Else
         'end 2015/10/20
            strSQL1 = strSQL1 & " AND NP02 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 1) & ") "
            strSQL5 = strSQL5 & " AND NP02 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 5) & ") "
         End If 'Added by Morgan 2015/10/20
         
            strSQL2 = strSQL2 & " AND NP02 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 2) & ") "
            StrSQL3 = StrSQL3 & " AND NP02 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 3) & ") "
            StrSQL4 = StrSQL4 & " AND NP02 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 4) & ") "
            pub_QL05 = pub_QL05 & ";" & Left(frm100106_1.Label1(4), 5) & frm100106_1.txt5(0) 'Add By Sindy 2010/01/22
       End If
       
       If Len(Trim(frm100106_1.txt4(2))) <> 0 Then
            strSQL1 = strSQL1 & " AND PA75>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
            strSQL2 = strSQL2 & " AND TM44>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
            StrSQL3 = StrSQL3 & " AND LC22>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
            strSQL5 = strSQL5 & " AND SP26>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
       End If
       If Len(Trim(frm100106_1.txt4(3))) <> 0 Then
            strSQL1 = strSQL1 & " AND PA75<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
            strSQL2 = strSQL2 & " AND TM44<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
            StrSQL3 = StrSQL3 & " AND LC22<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
            strSQL5 = strSQL5 & " AND SP26<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
       End If
       If Len(Trim(frm100106_1.txt4(2))) <> 0 Or Len(Trim(frm100106_1.txt4(3))) <> 0 Then
            pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(5) & frm100106_1.txt4(2) & "-" & frm100106_1.txt4(3)  'Add By Sindy 2010/01/22
       End If
       
'---------Move by  Lydia 2019/11/18 從If Len(Trim(frm100106_1.txt4(2))) <> 0 Then的上方移過來
       'Modify by Amy 2016/07/18 若有下申請人編號,則申請人1~5也要查
       If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            strSQL10(1) = strSQL1 & " AND PA27>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            strSQL10(2) = strSQL1 & " AND PA28>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            strSQL10(3) = strSQL1 & " AND PA29>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            strSQL10(4) = strSQL1 & " AND PA30>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            strSQL1 = strSQL1 & " AND PA26>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            
            StrSql20(1) = strSQL2 & " AND TM78>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql20(2) = strSQL2 & " AND TM79>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql20(3) = strSQL2 & " AND TM80>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql20(4) = strSQL2 & " AND TM81>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            strSQL2 = strSQL2 & " AND TM23>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            
            StrSql30(1) = StrSQL3 & " AND LC43>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql30(2) = StrSQL3 & " AND LC44>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql30(3) = StrSQL3 & " AND LC45>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql30(4) = StrSQL3 & " AND LC46>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSQL3 = StrSQL3 & " AND LC11>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            
            StrSql40(1) = StrSQL4 & " AND HC24>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql40(2) = StrSQL4 & " AND HC25>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql40(3) = StrSQL4 & " AND HC26>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql40(4) = StrSQL4 & " AND HC27>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSQL4 = StrSQL4 & " AND HC05>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            
            StrSql50(1) = strSQL5 & " AND SP58>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql50(2) = strSQL5 & " AND SP59>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql50(3) = strSQL5 & " AND SP65>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql50(4) = strSQL5 & " AND SP66>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            strSQL5 = strSQL5 & " AND SP08>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
       End If
       If Len(Trim(frm100106_1.txt4(1))) <> 0 Then
            strSQL1 = strSQL1 & " AND PA26<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            strSQL10(1) = strSQL10(1) & " AND PA27<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            strSQL10(2) = strSQL10(2) & " AND PA28<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            strSQL10(3) = strSQL10(3) & " AND PA29<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            strSQL10(4) = strSQL10(4) & " AND PA30<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            
            strSQL2 = strSQL2 & " AND TM23<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql20(1) = StrSql20(1) & " AND TM78<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql20(2) = StrSql20(2) & " AND TM79<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql20(3) = StrSql20(3) & " AND TM80<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql20(4) = StrSql20(4) & " AND TM81<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            
            StrSQL3 = StrSQL3 & " AND LC11<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql30(1) = StrSql30(1) & " AND LC11<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql30(2) = StrSql30(1) & " AND LC11<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql30(3) = StrSql30(1) & " AND LC11<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql30(4) = StrSql30(1) & " AND LC11<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            
            StrSQL4 = StrSQL4 & " AND HC05<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql40(1) = StrSql40(1) & " AND HC24<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql40(2) = StrSql40(2) & " AND HC25<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql40(3) = StrSql40(3) & " AND HC26<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql40(4) = StrSql40(4) & " AND HC27<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            
            strSQL5 = strSQL5 & " AND SP08<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql50(1) = StrSql50(1) & " AND SP58<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql50(2) = StrSql50(2) & " AND SP59<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql50(3) = StrSql50(3) & " AND SP65<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql50(4) = StrSql50(4) & " AND SP66<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
       End If
       'end 2016/07/18
       If Len(Trim(frm100106_1.txt4(0))) <> 0 Or Len(Trim(frm100106_1.txt4(1))) <> 0 Then
            pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(6) & frm100106_1.txt4(0) & "-" & frm100106_1.txt4(1) 'Add By Sindy 2010/01/22
       End If
'---------end 2019/11/18
      'Add by Morgan 2006/2/14 有下FCP管制人時只抓FCP資料
      If Len(Trim(frm100106_1.txt5(11).Text)) + Len(Trim(frm100106_1.txt5(12).Text)) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(8) & frm100106_1.txt5(11) & "-" & frm100106_1.txt5(12)   'Add By Sindy 2010/01/22
         'FCP管制人(抓代理人國籍的FCP管制人, 若無則抓申請人國籍的FCP管制人 )
         '2010/9/13 MODIFY BY SONIA 日期欄改百年日期排序問題
         'Modify by Amy 2016/07/18 若有下申請人編號,則申請人1~5也要查並顯示申請人資料
         'Modify by Amy 2018/09/20 申請人資料全顯示不限制字數 原:SUBSTRB(CU04,1,10)
         'Modified by Lydia 2019/11/01 +增加欄位SeColPA,SeColSP
         strSql = "SELECT '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
            "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日, DECODE(NP07,'411',NVL(N3.NA16,N2.NA16),NP10) NP10,NP01,np22,cp14, NP02, NP07,PA76,pa75,pa26 as ApplyNo " & _
             SeColPA & " FROM NEXTPROGRESS,CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
             " WHERE (PA57<>'Y' or pa57 is null) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14 AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL1 & _
             " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
             " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
             " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
             " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
          strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                  "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,DECODE(NP07,'411',NVL(N3.NA16,N2.NA16),NP10) NP10,NP01,np22,cp14, NP02, NP07,NULL,sp26,sp08 as ApplyNo " & _
                   SeColSP & " FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                   " WHERE (SP15<>'Y' or sp15 is null) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14 AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL5 & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
         'end 2018/09/20
         'end 2016/07/18
         '年費管制人順序:pa76->pa75->pa26.cu96->pa26(PA76可能是客戶編號)
         'Modify by Amy 2016/09/12 原PA26改為ApplyNo
         'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)
         'strSql = "SELECT Y.* FROM (" & _
            " select X.*,NVL(NVL(F1.FA10,C1.CU10),NVL(F2.FA10,NVL(F4.FA10,C2.CU10))) FA10" & _
            " from (" & strSql & ") X, FAGENT F1, FAGENT F2,FAGENT F4,CUSTOMER C1, CUSTOMER C2" & _
            " WHERE NP07='605' AND F1.FA01(+)=SUBSTR(PA76,1,8) AND F1.FA02(+)=SUBSTR(PA76||'0',9,1)" & _
            " AND C1.CU01(+)=SUBSTR(PA76,1,8) AND C1.CU02(+)=SUBSTR(PA76||'0',9,1)" & _
            " AND F2.FA01(+)=SUBSTR(PA75,1,8) AND F2.FA02(+)=SUBSTR(PA75||'0',9,1)" & _
            "" & _
            " AND C2.CU01(+)=SUBSTR(ApplyNo,1,8) AND C2.CU02(+)=SUBSTR(ApplyNo||'0',9,1)" & _
            " AND F4.FA01(+)=SUBSTR(C2.CU96,1,8) AND F4.FA02(+)=SUBSTR(C2.CU96||'0',9,1)" & _
            " Union All" & _
            " select X.*,NVL(FA10,CU10) FA10" & _
            " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
            " WHERE NP07<>'605' AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
            " AND CU01(+)=SUBSTR(ApplyNo,1,8) AND CU02(+)=SUBSTR(ApplyNo||'0',9,1)" & _
            ") Y, NATION WHERE NA01(+)=FA10"
         strSql = "SELECT Y.* FROM (" & _
            " select X.*,NVL(FA10,CU10) FA10" & _
            " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
            " WHERE FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
            " AND CU01(+)=SUBSTR(ApplyNo,1,8) AND CU02(+)=SUBSTR(ApplyNo||'0',9,1)" & _
            ") Y, NATION WHERE NA01(+)=FA10"
         'end 2020/5/12
         'Modify by Morgan 2006/2/14 從外層移進來並修改
         If Len(Trim(frm100106_1.txt5(11).Text)) <> 0 Then
              'Modified by Lydia 2023/09/11 區分FMP案管制人---by Phoebe
              'strSql = strSql & " AND NA16 >='" & frm100106_1.txt5(11).Text & "' "
              strSql = strSql & " and (((np02='FCP' or np02='FG') and nvl(fa10,'000') >'010' and NA16 >='" & frm100106_1.txt5(11).Text & "') " & _
                      " or ((np02='P' or np02='PS') and nvl(fa10,'000') >'010' and NA79 >='" & frm100106_1.txt5(11).Text & "'))"
         End If
         If Len(Trim(frm100106_1.txt5(12).Text)) <> 0 Then
              'Modified by Lydia 2023/09/11 區分FMP案管制人---by Phoebe
              'strSql = strSql & " AND NA16 <='" & frm100106_1.txt5(12).Text & "' "
              strSql = strSql & " and (((np02='FCP' or np02='FG') and nvl(fa10,'000') >'010' and NA16 <='" & frm100106_1.txt5(12).Text & "') " & _
                      " or ((np02='P' or np02='PS') and nvl(fa10,'000') >'010' and NA79 <='" & frm100106_1.txt5(12).Text & "'))"
         End If
      '2006/2/14 END
      Else
         '2010/9/13 MODIFY BY SONIA 日期欄改百年日期排序問題
         'Modify by Amy 2016/07/18 若有下申請人編號,則申請人1~5也要查並顯示申請人資料
         'Modify by Amy 2018/09/20 申請人資料全顯示不限制字數 原:SUBSTRB(CU04,1,10)
         'Modified by Lydia 2019/11/01 +增加欄位SeColPA
         strSql = "SELECT '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日, DECODE(NP02||NP07,'FCP411',NVL(N3.NA16,N2.NA16),NP10) NP10,NP01,np22,cp14, NP02, NP07 " & _
                ",PA26 as ApplyNo " & SeColPA & _
                " FROM NEXTPROGRESS,CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                " WHERE (PA57<>'Y' or pa57 is null) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14 AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL1 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
          If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
             strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日, DECODE(NP02||NP07,'FCP411',NVL(N3.NA16,N2.NA16),NP10) NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",PA27 as ApplyNo " & SeColPA & _
                   " FROM NEXTPROGRESS,CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                   " WHERE (PA57<>'Y' or pa57 is null) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14 AND SUBSTR(PA27,1,8)=CU01(+) AND DECODE(SUBSTR(PA27,9,1),NULL,'0',SUBSTR(PA27,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL10(1) & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
              strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日, DECODE(NP02||NP07,'FCP411',NVL(N3.NA16,N2.NA16),NP10) NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",PA28 as ApplyNo " & SeColPA & _
                   " FROM NEXTPROGRESS,CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                   " WHERE (PA57<>'Y' or pa57 is null) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14 AND SUBSTR(PA28,1,8)=CU01(+) AND DECODE(SUBSTR(PA28,9,1),NULL,'0',SUBSTR(PA28,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL10(2) & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
              strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日, DECODE(NP02||NP07,'FCP411',NVL(N3.NA16,N2.NA16),NP10) NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",PA29 as ApplyNo " & SeColPA & _
                   " FROM NEXTPROGRESS,CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                   " WHERE (PA57<>'Y' or pa57 is null) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14 AND SUBSTR(PA29,1,8)=CU01(+) AND DECODE(SUBSTR(PA29,9,1),NULL,'0',SUBSTR(PA29,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL10(3) & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日, DECODE(NP02||NP07,'FCP411',NVL(N3.NA16,N2.NA16),NP10) NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",PA30 as ApplyNo " & SeColPA & _
                   " FROM NEXTPROGRESS,CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                   " WHERE (PA57<>'Y' or pa57 is null) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14 AND SUBSTR(PA30,1,8)=CU01(+) AND DECODE(SUBSTR(PA30,9,1),NULL,'0',SUBSTR(PA30,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTR(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL10(4) & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
          End If
          'Modify By Sindy 2011/3/15 FCT,T,TF延展(102)和第二期(716)專用權須存在(TM17=Y)
          'Modified by Lydia 2019/11/01 +增加欄位SeColTM
          strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",TM23 as ApplyNo " & SeColTM & _
                   " FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                   " WHERE (TM29<>'Y' or tm29 is null) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(TM23,1,8)=CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) " & strSQL2 & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "") & _
                   " and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' "
          If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",TM78 as ApplyNo " & SeColTM & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (TM29<>'Y' or tm29 is null) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(TM78,1,8)=CU01(+) AND DECODE(SUBSTR(TM78,9,1),NULL,'0',SUBSTR(TM78,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) " & StrSql20(1) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "") & _
                     " and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' "
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",TM79 as ApplyNo " & SeColTM & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (TM29<>'Y' or tm29 is null) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(TM79,1,8)=CU01(+) AND DECODE(SUBSTR(TM79,9,1),NULL,'0',SUBSTR(TM79,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) " & StrSql20(2) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "") & _
                     " and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' "
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",TM80 as ApplyNo " & SeColTM & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (TM29<>'Y' or tm29 is null) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(TM80,1,8)=CU01(+) AND DECODE(SUBSTR(TM80,9,1),NULL,'0',SUBSTR(TM80,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) " & StrSql20(3) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "") & _
                     " and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' "
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",TM81 as ApplyNo " & SeColTM & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (TM29<>'Y' or tm29 is null) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(TM81,1,8)=CU01(+) AND DECODE(SUBSTR(TM81,9,1),NULL,'0',SUBSTR(TM81,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) " & StrSql20(4) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "") & _
                     " and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' "
          End If
          'Modified by Lydia 2019/11/01 +增加欄位SeColLC
          strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",LC11 as ApplyNo " & SeColLC & _
                   " FROM NEXTPROGRESS,CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                   " WHERE (LC08<>'Y' or lc08 is null) AND NP02=LC01(+) AND NP03=LC02(+) AND NP04=LC03(+) AND NP05=LC04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) " & StrSQL3 & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
          If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",LC43 as ApplyNo " & SeColLC & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (LC08<>'Y' or lc08 is null) AND NP02=LC01(+) AND NP03=LC02(+) AND NP04=LC03(+) AND NP05=LC04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(LC43,1,8)=CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) " & StrSql30(1) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                    ",LC44 as ApplyNo " & SeColLC & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (LC08<>'Y' or lc08 is null) AND NP02=LC01(+) AND NP03=LC02(+) AND NP04=LC03(+) AND NP05=LC04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(LC44,1,8)=CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) " & StrSql30(2) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                    "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                    ",LC45 as ApplyNo " & SeColLC & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (LC08<>'Y' or lc08 is null) AND NP02=LC01(+) AND NP03=LC02(+) AND NP04=LC03(+) AND NP05=LC04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(LC45,1,8)=CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) " & StrSql30(3) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                     ",LC46 as ApplyNo " & SeColLC & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (LC08<>'Y' or lc08 is null) AND NP02=LC01(+) AND NP03=LC02(+) AND NP04=LC03(+) AND NP05=LC04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(LC46,1,8)=CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) " & StrSql30(4) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
          End If
          'Modified by Lydia 2019/11/01 +增加欄位SeColHC
          strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                         AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",HC05 as ApplyNo " & SeColHC & _
                   " FROM NEXTPROGRESS,CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,STAFF_IDMAP " & _
                   " WHERE (HC09<>'Y' or hc09 is null) AND NP02=HC01(+) AND NP03=HC02(+) AND NP04=HC03(+) AND NP05=HC04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(HC05,1,8)=CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=CU02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSQL4 & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
          If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                         AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",HC24 as ApplyNo " & SeColHC & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,STAFF_IDMAP " & _
                     " WHERE (HC09<>'Y' or hc09 is null) AND NP02=HC01(+) AND NP03=HC02(+) AND NP04=HC03(+) AND NP05=HC04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(HC24,1,8)=CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=CU02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSql40(1) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                         AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                     ",HC25 as ApplyNo " & SeColHC & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,STAFF_IDMAP " & _
                     " WHERE (HC09<>'Y' or hc09 is null) AND NP02=HC01(+) AND NP03=HC02(+) AND NP04=HC03(+) AND NP05=HC04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(HC25,1,8)=CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=CU02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSql40(2) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                         AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",HC26 as ApplyNo " & SeColHC & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,STAFF_IDMAP " & _
                     " WHERE (HC09<>'Y' or hc09 is null) AND NP02=HC01(+) AND NP03=HC02(+) AND NP04=HC03(+) AND NP05=HC04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(HC26,1,8)=CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=CU02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSql40(3) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                         AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                    "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,NP10,NP01,np22,cp14, NP02, NP07 " & _
                    ",HC27 as ApplyNo " & SeColHC & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,STAFF_IDMAP " & _
                     " WHERE (HC09<>'Y' or hc09 is null) AND NP02=HC01(+) AND NP03=HC02(+) AND NP04=HC03(+) AND NP05=HC04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(HC27,1,8)=CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=CU02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSql40(4) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
          End If
          'Modified by Lydia 2019/11/01 +增加欄位SeColSP
          strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                   "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,DECODE(NP02||NP07,'FG411',NVL(N3.NA16,N2.NA16),NP10) NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",SP08 as ApplyNo " & SeColSP & _
                   " FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                   " WHERE (SP15<>'Y' or sp15 is null) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL5 & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                   " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
          If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                    "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,DECODE(NP02||NP07,'FG411',NVL(N3.NA16,N2.NA16),NP10) NP10,NP01,np22,cp14, NP02, NP07 " & _
                   ",SP58 as ApplyNo " & SeColSP & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (SP15<>'Y' or sp15 is null) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(SP58,1,8)=CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP28,9,1),NULL,'0',SUBSTR(SP28,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) AND FA10=N3.NA01(+) " & StrSql50(1) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                    "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,DECODE(NP02||NP07,'FG411',NVL(N3.NA16,N2.NA16),NP10) NP10,NP01,np22,cp14, NP02, NP07 " & _
                    ",SP59 as ApplyNo " & SeColSP & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (SP15<>'Y' or sp15 is null) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(SP59,1,8)=CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) AND FA10=N3.NA01(+) " & StrSql50(2) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                    "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,DECODE(NP02||NP07,'FG411',NVL(N3.NA16,N2.NA16),NP10) NP10,NP01,np22,cp14, NP02, NP07 " & _
                    ",SP65 as ApplyNo " & SeColSP & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (SP15<>'Y' or sp15 is null) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(SP65,1,8)=CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) AND FA10=N3.NA01(+) " & StrSql50(3) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,NP02||'-'||NP03||'-'||NP04||'-'||NP05 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,NP10) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,NP15 AS 備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                    "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,'' AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,DECODE(NP02||NP07,'FG411',NVL(N3.NA16,N2.NA16),NP10) NP10,NP01,np22,cp14, NP02, NP07 " & _
                    ",SP66 as ApplyNo " & SeColSP & _
                     " FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,STAFF S1,CASEPROPERTYMAP,CUSTOMER,FAGENT,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (SP15<>'Y' or sp15 is null) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP01=CP09(+) AND NP10=S2.ST01(+) AND CP14=S1.ST01(+) AND SIM02(+)=CP14  AND SUBSTR(SP66,1,8)=CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=FA02(+) AND NP02=CPM01(+) AND TO_CHAR(NP07)=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND NP01=SK01(+) AND FA10=N3.NA01(+) " & StrSql50(4) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND NP07 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND NP07 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
          End If
          'end 2018/09/20
          'end 2016/07/18
      End If
       '限制使用者所能使用的系統類別+案件性質
      'If frm100106_1.txt5(0).Text <> "ALL" Then
         'edit by nickc 2008/02/21 未收文的不用抓 cp57
         'strSQL = "Select AA.V AS V, AA.本所期限 AS 本所期限, AA.本所案號 AS 本所案號,AA.分所號 as 分所號, AA.案件名稱 AS 案件名稱, AA.案件性質 AS 案件性質, AA.承辦人 AS 承辦人, AA.智權人員 AS 智權人員, AA.收文日 AS 收文日, AA.法定期限 AS 法定期限, AA.備註 AS 備註, AA.申請人 AS 申請人, AA.是否出名 AS 是否出名, AA.延期日 AS 延期日, AA.申請國家 AS 申請國家, AA.申請人國籍 AS 申請人國籍, AA.發文日 AS 發文日, AA.代理人 AS 代理人, AA.彼所案號 AS 彼所案號, AA.申請案號 AS 申請案號, AA.承辦人備註 AS 承辦人備註, AA.取消收文日 AS 取消收文日, AA.NP10 AS NP10, AA.NP01 AS NP01, AA.np22 AS NP22, AA.cp14 AS CP14, AA.NP02 AS NP02, AA.NP07 AS NP07 " & _
                  " From ( " & strSQL & " ) AA, Staff SA, Staff_Group, Staff SB Where AA.NP02=SG02 And AA.NP07=SG03 And SA.ST11=SG01 And AA.NP10=SB.ST01(+) And SA.ST01='" & strUserNum & "' "
         'Modified by Morgan 2015/10/20 取消系統類別+案件性質的權限控制--秀玲檢查文件沒有該限制
         'strSql = "Select AA.V AS V, AA.本所期限 AS 本所期限, AA.本所案號 AS 本所案號,AA.分所號 as 分所號, AA.案件名稱 AS 案件名稱, AA.案件性質 AS 案件性質, AA.承辦人 AS 承辦人, AA.智權人員 AS 智權人員, AA.收文日 AS 收文日, AA.法定期限 AS 法定期限, AA.備註 AS 備註, AA.申請人 AS 申請人, AA.是否出名 AS 是否出名, AA.延期日 AS 延期日, AA.申請國家 AS 申請國家, AA.申請人國籍 AS 申請人國籍, AA.發文日 AS 發文日, AA.代理人 AS 代理人, AA.彼所案號 AS 彼所案號, AA.申請案號 AS 申請案號, AA.承辦人備註 AS 承辦人備註, '' AS 取消收文日, AA.NP10 AS NP10, AA.NP01 AS NP01, AA.np22 AS NP22, AA.cp14 AS CP14, AA.NP02 AS NP02, AA.NP07 AS NP07 " & _
                  " From ( " & strSql & " ) AA, Staff SA, Staff_Group, Staff SB Where AA.NP02=SG02 And AA.NP07=SG03 And SA.ST11=SG01 And AA.NP10=SB.ST01(+) And SA.ST01='" & strUserNum & "' "
         'Modify by Amy 2016/07/27 +申請人No 不顯示
         'Modified by Lydia 2019/11/01 利益衝突案件：於ApplyNo後面增加欄位
         'Modified by Lydia 2021/05/19 限定欄位長度
         'strSql = "Select AA.V AS V, AA.本所期限 AS 本所期限, AA.本所案號 AS 本所案號,AA.分所號 as 分所號, AA.案件名稱 AS 案件名稱, AA.案件性質 AS 案件性質, AA.承辦人 AS 承辦人, AA.智權人員 AS 智權人員, AA.收文日 AS 收文日, AA.法定期限 AS 法定期限, AA.備註 AS 備註, AA.申請人 AS 申請人, AA.是否出名 AS 是否出名, AA.延期日 AS 延期日, AA.申請國家 AS 申請國家, AA.申請人國籍 AS 申請人國籍, AA.發文日 AS 發文日, AA.代理人 AS 代理人, AA.彼所案號 AS 彼所案號, AA.申請案號 AS 申請案號, AA.承辦人備註 AS 承辦人備註, '' AS 取消收文日, AA.NP10 AS NP10, AA.NP01 AS NP01, AA.np22 AS NP22, AA.cp14 AS CP14, AA.NP02 AS NP02, AA.NP07 AS NP07 " & _
                  ",ApplyNo, cust01, cust02, cust03, cust04, cust05, fcno " & _
                  " From ( " & strSql & " ) AA "
         'end 2015/10/20
         'Modified by Morgan 2021/5/25 補 , Staff SB (依智權人員排序會用)
         strSql = "Select AA.V AS V, substr(AA.本所期限,1,10) AS 本所期限, AA.本所案號 AS 本所案號,AA.分所號 as 分所號, AA.案件名稱 AS 案件名稱, AA.案件性質 AS 案件性質,AA.承辦人 AS 承辦人, AA.智權人員 AS 智權人員" & _
                  ", substr(AA.收文日,1,10) AS 收文日, substr(AA.法定期限,1,10) AS 法定期限, substr(AA.備註,1,500) AS 備註, AA.申請人 AS 申請人, AA.是否出名 AS 是否出名, substr(AA.延期日,1,10) AS 延期日" & _
                  ", AA.申請國家 AS 申請國家, AA.申請人國籍 AS 申請人國籍, substr(AA.發文日,1,10) AS 發文日, AA.代理人 AS 代理人, AA.彼所案號 AS 彼所案號, AA.申請案號 AS 申請案號" & _
                  ",substr(AA.承辦人備註,1,500) AS 承辦人備註, '' AS 取消收文日, AA.NP10 AS NP10, AA.NP01 AS NP01, AA.np22 AS NP22, AA.cp14 AS CP14, AA.NP02 AS NP02, AA.NP07 AS NP07 " & _
                  ",ApplyNo, cust01, cust02, cust03, cust04, cust05, fcno " & _
                  " From ( " & strSql & " ) AA, Staff SB Where AA.NP10=SB.ST01(+) "
      'End If
       'Modify By Cheng 2003/05/23
       '是否依智權人員排序
       If frm100106_1.txt5(13).Text = "Y" Then
           'edit by nickc 2008/02/21 取消收文日放最下面
           'strSQL = strSQL & " ORDER BY SB.ST03, SB.ST01, 本所期限,本所案號 "
           '2010/9/17 modify by sonia 因改日期欄百年問題而修正
           'strSql = strSql & " ORDER BY SB.ST03, SB.ST01,nvl(取消收文日,'11/11/11'), 本所期限,本所案號 "
           strSql = strSql & " ORDER BY SB.ST03, SB.ST01,nvl(取消收文日,' 11/11/11'), 本所期限,本所案號 "
           pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(9) & frm100106_1.txt5(13)  'Add By Sindy 2010/01/22
       Else
           'edit by nickc 2008/02/21 取消收文日放最下面
           'strSQL = strSQL & " ORDER BY 本所期限,本所案號 "
           '2010/9/17 modify by sonia 因改日期欄百年問題而修正
           'strSql = strSql & " ORDER BY nvl(取消收文日,'11/11/11'),本所期限,本所案號 "
           strSql = strSql & " ORDER BY nvl(取消收文日,' 11/11/11'),本所期限,本所案號 "
       End If
       
       SetDataListWidth 'Added by Lydia 2021/05/19 清空欄位
       CheckOC
       adoRecordset.CursorLocation = adUseClient
       'Modified by Lydia 2019/11/01 改變型態
       'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
       If adoRecordset.RecordCount <> 0 Then
          dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
             'Added by Lydia 2019/11/01 逐案號判斷
             If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
                adoRecordset.MoveFirst
                Do While adoRecordset.EOF = False
                    '利益衝突案件：逐案號判斷
                    If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                        intCufaCnt = intCufaCnt + 1
                        adoRecordset.Delete
                    End If
                    adoRecordset.MoveNext
                Loop
                '利益衝突案件：限閱案件
                If intCufaCnt > 0 Then
                    pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                    MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
                End If
                InsertQueryLog (dblRow) 'Add By Sindy 2010/01/22
                If adoRecordset.RecordCount = 0 Then
                      GoTo JumpToNoData
                End If
             Else
               InsertQueryLog (dblRow) 'Add By Sindy 2010/01/22
             End If
            'end 2019/11/01
           
           cmdOK(0).Enabled = True
           cmdOK(1).Enabled = True
           cmdOK(2).Enabled = True
           cmdOK(3).Enabled = True
       Else
           InsertQueryLog (0) 'Add By Sindy 2010/01/22
JumpToNoData:   'Added by Lydia 2019/11/01
           'Add by Amy 2016/07/18
           If bolMsg = True Then ShowNoData
           Screen.MousePointer = vbDefault
           cmdOK(0).Enabled = False
           cmdOK(1).Enabled = False
           cmdOK(2).Enabled = False
           cmdOK(3).Enabled = False
           StrMenu = False
           Exit Function
       End If
       Set grdDataList.Recordset = adoRecordset
       Call SetDataListWidth(False) 'Added by Lydia 2021/05/19 預設欄位
       
       intK = grdDataList.Rows - 1
       CheckOC
       grdDataList.Visible = False
       For i = 1 To grdDataList.Rows - 1
            '910710 Sieg 107
            'Modified by Lydia 2021/05/19 改用變數取得
'            strTmp = Replace(grdDataList.TextMatrix(i, 2), "-", "")
'            If Left(strTmp, 3) = "FCP" Then
'               'Add by Morgan 2005/2/24 催審抓FCP管制人
'               If grdDataList.TextMatrix(i, 27) = "411" Then
'                  grdDataList.TextMatrix(i, 7) = GetStaffName(grdDataList.TextMatrix(i, 22))
'               Else
'                  If GetFCPSales(strTmp, strTmp1) Then
'                     grdDataList.TextMatrix(i, 7) = strTmp1
'                  End If
'               End If
'            ElseIf Left(strTmp, 2) = "FG" Then
'               'Add by Morgan 2005/2/24 催審抓FCP管制人
'               If grdDataList.TextMatrix(i, 27) = "411" Then
'                  grdDataList.TextMatrix(i, 7) = GetStaffName(grdDataList.TextMatrix(i, 22))
'               Else
'                  If GetFGSales(strTmp, strTmp1) Then
'                     grdDataList.TextMatrix(i, 7) = strTmp1
'                  End If
'               End If
'            End If
'            ' 相關案件性質  2012/8/9 ADD BY SONIA
'            If IsNull(Left(strTmp, 3)) = False Then
'               grdDataList.TextMatrix(i, 5) = grdDataList.TextMatrix(i, 5) & PUB_GetNextCasePropertyName(grdDataList.TextMatrix(i, 23), grdDataList.TextMatrix(i, 24), "1")
'            End If
'            '2012/8/9 END
'           grdDataList.row = i
'           grdDataList.col = grdDataList.Cols - 5
'           strSql = "SELECT " & SQLDate("DL02") & " FROM DATELIMIT WHERE DL01='" & grdDataList.Text & "' And DL06=" & Me.grdDataList.TextMatrix(i, 24) & " ORDER BY DL02"
            strTmp = Replace(grdDataList.TextMatrix(i, colCaseNo), "-", "")
            If Left(strTmp, 3) = "FCP" Then
               'Add by Morgan 2005/2/24 催審抓FCP管制人
               If grdDataList.TextMatrix(i, colCP10) = "411" Then
                  grdDataList.TextMatrix(i, colSalesName) = GetStaffName(grdDataList.TextMatrix(i, colCP13))
               Else
                  If GetFCPSales(strTmp, strTmp1) Then
                     grdDataList.TextMatrix(i, colSalesName) = strTmp1
                  End If
               End If
            ElseIf Left(strTmp, 2) = "FG" Then
               'Add by Morgan 2005/2/24 催審抓FCP管制人
               If grdDataList.TextMatrix(i, colCP10) = "411" Then
                  grdDataList.TextMatrix(i, colSalesName) = GetStaffName(grdDataList.TextMatrix(i, colCP13))
               Else
                  If GetFGSales(strTmp, strTmp1) Then
                     grdDataList.TextMatrix(i, colSalesName) = strTmp1
                  End If
               End If
            End If
            ' 相關案件性質  2012/8/9 ADD BY SONIA
            If IsNull(Left(strTmp, 3)) = False Then
               grdDataList.TextMatrix(i, colCp10Name) = grdDataList.TextMatrix(i, colCp10Name) & PUB_GetNextCasePropertyName(grdDataList.TextMatrix(i, colCp09), grdDataList.TextMatrix(i, colCP66), "1")
            End If
            '2012/8/9 END
           grdDataList.row = i
           grdDataList.col = colCp09
           strSql = "SELECT " & SQLDate("DL02") & " FROM DATELIMIT WHERE DL01='" & grdDataList.Text & "' And DL06=" & CNULL(Me.grdDataList.TextMatrix(i, colCP66)) & " ORDER BY DL02"
           'end 2021/05/19
           
           CheckOC
           adoRecordset.CursorLocation = adUseClient
           adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
           'Modified by Lydia 2021/05/19 改用變數取得: 13=>colDDate3
           grdDataList.col = colDDate3
           If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               adoRecordset.MoveLast
               If Not IsNull(adoRecordset.Fields(0)) Then
                    grdDataList.Text = adoRecordset.Fields(0)
               Else
                    grdDataList.Text = ""
               End If
           End If
           'Remove by Lydia 2021/05/19
           'grdDataList.col = 16
           'grdDataList.col = 1
           'end 2021/05/19
           'Add by Morgan 2004/8/16 6個月逾繳期限 "&" 紫色
           'Modified by Lydia 2021/05/19 改用變數取得:  2=> colCaseNo
           grdDataList.col = colCaseNo
           If grdDataList.CellBackColor = -2147483643 Then
               'Modified by Lydia 2021/05/19 改用變數取得:  2=> colCaseNo
               stCaseNo = grdDataList.TextMatrix(grdDataList.row, colCaseNo)
               For j = 1 To 4
                  iPos = InStr(stCaseNo, "-")
                  If iPos > 0 Then
                     stPA(j) = Left(stCaseNo, iPos - 1)
                     stCaseNo = Mid(stCaseNo, iPos + 1)
                  Else
                     stPA(j) = stCaseNo
                  End If
               Next j
               'Modified by Lydia 2021/05/19 改用變數取得
               'If (UCase(stPA(1)) = "FCP" Or UCase(stPA(1)) = "P") And grdDataList.TextMatrix(grdDataList.row, 27) = "605" Then
               '   '2010/9/20 modify by sonia 因日期加空格故加val
               '   strTmp = Val(ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, 9))))
               If (UCase(stPA(1)) = "FCP" Or UCase(stPA(1)) = "P") And grdDataList.TextMatrix(grdDataList.row, colCP10) = "605" Then
                  strTmp = Val(ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, colDDate2))))
               'end 2021/05/19
                  If strTmp <> "" Then
                    If PUB_IfCtrlDateExtended(stPA, strTmp) = True Then
                       'Modified by Lydia 2021/05/19 改用變數取得:  2=> colCaseNo
                       grdDataList.TextMatrix(grdDataList.row, colCaseNo) = "&" & grdDataList.TextMatrix(grdDataList.row, colCaseNo)
                       For j = 1 To grdDataList.Cols - 1
                          grdDataList.col = j
                          '紫色
                          grdDataList.CellBackColor = &HE600E6
                       Next j
                    End If
                  End If
               End If
           End If
           'Modified by Lydia 2021/05/19 改用變數取得:  2=> colCaseNo
           'grdDataList.col = 2
           grdDataList.col = colCaseNo
           If grdDataList.CellBackColor = -2147483643 Then
                'Modified by Lydia 2021/05/19 改用變數取得
                'If Mid(grdDataList.TextMatrix(grdDataList.row, 23), 1, 1) = "C" And grdDataList.TextMatrix(grdDataList.row, 16) = "" Then
                '    grdDataList.col = 2
                If Mid(grdDataList.TextMatrix(grdDataList.row, colCp09), 1, 1) = "C" And grdDataList.TextMatrix(grdDataList.row, colCP27) = "" Then
                    grdDataList.col = colCaseNo
                'end 2021/05/19
                    grdDataList.Text = "#" + grdDataList.Text
                    For j = 1 To grdDataList.Cols - 1
                        grdDataList.col = j
                        '黃色
                        grdDataList.CellBackColor = &HFFFF&
                    Next j
               'edit by nickc 2008/02/21 修正字串比對錯誤
               'ElseIf ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, 1)) < ChangeWStringToTString(ServerDate) And Trim(grdDataList.TextMatrix(grdDataList.row, 1)) <> "" Then
               'Modified by Lydia 2021/05/19 改用變數取得
               'ElseIf Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, 1))) < Val(ChangeWStringToTString(ServerDate)) And Trim(grdDataList.TextMatrix(grdDataList.row, 1)) <> "" Then
               '    grdDataList.col = 2
               ElseIf Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, colDDate1))) < Val(ChangeWStringToTString(strSrvDate(1))) And Trim(grdDataList.TextMatrix(grdDataList.row, colDDate1)) <> "" Then
                   grdDataList.col = colCaseNo
               'end 2021/05/19
                   grdDataList.Text = "*" + grdDataList.Text
                   For j = 1 To grdDataList.Cols - 1
                       grdDataList.col = j
                       '紅色
                       grdDataList.CellBackColor = &HFF&
                   Next j
                   GoTo Nextitem0
               Else
                  '2010/9/20 modify by sonia 因日期加空格故加val
                  'Modified by Lydia 2021/05/19 改用變數取得
                  'If Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, 1))) = Val(ChangeWStringToTString(ServerDate)) Then
                  '     grdDataList.col = 2
                  If Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, colDDate1))) = Val(ChangeWStringToTString(strSrvDate(1))) Then
                       grdDataList.col = colCaseNo
                  'end 2021/05/19
                       grdDataList.Text = "v" & grdDataList.Text
                       For j = 1 To grdDataList.Cols - 1
                           grdDataList.col = j
                           '橙色
                           grdDataList.CellBackColor = &H80FF&
                       Next j
                  'add nick 2004/07/13 加入年費無公告判斷
                  Else
                        'Modified by Lydia 2021/05/19 改用變數取得:  2=> colCaseNo
                        grdDataList.col = colCaseNo
                        If UCase(SystemNumber(grdDataList.Text, 1)) = "FCP" Or UCase(SystemNumber(grdDataList.Text, 1)) = "P" Then
                           'Modified by Lydia 2021/05/19 改用變數取得:  27=> colCP10
                           grdDataList.col = colCP10
                           If grdDataList.Text = "605" Then
                                 CheckOC
                                 'Modified by Lydia 2021/05/19 改用變數取得:  2=> colCaseNo
                                 grdDataList.col = colCaseNo
                                 strSql = "select pa09,pa14 from patent where pa01='" & SystemNumber(grdDataList.Text, 1) & _
                                                 "' and pa02='" & SystemNumber(grdDataList.Text, 2) & "' and pa03='" & _
                                                 SystemNumber(grdDataList.Text, 3) & "' and pa04='" & SystemNumber(grdDataList.Text, 4) & "' and pa09='000' "
                                 adoRecordset.CursorLocation = adUseClient
                                 adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                                 If adoRecordset.RecordCount <> 0 Then
                                    If CheckStr(adoRecordset.Fields(0).Value) = "000" And CheckStr(adoRecordset.Fields(1).Value) = "" Then
                                       'Modified by Lydia 2021/05/19 改用變數取得:  2=> colCaseNo
                                       grdDataList.col = colCaseNo
                                       grdDataList.Text = "!" & grdDataList.Text
                                       For j = 1 To grdDataList.Cols - 1
                                           grdDataList.col = j
                                           '綠色
                                           grdDataList.CellBackColor = &HC000&
                                       Next j
                                    End If
                                 End If
                                 CheckOC
                           End If
                        End If
                  End If
               End If
           End If
Nextitem0:
      Next i
      grdDataList.Visible = True
   Else
      'Modify By Sindy 2010/01/27 程式太長，切程式段
      If frm100106_1.opt2(1).Value = True Then
         '已收文未發文
         StrMenu = StrMenu_sub1(StrSQLa)
      Else
         '已收文已發文
         StrMenu = StrMenu_sub2(StrSQLa)
      End If
      If StrMenu = False Then Exit Function
   End If
End Function

'已收文未發文
Public Function StrMenu_sub1(StrSQLa As String) As Boolean
Dim strTmp As String
Dim stPA(1 To 4) As String, iPos As Integer, stCaseNo As String
Dim strDate1Col As String, strDate2Col As String 'Add By Sindy 2013/8/13
'Added by Lydia 2019/11/18 比照StrMeu => Add by Amy 2016/07/18 for 申請人1~5用
Dim strSQL10(1 To 4) As String, StrSql20(1 To 4) As String, StrSql30(1 To 4) As String, StrSql40(1 To 4) As String, StrSql50(1 To 4) As String
'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String
Dim dblRow As Double 'Add By Sindy 2025/9/3

      StrMenu_sub1 = True
      
      pub_QL05 = pub_QL05 & ";" & frm100106_1.Frame1.Caption & frm100106_1.opt2(1).Caption 'Add By Sindy 2010/01/22
      strSQL1 = ""
      strSQL2 = ""
      StrSQL3 = ""
      StrSQL4 = ""
      strSQL5 = ""
      'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
      SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
      SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
      SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
      SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
      SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
      m_AllSys = IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(, frm100106_1.txt5(0).Text))
      'Added by Lydia 2020/03/11 外專人員勾選含FMP外專管制期限時,要另外增加P,PS
      If frm100106_1.Check1.Value = 1 And Left(Pub_StrUserSt03, 2) = "F2" Then
          If InStr(m_AllSys & ",", ",P,") = 0 Then m_AllSys = m_AllSys & ",P,"
          If InStr(m_AllSys & ",", ",PS,") = 0 Then m_AllSys = m_AllSys & ",PS,"
          m_AllSys = Replace(m_AllSys, ",,", ",")
      End If
      'end 2020/03/11
      intCufaCnt = 0
      'end 2019/11/01
      
      'Add By Sindy 2013/8/13
      'Modified by Morgan 2024/1/8 +指定日期
      'strDate1Col = "SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限"
      strDate1Col = "SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限,SUBSTR(' '||sqldatet(CP142),-9) AS 指定日期"
      'end 2024/1/8
      strDate2Col = "SUBSTR(' '||sqldatet(CP07),-9) AS 法定期限"
      '以承辦期限查詢
      If frm100106_1.opt1(3).Value Then
         strDate1Col = "SUBSTR(' '||sqldatet(CP48),-9) AS 承辦期限"
         strDate2Col = "SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限"
          If Len(Trim(frm100106_1.txt6(0))) <> 0 Then
              strSQL1 = strSQL1 + " and CP48>=" & Val(ChangeTStringToWString(frm100106_1.txt6(0))) & " "
          End If
          pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(3).Caption & frm100106_1.txt6(0) 'Add By Sindy 2010/01/22
          If Len(Trim(frm100106_1.txt6(1))) <> 0 Then
              strSQL1 = strSQL1 + " and CP48<=" & Val(ChangeTStringToWString(frm100106_1.txt6(1))) & " "
              pub_QL05 = pub_QL05 & "-" & frm100106_1.txt6(1) 'Add By Sindy 2010/01/22
          Else
              strSQL1 = strSQL1 + " and CP48<=" & Val(ChangeTStringToWString(strSrvDate(1) - 19110000)) & " "
              pub_QL05 = pub_QL05 & "-" & (strSrvDate(1) - 19110000) 'Add By Sindy 2010/01/22
          End If
          '有完稿日者,代表已承辦,不需顯示出來
          strSQL1 = strSQL1 & " and ep09 is null"
          strSQL2 = strSQL2 & " and ep09 is null"
          StrSQL3 = StrSQL3 & " and ep09 is null"
          StrSQL4 = StrSQL4 & " and ep09 is null"
          strSQL5 = strSQL5 & " and ep09 is null"
      End If
      '2013/8/13 END
      '以本所期限查詢
      If frm100106_1.opt1(0).Value Then
          If Len(Trim(frm100106_1.txt1(0))) <> 0 Then
            'Added by Morgan 2024/1/8
            'Removed by Morgan 2024/3/11 速度太慢改在最後另外union
            'If strSrvDate(1) >= 指定日期啟用日 Then
            '   strSQL1 = strSQL1 + " and ( CP06>=" & Val(ChangeTStringToWString(frm100106_1.Txt1(0))) & " or cp142>=" & Val(ChangeTStringToWString(frm100106_1.Txt1(0))) & ") "
            'Else
            'end 2024/1/8
               strSQL1 = strSQL1 + " and CP06>=" & Val(ChangeTStringToWString(frm100106_1.txt1(0))) & " "
            'End If
          End If
          pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(0).Caption & frm100106_1.txt1(0) 'Add By Sindy 2010/01/22
          If Len(Trim(frm100106_1.txt1(1))) <> 0 Then
               'Added by Morgan 2024/1/8
               'Removed by Morgan 2024/3/11 速度太慢改在最後另外union
               'If strSrvDate(1) >= 指定日期啟用日 Then
               '   strSQL1 = strSQL1 + " and (CP06<=" & Val(ChangeTStringToWString(frm100106_1.Txt1(1))) & " or cp142<=" & Val(ChangeTStringToWString(frm100106_1.Txt1(1))) & ") "
               'Else
               'end 2024/1/8
                  strSQL1 = strSQL1 + " and CP06<=" & Val(ChangeTStringToWString(frm100106_1.txt1(1))) & " "
               'End If
              pub_QL05 = pub_QL05 & "-" & frm100106_1.txt1(1) 'Add By Sindy 2010/01/22
          Else
              strSQL1 = strSQL1 + " and CP06<=" & Val(ChangeTStringToWString(strSrvDate(1) - 19110000)) & " "
              pub_QL05 = pub_QL05 & "-" & (strSrvDate(1) - 19110000) 'Add By Sindy 2010/01/22
          End If
      End If
     '以本所案號查詢
     If frm100106_1.opt1(2).Value Then
         If Len(Trim(frm100106_1.txt3(0))) <> 0 Then
             strSQL1 = strSQL1 + " AND CP01='" & frm100106_1.txt3(0) & "' "
         End If
         pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(2).Caption & frm100106_1.txt3(0) 'Add By Sindy 2010/01/22
         If Len(Trim(frm100106_1.txt3(1))) <> 0 Then
             strSQL1 = strSQL1 + " AND CP02='" & frm100106_1.txt3(1) & "' "
         End If
         pub_QL05 = pub_QL05 & "-" & frm100106_1.txt3(1) 'Add By Sindy 2010/01/22
         If Len(Trim(frm100106_1.txt3(2))) <> 0 Then
             strSQL1 = strSQL1 + " AND CP03='" & frm100106_1.txt3(2) & "' "
             pub_QL05 = pub_QL05 & "-" & frm100106_1.txt3(2) 'Add By Sindy 2010/01/22
         Else
             strSQL1 = strSQL1 + " AND CP03='0' "
             pub_QL05 = pub_QL05 & "-" & "0" 'Add By Sindy 2010/01/22
         End If
         If Len(Trim(frm100106_1.txt3(3))) <> 0 Then
             strSQL1 = strSQL1 + " AND CP04='" & frm100106_1.txt3(3) & "' "
             pub_QL05 = pub_QL05 & "-" & frm100106_1.txt3(3) 'Add By Sindy 2010/01/22
         Else
             strSQL1 = strSQL1 + " AND CP04='00' "
             pub_QL05 = pub_QL05 & "-" & "00" 'Add By Sindy 2010/01/22
         End If
     End If
      If Len(Trim(frm100106_1.txt5(1))) <> 0 Then
         'Modify by Morgan 2007/9/27 也要抓外譯編號
         'strSQL1 = strSQL1 & " AND CP14='" & frm100106_1.txt5(1) & "' "
         strExc(1) = PUB_GetMapID(frm100106_1.txt5(1), 0)
         If strExc(1) <> "" Then
            strSQL1 = strSQL1 & " AND CP14 in ('" & frm100106_1.txt5(1) & "','" & strExc(1) & "')"
         Else
            strSQL1 = strSQL1 & " AND CP14='" & frm100106_1.txt5(1) & "' "
         End If
         'end 2007/9/27
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(2) & frm100106_1.txt5(1) & frm100106_1.LBL1(0) 'Add By Sindy 2010/01/22
      End If
      'Add by Morgan 2007/9/21 加承辦人組別
      If Len(Trim(frm100106_1.txt5(14))) <> 0 Then
           strSQL1 = strSQL1 & " AND exists(select * from staff SX where SX.ST01=NVL(SIM01,CP14) and SX.ST16='" & frm100106_1.txt5(14) & "')"
           pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(10) & frm100106_1.txt5(14)  'Add By Sindy 2010/01/22
      End If
      'end 2007/9/21
      'Modify by Amy 2016/09/12 避免抓到錯的index,所以+||''
      If Len(Trim(frm100106_1.txt5(2))) <> 0 Then
          strSQL1 = strSQL1 & " AND cp12||''>='" & frm100106_1.txt5(2) & "' "
      End If
      If Len(Trim(frm100106_1.txt5(3))) <> 0 Then
          strSQL1 = strSQL1 & " AND cp12||''<='" & frm100106_1.txt5(3) & "' "
      End If
      
       'Added by Morgan 2012/5/23 FMP案改可選擇
       If frm100106_1.Check1.Value = 1 Then
            pub_QL05 = pub_QL05 & ";" & frm100106_1.Check1.Caption
       Else
            'Modify By Sindy 2023/10/30 EP33要回歸用在英文核完日,改抓EP39.核稿完成日
            strSQL1 = strSQL1 & " and (cp01 not in ('P','PS','CFP','CPS') or substr(cp12,1,1)<>'F' or substr(s1.st03,1,1)<>'F' or (cp10<>'201' and ep09>0) or (cp10='201' and " & IIf(strSrvDate(1) >= FCP核完日改用EP39, "ep39", "ep33") & ">0)) "
       End If
       'end 2012/5/23
      
      If Len(Trim(frm100106_1.txt5(2))) <> 0 Or Len(Trim(frm100106_1.txt5(3))) <> 0 Then
          pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(3) & frm100106_1.txt5(2) & "-" & frm100106_1.txt5(3) 'Add By Sindy 2010/01/22
      End If
     If Len(Trim(frm100106_1.txt5(4))) <> 0 Then
        '2008/3/31 MODIFY BY SONIA 加控制查詢87027陳淑芳同時查20001台中所
        'strSQL1 = strSQL1 + " AND CP13='" & frm100106_1.txt5(4) & "' "
        If frm100106_1.txt5(4) = "87027" Then
           strSQL1 = strSQL1 + " AND CP13 IN ('87027','20001') "
        Else
           strSQL1 = strSQL1 + " AND CP13='" & frm100106_1.txt5(4) & "' "
        End If
        '2008/3/31 END
        pub_QL05 = pub_QL05 & ";" & frm100106_1.Label2(0) & frm100106_1.txt5(4) & frm100106_1.LBL1(1) 'Add By Sindy 2010/01/22
     End If
     '申請人國籍
     If Len(Trim(frm100106_1.txt5(9))) <> 0 Then
         strSQL1 = strSQL1 + " AND CU10>='" & frm100106_1.txt5(9) & "' "
     End If
     If Len(Trim(frm100106_1.txt5(10))) <> 0 Then
         strSQL1 = strSQL1 + " AND CU10<='" & frm100106_1.txt5(10) & "' "
     End If
     If Len(Trim(frm100106_1.txt5(9))) <> 0 Or Len(Trim(frm100106_1.txt5(10))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(1) & frm100106_1.txt5(9) & "-" & frm100106_1.txt5(10) 'Add By Sindy 2010/01/22
     End If
     '已收文未發文的資料應限制CP27 及CP57 IS NULL
     'edit by nickc 2008/02/15 秀玲說已收文的，不要控制取消收文
    'strSQL1 = strSQL1 + " and CP27 IS NULL and CP57 IS NULL  "
    strSQL1 = strSQL1 + " and CP27 IS NULL "
    strSQL2 = strSQL1
    StrSQL3 = strSQL1
    StrSQL4 = strSQL1
    strSQL5 = strSQL1
    
    strSQL1 = strSQL1 & " and cp10<>'926'" 'Added by Morgan 2015/10/20 排除核對已准專利
    strSQL1 = strSQL1 + " and cp10<>'1920'"  'Added by Lydia 2018/03/13 排除客戶提供文件
    
    If Len(Trim(frm100106_1.txt5(0))) <> 0 Then
         'Added by Morgan 2015/10/20
         '外專人員勾選含FMP外專管制期限時
         If frm100106_1.Check1.Value = 1 And Left(Pub_StrUserSt03, 2) = "F2" Then
            strSQL1 = strSQL1 & " AND (cp01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 1) & ") or (cp01 IN ('P','CFP') AND substr(cp12,1,1)='F')) "
            strSQL5 = strSQL5 & " AND (cp01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 5) & ") or (cp01 IN ('PS','CPS') AND substr(cp12,1,1)='F')) "
         Else
         'end 2015/10/20
            strSQL1 = strSQL1 & " AND cp01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 1) & ") "
            strSQL5 = strSQL5 & " AND cp01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 5) & ") "
         End If 'Added by Morgan 2015/10/20
         
         strSQL2 = strSQL2 & " AND cp01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 2) & ") "
         StrSQL3 = StrSQL3 & " AND cp01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 3) & ") "
         StrSQL4 = StrSQL4 & " AND cp01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 4) & ") "
         
         pub_QL05 = pub_QL05 & ";" & Left(frm100106_1.Label1(4), 5) & frm100106_1.txt5(0) 'Add By Sindy 2010/01/22
    End If

    If Len(Trim(frm100106_1.txt4(2))) <> 0 Then
         strSQL1 = strSQL1 & " AND PA75>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
         strSQL2 = strSQL2 & " AND TM44>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
         StrSQL3 = StrSQL3 & " AND LC22>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
         strSQL5 = strSQL5 & " AND SP26>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
    End If
    If Len(Trim(frm100106_1.txt4(3))) <> 0 Then
         strSQL1 = strSQL1 & " AND PA75<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
         strSQL2 = strSQL2 & " AND TM44<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
         StrSQL3 = StrSQL3 & " AND LC22<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
         strSQL5 = strSQL5 & " AND SP26<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
    End If
    If Len(Trim(frm100106_1.txt4(2))) <> 0 Or Len(Trim(frm100106_1.txt4(3))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(5) & frm100106_1.txt4(2) & "-" & frm100106_1.txt4(3)  'Add By Sindy 2010/01/22
    End If
    
'---------Move by  Lydia 2019/11/18 從If Len(Trim(frm100106_1.txt4(2))) <> 0 Then的上方移過來
    If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            strSQL10(1) = strSQL1 & " AND PA27>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            strSQL10(2) = strSQL1 & " AND PA28>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            strSQL10(3) = strSQL1 & " AND PA29>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            strSQL10(4) = strSQL1 & " AND PA30>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
         'end 2019/11/18
         strSQL1 = strSQL1 & " AND PA26>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "

         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql20(1) = strSQL2 & " AND TM78>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql20(2) = strSQL2 & " AND TM79>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql20(3) = strSQL2 & " AND TM80>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql20(4) = strSQL2 & " AND TM81>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
         'end 2019/11/18
         strSQL2 = strSQL2 & " AND TM23>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "

         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql30(1) = StrSQL3 & " AND LC43>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql30(2) = StrSQL3 & " AND LC44>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql30(3) = StrSQL3 & " AND LC45>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql30(4) = StrSQL3 & " AND LC46>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
         'end 2019/11/18
         StrSQL3 = StrSQL3 & " AND LC11>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "

         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql40(1) = StrSQL4 & " AND HC24>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql40(2) = StrSQL4 & " AND HC25>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql40(3) = StrSQL4 & " AND HC26>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql40(4) = StrSQL4 & " AND HC27>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
         'end 2019/11/18
         StrSQL4 = StrSQL4 & " AND HC05>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "

         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql50(1) = strSQL5 & " AND SP58>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql50(2) = strSQL5 & " AND SP59>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql50(3) = strSQL5 & " AND SP65>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql50(4) = strSQL5 & " AND SP66>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
         'end 2019/11/18
         strSQL5 = strSQL5 & " AND SP08>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
    End If
    If Len(Trim(frm100106_1.txt4(1))) <> 0 Then
         strSQL1 = strSQL1 & " AND PA26<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            strSQL10(1) = strSQL10(1) & " AND PA27<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            strSQL10(2) = strSQL10(2) & " AND PA28<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            strSQL10(3) = strSQL10(3) & " AND PA29<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            strSQL10(4) = strSQL10(4) & " AND PA30<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'end 2019/11/18
         strSQL2 = strSQL2 & " AND TM23<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql20(1) = StrSql20(1) & " AND TM78<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql20(2) = StrSql20(2) & " AND TM79<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql20(3) = StrSql20(3) & " AND TM80<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql20(4) = StrSql20(4) & " AND TM81<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'end 2019/11/18
         StrSQL3 = StrSQL3 & " AND LC11<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql30(1) = StrSql30(1) & " AND LC43<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql30(2) = StrSql30(1) & " AND LC44<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql30(3) = StrSql30(1) & " AND LC45<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql30(4) = StrSql30(1) & " AND LC46<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'end 2019/11/18
         StrSQL4 = StrSQL4 & " AND HC05<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql40(1) = StrSql40(1) & " AND HC24<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql40(2) = StrSql40(2) & " AND HC25<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql40(3) = StrSql40(3) & " AND HC26<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql40(4) = StrSql40(4) & " AND HC27<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'end 2019/11/18
         strSQL5 = strSQL5 & " AND SP08<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql50(1) = StrSql50(1) & " AND SP58<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql50(2) = StrSql50(2) & " AND SP59<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql50(3) = StrSql50(3) & " AND SP65<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql50(4) = StrSql50(4) & " AND SP66<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'end 2019/11/18
    End If
    If Len(Trim(frm100106_1.txt4(0))) <> 0 Or Len(Trim(frm100106_1.txt4(1))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(6) & frm100106_1.txt4(0) & "-" & frm100106_1.txt4(1) 'Add By Sindy 2010/01/22
    End If
'---------end 2019/11/18

   'Add by Morgan 2006/2/14 有下FCP管制人時只抓FCP資料
   If Len(Trim(frm100106_1.txt5(11).Text)) + Len(Trim(frm100106_1.txt5(12).Text)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(8) & frm100106_1.txt5(11) & "-" & frm100106_1.txt5(12)   'Add By Sindy 2010/01/22
      '2010/9/13 MODIFY BY SONIA 日期欄改百年日期排序問題
      'Modify by Amy 2018/09/20 申請人資料全顯示不限制字數 原:SUBSTRB(CU04,1,10)
      'Modified by Lydia 2019/11/01 +增加欄位SeColPA,SeColSP; PA26,SP08 => ApplyNo
      'Modified by Lydia 2021/05/19 在CP09後面加上CP66
      strSql = "SELECT '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10,PA76,pa75 " & _
                ",PA26 as ApplyNo " & SeColPA & " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                " WHERE (PA57<>'Y' or pa57 is null) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL1 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
       strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10,NULL,sp26 " & _
                ",SP08 as ApplyNo " & SeColSP & " FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                " WHERE (SP15<>'Y' or sp15 is null) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL5 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
      'end 2018/09/20
      '年費管制人順序:pa76->pa75->pa26.cu96->pa26(PA76可能是客戶編號)
      'Modified by Lydia 2019/11/18 原PA26改為ApplyNo
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)
      'strSql = "SELECT Y.* FROM (" & _
         " select X.*,NVL(NVL(F1.FA10,C1.CU10),NVL(F2.FA10,NVL(F4.FA10,C2.CU10))) FA10" & _
         " from (" & strSql & ") X, FAGENT F1, FAGENT F2,FAGENT F4,CUSTOMER C1, CUSTOMER C2" & _
         " WHERE CP10='605' AND F1.FA01(+)=SUBSTR(PA76,1,8) AND F1.FA02(+)=SUBSTR(PA76||'0',9,1)" & _
         " AND C1.CU01(+)=SUBSTR(PA76,1,8) AND C1.CU02(+)=SUBSTR(PA76||'0',9,1)" & _
         " AND F2.FA01(+)=SUBSTR(PA75,1,8) AND F2.FA02(+)=SUBSTR(PA75||'0',9,1)" & _
         "" & _
         " AND C2.CU01(+)=SUBSTR(ApplyNo,1,8) AND C2.CU02(+)=SUBSTR(ApplyNo||'0',9,1)" & _
         " AND F4.FA01(+)=SUBSTR(C2.CU96,1,8) AND F4.FA02(+)=SUBSTR(C2.CU96||'0',9,1)" & _
         " Union All" & _
         " select X.*,NVL(FA10,CU10) FA10" & _
         " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
         " WHERE CP10<>'605' AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
         " AND CU01(+)=SUBSTR(ApplyNo,1,8) AND CU02(+)=SUBSTR(ApplyNo||'0',9,1)" & _
         ") Y, NATION WHERE NA01(+)=FA10"
      strSql = "SELECT Y.* FROM (" & _
         " select X.*,NVL(FA10,CU10) FA10" & _
         " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
         " WHERE FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
         " AND CU01(+)=SUBSTR(ApplyNo,1,8) AND CU02(+)=SUBSTR(ApplyNo||'0',9,1)" & _
         ") Y, NATION WHERE NA01(+)=FA10"
      'end 2020/5/12
      'Modify by Morgan 2006/2/14 從外層移進來並修改
      If Len(Trim(frm100106_1.txt5(11).Text)) <> 0 Then
           'Modified by Lydia 2023/09/11 區分FMP案管制人---by Phoebe
           'strSql = strSql & " AND NA16 >='" & frm100106_1.txt5(11).Text & "' "
           strSql = strSql & " and (((CP01='FCP' or CP01='FG') and nvl(fa10,'000') >'010' and NA16 >='" & frm100106_1.txt5(11).Text & "') " & _
                   " or ((CP01='P' or CP01='PS') and nvl(fa10,'000') >'010' and NA79 >='" & frm100106_1.txt5(11).Text & "'))"
      End If
      If Len(Trim(frm100106_1.txt5(12).Text)) <> 0 Then
           'Modified by Lydia 2023/09/11 區分FMP案管制人---by Phoebe
           'strSql = strSql & " AND NA16 <='" & frm100106_1.txt5(12).Text & "' "
           strSql = strSql & " and (((CP01='FCP' or CP01='FG') and nvl(fa10,'000') >'010' and NA16 <='" & frm100106_1.txt5(12).Text & "') " & _
                   " or ((CP01='P' or CP01='PS') and nvl(fa10,'000') >'010' and NA79 <='" & frm100106_1.txt5(12).Text & "'))"
      End If
   '2006/2/14 END
   Else
      '2010/9/13 MODIFY BY SONIA 日期欄改百年日期排序問題
      'Modify by Amy 2018/09/20 申請人資料全顯示不限制字數 原:SUBSTRB(CU04,1,10)
      'Modified by Lydia 2019/11/01 +增加欄位SeColPA, ApplyNo
      'Modified by Lydia 2021/05/19 在CP09後面加上CP66
      strSql = "SELECT '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                ", PA26 as ApplyNo " & SeColPA & " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                " WHERE (PA57<>'Y' or pa57 is null) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL1 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
      'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查並顯示申請人資料
      If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            'Modified by Lydia 2021/05/19 在CP09後面加上CP66
            strSql = strSql & " union all SELECT '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", PA27 as ApplyNo " & SeColPA & " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                      " WHERE (PA57<>'Y' or pa57 is null) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA27,1,8)=CU01(+) AND DECODE(SUBSTR(PA27,9,1),NULL,'0',SUBSTR(PA27,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL10(1) & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all SELECT '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", PA28 as ApplyNo " & SeColPA & " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                      " WHERE (PA57<>'Y' or pa57 is null) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA28,1,8)=CU01(+) AND DECODE(SUBSTR(PA28,9,1),NULL,'0',SUBSTR(PA28,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL10(2) & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all SELECT '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", PA29 as ApplyNo " & SeColPA & " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                      " WHERE (PA57<>'Y' or pa57 is null) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA29,1,8)=CU01(+) AND DECODE(SUBSTR(PA29,9,1),NULL,'0',SUBSTR(PA29,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL10(3) & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all SELECT '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", PA30 as ApplyNo " & SeColPA & " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                      " WHERE (PA57<>'Y' or pa57 is null) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA30,1,8)=CU01(+) AND DECODE(SUBSTR(PA30,9,1),NULL,'0',SUBSTR(PA30,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL10(4) & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
      End If
      'end 2019/11/18
      
      'Modified by Lydia 2019/11/01 +增加欄位SeColTM, ApplyNo
      'Modified by Lydia 2021/05/19 在CP09後面加上CP66
      strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                ", TM23 as ApplyNo " & SeColTM & " FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                " WHERE (TM29<>'Y' or tm29 is null) AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(TM23,1,8)=CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & strSQL2 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "")
      'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查並顯示申請人資料
      If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            'Modified by Lydia 2021/05/19 在CP09後面加上CP66
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", TM78 as ApplyNo " & SeColTM & " FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                      " WHERE (TM29<>'Y' or tm29 is null) AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(TM78,1,8)=CU01(+) AND DECODE(SUBSTR(TM78,9,1),NULL,'0',SUBSTR(TM78,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql20(1) & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", TM79 as ApplyNo " & SeColTM & " FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                      " WHERE (TM29<>'Y' or tm29 is null) AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(TM79,1,8)=CU01(+) AND DECODE(SUBSTR(TM79,9,1),NULL,'0',SUBSTR(TM79,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql20(2) & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", TM80 as ApplyNo " & SeColTM & " FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                      " WHERE (TM29<>'Y' or tm29 is null) AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(TM80,1,8)=CU01(+) AND DECODE(SUBSTR(TM80,9,1),NULL,'0',SUBSTR(TM80,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql20(3) & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", TM81 as ApplyNo " & SeColTM & " FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                      " WHERE (TM29<>'Y' or tm29 is null) AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(TM81,1,8)=CU01(+) AND DECODE(SUBSTR(TM81,9,1),NULL,'0',SUBSTR(TM81,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql20(4) & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "")
      End If
      'end 2019/11/18
      
       'Modified by Lydia 2019/11/01 +增加欄位SeColLC, ApplyNo
       'Modified by Lydia 2021/05/19 在CP09後面加上CP66
       strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                ",LC11 as ApplyNo " & SeColLC & " FROM CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                " WHERE (LC08<>'Y' or lc08 is null) AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSQL3 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
      'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查並顯示申請人資料
      If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            'Modified by Lydia 2021/05/19 在CP09後面加上CP66
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",LC43 as ApplyNo " & SeColLC & " FROM CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (LC08<>'Y' or lc08 is null) AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(LC43,1,8)=CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql30(1) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",LC44 as ApplyNo " & SeColLC & " FROM CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (LC08<>'Y' or lc08 is null) AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(LC44,1,8)=CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql30(2) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",LC45 as ApplyNo " & SeColLC & " FROM CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (LC08<>'Y' or lc08 is null) AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(LC45,1,8)=CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql30(3) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",LC46 as ApplyNo " & SeColLC & " FROM CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (LC08<>'Y' or lc08 is null) AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(LC45,1,8)=CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql30(4) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
      End If
      'end 2019/11/18
      
      'Modified by Lydia 2019/11/01 +增加欄位SeColHC, ApplyNo
       'Modified by Lydia 2021/05/19 在CP09後面加上CP66
       strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                          AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                ",HC05 as ApplyNo " & SeColHC & " FROM CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,staff s1,engineerprogress,STAFF_IDMAP " & _
                " WHERE (HC09<>'Y' or hc09 is null) AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(HC05,1,8)=CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=CU02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSQL4 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
      'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查並顯示申請人資料
      If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            'Modified by Lydia 2021/05/19 在CP09後面加上CP66
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                          AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",HC24 as ApplyNo " & SeColHC & " FROM CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,staff s1,engineerprogress,STAFF_IDMAP " & _
                     " WHERE (HC09<>'Y' or hc09 is null) AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(HC24,1,8)=CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=CU02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSql40(1) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                          AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",HC25 as ApplyNo " & SeColHC & " FROM CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,staff s1,engineerprogress,STAFF_IDMAP " & _
                     " WHERE (HC09<>'Y' or hc09 is null) AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(HC25,1,8)=CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=CU02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSql40(2) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                          AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",HC26 as ApplyNo " & SeColHC & " FROM CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,staff s1,engineerprogress,STAFF_IDMAP " & _
                     " WHERE (HC09<>'Y' or hc09 is null) AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(HC26,1,8)=CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=CU02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSql40(3) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                          AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",HC27 as ApplyNo " & SeColHC & " FROM CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,staff s1,engineerprogress,STAFF_IDMAP " & _
                     " WHERE (HC09<>'Y' or hc09 is null) AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(HC27,1,8)=CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=CU02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSql40(4) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
      End If
      'end 2019/11/18
      
       'Modified by Lydia 2019/11/01 +增加欄位SeColSP, ApplyNo
       'Modified by Lydia 2021/05/19 在CP09後面加上CP66
       strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                ",SP08 as ApplyNo " & SeColSP & " FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                " WHERE (SP15<>'Y' or sp15 is null) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL5 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
       'end 2018/09/20
      'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查並顯示申請人資料
      If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            'Modified by Lydia 2021/05/19 在CP09後面加上CP66
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",SP58 as ApplyNo " & SeColSP & " FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (SP15<>'Y' or sp15 is null) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(SP58,1,8)=CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & StrSql50(1) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",SP59 as ApplyNo " & SeColSP & " FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (SP15<>'Y' or sp15 is null) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(SP59,1,8)=CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & StrSql50(2) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",SP65 as ApplyNo " & SeColSP & " FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (SP15<>'Y' or sp15 is null) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(SP65,1,8)=CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & StrSql50(3) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",SP66 as ApplyNo " & SeColSP & " FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND,STAFF_IDMAP " & _
                     " WHERE (SP15<>'Y' or sp15 is null) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) AND SIM02(+)=CP14 and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(SP66,1,8)=CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & StrSql50(4) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
      End If
      'end 2019/11/18
     End If
     
     'Added by Morgan 2024/3/11
     If InStr(strSql, "CP06>=") > 0 Or InStr(strSql, "CP06<=") > 0 Then
         strSql = strSql & " union " & Replace(Replace(strSql, "CP06>=", "CP142>="), "CP06<=", "CP142<=")
     End If
     'end 2024/3/11
     
     '限制使用者所能使用的系統類別+案件性質
     'Modify by Morgan 2004/11/10 要過濾已收不續辦的來函
     'Modified by Morgan 2015/10/20 取消系統類別+案件性質的權限控制--秀玲檢查文件沒有該限制
     'strSql = "Select AA.V AS V, " & _
               IIf(frm100106_1.opt1(3).Value, "AA.承辦期限 AS 承辦期限", "AA.本所期限 AS 本所期限") & _
               ",AA.本所案號 AS 本所案號,AA.分所號 as 分所號, AA.案件名稱 AS 案件名稱, AA.案件性質 AS 案件性質, AA.承辦人 AS 承辦人, AA.智權人員 AS 智權人員, AA.收文日 AS 收文日, " & _
               IIf(frm100106_1.opt1(3).Value, "AA.本所期限 AS 本所期限", "AA.法定期限 AS 法定期限") & _
               ", AA.進度備註 AS 進度備註, AA.申請人 AS 申請人, AA.是否出名 AS 是否出名, AA.延期日 AS 延期日, AA.申請國家 AS 申請國家, AA.申請人國籍 AS 申請人國籍, AA.發文日 AS 發文日, AA.代理人 AS 代理人, AA.彼所案號 AS 彼所案號, AA.申請案號 AS 申請案號, AA.承辦人備註 AS 承辦人備註, AA.取消收文日 AS 取消收文日, AA.cp13 AS CP13, AA.cp09 AS CP09, AA.cp14 AS CP14, AA.CP01 AS CP01, AA.CP10 AS CP10 " & _
              " From ( " & strSql & " ) AA, Staff SA, Staff_Group, Staff SB Where AA.CP01=SG02 And AA.CP10=SG03 And SA.ST11=SG01 And AA.CP13=SB.ST01(+) And SA.ST01='" & strUserNum & "' " & _
              " and not exists(select * from caseprogress a where a.cp43= AA.cp09 and a.cp10='907')"
     'Modified by Lydia 2019/11/01 利益衝突案件：於CP10後面增加欄位
     'Modified by Lydia 2021/05/19 限定欄位長度
     'strSql = "Select AA.V AS V, " & _
               IIf(frm100106_1.OPT1(3).Value, "AA.承辦期限 AS 承辦期限", "AA.本所期限 AS 本所期限") & _
               ",AA.本所案號 AS 本所案號,AA.分所號 as 分所號, AA.案件名稱 AS 案件名稱, AA.案件性質 AS 案件性質, AA.承辦人 AS 承辦人, AA.智權人員 AS 智權人員, AA.收文日 AS 收文日, " & _
               IIf(frm100106_1.OPT1(3).Value, "AA.本所期限 AS 本所期限", "AA.法定期限 AS 法定期限") & _
               ", AA.進度備註 AS 進度備註, AA.申請人 AS 申請人, AA.是否出名 AS 是否出名, AA.延期日 AS 延期日, AA.申請國家 AS 申請國家, AA.申請人國籍 AS 申請人國籍, AA.發文日 AS 發文日, AA.代理人 AS 代理人, AA.彼所案號 AS 彼所案號, AA.申請案號 AS 申請案號, AA.承辦人備註 AS 承辦人備註, AA.取消收文日 AS 取消收文日, AA.cp13 AS CP13, AA.cp09 AS CP09, AA.cp14 AS CP14, AA.CP01 AS CP01, AA.CP10 AS CP10 " & _
               ", ApplyNo, cust01, cust02, cust03, cust04, cust05, fcno " & _
              " From ( " & strSql & " ) AA Where not exists(select * from caseprogress a where a.cp43= AA.cp09 and a.cp10='907')"
     'end 2015/10/20
     'Modified by Morgan 2021/5/25 補 , Staff SB (依智權人員排序會用)
     'Modified by Morgan 2024/1/8 +指定日期
     'Modified by Morgan 2024/3/11 +distinct (因加上指定日期後資料可能會重複)
     strSql = "Select distinct AA.V AS V, " & _
               IIf(frm100106_1.opt1(3).Value, "substr(AA.承辦期限,1,10) AS 承辦期限", "substr(AA.本所期限,1,10) AS 本所期限") & _
               IIf(strSrvDate(1) >= 指定日期啟用日 And frm100106_1.opt1(0).Value, ",substr(AA.指定日期,1,10) AS 指定日期", "") & _
               ",AA.本所案號 AS 本所案號,AA.分所號 as 分所號, AA.案件名稱 AS 案件名稱, AA.案件性質 AS 案件性質, AA.承辦人 AS 承辦人, AA.智權人員 AS 智權人員, substr(AA.收文日,1,10) AS 收文日, " & _
               IIf(frm100106_1.opt1(3).Value, "substr(AA.本所期限,1,10) AS 本所期限", "substr(AA.法定期限,1,10) AS 法定期限") & _
               ", substr(AA.進度備註,1,500) AS 進度備註, AA.申請人 AS 申請人, AA.是否出名 AS 是否出名, substr(AA.延期日,1,10) AS 延期日, AA.申請國家 AS 申請國家, AA.申請人國籍 AS 申請人國籍, substr(AA.發文日,1,10) AS 發文日, " & _
               "AA.代理人 AS 代理人, AA.彼所案號 AS 彼所案號, AA.申請案號 AS 申請案號, substr(AA.承辦人備註,1,500) AS 承辦人備註, substr(AA.取消收文日,1,10) AS 取消收文日, AA.cp13 AS CP13, AA.cp09 AS CP09,  AA.cp66 AS CP66, " & _
               "AA.cp14 AS CP14, AA.CP01 AS CP01, AA.CP10 AS CP10 , ApplyNo, cust01, cust02, cust03, cust04, cust05, fcno " & _
              " From ( " & strSql & " ) AA, Staff SB Where AA.CP13=SB.ST01(+) and not exists(select * from caseprogress a where a.cp43= AA.cp09 and a.cp10='907')"
     
     '是否依智權人員排序
     If frm100106_1.txt5(13).Text = "Y" Then
         'edit by nickc 2008/02/21 取消收文日放最下面
         'strSQL = strSQL & " ORDER BY SB.ST03, SB.ST01, 本所期限,本所案號 "
         '2010/9/17 modify by sonia 因改日期欄百年問題故調整
         'strSql = strSql & " ORDER BY SB.ST03, SB.ST01, nvl(取消收文日,'11/11/11'),本所期限,本所案號 "
         strSql = strSql & " ORDER BY SB.ST03, SB.ST01, nvl(取消收文日,' 11/11/11')," & IIf(frm100106_1.opt1(3).Value, "承辦期限", "本所期限") & ",本所案號 "
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(9) & frm100106_1.txt5(13)  'Add By Sindy 2010/01/22
     Else
         'edit by nickc 2008/02/21 取消收文日放最下面
         'strSQL = strSQL & " ORDER BY 本所期限,本所案號 "
         '2010/9/17 modify by sonia 因改日期欄百年問題故調整
         'strSql = strSql & " ORDER BY nvl(取消收文日,'11/11/11'),本所期限,本所案號 "
         strSql = strSql & " ORDER BY nvl(取消收文日,' 11/11/11')," & IIf(frm100106_1.opt1(3).Value, "承辦期限", "本所期限") & ",本所案號 "
     End If
    
    SetDataListWidth 'Added by Lydia 2021/05/19 清空欄位
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    'Modified by Lydia 2019/11/01 改變型態
    'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
    If adoRecordset.RecordCount <> 0 Then
       dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
        'Added by Lydia 2019/11/01 逐案號判斷
        If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    adoRecordset.Delete
                End If
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
                pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            InsertQueryLog (dblRow) 'Add By Sindy 2010/01/22
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
        Else
           InsertQueryLog (dblRow) 'Add By Sindy 2010/01/22
        End If
        'end 2019/11/01
        
        cmdOK(0).Enabled = True
        cmdOK(1).Enabled = True
        cmdOK(2).Enabled = True
        cmdOK(3).Enabled = True
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/01/22
JumpToNoData:   'Added by Lydia 2019/11/01
        ShowNoData
        Screen.MousePointer = vbDefault
        cmdOK(0).Enabled = False
        cmdOK(1).Enabled = False
        cmdOK(2).Enabled = False
        cmdOK(3).Enabled = False
        StrMenu_sub1 = False
        Exit Function
    End If
    Set grdDataList.Recordset = adoRecordset
    Call SetDataListWidth(False) 'Added by Lydia 2021/05/19 預設欄位
    
    intK = grdDataList.Rows - 1
    CheckOC
    grdDataList.Visible = False
   For i = 1 To grdDataList.Rows - 1
         'Modified by Lydia 2021/05/19 改用變數取得
         'Me.grdDataList.TextMatrix(i, 5) = Me.grdDataList.TextMatrix(i, 5) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(i, 23), "1")
         grdDataList.TextMatrix(i, colCp10Name) = grdDataList.TextMatrix(i, colCp10Name) & PUB_GetNextCasePropertyName(grdDataList.TextMatrix(i, colCp09), grdDataList.TextMatrix(i, colCP66), "1")
         
         grdDataList.row = i
         'Modified by Lydia 2021/05/19 改用變數取得
         'grdDataList.col = grdDataList.Cols - 4
         grdDataList.col = colCp09
         strSql = "SELECT " & SQLDate("DL02") & " FROM DATELIMIT WHERE DL01='" & grdDataList.Text & "' ORDER BY DL02"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         'Modified by Lydia 2021/05/19 改用變數取得: 13=> colDDate3
         grdDataList.col = colDDate3
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
             adoRecordset.MoveLast
             If Not IsNull(adoRecordset.Fields(0)) Then
                  grdDataList.Text = adoRecordset.Fields(0)
             Else
                  grdDataList.Text = ""
             End If
         End If
           'Modified by Lydia 2021/05/19 改用變數取得
'           If Trim(grdDataList.TextMatrix(grdDataList.row, 21)) <> "" Then
'               'add by nickc 2008/02/21 灰色+x
'               grdDataList.TextMatrix(grdDataList.row, 2) = "x" & grdDataList.TextMatrix(grdDataList.row, 2)
           If Trim(grdDataList.TextMatrix(grdDataList.row, colCP57)) <> "" Then
               grdDataList.TextMatrix(grdDataList.row, colCaseNo) = "x" & grdDataList.TextMatrix(grdDataList.row, colCaseNo)
           'end 2021/05/19
               For j = 1 To grdDataList.Cols - 1
                  grdDataList.col = j
                  '灰色
                  'Modified by Lydia 2024/10/04 加深灰色&H8000000F>>&HC0C0C0
                  grdDataList.CellBackColor = &HC0C0C0
               Next j
           End If
           'Modified by Lydia 2021/05/19 改用變數取得: 3=>colCaseNo
           grdDataList.col = colCaseNo
           If grdDataList.CellBackColor = -2147483643 Then
               'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseNo
               stCaseNo = grdDataList.TextMatrix(grdDataList.row, colCaseNo)
               For j = 1 To 4
                   iPos = InStr(stCaseNo, "-")
                   If iPos > 0 Then
                      stPA(j) = Left(stCaseNo, iPos - 1)
                      stCaseNo = Mid(stCaseNo, iPos + 1)
                   Else
                      stPA(j) = stCaseNo
                   End If
               Next j
               'Modified by Lydia 2021/05/19 改用變數取得
'               If (UCase(stPA(1)) = "FCP" Or UCase(stPA(1)) = "P") And grdDataList.TextMatrix(grdDataList.row, 26) = "605" Then
'                   '2010/9/20 modify by sonia 因日期加空格故加val
'                   strTmp = Val(ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, 9))))
               If (UCase(stPA(1)) = "FCP" Or UCase(stPA(1)) = "P") And grdDataList.TextMatrix(grdDataList.row, colCP10) = "605" Then
                   strTmp = Val(ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, colDDate2))))
               'end 2021/05/19
                   If strTmp <> "" Then
                       If PUB_IfCtrlDateExtended(stPA, strTmp) = True Then
                           'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseNo
                           grdDataList.TextMatrix(grdDataList.row, colCaseNo) = "&" & grdDataList.TextMatrix(grdDataList.row, colCaseNo)
                           For j = 1 To grdDataList.Cols - 1
                              grdDataList.col = j
                              '紫色
                              grdDataList.CellBackColor = &HE600E6
                           Next j
                       End If
                   End If
               End If
           End If
           'Modified by Lydia 2021/05/19 改用變數取得: 3=>colCaseNo
           grdDataList.col = colCaseNo
           If grdDataList.CellBackColor = -2147483643 Then
               'add by nickc 2008/02/15 修正 加入黃色
               'Modified by Lydia 2021/05/19 改用變數取得
'               If Mid(Trim(grdDataList.TextMatrix(grdDataList.row, 23)), 1, 1) = "C" And Trim(grdDataList.TextMatrix(grdDataList.row, 16)) = "" Then
'                   'add by nickc 2008/02/21 黃色+#
'                   grdDataList.TextMatrix(grdDataList.row, 2) = "#" & grdDataList.TextMatrix(grdDataList.row, 2)
               If Mid(Trim(grdDataList.TextMatrix(grdDataList.row, colCp09)), 1, 1) = "C" And Trim(grdDataList.TextMatrix(grdDataList.row, colCP27)) = "" Then
                   grdDataList.TextMatrix(grdDataList.row, colCaseNo) = "#" & grdDataList.TextMatrix(grdDataList.row, colCaseNo)
               'end 2021/05/19
                   For j = 1 To grdDataList.Cols - 1
                       grdDataList.col = j
                       '黃色
                       grdDataList.CellBackColor = &HFFFF&
                   Next j
             'edit by nickc 2008/02/21 修正字串比對錯誤
             'ElseIf ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, 1)) < ChangeWStringToTString(ServerDate) And Trim(grdDataList.TextMatrix(grdDataList.row, 1)) <> "" Then
             'Modified by Lydia 2021/05/19 改用變數取得
'             ElseIf Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, 1))) < Val(ChangeWStringToTString(ServerDate)) And Trim(grdDataList.TextMatrix(grdDataList.row, 1)) <> "" Then
'                 grdDataList.col = 2
             'Modified by Morgan 2024/10/8 指定日期比照本所期限變色
             'ElseIf Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, colDDate1))) < Val(ChangeWStringToTString(ServerDate)) And Trim(grdDataList.TextMatrix(grdDataList.row, colDDate1)) <> "" Then
             ElseIf (Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, colDDate1))) < Val(ChangeWStringToTString(strSrvDate(1))) And Trim(grdDataList.TextMatrix(grdDataList.row, colDDate1)) <> "") _
               Or (Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, colDDate4))) < Val(ChangeWStringToTString(strSrvDate(1))) And Trim(grdDataList.TextMatrix(grdDataList.row, colDDate4)) <> "") Then
            'end 2024/10/8
                 grdDataList.col = colCaseNo
             'end 2021/05/19
                 grdDataList.Text = "*" + grdDataList.Text
                 For j = 1 To grdDataList.Cols - 1
                     grdDataList.col = j
                     '紅色
                     grdDataList.CellBackColor = &HFF&
                 Next j
                 GoTo NextItem1
             '2010/9/20 modify by sonia 因日期加空格故加val
             'Modified by Lydia 2021/05/19 改用變數取得
'             ElseIf Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, 1))) = Val(ChangeWStringToTString(ServerDate)) Then
'                     grdDataList.col = 2
             'Modified by Morgan 2024/10/8 指定日期比照本所期限變色
             'ElseIf Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, colDDate1))) = Val(ChangeWStringToTString(ServerDate)) Then
             ElseIf DBDATE(grdDataList.TextMatrix(grdDataList.row, colDDate1)) = strSrvDate(1) Or DBDATE(grdDataList.TextMatrix(grdDataList.row, colDDate4)) = strSrvDate(1) Then
             'end 2024/10/8
                     grdDataList.col = colCaseNo
             'end 2021/05/19
                     grdDataList.Text = "v" & grdDataList.Text
                     For j = 1 To grdDataList.Cols - 1
                         grdDataList.col = j
                         '橙色
                         grdDataList.CellBackColor = &H80FF&
                     Next j
                    GoTo NextItem1
           Else
               'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseNo
               grdDataList.col = colCaseNo
               If grdDataList.CellBackColor = -2147483643 Then
                   'Modified by Lydia 2021/05/19 改用變數取得
'                   If UCase(SystemNumber(grdDataList.TextMatrix(grdDataList.row, 2), 1)) = "FCP" Or UCase(SystemNumber(grdDataList.TextMatrix(grdDataList.row, 2), 1)) = "P" Then
'                       grdDataList.col = 26
                   If UCase(SystemNumber(grdDataList.TextMatrix(grdDataList.row, colCaseNo), 1)) = "FCP" Or UCase(SystemNumber(grdDataList.TextMatrix(grdDataList.row, colCaseNo), 1)) = "P" Then
                       grdDataList.col = colCP10
                   'end 2021/05/19
                       If grdDataList.Text = "605" Then
                           CheckOC
                           'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseNo
                           grdDataList.col = colCaseNo
                           strSql = "select pa09,pa14 from patent where pa01='" & SystemNumber(grdDataList.Text, 1) & _
                           "' and pa02='" & SystemNumber(grdDataList.Text, 2) & "' and pa03='" & _
                           SystemNumber(grdDataList.Text, 3) & "' and pa04='" & SystemNumber(grdDataList.Text, 4) & "' and pa09='000' "
                           adoRecordset.CursorLocation = adUseClient
                           adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                           If adoRecordset.RecordCount <> 0 Then
                               If CheckStr(adoRecordset.Fields(0).Value) = "000" And CheckStr(adoRecordset.Fields(1).Value) = "" Then
                                   'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseNo
                                   grdDataList.col = colCaseNo
                                   grdDataList.Text = "!" & grdDataList.Text
                                   For j = 1 To grdDataList.Cols - 1
                                       grdDataList.col = j
                                       '綠色
                                       grdDataList.CellBackColor = &HC000&
                                   Next j
                               End If
                           End If
                           CheckOC
                       End If
                   End If
               End If
           End If
           
NextItem1:
           'Add By Sindy 2025/2/11 淑華提,共同查詢'以期限管制日查詢這支，已有指定日期，請日期後帶出 之前/當天/之後 (同期限通知的顯示)
           If Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, colDDate4))) > 0 Then
               strSql = "SELECT CP09,CP142,decode(cp164,'1','當天','2','之前','3','之後','') FROM caseprogress " & _
                     "WHERE CP09 = '" & grdDataList.TextMatrix(grdDataList.row, colCp09) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  grdDataList.TextMatrix(grdDataList.row, colDDate4) = grdDataList.TextMatrix(grdDataList.row, colDDate4) & _
                        RsTemp.Fields(2)
               End If
           End If
           '2025/2/11 END
       End If
   Next i
   grdDataList.Visible = True
End Function

'已收文已發文
Public Function StrMenu_sub2(StrSQLa As String) As Boolean
Dim strTmp As String
Dim stPA(1 To 4) As String, iPos As Integer, stCaseNo As String
Dim strDate1Col As String, strDate2Col As String 'Add By Sindy 2013/8/13
'Added by Lydia 2019/11/18 比照StrMeu => Add by Amy 2016/07/18 for 申請人1~5用
Dim strSQL10(1 To 4) As String, StrSql20(1 To 4) As String, StrSql30(1 To 4) As String, StrSql40(1 To 4) As String, StrSql50(1 To 4) As String
'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String
Dim dblRow As Double 'Add By Sindy 2025/9/3
   
     StrMenu_sub2 = True
     
     pub_QL05 = pub_QL05 & ";" & frm100106_1.Frame1.Caption & frm100106_1.opt2(2).Caption 'Add By Sindy 2010/01/22
     CheckOC
     strSQL1 = ""
     strSQL2 = ""
     StrSQL3 = ""
     StrSQL4 = ""
     strSQL5 = ""
      'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
      SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
      SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
      SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
      SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
      SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
      m_AllSys = IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(, frm100106_1.txt5(0).Text))
      'Added by Lydia 2020/03/11 外專人員勾選含FMP外專管制期限時,要另外增加P,PS
      If frm100106_1.Check1.Value = 1 And Left(Pub_StrUserSt03, 2) = "F2" Then
          If InStr(m_AllSys & ",", ",P,") = 0 Then m_AllSys = m_AllSys & ",P,"
          If InStr(m_AllSys & ",", ",PS,") = 0 Then m_AllSys = m_AllSys & ",PS,"
          m_AllSys = Replace(m_AllSys, ",,", ",")
      End If
      'end 2020/03/11
      intCufaCnt = 0
      'end 2019/11/01
      
      'Add By Sindy 2013/8/13
      strDate1Col = "SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限"
      strDate2Col = "SUBSTR(' '||sqldatet(CP07),-9) AS 法定期限"
      '以承辦期限查詢
      If frm100106_1.opt1(3).Value Then
          strDate1Col = "SUBSTR(' '||sqldatet(CP48),-9) AS 承辦期限"
          strDate2Col = "SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限"
          If Len(Trim(frm100106_1.txt6(0))) <> 0 Then
              strSQL1 = strSQL1 + " and CP48>=" & Val(ChangeTStringToWString(frm100106_1.txt6(0))) & " "
          End If
          pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(3).Caption & frm100106_1.txt6(0) 'Add By Sindy 2010/01/22
          If Len(Trim(frm100106_1.txt6(1))) <> 0 Then
              strSQL1 = strSQL1 + " and CP48<=" & Val(ChangeTStringToWString(frm100106_1.txt6(1))) & " "
              pub_QL05 = pub_QL05 & "-" & frm100106_1.txt6(1) 'Add By Sindy 2010/01/22
          Else
              strSQL1 = strSQL1 + " and CP48<=" & Val(ChangeTStringToWString(strSrvDate(1) - 19110000)) & " "
              pub_QL05 = pub_QL05 & "-" & (strSrvDate(1) - 19110000) 'Add By Sindy 2010/01/22
          End If
          '有完稿日者,代表已承辦,不需顯示出來
          strSQL1 = strSQL1 & " and ep09 is null"
          strSQL2 = strSQL2 & " and ep09 is null"
          StrSQL3 = StrSQL3 & " and ep09 is null"
          StrSQL4 = StrSQL4 & " and ep09 is null"
          strSQL5 = strSQL5 & " and ep09 is null"
      End If
      '2013/8/13 END
      '以本所期限查詢
      If frm100106_1.opt1(0).Value Then
          If Len(Trim(frm100106_1.txt1(0))) <> 0 Then
              strSQL1 = strSQL1 + " and CP06>=" & Val(ChangeTStringToWString(frm100106_1.txt1(0))) & " "
          End If
          pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(0).Caption & frm100106_1.txt1(0) 'Add By Sindy 2010/01/22
          If Len(Trim(frm100106_1.txt1(1))) <> 0 Then
              strSQL1 = strSQL1 + " and CP06<=" & Val(ChangeTStringToWString(frm100106_1.txt1(1))) & " "
              pub_QL05 = pub_QL05 & "-" & frm100106_1.txt1(1) 'Add By Sindy 2010/01/22
          Else
              strSQL1 = strSQL1 + " and CP06<=" & Val(ChangeTStringToWString(strSrvDate(1) - 19110000)) & " "
              pub_QL05 = pub_QL05 & "-" & (strSrvDate(1) - 19110000) 'Add By Sindy 2010/01/22
          End If
      End If
     '以本所案號查詢
     If frm100106_1.opt1(2).Value Then
        If Len(Trim(frm100106_1.txt3(0))) <> 0 Then
            strSQL1 = strSQL1 + " AND CP01='" & frm100106_1.txt3(0) & "' "
        End If
        pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(2).Caption & frm100106_1.txt3(0) 'Add By Sindy 2010/01/22
        If Len(Trim(frm100106_1.txt3(1))) <> 0 Then
            strSQL1 = strSQL1 + " AND CP02='" & frm100106_1.txt3(1) & "' "
        End If
        pub_QL05 = pub_QL05 & "-" & frm100106_1.txt3(1) 'Add By Sindy 2010/01/22
        If Len(Trim(frm100106_1.txt3(2))) <> 0 Then
            strSQL1 = strSQL1 + " AND CP03='" & frm100106_1.txt3(2) & "' "
            pub_QL05 = pub_QL05 & "-" & frm100106_1.txt3(2) 'Add By Sindy 2010/01/22
        Else
            strSQL1 = strSQL1 + " AND CP03='0' "
            pub_QL05 = pub_QL05 & "-" & "0" 'Add By Sindy 2010/01/22
        End If
        If Len(Trim(frm100106_1.txt3(3))) <> 0 Then
            strSQL1 = strSQL1 + " AND CP04='" & frm100106_1.txt3(3) & "' "
            pub_QL05 = pub_QL05 & "-" & frm100106_1.txt3(3) 'Add By Sindy 2010/01/22
        Else
            strSQL1 = strSQL1 + " AND CP04='00' "
            pub_QL05 = pub_QL05 & "-" & "00" 'Add By Sindy 2010/01/22
        End If
    End If
    If Len(Trim(frm100106_1.txt5(1))) <> 0 Then
     strSQL1 = strSQL1 & " AND CP14='" & frm100106_1.txt5(1) & "' "
     pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(2) & frm100106_1.txt5(1) & frm100106_1.LBL1(0) 'Add By Sindy 2010/01/22
    End If
   'Add by Morgan 2007/9/21 加承辦人組別
   If Len(Trim(frm100106_1.txt5(14))) <> 0 Then
        strSQL1 = strSQL1 & " AND S1.ST16='" & frm100106_1.txt5(14) & "' "
        pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(10) & frm100106_1.txt5(14)  'Add By Sindy 2010/01/22
   End If
   'end 2007/9/21
    'Modify by Amy 2016/09/12 避免抓到錯的index,所以+||''
    If Len(Trim(frm100106_1.txt5(2))) <> 0 Then
         strSQL1 = strSQL1 & " AND cp12||''>='" & frm100106_1.txt5(2) & "' "
    End If
    If Len(Trim(frm100106_1.txt5(3))) <> 0 Then
         strSQL1 = strSQL1 & " AND cp12||''<='" & frm100106_1.txt5(3) & "' "
    End If
    
      'Added by Morgan 2015/10/21 FMP案改可選擇
      If frm100106_1.Check1.Value = 1 Then
           pub_QL05 = pub_QL05 & ";" & frm100106_1.Check1.Caption
      Else
           'Modify By Sindy 2023/10/30 EP33要回歸用在英文核完日,改抓EP39.核稿完成日
           strSQL1 = strSQL1 & " and (cp01 not in ('P','PS','CFP','CPS') or substr(cp12,1,1)<>'F' or substr(s1.st03,1,1)<>'F' or (cp10<>'201' and ep09>0) or (cp10='201' and " & IIf(strSrvDate(1) >= FCP核完日改用EP39, "ep39", "ep33") & ">0)) "
      End If
      'end 2015/10/21
       
    If Len(Trim(frm100106_1.txt5(2))) <> 0 Or Len(Trim(frm100106_1.txt5(3))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(3) & frm100106_1.txt5(2) & "-" & frm100106_1.txt5(3) 'Add By Sindy 2010/01/22
    End If
    If Len(Trim(frm100106_1.txt5(4))) <> 0 Then
       '2008/3/31 MODIFY BY SONIA 加控制查詢87027陳淑芳同時查20001台中所
       'strSQL1 = strSQL1 + " AND CP13='" & frm100106_1.txt5(4) & "' "
       If frm100106_1.txt5(4) = "87027" Then
          strSQL1 = strSQL1 + " AND CP13 IN ('87027','20001') "
       Else
          strSQL1 = strSQL1 + " AND CP13='" & frm100106_1.txt5(4) & "' "
       End If
       '2008/3/31 END
       pub_QL05 = pub_QL05 & ";" & frm100106_1.Label2(0) & frm100106_1.txt5(4) & frm100106_1.LBL1(1) 'Add By Sindy 2010/01/22
    End If
    '申請人國籍
    If Len(Trim(frm100106_1.txt5(9))) <> 0 Then
        strSQL1 = strSQL1 + " AND CU10>='" & frm100106_1.txt5(9) & "' "
    End If
    If Len(Trim(frm100106_1.txt5(10))) <> 0 Then
        strSQL1 = strSQL1 + " AND CU10<='" & frm100106_1.txt5(10) & "' "
    End If
    If Len(Trim(frm100106_1.txt5(9))) <> 0 Or Len(Trim(frm100106_1.txt5(10))) <> 0 Then
        pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(1) & frm100106_1.txt5(9) & "-" & frm100106_1.txt5(10) 'Add By Sindy 2010/01/22
    End If
   strSQL1 = strSQL1 + " and CP27 IS not NULL "
   strSQL2 = strSQL1
   StrSQL3 = strSQL1
   StrSQL4 = strSQL1
   strSQL5 = strSQL1
   
   strSQL1 = strSQL1 & " and cp10<>'926'" 'Added by Morgan 2015/10/20 排除核對已准專利
   strSQL1 = strSQL1 + " and cp10<>'1920'"  'Added by Lydia 2018/03/13 排除客戶提供文件
   
   If Len(Trim(frm100106_1.txt5(0))) <> 0 Then
         'Added by Morgan 2015/10/21
         '外專人員勾選含FMP外專管制期限時
         If frm100106_1.Check1.Value = 1 And Left(Pub_StrUserSt03, 2) = "F2" Then
            strSQL1 = strSQL1 & " AND (cp01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 1) & ") or (cp01 IN ('P','CFP') AND substr(cp12,1,1)='F')) "
            strSQL5 = strSQL5 & " AND (cp01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 5) & ") or (cp01 IN ('PS','CPS') AND substr(cp12,1,1)='F')) "
         Else
         'end 2015/10/21
            strSQL1 = strSQL1 & " AND cP01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 1) & ") "
            strSQL5 = strSQL5 & " AND cP01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 5) & ") "
         End If 'Added by Morgan 2015/10/21
         
        strSQL2 = strSQL2 & " AND cP01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 2) & ") "
        StrSQL3 = StrSQL3 & " AND cP01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 3) & ") "
        StrSQL4 = StrSQL4 & " AND cP01 IN (" & SQLGrpStr(IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(frm100106_1.txt5(0))), 4) & ") "
        pub_QL05 = pub_QL05 & ";" & Left(frm100106_1.Label1(4), 5) & frm100106_1.txt5(0) 'Add By Sindy 2010/01/22
   End If
   
   If Len(Trim(frm100106_1.txt4(2))) <> 0 Then
        strSQL1 = strSQL1 & " AND PA75>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
        strSQL2 = strSQL2 & " AND TM44>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
        StrSQL3 = StrSQL3 & " AND LC22>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
        strSQL5 = strSQL5 & " AND SP26>='" & GetNewFagent(frm100106_1.txt4(2)) & "' "
   End If
   If Len(Trim(frm100106_1.txt4(3))) <> 0 Then
        strSQL1 = strSQL1 & " AND PA75<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
        strSQL2 = strSQL2 & " AND TM44<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
        StrSQL3 = StrSQL3 & " AND LC22<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
        strSQL5 = strSQL5 & " AND SP26<='" & GetNewFagent(frm100106_1.txt4(3)) & "' "
   End If
   If Len(Trim(frm100106_1.txt4(2))) <> 0 Or Len(Trim(frm100106_1.txt4(3))) <> 0 Then
        pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(5) & frm100106_1.txt4(2) & "-" & frm100106_1.txt4(3)  'Add By Sindy 2010/01/22
   End If
   
'---------Move by  Lydia 2019/11/18 從If Len(Trim(frm100106_1.txt4(2))) <> 0 Then的上方移過來
    If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            strSQL10(1) = strSQL1 & " AND PA27>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            strSQL10(2) = strSQL1 & " AND PA28>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            strSQL10(3) = strSQL1 & " AND PA29>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            strSQL10(4) = strSQL1 & " AND PA30>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
         'end 2019/11/18
         strSQL1 = strSQL1 & " AND PA26>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "

         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql20(1) = strSQL2 & " AND TM78>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql20(2) = strSQL2 & " AND TM79>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql20(3) = strSQL2 & " AND TM80>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql20(4) = strSQL2 & " AND TM81>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
         'end 2019/11/18
         strSQL2 = strSQL2 & " AND TM23>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "

         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql30(1) = StrSQL3 & " AND LC43>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql30(2) = StrSQL3 & " AND LC44>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql30(3) = StrSQL3 & " AND LC45>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql30(4) = StrSQL3 & " AND LC46>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
         'end 2019/11/18
         StrSQL3 = StrSQL3 & " AND LC11>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "

         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql40(1) = StrSQL4 & " AND HC24>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql40(2) = StrSQL4 & " AND HC25>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql40(3) = StrSQL4 & " AND HC26>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql40(4) = StrSQL4 & " AND HC27>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
         'end 2019/11/18
         StrSQL4 = StrSQL4 & " AND HC05>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "

         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql50(1) = strSQL5 & " AND SP58>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql50(2) = strSQL5 & " AND SP59>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql50(3) = strSQL5 & " AND SP65>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
            StrSql50(4) = strSQL5 & " AND SP66>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
         'end 2019/11/18
         strSQL5 = strSQL5 & " AND SP08>='" & GetNewFagent(frm100106_1.txt4(0)) & "' "
    End If
    If Len(Trim(frm100106_1.txt4(1))) <> 0 Then
         strSQL1 = strSQL1 & " AND PA26<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            strSQL10(1) = strSQL10(1) & " AND PA27<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            strSQL10(2) = strSQL10(2) & " AND PA28<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            strSQL10(3) = strSQL10(3) & " AND PA29<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            strSQL10(4) = strSQL10(4) & " AND PA30<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'end 2019/11/18
         strSQL2 = strSQL2 & " AND TM23<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql20(1) = StrSql20(1) & " AND TM78<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql20(2) = StrSql20(2) & " AND TM79<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql20(3) = StrSql20(3) & " AND TM80<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql20(4) = StrSql20(4) & " AND TM81<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'end 2019/11/18
         StrSQL3 = StrSQL3 & " AND LC11<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql30(1) = StrSql30(1) & " AND LC11<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql30(2) = StrSql30(1) & " AND LC11<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql30(3) = StrSql30(1) & " AND LC11<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql30(4) = StrSql30(1) & " AND LC11<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'end 2019/11/18
         StrSQL4 = StrSQL4 & " AND HC05<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql40(1) = StrSql40(1) & " AND HC24<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql40(2) = StrSql40(2) & " AND HC25<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql40(3) = StrSql40(3) & " AND HC26<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql40(4) = StrSql40(4) & " AND HC27<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'end 2019/11/18
         strSQL5 = strSQL5 & " AND SP08<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查
            StrSql50(1) = StrSql50(1) & " AND SP58<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql50(2) = StrSql50(2) & " AND SP59<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql50(3) = StrSql50(3) & " AND SP65<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
            StrSql50(4) = StrSql50(4) & " AND SP66<='" & GetNewFagent(frm100106_1.txt4(1)) & "' "
         'end 2019/11/18
    End If
    If Len(Trim(frm100106_1.txt4(0))) <> 0 Or Len(Trim(frm100106_1.txt4(1))) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(6) & frm100106_1.txt4(0) & "-" & frm100106_1.txt4(1) 'Add By Sindy 2010/01/22
    End If
'---------end 2019/11/18

   'Add by Morgan 2006/2/14 有下FCP管制人時只抓FCP資料
   If Len(Trim(frm100106_1.txt5(11).Text)) + Len(Trim(frm100106_1.txt5(12).Text)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(8) & frm100106_1.txt5(11) & "-" & frm100106_1.txt5(12)   'Add By Sindy 2010/01/22
      '2010/9/13 MODIFY BY SONIA 日期欄改百年日期排序問題
      'Modify by Amy 2018/09/20 申請人資料全顯示不限制字數 原:SUBSTRB(CU04,1,10)
      'Modified by Lydia 2019/11/01 +增加欄位SeColPA,SeColSP; PA26,SP08 => ApplyNo
      'Modified by Lydia 2021/05/19 在CP09後面加上CP66
      strSql = "SELECT '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10,PA76,pa75 " & _
                ",PA26 as ApplyNo " & SeColPA & " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                " WHERE (PA57<>'Y' or pa57 is null) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL1 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
       strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10,NULL,sp26 " & _
                ",SP08 as ApplyNo " & SeColSP & " FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                " WHERE (SP15<>'Y' or sp15 is null) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL5 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
      'end 2018/09/20
      '年費管制人順序:pa76->pa75->pa26.cu96->pa26(PA76可能是客戶編號)
      'Modified by Lydia 2019/11/18 原PA26改為ApplyNo
      'Modified by Morgan 2020/5/12 年費管制人改抓案件代理人的管制人(不必考慮是否有年費代理人)
      'strSql = "SELECT Y.* FROM (" & _
         " select X.*,NVL(NVL(F1.FA10,C1.CU10),NVL(F2.FA10,NVL(F4.FA10,C2.CU10))) FA10" & _
         " from (" & strSql & ") X, FAGENT F1, FAGENT F2,FAGENT F4,CUSTOMER C1, CUSTOMER C2" & _
         " WHERE CP10='605' AND F1.FA01(+)=SUBSTR(PA76,1,8) AND F1.FA02(+)=SUBSTR(PA76||'0',9,1)" & _
         " AND C1.CU01(+)=SUBSTR(PA76,1,8) AND C1.CU02(+)=SUBSTR(PA76||'0',9,1)" & _
         " AND F2.FA01(+)=SUBSTR(PA75,1,8) AND F2.FA02(+)=SUBSTR(PA75||'0',9,1)" & _
         "" & _
         " AND C2.CU01(+)=SUBSTR(ApplyNo,1,8) AND C2.CU02(+)=SUBSTR(ApplyNo||'0',9,1)" & _
         " AND F4.FA01(+)=SUBSTR(C2.CU96,1,8) AND F4.FA02(+)=SUBSTR(C2.CU96||'0',9,1)" & _
         " Union All" & _
         " select X.*,NVL(FA10,CU10) FA10" & _
         " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
         " WHERE CP10<>'605' AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
         " AND CU01(+)=SUBSTR(ApplyNo,1,8) AND CU02(+)=SUBSTR(ApplyNo||'0',9,1)" & _
         ") Y, NATION WHERE NA01(+)=FA10"
      strSql = "SELECT Y.* FROM (" & _
         " select X.*,NVL(FA10,CU10) FA10" & _
         " from (" & strSql & ") X, FAGENT,CUSTOMER" & _
         " WHERE FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75||'0',9,1)" & _
         " AND CU01(+)=SUBSTR(ApplyNo,1,8) AND CU02(+)=SUBSTR(ApplyNo||'0',9,1)" & _
         ") Y, NATION WHERE NA01(+)=FA10"
      'end 2020/5/12
      'Modify by Morgan 2006/2/14 從外層移進來並修改
      If Len(Trim(frm100106_1.txt5(11).Text)) <> 0 Then
           'Modified by Lydia 2023/09/11 區分FMP案管制人---by Phoebe
           'strSql = strSql & " AND NA16 >='" & frm100106_1.txt5(11).Text & "' "
           strSql = strSql & " and (((CP01='FCP' or CP01='FG') and nvl(fa10,'000') >'010' and NA16 >='" & frm100106_1.txt5(11).Text & "') " & _
                   " or ((CP01='P' or CP01='PS') and nvl(fa10,'000') >'010' and NA79 >='" & frm100106_1.txt5(11).Text & "'))"
      End If
      If Len(Trim(frm100106_1.txt5(12).Text)) <> 0 Then
           'Modified by Lydia 2023/09/11 區分FMP案管制人---by Phoebe
           'strSql = strSql & " AND NA16 <='" & frm100106_1.txt5(12).Text & "' "
           strSql = strSql & " and (((CP01='FCP' or CP01='FG') and nvl(fa10,'000') >'010' and NA16 <='" & frm100106_1.txt5(12).Text & "') " & _
                   " or ((CP01='P' or CP01='PS') and nvl(fa10,'000') >'010' and NA79 <='" & frm100106_1.txt5(12).Text & "'))"
      End If
   '2006/2/14 END
   Else
      '2010/9/13 MODIFY BY SONIA 日期欄改百年日期排序問題
      'Modify by Amy 2018/09/20 申請人資料全顯示不限制字數 原:SUBSTRB(CU04,1,10)
      'Modified by Lydia 2019/11/01 +增加欄位SeColPA, ApplyNo
      'Modified by Lydia 2021/05/19 在CP09後面加上CP66
      strSql = "SELECT '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                ", PA26 as ApplyNo " & SeColPA & " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                " WHERE (PA57<>'Y' or pa57 is null) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL1 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
      'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查並顯示申請人資料
      If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
             'Modified by Lydia 2021/05/19 在CP09後面加上CP66
            strSql = strSql & " union all SELECT '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", PA27 as ApplyNo " & SeColPA & " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                      " WHERE (PA57<>'Y' or pa57 is null) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA27,1,8)=CU01(+) AND DECODE(SUBSTR(PA27,9,1),NULL,'0',SUBSTR(PA27,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL10(1) & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all SELECT '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", PA28 as ApplyNo " & SeColPA & " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                      " WHERE (PA57<>'Y' or pa57 is null) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA28,1,8)=CU01(+) AND DECODE(SUBSTR(PA28,9,1),NULL,'0',SUBSTR(PA28,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL10(2) & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all SELECT '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", PA29 as ApplyNo " & SeColPA & " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                      " WHERE (PA57<>'Y' or pa57 is null) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA29,1,8)=CU01(+) AND DECODE(SUBSTR(PA29,9,1),NULL,'0',SUBSTR(PA29,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL10(3) & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all SELECT '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", PA30 as ApplyNo " & SeColPA & " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                      " WHERE (PA57<>'Y' or pa57 is null) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(PA30,1,8)=CU01(+) AND DECODE(SUBSTR(PA30,9,1),NULL,'0',SUBSTR(PA30,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND DECODE(SUBSTr(PA75,9,1),NULL,'0',SUBSTR(PA75,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND PA09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL10(4) & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND PA09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                      " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND PA09 <= '" & frm100106_1.txt5(8).Text & "'", "")
      End If
      'end 2019/11/18
       
       'Modified by Lydia 2019/11/01 +增加欄位SeColTM, ApplyNo
       'Modified by Lydia 2021/05/19 在CP09後面加上CP66
       strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                 ", TM23 as ApplyNo " & SeColTM & " FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                " WHERE (TM29<>'Y' or tm29 is null) AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(TM23,1,8)=CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & strSQL2 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "")
      'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查並顯示申請人資料
      If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            'Modified by Lydia 2021/05/19 在CP09後面加上CP66
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", TM78 as ApplyNo " & SeColTM & " FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                     " WHERE (TM29<>'Y' or tm29 is null) AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(TM78,1,8)=CU01(+) AND DECODE(SUBSTR(TM78,9,1),NULL,'0',SUBSTR(TM78,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql20(1) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", TM79 as ApplyNo " & SeColTM & " FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                     " WHERE (TM29<>'Y' or tm29 is null) AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(TM79,1,8)=CU01(+) AND DECODE(SUBSTR(TM79,9,1),NULL,'0',SUBSTR(TM79,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql20(2) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", TM80 as ApplyNo " & SeColTM & " FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                     " WHERE (TM29<>'Y' or tm29 is null) AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(TM80,1,8)=CU01(+) AND DECODE(SUBSTR(TM80,9,1),NULL,'0',SUBSTR(TM80,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql20(3) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,DECODE(TM10,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(TM10,'000',TM45,CP45) AS 彼所案號,TM12 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ", TM81 as ApplyNo " & SeColTM & " FROM CASEPROGRESS,TRADEMARK,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                     " WHERE (TM29<>'Y' or tm29 is null) AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(TM81,1,8)=CU01(+) AND DECODE(SUBSTR(TM81,9,1),NULL,'0',SUBSTR(TM81,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=FA01(+) AND DECODE(SUBSTR(TM44,9,1),NULL,'0',SUBSTR(TM44,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND TM10=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql20(4) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND TM10 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND TM10 <= '" & frm100106_1.txt5(8).Text & "'", "")
      End If
      'end 2019/11/18
      
       'Modified by Lydia 2019/11/01 +增加欄位SeColLC, ApplyNo
       'Modified by Lydia 2021/05/19 在CP09後面加上CP66
       strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                ",LC11 as ApplyNo " & SeColLC & " FROM CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                " WHERE (LC08<>'Y' or lc08 is null) AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSQL3 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
      'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查並顯示申請人資料
      If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            'Modified by Lydia 2021/05/19 在CP09後面加上CP66
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",LC43 as ApplyNo " & SeColLC & " FROM CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                     " WHERE (LC08<>'Y' or lc08 is null) AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(LC43,1,8)=CU01(+) AND DECODE(SUBSTR(LC43,9,1),NULL,'0',SUBSTR(LC43,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql30(1) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",LC44 as ApplyNo " & SeColLC & " FROM CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                     " WHERE (LC08<>'Y' or lc08 is null) AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(LC44,1,8)=CU01(+) AND DECODE(SUBSTR(LC44,9,1),NULL,'0',SUBSTR(LC44,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql30(2) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",LC45 as ApplyNo " & SeColLC & " FROM CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                     " WHERE (LC08<>'Y' or lc08 is null) AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(LC45,1,8)=CU01(+) AND DECODE(SUBSTR(LC45,9,1),NULL,'0',SUBSTR(LC45,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql30(3) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,DECODE(LC15,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(LC15,'000',LC23,CP45) AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                     ",LC46 as ApplyNo " & SeColLC & " FROM CASEPROGRESS,LAWCASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                     " WHERE (LC08<>'Y' or lc08 is null) AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(LC46,1,8)=CU01(+) AND DECODE(SUBSTR(LC46,9,1),NULL,'0',SUBSTR(LC46,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=FA01(+) AND DECODE(SUBSTR(LC22,9,1),NULL,'0',SUBSTR(LC22,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND LC15=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) " & StrSql30(4) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND LC15 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND LC15 <= '" & frm100106_1.txt5(8).Text & "'", "")
      End If
      'end 2019/11/18
      
       'Modified by Lydia 2019/11/01 +增加欄位SeColHC, ApplyNo
       'Modified by Lydia 2021/05/19 在CP09後面加上CP66
       strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                          AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                 ",HC05 as ApplyNo " & SeColHC & " FROM CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,staff s1,engineerprogress " & _
                " WHERE (HC09<>'Y' or hc09 is null) AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(HC05,1,8)=CU01(+) AND DECODE(SUBSTR(HC05,9,1),NULL,'0',SUBSTR(HC05,9,1))=CU02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSQL4 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
      'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查並顯示申請人資料
      If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
            'Modified by Lydia 2021/05/19 在CP09後面加上CP66
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                          AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                      "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ",HC24 as ApplyNo " & SeColHC & " FROM CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,staff s1,engineerprogress " & _
                     " WHERE (HC09<>'Y' or hc09 is null) AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(HC24,1,8)=CU01(+) AND DECODE(SUBSTR(HC24,9,1),NULL,'0',SUBSTR(HC24,9,1))=CU02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSql40(1) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                          AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ",HC25 as ApplyNo " & SeColHC & " FROM CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,staff s1,engineerprogress " & _
                     " WHERE (HC09<>'Y' or hc09 is null) AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(HC25,1,8)=CU01(+) AND DECODE(SUBSTR(HC25,9,1),NULL,'0',SUBSTR(HC25,9,1))=CU02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSql40(2) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                          AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ",HC26 as ApplyNo " & SeColHC & " FROM CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,staff s1,engineerprogress " & _
                     " WHERE (HC09<>'Y' or hc09 is null) AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(HC26,1,8)=CU01(+) AND DECODE(SUBSTR(HC26,9,1),NULL,'0',SUBSTR(HC26,9,1))=CU02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSql40(3) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
            strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,CPM03                          AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                     "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日,''                                                         AS 代理人,' '                          AS 彼所案號,' ' AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                      ",HC27 as ApplyNo " & SeColHC & " FROM CASEPROGRESS,HIRECASE,NATION N1,NATION N2,STAFF S2,CASEPROPERTYMAP,CUSTOMER,staff s1,engineerprogress " & _
                     " WHERE (HC09<>'Y' or hc09 is null) AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(HC27,1,8)=CU01(+) AND DECODE(SUBSTR(HC27,9,1),NULL,'0',SUBSTR(HC27,9,1))=CU02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND N1.NA01='000' AND CU10=N2.NA01(+) " & StrSql40(4) & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND '000' >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                     " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND '000' <= '" & frm100106_1.txt5(8).Text & "'", "")
      End If
      'end 2019/11/18
      
       'Modified by Lydia 2019/11/01 +增加欄位SeColSP, ApplyNo
       'Modified by Lydia 2021/05/19 在CP09後面加上CP66
       strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                ",SP08 as ApplyNo " & SeColSP & " FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                " WHERE (SP15<>'Y' or sp15 is null) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & strSQL5 & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
       'end 2018/09/20
      'Added by Lydia 2019/11/18 若有下申請人編號,則申請人1~5也要查並顯示申請人資料
      If Len(Trim(frm100106_1.txt4(0))) <> 0 Then
             'Modified by Lydia 2021/05/19 在CP09後面加上CP66
             strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                ",SP58 as ApplyNo " & SeColSP & " FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                " WHERE (SP15<>'Y' or sp15 is null) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(SP58,1,8)=CU01(+) AND DECODE(SUBSTR(SP58,9,1),NULL,'0',SUBSTR(SP58,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & StrSql50(1) & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
             strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                ",SP59 as ApplyNo " & SeColSP & " FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                " WHERE (SP15<>'Y' or sp15 is null) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(SP59,1,8)=CU01(+) AND DECODE(SUBSTR(SP59,9,1),NULL,'0',SUBSTR(SP59,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & StrSql50(2) & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
             strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                ",SP65 as ApplyNo " & SeColSP & " FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                " WHERE (SP15<>'Y' or sp15 is null) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(SP65,1,8)=CU01(+) AND DECODE(SUBSTR(SP65,9,1),NULL,'0',SUBSTR(SP65,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & StrSql50(3) & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
             strSql = strSql & " union all select '' AS V," & strDate1Col & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,DECODE(SP09,'000',CPM03,CPM04) AS 案件性質,nvl(s1.st02,cp14) AS 承辦人,NVL(S2.ST02,cp13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日," & strDate2Col & ",CP64 AS 進度備註, NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) AS 申請人, " & _
                "DECODE(CP22,'Y','是','N','否',NULL,'是','否') AS 是否出名,'' AS 延期日,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍,SUBSTR(' '||sqldatet(CP27),-9) AS 發文日," & StrSQLa & "DECODE(SP09,'000',SP27,CP45) AS 彼所案號,SP11 AS 申請案號,ep12 AS 承辦人備註,SUBSTR(' '||sqldatet(CP57),-9) AS 取消收文日,cp13,cp09, CP66, cp14, CP01, CP10 " & _
                ",SP66 as ApplyNo " & SeColSP & " FROM CASEPROGRESS,SERVICEPRACTICE,NATION N1,NATION N2,NATION N3,STAFF S2,CASEPROPERTYMAP,CUSTOMER,FAGENT,staff s1,engineerprogress,SYSTEMKIND " & _
                " WHERE (SP15<>'Y' or sp15 is null) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND cP09=eP02(+) AND cp14=s1.st01(+) and CP13=S2.ST01(+) and cp14=ep05(+) AND SUBSTR(SP66,1,8)=CU01(+) AND DECODE(SUBSTR(SP66,9,1),NULL,'0',SUBSTR(SP66,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=FA01(+) AND DECODE(SUBSTR(SP26,9,1),NULL,'0',SUBSTR(SP26,9,1))=FA02(+) AND cp01=CPM01(+) AND cp10=CPM02(+) AND SP09=N1.NA01(+) AND CU10=N2.NA01(+) AND CP01=SK01(+) AND FA10=N3.NA01(+) " & StrSql50(4) & _
                " " & IIf(Len(Trim(frm100106_1.txt5(5).Text)) > 0, "AND CP10 >= '" & frm100106_1.txt5(5).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(6).Text)) > 0, "AND CP10 <= '" & frm100106_1.txt5(6).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(7).Text)) > 0, "AND SP09 >= '" & frm100106_1.txt5(7).Text & "'", "") & _
                " " & IIf(Len(Trim(frm100106_1.txt5(8).Text)) > 0, "AND SP09 <= '" & frm100106_1.txt5(8).Text & "'", "")
      End If
      'end 2019/11/18
   End If
     '限制使用者所能使用的系統類別+案件性質
     'Modified by Morgan 2015/10/21 取消系統類別+案件性質的權限控制--秀玲檢查文件沒有該限制
     'strSql = "Select AA.V AS V, " & _
              IIf(frm100106_1.opt1(3).Value, "AA.承辦期限 AS 承辦期限", "AA.本所期限 AS 本所期限") & _
              ", AA.本所案號 AS 本所案號,AA.分所號 as 分所號, AA.案件名稱 AS 案件名稱, AA.案件性質 AS 案件性質, AA.承辦人 AS 承辦人, AA.智權人員 AS 智權人員, AA.收文日 AS 收文日, " & _
              IIf(frm100106_1.opt1(3).Value, "AA.本所期限 AS 本所期限", "AA.法定期限 AS 法定期限") & _
              ", AA.進度備註 AS 進度備註, AA.申請人 AS 申請人, AA.是否出名 AS 是否出名, AA.延期日 AS 延期日, AA.申請國家 AS 申請國家, AA.申請人國籍 AS 申請人國籍, AA.發文日 AS 發文日, AA.代理人 AS 代理人, AA.彼所案號 AS 彼所案號, AA.申請案號 AS 申請案號, AA.承辦人備註 AS 承辦人備註, AA.取消收文日 AS 取消收文日, AA.cp13 AS CP13, AA.cp09 AS CP09, AA.cp14 AS CP14, AA.CP01 AS CP01, AA.CP10 AS CP10 " & _
              " From ( " & strSql & " ) AA, Staff SA, Staff_Group, Staff SB Where AA.CP01=SG02 And AA.CP10=SG03 And SA.ST11=SG01 And AA.CP13=SB.ST01(+) And SA.ST01='" & strUserNum & "' "
     'Modified by Lydia 2019/11/01 利益衝突案件：於CP10後面增加欄位
     'Modified by Lydia 2021/05/19 限定欄位長
     'strSql = "Select AA.V AS V, " & _
              IIf(frm100106_1.OPT1(3).Value, "AA.承辦期限 AS 承辦期限", "AA.本所期限 AS 本所期限") & _
              ", AA.本所案號 AS 本所案號,AA.分所號 as 分所號, AA.案件名稱 AS 案件名稱, AA.案件性質 AS 案件性質, AA.承辦人 AS 承辦人, AA.智權人員 AS 智權人員, AA.收文日 AS 收文日, " & _
              IIf(frm100106_1.OPT1(3).Value, "AA.本所期限 AS 本所期限", "AA.法定期限 AS 法定期限") & _
              ", AA.進度備註 AS 進度備註, AA.申請人 AS 申請人, AA.是否出名 AS 是否出名, AA.延期日 AS 延期日, AA.申請國家 AS 申請國家, AA.申請人國籍 AS 申請人國籍, AA.發文日 AS 發文日, AA.代理人 AS 代理人, AA.彼所案號 AS 彼所案號, AA.申請案號 AS 申請案號, AA.承辦人備註 AS 承辦人備註, AA.取消收文日 AS 取消收文日, AA.cp13 AS CP13, AA.cp09 AS CP09, AA.cp14 AS CP14, AA.CP01 AS CP01, AA.CP10 AS CP10 " & _
              ",ApplyNo, cust01, cust02, cust03, cust04, cust05, fcno " & _
              " From ( " & strSql & " ) AA"
     'end 2015/10/21
     strSql = "Select AA.V AS V, " & _
               IIf(frm100106_1.opt1(3).Value, "substr(AA.承辦期限,1,10) AS 承辦期限", "substr(AA.本所期限,1,10) AS 本所期限") & _
               ",AA.本所案號 AS 本所案號,AA.分所號 as 分所號, AA.案件名稱 AS 案件名稱, AA.案件性質 AS 案件性質, AA.承辦人 AS 承辦人, AA.智權人員 AS 智權人員, substr(AA.收文日,1,10) AS 收文日, " & _
               IIf(frm100106_1.opt1(3).Value, "substr(AA.本所期限,1,10) AS 本所期限", "substr(AA.法定期限,1,10) AS 法定期限") & _
               ", substr(AA.進度備註,1,500) AS 進度備註, AA.申請人 AS 申請人, AA.是否出名 AS 是否出名, substr(AA.延期日,1,10) AS 延期日, AA.申請國家 AS 申請國家, AA.申請人國籍 AS 申請人國籍, substr(AA.發文日,1,10) AS 發文日, " & _
               "AA.代理人 AS 代理人, AA.彼所案號 AS 彼所案號, AA.申請案號 AS 申請案號, substr(AA.承辦人備註,1,500) AS 承辦人備註, substr(AA.取消收文日,1,10) AS 取消收文日, AA.cp13 AS CP13, AA.cp09 AS CP09,  AA.cp66 AS CP66, " & _
               "AA.cp14 AS CP14, AA.CP01 AS CP01, AA.CP10 AS CP10 , ApplyNo, cust01, cust02, cust03, cust04, cust05, fcno " & _
              " From ( " & strSql & " ) AA, Staff SB Where AA.CP13=SB.ST01(+) "
     
     '是否依智權人員排序
     If frm100106_1.txt5(13).Text = "Y" Then
         'edit by nickc 2008/02/21 取消收文日放最下面
         'strSQL = strSQL & " ORDER BY SB.ST03, SB.ST01, 本所期限,本所案號 "
         
         'strSql = strSql & " ORDER BY SB.ST03, SB.ST01, nvl(取消收文日,'11/11/11'),本所期限,本所案號 "
         strSql = strSql & " ORDER BY SB.ST03, SB.ST01, nvl(取消收文日,' 11/11/11')," & IIf(frm100106_1.opt1(3).Value, "承辦期限", "本所期限") & ",本所案號 "
         pub_QL05 = pub_QL05 & ";" & frm100106_1.Label1(9) & frm100106_1.txt5(13)   'Add By Sindy 2010/01/22
     Else
         'edit by nickc 2008/02/21 取消收文日放最下面
         'strSQL = strSQL & " ORDER BY 本所期限,本所案號 "
         
         'strSql = strSql & " ORDER BY nvl(取消收文日,'11/11/11'),本所期限,本所案號 "
         strSql = strSql & " ORDER BY nvl(取消收文日,' 11/11/11')," & IIf(frm100106_1.opt1(3).Value, "承辦期限", "本所期限") & ",本所案號 "
     End If
     
    SetDataListWidth 'Added by Lydia 2021/05/19 清空欄位
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    'Modified by Lydia 2019/11/01 改變型態
    'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
    If adoRecordset.RecordCount <> 0 Then
       dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
         'Added by Lydia 2019/11/01 逐案號判斷
         If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    adoRecordset.Delete
                End If
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
                pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            InsertQueryLog (dblRow) 'Add By Sindy 2010/01/22
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
         Else
            InsertQueryLog (dblRow) 'Add By Sindy 2010/01/22
         End If
        'end 2019/11/01
        
        cmdOK(0).Enabled = True
        cmdOK(1).Enabled = True
        cmdOK(2).Enabled = True
        cmdOK(3).Enabled = True
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/01/22
JumpToNoData:   'Added by Lydia 2019/11/01
        ShowNoData
        Screen.MousePointer = vbDefault
        cmdOK(0).Enabled = False
        cmdOK(1).Enabled = False
        cmdOK(2).Enabled = False
        cmdOK(3).Enabled = False
        StrMenu_sub2 = False
        Exit Function
    End If
    Set grdDataList.Recordset = adoRecordset
    Call SetDataListWidth(False) 'Added by Lydia 2021/05/19 預設欄位
    intK = grdDataList.Rows - 1
    CheckOC
    grdDataList.Visible = False
    For i = 1 To grdDataList.Rows - 1
        'Modified by Lydia 2021/05/19 改用變數取得
        'Me.grdDataList.TextMatrix(i, 5) = Me.grdDataList.TextMatrix(i, 5) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(i, 23), "1")
        Me.grdDataList.TextMatrix(i, colCp10Name) = Me.grdDataList.TextMatrix(i, colCp10Name) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(i, colCp09), "1")
        
        grdDataList.row = i
        'Modified by Lydia 2021/05/19 改用變數取得
        'grdDataList.col = grdDataList.Cols - 4
        grdDataList.col = colCp09
        strSql = "SELECT " & SQLDate("DL02") & " FROM DATELIMIT WHERE DL01='" & grdDataList.Text & "' ORDER BY DL02"
        CheckOC
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        'Modified by Lydia 2021/05/19 改用變數取得: 13=>colDDate3
        grdDataList.col = colDDate3
        If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            adoRecordset.MoveLast
            If Not IsNull(adoRecordset.Fields(0)) Then
                 grdDataList.Text = adoRecordset.Fields(0)
            Else
                 grdDataList.Text = ""
            End If
        End If
          'Modified by Lydia 2021/05/19 改用變數取得
'          If Trim(grdDataList.TextMatrix(grdDataList.row, 21)) <> "" Then
'              'add by nickc 2008/02/21 灰色 +x
'              grdDataList.TextMatrix(grdDataList.row, 2) = "x" & grdDataList.TextMatrix(grdDataList.row, 2)
          If Trim(grdDataList.TextMatrix(grdDataList.row, colCP57)) <> "" Then
              grdDataList.TextMatrix(grdDataList.row, colCaseNo) = "x" & grdDataList.TextMatrix(grdDataList.row, colCaseNo)
          'end 2021/05/19
              For j = 1 To grdDataList.Cols - 1
                 grdDataList.col = j
                 '灰色
                 'Modified by Lydia 2024/10/04 加深灰色&H8000000F>>&HC0C0C0
                 grdDataList.CellBackColor = &HC0C0C0
              Next j
          End If
          'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseno
          grdDataList.col = colCaseNo
          If grdDataList.CellBackColor = -2147483643 Then
              'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseno
              stCaseNo = grdDataList.TextMatrix(grdDataList.row, colCaseNo)
              For j = 1 To 4
                  iPos = InStr(stCaseNo, "-")
                  If iPos > 0 Then
                     stPA(j) = Left(stCaseNo, iPos - 1)
                     stCaseNo = Mid(stCaseNo, iPos + 1)
                  Else
                     stPA(j) = stCaseNo
                  End If
              Next j
              'Modified by Lydia 2021/05/19 改用變數取得
'              If (UCase(stPA(1)) = "FCP" Or UCase(stPA(1)) = "P") And grdDataList.TextMatrix(grdDataList.row, 26) = "605" Then
'                  '2010/9/20 modify by sonia 因日期加空格故加val
'                  strTmp = Val(ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, 9))))
              If (UCase(stPA(1)) = "FCP" Or UCase(stPA(1)) = "P") And grdDataList.TextMatrix(grdDataList.row, colCP10) = "605" Then
                  strTmp = Val(ChangeTStringToWString(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, colDDate2))))
              'end 2021/05/19
                  If strTmp <> "" Then
                      If PUB_IfCtrlDateExtended(stPA, strTmp) = True Then
                          'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseno
                          grdDataList.TextMatrix(grdDataList.row, colCaseNo) = "&" & grdDataList.TextMatrix(grdDataList.row, colCaseNo)
                          For j = 1 To grdDataList.Cols - 1
                             grdDataList.col = j
                             '紫色
                             grdDataList.CellBackColor = &HE600E6
                          Next j
                      End If
                  End If
              End If
          End If
          'add by nickc 2008/02/15 修正 加入黃色
          'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseno
          grdDataList.col = colCaseNo
          If grdDataList.CellBackColor = -2147483643 Then
              'Modified by Lydia 2021/05/19 改用變數取得
'              If Mid(Trim(grdDataList.TextMatrix(grdDataList.row, 23)), 1, 1) = "C" And Trim(grdDataList.TextMatrix(grdDataList.row, 16)) = "" Then
'                  'add by nickc 2008/02/21 黃色 +#
'                  grdDataList.TextMatrix(grdDataList.row, 2) = "#" & grdDataList.TextMatrix(grdDataList.row, 2)
              If Mid(Trim(grdDataList.TextMatrix(grdDataList.row, colCp09)), 1, 1) = "C" And Trim(grdDataList.TextMatrix(grdDataList.row, colCP27)) = "" Then
                  grdDataList.TextMatrix(grdDataList.row, colCaseNo) = "#" & grdDataList.TextMatrix(grdDataList.row, colCaseNo)
              'end 2021/05/19
                  For j = 1 To grdDataList.Cols - 1
                      grdDataList.col = j
                      '黃色
                      grdDataList.CellBackColor = &HFFFF&
                  Next j
              'edit by nickc 2008/02/21 修正字串判斷錯誤
              'ElseIf ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, 1)) < ChangeWStringToTString(ServerDate) And Trim(grdDataList.TextMatrix(grdDataList.row, 1)) <> "" And grdDataList.TextMatrix(grdDataList.row, 16) = "" Then
              'Modified by Lydia 2021/05/19 改用變數取得
'              ElseIf Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, 1))) < Val(ChangeWStringToTString(ServerDate)) And Trim(grdDataList.TextMatrix(grdDataList.row, 1)) <> "" And grdDataList.TextMatrix(grdDataList.row, 16) = "" Then
'                    grdDataList.col = 2
              ElseIf Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, colDDate1))) < Val(ChangeWStringToTString(strSrvDate(1))) And Trim(grdDataList.TextMatrix(grdDataList.row, colDDate1)) <> "" And grdDataList.TextMatrix(grdDataList.row, colCP27) = "" Then
                    grdDataList.col = colCaseNo
              'end 2021/05/19
                    grdDataList.Text = "*" + grdDataList.Text
                    For j = 1 To grdDataList.Cols - 1
                        grdDataList.col = j
                        '紅色
                        grdDataList.CellBackColor = &HFF&
                    Next j
                    GoTo NextItem2
                '2010/9/20 modify by sonia 因日期加空格故加val
                'Modified by Lydia 2021/05/19 改用變數取得
'                ElseIf Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, 1))) = Val(ChangeWStringToTString(ServerDate)) And grdDataList.TextMatrix(grdDataList.row, 16) = "" Then
'                        grdDataList.col = 2
                ElseIf Val(ChangeTDateStringToTString(grdDataList.TextMatrix(grdDataList.row, colDDate1))) = Val(ChangeWStringToTString(strSrvDate(1))) And grdDataList.TextMatrix(grdDataList.row, colCP27) = "" Then
                        grdDataList.col = colCaseNo
                'end 2021/05/19
                        grdDataList.Text = "v" & grdDataList.Text
                        For j = 1 To grdDataList.Cols - 1
                            grdDataList.col = j
                            '橙色
                            grdDataList.CellBackColor = &H80FF&
                        Next j
                        GoTo NextItem2
              Else
                  'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseno
                  grdDataList.col = colCaseNo
                  If grdDataList.CellBackColor = -2147483643 Then
                      'Modified by Lydia 2021/05/19 改用變數取得
'                      If UCase(SystemNumber(grdDataList.TextMatrix(grdDataList.row, 2), 1)) = "FCP" Or UCase(SystemNumber(grdDataList.TextMatrix(grdDataList.row, 2), 1)) = "P" Then
'                          grdDataList.col = 26
                      If UCase(SystemNumber(grdDataList.TextMatrix(grdDataList.row, colCaseNo), 1)) = "FCP" Or UCase(SystemNumber(grdDataList.TextMatrix(grdDataList.row, colCaseNo), 1)) = "P" Then
                          grdDataList.col = colCP10
                      'end 2021/05/19
                          If grdDataList.Text = "605" Then
                              CheckOC
                              'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseno
                              grdDataList.col = colCaseNo
                              strSql = "select pa09,pa14 from patent where pa01='" & SystemNumber(grdDataList.Text, 1) & _
                              "' and pa02='" & SystemNumber(grdDataList.Text, 2) & "' and pa03='" & _
                              SystemNumber(grdDataList.Text, 3) & "' and pa04='" & SystemNumber(grdDataList.Text, 4) & "' and pa09='000' "
                              adoRecordset.CursorLocation = adUseClient
                              adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                              If adoRecordset.RecordCount <> 0 Then
                                  If CheckStr(adoRecordset.Fields(0).Value) = "000" And CheckStr(adoRecordset.Fields(1).Value) = "" Then
                                      'Modified by Lydia 2021/05/19 改用變數取得: 2=>colCaseno
                                      grdDataList.col = colCaseNo
                                      grdDataList.Text = "!" & grdDataList.Text
                                      For j = 1 To grdDataList.Cols - 1
                                          grdDataList.col = j
                                          '綠色
                                          grdDataList.CellBackColor = &HC000&
                                      Next j
                                  End If
                              End If
                              CheckOC
                          End If
                      End If
                  End If
              End If
          End If
NextItem2:
   Next i
   grdDataList.Visible = True
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm100106_3 = Nothing
End Sub

'Add By Sindy 2014/4/2
Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If grdDataList.MouseCol < 0 Or grdDataList.MouseRow < 0 Then Exit Sub
   grdDataList.col = grdDataList.MouseCol
   grdDataList.row = grdDataList.MouseRow
   If Me.grdDataList.row < 1 And Me.grdDataList.Text <> "V" Then
      If Me.grdDataList.Text = "無數字欄位" Then
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
End Sub

Private Sub grdDataList_SelChange()
   grdDataList.Visible = False
   grdDataList.row = grdDataList.MouseRow
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
   If grdDataList.Text = "V" Then
      grdDataList.Text = ""
   
        grdDataList.col = 0
        grdDataList.CellBackColor = QBColor(15)
   
   Else
        grdDataList.Text = "V"
   
        grdDataList.col = 0
        grdDataList.CellBackColor = &HFFC0C0
   
   End If
   End If
   grdDataList.Visible = True
End Sub

Sub StrMenu2()
   Dim StrTemp5 As String
   
   strSql = "SELECT TM44,TM23,TM05,TM06,TM07 FROM TRADEMARK WHERE TM29<>'Y' AND TM01='" & SystemNumber(StrTag, 1) & "' AND TM02='" & SystemNumber(StrTag, 2) & "' AND TM03='" & SystemNumber(StrTag, 3) & "' AND TM04='" & SystemNumber(StrTag, 4) & "' "
   strSql = strSql + "union all select PA75,PA26,PA05,PA06,PA07 FROM PATENT WHERE PA57<>'Y' AND PA01='" & SystemNumber(StrTag, 1) & "' AND PA02='" & SystemNumber(StrTag, 2) & "' AND PA03='" & SystemNumber(StrTag, 3) & "' AND PA04='" & SystemNumber(StrTag, 4) & "' "
   strSql = strSql + "union all select SP26,SP08,SP05,SP06,SP07 FROM SERVICEPRACTICE WHERE SP15<>'Y' AND SP01='" & SystemNumber(StrTag, 1) & "' AND SP02='" & SystemNumber(StrTag, 2) & "' AND SP03='" & SystemNumber(StrTag, 3) & "' AND SP04='" & SystemNumber(StrTag, 4) & "' "
   strSql = strSql + "union all select LC22,LC11,LC05,LC06,LC07 FROM LAWCASE WHERE LC08<>'Y' AND LC01='" & SystemNumber(StrTag, 1) & "' AND LC02='" & SystemNumber(StrTag, 2) & "' AND LC03='" & SystemNumber(StrTag, 3) & "' AND LC04='" & SystemNumber(StrTag, 4) & "' "
   strSql = strSql + "union all select '',HC05,HC06,'','' FROM HIRECASE WHERE HC09<>'Y' AND HC01='" & SystemNumber(StrTag, 1) & "' AND HC02='" & SystemNumber(StrTag, 2) & "' AND HC03='" & SystemNumber(StrTag, 3) & "' AND HC04='" & SystemNumber(StrTag, 4) & "' "
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       If Not IsNull(adoRecordset.Fields(0)) Then
           CheckOC2
           strSql = "SELECT FA10 FROM FAGENT WHERE FA01='" & Left(adoRecordset.Fields(0), 8) & "' AND FA02='" & Right(adoRecordset.Fields(0), 1) & "' "
           adoRecordset1.CursorLocation = adUseClient
           adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
           If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               If Not IsNull(adoRecordset1.Fields(0)) Then
                   StrTemp5 = adoRecordset1.Fields(0)
               Else
                   StrTemp5 = ""
               End If
           End If
           CheckOC2
       Else
           If Not IsNull(adoRecordset.Fields(1)) Then
               CheckOC2
               strSql = "SELECT CU10 FROM FAGENT WHERE CU01='" & Left(adoRecordset.Fields(1), 8) & "' AND CU02='" & Right(adoRecordset.Fields(1), 1) & "' "
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                   If Not IsNull(adoRecordset1.Fields(0)) Then
                       StrTemp5 = adoRecordset1.Fields(0)
                   Else
                       StrTemp5 = ""
                   End If
               End If
               CheckOC2
           Else
               StrTemp5 = ""
           End If
       End If
       If Not IsNull(adoRecordset.Fields(2)) Then
           StrChineseName = adoRecordset.Fields(2)
       Else
           StrChineseName = ""
       End If
       If Not IsNull(adoRecordset.Fields(3)) Then
           StrEnglishName = adoRecordset.Fields(3)
       Else
           StrEnglishName = ""
       End If
       If Not IsNull(adoRecordset.Fields(4)) Then
           StrJanpenName = adoRecordset.Fields(4)
       Else
           StrJanpenName = ""
       End If
   Else
       StrTemp5 = ""
   End If
   CheckOC
   If Len(Trim(StrTemp5)) <> 0 Then
       strSql = "SELECT NA16 FROM NATION WHERE NA01='" & StrTemp5 & "'"
       CheckOC
       adoRecordset.CursorLocation = adUseClient
       adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
           If Not IsNull(adoRecordset.Fields(0)) Then
               StrR03002 = adoRecordset.Fields(0)
               strSql = "SELECT ST02 FROM STAFF WHERE ST01='" & adoRecordset.Fields(0) & "'"
               CheckOC2
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                   StrTemp5 = adoRecordset1.Fields(0)
               Else
                   StrTemp5 = ""
               End If
               CheckOC2
           Else
               StrTemp5 = ""
           End If
       Else
           StrTemp5 = ""
       End If
   End If
   
   'Modified by Lydia 2021/05/19 改用變數取得: 9=>colDDate2
   Call frm100106_4.SetParent(Me, grdDataList.TextMatrix(i, colDDate2)) 'Added by Lydia 2020/03/11 傳入前一畫面和法定期限
   '910430  nick 加傳入mail的 員工代號
   frm100106_4.Hide
   StrTag = StrTemp5 + "-"
   frm100106_4.StrMailNum1 = StrTemp5
   'Modified by Lydia 2021/05/19 改用變數取得: 7=>colSalesName
   grdDataList.col = colSalesName
   StrTag = StrTag + grdDataList.Text + "-"
   'Modify by Morgan 2005/6/2
   'grdDataList.Col = 23
   'Modified by Lydia 2021/05/19 改用變數取得: 22=> colCP13
   grdDataList.col = colCP13
   frm100106_4.StrMailNum2 = grdDataList.Text
   'Modified by Lydia 2021/05/19 改用變數取得: 6=> colCP14name
   grdDataList.col = colCP14name
   StrTag = StrTag + grdDataList.Text + "-"
   'Modify by Morgan 2005/6/2
   'grdDataList.Col = 22
   'Modified by Lydia 2021/05/19 改用變數取得: 24=> colCP14
   grdDataList.col = colCP14
   frm100106_4.StrMailNum3 = grdDataList.Text
End Sub

Sub StrMenu3()
   cnnConnection.Execute "DELETE FROM R100106_T where id='" & strUserNum & "' "
   cnnConnection.Execute "INSERT INTO R100106_T VALUES ('" & ChgSQL(StrR03001) & "','" & ChgSQL(StrR03002) & "','" & ChgSQL(StrR03003) & "','" & ChgSQL(StrR03004) & "','" & ChgSQL(StrR03005) & "','" & ChgSQL(StrR03006) & "','" & ChgSQL(StrR03007) & "','" & ChgSQL(StrR03008) & "','" & ChgSQL(StrR03009) & "','" & ChgSQL(StrR03010) & "','" & ChgSQL(StrR03011) & "','" & strUserNum & "')"
End Sub

'列印期限管制表
Private Sub PrintData()

   Dim ii As Integer
    Page = 1
    With Me.grdDataList
        '若為列印管制表
        If m_blnExportFile = False Then
            PrintTitle
        '若為產生電子檔
        Else
            Print #1, "＜＜期限管制表＞＞"
            'Add By Sindy 2013/8/13
            If frm100106_1.opt1(3).Value = True Then
               Print #1, IIf(frm100106_1.opt1(3).Value, "承辦期限：" & frm100106_1.txt6(0).Text & "－" & frm100106_1.txt6(1).Text, "本所案號：" & frm100106_1.txt3(0).Text & "-" & frm100106_1.txt3(1).Text & "-" & Left(frm100106_1.txt3(2).Text & "0", 1) & "-" & Left(frm100106_1.txt3(3).Text & "00", 2))
            Else
            '2013/8/13 END
               Print #1, IIf(frm100106_1.opt1(0).Value, "本所期限：" & frm100106_1.txt1(0).Text & "－" & frm100106_1.txt1(1).Text, "本所案號：" & frm100106_1.txt3(0).Text & "-" & frm100106_1.txt3(1).Text & "-" & Left(frm100106_1.txt3(2).Text & "0", 1) & "-" & Left(frm100106_1.txt3(3).Text & "00", 2))
            End If
            Print #1, "查詢內容：" & IIf(frm100106_1.opt2(0).Value, frm100106_1.opt2(0).Caption, IIf(frm100106_1.opt2(1).Value, frm100106_1.opt2(1).Caption, frm100106_1.opt2(2).Caption))
            Print #1, String(80, "*")
        End If
        SeekTmp(0) = ""
        SeekTmp(1) = ""
        '記錄智權人員代號
        'edit by  nickc 2007/04/13
        'm_strSalesNo = "" & .TextMatrix(1, 21)
        'Modified by Lydia 2021/05/19 改用變數取得: 22=>colCP13
        m_strSalesNo = "" & .TextMatrix(1, colCP13)
        
        For ii = 1 To .Rows - 1
            For i = 0 To 20
                strTempA(i) = ""
            Next i
            'Modified by Lydia 2021/05/19 改用變數取得: 1=>colDDate1
            strTempA(0) = "" & .TextMatrix(ii, colDDate1) '本所期限/承辦期限
            'Modify by Morgan 2008/11/14
            'strTempA(21) = "" & .TextMatrix(ii, 8) '法定期限
            'Modified by Lydia 2021/05/19 改用變數取得: 9=>colDDate2
            strTempA(21) = "" & .TextMatrix(ii, colDDate2) '法定期限/本所期限
            
            'Modified by Lydia 2021/05/19 改用變數取得
'            strTempA(1) = "" & .TextMatrix(ii, 2)  '本所案號
'            strTempA(2) = StrConv(MidB(StrConv("" & .TextMatrix(ii, 4), vbFromUnicode), 1, 34), vbUnicode) '案件名稱
'            strTempA(3) = "" & .TextMatrix(ii, 8) '收文日
'            strTempA(4) = StrConv(MidB(StrConv("" & .TextMatrix(ii, 14), vbFromUnicode), 1, 8), vbUnicode)  '申請國家
'            strTempA(5) = "" & .TextMatrix(ii, 19) '申請案號
'            strTempA(6) = StrConv(MidB(StrConv("" & .TextMatrix(ii, 5), vbFromUnicode), 1, 10), vbUnicode)  '案件性質
'            strTempA(7) = Left("" & .TextMatrix(ii, 7), 3) '智權人員
'            strTempA(8) = Left("" & .TextMatrix(ii, 6), 3) '承辦人
'            GetOtherData "" & .TextMatrix(ii, 23)
            strTempA(1) = "" & .TextMatrix(ii, colCaseNo)  '本所案號
            strTempA(2) = StrConv(MidB(StrConv("" & .TextMatrix(ii, colCaseName), vbFromUnicode), 1, 34), vbUnicode) '案件名稱
            strTempA(3) = "" & .TextMatrix(ii, colCP05) '收文日
            strTempA(4) = StrConv(MidB(StrConv("" & .TextMatrix(ii, colPA10name), vbFromUnicode), 1, 8), vbUnicode)  '申請國家
            strTempA(5) = "" & .TextMatrix(ii, colPA11) '申請案號
            strTempA(6) = StrConv(MidB(StrConv("" & .TextMatrix(ii, colCp10Name), vbFromUnicode), 1, 10), vbUnicode)  '案件性質
            strTempA(7) = Left("" & .TextMatrix(ii, colSalesName), 3) '智權人員
            strTempA(8) = Left("" & .TextMatrix(ii, colCP14name), 3) '承辦人
            GetOtherData "" & .TextMatrix(ii, colCp09)
            'end 2021/05/19
            
            '若為列印管制表
            If m_blnExportFile = False Then
                If frm100106_1.txt5(13).Text = "Y" Then
                    '若智權人員不同時跳頁
                    'edit by nickc 2007/04/13
                    'If m_strSalesNo <> "" & .TextMatrix(ii, 21) Then
                    '    m_strSalesNo = "" & .TextMatrix(ii, 21)
                    'Modified by Lydia 2021/05/19 改用變數取得: 22=>colCP13
                    If m_strSalesNo <> "" & .TextMatrix(ii, colCP13) Then
                        PrintMemo 'Add by Amy 2016/07/18
                        m_strSalesNo = "" & .TextMatrix(ii, colCP13)
                        iPrint = iPrint + 1
                        Page = Page + 1
                        Printer.NewPage
                        PrintTitle
                    End If
                End If
                'Modified by Lydia 2018/02/14
'                If iPrint > 10600 Then
'                    PrintMemo 'Add by Amy 2016/07/18
'                    iPrint = iPrint + 1
'                    Page = Page + 1
'                    Printer.NewPage
'                    PrintTitle
'                End If
                Call PrintCheck("1")
                'end 2018/02/14
                PrintDatil
            '若為產生電子檔
            Else
                PrintDatilF
            End If
        Next ii
    End With
    '若為列印管制表
    If m_blnExportFile = False Then
        PrintMemo 'Add by Amy 2016/07/18
        Printer.EndDoc
        ShowPrintOk
    End If
End Sub

Private Sub GetPleft()
   Erase PLeft
   '第一列
   PLeft(0) = 0
   PLeft(21) = 1200
   PLeft(1) = 2400
   PLeft(2) = 4400
   PLeft(3) = 8700
   PLeft(4) = 10000
   PLeft(5) = 11000
   PLeft(6) = 13000
   PLeft(7) = 14500
   PLeft(8) = 15500
   '第二列
   PLeft(9) = 0
   PLeft(23) = 11500 'Add by Amy 2014/07/11 +客戶案件案號
   '第三列
   PLeft(10) = 0
   '第四列
   PLeft(22) = 0 + 250
   PLeft(11) = 4000
   PLeft(12) = 8000
   PLeft(13) = 12000
   '第五列
   PLeft(14) = 0 + 250
   '第六列
   PLeft(15) = 0
   '第七列
   PLeft(16) = 0 + 250
   PLeft(17) = 4000
   PLeft(18) = 8000
   PLeft(19) = 12000
   '第八列
   PLeft(20) = 0 + 250
End Sub

Private Sub PrintDatil()
   '第一列
   For i = 0 To 8
       Printer.CurrentX = PLeft(i)
       Printer.CurrentY = iPrint
       Printer.Print strTempA(i)
   Next i
   Printer.CurrentX = PLeft(21)
   Printer.CurrentY = iPrint
   Printer.Print strTempA(21)
   iPrint = iPrint + 300
   'Modified by Lydia 2018/02/14
'   If iPrint > 10600 Then
'       PrintMemo 'Add by Amy 2016/07/18
'       iPrint = iPrint + 1
'       Page = Page + 1
'       Printer.NewPage
'       PrintTitle
'   End If
   Call PrintCheck("1")
   'end 2018/02/14
   '第二列
   For i = 9 To 9
       Printer.CurrentX = PLeft(i)
       Printer.CurrentY = iPrint
       Printer.Print Left(strTempA(i), 80)  'Modify by Amy 2014/07/11 字數太長會與客戶案件案號重疊
   Next i
   'Add by Amy 2014/07/11 +客戶案件案號
   Printer.CurrentX = PLeft(23)
   Printer.CurrentY = iPrint
   Printer.Print strTempA(23)
   'end 2014/07/11
   iPrint = iPrint + 300
   'Modified by Lydia 2018/02/14
'   If iPrint > 10600 Then
'       PrintMemo 'Add by Amy 2016/07/18
'       iPrint = iPrint + 1
'       Page = Page + 1
'       Printer.NewPage
'       PrintTitle
'   End If
   Call PrintCheck("1")
   'end 2018/02/14
   If m_strFACUData = vbYes Then
       '第三列
       For i = 10 To 10
           Printer.CurrentX = PLeft(i)
           Printer.CurrentY = iPrint
           Printer.Print strTempA(i)
       Next i
       iPrint = iPrint + 300
       'Modified by Lydia 2018/02/14
'       If iPrint > 10600 Then
'           PrintMemo 'Add by Amy 2016/07/18
'           iPrint = iPrint + 1
'           Page = Page + 1
'           Printer.NewPage
'           PrintTitle
'       End If
       Call PrintCheck("1")
       'end 2018/02/14
       '第四列
       Printer.CurrentX = PLeft(22)
       Printer.CurrentY = iPrint
       Printer.Print strTempA(22)
       For i = 11 To 13
           Printer.CurrentX = PLeft(i)
           Printer.CurrentY = iPrint
           Printer.Print strTempA(i)
       Next i
       iPrint = iPrint + 300
       'Modified by Lydia 2018/02/14
'       If iPrint > 10600 Then
'           PrintMemo 'Add by Amy 2016/07/18
'           iPrint = iPrint + 1
'           Page = Page + 1
'           Printer.NewPage
'           PrintTitle
'       End If
       Call PrintCheck("1")
       'end 2018/02/14
       '第五列
       For i = 14 To 14
           Printer.CurrentX = PLeft(i)
           Printer.CurrentY = iPrint
           Printer.Print strTempA(i)
       Next i
       iPrint = iPrint + 300
       'Modified by Lydia 2018/02/14
'       If iPrint > 10600 Then
'           PrintMemo 'Add by Amy 2016/07/18
'           iPrint = iPrint + 1
'           Page = Page + 1
'           Printer.NewPage
'           PrintTitle
'       End If
       Call PrintCheck("1")
       'end 2018/02/14
      '若使用者等級非"S"開頭, 才可列印CF代理人資料
       If m_blnSales = False Then
           '第六列
           For i = 15 To 15
               Printer.CurrentX = PLeft(i)
               Printer.CurrentY = iPrint
               Printer.Print strTempA(i)
           Next i
           iPrint = iPrint + 300
           'Modified by Lydia 2018/02/14
'           If iPrint > 10600 Then
'               PrintMemo 'Add by Amy 2016/07/18
'               iPrint = iPrint + 1
'               Page = Page + 1
'               Printer.NewPage
'               PrintTitle
'           End If
           Call PrintCheck("1")
           'end 2018/02/14
           '第七列
           For i = 16 To 19
               Printer.CurrentX = PLeft(i)
               Printer.CurrentY = iPrint
               Printer.Print strTempA(i)
           Next i
           iPrint = iPrint + 300
           'Modified by Lydia 2018/02/14
'           If iPrint > 10600 Then
'               PrintMemo 'Add by Amy 2016/07/18
'               iPrint = iPrint + 1
'               Page = Page + 1
'               Printer.NewPage
'               PrintTitle
'           End If
           Call PrintCheck("1")
           'end 2018/02/14
           '第八列
           For i = 20 To 20
               Printer.CurrentX = PLeft(i)
               Printer.CurrentY = iPrint
               Printer.Print strTempA(i)
           Next i
           iPrint = iPrint + 300
       End If
       'Modified by Lydia 2018/02/14
'       If iPrint > 10600 Then
'           PrintMemo 'Add by Amy 2016/07/18
'           iPrint = iPrint + 1
'           Page = Page + 1
'           Printer.NewPage
'           PrintTitle
'       End If
       Call PrintCheck("1")
       'end 2018/02/14
   End If
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print String(220, "-")
   iPrint = iPrint + 300
   'Modified by Lydia 2018/02/14
'   If iPrint > 10600 Then
'       PrintMemo 'Add by Amy 2016/07/18
'       iPrint = iPrint + 1
'       Page = Page + 1
'       Printer.NewPage
'       PrintTitle
'   End If
   Call PrintCheck("1")
   'end 2018/02/14
End Sub

Private Sub PrintTitle()

   GetPleft
   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6500
   Printer.CurrentY = iPrint
   Printer.Print "期限管制表"
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 6300
   Printer.CurrentY = iPrint
   'Add By Sindy 2013/8/13
   If frm100106_1.opt1(3).Value = True Then
      Printer.Print "承辦期限：" & Format(ChangeTStringToTDateString(frm100106_1.txt6(0).Text) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(frm100106_1.txt6(1).Text)
   '2013/8/13 END
   ElseIf frm100106_1.opt1(0).Value = True Then
      Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(frm100106_1.txt1(0).Text) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(frm100106_1.txt1(1).Text)
   Else
      Printer.Print "本所案號：" & frm100106_1.txt3(0).Text & "-" & frm100106_1.txt3(1).Text & "-" & Left(frm100106_1.txt3(2).Text & "0", 1) & "-" & Left(frm100106_1.txt3(3).Text & "00", 2)
   End If
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print "查詢內容：" & IIf(frm100106_1.opt2(0).Value = True, "未收文", IIf(frm100106_1.opt2(1).Value = True, "已收文未發文", "已收文已發文"))
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print String(220, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   'Add By Sindy 2013/8/13
   If frm100106_1.opt1(3).Value = True Then
      Printer.Print "承辦期限"
   Else
   '2013/8/13 END
      Printer.Print "本所期限"
   End If
   Printer.CurrentX = PLeft(21)
   Printer.CurrentY = iPrint
   'Add By Sindy 2013/8/13
   If frm100106_1.opt1(3).Value = True Then
      Printer.Print "本所期限"
   Else
   '2013/8/13 END
      Printer.Print "法定期限"
   End If
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "案件中文名稱"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "收文日"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "申請國家"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "申請案號"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   iPrint = iPrint + 300
   
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "案件英文名稱"
     'Add by Amy 2014/07/011 +客戶案件案號
   Printer.CurrentX = PLeft(23)
   Printer.CurrentY = iPrint
   Printer.Print "客戶案件案號"
   'end 2014/07/11
   iPrint = iPrint + 300

   If m_strFACUData = vbYes Then
       Printer.CurrentX = PLeft(10)
       Printer.CurrentY = iPrint
       Printer.Print "FC代理人/客戶"
       iPrint = iPrint + 300
       
       Printer.CurrentX = PLeft(22)
       Printer.CurrentY = iPrint
       Printer.Print "FC彼所案號"
       Printer.CurrentX = PLeft(11)
       Printer.CurrentY = iPrint
       Printer.Print "電話"
       Printer.CurrentX = PLeft(12)
       Printer.CurrentY = iPrint
       Printer.Print "傳真"
       Printer.CurrentX = PLeft(13)
       Printer.CurrentY = iPrint
       Printer.Print "E-Mail"
       iPrint = iPrint + 300
       
       Printer.CurrentX = PLeft(14)
       Printer.CurrentY = iPrint
       Printer.Print "地址"
       iPrint = iPrint + 300
       
       '若使用者等級不為"S"開頭者, 才可印CF代理人資料
       If m_blnSales = False Then
           Printer.CurrentX = PLeft(15)
           Printer.CurrentY = iPrint
           Printer.Print "CF代理人"
           iPrint = iPrint + 300
           
           Printer.CurrentX = PLeft(16)
           Printer.CurrentY = iPrint
           Printer.Print "CF彼所案號"
           Printer.CurrentX = PLeft(17)
           Printer.CurrentY = iPrint
           Printer.Print "電話"
           Printer.CurrentX = PLeft(18)
           Printer.CurrentY = iPrint
           Printer.Print "傳真"
           Printer.CurrentX = PLeft(19)
           Printer.CurrentY = iPrint
           Printer.Print "E-Mail"
           iPrint = iPrint + 300
           
           Printer.CurrentX = PLeft(20)
           Printer.CurrentY = iPrint
           Printer.Print "地址"
           iPrint = iPrint + 300
       End If
   End If

   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print String(220, "-")
   iPrint = iPrint + 300
   
End Sub

Private Sub GetOtherData(strCP09 As String)

   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim strFaNo As String '國外代理人
   Dim strCUNo As String '申請人
   Dim strCP44 As String 'CF代理人
   Dim strTempString As String

   strFaNo = ""
   strCUNo = ""
   strCP44 = ""
   'Modify by Amy 2014/07/11 +客戶案件案號
   StrSQLa = "Select PA05, PA06, PA75, PA26, CP44, PA77,PA48 From CaseProgress, Patent Where CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP09='" & strCP09 & "' "
   StrSQLa = StrSQLa & " Union Select TM05, TM06, TM44, TM23, CP44, TM45,TM35 From CaseProgress, TradeMark Where CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And CP09='" & strCP09 & "' "
   StrSQLa = StrSQLa & " Union Select LC05, LC06, LC22, LC11, CP44, LC23,LC17 From CaseProgress, Lawcase Where CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04 And CP09='" & strCP09 & "' "
   StrSQLa = StrSQLa & " Union Select HC06, '', '', HC05, CP44, '','' From CaseProgress, Hirecase Where CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04 And CP09='" & strCP09 & "' "
   StrSQLa = StrSQLa & " Union Select SP05, SP06, SP26, SP08, CP44, SP27,SP29 From CaseProgress, ServicePractice Where CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And CP09='" & strCP09 & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       strTempA(9) = "" & rsA.Fields(1).Value
       strFaNo = "" & rsA.Fields(2).Value
       strCUNo = "" & rsA.Fields(3).Value
       'Add By Cheng 2003/05/23
       '若使用者等級為"S"開頭者, 一律只印申請人資料(不印FC代理人)
       If m_blnSales = True Then
           strFaNo = ""
       End If
       strCP44 = "" & rsA.Fields(4).Value ' For 已收文已發文
       strTempA(22) = "" & rsA.Fields(5).Value 'FC彼所案號
       strTempA(23) = "" & rsA.Fields(6).Value 'Add by Amy 2014/07/11 客戶案件案號
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   '若列印代理人及客戶資料
   If m_strFACUData = vbYes Then
      If strFaNo <> "" Then
          StrSQLa = "Select * From Fagent Where FA01='" & Left(strFaNo, 8) & "' And FA02='" & Right(strFaNo, 1) & "' "
          rsA.CursorLocation = adUseClient
          rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
          If rsA.RecordCount > 0 Then
              strTempA(10) = strFaNo & "  "
              If "" & rsA("FA10").Value < "010" Or "" & rsA("FA10").Value = "020" Or "" & rsA("FA10").Value = "013" Then
                  '中-->英-->日
                  strTempA(10) = strTempA(10) & IIf("" & rsA("FA04").Value <> "", "" & rsA("FA04").Value, IIf("" & rsA("FA05").Value <> "", Trim("" & rsA("FA05").Value & " " & rsA("FA63").Value & " " & rsA("FA64").Value & " " & rsA("FA65").Value), "" & rsA("FA06").Value))
              Else
                  '英-->中-->日
                  strTempA(10) = strTempA(10) & IIf("" & rsA("FA05").Value <> "", Trim("" & rsA("FA05").Value & " " & rsA("FA63").Value & " " & rsA("FA64").Value & " " & rsA("FA65").Value), IIf("" & rsA("FA04").Value <> "", "" & rsA("FA04").Value, "" & rsA("FA06").Value))
              End If
              strTempA(14) = ""
              If "" & rsA("FA10").Value < "010" Or "" & rsA("FA10").Value = "020" Or "" & rsA("FA10").Value = "013" Then
                  '中-->英-->日
                  'Modify by Morgan 2007/1/24 加FA70
                  strTempA(14) = strTempA(14) & IIf("" & rsA("FA17").Value <> "", "" & rsA("FA17").Value, IIf("" & rsA("FA18").Value <> "", Trim("" & rsA("FA18").Value & " " & rsA("FA19").Value & " " & rsA("FA20").Value & " " & rsA("FA21").Value & " " & rsA("FA22").Value & " " & rsA("FA70").Value), "" & rsA("FA23").Value))
              Else
                  '英-->中-->日
                  'Modify by Morgan 2007/1/24 加FA70
                  strTempA(14) = strTempA(14) & IIf("" & rsA("FA18").Value <> "", Trim("" & rsA("FA18").Value & " " & rsA("FA19").Value & " " & rsA("FA20").Value & " " & rsA("FA21").Value & " " & rsA("FA22").Value & " " & rsA("FA70").Value), IIf("" & rsA("FA17").Value <> "", "" & rsA("FA17").Value, "" & rsA("FA23").Value))
              End If
              strTempA(11) = "" & rsA("FA12").Value & IIf("" & rsA("FA13").Value <> "", "," & rsA("FA13").Value, "")
              strTempA(12) = "" & rsA("FA14").Value & IIf("" & rsA("FA15").Value <> "", "," & rsA("FA15").Value, "")
              strTempA(13) = "" & rsA("FA16").Value
          End If
          If rsA.State <> adStateClosed Then rsA.Close
          Set rsA = Nothing
      ElseIf strCUNo <> "" Then
          StrSQLa = "Select * From Customer Where CU01='" & Left(strCUNo, 8) & "' And CU02='" & Right(strCUNo, 1) & "' "
          rsA.CursorLocation = adUseClient
          rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
          If rsA.RecordCount > 0 Then
              strTempA(10) = strCUNo & "  "
              If "" & rsA("CU10").Value < "010" Or "" & rsA("CU10").Value = "020" Or "" & rsA("CU10").Value = "013" Then
                  '中-->英-->日
                  strTempA(10) = strTempA(10) & IIf("" & rsA("CU04").Value <> "", "" & rsA("CU04").Value, IIf("" & rsA("CU05").Value <> "", Trim("" & rsA("CU05").Value & " " & rsA("CU88").Value & " " & rsA("CU89").Value & " " & rsA("CU90").Value), "" & rsA("CU06").Value))
              Else
                  '英-->中-->日
                  strTempA(10) = strTempA(10) & IIf("" & rsA("CU05").Value <> "", Trim("" & rsA("CU05").Value & " " & rsA("CU88").Value & " " & rsA("CU89").Value & " " & rsA("CU90").Value), IIf("" & rsA("CU04").Value <> "", "" & rsA("CU04").Value, "" & rsA("CU06").Value))
              End If
              strTempA(14) = ""
              If "" & rsA("CU10").Value < "010" Or "" & rsA("CU10").Value = "020" Or "" & rsA("CU10").Value = "013" Then
                  '中-->英-->日
                  strTempString = IIf("" & rsA("CU31").Value <> "", "" & rsA("CU31").Value, "" & rsA("CU23").Value)
                  strTempA(14) = strTempA(14) & IIf(strTempString <> "", strTempString, IIf("" & rsA("CU24").Value <> "", Trim("" & rsA("CU24").Value & " " & rsA("CU25").Value & " " & rsA("CU26").Value & " " & rsA("CU27").Value & " " & rsA("CU28").Value), "" & rsA("CU29").Value))
              Else
                  '英-->中-->日
                  strTempString = IIf("" & rsA("CU31").Value <> "", "" & rsA("CU31").Value, "" & rsA("CU23").Value)
                  strTempA(14) = strTempA(14) & IIf("" & rsA("CU24").Value <> "", Trim("" & rsA("CU24").Value & " " & rsA("CU25").Value & " " & rsA("CU26").Value & " " & rsA("CU27").Value & " " & rsA("CU28").Value), IIf(strTempString <> "", strTempString, "" & rsA("CU29").Value))
              End If
              strTempA(11) = "" & rsA("CU16").Value & IIf("" & rsA("CU17").Value <> "", "," & rsA("CU17").Value, "")
              strTempA(12) = "" & rsA("CU18").Value & IIf("" & rsA("CU19").Value <> "", "," & rsA("CU19").Value, "")
              strTempA(13) = "" & rsA("CU20").Value
          End If
          If rsA.State <> adStateClosed Then rsA.Close
          Set rsA = Nothing
      End If
       If m_blnSales = True Then Exit Sub
       '若為列印未收文或已收文未發文, 重新取得CF代理人
       If frm100106_1.opt2(0).Value = True Or frm100106_1.opt2(1).Value = True Then
           StrSQLa = "Select CP44, CP45 From Caseprogress, (Select CP01 A1, CP02 A2, CP03 A3, CP04 A4 From CaseProgress Where CP09='" & strCP09 & "' ) A Where A.A1=CP01 AND A.A2=CP02 AND A.A3=CP03 AND A.A4=CP04 AND CP09 <'C' AND CP27 IS NOT NULL AND CP57 IS NULL ORDER BY CP27 DESC, CP09 DESC "
           rsA.CursorLocation = adUseClient
           rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
           If rsA.RecordCount > 0 Then
               strCP44 = "" & rsA.Fields(0).Value
               strTempA(16) = "" & rsA.Fields(1).Value
           Else
               strCP44 = ""
               strTempA(16) = ""
           End If
           If rsA.State <> adStateClosed Then rsA.Close
           Set rsA = Nothing
       End If
       If strCP44 <> "" Then
           StrSQLa = "Select * From Fagent Where FA01='" & Left(strCP44, 8) & "' And FA02='" & Right(strCP44, 1) & "' "
           rsA.CursorLocation = adUseClient
           rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
           If rsA.RecordCount > 0 Then
               strTempA(15) = strCP44 & "  "
               If "" & rsA("FA10").Value < "010" Or "" & rsA("FA10").Value = "020" Or "" & rsA("FA10").Value = "013" Then
                   '中-->英-->日
                   strTempA(15) = strTempA(15) & IIf("" & rsA("FA04").Value <> "", "" & rsA("FA04").Value, IIf("" & rsA("FA05").Value <> "", Trim("" & rsA("FA05").Value & " " & rsA("FA63").Value & " " & rsA("FA64").Value & " " & rsA("FA65").Value), "" & rsA("FA06").Value))
               Else
                   '英-->中-->日
                   strTempA(15) = strTempA(15) & IIf("" & rsA("FA05").Value <> "", Trim("" & rsA("FA05").Value & " " & rsA("FA63").Value & " " & rsA("FA64").Value & " " & rsA("FA65").Value), IIf("" & rsA("FA04").Value <> "", "" & rsA("FA04").Value, "" & rsA("FA06").Value))
               End If
               strTempA(20) = ""
               If "" & rsA("FA10").Value < "010" Or "" & rsA("FA10").Value = "020" Or "" & rsA("FA10").Value = "013" Then
                   '中-->英-->日
                   'Modify by Morgan 2007/1/24 加 FA70
                   strTempA(20) = strTempA(20) & IIf("" & rsA("FA17").Value <> "", "" & rsA("FA17").Value, IIf("" & rsA("FA18").Value <> "", Trim("" & rsA("FA18").Value & " " & rsA("FA19").Value & " " & rsA("FA20").Value & " " & rsA("FA21").Value & " " & rsA("FA22").Value & " " & rsA("FA70").Value), "" & rsA("FA23").Value))
               Else
                   '英-->中-->日
                   'Modify by Morgan 2007/1/24 加 FA70
                   strTempA(20) = strTempA(20) & IIf("" & rsA("FA18").Value <> "", Trim("" & rsA("FA18").Value & " " & rsA("FA19").Value & " " & rsA("FA20").Value & " " & rsA("FA21").Value & " " & rsA("FA22").Value & " " & rsA("FA70").Value), IIf("" & rsA("FA17").Value <> "", "" & rsA("FA17").Value, "" & rsA("FA23").Value))
               End If
               strTempA(17) = "" & rsA("FA12").Value & IIf("" & rsA("FA13").Value <> "", " ," & rsA("FA13").Value, "")
               strTempA(18) = "" & rsA("FA14").Value & IIf("" & rsA("FA15").Value <> "", " ," & rsA("FA15").Value, "")
               strTempA(19) = "" & rsA("FA16").Value
           End If
           If rsA.State <> adStateClosed Then rsA.Close
           Set rsA = Nothing
       End If
   End If

End Sub

Private Sub PrintDatilF()
   '第一列
   'Add By Sindy 2013/8/13
   If frm100106_1.opt1(3).Value = True Then
      Print #1, StrConv(LeftB(StrConv("承辦期限：" & strTempA(0) & String(40, " "), vbFromUnicode), 40), vbUnicode) & "本所期限：" & strTempA(21)
   Else
   '2013/8/13 END
      Print #1, StrConv(LeftB(StrConv("本所期限：" & strTempA(0) & String(40, " "), vbFromUnicode), 40), vbUnicode) & "法定期限：" & strTempA(21)
   End If
   Print #1, "本所案號：" & strTempA(1)
   Print #1, "案件中文名稱：" & strTempA(2)
   Print #1, StrConv(LeftB(StrConv("收文日：" & strTempA(3) & String(40, " "), vbFromUnicode), 40), vbUnicode) & "申請國家：" & strTempA(4)
   Print #1, StrConv(LeftB(StrConv("申請案號：" & strTempA(5) & String(40, " "), vbFromUnicode), 40), vbUnicode) & "案件性質：" & strTempA(6)
   Print #1, StrConv(LeftB(StrConv("智權人員：" & strTempA(7) & String(40, " "), vbFromUnicode), 40), vbUnicode) & "承辦人：" & strTempA(8)
   '第二列
   Print #1, "案件英文名稱：" & strTempA(9)
   '第三列
   Print #1, "FC代理人/客戶：" & strTempA(10)
   '第四列
   Print #1, "FC彼所案號：" & strTempA(22)
   Print #1, StrConv(LeftB(StrConv("電話：" & strTempA(11) & String(40, " "), vbFromUnicode), 40), vbUnicode) & "傳真：" & strTempA(12)
   Print #1, "E-Mail：" & strTempA(13)
   '第五列
   Print #1, "地址：" & strTempA(14)
   '若使用者等級為"S"開頭者, 才可列印CF代理人資料
   If m_blnSales = False Then
       '第六列
       Print #1, "CF代理人：" & strTempA(15)
       '第七列
       Print #1, "CF彼所案號：" & strTempA(16)
       Print #1, StrConv(LeftB(StrConv("電話：" & strTempA(17) & String(40, " "), vbFromUnicode), 40), vbUnicode) & "傳真：" & strTempA(18)
       Print #1, "E-Mail：" & strTempA(19)
       '第八列
       Print #1, "地址：" & strTempA(20)
   End If
   Print #1, String(80, "*")
End Sub

'取得使用者等級
'Modify by Amy 2016/07/18 改為Public
Public Function GetST05(ByVal strUserNum As String) As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
       
   GetST05 = ""
   StrSQLa = "Select ST05 From Staff Where ST01='" & strUserNum & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       GetST05 = "" & rsA.Fields(0).Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Add by Amy 2016/07/18
'按「個人客戶二個月之內期限未收文案件」鈕,列印期限管制表
Public Sub PrintData2()
    Dim ii As Integer
    Dim bolIsFirst As Boolean
   
    Page = 1: bolIsFirst = True: strOldApply = ""
    With Me.grdDataList
        PrintTitle2 (bolIsFirst)
        
        For ii = 1 To .Rows - 1
            For i = 0 To 20
                strTempA(i) = ""
            Next i
            'Modified by Lydia 2021/05/19 改用變數取得
'            strTempA(0) = "" & .TextMatrix(ii, 1) '本所期限/承辦期限
'            strTempA(21) = "" & .TextMatrix(ii, 9) '法定期限/本所期限
'            strTempA(1) = "" & .TextMatrix(ii, 2) '本所案號
'            strTempA(2) = StrConv(MidB(StrConv("" & .TextMatrix(ii, 4), vbFromUnicode), 1, 34), vbUnicode) '案件名稱
'            strTempA(3) = "" & .TextMatrix(ii, 8) '收文日
'            strTempA(4) = StrConv(MidB(StrConv("" & .TextMatrix(ii, 14), vbFromUnicode), 1, 8), vbUnicode)  '申請國家
'            strTempA(5) = "" & .TextMatrix(ii, 19) '申請案號
'            strTempA(6) = StrConv(MidB(StrConv("" & .TextMatrix(ii, 5), vbFromUnicode), 1, 10), vbUnicode)  '案件性質
'            strTempA(7) = Left("" & .TextMatrix(ii, 7), 3) '智權人員
'            strTempA(8) = Left("" & .TextMatrix(ii, 6), 3) '承辦人
'            GetOtherData "" & .TextMatrix(ii, 23)
'            strTempA(10) = StrConv(MidB(StrConv("" & .TextMatrix(ii, 10), vbFromUnicode), 1, 50), vbUnicode) '下一程序備註
            strTempA(0) = "" & .TextMatrix(ii, colDDate1) '本所期限/承辦期限
            strTempA(21) = "" & .TextMatrix(ii, colDDate2) '法定期限/本所期限
            strTempA(1) = "" & .TextMatrix(ii, colCaseNo) '本所案號
            strTempA(2) = StrConv(MidB(StrConv("" & .TextMatrix(ii, colCaseName), vbFromUnicode), 1, 34), vbUnicode) '案件名稱
            strTempA(3) = "" & .TextMatrix(ii, colCP05) '收文日
            strTempA(4) = StrConv(MidB(StrConv("" & .TextMatrix(ii, colPA10name), vbFromUnicode), 1, 8), vbUnicode)  '申請國家
            strTempA(5) = "" & .TextMatrix(ii, colPA11) '申請案號
            strTempA(6) = StrConv(MidB(StrConv("" & .TextMatrix(ii, colCp10Name), vbFromUnicode), 1, 10), vbUnicode)  '案件性質
            strTempA(7) = Left("" & .TextMatrix(ii, colSalesName), 3) '智權人員
            strTempA(8) = Left("" & .TextMatrix(ii, colCP14name), 3) '承辦人
            GetOtherData "" & .TextMatrix(ii, colCp09)
            strTempA(10) = StrConv(MidB(StrConv(PUB_MGridGetValue(ii, IIf(frm100106_1.opt2(0).Value = True, "備註", "進度備註"), grdDataList), vbFromUnicode), 1, 50), vbUnicode) '下一程序備註
            'end 2021/05/19
            
            '若智權人員不同時跳頁
            'Modified by Lydia 2021/05/19 改用變數取得
'            If strOldApply <> "" & .TextMatrix(ii, 28) Then
'                If bolIsFirst = True Then strOldApply = "" & .TextMatrix(ii, 28)
            If strOldApply <> "" & PUB_MGridGetValue(ii, "ApplyNo", grdDataList) Then
                If bolIsFirst = True Then strOldApply = "" & PUB_MGridGetValue(ii, "ApplyNo", grdDataList)
            'end 2021/05/19
                PrintMemo (True)
                If bolIsFirst = True Then
                    bolIsFirst = False
                Else
                    iPrint = iPrint + 1
                    Page = Page + 1
                    Printer.NewPage
                End If
                PrintTitle2
            End If
            'Modified by Lydia 2018/02/14
'            If iPrint > 10600 Then
'                PrintMemo (True)
'                iPrint = iPrint + 1
'                Page = Page + 1
'                Printer.NewPage
'                PrintTitle2
'            End If
            Call PrintCheck("2")
            'end 2018/02/14
            PrintDatil2
            '記錄申請人編號
            'Modified by Lydia 2021/05/19 改用變數取得
            'strOldApply = "" & .TextMatrix(1, 28)
            strOldApply = "" & "" & PUB_MGridGetValue(ii, "ApplyNo", grdDataList)
        Next ii
    End With
    PrintMemo (True)
    Printer.EndDoc
End Sub

Private Sub PrintMemo(Optional ByVal bolCusCase As Boolean = False)
    Dim i As Integer, iPos As Integer, iEnd As Integer
    Dim intCuX As Integer, intCuY As Integer
    Dim strMemo As String
    Dim strTmp(2) As String
    
    'Modified by Lydia 2018/02/14 列印管制表的頁面最下方備註會跨頁列印。
    'intCuX = 0: intCuY = 11000
    intCuX = 0: intCuY = m_Bottom + 400
    Printer.Font.Size = 12
    
    For i = 0 To Combo1.ListCount - 1
        If Left(Combo1.List(i), 2) <> "藍色" Then
            strTmp(1) = ""
            If bolCusCase = True Then
                If Left(Combo1.List(i), 8) <> "本所案號後有全形" And Left(Combo1.List(i), 2) <> "黃色" And Left(Combo1.List(i), 2) <> "灰色" Then
                    iPos = InStr(Combo1.List(i), "(") + 1
                    iEnd = InStr(Combo1.List(i), ")")
                    strTmp(1) = Replace(Mid(Combo1.List(i), iPos, iEnd - iPos), ")", "") & Mid(Combo1.List(i), iEnd + 1)
                End If
            Else
                If Left(Combo1.List(i), 8) = "本所案號後有全形" Then
                    strTmp(1) = Combo1.List(i)
                Else
                    iPos = InStr(Combo1.List(i), "(") + 1
                    iEnd = InStr(Combo1.List(i), ")")
                    strTmp(1) = Replace(Mid(Combo1.List(i), iPos, iEnd - iPos), ")", "") & Mid(Combo1.List(i), iEnd + 1)
                End If
            End If
            If GetTextLength(strTmp(0) & IIf(strTmp(1) = "", "", "　" & strTmp(1))) > 140 Then
                Printer.CurrentX = intCuX
                Printer.CurrentY = intCuY
                Printer.Print strTmp(0)
                strTmp(0) = "　" & strTmp(1)
                intCuX = 0
                intCuY = intCuY + 250
            ElseIf i = Combo1.ListCount - 1 Then
                Printer.CurrentX = intCuX
                Printer.CurrentY = intCuY
                Printer.Print strTmp(0) & IIf(strTmp(1) = "", "", "　" & strTmp(1))
            Else
                strTmp(0) = strTmp(0) & IIf(strTmp(1) = "", "", "　" & strTmp(1))
            End If
        End If
    Next i
End Sub

'取得申請人名稱
Private Function GetCustomerN(ByVal stCU01 As String, ByVal stCU02 As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim stSQL As String
    Dim intQ As Integer
    
    GetCustomerN = ""
    stSQL = "Select Nvl(CU04,Nvl(CU05||CU88||CU89||CU90,CU06)) as CusName " & _
                "From Customer Where CU01='" & stCU01 & "' And CU02='" & stCU02 & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
    If intQ = 1 Then
        GetCustomerN = "" & RsQ.Fields("CusName")
    End If
    RsQ.Close
End Function

Private Sub PrintTitle2(Optional ByVal bolFirst As Boolean = False)
   GetPleft2
   iPrint = 500
   If bolFirst = True Then Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6500
   Printer.CurrentY = iPrint
   Printer.Print "期限管制表"
   
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 6300
   Printer.CurrentY = iPrint
   If frm100106_1.opt1(3).Value = True Then
      Printer.Print "承辦期限：" & Format(ChangeTStringToTDateString(frm100106_1.txt6(0).Text) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(frm100106_1.txt6(1).Text)
   ElseIf frm100106_1.opt1(0).Value = True Then
      Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(frm100106_1.txt1(0).Text) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(frm100106_1.txt1(1).Text)
   Else
      Printer.Print "本所案號：" & frm100106_1.txt3(0).Text & "-" & frm100106_1.txt3(1).Text & "-" & Left(frm100106_1.txt3(2).Text & "0", 1) & "-" & Left(frm100106_1.txt3(3).Text & "00", 2)
   End If
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   
   iPrint = iPrint + 300
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print "查詢內容：未收文"
   Printer.CurrentX = 2400
   Printer.CurrentY = iPrint
   Printer.Print "申請人：" & PUB_StrToStr_byVal(GetCustomerN(Mid(strOldApply, 1, 8), Mid(strOldApply, 9, 1)), 100)
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(Page)
   
   iPrint = iPrint + 300
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print String(220, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   If frm100106_1.opt1(3).Value = True Then
      Printer.Print "承辦期限"
   Else
      Printer.Print "本所期限"
   End If
   Printer.CurrentX = PLeft(21)
   Printer.CurrentY = iPrint
   If frm100106_1.opt1(3).Value = True Then
      Printer.Print "本所期限"
   Else
      Printer.Print "法定期限"
   End If
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "案件中文名稱"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "申請國家"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "申請案號"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   iPrint = iPrint + 300
   
   Printer.CurrentX = PLeft(22)
   Printer.CurrentY = iPrint
   Printer.Print "客戶案件案號"
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iPrint
   Printer.Print "備註"
   iPrint = iPrint + 300
   
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print String(220, "-")
   iPrint = iPrint + 300
End Sub

Private Sub GetPleft2()
   Erase PLeft
   '第一列
   PLeft(0) = 0
   PLeft(21) = 1200
   PLeft(1) = 2400
   PLeft(2) = 4400
   PLeft(4) = 8700
   PLeft(5) = 10000
   PLeft(6) = 12000
   PLeft(7) = 13500
   PLeft(8) = 15000
   PLeft(9) = 16500
   '第二列
   PLeft(23) = 0 '客戶案件案號
   PLeft(10) = 4400
End Sub

Private Sub PrintDatil2()
   '第一列
   For i = 0 To 2
       Printer.CurrentX = PLeft(i)
       Printer.CurrentY = iPrint
       Printer.Print strTempA(i)
   Next i
   For i = 4 To 8
       Printer.CurrentX = PLeft(i)
       Printer.CurrentY = iPrint
       Printer.Print strTempA(i)
   Next i
   Printer.CurrentX = PLeft(21)
   Printer.CurrentY = iPrint
   Printer.Print strTempA(21)
   iPrint = iPrint + 300
   'Modified by Lydia 2018/02/14
'   If iPrint > 10600 Then
'       PrintMemo (True)
'       iPrint = iPrint + 1
'       Page = Page + 1
'       Printer.NewPage
'       PrintTitle2
'   End If
   Call PrintCheck("2")
   'end 2018/02/14
   '第二列
   Printer.CurrentX = PLeft(23)
   Printer.CurrentY = iPrint
   Printer.Print strTempA(23)
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iPrint
   Printer.Print strTempA(10)
   iPrint = iPrint + 300
   'Modified by Lydia 2018/02/14
'   If iPrint > 10600 Then
'       PrintMemo (True)
'       iPrint = iPrint + 1
'       Page = Page + 1
'       Printer.NewPage
'       PrintTitle2
'   End If
   Call PrintCheck("2")
   'end 2018/02/14
 
   Printer.CurrentX = 0
   Printer.CurrentY = iPrint
   Printer.Print String(220, "-")
   iPrint = iPrint + 300
   'Modified by Lydia 2018/02/14
'   If iPrint > 10600 Then
'       PrintMemo (True)
'       iPrint = iPrint + 1
'       Page = Page + 1
'       Printer.NewPage
'       PrintTitle2
'   End If
   Call PrintCheck("2")
   'end 2018/02/14
End Sub
'end 2016/07/18

'Added by Lydia 2018/02/14 檢查超過最大Y位置就換頁
Private Sub PrintCheck(ByVal iKind As String)
    If iPrint > m_Bottom Then
        If iKind = "1" Then
            PrintMemo
            iPrint = iPrint + 1
            Page = Page + 1
            Printer.NewPage
            PrintTitle
        Else
            PrintMemo (True)
            iPrint = iPrint + 1
            Page = Page + 1
            Printer.NewPage
            PrintTitle2
        End If
    End If
End Sub

'Added by Lydia 2021/05/28 Grid的勾選同步更新
Public Sub UpdateShowFlag(ByVal pRow As Integer)
'因為可選多筆連續輸入管制備註，存檔後自動帶下一筆；取消則回前畫面，但有勾選但未顯示的資料的勾選符號必須保留，這樣才知道做到哪一筆。
        
    grdDataList.row = pRow
    grdDataList.col = 0
    grdDataList.Text = "" '取消勾選V
    grdDataList.CellBackColor = QBColor(15)

End Sub


