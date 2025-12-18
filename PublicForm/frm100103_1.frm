VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100103_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "以案件名稱查詢"
   ClientHeight    =   5724
   ClientLeft      =   -480
   ClientTop       =   1752
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5724
   ScaleWidth      =   9312
   Begin VB.TextBox txtRows 
      Alignment       =   2  '置中對齊
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   228
      Left            =   8520
      TabIndex        =   19
      Top             =   1248
      Width           =   324
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "已設定代表圖(&I)"
      Height          =   350
      Index           =   6
      Left            =   6990
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   30
      Width           =   1425
   End
   Begin VB.CheckBox chk 
      Caption         =   "所有系統類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5970
      TabIndex        =   17
      Top             =   1170
      Width           =   1740
   End
   Begin VB.TextBox TXT3 
      Height          =   300
      Left            =   3285
      MaxLength       =   1
      TabIndex        =   3
      Top             =   420
      Width           =   315
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   4140
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   7303
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
      _Band(0).Cols   =   8
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   1116
      TabIndex        =   5
      Top             =   1140
      Width           =   4824
   End
   Begin VB.OptionButton Option1 
      Caption         =   "日文"
      Height          =   180
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   468
      Width           =   732
   End
   Begin VB.OptionButton Option1 
      Caption         =   "英文"
      Height          =   180
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   468
      Width           =   732
   End
   Begin VB.OptionButton Option1 
      Caption         =   "中文"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   468
      Value           =   -1  'True
      Width           =   732
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3450
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   30
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度(&C)"
      Height          =   350
      Index           =   3
      Left            =   5760
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   30
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   350
      Index           =   2
      Left            =   4230
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   30
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   8448
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   30
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "共　　　筆"
      Height          =   180
      Left            =   8232
      TabIndex        =   18
      Top             =   1272
      Width           =   900
   End
   Begin MSForms.TextBox Text2 
      Height          =   330
      Left            =   1116
      TabIndex        =   4
      Top             =   750
      Width           =   4815
      VariousPropertyBits=   671107099
      Size            =   "8493;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "●代表銷卷＊代表閉卷"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6030
      TabIndex        =   16
      Top             =   825
      Width           =   1830
   End
   Begin VB.Label Label4 
      Caption         =   "商標及查名案件請點選 中文 查詢, 輸入條件不受語文限制 !!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   456
      Left            =   108
      TabIndex        =   15
      Top             =   12
      Width           =   2988
   End
   Begin VB.Label Label1 
      Caption         =   "是否同時查詢對造案件(Y：是)"
      Height          =   180
      Left            =   3732
      TabIndex        =   14
      Top             =   468
      Width           =   2472
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   $"frm100103_1.frx":0000
      Height          =   180
      Left            =   150
      TabIndex        =   13
      Top             =   1200
      Width           =   6885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   150
      TabIndex        =   12
      Top             =   825
      Width           =   900
   End
End
Attribute VB_Name = "frm100103_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/29 改成Form2.0 ; GrdDataList改字型=新細明體-ExtB、Text2
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/10 日期欄已修改
Option Explicit

Dim strSql As String, StrTag As String, intK As Integer
Dim s As Integer, i As Integer, j As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String 'Add By Sindy 2016/2/22
'Added by Lydia 2019/11/01 利益衝突案件
Dim intCufaCnt As Integer '限閱案件X件


Private Sub SetDataListWidth()
'Modified by Lydia 2019/11/01
'GrdDataList.Cols = 9
Dim intField As Integer
intField = 15
grdDataList.Cols = intField
'end 2019/11/01
   
   grdDataList.row = 0
   grdDataList.col = 0: grdDataList.Text = "V"
   grdDataList.ColWidth(0) = 200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 1: grdDataList.Text = "本所案號"
   grdDataList.ColWidth(1) = 1550
   grdDataList.CellAlignment = flexAlignCenterCenter
   Dim iDep As String
   iDep = PUB_GetST06(strUserNum)
   grdDataList.col = 2: grdDataList.Text = "分所號"
   '電腦中心，跟分所才秀
   If PUB_GetST03(strUserNum) <> "M51" And iDep = "1" Then
       grdDataList.ColWidth(2) = 0
   Else
       grdDataList.ColWidth(2) = 620
   End If
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 3: grdDataList.Text = "案件名稱"
   grdDataList.ColWidth(3) = 1600
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 4: grdDataList.Text = "申請國家"
   grdDataList.ColWidth(4) = 1000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 5: grdDataList.Text = "商品類別"
   grdDataList.ColWidth(5) = 1200
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 6: grdDataList.Text = "申請人"
   grdDataList.ColWidth(6) = 1000
   grdDataList.CellAlignment = flexAlignCenterCenter
   grdDataList.col = 7: grdDataList.Text = "相關人"
   grdDataList.ColWidth(7) = 1000
   grdDataList.CellAlignment = flexAlignCenterCenter
   'add by nickc 2005/05/10
   grdDataList.col = 8: grdDataList.Text = ""
   grdDataList.ColWidth(8) = 0
   grdDataList.CellAlignment = flexAlignCenterCenter
    'Added by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
    For intI = 9 To intField - 1
         grdDataList.col = intI
         grdDataList.ColWidth(intI) = 0
    Next intI
    'end 2019/11/01
    
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
   
   Select Case cmdState
      Case 2
            Me.Enabled = False
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
                 Dim Str01 As String
                 grdDataList.col = 1
                 Str01 = SystemNumber(grdDataList, 1)
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
                     Case "CFL", "FCL", "L", "LIN", "ACS" '法務
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
                     Me.Enabled = True
                     Exit Sub
                 End If
               End If
            Next i
            Me.Enabled = True
      Case 3
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
      'Add By Sindy 2016/2/22
      Case 6 '已設定代表圖
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
                     strCP01 = SystemNumber(Pub_RplStr(grdDataList.Text), 1)
                     strCP02 = SystemNumber(Pub_RplStr(grdDataList.Text), 2)
                     strCP03 = SystemNumber(Pub_RplStr(grdDataList.Text), 3)
                     strCP04 = SystemNumber(Pub_RplStr(grdDataList.Text), 4)
                     frmPic001.oCP01 = strCP01
                     frmPic001.oCP02 = strCP02
                     frmPic001.oCP03 = strCP03
                     frmPic001.oCP04 = strCP04
                     frmPic001.StrMenu
                     frmPic001.CanScan
                     frmPic001.SetSeekCmdok 'Add by Amy 2018/07/18
                     frmPic001.Show vbModal
                     Call SetCmdImg(strCP01, strCP02, strCP03, strCP04) '檢查有無代表圖
                     Exit Sub
                  End If
               End If
            Next i
      '2016/2/22 END
      Case 1
            fnCloseAllFrm100
      Case Else
   End Select
End Sub

'Add By Sindy 2016/2/22 檢查有無代表圖
Private Sub SetCmdImg(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String)
    'Modify by Amy 2018/07/16  改寫至function
'   strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & strCP01 & "' and ibf02='" & strCP02 & "' and ibf03='" & strCP03 & "' and ibf04='" & strCP04 & "' and ibf05='1'"
'   CheckOC2
'   adoRecordset1.CursorLocation = adUseClient
'   adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
   If ChkImgByteFile(strCP01, strCP02, strCP03, strCP04) = True Then
      cmdOK(6).Caption = "已設定代表圖(&I)"
      cmdOK(6).BackColor = &HC0FFC0
   Else
      cmdOK(6).Caption = "未設定代表圖(&I)"
      cmdOK(6).BackColor = &HC0C0FF
   End If
'   CheckOC2
   'end2018/07/18
End Sub

'add by nickc 2007/05/04 加讓他勾的
Private Sub chk_Click()
   '若勾選所有系統類別
   If Me.chk.Value = vbChecked Then
       Me.Text5.Text = "ALL"
   '若取消勾選所有系統類別
   Else
       Me.Text5.Text = Systemkind_g
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   'add by nickc 2007/01/12
   If Len(Trim(Me.Text5.Text)) = 0 Then
       Me.Text5.Text = "ALL"
   End If
   '92.04.16 nick 紀錄作用按鍵
   cmdState = Index
   PubShowNextData
   Exit Sub
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   cmdOK(2).Enabled = False
   cmdOK(3).Enabled = False
   cmdOK(6).Enabled = False 'Add By Sindy 2016/2/22
   '2011/12/6 modify by sonia
   'Text5 = Systemkind_g
   Me.chk.Value = vbChecked
   Text5 = "ALL"
   '2011/12/6 end
   bolToEndByNick = False
   '92.04.16 nick
   cmdState = -1
End Sub

Private Sub cmdSearch_Click()
'Add By Cheng 2002/01/07
'宣告變數
Dim strText5 As String
   
   txtRows = "" 'Added by Morgan 2025/10/2
   
   'Modify By Cheng 2002/03/14
   ''Add By Cheng 2002/01/07
   'Text5_LostFocus
   'add by nickc 2007/01/12
   If Len(Trim(Me.Text5.Text)) = 0 Then
       Me.Text5.Text = "ALL"
   End If
   If (Len(Trim(Text2))) = 0 Then
       s = MsgBox("輸入條件不可空白", , "User 輸入錯誤")
       Text2.SetFocus
       Text2.SelStart = 0
       Text2.SelLength = Len(Text2)
       Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   GetData100103_1
   Screen.MousePointer = vbDefault
End Sub

Sub GetData100103_1()
'Add By Cheng 2002/03/14
Dim strTemp As String
'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String
Dim dblRow As Double 'Add By Sindy 2025/9/3
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/27 清除查詢印表記錄檔欄位
   Me.Enabled = False
   grdDataList.Clear
   grdDataList.Rows = 2
   SetDataListWidth
   'Add By Cheng 2002/03/14
   strTemp = IIf(Me.Text5.Text <> "ALL", Me.Text5.Text, GetAllSysKind(Me.Text5))
    'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
    SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
    SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
    SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
    SeColLC = " ,lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno "
    SeColHC = " ,hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno "
    intCufaCnt = 0
    'end 2019/11/01
        
   If Option1(0).Value = True Then              '中文
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption 'Add By Sindy 2010/10/27
      If Text5 <> "" Then
         pub_QL05 = pub_QL05 & ";" & Label3 & Text5 'Add By Sindy 2010/10/27
      End If
   '查詢商標基本檔
       'edit by nickc 2005/05/10
       'strSQL = "select ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,TM05 AS 案件名稱,nation.NA03 AS 申請國家,TM09 AS 商品類別,nvl(cu04,tm23) AS 申請人, '' AS 相關人 from trademark,nation,customer WHERE instr(Upper(TM05),'" & ChgSQL(UCase(Text2)) & "')>0 and tm01 in (" & SQLGrpStr(strTemp, 2) & ") and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) and tm10=na01(+) "
       'Modified by Lydia 2019/11/01 +增加欄位SeColTM
       strSql = "select ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,TM05 AS 案件名稱,nation.NA03 AS 申請國家,TM09 AS 商品類別,nvl(cu04,tm23) AS 申請人, '' AS 相關人,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as FSort" & SeColTM & _
                   " from trademark,nation,customer WHERE instr(Upper(TM05),'" & ChgSQL(UCase(Text2)) & "')>0 and tm01 in (" & SQLGrpStr(strTemp, 2) & ") and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) and tm10=na01(+) "
   '查詢專利基本檔
       'edit by nickc 2005/05/10
       'strSQL = strSQL + " union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,PA05 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu04,pa26) AS 申請人,'' AS 相關人 from patent,nation,customer WHERE instr(Upper(PA05),'" & ChgSQL(UCase(Text2)) & "')>0 and pa01 in (" & SQLGrpStr(strTemp, 1) & ") and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) and pa09=na01(+) "
       'Modified by Lydia 2019/11/01 +增加欄位SeColPA
       strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,PA05 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu04,pa26) AS 申請人,'' AS 相關人,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort" & SeColPA & _
                 " from patent,nation,customer WHERE instr(Upper(PA05),'" & ChgSQL(UCase(Text2)) & "')>0 and pa01 in (" & SQLGrpStr(strTemp, 1) & ") and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) and pa09=na01(+) "
   '查詢服務業務
       'edit by nickc 2005/05/10
       'strSQL = strSQL + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,SP05 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu04,sp08) AS 申請人,'' AS 相關人 from servicepractice,nation,customer WHERE instr(Upper(SP05),'" & ChgSQL(UCase(Text2)) & "')>0 and sp01 in (" & SQLGrpStr(strTemp, 5) & ") and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) and sp09=na01(+) "
       'Modified by Lydia 2019/11/01 +增加欄位SeColSP
       'Modify by Amy 2020/02/05 +SP73 商品類別
       strSql = strSql + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,SP05 AS 案件名稱,nation.NA03 AS 申請國家,NVL(SP73,'') AS 商品類別,nvl(cu04,sp08) AS 申請人,'' AS 相關人,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort" & SeColSP & _
                  " from servicepractice,nation,customer WHERE instr(Upper(SP05),'" & ChgSQL(UCase(Text2)) & "')>0 and sp01 in (" & SQLGrpStr(strTemp, 5) & ") and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) and sp09=na01(+) "
   '查詢法務案件
       'edit by nickc 2005/05/10
       'strSQL = strSQL + " union all select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,LC05 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu04,lc11) AS 申請人,'' AS 相關人 from lawcase,nation,customer WHERE instr(Upper(LC05),'" & ChgSQL(UCase(Text2)) & "')>0 and lc01 in (" & SQLGrpStr(strTemp, 3) & ") and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) and lc15=na01(+) "
       'Modified by Lydia 2019/11/01 +增加欄位SeColLC
       strSql = strSql + " union all select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,LC05 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu04,lc11) AS 申請人,'' AS 相關人,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   as FSort" & SeColLC & _
                  " from lawcase,nation,customer WHERE instr(Upper(LC05),'" & ChgSQL(UCase(Text2)) & "')>0 and lc01 in (" & SQLGrpStr(strTemp, 3) & ") and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) and lc15=na01(+) "
   '查詢顧問案件基本資料
       'edit by nickc 2005/05/10
       'strSQL = strSQL + " union all select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,na03 AS 申請國家,' ' AS 商品類別,nvl(cu04,hc05) AS 申請人,'' AS 相關人 from hirecase,nation,customer WHERE instr(upper(HC06),'" & UCase(ChgSQL(Text2)) & "')>0 and hc01 in (" & SQLGrpStr(strTemp, 4) & ") and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and '000'=na01(+) "
       'Modified by Lydia 2019/11/01 +增加欄位SeColHC
       strSql = strSql + " union all select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,na03 AS 申請國家,' ' AS 商品類別,nvl(cu04,hc05) AS 申請人,'' AS 相關人,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort" & SeColHC & _
                  " from hirecase,nation,customer WHERE instr(upper(HC06),'" & UCase(ChgSQL(Text2)) & "')>0 and hc01 in (" & SQLGrpStr(strTemp, 4) & ") and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and '000'=na01(+) "
       If txt3 = "Y" Then
   'edit by nickc 2005/05/10
   '        strSQL = strSQL + " union all select ' ' AS V,'N'||trademark.TM01 ||'-'|| trademark.TM02 ||'-'|| trademark.TM03 ||'-'|| trademark.TM04||DECODE(trademark.TM29,'Y','＊','') AS 本所案號,CP37 AS 案件名稱,nation.NA03 AS 申請國家,trademark.TM09 AS 商品類別,nvl(cu04,tm23) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,trademark,nation,customer WHERE instr(Upper(CP37),'" & ChgSQL(UCase(Text2)) & "')>0 AND CP01=tm01(+) AND CP02=tm02(+) AND CP03=tm03(+) AND CP04=tm04(+) AND TM10=na01(+) and cp01 in (" & SQLGrpStr(strTemp, 2) & ") and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) and cp37 is not null "
   '        strSQL = strSQL + " union all select ' ' AS V,'N'||patent.PA01 ||'-'|| patent.PA02 ||'-'|| patent.PA03 ||'-'|| patent.PA04||DECODE(patent.PA57,'Y','＊','') AS 本所案號,CP37 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu04,pa26) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,patent,nation,customer WHERE instr(Upper(CP37),'" & ChgSQL(UCase(Text2)) & "')>0 AND CP01=pa01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 1) & ") and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) and cp37 is not null "
   '        strSQL = strSQL + " union all select ' ' AS V,'N'||servicepractice.SP01 ||'-'|| servicepractice.SP02 ||'-'|| servicepractice.SP03 ||'-'|| servicepractice.SP04||DECODE(servicepractice.SP15,'Y','＊','') AS 本所案號,CP37 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu04,sp08) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,servicepractice,nation,customer WHERE instr(Upper(CP37),'" & ChgSQL(UCase(Text2)) & "')>0 AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND sp09=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 5) & ") and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) and cp37 is not null "
   '        strSQL = strSQL + " union all select ' ' AS V,'N'||lawcase.LC01 ||'-'|| lawcase.LC02 ||'-'|| lawcase.LC03 ||'-'|| lawcase.LC04||DECODE(lawcase.LC08,'Y','＊','') AS 本所案號,CP37 AS 案件名稱,NATION.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu04,lc11) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,lawcase,nation,customer WHERE instr(Upper(CP37),'" & ChgSQL(UCase(Text2)) & "')>0 AND cp01=LC01(+) AND cp01=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND lc15=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 3) & ") and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) and cp37 is not null "
   '        strSQL = strSQL + " union all select ' ' AS V,'N'||hirecase.HC01 ||'-'|| hirecase.HC02 ||'-'|| hirecase.HC03 ||'-'|| hirecase.HC04||DECODE(hirecase.HC09,'Y','＊','') AS 本所案號,CP37 AS 案件名稱,na03 AS 申請國家,' ' AS 商品類別,nvl(cu04,hc05) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,hirecase,nation,customer WHERE instr(Upper(CP37),'" & ChgSQL(UCase(Text2)) & "')>0 AND Cp01=hc01(+) AND Cp02=hC02(+) AND Cp03=hC03(+) AND Cp04=hc04(+) and '000'=na01(+) and cp01 in (" & SQLGrpStr(strTemp, 4) & ") and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and cp37 is not null "
           'Modified by Lydia 2019/11/01 +增加欄位SeColTM,SeColPA,SeColLC,SeColHC,SeColSP
           strSql = strSql + " union all select ' ' AS V,'N'||decode(trademark.tm28,'1','','N')||trademark.TM01 ||'-'|| trademark.TM02 ||'-'|| trademark.TM03 ||'-'|| trademark.TM04||DECODE(trademark.TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,CP37 AS 案件名稱,nation.NA03 AS 申請國家,trademark.TM09 AS 商品類別,nvl(cu04,tm23) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,trademark.TM01 ||'-'|| trademark.TM02 ||'-'|| trademark.TM03 ||'-'|| trademark.TM04||DECODE(trademark.TM29,'Y','＊','') as FSort" & SeColTM & _
                                   " from CASEPROGRESS,trademark,nation,customer WHERE instr(Upper(CP37),'" & ChgSQL(UCase(Text2)) & "')>0 AND CP01=tm01(+) AND CP02=tm02(+) AND CP03=tm03(+) AND CP04=tm04(+) AND TM10=na01(+) and cp01 in (" & SQLGrpStr(strTemp, 2) & ") and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) and cp37 is not null "
           strSql = strSql + " union all select ' ' AS V,'N'||decode(patent.pa23,'1','','N')||patent.PA01 ||'-'|| patent.PA02 ||'-'|| patent.PA03 ||'-'|| patent.PA04||DECODE(patent.PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,CP37 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu04,pa26) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,patent.PA01 ||'-'|| patent.PA02 ||'-'|| patent.PA03 ||'-'|| patent.PA04||DECODE(patent.PA57,'Y','＊','') as FSort" & SeColPA & _
                                    " from CASEPROGRESS,patent,nation,customer WHERE instr(Upper(CP37),'" & ChgSQL(UCase(Text2)) & "')>0 AND CP01=pa01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 1) & ") and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) and cp37 is not null "
           'Modify by Amy 2020/02/05 +SP73 商品類別
           strSql = strSql + " union all select ' ' AS V,'N'||servicepractice.SP01 ||'-'|| servicepractice.SP02 ||'-'|| servicepractice.SP03 ||'-'|| servicepractice.SP04||DECODE(servicepractice.SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,CP37 AS 案件名稱,nation.NA03 AS 申請國家,NVL(SP73,'') AS 商品類別,nvl(cu04,sp08) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,servicepractice.SP01 ||'-'|| servicepractice.SP02 ||'-'|| servicepractice.SP03 ||'-'|| servicepractice.SP04||DECODE(servicepractice.SP15,'Y','＊','') as FSort" & SeColSP & _
                                    " from CASEPROGRESS,servicepractice,nation,customer WHERE instr(Upper(CP37),'" & ChgSQL(UCase(Text2)) & "')>0 AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND sp09=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 5) & ") and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) and cp37 is not null "
           strSql = strSql + " union all select ' ' AS V,'N'||lawcase.LC01 ||'-'|| lawcase.LC02 ||'-'|| lawcase.LC03 ||'-'|| lawcase.LC04||DECODE(lawcase.LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,CP37 AS 案件名稱,NATION.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu04,lc11) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,lawcase.LC01 ||'-'|| lawcase.LC02 ||'-'|| lawcase.LC03 ||'-'|| lawcase.LC04||DECODE(lawcase.LC08,'Y','＊','') as FSort" & SeColLC & _
                                    " from CASEPROGRESS,lawcase,nation,customer WHERE instr(Upper(CP37),'" & ChgSQL(UCase(Text2)) & "')>0 AND cp01=LC01(+) AND cp01=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND lc15=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 3) & ") and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) and cp37 is not null "
           strSql = strSql + " union all select ' ' AS V,'N'||hirecase.HC01 ||'-'|| hirecase.HC02 ||'-'|| hirecase.HC03 ||'-'|| hirecase.HC04||DECODE(hirecase.HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,CP37 AS 案件名稱,na03 AS 申請國家,' ' AS 商品類別,nvl(cu04,hc05) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,hirecase.HC01 ||'-'|| hirecase.HC02 ||'-'|| hirecase.HC03 ||'-'|| hirecase.HC04||DECODE(hirecase.HC09,'Y','＊','') as FSort" & SeColHC & _
                                   " from CASEPROGRESS,hirecase,nation,customer WHERE instr(Upper(CP37),'" & ChgSQL(UCase(Text2)) & "')>0 AND Cp01=hc01(+) AND Cp02=hC02(+) AND Cp03=hC03(+) AND Cp04=hc04(+) and '000'=na01(+) and cp01 in (" & SQLGrpStr(strTemp, 4) & ") and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and cp37 is not null "
           'end 2019/11/01
       End If
       'edit by nickc 2005/05/10
       'strSQL = strSQL + " ORDER BY 本所案號 "
       strSql = strSql + " ORDER BY FSort,本所案號 "
   Else
       If Option1(1).Value = True Then                 '英文
            pub_QL05 = pub_QL05 & ";" & Option1(1).Caption 'Add By Sindy 2010/10/27
            If Text5 <> "" Then
               pub_QL05 = pub_QL05 & ";" & Label3 & Text5 'Add By Sindy 2010/10/27
            End If
       '查詢商標基本檔
           'EDIT BY nickc 2005/05/10
           'strSQL = "select ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,TM06 AS 案件名稱,nation.na04 AS 申請國家,TM09 AS 商品類別,nvl(cu05||' '||cu88||'-'||cu89||'-'||cu90,tm23) AS 申請人, '' AS 相關人 from trademark,nation,customer WHERE instr(upper(TM06),'" & UCase(ChgSQL(Text2)) & "')>0 and tm10=na01(+) and tm01 in (" & SQLGrpStr(strTemp, 2) & ") and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) "
           'Modified by Lydia 2019/11/01 +增加欄位SeColTM
           strSql = "select ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,TM05 AS 案件名稱,nation.na04 AS 申請國家,TM09 AS 商品類別,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),tm23) AS 申請人, '' AS 相關人,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as FSort" & SeColTM & _
                        " from trademark,nation,customer WHERE instr(upper(TM05),'" & UCase(ChgSQL(Text2)) & "')>0 and tm10=na01(+) and tm01 in (" & SQLGrpStr(strTemp, 2) & ") and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) "
       '查詢專利基本檔
           'edit by nickc 2005/05/10
           'strSQL = strSQL + " union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,PA06 AS 案件名稱,nation.na04 AS 申請國家,' ' AS 商品類別,nvl(cu05||' '||cu88||'-'||cu89||'-'||cu90,pa26) AS 申請人,'' AS 相關人 from patent,nation,customer WHERE instr(upper(PA06),'" & UCase(ChgSQL(Text2)) & "')>0 and pa09=na01(+) and pa01 in (" & SQLGrpStr(strTemp, 1) & ") and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) "
           'Modified by Lydia 2019/11/01 +增加欄位SeColPA
           strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,PA06 AS 案件名稱,nation.na04 AS 申請國家,' ' AS 商品類別,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa26) AS 申請人,'' AS 相關人,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort" & SeColPA & _
                       " from patent,nation,customer WHERE instr(upper(PA06),'" & UCase(ChgSQL(Text2)) & "')>0 and pa09=na01(+) and pa01 in (" & SQLGrpStr(strTemp, 1) & ") and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) "
       '查詢服務業務
           'edit by nickc 2005/05/10
           'strSQL = strSQL + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,SP06 AS 案件名稱,nation.na04 AS 申請國家,' ' AS 商品類別,nvl(cu05||' '||cu88||'-'||cu89||'-'||cu90,sp08) AS 申請人,'' AS 相關人 from servicepractice,nation,customer WHERE instr(upper(SP06),'" & UCase(ChgSQL(Text2)) & "')>0 and sp09=na01(+) and sp01 in (" & SQLGrpStr(strTemp, 5) & ") and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) "
            'Modified by Lydia 2019/11/01 +增加欄位SeColSP
            'Modify by Amy 2020/02/05 +SP73 商品類別
           strSql = strSql + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,SP06 AS 案件名稱,nation.na04 AS 申請國家,NVL(SP73,'') AS 商品類別,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp08) AS 申請人,'' AS 相關人,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort" & SeColSP & _
                      " from servicepractice,nation,customer WHERE instr(upper(SP06),'" & UCase(ChgSQL(Text2)) & "')>0 and sp09=na01(+) and sp01 in (" & SQLGrpStr(strTemp, 5) & ") and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) "
       '查詢法務案件
           'edit by nickc 2005/05/10
           'strSQL = strSQL + " union all select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,LC06 AS 案件名稱,nation.na04 AS 申請國家,' ' AS 商品類別,nvl(cu05||' '||cu88||'-'||cu89||'-'||cu90,lc11) AS 申請人,'' AS 相關人 from lawcase,nation,customer WHERE instr(upper(LC06),'" & UCase(ChgSQL(Text2)) & "')>0 and lc15=na01(+) and lc01 in (" & SQLGrpStr(strTemp, 3) & ") and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) "
            'Modified by Lydia 2019/11/01 +增加欄位SeColLC
           strSql = strSql + " union all select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,LC06 AS 案件名稱,nation.na04 AS 申請國家,' ' AS 商品類別,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc11) AS 申請人,'' AS 相關人,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   as FSort" & SeColLC & _
                      " from lawcase,nation,customer WHERE instr(upper(LC06),'" & UCase(ChgSQL(Text2)) & "')>0 and lc15=na01(+) and lc01 in (" & SQLGrpStr(strTemp, 3) & ") and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) "
       '查詢顧問案件基本資料
           'edit by nickc 2005/05/10
           'strSQL = strSQL + " union all select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,na04 AS 申請國家,' ' AS 商品類別,nvl(cu05||' '||cu88||'-'||cu89||'-'||cu90,hc05) AS 申請人,'' AS 相關人 from hirecase,customer,nation WHERE instr(upper(HC06),'" & UCase(ChgSQL(Text2)) & "')>0 and hc01 in (" & SQLGrpStr(strTemp, 4) & ") and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and '000'=na01(+) "
            'Modified by Lydia 2019/11/01 +增加欄位SeColHC
           strSql = strSql + " union all select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,na04 AS 申請國家,' ' AS 商品類別,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc05) AS 申請人,'' AS 相關人,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort" & SeColHC & _
                      " from hirecase,customer,nation WHERE instr(upper(HC06),'" & UCase(ChgSQL(Text2)) & "')>0 and hc01 in (" & SQLGrpStr(strTemp, 4) & ") and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and '000'=na01(+) "
           If txt3 = "Y" Then
   
   'edit by nick 2005/05/10
   '                strSQL = strSQL + " union all select ' ' AS V,'N'||trademark.TM01 ||'-'|| trademark.TM02 ||'-'|| trademark.TM03 ||'-'|| trademark.TM04||DECODE(trademark.TM29,'Y','＊','') AS 本所案號,CP38 AS 案件名稱,nation.na04 AS 申請國家,trademark.TM09 AS 商品類別,nvl(cu05||' '||cu88||'-'||cu89||'-'||cu90,tm23) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,trademark,nation,customer WHERE instr(upper(CP38),'" & UCase(ChgSQL(Text2)) & "')>0 AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND tm10=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 2) & ") and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) "
   '                strSQL = strSQL + " union all select ' ' AS V,'N'||patent.PA01 ||'-'|| patent.PA02 ||'-'|| patent.PA03 ||'-'|| patent.PA04||DECODE(patent.PA57,'Y','＊','') AS 本所案號,CP38 AS 案件名稱,nation.na04 AS 申請國家,' ' AS 商品類別,nvl(cu05||' '||cu88||'-'||cu89||'-'||cu90,pa26) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,patent,nation,customer WHERE instr(upper(CP38),'" & UCase(ChgSQL(Text2)) & "')>0 AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 1) & ") and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) "
   '                strSQL = strSQL + " union all select ' ' AS V,'N'||servicepractice.SP01 ||'-'|| servicepractice.SP02 ||'-'|| servicepractice.SP03 ||'-'|| servicepractice.SP04||DECODE(servicepractice.SP15,'Y','＊','') AS 本所案號,CP38 AS 案件名稱,nation.na04 AS 申請國家,' ' AS 商品類別,nvl(cu05||' '||cu88||'-'||cu89||'-'||cu90,sp08) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,servicepractice,nation,customer WHERE instr(upper(CP38),'" & UCase(ChgSQL(Text2)) & "')>0 AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND SP09=na01(+) and cp01 in (" & SQLGrpStr(strTemp, 5) & ") and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) "
   '                strSQL = strSQL + " union all select ' ' AS V,'N'||lawcase.LC01 ||'-'|| lawcase.LC02 ||'-'|| lawcase.LC03 ||'-'|| lawcase.LC04||DECODE(lawcase.LC08,'Y','＊','') AS 本所案號,CP38 AS 案件名稱,NATION.na04 AS 申請國家,' ' AS 商品類別,nvl(cu05||' '||cu88||'-'||cu89||'-'||cu90,lc11) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,lawcase,nation,customer WHERE instr(upper(CP38),'" & UCase(ChgSQL(Text2)) & "')>0 AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND lc15=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 3) & ") and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) "
   '                strSQL = strSQL + " union all select ' ' AS V,'N'||hirecase.HC01 ||'-'|| hirecase.HC02 ||'-'|| hirecase.HC03 ||'-'|| hirecase.HC04||DECODE(hirecase.HC09,'Y','＊','') AS 本所案號,CP38 AS 案件名稱,na04 AS 申請國家,' ' AS 商品類別,nvl(cu05||' '||cu88||'-'||cu89||'-'||cu90,hc05) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,hirecase,nation,customer WHERE instr(upper(CP38),'" & UCase(ChgSQL(Text2)) & "')>0 AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) and cp01 in (" & SQLGrpStr(strTemp, 4) & ") and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and '000'=na01(+) "
                   'Modified by Lydia 2019/11/01 +增加欄位SeColTM,SeColPA,SeColLC,SeColHC,SeColSP
                   strSql = strSql + " union all select ' ' AS V,'N'||decode(tm28,'1','','N')||trademark.TM01 ||'-'|| trademark.TM02 ||'-'|| trademark.TM03 ||'-'|| trademark.TM04||DECODE(trademark.TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,CP38 AS 案件名稱,nation.na04 AS 申請國家,trademark.TM09 AS 商品類別,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),tm23) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,trademark.TM01 ||'-'|| trademark.TM02 ||'-'|| trademark.TM03 ||'-'|| trademark.TM04||DECODE(trademark.TM29,'Y','＊','') as FSort" & SeColTM & _
                               " from CASEPROGRESS,trademark,nation,customer WHERE instr(upper(CP38),'" & UCase(ChgSQL(Text2)) & "')>0 AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND tm10=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 2) & ") and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) "
                   strSql = strSql + " union all select ' ' AS V,'N'||decode(pa23,'1','','N')||patent.PA01 ||'-'|| patent.PA02 ||'-'|| patent.PA03 ||'-'|| patent.PA04||DECODE(patent.PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,CP38 AS 案件名稱,nation.na04 AS 申請國家,' ' AS 商品類別,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa26) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,patent.PA01 ||'-'|| patent.PA02 ||'-'|| patent.PA03 ||'-'|| patent.PA04||DECODE(patent.PA57,'Y','＊','') as FSort" & SeColPA & _
                               " from CASEPROGRESS,patent,nation,customer WHERE instr(upper(CP38),'" & UCase(ChgSQL(Text2)) & "')>0 AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 1) & ") and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) "
                   'Modify by Amy 2020/02/05 +SP73 商品類別
                   strSql = strSql + " union all select ' ' AS V,'N'||servicepractice.SP01 ||'-'|| servicepractice.SP02 ||'-'|| servicepractice.SP03 ||'-'|| servicepractice.SP04||DECODE(servicepractice.SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,CP38 AS 案件名稱,nation.na04 AS 申請國家,NVL(SP73,'') AS 商品類別,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),sp08) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,servicepractice.SP01 ||'-'|| servicepractice.SP02 ||'-'|| servicepractice.SP03 ||'-'|| servicepractice.SP04||DECODE(servicepractice.SP15,'Y','＊','') as FSort" & SeColSP & _
                               " from CASEPROGRESS,servicepractice,nation,customer WHERE instr(upper(CP38),'" & UCase(ChgSQL(Text2)) & "')>0 AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND SP09=na01(+) and cp01 in (" & SQLGrpStr(strTemp, 5) & ") and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) "
                   strSql = strSql + " union all select ' ' AS V,'N'||lawcase.LC01 ||'-'|| lawcase.LC02 ||'-'|| lawcase.LC03 ||'-'|| lawcase.LC04||DECODE(lawcase.LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,CP38 AS 案件名稱,NATION.na04 AS 申請國家,' ' AS 商品類別,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),lc11) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,lawcase.LC01 ||'-'|| lawcase.LC02 ||'-'|| lawcase.LC03 ||'-'|| lawcase.LC04||DECODE(lawcase.LC08,'Y','＊','') as FSort" & SeColLC & _
                               " from CASEPROGRESS,lawcase,nation,customer WHERE instr(upper(CP38),'" & UCase(ChgSQL(Text2)) & "')>0 AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND lc15=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 3) & ") and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) "
                   strSql = strSql + " union all select ' ' AS V,'N'||hirecase.HC01 ||'-'|| hirecase.HC02 ||'-'|| hirecase.HC03 ||'-'|| hirecase.HC04||DECODE(hirecase.HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,CP38 AS 案件名稱,na04 AS 申請國家,' ' AS 商品類別,nvl(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),hc05) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,hirecase.HC01 ||'-'|| hirecase.HC02 ||'-'|| hirecase.HC03 ||'-'|| hirecase.HC04||DECODE(hirecase.HC09,'Y','＊','') as FSort" & SeColHC & _
                               " from CASEPROGRESS,hirecase,nation,customer WHERE instr(upper(CP38),'" & UCase(ChgSQL(Text2)) & "')>0 AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=HC04(+) and cp01 in (" & SQLGrpStr(strTemp, 4) & ") and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and '000'=na01(+) "
                   'end 2019/11/01
           End If
            'strSQL = strSQL + "  ORDER BY 本所案號 "
            strSql = strSql + "  ORDER BY FSort,本所案號 "
       Else
           If Option1(2).Value = True Then                   '日文
               pub_QL05 = pub_QL05 & ";" & Option1(2).Caption 'Add By Sindy 2010/10/27
               If Text5 <> "" Then
                  pub_QL05 = pub_QL05 & ";" & Label3 & Text5 'Add By Sindy 2010/10/27
               End If
           '查詢商標基本檔
               'edit by nickc 2005/05/10
               'strSQL = "select ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,TM07 AS 案件名稱,nation.NA03 AS 申請國家,TM09 AS 商品類別,nvl(cu06,tm23) AS 申請人, '' AS 相關人 from trademark,nation,customer WHERE instr(TM07,'" & ChgSQL(Text2) & "')>0 and tm10=na01(+) and tm01 in (" & SQLGrpStr(strTemp, 2) & ") and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) "
               'Modified by Lydia 2019/11/01 +增加欄位SeColTM
               strSql = "select ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,TM05 AS 案件名稱,nation.NA03 AS 申請國家,TM09 AS 商品類別,nvl(cu06,tm23) AS 申請人, '' AS 相關人,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') as FSort" & SeColTM & _
                            " from trademark,nation,customer WHERE instr(TM05,'" & ChgSQL(Text2) & "')>0 and tm10=na01(+) and tm01 in (" & SQLGrpStr(strTemp, 2) & ") and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) "
           '查詢專利基本檔
               'edit by nickc 2005/05/10
               'strSQL = strSQL + " union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,PA07 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu06,pa26) AS 申請人,'' AS 相關人 from patent,nation,customer WHERE instr(PA07,'" & ChgSQL(Text2) & "')>0 and pa09=na01(+) and pa01 in (" & SQLGrpStr(strTemp, 1) & ") and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) "
               'Modified by Lydia 2019/11/01 +增加欄位SeColPA
               strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,PA07 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu06,pa26) AS 申請人,'' AS 相關人,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') as FSort" & SeColPA & _
                            " from patent,nation,customer WHERE instr(PA07,'" & ChgSQL(Text2) & "')>0 and pa09=na01(+) and pa01 in (" & SQLGrpStr(strTemp, 1) & ") and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) "
           '查詢服務業務
               'edit by nickc 2005/05/10
               'strSQL = strSQL + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,SP07 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu06,sp08) AS 申請人,'' AS 相關人 from servicepractice,nation,customer WHERE instr(SP07,'" & ChgSQL(Text2) & "')>0 and sp09=na01(+) and sp01 in (" & SQLGrpStr(strTemp, 5) & ") and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) "
               'Modified by Lydia 2019/11/01 +增加欄位SeColSP
               'Modify by Amy 2020/02/05 +SP73 商品類別
               strSql = strSql + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,SP07 AS 案件名稱,nation.NA03 AS 申請國家,NVL(SP73,'') AS 商品類別,nvl(cu06,sp08) AS 申請人,'' AS 相關人,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  as FSort" & SeColSP & _
                            " from servicepractice,nation,customer WHERE instr(SP07,'" & ChgSQL(Text2) & "')>0 and sp09=na01(+) and sp01 in (" & SQLGrpStr(strTemp, 5) & ") and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) "
           '查詢法務案件
               'edit by nickc 2005/05/10
               'strSQL = strSQL + " union all select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,LC07 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu06,lc11) AS 申請人,'' AS 相關人 from lawcase,nation,customer WHERE instr(LC07,'" & ChgSQL(Text2) & "')>0 and lc15=na01(+) and lc01 in (" & SQLGrpStr(strTemp, 3) & ") and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) "
               'Modified by Lydia 2019/11/01 +增加欄位SeColLC
               strSql = strSql + " union all select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,LC07 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu06,lc11) AS 申請人,'' AS 相關人,LC01||'-'||LC02||'-'||LC03||'-'||LC04||DECODE(LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●')   as FSort" & SeColLC & _
                            " from lawcase,nation,customer WHERE instr(LC07,'" & ChgSQL(Text2) & "')>0 and lc15=na01(+) and lc01 in (" & SQLGrpStr(strTemp, 3) & ") and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) "
           '查詢顧問案件基本資料
               'edit by nickc 2005/05/10
               'strSQL = strSQL + " union all select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,na03 AS 申請國家,' ' AS 商品類別,nvl(cu06,hc05) AS 申請人,'' AS 相關人 from hirecase,customer,nation WHERE instr(upper(HC06),'" & UCase(ChgSQL(Text2)) & "')>0 and hc01 in (" & SQLGrpStr(strTemp, 4) & ") and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and '000'=na01(+) "
               'Modified by Lydia 2019/11/01 +增加欄位SeColHC
               strSql = strSql + " union all select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,na03 AS 申請國家,' ' AS 商品類別,nvl(cu06,hc05) AS 申請人,'' AS 相關人,HC01||'-'||HC02||'-'||HC03||'-'||HC04||DECODE(HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●')  as FSort" & SeColHC & _
                            " from hirecase,customer,nation WHERE instr(upper(HC06),'" & UCase(ChgSQL(Text2)) & "')>0 and hc01 in (" & SQLGrpStr(strTemp, 4) & ") and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and '000'=na01(+) "
               If txt3 = "Y" Then
                  '查詢案件進度檔
   'edit by nickc 2005/05/10
   '                   strSQL = strSQL + " union all select ' ' AS V,'N'||trademark.TM01 ||'-'|| trademark.TM02 ||'-'|| trademark.TM03 ||'-'|| trademark.TM04||DECODE(trademark.TM29,'Y','＊','') AS 本所案號,CP39 AS 案件名稱,nation.NA03 AS 申請國家,trademark.TM09 AS 商品類別,nvl(cu06,tm23) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,trademark,nation,customer WHERE instr(CP39,'" & ChgSQL(Text2) & "')>0 AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND tm10=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 2) & ") and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) "
   '                   strSQL = strSQL + " union all select ' ' AS V,'N'||patent.PA01 ||'-'|| patent.PA02 ||'-'|| patent.PA03 ||'-'|| patent.PA04||DECODE(patent.PA57,'Y','＊','') AS 本所案號,CP39 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu06,pa26) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,patent,nation,customer WHERE instr(CP39,'" & ChgSQL(Text2) & "')>0 AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 1) & ") and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) "
   '                   strSQL = strSQL + " union all select ' ' AS V,'N'||servicepractice.SP01 ||'-'|| servicepractice.SP02 ||'-'|| servicepractice.SP03 ||'-'|| servicepractice.SP04||DECODE(servicepractice.SP15,'Y','＊','') AS 本所案號,CP39 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu06,sp08) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,customer,servicepractice,nation WHERE instr(CP39,'" & ChgSQL(Text2) & "')>0 AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND sp09=NA01(+) and cp01 iN (" & SQLGrpStr(strTemp, 5) & ") and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) "
   '                   strSQL = strSQL + " union all select ' ' AS V,'N'||lawcase.LC01 ||'-'|| lawcase.LC02 ||'-'|| lawcase.LC03 ||'-'|| lawcase.LC04||DECODE(lawcase.LC08,'Y','＊','') AS 本所案號,CP39 AS 案件名稱,NATION.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu06,lc11) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,lawcase,customer,nation WHERE instr(CP39,'" & ChgSQL(Text2) & "')>0 AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND lc15=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 3) & ") and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) "
   '                   strSQL = strSQL + " union all select ' ' AS V,'N'||hirecase.HC01 ||'-'|| hirecase.HC02 ||'-'|| hirecase.HC03 ||'-'|| hirecase.HC04||DECODE(hirecase.HC09,'Y','＊','') AS 本所案號,CP39 AS 案件名稱,na03 AS 申請國家,' ' AS 商品類別,nvl(cu06,hc05) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人 from CASEPROGRESS,hirecase,nation,customer WHERE instr(CP39,'" & ChgSQL(Text2) & "')>0 AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=hC04(+) and cp01 in (" & SQLGrpStr(strTemp, 4) & ") and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and '000'=na01(+) "
                      'Modified by Lydia 2019/11/01 +增加欄位SeColTM,SeColPA,SeColLC,SeColHC,SeColSP
                      strSql = strSql + " union all select ' ' AS V,'N'||decode(tm28,'1','','N')||trademark.TM01 ||'-'|| trademark.TM02 ||'-'|| trademark.TM03 ||'-'|| trademark.TM04||DECODE(trademark.TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,CP39 AS 案件名稱,nation.NA03 AS 申請國家,trademark.TM09 AS 商品類別,nvl(cu06,tm23) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,trademark.TM01 ||'-'|| trademark.TM02 ||'-'|| trademark.TM03 ||'-'|| trademark.TM04||DECODE(trademark.TM29,'Y','＊','') as FSort" & SeColTM & _
                                " from CASEPROGRESS,trademark,nation,customer WHERE instr(CP39,'" & ChgSQL(Text2) & "')>0 AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND tm10=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 2) & ") and substr(tm23,1,8)=cu01(+) and decode(substr(tm23,9,1),null,'0',substr(tm23,9,1))=cu02(+) "
                      strSql = strSql + " union all select ' ' AS V,'N'||decode(pa23,'1','','N')||patent.PA01 ||'-'|| patent.PA02 ||'-'|| patent.PA03 ||'-'|| patent.PA04||DECODE(patent.PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,CP39 AS 案件名稱,nation.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu06,pa26) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,patent.PA01 ||'-'|| patent.PA02 ||'-'|| patent.PA03 ||'-'|| patent.PA04||DECODE(patent.PA57,'Y','＊','') as FSort" & SeColSP & _
                                " from CASEPROGRESS,patent,nation,customer WHERE instr(CP39,'" & ChgSQL(Text2) & "')>0 AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 1) & ") and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),null,'0',substr(pa26,9,1))=cu02(+) "
                      'Modify by Amy 2020/02/05 +SP73 商品類別
                      strSql = strSql + " union all select ' ' AS V,'N'||servicepractice.SP01 ||'-'|| servicepractice.SP02 ||'-'|| servicepractice.SP03 ||'-'|| servicepractice.SP04||DECODE(servicepractice.SP15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,CP39 AS 案件名稱,nation.NA03 AS 申請國家,NVL(SP73,'') AS 商品類別,nvl(cu06,sp08) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,servicepractice.SP01 ||'-'|| servicepractice.SP02 ||'-'|| servicepractice.SP03 ||'-'|| servicepractice.SP04||DECODE(servicepractice.SP15,'Y','＊','') as FSort" & SeColSP & _
                                " from CASEPROGRESS,customer,servicepractice,nation WHERE instr(CP39,'" & ChgSQL(Text2) & "')>0 AND cP01=sP01(+) AND cP02=sP02(+) AND cP03=sP03(+) AND cP04=sP04(+) AND sp09=NA01(+) and cp01 iN (" & SQLGrpStr(strTemp, 5) & ") and substr(sp08,1,8)=cu01(+) and decode(substr(sp08,9,1),null,'0',substr(sp08,9,1))=cu02(+) "
                      strSql = strSql + " union all select ' ' AS V,'N'||lawcase.LC01 ||'-'|| lawcase.LC02 ||'-'|| lawcase.LC03 ||'-'|| lawcase.LC04||DECODE(lawcase.LC08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,CP39 AS 案件名稱,NATION.NA03 AS 申請國家,' ' AS 商品類別,nvl(cu06,lc11) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,lawcase.LC01 ||'-'|| lawcase.LC02 ||'-'|| lawcase.LC03 ||'-'|| lawcase.LC04||DECODE(lawcase.LC08,'Y','＊','') as FSort" & SeColLC & _
                                " from CASEPROGRESS,lawcase,customer,nation WHERE instr(CP39,'" & ChgSQL(Text2) & "')>0 AND cp01=LC01(+) AND cp02=LC02(+) AND cp03=LC03(+) AND cp04=LC04(+) AND lc15=NA01(+) and cp01 in (" & SQLGrpStr(strTemp, 3) & ") and substr(lc11,1,8)=cu01(+) and decode(substr(lc11,9,1),null,'0',substr(lc11,9,1))=cu02(+) "
                      strSql = strSql + " union all select ' ' AS V,'N'||hirecase.HC01 ||'-'|| hirecase.HC02 ||'-'|| hirecase.HC03 ||'-'|| hirecase.HC04||DECODE(hirecase.HC09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,CP39 AS 案件名稱,na03 AS 申請國家,' ' AS 商品類別,nvl(cu06,hc05) AS 申請人,NVL(NVL(CP40,CP41),CP42) AS 相關人,hirecase.HC01 ||'-'|| hirecase.HC02 ||'-'|| hirecase.HC03 ||'-'|| hirecase.HC04||DECODE(hirecase.HC09,'Y','＊','') as FSort" & SeColHC & _
                                " from CASEPROGRESS,hirecase,nation,customer WHERE instr(CP39,'" & ChgSQL(Text2) & "')>0 AND cp01=HC01(+) AND cp02=HC02(+) AND cp03=HC03(+) AND cp04=hC04(+) and cp01 in (" & SQLGrpStr(strTemp, 4) & ") and substr(hc05,1,8)=cu01(+) and decode(substr(hc05,9,1),null,'0',substr(hc05,9,1))=cu02(+) and '000'=na01(+) "
                      'end 2019/11/01
               End If
               'edit by nickc 2005/05/10
               'strSQL = strSQL + " ORDER BY 本所案號 "
               strSql = strSql + " ORDER BY FSort,本所案號 "
            End If
        End If
   End If
   pub_QL05 = pub_QL05 & ";" & Label2 & Text2 'Add By Sindy 2010/10/27
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   'Modified by Lydia 2019/11/01 改變型態
   'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic

   If adoRecordset.RecordCount <> 0 Then
      dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
      txtRows = dblRow 'Added by Morgan 2025/10/2
      
        'Added by Lydia 2019/11/01 逐案號判斷
        If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, strTemp, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
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
            InsertQueryLog (dblRow) 'Add By Sindy 2010/10/27
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
        Else
            InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/10/27
        End If
       'end 2019/11/01
       
       grdDataList.Rows = adoRecordset.RecordCount + 1
       If Not cmdOK(2).Enabled Then cmdOK(2).Enabled = True
       If Not cmdOK(3).Enabled Then cmdOK(3).Enabled = True
       If Not cmdOK(6).Enabled Then cmdOK(6).Enabled = True 'Add By Sindy 2016/2/22
   Else
       InsertQueryLog (0) 'Add By Sindy 2010/10/27
JumpToNoData:   'Added by Lydia 2019/11/01
       cmdOK(2).Enabled = False
       cmdOK(3).Enabled = False
       cmdOK(6).Enabled = False 'Add By Sindy 2016/2/22
       CheckOC
       grdDataList.Visible = True
       Me.Enabled = True
       ShowNoData
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   '把資料放進   GRID
   '911029 nick edit
   If adoRecordset.RecordCount <> 0 Then
       Set grdDataList.Recordset = adoRecordset
   End If
   CheckOC
   '若僅有一筆資料則自動勾選並重設底色
   If Me.grdDataList.Rows = 2 And Len(Me.grdDataList.TextMatrix(1, 1)) > 0 Then
      grdDataList.row = 1
      grdDataList.col = 0
      grdDataList.Text = "V"
      For i = 0 To grdDataList.Cols - 1
          grdDataList.col = i
          grdDataList.CellBackColor = &HFFC0C0
      Next i
      grdDataList.col = 1
      strCP01 = SystemNumber(Pub_RplStr(grdDataList.Text), 1)
      strCP02 = SystemNumber(Pub_RplStr(grdDataList.Text), 2)
      strCP03 = SystemNumber(Pub_RplStr(grdDataList.Text), 3)
      strCP04 = SystemNumber(Pub_RplStr(grdDataList.Text), 4)
      Call SetCmdImg(strCP01, strCP02, strCP03, strCP04) '檢查有無代表圖
   End If
   Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100103_1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
           If Option1(0).Value = True Then
               Option1(1).Value = False
               Option1(2).Value = False
               'edit by nickc 2007/06/06 切換輸入法改用API
               'Text2.IMEMode = 1
               OpenIme
               Text2.SetFocus
               Text2_GotFocus
           End If
      Case 1
           If Option1(1).Value = True Then
              Option1(0).Value = False
              Option1(2).Value = False
              'edit by nickc 2007/06/06 切換輸入法改用API
              'Text2.IMEMode = 2
              CloseIme
              Text2.SetFocus
              Text2_GotFocus
           End If
      Case 2
           If Option1(2).Value = True Then
              Option1(0).Value = False
              Option1(1).Value = False
              'edit by nickc 2007/06/06 切換輸入法改用API
              'Text2.IMEMode = 1
              OpenIme
              Text2.SetFocus
              Text2_GotFocus
           End If
      Case Else
   End Select
End Sub

Private Sub grdDataList_SelChange()
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
   Else
      grdDataList.Text = "V"
      For i = 0 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = &HFFC0C0
      Next i
      'Add By Sindy 2016/2/22 檢查有無代表圖
      grdDataList.col = 1
      strCP01 = SystemNumber(Pub_RplStr(grdDataList.Text), 1)
      strCP02 = SystemNumber(Pub_RplStr(grdDataList.Text), 2)
      strCP03 = SystemNumber(Pub_RplStr(grdDataList.Text), 3)
      strCP04 = SystemNumber(Pub_RplStr(grdDataList.Text), 4)
      Call SetCmdImg(strCP01, strCP02, strCP03, strCP04)
      '2016/2/22 END
   End If
   End If
   grdDataList.Visible = True
End Sub

Private Sub Text2_GotFocus()
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2)
   'edit by nickc 2007/06/06 切換輸入法改用API
   OpenIme
End Sub

Private Sub Text2_LostFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   CloseIme
End Sub

Private Sub Text5_GotFocus()
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5)
   CloseIme
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_LostFocus()
   'Modify By Cheng 2002/03/14
   ''Add By Cheng 2002/01/07
   'Me.Text5.Text = GetAllSysKind(Me.Text5)
End Sub

'add by sonia 2014/10/29
Private Sub TXT3_GotFocus()
   txt3.SelStart = 0
   txt3.SelLength = Len(txt3)
   CloseIme
End Sub
'end 2014/10/29

Private Sub txt3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TXT3_LostFocus()
   If txt3 <> "Y" And Trim(txt3) <> "" Then
      s = MsgBox("只能輸入 Y OR 空白!!")
   End If
End Sub
