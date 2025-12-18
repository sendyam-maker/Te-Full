VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040117 
   BorderStyle     =   1  '單線固定
   Caption         =   "發後補看作業-函知客戶"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8955
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   315
      Left            =   150
      TabIndex        =   13
      Top             =   390
      Visible         =   0   'False
      Width           =   2100
      Begin MSForms.TextBox txtInput 
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   1635
         VariousPropertyBits=   679493659
         MaxLength       =   100
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "全選"
      Height          =   345
      Left            =   3015
      TabIndex        =   11
      Top             =   510
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "顯示代表圖(&I)"
      Height          =   345
      Index           =   4
      Left            =   3690
      TabIndex        =   10
      Top             =   510
      Width           =   1275
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H0000FF00&
      Caption         =   "確認(&O)"
      Height          =   345
      Left            =   4995
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   510
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "重整"
      Height          =   345
      Left            =   3015
      TabIndex        =   7
      Top             =   120
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "本程序卷宗(&I)"
      Height          =   345
      Index           =   2
      Left            =   6300
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "完整卷宗(&H)"
      Height          =   345
      Index           =   3
      Left            =   3690
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "案件進度(&C)"
      Height          =   345
      Index           =   1
      Left            =   4995
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   0
      Left            =   7605
      TabIndex        =   2
      Top             =   120
      Width           =   1005
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4725
      Left            =   135
      TabIndex        =   3
      Top             =   900
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   8334
      _Version        =   393216
      Cols            =   10
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|發文日|本所案號|案件名稱|國家|種類|案件性質|本所期限|承辦人|智權人員"
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
      _Band(0).Cols   =   10
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      TabIndex        =   12
      Top             =   120
      Width           =   1725
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3043;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "0 / 0"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   6795
      TabIndex        =   9
      Top             =   600
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "補看人員："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(註:雙擊開啟本程序卷宗畫面)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   660
      Width           =   2325
   End
End
Attribute VB_Name = "frm040117"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/18 改成Form2.0 ; txtInput
'Memo by Lydia 2021/10/07 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、Combo1
'Created by Morgan 2014/12/16
Option Explicit

Dim iPrevRow As Integer '前次點選列
Dim lTotRows As Long, lSelRows As Long
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_InputCol As Integer, m_InputRow As Integer

Public cmdState As Integer
Public cmdAction As Integer 'Added by Morgan 2018/10/29 卷宗區回傳值 1=確認 2=下一筆 9=結束
Public m_ProState As String 'Add By Sindy 2020/12/7 系統作業


Private Sub cmdOK_Click()
   Dim iRow As Integer, bContinue As Boolean
   Dim iIdx As Integer
   Dim bolShowForm As Boolean
   Dim strCP09 As String

   SetMouseBusy
   bContinue = False
   With MSHFlexGrid1
   iIdx = GetFieldId("cp10", Me.MSHFlexGrid1)
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         bContinue = True
         Exit For
      End If
   Next
   End With
   If bContinue = False Then
      MsgBox "請先勾選(V)資料列！", vbInformation
   Else
      Process
   End If

EXITSUB:

   SetMouseReady
End Sub

Public Sub PubShowNextData()
   Dim StrTag As String
   Dim ii As Integer
   
   If iPrevRow = 0 Then Exit Sub
   Select Case cmdState
   
   Case 1
      Me.Enabled = False
      If fnSaveParentForm(Me) = False Then
         Me.Enabled = True
         Exit Sub
      End If
      Screen.MousePointer = vbHourglass
      frm100101_2.Show
      frm100101_2.Tag = Pub_RplStr(MSHFlexGrid1.TextMatrix(iPrevRow, 1))
      frm100101_2.cmdok(5).Visible = False '下一筆按鈕隱藏
      frm100101_2.StrMenu
      Screen.MousePointer = vbDefault
      Me.Enabled = True
   
   'Modified by Morgan 2018/12/18 增加多選及卷宗區可確認/下一筆功能
   Case 2, 3 '2本程序卷宗 3完整卷宗
      '卷宗區點結束
      If cmdAction = 9 Then cmdAction = 0: Exit Sub
      
      Screen.MousePointer = vbHourglass
      With MSHFlexGrid1
      '卷宗區點補看確認
      If cmdAction = 1 Then
         If UpdateDB(GetValue(iPrevRow, "CP09"), GetValue(iPrevRow, "備註")) = True Then
            .TextMatrix(iPrevRow, 0) = "X"
            .RowHeight(iPrevRow) = 0
            iPrevRow = 0
            
            lSelRows = lSelRows - 1
            lTotRows = lTotRows - 1
            lblCount = lSelRows & " / " & lTotRows
            DoEvents
         Else
            Exit Sub
         End If
         
      '卷宗區點下一筆
      'ElseIf cmdAction = 2 Then
      '   '清除前次勾選
      '   If .TextMatrix(iPrevRow, 0) = "V" Then
      '      .TextMatrix(iPrevRow, 0) = ""
      '   End If
      End If
            
      '檢查下一筆勾選
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) = "V" Then
            SelectRow ii
            cmdAction = 0
            Exit For
         End If
      Next
      End With
      
      '卷宗區回來且已無勾選
      If cmdAction <> 0 Then
         cmdAction = 0
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      
      StrTag = ""
      '本程序卷宗
      If cmdState = 2 Then
         'Modified by Morgan 2018/10/17 通知申請號要看相關收文號的卷宗區--游經理
         If GetValue(iPrevRow, "CP10") = "1101" Then
            StrTag = GetValue(iPrevRow, "CP43")
         Else
            StrTag = GetValue(iPrevRow, "CP09")
         End If
         
      '完整卷宗
      Else
         StrTag = GetValue(iPrevRow, "本所案號")
         If UBound(Split(StrTag, "-")) = 1 Then
            StrTag = StrTag & "-0-00"
         End If
      End If
      
      If StrTag <> "" Then
         cmdAction = 9 '預設結束
         frm100101_L.m_strKey = StrTag
         frm100101_L.SetParent Me
         If frm100101_L.QueryData = True Then
            frm100101_L.cmdok(1).Visible = True
            'frm100101_L.cmdOK(2).Visible = True
            frm100101_L.Show
            Me.Hide
         Else
            Unload frm100101_L
         End If
      End If
      Screen.MousePointer = vbDefault
      
   Case 4 '代表圖
      
      frmPic001.oCP01 = GetValue(iPrevRow, "pa01")
      frmPic001.oCP02 = GetValue(iPrevRow, "pa02")
      frmPic001.oCP03 = GetValue(iPrevRow, "pa03")
      frmPic001.oCP04 = GetValue(iPrevRow, "pa04")
      frmPic001.StrMenu
      frmPic001.cmdok(0).Visible = False
      frmPic001.cmdok(1).Visible = False
      frmPic001.cmdok(2).Visible = False
      frmPic001.cmdok(4).Visible = False
      frmPic001.cmdok(5).Visible = False
      frmPic001.cmdok(6).Visible = False
      frmPic001.Label12.Visible = False
      'Add by Amy 2018/07/19
      frmPic001.cmdok(2).Enabled = False
      frmPic001.SetSeekCmdok
      'end 2018/07/19
      frmPic001.Show vbModal
   End Select
End Sub

Private Sub Combo1_Click()
   If Combo1.Tag <> Combo1 Then
      Command5.Value = True
      Combo1.Tag = Combo1
   End If
End Sub


Private Sub Command1_Click(Index As Integer)
   Select Case Index
   Case 0
      Unload Me
   Case 1, 2, 3, 4
      cmdState = Index
      PubShowNextData
      Exit Sub
   End Select
End Sub

Private Sub Command2_Click()
   Dim iRow As Integer
   
   With MSHFlexGrid1
      .Visible = False
            
      lSelRows = 0
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, 0) = "" Then
            .TextMatrix(iRow, 0) = "V"
         End If
         lSelRows = lSelRows + 1
      Next
      
      lblCount = lSelRows & " / " & lTotRows
      .Visible = True
   End With
End Sub

Private Sub Form_Activate()
   Static bDone As Boolean
   If bDone = False Then
      Combo1_Click
      bDone = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetCombo1
End Sub

Private Sub Command5_Click()
   SetMouseBusy
   QueryData
   SetMouseReady
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   If Pub_StrUserSt03 = "P13" Then
      arrGridHeadWidth = Array(240, 1400, 800, 1200, 640, 640, 1200, 1000, 0, 650, 650, 1900)
   Else
      arrGridHeadWidth = Array(240, 1400, 800, 1200, 640, 640, 1200, 1000, 650, 0, 650, 1900)
   End If
   iUbound = UBound(arrGridHeadWidth)

   With MSHFlexGrid1
   If pReset = True Then
      .Clear
      .Rows = 2

      iPrevRow = 0
      lTotRows = 0
      lSelRows = 0
      lblCount = lSelRows & " / " & lTotRows
   End If
   .FixedCols = 2
   .FormatString = "V|本所案號|發文日|案件名稱|國家|種類|案件性質|相關號性質|承辦人|繪圖人|智權人|備註"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub

Private Sub QueryData()
   Dim iRow As Integer
   Dim stCon As String, stCon2 As String, stCon1101 As String
   'Dim idx As Integer, idxCP09 As Integer

   SetGrid True
   
   'Added by Morgan 2014/12/31 +繪圖發後補看
   If Pub_StrUserSt03 = "P13" Then
      If Trim(Left(Combo1.Text, 6)) <> "" Then
         stCon = " and lp22= '" & Trim(Left("" & Combo1.Text, 6)) & "'"
      Else
         stCon = " and lp22 is not null"
      End If
      stCon = stCon & " and lp23=0"
      
      strExc(0) = "select '' V,pa01||'-'||pa02||'-'||pa03||'-'||pa04 本所案號,sqldatet(c1.cp27) 發文日" & _
         ",pa05 案件名稱,na03 國家,Decode(PA09,'000',PTM03,PTM04) 種類,Decode(PA09,'000',m1.cpm03,m1.cpm04) 案件性質" & _
         ",Decode(PA09,'000',m2.cpm03,m2.cpm04) 相關號性質" & _
         ",s1.st02 承辦人,s3.st02 繪圖人,s2.st02 智權人,'' 備註,c1.cp27,c1.cp09,c1.cp10,c1.cp43,pa01,pa02,pa03,pa04" & _
         " From letterprogress l1,caseprogress c1,caseprogress c2, patent, casepropertymap m1, casepropertymap m2,nation,PatentTradeMarkMap,staff s1,staff s2,staff s3" & _
         " where lp10='Y' and (lp15='Y' or lp11='0' or lp11='2')" & stCon & " and c1.cp09(+)=lp01 and c2.cp09(+)=c1.cp43" & _
         " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04" & _
         " and m1.cpm01(+)=c1.cp01 and m1.cpm02(+)=c1.cp10 and m2.cpm01(+)=c2.cp01 and m2.cpm02(+)=c2.cp10 and na01(+)=pa09 And '1'=PTM01(+) AND PA08=PTM02(+)" & _
         " and s1.st01(+)=c1.cp14 and s2.st01(+)=c1.cp13 and s3.st01(+)=nvl(c1.cp29,c2.cp29)" & _
         " order by c1.cp27,c1.cp09"
      
   Else
   'end 2014/12/31
      
      If Trim(Left(Combo1.Text, 6)) <> "" Then
         stCon = " and lp20='" & Trim(Left("" & Combo1.Text, 6)) & "'"
      Else
         stCon = " and lp20 is not null"
      End If
      stCon = stCon & " and lp21=0"
      
      'Add By Sindy 2020/12/7 內商客戶函知
      If m_ProState = "T" Then
         strExc(0) = "select '' V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,sqldatet(cp27) 發文日" & _
            ",tm05 案件名稱,na03 國家,Decode(tm10,'000',PTM03,PTM04) 種類,cpm03 案件性質,'' 相關號性質,s1.st02 承辦人,s3.st02 繪圖人,s2.st02 智權人,'' 備註" & _
            ",cp27,cp09,cp10,cp43,tm01 as pa01,tm02 as pa02,tm03 as pa03,tm04 as pa04" & _
            " From letterprogress l1,caseprogress c1, trademark, casepropertymap,nation,PatentTradeMarkMap,staff s1,staff s2,staff s3" & _
            " where lp10='Y' and (lp15='Y' or lp11='0' or lp11='2') and lp01<'C'" & stCon & " and cp09(+)=lp01" & _
            " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=tm10 And '2'=PTM01(+) AND tm08=PTM02(+)" & _
            " and s1.st01(+)=cp14 and s2.st01(+)=cp13 and s3.st01(+)=cp29"
         strExc(0) = strExc(0) & " union all select '' V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,sqldatet(cp27) 發文日" & _
            ",sp05 案件名稱,na03 國家,' ' 種類,cpm03 案件性質,'' 相關號性質,s1.st02 承辦人,s3.st02 繪圖人,s2.st02 智權人,'' 備註" & _
            ",cp27,cp09,cp10,cp43,sp01 as pa01,sp02 as pa02,sp03 as pa03,sp04 as pa04" & _
            " From letterprogress l1,caseprogress c1, servicepractice, casepropertymap,nation,staff s1,staff s2,staff s3" & _
            " where lp10='Y' and (lp15='Y' or lp11='0' or lp11='2') and lp01<'C'" & stCon & " and cp09(+)=lp01" & _
            " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=sp09" & _
            " and s1.st01(+)=cp14 and s2.st01(+)=cp13 and s3.st01(+)=cp29"
         strExc(0) = strExc(0) & " order by cp27,cp09"
      Else
      '2020/12/7 END
      
         'Added by Morgan 2015/9/7 只抓工程師承辦--游經理
         'Modified by Morgan 2018/10/12 +CFP的發後補看
         '1.第1次有客戶函的通知申請案號(1101)。
         '2.所有工程師承辦有客戶函的進度。
         '3.有客戶函的專利權公告作廢(1606)、消滅(1604)、公告異議期滿通知(1223)、形式審查合格通知(1213)。(程序承辦)
         '4.詢問代理人(957)。(無客戶函)
         'stCon = stCon & " and ((substr(s1.st03,1,2)='P1' and s1.ST03<>'P12' and s1.ST03<>'P13')"
         stCon2 = " and ((substr(s1.st03,1,2)='P1' and s1.ST03<>'P12' and s1.ST03<>'P13') or (c1.cp01='CFP' and c1.cp10 in('1606','1604','1223','1213','1101'))) "
      
         'Modified by Morgan 2015/4/13 +親送,不寄
         'Modified by Morgan 2018/10/12 +相關號性質
         'Modify By Sindy 2020/12/7 + and pa01 is not null
         strExc(0) = "select '' V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,sqldatet(cp27) 發文日" & _
            ",pa05 案件名稱,na03 國家,Decode(PA09,'000',PTM03,PTM04) 種類,cpm03 案件性質,'' 相關號性質,s1.st02 承辦人,s3.st02 繪圖人,s2.st02 智權人,'' 備註" & _
            ",cp27,cp09,cp10,cp43,pa01,pa02,pa03,pa04" & _
            " From letterprogress l1,caseprogress c1, patent, casepropertymap,nation,PatentTradeMarkMap,staff s1,staff s2,staff s3" & _
            " where lp10='Y' and (lp15='Y' or lp11='0' or lp11='2') and lp01<'C'" & stCon & stCon2 & " and cp09(+)=lp01" & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=pa09 And '1'=PTM01(+) AND PA08=PTM02(+)" & _
            " and s1.st01(+)=cp14 and s2.st01(+)=cp13 and s3.st01(+)=cp29"
            
         'Modified by Morgan 2018/10/12 +相關號性質,工程師改抓C類,通知申請號改單獨過濾(下面)
         'Modify By Sindy 2020/12/7 + and pa01 is not null
         strExc(0) = strExc(0) & " union all select '' V,pa01||'-'||pa02||'-'||pa03||'-'||pa04 本所案號,sqldatet(c1.cp27) 發文日" & _
            ",pa05 案件名稱,na03 國家,Decode(PA09,'000',PTM03,PTM04) 種類,Decode(PA09,'000',m1.cpm03,m1.cpm04) 案件性質" & _
            ",Decode(PA09,'000',m2.cpm03,m2.cpm04) 相關號性質" & _
            ",s1.st02 承辦人,s3.st02 繪圖人,s2.st02 智權人,'' 備註,c1.cp27,c1.cp09,c1.cp10,c1.cp43,pa01,pa02,pa03,pa04" & _
            " From letterprogress l1,caseprogress c1,caseprogress c2, patent, casepropertymap m1, casepropertymap m2,nation,PatentTradeMarkMap,staff s1,staff s2,staff s3" & _
            " where lp10='Y' and (lp15='Y' or lp11='0' or lp11='2') and lp01>'C' and c1.cp09(+)=lp01 and c1.cp10<>'1101' and c2.cp09(+)=c1.cp43" & stCon & stCon2 & _
            " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04 and pa01 is not null" & _
            " and m1.cpm01(+)=c1.cp01 and m1.cpm02(+)=c1.cp10 and m2.cpm01(+)=c2.cp01 and m2.cpm02(+)=c2.cp10 and na01(+)=pa09 And '1'=PTM01(+) AND PA08=PTM02(+)" & _
            " and s1.st01(+)=c1.cp14 and s2.st01(+)=c1.cp13 and s3.st01(+)=c1.cp29"
            
         'Added by Morgan 2018/10/12
         '通知申請號(第1次有通知客戶函的)
         'Modified by Morgan 2018/11/19 P案不用看--游經理
         'Modify By Sindy 2020/12/7 + and pa01 is not null
         strExc(0) = strExc(0) & " union all select '' V,pa01||'-'||pa02||'-'||pa03||'-'||pa04 本所案號,sqldatet(c1.cp27) 發文日" & _
            ",pa05 案件名稱,na03 國家,Decode(PA09,'000',PTM03,PTM04) 種類,Decode(PA09,'000',m1.cpm03,m1.cpm04) 案件性質" & _
            ",Decode(PA09,'000',m2.cpm03,m2.cpm04) 相關號性質" & _
            ",s1.st02 承辦人,s3.st02 繪圖人,s2.st02 智權人,'' 備註,c1.cp27,c1.cp09,c1.cp10,c1.cp43,pa01,pa02,pa03,pa04" & _
            " From letterprogress l1,caseprogress c1,caseprogress c2, patent, casepropertymap m1, casepropertymap m2,nation,PatentTradeMarkMap,staff s1,staff s2,staff s3" & _
            " where lp10='Y' and (lp15='Y' or lp11='0' or lp11='2') and lp01>'C' and c1.cp01='CFP' and c1.cp09(+)=lp01 and c1.cp10='1101' and c2.cp09(+)=c1.cp43" & stCon & stCon2 & _
            " and not exists(select * From caseprogress c3,letterprogress l2 where c3.cp43=c1.cp43 and c3.cp10=c1.cp10 and c3.cp05<c1.cp05 and l2.lp01(+)=c3.cp09 and l2.lp10='Y')" & _
            " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04 and pa01 is not null" & _
            " and m1.cpm01(+)=c1.cp01 and m1.cpm02(+)=c1.cp10 and m2.cpm01(+)=c2.cp01 and m2.cpm02(+)=c2.cp10 and na01(+)=pa09 And '1'=PTM01(+) AND PA08=PTM02(+)" & _
            " and s1.st01(+)=c2.cp14 and s2.st01(+)=c2.cp13 and s3.st01(+)=c2.cp29"
            
         'CFP詢問代理人(957)
         'Modify By Sindy 2020/12/7 + and pa01 is not null
         strExc(0) = strExc(0) & " union all select '' V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,sqldatet(cp27) 發文日" & _
            ",pa05 案件名稱,na03 國家,Decode(PA09,'000',PTM03,PTM04) 種類,cpm03 案件性質,'' 相關號性質,s1.st02 承辦人,s3.st02 繪圖人,s2.st02 智權人,'' 備註" & _
            ",cp27,cp09,cp10,cp43,pa01,pa02,pa03,pa04" & _
            " From letterprogress l1,caseprogress c1, patent, casepropertymap,nation,PatentTradeMarkMap,staff s1,staff s2,staff s3" & _
            " where lp01<'C'" & stCon & " and cp09(+)=lp01 and cp01='CFP' and cp10='957' and cp27>19221111" & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=pa09 And '1'=PTM01(+) AND PA08=PTM02(+)" & _
            " and s1.st01(+)=cp14 and s2.st01(+)=cp13 and s3.st01(+)=cp29"
         'end 2018/10/12
         strExc(0) = strExc(0) & " order by cp27,cp09"
      End If
   End If
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid1
      .Visible = False
      .FixedCols = 0
      Set .Recordset = RsTemp
      lTotRows = RsTemp.RecordCount
      lblCount = lSelRows & " / " & lTotRows
      
      SetGrid
      .col = 1: .row = 1
      SelectRow 1
      .Visible = True

      '筆數多時速度有點慢,先不抓相關收文號性質
      'idx = GetFieldId("案件性質", MSHFlexGrid1)
      'idxCP09 = GetFieldId("CP09", MSHFlexGrid1)
      'For iRow = 1 To .Rows - 1
      '   .TextMatrix(iRow, idx) = .TextMatrix(iRow, idx) & PUB_GetRelateCasePropertyName(.TextMatrix(iRow, idxCP09), "1")
      'Next
      End With
   Else
      MsgBox "無待補看資料！", vbExclamation
   End If
End Sub

Private Function GetFieldId(pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As Integer
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetFieldId = iRow
         Exit For
      End If
   Next
   End With
End Function

Private Function GetValue(pRow As Integer, pFieldName As String) As String
   Dim iRow As Integer
   With MSHFlexGrid1
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, iRow)
         Exit For
      End If
   Next
   End With
End Function

Private Function SetValue(pRow As Integer, pFieldName As String, pValue As String) As Boolean
   Dim iRow As Integer
   With MSHFlexGrid1
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         .TextMatrix(pRow, iRow) = pValue
         SetValue = True
         Exit Function
      End If
   Next
   End With
End Function

Private Sub Process()
   Dim iRow As Integer, idxCP09 As Integer, idxMemo As Integer

On Error GoTo ErrHnd

   idxCP09 = GetFieldId("cp09", MSHFlexGrid1)
   idxMemo = GetFieldId("備註", MSHFlexGrid1)
   With MSHFlexGrid1
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) = "V" Then
         If UpdateDB(.TextMatrix(iRow, idxCP09), .TextMatrix(iRow, idxMemo)) = True Then
            If iRow = iPrevRow Then SelectRow 0
            .TextMatrix(iRow, 0) = "X"
            .RowHeight(iRow) = 0
            lSelRows = lSelRows - 1
            lTotRows = lTotRows - 1
            lblCount = lSelRows & " / " & lTotRows
            DoEvents
         Else
            Exit For
         End If
      End If
   Next
   End With
   Exit Sub

ErrHnd:
   MsgBox Err.Description, vbCritical
End Sub

Private Function UpdateDB(pKey As String, pMemo As String) As Boolean
   cnnConnection.BeginTrans
   
On Error GoTo ErrHndT
    'Modified by Morgan 2014/12/31+繪圖發後補看
    If Pub_StrUserSt03 = "P13" Then
       strSql = "update letterprogress set " & "lp22='" & strUserNum & "',lp23=" & strSrvDate(1) & ",lp25='" & ChgSQL(pMemo) & "' where lp01='" & pKey & "' and lp23=0"
    Else
       strSql = "update letterprogress set " & "lp20='" & strUserNum & "',lp21=" & strSrvDate(1) & ",lp24='" & ChgSQL(pMemo) & "' where lp01='" & pKey & "' and lp21=0"
    End If
    'end 2014/12/31
    cnnConnection.Execute strSql, intI
    cnnConnection.CommitTrans
    UpdateDB = True
    
    Exit Function
    
ErrHndT:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

Private Sub SelectRow(pRow As Integer)
   Dim nCol As Integer, iCol As Integer
   With MSHFlexGrid1
   nCol = .col
   If iPrevRow > 0 Then
      If iPrevRow <> pRow Then
         .row = iPrevRow
         If .FixedCols > 0 Then
            .col = .FixedCols - 1
            .CellBackColor = .BackColorFixed
            .CellForeColor = .ForeColor
         End If
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            .CellBackColor = .BackColor
         Next
      End If
   End If
   If pRow > 0 Then
      .row = pRow
      If .FixedCols > 0 Then
         .col = .FixedCols - 1
         .CellBackColor = .BackColorSel
         .CellForeColor = .ForeColorSel
      End If
      For iCol = .FixedCols To .Cols - 1
        .col = iCol
        .CellBackColor = &HFFC0C0
      Next
   End If
   .col = nCol
   iPrevRow = pRow
   End With
End Sub

Private Sub SetCombo1()
   Combo1.Clear
   If Pub_StrUserSt03 = "M51" Then
      Combo1.AddItem "      " & "全部"
   End If
   Combo1.AddItem strUserNum & " " & strUserName
   
   'Add By Sindy 2021/11/11 內商程序客戶函發後補看人員掛林律師,但江協理要可以進去補看作業操作
   If strUserNum = "98020" Then
      Combo1.AddItem "98003" & " " & GetPrjSalesNM("98003")
   End If
   '2021/11/11 END
   
   '檢查當時是否需要為他人職代
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040117 = Nothing
End Sub

Private Sub MSHFlexGrid1_DblClick()
   If MSHFlexGrid1.MouseRow > 0 Then
      Command1_Click 2
   End If
End Sub
'
Private Sub MSHFlexGrid1_Click()

   Dim nCol As Integer, nRow As Integer, iRow As Integer
   Dim stValue As String
   Dim stCP09 As String

   With MSHFlexGrid1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow = 0 Then
   
'Removed by Morgan 2015/4/13 取消全選--游經理
'      If nCol = 0 Then
'         For iRow = 1 To .Rows - 1
'            If .TextMatrix(iRow, 0) = "" Then
'               stValue = "V"
'               Exit For
'            '已刪除資料標示為 X
'            ElseIf .TextMatrix(iRow, 0) = "V" Then
'               stValue = ""
'               Exit For
'            End If
'         Next
'
'         lSelRows = 0
'         For iRow = 1 To .Rows - 1
'            If .TextMatrix(iRow, 0) <> "X" Then
'               If .TextMatrix(iRow, 0) <> stValue Then
'                  .TextMatrix(iRow, 0) = stValue
'               End If
'            End If
'            If .TextMatrix(iRow, 0) = "V" Then
'               lSelRows = lSelRows + 1
'            End If
'         Next
'
'         lblCount = lSelRows & " / " & lTotRows
'      Else

         '紀錄前次點選的收文號
         If iPrevRow > 0 Then
            stCP09 = GetValue(iPrevRow, "cp09")
         End If

         .col = nCol
         If m_blnColOrderAsc = False Then '字串降冪
            .Sort = 5 '字串昇冪
            m_blnColOrderAsc = True
         Else
            .Sort = 6 '字串降冪
            m_blnColOrderAsc = False
         End If

         '重設排序後前次點選的位置
         If iPrevRow > 0 Then
            For iRow = 1 To .Rows - 1
               If stCP09 = GetValue(iRow, "cp09") Then
                  iPrevRow = iRow
                  Exit For
               End If
            Next
         End If
         
'      End If'Removed by Morgan 2015/4/13 取消全選--游經理

   ElseIf nRow > 0 And .TextMatrix(nRow, 1) <> "" Then
      .row = nRow
      .col = nCol
      If nCol = 0 Then
         ClickGrid MSHFlexGrid1
      End If
      SelectRow nRow
      
      '有確認的點選備註欄可輸入
      If .TextMatrix(.row, 0) = "V" Then
         SetBox MSHFlexGrid1
      End If
   End If

   .Visible = True
   End With
End Sub

Private Sub SetMouseBusy()
   Screen.MousePointer = vbHourglass
   MSHFlexGrid1.MousePointer = vbHourglass
End Sub

Private Sub SetMouseReady()
   Screen.MousePointer = vbDefault
   MSHFlexGrid1.MousePointer = vbDefault
End Sub

Private Sub ClickGrid(FlexGrid As MSHFlexGrid)
   With FlexGrid
   If .Text = "V" Then
      lSelRows = lSelRows - 1
      .Text = ""
   '已刪除資料標示為 X
   ElseIf .Text = "" Then
      lSelRows = lSelRows + 1
      .Text = "V"
   End If
   lblCount = lSelRows & " / " & lTotRows
   End With
End Sub

Private Sub SetBox(ByRef FlexGrid As MSHFlexGrid)
   Dim lngLeft As Long, lngTop As Long, iCol As Integer, ii As Integer

   iCol = GetFieldId("備註", FlexGrid)
   With FlexGrid
      If .col = iCol Then
         txtInput.FontName = .CellFontName
         txtInput.FontSize = .CellFontSize
         'Modify by Lydia 2022/02/18 Form2.0 無Alignment屬性
         'txtInput.Alignment = .CellAlignment \ 5
         txtInput.TextAlign = 1
         txtInput.Text = .TextMatrix(.row, .col)
         txtInput.Tag = txtInput.Text
         'Modify by Lydia 2022/02/18 Form2.0 txtInput 圖層會在最下方,故加frame1
         Frame1.Width = .ColWidth(.col)
         Frame1.Height = .RowHeight(.row)
         'end 2022/02/18
         txtInput.Width = .ColWidth(.col)
         txtInput.Height = .RowHeight(.row)
         txtInput.Tag = txtInput.Text
         lngLeft = .Left + 25
         lngTop = .Top + .RowHeight(0) + 25
         lngLeft = lngLeft + .ColPos(.col)
         For ii = .TopRow To .row - 1
            lngTop = lngTop + .RowHeight(ii)
         Next
         'Modify by Lydia 2022/02/18 Form2.0 txtInput 圖層會在最下方,故加frame1
         'txtInput.Left = lngLeft: txtInput.Top = lngTop
         Frame1.Left = lngLeft: Frame1.Top = lngTop - 20
         
         Frame1.Visible = True 'Add by Lydia 2022/02/18 Form2.0 txtInput 圖層會在最下方,故加frame1
         txtInput.Visible = True
         txtInput.SetFocus
         
         m_InputRow = .row
         m_InputCol = .col
      End If
   End With
End Sub

Private Sub MSHFlexGrid1_Scroll()
   'Modify by Lydia 2022/02/18 Form2.0 txtInput 圖層會在最下方,故加frame1
   'If txtInput.Visible = True Then
   If Frame1.Visible = True Then
      MSHFlexGrid1.TextMatrix(m_InputRow, m_InputCol) = txtInput.Text
      'Modify by Lydia 2022/02/18 Form2.0 txtInput 圖層會在最下方,故加frame1
      'txtInput.Visible = False
      Frame1.Visible = False
   End If
End Sub

Private Sub txtInput_Change()
   txtInput = PUB_StrToStr(txtInput, txtInput.MaxLength)
End Sub

Private Sub txtInput_GotFocus()
   TextInverse txtInput
   OpenIme
End Sub

'Mark by Lydia 2022/02/18 按Enter字會消失
'Private Sub txtInput_KeyPress(KeyAscii As Integer)
'   Dim iCol As Integer, iRow As Integer
'
'   If KeyAscii = vbKeyReturn Then
'      MSHFlexGrid1.TextMatrix(m_InputRow, m_InputCol) = txtInput.Text
'      txtInput.Visible = False
'   ElseIf KeyAscii = vbKeyEscape Then
'      txtInput = txtInput.Tag
'      TextInverse txtInput
'   End If
'End Sub
'end 2022/02/18

'Add by Lydia 2022/02/18 從KeyPress搬過來修改
Private Sub txtInput_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim Cancel  As Boolean
   If KeyCode = vbKeyReturn Then
      Cancel = False
      txtInputValidate Cancel
      If Cancel = False Then
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, MSHFlexGrid1.col) = txtInput.Text
         MSHFlexGrid1.SetFocus
         MSHFlexGrid1.Refresh
         Frame1.Visible = False
      End If
   ElseIf KeyCode = vbKeyEscape Then
      MSHFlexGrid1.SetFocus
      Frame1.Visible = False
   End If
End Sub

Private Sub txtInput_LostFocus()
   'Modify by Lydia 2022/02/18 Form2.0 txtInput 圖層會在最下方,故加frame1
   'If txtInput.Visible = True Then
   If Frame1.Visible = True Then
      MSHFlexGrid1.TextMatrix(m_InputRow, m_InputCol) = txtInput.Text
      'Modify by Lydia 2022/02/18 Form2.0 txtInput 圖層會在最下方,故加frame1
      'txtInput.Visible = False
      Frame1.Visible = True
   End If
End Sub

'Added by Lydia 2022/02/18
Private Sub txtInputValidate(Cancel As Boolean)
Cancel = False
If CheckLengthIsOK(txtInput.Text, txtInput.MaxLength) = False Then
    txtInput.SetFocus
    txtInput_GotFocus
    Cancel = True
End If

'檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    txtInput.SetFocus
    txtInput_GotFocus
    Cancel = True
End If

End Sub
