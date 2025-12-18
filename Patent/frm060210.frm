VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060210 
   BorderStyle     =   1  '單線固定
   Caption         =   "程序大項工作期限通知"
   ClientHeight    =   5352
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9384
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5352
   ScaleWidth      =   9384
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1230
      TabIndex        =   11
      Text            =   "Combo2"
      Top             =   510
      Width           =   2265
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060210.frx":0000
      Left            =   6120
      List            =   "frm060210.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   9
      Top             =   510
      Width           =   3195
   End
   Begin VB.ComboBox cboPrinter 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   915
      TabIndex        =   7
      Top             =   5025
      Width           =   3675
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   400
      Left            =   4455
      TabIndex        =   6
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   400
      Index           =   3
      Left            =   7705
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   2
      Left            =   6155
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   60
      Width           =   1500
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5305
      TabIndex        =   2
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8550
      TabIndex        =   0
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4095
      Left            =   45
      TabIndex        =   1
      Top             =   870
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   7218
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V |E化情形 |核准函發文日 |承辦期限 |定稿日 |管制人 |承辦人 |本所案號 |備註"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin MSForms.ComboBox Combo3 
      Height          =   315
      Left            =   1230
      TabIndex        =   12
      Top             =   120
      Width           =   2265
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3995;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   14
      Top             =   150
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "工作項目："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   150
      TabIndex        =   13
      Top             =   510
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "顏色符號說明："
      Height          =   180
      Left            =   4815
      TabIndex        =   10
      Top             =   570
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   5055
      Width           =   975
   End
   Begin VB.Label lblCnt 
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7620
      TabIndex        =   5
      Top             =   5160
      Width           =   1710
   End
End
Attribute VB_Name = "frm060210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改 (Printer列印未改)
'Create by Sindy 2017/1/12
''Memo by Lydia 2019/05/31 原本「告准未發文期限通知」，更名為「程序大項工作期限通知」
Option Explicit

Dim bolBarShow As Boolean
Public cmdState As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim iPage As Integer, iPrint As Integer, PLeft() As Integer
Dim m_iCols As Integer, m_iPrtCols As Integer
Private Const ciFontSize = 12, ciTitleFontSize = 22
Private Const ciStartX = 500, ciColGap = 250, ciStartY = 500
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim strPrinter As String
Dim m_adoRst As ADODB.Recordset
Dim m_stSort As String '排序方式
'Added by Lydia 2019/05/31 要處理的工作大項種類(1.告准函1917、2.專利證書1603、3.公開公報1229、4.專利權消滅1604、5.通知年費逾期1605
'Modified by Lydia 2019/08/16 + 6.期限通知-年費
Dim iKind As String  '1~5 => 2019/08/16 1~6　'2025/06/18 改為1~7
Dim iPty(1 To 7) As String '案件性質 '2019/08/16 5=>6  '2025/06/18 改為6=>7
Dim colCaseNo As Integer '本所案號欄位置
Dim colCP48 As Integer '承辦期限欄位置


Private Sub SetRst2Grid()
   Set grdDataList.Recordset = m_adoRst
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
Dim i As Integer, StrTag As String, lngColor As Long, ii As Integer
Dim Str01 As String
   
On Error GoTo ErrorHandler
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 0
         grdDataList.Text = ""
         For ii = 0 To grdDataList.Cols - 1
            grdDataList.col = ii
            grdDataList.CellBackColor = grdDataList.BackColor
         Next
         
         'Modified by Lydia 2019/05/31 改成變數
         'StrTag = grdDataList.TextMatrix(i, 7)
         StrTag = grdDataList.TextMatrix(i, colCaseNo)
         If Left(Right(StrTag, 7), 1) = "-" Then
            StrTag = StrTag & "-0-00"
         ElseIf Left(Right(StrTag, 2), 1) = "-" Then
            StrTag = StrTag & "-00"
         End If
         
         If Left(StrTag, 1) < "A" Or Left(StrTag, 1) > "Z" Then
            StrTag = Right(StrTag, Len(StrTag) - 1)
         End If
         Str01 = SystemNumber(StrTag, 1)
         If fnSaveParentForm(Me) = False Then
            Exit For
         End If
         Me.Show
         Select Case cmdState
            Case 2 '案件基本資料
               Select Case Pub_RplStr(Str01)
                  Case "CFP", "FCP", "P"   '專利
                     frm100101_3.Show
                     frm100101_3.Tag = StrTag
                     frm100101_3.StrMenu
                     
                  Case "FG"
                     frm100101_B.Show
                     frm100101_3.Tag = StrTag
                     frm100101_B.StrMenu
               End Select
               
            Case 3 '案件進度
               frm100101_2.Show
               frm100101_2.Tag = StrTag
               frm100101_2.StrMenu
         End Select
         Exit For
      End If
   Next i
   
ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "(" & Err.Number & ")" & Err.Description
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub cmdPrint_Click()
   If grdDataList.Rows - 1 > 0 Then
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      PUB_RestorePrinter cboPrinter.Text
      DoPrint
      PUB_RestorePrinter strPrinter
      Me.Enabled = True
      Screen.MousePointer = vbDefault
   End If
End Sub

Public Sub cmdQuery_Click()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   doQuery
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   bolBarShow = Forms(0).StatusBar1.Visible
   Forms(0).StatusBar1.Visible = True
End Sub

Private Sub Form_Deactivate()
   Forms(0).StatusBar1.Visible = bolBarShow
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

'語法內有用組合欄位為條件以控制使用特定index(避掉某些不適當的)
Private Sub doQuery(Optional bolShow As Boolean = True)
Dim strConSql As String 'Added by Lydia 2019/05/31

   'Added by Lydia 2019/05/31 工作項目
   '1.畫面上增加"工作項目"、"承辦人員"的下拉選項，可以切換個人負責的工作，切換項目會自動啟動查詢。
   '2.工作項目:下拉選單"1.告准函、2.專利證書、3.公開公報、4.專利權消滅、5.通知年費逾期"。
   iKind = Left(Combo2.Text, 1)

   'Modified by Lydia 2019/05/31 避免查詢無資料,造成點選列判斷位置有誤
   'SetGrid
   Call SetGrid(True)
   
   'Added by Lydia 2019/05/31 指定承辦人
   If Combo3.Text <> "" Then
         strConSql = strConSql & " and c1.cp14='" & Trim(Left(Combo3, 6)) & "' "
   End If
   
    'Modified by Morgan 2017/8/17 ADODB.Recordset 用 Sort 方法排序時 O12 的函數的回傳值會超過 欄位長度上限,加 substr 限制長度
    'Modified by Lydia 2019/05/31 區分工作項目+配合其他案件性質
    'strExc(0) = "Select '' V,substr(GETEMAILFLAG(c1.cp09),1,2) E化情形,substr(sqldatet(c2.cp27),1,10) 核准函發文日,substr(sqldatet(c1.cp48),1,10) 承辦期限,substr(sqldatet(c1.cp85),1,10) 定稿日,S1.ST02 管制人" & _
                ",S2.ST02 承辦人,c1.CP01||'-'||c1.CP02||DECODE(c1.CP03||c1.CP04,'000','','-'||c1.CP03||'-'||c1.CP04) 本所案號" & _
                ",substr(c1.cp64,1,10) 備註" & _
                " from caseprogress c1,caseprogress c2,staff s1,staff s2,PATENT,FAGENT,Nation" & _
                " where c1.cp01='FCP' and c1.cp10='1917' and c1.cp27||c1.cp57 is null" & _
                " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04" & _
                " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
                " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & _
                " and na16=s1.st01(+) and c1.cp14=s2.st01(+)" & _
                " and c1.cp43=c2.cp09(+)" & _
                " order by GETEMAILFLAG(c1.cp09),pa01,pa02,pa03,pa04"
   Select Case iKind
        Case "1" '1.告准函1917
            'Modified by Lydia 2019/06/17 本所案號前加註銷卷＊/閉卷●
            strExc(0) = "Select '' V,substr(nvl(GETEMAILFLAG(c1.cp09),' '),1,1) E化情形,substr(sqldatet(c2.cp27),1,10) 核准函發文日,substr(sqldatet(c1.cp05),1,10) 來函收文日,substr(sqldatet(c1.cp48),1,10) 承辦期限,substr(sqldatet(c1.cp85),1,10) 定稿日,S1.ST02 管制人" & _
                        ",S2.ST02 承辦人,decode(pa57,null,'','＊')||decode(pa108,null,'','●')||c1.CP01||'-'||c1.CP02||DECODE(c1.CP03||c1.CP04,'000','','-'||c1.CP03||'-'||c1.CP04) 本所案號" & _
                        ",substr(c1.cp64,1,80) 備註,c1.cp09 C1總收文號" & _
                        " from caseprogress c1,caseprogress c2,staff s1,staff s2,PATENT,FAGENT,Nation" & _
                        " where c1.cp01='FCP' and c1.cp10='" & iPty(iKind) & "' and c1.cp158=0 and c1.cp159=0" & _
                        " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04" & _
                        " AND PA57||PA108 IS NULL AND PA75 IS NOT NULL" & _
                        " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & _
                        " and na16=s1.st01(+) and c1.cp14=s2.st01(+)" & _
                        " and c1.cp43=c2.cp09(+) " & strConSql & _
                        " order by substr(nvl(GETEMAILFLAG(c1.cp09),' '),1,1),pa01,pa02,pa03,pa04"
        Case Else '其他:2.專利證書1603、3.公開公報1229、4.專利權消滅1604、5.通知年費逾期1605
            'Added by Lydia 2019/08/16 區分通知期限
            If InStr(iPty(iKind), "-") > 0 Then
                strExc(1) = Mid(iPty(iKind), 1, InStr(iPty(iKind), "-") - 1) '
                '相關收文號的案件性質
                strExc(2) = Mid(iPty(iKind), InStr(iPty(iKind), "-") + 1)
                strExc(2) = " and np07='" & strExc(2) & "' "
                strExc(0) = "Select '' V,' ' E化情形,' ' 核准函發文日,substr(sqldatet(c1.cp05),1,10) 來函收文日,substr(sqldatet(c1.cp48),1,10) 承辦期限,substr(sqldatet(c1.cp85),1,10) 定稿日,S1.ST02 管制人" & _
                            ",S2.ST02 承辦人,decode(pa57,null,'','＊')||decode(pa108,null,'','●')||c1.CP01||'-'||c1.CP02||DECODE(c1.CP03||c1.CP04,'000','','-'||c1.CP03||'-'||c1.CP04) 本所案號" & _
                            ",substr(c1.cp64,1,80) 備註,c1.cp09 C1總收文號" & _
                            " from caseprogress c1,staff s1,staff s2,PATENT,FAGENT,Nation,nextprogress" & _
                            " where c1.cp01='FCP' and c1.cp10='" & strExc(1) & "' and c1.cp158=0 and c1.cp43=np01(+) and c1.cp30=np22(+)" & strExc(2) & _
                            " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04" & _
                            " AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & _
                            " and na16=s1.st01(+) and c1.cp14=s2.st01(+)" & strConSql & _
                            " order by c1.cp48,pa01,pa02,pa03,pa04"
            Else
            'end 2019/08/16
                'Modifeid by Lydia 2019/06/17 拿掉PA57||PA108 IS NULL ; 若是已上閉卷的案件，各項大批進度檔發文日請先上"111111"(目前年費逾繳已會先上111111)'
                                                                                                        '，若有特殊案件已閉卷尚需報告，則程序會個案下定稿，到時自動將發文日"111111"拿掉，如此案件又可自大批上發文。
                'Modified by Lydia 2019/06/17 本所案號前加註銷卷＊/閉卷●
                strExc(0) = "Select '' V,' ' E化情形,' ' 核准函發文日,substr(sqldatet(c1.cp05),1,10) 來函收文日,substr(sqldatet(c1.cp48),1,10) 承辦期限,substr(sqldatet(c1.cp85),1,10) 定稿日,S1.ST02 管制人" & _
                            ",S2.ST02 承辦人,decode(pa57,null,'','＊')||decode(pa108,null,'','●')||c1.CP01||'-'||c1.CP02||DECODE(c1.CP03||c1.CP04,'000','','-'||c1.CP03||'-'||c1.CP04) 本所案號" & _
                            ",substr(c1.cp64,1,80) 備註,c1.cp09 C1總收文號" & _
                            " from caseprogress c1,staff s1,staff s2,PATENT,FAGENT,Nation" & _
                            " where c1.cp01='FCP' and c1.cp10='" & iPty(iKind) & "' and c1.cp158=0 and c1.cp159=0" & _
                            " AND PA01(+)=c1.CP01 AND PA02(+)=c1.CP02 AND PA03(+)=c1.CP03 AND PA04(+)=c1.CP04" & _
                            " AND PA75 IS NOT NULL AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9) AND NA01(+)=FA10" & _
                            " and na16=s1.st01(+) and c1.cp14=s2.st01(+)" & strConSql & _
                            " order by c1.cp48,pa01,pa02,pa03,pa04"
            End If 'end 2019/08/16
   End Select
   'end 2019/05/31
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If RsTemp Is Nothing Then Exit Sub
   'Remove by Lydia 2019/05/31 後面有設來源
   'Set grdDataList.Recordset = RsTemp
   'RecordShow
   'end 2019/05/31
   If RsTemp.RecordCount = 0 Then
      'Remove by Lydia 2019/05/31 若無資料再設到Grid中,會造成點選位置計算錯誤( Morgan : 最近發現的問題)
      'Set m_adoRst = RsTemp.Clone
      'SetRst2Grid
      'end 2019/05/31
      If bolShow = True Then
         MsgBox "查無資料！", vbInformation
      End If
      lblCnt.Caption = "共 0 筆"
   Else
      Set m_adoRst = RsTemp.Clone
      m_stSort = "E化情形 asc,本所案號 asc"
      m_adoRst.Sort = m_stSort
      SetRst2Grid
      Call SetGrid 'Added by Lydia 2019/05/31
      SetColor
      m_blnColOrderAsc = True
   End If
End Sub

'Modified by Lydia 2019/05/31
'Private Sub SetGrid()
Private Sub SetGrid(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   'Modified by Lydia 2019/05/31 區分工作項目
   'arrGridHeadText = Array("V", "E化情形", "核准函發文日", "承辦期限", "定稿日", "管制人", "承辦人", "本所案號", "備註")
   'arrGridHeadWidth = Array(200, 750, 1250, 900, 900, 900, 900, 1200, 2000)
   arrGridHeadText = Array("V", "E化情形", "核准函發文日", "來函收文日", "承辦期限", "定 稿 日", "管制人", "承辦人", "本  所  案  號", "備　　註", "C1總收文號")
   If iKind = "1" Then '告准函
       arrGridHeadWidth = Array(240, 750, 1300, 0, 1000, 1000, 1000, 1000, 1500, 2000, 0)
   ElseIf iKind = "2" Then '專利證書(顯示定稿日)
       arrGridHeadWidth = Array(240, 0, 0, 1100, 1000, 1000, 0, 1000, 1500, 2000, 0)
   Else
       arrGridHeadWidth = Array(240, 0, 0, 1100, 1000, 0, 0, 1000, 1500, 2000, 0)
   End If
   'end 2019/05/31
   
   grdDataList.Visible = False
   grdDataList.Cols = UBound(arrGridHeadText) + 1
   'Modified by Lydia 2019/05/31 避免查詢無資料,造成點選列判斷位置有誤
   'GrdDataList.Rows = 2
   If pReset = True Then
         grdDataList.Clear
         grdDataList.Rows = 2
   End If
   'end 2019/05/31
   
   For iRow = 0 To grdDataList.Cols - 1
      grdDataList.row = 0
      grdDataList.col = iRow
      grdDataList.Text = arrGridHeadText(iRow)
      grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grdDataList.CellAlignment = flexAlignCenterCenter
   Next
   'Added by Lydia 2019/05/31 預設欄位置
   If m_iCols = 0 Then
      colCaseNo = PUB_MGridGetId("本  所  案  號", grdDataList)
      colCP48 = PUB_MGridGetId("承辦期限", grdDataList)
      m_iCols = UBound(arrGridHeadText)
   End If
   
   grdDataList.Visible = True
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   PUB_SetPrinter Me.Name, cboPrinter, strPrinter
   
   '畫面沒離開時沒寫Log會造成逾時重新登入後重複執行
   PUB_AddExcuteLog Me.Name
   
   Combo1.Clear
   Combo1.AddItem "紅色(*)：表示逾管控期限"
   Combo1.AddItem "綠色(v): 表示當日期限"
   Combo1.AddItem "藍色: 表示點選資料"
   Combo1.ListIndex = 0
   
   'Added by Lydia 2019/05/31
   Combo2.Clear
   Combo2.AddItem "1. 告准函"
   iPty(1) = "1917"
   Combo2.AddItem "2. 專利證書"
   iPty(2) = "1603"
   Combo2.AddItem "3. 公開公報"
   iPty(3) = "1229"
   Combo2.AddItem "4. 專利權消滅"
   iPty(4) = "1604"
   Combo2.AddItem "5. 通知年費逾期"
   iPty(5) = "1605"
   'Added by Lydia 2019/08/16
   Combo2.AddItem "6. 期限通知-年費"
   iPty(6) = "1913-605"
   'Added by Lydia 2025/06/18
   Combo2.AddItem "7. 期限通知-實體審查"
   iPty(7) = "1913-416"
   
   Call SetCombo3
   'end 2019/05/31
   
   Call doQuery(False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060210 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub

Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim iCol As Integer
   
   iCol = grdDataList.MouseCol
   If grdDataList.MouseRow < 1 Then
      grdDataList.Visible = False
      Set grdDataList.Recordset = Nothing
      If m_blnColOrderAsc = True Then
         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " desc"
         m_blnColOrderAsc = False
      Else
         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " asc"
         m_blnColOrderAsc = True
      End If
      SetRst2Grid
      SetColor
      grdDataList.Visible = True
   End If
End Sub

Private Sub grdDataList_SelChange()
Dim ii As Integer
   
   With grdDataList
      If .MouseRow > 0 Then
         .Visible = False
         .row = .MouseRow
         .col = 0
         If .Text = "V" Then
            .Text = ""
            For ii = 0 To .Cols - 1
               .col = ii
               .CellBackColor = .BackColor
            Next
         Else
            .Text = "V"
            For ii = 0 To .Cols - 1
               .col = ii
               .CellBackColor = &HFFC0C0
            Next
         End If
         .Visible = True
      End If
   End With
End Sub

Private Sub SetColor()
Dim lngToday As Long, lngCP48 As Long
Dim ii As Integer, jj As Integer, dblCnt As Double
   
   dblCnt = 0
   With grdDataList
   If .Rows > 1 Then
      .Visible = False
      lngToday = Val(strSrvDate(2))
      For ii = 1 To .Rows - 1
         .row = ii
         '固定欄位變回白色
         For jj = 0 To 2
            .col = jj
            .CellBackColor = &HFFFFFF
            '.CellAlignment = flexAlignRightTop
            '.CellFontSize = 9
         Next
         
         .RowHeight(ii) = 255
         'Modified by Lydia 2019/05/31 改成變數
         'lngCP48 = Val(Replace(.TextMatrix(ii, 3), "/", ""))
         lngCP48 = Val(Replace(.TextMatrix(ii, colCP48), "/", ""))
         
         '逾管控期限
         If lngCP48 < lngToday And lngCP48 > 0 Then
            'Modified by Lydia 2019/05/31 改成變數
            '.TextMatrix(ii, 7) = "*" & .TextMatrix(ii, 7)
            .TextMatrix(ii, colCaseNo) = "*" & .TextMatrix(ii, colCaseNo)
            For jj = 1 To .Cols - 1
               .col = jj
               '紅
               .CellBackColor = &HFF&
            Next
         '當日期限
         ElseIf lngCP48 = lngToday And lngCP48 > 0 Then
            'Modified by Lydia 2019/05/31 改成變數
            '.TextMatrix(ii, 7) = "v" & .TextMatrix(ii, 7)
            .TextMatrix(ii, colCaseNo) = "v" & .TextMatrix(ii, colCaseNo)
            For jj = 1 To .Cols - 1
               .col = jj
               '綠
               .CellBackColor = &HC000&
            Next
         End If
         
         If .RowHeight(ii) > 0 Then
            dblCnt = dblCnt + 1
         End If
      Next
      .TopRow = 1
      .Visible = True
   End If
   End With

   lblCnt.Caption = "共 " & dblCnt & " 筆"
End Sub

Private Sub DoPrint()
Dim iOrientation As Integer, iRow As Integer, iCol As Integer
Dim strTemp() As String
   
   iOrientation = Printer.Orientation
   If iKind = "1" Or iKind = "2" Then 'Added by Lydia 2019/05/31 區分工作項目
       Printer.Orientation = 2 '橫印: 告准函,專利證書
   'Added by Lydia 2019/05/31
   Else
       Printer.Orientation = 1 '直印
   End If
   'end 2019/05/31
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With grdDataList
      GetPleft
      m_iPrtCols = m_iCols
      ReDim strTemp(1 To m_iPrtCols)
      
      iPage = 1
      PrintPageHeader
      PrintPageHeader1
      For iRow = 1 To .Rows - 1
         For iCol = LBound(strTemp) To UBound(strTemp)
            strTemp(iCol) = .TextMatrix(iRow, iCol)
         Next
         PrintDetail strTemp
      Next
      Call PrintReportFooter(.Rows - 1)
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
End Sub

Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   'm_iCols = 8 'Mark by Lydia 2019/05/31 改在SetGrid設定
   ReDim PLeft(1 To m_iCols)
   PLeft(1) = ciStartX
   For intI = 2 To m_iCols
      If grdDataList.ColWidth(intI - 1) > 0 Then
         PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(grdDataList.TextMatrix(0, intI - 1)) + ciColGap
      Else
         PLeft(intI) = PLeft(intI - 1)
      End If
   Next
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      'Modified by Lydia 2019/05/31
      'Printer.Print String(130, "-")
      Call PrintLine
      
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
End Sub

Sub PrintDetail(strData() As String)
   Dim iCol As Integer
   Dim strTemp As String 'Added by Lydia 2019/06/17
   
   PrintNewLine
   For iCol = LBound(strData) To UBound(strData)
      If Me.grdDataList.ColWidth(iCol) > 0 Then
         Printer.CurrentX = PLeft(iCol)
         Printer.CurrentY = iPrint
         strTemp = PUB_StringFilter(strData(iCol)) 'Added by Lydia 2019/06/17 去除跳行符號
         'Added by Lydia 2019/05/31 備註限長度
         If iCol = 9 Then
             If iKind = "2" Then
                  'Modified by Lydia 2019/06/17 strData(iCol)=> strtemp
                 Printer.Print convForm(strTemp, 80)
             Else
                 Printer.Print convForm(strTemp, 46)
             End If
         Else
         'end 2019/05/31
              Printer.Print strTemp
         End If 'end 2019/05/31
      End If
   Next
End Sub

Sub PrintPageHeader()
Dim strPTmp As String
   
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = Replace(Me.Caption, "通知", "清單")
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   'Added by Lydia 2019/05/31
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "工作項目：" & Trim(Mid(Combo2.Text, 3))
   'end 2019/05/31
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
    'Modified by Lydia 2019/05/31
    'Printer.Print String(130, "-")
    Call PrintLine
End Sub

Sub PrintPageHeader1()
   Call PrintNewLine(False, 1)
   For intI = 1 To m_iPrtCols
     If Me.grdDataList.ColWidth(intI) > 0 Then
        Printer.CurrentX = PLeft(intI)
        Printer.CurrentY = iPrint
        Printer.Print grdDataList.TextMatrix(0, intI)
     End If
   Next
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   'Modified by Lydia 2019/05/31
   'Printer.Print String(130, "-")
   Call PrintLine
End Sub

'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)
   Call PrintNewLine(True, 1)
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   'Modified by Lydia 2019/05/31
   'Printer.Print String(130, "-")
   Call PrintLine
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "合計： " & iRecCount & " 筆"
   Printer.EndDoc
End Sub

'Added by Lydia 2019/05/31
Private Sub SetCombo3()
Dim strTmp As String

   Combo3.Clear
   strExc(0) = "select st01,st02 from staff a where st03='F22' and st04='1' order by 1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not RsTemp.EOF
         If .Fields("st01") = strUserNum Then
            Combo3.AddItem .Fields("st01") & " " & .Fields("st02"), 0
            strTmp = .Fields("st01") & " " & .Fields("st02")
         Else
            Combo3.AddItem .Fields("st01") & " " & .Fields("st02")
         End If
      .MoveNext
      Loop
      End With
   End If
   
   If strTmp <> "" Then
      Combo3.ListIndex = 0
   Else
      Combo3.ListIndex = Combo3.ListCount - 1
   End If
   
   '抓操作者有整批未發文的案件性質,預設工作項目種類
   strExc(1) = ""
   strExc(2) = ""
   For intI = 1 To UBound(iPty)
       'Modified by Lydia 2019/08/16 改成組合語法
       'strExc(1) = strExc(1) & ", '" & iPty(intI) & "' , '" & intI & "' "
       'strExc(2) = strExc(2) & "," & iPty(intI)
       If iPty(intI) <> "" Then
           If InStr(iPty(intI), "-") > 0 Then '通知期限
               strExc(1) = strExc(1) & "Union select '" & intI & "' as ord1, cp10,count(*) cnt " & _
                               "from caseprogress,nextprogress where cp01='FCP'  and cp05>=20190501 and cp158=0 and cp159=0 and cp14='" & Trim(Left(Combo3.Text, 6)) & "' and cp10='" & Mid(iPty(intI), 1, InStr(iPty(intI), "-") - 1) & "' " & _
                               "and cp43=np01(+) and cp30=np22(+) and np07='" & Mid(iPty(intI), InStr(iPty(intI), "-") + 1) & "' group by cp10 "
           Else
               strExc(1) = strExc(1) & "Union select '" & intI & "' as ord1, cp10,count(*) cnt " & _
                               "from caseprogress where cp01='FCP'  and cp05>=20190501 and cp158=0 and cp159=0 and cp14='" & Trim(Left(Combo3.Text, 6)) & "' and cp10='" & iPty(intI) & "' " & _
                               "group by cp10 "
           End If
       End If
       'end 2019/08/16
   Next intI
   'Modified by Lydia 2019/08/16
   'strExc(2) = GetAddStr(Mid(strExc(2), 2))
   'strExc(0) = "select decode(cp10 " & strExc(1) & ",'9' ) ord1,cp10 ,count(*) cnt " & _
   '                 "from caseprogress where cp01='FCP'  and cp05>=20190501 and cp158=0 and cp159=0 and cp14='" & Trim(Left(Combo3.Text, 6)) & "' and cp10 in (" & strExc(2) & ") " & _
   '                 "group by decode(cp10 " & strExc(1) & ",'9' ),cp10 "
   'strExc(0) = strExc(0) & " order by ord1 "
   strExc(0) = "select * from (" & Mid(strExc(1), 6) & ") where cnt > 0 order by ord1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       RsTemp.MoveFirst
       Combo2.ListIndex = Val(RsTemp.Fields("ord1")) - 1
   Else  '無=>預設告准函
       Combo2.ListIndex = 0
   End If
   
End Sub

Private Sub Combo2_Click()
     '切換人員先不啟動查詢
    If (Combo3.Tag <> "" And Combo3.Tag <> Combo3.Text) Or (Combo2.Tag <> "" And Combo2.Tag <> Combo2.Text) Then
        m_iCols = 0 'Added by Lydia 2019/08/16 重新本所案號的欄位值
        Call doQuery(True)
    End If
    Combo3.Tag = Combo3.Text
    Combo2.Tag = Combo2.Text
End Sub

'列印分隔線
Private Sub PrintLine()
   If iKind = "1" Or iKind = "2" Then '橫印: 告准函,專利證書
       Printer.Print String(130, "-")
   Else
       Printer.Print String(92, "-")
   End If
End Sub

