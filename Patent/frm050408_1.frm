VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050408_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "互惠代理人案件統計表"
   ClientHeight    =   5736
   ClientLeft      =   168
   ClientTop       =   960
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9312
   Begin VB.CommandButton CmdProcExcel 
      Caption         =   "產生Excel"
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   30
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印全部區間"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   5
      Left            =   1395
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   30
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印該國區間"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   4
      Left            =   90
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   30
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印該國(&P)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   3
      Left            =   2790
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   30
      Width           =   1110
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1170
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   780
      Width           =   2355
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   1
      Top             =   480
      Width           =   1290
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   2205
      MaxLength       =   8
      TabIndex        =   0
      Top             =   480
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印全部(&A)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   2
      Left            =   3915
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   30
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   8460
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   30
      Width           =   750
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   7335
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   30
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3855
      Left            =   45
      TabIndex        =   5
      Top             =   1110
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   6795
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblExtra 
      Height          =   225
      Left            =   5535
      TabIndex        =   14
      Top             =   840
      Width           =   3660
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "3. 點選給案量欄位會顯示該欄日期區間之案件及盈虧"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   585
      TabIndex        =   12
      Top             =   5490
      Width           =   4470
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "2. 點選代理人編號欄位會顯示該日期區間代理人之FCP請款/收款及CFP應付/付款金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   585
      TabIndex        =   11
      Top             =   5250
      Width           =   7170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "1. 游標指到會變黑色的欄位表示可點選顯示參考資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   585
      TabIndex        =   10
      Top             =   5010
      Width           =   4470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PS:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   270
      TabIndex        =   9
      Top             =   5010
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人國籍:"
      Height          =   180
      Left            =   135
      TabIndex        =   7
      Top             =   810
      Width           =   945
   End
   Begin VB.Line Line1 
      X1              =   3510
      X2              =   3690
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "代理人應收帳款統計區間:"
      Height          =   180
      Left            =   135
      TabIndex        =   6
      Top             =   510
      Width           =   2025
   End
End
Attribute VB_Name = "frm050408_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/11 改成Form2.0 ; grdDataList改字型=新細明體-ExtB; Printer列印未改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo by Morgan2010/12/28 申請案號欄已修改
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2008/3/4
Option Explicit

Dim m_iCol As Integer, m_iRow As Integer
Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 10
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 150
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Public m_adoRst As ADODB.Recordset
Public m_stDate1 As String
Public m_stDate2 As String
Dim m_bComboNoAct As Boolean
Dim m_bolByAgent As Boolean  '是否以代理人統計
Dim strCF As String 'Add By Sindy 2013/5/24
Dim strFC As String 'Add By Sindy 2013/5/24

Public Sub InitGrid()
   SetGrid
   SetCombo
End Sub

Private Sub SetGrid(Optional p_Country As String)
   Set grdDataList.Recordset = Nothing
   SetDataListWidth False
   If p_Country <> "" Then
      m_adoRst.Filter = " NA03='" & p_Country & "'"
   Else
      m_adoRst.Filter = 0
   End If
   Set grdDataList.Recordset = m_adoRst
   SetDataListWidth True
End Sub

'p_bolHeaderOnly:是否只設定表頭 true=是 false=資料一併清除
Public Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   'Memo by Lydia 2025/06/27 ＊＊若表單frm050408_1的欄位有變動，呼叫Pub_Frm050408_GetStatistic也要變動＊＊
   Dim ii As Integer
   With grdDataList
      .Visible = False
      If lblExtra <> "" Then
         .Cols = 21
      Else
         .Cols = 18
      End If
      If p_bolHeaderOnly = False Then
         .Clear
         .Rows = 2: .FixedRows = 1: .FixedCols = 0
      Else
         .FixedCols = 2
      End If
      .row = 0
      .RowHeight(0) = 850
      .RowHeightMin = 450
      ii = 0
      .col = ii: .ColWidth(.col) = 2115: .Text = "代理人名稱" '(0)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 1100: .Text = "代理人編號" '(1)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 690: .Text = "聯絡人資訊"  '(2)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 510: .Text = "國籍"        '(3)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 525: .Text = strFC & "總收案量" '"FCP總收案量"
      .col = ii: .ColWidth(.col) = 525: .Text = "全部" & strFC  '(4)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 525: .Text = strCF & "總給案量" '"CFP總給案量"
      .col = ii: .ColWidth(.col) = 525: .Text = "全部" & strCF '(5)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 680: .Text = "前年度" & strFC & "收案量" '"前年度FCP收案量"
      .col = ii: .ColWidth(.col) = 680: .Text = "前年度" & strFC '(6)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 680: .Text = "前年度" & strCF & "給案量" '"前年度CFP給案量"
      .col = ii: .ColWidth(.col) = 680: .Text = "前年度" & strCF '(7)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 680: .Text = "去年度" & strFC & "收案量" '"去年度FCP收案量"
      .col = ii: .ColWidth(.col) = 680: .Text = "去年度" & strFC '(8)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 680: .Text = "去年度" & strCF & "給案量" '"去年度CFP給案量"
      .col = ii: .ColWidth(.col) = 680: .Text = "去年度" & strCF '(9)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 700: .Text = "當年度1-6月" & strFC & "收案量" '"當年度1-6月FCP收案量"
      .col = ii: .ColWidth(.col) = 700: .Text = "當年度1-6月" & strFC '(10)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 700: .Text = "當年度1-6月" & strCF & "給案量" '"當年度1-6月CFP給案量"
      .col = ii: .ColWidth(.col) = 700: .Text = "當年度1-6月" & strCF '(11)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 700: .Text = "當年度1-6月收給案量差"
      .col = ii: .ColWidth(.col) = 700: .Text = "當年度1-6月" & strFC & "-" & strCF '(12)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 700: .Text = "當年度7-12月" & strFC & "收案量" '"當年度7-12月FCP收案量"
      .col = ii: .ColWidth(.col) = 700: .Text = "當年度7-12月" & strFC  '(13)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 700: .Text = "當年度7-12月" & strCF & "給案量" '"當年度7-12月CFP給案量"
      .col = ii: .ColWidth(.col) = 700: .Text = "當年度7-12月" & strCF  '(14)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 700: .Text = "當年度7-12月收給案量差"
      .col = ii: .ColWidth(.col) = 700: .Text = "當年度7-12月" & strFC & "-" & strCF '(15)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      '.col = ii: .ColWidth(.col) = 700: .Text = "該半年度建議" & strCF & "給案量" '"該半年度建議CFP給案量"
      .col = ii: .ColWidth(.col) = 700: .Text = "該半年度建議" & strCF '(16)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 510: .Text = "備註"    '(17)
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      If lblExtra <> "" Then
          ii = ii + 1
         'Modified by Lydia 2023/09/12 與Excel抬頭一致
         '.col = ii: .ColWidth(.col) = 700: .Text = "指定日期區間" & strFC & "收案量" '"指定日期區間FCP收案量"
         .col = ii: .ColWidth(.col) = 700: .Text = "指定日期區間" & strFC '(18)
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .ColAlignment(.col) = flexAlignRightCenter
         ii = ii + 1
         'Modified by Lydia 2023/09/12 與Excel抬頭一致
         '.col = ii: .ColWidth(.col) = 700: .Text = "指定日期區間" & strCF & "給案量" '"指定日期區間CFP給案量"
         .col = ii: .ColWidth(.col) = 700: .Text = "指定日期區間" & strCF '(19)
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .ColAlignment(.col) = flexAlignRightCenter
         ii = ii + 1
         'Modified by Lydia 2023/09/12 與Excel抬頭一致
         '.col = ii: .ColWidth(.col) = 700: .Text = "指定日期區間收給案量差"
         .col = ii: .ColWidth(.col) = 700: .Text = "指定日期區間" & strFC & "-" & strCF '(20)
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .ColAlignment(.col) = flexAlignRightCenter
      End If
      .Refresh
      .Visible = True
   End With
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim ii As Integer, jj As Integer, bOK As Boolean
      
   ClearQueryLog ("frm050408") 'Added by Lydia 2025/07/30
   
   Select Case Index
      Case 0
         frm050408.Show
         Unload Me
      Case 1
         Unload frm050408
         Unload Me
      'Add by Morgan 2008/6/10
      Case 2 '列印全部
         pub_QL05 = frm050408.m_strQL05 & ";列印全部" 'Added by Lydia 2025/07/30
         jj = Combo1.ListIndex
         For ii = 0 To Combo1.ListCount - 1
            SetGrid Combo1.List(ii)
            If DoPrint = True Then
               bOK = True
            End If
            
         Next
         SetGrid Combo1.List(jj)
         If bOK = True Then
            MsgBox "列印完成！"
         End If
         
      Case 3
         pub_QL05 = frm050408.m_strQL05 & ";列印該國:" & Combo1.Text 'Added by Lydia 2025/07/30
         If DoPrint = True Then
            MsgBox "列印完成！"
         End If
         
      'Add by Morgan 2008/6/24
      Case 5 '列印該國區間 'Memo by Lydia 2025/07/30 前一畫面有輸入「指定日期區間」才會出現
         pub_QL05 = frm050408.m_strQL05 & ";列印該國區間:" & Combo1.Text 'Added by Lydia 2025/07/30
         jj = Combo1.ListIndex
         For ii = 0 To Combo1.ListCount - 1
            SetGrid Combo1.List(ii)
            If DoPrint1 = True Then
               bOK = True
            End If
         Next
         SetGrid Combo1.List(jj)
         If bOK = True Then
            MsgBox "列印完成！"
         End If
         
      'Add by Morgan 2008/6/24
      Case 4 '列印全部區間  'Memo by Lydia 2025/07/30 前一畫面有輸入「指定日期區間」才會出現
         pub_QL05 = frm050408.m_strQL05 & ";列印全部區間:" & Combo1.Text 'Added by Lydia 2025/07/30
         If DoPrint1 = True Then
            MsgBox "列印完成！"
         End If
   End Select
   
End Sub

Private Sub Combo1_Click()
   If m_bComboNoAct = False Then
      SetGrid Combo1.Text
      m_iRow = 0: m_iCol = 0 'Added by Lydia 2025/02/05 切換國籍清空變數
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_bolByAgent = frm050408.m_bolByAgent
   'Add By Sindy 2013/5/24
   If frm050408.txtKind = "1" Then '專利
      strCF = "CFP"
      strFC = "FCP"
      Label5.Caption = "2. 點選代理人編號欄位會顯示該日期區間代理人之FCP請款/收款及CFP應付/付款金額"
   ElseIf frm050408.txtKind = "2" Then '商標
      strCF = "CFT"
      strFC = "FCT"
      Label5.Caption = "2. 點選代理人編號欄位會顯示該日期區間代理人之FCT請款/收款及CFT應付/付款金額"
   End If
   '2013/5/24 End
   
   'Added by Lydia 2025/06/06
   If frm050408.Tag <> "" Then
      Me.Caption = "互惠期間統計表"
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050408_1 = Nothing
End Sub

Private Sub GrdDataList_Click()
   Dim iRow As Integer, iCol As Integer
   With grdDataList
      iRow = .MouseRow
      iCol = .MouseCol
      If iRow > 0 Then
         Select Case iCol
            Case 4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 18, 19
               .Enabled = False
               Screen.MousePointer = vbHourglass 'Added by Lydia 2025/06/27
               GetStatistic iRow, iCol
               .Enabled = True
               Screen.MousePointer = vbDefault 'Added by Lydia 2025/06/27
            Case 1
               .Enabled = False
               Screen.MousePointer = vbHourglass 'Added by Lydia 2025/06/27
               GetStatistic1 iRow, iCol
               .Enabled = True
               Screen.MousePointer = vbDefault 'Added by Lydia 2025/06/27
         End Select
      End If
   End With
End Sub

'代理人應收帳款統計
Private Sub GetStatistic1(p_iRow As Integer, p_iCol As Integer)
   Dim stAgentNo As String, stDate1 As String, stDate2 As String, iPos As Integer
   Dim stVTB1 As String, stVTB2 As String, stVTB3 As String, stVTB4 As String, stVTB5 As String, stVTB6 As String
   Dim stVTB7 As String, stCol As String, dblNet As Double
   'Added by Lydia 2024/01/30
   Dim strCFcon As String, strFCcon As String
   If frm050408.txtKind = "1" Then '專利
       strFCcon = " and a1k13 in ('FCP','P') "
       '與統計表只抓母案不同,損益包含子案
       strCFcon = " and (substr(axf03,1,3)='CFP' or (substr(axf03,1,1)='P' and instr('0123456789',substr(axf03,2,1))> 0)) "
   ElseIf frm050408.txtKind = "2" Then '商標
       strFCcon = " and a1k13 in ('FCT','T') "
       '與統計表只抓母案不同,損益包含子案
       strCFcon = " and (substr(axf03,1,3)='CFT' or (substr(axf03,1,1)='T' and instr('0123456789',substr(axf03,2,1))> 0)) "
   End If
   'end 2024/01/30
   
   If txt1(0) = "" Then
      MsgBox "請輸入帳款統計起日！"
      txt1(0).SetFocus
   ElseIf ChkDate(txt1(0)) = False Then
      Exit Sub
      txt1(0).SetFocus
   End If
   
   If txt1(0) = "" Then
      MsgBox "請輸入帳款統計迄日！"
      txt1(1).SetFocus
   ElseIf ChkDate(txt1(1)) = False Then
      Exit Sub
      txt1(1).SetFocus
   End If
   
   stDate1 = TransDate(txt1(0), 1)
   stDate2 = TransDate(txt1(1), 1)
   
   stAgentNo = grdDataList.TextMatrix(p_iRow, 1)
   iPos = InStr(stAgentNo, "-")
   If iPos = 0 Then
      stAgentNo = Left(stAgentNo & "000", 9)
   Else
      stAgentNo = Left(Left(stAgentNo, iPos - 1) & "000", 9)
   End If
   
   'Added by Lydia 2025/07/30 查詢印表記錄檔欄位
   ClearQueryLog ("frm050408")
   pub_QL05 = frm050408.m_strQL05 & ";代理人應收帳款統計:" & stAgentNo
   'end 2025/07/30
   
   'FC請款金額,請款規費
   'Modify By Sindy 2012/10/12
'   stVTB1 = "select sum(a1k11 - nvl(a1k06, 0) * a1k10) A1,sum(a1k09) A2" & _
'      " from acc1k0 where a1k12 is null and a1k25 is null and a1k13='FCP'" & _
'      " and '" & stAgentNo & "' in (a1k03,a1k27,a1k28)" & _
'      " and a1k02>=" & stDate1 & " and a1k02<=" & stDate2
   'Modified by Lydia 2024/01/30 改用變數
   'stVTB1 = "select sum(a1k11 - nvl(a1k06, 0)) A1,sum(a1k09) A2" & _
      " from acc1k0 where a1k12 is null and a1k25 is null and a1k13='" & strFC & "'" & _
      " and '" & stAgentNo & "' in (a1k03,a1k27,a1k28)" & _
      " and a1k02>=" & stDate1 & " and a1k02<=" & stDate2
   stVTB1 = "select sum(a1k11 - nvl(a1k06, 0)) A1,sum(a1k09) A2" & _
      " from acc1k0 where a1k12 is null and a1k25 is null " & strFCcon & _
      " and '" & stAgentNo & "' in (a1k03,a1k27,a1k28)" & _
      " and a1k02>=" & stDate1 & " and a1k02<=" & stDate2
      
   'FC抵帳金額(算收款)
   'Modify By Sindy 2012/10/12
'   stVTB2 = "select sum((a1k08 - nvl(a1k06, 0)) * nvl(a1g02, 0)) B1" & _
'      " from acc1k0,acc1g0 where a1k12 is null and a1k25 is null and a1k13='FCP'" & _
'      " and '" & stAgentNo & "' in (a1k03,a1k27,a1k28)" & _
'      " and a1k02>=" & stDate1 & " and a1k02<=" & stDate2 & " and a1k17 is not null" & _
'      " and a1g01(+)=a1k17"
   'Modified by Lydia 2024/01/30 改用變數
   'stVTB2 = "select sum((a1k08 - nvl(a1k31, 0)) * nvl(a1g02, 0)) B1" & _
      " from acc1k0,acc1g0 where a1k12 is null and a1k25 is null and a1k13='" & strFC & "'" & _
      " and '" & stAgentNo & "' in (a1k03,a1k27,a1k28)" & _
      " and a1k02>=" & stDate1 & " and a1k02<=" & stDate2 & " and a1k17 is not null" & _
      " and a1g01(+)=a1k17"
   stVTB2 = "select sum((a1k08 - nvl(a1k31, 0)) * nvl(a1g02, 0)) B1" & _
      " from acc1k0,acc1g0 where a1k12 is null and a1k25 is null " & strFCcon & _
      " and '" & stAgentNo & "' in (a1k03,a1k27,a1k28)" & _
      " and a1k02>=" & stDate1 & " and a1k02<=" & stDate2 & " and a1k17 is not null" & _
      " and a1g01(+)=a1k17"
      
   'FC收款金額
   'Modified by Lydia 2024/01/30 改用變數
   'stVTB3 = "select sum(a0z04 * a0y04) C1" & _
      " from acc1k0,acc0z0,acc0y0 where a1k12 is null and a1k25 is null and a1k13='" & strFC & "'" & _
      " and '" & stAgentNo & "' in (a1k03,a1k27,a1k28)" & _
      " and a1k02>=" & stDate1 & " and a1k02<=" & stDate2 & " and a1k17 is null" & _
      " and a0z02(+)=a1k01 and a0y01(+)=a0z01"
   stVTB3 = "select sum(a0z04 * a0y04) C1" & _
      " from acc1k0,acc0z0,acc0y0 where a1k12 is null and a1k25 is null " & strFCcon & _
      " and '" & stAgentNo & "' in (a1k03,a1k27,a1k28)" & _
      " and a1k02>=" & stDate1 & " and a1k02<=" & stDate2 & " and a1k17 is null" & _
      " and a0z02(+)=a1k01 and a0y01(+)=a0z01"
      
   'CF應付金額(帳單)
   'Modified by Lydia 2024/01/30 改用變數
   'stVTB4 = "select sum(decode(A1G01,null,decode(a1901,null,AXF15,AXF04*A1906),AXF04*A1G03))  D1" & _
      " from acc150,acc151,acc190,acc1g0 where a1507 is null and a1503='" & stAgentNo & "'" & _
      " and a1502 >= " & stDate1 & " And a1502 <= " & stDate2 & _
      " and axf01(+)=a1501 and substr(axf03,1,3)='" & strCF & "'" & _
      " and A1902(+)=AXF01 AND A1G01(+)=A1512"
   stVTB4 = "select sum(decode(A1G01,null,decode(a1901,null,AXF15,AXF04*A1906),AXF04*A1G03))  D1" & _
      " from acc150,acc151,acc190,acc1g0 where a1507 is null and a1503='" & stAgentNo & "'" & _
      " and a1502 >= " & stDate1 & " And a1502 <= " & stDate2 & _
      " and axf01(+)=a1501 " & strCFcon & " and A1902(+)=AXF01 AND A1G01(+)=A1512"
      
   'CF抵帳金額(算付款)
   'Modified by Lydia 2024/01/30 改用變數
   'stVTB5 = "select sum(AXF04*A1G03) E1" & _
      " from acc150,acc151,acc1g0 where a1507 is null and a1503='" & stAgentNo & "'" & _
      " and a1502>=" & stDate1 & " and a1502<=" & stDate2 & " and a1512 is not null" & _
      " and axf01(+)=a1501 and substr(axf03,1,3)='" & strCF & "' and a1g01(+)=a1512"
   stVTB5 = "select sum(AXF04*A1G03) E1" & _
      " from acc150,acc151,acc1g0 where a1507 is null and a1503='" & stAgentNo & "'" & _
      " and a1502>=" & stDate1 & " and a1502<=" & stDate2 & " and a1512 is not null" & _
      " and axf01(+)=a1501 " & strCFcon & " and a1g01(+)=a1512"
      
   'CF付款金額
   'Modified by Lydia 2024/01/30 改用變數
   'stVTB6 = "select sum(AXF04*A1906) F1" & _
     " from acc150,acc151,acc190 where a1507 is null and a1503='" & stAgentNo & "'" & _
      " and a1502 >= " & stDate1 & " And a1502 <= " & stDate2 & _
      " and axf01(+)=a1501 and substr(axf03,1,3)='" & strCF & "'" & _
      " and A1902(+)=AXF01 AND A1901 is not null"
         stVTB6 = "select sum(AXF04*A1906) F1" & _
     " from acc150,acc151,acc190 where a1507 is null and a1503='" & stAgentNo & "'" & _
      " and a1502 >= " & stDate1 & " And a1502 <= " & stDate2 & _
      " and axf01(+)=a1501 " & strCFcon & " and A1902(+)=AXF01 AND A1901 is not null"
      
   'Add by Morgan 2008/7/23
   'CF 盈虧
   For intI = 1 To 5
      If intI < 4 Then
         stCol = "cp" & (60 + intI)
      Else
         stCol = "cp" & (83 + intI)
      End If
      If intI > 1 Then
         stVTB7 = stVTB7 & " union "
      End If
      'Modified by Lydia 2024/01/30
      'stVTB7 = stVTB7 & "select cp01,cp02,cp03,cp04,cp09,cp16,cp77,cp18,cp31" & _
         " from caseprogress where cp01='" & strCF & "' and " & stCol & " in (" & _
         " select a1501 from acc150 where a1507 is null and a1503='" & stAgentNo & "'" & _
         " and a1502>=" & stDate1 & " And a1502<=" & stDate2 & ")"
      stVTB7 = stVTB7 & "select cp01,cp02,cp03,cp04,cp09,cp16,cp77,cp18,cp31" & _
         " from caseprogress where " & IIf(frm050408.txtKind = "1", "cp01 in ('CFP','P')", "cp01 in ('CFT','T')") & " and " & stCol & " in (" & _
         " select a1501 from acc150 where a1507 is null and a1503='" & stAgentNo & "'" & _
         " and a1502>=" & stDate1 & " And a1502<=" & stDate2 & ")"
   Next
   strExc(0) = "select * from (" & stVTB1 & ") A,(" & stVTB2 & ") B,(" & stVTB3 & ") C,(" & stVTB4 & ") D,(" & stVTB5 & ") E,(" & stVTB6 & ") F "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      InsertQueryLog (RsTemp.RecordCount) 'Added by Lydia 2025/07/30
      With frm050408_3
         'Add By Sindy 2013/5/24
         'Modified by Morgan 2015/12/1
         'If strCF = "FCP" Then '專利
         If frm050408.txtKind = "1" Then '專利
         'end 2015/12/1
            .Label6.Caption = "FCP請款金額:"
            .Label11.Caption = "FCP請款服務費:"
            .Label7.Caption = "FCP收款金額:"
            .Label8.Caption = "CFP應付金額:"
            .Label9.Caption = "CFP付款金額:"
            .Label12.Caption = "CFP盈虧:"
            .Label10.Caption = "FCP請款 - CFP應付:"
         'Modified by Morgan 2015/12/1
         'ElseIf strCF = "FCT" Then '商標
         ElseIf frm050408.txtKind = "2" Then '商標
         'end 2015/12/1
            .Label6.Caption = "FCT請款金額:"
            .Label11.Caption = "FCT請款服務費:"
            .Label7.Caption = "FCT收款金額:"
            .Label8.Caption = "CFT應付金額:"
            .Label9.Caption = "CFT付款金額:"
            .Label12.Caption = "CFT盈虧:"
            .Label10.Caption = "FCT請款 - CFT應付:"
         End If
         '2013/5/24 End
         .lblAgentName = grdDataList.TextMatrix(p_iRow, 0)
         .lblAgentNo = grdDataList.TextMatrix(p_iRow, 1)
         .lblCondition = txt1(0) & " - " & txt1(1)
         .lblContact = grdDataList.TextMatrix(p_iRow, 2)
         .lblNation = grdDataList.TextMatrix(p_iRow, 3)
         .Text1 = Format(Val("" & RsTemp.Fields("A1")), "#,##0")
         .Text2 = Format(Val("" & RsTemp.Fields("B1")) + Val("" & RsTemp.Fields("C1")), "#,##0")
         .Text3 = Format(Val("" & RsTemp.Fields("D1")), "#,##0")
         .Text4 = Format(Val("" & RsTemp.Fields("E1")) + Val("" & RsTemp.Fields("F1")), "#,##0")
         .Text5 = Format(Format(.Text1) - Format(.Text3), "#,##0")
         'Add by Morgan 2008/7/23
         .Text6 = Format(Val("" & RsTemp.Fields("A1")) - Val("" & RsTemp.Fields("A2")), "#,##0")
         'CF 盈虧
         strExc(0) = "select cp01,cp02,cp03,cp04,sum(decode(nvl(cp16,0)-nvl(cp77,0),0,0,nvl(cp16,0)-nvl(CP18,0)*1000)) Net,max(cp31) CP31" & _
            " from (" & stVTB7 & ") group by cp01,cp02,cp03,cp04"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            .MoveFirst
            dblNet = 0
            Do While Not .EOF
               dblNet = dblNet + Val("" & .Fields("Net"))
               If Not IsNull(.Fields("CP31")) Then
                  dblNet = dblNet - GetFloatPrepareCase(.Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04"))
               End If
               .MoveNext
            Loop
            End With
            .Text7 = Format(dblNet - Format(.Text3), "#,##0")
         End If
         'end 2008/7/23
         Screen.MousePointer = vbDefault 'Added by Lydia 2025/06/27
         .Show vbModal
      End With
   'Added by Lydia 2025/07/30
   Else
      InsertQueryLog (0) 'Added by Lydia 2025/07/30
   'end 2025/07/30
   End If
End Sub

'案件統計
Private Sub GetStatistic(p_iRow As Integer, p_iCol As Integer)
   Dim stCon As String, iSys As Integer, iPos As Integer, stCP44 As String, stCP116 As String
   Dim stVTB0 As String, stVTB1 As String, stVTB2 As String, stVTB3 As String
   Dim stDate As String, iYear As Integer
   Dim stAgentNo As String, stCondition As String, dblNet As Double, stCaseNo As String
   Dim iRow As Integer, iCol As Integer
   Dim bolNew As Boolean 'Added by Lydia 2025/06/27
   
   stAgentNo = grdDataList.TextMatrix(p_iRow, 1)
   'Modified by Lydia 2025/06/27 改成共用模組
   bolNew = False
If bolNew = False Then
   stDate = strSrvDate(1)
   iYear = Left(stDate, 4)
   iPos = InStr(stAgentNo, "-")
   If iPos = 0 Then
      stCP44 = Left(stAgentNo & "000", 9)
      stCP116 = ""
   Else
      stCP44 = Left(Left(stAgentNo, iPos - 1) & "000", 9)
      stCP116 = Right("00" & Mid(stAgentNo, iPos + 1), 2)
   End If

   stCon = ""
   Select Case p_iCol
      Case 4, 6, 8, 10, 13, 18 'FC
         iSys = 1
         If strFC = "FCP" Then
            'Modified by Lydia 2018/06/21 改新申請案性質
            'stCon = stCon & " and instr('101,102,103,104,105,307',cp10)>0 and cp09<'B'"
            'Modified by Lydia 2023/09/12 改新申請案性質同excel
            'stCon = stCon & " and instr('" & NewCasePtyList & "',cp10)>0 and cp09<'B'"
            stCon = stCon & " and (instr('" & NewCasePtyList & "', cp10) > 0 or cp10 like '3%') and cp09<'B'"
            'Modified by Lydia 2024/01/30 (Widen)因應目前互惠的策略已擴大到事務所層級，再麻煩進行程式的調整到所有國家地區>> 加抓P案
            'stCon = stCon & " and pa01||''='FCP' and pa75='" & stCP44 & "'"
            stCon = stCon & " and (pa01||''='FCP' or pa01||''='P') and pa75='" & stCP44 & "'"
            'Modify by Morgan 2008/7/22 加以代理人統計
            If m_bolByAgent = False Then
               If stCP116 = "" Then
                  stCon = stCon & " and pa144 is null"
               Else
                  stCon = stCon & " and pa144='" & stCP116 & "'"
               End If
            End If
         'Modify By Sindy 2013/5/24
         ElseIf strFC = "FCT" Then
            stCon = stCon & " and instr('101',cp10)>0 and cp09<'B'"
            'Modified by Lydia 2024/01/30 (Widen)因應目前互惠的策略已擴大到事務所層級，再麻煩進行程式的調整到所有國家地區>> 加抓T案
            'stCon = stCon & " and tm01||''='FCT' and tm44='" & stCP44 & "'"
            stCon = stCon & " and (tm01||''='FCT' or tm01||''='T') and tm44='" & stCP44 & "'"
            '加以代理人統計
            If m_bolByAgent = False Then
               If stCP116 = "" Then
                  stCon = stCon & " and tm119 is null"
               Else
                  stCon = stCon & " and tm119='" & stCP116 & "'"
               End If
            End If
         End If
         '2013/5/24 End
         stCon = stCon & " and cp05<=" & stDate
         If p_iCol = 6 Then '前年
            stCon = stCon & " and cp05 between " & (iYear - 2) & "0101 and " & (iYear - 2) & "1231"
            stCondition = strFC & " " & (iYear - 1911 - 2) & " 年"
         ElseIf p_iCol = 8 Then '去年
            stCon = stCon & " and cp05 between " & (iYear - 1) & "0101 and " & (iYear - 1) & "1231"
            stCondition = strFC & " " & (iYear - 1911 - 1) & " 年"
         ElseIf p_iCol = 10 Then '當年1-6月
            stCon = stCon & " and cp05 between " & iYear & "0101 and " & iYear & "0630"
            stCondition = strFC & " " & (iYear - 1911) & " 年 1-6 月"
         ElseIf p_iCol = 13 Then '當年7-12月
            stCon = stCon & " and cp05 between " & iYear & "0701 and " & iYear & "1231"
            stCondition = strFC & " " & (iYear - 1911) & " 年 7-12 月"
         ElseIf p_iCol = 18 Then '指定區間
            stCon = stCon & " and cp05 between " & m_stDate1 & " And " & m_stDate2
            stCondition = strFC & " " & (m_stDate1 - 19110000) & " － " & (m_stDate2 - 19110000)
         Else
            stCondition = strFC & " 全部"
         End If

      Case 5, 7, 9, 11, 14, 19 'CF
         iSys = 2
         If strCF = "CFP" Then
            'Modified by Lydia 2023/09/12 改新申請案性質同excel
            'stCon = stCon & " and cp01||cp04='CFP00' and instr('" & NewCasePtyList & "',cp10)>0 and cp09<'B'"
            'Modified by Lydia 2024/01/30 (Widen)因應目前互惠的策略已擴大到事務所層級，再麻煩進行程式的調整到所有國家地區>> 加抓P案
            'stCon = stCon & " and cp01||cp04='CFP00' and (instr('" & NewCasePtyList & "', cp10) > 0 or cp10 like '3%') and cp09<'B'"
            stCon = stCon & " and instr('CFP00,P00',cp01||cp04)>0  and (instr('" & NewCasePtyList & "', cp10) > 0 or cp10 like '3%') and cp09<'B'"
         'Modify By Sindy 2013/5/24
         ElseIf strCF = "CFT" Then
            'Modified by Lydia 2024/01/30 (Widen)因應目前互惠的策略已擴大到事務所層級，再麻煩進行程式的調整到所有國家地區>> 加抓T案
            'stCon = stCon & " and cp01||cp04='CFT00' and instr('101',cp10)>0 and cp09<'B'"
            'Modified by Lydia 2025/02/05 debug instr('CFT00,T00',cp01||cp04)>0' >>instr('CFT00,T00',cp01||cp04)>0
            stCon = stCon & " and instr('CFT00,T00',cp01||cp04)>0 and instr('101',cp10)>0 and cp09<'B'"
         End If
         '2013/5/24 End
         stCon = stCon & " and cp44='" & stCP44 & "'"
         'Modify by Morgan 2008/7/22 加以代理人統計
         If m_bolByAgent = False Then
            If stCP116 = "" Then
               stCon = stCon & " and cp116 is null"
            Else
               stCon = stCon & " and cp116='" & stCP116 & "'"
            End If
         End If
         stCon = stCon & " and cp27<=" & stDate

         If p_iCol = 7 Then '前年
            stCon = stCon & " and cp27 between " & (iYear - 2) & "0101 and " & (iYear - 2) & "1231"
            stCondition = strCF & " " & (iYear - 1911 - 2) & " 年"
         ElseIf p_iCol = 9 Then '去年
            stCon = stCon & " and cp27 between " & (iYear - 1) & "0101 and " & (iYear - 1) & "1231"
            stCondition = strCF & " " & (iYear - 1911 - 1) & " 年"
         ElseIf p_iCol = 11 Then '當年1-6月
            stCon = stCon & " and cp27 between " & iYear & "0101 and " & iYear & "0630"
            stCondition = strCF & " " & (iYear - 1911) & " 年 1-6 月"
         ElseIf p_iCol = 14 Then '當年7-12月
            stCon = stCon & " and cp27 between " & iYear & "0701 and " & iYear & "1231"
            stCondition = strCF & " " & (iYear - 1911) & " 年 7-12 月"
         ElseIf p_iCol = 19 Then '指定區間
            stCon = stCon & " and cp27 between " & m_stDate1 & " And " & m_stDate2
            stCondition = strCF & " " & (m_stDate1 - 19110000) & " － " & (m_stDate2 - 19110000)
         Else
            stCondition = strCF & " 全部"
         End If
   End Select

   If iSys = 1 Then
      If strFC = "FCP" Then
         stVTB0 = "SELECT distinct PA01,PA02,PA03,PA04 FROM PATENT,CASEPROGRESS" & _
            " where cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & stCon & _
            " AND cp57 is null" & stCon
      'Modify By Sindy 2013/5/24
      ElseIf strFC = "FCT" Then
         stVTB0 = "SELECT distinct TM01,TM02,TM03,TM04 FROM Trademark,CASEPROGRESS" & _
            " where cp01(+)=TM01 and cp02(+)=TM02 and cp03(+)=TM03 and cp04(+)=TM04" & stCon & _
            " AND cp57 is null" & stCon
      End If
      '2013/5/24 End

      'Modify by Morgan 2008/6/5 FCP只顯示案件盈虧
      'strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) C01" & _
         ",null C02,cp09,cpm03,null C03" & _
         " from (" & stVTB0 & ") X,caseprogress,casepropertymap" & _
         " where cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
         " and (cp16>0 or cp60 is not null) and cpm01(+)=cp01 and cpm02(+)=cp10"

      If strFC = "FCP" Then
         stVTB1 = "SELECT distinct cp01,cp02,cp03,cp04,cp60 FROM PATENT,CASEPROGRESS" & _
            " where cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
            " AND cp57 is null and cp60 is not null" & stCon
      'Modify By Sindy 2013/5/24
      ElseIf strFC = "FCT" Then
         stVTB1 = "SELECT distinct cp01,cp02,cp03,cp04,cp60 FROM Trademark,CASEPROGRESS" & _
            " where cp01(+)=TM01 and cp02(+)=TM02 and cp03(+)=TM03 and cp04(+)=TM04" & _
            " AND cp57 is null and cp60 is not null" & stCon
      End If
      '2013/5/24 End

      '--案件盈虧=請款金額-折讓-規費-作業失誤
      '請款金額-折讓-規費
      'Modify By Sindy 2012/10/12
'      stVTB2 = "select cp01,cp02,cp03,cp04,sum(nvl(a1k11,0)-nvl(a1k06,0)*nvl(a1k10,0)-nvl(a1k09,0)) A1" & _
'         " From (" & stVTB1 & ") X,acc1k0 a Where a1k01(+)=cp60" & _
'         " group by X.cp01,X.cp02,X.cp03,X.cp04"
      stVTB2 = "select cp01,cp02,cp03,cp04,sum(nvl(a1k11,0)-nvl(a1k06,0)-nvl(a1k09,0)) A1" & _
         " From (" & stVTB1 & ") X,acc1k0 a Where a1k01(+)=cp60" & _
         " group by X.cp01,X.cp02,X.cp03,X.cp04"

      If strFC = "FCP" Then
         '作業失誤=cp17(cp16=0 & cp17>0)
         stVTB3 = "select cp01,cp02,cp03,cp04,sum(cp17) B1" & _
            " From (" & stVTB0 & ") X,caseprogress a" & _
            " Where a.cp01(+)=X.pa01 and a.cp02(+)=X.pa02 and a.cp03(+)=X.pa03 and a.cp04(+)=X.pa04" & _
            " and cp16=0 and cp17>0" & _
            " group by cp01,cp02,cp03,cp04"

         strExc(0) = "select X.pa01||'-'||X.pa02||decode(X.pa03||X.pa04,'000','','-'||X.pa03||'-'||X.pa04) C01" & _
            ",Y.pa05 C02,to_char(nvl(A1,0)-nvl(B1,0),'9,999,999.00') C03" & _
            " from (" & stVTB0 & ") X,(" & stVTB2 & ") A,(" & stVTB3 & ") B,patent Y" & _
            " where A.cp01(+)=X.pa01 and A.cp02(+)=X.pa02 and A.cp03(+)=X.pa03 and A.cp04(+)=X.pa04" & _
            " and B.cp01(+)=X.pa01 and B.cp02(+)=X.pa02 and B.cp03(+)=X.pa03 and B.cp04(+)=X.pa04" & _
            " and Y.pa01(+)=X.pa01 and Y.pa02(+)=X.pa02 and Y.pa03(+)=X.pa03 and Y.pa04(+)=X.pa04 order by 1,2"
      'Modify By Sindy 2013/5/24
      ElseIf strFC = "FCT" Then
         '作業失誤=cp17(cp16=0 & cp17>0)
         stVTB3 = "select cp01,cp02,cp03,cp04,sum(cp17) B1" & _
            " From (" & stVTB0 & ") X,caseprogress a" & _
            " Where a.cp01(+)=X.tm01 and a.cp02(+)=X.tm02 and a.cp03(+)=X.tm03 and a.cp04(+)=X.tm04" & _
            " and cp16=0 and cp17>0" & _
            " group by cp01,cp02,cp03,cp04"

         strExc(0) = "select X.tm01||'-'||X.tm02||decode(X.tm03||X.tm04,'000','','-'||X.tm03||'-'||X.tm04) C01" & _
            ",Y.tm05 C02,to_char(nvl(A1,0)-nvl(B1,0),'9,999,999.00') C03" & _
            " from (" & stVTB0 & ") X,(" & stVTB2 & ") A,(" & stVTB3 & ") B,Trademark Y" & _
            " where A.cp01(+)=X.tm01 and A.cp02(+)=X.tm02 and A.cp03(+)=X.tm03 and A.cp04(+)=X.tm04" & _
            " and B.cp01(+)=X.tm01 and B.cp02(+)=X.tm02 and B.cp03(+)=X.tm03 and B.cp04(+)=X.tm04" & _
            " and Y.tm01(+)=X.tm01 and Y.tm02(+)=X.tm02 and Y.tm03(+)=X.tm03 and Y.tm04(+)=X.tm04 order by 1,2"
      End If
      '2013/5/24 End

   Else
      stVTB0 = "SELECT CP01,CP02,CP03,CP04,max(CP31) SF FROM CASEPROGRESS" & _
         " WHERE cp57 is null" & stCon & " group by cp01,cp02,cp03,cp04"

      '--案件收(不扣安全基金)(若有銷規費時會錯，案件盈虧查詢也是)
      stVTB1 = "select X.cp01,X.cp02,X.cp03,X.cp04,sum(decode(nvl(cp16,0)-nvl(cp77,0),0,0,nvl(cp16,0)-nvl(CP18,0)*1000)) A1" & _
         " From (" & stVTB0 & ") X,caseprogress a Where a.cp01(+)=X.cp01 and a.cp02(+)=X.cp02 and a.cp03(+)=X.cp03 and a.cp04(+)=X.cp04" & _
         " and (cp16>0 or cp61 is not null)" & _
         " group by X.cp01,X.cp02,X.cp03,X.cp04"

      '--案件付(只能用本所號抓因為舊系統沒串到收文號)
      stVTB2 = "select cp01,cp02,cp03,cp04,sum(decode(A1G01,null,decode(a1901,null,AXF15,AXF04*A1906),AXF04*A1G03)) B1" & _
         " From (" & stVTB0 & ") X,acc151,acc150,acc190,acc1g0" & _
         " where axf03(+)=cp01||cp02||cp03||cp04 and a1501(+)=axf01 and a1507 is null" & _
         " and A1902(+)=AXF01 AND A1G01(+)=A1512" & _
         " group by cp01,cp02,cp03,cp04"

      'Modify by Morgan 2008/6/5 案件盈虧與收文盈虧分畫面
      '--程序收付
      'stVTB3 = "select X.cp01,X.cp02,X.cp03,X.cp04,cp09,cp10" & _
         ",nvl(min(cp16-nvl(cp77,0)-nvl(CP18,0)*1000),0)-nvl(sum(decode(a1512,null,decode(a1901,null,AXF15,AXF04*A1906),AXF04*A1G03)),0) C1" & _
         " From (" & stVTB0 & ") X,caseprogress a, acc151, acc150,acc190,acc1g0" & _
         " Where a.cp01(+)=X.cp01 and a.cp02(+)=X.cp02 and a.cp03(+)=X.cp03 and a.cp04(+)=X.cp04" & _
         " and (cp16>0 or cp61 is not null)" & _
         " and axf02(+)=cp09 and a1501(+)=axf01 and a1507 is null" & _
         " and A1902(+)=AXF01 AND A1G01(+)=A1512" & _
         " group by X.cp01,X.cp02,X.cp03,X.cp04,cp09,cp10"

      'strExc(0) = "select A.cp01||'-'||A.cp02||decode(A.cp03||A.cp04,'000','','-'||A.cp03||'-'||A.cp04) C01" & _
         ",to_char(nvl(A1,0)-nvl(B1,0),'9,999,999.00') C02,cp09,cpm03,to_char(nvl(C1,0),'9,999,999.00') C03" & _
         " from (" & stVTB1 & ") A,(" & stVTB2 & ") B,(" & stVTB3 & ") C,casepropertymap" & _
         " where B.cp01(+)=A.cp01 and B.cp02(+)=A.cp02 and B.cp03(+)=A.cp03 and B.cp04(+)=A.cp04" & _
         " and C.cp01(+)=A.cp01 and C.cp02(+)=A.cp02 and C.cp03(+)=A.cp03 and C.cp04(+)=A.cp04" & _
         " and cpm01(+)=c.cp01 and cpm02(+)=c.cp10 order by 1,3"

      If strFC = "FCP" Then
         strExc(0) = "select X.cp01||'-'||X.cp02||decode(X.cp03||X.cp04,'000','','-'||X.cp03||'-'||X.cp04) C01" & _
            ",pa05 C02,to_char(nvl(A1,0)-nvl(B1,0),'9,999,999.00') C03,X.SF,0 SFV,X.cp01,X.cp02,X.cp03,X.cp04" & _
            " from (" & stVTB0 & ") X,(" & stVTB1 & ") A,(" & stVTB2 & ") B,patent" & _
            " where A.cp01(+)=X.cp01 and A.cp02(+)=X.cp02 and A.cp03(+)=X.cp03 and A.cp04(+)=X.cp04" & _
            " and B.cp01(+)=X.cp01 and B.cp02(+)=X.cp02 and B.cp03(+)=X.cp03 and B.cp04(+)=X.cp04" & _
            " and pa01(+)=X.cp01 and pa02(+)=X.cp02 and pa03(+)=X.cp03 and pa04(+)=X.cp04 order by 1,2"
         'end 2008/6/5
      'Modify By Sindy 2013/5/24
      ElseIf strFC = "FCT" Then
         strExc(0) = "select X.cp01||'-'||X.cp02||decode(X.cp03||X.cp04,'000','','-'||X.cp03||'-'||X.cp04) C01" & _
            ",tm05 C02,to_char(nvl(A1,0)-nvl(B1,0),'9,999,999.00') C03,X.SF,0 SFV,X.cp01,X.cp02,X.cp03,X.cp04" & _
            " from (" & stVTB0 & ") X,(" & stVTB1 & ") A,(" & stVTB2 & ") B,Trademark" & _
            " where A.cp01(+)=X.cp01 and A.cp02(+)=X.cp02 and A.cp03(+)=X.cp03 and A.cp04(+)=X.cp04" & _
            " and B.cp01(+)=X.cp01 and B.cp02(+)=X.cp02 and B.cp03(+)=X.cp03 and B.cp04(+)=X.cp04" & _
            " and tm01(+)=X.cp01 and tm02(+)=X.cp02 and tm03(+)=X.cp03 and tm04(+)=X.cp04 order by 1,2"
      End If
      '2013/5/24 End
   End If
Else
   '＊＊若表單frm050408_1的欄位有變動，呼叫Pub_Frm050408_GetStatistic也要變動＊＊
   Select Case p_iCol
      Case 4, 6, 8, 10, 13, 18 'FC
         iSys = 1
      Case 5, 7, 9, 11, 14, 19 'CF
         iSys = 2
   End Select
   Call Pub_Frm050408_GetStatistic(frm050408.txtKind, frm050408.txtYear, frm050408.cboPeriod.ListIndex + 1, m_bolByAgent, strFC, strCF, p_iCol, stAgentNo, stCondition, m_stDate1, m_stDate2)
   'R004欄位長度500，專門放案件名稱
   strExc(0) = "select r001 as c01, r004 as c02, r003 as c03, r002 as sf, r005 as sfv, r006 as cp01, r007 as cp02, r008 as cp03, r009 as cp04 from rdatafactory where formname='frm050408_2' and id = '" & strUserNum & "' order by seqno,rowseq "
   'end 2025/06/27
End If
   'Added by Lydia 2025/07/30 查詢印表記錄檔欄位
   ClearQueryLog ("frm050408")
   pub_QL05 = frm050408.m_strQL05 & ";" & stCondition
   'end 2025/07/30
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      InsertQueryLog (RsTemp.RecordCount) 'Added by Lydia 2025/07/30
      With frm050408_2
         .lblAgentName = grdDataList.TextMatrix(p_iRow, 0)
         .lblAgentNo = grdDataList.TextMatrix(p_iRow, 1)
         .lblCondition = stCondition
         .lblContact = grdDataList.TextMatrix(p_iRow, 2)
         .lblNation = grdDataList.TextMatrix(p_iRow, 3)
         .grdDataList.Visible = False
         .SetDataListWidth
         Set .grdDataList.Recordset = RsTemp.Clone
         
         'FC不可點選程序盈虧
         If iSys = 1 Then
            .m_bolNoDetail = True
         End If
         With .grdDataList
            dblNet = 0
            stCaseNo = ""
            For iRow = 1 To .Rows - 1
               If stCaseNo <> .TextMatrix(iRow, 0) Then
                  stCaseNo = .TextMatrix(iRow, 0)
                  
                  'Added by Lydia 2025/06/27 改成共用模組，模組已計算
                  If bolNew = True Then
                  Else
                  'end 2025/06/27
                     'Add by Morgan 2008/7/22
                     'CFP新案要扣安全基金
                     If Left(stCaseNo, 3) = "CFP" Then
                        If .TextMatrix(iRow, 3) = "Y" Then
                           .TextMatrix(iRow, 4) = GetFloatPrepareCase(.TextMatrix(iRow, 5), .TextMatrix(iRow, 6), .TextMatrix(iRow, 7), .TextMatrix(iRow, 8))
                           .TextMatrix(iRow, 2) = Format(.TextMatrix(iRow, 2)) - Val(.TextMatrix(iRow, 4))
                           .TextMatrix(iRow, 4) = Format(.TextMatrix(iRow, 4), "#,###.00")
                           .TextMatrix(iRow, 2) = Format(.TextMatrix(iRow, 2), "#,###.00")
                        End If
                     End If
                     'end 2008/7/22
                  End If 'end 2025/06/27
                  dblNet = dblNet + Format(.TextMatrix(iRow, 2))
               End If
            Next
         End With
         Screen.MousePointer = vbDefault 'Added by Lydia 2025/06/27
         .SetDataListWidth True
         .lblNetTot = Format(dblNet, "#,###.00")
         .grdDataList.Visible = True
         .Show vbModal
      End With
   Else
      InsertQueryLog (0) 'Added by Lydia 2025/07/30
      MsgBox "無資料！"
   End If
End Sub

Private Sub grdDataList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim iRow As Integer, iCol As Integer, lBackColor As Long
   
   With grdDataList
      iRow = .MouseRow
      iCol = .MouseCol
      If iRow < 0 Or iCol < 0 Then Exit Sub
      If iRow = m_iRow And iCol = m_iCol Then
         Exit Sub
      End If
      If m_iCol <> 0 Then
         .row = m_iRow: .col = m_iCol
         If m_iCol < 3 Then
            .CellForeColor = .ForeColorFixed
            .CellBackColor = .BackColorFixed
         Else
            .CellForeColor = .ForeColor
            .CellBackColor = .BackColor
         End If
         m_iRow = 0: m_iCol = 0
      End If
      
      If iRow > 0 Then
         Select Case iCol
            Case 1, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 18, 19
               .row = iRow: .col = iCol
               lBackColor = .CellBackColor
               .CellBackColor = .CellForeColor
               .CellForeColor = lBackColor
               m_iCol = .col
               m_iRow = .row
            Case Else
         End Select
      End If
   End With
End Sub

Private Function DoPrint() As Boolean
   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
   Dim strTemp(1 To 18) As String
   
   GetPleft
   iOrientation = Printer.Orientation
   Printer.Orientation = 2
   lngPageHeight = 11900
   lngPageWidth = 16800
   lngLineHeight = 300
   With grdDataList
      iPage = 1
      InsertQueryLog (.Rows - 1) 'Added by Lydia 2025/07/30
      PrintPageHeader
      For iRow = 1 To .Rows - 1
         For iCol = 1 To 18
            Select Case iCol
               Case 1
                  strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 20)
               Case 3
                  strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 10)
               Case 18
                  'Modified by Lydia 2017/06/09 備註全顯示
                  'strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 4)
                  'Modified by Lydia 2022/04/06 Y2099001在2022上半年的備註超長
                  'strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 150)
                  strTemp(iCol) = Trim(.TextMatrix(iRow, iCol - 1))
               Case Else
                  strTemp(iCol) = .TextMatrix(iRow, iCol - 1)
            End Select
         Next
         PrintDetail strTemp
      Next
      Call PrintReportFooter(.Rows - 1)
      Printer.EndDoc
   End With
   Printer.Orientation = iOrientation
   DoPrint = True
End Function

'Add by Morgan 2008/6/24
Private Function DoPrint1() As Boolean
   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
   Dim strTemp(1 To 8) As String
      
   GetPleft1
   
   iOrientation = Printer.Orientation
   Printer.Orientation = 2
   lngPageHeight = 11900
   lngPageWidth = 16800
   lngLineHeight = 300
   With grdDataList
      iPage = 1
      InsertQueryLog (.Rows - 1) 'Added by Lydia 2025/07/30
      PrintPageHeader1
      For iRow = 1 To .Rows - 1
         For iCol = 1 To 4
            Select Case iCol
               Case 1
                  strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 30)
               Case 3
                  strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 30)
               Case Else
                  strTemp(iCol) = .TextMatrix(iRow, iCol - 1)
            End Select
         Next
         strTemp(5) = .TextMatrix(iRow, 18)
         strTemp(6) = .TextMatrix(iRow, 19)
         strTemp(7) = .TextMatrix(iRow, 20)
         'Modified by Lydia 2017/06/09 備註全顯示
         'strTemp(8) = Left(.TextMatrix(iRow, 17), 20)
         'Modified by Lydia 2022/04/06 Y2099001在2022上半年的備註超長
         'strTemp(8) = Left(.TextMatrix(iRow, 17), 150)
         strTemp(8) = Trim(.TextMatrix(iRow, 17))
         PrintDetail1 strTemp
      Next
      Call PrintReportFooter1(.Rows - 1)
      Printer.EndDoc
   End With
   Printer.Orientation = iOrientation
   DoPrint1 = True
End Function

Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   ReDim PLeft(1 To 18)
   PLeft(1) = ciStartX
   PLeft(2) = PLeft(1) + ciColGap + Printer.TextWidth(String(10, "　"))
   PLeft(3) = PLeft(2) + ciColGap + Printer.TextWidth(String(5, "　"))
   PLeft(4) = PLeft(3) + ciColGap + Printer.TextWidth(String(5, "　"))
   PLeft(5) = PLeft(4) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(6) = PLeft(5) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(7) = PLeft(6) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(8) = PLeft(7) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(9) = PLeft(8) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(10) = PLeft(9) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(11) = PLeft(10) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(12) = PLeft(11) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(13) = PLeft(12) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(14) = PLeft(13) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(15) = PLeft(14) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(16) = PLeft(15) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(17) = PLeft(16) + ciColGap + Printer.TextWidth(String(3, "　"))
   PLeft(18) = PLeft(17) + ciColGap + Printer.TextWidth(String(3, "　"))
End Sub

Sub GetPleft1()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   ReDim PLeft(1 To 8)
   PLeft(1) = ciStartX
   PLeft(2) = PLeft(1) + 2 * ciColGap + Printer.TextWidth(String(15, "　"))
   PLeft(3) = PLeft(2) + 2 * ciColGap + Printer.TextWidth(String(5, "　"))
   PLeft(4) = PLeft(3) + 2 * ciColGap + Printer.TextWidth(String(15, "　"))
   PLeft(5) = PLeft(4) + 2 * ciColGap + Printer.TextWidth(String(5, "　"))
   PLeft(6) = PLeft(5) + 2 * ciColGap + Printer.TextWidth(String(5, "　"))
   PLeft(7) = PLeft(6) + 2 * ciColGap + Printer.TextWidth(String(5, "　"))
   PLeft(8) = PLeft(7) + 2 * ciColGap + Printer.TextWidth(String(5, "　"))
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = frm050408.Caption
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   iPrint = iPrint + lngLineHeight
   strPTmp = "統計年度區間：" & frm050408.txtYear & "年" & frm050408.cboPeriod
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   'Added by Lydia 2022/01/11 增加統計對象：條件顯示---Winden 於110/12/9口頭提出
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint + lngLineHeight
   Printer.Print "統計對象：" & frm050408.cboTarget
   'end 2022/01/11
   
   iPrint = iPrint + lngLineHeight
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   iPrint = iPrint + lngLineHeight
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   iPrint = iPrint + lngLineHeight
   PLine
   
   iPrint = iPrint + lngLineHeight
   '第一列
   For intI = 11 To 17
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      'Modified by Lydia 2023/09/12 與Excel抬頭一致
      'strPTmp = Left(grdDataList.TextMatrix(0, intI - 1), 3)
      strPTmp = Left(grdDataList.TextMatrix(0, intI - 1), 3)
      Printer.Print strPTmp
   Next
   
   iPrint = iPrint + lngLineHeight
   '第二列
   'Modified by Lydia 2023/09/12 與Excel抬頭一致
   'For intI = 7 To 17
   For intI = 5 To 17
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Select Case intI
         'Modified by Lydia 2023/09/12 與Excel抬頭一致
         'Case 7 To 10
            'strPTmp = Left(grdDataList.TextMatrix(0, intI - 1), 3)
         Case 5 To 10
         'end 2023/09/12
            strPTmp = Left(grdDataList.TextMatrix(0, intI - 1), 2)
         Case 11 To 13
            strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 4, 4)
         Case 14 To 16
            strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 4, 5)
         Case 17
            strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 4, 3)
      End Select
      Printer.Print strPTmp
   Next
   
   iPrint = iPrint + lngLineHeight
   '第三列
   For intI = 5 To 17
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Select Case intI
         Case 5 To 6
            'Modified by Lydia 2023/09/12 與Excel抬頭一致
            'strPTmp = Left(grdDataList.TextMatrix(0, intI - 1), 4)
             strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 3)
         Case 7 To 10
            strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 4, 3)
         Case 11 To 13
            strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 8, 3)
         Case 14 To 16
            strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 9, 3)
         Case 17
            strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 7, 3)
      End Select
      Printer.Print strPTmp
   Next
   
   iPrint = iPrint + lngLineHeight
   '第四列
   'Modified by Lydia 2017/06/09 備註換行
   'For intI = 1 To 18
   For intI = 1 To 17
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Select Case intI
         Case 5 To 6
            'Modified by Lydia 2023/09/12 與Excel抬頭一致
            'strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 5)
            strPTmp = ""
         Case 7 To 10
            strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 7)
         Case 11 To 13
            strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 11)
         Case 14 To 16
            strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 12)
         Case 17
            strPTmp = Mid(grdDataList.TextMatrix(0, intI - 1), 10)
         Case Else
            strPTmp = grdDataList.TextMatrix(0, intI - 1)
'            'Added by Lydia 2017/06/09 備註換行
'            If Trim(strPTmp) <> "" And intI = 18 Then
'               iPrint = iPrint + lngLineHeight
'               Printer.CurrentX = PLeft(1)
'               Printer.CurrentY = iPrint
'            End If
'            'end 2017/06/09
      End Select
      Printer.Print strPTmp
   Next
   iPrint = iPrint + lngLineHeight
   PLine
End Sub

Sub PrintPageHeader1()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = frm050408.Caption
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   iPrint = iPrint + lngLineHeight
   strPTmp = "統計年度區間：" & frm050408.txtYear & "年" & frm050408.cboPeriod
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp

       
   iPrint = iPrint + lngLineHeight
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   
   strPTmp = lblExtra
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   iPrint = iPrint + lngLineHeight
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   iPrint = iPrint + lngLineHeight
   PLine
   iPrint = iPrint + lngLineHeight
   
   For intI = 1 To 4
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Printer.Print grdDataList.TextMatrix(0, intI - 1)
   Next
   
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   'Modified by Lydia 2023/09/12 與Excel抬頭一致
   'Printer.Print strFC & "收案量" '"FCP收案量"
   Printer.Print strFC
   
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   'Modified by Lydia 2023/09/12 與Excel抬頭一致
   'Printer.Print strCF & "給案量" '"CFP給案量"
   Printer.Print strCF
   
   'Remove by Lydia 2023/09/12 與Excel抬頭一致
   'Printer.CurrentX = PLeft(7)
   'Printer.CurrentY = iPrint
   'Printer.Print "收給案量差"
   
   'Remove by Lydia 2017/06/09 備註換行
   'Printer.CurrentX = PLeft(8)
   'Printer.CurrentY = iPrint
   'Printer.Print "備註"
   
   iPrint = iPrint + lngLineHeight
   PLine
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      PLine
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      iPrint = iPrint + lngLineHeight
   End If
End Sub

Private Sub PrintNewLine1(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      PLine
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader1
      iPrint = iPrint + lngLineHeight
   End If
End Sub

Private Sub PLine()
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print String(156, "-")
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    Dim tmpArr As Variant, intP As Integer, strMid As String, strText As String  'Added by Lydia 2022/04/06
    
    PrintNewLine
    
    For iCol = 1 To UBound(strData)
      'Modified by Lydia 2017/06/09 備註換行
      'If iCol < 4 Or iCol = 18 Then
      If iCol < 4 Then
         Printer.CurrentX = PLeft(iCol)
         Printer.CurrentY = iPrint
         Printer.Print strData(iCol)
      'Added by Lydia 2017/06/09 備註換行
      ElseIf iCol = 18 Then
         If strData(iCol) <> "" Then
            'Modifified by Lydia 2022/04/06 判斷是否換行
            'PrintNewLine
            'Printer.CurrentX = PLeft(1)
            'Printer.CurrentY = iPrint
            'Printer.Print "備註：" & strData(iCol)
            tmpArr = Split(strData(iCol), vbCrLf)
            For intP = 0 To UBound(tmpArr)
                If Trim(tmpArr(intP)) <> "" Then
                    If intP = 0 Then
                        strMid = "備註：" & Trim(tmpArr(intP))
                    Else
                        strMid = " 　　" & Trim(tmpArr(intP))
                    End If
                    Do While strMid <> ""
                       strText = PUB_StrToStr(strMid, 150)
                       PrintNewLine
                       Printer.CurrentX = PLeft(1)
                       Printer.CurrentY = iPrint
                       Printer.Print strText
                       strMid = Trim(Replace(strMid, strText, ""))
                       If strMid <> "" Then strMid = "　　" & strMid
                    Loop
                End If
            Next
            'end 2022/04/06
         End If
      'end 2017/06/09
      Else
        Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - ciColGap
        Printer.CurrentY = iPrint
        Printer.Print strData(iCol)
      End If
    Next
End Sub

Sub PrintDetail1(strData() As String)
    Dim iCol As Integer
    Dim tmpArr As Variant, intP As Integer, strMid As String, strText As String  'Added by Lydia 2022/04/06
    
    PrintNewLine1

    For iCol = 1 To UBound(strData)
      'Modified by Lydia 2017/06/09 備註換行
      'If iCol < 5 Or iCol = 8 Then
      If iCol < 5 Then
         Printer.CurrentX = PLeft(iCol)
         Printer.CurrentY = iPrint
         Printer.Print strData(iCol)
      'Added by Lydia 2017/06/09 備註換行
      ElseIf iCol = 8 Then
         If strData(iCol) <> "" Then
            'Modifified by Lydia 2022/04/06 判斷是否換行
            'PrintNewLine
            'Printer.CurrentX = PLeft(1)
            'Printer.CurrentY = iPrint
            'Printer.Print "備註：" & strData(iCol)
            tmpArr = Split(strData(iCol), vbCrLf)
            For intP = 0 To UBound(tmpArr)
                If Trim(tmpArr(intP)) <> "" Then
                    If intP = 0 Then
                        strMid = "備註：" & Trim(tmpArr(intP))
                    Else
                        strMid = " 　　　" & Trim(tmpArr(intP))
                    End If
                    Do While strMid <> ""
                       strText = PUB_StrToStr(strMid, 150)
                       PrintNewLine
                       Printer.CurrentX = PLeft(1)
                       Printer.CurrentY = iPrint
                       Printer.Print strText
                       strMid = Trim(Replace(strMid, strText, ""))
                       If strMid <> "" Then strMid = "　　" & strMid
                    Loop
                End If
            Next
            'end 2022/04/06
         End If
      'end 2017/06/09
      Else
        Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - ciColGap
        Printer.CurrentY = iPrint
        Printer.Print strData(iCol)
      End If
    Next
End Sub

'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 1)
    PLine
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub

'列印表尾
Private Sub PrintReportFooter1(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine1(True, 1)
    PLine
    PrintNewLine1
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub

Public Sub SetCombo()
   Dim iRow As Integer, sCountry As String
   m_bComboNoAct = True
   With grdDataList
      Combo1.Clear
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, 3) <> sCountry Then
            sCountry = .TextMatrix(iRow, 3)
            Combo1.AddItem sCountry
         End If
      Next
   End With
   m_bComboNoAct = False
   Combo1.ListIndex = 0
End Sub

'Added by Lydia 2020/09/22 產生Excel檔; 固定格式不受Grid的影響
Private Sub CmdProcExcel_Click()
    Call Pub_ChkExcelPath 'Added by Lydia 2021/07/01 檢查xls資料夾的模組
    
    'Memo by Lydia 2021/04/23 調整版面，舊版面另外抽出保留在ProcExcelSave_2020
    'Modified by Lydia 2022/05/12 (2022)調整版面
    'Call ProcExcelSave_2021
    CmdProcExcel.Enabled = False
    Screen.MousePointer = vbHourglass
    'Modified by Lydia 2023/09/12 (2023)調整版面
    'Call ProcExcelSave_2022
    Call ProcExcelSave_2023
    Screen.MousePointer = vbDefault
    CmdProcExcel.Enabled = True
End Sub

'Added by Lydia 2021/0/4/23 保留前一版面
Private Sub ProcExcelSave_2020()
Dim strTitleL(0 To 4) As String
Dim strCon1 As String
Dim intJ As Integer
Dim rsRD As New ADODB.Recordset
Dim rsAD As New ADODB.Recordset
Dim xlsPoint1 As New Excel.Application
Dim wksPoint1 As New Worksheet
Dim strFileName As String '檔案名稱
Dim iRow As Integer
Dim xCols As Integer
Dim MaxCols As Integer '最大行位置
Dim tmpArr1 As Variant, tmpArr2 As Variant
Dim strTmp As String
Dim m_bolAgent As Boolean

On Error GoTo ErrHandle
    
    If frm050408.cboTarget.ListIndex = 0 Then
        m_bolAgent = True
    Else
        m_bolAgent = False
    End If
    
    '設定抬頭為4列
    'X1月=1-6月,X2月=7-12月,Y2年=109下半年
    'B1=FCP,B2=CFP,B3=專利
    'M = 合併儲存格
    'Modified by Lydia 2020/12/22 資料欄位是固定的
    'strTitleL(1) = "代理人名稱,代理人編號,聯絡人資訊,國籍   ,空白        ,空白       ,前年度 ,前年度 ,去年度,去年度,當年度X1月,當年度X1月,當年度X1月,當年度X2月,當年度X2月,當年度X2月,空白               ,備註     ,空白                     ,空白             ,空白  "
    strTitleL(1) = "代理人名稱,代理人編號,聯絡人資訊,國籍   ,空白        ,空白       ,前年度 ,前年度 ,去年度,去年度,當年度1-6月,當年度1-6月,當年度1-6月,當年度7-12月,當年度7-12月,當年度7-12月,空白               ,備註     ,空白                     ,空白             ,空白  "
    strTitleL(2) = "空白           ,空白           ,空白            ,空白   ,B1           ,B2           ,B1        ,B2        ,B1        ,B2        ,B1               ,B2               ,B3               ,B1               ,B2               ,B3               ,該半年度建議,空白     ,維持/刪除/           ,Y2年             ,Y2年  "
    strTitleL(3) = "空白           ,空白           ,空白            ,空白   ,C1           ,C2           ,C1        ,C2        ,C1        ,C2        ,C1               ,C2               ,C3               ,C1               ,C2               ,C3               ,B2給案量        ,空白    ,給案量調整/新增 ,建議給案量  ,備註  "
    strTitleL(4) = "M空白        ,M空白       ,M空白         ,M空白,總收案量,總收案量,收案量 ,給案量,收案量,給案量,收案量          ,給案量       ,收給案量差,收案量        ,給案量        ,收給案量差,空白                ,M空白 ,空白                     ,空白              ,空白  "
    '欄寬
    strTitleL(0) = "32,11,15,7,9,9,10,10,10,10,13,13,13,13,13,13,13,15,15,11,10"
    
    '替換為正確內容
    strExc(1) = "": strExc(2) = "": strExc(3) = "": strExc(4) = "": strExc(5) = "": strExc(6) = ""
    'Remove by Lydia 2020/12/22
    'strExc(1) = IIf(frm050408.cboPeriod.ListIndex = 0, "1-6", "7-12")
    'strExc(2) = IIf(frm050408.cboPeriod.ListIndex = 0, "7-12", "1-6")
    'end 2020/12/22
    strExc(3) = frm050408.txtYear & IIf(frm050408.cboPeriod.ListIndex = 0, "上半年", "下半年")
    
    'strTitleL(1) = Replace(Replace(strTitleL(1), "X1", strExc(1)), "X2", strExc(2)) 'Remove by Lydia 2020/12/22
    strTitleL(2) = Replace(strTitleL(2), "Y2年", strExc(3))
    If frm050408.txtKind = "1" Then '專利為主
        strTitleL(2) = Replace(Replace(Replace(strTitleL(2), "B1", "FCP"), "B2", "CFP"), "B3", "專利")
        strTitleL(3) = Replace(Replace(Replace(Replace(strTitleL(3), "C1", "(FCT)"), "C2", "(CFT)"), "C3", "(商標)"), "B2", "CFP")
    ElseIf frm050408.txtKind = "2" Then  '商標為主
        strTitleL(2) = Replace(Replace(Replace(strTitleL(2), "B1", "FCT"), "B2", "CFT"), "B3", "商標")
        strTitleL(3) = Replace(Replace(Replace(Replace(strTitleL(3), "C1", "(FCP)"), "C2", "(CFP)"), "C3", "(專利)"), "B2", "CFT")
    End If
      
    strCon1 = frm050408.strConFirst
    intI = 1
    Set rsRD = ClsLawReadRstMsg(intI, strCon1)
    If intI = 1 Then
        strFileName = strExcelPath & strSrvDate(1) & "_互惠代理人案件統計表" & frm050408.txtYear & "年" & IIf(frm050408.cboPeriod.ListIndex = 0, "上半年", "下半年") & IIf(frm050408.txtKind = "1", "專利", "商標") & IIf(m_bolAgent = True, "(代理人)", "(聯絡人)") & MsgText(43)
        If Dir(strFileName) <> "" Then
            '檢查檔案是否正在使用中
            If PUB_ChkFileOpening(strFileName) = True Then
                 MsgBox strFileName & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
                 Exit Sub
            End If
            Kill strFileName
        End If
        
        xlsPoint1.Workbooks.add
        xlsPoint1.Visible = False '預設不顯示
        iRow = 0
        xCols = 1
        xlsPoint1.SheetsInNewWorkbook = 3 'Modify by Amy 2021/06/22 預設工作表數量
        Set wksPoint1 = xlsPoint1.Worksheets(1)
        xlsPoint1.Sheets(1).Select '選擇工作表
        tmpArr2 = Split(strTitleL(0), ",") '欄寬
        MaxCols = UBound(tmpArr2) + 1
        '凍結窗格要在合併儲存格之前, 不然A欄會自動隱藏
        wksPoint1.Range("E5").Select
        xlsPoint1.ActiveWindow.FreezePanes = True

        '欄位抬頭
        For intI = 1 To 4
            iRow = iRow + 1
            tmpArr1 = Empty
            tmpArr1 = Split(strTitleL(intI), ",")
            For intJ = 0 To UBound(tmpArr1)
                If Trim(tmpArr1(intJ)) <> "" Then
                   strExc(3) = Pub_NumberToSystem26(intJ + 1)
                   strCon1 = Trim(Replace(tmpArr1(intJ), "空白", ""))
                   wksPoint1.Range(strExc(3) & iRow).Value = IIf(strCon1 = "M", "", strCon1)
                   If Asc(strExc(3)) >= Asc("S") And Asc(strExc(3)) <= Asc("U") Then
                        wksPoint1.Range(strExc(3) & iRow).Font.Color = vbRed '字體: 紅色
                   End If
                   If intI = 3 Then '副:資料
                       If Asc(strExc(3)) >= Asc("E") And Asc(strExc(3)) <= Asc("P") Then
                          wksPoint1.Range(strExc(3) & iRow).Font.ColorIndex = 46 '字體: 橙色
                       End If
                   End If
                   If intI = 4 Then '最後處理抬頭設定
                       wksPoint1.Range(strExc(3) & ":" & strExc(3)).ColumnWidth = Val("" & tmpArr2(intJ))
                       wksPoint1.Range(strExc(3) & ":" & strExc(3)).NumberFormatLocal = "@"
                       If Asc(strExc(3)) >= Asc("E") And Asc(strExc(3)) <= Asc("Q") Then
                          wksPoint1.Range(strExc(3) & ":" & strExc(3)).HorizontalAlignment = xlCenter '數字,水平置中
                       End If
                       If strCon1 = "M" Then '合併儲存格
                           xlsPoint1.Range(strExc(3) & iRow - 3 & ":" & strExc(3) & iRow).Select
                           With xlsPoint1.Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .ShrinkToFit = False
                                .MergeCells = True
                           End With
                       End If
                   End If
                End If
            Next intJ
        Next intI

       wksPoint1.Range("Q" & iRow - 3 & ":" & "U" & iRow).Interior.ColorIndex = 6 '底色: 黃色
       iRow = iRow + 1
       
       ReDim tmpArr1(1 To MaxCols)
       strExc(3) = Pub_NumberToSystem26(MaxCols)
       rsRD.MoveFirst
       Do While Not rsRD.EOF
            For intJ = 1 To MaxCols
                If intJ >= MaxCols - 2 Then  '後3欄為人工填寫
                    tmpArr1(intJ) = ""
                Else
                    tmpArr1(intJ) = "" & rsRD.Fields(intJ - 1)
                End If
            Next intJ
            wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value = tmpArr1
            wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).NumberFormat = "@"
            
            '另外抓相對的專利/商標案件，若本身沒設相對的專利/商標互惠，還是要抓資料，給案量欄位設空白
            If "" & rsRD.Fields("C2") <> "" Then
                strCon1 = Mid(frm050408.strConSecond, 1, InStr(UCase(frm050408.strConSecond), "ORDER") - 1)
                strCon1 = strCon1 & " AND FC01=" & CNULL(Left(rsRD.Fields("C2"), 8))
                intI = 1
                Set rsAD = ClsLawReadRstMsg(intI, strCon1)
                If intI = 1 Then
                    iRow = iRow + 1
                    For intJ = 1 To MaxCols
                        If intJ < 5 Then '代理人資料
                            If m_bolAgent = False <> 0 And (intJ = 2 Or intJ = 3) And InStr("" & rsAD.Fields("C2"), "-") > 0 Then
                               tmpArr1(intJ) = "" & rsAD.Fields(intJ - 1) '聯絡人：因為專利和商標設定的聯絡人可能不一致
                            Else
                               tmpArr1(intJ) = ""  '空白
                            End If
                        ElseIf intJ >= MaxCols - 2 Then  '後3欄為人工填寫
                            tmpArr1(intJ) = ""
                        Else
                            If intJ >= 5 And intJ <= 17 Then '數量用()區隔
                                tmpArr1(intJ) = "(" & rsAD.Fields(intJ - 1) & ")"
                            Else
                                tmpArr1(intJ) = "" & rsAD.Fields(intJ - 1)
                            End If
                        End If
                    Next intJ

                    wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value = tmpArr1
                    wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).NumberFormat = "@"
                    wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Font.ColorIndex = 46
                End If
            End If
            rsRD.MoveNext
            iRow = iRow + 1
       Loop
       wksPoint1.Range("A1").Select
       xlsPoint1.ActiveWindow.ScrollColumn = 5 '將非凍結欄位的捲軸拉回到最前面
       
       '判斷版本
       If Val(xlsPoint1.Version) < 12 Then
            xlsPoint1.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
       Else
            xlsPoint1.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
       End If

       xlsPoint1.Workbooks.Close
       xlsPoint1.Quit
       Set wksPoint1 = Nothing
       Set xlsPoint1 = Nothing
       'Modify by Amy 2021/06/22 原:strFileName 改中文字顯示
       MsgBox "Excel檔案產生完成！" & vbCrLf & "檔案位置：" & strExcelPathN & " " & Replace(strFileName, strExcelPath, "")
    End If
    
    Set rsRD = Nothing
    Set rsAD = Nothing
    Exit Sub
    
ErrHandle:
    If Err.Number <> 0 Then
        MsgBox "產生Excel失敗：" & vbCrLf & Err.Description, vbCritical
    End If
End Sub

'Added by Lydia 2021/0/4/23 (2021)調整版面
Private Sub ProcExcelSave_2021()
Dim strTitleL(0 To 4) As String
Dim strCon1 As String
Dim intJ As Integer
Dim rsRD As New ADODB.Recordset
Dim rsAD As New ADODB.Recordset
Dim xlsPoint1 As New Excel.Application
Dim wksPoint1 As New Worksheet
Dim strFileName As String '檔案名稱
Dim iRow As Integer
Dim xCols As Integer
Dim MaxCols As Integer '最大行位置
Dim tmpArr1 As Variant, tmpArr2 As Variant
Dim strTmp As String
Dim m_bolAgent As Boolean
Dim colM1 As Integer '最後n欄為人工填寫

On Error GoTo ErrHandle
    
    If frm050408.cboTarget.ListIndex = 0 Then
        m_bolAgent = True
    Else
        m_bolAgent = False
    End If
    
    '設定抬頭為4列
    'X1月=1-6月,X2月=7-12月,Y2年=109下半年
    'B1=FCP,B2=CFP,B3=專利
    'M = 合併儲存格
    'Modified by Lydia 2022/01/11 代理人名稱=> 代理人名稱(統計對象：xxx)
    strTitleL(1) = "代理人名稱　（統計對象：" & frm050408.cboTarget & "）,代理人編號,聯絡人資訊,國籍   ,空白        ,空白       ,前年度 ,前年度 ,去年度,去年度,當年度1-6月,當年度1-6月,當年度1-6月,當年度7-12月,當年度7-12月,當年度7-12月,空白               ,現有備註  ,請選擇： ,空白             ,隔半年度備註,空白,空白,空白,空白"
    strTitleL(2) = "空白           ,空白           ,空白            ,空白   ,B1           ,B2           ,B1        ,B2        ,B1        ,B2        ,B1               ,B2               ,B3               ,B1               ,B2               ,B3               ,該半年度建議,空白     ,維持/刪除/           ,隔半年度建議,空白              ,B1    ,B2   ,業拓,總經理"
    strTitleL(3) = "空白           ,空白           ,空白            ,空白   ,C1           ,C2           ,C1        ,C2        ,C1        ,C2        ,C1               ,C2               ,C3               ,C1               ,C2               ,C3               ,B2給案量        ,空白    ,給案量調整(減少/ ,B2給案量  ,空白                  ,意見,意見,補充,意見"
    strTitleL(4) = "M空白        ,M空白       ,M空白         ,M空白,總收案量,總收案量,收案量 ,給案量,收案量,給案量,收案量          ,給案量       ,收給案量差,收案量        ,給案量        ,收給案量差,空白                ,M空白 ,增加)/新增       ,空白              ,M空白                ,空白,空白,空白,空白"
    '欄寬
    strTitleL(0) = "32,11,15,7,9,9,10,10,10,10,13,13,13,13,13,13,13,20,15,15,15,15,15,15,20"
        
    '替換為正確內容
    strExc(1) = "": strExc(2) = "": strExc(3) = "": strExc(4) = "": strExc(5) = "": strExc(6) = ""
    strExc(3) = frm050408.txtYear & IIf(frm050408.cboPeriod.ListIndex = 0, "上半年", "下半年")
     
    strTitleL(2) = Replace(strTitleL(2), "Y2年", strExc(3))
    If frm050408.txtKind = "1" Then '專利為主
        strTitleL(2) = Replace(Replace(Replace(strTitleL(2), "B1", "FCP"), "B2", "CFP"), "B3", "專利")
        strTitleL(3) = Replace(Replace(Replace(Replace(strTitleL(3), "C1", "FCT"), "C2", "CFT"), "C3", "商標"), "B2", "CFP")
    ElseIf frm050408.txtKind = "2" Then  '商標為主
        strTitleL(2) = Replace(Replace(Replace(strTitleL(2), "B1", "FCT"), "B2", "CFT"), "B3", "商標")
        strTitleL(3) = Replace(Replace(Replace(Replace(strTitleL(3), "C1", "FCP"), "C2", "CFP"), "C3", "專利"), "B2", "CFT")
    End If
      
    colM1 = 6 '最後7欄為人工填寫
    
    strCon1 = frm050408.strConFirst
    intI = 1
    Set rsRD = ClsLawReadRstMsg(intI, strCon1)
    If intI = 1 Then
        strFileName = strExcelPath & strSrvDate(1) & "_互惠代理人案件統計表" & frm050408.txtYear & "年" & IIf(frm050408.cboPeriod.ListIndex = 0, "上半年", "下半年") & IIf(frm050408.txtKind = "1", "專利", "商標") & IIf(m_bolAgent = True, "(代理人)", "(聯絡人)") & MsgText(43)
        If Dir(strFileName) <> "" Then
            '檢查檔案是否正在使用中
            If PUB_ChkFileOpening(strFileName) = True Then
                 MsgBox strFileName & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
                 Exit Sub
            End If
            Kill strFileName
        End If
        
        xlsPoint1.Workbooks.add
        xlsPoint1.Visible = False '預設不顯示
        iRow = 0
        xCols = 1
        xlsPoint1.SheetsInNewWorkbook = 3 'Modify by Amy 2021/06/22 預設工作表數量
        Set wksPoint1 = xlsPoint1.Worksheets(1)
        xlsPoint1.Sheets(1).Select '選擇工作表
        tmpArr2 = Split(strTitleL(0), ",") '欄寬
        MaxCols = UBound(tmpArr2) + 1
        '凍結窗格要在合併儲存格之前, 不然A欄會自動隱藏
        wksPoint1.Range("E5").Select
        xlsPoint1.ActiveWindow.FreezePanes = True
        '抬頭底色+字體設定
        wksPoint1.Range("E2:P2").Interior.ColorIndex = 15  '底色:灰
        wksPoint1.Range("Q1:Q4").Font.Color = vbRed '字體: 紅色
        wksPoint1.Range("S1:S4").Interior.ColorIndex = 6 '底色:黃
        wksPoint1.Range("T1:T4").Interior.ColorIndex = 6
        wksPoint1.Range("U1:U4").Interior.ColorIndex = 6
        wksPoint1.Range("S2:S4").Font.Color = vbRed '字體: 紅色
        wksPoint1.Range("T1:T4").Font.Color = vbRed
        wksPoint1.Range("U1:U4").Font.Color = vbRed
        wksPoint1.Range("R1:R1").Font.Color = vbRed
        If frm050408.cboPeriod.ListIndex = 0 Then '上半年
           wksPoint1.Range("L1:L2").Font.Color = vbRed '字體: 紅色
           wksPoint1.Range("L4").Font.Color = vbRed
        ElseIf frm050408.cboPeriod.ListIndex = 1 Then  '下半年
           wksPoint1.Range("O1:O2").Font.Color = vbRed '字體: 紅色
           wksPoint1.Range("O4").Font.Color = vbRed
        End If
        
        '欄位抬頭
        For intI = 1 To 4
            iRow = iRow + 1
            tmpArr1 = Empty
            tmpArr1 = Split(strTitleL(intI), ",")
            For intJ = 0 To UBound(tmpArr1)
                If Trim(tmpArr1(intJ)) <> "" Then
                   strExc(3) = Pub_NumberToSystem26(intJ + 1)
                   strCon1 = Trim(Replace(tmpArr1(intJ), "空白", ""))
                   wksPoint1.Range(strExc(3) & iRow).Value = IIf(strCon1 = "M", "", strCon1)
                   If intI = 4 Then '最後處理抬頭設定
                       wksPoint1.Range(strExc(3) & ":" & strExc(3)).ColumnWidth = Val("" & tmpArr2(intJ))
                       wksPoint1.Range(strExc(3) & ":" & strExc(3)).NumberFormatLocal = "@"
                       If Asc(strExc(3)) >= Asc("E") And Asc(strExc(3)) <= Asc("Q") Then
                          wksPoint1.Range(strExc(3) & ":" & strExc(3)).HorizontalAlignment = xlCenter '數字,水平置中
                       End If
                       If strCon1 = "M" Then '合併儲存格
                           xlsPoint1.Range(strExc(3) & iRow - 3 & ":" & strExc(3) & iRow).Select
                           With xlsPoint1.Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .ShrinkToFit = False
                                .MergeCells = True
                           End With
                       End If
                   End If
                End If
            Next intJ
        Next intI

       iRow = iRow + 1
       
       ReDim tmpArr1(1 To MaxCols)
       strExc(3) = Pub_NumberToSystem26(MaxCols)
       rsRD.MoveFirst
       Do While Not rsRD.EOF
            For intJ = 1 To MaxCols
                If intJ >= MaxCols - colM1 Then '最後n欄為人工填寫
                    tmpArr1(intJ) = ""
                Else
                    tmpArr1(intJ) = "" & rsRD.Fields(intJ - 1)
                End If
                '第一行資料底色設為灰色
                If intJ = 1 And (iRow Mod 2 = 1) Then
                    wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Interior.ColorIndex = 15  '底色:灰
                End If
            Next intJ
            wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value = tmpArr1
            wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).NumberFormat = "@"
            '備註+人工輸入：設定自動換行
            wksPoint1.Range(Pub_NumberToSystem26(MaxCols - colM1 - 1) & iRow & ":" & strExc(3) & iRow).WrapText = True
            '另外抓相對的專利/商標案件，若本身沒設相對的專利/商標互惠，還是要抓資料，給案量欄位設空白
            If "" & rsRD.Fields("C2") <> "" Then
                strCon1 = Mid(frm050408.strConSecond, 1, InStr(UCase(frm050408.strConSecond), "ORDER") - 1)
                strCon1 = strCon1 & " AND FC01=" & CNULL(Left(rsRD.Fields("C2"), 8))
                intI = 1
                Set rsAD = ClsLawReadRstMsg(intI, strCon1)
                If intI = 1 Then
                    iRow = iRow + 1
                    For intJ = 1 To MaxCols
                        If intJ < 5 Then '代理人資料
                            If m_bolAgent = False And (intJ = 2 Or intJ = 3) And InStr("" & rsAD.Fields("C2"), "-") > 0 Then
                               tmpArr1(intJ) = "" & rsAD.Fields(intJ - 1) '聯絡人：因為專利和商標設定的聯絡人可能不一致
                            Else
                               tmpArr1(intJ) = ""  '空白
                            End If
                        ElseIf intJ >= MaxCols - colM1 Then   '最後n欄為人工填寫
                            tmpArr1(intJ) = ""
                        Else
                            tmpArr1(intJ) = "" & rsAD.Fields(intJ - 1)
                        End If
                    Next intJ

                    wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value = tmpArr1
                    wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).NumberFormat = "@"
                    '備註+人工輸入：設定自動換行
                    wksPoint1.Range(Pub_NumberToSystem26(MaxCols - colM1 - 1) & iRow & ":" & strExc(3) & iRow).WrapText = True
                End If
            End If
            rsRD.MoveNext
            iRow = iRow + 1
       Loop
       wksPoint1.Range("A1").Select
       xlsPoint1.ActiveWindow.ScrollColumn = 5 '將非凍結欄位的捲軸拉回到最前面
       
       '判斷版本
       If Val(xlsPoint1.Version) < 12 Then
            xlsPoint1.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
       Else
            xlsPoint1.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
       End If

       xlsPoint1.Workbooks.Close
       xlsPoint1.Quit
       Set wksPoint1 = Nothing
       Set xlsPoint1 = Nothing
       'Modify by Amy 2021/06/22 原:strFileName 改中文字顯示
       MsgBox "Excel檔案產生完成！" & vbCrLf & "檔案位置：" & strExcelPathN & " " & Replace(strFileName, strExcelPath, "")
       
    End If
    
    Set rsRD = Nothing
    Set rsAD = Nothing
    Exit Sub
    
ErrHandle:
    If Err.Number <> 0 Then
        MsgBox "產生Excel失敗：" & vbCrLf & Err.Description, vbCritical
    End If
End Sub

'Added by Lydia 2022/05/12 (2022)調整版面
Private Sub ProcExcelSave_2022()
Dim strTitle() As String
Dim MaxCols As Integer '最大行位置
Dim strCon1 As String
Dim intJ As Integer
Dim rsRD As New ADODB.Recordset
Dim rsAD As New ADODB.Recordset
Dim xlsPoint1 As New Excel.Application
Dim wksPoint1 As New Worksheet
Dim strFileName As String '檔案名稱
Dim iRow As Integer
Dim xCols As Integer
Dim tmpArr1 As Variant, tmpArr2 As Variant
Dim m_bolAgent As Boolean
Dim strInputCol As String '第X行為人工填寫
Dim strGrp As String '換國家加上邊界

On Error GoTo ErrHandle
    
    If frm050408.cboTarget.ListIndex = 0 Then
        m_bolAgent = True
    Else
        m_bolAgent = False
    End If
    
    '設定抬頭為4列
    'B1=FCP,B2=CFP,B3=專利
    MaxCols = 28
    ReDim strTitle(1 To intMax) As String
    '位置|左側框線|第1行~第4行|欄寬
    strTitle(1) = "A|無|代理人名稱　（統計對象：" & frm050408.cboTarget & "）|空白|空白|M空白|32"
    strTitle(2) = "B|無|代理人編號|空白|空白|M空白|11"
    strTitle(3) = "C|無|聯絡人資訊|空白|空白|M空白|15"
    strTitle(4) = "D|無|國籍|空白|空白|M空白|7"
    strTitle(5) = "E|有|全部|B1|C1|收案量|9"
    strTitle(6) = "F|無|全部|B2|C2|收案量|9"
    strTitle(7) = "G|有|前年度|B1|C1|收案量|10"
    strTitle(8) = "H|無|前年度|B2|C2|給案量|10"
    strTitle(9) = "I|有|去年度|B1|C1|收案量|10"
    strTitle(10) = "J|無|去年度|B2|C2|給案量|10"
    strTitle(11) = "K|有|當年度1-6月|B1|C1|收案量|13"
    strTitle(12) = "L|無|當年度1-6月|B2|C2|給案量|13"
    strTitle(13) = "M|無|當年度1-6月|B3|C3|收給案量差|13"
    strTitle(14) = "N|有|當年度7-12月|B1|C1|收案量|13"
    strTitle(15) = "O|無|當年度7-12月|B2|C2|給案量|13"
    strTitle(16) = "P|無|當年度7-12月|B3|C3|收給案量差|13"
    strTitle(17) = "Q|有|該半年度建議@B2給案量|空白|空白|M空白|13"
    strTitle(18) = "R|有|現有備註|空白|空白|M空白|20"
    strTitle(19) = "S|有|請選擇：@維持/刪除/@給案量調整(減少/@增加)/新增|空白|空白|M空白|19"
    strTitle(20) = "T|無|隔半年度建議@B2給案量|空白|空白|M空白|15"
    strTitle(21) = "U|無|隔半年度備註|空白|空白|M空白|15"
    strTitle(22) = "V|無|提出年度|空白|空白|M空白|15"
    strTitle(23) = "W|無|提出部門|空白|空白|M空白|15"
    strTitle(24) = "X|無|提出人|空白|空白|M空白|15"
    strTitle(25) = "Y|無|空白|B1|意見|空白|30"
    strTitle(26) = "Z|無|空白|B2|意見|空白|30"
    strTitle(27) = "AA|無|空白|業拓|意見|空白|30"
    strTitle(28) = "AB|無|空白|互惠協調會|意見|空白|30"
    '第X欄為人工填寫
    strInputCol = ",S,T,U,Y,Z,AA,AB,"

    strCon1 = frm050408.strConFirst
    intI = 1
    Set rsRD = ClsLawReadRstMsg(intI, strCon1)
    If intI = 1 Then
        strFileName = strExcelPath & strSrvDate(1) & "_互惠代理人案件統計表" & frm050408.txtYear & "年" & IIf(frm050408.cboPeriod.ListIndex = 0, "上半年", "下半年") & IIf(frm050408.txtKind = "1", "專利", "商標") & IIf(m_bolAgent = True, "(代理人)", "(聯絡人)") & MsgText(43)
        If Dir(strFileName) <> "" Then
            '檢查檔案是否正在使用中
            If PUB_ChkFileOpening(strFileName) = True Then
                 MsgBox strFileName & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
                 Exit Sub
            End If
            Kill strFileName
        End If
        
        xlsPoint1.Workbooks.add
        xlsPoint1.Visible = False '預設不顯示
        iRow = 0
        xCols = 1
        xlsPoint1.SheetsInNewWorkbook = 3 '預設工作表數量
        Set wksPoint1 = xlsPoint1.Worksheets(1)
        xlsPoint1.Sheets(1).Select '選擇工作表
        '凍結窗格要在合併儲存格之前, 不然A欄會自動隱藏
        wksPoint1.Range("E5").Select
        xlsPoint1.ActiveWindow.FreezePanes = True
        '抬頭底色+字體設定
        wksPoint1.Range("E2:P2").Interior.ColorIndex = 15  '底色:灰
        wksPoint1.Range("S1:S4").Interior.ColorIndex = 6 '底色:黃
        wksPoint1.Range("T1:T4").Interior.ColorIndex = 6
        wksPoint1.Range("U1:U4").Interior.ColorIndex = 6

        '欄位抬頭
        For intI = 1 To MaxCols
           If frm050408.txtKind = "1" Then '專利為主
               strTitle(intI) = Replace(Replace(Replace(strTitle(intI), "B1", "FCP"), "B2", "CFP"), "B3", "專利")
               strTitle(intI) = Replace(Replace(Replace(Replace(strTitle(intI), "C1", "FCT"), "C2", "CFT"), "C3", "商標"), "B2", "CFP")
           ElseIf frm050408.txtKind = "2" Then  '商標為主
               strTitle(intI) = Replace(Replace(Replace(strTitle(intI), "B1", "FCT"), "B2", "CFT"), "B3", "商標")
               strTitle(intI) = Replace(Replace(Replace(Replace(strTitle(intI), "C1", "FCP"), "C2", "CFP"), "C3", "專利"), "B2", "CFT")
           End If
           strTitle(intI) = Replace(strTitle(intI), "@", vbCrLf)
           tmpArr1 = Empty
           tmpArr1 = Split(strTitle(intI), "|")
           strExc(3) = tmpArr1(0)
           iRow = 0
           For intJ = 1 To UBound(tmpArr1)
               If intJ = 1 Then '左側框線
                  If tmpArr1(intJ) = "有" Then
                      wksPoint1.Range(strExc(3) & ":" & strExc(3)).Select
                      With xlsPoint1.Selection.Borders(xlEdgeLeft)
                          .LineStyle = xlContinuous
                          .Weight = xlMedium
                      End With
                  End If
               ElseIf intJ > 1 And intJ < 6 Then
                  iRow = iRow + 1
                  strCon1 = Trim(Replace(tmpArr1(intJ), "空白", ""))
                  wksPoint1.Range(strExc(3) & iRow).Value = IIf(strCon1 = "M", "", strCon1)
               ElseIf intJ = 6 Then '欄寬
                  wksPoint1.Range(strExc(3) & ":" & strExc(3)).ColumnWidth = Val("" & tmpArr1(intJ))
                  wksPoint1.Range(strExc(3) & ":" & strExc(3)).NumberFormatLocal = "@"
                  If Asc(strExc(3)) >= Asc("E") And Asc(strExc(3)) <= Asc("Q") Then
                      wksPoint1.Range(strExc(3) & ":" & strExc(3)).HorizontalAlignment = xlCenter '數字,水平置中
                  End If
                  If tmpArr1(5) = "M空白" Then   '合併儲存格
                      xlsPoint1.Range(strExc(3) & iRow - 3 & ":" & strExc(3) & iRow).Select
                      With xlsPoint1.Selection
                           .HorizontalAlignment = xlCenter
                           .VerticalAlignment = xlCenter
                           .WrapText = True
                           .Orientation = 0
                           .AddIndent = False
                           .ShrinkToFit = False
                           .MergeCells = True
                       End With
                    Else '置中
                      xlsPoint1.Range(strExc(3) & iRow - 3 & ":" & strExc(3) & iRow).Select
                      With xlsPoint1.Selection
                           .HorizontalAlignment = xlCenter
                           .VerticalAlignment = xlCenter
                       End With
                    End If
               End If
           Next intJ
        Next intI
        
       '代理人編號,國籍：置中
        xlsPoint1.Range("B:B").HorizontalAlignment = xlCenter
        xlsPoint1.Range("D:D").HorizontalAlignment = xlCenter
       '備註+人工輸入：設定自動換行
       For intJ = 18 To MaxCols
          xlsPoint1.Range(Pub_NumberToSystem26(intJ) & ":" & Pub_NumberToSystem26(intJ)).HorizontalAlignment = xlCenter
          xlsPoint1.Range(Pub_NumberToSystem26(intJ) & ":" & Pub_NumberToSystem26(intJ)).VerticalAlignment = xlCenter
          xlsPoint1.Range(Pub_NumberToSystem26(intJ) & ":" & Pub_NumberToSystem26(intJ)).WrapText = True
       Next intJ
  
       iRow = iRow + 1
       ReDim tmpArr1(1 To MaxCols)
       strExc(3) = Pub_NumberToSystem26(MaxCols)
       tmpArr2 = Empty
       tmpArr2 = Split(strInputCol, ",")
       rsRD.MoveFirst
       Do While Not rsRD.EOF
            xCols = 0
            For intJ = 1 To MaxCols
                If InStr(strInputCol, "," & Pub_NumberToSystem26(intJ) & ",") > 0 Then   '排除第n欄為人工填寫
                    tmpArr1(intJ) = ""
                Else
                     tmpArr1(intJ) = "" & rsRD.Fields(xCols)
                     xCols = xCols + 1
                End If
                '第一行資料底色設為灰色
                If intJ = 1 And (iRow Mod 2 = 1) Then
                    wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Interior.ColorIndex = 15  '底色:灰
                End If
            Next intJ
            wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value = tmpArr1
            wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).NumberFormat = "@"
            '加上邊框
            If strGrp <> tmpArr1(4) Then
                wksPoint1.Range("A" & iRow & ":" & Pub_NumberToSystem26(MaxCols) & iRow).Select
                With xlsPoint1.Selection.Borders(xlEdgeTop)
                     .LineStyle = xlContinuous
                     .Weight = xlMedium
                End With
                strGrp = tmpArr1(4)
            End If
            '另外抓相對的專利/商標案件，若本身沒設相對的專利/商標互惠，還是要抓資料，給案量欄位設空白
            If "" & rsRD.Fields("C2") <> "" Then
                strCon1 = Mid(frm050408.strConSecond, 1, InStr(UCase(frm050408.strConSecond), "ORDER") - 1)
                strCon1 = strCon1 & " AND FC01=" & CNULL(Left(rsRD.Fields("C2"), 8))
                intI = 1
                Set rsAD = ClsLawReadRstMsg(intI, strCon1)
                If intI = 1 Then
                    xCols = 0
                    iRow = iRow + 1
                    For intJ = 1 To MaxCols
                        If intJ < 5 Then '代理人資料
                            If m_bolAgent = False And (intJ = 2 Or intJ = 3) And InStr("" & rsAD.Fields("C2"), "-") > 0 Then
                               tmpArr1(intJ) = "" & rsAD.Fields(xCols) '聯絡人：因為專利和商標設定的聯絡人可能不一致
                            Else
                               tmpArr1(intJ) = ""  '空白
                            End If
                            xCols = xCols + 1
                        Else
                           If InStr(strInputCol, "," & Pub_NumberToSystem26(intJ) & ",") > 0 Then   '排除第n欄為人工填寫
                               tmpArr1(intJ) = ""
                           Else
                                tmpArr1(intJ) = "" & rsAD.Fields(xCols)
                                xCols = xCols + 1
                           End If
                        End If
                    Next intJ

                    wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value = tmpArr1
                    wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).NumberFormat = "@"
                End If
            End If
            rsRD.MoveNext
            iRow = iRow + 1
       Loop
       wksPoint1.Range("A1").Select
       xlsPoint1.ActiveWindow.ScrollColumn = 5 '將非凍結欄位的捲軸拉回到最前面
       
       '判斷版本
       If Val(xlsPoint1.Version) < 12 Then
            xlsPoint1.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
       Else
            xlsPoint1.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
       End If

       xlsPoint1.Workbooks.Close
       xlsPoint1.Quit
       Set wksPoint1 = Nothing
       Set xlsPoint1 = Nothing
       MsgBox "Excel檔案產生完成！" & vbCrLf & "檔案位置：" & strExcelPathN & " " & Replace(strFileName, strExcelPath, "")
       
    End If
    
    Set rsRD = Nothing
    Set rsAD = Nothing
    Exit Sub
    
ErrHandle:
    If Err.Number <> 0 Then
        MsgBox "產生Excel失敗：" & vbCrLf & Err.Description, vbCritical
    End If
End Sub

'Added by Lydia 2023/09/12 (2023)調整版面
Private Sub ProcExcelSave_2023()
Dim strTitle() As String
Dim MaxCols As Integer '最大行位置
Dim strCon1 As String, strCon2 As String
Dim intJ As Integer
Dim rsRD As New ADODB.Recordset
Dim rsAD As New ADODB.Recordset
Dim xlsPoint1 As New Excel.Application
Dim wksPoint1 As New Worksheet
Dim strFileName As String '檔案名稱
Dim iRow As Integer
Dim xCols As Integer
Dim tmpArr1 As Variant
Dim m_bolAgent As Boolean
Dim strInputCol As String '第X行為人工填寫
Dim strGrp As String '換國家加上邊界
Dim strFCname As String '互惠代理人的名稱
Dim iCall As Integer, intAdd As Integer '分兩個工作表,欄位數不同

On Error GoTo ErrHandle

    If frm050408.cboTarget.ListIndex = 0 Then
        m_bolAgent = True
    Else
        m_bolAgent = False
    End If

    '取得關聯企業=> 模組化
    Call Pub_GetFCfrm050408(strUserNum, frm050408.txtKind, "1", frm050408.txtYear, frm050408.cboPeriod.ListIndex + 1, IIf(frm050408.cboTarget.ListIndex = 0, True, False))

    strSql = "delete from R050408_2023 where id='" & strUserNum & "' "
    cnnConnection.Execute strSql
    
    '資料全部丟暫存檔
      'create table R050408_2023 (
      'ID VARCHAR2(6 CHAR),
      'KIND VARCHAR2(6 CHAR), ---1.畫面輸入案件類別(專利/商標,2.另一種案件類別(商標/專利)
      'AN1 VARCHAR2(1 CHAR), ---篩選(關聯企業)
      'AN2 VARCHAR2(6 CHAR), ---篩選(同族企業)=>原本互惠代理人的排序
      'NA01 VARCHAR2(3 CHAR),
      'C1 VARCHAR2(130 CHAR), ---代理人名稱(放大欄位長度預防程式出錯)
      'C2 VARCHAR2(20 CHAR),  ---代理人編號-接洽人編號
      'C3 VARCHAR2(100 char), ---接洽人名稱
      'NA03 VARCHAR2(20 CHAR),
      'FC_TOT Number(6, 0), ---全部統計(FC)
      'CF_TOT NUMBER(6,0),
      'FC_L2 Number(6, 0), ---前年度統計(FC)
      'CF_L2 NUMBER(6,0),
      'FC_L1 Number(6, 0), ---去年度統計(FC)
      'CF_L1 NUMBER(6,0),
      'FC_C1 NUMBER(6,0),  ---當年度1-6月統計(FC)
      'CF_C1 NUMBER(6,0),
      'DIFF01 NUMBER(6,0),
      'FC_C2 NUMBER(6,0),  ---當年度7-12月統計(FC)
      'CF_C2 NUMBER(6,0),
      'DIFF02 NUMBER(6,0),
      'FC07 Number(6, 0), ---給案量
      'FC08 VARCHAR2(1000 CHAR),  ---備註(放大欄位長度預防程式出錯)
      'FC16 VARCHAR2(100 CHAR),    ---提出年度(放大欄位長度預防程式出錯)
      'FC17DEPT VARCHAR2(20 CHAR), ---提出人員部門
      'FC17NAME VARCHAR2(12 CHAR) ---提出人員名稱
      ');
    
    '1.畫面輸入案件類別(專利/商標)
    '排除指定日期條件=> , frm050408.Text1, frm050408.Text2
    Call Pub_GetSqlfrm050408(strUserNum, frm050408.txtKind, "0", frm050408.txtYear, frm050408.cboPeriod.ListIndex + 1, m_bolAgent, strExc(3), strExc(4), strSrvDate(1))
    strCon2 = "INSERT INTO R050408_2023 (ID,KIND,AN1,AN2,NA01,C1,C2,C3,NA03,FC_TOT,CF_TOT,FC_L2,CF_L2,FC_L1,CF_L1,FC_C1,CF_C1,DIFF01,FC_C2,CF_C2,DIFF02,FC07,FC08,FC16,FC17DEPT,FC17NAME) "
    strCon1 = Mid(strExc(3), 1, InStr(UCase(strExc(3)), "ORDER") - 1)
    strSql = strCon2 & Replace(strCon1, "SELECT NVL(FA05,NVL(FA06,FA04))", "SELECT '" & strUserNum & "','" & IIf(frm050408.txtKind = "1", "CFP", "CFT") & "','0', ROWSEQ, SUBSTR(NA01,1,3) NA01, NVL(FA05,NVL(FA06,FA04))")
    cnnConnection.Execute strSql, intI
    '2.另一種案件類別(商標/專利)
    strCon1 = Mid(strExc(4), 1, InStr(UCase(strExc(4)), "ORDER") - 1)
    strSql = strCon2 & Replace(strCon1, "SELECT NVL(FA05,NVL(FA06,FA04))", "SELECT '" & strUserNum & "','" & IIf(frm050408.txtKind = "1", "CFT", "CFP") & "','0', ROWSEQ, SUBSTR(NA01,1,3) NA01, NVL(FA05,NVL(FA06,FA04))")
    cnnConnection.Execute strSql, intI
    '---關聯企業
    Call Pub_GetSqlfrm050408(strUserNum, frm050408.txtKind, "1", frm050408.txtYear, frm050408.cboPeriod.ListIndex + 1, m_bolAgent, strExc(3), strExc(4), strSrvDate(1))
    strCon1 = Mid(strExc(3), 1, InStr(UCase(strExc(3)), "ORDER") - 1)
    strSql = strCon2 & Replace(strCon1, "SELECT NVL(FA05,NVL(FA06,FA04))", "SELECT '" & strUserNum & "','" & IIf(frm050408.txtKind = "1", "CFP", "CFT") & "','1', ROWSEQ, SUBSTR(NA01,1,3) NA01, NVL(FA05,NVL(FA06,FA04))")
    cnnConnection.Execute strSql, intI
    strCon1 = Mid(strExc(4), 1, InStr(UCase(strExc(4)), "ORDER") - 1)
    strSql = strCon2 & Replace(strCon1, "SELECT NVL(FA05,NVL(FA06,FA04))", "SELECT '" & strUserNum & "','" & IIf(frm050408.txtKind = "1", "CFT", "CFP") & "','1', ROWSEQ, SUBSTR(NA01,1,3) NA01, NVL(FA05,NVL(FA06,FA04))")
    cnnConnection.Execute strSql, intI

    '設定抬頭為3行 ; (專利)B1=FCP,B2=CFP,B3=FCP-CFP ; C1~C2 (若B1=專利,C1~C3為商標)
    MaxCols = 28
    ReDim strTitle(1 To MaxCols) As String
    '(位置改用index計算)左側框線|第1行~第3行|欄寬
    strTitle(1) = "無|空白|代理人名稱|（統計對象：" & frm050408.cboTarget & "）|32"
    strTitle(2) = "無|空白|代理人編號|空白|11"
    strTitle(3) = "無|空白|聯絡人資訊|空白|15"
    strTitle(4) = "無|空白|國籍|空白|7"
    strTitle(5) = "有|全部|B1|C1|9"
    strTitle(6) = "無|全部|B2|C2|9"
    strTitle(7) = "有|前年度|B1|C1|10"
    strTitle(8) = "無|前年度|B2|C2|10"
    strTitle(9) = "有|去年度|B1|C1|10"
    strTitle(10) = "無|去年度|B2|C2|10"
    strTitle(11) = "有|當年度1-6月|B1|C1|13"
    strTitle(12) = "無|當年度1-6月|B2|C2|13"
    strTitle(13) = "無|當年度1-6月|B3|C3|13"
    strTitle(14) = "有|當年度7-12月|B1|C1|13"
    strTitle(15) = "無|當年度7-12月|B2|C2|13"
    strTitle(16) = "無|當年度7-12月|B3|C3|13"
    strTitle(17) = "有|該半年度建議|B2|空白|13"
    strTitle(18) = "有|空白|現有備註|空白|20"
    strTitle(19) = "無|空白|提出年度|空白|15"
    strTitle(20) = "無|空白|提出部門|空白|15"
    strTitle(21) = "無|空白|提出人|空白|15"
    strTitle(22) = "有|請填寫：|維持/刪除/新增|給案量調整(減少/增加)|25"
    strTitle(23) = "無|次半年度建議|B2|空白|15"
    strTitle(24) = "無|次半年度備註|空白|M空白|15"
    strTitle(25) = "無|空白|B1意見|空白|30"
    strTitle(26) = "無|空白|B2意見|空白|30"
    strTitle(27) = "無|空白|業拓意見|空白|30"
    strTitle(28) = "無|空白|互惠協調會意見|空白|30"
    '第X欄為人工填寫
    strInputCol = ",22,23,24,25,26,27,28,"
    
    For iCall = 1 To 2 '1:含關聯企業編號, 2=不含關聯企業
       strCon1 = "SELECT C1,C2,C3,NA03,FC_TOT,CF_TOT,FC_L2,CF_L2,FC_L1,CF_L1,FC_C1,CF_C1,DIFF01,FC_C2,CF_C2,DIFF02,FC07,FC08,FC16,FC17DEPT,FC17NAME,KIND,AN2,NA01" & _
                 " FROM R050408_2023 WHERE ID='" & strUserNum & "' AND AN1='0'" & _
                 " ORDER BY TO_NUMBER(AN2) ASC,DECODE(KIND,'" & IIf(frm050408.txtKind = "1", "CFP", "CFT") & "','1','2') ASC "
       intI = 1
       Set rsRD = ClsLawReadRstMsg(intI, strCon1)
       If intI = 1 Then
            'Added by Lydia 2025/07/30 查詢印表記錄檔欄位
            If iCall = 2 Then
               ClearQueryLog ("frm050408")
               pub_QL05 = frm050408.m_strQL05 & ";產生Excel"
               InsertQueryLog (rsRD.RecordCount)
            End If
            'end 2025/07/30
           If strFileName = "" Then
              'Modified by Lydia 2025/06/06
              'strFileName = strExcelPath & strSrvDate(1) & "_互惠代理人案件統計表" & frm050408.txtYear & "年" & IIf(frm050408.cboPeriod.ListIndex = 0, "上半年", "下半年") & IIf(frm050408.txtKind = "1", "專利", "商標") & IIf(m_bolAgent = True, "(代理人)", "(聯絡人)") & MsgText(43)
              strFileName = strExcelPath & strSrvDate(1) & "_" & IIf(frm050408.Tag <> "", "互惠期間統計表", "互惠代理人案件統計表") & frm050408.txtYear & "年" & IIf(frm050408.cboPeriod.ListIndex = 0, "上半年", "下半年") & IIf(frm050408.txtKind = "1", "專利", "商標") & IIf(m_bolAgent = True, "(代理人)", "(聯絡人)") & IIf(frm050408.Tag <> "", "_" & frm050408.Tag, "") & MsgText(43)
              If Dir(strFileName) <> "" Then
                 '檢查檔案是否正在使用中
                 If PUB_ChkFileOpening(strFileName) = True Then
                     MsgBox strFileName & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
                     Exit Sub
                 End If
                 Kill strFileName
              End If
              xlsPoint1.Visible = False '預設不顯示
              iRow = 0
              xlsPoint1.SheetsInNewWorkbook = 3 '預設工作表數量
              xlsPoint1.Workbooks.add
           End If

           Set wksPoint1 = xlsPoint1.Worksheets(iCall)
           xlsPoint1.Sheets(iCall).Select '選擇工作表
           xlsPoint1.ActiveWindow.Zoom = 60 '顯示縮小到60%
           
           If frm050408.txtKind = "1" Then
              strExc(1) = "專利"
           Else
              strExc(1) = "商標"
           End If
           xlsPoint1.Worksheets(iCall).Name = IIf(iCall = 2, strExc(1), strExc(1) & " (含關聯企業編號且前年度至今有案件往來者)") '工作表名稱
           xlsPoint1.ActiveWindow.FreezePanes = False '取消凍結窗格
           xlsPoint1.ActiveWindow.SplitColumn = IIf(iCall = 2, 4, 6)
           xlsPoint1.ActiveWindow.SplitRow = 3
           xlsPoint1.ActiveWindow.FreezePanes = True '凍結窗格
            
           If iCall = 1 Then '含關聯企業編號: 前方多2欄
              wksPoint1.Range("A:A").ColumnWidth = 12
              wksPoint1.Range("A:A").Font.Name = "微軟正黑體"
              wksPoint1.Range("A1").Value = "篩選" & vbCrLf & "(關聯企業)"
              xlsPoint1.Range("A1:A3").Select
              With xlsPoint1.Selection
                 .HorizontalAlignment = xlCenter
                 .VerticalAlignment = xlCenter
                 .WrapText = True
                 .Orientation = 0
                 .AddIndent = False
                 .ShrinkToFit = False
                 .MergeCells = True
              End With
              wksPoint1.Range("B:B").ColumnWidth = 12
              wksPoint1.Range("B:B").Font.Name = "微軟正黑體"
              wksPoint1.Range("B1").Value = "篩選" & vbCrLf & "(同族企業)"
              xlsPoint1.Range("B1:B3").Select
              With xlsPoint1.Selection
                 .HorizontalAlignment = xlCenter
                 .VerticalAlignment = xlCenter
                 .WrapText = True
                 .Orientation = 0
                 .AddIndent = False
                 .ShrinkToFit = False
                 .MergeCells = True
              End With
              intAdd = 2
           Else
              intAdd = 0
           End If
           
           '抬頭底色+字體設定
           wksPoint1.Range(Chr(Asc("A") + intAdd) & "2:" & Chr(Asc("U") + intAdd) & "2").Interior.ColorIndex = 15 '年度統計數字欄位底色:灰
           '不含關聯企業: V~最後
           wksPoint1.Range(Pub_NumberToSystem26(22 + intAdd) & "1:" & Pub_NumberToSystem26(MaxCols + intAdd) & "3").Interior.ColorIndex = 6  '底色:黃
           
           '欄位抬頭
           For intI = 1 To MaxCols
              strExc(4) = strTitle(intI)
              If frm050408.txtKind = "1" Then '專利為主
                  strExc(4) = Replace(Replace(Replace(strExc(4), "B1", "FCP"), "B2", "CFP"), "B3", "FCP-CFP")
                  strExc(4) = Replace(Replace(Replace(Replace(strExc(4), "C1", "FCT"), "C2", "CFT"), "C3", "FCT-CFT"), "B2", "CFP")
              ElseIf frm050408.txtKind = "2" Then  '商標為主
                  strExc(4) = Replace(Replace(Replace(strExc(4), "B1", "FCT"), "B2", "CFT"), "B3", "FCT-CFT")
                  strExc(4) = Replace(Replace(Replace(Replace(strExc(4), "C1", "FCP"), "C2", "CFP"), "C3", "FCP-CFP"), "B2", "CFT")
              End If
              strExc(4) = Replace(strExc(4), "@", vbCrLf)
              tmpArr1 = Empty
              tmpArr1 = Split(strExc(4), "|")
              strExc(3) = Pub_NumberToSystem26(intI + intAdd)  '(位置改用index計算)
              iRow = 0
              For intJ = 0 To UBound(tmpArr1)
                  If intJ = 0 Then '左側框線
                     If tmpArr1(intJ) = "有" Then
                         wksPoint1.Range(strExc(3) & ":" & strExc(3)).Select
                         With xlsPoint1.Selection.Borders(xlEdgeLeft)
                             .LineStyle = xlContinuous
                             .Weight = xlMedium
                         End With
                     End If
                     wksPoint1.Range(strExc(3) & ":" & strExc(3)).Font.Name = "微軟正黑體"
                  ElseIf intJ < 4 Then
                     iRow = iRow + 1
                     strCon1 = Trim(Replace(tmpArr1(intJ), "空白", ""))
                     wksPoint1.Range(strExc(3) & iRow).Value = IIf(strCon1 = "M", "", strCon1)
                  ElseIf intJ = 4 Then '欄寬
                     wksPoint1.Range(strExc(3) & ":" & strExc(3)).ColumnWidth = Val("" & tmpArr1(intJ))
                     If tmpArr1(3) = "M空白" Then   '合併儲存格
                         xlsPoint1.Range(strExc(3) & iRow - 2 & ":" & strExc(3) & iRow).Select
                         With xlsPoint1.Selection
                           .HorizontalAlignment = xlCenter
                           .VerticalAlignment = xlCenter
                           .WrapText = True
                           .Orientation = 0
                           .AddIndent = False
                           .ShrinkToFit = False
                           .MergeCells = True
                         End With
                     Else '置中
                         xlsPoint1.Range(strExc(3) & iRow - 2 & ":" & strExc(3) & iRow).Select
                         With xlsPoint1.Selection
                           .HorizontalAlignment = xlCenter
                           .VerticalAlignment = xlCenter
                         End With
                     End If
                  End If
              Next intJ
           Next intI
          '先丟值格式後調整: 代理人編號,國籍：置中
          If intAdd > 0 Then
             xlsPoint1.Range("A:A").HorizontalAlignment = xlCenter
             xlsPoint1.Range("B:B").HorizontalAlignment = xlCenter
          End If
          xlsPoint1.Range(Chr(Asc("B") + intAdd) & ":" & Chr(Asc("B") + intAdd)).HorizontalAlignment = xlCenter
          xlsPoint1.Range(Chr(Asc("D") + intAdd) & ":" & Chr(Asc("D") + intAdd)).HorizontalAlignment = xlCenter
          '數字,水平置中(5~17) 備註+人工輸入：設定自動換行
          For intJ = 5 To MaxCols
             If intJ = 18 Then '「現有備註」欄，除了表頭外，下方調整為靠左對齊
                xlsPoint1.Range(Pub_NumberToSystem26(intJ + intAdd) & ":" & Pub_NumberToSystem26(intJ + intAdd)).HorizontalAlignment = xlLeft
                xlsPoint1.Range(Pub_NumberToSystem26(intJ + intAdd) & ":" & Pub_NumberToSystem26(intJ + intAdd)).VerticalAlignment = xlTop
                xlsPoint1.Range(Pub_NumberToSystem26(intJ + intAdd) & "1:" & Pub_NumberToSystem26(intJ + intAdd) & "3").HorizontalAlignment = xlCenter
             Else
                xlsPoint1.Range(Pub_NumberToSystem26(intJ + intAdd) & ":" & Pub_NumberToSystem26(intJ + intAdd)).HorizontalAlignment = xlCenter
                xlsPoint1.Range(Pub_NumberToSystem26(intJ + intAdd) & ":" & Pub_NumberToSystem26(intJ + intAdd)).VerticalAlignment = xlCenter
             End If
             If intJ > 17 Then
                xlsPoint1.Range(Pub_NumberToSystem26(intJ + intAdd) & ":" & Pub_NumberToSystem26(intJ + intAdd)).WrapText = True
             End If
          Next intJ
          
          iRow = iRow + 1
          ReDim tmpArr1(1 To MaxCols + intAdd)
          strExc(3) = Pub_NumberToSystem26(MaxCols + intAdd)
          rsRD.MoveFirst
          Do While Not rsRD.EOF
'------------互惠代理人的資料處理
             xCols = 0
             If intAdd > 0 Then
                tmpArr1(1) = Val(0)
                tmpArr1(2) = Val(rsRD.Fields("AN2"))
             End If
             For intJ = 1 To MaxCols
                If iRow Mod 2 = 1 And (intJ >= 18 Or intJ <= 4) Then    '基本資料只在第一行(偶數)顯示
                    If m_bolAgent = False And (intJ = 2 Or intJ = 3) And InStr("" & rsRD.Fields("C2"), "-") > 0 Then
                       tmpArr1(intJ + intAdd) = "" & rsRD.Fields(xCols)  '聯絡人：因為專利和商標設定的聯絡人可能不一致
                    Else
                       tmpArr1(intJ + intAdd) = ""
                    End If
                    xCols = xCols + 1
                ElseIf InStr(strInputCol, "," & Format(intJ, "00") & ",") > 0 Then  '排除第n欄為人工填寫
                    tmpArr1(intJ + intAdd) = ""
                ElseIf intJ >= 5 And intJ <= 17 Then '數值
                    tmpArr1(intJ + intAdd) = Val("" & rsRD.Fields(xCols))
                    xCols = xCols + 1
                Else
                    tmpArr1(intJ + intAdd) = "" & rsRD.Fields(xCols)
                    xCols = xCols + 1
                End If
                '第一行(偶數)資料底色設為灰色
                If intJ = 1 And iRow Mod 2 = 0 Then
                    wksPoint1.Range(Chr(Asc("A") + intAdd) & iRow & ":" & strExc(3) & iRow).Interior.ColorIndex = 15   '底色:灰
                End If
             Next intJ
             wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value = tmpArr1
             '當儲存格格式為通用格式，選取儲存格的值再代入到選取的儲存格=>自動化為數值欄位
             wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value = wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value
             strFCname = "" & rsRD.Fields("C1")
             '不同國家加上邊框
             If strGrp <> tmpArr1(4 + intAdd) And iRow Mod 2 = 0 Then
                wksPoint1.Range("A" & iRow & ":" & Pub_NumberToSystem26(MaxCols + intAdd) & iRow).Select
                With xlsPoint1.Selection.Borders(xlEdgeTop)
                   .LineStyle = xlContinuous
                   .Weight = xlMedium
                End With
                strGrp = tmpArr1(4 + intAdd)
             End If
'------------關聯企業有案件數才列出(印完互惠2筆資料)
             If iCall = 1 And iRow Mod 2 = 1 Then
                strCon1 = "SELECT C1,C2,C3,NA03,FC_TOT,CF_TOT,FC_L2,CF_L2,FC_L1,CF_L1,FC_C1,CF_C1,DIFF01,FC_C2,CF_C2,DIFF02,FC07,FC08,FC16,FC17DEPT,FC17NAME,KIND,AN2,NA01" & _
                          " FROM R050408_2023 WHERE ID='" & strUserNum & "' AND AN1='1' AND AN2 = '" & rsRD.Fields("AN2") & "' AND C2 IN (SELECT C2 FROM (" & _
                          " SELECT C2,SUM(FC_TOT+CF_TOT+FC_L2+CF_L2+FC_L1+CF_L1+FC_C1+CF_C1+FC_C2+CF_C2) TOT1 FROM R050408_2023 WHERE ID='" & strUserNum & "' AND AN1='1' AND AN2 = '" & rsRD.Fields("AN2") & "' GROUP BY C2" & _
                          " ) WHERE TOT1 <> 0) ORDER BY C2 ASC,DECODE(KIND,'" & IIf(frm050408.txtKind = "1", "CFP", "CFT") & "','1','2') ASC"
                intI = 1
                Set rsAD = ClsLawReadRstMsg(intI, strCon1)
                If intI = 1 Then
                   rsAD.MoveFirst
                   Do While Not rsAD.EOF
                      iRow = iRow + 1
                      tmpArr1(1) = "1"
                      tmpArr1(2) = "" & rsRD.Fields("AN2")
                      xCols = 0
                      For intJ = 1 To MaxCols
                         If iRow Mod 2 = 1 And (intJ >= 18 Or intJ <= 4) Then    '基本資料只在第一行(偶數)顯示
                            tmpArr1(intJ + intAdd) = ""
                            If intJ <= 4 Then
                               xCols = xCols + 1
                            End If
                         ElseIf InStr(strInputCol, "," & Format(intJ, "00") & ",") > 0 Then  '排除第n欄為人工填寫
                            tmpArr1(intJ + intAdd) = ""
                         ElseIf intJ >= 5 And intJ <= 17 Then '數值
                            tmpArr1(intJ + intAdd) = Val("" & rsAD.Fields(xCols))
                            xCols = xCols + 1
                         Else
                            tmpArr1(intJ + intAdd) = "" & rsAD.Fields(xCols)
                            xCols = xCols + 1
                         End If
                         '第一行(偶數)資料底色設為綠色
                         If intJ = 1 And iRow Mod 2 = 0 Then
                            wksPoint1.Range(Chr(Asc("A") + intAdd) & iRow & ":" & strExc(3) & iRow).Interior.ColorIndex = 35   '底色:綠色
                         End If
                      Next intJ
                      wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value = tmpArr1
                      '當儲存格格式為通用格式，選取儲存格的值再代入到選取的儲存格=>自動化為數值欄位
                      wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value = wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value
                      rsAD.MoveNext
                   Loop
                   '互惠+關聯的加總
                   strCon1 = "SELECT KIND,SUM(FC_TOT) FCTOT,SUM(CF_TOT) CF_TOT,SUM(FC_L2) FC_L2,SUM(CF_L2) CF_L2,SUM(FC_L1) FC_L1,SUM(CF_L1) CF_L1 " & _
                             ",SUM(FC_C1) FC_C1,SUM(CF_C1) CF_C1,SUM(DIFF01) DIFF01,SUM(FC_C2) FC_C2,SUM(CF_C2) CF_C2,SUM(DIFF02) DIFF02,SUM(FC07) FC07 " & _
                             "FROM R050408_2023 WHERE ID='" & strUserNum & "' AND AN2 = '" & rsRD.Fields("AN2") & "' GROUP BY KIND ORDER BY DECODE(KIND,'" & IIf(frm050408.txtKind = "1", "CFP", "CFT") & "','1','2') "
                   intI = 1
                   Set rsAD = ClsLawReadRstMsg(intI, strCon1)
                   If intI = 1 Then
                      rsAD.MoveFirst
                      Do While Not rsAD.EOF
                         iRow = iRow + 1
                         tmpArr1(1) = "2"
                         tmpArr1(2) = "" & rsRD.Fields("AN2")
                         xCols = 0
                         For intJ = 1 To MaxCols
                            If intJ > 17 Then   '排除第n欄為人工填寫
                               tmpArr1(intJ + intAdd) = ""
                            ElseIf intJ >= 5 And intJ <= 17 Then '數值
                               tmpArr1(intJ + intAdd) = Val("" & rsAD.Fields(xCols))
                               xCols = xCols + 1
                            Else
                               If intJ = 1 Then
                                  '第一行(偶數)資料底色設為淡橘色
                                  If iRow Mod 2 = 0 Then
                                     tmpArr1(intJ + intAdd) = strFCname
                                     tmpArr1(intJ + intAdd + 1) = "全部"
                                     tmpArr1(intJ + intAdd + 2) = ""
                                     tmpArr1(intJ + intAdd + 3) = ""
                                     wksPoint1.Range(Chr(Asc("A") + intAdd) & iRow & ":" & strExc(3) & iRow).Interior.ColorIndex = 40        '底色:淡橘色
                                  Else
                                     tmpArr1(intJ + intAdd) = ""
                                     tmpArr1(intJ + intAdd + 1) = ""
                                     tmpArr1(intJ + intAdd + 2) = ""
                                     tmpArr1(intJ + intAdd + 3) = ""
                                  End If
                                  xCols = xCols + 1
                                  intJ = 4
                               End If
                            End If
                         Next intJ
                         wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value = tmpArr1
                         '當儲存格格式為通用格式，選取儲存格的值再代入到選取的儲存格=>自動化為數值欄位
                         wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value = wksPoint1.Range("A" & iRow & ":" & strExc(3) & iRow).Value
                         rsAD.MoveNext
                      Loop
                   End If
                End If
             End If
             rsRD.MoveNext
             iRow = iRow + 1
          Loop
          '調整為能使文字全部顯示之欄寬: 代理人名稱、代理人編號、聯絡人資訊、國籍、該半年度建議
          For intJ = 1 To 4 + intAdd
             wksPoint1.Columns(Pub_NumberToSystem26(intJ) & ":" & Pub_NumberToSystem26(intJ)).EntireColumn.AutoFit
          Next intJ
          wksPoint1.Columns(Pub_NumberToSystem26(17 + intAdd) & ":" & Pub_NumberToSystem26(17 + intAdd)).EntireColumn.AutoFit
          wksPoint1.Range("3:3").RowHeight = 16.2
          xlsPoint1.ActiveWindow.ScrollRow = 1
          xlsPoint1.ActiveWindow.ScrollColumn = 5 + intAdd '將非凍結欄位的捲軸拉回到最前面
       End If
    Next iCall
          
    Set wksPoint1 = xlsPoint1.Worksheets(3)
    xlsPoint1.Sheets(3).Select '選擇工作表
    xlsPoint1.Worksheets(3).Name = "說明"
    wksPoint1.Range("A1").Value = "本報表採用廣義新案數定義"
    tmpArr1 = Array("定義", "", "專利", "", "商標")
    wksPoint1.Range("A3:E3").Value = tmpArr1
    tmpArr1 = Array("狹義新案數", "101", "發明申請", "101", "申請")
    wksPoint1.Range("A4:E4").Value = tmpArr1
    tmpArr1 = Array("", "102", "新型申請", "", "")
    wksPoint1.Range("A5:E5").Value = tmpArr1
    tmpArr1 = Array("", "103", "設計申請", "", "")
    wksPoint1.Range("A6:E6").Value = tmpArr1
    tmpArr1 = Array("廣義新案數", "101", "發明申請", "101", "申請")
    wksPoint1.Range("A7:E7").Value = tmpArr1
    tmpArr1 = Array("", "102", "新型申請", "", "")
    wksPoint1.Range("A8:E8").Value = tmpArr1
    tmpArr1 = Array("", "103", "設計申請", "", "")
    wksPoint1.Range("A9:E9").Value = tmpArr1
    tmpArr1 = Array("", "125", "衍生設計申請", "", "")
    wksPoint1.Range("A10:E10").Value = tmpArr1
    tmpArr1 = Array("", "301~306", "改請", "", "")
    wksPoint1.Range("A11:E11").Value = tmpArr1
    tmpArr1 = Array("", "307", "分割", "", "")
    wksPoint1.Range("A12:E12").Value = tmpArr1

    For intJ = 1 To 5
       strExc(3) = Pub_NumberToSystem26(intJ)
       wksPoint1.Range(strExc(3) & ":" & strExc(3)).Font.Name = "微軟正黑體"
       Select Case intJ
          Case 1
             wksPoint1.Range(strExc(3) & ":" & strExc(3)).ColumnWidth = 28
             '狹義新案數
             xlsPoint1.Range(strExc(3) & "4:" & strExc(3) & "6").Select
             With xlsPoint1.Selection
                 .HorizontalAlignment = xlLeft
                 .VerticalAlignment = xlTop
                 .WrapText = True
                 .Orientation = 0
                 .AddIndent = False
                 .ShrinkToFit = False
                 .MergeCells = True
                 .Font.Bold = True
             End With
             '廣義新案數
             xlsPoint1.Range(strExc(3) & "7:" & strExc(3) & "12").Select
             With xlsPoint1.Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
                .Font.Bold = True
             End With
             wksPoint1.Range(strExc(3) & "3:" & strExc(3) & "12").Borders.LineStyle = xlContinuous
             wksPoint1.Range(strExc(3) & "3:" & strExc(3) & "12").Borders.Weight = xlThin
             wksPoint1.Range(strExc(3) & "3:" & strExc(3) & "12").HorizontalAlignment = xlRight
          Case 2, 4 '案件性質CP10
             If intJ = 2 Then
                wksPoint1.Range(strExc(3) & ":" & strExc(3)).ColumnWidth = 10
             Else
                wksPoint1.Range(strExc(3) & ":" & strExc(3)).ColumnWidth = 7
             End If
             wksPoint1.Range(strExc(3) & "3:" & strExc(3) & "12").HorizontalAlignment = xlRight
          Case 3, 5 '案件性質名稱
             If intJ = 3 Then
                wksPoint1.Range(strExc(3) & ":" & strExc(3)).ColumnWidth = 16
             Else
                wksPoint1.Range(strExc(3) & ":" & strExc(3)).ColumnWidth = 12
             End If
             wksPoint1.Range(strExc(3) & "3:" & strExc(3) & "12").HorizontalAlignment = xlRight
             xlsPoint1.Range(Pub_NumberToSystem26(intJ - 1) & "3:" & strExc(3) & "3").Select
             With xlsPoint1.Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
                .Font.Bold = True
                .HorizontalAlignment = xlRight
             End With
             wksPoint1.Range(Pub_NumberToSystem26(intJ - 1) & "3:" & strExc(3) & "3").Borders.LineStyle = xlContinuous
             wksPoint1.Range(Pub_NumberToSystem26(intJ - 1) & "3:" & strExc(3) & "3").Borders.Weight = xlThin
             With wksPoint1.Range(Pub_NumberToSystem26(intJ - 1) & "4:" & strExc(3) & "6")
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin '細線
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThin
             End With
             With wksPoint1.Range(Pub_NumberToSystem26(intJ - 1) & "7:" & strExc(3) & "12")
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin '細線
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThin
             End With
       End Select
    Next intJ
    xlsPoint1.Sheets(1).Select '選擇工作表
    'Added by Lydia 2025/06/06
    If frm050408.Tag <> "" Then
       '刪除前2欄
       xlsPoint1.Range("A:B").Select
       xlsPoint1.Selection.EntireColumn.Delete
       '刪除建議給案量以後的所有欄位
       xlsPoint1.Range("Q:AB").Select
       xlsPoint1.Selection.EntireColumn.Delete
       '刪除不含關聯企業的工作表
       xlsPoint1.Sheets(1).Select
       xlsPoint1.Application.DisplayAlerts = False
       xlsPoint1.Worksheets(2).Delete
       xlsPoint1.Application.DisplayAlerts = True
       xlsPoint1.Sheets(1).Select
    End If
    'end 2025/06/06
    xlsPoint1.Range("A1").Select
    
    '判斷版本
    If Val(xlsPoint1.Version) < 12 Then
       xlsPoint1.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
    Else
       xlsPoint1.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
    End If

    xlsPoint1.Workbooks.Close
    xlsPoint1.Quit
    Set wksPoint1 = Nothing
    Set xlsPoint1 = Nothing
    MsgBox "Excel檔案產生完成！" & vbCrLf & "檔案位置：" & strExcelPathN & " " & Replace(strFileName, strExcelPath, "")
    
    Set rsRD = Nothing
    Set rsAD = Nothing
    Exit Sub
    
ErrHandle:
    If Err.Number <> 0 Then
        MsgBox "產生Excel失敗：" & vbCrLf & Err.Description, vbCritical
    End If
End Sub


