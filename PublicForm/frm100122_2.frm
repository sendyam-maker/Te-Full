VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100122_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員收/發文量比較查詢"
   ClientHeight    =   5712
   ClientLeft      =   1860
   ClientTop       =   3588
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5712
   ScaleWidth      =   9300
   Begin VB.CommandButton cmdOK 
      Caption         =   "開啟Word(&W)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   5220
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   10
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印(&P)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6480
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   10
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7260
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8490
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   10
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   870
      Width           =   9225
      _ExtentX        =   16277
      _ExtentY        =   8488
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      WordWrap        =   -1  'True
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
      _Band(0).Cols   =   1
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   2
      Left            =   3900
      TabIndex        =   8
      Top             =   450
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "類別："
      Height          =   180
      Left            =   3180
      TabIndex        =   7
      Top             =   450
      Width           =   540
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   1
      Left            =   1140
      TabIndex        =   6
      Top             =   660
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "點件數："
      Height          =   180
      Left            =   60
      TabIndex        =   5
      Top             =   645
      Width           =   720
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   2
      Top             =   435
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   60
      TabIndex        =   1
      Top             =   450
      Width           =   900
   End
End
Attribute VB_Name = "frm100122_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/26 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim strSql As String, i As Integer, j As Integer, s As Integer, strTemp As Variant, intK As Integer
Dim PLeft(0 To 9) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 13) As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim strSQL11 As String, strSQL12 As String, strSQL13 As String, strSQL21 As String, strSQL22 As String, strSQL23 As String
'Add By Cheng 2003/08/12
Dim m_strFileName As String '檔案名稱
Dim m_strFilePathName As String '路徑名稱

Private Sub SetDataListWidth()
With Me.grdDataList1
    .Cols = 14
    .row = 0
    .col = 0: .Text = "業務區"
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 1: .Text = "智權人員"
    .ColWidth(1) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 2: .Text = "系統類別"
    .ColWidth(2) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = "比較時段1" & vbCrLf & ChangeTStringToTDateString(frm100122_1.Txt1(21).Text) & vbCrLf & " | " & vbCrLf & ChangeTStringToTDateString(frm100122_1.Txt1(22).Text)
    .ColWidth(3) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = "增減"
    .ColWidth(4) = 600
    .CellAlignment = flexAlignCenterCenter
    .col = 5: .Text = "百分比"
    .ColWidth(5) = 1000
    .CellAlignment = flexAlignCenterCenter
    'edit by nickc 2006/02/09
    '.col = 6: .Text = "統計時段" & vbCrLf & ChangeTStringToTDateString(frm100122_1.txt1(1).Text) & vbCrLf & " | " & vbCrLf & ChangeTStringToTDateString(frm100122_1.txt1(2).Text)
    .col = 6: .Text = "比較時段2" & vbCrLf & ChangeTStringToTDateString(frm100122_1.Txt1(1).Text) & vbCrLf & " | " & vbCrLf & ChangeTStringToTDateString(frm100122_1.Txt1(2).Text)
    .ColWidth(6) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = "增減"
    .ColWidth(7) = 600
    .CellAlignment = flexAlignCenterCenter
    .col = 8: .Text = "百分比"
    .ColWidth(8) = 1000
    .CellAlignment = flexAlignCenterCenter
    'edit by nickc 2006/02/09
    '.col = 9: .Text = "比較時段2" & vbCrLf & ChangeTStringToTDateString(frm100122_1.txt1(23).Text) & vbCrLf & " | " & vbCrLf & ChangeTStringToTDateString(frm100122_1.txt1(24).Text)
    .col = 9: .Text = "比較時段3" & vbCrLf & ChangeTStringToTDateString(frm100122_1.Txt1(23).Text) & vbCrLf & " | " & vbCrLf & ChangeTStringToTDateString(frm100122_1.Txt1(24).Text)
    .ColWidth(9) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 10: .Text = "部門別"
    .ColWidth(10) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 11: .Text = "員工代號"
    .ColWidth(11) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 12: .Text = "序號"
    .ColWidth(12) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 13: .Text = "所別"
    .ColWidth(13) = 0
    .CellAlignment = flexAlignCenterCenter
    .RowHeight(0) = 255 * 3.5
    .MergeCells = flexMergeRestrictRows
    .MergeCol(0) = True: .MergeCol(1) = True
End With
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Dim strFileName  As String

Select Case cmdState
Case 0 '回前畫面
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 1 '結束
    fnCloseAllFrm100
Case 2 '列印
    PrintData
Case 3 '電子檔
    If Me.grdDataList1.Rows > 2 Then
        Screen.MousePointer = vbHourglass
'        m_strFileName = InputBox("請輸入欲匯出的檔案名稱???" & vbCrLf & vbCrLf & "檔案存放位置為==> X:\" & strUserNum)
'        If Trim(m_strFileName) <> "" Then
'            m_strFilePathName = "X:\" & strUserNum & "\" & m_strFileName & ".doc"
'            DeleteFile
            OpenWord
'            MsgBox "檔案 " & m_strFileName & ".doc 匯出成功!!!", vbExclamation + vbOKOnly
'        Else
'            MsgBox "您已取消作業或未輸入欲匯出的檔案名稱!!!", vbExclamation + vbOKOnly
'        End If
        Screen.MousePointer = vbDefault
    Else
        MsgBox "無資料可產生電子檔!!!", vbExclamation + vbOKOnly
    End If
Case Else
End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
End Sub

Private Sub PrintData()
Dim intItem As Integer
Dim ii As Integer
Dim strSalesZone As String '業務區
Dim strSales As String '智權人員
Dim strOffice As String '所別

Screen.MousePointer = vbHourglass
'若有資料
If Me.grdDataList1.Rows > 2 Then
    strSalesZone = " "
    strSales = " "
    strOffice = " "
    Page = 1
    PrintTitle
    With Me.grdDataList1
        For intItem = 1 To .Rows - 1
            For ii = 0 To .Cols - 1
                strTemp3(ii) = .TextMatrix(intItem, ii)
            Next ii
            If strOffice <> strTemp3(13) Then
                strOffice = strTemp3(13)
                strSalesZone = strTemp3(0)
                strSales = strTemp3(1)
            ElseIf strSalesZone <> strTemp3(0) Then
                strSalesZone = strTemp3(0)
                strSales = strTemp3(1)
            Else
                strTemp3(0) = ""
                If strSales <> strTemp3(1) Then
                    strSales = strTemp3(1)
                Else
                    strTemp3(1) = ""
                End If
            End If
            If strTemp3(2) = "小計" Then strTemp3(0) = ""
            If iPrint >= 10000 Then
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            If .TextMatrix(intItem, 2) = "小計" Then
                Printer.CurrentX = PLeft(2)
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                iPrint = iPrint + 300
            End If
            If .TextMatrix(intItem, 1) = "區小計" Or .TextMatrix(intItem, 1) = "所別小計" Or .TextMatrix(intItem, 1) = "總計" Then
                Printer.CurrentX = PLeft(1)
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                iPrint = iPrint + 300
            End If
            PrintDatil
            If .TextMatrix(intItem, 2) = "小計" Then
                iPrint = iPrint + 300
            End If
            If .TextMatrix(intItem, 1) = "區小計" Or .TextMatrix(intItem, 1) = "所別小計" Or .TextMatrix(intItem, 1) = "總計" Then
                iPrint = iPrint + 300
            End If
        Next intItem
    End With
    Printer.EndDoc
    ShowPrintOk
Else
    MsgBox "沒有資料可以列印 !", vbCritical
End If
Screen.MousePointer = vbDefault
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
'Modified by Lydia 2016/09/06
'Printer.Print "智權人員" & IIf(frm100122_1.Txt1(0).Text = "1", "收", "發") & "文量比較表"
Printer.Print "智權人員收發文量比較表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "案件性質：" & Me.lbl1(0).Caption
Printer.CurrentX = PLeft(5) - 500
Printer.CurrentY = iPrint
Printer.Print "類別：" & Me.lbl1(2).Caption
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "點件數：" & Me.lbl1(1).Caption
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "比較時段1"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
'edit by nickc 2006/02/09
'Printer.Print "統計時段"
Printer.Print "比較時段2"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
'edit by nickc 2006/02/09
'Printer.Print "比較時段2"
Printer.Print "比較時段3"
iPrint = iPrint + 300
     
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print ChangeTStringToTDateString(frm100122_1.Txt1(21).Text)
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print ChangeTStringToTDateString(frm100122_1.Txt1(1).Text)
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print ChangeTStringToTDateString(frm100122_1.Txt1(23).Text)
iPrint = iPrint + 300
     
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "　　|"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "　　|"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "　　|"
iPrint = iPrint + 300
     
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "業務區"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "系統類別"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print ChangeTStringToTDateString(frm100122_1.Txt1(22).Text)
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "　　增減"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "　百分比"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print ChangeTStringToTDateString(frm100122_1.Txt1(2).Text)
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "　　增減"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "　百分比"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print ChangeTStringToTDateString(frm100122_1.Txt1(24).Text)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500 '業務區
PLeft(1) = PLeft(0) + 125 * 15 '智權人員
PLeft(2) = PLeft(1) + 125 * 9 '系統類別
PLeft(3) = PLeft(2) + 125 * 10 '比較時段1
PLeft(4) = PLeft(3) + 125 * 12 '　　增減
PLeft(5) = PLeft(4) + 125 * 9 + 500 '　百分比
PLeft(6) = PLeft(5) + 125 * 9 + 500 '統計時段
PLeft(7) = PLeft(6) + 125 * 9 + 500 '　　增減
PLeft(8) = PLeft(7) + 125 * 9 + 500   '　百分比
PLeft(9) = PLeft(8) + 125 * 9 + 500 '比較時段2
End Sub

Sub PrintDatil()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp3(0)

Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp3(1)

Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print strTemp3(2)

Printer.CurrentX = PLeft(3) + Printer.TextWidth("比較時段") - Printer.TextWidth(strTemp3(3))
Printer.CurrentY = iPrint
Printer.Print strTemp3(3)

Printer.CurrentX = PLeft(4) + Printer.TextWidth("　　增減") - Printer.TextWidth(strTemp3(4))
Printer.CurrentY = iPrint
Printer.Print strTemp3(4)

Printer.CurrentX = PLeft(5) + Printer.TextWidth("　百分比") - Printer.TextWidth(strTemp3(5))
Printer.CurrentY = iPrint
Printer.Print strTemp3(5)

Printer.CurrentX = PLeft(6) + Printer.TextWidth("統計時段") - Printer.TextWidth(strTemp3(6))
Printer.CurrentY = iPrint
Printer.Print strTemp3(6)

Printer.CurrentX = PLeft(7) + Printer.TextWidth("　　增減") - Printer.TextWidth(strTemp3(7))
Printer.CurrentY = iPrint
Printer.Print strTemp3(7)

Printer.CurrentX = PLeft(8) + Printer.TextWidth("　百分比") - Printer.TextWidth(strTemp3(8))
Printer.CurrentY = iPrint
Printer.Print strTemp3(8)

Printer.CurrentX = PLeft(9) + Printer.TextWidth("比較時段") - Printer.TextWidth(strTemp3(9))
Printer.CurrentY = iPrint
Printer.Print strTemp3(9)

iPrint = iPrint + 300
End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
SetDataListWidth
'92.04.16 nick
cmdState = -1
End Sub

Sub StrMenu()
Dim ii As Integer
Dim strSaleZone As String '業務區
Dim strSales As String '智權人員
Dim strOffice As String '所別
Dim intBegin As Integer
Dim dblSalesAmt1 As Double, dblSalesAmt2 As Double, dblSalesAmt3 As Double
Dim dblSaleZoneAmt1 As Double, dblSaleZoneAmt2 As Double, dblSaleZoneAmt3 As Double
Dim dblOfficeAmt1 As Double, dblOfficeAmt2 As Double, dblOfficeAmt3 As Double
Dim dblTotAmt1 As Double, dblTotAmt2 As Double, dblTotAmt3 As Double
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

Me.Enabled = False
'讀出資料
If DoTemp = False Then
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Screen.MousePointer = vbDefault
   Exit Sub
End If
'顯示表單資料
'Modified by Lydia 2016/09/06 +是否含多國案
'lbl1(0).Caption = IIf(frm100122_1.txt1(13) = "1", "新申請案", "全部")
strExc(0) = IIf(frm100122_1.Txt1(13) = "1", "新申請案", "全部")
strExc(0) = strExc(0) & IIf(frm100122_1.Txt1(9) = "Y", "(含多國案)", "")
lbl1(0).Caption = strExc(0)
'end 2016/09/06
lbl1(1).Caption = IIf(frm100122_1.Txt1(5) = "1", "件數", "點數")
lbl1(2).Caption = IIf(frm100122_1.Txt1(0) = "1", "收文", "發文")
'選擇對象
Select Case frm100122_1.Txt1(17).Text
Case "1" '各區
    'Modified by Lydia 2016/09/06 非智權部單位全部歸北所
    'strSql = "Select A0902,       '', R1001224, Sum(R1001225), '', '', Sum(R1001228), '', '', Sum(R100122B), R1001221,               '', R1001223, ST06 From R100122, Acc090, Staff Where R1001221=A0901 And R1001222=ST01(+) And ID='" & strUserNum & "' Group By A0902, R1001224, R1001221, R1001223, ST06 Order By R1001221, R1001223 "
    'Added by Lydia 2023/12/25
    If strSrvDate(1) >= 新部門啟用日 Then
       strSql = "Select nvl(a0922,A0902) as a0902, '', R1001224, Sum(R1001225), '', '', Sum(R1001228), '', '', Sum(R100122B), R1001221, '', R1001223, DECODE(SUBSTR(R1001221,1,1),'S',ST06,'1') ST06,DECODE(SUBSTR(R1001221,1,1),'S','0','1') ORD2 From R100122, Acc090, Staff, Acc090New " & _
              "Where R1001221=A0901 And R1001222=ST01(+) And ID='" & strUserNum & "' and st93=a0921(+) Group By nvl(a0922,A0902), R1001224, R1001221, R1001223, DECODE(SUBSTR(R1001221,1,1),'S',ST06,'1'),DECODE(SUBSTR(R1001221,1,1),'S','0','1') Order By ST06,ORD2,R1001221, R1001223 "
    Else
    'end 2023/12/25
       strSql = "Select A0902, '', R1001224, Sum(R1001225), '', '', Sum(R1001228), '', '', Sum(R100122B), R1001221, '', R1001223, DECODE(SUBSTR(R1001221,1,1),'S',ST06,'1') ST06,DECODE(SUBSTR(R1001221,1,1),'S','0','1') ORD2 From R100122, Acc090, Staff " & _
                "Where R1001221=A0901 And R1001222=ST01(+) And ID='" & strUserNum & "' Group By A0902, R1001224, R1001221, R1001223, DECODE(SUBSTR(R1001221,1,1),'S',ST06,'1'),DECODE(SUBSTR(R1001221,1,1),'S','0','1') Order By ST06,ORD2,R1001221, R1001223 "
    End If
Case Else '個人
    'Modified by Lydia 2016/09/06 非智權部單位全部歸北所
    'strSql = "Select A0902, ST02, R1001224, Sum(R1001225), '', '', Sum(R1001228), '', '', Sum(R100122B), R1001221, R1001222, R1001223, ST06 From R100122, Acc090, Staff Where R1001221=A0901(+) And R1001222=ST01(+) And ID='" & strUserNum & "' Group By A0902, ST02, R1001224, R1001221, R1001222, R1001223, ST06 Order By R1001221, R1001222, R1001223 "
    'Added by Lydia 2023/12/25
    If strSrvDate(1) >= 新部門啟用日 Then
       strSql = "Select nvl(a0922,A0902) as a0902, ST02, R1001224, Sum(R1001225), '', '', Sum(R1001228), '', '', Sum(R100122B), R1001221, R1001222, R1001223, DECODE(SUBSTR(R1001221,1,1),'S',ST06,'1') ST06,DECODE(SUBSTR(R1001221,1,1),'S','0','1') ORD2 From R100122, Acc090, Staff, Acc090New " & _
                "Where R1001221=A0901(+) And R1001222=ST01(+) And ID='" & strUserNum & "' and st93=a0921(+) Group By nvl(a0922,A0902), ST02, R1001224, R1001221, R1001222, R1001223, DECODE(SUBSTR(R1001221,1,1),'S',ST06,'1'),DECODE(SUBSTR(R1001221,1,1),'S','0','1') Order By ST06,ORD2,R1001221, R1001222, R1001223 "
    Else
    'end 2023/12/25
       strSql = "Select A0902, ST02, R1001224, Sum(R1001225), '', '', Sum(R1001228), '', '', Sum(R100122B), R1001221, R1001222, R1001223, DECODE(SUBSTR(R1001221,1,1),'S',ST06,'1') ST06,DECODE(SUBSTR(R1001221,1,1),'S','0','1') ORD2 From R100122, Acc090, Staff " & _
                "Where R1001221=A0901(+) And R1001222=ST01(+) And ID='" & strUserNum & "' Group By A0902, ST02, R1001224, R1001221, R1001222, R1001223, DECODE(SUBSTR(R1001221,1,1),'S',ST06,'1'),DECODE(SUBSTR(R1001221,1,1),'S','0','1') Order By ST06,ORD2,R1001221, R1001222, R1001223 "
    
    End If
End Select
CheckOC
grdDataList1.Visible = False
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
Set grdDataList1.Recordset = adoRecordset
SetDataListWidth
CheckOC
With Me.grdDataList1
    dblSalesAmt1 = 0: dblSalesAmt2 = 0: dblSalesAmt3 = 0
    dblSaleZoneAmt1 = 0: dblSaleZoneAmt2 = 0: dblSaleZoneAmt3 = 0
    dblOfficeAmt1 = 0: dblOfficeAmt2 = 0: dblOfficeAmt3 = 0
    dblTotAmt1 = 0: dblTotAmt2 = 0: dblTotAmt3 = 0
    intBegin = 1
ReDo:
    strSaleZone = .TextMatrix(intBegin, 0)
    strSales = .TextMatrix(intBegin, 1)
    strOffice = .TextMatrix(intBegin, 13)
    For ii = intBegin To .Rows - 1
        .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
        .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
        .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
        .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
        .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
        .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
        .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
        '若所別不同
        If strOffice <> .TextMatrix(ii, 13) Then
            If frm100122_1.Txt1(17).Text = "2" Then '個人
                .AddItem "", ii
                .TextMatrix(ii, 0) = strSaleZone
                .TextMatrix(ii, 2) = "小計"
                .TextMatrix(ii, 3) = dblSalesAmt1
                .TextMatrix(ii, 6) = dblSalesAmt2
                .TextMatrix(ii, 9) = dblSalesAmt3
                'edit by nickc 2006/02/09
                '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
                .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
                .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
                .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
                .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
                .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
                .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
                .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
                .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
                .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
                .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
                .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
                ii = ii + 1
            End If
            .AddItem "", ii
            .TextMatrix(ii, 1) = "區小計"
            .TextMatrix(ii, 3) = dblSaleZoneAmt1
            .TextMatrix(ii, 6) = dblSaleZoneAmt2
            .TextMatrix(ii, 9) = dblSaleZoneAmt3
            'edit by nickc 2006/02/09
            '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
            .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
            .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
            .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
            .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
            .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
            .Refresh
            'Modified by Lydia 2016/09/06 非智權部單位全部歸北所
            'StrSQLa = "Select '', '', R1001224, Sum(R1001225), '', '', Sum(R1001228), '', '', Sum(R100122B), '', '', R1001223, ST06 From R100122, Acc090, Staff Where R1001221=A0901 And R1001222=ST01(+) And ST06='" & strOffice & "' And ID='" & strUserNum & "' Group By R1001224, R1001223, ST06 Order By R1001223 "
             StrSQLa = "Select '', '', R1001224, Sum(R1001225), '', '', Sum(R1001228), '', '', Sum(R100122B), '', '', R1001223, '" & strOffice & "' as ST06 From R100122, Acc090, Staff Where R1001221=A0901 And R1001222=ST01(+) And DECODE(SUBSTR(R1001221,1,1),'S',ST06,'1')='" & strOffice & "' And ID='" & strUserNum & "' Group By R1001224, R1001223 Order By R1001223 "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                While Not rsA.EOF
                    ii = ii + 1
                    .AddItem "", ii
                    .TextMatrix(ii, 0) = IIf(strOffice = "1", "北所", IIf(strOffice = "2", "中所", IIf(strOffice = "3", "南所", IIf(strOffice = "4", "高所", "廣東所"))))
                    .TextMatrix(ii, 1) = ""
                    .TextMatrix(ii, 2) = "" & rsA.Fields(2).Value
                    .TextMatrix(ii, 3) = Val("" & rsA.Fields(3).Value)
                    .TextMatrix(ii, 6) = Val("" & rsA.Fields(6).Value)
                    .TextMatrix(ii, 9) = Val("" & rsA.Fields(9).Value)
                    'edit by nickc 2006/02/09
                    '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
                    .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
                    .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
                    .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
                    .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
                    .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
                    .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
                    .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
                    .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
                    .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
                    .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
                    .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
                    .Refresh
                                    
                    rsA.MoveNext
                Wend
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            
            ii = ii + 1
            .AddItem "", ii
            .TextMatrix(ii, 0) = "　　　　"
            .TextMatrix(ii, 1) = "所別小計"
            .TextMatrix(ii, 3) = dblOfficeAmt1
            .TextMatrix(ii, 6) = dblOfficeAmt2
            .TextMatrix(ii, 9) = dblOfficeAmt3
            'edit by nickc 2006/02/09
            '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
            .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
            .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
            .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
            .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
            .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
            .Refresh
            
            strSaleZone = .TextMatrix(ii, 0)
            strSales = .TextMatrix(ii, 1)
            strOffice = .TextMatrix(ii, 13)
            intBegin = ii + 1
            dblSalesAmt1 = 0: dblSalesAmt2 = 0: dblSalesAmt3 = 0
            dblSaleZoneAmt1 = 0: dblSaleZoneAmt2 = 0: dblSaleZoneAmt3 = 0
            dblOfficeAmt1 = 0: dblOfficeAmt2 = 0: dblOfficeAmt3 = 0
            GoTo ReDo
            
        '若業務區不同
        ElseIf strSaleZone <> .TextMatrix(ii, 0) Then
            If frm100122_1.Txt1(17).Text = "2" Then
                .AddItem "", ii
                .TextMatrix(ii, 0) = strSaleZone
                .TextMatrix(ii, 2) = "小計"
                .TextMatrix(ii, 3) = dblSalesAmt1
                .TextMatrix(ii, 6) = dblSalesAmt2
                .TextMatrix(ii, 9) = dblSalesAmt3
                'edit by nickc 2006/02/09
                '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
                .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
                .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
                .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
                .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
                .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
                .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
                .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
                .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
                .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
                .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
                .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
                ii = ii + 1
            End If
            .AddItem "", ii
            .TextMatrix(ii, 1) = "區小計"
            .TextMatrix(ii, 3) = dblSaleZoneAmt1
            .TextMatrix(ii, 6) = dblSaleZoneAmt2
            .TextMatrix(ii, 9) = dblSaleZoneAmt3
            'edit by nickc 2006/02/09
            '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
            .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
            .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
            .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
            .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
            .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
            .Refresh
            strSaleZone = .TextMatrix(ii, 0)
            strSales = .TextMatrix(ii, 1)
            intBegin = ii + 1
            dblSalesAmt1 = 0: dblSalesAmt2 = 0: dblSalesAmt3 = 0
            dblSaleZoneAmt1 = 0: dblSaleZoneAmt2 = 0: dblSaleZoneAmt3 = 0
            GoTo ReDo
        '若智權人員不同
        ElseIf strSales <> .TextMatrix(ii, 1) Then
            .AddItem "", ii
            .TextMatrix(ii, 0) = strSaleZone
            .TextMatrix(ii, 2) = "小計"
            .TextMatrix(ii, 3) = dblSalesAmt1
            .TextMatrix(ii, 6) = dblSalesAmt2
            .TextMatrix(ii, 9) = dblSalesAmt3
            'edit by nickc 2006/02/09
            '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
            .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
            .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
            .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
            .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
            .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
            .Refresh
            strSales = .TextMatrix(ii, 1)
            intBegin = ii + 1
            dblSalesAmt1 = 0: dblSalesAmt2 = 0: dblSalesAmt3 = 0
            GoTo ReDo
        End If
        
        If .TextMatrix(ii, 3) = "" Then .TextMatrix(ii, 3) = "0"
        If .TextMatrix(ii, 9) = "" Then .TextMatrix(ii, 9) = "0"
        'edit by nickc 2006/02/09
        '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
        .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
        .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
        .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
        .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
        
        dblSalesAmt1 = dblSalesAmt1 + Val(.TextMatrix(ii, 3))
        dblSalesAmt2 = dblSalesAmt2 + Val(.TextMatrix(ii, 6))
        dblSalesAmt3 = dblSalesAmt3 + Val(.TextMatrix(ii, 9))
        
        dblSaleZoneAmt1 = dblSaleZoneAmt1 + Val(.TextMatrix(ii, 3))
        dblSaleZoneAmt2 = dblSaleZoneAmt2 + Val(.TextMatrix(ii, 6))
        dblSaleZoneAmt3 = dblSaleZoneAmt3 + Val(.TextMatrix(ii, 9))
    
        dblOfficeAmt1 = dblOfficeAmt1 + Val(.TextMatrix(ii, 3))
        dblOfficeAmt2 = dblOfficeAmt2 + Val(.TextMatrix(ii, 6))
        dblOfficeAmt3 = dblOfficeAmt3 + Val(.TextMatrix(ii, 9))
    
        dblTotAmt1 = dblTotAmt1 + Val(.TextMatrix(ii, 3))
        dblTotAmt2 = dblTotAmt2 + Val(.TextMatrix(ii, 6))
        dblTotAmt3 = dblTotAmt3 + Val(.TextMatrix(ii, 9))
    
    Next ii
    If frm100122_1.Txt1(17).Text = "2" Then
        .AddItem "", ii
        .TextMatrix(ii, 0) = strSaleZone
        .TextMatrix(ii, 2) = "小計"
        .TextMatrix(ii, 3) = dblSalesAmt1
        .TextMatrix(ii, 6) = dblSalesAmt2
        .TextMatrix(ii, 9) = dblSalesAmt3
        'edit by nickc 2006/02/09
        '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
        .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
        .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
        .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
        .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
        .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
        .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
        .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
        .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
        .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
        .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
        .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
        ii = ii + 1
    End If
    .AddItem "", ii
    .TextMatrix(ii, 1) = "區小計"
    .TextMatrix(ii, 3) = dblSaleZoneAmt1
    .TextMatrix(ii, 6) = dblSaleZoneAmt2
    .TextMatrix(ii, 9) = dblSaleZoneAmt3
    'edit by nickc 2006/02/09
    '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
    .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
    .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
    .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
    .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
    .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
    .Refresh
    'Modified by Lydia 2016/09/06 非智權部單位全部歸北所
    'StrSQLa = "Select '', '', R1001224, Sum(R1001225), '', '', Sum(R1001228), '', '', Sum(R100122B), '', '', R1001223, ST06 From R100122, Acc090, Staff Where R1001221=A0901 And R1001222=ST01(+) And ST06='" & strOffice & "' And ID='" & strUserNum & "' Group By R1001224, R1001223, ST06 Order By R1001223 "
    StrSQLa = "Select '', '', R1001224, Sum(R1001225), '', '', Sum(R1001228), '', '', Sum(R100122B), '', '', R1001223, '" & strOffice & "' as ST06 From R100122, Acc090, Staff Where R1001221=A0901 And R1001222=ST01(+) And DECODE(SUBSTR(R1001221,1,1),'S',ST06,'1')='" & strOffice & "' And ID='" & strUserNum & "' Group By R1001224, R1001223 Order By R1001223 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            ii = ii + 1
            .AddItem "", ii
            .TextMatrix(ii, 0) = IIf(strOffice = "1", "北所", IIf(strOffice = "2", "中所", IIf(strOffice = "3", "南所", IIf(strOffice = "4", "高所", "廣東所"))))
            .TextMatrix(ii, 1) = ""
            .TextMatrix(ii, 2) = "" & rsA.Fields(2).Value
            .TextMatrix(ii, 3) = Val("" & rsA.Fields(3).Value)
            .TextMatrix(ii, 6) = Val("" & rsA.Fields(6).Value)
            .TextMatrix(ii, 9) = Val("" & rsA.Fields(9).Value)
            'edit by nickc  2006/02/09
            '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
            .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
            .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
            .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
            .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
            .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
            .Refresh
                            
            rsA.MoveNext
        Wend
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    
    ii = ii + 1
    .AddItem "", ii
    .TextMatrix(ii, 0) = "　　　　"
    .TextMatrix(ii, 1) = "所別小計"
    .TextMatrix(ii, 3) = dblOfficeAmt1
    .TextMatrix(ii, 6) = dblOfficeAmt2
    .TextMatrix(ii, 9) = dblOfficeAmt3
    'edit by nickc 2006/02/09
    '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
    .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
    .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
    .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
    .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
    .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
    .Refresh
    
    StrSQLa = "Select '', '', R1001224, Sum(R1001225), '', '', Sum(R1001228), '', '', Sum(R100122B), '', '', R1001223, '' From R100122, Acc090, Staff Where R1001221=A0901 And R1001222=ST01(+) And ID='" & strUserNum & "' Group By R1001224, R1001223 Order By R1001223 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            ii = ii + 1
            .AddItem "", ii
            .TextMatrix(ii, 0) = "全所"
            .TextMatrix(ii, 1) = ""
            .TextMatrix(ii, 2) = "" & rsA.Fields(2).Value
            .TextMatrix(ii, 3) = Val("" & rsA.Fields(3).Value)
            .TextMatrix(ii, 6) = Val("" & rsA.Fields(6).Value)
            .TextMatrix(ii, 9) = Val("" & rsA.Fields(9).Value)
            'edit by nickc 2006/02/09
            '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
            .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
            .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
            .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
            .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
            .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
            .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
            .Refresh
                            
            rsA.MoveNext
        Wend
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    
    ii = ii + 1
    .AddItem "", ii
    .TextMatrix(ii, 0) = "　　"
    .TextMatrix(ii, 1) = "總計"
    .TextMatrix(ii, 3) = dblTotAmt1
    .TextMatrix(ii, 6) = dblTotAmt2
    .TextMatrix(ii, 9) = dblTotAmt3
    'eidt by nickc 2006/02/09
    '.TextMatrix(ii, 4) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 3))
    .TextMatrix(ii, 4) = Val(.TextMatrix(ii, 3)) - Val(.TextMatrix(ii, 6))
    .TextMatrix(ii, 5) = Format(Val(.TextMatrix(ii, 4)) / Val(IIf(.TextMatrix(ii, 3) = "0", "1", .TextMatrix(ii, 3))) * 100, "##0.00") & "%"
    .TextMatrix(ii, 7) = Val(.TextMatrix(ii, 6)) - Val(.TextMatrix(ii, 9))
    .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 7)) / Val(IIf(.TextMatrix(ii, 9) = "0", "1", .TextMatrix(ii, 9))) * 100, "##0.00") & "%"
    .row = ii: .col = 3: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 4: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 5: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 6: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 7: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 8: .CellAlignment = flexAlignRightCenter
    .row = ii: .col = 9: .CellAlignment = flexAlignRightCenter
    .Refresh
    
    For ii = 1 To .Rows - 1
        If frm100122_1.Txt1(5).Text = "1" Then
            .TextMatrix(ii, 4) = Format(.TextMatrix(ii, 4), "##0")
            .TextMatrix(ii, 7) = Format(.TextMatrix(ii, 7), "##0")
        Else
            .TextMatrix(ii, 4) = Format(.TextMatrix(ii, 4), "##0.0")
            .TextMatrix(ii, 7) = Format(.TextMatrix(ii, 7), "##0.0")
        End If
        If Val(.TextMatrix(ii, 4)) > 0 Then .TextMatrix(ii, 4) = "+" & .TextMatrix(ii, 4)
        If Val(.TextMatrix(ii, 7)) > 0 Then .TextMatrix(ii, 7) = "+" & .TextMatrix(ii, 7)
    Next ii
End With
grdDataList1.Visible = True
Me.Enabled = True
End Sub

Function DoTemp() As Boolean
'911024 nick 指定使用 index
Dim IndexString As String
Dim intK As Integer, IntK1 As Integer
frm100122_1.Hide
cnnConnection.Execute "DELETE FROM R100122 where id='" & strUserNum & "' "
strSQL11 = "": strSQL12 = "": strSQL13 = "": strSQL21 = "": strSQL22 = "": strSQL23 = ""
'組合條件
'系統類別
If Len(Trim(frm100122_1.Txt1(3))) <> 0 Then
    strSQL11 = strSQL11 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100122_1.Txt1(3).Text <> "ALL", frm100122_1.Txt1(3).Text, frm100122_1.m_strSystemkindByUser), 1) & ") "
    strSQL12 = strSQL12 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100122_1.Txt1(3).Text <> "ALL", frm100122_1.Txt1(3).Text, frm100122_1.m_strSystemkindByUser), 1) & ") "
    strSQL13 = strSQL13 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100122_1.Txt1(3).Text <> "ALL", frm100122_1.Txt1(3).Text, frm100122_1.m_strSystemkindByUser), 1) & ") "
    
    strSQL21 = strSQL21 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100122_1.Txt1(3).Text <> "ALL", frm100122_1.Txt1(3).Text, frm100122_1.m_strSystemkindByUser), 2) & ") "
    strSQL22 = strSQL22 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100122_1.Txt1(3).Text <> "ALL", frm100122_1.Txt1(3).Text, frm100122_1.m_strSystemkindByUser), 2) & ") "
    strSQL23 = strSQL23 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100122_1.Txt1(3).Text <> "ALL", frm100122_1.Txt1(3).Text, frm100122_1.m_strSystemkindByUser), 2) & ") "
End If
'是否含FC資料
If frm100122_1.Txt1(4).Text = "" Then
    strSQL11 = strSQL11 & " AND CP01<>'FCP'  And CP01<>'FCT' "
    strSQL12 = strSQL12 & " AND CP01<>'FCP'  And CP01<>'FCT' "
    strSQL13 = strSQL13 & " AND CP01<>'FCP'  And CP01<>'FCT' "
    
    strSQL21 = strSQL21 & " AND CP01<>'FCP'  And CP01<>'FCT' "
    strSQL22 = strSQL22 & " AND CP01<>'FCP'  And CP01<>'FCT' "
    strSQL23 = strSQL23 & " AND CP01<>'FCP'  And CP01<>'FCT' "
End If
'查詢別-收文
If frm100122_1.Txt1(0) = "1" Then
    'Modified by Lydia 2016/09/06 +判斷未取消收文 CP159=0
    strSQL11 = strSQL11 & " AND CP05>=" & Val(ChangeTStringToWString(frm100122_1.Txt1(21))) & " "
    strSQL11 = strSQL11 & " AND CP05<=" & Val(ChangeTStringToWString(frm100122_1.Txt1(22))) & " And CP159=0 "
    strSQL12 = strSQL12 & " AND CP05>=" & Val(ChangeTStringToWString(frm100122_1.Txt1(1))) & " "
    strSQL12 = strSQL12 & " AND CP05<=" & Val(ChangeTStringToWString(frm100122_1.Txt1(2))) & " And CP159=0 "
    strSQL13 = strSQL13 & " AND CP05>=" & Val(ChangeTStringToWString(frm100122_1.Txt1(23))) & " "
    strSQL13 = strSQL13 & " AND CP05<=" & Val(ChangeTStringToWString(frm100122_1.Txt1(24))) & " And CP159=0 "

    strSQL21 = strSQL21 & " AND CP05>=" & Val(ChangeTStringToWString(frm100122_1.Txt1(21))) & " "
    strSQL21 = strSQL21 & " AND CP05<=" & Val(ChangeTStringToWString(frm100122_1.Txt1(22))) & " And CP159=0 "
    strSQL22 = strSQL22 & " AND CP05>=" & Val(ChangeTStringToWString(frm100122_1.Txt1(1))) & " "
    strSQL22 = strSQL22 & " AND CP05<=" & Val(ChangeTStringToWString(frm100122_1.Txt1(2))) & " And CP159=0 "
    strSQL23 = strSQL23 & " AND CP05>=" & Val(ChangeTStringToWString(frm100122_1.Txt1(23))) & " "
    strSQL23 = strSQL23 & " AND CP05<=" & Val(ChangeTStringToWString(frm100122_1.Txt1(24))) & " And CP159=0 "
'查詢別-發文
Else
    strSQL11 = strSQL11 & " AND CP27>=" & Val(ChangeTStringToWString(frm100122_1.Txt1(21))) & " "
    strSQL11 = strSQL11 & " AND CP27<=" & Val(ChangeTStringToWString(frm100122_1.Txt1(22))) & " "
    strSQL12 = strSQL12 & " AND CP27>=" & Val(ChangeTStringToWString(frm100122_1.Txt1(1))) & " "
    strSQL12 = strSQL12 & " AND CP27<=" & Val(ChangeTStringToWString(frm100122_1.Txt1(2))) & " "
    strSQL13 = strSQL13 & " AND CP27>=" & Val(ChangeTStringToWString(frm100122_1.Txt1(23))) & " "
    strSQL13 = strSQL13 & " AND CP27<=" & Val(ChangeTStringToWString(frm100122_1.Txt1(24))) & " "

    strSQL21 = strSQL21 & " AND CP27>=" & Val(ChangeTStringToWString(frm100122_1.Txt1(21))) & " "
    strSQL21 = strSQL21 & " AND CP27<=" & Val(ChangeTStringToWString(frm100122_1.Txt1(22))) & " "
    strSQL22 = strSQL22 & " AND CP27>=" & Val(ChangeTStringToWString(frm100122_1.Txt1(1))) & " "
    strSQL22 = strSQL22 & " AND CP27<=" & Val(ChangeTStringToWString(frm100122_1.Txt1(2))) & " "
    strSQL23 = strSQL23 & " AND CP27>=" & Val(ChangeTStringToWString(frm100122_1.Txt1(23))) & " "
    strSQL23 = strSQL23 & " AND CP27<=" & Val(ChangeTStringToWString(frm100122_1.Txt1(24))) & " "
End If
'案件性質
'若為新申請案
If frm100122_1.Txt1(13).Text = "1" Then
    'Modified by Lydia 2016/09/06 改案件性質為新申請案共用變數+改請(3xx)
'    strSQL11 = strSQL11 & " And (To_Number(CP10)>=101 And To_Number(CP10)<=105) "
'    strSQL12 = strSQL12 & " And (To_Number(CP10)>=101 And To_Number(CP10)<=105) "
'    strSQL13 = strSQL13 & " And (To_Number(CP10)>=101 And To_Number(CP10)<=105) "
    strSQL11 = strSQL11 & " And (instr('" & NewCasePtyList & "',CP10)>0 or substr(CP10,1,1)='3') "
    strSQL12 = strSQL12 & " And (instr('" & NewCasePtyList & "',CP10)>0 or substr(CP10,1,1)='3') "
    strSQL13 = strSQL13 & " And (instr('" & NewCasePtyList & "',CP10)>0 or substr(CP10,1,1)='3') "
    'end 2016/09/06
    strSQL21 = strSQL21 & " And CP10='101' "
    strSQL22 = strSQL22 & " And CP10='101' "
    strSQL23 = strSQL23 & " And CP10='101' "
End If
'點件數
'若為件數
If frm100122_1.Txt1(5).Text = "1" Then
    strSQL11 = strSQL11 & " And CP26 Is Null "
    strSQL12 = strSQL12 & " And CP26 Is Null "
    strSQL13 = strSQL13 & " And CP26 Is Null "
    
    strSQL21 = strSQL21 & " And CP26 Is Null "
    strSQL22 = strSQL22 & " And CP26 Is Null "
    strSQL23 = strSQL23 & " And CP26 Is Null "
End If
'業務區
If frm100122_1.Txt1(6).Text <> "" Then
    strSQL11 = strSQL11 & " And ST15>='" & frm100122_1.Txt1(6).Text & "' "
    strSQL12 = strSQL12 & " And ST15>='" & frm100122_1.Txt1(6).Text & "' "
    strSQL13 = strSQL13 & " And ST15>='" & frm100122_1.Txt1(6).Text & "' "

    strSQL21 = strSQL21 & " And ST15>='" & frm100122_1.Txt1(6).Text & "' "
    strSQL22 = strSQL22 & " And ST15>='" & frm100122_1.Txt1(6).Text & "' "
    strSQL23 = strSQL23 & " And ST15>='" & frm100122_1.Txt1(6).Text & "' "
End If
If frm100122_1.Txt1(7).Text <> "" Then
    strSQL11 = strSQL11 & " And ST15<='" & frm100122_1.Txt1(7).Text & "' "
    strSQL12 = strSQL12 & " And ST15<='" & frm100122_1.Txt1(7).Text & "' "
    strSQL13 = strSQL13 & " And ST15<='" & frm100122_1.Txt1(7).Text & "' "

    strSQL21 = strSQL21 & " And ST15<='" & frm100122_1.Txt1(7).Text & "' "
    strSQL22 = strSQL22 & " And ST15<='" & frm100122_1.Txt1(7).Text & "' "
    strSQL23 = strSQL23 & " And ST15<='" & frm100122_1.Txt1(7).Text & "' "
End If
'智權人員
If frm100122_1.Txt1(8).Text <> "" Then
    strSQL11 = strSQL11 & " And CP13='" & frm100122_1.Txt1(8).Text & "' "
    strSQL12 = strSQL12 & " And CP13='" & frm100122_1.Txt1(8).Text & "' "
    strSQL13 = strSQL13 & " And CP13='" & frm100122_1.Txt1(8).Text & "' "
    
    strSQL21 = strSQL21 & " And CP13='" & frm100122_1.Txt1(8).Text & "' "
    strSQL22 = strSQL22 & " And CP13='" & frm100122_1.Txt1(8).Text & "' "
    strSQL23 = strSQL23 & " And CP13='" & frm100122_1.Txt1(8).Text & "' "
End If

'Added by Lydia 2016/09/06 是否含多國案
If frm100122_1.Txt1(9).Text <> "Y" Then
    strSQL11 = strSQL11 & " And CP21 is Null "
    strSQL12 = strSQL12 & " And CP21 is Null "
    strSQL13 = strSQL13 & " And CP21 is Null "
    
    strSQL21 = strSQL21 & " And CP21 is Null "
    strSQL22 = strSQL22 & " And CP21 is Null "
    strSQL23 = strSQL23 & " And CP21 is Null "
End If

'若計件件數
If frm100122_1.Txt1(5).Text = "1" Then
    '比較時段1
    strSql = "Select ST15, CP13, Decode(PA09,'000','1','2'), CP01||Decode(PA09,'000','(台灣)','(非台灣)'), Count(*), '', '', 0, '', '', 0, '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='P' " & strSQL11 & " Group By ST15, CP13, Decode(PA09,'000','1','2'), CP01||Decode(PA09,'000','(台灣)','(非台灣)') "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '5', CP01, Count(*), '', '', 0, '', '', 0, '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='CFP' " & strSQL11 & " Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '7', CP01, Count(*), '', '', 0, '', '', 0, '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='FCP' " & strSQL11 & " Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    
    strSql = "Select ST15, CP13, Decode(TM10,'000','3','4'), CP01||Decode(TM10,'000','(台灣)','(非台灣)'), Count(*), '', '', 0, '', '', 0, '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='T' " & strSQL21 & " Group By ST15, CP13, Decode(TM10,'000','3','4'), CP01||Decode(TM10,'000','(台灣)','(非台灣)') "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '6', CP01, Count(*), '', '', 0, '', '', 0, '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='CFT' " & strSQL21 & " Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '8', CP01, Count(*), '', '', 0, '', '', 0, '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='FCT' " & strSQL21 & " Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    '統計時段
    strSql = "Select ST15, CP13, Decode(PA09,'000','1','2'), CP01||Decode(PA09,'000','(台灣)','(非台灣)'), 0, '', '', Count(*), '', '', 0, '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='P' " & strSQL12 & " Group By ST15, CP13, Decode(PA09,'000','1','2'), CP01||Decode(PA09,'000','(台灣)','(非台灣)') "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '5', CP01, 0, '', '', Count(*), '', '', 0, '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='CFP' " & strSQL12 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '7', CP01, 0, '', '', Count(*), '', '', 0, '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='FCP' " & strSQL12 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    
    strSql = "Select ST15, CP13, Decode(TM10,'000','3','4'), CP01||Decode(TM10,'000','(台灣)','(非台灣)'), 0, '', '', Count(*), '', '', 0, '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='T' " & strSQL22 & " Group By ST15, CP13, Decode(TM10,'000','3','4'), CP01||Decode(TM10,'000','(台灣)','(非台灣)') "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '6', CP01, 0, '', '', Count(*), '', '', 0, '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='CFT' " & strSQL22 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '8', CP01, 0, '', '', Count(*), '', '', 0, '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='FCT' " & strSQL22 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    '比較時段2
    strSql = "Select ST15, CP13, Decode(PA09,'000','1','2'), CP01||Decode(PA09,'000','(台灣)','(非台灣)'), 0, '', '', 0, '', '', Count(*), '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='P' " & strSQL13 & " Group By ST15, CP13, Decode(PA09,'000','1','2'), CP01||Decode(PA09,'000','(台灣)','(非台灣)') "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '5', CP01, 0, '', '', 0, '', '', Count(*), '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='CFP' " & strSQL13 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '7', CP01, 0, '', '', 0, '', '', Count(*), '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='FCP' " & strSQL13 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    
    strSql = "Select ST15, CP13, Decode(TM10,'000','3','4'), CP01||Decode(TM10,'000','(台灣)','(非台灣)'), 0, '', '', 0, '', '', Count(*), '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='T' " & strSQL23 & " Group By ST15, CP13, Decode(TM10,'000','3','4'), CP01||Decode(TM10,'000','(台灣)','(非台灣)') "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '6', CP01, 0, '', '', 0, '', '', Count(*), '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='CFT' " & strSQL23 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '8', CP01, 0, '', '', 0, '', '', Count(*), '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='FCT' " & strSQL23 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
'若計算點數
Else
    '比較時段1
    strSql = "Select ST15, CP13, Decode(PA09,'000','1','2'), CP01||Decode(PA09,'000','(台灣)','(非台灣)'), Sum(CP18), '', '', 0, '', '', 0, '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='P' " & strSQL11 & " Group By ST15, CP13, Decode(PA09,'000','1','2'), CP01||Decode(PA09,'000','(台灣)','(非台灣)') "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '5', CP01, Sum(CP18), '', '', 0, '', '', 0, '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='CFP' " & strSQL11 & " Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '7', CP01, Sum(CP18), '', '', 0, '', '', 0, '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='FCP' " & strSQL11 & " Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    
    strSql = "Select ST15, CP13, Decode(TM10,'000','3','4'), CP01||Decode(TM10,'000','(台灣)','(非台灣)'), Sum(CP18), '', '', 0, '', '', 0, '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='T' " & strSQL21 & " Group By ST15, CP13, Decode(TM10,'000','3','4'), CP01||Decode(TM10,'000','(台灣)','(非台灣)') "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '6', CP01, Sum(CP18), '', '', 0, '', '', 0, '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='CFT' " & strSQL21 & " Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '8', CP01, Sum(CP18), '', '', 0, '', '', 0, '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='FCT' " & strSQL21 & " Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    '統計時段
    strSql = "Select ST15, CP13, Decode(PA09,'000','1','2'), CP01||Decode(PA09,'000','(台灣)','(非台灣)'), 0, '', '', Sum(CP18), '', '', 0, '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='P' " & strSQL12 & " Group By ST15, CP13, Decode(PA09,'000','1','2'), CP01||Decode(PA09,'000','(台灣)','(非台灣)') "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '5', CP01, 0, '', '', Sum(CP18), '', '', 0, '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='CFP' " & strSQL12 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '7', CP01, 0, '', '', Sum(CP18), '', '', 0, '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='FCP' " & strSQL12 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    
    strSql = "Select ST15, CP13, Decode(TM10,'000','3','4'), CP01||Decode(TM10,'000','(台灣)','(非台灣)'), 0, '', '', Sum(CP18), '', '', 0, '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='T' " & strSQL22 & " Group By ST15, CP13, Decode(TM10,'000','3','4'), CP01||Decode(TM10,'000','(台灣)','(非台灣)') "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '6', CP01, 0, '', '', Sum(CP18), '', '', 0, '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='CFT' " & strSQL22 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '8', CP01, 0, '', '', Sum(CP18), '', '', 0, '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='FCT' " & strSQL22 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    '比較時段2
    strSql = "Select ST15, CP13, Decode(PA09,'000','1','2'), CP01||Decode(PA09,'000','(台灣)','(非台灣)'), 0, '', '', 0, '', '', Sum(CP18), '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='P' " & strSQL13 & " Group By ST15, CP13, Decode(PA09,'000','1','2'), CP01||Decode(PA09,'000','(台灣)','(非台灣)') "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '5', CP01, 0, '', '', 0, '', '', Sum(CP18), '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='CFP' " & strSQL13 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '7', CP01, 0, '', '', 0, '', '', Sum(CP18), '" & strUserNum & "' From  CaseProgress, Patent, Staff Where PA01=CP01 And PA02=CP02 And PA03=CP03 And PA04=CP04 And CP13=ST01 And CP01='FCP' " & strSQL13 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    
    strSql = "Select ST15, CP13, Decode(TM10,'000','3','4'), CP01||Decode(TM10,'000','(台灣)','(非台灣)'), 0, '', '', 0, '', '', Sum(CP18), '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='T' " & strSQL23 & " Group By ST15, CP13, Decode(TM10,'000','3','4'), CP01||Decode(TM10,'000','(台灣)','(非台灣)') "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '6', CP01, 0, '', '', 0, '', '', Sum(CP18), '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='CFT' " & strSQL23 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
    strSql = "Select ST15, CP13, '8', CP01, 0, '', '', 0, '', '', Sum(CP18), '" & strUserNum & "' From  CaseProgress, Trademark, Staff Where TM01=CP01 And TM02=CP02 And TM03=CP03 And TM04=CP04 And CP13=ST01 And CP01='FCT' " & strSQL23 & "Group By ST15, CP13, CP01 "
    cnnConnection.Execute "Insert Into R100122 " & strSql
End If
strSql = "Select * From R100122 Where ID ='" & strUserNum & "' And Rownum < 2 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    '無動作
Else
   ShowNoData
   Screen.MousePointer = vbDefault
   DoTemp = False
   Exit Function
End If
CheckOC
DoTemp = True
End Function

Private Sub Form_Unload(Cancel As Integer)
Set frm100122_2 = Nothing
End Sub

Private Sub OpenWord()
Dim ii As Integer
Dim jj As Integer
Dim strSalesZone As String
Dim strSales As String
Dim intPage As Integer
Dim blnFirstPage As Boolean

' 顯示Word程式
On Error GoTo ERRORSECTION2
    blnFirstPage = True
    If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
    g_WordAp.Documents.add
    g_WordAp.Visible = True
    With g_WordAp.Application
        .WindowState = wdWindowStateMaximize
        '版面轉成橫向
        With .ActiveDocument.PageSetup
            .LineNumbering.Active = False
            .Orientation = wdOrientLandscape
            .TopMargin = CentimetersToPoints(3.17)
            .BottomMargin = CentimetersToPoints(3.17)
            .LeftMargin = CentimetersToPoints(2.54)
            .RightMargin = CentimetersToPoints(2.54)
            .Gutter = CentimetersToPoints(0)
            .HeaderDistance = CentimetersToPoints(1.5)
            .FooterDistance = CentimetersToPoints(1.75)
            .PageWidth = CentimetersToPoints(29.7)
            .PageHeight = CentimetersToPoints(21)
            .FirstPageTray = wdPrinterDefaultBin
            .OtherPagesTray = wdPrinterDefaultBin
            .SectionStart = wdSectionNewPage
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .VerticalAlignment = wdAlignVerticalTop
            .SuppressEndnotes = False
            .MirrorMargins = False
            .TwoPagesOnOne = False
            .GutterOnTop = False
            .CharsLine = 58
            .LinesPage = 23
        End With
        '加表格
        .ActiveDocument.Tables.add Range:=.Selection.Range, NumRows:=(Me.grdDataList1.Rows - 1) + IIf((Me.grdDataList1.Rows - 1) Mod 14 <> 0, Fix((Me.grdDataList1.Rows - 1) / 14) + 1, Fix((Me.grdDataList1.Rows - 1) / 14)) * 7, NumColumns:=10
        .Selection.Cells.Height = 19
        DoEvents
TitleParagraph:
        .Selection.SelectRow
        .Selection.Cells.Merge
        'Modified by Lydia 2016/09/06
        '.Selection.TypeText "智權人員" & IIf(frm100122_1.Txt1(0).Text = "1", "收", "發") & "文量比較表"
        .Selection.TypeText "智權人員收發文量比較表"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.MoveRight Unit:=wdCell
        '三欄
        .Selection.SelectRow
        .Selection.Cells.Merge
        .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
        .Selection.TypeText Me.Label1.Caption & Me.lbl1(0).Caption
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.MoveRight Unit:=wdCell
        
        .Selection.TypeText Me.Label2.Caption & Me.lbl1(1).Caption
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.MoveRight Unit:=wdCell
        
        .Selection.TypeText "列印日期：" & Format(strSrvDate(2), "###/##/##")
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.MoveRight Unit:=wdCell
                
        .Selection.SelectRow
        .Selection.Cells.Merge
        .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
        .Selection.TypeText Me.Label3.Caption & Me.lbl1(2).Caption
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        intPage = intPage + 1
        .Selection.TypeText "頁　　數：" & intPage
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.MoveRight Unit:=wdCell
            
        .Selection.SelectRow
        .Selection.Cells.Merge
        .Selection.Cells.Split NumRows:=1, NumColumns:=10, MergeBeforeSplit:=False
        .Selection.TypeText "業務區"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "智權人員"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "系統類別"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "比較時段1"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "增減"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "百分比"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        'edit by nickc 2006/02/09
        '.Selection.TypeText "統計時段"
        .Selection.TypeText "比較時段2"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "增減"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "百分比"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        'edit by nickc 2006/02/09
        '.Selection.TypeText "比較時段2"
        .Selection.TypeText "比較時段3"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        
        .Selection.SelectRow
        .Selection.Cells.Merge
        .Selection.Cells.Split NumRows:=1, NumColumns:=10, MergeBeforeSplit:=False
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText ChangeTStringToTDateString(frm100122_1.Txt1(21).Text)
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText ChangeTStringToTDateString(frm100122_1.Txt1(1).Text)
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText ChangeTStringToTDateString(frm100122_1.Txt1(23).Text)
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        
        .Selection.SelectRow
        .Selection.Cells.Merge
        .Selection.Cells.Split NumRows:=1, NumColumns:=10, MergeBeforeSplit:=False
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText "｜"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText "｜"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText "｜"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        
        .Selection.SelectRow
        .Selection.Cells.Merge
        .Selection.Cells.Split NumRows:=1, NumColumns:=10, MergeBeforeSplit:=False
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText ChangeTStringToTDateString(frm100122_1.Txt1(22).Text)
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText ChangeTStringToTDateString(frm100122_1.Txt1(2).Text)
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText ChangeTStringToTDateString(frm100122_1.Txt1(24).Text)
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.MoveRight Unit:=wdCell
        
        If intPage > 1 Then GoTo ReDoFor
        strSalesZone = ""
        strSales = ""
        For ii = 1 To Me.grdDataList1.Rows - 1
            If blnFirstPage = False And ii Mod 14 = 1 Then
                GoTo TitleParagraph
            End If
ReDoFor:
            blnFirstPage = False
'            For jj = 0 To Me.grdDataList1.Cols - 1
            For jj = 0 To 9
                If jj = 0 Then
                    If strSalesZone <> Me.grdDataList1.TextMatrix(ii, jj) Then
                        .Selection.TypeText Me.grdDataList1.TextMatrix(ii, jj)
                        strSalesZone = Me.grdDataList1.TextMatrix(ii, jj)
                    End If
                ElseIf jj = 1 Then
                    If strSales <> Me.grdDataList1.TextMatrix(ii, jj) Then
                        .Selection.TypeText Me.grdDataList1.TextMatrix(ii, jj)
                        strSales = Me.grdDataList1.TextMatrix(ii, jj)
                    End If
                Else
                    .Selection.TypeText Me.grdDataList1.TextMatrix(ii, jj)
                End If
                If jj <= 2 Then
                    .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                Else
                    .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
                End If
                If ii < Me.grdDataList1.Rows - 1 Then
                    .Selection.MoveRight Unit:=wdCell
                Else
'                    If jj <> Me.grdDataList1.Cols - 1 Then
                    If jj <> 9 Then
                        .Selection.MoveRight Unit:=wdCell
                    End If
                End If
            Next jj
        Next ii
        .Selection.WholeStory
        .Selection.Font.Name = "標楷體"
        .Selection.WholeStory
        .Selection.Font.Name = "Courier"
    End With
    Exit Sub
   
ERRORSECTION2:
Select Case Err.Number
Case 91:
    'Debug.Print "ERRORSECTION2:新增一個Word 頁面"
    g_WordAp.Documents.add
    Resume Next
Case 462:
    'Debug.Print "ERRORSECTION2:新增一個Word Application物件"
    Set g_WordAp = New Word.Application
    g_WordAp.Documents.add
    Resume Next
Case Else:
    MsgBox "錯誤 : " & Err.Description, vbCritical
    Exit Sub
End Select

End Sub

Private Sub DeleteFile()
On Error Resume Next

Kill m_strFilePathName
If Err.Number <> 0 Then Err.Clear

End Sub
