VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100105_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "以收/發文量查詢"
   ClientHeight    =   5736
   ClientLeft      =   5448
   ClientTop       =   3396
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8952
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   5490
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   5430
      Width           =   3450
   End
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "列印(&P)"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   5010
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6510
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
      Left            =   7740
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   10
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList1 
      Height          =   4515
      Left            =   30
      TabIndex        =   0
      Top             =   870
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   7959
      _Version        =   393216
      Cols            =   3
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
      _Band(0).Cols   =   3
   End
   Begin VB.Label lblMemo 
      Caption         =   "lblMemo"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3240
      TabIndex        =   28
      Top             =   645
      Width           =   3615
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "只統計新申請案："
      Height          =   180
      Left            =   7020
      TabIndex        =   27
      Top             =   1580
      Width           =   1440
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4590
      TabIndex        =   26
      Top             =   5460
      Width           =   855
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   9
      Left            =   7035
      TabIndex        =   25
      Top             =   4770
      Width           =   1695
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "中間接進來新案合計："
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   7020
      TabIndex        =   24
      Top             =   4560
      Width           =   1800
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "FCP工程師組別合計："
      Height          =   180
      Left            =   7020
      TabIndex        =   23
      Top             =   4050
      Width           =   1740
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   8
      Left            =   7035
      TabIndex        =   22
      Top             =   4260
      Width           =   1695
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   7
      Left            =   7035
      TabIndex        =   21
      Top             =   3750
      Width           =   1695
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "商申類別總計："
      Height          =   180
      Left            =   7020
      TabIndex        =   20
      Top             =   3540
      Width           =   1260
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   6
      Left            =   7410
      TabIndex        =   19
      Top             =   1280
      Width           =   1290
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   7020
      TabIndex        =   18
      Top             =   1060
      Width           =   900
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   5
      Left            =   7410
      TabIndex        =   17
      Top             =   800
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   7020
      TabIndex        =   16
      Top             =   560
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(1.收文 2.發文)"
      Height          =   180
      Left            =   1950
      TabIndex        =   15
      Top             =   645
      Width           =   1155
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   4
      Left            =   7035
      TabIndex        =   14
      Top             =   3225
      Width           =   1695
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   3
      Left            =   7035
      TabIndex        =   13
      Top             =   2715
      Width           =   1695
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   2
      Left            =   7035
      TabIndex        =   12
      Top             =   2205
      Width           =   1695
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "總計："
      Height          =   180
      Left            =   7020
      TabIndex        =   11
      Top             =   3015
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "多國案件合計："
      Height          =   180
      Left            =   7020
      TabIndex        =   10
      Top             =   2505
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "不計件數合計："
      Height          =   180
      Left            =   7020
      TabIndex        =   9
      Top             =   1995
      Width           =   1260
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   1
      Left            =   1140
      TabIndex        =   8
      Top             =   660
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "查詢別："
      Height          =   180
      Left            =   60
      TabIndex        =   7
      Top             =   645
      Width           =   720
   End
   Begin VB.Label lbl1 
      Caption         =   " "
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   2
      Top             =   432
      Width           =   3180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "查詢期間："
      Height          =   180
      Left            =   60
      TabIndex        =   1
      Top             =   435
      Width           =   900
   End
End
Attribute VB_Name = "frm100105_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/19 改成Form2.0(grdDataList改Fonts)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/10 日期欄已修改
'2009/11/16 整理 by sonia
'所有統計條件語法都顯示系統類別+統計條件+數量+中間新案數+商申類別數,再以欄位寬度=0決定是否顯示出來
Option Explicit

Dim strSql As String, i As Integer, j As Long, s As Integer, strTemp As Variant, intK As Integer
Dim strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL11 As String, strSQL22 As String, strSQL33 As String, strSQL44 As String, strSQL55 As String
Dim PLeft(0 To 8) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 10) As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim u As Long  'modify by Toni 2008/10/16
Dim strPrinter As String 'Add By Sindy 2015/2/9

Private Sub SetDataListWidth()
   'Modify By Sindy 2010/3/18
   'grdDataList1.Cols = 6   '2009/11/19 add by sonia
   grdDataList1.Cols = 7
   
   grdDataList1.row = 0
   grdDataList1.col = 0: grdDataList1.Text = "系統類別"
   Select Case frm100105_1.txt1(17)
      Case "2"
         grdDataList1.Text = "業務區"
         grdDataList1.ColWidth(0) = 1200
      Case "5", "6", "7", "9" '分系統類別
         grdDataList1.ColWidth(0) = 800
      'Added by Lydia 2025/08/06
      Case "C"
         grdDataList1.Text = "部門別"
         grdDataList1.ColWidth(0) = 1200
      'end 2025/08/06
      Case Else
         grdDataList1.ColWidth(0) = 0
   End Select
   grdDataList1.CellAlignment = flexAlignCenterCenter
   
   grdDataList1.col = 1: grdDataList1.Text = "統計條件"
   Select Case frm100105_1.txt1(17)
      Case "5"
         grdDataList1.ColWidth(1) = 1000
      'add by nickc 2007/05/15  加入代理人編號
      Case "6", "7"
         grdDataList1.ColWidth(1) = 3000
      'add by Toni 2008/10/16 加FCP工程師組別
      Case "9"
         grdDataList1.ColWidth(1) = 1300
      'Added by Lydia 2018/02/12 申請國家或洲別
      Case "B"
         grdDataList1.ColWidth(1) = 1500
      Case Else
         grdDataList1.ColWidth(1) = 1200
   End Select
   grdDataList1.CellAlignment = flexAlignCenterCenter
   
   grdDataList1.col = 2: grdDataList1.Text = "數量"
   grdDataList1.ColWidth(2) = 800
   grdDataList1.CellAlignment = flexAlignCenterCenter
   
   '2009/11/19 add by sonia
   grdDataList1.col = 3: grdDataList1.Text = "中間新案數"
   grdDataList1.ColWidth(3) = 1000
   grdDataList1.CellAlignment = flexAlignCenterCenter
   
   grdDataList1.col = 4: grdDataList1.Text = "類別數"
   'Added by Lydia 2017/02/03 洲別不顯示類別數
'   If frm100105_1.txt1(17) = "B" Then
'      grdDataList1.ColWidth(4) = 0
'   Else
      grdDataList1.ColWidth(4) = 1000
   'End If
   'end 2017/02/03
   grdDataList1.CellAlignment = flexAlignCenterCenter
   
   grdDataList1.ColWidth(5) = 0   '放代理人編號
   
   'Add By Sindy 2010/3/18
   grdDataList1.col = 6: grdDataList1.Text = "已收訴願量"
   grdDataList1.ColWidth(6) = 0
   grdDataList1.CellAlignment = flexAlignCenterCenter
End Sub

Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
      Case Else
   End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

'列印    第二階段   nick   91.07.31
Private Sub cmdPrinter_Click()
Dim k As Integer
   
   Screen.MousePointer = vbHourglass
   If grdDataList1.Rows <> 1 Then
      PUB_RestorePrinter Combo1 'Add By Sindy 2015/2/9
      Page = 1
      PrintTitle
      Dim intItem As Integer
      With grdDataList1
         For intItem = 1 To IIf(.Rows - 1 < 12, 12, .Rows - 1)   '右邊合計欄之行數要全印
            If iPrint >= 15000 Then
               Printer.NewPage
               Page = Page + 1
               PrintTitle
            End If
            If intItem > .Rows - 1 Then
               Erase strTemp3
            Else
               '2009/11/20 modify by sonia 全都搬至strTemp3,列印時再依gride欄位寬度決定是否列印
               Erase strTemp3
               For i = 0 To .Cols - 1
                  strTemp3(i) = Me.grdDataList1.TextMatrix(intItem, i)
               Next i
               '2009/11/20 end
            End If
            
            'Modify By Sindy 2010/3/18 案件性質增加一個欄位
            If frm100105_1.txt1(17) = "5" Then
               k = 7
            Else
            '2010/3/18 End
               k = 6
            End If
            If intItem = 1 Then
               strTemp3(k) = "不計件數合計："
            ElseIf intItem = 2 Then
               strTemp3(k + 1) = lbl1(2).Caption
            ElseIf intItem = 3 Then
               strTemp3(k) = "多國案件合計："
            ElseIf intItem = 4 Then
               strTemp3(k + 1) = lbl1(3).Caption
            ElseIf intItem = 5 Then
               strTemp3(k) = "總　計："
            ElseIf intItem = 6 Then
               strTemp3(k + 1) = lbl1(4).Caption
            ElseIf intItem = 7 Then
               strTemp3(k) = "商申類別總計："
            ElseIf intItem = 8 Then
               strTemp3(k + 1) = lbl1(7).Caption
            'add by Toni 2008/10/16
            ElseIf intItem = 9 Then
               strTemp3(k) = "FCP工程師組別："
            ElseIf intItem = 10 Then
               strTemp3(k + 1) = lbl1(8).Caption
            '2009/11/16 ADD BY SONIA
            ElseIf intItem = 11 Then
               strTemp3(k) = "中間接進來新案："
            ElseIf intItem = 12 Then
               strTemp3(k + 1) = lbl1(9).Caption
            End If
            PrintDatil
         Next intItem
      End With
      'Added by Lydia 2017/02/07 顯示提示
      If frm100105_1.txt1(17) = "B" Then
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print lblMemo.Caption
      End If
      'end 2017/02/07
      
      Printer.EndDoc
      PUB_RestorePrinter strPrinter 'Add By Sindy 2015/2/9
      ShowPrintOk
   Else
      MsgBox "沒有資料可以列印 !", vbCritical
   End If
   Screen.MousePointer = vbDefault
End Sub

Sub PrintTitle()
   GetPleft
   iPrint = 500
   Printer.Orientation = 1
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4500
   Printer.CurrentY = iPrint
   Printer.Print "收/發文量查詢"
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "系統類別：" & frm100105_1.txt1(3).Text
   
   iPrint = iPrint + 300
   'Add by Morgan 2007/1/16
   If frm100105_1.txt1(28) <> "" Then
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      If frm100105_1.txt1(28) = "A" Then
         Printer.Print "FC代理人性質：A.律師事務所"
      ElseIf frm100105_1.txt1(28) = "B" Then
         Printer.Print "FC代理人性質：B.公司直接委辦"
      Else
         Printer.Print "FC代理人性質：C.其他"
      End If
   End If
   'end 2007/1/16
   Printer.CurrentX = 4300
   Printer.CurrentY = iPrint
   Printer.Print IIf(frm100105_1.txt1(0) = "1", "收文日：", "發文日：") & Format(ChangeTStringToTDateString(frm100105_1.txt1(1)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(frm100105_1.txt1(2))
   'Added by Lydia 2016/02/25
   If Len(Trim(frm100105_1.txt1(37))) <> 0 Then
      Printer.CurrentX = 8500
      Printer.CurrentY = iPrint
      Printer.Print "僅統計新申請案"
   End If
   'end 2016/02/25
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 4300
   Printer.CurrentY = iPrint
   Printer.Print "申請國家 : " & Me.lbl1(5).Caption
   Printer.CurrentX = 8500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   iPrint = iPrint + 300
   'Add By Cheng 2003/09/12
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "業務區別：" & frm100105_1.txt1(6).Text & IIf(frm100105_1.txt1(6).Text <> "" Or frm100105_1.txt1(7).Text <> "", "－", "") & frm100105_1.txt1(7).Text
   'End 2003/09/12
   Printer.CurrentX = 4300
   Printer.CurrentY = iPrint
   Printer.Print "案件性質 : " & Me.lbl1(6).Caption
   Printer.CurrentX = 8500
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(90, "-")
   iPrint = iPrint + 300
   
   '2009/11/16 MODIFY BY SONIA 加印中間新案數欄
   Select Case frm100105_1.txt1(17)
      Case "1" '業務區
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "業務區"
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "數量"
         Printer.CurrentX = PLeft(2) - 200
         Printer.CurrentY = iPrint
         Printer.Print "中間新案數"
         If Me.grdDataList1.ColWidth(4) <> 0 Then
            Printer.CurrentX = PLeft(3) - 200
            Printer.CurrentY = iPrint
            Printer.Print "類別數"
         End If
      Case "2" '智權人員
         '2009/11/20 add by sonia 加業務區
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "業務區"
         '2009/11/20 end
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "智權人員"
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print "數量"
         Printer.CurrentX = PLeft(3) - 200
         Printer.CurrentY = iPrint
         Printer.Print "中間新案數"
         If Me.grdDataList1.ColWidth(4) <> 0 Then
            Printer.CurrentX = PLeft(4) - 200
            Printer.CurrentY = iPrint
            Printer.Print "類別數"
         End If
      Case "3" '申請國家
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "申請國家"
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "數量"
         Printer.CurrentX = PLeft(2) - 200
         Printer.CurrentY = iPrint
         Printer.Print "中間新案數"
         If Me.grdDataList1.ColWidth(4) <> 0 Then
            Printer.CurrentX = PLeft(3) - 200
            Printer.CurrentY = iPrint
            Printer.Print "類別數"
         End If
      Case "4" '申請人國籍
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "申請人國籍"
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "數量"
         Printer.CurrentX = PLeft(2) - 200
         Printer.CurrentY = iPrint
         Printer.Print "中間新案數"
         If Me.grdDataList1.ColWidth(4) <> 0 Then
            Printer.CurrentX = PLeft(3) - 200
            Printer.CurrentY = iPrint
            Printer.Print "類別數"
         End If
      Case "5" '案件性質
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "系統類別"
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "案件性質"
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print "數量"
         Printer.CurrentX = PLeft(3) - 200
         Printer.CurrentY = iPrint
         Printer.Print "中間新案數"
         If Me.grdDataList1.ColWidth(4) <> 0 Then
            Printer.CurrentX = PLeft(4) - 200
            Printer.CurrentY = iPrint
            Printer.Print "類別數"
         End If
         'Add By Sindy 2010/3/18
         If Me.grdDataList1.ColWidth(6) <> 0 Then
            Printer.CurrentX = PLeft(5) - 200
            Printer.CurrentY = iPrint
            Printer.Print "已收訴願量"
         End If
      'add by nickc 2006/09/01
      Case "6", "7" '代理人
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "系統類別"
         Printer.CurrentX = PLeft(1) + 200
         Printer.CurrentY = iPrint
         If frm100105_1.txt1(17) = "6" Then
            Printer.Print "FC代理人"
         Else
            Printer.Print "CF代理人"
         End If
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print "數量"
         Printer.CurrentX = PLeft(3) - 200
         Printer.CurrentY = iPrint
         Printer.Print "中間新案數"
         'edit by nickc 2007/05/15 加入代理人編號
         'If Me.grdDataList1.Cols = 4 Then
         If Me.grdDataList1.ColWidth(4) <> 0 Then
            Printer.CurrentX = PLeft(4) - 200
            Printer.CurrentY = iPrint
            Printer.Print "類別數"
         End If
      'add by nickc 2007/11/27
      Case "8" '代理人國籍
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "代理人國籍"
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "數量"
         Printer.CurrentX = PLeft(2) - 200
         Printer.CurrentY = iPrint
         Printer.Print "中間新案數"
         If Me.grdDataList1.ColWidth(4) <> 0 Then
            Printer.CurrentX = PLeft(3) - 200
            Printer.CurrentY = iPrint
            Printer.Print "類別數"
         End If
      'add by Toni 2008/10/16
      Case "9" '增加FCP工程師組別
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "系統類別"
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "FCP工程師組別"
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print "數量"
         Printer.CurrentX = PLeft(3) - 200
         Printer.CurrentY = iPrint
         Printer.Print "中間新案數"
         If Me.grdDataList1.ColWidth(4) <> 0 Then
            Printer.CurrentX = PLeft(4) - 200
            Printer.CurrentY = iPrint
            Printer.Print "類別數"
         End If
      'Add By Sindy 2014/7/9
      Case "A" '專利案件屬性
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "專利案件屬性"
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "數量"
         Printer.CurrentX = PLeft(2) - 200
         Printer.CurrentY = iPrint
         Printer.Print "中間新案數"
         If Me.grdDataList1.ColWidth(4) <> 0 Then
            Printer.CurrentX = PLeft(3) - 200
            Printer.CurrentY = iPrint
            Printer.Print "類別數"
         End If
      '2014/7/9 END
      'Added by Lydia 2017/02/03
      Case "B" '洲別
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "洲別"
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "數量"
         Printer.CurrentX = PLeft(2) - 200
         Printer.CurrentY = iPrint
         Printer.Print "中間新案數"
         If Me.grdDataList1.ColWidth(4) <> 0 Then
            Printer.CurrentX = PLeft(3) - 200
            Printer.CurrentY = iPrint
            Printer.Print "類別數"
         End If
      'Added by Lydia 2025/08/06
      Case "C" '承辦人
         Printer.CurrentX = PLeft(0)
         Printer.CurrentY = iPrint
         Printer.Print "部門別"
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print "承辦人"
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iPrint
         Printer.Print "數量"
         Printer.CurrentX = PLeft(3) - 200
         Printer.CurrentY = iPrint
         Printer.Print "中間新案數"
         If Me.grdDataList1.ColWidth(4) <> 0 Then
            Printer.CurrentX = PLeft(4) - 200
            Printer.CurrentY = iPrint
            Printer.Print "類別數"
         End If
      'end 2025/08/06
      Case Else
   End Select
   iPrint = iPrint + 300
   
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(90, "-")
   iPrint = iPrint + 300
End Sub

Sub GetPleft()
   Erase PLeft
   PLeft(0) = 500
   '2009/11/23 MODIFY BY SONIA 改依統計條件設定
   Select Case frm100105_1.txt1(17)
      'Modify By Sindy 2014/7/9 +A
      'Modified by Lydia 2017/02/03 +B
      Case "1", "3", "4", "8", "A", "B" '1業務區,3申請國家,4申請人國籍,8代理人國籍,A專利案件屬性,B.洲別
         PLeft(1) = 2500
         PLeft(2) = 4000
         PLeft(3) = 6000
         PLeft(4) = 8000
      'Modified by Lydia 2025/08/06 +承辦人C
      Case "2", "C" '智權人員
         PLeft(1) = 2500
         PLeft(2) = 4000
         PLeft(3) = 6000
         PLeft(4) = 8000
         PLeft(5) = 9000
      Case "5" '案件性質
         PLeft(1) = 1500
         PLeft(2) = 4000
         PLeft(3) = 5500
         PLeft(4) = 7000
         PLeft(5) = 8500
      Case "6", "7"  '代理人
         PLeft(1) = 1500
         PLeft(2) = 6700
         PLeft(3) = 7500
         PLeft(4) = 8900
      Case "9" '增加FCP工程師組別
         PLeft(1) = 2000
         PLeft(2) = 4000
         PLeft(3) = 6000
         PLeft(4) = 8000
   End Select
   'Modify By Sindy 2010/3/18
   If frm100105_1.txt1(17) = "5" Then
      PLeft(7) = 9800   '報表右邊統計數字標題
      PLeft(8) = 11000   '報表右邊統計數字
   Else
   '2010/3/18 End
      PLeft(6) = 9800   '報表右邊統計數字標題
      PLeft(7) = 11000   '報表右邊統計數字
   End If
End Sub

'Sub PrintDatil()
'   '2009/11/20 modify by sonia 依gride欄位寬度決定是否列印
'   j = 0
'   For i = 0 To 7
'      If i = 6 Then           '報表右邊統計數字標題
'         Printer.CurrentX = PLeft(i)
'         Printer.CurrentY = iPrint
'         Printer.Print strTemp3(i)
'         j = j + 1
'      ElseIf i = 7 Then       '報表右邊統計數字
'         Printer.CurrentX = PLeft(i) - Printer.TextWidth(strTemp3(i))
'         Printer.CurrentY = iPrint
'         Printer.Print strTemp3(i)
'         j = j + 1
'      ElseIf Me.grdDataList1.ColWidth(i) <> 0 Then
'         If IsNumeric(strTemp3(i)) Then
'            Printer.CurrentX = PLeft(j) - Printer.TextWidth(strTemp3(i)) '+ 500
'         Else
'            Printer.CurrentX = PLeft(j)
'         End If
'         Printer.CurrentY = iPrint
'         If frm100105_1.txt1(17) = "6" Or frm100105_1.txt1(17) = "7" Then
'            Printer.Print Mid(strTemp3(i), 1, 40)
'         Else
'            Printer.Print strTemp3(i)
'         End If
'         j = j + 1
'      End If
'   Next i
'   iPrint = iPrint + 300
'End Sub

Sub PrintDatil()
   '2009/11/20 modify by sonia 依gride欄位寬度決定是否列印
   j = 0
   'Modify By Sindy 2010/3/18
   'For i = 0 To 7
   For i = 0 To 8
      If frm100105_1.txt1(17) <> "5" And i > 7 Then Exit For
      If (frm100105_1.txt1(17) <> "5" And i = 6) Or _
         (frm100105_1.txt1(17) = "5" And i = 7) Then '報表右邊統計數字標題
         Printer.CurrentX = PLeft(i)
         Printer.CurrentY = iPrint
         Printer.Print strTemp3(i)
         j = j + 1
      ElseIf (frm100105_1.txt1(17) <> "5" And i = 7) Or _
         (frm100105_1.txt1(17) = "5" And i = 8) Then '報表右邊統計數字
         Printer.CurrentX = PLeft(i) - Printer.TextWidth(strTemp3(i))
         Printer.CurrentY = iPrint
         Printer.Print strTemp3(i)
         j = j + 1
      ElseIf Me.grdDataList1.ColWidth(i) <> 0 Then
         If IsNumeric(strTemp3(i)) Then
            Printer.CurrentX = PLeft(j) - Printer.TextWidth(strTemp3(i)) + 500
         Else
            Printer.CurrentX = PLeft(j)
         End If
         Printer.CurrentY = iPrint
         If (frm100105_1.txt1(17) = "6" Or frm100105_1.txt1(17) = "7") And i = 1 Then
            Printer.Print Mid(strTemp3(i), 1, 40)
         Else
            Printer.Print strTemp3(i)
         End If
         j = j + 1
      End If
   Next i
   iPrint = iPrint + 300
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   '92.04.16 nick
   cmdState = -1
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add by Sindy 2015/2/9
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Sindy 2015/2/9
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2015/2/9 END
   Set frm100105_2 = Nothing
End Sub

Sub StrMenu()
Dim StrSQLa As String
Dim strSQL2 As String      'add by nickc 2006/09/01
Dim StrSQLaN As String     'Add By Sindy 2009/11/06
Dim rsA As New ADODB.Recordset
Dim ii As Integer
'Dim dblTMKindCnt As Double '商申類別數
Dim jj As Integer          '2009/11/19 add by sonia
Dim m_Condition As String  '2009/11/19 add by sonia
Dim strPA158 As String 'Add By Sindy 2014/7/10

   Me.Enabled = False
   '讀出資料
   If DoTemp = False Then
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   '顯示表單資料
   lbl1(0).Caption = frm100105_1.txt1(1) + "－" + frm100105_1.txt1(2)
   lbl1(1).Caption = frm100105_1.txt1(0)
   lblMemo.Caption = "" 'Added by Lydia 2017/02/07
   
   'Add By Cheng 2003/05/28
   '申請國家
   Me.lbl1(5).Caption = ""
   If frm100105_1.txt1(9).Text <> "" Or frm100105_1.txt1(10).Text <> "" Then
      Me.lbl1(5).Caption = frm100105_1.txt1(9).Text & "－" & frm100105_1.txt1(10).Text
   End If
   '案件性質
   Me.lbl1(6).Caption = ""
   If frm100105_1.txt1(13).Text <> "" Or frm100105_1.txt1(14).Text <> "" Then
      Me.lbl1(6).Caption = frm100105_1.txt1(13).Text & "－" & frm100105_1.txt1(14).Text
   End If
   
    'Add by Lydia 2015/02/12 + 是否只統計新申請案
    If Len(Trim(frm100105_1.txt1(37))) <> 0 Then Me.Label13.Caption = Me.Label13.Caption & "Y"

   '清除畫面右邊統計數
   lbl1(2).Caption = ""    '不計件數合計
   lbl1(3).Caption = ""    '多國案件合計
   lbl1(4).Caption = ""    '總計
   lbl1(7).Caption = ""    '商申類別總計
   lbl1(8).Caption = ""    'FCP工程師組別合計
   lbl1(9).Caption = ""    '中間接進來新案合計
      
   'Modify By Cheng 2002/12/18
   '若選出的資料皆為NULL會造成"資料提供者或其他服務回傳電子郵件狀態"的訊息
   'strSQL = "SELECT R02004,R02005 FROM R100105 WHERE ID='" & strUserNum & "' "
   'edit by nick 2004/10/12
   'strSQL = "SELECT R02004,R02005 FROM R100105 WHERE (R02004 IS NOT NULL OR R02005 IS NOT NULL) AND ID='" & strUserNum & "' "
   'strSQL = "select x.i,y.i from (SELECT count(R02004) as i FROM R100105 WHERE R02004 IS NOT NULL  AND ID='" & strUserNum & "') x,(SELECT count(R02005) as i FROM R100105 WHERE R02005 IS NOT NULL  AND ID='" & strUserNum & "') y "
   strSql = "select x.i,y.i,u.i from (SELECT count(R02004) as i FROM R100105 WHERE R02004 IS NOT NULL  AND ID='" & strUserNum & "') x,(SELECT count(R02005) as i FROM R100105 WHERE R02005 IS NOT NULL  AND ID='" & strUserNum & "') y,(SELECT count(R02016) as i FROM R100105 WHERE R02016 IS NOT NULL  AND ID='" & strUserNum & "') u "
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
   'edit by nick 2004/10/12 因strSQL語法改寫
   '   Do While adoRecordset.EOF = False
   '      If UCase(adoRecordset.Fields(0)) = "N" Then
   '         i = i + 1
   '      End If
   '      If UCase(adoRecordset.Fields(1)) = "Y" Then
   '         j = j + 1
   '      End If
   '      adoRecordset.MoveNext
   '   Loop
      lbl1(2).Caption = str(adoRecordset.Fields(0))     '不計件數合計
      lbl1(3).Caption = str(adoRecordset.Fields(1))     '多國案件合計
   End If
   
'   grdDataList1.Cols = 2
   m_Condition = ""
   'Modified by Morgan 2018/7/30 +補 Order by 語法(O12不會自動以 group by 欄位排序)
   Select Case frm100105_1.txt1(17)
      Case "1"
           m_Condition = "業務區"   '2009/11/19 add by sonia
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加系統類別及商申類別數欄但存空字串
           strSql = "SELECT '' AS 系統類別,R02006 AS 業務區,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105 WHERE ID='" & strUserNum & "' GROUP BY R02006 ORDER BY R02006"
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加系統類別及商申類別數欄但存空字串
           StrSQLa = "SELECT '' AS 系統類別,R02006 AS 業務區,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105_1 WHERE ID='" & strUserNum & "' GROUP BY R02006 ORDER BY R02006"
      Case "2"
           m_Condition = "智權人員"   '2009/11/19 add by sonia
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加系統類別及商申類別數欄但存空字串
           '2009/11/20 MODIFY BY SONIA 因排序條件改抓業務區+智權人員故第一欄改業務區
           'strSQL = "SELECT '' AS 系統類別,R02007 AS 智權人員,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 商申類別數 " & _
                           " FROM R100105 WHERE ID='" & strUserNum & "' GROUP BY R02007"
           strSql = "SELECT R02006 AS 業務區,R02007 AS 智權人員,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105 WHERE ID='" & strUserNum & "' GROUP BY R02006,R02007 ORDER BY R02006,R02007"
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加系統類別及商申類別數欄但存空字串
           '2009/11/20 MODIFY BY SONIA 因排序條件改抓業務區+智權人員故第一欄改業務區
           'StrSQLa = "SELECT '' AS 系統類別,R02007 AS 智權人員,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 商申類別數 " & _
                           " FROM R100105_1 WHERE ID='" & strUserNum & "' GROUP BY R02007"
           StrSQLa = "SELECT R02006 AS 業務區,R02007 AS 智權人員,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105_1 WHERE ID='" & strUserNum & "' GROUP BY R02006,R02007 ORDER BY R02006,R02007"
      Case "3"
           m_Condition = "申請國家"   '2009/11/19 add by sonia
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加系統類別及商申類別數欄但存空字串
           'Added by Lydia 2017/02/07 法務顧問案歸台灣
           strSql = "UPDATE R100105 SET R02008='000' WHERE ID='" & strUserNum & "' AND R02003='LA' AND R02008 IS NULL"
           cnnConnection.Execute strSql
           strSql = "UPDATE R100105_1 SET R02008='000' WHERE ID='" & strUserNum & "' AND R02003='LA' AND R02008 IS NULL"
           cnnConnection.Execute strSql
           'end 2017/02/07
           strSql = "SELECT '' AS 系統類別,SUBSTR(R02008,1,3) AS 申請國家,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105 WHERE ID='" & strUserNum & "' GROUP BY SUBSTR(R02008,1,3) ORDER BY SUBSTR(R02008,1,3)"
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加系統類別及商申類別數欄但存空字串
           StrSQLa = "SELECT '' AS 系統類別,SUBSTR(R02008,1,3) AS 申請國家,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105_1 WHERE ID='" & strUserNum & "' GROUP BY SUBSTR(R02008,1,3) ORDER BY SUBSTR(R02008,1,3)"
      Case "4"
           m_Condition = "申請人國籍"   '2009/11/19 add by sonia
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加系統類別及商申類別數欄但存空字串
           strSql = "SELECT '' AS 系統類別,substr(R02009,1,3) AS 申請人國籍,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105 WHERE ID='" & strUserNum & "' GROUP BY substr(R02009,1,3) ORDER BY substr(R02009,1,3)"
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加系統類別及商申類別數欄但存空字串
           StrSQLa = "SELECT '' AS 系統類別,substr(R02009,1,3) AS 申請人國籍,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105_1 WHERE ID='" & strUserNum & "' GROUP BY substr(R02009,1,3) ORDER BY substr(R02009,1,3)"
      Case "5"
           m_Condition = "案件性質"   '2009/11/19 add by sonia
           '2009/11/19 modify by sonia 加商申類別數欄但存空字串
           'Memo by Lydia 2024/12/02 若統計類別=5的條件有異動，請一併變更frm090642-收文量、發文量
           strSql = "SELECT R02003 AS 系統類別,R02010 AS 案件性質,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 新案件數,'' AS 類別數 " & _
                           " FROM R100105 WHERE ID='" & strUserNum & "' GROUP BY R02003,R02010 ORDER BY R02003,R02010"
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加商申類別數欄但存空字串
           StrSQLa = "SELECT R02003 AS 系統類別,R02010 AS 案件性質,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105_1 WHERE ID='" & strUserNum & "' GROUP BY R02003,R02010 ORDER BY R02003,R02010"
      'add by nickc 2006/09/01
      Case "6", "7"
           '2009/11/19 ADD BY SONIA
           If frm100105_1.txt1(17) = "6" Then
               m_Condition = "FC代理人"
           ElseIf frm100105_1.txt1(17) = "7" Then
               m_Condition = "CF代理人"
           End If
           '2009/11/19 END
           strSQL2 = "decode(fa10,'000',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'020',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),'013',NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65))"
           'edit by nickc 2007/05/15 加入代理人編號
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加商申類別數欄但存空字串
           '2009/11/20 MODIFY BY SONIA 排序條件改系統類別+大寫(代理人名稱)但資料仍顯示原資料
           'strSQL = "SELECT R02003 AS 系統類別," & strSQL2 & " AS 代理人,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 商申類別數 ," & _
                           "fa01||fa02" & _
                           " FROM R100105,fagent WHERE ID='" & strUserNum & "' and substr(r02011,1,8)=fa01(+) and substr(r02011,9,1)=fa02(+)  GROUP BY R02003," & strSQL2 & ",fa01||fa02 "
           strSql = "SELECT R02003 AS 系統類別," & strSQL2 & " AS 代理人,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 ," & _
                           "fa01||fa02,UPPER(" & strSQL2 & ") " & _
                           " FROM R100105,fagent WHERE ID='" & strUserNum & "' and substr(r02011,1,8)=fa01(+) and substr(r02011,9,1)=fa02(+)  GROUP BY R02003,UPPER(" & strSQL2 & ")," & strSQL2 & ",fa01||fa02 ORDER BY R02003,UPPER(" & strSQL2 & ")"
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加商申類別數欄但存空字串
           '2009/11/20 MODIFY BY SONIA 排序條件改系統類別+大寫(代理人名稱)
           'StrSQLa = "SELECT R02003 AS 系統類別," & strSQL2 & " AS 代理人,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 商申類別數 ," & _
                           "fa01||fa02" & _
                           " FROM R100105_1,fagent WHERE ID='" & strUserNum & "' and substr(r02011,1,8)=fa01(+) and substr(r02011,9,1)=fa02(+) GROUP BY R02003," & strSQL2 & ",fa01||fa02 "
           StrSQLa = "SELECT R02003 AS 系統類別," & strSQL2 & " AS 代理人,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 ," & _
                           "fa01||fa02,UPPER(" & strSQL2 & ") " & _
                           " FROM R100105_1,fagent WHERE ID='" & strUserNum & "' and substr(r02011,1,8)=fa01(+) and substr(r02011,9,1)=fa02(+) GROUP BY R02003,UPPER(" & strSQL2 & ")," & strSQL2 & ",fa01||fa02 ORDER BY R02003,UPPER(" & strSQL2 & ")"
      'add by nickc 2007/11/27 加入代理人國籍
      Case "8"
           m_Condition = "代理人國籍"   '2009/11/19 add by sonia
           strSQL2 = "substr(fa10,1,3)"
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加系統類別及商申類別數欄但存空字串
           strSql = "SELECT '' AS 系統類別," & strSQL2 & " AS 代理人國籍,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105,fagent WHERE ID='" & strUserNum & "' and substr(r02011,1,8)=fa01(+) and substr(r02011,9,1)=fa02(+)  GROUP BY " & strSQL2 & " ORDER BY " & strSQL2
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加系統類別及商申類別數欄但存空字串
           StrSQLa = "SELECT '' AS 系統類別," & strSQL2 & " AS 代理人國籍,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105_1,fagent WHERE ID='" & strUserNum & "' and substr(r02011,1,8)=fa01(+) and substr(r02011,9,1)=fa02(+) GROUP BY " & strSQL2 & " ORDER BY " & strSQL2
           '2008/11/4 END
      'add by toni 2008/10/15 增加FCP工程師組別
      Case "9"
           m_Condition = "FCP工程師組別"   '2009/11/19 add by sonia
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加商申類別數欄但存空字串
           strSql = "SELECT R02003 AS 系統類別,R02016 AS FCP工程師組別,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105 WHERE ID='" & strUserNum & "' GROUP BY R02003,R02016 ORDER BY R02003,R02016"
           'Modify By Sindy 2009/11/06 加中間新案數欄
           '2009/11/19 modify by sonia 加商申類別數欄但存空字串
           StrSQLa = "SELECT R02003 AS 系統類別,R02016 AS FCP工程師組別,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105_1 WHERE ID='" & strUserNum & "' GROUP BY R02003,R02016 ORDER BY R02003,R02016"
      'end 2008/10/15
      'Add By Sindy 2014/7/9
      Case "A"
           m_Condition = "專利案件屬性"
           strPA158 = "decode(length(R02006),1,'',decode(R02006,'21','11','22','12','23','13',R02006))"
           strSql = "SELECT '' AS 系統類別," & strPA158 & " AS 專利案件屬性,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105 WHERE ID='" & strUserNum & "' GROUP BY " & strPA158 & " ORDER BY " & strPA158
           StrSQLa = "SELECT '' AS 系統類別," & strPA158 & " AS 專利案件屬性,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105_1 WHERE ID='" & strUserNum & "' GROUP BY " & strPA158 & " ORDER BY " & strPA158
      '2014/7/9 END
      'Added by Lydia 2017/02/03
      Case "B"
           'Modified by Lydia 2018/02/12 +申請
           m_Condition = "申請國家或洲別"
           lblMemo.Caption = "PS.亞洲含台灣、大陸和日本；美洲含美國"
           'Added by Lydia 2017/02/07 法務顧問案歸台灣
           strSql = "UPDATE R100105 SET R02008='000' WHERE ID='" & strUserNum & "' AND R02003='LA' AND R02008 IS NULL"
           cnnConnection.Execute strSql
           strSql = "UPDATE R100105_1 SET R02008='000' WHERE ID='" & strUserNum & "' AND R02003='LA' AND R02008 IS NULL"
           cnnConnection.Execute strSql
           'end 2017/02/07
           '台灣
           strSql = "SELECT '01' 系統類別,SUBSTR(R02008,1,3) AS 申請國家或洲別 ,COUNT(*) 數量," & _
                    "SUM(DECODE(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數,NA03 " & _
                    "FROM R100105,NATION WHERE ID='" & strUserNum & "' AND R02008='000' AND R02008=NA01(+) GROUP BY SUBSTR(R02008,1,3),NA03 "
           '大陸
           strSql = strSql & "UNION ALL SELECT '02' 系統類別,SUBSTR(R02008,1,3) AS 申請國家或洲別 ,COUNT(*) 數量," & _
                    "SUM(DECODE(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數,NA03 " & _
                    "FROM R100105,NATION WHERE ID='" & strUserNum & "' AND R02008='020' AND R02008=NA01(+) GROUP BY SUBSTR(R02008,1,3),NA03 "
           '美國
           strSql = strSql & "UNION ALL SELECT '03' 系統類別,SUBSTR(R02008,1,3) AS 申請國家或洲別 ,COUNT(*) 數量," & _
                    "SUM(DECODE(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數,NA03 " & _
                    "FROM R100105,NATION WHERE ID='" & strUserNum & "' AND R02008='101' AND R02008=NA01(+) GROUP BY SUBSTR(R02008,1,3),NA03 "
           '日本
           strSql = strSql & "UNION ALL SELECT '04' 系統類別,SUBSTR(R02008,1,3) AS 申請國家或洲別 ,COUNT(*) 數量," & _
                    "SUM(DECODE(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數,NA03 " & _
                    "FROM R100105,NATION WHERE ID='" & strUserNum & "' AND R02008='011' AND R02008=NA01(+) GROUP BY SUBSTR(R02008,1,3),NA03 "
           '各洲
           strSql = strSql & "UNION ALL SELECT '05' 系統類別, DECODE(SUBSTR(NA02,1,2),'A0','C0','B0','C0',SUBSTR(NA02,1,2)) AS 申請國家或洲別 ,COUNT(*) 數量, " & _
                    "SUM(DECODE(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數,DECODE(SUBSTR(NA02,1,2),'A0','亞洲','B0','亞洲','C0','亞洲','C1','美洲','C2','歐洲','C3','非洲','C4','大洋洲',NA02) " & _
                    "FROM R100105,NATION WHERE ID='" & strUserNum & "' AND R02008=NA01(+) GROUP BY DECODE(SUBSTR(NA02,1,2),'A0','C0','B0','C0',SUBSTR(NA02,1,2)), " & _
                    "DECODE(SUBSTR(NA02,1,2),'A0','亞洲','B0','亞洲','C0','亞洲','C1','美洲','C2','歐洲','C3','非洲','C4','大洋洲',NA02) " & _
                    "ORDER BY 1,2 "
           StrSQLa = Replace(strSql, "R100105", "R100105_1")
      'Added by Lydia 2025/08/06
      Case "C"   '承辦人
           strSql = "SELECT '' AS 系統類別,R02007 AS 承辦人,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105 WHERE ID='" & strUserNum & "' GROUP BY R02007 ORDER BY R02007"
           StrSQLa = "SELECT '' AS 系統類別,R02007 AS 承辦人,COUNT(*) AS 數量," & _
                           "sum(decode(R02017,'Y',1,0)) AS 中間新案數,'' AS 類別數 " & _
                           " FROM R100105_1 WHERE ID='" & strUserNum & "' GROUP BY R02007 ORDER BY R02007"
      'end 2025/08/06
      Case Else
   End Select
   
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set grdDataList1.Recordset = adoRecordset
'   IntK = adoRecordset.RecordCount
   '2009/11/19 ADD BY SONIA
   SetDataListWidth     '表頭置中
   Me.grdDataList1.TextMatrix(0, 1) = m_Condition  '放入統計條件
   '2009/11/19 end
   
   'Add By Cheng 2003/12/31 加類別數
   If rsA.State <> adStateClosed Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      While Not rsA.EOF
         For ii = 1 To Me.grdDataList1.Rows - 1
            If Me.grdDataList1.TextMatrix(ii, 0) = "" & rsA.Fields(0).Value And Me.grdDataList1.TextMatrix(ii, 1) = "" & rsA.Fields(1).Value Then
               Me.grdDataList1.TextMatrix(ii, 4) = "" & rsA.Fields(2).Value
               Me.grdDataList1.row = ii
               Me.grdDataList1.col = Me.grdDataList1.Cols - 1
               Exit For
            End If
         Next ii
         rsA.MoveNext
      Wend
   '2009/11/19 add by sonia 無類別數欄則隱藏
   Else
      Me.grdDataList1.ColWidth(4) = 0
   '2009/11/19 end
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   'Add By Sindy 2010/3/18 已收訴願量
   If frm100105_1.txt1(17) = "5" Then
      For ii = 1 To Me.grdDataList1.Rows - 1
         If Trim(Me.grdDataList1.TextMatrix(ii, 0)) = "FCT" And Trim(Me.grdDataList1.TextMatrix(ii, 1)) = "1002" Then
            CheckOC
            If frm100105_1.txt1(0) = "1" Then
               '收文量
               strSQL2 = "SELECT count(*) " & _
                                 "FROM R100105 A,NEXTPROGRESS,CASEPROGRESS B " & _
                                 "WHERE A.R02003='FCT' AND A.R02010='1002' " & _
                                 "AND A.R02001>='" & ChangeTStringToTDateString(frm100105_1.txt1(1)) & "' AND A.R02001<='" & ChangeTStringToTDateString(frm100105_1.txt1(2)) & "' " & _
                                 "AND A.R02015=NP01 AND A.R02003=NP02 AND A.R02014=NP03||NP04||NP05 " & _
                                 "AND '401'=NP07 AND NP07=B.CP10 " & _
                                 "AND A.R02015=B.CP43 " & _
                                 "AND ID='" & strUserNum & "' " & _
                                 "AND B.CP04='00' AND B.CP09<'B' "
            Else
               '發文量
               strSQL2 = "SELECT count(*) " & _
                                 "FROM R100105 A,NEXTPROGRESS,CASEPROGRESS B " & _
                                 "WHERE A.R02003='FCT' AND A.R02010='1002' " & _
                                 "AND A.R02002>='" & ChangeTStringToTDateString(frm100105_1.txt1(1)) & "' AND A.R02002<='" & ChangeTStringToTDateString(frm100105_1.txt1(2)) & "' " & _
                                 "AND A.R02015=NP01 AND A.R02003=NP02 AND A.R02014=NP03||NP04||NP05 " & _
                                 "AND '401'=NP07 AND NP07=B.CP10 " & _
                                 "AND A.R02015=B.CP43 " & _
                                 "AND ID='" & strUserNum & "' " & _
                                 "AND B.CP04='00' AND B.CP09<'B' "
            End If
            rsA.CursorLocation = adUseClient
            rsA.Open strSQL2, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Me.grdDataList1.TextMatrix(ii, 6) = "" & rsA.Fields(0)
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Me.grdDataList1.row = ii
            Me.grdDataList1.col = Me.grdDataList1.Cols - 1
            Me.grdDataList1.ColWidth(6) = 1000
            Exit For
         End If
      Next ii
   End If
   
   'Add By Sindy 2009/11/06
   '中間接進來新案合計
   CheckOC
   StrSQLaN = "SELECT sum(decode(R02017,'Y',1,0)) AS 中間接進來新案合計" & _
                   " FROM R100105 WHERE ID='" & strUserNum & "' "
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open StrSQLaN, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      lbl1(9).Caption = str(adoRecordset.Fields(0).Value)
   End If
   '2009/11/06 End
   'Modify By Sindy 2013/4/18 商申類別總計
   CheckOC
   StrSQLaN = "SELECT count(*)" & _
                   " FROM R100105_1 WHERE ID='" & strUserNum & "' and R02010='101'"
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open StrSQLaN, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      lbl1(7).Caption = str(adoRecordset.Fields(0).Value)
   End If
   '2013/4/18 End
   
   j = 0 ': dblTMKindCnt = 0
   For i = 1 To grdDataList1.Rows - 1
      '2009/11/19 modify by sonia 統計條件統一放 col=1
      grdDataList1.col = 1
      grdDataList1.row = i
      'Modify By Sindy 2014/7/10 +And frm100105_1.txt1(17) <> "A"
      If frm100105_1.txt1(17) <> "6" And frm100105_1.txt1(17) <> "7" And frm100105_1.txt1(17) <> "A" Then
         Select Case frm100105_1.txt1(17)
            Case "1"
               strSql = "SELECT A0902 FROM ACC090 WHERE A0901='" & grdDataList1.Text & "'"
            'Modified by Lydia 2025/08/06 +承辦人C
            Case "2", "C"
               '2009/11/20 MODIFY BY SONIA 因排序條件改抓業務區+智權人員故第一欄改業務區
               'strSQL = "SELECT ST02 FROM STAFF WHERE  ST01='" & grdDataList1.Text & "'"
               strSql = Me.grdDataList1.TextMatrix(i, 0)
               'Modified by Lydia 2023/12/20 用Left Join
               'strSql = "SELECT ST02,A0902 FROM STAFF,ACC090 WHERE A0901='" & strSql & "' AND ST01='" & grdDataList1.Text & "'"
               strSql = "SELECT ST02,A0902 FROM STAFF,ACC090 WHERE ST01='" & grdDataList1.Text & "' AND ST15=A0901(+) "
            Case "3", "4", "8"
               strSql = "SELECT NA03 FROM NATION WHERE NA01='" & grdDataList1.Text & "'"
            Case "5"
               strSql = Me.grdDataList1.TextMatrix(i, 0)
               strSql = "SELECT Decode(CPM03,'（無）',CPM04,CPM03) FROM CASEPROPERTYMAP WHERE CPM01='" & strSql & "' AND CPM02='" & grdDataList1.Text & "'"
            'add by toni 2008/10/15 增加FCP工程師組別
            Case "9"
                strSql = Me.grdDataList1.TextMatrix(i, 0)
                If Trim(grdDataList1.Text) = "" Then
                   strSql = ""
                Else
                   'Modified by Morgan 2012/3/8 不必限制系統否則FMP案不會顯示
                   'strSql = "Select '" & PUB_GetFCPGrpName(grdDataList1.Text) & "' from dual where '" & grdDataList1.Text & "' is not null and '" & strSql & "' in ('FCP','FG')"
                   strSql = "Select '" & PUB_GetFCPGrpName(grdDataList1.Text) & "' from dual where '" & grdDataList1.Text & "' is not null"
                End If
            'end 2008/10/15
            'Added by Lydia 2017/02/07
            Case "B"
                Me.grdDataList1.TextMatrix(i, 1) = Me.grdDataList1.TextMatrix(i, 5)
         End Select
         CheckOC
         'Modified by Lydia 2017/02/07 洲別不計FCP工程師組別 (+ Or frm100105_1.txt1(17) = "B")
         If strSql = "" Or frm100105_1.txt1(17) = "B" Then  '無FCP工程師組別者不計入右邊FCP工程師組別合計
         Else
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
               If Not IsNull(adoRecordset.Fields(0)) Then
                  grdDataList1.Text = adoRecordset.Fields(0)
               End If
               'add by toni 2008/10/16
               If frm100105_1.txt1(17) = "9" Then
                  grdDataList1.col = 2
                  u = u + Val(grdDataList1.Text)
               '2009/11/20 ADD BY SONIA col=0放業務區名稱
               'Modified by Lydia 2025/08/06 +承辦人C
               ElseIf frm100105_1.txt1(17) = "2" Or frm100105_1.txt1(17) = "C" Then
                  grdDataList1.col = 0
                  grdDataList1.Text = adoRecordset.Fields(1)
               '2009/11/20 END
               End If
               '2008/10/16 end
            End If
            CheckOC
         End If
      'add by nickc 2007/05/15 加入代理人編號
      Else
         If frm100105_1.txt1(17) = "6" Or frm100105_1.txt1(17) = "7" Then
            Me.grdDataList1.TextMatrix(i, 1) = Me.grdDataList1.TextMatrix(i, 5) & " " & Me.grdDataList1.TextMatrix(i, 1)
            'Add By Sindy 2010/3/18
            Me.grdDataList1.TextMatrix(i, 5) = " "
            Me.grdDataList1.TextMatrix(i, 6) = " "
            '2010/3/18 End
         'Add By Sindy 2014/7/10
         ElseIf frm100105_1.txt1(17) = "A" Then
            If Me.grdDataList1.TextMatrix(i, 1) = "11" Then
               Me.grdDataList1.TextMatrix(i, 1) = "機械"
            ElseIf Me.grdDataList1.TextMatrix(i, 1) = "12" Then
               Me.grdDataList1.TextMatrix(i, 1) = "電子電機"
            ElseIf Me.grdDataList1.TextMatrix(i, 1) = "13" Then
               Me.grdDataList1.TextMatrix(i, 1) = "化學生醫"
            ElseIf Me.grdDataList1.TextMatrix(i, 1) = "31" Then
               Me.grdDataList1.TextMatrix(i, 1) = "整體"
            ElseIf Me.grdDataList1.TextMatrix(i, 1) = "32" Then
               Me.grdDataList1.TextMatrix(i, 1) = "部分"
            ElseIf Me.grdDataList1.TextMatrix(i, 1) = "33" Then
               Me.grdDataList1.TextMatrix(i, 1) = "圖像"
            ElseIf Me.grdDataList1.TextMatrix(i, 1) = "34" Then
               Me.grdDataList1.TextMatrix(i, 1) = "成組"
            End If
         '2014/7/10 END
         End If
      End If
      
      '類別數=>總計
      grdDataList1.col = 2
      'Added by Lydia 2017/02/03 因為洲別有另外顯示台灣、大陸、美國和日本，所以總計要排除
      If frm100105_1.txt1(17) = "B" And Val(Me.grdDataList1.TextMatrix(i, 0)) < 5 Then
      Else
         j = j + Val(grdDataList1.Text)
      End If
      'end 2017/02/03
      
      'dblTMKindCnt = dblTMKindCnt + Val(Me.grdDataList1.TextMatrix(i, 4))
      '數字欄右靠
      For jj = 2 To Me.grdDataList1.Cols - 1
         Me.grdDataList1.row = i
         Me.grdDataList1.col = jj
         Me.grdDataList1.CellAlignment = flexAlignRightCenter
      Next jj
   
   Next i
   lbl1(4).Caption = str(j)
   'lbl1(7).Caption = str(dblTMKindCnt)
   lbl1(8).Caption = str(u) 'add by toni 2008/10/16
   
   grdDataList1.Visible = True
   Me.Enabled = True
End Sub

Function DoTemp() As Boolean
'911024 nick 指定使用 index
Dim IndexString As String
Dim intK As Integer, IntK1 As Integer
'Add By Cheng 2003/12/31
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim dblTMKindCnt As Double '類別數
Dim ii As Integer
Dim strCP31 As String      '2009/11/20 ADD BY SONIA
Dim strMCTF As String, strCU13 As String 'Add by Amy 2017/02/24
   
   'If frm100105_1.txt1(0) = "1" Then
   '    IndexString = " /*+ index(caseprogress idxcp010203040526) */ "
   '    'IndexString = " /*+ index(caseprogress idxcp010526) */ "
   'Else
   '    IndexString = " /*+ index(caseprogress idxcp275710) */ "
   'End If
   frm100105_1.Hide
   
   j = 0
   cnnConnection.Execute "DELETE FROM R100105 where id='" & strUserNum & "' "
   'Add By Cheng 2003/12/31 For 商申類別數
   cnnConnection.Execute "DELETE FROM R100105_1 where id='" & strUserNum & "' "
   
   strSQL1 = ""
   strSQL2 = ""
   StrSQL3 = ""
   StrSQL4 = ""
   strSQL5 = ""
   
   If frm100105_1.txt1(0) = "1" Then
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label1 & "收文" 'Add By Sindy 2010/01/22
   ElseIf frm100105_1.txt1(0) = "2" Then
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label1 & "發文" 'Add By Sindy 2010/01/22
   End If
   
   '組合條件
   '查詢別-收文
   If frm100105_1.txt1(0) = "1" Then
      If Len(Trim(frm100105_1.txt1(1))) <> 0 Then
         strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(frm100105_1.txt1(1))) & " "
      End If
      If Len(Trim(frm100105_1.txt1(2))) <> 0 Then
         strSQL1 = strSQL1 & " AND CP05<=" & Val(ChangeTStringToWString(frm100105_1.txt1(2))) & " "
      End If
      
      'Added by Lydia 2016/09/06  +判斷未取消收文 CP159=0
      strSQL1 = strSQL1 & " AND CP159=0 "
      
      If Len(Trim(frm100105_1.txt1(1))) <> 0 Or Len(Trim(frm100105_1.txt1(2))) <> 0 Then
         pub_QL05 = pub_QL05 & ";收文" & frm100105_1.Label2 & frm100105_1.txt1(1) & "-" & frm100105_1.txt1(2) 'Add By Sindy 2010/01/22
      End If
   '查詢別-發文
   Else
      If Len(Trim(frm100105_1.txt1(1))) <> 0 Then
         strSQL1 = strSQL1 + " AND CP27>=" & Val(ChangeTStringToWString(frm100105_1.txt1(1))) & " "
      End If
      If Len(Trim(frm100105_1.txt1(2))) <> 0 Then
         strSQL1 = strSQL1 & " AND CP27<=" & Val(ChangeTStringToWString(frm100105_1.txt1(2))) & " "
      End If
      If Len(Trim(frm100105_1.txt1(1))) <> 0 Or Len(Trim(frm100105_1.txt1(2))) <> 0 Then
         pub_QL05 = pub_QL05 & ";發文" & frm100105_1.Label2 & frm100105_1.txt1(1) & "-" & frm100105_1.txt1(2) 'Add By Sindy 2010/01/22
      End If
   End If
   
   If Len(Trim(frm100105_1.txt1(4))) = 0 Then
      strSQL1 = strSQL1 + " AND cp26 IS NULL "
   Else
      pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label14, 11) & frm100105_1.txt1(4) 'Add By Sindy 2010/01/22
   End If
   If Len(Trim(frm100105_1.txt1(5))) = 0 Then
      strSQL1 = strSQL1 + " AND cp21 IS NULL "
   Else
      pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label9(0), 9) & frm100105_1.txt1(5) 'Add By Sindy 2010/01/22
   End If
   If Len(Trim(frm100105_1.txt1(6))) <> 0 Then
      strSQL1 = strSQL1 + " AND cp12>='" & frm100105_1.txt1(6) & "' "
   End If
   If Len(Trim(frm100105_1.txt1(7))) <> 0 Then
      strSQL1 = strSQL1 + " AND cp12<='" & frm100105_1.txt1(7) & "' "
   End If
   If Len(Trim(frm100105_1.txt1(6))) <> 0 Or Len(Trim(frm100105_1.txt1(7))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label4 & frm100105_1.txt1(6) & "-" & frm100105_1.txt1(7) 'Add By Sindy 2010/01/22
   End If
   If Len(Trim(frm100105_1.txt1(8))) <> 0 Then
      'Modify by Amy 2017/02/24 智權人員加MCTF
      If Left(frm100105_1.txt1(8), 4) = "MCTM" Or Left(frm100105_1.txt1(8), 4) = "MCTF" Then
        '下MCTM則三組分組都抓
        'Modify by Amy 2019/07/19 多增加MCTF04/05 且可能有離職人員無法記錄到,故增加 MCTMember
        If Left(frm100105_1.txt1(8), 5) = "MCTM" Then
'              strMCTF = ",'" & Replace(Pub_GetSpecMan("MCTF0", True), ";", "','") & "','MCTF01','MCTF02','MCTF03' "
'              strCU13 = " And SubStr(F1.fa120,1,5)='MCTF0' "
              strMCTF = ",'" & Replace(Pub_GetSpecMan("MCTMember", True), ";", "','") & "' "
              strCU13 = " And SubStr(cp161,1,4)='MCTF' "
        Else
              strMCTF = ",'" & Replace(Pub_GetSpecMan(frm100105_1.txt1(8)), ";", "','") & "','" & frm100105_1.txt1(8) & "' "
'              strCU13 = " And F1.fa120='" & frm100105_1.Txt1(8) & "' "
              strCU13 = " And cp161='" & frm100105_1.txt1(8) & "' "
        End If
        'strSQL1 = strSQL1 + " AND CP13 in (" & Mid(strMCTF, 2) & ") "
        strSQL1 = strSQL1 + " AND CP161 in (" & Mid(strMCTF, 2) & ") "
        'end 2019/07/19
        strSQL2 = strSQL2 + " And TM44 is not null" & strCU13
        strSQL22 = strSQL22 + " And TM44 is not null" & strCU13
        'Add by Amy 2017/06/06 法務也加 ex:1060508 86048收文之L-005714不加
        StrSQL3 = StrSQL3 + " And LC22 is not null" & strCU13
        strSQL33 = strSQL33 + " And LC22 is not null" & strCU13
        'end 2017/06/06
        strSQL5 = strSQL5 + " And SP26 is not null" & strCU13
        strSQL55 = strSQL55 + " And SP26 is not null" & strCU13
      Else
            strSQL1 = strSQL1 + " AND cp13='" & frm100105_1.txt1(8) & "' "
      End If
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label13 & frm100105_1.txt1(8) & frm100105_1.lbl1(0) 'Add By Sindy 2010/01/22
   End If
   If Len(Trim(frm100105_1.txt1(11))) <> 0 Then
      strSQL1 = strSQL1 + " AND cu10>='" & frm100105_1.txt1(11) & "' "
   End If
   If Len(Trim(frm100105_1.txt1(12))) <> 0 Then
      strSQL1 = strSQL1 + " AND cu10<='" & frm100105_1.txt1(12) & "z' "
   End If
   If Len(Trim(frm100105_1.txt1(11))) <> 0 Or Len(Trim(frm100105_1.txt1(12))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label5(1) & frm100105_1.txt1(11) & "-" & frm100105_1.txt1(12) 'Add By Sindy 2010/01/22
   End If
   If Len(Trim(frm100105_1.txt1(13))) <> 0 Then
      strSQL1 = strSQL1 + " AND cp10>='" & frm100105_1.txt1(13) & "' "
   End If
   If Len(Trim(frm100105_1.txt1(14))) <> 0 Then
      strSQL1 = strSQL1 + " AND cp10<='" & frm100105_1.txt1(14) & "' "
   End If
   If Len(Trim(frm100105_1.txt1(13))) <> 0 Or Len(Trim(frm100105_1.txt1(14))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label6 & frm100105_1.txt1(13) & "-" & frm100105_1.txt1(14) 'Add By Sindy 2010/01/22
   End If
   
    'Add by Lydia 2015/02/12 + 是否只統計新申請案 (服務業務之新申請案案件性質為801、802、805、806) 'Memo by Lydia 2024/12/02 若統計類別=5的條件有異動，請一併變更frm090642-收文量、發文量
    If Len(Trim(frm100105_1.txt1(37))) <> 0 Then
        strExc(1) = "": strExc(2) = "": strExc(3) = ""
        'Modified by Lydia 2016/02/24 判斷跨部門權限
       ' strExc(1) = SQLGrpStr(GetSystemKindByNick, 1) '專利
       ' strExc(2) = SQLGrpStr(GetSystemKindByNick, 2) '商標
       ' strExc(3) = SQLGrpStr(GetSystemKindByNick, 5) '服務
        If PUB_CheckSKAddCross(strUserNum, Systemkind_g, True, "ALL", strExc(4), False) Then
        End If
        strExc(1) = SQLGrpStr(strExc(4), 1) '專利
        strExc(2) = SQLGrpStr(strExc(4), 2) '商標
        strExc(3) = SQLGrpStr(strExc(4), 5) '服務
        'end 2016/02/24
        strExc(1) = Replace(strExc(1), ",' '", "")
        strExc(2) = Replace(strExc(2), ",' '", "")
        strExc(3) = Replace(strExc(3), ",' '", "")
        strExc(0) = ""
            If Len(strExc(1)) > 0 Then
               'Modified by Lydia 2016/08/01 + 含改請
               'strExc(0) = "(cp01 in (" & strExc(1) & ") and instr('" & NewCasePtyList & "',CP10)>0 ) "
               'Modified by Lydia 2025/09/19 改模組
               'strExc(0) = "(cp01 in (" & strExc(1) & ") and (instr('" & NewCasePtyList & "',CP10)>0 or substr(CP10,1,1)='3')) "
               strExc(0) = "(cp01 in (" & strExc(1) & ") and " & PUB_GetForNewCaseSql("1") & ") "
            End If
            If Len(strExc(2)) > 0 Then
               If Len(strExc(0)) > 0 Then strExc(0) = strExc(0) & " or "
               'Modified by Lydia 2025/09/19 改模組
               'strExc(0) = strExc(0) & "(cp01 in (" & strExc(2) & ") and CP10='101') "
               strExc(0) = strExc(0) & "(cp01 in (" & strExc(2) & ") " & PUB_GetForNewCaseSql("2") & ") "
            End If
            If Len(strExc(3)) > 0 Then
               If Len(strExc(0)) > 0 Then strExc(0) = strExc(0) & " or "
               'Modified by Lydia 2025/09/19 改模組
               'strExc(0) = strExc(0) & "(cp01 in (" & strExc(3) & ") and instr('801,802,805,806',CP10)>0) "
               strExc(0) = strExc(0) & "(cp01 in (" & strExc(3) & ") " & PUB_GetForNewCaseSql("5") & ") "
            End If
            strSQL1 = strSQL1 & " and (" & strExc(0) & ") "
       pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label9(9), 10) & frm100105_1.txt1(37)
    End If
    'end 2015/02/12
    
   'Add By Cheng 2002/11/22
   If frm100105_1.txt1(19).Text = "" And frm100105_1.txt1(20).Text = "" Then
      strSQL1 = strSQL1 + " AND CP09< 'B' "
   ElseIf frm100105_1.txt1(19).Text = "Y" And frm100105_1.txt1(20).Text = "" Then
      strSQL1 = strSQL1 + " AND CP09< 'C' "
   ElseIf frm100105_1.txt1(19).Text = "" And frm100105_1.txt1(20).Text = "Y" Then
      strSQL1 = strSQL1 + " AND ( CP09< 'B' OR CP09 >= 'C' ) "
   End If
   If Len(Trim(frm100105_1.txt1(19))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label18, 10) & frm100105_1.txt1(19) 'Add By Sindy 2010/01/22
   End If
   If Len(Trim(frm100105_1.txt1(20))) <> 0 Then
       pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label19, 8) & frm100105_1.txt1(20) 'Add By Sindy 2010/01/22
   End If
   
   'Add By Cheng 2003/07/02
   '承辦人
   If frm100105_1.txt1(18).Text <> "" Then
      strSQL1 = strSQL1 + " AND CP14='" & frm100105_1.txt1(18).Text & "' "
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label17 & frm100105_1.txt1(18) & frm100105_1.lbl1(3) 'Add By Sindy 2010/01/22
   End If
   
   'Add by Morgan 2011/1/17
   '只統計電子送件資料
   If frm100105_1.txt1(33).Text = "Y" Then
      'Modified by Lydia 2018/09/13 電子送件含自動扣款=A
      'strSQL1 = strSQL1 + " AND CP118='Y' "
      strSQL1 = strSQL1 + " AND CP118 is not null "
      pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label9(6), 10) & frm100105_1.txt1(33)
   End If
   
   strSQL1 = strSQL1 & " AND CP01||CP02<>'TT999999' " 'Added by Lydia 2022/09/27 排除TT-999999案號 'Memo by Lydia 2024/12/02 若統計類別=5的條件有異動，請一併變更frm090642-收文量、發文量
   StrSQL4 = strSQL1
   strSQL44 = strSQL1
   
   '2005/10/18 ADD BY SONIA
   If Len(Trim(frm100105_1.txt1(24))) <> 0 Then
      strSQL1 = strSQL1 + " AND F1.FA10>='" & frm100105_1.txt1(24) & "' "
   End If
   If Len(Trim(frm100105_1.txt1(25))) <> 0 Then
      strSQL1 = strSQL1 + " AND F1.FA10<='" & frm100105_1.txt1(25) & "z' "
   End If
   If Len(Trim(frm100105_1.txt1(24))) <> 0 Or Len(Trim(frm100105_1.txt1(25))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label5(2) & frm100105_1.txt1(24) & "-" & frm100105_1.txt1(25) 'Add By Sindy 2010/01/22
   End If
   '2005/10/18 END
   
   'Add by Morgan 2007/1/16
   If Len(Trim(frm100105_1.txt1(28))) <> 0 Then
      strSQL1 = strSQL1 + " AND F1.FA76='" & frm100105_1.txt1(28) & "' "
      If frm100105_1.txt1(28) = "A" Then
         pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label23, 8) & "律師事務所" 'Add By Sindy 2010/01/22
      ElseIf frm100105_1.txt1(28) = "B" Then
         pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label23, 8) & "公司直接委辦" 'Add By Sindy 2010/01/22
      ElseIf frm100105_1.txt1(28) = "C" Then
         pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label23, 8) & "其他" 'Add By Sindy 2010/01/22
      End If
   End If
   'end 2007/1/16
   
   '2005/10/21 ADD BY SONIA
   If Len(Trim(frm100105_1.txt1(26))) <> 0 Then
      strSQL1 = strSQL1 + " AND F2.FA10>='" & frm100105_1.txt1(26) & "' "
   End If
   If Len(Trim(frm100105_1.txt1(27))) <> 0 Then
      strSQL1 = strSQL1 + " AND F2.FA10<='" & frm100105_1.txt1(27) & "z' "
   End If
   If Len(Trim(frm100105_1.txt1(26))) <> 0 Or Len(Trim(frm100105_1.txt1(27))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label5(3) & frm100105_1.txt1(26) & "-" & frm100105_1.txt1(27) 'Add By Sindy 2010/01/22
   End If
   '2005/10/21 END
   
   'add by nickc 2007/01/12
   If Len(Trim(frm100105_1.txt1(3))) = 0 Then
      frm100105_1.txt1(3) = "ALL"
   End If
   
   strSQL2 = strSQL2 & strSQL1 'Modify by Amy 2017/02/24 智權人員加MCTF
   StrSQL3 = StrSQL3 & strSQL1 'Modify by Amy 2017/06/06 智權人員加MCTF
   strSQL5 = strSQL5 & strSQL1 'Modify by Amy 2017/02/24 智權人員加MCTF
   strSQL11 = strSQL1
   strSQL22 = strSQL22 & strSQL1 'Modify by Amy 2017/02/24 智權人員加MCTF
   strSQL33 = strSQL33 & strSQL1 'Modify by Amy 2017/06/06 智權人員加MCTF
   strSQL55 = strSQL55 & strSQL1 'Modify by Amy 2017/02/24 智權人員加MCTF
   
   If Len(Trim(frm100105_1.txt1(3))) <> 0 Then
   'Modified by Lydia 2016/02/25 "ALL"=使用者所有部門查詢權限
'      strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.Txt1(3).Text <> "ALL", frm100105_1.Txt1(3).Text, GetAllSysKind(frm100105_1.Txt1(3))), 1) & ") "
'      strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.Txt1(3).Text <> "ALL", frm100105_1.Txt1(3).Text, GetAllSysKind(frm100105_1.Txt1(3))), 2) & ") "
'      StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.Txt1(3).Text <> "ALL", frm100105_1.Txt1(3).Text, GetAllSysKind(frm100105_1.Txt1(3))), 3) & ") "
'      StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.Txt1(3).Text <> "ALL", frm100105_1.Txt1(3).Text, GetAllSysKind(frm100105_1.Txt1(3))), 4) & ") "
'      strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.Txt1(3).Text <> "ALL", frm100105_1.Txt1(3).Text, GetAllSysKind(frm100105_1.Txt1(3))), 5) & ") "
'      strSQL11 = strSQL11 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.Txt1(3).Text <> "ALL", frm100105_1.Txt1(3).Text, GetAllSysKind(frm100105_1.Txt1(3))), 1) & ") "
'      strSQL22 = strSQL22 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.Txt1(3).Text <> "ALL", frm100105_1.Txt1(3).Text, GetAllSysKind(frm100105_1.Txt1(3))), 2) & ") "
'      strSQL33 = strSQL33 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.Txt1(3).Text <> "ALL", frm100105_1.Txt1(3).Text, GetAllSysKind(frm100105_1.Txt1(3))), 3) & ") "
'      strSQL44 = strSQL44 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.Txt1(3).Text <> "ALL", frm100105_1.Txt1(3).Text, GetAllSysKind(frm100105_1.Txt1(3))), 4) & ") "
'      strSQL55 = strSQL55 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.Txt1(3).Text <> "ALL", frm100105_1.Txt1(3).Text, GetAllSysKind(frm100105_1.Txt1(3))), 5) & ") "
      If PUB_CheckSKAddCross(strUserNum, Systemkind_g, True, "ALL", strExc(4), False) Then
      End If
      'Memo by Lydia 2024/12/02 若統計類別=5的條件有異動，請一併變更frm090642-收文量、發文量
      strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.txt1(3).Text <> "ALL", frm100105_1.txt1(3).Text, strExc(4)), 1) & ") "
      strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.txt1(3).Text <> "ALL", frm100105_1.txt1(3).Text, strExc(4)), 2) & ") "
      StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.txt1(3).Text <> "ALL", frm100105_1.txt1(3).Text, strExc(4)), 3) & ") "
      StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.txt1(3).Text <> "ALL", frm100105_1.txt1(3).Text, strExc(4)), 4) & ") "
      strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.txt1(3).Text <> "ALL", frm100105_1.txt1(3).Text, strExc(4)), 5) & ") "
      strSQL11 = strSQL11 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.txt1(3).Text <> "ALL", frm100105_1.txt1(3).Text, strExc(4)), 1) & ") "
      strSQL22 = strSQL22 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.txt1(3).Text <> "ALL", frm100105_1.txt1(3).Text, strExc(4)), 2) & ") "
      strSQL33 = strSQL33 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.txt1(3).Text <> "ALL", frm100105_1.txt1(3).Text, strExc(4)), 3) & ") "
      strSQL44 = strSQL44 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.txt1(3).Text <> "ALL", frm100105_1.txt1(3).Text, strExc(4)), 4) & ") "
      strSQL55 = strSQL55 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100105_1.txt1(3).Text <> "ALL", frm100105_1.txt1(3).Text, strExc(4)), 5) & ") "
      'end 2016/02/25
      pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label3, 5) & frm100105_1.txt1(3) 'Add By Sindy 2010/01/22
   End If
   If Len(Trim(frm100105_1.txt1(9))) <> 0 Then
      strSQL1 = strSQL1 + " AND pa09>='" & frm100105_1.txt1(9) & "' "
      strSQL2 = strSQL2 + " AND tm10>='" & frm100105_1.txt1(9) & "' "
      StrSQL3 = StrSQL3 + " AND lc15>='" & frm100105_1.txt1(9) & "' "
      StrSQL4 = StrSQL4 + " AND '000'>='" & frm100105_1.txt1(9) & "' " 'Add By Sindy 2011/2/8
      strSQL5 = strSQL5 + " AND sp09>='" & frm100105_1.txt1(9) & "' "
      strSQL11 = strSQL11 + " AND pa09>='" & frm100105_1.txt1(9) & "' "
      strSQL22 = strSQL22 + " AND tm10>='" & frm100105_1.txt1(9) & "' "
      strSQL33 = strSQL33 + " AND lc15>='" & frm100105_1.txt1(9) & "' "
      strSQL44 = strSQL44 + " AND '000'>='" & frm100105_1.txt1(9) & "' " 'Add By Sindy 2011/2/8
      strSQL55 = strSQL55 + " AND sp09>='" & frm100105_1.txt1(9) & "' "
   End If
   If Len(Trim(frm100105_1.txt1(10))) <> 0 Then
       strSQL1 = strSQL1 + " AND pa09<='" & frm100105_1.txt1(10) & "' "
       strSQL2 = strSQL2 + " AND tm10<='" & frm100105_1.txt1(10) & "' "
       StrSQL3 = StrSQL3 + " AND lc15<='" & frm100105_1.txt1(10) & "' "
       StrSQL4 = StrSQL4 + " AND '000'<='" & frm100105_1.txt1(10) & "' " 'Add By Sindy 2011/2/8
       strSQL5 = strSQL5 + " AND sp09<='" & frm100105_1.txt1(10) & "' "
       strSQL11 = strSQL11 + " AND pa09<='" & frm100105_1.txt1(10) & "' "
       strSQL22 = strSQL22 + " AND tm10<='" & frm100105_1.txt1(10) & "' "
       strSQL33 = strSQL33 + " AND lc15<='" & frm100105_1.txt1(10) & "' "
       strSQL44 = strSQL44 + " AND '000'<='" & frm100105_1.txt1(10) & "' " 'Add By Sindy 2011/2/8
       strSQL55 = strSQL55 + " AND sp09<='" & frm100105_1.txt1(10) & "' "
   End If
   If Len(Trim(frm100105_1.txt1(9))) <> 0 Or Len(Trim(frm100105_1.txt1(10))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label5(0) & frm100105_1.txt1(9) & "-" & frm100105_1.txt1(10) 'Add By Sindy 2010/01/22
   End If
   'Add By Cheng 2004/03/08
   '專利/商標種類
   If Len(Trim(frm100105_1.txt1(21))) <> 0 Then
      strSQL1 = strSQL1 + " AND pa08>='" & frm100105_1.txt1(21) & "' "
      strSQL2 = strSQL2 + " AND tm08>='" & frm100105_1.txt1(21) & "' "
      strSQL11 = strSQL11 + " AND pa08>='" & frm100105_1.txt1(21) & "' "
      strSQL22 = strSQL22 + " AND tm08>='" & frm100105_1.txt1(21) & "' "
   End If
   If Len(Trim(frm100105_1.txt1(22))) <> 0 Then
      strSQL1 = strSQL1 + " AND pa08<='" & frm100105_1.txt1(22) & "' "
      strSQL2 = strSQL2 + " AND tm08<='" & frm100105_1.txt1(22) & "' "
      strSQL11 = strSQL11 + " AND pa08<='" & frm100105_1.txt1(22) & "' "
      strSQL22 = strSQL22 + " AND tm08<='" & frm100105_1.txt1(22) & "' "
   End If
   If Len(Trim(frm100105_1.txt1(21))) <> 0 Or Len(Trim(frm100105_1.txt1(22))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label20 & frm100105_1.txt1(21) & "-" & frm100105_1.txt1(22) 'Add By Sindy 2010/01/22
   End If
   '2004/03/08 End
   
   '2009/4/2 ADD BY SONIA 加准駁條件
   If Len(Trim(frm100105_1.txt1(32))) <> 0 Then
      If frm100105_1.txt1(32) = "1" Then
         pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label27, 9) & "准" 'Add By Sindy 2010/01/22
      ElseIf frm100105_1.txt1(32) = "2" Then
         pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label27, 9) & "駁" 'Add By Sindy 2010/01/22
      ElseIf frm100105_1.txt1(32) = "3" Then
         pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label27, 9) & "未准(含駁)" 'Add By Sindy 2010/01/22
      ElseIf frm100105_1.txt1(32) = "4" Then
         pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label27, 9) & "無准駁" 'Add By Sindy 2010/01/22
      End If
      Select Case Trim(frm100105_1.txt1(32))
         Case "1", "2"
            strSQL1 = strSQL1 + " AND pa16='" & frm100105_1.txt1(32) & "' "
            strSQL2 = strSQL2 + " AND tm16='" & frm100105_1.txt1(32) & "' "
            strSQL11 = strSQL11 + " AND pa16='" & frm100105_1.txt1(32) & "' "
            strSQL22 = strSQL22 + " AND tm16='" & frm100105_1.txt1(32) & "' "
         Case "3"
            strSQL1 = strSQL1 + " AND (pa16 is null or pa16<>'1') "
            strSQL2 = strSQL2 + " AND (tm16 is null or tm16<>'1') "
            strSQL11 = strSQL11 + " AND (pa16 is null or pa16<>'1') "
            strSQL22 = strSQL22 + " AND (tm16 is null or tm16<>'1') "
         Case "4"
            strSQL1 = strSQL1 + " AND pa16 is null "
            strSQL2 = strSQL2 + " AND tm16 is null "
            strSQL11 = strSQL11 + " AND pa16 is null "
            strSQL22 = strSQL22 + " AND tm16 is null "
      End Select
   End If
   '2009/4/2 END
   
   'edit by nickc 2007/04/24 改成區間
   If Len(Trim(frm100105_1.txt1(15))) <> 0 And Len(Trim(frm100105_1.txt1(29))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label7 & frm100105_1.txt1(15) & "-" & frm100105_1.txt1(29) 'Add By Sindy 2010/01/22
      strSQL1 = strSQL1 + " AND pa75>='" & GetNewFagent(frm100105_1.txt1(15)) & "' and pa75<='" & GetNewFagent(frm100105_1.txt1(29)) & "' "
      strSQL2 = strSQL2 + " AND tm44>='" & GetNewFagent(frm100105_1.txt1(15)) & "' and tm44<='" & GetNewFagent(frm100105_1.txt1(29)) & "' "
      StrSQL3 = StrSQL3 + " AND lc22>='" & GetNewFagent(frm100105_1.txt1(15)) & "' and lc22<='" & GetNewFagent(frm100105_1.txt1(29)) & "' "
      strSQL5 = strSQL5 + " AND sp26>='" & GetNewFagent(frm100105_1.txt1(15)) & "' and sp26<='" & GetNewFagent(frm100105_1.txt1(29)) & "' "
      strSQL11 = strSQL11 + " AND cp44>='" & GetNewFagent(frm100105_1.txt1(15)) & "' and cp44<='" & GetNewFagent(frm100105_1.txt1(29)) & "' "
      strSQL22 = strSQL22 + " AND cp44>='" & GetNewFagent(frm100105_1.txt1(15)) & "' and cp44<='" & GetNewFagent(frm100105_1.txt1(29)) & "' "
      strSQL33 = strSQL33 + " AND cp44>='" & GetNewFagent(frm100105_1.txt1(15)) & "' and cp44<='" & GetNewFagent(frm100105_1.txt1(29)) & "' "
      strSQL44 = strSQL44 + " AND cp44>='" & GetNewFagent(frm100105_1.txt1(15)) & "' and cp44<='" & GetNewFagent(frm100105_1.txt1(29)) & "' "
      strSQL55 = strSQL55 + " AND cp44>='" & GetNewFagent(frm100105_1.txt1(15)) & "' and cp44<='" & GetNewFagent(frm100105_1.txt1(29)) & "' "
   End If
   
   'edit by nickc 2007/04/24 改成區間
   If Len(Trim(frm100105_1.txt1(16))) <> 0 And Len(Trim(frm100105_1.txt1(30))) <> 0 Then
      'Modify By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label12 & frm100105_1.txt1(16) & "-" & frm100105_1.txt1(30) 'Add By Sindy 2010/01/22
      strSQL1 = strSQL1 + " AND ((pa26>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and pa26<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (pa27>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and pa27<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (pa28>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and pa28<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (pa29>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and pa29<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (pa30>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and pa30<='" & GetNewFagent(frm100105_1.txt1(30)) & "')) "
      strSQL2 = strSQL2 + " AND ((tm23>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and tm23<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (tm78>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and tm78<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (tm79>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and tm79<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (tm80>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and tm80<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (tm81>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and tm81<='" & GetNewFagent(frm100105_1.txt1(30)) & "')) "
      StrSQL3 = StrSQL3 + " AND ((lc11>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and lc11<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (lc43>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and lc43<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (lc44>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and lc44<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (lc45>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and lc45<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (lc46>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and lc46<='" & GetNewFagent(frm100105_1.txt1(30)) & "')) "
      StrSQL4 = StrSQL4 + " AND ((hc05>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and hc05<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (hc24>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and hc24<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (hc25>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and hc25<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (hc26>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and hc26<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (hc27>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and hc27<='" & GetNewFagent(frm100105_1.txt1(30)) & "')) "
      strSQL5 = strSQL5 + " AND ((sp08>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and sp08<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (sp58>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and sp58<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (sp59>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and sp59<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (sp65>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and sp65<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (sp66>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and sp66<='" & GetNewFagent(frm100105_1.txt1(30)) & "')) "
      strSQL11 = strSQL11 + " AND ((pa26>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and pa26<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (pa27>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and pa27<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (pa28>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and pa28<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (pa29>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and pa29<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (pa30>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and pa30<='" & GetNewFagent(frm100105_1.txt1(30)) & "')) "
      strSQL22 = strSQL22 + " AND ((tm23>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and tm23<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (tm78>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and tm78<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (tm79>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and tm79<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (tm80>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and tm80<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (tm81>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and tm81<='" & GetNewFagent(frm100105_1.txt1(30)) & "')) "
      strSQL33 = strSQL33 + " AND ((lc11>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and lc11<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (lc43>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and lc43<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (lc44>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and lc44<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (lc45>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and lc45<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (lc46>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and lc46<='" & GetNewFagent(frm100105_1.txt1(30)) & "')) "
      strSQL44 = strSQL44 + " AND ((hc05>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and hc05<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (hc24>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and hc24<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (hc25>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and hc25<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (hc26>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and hc26<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (hc27>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and hc27<='" & GetNewFagent(frm100105_1.txt1(30)) & "')) "
      strSQL55 = strSQL55 + " AND ((sp08>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and sp08<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (sp58>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and sp58<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (sp59>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and sp59<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (sp65>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and sp65<='" & GetNewFagent(frm100105_1.txt1(30)) & "') or (sp66>='" & GetNewFagent(frm100105_1.txt1(16)) & "' and sp66<='" & GetNewFagent(frm100105_1.txt1(30)) & "')) "
   End If
   
   'add by nick 2005/02/04
   If Trim(frm100105_1.txt1(23).Text) = "Y" Then
      strSQL1 = strSQL1 & " and pa46='Y' and pa09<>'056' "
      strSQL11 = strSQL11 & " and pa46='Y' and pa09<>'056' "
      pub_QL05 = pub_QL05 & ";" & Left(frm100105_1.Label9(4), 10) & frm100105_1.txt1(23) 'Add By Sindy 2010/01/22
   End If
   
   'add by toni 2008/10/15 FCP工程師組別
   If Len(Trim(frm100105_1.txt1(31))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100105_1.Label26 & frm100105_1.txt1(31) & frm100105_1.lblName 'Add By Sindy 2010/01/22
      strSQL1 = strSQL1 + " AND pa150 ='" & frm100105_1.txt1(31) & "' "
      strSQL2 = strSQL2 + "and TM01='" & frm100105_1.txt1(31) & "'"
      StrSQL3 = StrSQL3 + "and LC01='" & frm100105_1.txt1(31) & "'"
      StrSQL4 = StrSQL4 + "and HC01='" & frm100105_1.txt1(31) & "'"
      'Modify by Morgan 2009/9/14 FG也要分組別
      'strSQL5 = strSQL5 + "and SP01='" & frm100105_1.txt1(31) & "'"
      strSQL5 = strSQL5 + "and SP79='" & frm100105_1.txt1(31) & "'"
      strSQL11 = strSQL11 + " AND pa150 ='" & frm100105_1.txt1(31) & "' "
      strSQL22 = strSQL22 + "and TM01='" & frm100105_1.txt1(31) & "'"
      strSQL33 = strSQL33 + "and LC01='" & frm100105_1.txt1(31) & "'"
      strSQL44 = strSQL44 + "and HC01='" & frm100105_1.txt1(31) & "'"
      'Modify by Morgan 2009/9/14 FG也要分組別
      'StrSQL55 = StrSQL55 + "and SP01='" & frm100105_1.txt1(31) & "'"
      strSQL55 = strSQL55 + "and SP79='" & frm100105_1.txt1(31) & "'"
   End If
   'end 2008/10/15
   
   'Add By Sindy 2012/3/9 +國際分類
   If Len(Trim(frm100105_1.txt1(34))) <> 0 Then
       strSQL1 = strSQL1 + " AND pa160='" & frm100105_1.txt1(34) & "' "
       strSQL11 = strSQL11 + " AND pa160='" & frm100105_1.txt1(34) & "' "
       pub_QL05 = pub_QL05 & ";" & frm100105_1.Label28 & frm100105_1.txt1(34)
   End If
   
   'Add By Sindy 2014/7/9 +專利案件屬性
   If Trim(frm100105_1.txt1(17)) = "A" Then
      '專利案件屬性,其他主檔不需讀取資料
      strSQL2 = strSQL2 + "and TM01='P'"
      StrSQL3 = StrSQL3 + "and LC01='P'"
      StrSQL4 = StrSQL4 + "and HC01='P'"
      strSQL5 = strSQL5 + "and SP79='P'"
      strSQL22 = strSQL22 + "and TM01='P'"
      strSQL33 = strSQL33 + "and LC01='P'"
      strSQL44 = strSQL44 + "and HC01='P'"
      strSQL55 = strSQL55 + "and SP79='P'"
   End If
   '2014/7/9 END
   
   'Add By Sindy 2010/01/22
   If Len(Trim(frm100105_1.txt1(17))) <> 0 Then
      If Trim(frm100105_1.txt1(17)) = "1" Then
         pub_QL05 = pub_QL05 & ";" & frm100105_1.Label11 & "業務區"
      ElseIf Trim(frm100105_1.txt1(17)) = "2" Then
         pub_QL05 = pub_QL05 & ";" & frm100105_1.Label11 & "智權人員"
      ElseIf Trim(frm100105_1.txt1(17)) = "3" Then
         pub_QL05 = pub_QL05 & ";" & frm100105_1.Label11 & "申請國家"
      ElseIf Trim(frm100105_1.txt1(17)) = "4" Then
         pub_QL05 = pub_QL05 & ";" & frm100105_1.Label11 & "申請人國籍"
      ElseIf Trim(frm100105_1.txt1(17)) = "5" Then
         pub_QL05 = pub_QL05 & ";" & frm100105_1.Label11 & "案件性質"
      ElseIf Trim(frm100105_1.txt1(17)) = "6" Then
         pub_QL05 = pub_QL05 & ";" & frm100105_1.Label11 & "FC代理人"
      ElseIf Trim(frm100105_1.txt1(17)) = "7" Then
         pub_QL05 = pub_QL05 & ";" & frm100105_1.Label11 & "CF代理人"
      ElseIf Trim(frm100105_1.txt1(17)) = "8" Then
         pub_QL05 = pub_QL05 & ";" & frm100105_1.Label11 & "FC代理人國籍"
      ElseIf Trim(frm100105_1.txt1(17)) = "9" Then
         pub_QL05 = pub_QL05 & ";" & frm100105_1.Label11 & "FCP工程師組別"
      'Add By Sindy 2014/7/9
      ElseIf Trim(frm100105_1.txt1(17)) = "A" Then
         pub_QL05 = pub_QL05 & ";" & frm100105_1.Label11 & "專利案件屬性"
      '2014/7/9 END
      'Added by Lydia 2017/02/03
      ElseIf Trim(frm100105_1.txt1(17)) = "B" Then
         pub_QL05 = pub_QL05 & ";" & frm100105_1.Label11 & "洲別"
      'end 2017/02/03
      'Added by Lydia 2025/08/06
      ElseIf Trim(frm100105_1.txt1(17)) = "C" Then
         pub_QL05 = pub_QL05 & ";" & frm100105_1.Label11 & "承辦人"
      End If
      'end 2025/08/06
   End If
   
   '2009/11/20 ADD BY SONIA 寫入工作檔時先判斷是否中間新案
   'Memo by Lydia 2024/03/05 每日批次frmAutoBatchDay.StrMenu131參考FCT,CFT,S案的統計量，如有修改請一併檢查
   strCP31 = "decode(CP31,'Y',decode(instr('CFT,FCT,T,TF',CP01),0,decode(instr('CFP,FCP,P',CP01),0,'',decode(instr('" & NewCasePtyList & ",801,803',CP10),0,'Y','')),decode(instr('101,308,601,603,605,618',CP10),0,'Y','')),'') "
   '2009/11/20 END
   
   'Modify By Cheng 2002/05/09將ALL取消, 但為了避免選出的資料由於重覆而合併在一起, 故多加總收文號
   '911024 nick 指定使用 index 及 join  2009/11/18 MODIFY BY SONIA發現不知何時取消沒用
   '2005/10/18 MODIFY BY SONIA 加入FC代理人國籍條件
   '2005/10/21 MODIFY BY SONIA 再加CF代理人國籍條件
   'Modify by Sindy 2009/11/06 +CP31
   '2009/11/18 MODIFY BY SONIA 取消IndexString
   '2009/11/20 MODIFY BY SONIA 寫入工作檔時先判斷是否中間新案
   'Modify By Sindy 2014/7/10 專利的CP12==>" & IIf(frm100105_1.txt1(17) = "A", "pa08||pa158", "cp12") & "
   'Memo by Lydia 2024/12/02 若統計類別=5的條件有異動，請一併變更frm090642-收文量、發文量
   strSql = "SELECT " & SQLDate("CP05") & "," & SQLDate("CP27") & ",CP01,CP26,CP21," & IIf(frm100105_1.txt1(17) = "A", "pa08||pa158", "cp12") & ",CP13,pa09,cu10,CP10," & IIf(frm100105_1.txt1(17) = "7", "cp44", "pa75") & ",pa26,'',CP02||CP03||CP04,'" & strUserNum & "',CP09,PA150," & strCP31 & " FROM CASEPROGRESS,patent,customer,FAGENT F1,FAGENT F2 where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND SUBSTR(PA75,1,8)=F1.FA01(+) AND SUBSTR(PA75,9,1)=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND SUBSTR(CP44,9,1)=F2.FA02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL1
   'Memo by Lydia 2024/03/05 每日批次frmAutoBatchDay.StrMenu131參考FCT,CFT案的統計量，如有修改請一併檢查
   strSql = strSql & " union  select " & SQLDate("CP05") & "," & SQLDate("CP27") & ",CP01,CP26,CP21,CP12,CP13,tm10,cu10,CP10," & IIf(frm100105_1.txt1(17) = "7", "cp44", "tm44") & ",tm23,'',CP02||CP03||CP04,'" & strUserNum & "',CP09,''," & strCP31 & " FROM CASEPROGRESS,trademark,customer,FAGENT F1,FAGENT F2 where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) AND SUBSTR(TM44,1,8)=F1.FA01(+) AND SUBSTR(TM44,9,1)=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND SUBSTR(CP44,9,1)=F2.FA02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
   strSql = strSql & " union  select " & SQLDate("CP05") & "," & SQLDate("CP27") & ",CP01,CP26,CP21,CP12,CP13,lc15,cu10,CP10," & IIf(frm100105_1.txt1(17) = "7", "cp44", "lc22") & ",lc11,'',CP02||CP03||CP04,'" & strUserNum & "',CP09,''," & strCP31 & " FROM CASEPROGRESS,lawcase,customer,FAGENT F1,FAGENT F2 where cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) AND SUBSTR(LC22,1,8)=F1.FA01(+) AND SUBSTR(LC22,9,1)=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND SUBSTR(CP44,9,1)=F2.FA02(+) and " & SQLNewFag("lc11", "cu") & " " & StrSQL3
   'Modify By Sindy 2011/2/8 若有輸入FC或CF代理人相關查詢條件, 則此句SQL不使用, 因顧問無FC代理人
   If (Len(Trim(frm100105_1.txt1(15))) = 0 And Len(Trim(frm100105_1.txt1(29))) = 0) And _
      Len(Trim(frm100105_1.txt1(24))) = 0 And Len(Trim(frm100105_1.txt1(25))) = 0 And _
      Len(Trim(frm100105_1.txt1(28))) = 0 And Len(Trim(frm100105_1.txt1(26))) = 0 And _
      Len(Trim(frm100105_1.txt1(27))) = 0 Then
   '2011/2/8 End
      strSql = strSql & " union  select " & SQLDate("CP05") & "," & SQLDate("CP27") & ",CP01,CP26,CP21,CP12,CP13,''  ,cu10,CP10," & IIf(frm100105_1.txt1(17) = "7", "cp44", "''") & ",hc05,'',CP02||CP03||CP04,'" & strUserNum & "',CP09,''," & strCP31 & " FROM CASEPROGRESS,hirecase,customer where cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and " & SQLNewFag("hc05", "cu") & " " & StrSQL4
   End If
   'Modify by Morgan 2009/9/14 +SP79
   'Memo by Lydia 2024/03/05 每日批次frmAutoBatchDay.StrMenu131參考S案的統計量，如有修改請一併檢查
   strSql = strSql & " union  select " & SQLDate("CP05") & "," & SQLDate("CP27") & ",CP01,CP26,CP21,CP12,CP13,sp09,cu10,CP10," & IIf(frm100105_1.txt1(17) = "7", "cp44", "sp26") & ",sp08,'',CP02||CP03||CP04,'" & strUserNum & "',CP09,SP79," & strCP31 & " FROM CASEPROGRESS,servicepractice,customer,FAGENT F1,FAGENT F2 where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND SUBSTR(SP26,9,1)=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND SUBSTR(CP44,9,1)=F2.FA02(+) and " & SQLNewFag("sp08", "cu") & " " & strSQL5
   strSql = strSql & " union  select " & SQLDate("CP05") & "," & SQLDate("CP27") & ",CP01,CP26,CP21," & IIf(frm100105_1.txt1(17) = "A", "pa08||pa158", "cp12") & ",CP13,pa09,cu10,CP10," & IIf(frm100105_1.txt1(17) = "7", "cp44", "pa75") & ",pa26,'',CP02||CP03||CP04,'" & strUserNum & "',CP09,PA150," & strCP31 & " FROM CASEPROGRESS,patent,customer,FAGENT F1,FAGENT F2 where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND SUBSTR(PA75,1,8)=F1.FA01(+) AND SUBSTR(PA75,9,1)=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND SUBSTR(CP44,9,1)=F2.FA02(+) and " & SQLNewFag("pa26", "cu") & " " & strSQL11
   strSql = strSql & " union  select " & SQLDate("CP05") & "," & SQLDate("CP27") & ",CP01,CP26,CP21,CP12,CP13,tm10,cu10,CP10," & IIf(frm100105_1.txt1(17) = "7", "cp44", "tm44") & ",tm23,'',CP02||CP03||CP04,'" & strUserNum & "',CP09,''," & strCP31 & " FROM CASEPROGRESS,trademark,customer,FAGENT F1,FAGENT F2 where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) AND SUBSTR(TM44,1,8)=F1.FA01(+) AND SUBSTR(TM44,9,1)=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND SUBSTR(CP44,9,1)=F2.FA02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL22
   strSql = strSql & " union  select " & SQLDate("CP05") & "," & SQLDate("CP27") & ",CP01,CP26,CP21,CP12,CP13,lc15,cu10,CP10," & IIf(frm100105_1.txt1(17) = "7", "cp44", "lc22") & ",lc11,'',CP02||CP03||CP04,'" & strUserNum & "',CP09,''," & strCP31 & " FROM CASEPROGRESS,lawcase,customer,FAGENT F1,FAGENT F2 where cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) AND SUBSTR(LC22,1,8)=F1.FA01(+) AND SUBSTR(LC22,9,1)=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND SUBSTR(CP44,9,1)=F2.FA02(+) and " & SQLNewFag("lc11", "cu") & " " & strSQL33
   'Modify By Sindy 2011/2/8 若有輸入FC或CF代理人相關查詢條件, 則此句SQL不使用, 因顧問無FC代理人
   If Len(Trim(frm100105_1.txt1(24))) = 0 And Len(Trim(frm100105_1.txt1(25))) = 0 And _
      Len(Trim(frm100105_1.txt1(28))) = 0 And Len(Trim(frm100105_1.txt1(26))) = 0 And _
      Len(Trim(frm100105_1.txt1(27))) = 0 Then
   '2011/2/8 End
      strSql = strSql & " union  select " & SQLDate("CP05") & "," & SQLDate("CP27") & ",CP01,CP26,CP21,CP12,CP13,''  ,cu10,CP10," & IIf(frm100105_1.txt1(17) = "7", "cp44", "''") & ",hc05,'',CP02||CP03||CP04,'" & strUserNum & "',CP09,''," & strCP31 & " FROM CASEPROGRESS,hirecase,customer where cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and " & SQLNewFag("hc05", "cu") & " " & strSQL44
   End If
   'Modify by Morgan 2009/9/14 +SP79
   strSql = strSql & " union  select " & SQLDate("CP05") & "," & SQLDate("CP27") & ",CP01,CP26,CP21,CP12,CP13,sp09,cu10,CP10," & IIf(frm100105_1.txt1(17) = "7", "cp44", "sp26") & ",sp08,'',CP02||CP03||CP04,'" & strUserNum & "',CP09,SP79," & strCP31 & " FROM CASEPROGRESS,servicepractice,customer,FAGENT F1,FAGENT F2 where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND SUBSTR(SP26,9,1)=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND SUBSTR(CP44,9,1)=F2.FA02(+) and " & SQLNewFag("sp08", "cu") & " " & strSQL55
   
   cnnConnection.Execute "insert into r100105 " & strSql    '寫數量work檔
   
   'Added by Lydia 2025/08/06 統計條件:承辦人(更新資料)
   If frm100105_1.txt1(17) = "C" Then
      strSql = "UPDATE r100105 SET r02006=(SELECT st03 FROM caseprogress,staff WHERE cp09=r02015 AND cp14=st01(+)) WHERE ID ='" & strUserNum & "' "
      cnnConnection.Execute strSql
      strSql = "UPDATE r100105 SET r02007=(SELECT st01 FROM caseprogress,staff WHERE cp09=r02015 AND cp14=st01(+)) WHERE ID ='" & strUserNum & "' "
      cnnConnection.Execute strSql
   End If
   'end 2025/08/06
   
   strSql = "select * from r100105 where id ='" & strUserNum & "' And RowNum <= 1 "
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/01/22
      '抓商品類別寫商申類別數量work檔
      'Modify by toni 加欄位R02016  2008/10/15
      'Modify By Sindy 2009/11/06 加欄位R02017
      'Modify By Sindy 2011/1/31 增加「查名」案件之類別數統計
      'modify by sonia 2013/7/16 取消僅台灣案才統計類別數的限制And R02008='000'(陳經理要統計CFT)
      StrSQLa = "Select R02001, R02002, R02003, R02004, R02005, R02006, R02007, R02008, R02009, R02010, R02011, R02012, R02013, R02014, ID, R02015,R02016,R02017,TM09 From R100105, CaseProgress, Trademark " & _
                      " Where R02015=CP09 And CP01=TM01 And CP02=TM02 And CP03=TM03 And Cp04=TM04 And R02010 in ('101','001') And  id ='" & strUserNum & "' " & _
                      " union " & _
                      "Select R02001, R02002, R02003, R02004, R02005, R02006, R02007, R02008, R02009, R02010, R02011, R02012, R02013, R02014, ID, R02015,R02016,R02017,SP73 as TM09 From R100105, CaseProgress, ServicePractice " & _
                      " Where R02015=CP09 And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And R02010 in ('001') And  id ='" & strUserNum & "' "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         While Not rsA.EOF
            '若有商品類別
            If "" & rsA("TM09").Value <> "" Then
               dblTMKindCnt = UBound(Split("" & rsA("TM09").Value, ",")) + 1
               For ii = 1 To dblTMKindCnt
                   'Modify By Sindy 2009/11/06 +rsA.Fields(17).Value
                   cnnConnection.Execute "insert into r100105_1 values ('" & rsA.Fields(0).Value & "','" & rsA.Fields(1).Value & "','" & rsA.Fields(2).Value & "','" & rsA.Fields(3).Value & "','" & rsA.Fields(4).Value & "','" & rsA.Fields(5).Value & "','" & rsA.Fields(6).Value & "','" & rsA.Fields(7).Value & "','" & rsA.Fields(8).Value & "','" & rsA.Fields(9).Value & "','" & rsA.Fields(10).Value & "','" & rsA.Fields(11).Value & "','" & rsA.Fields(12).Value & "','" & rsA.Fields(13).Value & "','" & rsA.Fields(14).Value & "','" & rsA.Fields(15).Value & "','" & rsA.Fields(16).Value & "','" & rsA.Fields(17).Value & "') "
               Next ii
            '若無商品類別(要算一件)
            Else
               'Modify By Sindy 2009/11/06 +rsA.Fields(17).Value
               cnnConnection.Execute "insert into r100105_1 values ('" & rsA.Fields(0).Value & "','" & rsA.Fields(1).Value & "','" & rsA.Fields(2).Value & "','" & rsA.Fields(3).Value & "','" & rsA.Fields(4).Value & "','" & rsA.Fields(5).Value & "','" & rsA.Fields(6).Value & "','" & rsA.Fields(7).Value & "','" & rsA.Fields(8).Value & "','" & rsA.Fields(9).Value & "','" & rsA.Fields(10).Value & "','" & rsA.Fields(11).Value & "','" & rsA.Fields(12).Value & "','" & rsA.Fields(13).Value & "','" & rsA.Fields(14).Value & "','" & rsA.Fields(15).Value & "','" & rsA.Fields(16).Value & "','" & rsA.Fields(17).Value & "') "
            End If
            rsA.MoveNext
         Wend
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/01/22
      ShowNoData
      Screen.MousePointer = vbDefault
      DoTemp = False
      Exit Function
   End If
   CheckOC
   DoTemp = True
End Function

