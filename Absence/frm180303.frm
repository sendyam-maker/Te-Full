VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180303 
   BorderStyle     =   1  '單線固定
   Caption         =   "打卡資料查詢"
   ClientHeight    =   6320
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   8955
   Tag             =   "加班資料"
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   540
      Left            =   2910
      TabIndex        =   25
      Top             =   5790
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   26
         Top             =   210
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   225
         Index           =   1
         Left            =   75
         TabIndex        =   27
         Top             =   270
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   5790
      TabIndex        =   13
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton cmdB14 
      Caption         =   "異常處理結果"
      Height          =   360
      Left            =   5100
      TabIndex        =   10
      Top             =   420
      Width           =   1665
   End
   Begin VB.CommandButton cmdABS 
      Caption         =   "查詢當日請假資料"
      Height          =   360
      Left            =   6840
      TabIndex        =   11
      Top             =   420
      Width           =   2025
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "當日打卡明細"
      Height          =   360
      Left            =   7200
      TabIndex        =   12
      Top             =   810
      Width           =   1665
   End
   Begin VB.TextBox txtST06 
      Height          =   300
      Index           =   0
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1020
      Width           =   495
   End
   Begin VB.TextBox txtST06 
      Height          =   300
      Index           =   1
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1020
      Width           =   495
   End
   Begin VB.TextBox txtB1003 
      Height          =   300
      Index           =   1
      Left            =   2190
      MaxLength       =   6
      TabIndex        =   3
      Top             =   360
      Width           =   1005
   End
   Begin VB.TextBox txtDept 
      Height          =   300
      Index           =   1
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   5
      Top             =   690
      Width           =   495
   End
   Begin VB.TextBox txtDept 
      Height          =   300
      Index           =   0
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   4
      Top             =   690
      Width           =   495
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Index           =   1
      Left            =   2190
      MaxLength       =   7
      TabIndex        =   1
      Top             =   30
      Width           =   1005
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Index           =   0
      Left            =   1050
      MaxLength       =   7
      TabIndex        =   0
      Top             =   30
      Width           =   1005
   End
   Begin VB.TextBox txtB1003 
      Height          =   300
      Index           =   0
      Left            =   1050
      MaxLength       =   6
      TabIndex        =   2
      Top             =   360
      Width           =   1005
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7890
      TabIndex        =   9
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   360
      Left            =   6840
      TabIndex        =   8
      Top             =   30
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3765
      Left            =   60
      TabIndex        =   14
      Top             =   1410
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   6632
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|日期|部門|員工姓名|上班打卡|下班打卡|有無請假|上班異常|下班異常"
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
      _Band(0).Cols   =   9
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   705
      Left            =   3660
      TabIndex        =   20
      Top             =   60
      Visible         =   0   'False
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   1252
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin VB.ComboBox cboDept 
      Height          =   260
      Index           =   1
      Left            =   5700
      TabIndex        =   29
      Text            =   "cboDept"
      Top             =   720
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.ComboBox cboDept 
      Height          =   260
      Index           =   0
      Left            =   3690
      TabIndex        =   30
      Text            =   "cboDept"
      Top             =   720
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "註：99996.來賓"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   4110
      TabIndex        =   28
      Top             =   5220
      Width           =   1370
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "共   筆"
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   8100
      TabIndex        =   24
      Top             =   5220
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "注意：請勾選下面資料列，再進行各項查詢"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   0
      Left            =   5190
      TabIndex        =   23
      Top             =   1200
      Width           =   3420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1.Caption = 接收打卡時間公告"
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   60
      TabIndex        =   22
      Top             =   5220
      Width           =   2700
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "備註：您有權限查詢的部門別為"
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   60
      TabIndex        =   21
      Top             =   5430
      Width           =   8010
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "(1.北所 2.中所 3.南所 4.高所 5.其他)"
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   2250
      TabIndex        =   18
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Line Line4 
      X1              =   1469.821
      X2              =   1950.089
      Y1              =   1170.074
      Y2              =   1170.074
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "所　　別："
      Height          =   180
      Left            =   105
      TabIndex        =   19
      Top             =   1080
      Width           =   900
   End
   Begin VB.Line Line3 
      X1              =   1980.106
      X2              =   2220.24
      Y1              =   509.596
      Y2              =   509.596
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "部  門  別："
      Height          =   180
      Left            =   105
      TabIndex        =   17
      Top             =   750
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   1469.821
      X2              =   1950.089
      Y1              =   840.335
      Y2              =   840.335
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  '內實線
      X1              =   2010.122
      X2              =   2220.24
      Y1              =   149.881
      Y2              =   149.881
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "日　　期："
      Height          =   180
      Left            =   105
      TabIndex        =   16
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Left            =   105
      TabIndex        =   15
      Top             =   420
      Width           =   900
   End
End
Attribute VB_Name = "frm180303"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2023/12/19 修改抓新部門程式
'Memo By Sindy 2021/5/28 Form2.0已修改
'Created by Morgan 2013/6/10
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Public bolClose As Boolean
Dim i As Long, j As Integer
Public m_IsAbsBossST03 As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iLine As Integer
Public m_strEmp As String 'Add By Sindy 2021/12/21 所屬簽核的人員


Private Sub SetGrd2()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("V", "員工代號", "表單編號", "TableID", "SA02", "SA03")
   arrGridHeadWidth = Array(800, 800, 800, 800, 800, 800)
   'grd2.Visible = False
   grd2.Cols = UBound(arrGridHeadText) + 1
   grd2.Rows = 2
   For iRow = 0 To grd2.Cols - 1
      grd2.row = 0
      grd2.col = iRow
      grd2.Text = arrGridHeadText(iRow)
      grd2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd2.CellAlignment = flexAlignCenterCenter
   Next
   'grd2.Visible = True
End Sub

'查詢出缺勤明細資料
Public Sub PubShowNextData()
Dim i As Integer
Dim bolSelV As Boolean
   
   bolSelV = False
   Me.Enabled = False
   For i = 1 To grd2.Rows - 1
      grd2.col = 0
      grd2.row = i
      If Trim(grd2.Text) = "V" Then
         bolSelV = True
         grd2.Text = ""
         grd2.col = 2 '表單編號
         Screen.MousePointer = vbHourglass
         Me.Hide
         Call frm180301_03.SetParent(Me)
         If grd2.TextMatrix(i, 3) = "1" Then '出缺勤
            frm180301_03.txtB1001 = Pub_RplStr(grd2.Text)
            frm180301_03.QueryData
         Else
            frm180301_03.txtB1003 = Pub_RplStr(grd2.TextMatrix(i, 1))
            frm180301_03.m_SA02 = Pub_RplStr(grd2.TextMatrix(i, 4))
            frm180301_03.m_SA03 = Pub_RplStr(grd2.TextMatrix(i, 5))
            If grd2.TextMatrix(i, 3) = "2" Then '請假
               frm180301_03.QueryData_2
            ElseIf grd2.TextMatrix(i, 3) = "3" Then '加班
               frm180301_03.QueryData_3
            ElseIf grd2.TextMatrix(i, 3) = "4" Then '出差
               frm180301_03.QueryData_4
            End If
         End If
         frm180301_03.Show
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         Exit Sub
      End If
   Next i
   Me.Enabled = True
   If bolSelV = False Then
      Call cmdABS_Click
   End If
End Sub

Private Sub cmdABS_Click()
Dim rsTmp As New ADODB.Recordset
'Dim bolSelisV As Boolean
'
'   bolSelisV = False
   grd2.Clear
   SetGrd2
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If GRD1.TextMatrix(i, 0) = "V" Then
'         bolSelisV = True
         GRD1.Text = ""
         For j = 0 To GRD1.Cols - 1
            GRD1.col = j
            GRD1.CellBackColor = QBColor(15)
         Next j
         Call SetColColor(i)
'         If QueryData_ABS(GRD1.TextMatrix(i, 9), GRD1.TextMatrix(i, 1)) = True Then
'            Exit Sub
'         End If
         If PUB_QueryData_ABS(GRD1.TextMatrix(i, 9), GRD1.TextMatrix(i, 1), rsTmp) = True Then
            Set grd2.Recordset = rsTmp
            Call PubShowNextData
            Exit Sub
         End If
      End If
   Next i
'   If bolSelisV = False Then
'      MsgBox "請勾選欲查詢的資料！"
'   End If
End Sub

'異常處理明細
Public Sub cmdB14_Click()
'Dim bolSelisV As Boolean
'
'   bolSelisV = False
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If GRD1.TextMatrix(i, 0) = "V" Then
'         bolSelisV = True
         GRD1.Text = ""
         For j = 0 To GRD1.Cols - 1
            GRD1.col = j
            GRD1.CellBackColor = QBColor(15)
         Next j
         Call SetColColor(i)
         If GRD1.TextMatrix(i, 7) <> "" Or GRD1.TextMatrix(i, 8) <> "" Then
            Me.Hide
            Call frm180303_2.SetParent(Me)
            frm180303_2.m_B1401 = GRD1.TextMatrix(i, 9)
            frm180303_2.m_B1402 = GRD1.TextMatrix(i, 1)
            If GRD1.TextMatrix(i, 7) <> "" Then
               frm180303_2.m_B1403_A = True
            Else
               frm180303_2.m_B1403_A = False
            End If
            If GRD1.TextMatrix(i, 8) <> "" Then
               frm180303_2.m_B1403_P = True
            Else
               frm180303_2.m_B1403_P = False
            End If
            frm180303_2.Show
            Exit Sub
         Else
            'ShowNoData
            MsgBox "當日無打卡異常資料！"
         End If
      End If
   Next i
'   If bolSelisV = False Then
'      MsgBox "請勾選欲查詢的資料！"
'   End If
End Sub

'打卡明細
Private Sub cmdDetail_Click()
'Dim bolSelisV As Boolean
'
'   bolSelisV = False
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If GRD1.TextMatrix(i, 0) = "V" Then
'         bolSelisV = True
         GRD1.Text = ""
         For j = 0 To GRD1.Cols - 1
            GRD1.col = j
            GRD1.CellBackColor = QBColor(15)
         Next j
         Call SetColColor(i)
         bolClose = False
         Call frm180303_1.SetParent(Me)
         frm180303_1.m_B1401 = GRD1.TextMatrix(i, 9)
         frm180303_1.m_B1402 = GRD1.TextMatrix(i, 1)
         If frm180303_1.QueryData = True Then
            frm180303_1.Show vbModal '強制回應表單
         Else
            Unload frm180303_1
         End If
         If bolClose = True Then
            Exit Sub
         End If
      End If
   Next i
'   If bolSelisV = False Then
'      MsgBox "請勾選欲查詢的資料！"
'   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

'Add By Sindy 2015/12/23
Private Sub cmdPrint_Click()
   
   If Not (GRD1.Rows >= 2 And GRD1.TextMatrix(1, 3) <> "") Then
      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 2 '1.直印 2.橫印
   Printer.PaperSize = 9  'PDF
   
   iLine = 1
   PrintTitle '列印表頭
   For i = GRD1.Rows - 1 To 1 Step -1
      For m_i = 1 To 8
          strTemp(m_i) = ""
      Next m_i
      strTemp(1) = GRD1.TextMatrix(i, 1)
      strTemp(2) = GRD1.TextMatrix(i, 2)
      strTemp(3) = GRD1.TextMatrix(i, 3)
      strTemp(4) = GRD1.TextMatrix(i, 4)
      strTemp(5) = GRD1.TextMatrix(i, 5)
      strTemp(6) = GRD1.TextMatrix(i, 6)
      strTemp(7) = convForm(CheckStr(GRD1.TextMatrix(i, 7)), 40)
      strTemp(8) = convForm(CheckStr(GRD1.TextMatrix(i, 8)), 40)
           
      PrintDetail '列印表中
      
      If iLine >= 37 Then
         Printer.NewPage
         iLine = 1
         PrintTitle '列印表頭
      End If
   Next i
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintTitle()
Dim strText As String

GetPleft

If iLine = 1 Then iLine = iLine + 1

Printer.Font.Size = 18
Printer.Font.Underline = True
Printer.FontBold = True
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("打卡記錄明細表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "打卡記錄明細表"

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False
iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

strText = "打卡日期：" & ChangeTStringToTDateString(txtDate(0)) & " ~ " & ChangeTStringToTDateString(txtDate(1))
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strText) / 2)
Printer.CurrentY = iLine * 300
Printer.Print strText

Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "日期"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "部門"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "員工姓名"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "上班打卡"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "下班打卡"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "有無請假"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iLine * 300
Printer.Print "上班異常"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iLine * 300
Printer.Print "下班異常"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(210, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 1800
PLeft(3) = 3100
PLeft(4) = 4500
PLeft(5) = 6000
PLeft(6) = 7200
PLeft(7) = 8400
PLeft(8) = 11800
End Sub

Sub PrintDetail()
Dim ii As Integer
   
   For ii = 1 To 8
      Printer.CurrentX = PLeft(ii)
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(ii)
   Next ii
   iLine = iLine + 1
End Sub

Private Sub cmdQuery_Click()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCon As String, strConAsc As String, strConB14 As String, strConSCD As String
Dim subSQL0 As String, subSQL1 As String, subSQL2 As String, subSQL3 As String, subSQL4 As String
   
   m_blnColOrderAsc = True
   
   GRD1.Clear
   SetGrd
   strCon = "": strConAsc = "": strConB14 = "": strConSCD = ""
   If Val(txtDate(0)) = 0 Or Val(txtDate(1)) = 0 Then
      MsgBox "請輸入起迄日期！", vbExclamation, "操作錯誤！"
      If Val(txtDate(0)) = 0 Then txtDate(0).SetFocus
      If Val(txtDate(1)) = 0 Then txtDate(1).SetFocus
      Exit Sub
   End If
   '員工代號
   If txtB1003(0) <> "" And txtB1003(1) <> "" Then
      strCon = strCon & " and s1.ST01>='" & txtB1003(0) & "' and s1.ST01<='" & txtB1003(1) & "'"
      strConAsc = strConAsc & " and s1.ST01>='" & txtB1003(0) & "' and s1.ST01<='" & txtB1003(1) & "'"
      strConB14 = strConB14 & " and b1401>='" & txtB1003(0) & "' and b1401<='" & txtB1003(1) & "'"
      strConSCD = strConSCD & " and scd01>='" & txtB1003(0) & "' and scd01<='" & txtB1003(1) & "'"
   End If
   '部門別
   If txtDept(0) <> "" And txtDept(1) <> "" Then
      'Modify By Sindy 2023/12/19
      If strSrvDate(1) >= 新部門啟用日 Then
         strCon = strCon & " and s1.ST93>='" & txtDept(0) & "' and s1.ST93<='" & txtDept(1) & "'"
         strConAsc = strConAsc & " and s1.ST93>='" & txtDept(0) & "' and s1.ST93<='" & txtDept(1) & "'"
      Else
      '2023/12/19 END
         strCon = strCon & " and s1.ST03>='" & txtDept(0) & "' and s1.ST03<='" & txtDept(1) & "'"
         strConAsc = strConAsc & " and s1.ST03>='" & txtDept(0) & "' and s1.ST03<='" & txtDept(1) & "'"
      End If
   End If
   'Modify By Sindy 2021/12/21
   '所屬簽核的人員
   If m_IsAbsBossST03 <> "" Then
      If m_strEmp <> "" Then
         strCon = strCon & " and s1.ST01 in(" & m_strEmp & ")"
      End If
   End If
   '2021/12/21 END
   
   'ADD BY SONIA 2013/7/26 部門別權限限制
   If m_IsAbsBossST03 <> "" Then
      'Modify By Sindy 2023/12/19
      If strSrvDate(1) >= 新部門啟用日 Then
         strCon = strCon & " and s1.ST93 in (" & m_IsAbsBossST03 & ")"
      Else
      '2023/12/19 END
         strCon = strCon & " and s1.ST03 in (" & m_IsAbsBossST03 & ")"
      End If
   End If
   '2013/7/26 END
   '所別
   If txtST06(0) <> "" And txtST06(1) <> "" Then
      strCon = strCon & " and s1.ST06>='" & txtST06(0) & "' and s1.ST06<='" & txtST06(1) & "'"
      strConAsc = strConAsc & " and s1.ST06>='" & txtST06(0) & "' and s1.ST06<='" & txtST06(1) & "'"
   End If
   
   Screen.MousePointer = vbHourglass
   
   '有無請假
   subSQL1 = "(SELECT SA01 as Userid,SA02 as date1,SA04 as date2,'Y' as IsY FROM staff_Absence,staff s1 WHERE SA01=s1.st01(+)" & _
             " and ((SA02>=" & DBDATE(txtDate(0)) & " and SA02<=" & DBDATE(txtDate(1)) & ") or (SA04>=" & DBDATE(txtDate(0)) & " and SA04<=" & DBDATE(txtDate(1)) & ") or (" & DBDATE(txtDate(0)) & " between SA02 and SA04) or (" & DBDATE(txtDate(1)) & " between SA02 and SA04))" & strConAsc & _
             " union SELECT SB01 as Userid,SB02 as date1,SB04 as date2,'Y' as IsY FROM staff_busi_trip,staff s1 WHERE SB01=s1.st01(+)" & _
             " and ((SB02>=" & DBDATE(txtDate(0)) & " and SB02<=" & DBDATE(txtDate(1)) & ") or (SB04>=" & DBDATE(txtDate(0)) & " and SB04<=" & DBDATE(txtDate(1)) & ") or (" & DBDATE(txtDate(0)) & " between SB02 and SB04) or (" & DBDATE(txtDate(1)) & " between SB02 and SB04))" & strConAsc & _
             " union SELECT B1003 as Userid,B1004 as date1,B1006 as date2,'Y' as IsY FROM ABS010,staff s1 WHERE B1002 in('01','03') and B1018 not in('" & 退回 & "','" & 註銷 & "','" & 已核准 & "') and B1003=s1.st01(+)" & _
             " and ((B1004>=" & DBDATE(txtDate(0)) & " and B1004<=" & DBDATE(txtDate(1)) & ") or (B1006>=" & DBDATE(txtDate(0)) & " and B1006<=" & DBDATE(txtDate(1)) & ") or (" & DBDATE(txtDate(0)) & " between B1004 and B1006) or (" & DBDATE(txtDate(1)) & " between B1004 and B1006))" & strConAsc & ") V1"
   '異常確認:
   strSql = "decode(ac03,null,'有異常',ac03||decode(b1408,null,'','，'||st02||' '||sqldatet(b1410)||' '||decode(b1409,'Y','同意','N','不同意',''))||decode(b1411,null,'','，'||decode(b1411,'A','系統確認','B','人事處先確認','C','人事處已確認','')||sqldatet(b1412)||' '||sqltime6(b1413)))"
   '上班異常
   subSQL2 = "(select b1401,b1402,b1404,b1405," & strSql & " as AErr,decode(b1405,'4',decode(b1409,'Y','洽公',''),'6','洽公',ac03) as A05txt" & _
             " From abs014,allcode,staff" & _
             " where ac01(+)='10' and b1405=ac02(+) and b1408=st01(+)" & _
             " and b1402>=" & DBDATE(txtDate(0)) & " and b1402<=" & DBDATE(txtDate(1)) & strConB14 & _
             " and b1403='A') V2"
   '下班異常
   subSQL3 = "(select b1401,b1402,b1404,b1405," & strSql & " as PErr,decode(b1405,'4',decode(b1409,'Y','洽公',''),'6','洽公',ac03) as P05txt" & _
             " From abs014,allcode,staff" & _
             " where ac01(+)='10' and b1405=ac02(+) and b1408=st01(+)" & _
             " and b1402>=" & DBDATE(txtDate(0)) & " and b1402<=" & DBDATE(txtDate(1)) & strConB14 & _
             " and b1403='P') V3"
   '打卡資料查詢(員工代號,打卡日期)
   subSQL0 = "(select distinct b1401,b1402 from(" & _
             " select b1401,b1402" & _
             " From abs014" & _
             " Where b1402>=" & DBDATE(txtDate(0)) & " and b1402<=" & DBDATE(txtDate(1)) & strConB14 & _
             " Union All" & _
             " select scd01 as b1401,pr01 as b1402" & _
             " From pollrecord, staffcarddata " & _
             " Where pr03 = scd02 and pr01>=" & DBDATE(txtDate(0)) & " and pr01<=" & DBDATE(txtDate(1)) & strConSCD & _
             ")) A1"
   '最早最晚打卡時間
   subSQL4 = "(select scd01,pr01,nvl(min(pr02),0) as min_pr02,nvl(max(pr02),0) as max_pr02 from pollrecord, staffcarddata where pr03=scd02(+) and  pr01>=" & DBDATE(txtDate(0)) & " and pr01<=" & DBDATE(txtDate(1)) & strConSCD & " group by scd01,pr01) V4"
   '組合全部SQL
   'Modify By Sindy 2023/12/19
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "select ' ' as V,sqldatet(A1.b1402) as 日期,nvl(A0922,'(舊)'||A0902) as 部門,s1.ST02 as 員工姓名," & _
               "decode(V2.b1401,'',sqltime(min(V4.min_pr02)),decode(V2.b1404,'',decode(V2.A05txt,'','未打卡',V2.A05txt),sqltime(V2.b1404))) as 上班打卡," & _
               "decode(V3.b1401,'',decode(A1.b1402," & strSrvDate(1) & ",'', decode(sqltime(min(V4.min_pr02))||V2.b1401,sqltime(max(V4.max_pr02)),'',sqltime(max(V4.max_pr02))) ),decode(V3.b1404,'',decode(V3.P05txt,'','未打卡',V3.P05txt),nvl(sqltime(max(V4.max_pr02)),sqltime(V3.b1404)))) as 下班打卡," & _
               "V1.IsY as 有無請假," & _
               "decode(V2.A05txt,'',V2.AErr,replace(V2.AErr,V2.b1404||V2.A05txt||'，','')) as 上班異常," & _
               "decode(V3.P05txt,'',V3.PErr,replace(V3.PErr,V3.b1404||V3.P05txt||'，','')) as 下班異常," & _
               "A1.b1401,s1.ST03,min(V4.min_pr02),max(V4.max_pr02),V2.AErr,V3.PErr" & _
               " from " & subSQL0 & ",staff s1,ACC090NEW,ACC090," & subSQL1 & "," & subSQL2 & "," & subSQL3 & "," & subSQL4 & _
               " where A1.b1401=s1.st01(+)" & _
               " and s1.ST93=A0921(+) and s1.ST03=A0901(+)" & _
               " and A1.b1402>=" & DBDATE(txtDate(0)) & " and A1.b1402<=" & DBDATE(txtDate(1)) & strCon & _
               " and A1.B1401=V1.Userid(+) and A1.B1402 between V1.date1(+) and V1.date2(+)" & _
               " and A1.B1401=V2.b1401(+) and A1.B1402=V2.b1402(+)" & _
               " and A1.B1401=V3.b1401(+) and A1.B1402=V3.b1402(+)" & _
               " and A1.B1401=V4.scd01(+) and A1.B1402=V4.pr01(+)" & _
               " group by A1.b1402,A0922,A0921,A0902,s1.ST02,s1.ST06,A1.b1401,V1.IsY,V2.AErr,V3.PErr,s1.ST03,V2.b1401,V2.b1404,V3.b1401,V3.b1404,V2.A05txt,V3.P05txt" & _
               " order by s1.ST06,A0921,A1.b1401,A1.b1402 desc"
   Else
   '2023/12/19 END
      strSql = "select ' ' as V,sqldatet(A1.b1402) as 日期,A0902 as 部門,s1.ST02 as 員工姓名," & _
               "decode(V2.b1401,'',sqltime(min(V4.min_pr02)),decode(V2.b1404,'',decode(V2.A05txt,'','未打卡',V2.A05txt),sqltime(V2.b1404))) as 上班打卡," & _
               "decode(V3.b1401,'',decode(A1.b1402," & strSrvDate(1) & ",'', decode(sqltime(min(V4.min_pr02))||V2.b1401,sqltime(max(V4.max_pr02)),'',sqltime(max(V4.max_pr02))) ),decode(V3.b1404,'',decode(V3.P05txt,'','未打卡',V3.P05txt),nvl(sqltime(max(V4.max_pr02)),sqltime(V3.b1404)))) as 下班打卡," & _
               "V1.IsY as 有無請假," & _
               "decode(V2.A05txt,'',V2.AErr,replace(V2.AErr,V2.b1404||V2.A05txt||'，','')) as 上班異常," & _
               "decode(V3.P05txt,'',V3.PErr,replace(V3.PErr,V3.b1404||V3.P05txt||'，','')) as 下班異常," & _
               "A1.b1401,s1.ST03,min(V4.min_pr02),max(V4.max_pr02),V2.AErr,V3.PErr" & _
               " from " & subSQL0 & ",staff s1,ACC090," & subSQL1 & "," & subSQL2 & "," & subSQL3 & "," & subSQL4 & _
               " where A1.b1401=s1.st01(+)" & _
               " and s1.ST03=A0901(+)" & _
               " and A1.b1402>=" & DBDATE(txtDate(0)) & " and A1.b1402<=" & DBDATE(txtDate(1)) & strCon & _
               " and A1.B1401=V1.Userid(+) and A1.B1402 between V1.date1(+) and V1.date2(+)" & _
               " and A1.B1401=V2.b1401(+) and A1.B1402=V2.b1402(+)" & _
               " and A1.B1401=V3.b1401(+) and A1.B1402=V3.b1402(+)" & _
               " and A1.B1401=V4.scd01(+) and A1.B1402=V4.pr01(+)" & _
               " group by A1.b1402,A0902,A0901,s1.ST02,s1.ST06,A1.b1401,V1.IsY,V2.AErr,V3.PErr,s1.ST03,V2.b1401,V2.b1404,V3.b1401,V3.b1404,V2.A05txt,V3.P05txt" & _
               " order by s1.ST06,A0901,A1.b1401,A1.b1402 desc"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      'Label1.Caption = "共 " & rsTmp.RecordCount & " 筆"
      '逐筆檢查資料
      For i = 1 To GRD1.Rows - 1
'         If GRD1.TextMatrix(i, 3) = "楊毓純" Then
'            MsgBox "測試!!!"
'         End If
         
         '上班打卡需要特殊處理顯示的時間
         If InStr(GRD1.TextMatrix(i, 4), ":") = 0 And GRD1.TextMatrix(i, 4) <> "" Then
'            If GRD1.TextMatrix(i, 4) = "請假" Then
'               If GRD1.TextMatrix(i, 11) <> "" Then
'                  GRD1.TextMatrix(i, 4) = Format(GRD1.TextMatrix(i, 11), "00:00:00")
'                  GRD1.TextMatrix(i, 7) = GRD1.TextMatrix(i, 13)
'               End If
'            Else
               If GRD1.TextMatrix(i, 11) <> "" Then
                  If (Val(GRD1.TextMatrix(i, 12)) - Val(GRD1.TextMatrix(i, 11))) >= 3000 Or _
                     ((Val(GRD1.TextMatrix(i, 12)) - Val(GRD1.TextMatrix(i, 11))) < 3000 And DBDATE(GRD1.TextMatrix(i, 1)) = strSrvDate(1)) Then
                     GRD1.TextMatrix(i, 4) = Format(GRD1.TextMatrix(i, 11), "00:00:00")
                     GRD1.TextMatrix(i, 7) = GRD1.TextMatrix(i, 13)
                  End If
               End If
'            End If
         End If
         If Left(GRD1.TextMatrix(i, 7), 3) = "忘打卡" And InStr(GRD1.TextMatrix(i, 4), ":") > 0 Then
            GRD1.TextMatrix(i, 4) = "忘打卡"
         End If
         If InStr(GRD1.TextMatrix(i, 4), ":") = 0 And GRD1.TextMatrix(i, 4) <> "" Then
            If InStr(GRD1.TextMatrix(i, 7), "，") > 0 Then
               GRD1.TextMatrix(i, 7) = Mid(GRD1.TextMatrix(i, 13), InStrRev(GRD1.TextMatrix(i, 13), "，") + 1)
            End If
         Else
            'Add By Sindy 2013/8/9
            If InStr(GRD1.TextMatrix(i, 7), "洽公請主管批示，") > 0 And InStr(GRD1.TextMatrix(i, 7), "確認") > 0 Then
               GRD1.TextMatrix(i, 7) = Replace(GRD1.TextMatrix(i, 7), "洽公請主管批示，", "洽公，")
            ElseIf InStr(GRD1.TextMatrix(i, 7), "因公未打卡，") > 0 And InStr(GRD1.TextMatrix(i, 7), "確認") > 0 Then
               GRD1.TextMatrix(i, 7) = Replace(GRD1.TextMatrix(i, 7), "因公未打卡，", "洽公，")
            End If
         End If
         '下班打卡需要特殊處理顯示的時間
         If InStr(GRD1.TextMatrix(i, 5), ":") = 0 And GRD1.TextMatrix(i, 5) <> "" Then
'            If GRD1.TextMatrix(i, 5) = "請假" Then
'               If GRD1.TextMatrix(i, 12) <> "" Then
'                  GRD1.TextMatrix(i, 5) = Format(GRD1.TextMatrix(i, 12), "00:00:00")
'                  GRD1.TextMatrix(i, 8) = GRD1.TextMatrix(i, 14)
'               End If
'            Else
               If GRD1.TextMatrix(i, 12) <> "" Then
                  If (Val(GRD1.TextMatrix(i, 12)) - Val(GRD1.TextMatrix(i, 11))) > 3000 Then
                     GRD1.TextMatrix(i, 5) = Format(GRD1.TextMatrix(i, 12), "00:00:00")
                     GRD1.TextMatrix(i, 8) = GRD1.TextMatrix(i, 14)
                  End If
               End If
'            End If
         End If
         If Left(GRD1.TextMatrix(i, 8), 3) = "忘打卡" And InStr(GRD1.TextMatrix(i, 5), ":") > 0 Then
            GRD1.TextMatrix(i, 5) = "忘打卡"
            '若上班也異常時,將打卡時間放至上班打卡欄位裡
            If GRD1.TextMatrix(i, 11) <> "" Then
               GRD1.TextMatrix(i, 4) = Format(GRD1.TextMatrix(i, 11), "00:00:00")
               GRD1.TextMatrix(i, 7) = GRD1.TextMatrix(i, 13)
            End If
         End If
         If InStr(GRD1.TextMatrix(i, 5), ":") = 0 And GRD1.TextMatrix(i, 5) <> "" Then
            If InStr(GRD1.TextMatrix(i, 8), "，") > 0 Then
               GRD1.TextMatrix(i, 8) = Mid(GRD1.TextMatrix(i, 14), InStrRev(GRD1.TextMatrix(i, 14), "，") + 1)
            End If
         Else
            'Add By Sindy 2013/8/9
            If InStr(GRD1.TextMatrix(i, 8), "洽公請主管批示，") > 0 And InStr(GRD1.TextMatrix(i, 8), "確認") > 0 Then
               GRD1.TextMatrix(i, 8) = Replace(GRD1.TextMatrix(i, 8), "洽公請主管批示，", "洽公，")
            ElseIf InStr(GRD1.TextMatrix(i, 8), "因公未打卡，") > 0 And InStr(GRD1.TextMatrix(i, 8), "確認") > 0 Then
               GRD1.TextMatrix(i, 8) = Replace(GRD1.TextMatrix(i, 8), "因公未打卡，", "洽公，")
            End If
         End If
      Next i
      Label4.Caption = "共 " & rsTmp.RecordCount & " 筆"
   Else
      ShowNoData
      Label4.Caption = "共 0 筆"
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
'   If rsTmp.RecordCount > 0 Then
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   GRD1.Visible = True
   
   For i = 1 To GRD1.Rows - 1
      Call SetColColor(i)
   Next i
   
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
   Screen.MousePointer = vbDefault
End Sub

'異常資料以紅色標註
Private Sub SetColColor(intRow As Long)
   GRD1.row = intRow
   If GRD1.TextMatrix(intRow, 7) = "有異常" Then
      GRD1.col = 4
      GRD1.CellBackColor = &H8080FF
      GRD1.col = 7
      GRD1.CellBackColor = &H8080FF
   End If
   If GRD1.TextMatrix(intRow, 8) = "有異常" Then
      GRD1.col = 5
      GRD1.CellBackColor = &H8080FF
      GRD1.col = 8
      GRD1.CellBackColor = &H8080FF
   End If
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
      
   MoveFormToCenter Me
   '前一個月的第一天
   'txtDate(0) = Left(ChangeWStringToTString(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##")))), 5) & "01"
'   '取得前一日的工作天
'   For i = 1 To 12
'      txtDate(0) = Val(DBDATE(DateAdd("d", Val("-" & i), ChangeWStringToWDateString(strSrvDate(1))))) - 19110000
'      If ChkWorkDay(DBDATE(DateAdd("d", Val("-" & i), ChangeWStringToWDateString(strSrvDate(1))))) = True Then
'         Exit For
'      End If
'   Next i
   
   'Modify By Sindy 2025/3/19
   Call PUB_SetQFormCol_ABS(m_IsAbsBossST03, m_strEmp, Me.Name, txtDate(0), txtDate(1), txtB1003(0), txtB1003(1), _
      cboDept(0), cboDept(1), txtDept(0), txtDept(1), txtST06(0), txtST06(1), Me.Label5)
   '2025/3/19 END
   
   Label1.Caption = 接收打卡時間公告
   
   'Add By Sindy 2015/12/23 列印功能僅開放人事處和電腦中心
   If Pub_StrUserSt03 = "M21" Or Pub_StrUserSt03 = "M51" Then
      CmdPrint.Visible = True
      Frame1.Visible = True
   Else
      CmdPrint.Visible = False
      Frame1.Visible = False
   End If
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSql Then
         SeekPrint = i
      End If
   Next i
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
   '2015/12/23 END
   
   SetGrd
   Call cmdQuery_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180303 = Nothing
End Sub

' 初始化列表
Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("V", "日期", "部門", "員工姓名", "上班打卡", "下班打卡", _
                           "有無請假", "上班異常", "下班異常", "b1401", "s1.ST03", "min_pr02", _
                           "max_pr02", "V2.AErr", "V3.PErr")
   arrGridHeadWidth = Array(200, 800, 1000, 800, 800, 800, _
                           800, 1650, 1650, 0, 0, 0, _
                           0, 0, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   If GRD1.TextMatrix(GRD1.MouseRow, 1) <> "" Then
      If GRD1.Text = "V" Then
         GRD1.Text = ""
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         Next i
         Call SetColColor(GRD1.MouseRow)
      Else
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
End If
GRD1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   If nCol = 2 Then nCol = 10 '部門別置換為使用部門別代碼做排序
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If Me.GRD1.Text = "部門別" Then
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub txtB1003_GotFocus(Index As Integer)
   InverseTextBox txtB1003(Index)
End Sub

Private Sub txtB1003_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'2013/7/26 ADD BY SONIA
Private Sub txtB1003_LostFocus(Index As Integer)
   Select Case Index
      Case 0
         txtB1003(1) = txtB1003(0)
   End Select
End Sub
'2013/7/26 END

Private Sub txtB1003_Validate(Index As Integer, Cancel As Boolean)
   If txtB1003(Index).Text <> "" Then
      If txtB1003(Index).Text <> "99996" Then 'Add By Sindy 2017/4/17 + 99996.來賓
         If ChkStaffID(txtB1003(Index)) Then
            Call txtB1003_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      Else
         txtDept(0).Text = ""
         txtDept(1).Text = ""
      End If
   End If
   If Index = 0 Then
      If txtB1003(Index) <> "" And txtB1003(Index + 1) = "" Then
         txtB1003(Index + 1) = txtB1003(Index)
      End If
      If txtB1003(Index) > txtB1003(Index + 1) Then
         txtB1003(Index + 1) = txtB1003(Index)
      End If
   ElseIf Index = 1 Then
      If txtB1003(Index) <> "" And txtB1003(Index - 1) = "" Then
         txtB1003(Index - 1) = txtB1003(Index)
      End If
      If txtB1003(Index - 1) <> "" And txtB1003(Index) <> "" Then
         If RunNick(txtB1003(Index - 1), txtB1003(Index)) Then
            Call txtB1003_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtDept_GotFocus(Index As Integer)
   InverseTextBox txtDept(Index)
End Sub

Private Sub txtDept_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDept_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
      If txtDept(Index) <> "" And txtDept(Index + 1) = "" Then
         txtDept(Index + 1) = txtDept(Index)
      End If
      If txtDept(Index) > txtDept(Index + 1) Then
         txtDept(Index + 1) = txtDept(Index)
      End If
   ElseIf Index = 1 Then
      If txtDept(Index) <> "" And txtDept(Index - 1) = "" Then
         txtDept(Index - 1) = txtDept(Index)
      End If
      If txtDept(Index - 1) <> "" And txtDept(Index) <> "" Then
         If RunNick(txtDept(Index - 1), txtDept(Index)) Then
            Call txtDept_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
   InverseTextBox txtDate(Index)
End Sub

Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
   If txtDate(Index).Text <> "" Then
      If ChkDate(txtDate(Index)) = False Then
         Call txtDate_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
   End If
   If Index = 0 Then
      If txtDate(Index) <> "" And txtDate(Index + 1) = "" Then
         txtDate(Index + 1) = txtDate(Index)
      End If
      If Val(txtDate(Index)) > Val(txtDate(Index + 1)) Then
         txtDate(Index + 1) = txtDate(Index)
      End If
   ElseIf Index = 1 Then
      If txtDate(Index) <> "" And txtDate(Index - 1) = "" Then
         txtDate(Index - 1) = txtDate(Index)
      End If
      If txtDate(Index - 1) <> "" And txtDate(Index) <> "" Then
         If RunNick2(txtDate(Index - 1), txtDate(Index)) Then
            Call txtDate_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtST06_GotFocus(Index As Integer)
   InverseTextBox txtST06(Index)
End Sub

Private Sub txtST06_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtST06_Validate(Index As Integer, Cancel As Boolean)
   If txtST06(Index) <> "" Then
      If CheckLengthIsOK(txtST06(Index), txtST06(Index).MaxLength) = False Then
          Call txtST06_GotFocus(Index)
          Cancel = True
          Exit Sub
      End If
      If Trim(txtST06(Index)) <> "" Then
         If txtST06(Index) <> "1" And txtST06(Index) <> "2" And txtST06(Index) <> "3" And _
            txtST06(Index) <> "4" And txtST06(Index) <> "5" Then
            MsgBox "所別代碼有誤!!!", vbExclamation + vbOKOnly
            Call txtST06_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
   If Index = 0 Then
      If txtST06(Index) <> "" And txtST06(Index + 1) = "" Then
         txtST06(Index + 1) = txtST06(Index)
      End If
      If txtST06(Index) > txtST06(Index + 1) Then
         txtST06(Index + 1) = txtST06(Index)
      End If
   ElseIf Index = 1 Then
      If txtST06(Index) <> "" And txtST06(Index - 1) = "" Then
         txtST06(Index - 1) = txtST06(Index)
      End If
      If txtST06(Index - 1) <> "" And txtST06(Index) <> "" Then
         If RunNick(txtST06(Index - 1), txtST06(Index)) Then
            Call txtST06_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
CloseIme
End Sub
