VERSION 5.00
Begin VB.Form frm12040128 
   BorderStyle     =   1  '單線固定
   Caption         =   "刪除記錄統計表"
   ClientHeight    =   2610
   ClientLeft      =   1605
   ClientTop       =   1785
   ClientWidth     =   5625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5625
   Begin VB.ComboBox cmbPrinter 
      Height          =   276
      ItemData        =   "frm12040128.frx":0000
      Left            =   1368
      List            =   "frm12040128.frx":0002
      TabIndex        =   3
      Top             =   1536
      Width           =   4092
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   4572
      TabIndex        =   5
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3744
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox textOption 
      Height          =   264
      Left            =   1368
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1176
      Width           =   372
   End
   Begin VB.TextBox textDateTo 
      Height          =   264
      Left            =   2688
      MaxLength       =   7
      TabIndex        =   1
      Top             =   816
      Width           =   975
   End
   Begin VB.TextBox textDateFrom 
      Height          =   264
      Left            =   1368
      MaxLength       =   7
      TabIndex        =   0
      Top             =   816
      Width           =   972
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "＊明細表為大報表 !!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   300
      TabIndex        =   10
      Top             =   2010
      Width           =   2820
   End
   Begin VB.Label Label10 
      Caption         =   "印表機 :"
      Height          =   252
      Index           =   0
      Left            =   288
      TabIndex        =   9
      Top             =   1536
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "(1.統計表2.明細表3.清除歷史資料)"
      Height          =   252
      Left            =   1968
      TabIndex        =   8
      Top             =   1176
      Width           =   3372
   End
   Begin VB.Line Line1 
      X1              =   2448
      X2              =   2568
      Y1              =   936
      Y2              =   936
   End
   Begin VB.Label Label2 
      Caption         =   "作業方式："
      Height          =   252
      Left            =   288
      TabIndex        =   7
      Top             =   1176
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "刪除日期："
      Height          =   252
      Left            =   288
      TabIndex        =   6
      Top             =   816
      Width           =   972
   End
End
Attribute VB_Name = "frm12040128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Const m_CharWidth = 120
Const m_CharHeight = 240

Dim m_PaperSize As String
' 宣告報表表頭的欄位其資料型態
Private Type REPORTFIELD
   Name As String
   Left As Long
   Width As Long
End Type
' 表頭欄位的內容
Dim m_Field(8) As REPORTFIELD
' 報表左方留白的寬度
Dim m_LeftMargin As Integer
' 報表上方留白的高度
Dim m_TopMargin As Integer
' 報表頁首的高度
Dim m_HeaderHeight As Integer
' 報表文件的寬度
Dim m_ReportWidth As Integer
' 報表文件中可容納的資料列數
Dim m_ReportDataRows As Integer
' 預設印表機
Dim m_DefaultPrinter As String
' 執行的作業
Dim m_Option As Integer
' 起始日
Dim m_DateFrom As String
' 終止日
Dim m_DateTo As String

Private Type ITEMDATA
   ItemName As String
   ItemNo As String
   ItemDept As String
   ItemDeptNo As String
   ItemCount As Integer
End Type
Dim m_ItemList() As ITEMDATA
Dim m_ItemListCount As Integer
'Add By Cheng 2002/09/10
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub InsertItem(ByVal strItem As String)
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   bFind = False
   For nIndex = 0 To m_ItemListCount - 1
      If m_ItemList(nIndex).ItemNo = strItem Then
         bFind = True
         m_ItemList(nIndex).ItemCount = m_ItemList(nIndex).ItemCount + 1
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_ItemList(m_ItemListCount + 1)
      m_ItemList(m_ItemListCount).ItemNo = strItem
      m_ItemList(m_ItemListCount).ItemName = Empty
      m_ItemList(m_ItemListCount).ItemName = GetStaffName(strItem, True)
      m_ItemList(m_ItemListCount).ItemDeptNo = Empty
      m_ItemList(m_ItemListCount).ItemDeptNo = GetStaffDepartment(strItem)
      If IsEmptyText(m_ItemList(m_ItemListCount).ItemDeptNo) = False Then
         m_ItemList(m_ItemListCount).ItemDept = GetDepartmentName(m_ItemList(m_ItemListCount).ItemDeptNo)
      End If
      m_ItemList(m_ItemListCount).ItemCount = 1
      m_ItemListCount = m_ItemListCount + 1
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim Prn As Printer
   Dim nIndex As Integer
   Dim nSel As Integer
   
   m_DefaultPrinter = Printer.DeviceName
   MoveFormToCenter Me
   
   nSel = 0
   nIndex = 0
   For Each Prn In Printers
      cmbPrinter.AddItem Prn.DeviceName
      If Prn.DeviceName = m_DefaultPrinter Then
         nSel = nIndex
      End If
      nIndex = nIndex + 1
   Next
   cmbPrinter.ListIndex = nSel
   
   cmbPrinter.Enabled = False
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'   Dim nIndex As Integer
'   Dim Prn As Printer
'   '搜尋 Printer
'   For Each Prn In Printers
'      If Prn.DeviceName = m_DefaultPrinter Then
'         Set Printer = Prn
'         Exit For
'      End If
'   Next
'End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim bData As Boolean
   Dim Prn As Printer
   
   If CheckDataValid() = True Then
      If textOption = "1" Then
         m_PaperSize = "A4"
      Else
         m_PaperSize = "REPORT"
         
         '搜尋 Printer, 設定 Printer
         For Each Prn In Printers
            If Prn.DeviceName = cmbPrinter.Text Then
               Set Printer = Prn
               Exit For
            End If
         Next
      End If
      
      Select Case textOption
         Case 1: m_Option = 1
         Case 2: m_Option = 2
         Case 3: m_Option = 3
      End Select
      If IsEmptyText(textDateFrom) = False Then
         m_DateFrom = DBDATE(textDateFrom)
      Else
         m_DateFrom = textDateFrom
      End If
      If IsEmptyText(textDateTo) = False Then
         m_DateTo = DBDATE(textDateTo)
      Else
         m_DateTo = textDateTo
      End If
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 產生並列印資料
      bData = GenerateData()
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      If bData = True Then
         strTit = "資料處理"
         strMsg = "列印結束"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Else
         strTit = "資料處理"
         strMsg = "沒有符合條件的資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      
      ' 還原為預設印表機
      For Each Prn In Printers
      If Prn.DeviceName = m_DefaultPrinter Then
         Set Printer = Prn
         Exit For
      End If
   Next
   End If
End Sub

' 列印分隔線
Public Sub PrintSplitLine(ByVal nRow As Integer)
   Dim nCount As Integer
   For nCount = 0 To m_ReportWidth - 1
      Printer.CurrentX = (m_LeftMargin + nCount) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "-"
   Next nCount
End Sub

' 列印分隔線
Public Sub PrintTerminateLine(ByVal nRow As Integer)
   Dim nCount As Integer
   For nCount = 0 To m_ReportWidth - 1
      Printer.CurrentX = (m_LeftMargin + nCount) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "="
   Next nCount
End Sub

Public Sub BuildTitle(ByVal nDetail As Integer)
   Dim nIndex As Integer
   Dim nLeft As Integer
   
   nLeft = m_LeftMargin
   
   Select Case m_Option
      Case 1:
         For nIndex = 0 To 2
            m_Field(nIndex).Left = nLeft
            Select Case nIndex
               Case 0:
                  m_Field(nIndex).Name = "部門名稱"
                  m_Field(nIndex).Width = 20
               Case 1:
                  m_Field(nIndex).Name = "失誤人員"
                  m_Field(nIndex).Width = 20
               Case 2:
                  m_Field(nIndex).Name = "失誤筆數"
                  m_Field(nIndex).Width = 8
            End Select
            nLeft = nLeft + m_Field(nIndex).Width
         Next nIndex
      Case 2:
         Select Case nDetail
            Case 0:
               For nIndex = 0 To 8: m_Field(nIndex).Name = Empty: Next nIndex
               For nIndex = 0 To 4
                  m_Field(nIndex).Left = nLeft
                  Select Case nIndex
                     Case 0:
                        m_Field(nIndex).Name = "刪除日期"
                        m_Field(nIndex).Width = 20
                     Case 1:
                        m_Field(nIndex).Name = "失誤人員"
                        m_Field(nIndex).Width = 20
                     Case 2:
                        m_Field(nIndex).Name = "原資料產生日期"
                        m_Field(nIndex).Width = 16
                     Case 3:
                        m_Field(nIndex).Name = "原資料產生人員"
                        m_Field(nIndex).Width = 30
                     Case 4:
                        m_Field(nIndex).Name = "刪除備註"
                        m_Field(nIndex).Width = 50
                  End Select
                  nLeft = nLeft + m_Field(nIndex).Width
               Next nIndex
            Case 1:
               For nIndex = 0 To 8: m_Field(nIndex).Name = Empty: Next nIndex
               For nIndex = 0 To 8
                  m_Field(nIndex).Left = nLeft
                  Select Case nIndex
                     Case 0:
                        m_Field(nIndex).Name = "本所案號"
                        m_Field(nIndex).Width = 16
                     Case 1:
                        m_Field(nIndex).Name = "申請人"
                        m_Field(nIndex).Width = 20
                     Case 2:
                        m_Field(nIndex).Name = "申請國家"
                        m_Field(nIndex).Width = 16
                     Case 3:
                        m_Field(nIndex).Name = "申請案號"
                        m_Field(nIndex).Width = 20
                     Case 4:
                        m_Field(nIndex).Name = "專利/商標"
                        m_Field(nIndex).Width = 14
                     Case 5:
                        m_Field(nIndex).Name = "目前准駁"
                        m_Field(nIndex).Width = 10
                     Case 6:
                        m_Field(nIndex).Name = "FC代理人"
                        m_Field(nIndex).Width = 20
                     Case 7:
                        m_Field(nIndex).Name = "年費/延展代理人"
                        m_Field(nIndex).Width = 20
                     Case 8:
                        m_Field(nIndex).Name = "分所案號"
                        m_Field(nIndex).Width = 14
                  End Select
                  nLeft = nLeft + m_Field(nIndex).Width
               Next nIndex
            Case 2:
               For nIndex = 0 To 8: m_Field(nIndex).Name = Empty: Next nIndex
               For nIndex = 0 To 8
                  m_Field(nIndex).Left = nLeft
                  Select Case nIndex
                     Case 0:
                        m_Field(nIndex).Name = "總收文號"
                        m_Field(nIndex).Width = 10
                     Case 1:
                        m_Field(nIndex).Name = "案件性質"
                        m_Field(nIndex).Width = 20
                     Case 2:
                        m_Field(nIndex).Name = "收文日"
                        m_Field(nIndex).Width = 12
                     Case 3:
                        m_Field(nIndex).Name = "智權人員"
                        m_Field(nIndex).Width = 20
                     Case 4:
                        m_Field(nIndex).Name = "本所期限"
                        m_Field(nIndex).Width = 12
                     Case 5:
                        m_Field(nIndex).Name = "法定期限"
                        m_Field(nIndex).Width = 12
                     Case 6:
                        m_Field(nIndex).Name = "費用"
                        m_Field(nIndex).Width = 12
                     Case 7:
                        m_Field(nIndex).Name = "規費"
                        m_Field(nIndex).Width = 12
                     Case 8:
                        m_Field(nIndex).Name = "收據/請款編號"
                        m_Field(nIndex).Width = 16
                  End Select
                  nLeft = nLeft + m_Field(nIndex).Width
               Next nIndex
            Case 3:
               For nIndex = 0 To 8: m_Field(nIndex).Name = Empty: Next nIndex
               For nIndex = 0 To 0
                  m_Field(nIndex).Left = nLeft
                  Select Case nIndex
                     Case 0:
                        m_Field(nIndex).Name = "案件中文名稱"
                        m_Field(nIndex).Width = 60
                  End Select
                  nLeft = nLeft + m_Field(nIndex).Width
               Next nIndex
         End Select
   End Select
End Sub

Public Sub PrintTitle(ByVal nRow As Integer)
   Dim nIndex As Integer
   Dim nCenter As Integer
   Dim strTemp As String
   For nIndex = 0 To 8
      'nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
      'strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
      'Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
      strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
      Printer.CurrentX = m_Field(nIndex).Left * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print strTemp
   Next nIndex
End Sub

' 設定報表欄位的左方位置及其名稱
Public Sub BuildField()
   Dim nIndex As Integer
   Dim nFieldWidth As Integer
   Dim nLeft As Integer

   Select Case m_PaperSize
      Case "REPORT"
         m_LeftMargin = 1
         m_TopMargin = 3
         m_ReportWidth = 154
         m_ReportDataRows = 45
         nFieldWidth = 9
      Case Else
         m_LeftMargin = 3
         m_TopMargin = 2
         If m_Option = 1 Then
            m_ReportWidth = 80
            m_ReportDataRows = 50
         Else
            m_ReportWidth = 130
            m_ReportDataRows = 30
         End If
         nFieldWidth = 7
   End Select

   If m_Option = 1 Then
      nLeft = m_LeftMargin
      For nIndex = 0 To 2
         m_Field(nIndex).Left = nLeft
         Select Case nIndex
            Case 0:
               m_Field(nIndex).Name = "部門名稱"
               m_Field(nIndex).Width = 20
            Case 1:
               m_Field(nIndex).Name = "失誤人員"
               m_Field(nIndex).Width = 20
            Case 2:
               m_Field(nIndex).Name = "失誤筆數"
               m_Field(nIndex).Width = 8
         End Select
         nLeft = nLeft + m_Field(nIndex).Width
      Next nIndex
   End If
End Sub

' 列印表頭
Private Sub PrintPageHeader(ByVal nPage As Integer)
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim nRow As Integer
   Dim nX As Long
   Dim nY As Long
   Dim nCenter As Long
   Dim strTemp As String

   ' 表頭
   nRow = 1
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.FontSize = 24
   Printer.Font.Underline = True
   nX = m_LeftMargin + m_ReportWidth / 2 - 14
   Printer.CurrentX = nX * m_CharWidth
   Select Case m_Option
      Case 1:
         Printer.Print "刪除記錄統計表"
      Case 2:
         Printer.Print "刪除記錄明細表"
   End Select
   
   Printer.Font.Underline = False
   ' 下二列
   nRow = nRow + 2
   
   Printer.FontSize = 12
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "列印人 : " & strUserName
      
   nX = m_LeftMargin + m_ReportWidth - 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "製表日期 : " & Format(Date, "EE/MM/DD")
   
   ' 下一列
   nRow = nRow + 1

   ' 來函收文日
   nX = m_LeftMargin + m_ReportWidth / 2 - 16
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "刪除日期 : "
   ' 印日期的起迄
   nX = nX + 12
   If IsEmptyText(m_DateFrom) = False Then
      Printer.CurrentX = nX * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print Format(TAIWANDATE(m_DateFrom), "###/##/##")
   End If
   nX = nX + 8
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "-"
   nX = nX + 2
   If IsEmptyText(m_DateTo) = False Then
      Printer.CurrentX = nX * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print Format(TAIWANDATE(m_DateTo), "###/##/##")
   End If
   
   ' 頁次
   nX = m_LeftMargin + m_ReportWidth - 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "頁"
   
   nX = m_LeftMargin + m_ReportWidth - 14
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "次 : " & nPage
   
   If m_Option = 1 Then
      ' 列印分隔線
      nRow = nRow + 1
      PrintSplitLine nRow
      
      nRow = nRow + 1
      'For nIndex = 0 To 1
      '   nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
      '   strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
      '   Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
      '   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      '   Printer.Print strTemp
      'Next nIndex
      For nIndex = 0 To 2
         strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
         Printer.CurrentX = m_Field(nIndex).Left * m_CharWidth
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print strTemp
      Next nIndex
      
      ' 列印分隔線
      nRow = nRow + 1
      PrintSplitLine nRow
   Else
      ' 列印分隔線
      nRow = nRow + 1
      'PrintSplitLine nRow
      PrintTerminateLine nRow
   End If
      
   m_HeaderHeight = nRow
End Sub

Private Sub GenerateReport1(ByRef rsTmp As ADODB.Recordset)
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(8) As String
   Dim nType As Integer
   Dim nIndex As Integer
   Dim nCenter As Long
   Dim nLeft As Long
   Dim nRight As Long
   Dim nPos As Long
   Dim nField As Integer
   Dim nAmount As Integer
   Dim nX As Long
   Dim nY As Long
   Dim tmpItem As ITEMDATA
   
   'Add By Cheng 2002/01/09
   Dim strDepName As String
   
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      If IsNull(rsTmp.Fields("DD23")) = False Then
         InsertItem rsTmp.Fields("DD23")
      Else
         InsertItem Empty
      End If
      rsTmp.MoveNext
   Loop
   
   ' 排序
   For nX = 0 To m_ItemListCount - 1
      For nY = nX To m_ItemListCount - 1
         If m_ItemList(nX).ItemDeptNo > m_ItemList(nY).ItemDeptNo Then
            tmpItem = m_ItemList(nX)
            m_ItemList(nX) = m_ItemList(nY)
            m_ItemList(nY) = tmpItem
         ElseIf m_ItemList(nX).ItemDeptNo = m_ItemList(nY).ItemDeptNo Then
            tmpItem = m_ItemList(nX)
            m_ItemList(nX) = m_ItemList(nY)
            m_ItemList(nY) = tmpItem
         End If
      Next nY
   Next nX
   
   ' 紙張大小, 方向
   Select Case m_PaperSize
      Case "A4":
         Printer.PaperSize = vbPRPSA4
         If m_Option = 1 Then
            Printer.Orientation = vbPRORPortrait
         Else
            Printer.Orientation = vbPRORLandscape
         End If
      Case "REPORT":
         Printer.PaperSize = vbPRPSFanfoldUS
      Case Else:
         Printer.PaperSize = vbPRPSA4
         If m_Option = 1 Then
            Printer.Orientation = vbPRORPortrait
         Else
            Printer.Orientation = vbPRORLandscape
         End If
   End Select
   
   BuildField
      
   ' 印表頭
   nPage = 1
   PrintPageHeader nPage
   nRow = 1

   'Add By Cheng 2002/01/09
   '若部門名稱相同則不用顯示
   strDepName = ""

   For nIndex = 0 To m_ItemListCount - 1
      ' 若列數超過頁面的高度限制時則換頁
      If nRow > m_ReportDataRows Then
         Printer.NewPage
         nPage = nPage + 1
         PrintPageHeader nPage
         nRow = 1
         'Add By Cheng 2002/01/09
         strDepName = ""
      End If

      ' 清除欄位
      For nField = 0 To 2: fld(nField) = Empty: Next nField
      
      'Modify By Cheng 2002/01/09
      If strDepName <> m_ItemList(nIndex).ItemDept Then
         
         fld(0) = m_ItemList(nIndex).ItemDept
         
         strDepName = fld(0)
      Else
         fld(0) = ""
      End If

      If IsEmptyText(m_ItemList(nIndex).ItemName) = True Then
         fld(1) = m_ItemList(nIndex).ItemNo
      Else
         fld(1) = m_ItemList(nIndex).ItemName
      End If
      fld(2) = m_ItemList(nIndex).ItemCount
      
      ' 輸出
      For nField = 0 To 2
         Select Case nField
            Case 0, 1:
               Printer.CurrentX = m_Field(nField).Left * m_CharWidth
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
            Case Else:
               nLeft = m_Field(nField).Left + (m_Field(nField).Width - 2) - StrLength(fld(nField))
               Printer.CurrentX = nLeft * m_CharWidth
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
         End Select
      Next nField
      ' 列數加一
      nRow = nRow + 1
   Next nIndex
   
   ' 列印分隔列
   PrintSplitLine m_HeaderHeight + nRow
   
   ' 列數加一
   nRow = nRow + 1
      
   ' 清除欄位
   For nField = 0 To 2: fld(nField) = Empty: Next nField
   fld(0) = "總筆數 : "
   nAmount = 0
   For nField = 0 To m_ItemListCount - 1
      nAmount = nAmount + m_ItemList(nField).ItemCount
   Next nField
   fld(2) = CStr(nAmount)
   ' 輸出
   For nField = 0 To 2
      Select Case nField
         Case 0, 1:
            Printer.CurrentX = m_Field(nField).Left * m_CharWidth
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
         Case Else:
            nLeft = m_Field(nField).Left + (m_Field(nField).Width - 2) - StrLength(fld(nField))
            Printer.CurrentX = nLeft * m_CharWidth
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
      End Select
   Next nField
   
   ' 列數加一
   nRow = nRow + 1
   ' 列印分隔線
   PrintTerminateLine m_HeaderHeight + nRow

   Printer.EndDoc
   
   If m_ItemListCount > 0 Then
      Erase m_ItemList
   End If
   m_ItemListCount = 0
End Sub

Private Sub GenerateReport2(ByRef rsTmp As ADODB.Recordset)
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(12) As String
   Dim nType As Integer
   Dim nIndex As Integer
   Dim nCenter As Long
   Dim nLeft As Long
   Dim nRight As Long
   Dim nPos As Long
   Dim nField As Integer
   Dim strNation As String 'Add By Sindy 2015/8/13
   
   BuildField
   
   ' 紙張大小, 方向
   Select Case m_PaperSize
      Case "A4":
         Printer.PaperSize = vbPRPSA4
         If m_Option = 1 Then
            Printer.Orientation = vbPRORPortrait
         Else
            Printer.Orientation = vbPRORLandscape
         End If
      Case "REPORT":
         'modify by sonia 2018/2/22
         'Printer.PaperSize = vbPRPSFanfoldUS
         Printer.PaperSize = PUB_GetPaperSize(15)
      Case Else:
         Printer.PaperSize = vbPRPSA4
         If m_Option = 1 Then
            Printer.Orientation = vbPRORPortrait
         Else
            Printer.Orientation = vbPRORLandscape
         End If
   End Select
   
   ' 印表頭
   nPage = 1
   PrintPageHeader nPage
   nRow = 1
   
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      ' 若列數超過頁面的高度限制時則換頁
      If nRow > m_ReportDataRows Then
         Printer.NewPage
         nPage = nPage + 1
         PrintPageHeader nPage
         nRow = 1
      End If
      
      ' 第一列的Title
      BuildTitle 0
      PrintTitle m_HeaderHeight + nRow
      ' 下一列
      'nRow = nRow + 1
      ' 列印分隔線
      'PrintSplitLine m_HeaderHeight + nRow
      ' 下一列
      nRow = nRow + 1
      ' 清除欄位
      For nField = 0 To 12: fld(nField) = Empty: Next nField
      ' 放入資料
      If IsNull(rsTmp.Fields("DD27")) = False Then: fld(0) = rsTmp.Fields("DD27")
      If IsNull(rsTmp.Fields("DD23")) = False Then: fld(1) = rsTmp.Fields("DD23")
      If IsNull(rsTmp.Fields("DD25")) = False Then: fld(2) = rsTmp.Fields("DD25")
      If IsNull(rsTmp.Fields("DD26")) = False Then: fld(3) = rsTmp.Fields("DD26")
      If IsNull(rsTmp.Fields("DD24")) = False Then: fld(4) = rsTmp.Fields("DD24")
      ' 輸出
      'For nField = 0 To 4
      '   Select Case nField
      '      Case 1, 3:
      '         Printer.CurrentX = m_Field(nField).Left * m_CharWidth
      '         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
      '         Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
      '      Case Else:
      '         nLeft = m_Field(nField).Left + (m_Field(nField).Width / 2) - (StrLength(fld(nField)) / 2)
      '         Printer.CurrentX = nLeft * m_CharWidth
      '         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
      '         Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
      '   End Select
      'Next nField
      For nField = 0 To 4
         Printer.CurrentX = m_Field(nField).Left * m_CharWidth
         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
         Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
      Next nField
      ' 下一列
      nRow = nRow + 1
      ' 列印分隔線
      PrintSplitLine m_HeaderHeight + nRow
      ' 下一列
      nRow = nRow + 1
      
      ' 第二列的Title
      BuildTitle 1
      PrintTitle m_HeaderHeight + nRow
      ' 下一列
      'nRow = nRow + 1
      ' 列印分隔線
      'PrintSplitLine m_HeaderHeight + nRow
      ' 下一列
      nRow = nRow + 1
      ' 清除欄位
      For nField = 0 To 12: fld(nField) = Empty: Next nField
      If rsTmp.Fields("DD01") = "TF" Then
         fld(0) = rsTmp.Fields("DD01") & "-" & Mid(rsTmp.Fields("DD02"), 1, 5) & "-" & Mid(rsTmp.Fields("DD02"), 6, 1) & "-" & rsTmp.Fields("DD03") & "-" & rsTmp.Fields("DD04")
      Else
         fld(0) = rsTmp.Fields("DD01") & "-" & rsTmp.Fields("DD02") & "-" & rsTmp.Fields("DD03") & "-" & rsTmp.Fields("DD04")
      End If
      strNation = GetPrjNation(rsTmp.Fields("DD01") & "-" & rsTmp.Fields("DD02") & "-" & rsTmp.Fields("DD03") & "-" & rsTmp.Fields("DD04")) 'Add By Sindy 2015/8/13
      If IsNull(rsTmp.Fields("DD06")) = False Then: fld(1) = rsTmp.Fields("DD06")
      If IsNull(rsTmp.Fields("DD07")) = False Then: fld(2) = rsTmp.Fields("DD07")
      If IsNull(rsTmp.Fields("DD08")) = False Then: fld(3) = rsTmp.Fields("DD08")
      Select Case rsTmp.Fields("DD01")
         Case "T", "TF", "FCT", "CFT":
            If IsNull(rsTmp.Fields("DD10")) = False Then
               'Modify By Sindy 2015/8/13
               'fld(4) = GetTradeMarkName(rsTmp.Fields("DD10"), 0)
               fld(4) = GetTradeMarkName(rsTmp.Fields("DD10"), IIf(strNation = "020", 1, 0))
               '2015/8/13 END
            End If
         Case "P", "CFP", "FCP":
            If IsNull(rsTmp.Fields("DD10")) = False Then: fld(4) = GetPatentName(rsTmp.Fields("DD10"), 0)
      End Select
      If IsNull(rsTmp.Fields("DD10")) = False Then
         If IsEmptyText(fld(4)) = True Then
            fld(4) = rsTmp.Fields("DD10")
         End If
      End If
      If IsNull(rsTmp.Fields("DD11")) = False Then
         Select Case rsTmp.Fields("DD11")
            Case "1":
               fld(5) = "准"
            Case "2":
               fld(5) = "駁"
            Case Else:
               fld(5) = rsTmp.Fields("DD11")
         End Select
      End If
      If IsNull(rsTmp.Fields("DD12")) = False Then: fld(6) = rsTmp.Fields("DD12")
      If IsNull(rsTmp.Fields("DD13")) = False Then: fld(7) = rsTmp.Fields("DD13")
      If IsNull(rsTmp.Fields("DD09")) = False Then: fld(8) = rsTmp.Fields("DD09")
      ' 輸出
      'For nField = 0 To 8
      '   Select Case nField
      '      Case 0, 1, 2, 3, 4, 6, 7, 8:
      '         Printer.CurrentX = m_Field(nField).Left * m_CharWidth
      '         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
      '         Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
      '      Case Else:
      '         nLeft = m_Field(nField).Left + (m_Field(nField).Width / 2) - (StrLength(fld(nField)) / 2)
      '         Printer.CurrentX = nLeft * m_CharWidth
      '         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
      '         Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
      '   End Select
      'Next nField
      For nField = 0 To 8
         Printer.CurrentX = m_Field(nField).Left * m_CharWidth
         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
         Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
      Next nField
      ' 下一列
      nRow = nRow + 1
      ' 列印分隔線
      PrintSplitLine m_HeaderHeight + nRow
      ' 下一列
      nRow = nRow + 1
      
      If IsNull(rsTmp.Fields("DD14")) = False Then
         If IsEmptyText(rsTmp.Fields("DD14")) = False Then
            ' 第三列的Title
            BuildTitle 2
            PrintTitle m_HeaderHeight + nRow
            ' 下一列
            'nRow = nRow + 1
            ' 列印分隔線
            'PrintSplitLine m_HeaderHeight + nRow
            ' 下一列
            nRow = nRow + 1
            ' 清除欄位
            For nField = 0 To 12: fld(nField) = Empty: Next nField
            
            If IsNull(rsTmp.Fields("DD14")) = False Then: fld(0) = rsTmp.Fields("DD14")
            If IsNull(rsTmp.Fields("DD15")) = False Then: fld(1) = rsTmp.Fields("DD15")
            If IsNull(rsTmp.Fields("DD18")) = False Then: fld(2) = rsTmp.Fields("DD18")
            If IsNull(rsTmp.Fields("DD19")) = False Then: fld(3) = rsTmp.Fields("DD19")
            If IsNull(rsTmp.Fields("DD16")) = False Then: fld(4) = rsTmp.Fields("DD16")
            If IsNull(rsTmp.Fields("DD17")) = False Then: fld(5) = rsTmp.Fields("DD17")
            If IsNull(rsTmp.Fields("DD20")) = False Then: fld(6) = rsTmp.Fields("DD20")
            If IsNull(rsTmp.Fields("DD21")) = False Then: fld(7) = rsTmp.Fields("DD21")
            If IsNull(rsTmp.Fields("DD22")) = False Then: fld(8) = rsTmp.Fields("DD22")
            ' 輸出
            'For nField = 0 To 8
            '   Select Case nField
            '      Case 0, 1, 3, 8:
            '         Printer.CurrentX = m_Field(nField).Left * m_CharWidth
            '         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            '         Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
            '      Case 6, 7:
            '         nLeft = m_Field(nField).Left + (m_Field(nField).Width - 4) - StrLength(fld(nField))
            '         Printer.CurrentX = nLeft * m_CharWidth
            '         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            '         Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
            '      Case Else:
            '         nLeft = m_Field(nField).Left + (m_Field(nField).Width / 2) - (StrLength(fld(nField)) / 2)
            '         Printer.CurrentX = nLeft * m_CharWidth
            '         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            '         Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
            '   End Select
            'Next nField
            ' 輸出
            For nField = 0 To 8
               Printer.CurrentX = m_Field(nField).Left * m_CharWidth
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
            Next nField
            ' 下一列
            nRow = nRow + 1
         End If
      End If
         
      If IsNull(rsTmp.Fields("DD05")) = False Then
         If IsEmptyText(rsTmp.Fields("DD05")) = False Then
            ' 第四列的Title
            BuildTitle 3
            PrintTitle m_HeaderHeight + nRow
            ' 下一列
            nRow = nRow + 1
            ' 列印分隔線
            'PrintSplitLine m_HeaderHeight + nRow
            ' 下一列
            'nRow = nRow + 1
            ' 清除欄位
            For nField = 0 To 12: fld(nField) = Empty: Next nField
                        
            If IsNull(rsTmp.Fields("DD05")) = False Then: fld(0) = rsTmp.Fields("DD05")
            'For nField = 0 To 0
            '   Select Case nField
            '      Case 0:
            '         Printer.CurrentX = m_Field(nField).Left * m_CharWidth
            '         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            '         Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
            '      Case Else:
            '         nLeft = m_Field(nField).Left + (m_Field(nField).Width / 2) - (StrLength(fld(nField)) / 2)
            '         Printer.CurrentX = nLeft * m_CharWidth
            '         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            '         Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
            '   End Select
            'Next nField
            For nField = 0 To 0
               Printer.CurrentX = m_Field(nField).Left * m_CharWidth
               Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
               Printer.Print LeftStr(fld(nField), m_Field(nField).Width)
            Next nField
            
            ' 下一列
            nRow = nRow + 1
            
         End If
      End If
      
      ' 列印分隔線
      PrintTerminateLine m_HeaderHeight + nRow
      
      ' 下一列
      nRow = nRow + 1
      
      rsTmp.MoveNext
   Loop
   
   Printer.EndDoc
   
End Sub

Private Function GenerateData() As Boolean
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Dim strSubSQL As String
   Dim bData As Boolean
   Dim nRecords As Long
   
   bData = False
   Select Case m_Option
      ' 統計表
      Case 1:
         strSubSQL = Empty
         strSql = "SELECT * FROM DATADELETERECORD "
         If IsEmptyText(m_DateFrom) = False Then
            If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & "AND "
            strSubSQL = strSubSQL & "DD27 >= " & DBDATE(m_DateFrom) & " "
         End If
         If IsEmptyText(m_DateTo) = False Then
            If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & "AND "
            strSubSQL = strSubSQL & "DD27 <= " & DBDATE(m_DateTo) & " "
         End If
         If IsEmptyText(strSubSQL) = False Then
            strSql = strSql & "WHERE " & strSubSQL
         End If
         Set rsTmp = New ADODB.Recordset
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            bData = True
            GenerateReport1 rsTmp
         End If
         rsTmp.Close
         Set rsTmp = Nothing
      ' 明細表
      Case 2:
         strSql = "SELECT DD01,DD02,DD03,DD04,DD05,NVL(NVL(NVL(C1.CU04,C1.CU05),C1.CU06),DD06) AS DD06,NVL(N1.NA03, DD07) AS DD07,DD08,DD09,DD10,DD11,NVL(NVL(NVL(F1.FA04,F1.FA05),F1.FA06),DD12) AS DD12,NVL(NVL(NVL(F2.FA04,F2.FA05),F2.FA06),DD13) AS DD13,DD14,NVL(C2.CPM03,DD15) AS DD15,SUBSTR(' '||sqldatet(DD16),-9) AS DD16,SUBSTR(' '||sqldatet(DD17),-9) AS DD17,SUBSTR(' '||sqldatet(DD18),-9) AS DD18,NVL(S1.ST02,DD19) AS DD19,DD20,DD21,DD22,NVL(S2.ST02,DD23) AS DD23,DD24,SUBSTR(' '||sqldatet(DD25),-9) AS DD25,NVL(S3.ST02,DD26) AS DD26,SUBSTR(' '||sqldatet(DD27),-9) AS DD27 " & _
                  "FROM DATADELETERECORD, CUSTOMER C1, CASEPROPERTYMAP C2, NATION N1, STAFF S1, STAFF S2, STAFF S3, FAGENT F1, FAGENT F2 " & _
                  "WHERE DD07 = N1.NA01(+) AND " & _
                        "SUBSTR(DD06,1,8) = C1.CU01(+) AND " & _
                        "SUBSTR(DD06,9,1) = C1.CU02(+) AND " & _
                        "SUBSTR(DD12,1,8) = F1.FA01(+) AND " & _
                        "SUBSTR(DD12,9,1) = F1.FA02(+) AND " & _
                        "SUBSTR(DD13,1,8) = F2.FA01(+) AND " & _
                        "SUBSTR(DD13,9,1) = F2.FA02(+) AND " & _
                        "DD01 = C2.CPM01(+) AND " & _
                        "DD15 = C2.CPM02(+) AND " & _
                        "DD19 = S1.ST01(+) AND " & _
                        "DD23 = S2.ST01(+) AND " & _
                        "DD26 = S3.ST01(+) "
         If IsEmptyText(m_DateFrom) = False Then
            strSql = strSql & "AND DD27 >= " & DBDATE(m_DateFrom) & " "
         End If
         If IsEmptyText(m_DateTo) = False Then
            strSql = strSql & "AND DD27 <= " & DBDATE(m_DateTo) & " "
         End If
         strSql = strSql & "ORDER BY DD27, DD23, DD01,DD02,DD03,DD04 "
         Set rsTmp = New ADODB.Recordset
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            bData = True
            GenerateReport2 rsTmp
         End If
         rsTmp.Close
         Set rsTmp = Nothing
      ' 刪除資料
      Case 3:
         strSubSQL = Empty
         strSql = "DELETE FROM DATADELETERECORD "
         If IsEmptyText(m_DateFrom) = False Then
            If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & "AND "
            strSubSQL = strSubSQL & "DD27 >= " & DBDATE(m_DateFrom) & " "
         End If
         If IsEmptyText(m_DateTo) = False Then
            If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & "AND "
            strSubSQL = strSubSQL & "DD27 <= " & DBDATE(m_DateTo) & " "
         End If
         If IsEmptyText(strSubSQL) = False Then
            strSql = strSql & "WHERE " & strSubSQL
         End If
         nRecords = 0
         cnnConnection.Execute strSql, nRecords
         
         If nRecords > 0 Then
            bData = True
         End If
   End Select
   
   GenerateData = bData
End Function

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   'Add By Cheng 2002/09/10
   blnClkSure = False
   
   ' 刪除日期起日不可空白
   If IsEmptyText(textDateFrom) = True Then
      strTit = "檢核資料"
      strMsg = "刪除日期起日不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textDateFrom.SetFocus
      GoTo EXITSUB
   End If
   
   ' 刪除日期迄日不可空白
   If IsEmptyText(textDateTo) = True Then
      strTit = "檢核資料"
      strMsg = "刪除日期迄日不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textDateTo.SetFocus
      GoTo EXITSUB
   End If

   'Add By Cheng 2002/03/21
   If PUB_CheckKeyInDate(Me.textDateFrom) = -1 Then
      Me.textDateFrom.SetFocus
      textDateFrom_GotFocus
      GoTo EXITSUB
   End If
   If PUB_CheckKeyInDate(Me.textDateTo) = -1 Then
      Me.textDateTo.SetFocus
      textDateTo_GotFocus
      GoTo EXITSUB
   End If

   ' 刪除日期範圍
   If IsEmptyText(textDateFrom) = False And IsEmptyText(textDateTo) = False Then
      If Val(DBDATE(textDateFrom)) > Val(DBDATE(textDateTo)) Then
         strTit = "檢核資料"
         strMsg = "刪除日期範圍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         'Add By Cheng 2002/09/10
         blnClkSure = True
         textDateFrom.SetFocus
         textDateFrom_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 作業方式
   If IsEmptyText(textOption) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入作業方式"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textOption.SetFocus
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm12040128 = Nothing
End Sub

' 刪除日期起日
Private Sub textDateFrom_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDateFrom) = False Then
      If CheckIsTaiwanDate(textDateFrom, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "刪除日期起日格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDateFrom_GotFocus
      End If
   End If
End Sub

Private Sub textDateTo_LostFocus()
   'Add By Cheng 2002/09/10
   If blnClkSure = False Then
      If Me.textDateFrom.Text <> "" And Me.textDateTo.Text <> "" Then
         If Val(Me.textDateFrom.Text) > Val(Me.textDateTo.Text) Then
            MsgBox "刪除日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.textDateFrom.SetFocus
            textDateFrom_GotFocus
            Exit Sub
         End If
      End If
   Else
      blnClkSure = False
   End If
End Sub

' 刪除日期止日
Private Sub textDateTo_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDateTo) = False Then
      If CheckIsTaiwanDate(textDateTo, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "刪除日期止日格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDateTo_GotFocus
      End If
   End If
End Sub

' 作業方式
Private Sub textOption_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim Prn As Printer
   Dim nIndex As Integer
   Dim nSel As Integer
   
   Cancel = False
   If IsEmptyText(textOption) = False Then
      Select Case textOption
         Case "1", "2", "3":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "作業方式只可輸入1, 2 或 3"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textOption_GotFocus
      End Select
      
      cmbPrinter.Enabled = True
      
      ' 設定預設印表機為列印的印表機
      If textOption = "1" Then
         nSel = cmbPrinter.ListIndex
         nIndex = 0
         For nIndex = 0 To cmbPrinter.ListCount - 1
            If cmbPrinter.List(nIndex) = m_DefaultPrinter Then
               nSel = nIndex
            End If
         Next
         cmbPrinter.ListIndex = nSel
         cmbPrinter.Enabled = False
      End If
   End If
End Sub

Private Sub textDateFrom_GotFocus()
   InverseTextBox textDateFrom
End Sub

Private Sub textDateTo_GotFocus()
   InverseTextBox textDateTo
End Sub

Private Sub textOption_GotFocus()
   InverseTextBox textOption
End Sub

Private Function LeftStr(ByVal strData As String, ByVal nLen As Integer) As String
   LeftStr = StrConv(MidB(StrConv(strData, vbFromUnicode), 1, nLen), vbUnicode)
End Function


