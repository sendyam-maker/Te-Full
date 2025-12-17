VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160305 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式名單"
   ClientHeight    =   4220
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4220
   ScaleWidth      =   5700
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   2940
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1710
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   2100
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1710
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2100
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1380
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   1380
      MaxLength       =   1
      TabIndex        =   0
      Top             =   720
      Width           =   195
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2100
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1050
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   1
      Left            =   1380
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1050
      Width           =   615
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4620
      TabIndex        =   8
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3660
      TabIndex        =   7
      Top             =   60
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   60
      TabIndex        =   9
      Top             =   3540
      Width           =   5500
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   180
         Width           =   4500
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   3790
      Left            =   690
      TabIndex        =   16
      Top             =   4530
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   6685
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label7 
      Caption         =   $"frm160305.frx":0000
      Height          =   1300
      Left            =   480
      TabIndex        =   17
      Top             =   2160
      Width           =   4000
   End
   Begin VB.Line Line1 
      X1              =   2730
      X2              =   3120
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "上次簽保證書日期："
      Height          =   180
      Left            =   480
      TabIndex        =   15
      Top             =   1740
      Width           =   1620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "計算年資截止日期："
      Height          =   180
      Left            =   480
      TabIndex        =   14
      Top             =   1410
      Width           =   1620
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   2280
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1.資深人員 2.保證書 3.舊制退休金退休人員)"
      Height          =   180
      Left            =   1620
      TabIndex        =   13
      Top             =   750
      Width           =   3495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "名單類別："
      Height          =   180
      Left            =   480
      TabIndex        =   12
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門代號："
      Height          =   180
      Left            =   480
      TabIndex        =   11
      Top             =   1080
      Width           =   900
   End
End
Attribute VB_Name = "frm160305"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
'Create by SINDY 2009/01/14
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim m_Year As Double
Dim m_intMaxYear As Integer, m_intCnt As String


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
   '        If Trim(txt1(0)) & Trim(txt1(1)) & Trim(txt1(2)) & Trim(txt1(3)) & Trim(txt1(4)) & Trim(txt1(5)) = "" Then
   '            MsgBox "請至少輸入一項列印條件！", vbInformation, "操作錯誤！"
   '            txt1(0).SetFocus
   '            Exit Sub
   '        End If
           If Trim(txt1(0)) = "" Then
               MsgBox "名單類別不可以空白！", vbInformation, "操作錯誤！"
               txt1(0).SetFocus
               Exit Sub
           End If
           '資深人員名單
           'modify by sonia 2016/3/24 +3.舊制退休金退休人員
           If txt1(0) = "1" Or txt1(0) = "3" Then
               If Trim(txt1(3)) = "" Then
                   MsgBox "計算年資截止日期不可以空白！", vbInformation, "操作錯誤！"
                   txt1(3).SetFocus
                   Exit Sub
               End If
           '保證書名單
           ElseIf txt1(0) = "2" Then
               If Trim(txt1(4)) = "" Then
                   MsgBox "上次簽保證書起始日期不可以空白！", vbInformation, "操作錯誤！"
                   txt1(4).SetFocus
                   Exit Sub
               End If
               If Trim(txt1(5)) = "" Then
                   MsgBox "上次簽保證書終止日期不可以空白！", vbInformation, "操作錯誤！"
                   txt1(5).SetFocus
                   Exit Sub
               End If
           End If
           
           Screen.MousePointer = vbHourglass
           m_StrSQL = ""
           '資深人員名單
           'modify by sonia 2016/3/24 +3.舊制退休金退休人員
           If txt1(0) = "1" Or txt1(0) = "3" Then
               If txt1(1) <> "" Then
                   'Modify By Sindy 2023/12/27 部門調整改抓ST93
                   m_StrSQL = m_StrSQL & " and st93>='" & txt1(1) & "' "
               End If
               If txt1(2) <> "" Then
                   'Modify By Sindy 2023/12/27 部門調整改抓ST93
                   m_StrSQL = m_StrSQL & " and st93<='" & txt1(2) & "' "
               End If
           '保證書名單
           ElseIf txt1(0) = "2" Then
               If txt1(1) <> "" Then
                   'Modify By Sindy 2023/12/27 部門調整改抓ST93
                   m_StrSQL = m_StrSQL & " and st93>='" & txt1(1) & "' "
               End If
               If txt1(2) <> "" Then
                   'Modify By Sindy 2023/12/27 部門調整改抓ST93
                   m_StrSQL = m_StrSQL & " and st93<='" & txt1(2) & "' "
               End If
               If txt1(4) <> "" Then
                   m_StrSQL = m_StrSQL & " and st32>='" & DBDATE(txt1(4)) & "' "
               End If
               If txt1(5) <> "" Then
                   m_StrSQL = m_StrSQL & " and st32<='" & DBDATE(txt1(5)) & "' "
               End If
           End If
           '資深人員名單
           If txt1(0) = "1" Then
               StrMenu1
           '保證書名單
           ElseIf txt1(0) = "2" Then
               StrMenu2
           'add by sonia 2016/3/24 舊制退休金退休人員
           ElseIf txt1(0) = "3" Then
               StrMenu3
           'end 2016/3/24
           End If
           Screen.MousePointer = vbDefault
      Case 1
           Unload Me
   End Select
End Sub

Private Sub SetGrd2()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   GRD2.Clear
   GRD2.Rows = 1
   arrGridHeadText = Array("員工編號", "姓名", "年資", "排序欄位")
   arrGridHeadWidth = Array(800, 800, 800, 1200)
   'grd2.Visible = False
   GRD2.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD2.Cols - 1
      GRD2.row = 0
      GRD2.col = iRow
      GRD2.Text = arrGridHeadText(iRow)
      GRD2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD2.CellAlignment = flexAlignCenterCenter
   Next iRow
   'grd2.Visible = True
End Sub

Function GetData() As Boolean
Dim dblYear As Double

   Call SetGrd2
   'Modify By Sindy 2023/12/27 部門調整改抓ST93
   'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
   m_str = "SELECT ST01,ST02 " & _
                "From Staff,SalaryData " & _
                "WHERE ST04='1' " & _
                "and ST01=SD01 " & _
                "and ((sd02 not in('P','F') or sd02 is null) or ST01='68007') and not(substr(st01,5,1)>='A') " & _
                "and ST93 not in('R04') " & m_StrSQL & _
                "Order By ST01 "
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      GetData = True
      With m_rs
         m_rs.MoveFirst
         Do While Not m_rs.EOF
            dblYear = CalYear(CheckStr(m_rs.Fields(0)), DBDATE(txt1(3))) '取得年資
            If m_rs.Fields(0) = "84002" Then
               MsgBox m_rs.Fields(0)
            End If
            GRD2.AddItem CheckStr(m_rs.Fields(0)) & Chr(9) & _
                                  CheckStr(m_rs.Fields(1)) & Chr(9) & _
                                  dblYear & Chr(9) & _
                                  (Mid(CStr(Format(dblYear, "#000.0")), 2, 2) & CheckStr(m_rs.Fields(0)))
            m_rs.MoveNext
         Loop
      End With
   Else
      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      GetData = False
   End If
End Function

Sub StrMenu1()
Dim i As Integer, j As Integer
Dim intRow As Integer, Index As Integer
Dim intItem As Integer
Dim iLst As Integer 'Added by Morgan 2023/11/15

   '取得各員工年資資料
   If GetData = False Then
      Exit Sub
   End If
   
   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   'Printer.PaperSize = 9  'PDF
   
   GRD2.col = 3
   GRD2.Sort = 2 '一般降冪。執行估計文字不管是字串或者是數字的降冪排序。
   m_intMaxYear = Mid(Trim(GRD2.TextMatrix(1, 3)), 1, 2)
   m_intCnt = m_intMaxYear \ 5
   GRD2.col = 3
   'Modified by Morgan 2023/11/15 依員工號排序(與獎金輸入一致)
   'grd2.Sort = 1 '一般升冪。執行估計文字不管是字串或者是數字的升冪排序。
   GRD2.Sort = 7 '字串昇冪。執行區分字串考慮大小寫比較的昇冪排序。
   'end 2023/11/15
   
   '預設值
   iLine = 1
   PrintTitle (0) '列印表頭
   iLst = 1 'Added by Morgan 2023/11/15
   For i = 1 To m_intCnt
      m_Year = i * 5
      For Index = 1 To 5
         strTemp(Index) = ""
      Next Index
      Index = 0
      intRow = 0
      intItem = 0
      'Modified by Morgan 2023/11/15
      'For j = 1 To GRD2.Rows - 1
      For j = iLst To GRD2.Rows - 1
         iLst = j
      'end 2023/11/15
         'Debug.Print Trim(GRD2.TextMatrix(j, 3)) & "  " & Val(Mid(Trim(GRD2.TextMatrix(j, 3)), 1, 2))
         If Val(Mid(Trim(GRD2.TextMatrix(j, 3)), 1, 2)) = m_Year Then
            intItem = intItem + 1
            Index = Index + 1
            intRow = intRow + 1
            If intRow = 1 Then
               If iLine > 50 Then
                  Printer.NewPage
                  iLine = 1
                  PrintTitle (0) '列印表頭
               Else
               End If
               PrintTitle (1) '列印欄位名稱
            End If
            strTemp(Index) = CStr(intItem) & ". " & Trim(GRD2.TextMatrix(j, 1))
            If Index = 5 Then '每五名新增一筆資料
               Call PrintData
               Index = 0
            End If
         'Modified by Morgan 2023/11/15
         'Else
         '   If IsNumeric(Trim(GRD2.TextMatrix(j, 3))) = True Then
         '      If Mid(Trim(GRD2.TextMatrix(j, 3)), 1, 2) > m_Year Then
         '         GoTo NextRec
         '      End If
         '   Else
         '      GoTo NextRec
         '   End If
         ElseIf Val(Mid(Trim(GRD2.TextMatrix(j, 3)), 1, 2)) > m_Year Then
            Exit For
            
         'end 2023/11/15
         End If
NextRec:
      Next j
      If Index <> 0 Then Call PrintData
   Next i
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintData()
   If iLine > 50 Or iLine = 1 Then
      'If .AbsolutePosition <> .RecordCount Then
         If iLine <> 1 Then Printer.NewPage
         iLine = 1
         PrintTitle (0) '列印表頭
      'End If
   End If
   
   PrintDetail '列印表中
   
   For m_i = 1 To 5
      strTemp(m_i) = ""
   Next m_i
End Sub

Sub PrintTitle(Index As Integer)
Dim strText As String, intYear As Integer

   GetPleft
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   If Index = 0 Then
      For m_i = 1 To m_intCnt
         intYear = m_i * 5
         If m_i = 1 Then
            strText = PUB_ChgNumber2Chinese(CStr(intYear)) & "年"
         ElseIf m_i <> m_intCnt Then
            strText = strText & "、" & PUB_ChgNumber2Chinese(CStr(intYear)) & "年"
         Else
            strText = strText & "及" & PUB_ChgNumber2Chinese(CStr(intYear)) & "年"
         End If
      Next m_i
      
      Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("服務滿" & strText & "同仁名單") / 2)
      Printer.CurrentY = iLine * 300
      Printer.Print "服務滿" & strText & "同仁名單"
      
      iLine = iLine + 2
      Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
      Printer.CurrentY = iLine * 300
      Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   Else
      iLine = iLine + 1
      Printer.CurrentX = 500
      Printer.CurrentY = iLine * 300
      Printer.Print "服務滿" & PUB_ChgNumber2Chinese(CStr(m_Year)) & "年名單："
      
      iLine = iLine + 2
      For m_i = 1 To 5
         Printer.CurrentX = PLeft(m_i) - Printer.TextWidth("姓　名")
         Printer.CurrentY = iLine * 300
         Printer.Print "姓　名"
      Next m_i
   '   Printer.CurrentX = PLeft(1) - Printer.TextWidth("姓　名")
   '   Printer.CurrentY = iLine * 300
   '   Printer.Print "姓　名"
   '   Printer.CurrentX = PLeft(2) - Printer.TextWidth("姓　名")
   '   Printer.CurrentY = iLine * 300
   '   Printer.Print "姓　名"
   '   Printer.CurrentX = PLeft(3) - Printer.TextWidth("姓　名")
   '   Printer.CurrentY = iLine * 300
   '   Printer.Print "姓　名"
   '   Printer.CurrentX = PLeft(4) - Printer.TextWidth("姓　名")
   '   Printer.CurrentY = iLine * 300
   '   Printer.Print "姓　名"
   '   Printer.CurrentX = PLeft(5) - Printer.TextWidth("姓　名")
   '   Printer.CurrentY = iLine * 300
   '   Printer.Print "姓　名"
      
      iLine = iLine + 1
      Printer.CurrentX = 500
      Printer.CurrentY = iLine * 300
      Printer.Print String(140, "-")
   End If
   
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 2200
   PLeft(2) = 4400
   PLeft(3) = 6600
   PLeft(4) = 8800
   PLeft(5) = 11000
End Sub

Sub PrintDetail()
Dim m_j As Integer

   For m_j = 1 To 5
      Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j))
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine = iLine + 1
End Sub

Sub StrMenu2()

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   'Printer.PaperSize = 9  'PDF
   'Modify By Sindy 2023/12/27 部門調整改抓ST93
   'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
   m_str = "SELECT a0922,ST01,ST02,sqldateT(ST32) " & _
                "From Staff,acc090NEW,SalaryData " & _
                "WHERE ST04='1' and ST01=SD01 and (sd02 not in('P','F') or sd02 is null) and not(substr(st01,5,1)>='A') " & _
                "AND ST93=a0921(+) " & m_StrSQL & _
                "Order By ST93,ST01 "
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       With m_rs
           m_rs.MoveFirst
           
           '預設值
           iLine = 1
           strType = "" '切頁條件
           
           Do While Not m_rs.EOF
               
               For m_i = 1 To 4
                   strTemp(m_i) = ""
               Next m_i
               
               strTemp(1) = CheckStr(m_rs.Fields(0))
               If strType = strTemp(1) Then
                  strTemp(1) = ""
               End If
               strTemp(2) = CheckStr(m_rs.Fields(1))
               strTemp(3) = CheckStr(m_rs.Fields(2))
               strTemp(4) = CheckStr(m_rs.Fields(3))
               
               If iLine > 50 Or iLine = 1 Then
                  'If .AbsolutePosition <> .RecordCount Then
                     If strType <> "" Then Printer.NewPage
                     iLine = 1
                     PrintTitle2 '列印表頭
                  'End If
               End If
               
               PrintDetail2 '列印表中
               
               strType = CheckStr(m_rs.Fields(0))
               m_rs.MoveNext
           Loop
       End With
   Else
      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintTitle2()
   GetPleft2
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("重簽保證書同仁名單") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "重簽保證書同仁名單"
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "上次簽保證書日期：" & ChangeTStringToTDateString(txt1(4)) & "--" & ChangeTStringToTDateString(txt1(5))
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "部　門"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "員工編號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print "上次簽保證書日"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft2()
   PLeft(1) = 1000
   PLeft(2) = 4500
   PLeft(3) = 6000
   PLeft(4) = 7500
End Sub

Sub PrintDetail2()
Dim m_j As Integer

   For m_j = 1 To 4
      Printer.CurrentX = PLeft(m_j)
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine = iLine + 1
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160305 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 3, 4, 5
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 1, 2
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
            If txt1(Index) <> "" Then
               Select Case txt1(Index)
               'modify by sonia 2016/3/24 +3.舊制退休金退休人員
               Case "1", "2", "3"
               Case Else
                   MsgBox "名單類別只可以輸入 1 ~ 3！", vbInformation, "輸入錯誤！"
                   Call txt1_GotFocus(Index)
                   Cancel = True
                   Exit Sub
               End Select
            End If
      Case 1, 2
         If Index = 1 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 2 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 3
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index)) = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
      Case 4, 5
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index)) = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
         If Index = 4 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 5 Then
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub

'add by sonia 2016/3/24 +3.舊制退休金退休人員
Sub StrMenu3()
Dim dblYear As Double

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   
   'modify by sonia 2017/8/1 +劉經理說再加72010簡美玉
   'Modify By Sindy 2023/12/27 部門調整改抓ST93
   'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
   m_str = "SELECT a0802,a0922,sd01,st02,sd16,substr(" & txt1(3) + 19110000 & ",1,4)-substr(st23,1,4)" & _
                " From Staff,SalaryData,acc080,acc090NEW" & _
                " WHERE ST04='1' and st01<'F' and (sd01='99029' or sd01='72010' or (sd01>='75001' and sd01<='98015')) and not(substr(st01,5,1)>='A')" & _
                  " and ST01=SD01(+) AND ST93=A0921(+) AND SD19=A0801(+)" & m_StrSQL & _
                "Order By sd19,ST01 "
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       With m_rs
           m_rs.MoveFirst
           
           '預設值
           iLine = 1
           strType = "" '切頁條件
           
           Do While Not m_rs.EOF
               
               For m_i = 1 To 10
                   strTemp(m_i) = ""
               Next m_i
               
               strTemp(1) = CheckStr(m_rs.Fields(0))     '公司
               strTemp(2) = CheckStr(m_rs.Fields(1))     '部門
               strTemp(3) = CheckStr(m_rs.Fields(2))     '員工編號
               strTemp(4) = CheckStr(m_rs.Fields(3))     '姓名
               strTemp(5) = CheckStr(m_rs.Fields(4))     '適用新制
               strTemp(7) = CheckStr(m_rs.Fields(5))     '年齡
               strTemp(6) = CalYear(CheckStr(m_rs.Fields(2)), DBDATE(txt1(3))) '取得年資
               
               If Val(strTemp(7)) >= 64 Then GoTo PrintData                        '年滿 64歲
               If strTemp(6) >= 15 And Val(strTemp(7)) >= 54 Then GoTo PrintData   '工作15年上年滿 54歲
               If strTemp(6) >= 24 Then GoTo PrintData                             '工作 24年以上
               If strTemp(6) >= 10 And Val(strTemp(7)) >= 59 Then GoTo PrintData   '工作10年以上年滿 59歲
                  
               GoTo Nextstep  '不符合上述條件不印
               
PrintData:
               '公司欄與前一筆相同印空白,不同時印區隔線
               If strType = strTemp(1) Then
                  strTemp(1) = ""
               ElseIf strType <> "" Then
                  Printer.CurrentX = 500
                  Printer.CurrentY = iLine * 300
                  Printer.Print String(140, "-")
                  iLine = iLine + 1
               End If
               
               If iLine > 50 Or iLine = 1 Then
                  If strType <> "" Then Printer.NewPage
                  iLine = 1
                  PrintTitle3 '列印表頭
               End If
               
               PrintDetail3 '列印表中
               
               strType = CheckStr(m_rs.Fields(0))
Nextstep:
               m_rs.MoveNext
           Loop
       End With
   Else
      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintTitle3()
   GetPleft3
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("舊制退休金退休人員名單") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "舊制退休金退休人員名單"
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "計算年資截止日期：" & ChangeTStringToTDateString(txt1(3))
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "公  司"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "部　門"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print "編號"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   Printer.CurrentX = PLeft(5) - 200
   Printer.CurrentY = iLine * 300
   Printer.Print "新制"
   Printer.CurrentX = PLeft(6) - 200
   Printer.CurrentY = iLine * 300
   Printer.Print "年資"
   Printer.CurrentX = PLeft(7) - 200
   Printer.CurrentY = iLine * 300
   Printer.Print "年齡"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft3()
   PLeft(1) = 500
   PLeft(2) = 4000
   PLeft(3) = 5600
   PLeft(4) = 6400
   PLeft(5) = 7900
   PLeft(6) = 8900
   PLeft(7) = 9900
End Sub

Sub PrintDetail3()
Dim m_j As Integer

   For m_j = 1 To 7
      Printer.CurrentX = PLeft(m_j)
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine = iLine + 1
End Sub
'end 2016/3/24
