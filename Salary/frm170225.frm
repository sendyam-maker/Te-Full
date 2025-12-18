VERSION 5.00
Begin VB.Form frm170225 
   BorderStyle     =   1  '單線固定
   Caption         =   "端午,中秋代金入帳明細"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4770
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   2
      Left            =   2010
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1200
      Width           =   250
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   2010
      MaxLength       =   3
      TabIndex        =   0
      Top             =   870
      Width           =   675
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2010
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1530
      Width           =   250
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   60
      TabIndex        =   3
      Top             =   2370
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   4
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   5
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2610
      TabIndex        =   8
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3630
      TabIndex        =   9
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "資料類別：         (1.端午 2.中秋)  "
      Height          =   180
      Index           =   2
      Left            =   1080
      TabIndex        =   10
      Top             =   1260
      Width           =   2550
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "代金年度："
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   7
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "排序條件：         (1.公司別 2.入帳類別)"
      Height          =   180
      Left            =   1080
      TabIndex        =   6
      Top             =   1560
      Width           =   3000
   End
End
Attribute VB_Name = "frm170225"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2009/2/5 add by sonia
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If txt1(0) = "" Then
             MsgBox "代金年度不可以空白！", vbInformation, "操作錯誤！"
             txt1(0).SetFocus
             Exit Sub
         End If
         If txt1(0) <> "" Then
             If ChkDate(txt1(0) & "0101") = False Then
                 txt1(0).SetFocus
                 Exit Sub
             End If
         End If
         If txt1(2) = "" Then
             MsgBox "資料類別不可以空白！", vbInformation, "操作錯誤！"
             txt1(1).SetFocus
             Exit Sub
         End If
         If txt1(2) <> "" Then
             If txt1(1) <> "1" And txt1(1) <> "2" Then
                MsgBox "資料類別輸入錯誤！", vbInformation, "操作錯誤！"
                txt1(1).SetFocus
                Exit Sub
             End If
         End If
         If txt1(1) = "" Then
             MsgBox "排序條件不可以空白！", vbInformation, "操作錯誤！"
             txt1(1).SetFocus
             Exit Sub
         End If
         If txt1(1) <> "" Then
             If txt1(1) <> "1" And txt1(1) <> "2" Then
                MsgBox "排序條件輸入錯誤！", vbInformation, "操作錯誤！"
                txt1(1).SetFocus
                Exit Sub
             End If
         End If
         
         Screen.MousePointer = vbHourglass
         StrMenu
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Sub StrMenu()
Dim dblAmt As Double      '小計
Dim dblTotAmt As Double   '合計
Dim dblCnt As Double
Dim dblTotCnt As Double
Dim strYM As String

   strYM = Left(ChangeTStringToWString(txt1(0) & "0101"), 4)
   
   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   
   m_StrSQL = ""
   If txt1(1) = "1" Then
       m_StrSQL = m_StrSQL & " AND OB02='" & txt1(2) & "' Order by T3,SD01 "
   ElseIf txt1(1) = "2" Then
       m_StrSQL = m_StrSQL & " AND OB02='" & txt1(2) & "' Order by SD05,SD01 "
   End If
   '2009/9/30 MODIFY BY SONIA 加OB05>0  因為2009年96011於九月底離職
   'modify by sonia 2016/1/18 公司別改用ohbonus之ob12不再用sd19
   'm_str = "SELECT SD01,ST02,SD05,SD06,ob03 T1,nvl(ob05,0) T2,sd19 T3,ob03 T4,a0802 " & _
            "FROM Staff,SalaryData,ohbonus,acc080 " & _
            "WHERE substr(ob01,1,4)='" & strYM & "' and ob05>0 " & _
            "AND ob03=st01(+) AND ob03=SD01(+) AND sd19=a0801(+) " & m_StrSQL
   m_str = "SELECT SD01,ST02,SD05,SD06,ob03 T1,nvl(ob05,0) T2,ob12 T3,ob03 T4,a0802 " & _
            "FROM Staff,SalaryData,ohbonus,acc080 " & _
            "WHERE substr(ob01,1,4)='" & strYM & "' and ob05>0 " & _
            "AND ob03=st01(+) AND ob03=SD01(+) AND ob12=a0801(+) " & m_StrSQL
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         iLine = 1
         PrintTitle '列印表頭
         strType = "" '切頁條件
         dblAmt = 0
         dblTotAmt = 0
         dblCnt = 0
         dblTotCnt = 0
         Do While Not m_rs.EOF
             
            For m_i = 1 To 10
               strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields("SD01"))
            strTemp(2) = CheckStr(m_rs.Fields("ST02"))
            strTemp(3) = CheckStr(m_rs.Fields("SD05")) '入帳類別
            strTemp(4) = CheckStr(m_rs.Fields("SD06"))
            strTemp(5) = CheckStr(m_rs.Fields("T4"))
            strTemp(6) = CheckStr(m_rs.Fields("T2"))
            strTemp(7) = CheckStr(m_rs.Fields("a0802")) '公司別
            
            If strType <> "" Then
               If iLine > 50 Or _
                     (strType <> strTemp(7) And txt1(1) = "1") Or _
                     (strType <> strTemp(3) And txt1(1) = "2") Then
                     
                  If (strType <> strTemp(7) And txt1(1) = "1") Or _
                     (strType <> strTemp(3) And txt1(1) = "2") Then
                     
                     Printer.CurrentX = 500
                     Printer.CurrentY = iLine * 300
                     Printer.Print String(140, "-")
                     
                     iLine = iLine + 1
                     Printer.CurrentX = PLeft(4)
                     Printer.CurrentY = iLine * 300
                     Printer.Print "小計：(" & dblCnt & "人)"
                     Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmt, "##,##0"))
                     Printer.CurrentY = iLine * 300
                     Printer.Print Format(dblAmt, "##,##0")
                     
                     dblAmt = 0 '小計
                     dblCnt = 0
                  End If
                  
                  'If .AbsolutePosition <> .RecordCount Then
                      Printer.NewPage
                      iLine = 1
                      PrintTitle '列印表頭
                  'End If
               End If
            End If
            
            PrintDetail '列印表中
            
            If txt1(1) = "1" Then
               '公司別
               strType = strTemp(7)
            ElseIf txt1(1) = "2" Then
               '入帳類別
               strType = strTemp(3)
            End If
            
            dblAmt = dblAmt + strTemp(6)  '小計
            dblTotAmt = dblTotAmt + strTemp(6)  '合計
            dblCnt = dblCnt + 1
            dblTotCnt = dblTotCnt + 1
            m_rs.MoveNext
         Loop
          
         '列印表尾
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 300
         Printer.Print String(140, "-")
         
         iLine = iLine + 1
         Printer.CurrentX = PLeft(4)
         Printer.CurrentY = iLine * 300
         Printer.Print "小計：(" & dblCnt & "人)"
         Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmt, "##,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblAmt, "##,##0")
         
         iLine = iLine + 1
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 300
         Printer.Print String(140, "-")
         
         iLine = iLine + 1
         Printer.CurrentX = PLeft(4)
         Printer.CurrentY = iLine * 300
         Printer.Print "合計：(" & dblTotCnt & "人)"
         Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblTotAmt, "##,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt, "##,##0")
      End With
   Else
      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintTitle()
   GetPleft
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("端午代金入帳明細") / 2)
   Printer.CurrentY = iLine * 300
   If txt1(2) = "1" Then
      Printer.Print "端午代金入帳明細"
   Else
      Printer.Print "中秋代金入帳明細"
   End If
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("代金年度：" & txt1(0) & "  年") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "代金年度：" & txt1(0) & "  年"
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   If txt1(1) = "1" Then
      Printer.Print "公司別"
   ElseIf txt1(1) = "2" Then
      Printer.Print "入帳類別"
   End If
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "帳　　號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print "員工代號"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   Printer.CurrentX = PLeft(5) - Printer.TextWidth("金　額")
   Printer.CurrentY = iLine * 300
   Printer.Print "金　額"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 500
   PLeft(2) = 3500
   PLeft(3) = 6000
   PLeft(4) = 7500
   PLeft(5) = 10500
End Sub

Sub PrintDetail()
   '1.公司別/2.入帳類別
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   If txt1(1) = "1" Then
      If iLine = 7 Then
         Printer.Print strTemp(7)
      End If
   ElseIf txt1(1) = "2" Then
      If iLine = 7 Then
         If strTemp(3) = "1" Then
            Printer.Print "現金"
         ElseIf strTemp(3) = "2" Then
            Printer.Print "北所"
         ElseIf strTemp(3) = "3" Then
            Printer.Print "匯款"
         ElseIf strTemp(3) = "4" Then
            Printer.Print "中所"
         ElseIf strTemp(3) = "5" Then
            Printer.Print "南所"
         ElseIf strTemp(3) = "6" Then
            Printer.Print "高所"
         ElseIf strTemp(3) = "7" Then
            Printer.Print "其他"
         Else
            Printer.Print strTemp(3)
         End If
      End If
   End If
   '帳　　號
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(4)
   '員工代號
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(5)
   '姓　　名
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   '金　額
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(6), "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "##,##0")
   
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
   Set frm170225 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 3
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "0101") = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
