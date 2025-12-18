VERSION 5.00
Begin VB.Form frm170212 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工貸款償還明細"
   ClientHeight    =   2424
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4764
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2424
   ScaleWidth      =   4764
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   5
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   6
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2580
      TabIndex        =   2
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3660
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   2010
      MaxLength       =   5
      TabIndex        =   0
      Top             =   870
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "償還年月："
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   930
      Width           =   900
   End
End
Attribute VB_Name = "frm170212"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2009/2/6 add by sonia
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblTotAmt As Double  '合計
Dim dblTotCnt As Double

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If txt1(0) = "" Then
            MsgBox "償還年月不可以空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
         End If
         If txt1(0) <> "" Then
            If Len(txt1(0)) <= 3 Then
               MsgBox "償還年月輸入錯誤！", vbInformation, "操作錯誤！"
               txt1(0).SetFocus
               Exit Sub
            End If
            If ChkDate(txt1(0) & "01") = False Then
               txt1(0).SetFocus
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
Dim strYM As String

   strYM = Val(txt1(0)) + 191100
   
   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   m_str = "SELECT substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) 員工編號,ST02 姓名,SUM(NVL(SM19,0)) 償還金額 " & _
         "FROM SALARYMONTH,STAFF WHERE substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=ST01(+) and nvl(sm19,0)>0  " & _
         "AND SM02= " & strYM
   m_str = m_str & " GROUP BY substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4),ST02 order by 員工編號"
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         iLine = 1
         strType = "" '切頁條件
         dblTotAmt = 0: dblTotCnt = 0
         
         Do While Not m_rs.EOF
             
            For m_i = 1 To 10
               strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields(0))  '員工編號
            strTemp(2) = CheckStr(m_rs.Fields(1))  '姓名
            strTemp(3) = CheckStr(m_rs.Fields(2))  '償還金額
            
            If iLine > 50 Or iLine = 1 Then
                     
               If strType <> "" Then Printer.NewPage
               iLine = 1
               PrintTitle '列印表頭
            End If
            
            PrintDetail '列印表中
            
            'strType = strTemp(1) '暫不跳頁
            
            dblTotAmt = dblTotAmt + strTemp(3)  '合計
            dblTotCnt = dblTotCnt + 1
            m_rs.MoveNext
         Loop
          
         '列印表尾
         
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 300
         Printer.Print String(140, "-")
         
         iLine = iLine + 1
         Printer.CurrentX = PLeft(1) - Printer.TextWidth(dblTotCnt & " 人 ")
         Printer.CurrentY = iLine * 300
         Printer.Print dblTotCnt & " 人 "
         Printer.CurrentX = PLeft(2) - Printer.TextWidth("合　計")
         Printer.CurrentY = iLine * 300
         Printer.Print "合　計："
         Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblTotAmt, "#,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt, "#,###,##0")
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
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("員工貸款償還明細") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "員工貸款償還明細"
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("000 年 00 月") / 2)
   Printer.CurrentY = iLine * 300
   If Len(txt1(0)) = 5 Then
      Printer.Print Left(Trim(txt1(0)), 3) & "  年  " & Right(Trim(txt1(0)), 2) & "  月"
   Else
      Printer.Print Left(Trim(txt1(0)), 2) & "  年  " & Right(Trim(txt1(0)), 2) & "  月"
   End If
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = PLeft(1) - Printer.TextWidth("員工編號")
   Printer.CurrentY = iLine * 300
   Printer.Print "員工編號"
   Printer.CurrentX = PLeft(2) - Printer.TextWidth("姓　名")
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("償還金額")
   Printer.CurrentY = iLine * 300
   Printer.Print "償還金額"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 4000
   PLeft(2) = 5500
   PLeft(3) = 7700
End Sub

Sub PrintDetail()
   Printer.CurrentX = PLeft(1) - 700
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2) - 700
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(3), "#,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(3), "#,###,##0")
   
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
   Set frm170212 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "01") = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub


