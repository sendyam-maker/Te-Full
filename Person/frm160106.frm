VERSION 5.00
Begin VB.Form frm160106 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人出差紀錄列印"
   ClientHeight    =   3060
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4950
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   1740
      TabIndex        =   10
      Top             =   1530
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3000
      TabIndex        =   4
      Top             =   30
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3945
      TabIndex        =   5
      Top             =   30
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   0
      Top             =   810
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   1
      Left            =   1740
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1170
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1170
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   2430
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   3
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   7
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(可依想查的輸入3,4)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   1740
      TabIndex        =   13
      Top             =   1860
      Width           =   1605
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "(1:長程 2:短程 3:大陸 4:國外)"
      Height          =   180
      Left            =   2640
      TabIndex        =   12
      Top             =   1590
      Width           =   2235
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "差程："
      Height          =   180
      Left            =   1170
      TabIndex        =   11
      Top             =   1590
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   810
      TabIndex        =   9
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "出差日期："
      Height          =   180
      Left            =   810
      TabIndex        =   8
      Top             =   1200
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2370
      X2              =   2730
      Y1              =   1290
      Y2              =   1290
   End
End
Attribute VB_Name = "frm160106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
'Create by Sindy 2009/01/12
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_str2 As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
        If txt1(0) = "" And txt1(1) = "" And txt1(2) = "" Then
            MsgBox "請至少輸入一項列印條件！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        If (txt1(1) = "" And txt1(2) <> "") Then
            MsgBox "起始日期不可空白！", vbInformation, "操作錯誤！"
            txt1(1).SetFocus
            Exit Sub
        End If
        If (txt1(1) <> "" And txt1(2) = "") Then
            MsgBox "迄止日期不可空白！", vbInformation, "操作錯誤！"
            txt1(2).SetFocus
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            m_StrSQL = m_StrSQL & " and sb01='" & txt1(0) & "' "
        End If
        If txt1(1) <> "" Then
            m_StrSQL = m_StrSQL & " and sb02 >= '" & ChangeTStringToWString(txt1(1)) & "' "
        End If
        If txt1(2) <> "" Then
            m_StrSQL = m_StrSQL & " and sb02 <= '" & ChangeTStringToWString(txt1(2)) & "' "
        End If
        'Add By Sindy 2017/6/23
        If txt1(3) <> "" Then
            m_StrSQL = m_StrSQL & " and sb08 in(" & txt1(3) & ") "
        End If
        '2017/6/23 END
        StrMenu1
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
Case Else
End Select
End Sub


Sub StrMenu1()
Dim i As Integer

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF
'Modify By Sindy 2023/12/28 部門調整改抓ST93
m_str = "select sb01,s1.st02,decode(sb08,'1','長程','2','短程','3','大陸','4','國外',''),sqldatet(sb02)||' '||substr(rtrim(ltrim(to_char(sb03,'0000'))),1,2)||':'||substr(rtrim(ltrim(to_char(sb03,'0000'))),3,2)||'--'||sqldatet(sb04)||' '||substr(rtrim(ltrim(to_char(sb05,'0000'))),1,2)||':'||substr(rtrim(ltrim(to_char(sb05,'0000'))),3,2),sb06,sb07,sb09,sb10 " & _
             "from staff_busi_trip,staff s1 " & _
             "where sb01=s1.st01(+) " & m_StrSQL & _
             "order by s1.st93,sb01,sb02 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        .MoveFirst
        
        iLine = 1
        strType = ""
        i = 0
        
        Do While Not .EOF
            For m_i = 1 To 7
                strTemp(m_i) = ""
            Next m_i
            
            '流水號
            i = i + 1
            strTemp(1) = Right("0000" & CStr(i), 4)
            strTemp(2) = CheckStr(m_rs.Fields(2))
            strTemp(3) = CheckStr(m_rs.Fields(3))
            strTemp(4) = CheckStr(m_rs.Fields(4))
            strTemp(5) = CheckStr(m_rs.Fields(5))
            strTemp(6) = CheckStr(m_rs.Fields(6))
'            strTemp(7) = CheckStr(m_rs.Fields(7))
            
            If iLine > 34 Or iLine = 1 Or _
               (strType <> "" And strType <> CheckStr(m_rs.Fields(0))) Then
               'If .AbsolutePosition <> .RecordCount Then
                   If strType <> "" Then Printer.NewPage
                   iLine = 1
                   PrintTitle '列印表頭
               'End If
            End If
            PrintDetail
            
            strType = CheckStr(m_rs.Fields(0))
            .MoveNext
        Loop
    End With
Else
    ShowNoData
    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 1500
PLeft(3) = 2500
PLeft(4) = 7500
PLeft(5) = 8500
PLeft(6) = 9000
'PLeft(7) = 11500
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("個人出差資料表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "個人出差資料表"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "員工姓名：" & CheckStr(m_rs.Fields(0)) & "　" & CheckStr(m_rs.Fields(1))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "流水號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "差程"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "時間起迄"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("日")
Printer.CurrentY = iLine * 300
Printer.Print "日"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("時")
Printer.CurrentY = iLine * 300
Printer.Print "時"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "地點"
'Printer.CurrentX = PLeft(7)
'Printer.CurrentY = iLine * 300
'Printer.Print "職務代理人"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(205, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 6 '7
   If m_j = 4 Or m_j = 5 Then
      Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j))
   Else
      Printer.CurrentX = PLeft(m_j)
   End If
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
Set frm160106 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    InverseTextBox txt1(Index)
    CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 1, 2
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 0
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 1, 2
         If txt1(Index).Text <> "" Then
            If ChkDate(txt1(Index)) = False Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 1 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 2 Then
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
