VERSION 5.00
Begin VB.Form frm160108 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人請假紀錄列印"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4980
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2970
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3915
      TabIndex        =   5
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   1950
      MaxLength       =   6
      TabIndex        =   0
      Top             =   930
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   1
      Left            =   1950
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1290
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2790
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1290
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   6
      Top             =   2400
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   1020
      TabIndex        =   9
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "請假日期："
      Height          =   180
      Left            =   1020
      TabIndex        =   8
      Top             =   1320
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2580
      X2              =   2940
      Y1              =   1410
      Y2              =   1410
   End
End
Attribute VB_Name = "frm160108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
'Create by Sindy 2009/01/13
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
Dim dblAmt As Double, dblAmt2 As Double


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
            m_StrSQL = m_StrSQL & " and sa01='" & txt1(0) & "' "
        End If
        If txt1(1) <> "" Then
            m_StrSQL = m_StrSQL & " and sa02>='" & DBDATE(txt1(1)) & "' "
        End If
        If txt1(2) <> "" Then
            m_StrSQL = m_StrSQL & " and sa02<='" & DBDATE(txt1(2)) & "' "
        End If
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
Printer.Orientation = 1 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

m_str = "select s1.ST01,s1.ST02,sa06||' '||ac03, " & _
             "sqldatet(sa02)||' '||substr(rtrim(ltrim(to_char(sa03,'0000'))),1,2)||':'||substr(rtrim(ltrim(to_char(sa03,'0000'))),3,2)||'--'||sqldatet(sa04)||' '||substr(rtrim(ltrim(to_char(sa05,'0000'))),1,2)||':'||substr(rtrim(ltrim(to_char(sa05,'0000'))),3,2), " & _
             "sa07,sa08,SA09 " & _
             "from staff s1,staff_absence,allcode " & _
             "where sa01=s1.st01(+) " & _
             "and '04'=ac01(+) and sa06=ac02(+) " & m_StrSQL & _
             "order by sa01,sa02,sa03 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        .MoveFirst
        
        iLine = 1
        strType = ""
        i = 0
        dblAmt = 0
        dblAmt2 = 0
        
        Do While Not .EOF
            For m_i = 1 To 6
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
            
            If iLine > 48 Or iLine = 1 Or _
               (strType <> "" And strType <> CheckStr(m_rs.Fields(0))) Then
               
               If strType <> "" And strType <> CheckStr(m_rs.Fields(0)) Then
                  PrintEnd '小計
               End If
               
               'If .AbsolutePosition <> .RecordCount Then
                   If strType <> "" Then Printer.NewPage
                   iLine = 1
                   PrintTitle '列印表頭
               'End If
            End If
            PrintDetail
            
            dblAmt = dblAmt + CDbl(strTemp(4))
            dblAmt2 = dblAmt2 + CDbl(strTemp(5))
            strType = CheckStr(m_rs.Fields(0))
            .MoveNext
        Loop
        PrintEnd '小計
    End With
Else
    ShowNoData
    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
   Printer.CurrentX = 6500
   Printer.CurrentY = iLine * 300
   Printer.Print "小　計："
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(dblAmt)
   Printer.CurrentY = iLine * 300
   Printer.Print dblAmt
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(dblAmt2)
   Printer.CurrentY = iLine * 300
   Printer.Print dblAmt2
   
   iLine = iLine + 1
   dblAmt = 0
   dblAmt2 = 0
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 1500
PLeft(3) = 3500
PLeft(4) = 8500
PLeft(5) = 9500
PLeft(6) = 10000
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("個人請假資料表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "個人請假資料表"
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
Printer.Print "假　別"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "時　間　起　迄"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("日")
Printer.CurrentY = iLine * 300
Printer.Print "日"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("時")
Printer.CurrentY = iLine * 300
Printer.Print "時"
'Printer.CurrentX = PLeft(6)
'Printer.CurrentY = iLine * 300
'Printer.Print "職務代理人"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 5 '6
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
Set frm160108 = Nothing
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
