VERSION 5.00
Begin VB.Form frm160112 
   BorderStyle     =   1  '單線固定
   Caption         =   "各類代號資料"
   ClientHeight    =   3080
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3080
   ScaleWidth      =   4980
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   9
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
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2940
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3900
      TabIndex        =   5
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   1
      Left            =   1830
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1650
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2550
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1650
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   1830
      MaxLength       =   1
      TabIndex        =   0
      Top             =   960
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代　　號："
      Height          =   180
      Left            =   870
      TabIndex        =   8
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "明細表類別："
      Height          =   180
      Left            =   690
      TabIndex        =   7
      Top             =   990
      Width           =   1080
   End
   Begin VB.Label Label4 
      Caption         =   "(1.部門 2.職稱 3.職位 4.學歷           5.假別 6.異動原因 7.出生地           8.獎懲)"
      Height          =   600
      Left            =   2070
      TabIndex        =   6
      Top             =   990
      Width           =   2595
   End
   Begin VB.Line Line2 
      X1              =   2340
      X2              =   2700
      Y1              =   1770
      Y2              =   1770
   End
End
Attribute VB_Name = "frm160112"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
Option Explicit

Dim m_StrSQL As String
Dim m_str  As String
Dim m_rs As New ADODB.Recordset
Dim m_i As Integer
Dim PLeft(1 To 3) As Integer
Dim strTemp(1 To 3) As String
Dim iPgae As Integer, iLine As Integer


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
        If txt1(0) = "" Then
            MsgBox "明細表類別不可以空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        
        Set Printer = Printers(Combo1.ListIndex)
        Printer.EndDoc
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(1) <> "" Then
            If Val(txt1(0)) = 1 Then
                'Modify By Sindy 2023/12/28 部門調整改抓ST93
                'm_StrSQL = m_StrSQL & " and a0901>='" & txt1(1) & "' "
                m_StrSQL = m_StrSQL & " and a0921>='" & txt1(1) & "' "
            Else
                m_StrSQL = m_StrSQL & " and ac02>='" & txt1(1) & "' "
            End If
        End If
        If txt1(2) <> "" Then
            If Val(txt1(0)) = 1 Then
                'Modify By Sindy 2023/12/28 部門調整改抓ST93
                'm_StrSQL = m_StrSQL & " and a0901<='" & txt1(2) & "' "
                m_StrSQL = m_StrSQL & " and a0921<='" & txt1(2) & "' "
            Else
                m_StrSQL = m_StrSQL & " and ac02<='" & txt1(2) & "' "
            End If
        End If
        StrMenu
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
Case Else
End Select
End Sub


Sub StrMenu()
If Val(txt1(0)) = 1 Then
   'Modify By Sindy 2023/12/28 部門調整改抓ST93
   'm_str = "select a0901,a0902 from acc090 where 1=1 " & m_StrSQL & " order by a0901 "
   m_str = "select a0921,a0922,a0923 from acc090NEW where 1=1 " & m_StrSQL & " order by a0921 "
ElseIf Val(txt1(0)) = 8 Then
   m_str = "select ac02,ac03 from allcode where ac01='" & Format(Val(txt1(0)), "00") & "' " & m_StrSQL & " order by ac02 "
Else
   m_str = "select ac02,ac03 from allcode where ac01='" & Format(Val(txt1(0)) - 1, "00") & "' " & m_StrSQL & " order by ac02 "
End If
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    PrintData
Else
    ShowNoData
    Exit Sub
End If
End Sub

Sub PrintData()
Printer.Orientation = 1
With m_rs
    .MoveFirst
    PrintTitle
    Do While Not .EOF
        For m_i = 1 To 3 '2
            strTemp(m_i) = ""
        Next m_i
        strTemp(1) = CheckStr(.Fields(0))
        strTemp(2) = CheckStr(.Fields(1))
        'Add By Sindy 2023/12/28
        If txt1(0).Text = "1" Then '1.部門
            strTemp(3) = CheckStr(.Fields(2))
        End If
        '2023/12/28 END
        PrintDetail
        If iLine >= 52 Then
            If .AbsolutePosition <> .RecordCount Then
                Printer.NewPage
                PrintTitle
            End If
        End If
        .MoveNext
    Loop
End With
Printer.EndDoc
ShowPrintOk
End Sub

Sub GetPleft()
'Add By Sindy 2023/12/28
If txt1(0).Text = "1" Then '1.部門
   PLeft(1) = 2000
   PLeft(2) = 4000
   PLeft(3) = 6000
Else
'2023/12/28 END
   PLeft(1) = 3000
   PLeft(2) = 5000
End If
End Sub

Sub PrintTitle()
Dim oStr As String
Dim oStr1 As String
oStr = "代號資料明細表"
Select Case Val(txt1(0))
Case 1
        oStr1 = "部門"
Case 2
        oStr1 = "職稱"
Case 3
        oStr1 = "職位"
Case 4
        oStr1 = "學歷"
Case 5
        oStr1 = "假別"
Case 6
        oStr1 = "異動原因"
Case 7
        oStr1 = "出生地"
Case 8
        oStr1 = "獎懲"
End Select
GetPleft
Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(oStr1 & oStr) / 2)
Printer.CurrentY = 300
Printer.Print oStr1 & oStr
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print oStr1 & "代號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print oStr1 & "名稱"
'Add By Sindy 2023/12/28
If txt1(0).Text = "1" Then '1.部門
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print oStr1 & "全名"
End If
'2023/12/28 END
iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
Dim intCol As Integer

'Add By Sindy 2023/12/28
If txt1(0).Text = "1" Then '1.部門
   intCol = 3
Else
   intCol = 2
End If
'2023/12/28 END
For m_j = 1 To intCol '2
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
Set frm160112 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    InverseTextBox txt1(Index)
    CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
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
               Case "1", "2", "3", "4", "5", "6", "7", "8"
               Case Else
                   MsgBox "明細表類別只可以輸入 1 到 8！", vbInformation, "輸入錯誤！"
                   Call txt1_GotFocus(Index)
                   Cancel = True
                   Exit Sub
               End Select
           End If
   Case 2
           If txt1(Index) <> "" Then
               If RunNick(txt1(Index - 1), txt1(Index)) Then
                   Call txt1_GotFocus(Index)
                   Cancel = True
                   Exit Sub
               End If
           End If
   Case Else
   End Select
End Sub
