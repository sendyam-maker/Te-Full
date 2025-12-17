VERSION 5.00
Begin VB.Form frm160111 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工月卡名條"
   ClientHeight    =   2870
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2870
   ScaleWidth      =   4980
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   2070
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   10
      Top             =   2250
      Width           =   4875
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
         TabIndex        =   11
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3945
      TabIndex        =   7
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   2070
      MaxLength       =   3
      TabIndex        =   0
      Top             =   840
      Width           =   465
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2610
      MaxLength       =   3
      TabIndex        =   1
      Top             =   840
      Width           =   465
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   2
      Left            =   2070
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2760
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "月卡年月："
      Height          =   180
      Index           =   2
      Left            =   1140
      TabIndex        =   12
      Top             =   1590
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門代號："
      Height          =   180
      Left            =   1140
      TabIndex        =   9
      Top             =   870
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   8
      Top             =   1230
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2310
      X2              =   3300
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      X1              =   2370
      X2              =   2730
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "frm160111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
Option Explicit

Dim SeekPrint As Integer, SeekPrintL As Integer, i As Integer, j As Integer
Dim m_StrSQL As String
Dim m_str  As String
Dim m_rs As New ADODB.Recordset
Dim m_i As Integer
Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 7) As String
Dim iPgae As Integer, iLine As Integer


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
        If txt1(1) <> "" Then
            If txt1(0) = "" Then
                MsgBox "部門代號區間請輸入完整！", vbInformation, "操作錯誤！"
                txt1(0).SetFocus
                Exit Sub
            Else
                If RunNick(txt1(0), txt1(1)) Then
                    txt1(0).SetFocus
                    Exit Sub
                End If
            End If
        End If
        If txt1(3) <> "" Then
            If txt1(2) = "" Then
                MsgBox "員工編號區間請輸入完整！", vbInformation, "操作錯誤！"
                txt1(2).SetFocus
                Exit Sub
            Else
                If RunNick(txt1(2), txt1(3)) Then
                    txt1(2).SetFocus
                    Exit Sub
                End If
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            m_StrSQL = m_StrSQL & " and st03>='" & txt1(0) & "' "
        End If
        If txt1(1) <> "" Then
            m_StrSQL = m_StrSQL & " and st03<='" & txt1(1) & "' "
        End If
        If txt1(2) <> "" Then
            m_StrSQL = m_StrSQL & " and st01>='" & txt1(2) & "' "
        End If
        If txt1(3) <> "" Then
            m_StrSQL = m_StrSQL & " and st01<='" & txt1(3) & "' "
        End If
        m_StrSQL = m_StrSQL & " and st04='1' "
        StrMenu
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
Case Else
End Select
End Sub

Sub StrMenu()
Dim RowHeight As Integer

'設定印表機
Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1 '1.直印 2.橫印
Printer.PaperSize = PUB_GetPaperSize(12)  '打卡卡片名條
'Modify By Sindy 2011/6/3 台一開發林宜鋒是不用打卡 , 請取消卡片名條列印
'Modify By Sindy 2016/10/11 L01.法務處律師只有3個人不打卡,其他都要(and st03 not in ('R04','L01')取消L01)
m_str = "select st01,st02,a0902,st22 " & _
             "from staff,acc090,SalaryData " & _
             "where ST01=SD01 " & _
             "and ((SD02 not in('P','F') or SD02 is null) or ST01='68007') " & _
             "and ST01 not in('A0018','" & Replace(Pub_GetSpecMan("不用打卡的律師"), ";", "','") & "') " & _
             "and st03 not in ('R04') and st06='1' " & _
             "and st03=a0901(+) " & m_StrSQL & " order by st03,st01 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        .MoveFirst
        
        iLine = 1
        RowHeight = 950
        
        Do While Not .EOF
            For m_i = 1 To 5
               strTemp(m_i) = ""
            Next m_i
            
            '每一列的左編號
            strTemp(1) = CheckStr(m_rs.Fields("st01"))
            
            'Modify By Sindy 98/03/24
            If strTemp(1) = "96009" Then
               strTemp(2) = "STEWART"
            '98/03/24 End
            '2010/7/26 ADD BY SONIA
            ElseIf strTemp(1) = "99029" Then
               strTemp(2) = "IAIN"
            '2010/7/26 End
            Else
               strTemp(2) = CheckStr(m_rs.Fields("st02"))
            End If
            
            If Len(Trim(txt1(4))) = 4 Then strTemp(3) = Left(Trim(txt1(4)), 2) & " 年 " & Right(Trim(txt1(4)), 2) & " 月"
            If Len(Trim(txt1(4))) = 5 Then strTemp(3) = Left(Trim(txt1(4)), 3) & " 年 " & Right(Trim(txt1(4)), 2) & " 月"
            
            '每一列的右編號
            strTemp(4) = CheckStr(m_rs.Fields("st01"))
            
            'Modify By Sindy 98/03/24
            If strTemp(4) = "96009" Then
               strTemp(5) = "STEWART"
            '98/03/24 End
            '2010/7/26 ADD BY SONIA
            ElseIf strTemp(4) = "99029" Then
               strTemp(5) = "IAIN"
            '2010/7/26 End
            Else
               strTemp(5) = CheckStr(m_rs.Fields("st02"))
            End If
            
            If iLine > 15 Then
               Printer.NewPage
               iLine = 1
            End If
            
            '開始列印
            '員工代號及姓名
            Printer.Font.Size = 20
            Printer.CurrentX = 300
            Printer.CurrentY = iLine * RowHeight - RowHeight
            Printer.Print strTemp(1)
            Printer.CurrentX = 1600 '1800 姓名1
            Printer.CurrentY = iLine * RowHeight - RowHeight
            Printer.Print strTemp(2)
            Printer.CurrentX = 5100 '5300
            Printer.CurrentY = iLine * RowHeight - RowHeight
            Printer.Print strTemp(4)
            Printer.CurrentX = 6400 '6600 '6800 姓名2
            Printer.CurrentY = iLine * RowHeight - RowHeight
            Printer.Print strTemp(5)
            '月卡年月
            Printer.Font.Size = 10
            Printer.CurrentX = 3500
            Printer.CurrentY = iLine * RowHeight - 750
            Printer.Print strTemp(3)
            Printer.CurrentX = 8300 '8500
            Printer.CurrentY = iLine * RowHeight - 750
            Printer.Print strTemp(3)
            
            iLine = iLine + 1
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

Private Sub Form_Load()
MoveFormToCenter Me
strSql = Printer.DeviceName
SeekPrintL = Printer.Orientation
j = 0
For i = 0 To Printers.Count - 1
    Set Printer = Printers(i)
    Combo1.AddItem Printer.DeviceName, j
    j = j + 1
    If Printer.DeviceName = strSql Then
        SeekPrint = i
    End If
Next i
Set Printer = Printers(SeekPrint)
Combo1.Text = Combo1.List(0)

'預設月卡年月
txt1(4).Text = Left(Trim(ChangeWStringToTString(strSrvDate(1))), Len(Trim(ChangeWStringToTString(strSrvDate(1)))) - 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
Set frm160111 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    InverseTextBox txt1(Index)
    CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 4
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 0, 1, 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 2, 3
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 3 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
