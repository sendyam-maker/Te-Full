VERSION 5.00
Begin VB.Form frm160304 
   BorderStyle     =   1  '單線固定
   Caption         =   "歷年考績"
   ClientHeight    =   3200
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3200
   ScaleWidth      =   4960
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2790
      MaxLength       =   3
      TabIndex        =   1
      Top             =   840
      Width           =   555
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   9
      Top             =   2580
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   6
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
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   2940
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1560
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   2130
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1560
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2790
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1200
      Width           =   555
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2130
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1200
      Width           =   555
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4005
      TabIndex        =   8
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3060
      TabIndex        =   7
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   2130
      MaxLength       =   3
      TabIndex        =   0
      Top             =   840
      Width           =   555
   End
   Begin VB.Line Line3 
      X1              =   2640
      X2              =   3030
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Line Line2 
      X1              =   2850
      X2              =   3090
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line Line1 
      X1              =   2700
      X2              =   3090
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   1200
      TabIndex        =   13
      Top             =   1590
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門代號："
      Height          =   180
      Left            =   1200
      TabIndex        =   12
      Top             =   1230
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "年度："
      Height          =   180
      Left            =   1560
      TabIndex        =   11
      Top             =   870
      Width           =   540
   End
End
Attribute VB_Name = "frm160304"
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
Dim strType As String, strType2 As String, strType3 As String


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
        If Trim(txt1(0)) & Trim(txt1(1)) & Trim(txt1(2)) & _
            Trim(txt1(3)) & Trim(txt1(4)) & Trim(txt1(5)) = "" Then
            MsgBox "請輸入至少一項列印條件！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        If Trim(txt1(0)) & Trim(txt1(1)) <> "" Then
            If Trim(txt1(0)) = "" Then txt1(0).SetFocus: MsgBox "請輸入起始年度！", vbInformation, "操作錯誤！": Exit Sub
            If Trim(txt1(1)) = "" Then txt1(1).SetFocus: MsgBox "請輸入終止年度！", vbInformation, "操作錯誤！": Exit Sub
        End If
        If Trim(txt1(2)) & Trim(txt1(3)) <> "" Then
            If Trim(txt1(2)) = "" Then txt1(2).SetFocus: MsgBox "請輸入起始部門代號！", vbInformation, "操作錯誤！": Exit Sub
            If Trim(txt1(3)) = "" Then txt1(3).SetFocus: MsgBox "請輸入終止部門代號！", vbInformation, "操作錯誤！": Exit Sub
        End If
        If Trim(txt1(4)) & Trim(txt1(5)) <> "" Then
            If Trim(txt1(4)) = "" Then txt1(4).SetFocus: MsgBox "請輸入起始員工編號！", vbInformation, "操作錯誤！": Exit Sub
            If Trim(txt1(5)) = "" Then txt1(5).SetFocus: MsgBox "請輸入終止員工編號！", vbInformation, "操作錯誤！": Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            m_StrSQL = m_StrSQL & " and ym01 >='" & Val(txt1(0)) + 1911 & "' "
        End If
        If txt1(1) <> "" Then
            m_StrSQL = m_StrSQL & " and ym01 <='" & Val(txt1(1)) + 1911 & "' "
        End If
        If txt1(2) <> "" Then
            'Modify By Sindy 2023/12/27 部門調整改抓ST93
            m_StrSQL = m_StrSQL & " and st93 >='" & txt1(2) & "' "
        End If
        If txt1(3) <> "" Then
            'Modify By Sindy 2023/12/27 部門調整改抓ST93
            m_StrSQL = m_StrSQL & " and st93 <='" & txt1(3) & "' "
        End If
        If txt1(4) <> "" Then
            m_StrSQL = m_StrSQL & " and ym03 >='" & txt1(4) & "' "
        End If
        If txt1(5) <> "" Then
            m_StrSQL = m_StrSQL & " and ym03 <='" & txt1(5) & "' "
        End If
        StrMenu1
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
End Select
End Sub


Sub StrMenu1()

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   'Printer.PaperSize = 9  'PDF
   
   '甲等不列印
   'modify by sonia 2015/12/25 ym02 in ('1','3','4') 改 (ym02 is null or ym02<>'2')
   'Modify By Sindy 2023/12/27 部門調整改抓ST93
   'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
   m_str = "SELECT ym01,ym02,nvl(A0922,'(舊)'||A0902) a0902,st02 " & _
                "From YearMerit, Staff, acc090, acc090NEW " & _
                "WHERE ym03=st01(+) " & _
                "AND st03=a0901(+) AND st93=a0921(+) " & _
                "AND (ym02 is null or ym02<>'2') " & m_StrSQL & _
                "and not(substr(st01,5,1)>='A') " & _
                "Order By ym01,ym02,nvl(st93,st03),st01 "
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       With m_rs
           m_rs.MoveFirst
           
           '預設值
           iLine = 1
           strType = "" '切頁條件
           strType2 = ""
           strType3 = ""
           
           Do While Not m_rs.EOF
               
               For m_i = 1 To 4
                   strTemp(m_i) = ""
               Next m_i
               
               strTemp(1) = CheckStr(m_rs.Fields(0)) - 1911
               If CheckStr(m_rs.Fields(1)) = "1" Then
                  strTemp(2) = "優等"
               ElseIf CheckStr(m_rs.Fields(1)) = "3" Then
                  strTemp(2) = "乙等"
               ElseIf CheckStr(m_rs.Fields(1)) = "4" Then
                  strTemp(2) = "丙等"
               'add by sonia 2015/12/25
               ElseIf CheckStr(m_rs.Fields(1)) = "*" Then
                  strTemp(2) = "不參加考核"
               'end 2015/12/25
               Else
                  strTemp(2) = CheckStr(m_rs.Fields(1))
               End If
               strTemp(3) = CheckStr(m_rs.Fields(2))
               strTemp(4) = CheckStr(m_rs.Fields(3))
               
               If strType = CheckStr(m_rs.Fields(0)) Then
                  strTemp(1) = ""
                  If strType2 = CheckStr(m_rs.Fields(1)) Then
                     strTemp(2) = ""
                     If strType3 = CheckStr(m_rs.Fields(2)) Then
                        strTemp(3) = ""
                     End If
                  End If
               End If
               
               If iLine > 50 Or iLine = 1 Then
                  'If .AbsolutePosition <> .RecordCount Then
                     If strType <> "" Then Printer.NewPage
                     iLine = 1
                     PrintTitle '列印表頭
                  'End If
               End If
               
               PrintDetail '列印表中
               
               strType = CheckStr(m_rs.Fields(0))
               strType2 = CheckStr(m_rs.Fields(1))
               strType3 = CheckStr(m_rs.Fields(2))
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

Sub PrintTitle()
GetPleft

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("歷年考績表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "歷年考績表"

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "考績年度"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "考績等級"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "部　門"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "姓　名"

iLine = iLine + 1
Printer.CurrentX = 300
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 1500
PLeft(2) = 3000
PLeft(3) = 4500
PLeft(4) = 7500
End Sub

Sub PrintDetail()
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
   Set frm160304 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 2, 3, 4, 5
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "0101") = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
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
      Case 4, 5
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
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
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
