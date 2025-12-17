VERSION 5.00
Begin VB.Form frm180501 
   BorderStyle     =   1  '單線固定
   Caption         =   "職務代理人資料表"
   ClientHeight    =   3170
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   5390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3170
   ScaleWidth      =   5390
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3480
      TabIndex        =   10
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4425
      TabIndex        =   9
      Top             =   60
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   7
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   8
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   2130
      MaxLength       =   3
      TabIndex        =   0
      Top             =   810
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   1
      Top             =   810
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   2130
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1140
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   2970
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1140
      Width           =   705
   End
   Begin VB.Line Line2 
      X1              =   2550
      X2              =   3030
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "部  門  別："
      Height          =   180
      Left            =   1200
      TabIndex        =   5
      Top             =   870
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   1170
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2850
      X2              =   3090
      Y1              =   1260
      Y2              =   1260
   End
End
Attribute VB_Name = "frm180501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Create by SINDY 2011/9/26
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 20) As Integer
Dim strTemp(1 To 20) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(1) <> "" Then
            'Modify By Sindy 2023/12/27 部門調整改抓ST93
            'm_StrSQL = m_StrSQL & "AND s1.ST03>='" & txt1(1) & "' "
            m_StrSQL = m_StrSQL & "AND s1.ST93>='" & txt1(1) & "' "
        End If
        If txt1(2) <> "" Then
            'Modify By Sindy 2023/12/27 部門調整改抓ST93
            'm_StrSQL = m_StrSQL & "AND s1.ST03<='" & txt1(2) & "' "
            m_StrSQL = m_StrSQL & "AND s1.ST93<='" & txt1(2) & "' "
        End If
        If txt1(3) <> "" Then
            m_StrSQL = m_StrSQL & "AND s1.ST01>='" & txt1(3) & "' "
        End If
        If txt1(4) <> "" Then
            m_StrSQL = m_StrSQL & "AND s1.ST01<='" & txt1(4) & "' "
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
Printer.Orientation = 2 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF
'Modified by Lydia 2017/03/28 ST14改成多個編號 s1.st14<>'99997'=> instr(s1.st14,'99997')=0
'Modify By Sindy 2023/12/27 部門調整改抓ST93
m_str = "select A0922,s1.ST02,s2.ST02,s3.ST02,s4.ST02,s5.ST02,s6.ST02,s7.ST02 " & _
        ",s17.ST02||decode(nvl(B0116,''),'','','('||B0116||')'),s19.ST02||decode(nvl(B0118,''),'','','('||B0118||')') " & _
        ",s21.ST02||decode(nvl(B0120,''),'','','('||B0120||')'),s23.ST02||decode(nvl(B0122,''),'','','('||B0122||')') " & _
        ",s8.ST02||decode(nvl(B0112,''),'','','('||B0112||')'),s9.ST02||decode(nvl(B0113,''),'','','('||B0113||')') " & _
        ",s10.ST02||decode(nvl(B0114,''),'','','('||B0114||')'),s11.ST02||decode(nvl(B0115,''),'','','('||B0115||')') " & _
        "from ABS001,ACC090NEW,Staff s1,Staff s2,Staff s3,Staff s4,Staff s5,Staff s6,Staff s7 " & _
        ",Staff s8,Staff s9,Staff s10,Staff s11 " & _
        ",Staff s17,Staff s19,Staff s21,Staff s23 " & _
        "where s1.ST01=B0101(+) and s1.ST93=A0921(+) " & _
        "and s1.ST04='1' and (instr(s1.st14,'99997')=0 or s1.ST14 is null) and substr(s1.ST01,1,1) in(" & ST01CodeNum1 & ") and substr(s1.ST01,4,1)<>'9' and s1.st01 not in('60000','96029','96030','86026','67004','68007','63001') " & _
        "and B0102=s2.ST01(+) and B0103=s3.ST01(+) and B0104=s4.ST01(+) and B0105=s5.ST01(+) and B0106=s6.ST01(+) and B0107=s7.ST01(+) " & _
        "and B0108=s8.ST01(+) and B0109=s9.ST01(+) and B0110=s10.ST01(+) and B0111=s11.ST01(+) " & _
        "and B0117=s17.ST01(+) and B0119=s19.ST01(+) and B0121=s21.ST01(+) and B0123=s23.ST01(+) " & m_StrSQL & _
        "order by s1.ST93,B0101 "
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
         For m_i = 1 To 16
             strTemp(m_i) = ""
         Next m_i
         strTemp(1) = CheckStr(m_rs.Fields(0)) 'Left(CheckStr(m_rs.Fields(0)), 5)
         strTemp(2) = CheckStr(m_rs.Fields(1))
         strTemp(3) = CheckStr(m_rs.Fields(2))
         strTemp(4) = CheckStr(m_rs.Fields(3))
         strTemp(5) = CheckStr(m_rs.Fields(4))
         strTemp(6) = CheckStr(m_rs.Fields(5))
         strTemp(7) = CheckStr(m_rs.Fields(6))
         strTemp(8) = CheckStr(m_rs.Fields(7))
         strTemp(9) = CheckStr(m_rs.Fields(8))
         strTemp(10) = CheckStr(m_rs.Fields(9))
         strTemp(11) = CheckStr(m_rs.Fields(10))
         strTemp(12) = CheckStr(m_rs.Fields(11))
         strTemp(13) = CheckStr(m_rs.Fields(12))
         strTemp(14) = CheckStr(m_rs.Fields(13))
         strTemp(15) = CheckStr(m_rs.Fields(14))
         strTemp(16) = CheckStr(m_rs.Fields(15))
         
         If iLine > 36 Or iLine = 1 Then
            If strType <> "" Then Printer.NewPage
            iLine = 1
            PrintTitle '列印表頭
         End If
         
         PrintDetail '列印明細
         
         strType = CheckStr(m_rs.Fields(0))
         m_rs.MoveNext
      Loop
   End With
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "註：職代(1-1)(1-2)：第一組職代　職代(2-1)(2-2)：第二組職代　職代(3-1)(3-2)：第三組職代"
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "　　主管括號內的數字是指幾天(不含)以上須經過該主管審核"
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "　　案職代(1-1)(1-2)：案件第一組職代　案職代(2-1)(2-2)：案件第二組職代　＜註：(1)台灣案(2)非台灣案＞"
Else
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 14
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("職務代理人資料表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "職務代理人資料表"

Printer.Font.Size = 10
iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "部門別"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "員工姓名"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "職代(1-1)"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "職代(1-2)"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "職代(2-1)"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "職代(2-2)"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iLine * 300
Printer.Print "職代(3-1)"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iLine * 300
Printer.Print "職代(3-2)"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iLine * 300
Printer.Print "案職代(1-1)"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iLine * 300
Printer.Print "案職代(1-2)"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iLine * 300
Printer.Print "案職代(2-1)"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iLine * 300
Printer.Print "案職代(2-2)"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iLine * 300
Printer.Print "主管1"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iLine * 300
Printer.Print "主管2"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iLine * 300
Printer.Print "主管3"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iLine * 300
Printer.Print "主管4"

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(255, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 300
PLeft(2) = 1500
PLeft(3) = 2500
PLeft(4) = 3500
PLeft(5) = 4500
PLeft(6) = 5500
PLeft(7) = 6500
PLeft(8) = 7500
PLeft(9) = 8500
PLeft(10) = 9500
PLeft(11) = 10500
PLeft(12) = 11500
PLeft(13) = 12500
PLeft(14) = 13500
PLeft(15) = 14500
PLeft(16) = 15500
End Sub

Sub PrintDetail()
Dim i As Integer
   
   For i = 1 To 16
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(i)
   Next i
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
   Set frm180501 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
'      Case 0
'         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 1, 2, 3, 4
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
'      Case 0
'         If txt1(index) <> "" Then
'            If ChkDate(txt1(index) & "01") = False Then
'                Call txt1_GotFocus(index)
'                Cancel = True
'                Exit Sub
'            End If
'         End If
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
      Case 3, 4
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 3 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 4 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
