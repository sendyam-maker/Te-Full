VERSION 5.00
Begin VB.Form frm180502 
   BorderStyle     =   1  '單線固定
   Caption         =   "每日假單簽收明細表"
   ClientHeight    =   3270
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   6050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6050
   Begin VB.TextBox txtST06 
      Height          =   300
      Index           =   0
      Left            =   1710
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1740
      Width           =   495
   End
   Begin VB.TextBox txtST06 
      Height          =   300
      Index           =   1
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1740
      Width           =   495
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Index           =   0
      Left            =   1710
      MaxLength       =   7
      TabIndex        =   0
      Top             =   750
      Width           =   945
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Index           =   1
      Left            =   2790
      MaxLength       =   7
      TabIndex        =   1
      Top             =   750
      Width           =   945
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   4140
      TabIndex        =   8
      Top             =   30
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   5085
      TabIndex        =   9
      Top             =   30
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   60
      TabIndex        =   12
      Top             =   2640
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   13
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   14
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1710
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   2340
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1410
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   2550
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1410
      Width           =   705
   End
   Begin VB.Line Line3 
      X1              =   2130
      X2              =   2610
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "所　　別："
      Height          =   180
      Left            =   780
      TabIndex        =   17
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "(1.北所 2.中所 3.南所 4.高所 5.其他)"
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   2880
      TabIndex        =   16
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "簽收日期："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   780
      TabIndex        =   15
      Top             =   810
      Width           =   900
   End
   Begin VB.Line Line4 
      X1              =   2550
      X2              =   3090
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line Line2 
      X1              =   2130
      X2              =   2610
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "部  門  別："
      Height          =   180
      Left            =   780
      TabIndex        =   11
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   0
      Left            =   780
      TabIndex        =   10
      Top             =   1440
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2430
      X2              =   2670
      Y1              =   1530
      Y2              =   1530
   End
End
Attribute VB_Name = "frm180502"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Create by SINDY 2011/9/26
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL_SA As String, m_StrSQL_SB As String, m_StrSQL_So As String
Dim m_i As Integer
Dim PLeft(1 To 20) As Integer
Dim strTemp(1 To 20) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
        Screen.MousePointer = vbHourglass
        m_StrSQL_SA = "": m_StrSQL_SB = "": m_StrSQL_So = ""
        '簽收日期
        If txtDate(0) <> "" Then
            m_StrSQL_SA = m_StrSQL_SA & "AND SA11>=" & DBDATE(txtDate(0)) & " "
            m_StrSQL_SB = m_StrSQL_SB & "AND SB12>=" & DBDATE(txtDate(0)) & " "
            m_StrSQL_So = m_StrSQL_So & "AND So08>=" & DBDATE(txtDate(0)) & " "
        End If
        If txtDate(1) <> "" Then
            m_StrSQL_SA = m_StrSQL_SA & "AND SA11<=" & DBDATE(txtDate(1)) & " "
            m_StrSQL_SB = m_StrSQL_SB & "AND SB12<=" & DBDATE(txtDate(1)) & " "
            m_StrSQL_So = m_StrSQL_So & "AND So08<=" & DBDATE(txtDate(1)) & " "
        End If
        '部門別
        If txt1(1) <> "" Then
            'Modify By Sindy 2023/12/27 部門調整改抓ST93
            m_StrSQL_SA = m_StrSQL_SA & "AND ST93>='" & DBDATE(txt1(1)) & "' "
            m_StrSQL_SB = m_StrSQL_SB & "AND ST93>='" & DBDATE(txt1(1)) & "' "
            m_StrSQL_So = m_StrSQL_So & "AND ST93>='" & DBDATE(txt1(1)) & "' "
        End If
        If txt1(2) <> "" Then
            'Modify By Sindy 2023/12/27 部門調整改抓ST93
            m_StrSQL_SA = m_StrSQL_SA & "AND ST93<='" & DBDATE(txt1(2)) & "' "
            m_StrSQL_SB = m_StrSQL_SB & "AND ST93<='" & DBDATE(txt1(2)) & "' "
            m_StrSQL_So = m_StrSQL_So & "AND ST93<='" & DBDATE(txt1(2)) & "' "
        End If
        '員工代號
        If txt1(3) <> "" Then
            m_StrSQL_SA = m_StrSQL_SA & "AND ST01>='" & DBDATE(txt1(3)) & "' "
            m_StrSQL_SB = m_StrSQL_SB & "AND ST01>='" & DBDATE(txt1(3)) & "' "
            m_StrSQL_So = m_StrSQL_So & "AND ST01>='" & DBDATE(txt1(3)) & "' "
        End If
        If txt1(4) <> "" Then
            m_StrSQL_SA = m_StrSQL_SA & "AND ST01<='" & DBDATE(txt1(4)) & "' "
            m_StrSQL_SB = m_StrSQL_SB & "AND ST01<='" & DBDATE(txt1(4)) & "' "
            m_StrSQL_So = m_StrSQL_So & "AND ST01<='" & DBDATE(txt1(4)) & "' "
        End If
        '所別
        If txtST06(0) <> "" Then
            m_StrSQL_SA = m_StrSQL_SA & "AND ST06>='" & txtST06(0) & "' "
            m_StrSQL_SB = m_StrSQL_SB & "AND ST06>='" & txtST06(0) & "' "
            m_StrSQL_So = m_StrSQL_So & "AND ST06>='" & txtST06(0) & "' "
        End If
        If txtST06(1) <> "" Then
            m_StrSQL_SA = m_StrSQL_SA & "AND ST06<='" & txtST06(1) & "' "
            m_StrSQL_SB = m_StrSQL_SB & "AND ST06<='" & txtST06(1) & "' "
            m_StrSQL_So = m_StrSQL_So & "AND ST06<='" & txtST06(1) & "' "
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
'Modify By Sindy 2023/12/27 部門調整改抓ST93
m_str = "SELECT 1,SA11,nvl(ST93,ST03),nvl(A0922,'(舊)'||A0902),ST01,ST02,sqldateT(SA02),SA03,sqldateT(SA04),SA05,ac03,SA07,SA08,SA09,B1028,B1029 " & _
        "FROM staff_Absence,staff,allcode,acc090,acc090NEW,ABS010 " & _
        "WHERE SA01=ST01(+) and ac01(+)='04' and SA06=ac02(+) and ST03=A0901(+) and ST93=A0921(+) and SA09=B1001(+) " & m_StrSQL_SA & _
        "Union " & _
        "SELECT 2,SB12,nvl(ST93,ST03),nvl(A0922,'(舊)'||A0902),ST01,ST02,sqldateT(SB02),SB03,sqldateT(SB04),SB05,' ',SB06,SB07,SB10,B1028,B1029 " & _
        "FROM staff_busi_trip,staff,acc090,acc090NEW,ABS010 " & _
        "WHERE SB01=ST01(+) and ST03=A0901(+) and ST93=A0921(+) and SB10=B1001(+) " & m_StrSQL_SB & _
        "Union " & _
        "SELECT 3,So08,nvl(ST93,ST03),nvl(A0922,'(舊)'||A0902),ST01,ST02,sqldateT(So02),So03,sqldateT(So02),So04,' ',0,nvl(So05,So06),So13,B1028,B1029 " & _
        "FROM Staff_Overtime,staff,acc090,acc090NEW,ABS010 " & _
        "WHERE So01=ST01(+) and ST03=A0901(+) and ST93=A0921(+) and So13=B1001(+) " & m_StrSQL_So & _
        "order by 2,1,4,ST01 asc "
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
         strTemp(1) = Left(CheckStr(m_rs.Fields(3)), 5)
         strTemp(2) = CheckStr(m_rs.Fields(4))
         strTemp(3) = CheckStr(m_rs.Fields(5))
         If m_rs.Fields(0) = "1" Then
            strTemp(4) = "請假"
         ElseIf m_rs.Fields(0) = "2" Then
            strTemp(4) = "出差"
         ElseIf m_rs.Fields(0) = "3" Then
            strTemp(4) = "加班"
         End If
         strTemp(5) = CheckStr(m_rs.Fields(10))
         strTemp(6) = CheckStr(m_rs.Fields(6)) & "  " & Right("0" & Format(CheckStr(m_rs.Fields(7)), "##:##"), 5)
         strTemp(7) = CheckStr(m_rs.Fields(8)) & "  " & Right("0" & Format(CheckStr(m_rs.Fields(9)), "##:##"), 5)
         strTemp(8) = CheckStr(m_rs.Fields(11))
         strTemp(9) = CheckStr(m_rs.Fields(12))
         strTemp(10) = CheckStr(m_rs.Fields(13))
         If Not IsNull(m_rs.Fields("B1028")) Then
            strTemp(11) = Right("0" & Format(CheckStr(m_rs.Fields("B1028")), "##:##"), 5)
         End If
         If Not IsNull(m_rs.Fields("B1029")) Then
            strTemp(12) = Right("0" & Format(CheckStr(m_rs.Fields("B1029")), "##:##"), 5)
         End If
         
         If iLine > 36 Or iLine = 1 Then
            If strType <> "" Then Printer.NewPage
            iLine = 1
            Call PrintTitle(m_rs.Fields(1))  '列印表頭
         End If
         
         PrintDetail '列印明細
         
         strType = CheckStr(m_rs.Fields(1))
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

Sub PrintTitle(strDate As String)
GetPleft

Printer.Font.Size = 18
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("每日假單簽收明細表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "每日假單簽收明細表"

Printer.Font.Size = 12
iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeWStringToTDateString(strSrvDate(1))

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "簽收日期：" & ChangeWStringToTDateString(DBDATE(strDate))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "部門別"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "員工代號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "員工姓名"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "類別"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "假別"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "起始日期"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iLine * 300
Printer.Print "迄止日期"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iLine * 300
Printer.Print "日"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iLine * 300
Printer.Print "時數"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iLine * 300
Printer.Print "電子簽收"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iLine * 300
Printer.Print "起日上班時段"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iLine * 300
Printer.Print "迄日下班時段"

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(215, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 2000
PLeft(3) = 3250
PLeft(4) = 4500
PLeft(5) = 5500
PLeft(6) = 6500
PLeft(7) = 8500
PLeft(8) = 10500
PLeft(9) = 11000
PLeft(10) = 11750
PLeft(11) = 13000
PLeft(12) = 14750
End Sub

Sub PrintDetail()
Dim i As Integer
   
   For i = 1 To 12
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
   
   txtDate(0) = strSrvDate(2) 'CStr((Val(Left(strSrvDate(1), 4)) - 1911)) & "0101"
   txtDate(1) = strSrvDate(2)
   
   'M51電腦中心,M21人事處開放所別權限
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M21" Then
      txtST06(0).Enabled = True
      txtST06(1).Enabled = True
   Else 'If Pub_StrUserSt03 = "M71" Then 'M71管理部分所鎖住所別
      txtST06(0) = PUB_GetST06(strUserNum)
      txtST06(1) = PUB_GetST06(strUserNum)
      txtST06(0).Enabled = False
      txtST06(1).Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180502 = Nothing
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

Private Sub txtDate_GotFocus(Index As Integer)
   InverseTextBox txtDate(Index)
End Sub

Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
   If txtDate(Index).Text <> "" Then
      If ChkDate(txtDate(Index)) = False Then
         Call txtDate_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
   End If
   If Index = 0 Then
      If txtDate(Index) <> "" And txtDate(Index + 1) = "" Then
         txtDate(Index + 1) = txtDate(Index)
      End If
      If Val(txtDate(Index)) > Val(txtDate(Index + 1)) Then
         txtDate(Index + 1) = txtDate(Index)
      End If
   ElseIf Index = 1 Then
      If txtDate(Index) <> "" And txtDate(Index - 1) = "" Then
         txtDate(Index - 1) = txtDate(Index)
      End If
      If txtDate(Index - 1) <> "" And txtDate(Index) <> "" Then
         If RunNick2(txtDate(Index - 1), txtDate(Index)) Then
            Call txtDate_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtST06_GotFocus(Index As Integer)
   InverseTextBox txtST06(Index)
End Sub

Private Sub txtST06_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtST06_Validate(Index As Integer, Cancel As Boolean)
   If txtST06(Index) <> "" Then
      If CheckLengthIsOK(txtST06(Index), txtST06(Index).MaxLength) = False Then
          Call txtST06_GotFocus(Index)
          Cancel = True
          Exit Sub
      End If
      If Trim(txtST06(Index)) <> "" Then
         If txtST06(Index) <> "1" And txtST06(Index) <> "2" And txtST06(Index) <> "3" And _
            txtST06(Index) <> "4" And txtST06(Index) <> "5" Then
            MsgBox "所別代碼有誤!!!", vbExclamation + vbOKOnly
            Call txtST06_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
   If Index = 0 Then
      If txtST06(Index) <> "" And txtST06(Index + 1) = "" Then
         txtST06(Index + 1) = txtST06(Index)
      End If
      If txtST06(Index) > txtST06(Index + 1) Then
         txtST06(Index + 1) = txtST06(Index)
      End If
   ElseIf Index = 1 Then
      If txtST06(Index) <> "" And txtST06(Index - 1) = "" Then
         txtST06(Index - 1) = txtST06(Index)
      End If
      If txtST06(Index - 1) <> "" And txtST06(Index) <> "" Then
         If RunNick(txtST06(Index - 1), txtST06(Index)) Then
            Call txtST06_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
CloseIme
End Sub
