VERSION 5.00
Begin VB.Form frm160307 
   BorderStyle     =   1  '單線固定
   Caption         =   "忘記打卡次數"
   ClientHeight    =   3190
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3190
   ScaleWidth      =   5040
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   1
      Left            =   1710
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "1"
      Top             =   1260
      Width           =   195
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3945
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1710
      MaxLength       =   3
      TabIndex        =   0
      Top             =   900
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   60
      TabIndex        =   5
      Top             =   2490
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   2
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "類別："
      Height          =   180
      Left            =   1080
      TabIndex        =   9
      Top             =   1290
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "(1.分所 2.部門)"
      Height          =   240
      Left            =   2010
      TabIndex        =   8
      Top             =   1260
      Width           =   2595
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "年度："
      Height          =   180
      Left            =   1080
      TabIndex        =   7
      Top             =   960
      Width           =   540
   End
End
Attribute VB_Name = "frm160307"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
'Create by SINDY 2010/01/15
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_str2 As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 11) As String '次數
Dim strDeptNm(1 To 11) As String '部門
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim intRow As Integer
Dim strItem As String


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If txt1(0) = "" Then
            MsgBox "年度不可以空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
         End If
         If txt1(1) = "" Then
            MsgBox "類別不可以空白！", vbInformation, "操作錯誤！"
            txt1(1).SetFocus
            Exit Sub
         End If
         
         Screen.MousePointer = vbHourglass
         m_StrSQL = "and sa02 between " & Val(txt1(0)) + 1911 & "0101 and " & Val(txt1(0)) + 1911 & "1231 "
         Call StrMenu
         Screen.MousePointer = vbDefault
      Case 1
           Unload Me
   End Select
   Printer.Font.Size = 12
End Sub

'明細表
Sub StrMenu()
Dim int_i As Integer
Dim decTot As Double
Dim varDept As Variant
Dim i As Integer
   
   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   Printer.PaperSize = 9  'PDF
   
   For m_i = 1 To 5
       strTemp(m_i) = ""
   Next m_i
   
   decTot = 0
   'Modify By Sindy 2023/12/27 部門調整改抓ST93
   If txt1(1) = "1" Then '1.分所
      m_str = "select st06,sum(decode(sa03,null,0,sa03)) " & _
                  "From Staff_Assist, staff " & _
                  "where sa01=st01 " & m_StrSQL & _
                  "and st93<>'R04' " & _
                  "group by st06 order by st06 "
      If m_rs.State = 1 Then m_rs.Close
      m_rs.CursorLocation = adUseClient
      m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
      If Not m_rs.EOF And Not m_rs.BOF Then
         With m_rs
            m_rs.MoveFirst
            m_i = 0
            Do While Not m_rs.EOF
               m_i = m_i + 1
               strTemp(m_i) = "" & .Fields(1) '次數
               decTot = decTot + Val("" & .Fields(1))
               m_rs.MoveNext
            Loop
         End With
      End If
      
      m_str = "select sum(decode(sa03,null,0,sa03)) " & _
                  "From Staff_Assist, staff " & _
                  "where sa01=st01 " & m_StrSQL & _
                  "and st93='R04' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, m_str)
      If intI = 1 Then
         strTemp(5) = "" & RsTemp.Fields(0) '台一投資 次數
         decTot = decTot + Val("" & RsTemp.Fields(0))
      End If
      
   Else '2.部門
      'Modify By Sindy 2023/12/27
      varDept = Split(A0925For1Code, ",") '11個單位
      For i = 0 To UBound(varDept)
         m_str = "select sum(decode(sa03,null,0,sa03))," & A0925CName & _
                 " From Staff_Assist, staff" & _
                 " where sa01=st01 " & m_StrSQL & _
                 " and substr(st93,1,1)='" & varDept(i) & "'" & _
                 " group by " & A0925CName
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, m_str)
         If intI = 1 Then
            strTemp(i + 1) = "" & RsTemp.Fields(0) '次數
            decTot = decTot + Val("" & RsTemp.Fields(0))
            strDeptNm(i + 1) = "" & RsTemp.Fields(1) '大部門名稱
         Else
            m_str = "select 0," & A0925CName & _
                    " From staff" & _
                    " where substr(st93,1,1)='" & varDept(i) & "'" & _
                    " group by " & A0925CName
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, m_str)
            If intI = 1 Then
               strTemp(i + 1) = "" & RsTemp.Fields(0) '次數
               decTot = decTot + Val("" & RsTemp.Fields(0))
               strDeptNm(i + 1) = "" & RsTemp.Fields(1) '大部門名稱
            End If
         End If
      Next
      '2023/12/27 END
'      m_str = "select sum(decode(sa03,null,0,sa03)) " & _
'                     "From Staff_Assist, staff " & _
'                     "where sa01=st01 " & m_StrSQL & _
'                     "and substr(st03,1,1)='F' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, m_str)
'      If intI = 1 Then
'         strTemp(1) = "" & RsTemp.Fields(0) '國外部 次數
'         decTot = decTot + Val("" & RsTemp.Fields(0))
'      End If
'      m_str = "select sum(decode(sa03,null,0,sa03)) " & _
'                     "From Staff_Assist, staff " & _
'                     "where sa01=st01 " & m_StrSQL & _
'                     "and (substr(st03,1,1)='L' or substr(st03,1,1)='P') "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, m_str)
'      If intI = 1 Then
'         strTemp(2) = "" & RsTemp.Fields(0) '專業部 次數
'         decTot = decTot + Val("" & RsTemp.Fields(0))
'      End If
'      m_str = "select sum(decode(sa03,null,0,sa03)) " & _
'                     "From Staff_Assist, staff " & _
'                     "where sa01=st01 " & m_StrSQL & _
'                     "and (substr(st03,1,1)<>'F' and substr(st03,1,1)<>'L' and substr(st03,1,1)<>'P') and substr(st03,1,1)<>'S' and st03<>'R04' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, m_str)
'      If intI = 1 Then
'         strTemp(3) = "" & RsTemp.Fields(0) '管理部 次數
'         decTot = decTot + Val("" & RsTemp.Fields(0))
'      End If
'      m_str = "select sum(decode(sa03,null,0,sa03)) " & _
'                     "From Staff_Assist, staff " & _
'                     "where sa01=st01 " & m_StrSQL & _
'                     "and substr(st03,1,1)='S' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, m_str)
'      If intI = 1 Then
'         strTemp(4) = "" & RsTemp.Fields(0) '智權部 次數
'         decTot = decTot + Val("" & RsTemp.Fields(0))
'      End If
   End If
   
'   m_str = "select sum(decode(sa03,null,0,sa03)) " & _
'               "From Staff_Assist, staff " & _
'               "where sa01=st01 " & m_StrSQL & _
'               "and st03='R04' "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, m_str)
'   If intI = 1 Then
'      strTemp(10) = "" & RsTemp.Fields(0) '台一投資 次數
'      decTot = decTot + Val("" & RsTemp.Fields(0))
'   End If
   
   If decTot = 0 Then
      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   iLine = 0
   PrintTitle '列印表頭
   PrintDetail
   
   iLine = iLine + 1
   Printer.Line (500, iLine * 300)-(11000, iLine * 300), , B
   iLine = iLine + 1
   
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "合計"
   Printer.CurrentX = PLeft(2) - Printer.TextWidth(decTot)
   Printer.CurrentY = iLine * 300
   Printer.Print decTot
   
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintTitle()

GetPleft

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

iLine = iLine + 2
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(txt1(0) & "年度忘記打卡次數") / 2)
Printer.CurrentY = iLine * 300
Printer.Print txt1(0) & "年度忘記打卡次數"

iLine = iLine + 2
Printer.Font.Size = 12
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 3
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
If txt1(1) = "1" Then
   Printer.Print "分所"
Else
   Printer.Print "部門"
End If

iLine = iLine + 1
Printer.Line (500, iLine * 300)-(11000, iLine * 300), , B
iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 4250
PLeft(2) = 7000
End Sub

Sub PrintDetail()
Dim i As Integer
Dim strText(1 To 5) As String
   
   If txt1(1) = "1" Then '分所
      For i = 1 To 5
         strText(1) = "北所"
         strText(2) = "中所"
         strText(3) = "南所"
         strText(4) = "高所"
         strText(5) = "臺一投資"    'modify by sonia 2021/2/25 改名稱
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iLine * 300
         Printer.Print strText(i)
         Printer.CurrentX = PLeft(2) - Printer.TextWidth(strTemp(i))
         Printer.CurrentY = iLine * 300
         Printer.Print strTemp(i)
         If i <> 5 Then iLine = iLine + 2
      Next i
   Else '部門
      For i = 1 To 11
'         strText(1) = "國外部"
'         strText(2) = "專業部"
'         strText(3) = "管理部"
'         strText(4) = "智權部"
'         strText(5) = "臺一投資"    'modify by sonia 2021/2/25 改名稱
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iLine * 300
         Printer.Print strDeptNm(i)
         Printer.CurrentX = PLeft(2) - Printer.TextWidth(strTemp(i))
         Printer.CurrentY = iLine * 300
         Printer.Print strTemp(i)
         If i <> 11 Then iLine = iLine + 2
      Next i
   End If
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
   Set frm160307 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(0).Text = "" Then Exit Sub
         If CheckIsTaiwanDate(txt1(0).Text & "0101", False) = False Then
             Cancel = True
             MsgBox "請輸入民國年度！", vbInformation, "輸入新年度錯誤"
             Exit Sub
         End If
      Case 1
            If txt1(Index) <> "" Then
               Select Case txt1(Index)
               Case "1", "2"
               Case Else
                   MsgBox "類別只可以輸入 1 或 2！", vbInformation, "輸入錯誤！"
                   Call txt1_GotFocus(Index)
                   Cancel = True
                   Exit Sub
               End Select
            End If
      Case Else
   End Select
End Sub
