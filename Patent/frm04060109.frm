VERSION 5.00
Begin VB.Form frm04060109 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內公報代理人資料列印"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4980
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   9
      Top             =   2520
      Width           =   4875
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
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2790
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1410
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   2
      Left            =   1950
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1410
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2670
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1050
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   1950
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1050
      Width           =   615
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3975
      TabIndex        =   6
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3030
      TabIndex        =   5
      Top             =   90
      Width           =   915
   End
   Begin VB.Line Line2 
      X1              =   2580
      X2              =   2940
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "建檔時公告日："
      Height          =   180
      Left            =   660
      TabIndex        =   8
      Top             =   1440
      Width           =   1260
   End
   Begin VB.Line Line1 
      X1              =   2190
      X2              =   3180
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "代理人代號："
      Height          =   180
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   1080
      Width           =   1080
   End
End
Attribute VB_Name = "frm04060109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Create by Sindy 2011/5/30
Option Explicit

'strTA01   P:專利   T:商標
Public strTA01 As String

Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 4) As Integer
Dim strTemp(1 To 4) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
        If txt1(0) = "" And txt1(1) = "" And txt1(2) = "" And txt1(3) = "" Then
            MsgBox "請至少輸入一項列印條件！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            'Modified by Morgan 2016/12/29 條件下1~1### 會少印0#及2#-9#的代理人--蕭茹曣
            'm_StrSQL = m_StrSQL & " AND TA02 >= '" & txt1(0) & "' "
            m_StrSQL = m_StrSQL & " AND to_number(TA02) >= " & Val(txt1(0)) & " "
        End If
        If txt1(1) <> "" Then
            'Modified by Morgan 2016/12/29 條件下1~1### 會少印0#及2#-9#的代理人--蕭茹曣
            'm_StrSQL = m_StrSQL & " AND TA02 <= '" & txt1(1) & "' "
            m_StrSQL = m_StrSQL & " AND to_number(TA02) <= " & Val(txt1(1)) & " "
        End If
        If txt1(2) <> "" Then
            m_StrSQL = m_StrSQL & " AND TA05 >= '" & ChangeTStringToWString(txt1(2)) & "' "
        End If
        If txt1(3) <> "" Then
            m_StrSQL = m_StrSQL & " AND TA05 <= '" & ChangeTStringToWString(txt1(3)) & "' "
        End If
        StrMenu1
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
Case Else
End Select
End Sub

Sub StrMenu1()

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

m_str = "select TA02,TA03,TA04,sqldatet(TA05) " & _
          "From tagent " & _
         "where TA01='" & strTA01 & "' " & m_StrSQL & _
         "order by to_number(TA02) asc "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        .MoveFirst
        
        iLine = 1
        strType = ""
        
        Do While Not .EOF
            For m_i = 1 To 4
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(.Fields(0))
            strTemp(2) = CheckStr(.Fields(1))
            strTemp(3) = CheckStr(.Fields(2))
            strTemp(4) = CheckStr(.Fields(3))
            
            'Modified by Morgan 2016/12/29 跳頁異常--蕭茹曣
            'If iLine > 54 Or iLine = 1 Then
            If (iLine + 2) * 300 > Printer.ScaleHeight Or iLine = 1 Then
            'end 2016/12/29
               If strType <> "" Then Printer.NewPage
               iLine = 1
               PrintTitle '列印表頭
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
PLeft(2) = 2000
PLeft(3) = 5000
PLeft(4) = 8500
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("國內公報代理人資料") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "國內公報代理人資料"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "代號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "代理人名稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "事務所名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "建檔時公告日"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")
iLine = iLine + 1
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
Set frm04060109 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    InverseTextBox txt1(Index)
    CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 2, 3
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 0, 1
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
         If txt1(Index).Text <> "" Then
            If ChkDate(txt1(Index)) = False Then
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
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
