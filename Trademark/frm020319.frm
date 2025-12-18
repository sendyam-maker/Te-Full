VERSION 5.00
Begin VB.Form frm020319 
   BorderStyle     =   1  '單線固定
   Caption         =   "下載商標圖參考報表"
   ClientHeight    =   2568
   ClientLeft      =   3648
   ClientTop       =   1968
   ClientWidth     =   5364
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2568
   ScaleWidth      =   5364
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1950
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1560
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   3285
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1560
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   3285
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1200
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1950
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1200
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1950
      MaxLength       =   7
      TabIndex        =   0
      Top             =   825
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   3285
      MaxLength       =   7
      TabIndex        =   1
      Top             =   825
      Width           =   1110
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3120
      TabIndex        =   6
      Top             =   105
      Width           =   972
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   4170
      TabIndex        =   7
      Top             =   105
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "審定來函日："
      Height          =   180
      Index           =   2
      Left            =   825
      TabIndex        =   11
      Top             =   1605
      Width           =   1095
   End
   Begin VB.Line Line3 
      X1              =   3240
      X2              =   3090
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "(三選一查詢)"
      ForeColor       =   &H00000080&
      Height          =   180
      Index           =   1
      Left            =   840
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   3240
      X2              =   3090
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   3240
      X2              =   3090
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "專用期止日："
      Height          =   180
      Index           =   0
      Left            =   825
      TabIndex        =   9
      Top             =   1245
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "申請日："
      Height          =   180
      Index           =   3
      Left            =   825
      TabIndex        =   8
      Top             =   870
      Width           =   1095
   End
End
Attribute VB_Name = "frm020319"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
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


Private Sub cmdOK_Click(Index As Integer)
Dim intCnt As Integer
   Select Case Index
      Case 0
            Printer.Orientation = 2
            DoEvents
            
'            If txt1(0) = "" And txt1(1) = "" And txt1(2) = "" And txt1(3) = "" Then
'                MsgBox "請至少輸入一項列印條件！", vbInformation, "操作錯誤！"
'                txt1(0).SetFocus
'                Exit Sub
'            End If
            If (txt1(0) = "" And txt1(2) = "" And txt1(4) = "") Or _
               (txt1(0) <> "" And txt1(2) <> "" And txt1(4) <> "") Then
                MsgBox "請三選一輸入！", vbInformation, "操作錯誤！"
                txt1(0).SetFocus
                Exit Sub
            End If
            intCnt = 0
            If (txt1(0) <> "" Or txt1(1) <> "") Then intCnt = intCnt + 1
            If (txt1(2) <> "" Or txt1(3) <> "") Then intCnt = intCnt + 1
            If (txt1(4) <> "" Or txt1(5) <> "") Then intCnt = intCnt + 1
            If intCnt > 1 Then
                MsgBox "請三選一輸入！", vbInformation, "操作錯誤！"
                txt1(0).SetFocus
                Exit Sub
            End If
            
            If (txt1(0) = "" And txt1(1) <> "") Then
                MsgBox "申請(起始)日不可空白！", vbInformation, "操作錯誤！"
                txt1(0).SetFocus
                Exit Sub
            End If
            If (txt1(0) <> "" And txt1(1) = "") Then
                MsgBox "申請(迄止)日不可空白！", vbInformation, "操作錯誤！"
                txt1(1).SetFocus
                Exit Sub
            End If
            If (txt1(2) = "" And txt1(3) <> "") Then
                MsgBox "專用期止(起始)日不可空白！", vbInformation, "操作錯誤！"
                txt1(2).SetFocus
                Exit Sub
            End If
            If (txt1(2) <> "" And txt1(3) = "") Then
                MsgBox "專用期止(迄止)日不可空白！", vbInformation, "操作錯誤！"
                txt1(3).SetFocus
                Exit Sub
            End If
            If (txt1(4) = "" And txt1(5) <> "") Then
                MsgBox "審定來函(起始)日不可空白！", vbInformation, "操作錯誤！"
                txt1(4).SetFocus
                Exit Sub
            End If
            If (txt1(4) <> "" And txt1(5) = "") Then
                MsgBox "審定來函(迄止)日不可空白！", vbInformation, "操作錯誤！"
                txt1(5).SetFocus
                Exit Sub
            End If
            
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/15 清除查詢印表記錄檔欄位
            Call PrintData
            Me.Enabled = True
            Screen.MousePointer = vbDefault
      Case 1
           Unload Me
      Case Else
   End Select
End Sub

Sub PrintData()
Dim dblCnt As Double

'依申請案號(發審定書時使用)
'Modified by Morgan 2023/9/23
'改語法 AND TM01||TM02||TM03||TM04 not in (SELECT IBF01||IBF02||IBF03||IBF04 FROM imgbytefile WHERE IBF01='T')
'--> AND not exists( SELECT * FROM imgbytefile WHERE IBF01=tm01 and IBF02=tm02 and IBF03=tm03 and IBF04=TM04)
If txt1(0) <> "" And txt1(1) <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(0) & "-" & txt1(1) 'Add By Sindy 2010/10/15
   m_str = "SELECT TM12,substr(TM05,1,20),TM09,TM01||'-'||TM02||'-'||TM03||'-'||TM04,sqldateT(TM11),substr(CU04,1,20) " & _
                    " From Trademark, Customer " & _
               " WHERE TM11 >= " & ChangeTStringToWString(txt1(0)) & " AND TM11 <= " & ChangeTStringToWString(txt1(1)) & _
                    " AND TM01='T' " & _
                    " AND TM10='000' " & _
                    " AND TM29 is null " & _
                    " AND substr(TM23,1,8)=CU01(+) " & _
                    " AND substr(TM23,9,1)=CU02(+) " & _
                    " AND not exists( SELECT * FROM imgbytefile WHERE IBF01=tm01 and IBF02=tm02 and IBF03=tm03 and IBF04=TM04) " & _
               " Order By TM11,TM12"
'依註冊號(延展時用)
ElseIf txt1(2) <> "" And txt1(3) <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/10/15
   m_str = "SELECT TM15,substr(TM05,1,20),TM09,TM01||'-'||TM02||'-'||TM03||'-'||TM04,TM12,substr(CU04,1,20) " & _
                    " From Trademark, Customer " & _
               " WHERE TM22 >= " & ChangeTStringToWString(txt1(2)) & " AND TM22 <= " & ChangeTStringToWString(txt1(3)) & _
                    " AND TM01='T' " & _
                    " AND TM10='000' " & _
                    " AND TM29 is null " & _
                    " AND substr(TM23,1,8)=CU01(+) " & _
                    " AND substr(TM23,9,1)=CU02(+) " & _
                    " AND not exists( SELECT * FROM imgbytefile WHERE IBF01=tm01 and IBF02=tm02 and IBF03=tm03 and IBF04=TM04) " & _
               " Order By TM22,TM15"
'審定來函日(依申請案號)
ElseIf txt1(4) <> "" And txt1(5) <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/10/15
   m_str = "SELECT TM12,substr(TM05,1,20),TM09,TM01||'-'||TM02||'-'||TM03||'-'||TM04,sqldateT(TM13),substr(CU04,1,20) " & _
                    " From Trademark, Customer " & _
               " WHERE TM13 >= " & ChangeTStringToWString(txt1(4)) & " AND TM13 <= " & ChangeTStringToWString(txt1(5)) & _
                    " AND TM01='T' " & _
                    " AND TM10='000' " & _
                    " AND TM29 is null " & _
                    " AND substr(TM23,1,8)=CU01(+) " & _
                    " AND substr(TM23,9,1)=CU02(+) " & _
                    " AND not exists( SELECT * FROM imgbytefile WHERE IBF01=tm01 and IBF02=tm02 and IBF03=tm03 and IBF04=TM04) " & _
                    " AND TM16='1' " & _
               " Order By TM13,TM12"
End If
dblCnt = 0
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        InsertQueryLog (m_rs.RecordCount) 'Add By Sindy 2010/10/15
        .MoveFirst
        
        iLine = 1
        strType = ""
        
        Do While Not .EOF
            For m_i = 1 To 6
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields(0))
            strTemp(2) = CheckStr(m_rs.Fields(1))
            strTemp(3) = CheckStr(m_rs.Fields(2))
            strTemp(4) = CheckStr(m_rs.Fields(3))
            strTemp(5) = CheckStr(m_rs.Fields(4))
            strTemp(6) = CheckStr(m_rs.Fields(5))
            dblCnt = dblCnt + 1
            If iLine > 35 Or iLine = 1 Then 'Or
               '(strType <> "" And strType <> CheckStr(m_rs.Fields(0))) Then
               'If .AbsolutePosition <> .RecordCount Then
                   If strType <> "" Then Printer.NewPage
                   iLine = 1
                   PrintTitle '列印表頭
               'End If
            End If
            PrintDetail
            
            strType = CheckStr(m_rs.Fields(3))
            .MoveNext
        Loop
        '合計
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iLine * 300
        Printer.Print String(205, "-")
        iLine = iLine + 1
        Printer.CurrentX = PLeft(6)
        Printer.CurrentY = iLine * 300
        Printer.Print "共  " & dblCnt & "  件"
    End With
Else
    InsertQueryLog (0) 'Add By Sindy 2010/10/15
    ShowNoData
    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub GetPleft()
PLeft(1) = 600
PLeft(2) = 2000
PLeft(3) = 6500
PLeft(4) = 8000
PLeft(5) = 9750
PLeft(6) = 11500
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("下載商標圖參考報表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "下載商標圖參考報表"
Printer.Font.Size = 12
Printer.FontBold = False
Printer.Font.Underline = False
iLine = iLine + 2
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "列印人：" & GetStaffName(strUserNum, False)
'依申請案號(發審定書時使用)
If txt1(0) <> "" And txt1(1) <> "" Then
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("申請日：" & ChangeTStringToTDateString(txt1(0)) & " - " & ChangeTStringToTDateString(txt1(1))) / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "申請日：" & ChangeTStringToTDateString(txt1(0)) & " - " & ChangeTStringToTDateString(txt1(1))
'依註冊號(延展時用)
ElseIf txt1(2) <> "" And txt1(3) <> "" Then
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("專用期止日：" & ChangeTStringToTDateString(txt1(2)) & " - " & ChangeTStringToTDateString(txt1(3))) / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "專用期止日：" & ChangeTStringToTDateString(txt1(2)) & " - " & ChangeTStringToTDateString(txt1(3))
ElseIf txt1(4) <> "" And txt1(5) <> "" Then
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("審定來函日：" & ChangeTStringToTDateString(txt1(4)) & " - " & ChangeTStringToTDateString(txt1(5))) / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "審定來函日：" & ChangeTStringToTDateString(txt1(4)) & " - " & ChangeTStringToTDateString(txt1(5))
End If
iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
'依申請案號(發審定書時使用)
If txt1(0) <> "" And txt1(1) <> "" Then
   Printer.Print "申請案號"
'依註冊號(延展時用)
ElseIf txt1(2) <> "" And txt1(3) <> "" Then
   Printer.Print "註冊號"
ElseIf txt1(4) <> "" And txt1(5) <> "" Then
   Printer.Print "申請案號"
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "商標名稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "類別"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
'依申請案號(發審定書時使用)
If txt1(0) <> "" And txt1(1) <> "" Then
   Printer.Print "申請日"
'依註冊號(延展時用)
ElseIf txt1(2) <> "" And txt1(3) <> "" Then
   Printer.Print "申請案號"
ElseIf txt1(4) <> "" And txt1(5) <> "" Then
   Printer.Print "審定來函日"
End If
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "申請人"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(205, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 6
'   If m_j = 4 Or m_j = 5 Then
'      Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j))
'   Else
      Printer.CurrentX = PLeft(m_j)
'   End If
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm020319 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    InverseTextBox txt1(Index)
    CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2, 3
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If txt1(Index).Text <> "" Then
            If ChkDate(txt1(Index)) = False Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 0 Then
'            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
'               txt1(Index + 1) = txt1(Index)
'            End If
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
'            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
'               txt1(Index + 1) = txt1(Index)
'            End If
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
