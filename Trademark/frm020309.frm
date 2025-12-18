VERSION 5.00
Begin VB.Form frm020309 
   BorderStyle     =   1  '單線固定
   Caption         =   "商品類別/組群案件明細表"
   ClientHeight    =   1755
   ClientLeft      =   3510
   ClientTop       =   3075
   ClientWidth     =   3255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3255
   Begin VB.CommandButton CMDOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2472
      TabIndex        =   6
      Top             =   12
      Width           =   756
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   12
      Width           =   756
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   2
      Left            =   2196
      MaxLength       =   7
      TabIndex        =   2
      Top             =   768
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   1
      Left            =   1116
      MaxLength       =   7
      TabIndex        =   1
      Top             =   768
      Width           =   990
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   4
      Left            =   1116
      TabIndex        =   4
      Top             =   1440
      Width           =   2085
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   3
      Left            =   1116
      TabIndex        =   3
      Top             =   1104
      Width           =   2100
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   0
      Left            =   1116
      TabIndex        =   0
      Top             =   432
      Width           =   2055
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1620
      X2              =   2820
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label1 
      Caption         =   "申請日："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   816
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1164
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "商品組群："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1488
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   492
      Width           =   948
   End
End
Attribute VB_Name = "frm020309"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay(0 To 2) As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 15) As String, strTemp3 As String, TestOk As Boolean
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     Printer.Orientation = 2
     DoEvents
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         'Add By Cheng 2002/03/21
         If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
            Me.txt1(1).SetFocus
            txt1_GotFocus 1
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
            Me.txt1(2).SetFocus
            txt1_GotFocus 2
            Exit Sub
         End If
         
         If Len(txt1(2)) = 0 Then
             s = MsgBox("申請日區間不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             txt1_GotFocus (1)
             Exit Sub
         Else
             ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/4 清除查詢印表記錄檔欄位
             Screen.MousePointer = vbHourglass
             Me.Enabled = False
             Process
             Me.Enabled = True
             Screen.MousePointer = vbDefault
         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "delete from r020309 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 & " AND TM01 IN (" & SQLGrpStr(txt1(0), 2) & ")"
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/10/4
End If
StrSQL6 = ""
If Len(txt1(1)) <> 0 Then
      strSQL1 = strSQL1 + " and TM11>=" & Val(ChangeTStringToWString(txt1(1))) & ""
End If
If Len(Trim(txt1(2))) <> 0 Then
   strSQL1 = strSQL1 + " AND TM11<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(txt1(1)) <> 0 Or Len(Trim(txt1(2))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/4
End If
If Len(txt1(3)) <> 0 Then
    StrSQL6 = " AND ("
    If Len(txt1(0)) <> 0 Then
        strTemp1 = Split(txt1(3), ",")
        For i = 0 To UBound(strTemp1)
            strSQL2 = strSQL2 + " instr(TM09,'" & strTemp1(i) & "')>0  "
            If i <> UBound(strTemp1) Then
                strSQL2 = strSQL2 + " OR "
            End If
        Next i
        'StrSQL1 = StrSQL1 + " TM01=' ') "
        strSQL2 = strSQL2 + " ) "
    End If
    'Modified by Lydia 2018/03/16 修正; 同時放大R061005為V2(699)
    'strSQL1 = StrSQL6 + strSQL2
    strSQL1 = strSQL1 + StrSQL6 + strSQL2
    pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) 'Add By Sindy 2010/10/4
End If
strSQL2 = ""
If Len(txt1(4)) <> 0 Then
    StrSQL6 = " AND ("
    If Len(txt1(0)) <> 0 Then
        strTemp1 = Split(txt1(4), ",")
        For i = 0 To UBound(strTemp1)
            strSQL2 = strSQL2 + " instr(TM32,'" & strTemp1(i) & "')>0 "
            If i <> UBound(strTemp1) Then
                strSQL2 = strSQL2 + " OR "
            End If
        Next i
        'StrSQL1 = StrSQL1 + " TM01=' ') "
        strSQL2 = strSQL2 + ") "
    End If
    'Modified by Lydia 2018/03/16
    'strSQL1 = StrSQL6 + strSQL2
    strSQL1 = strSQL1 + StrSQL6 + strSQL2
    pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(4) 'Add By Sindy 2010/10/4
End If
CheckOC
strSql = "SELECT " & SQLDate("TM11") & ",TM01||'-'||TM02||'-'||TM03||'-'||TM04,NVL(TM05,NVL(TM06,TM07)),TM32,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),TM09 FROM TRADEMARK,CUSTOMER WHERE SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) " & strSQL1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/4
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 5
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If Len(strTemp(5)) <> 0 Then
                strTemp1 = Split(strTemp(5), ",")
                For i = 0 To UBound(strTemp1)
                    If Len(Trim(strTemp1(i))) <> 0 Then
                        strSql = "INSERT INTO R020309 VALUES ('" & ChgSQL(strTemp1(i)) & "','" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                        cnnConnection.Execute strSql
                    End If
                Next i
            Else
                strSql = "INSERT INTO R020309 VALUES ('   ','" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
            End If
            .MoveNext
            DoEvents
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/4
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
PrintData
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
strSql = "SELECT * FROM R020309 WHERE ID='" & strUserNum & "' ORDER BY R061001,R061002,R061003 "
CheckOC
Page = 1
SavDay1 = ""
SavDay2 = "        "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        SavDay1 = CheckStr(.Fields(0))
        SavDay2 = "  "
        PrintTitle
        Do While .EOF = False
            For i = 0 To 5
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If SavDay1 <> strTemp(0) Then
                Page = Page + 1
                Printer.NewPage
                SavDay1 = strTemp(0)
                SavDay2 = "  "
                PrintTitle
            Else
                If SavDay2 = strTemp(1) Then
                    strTemp(1) = ""
                Else
                    SavDay2 = strTemp(1)
                End If
            End If
            strTemp(3) = StrToStr(strTemp(3), 8)
            strTemp(4) = StrToStr(strTemp(4), 18)
            strTemp(5) = StrToStr(strTemp(5), 9)
            PrintDatil
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
Printer.EndDoc
CheckOC
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "商品類別/組群明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6200
Printer.CurrentY = iPrint
Printer.Print "申請日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "商品類別：" & SavDay1
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "申請日"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "組       群"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "申請人"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub PrintDatil()
For i = 1 To 5
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 500
PLeft(2) = 1700
PLeft(3) = 3700
PLeft(4) = 6000
PLeft(5) = 14000
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm020309 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdOK(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     strTemp1 = Split(UCase(GetSystemKindByNick), ",")
     strTemp2 = Split(UCase(txt1(0)), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp2(i) = strTemp1(j) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
     Next i
Case 1, 2
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 2 Then
     If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   End If
Case Else
End Select
End Sub

