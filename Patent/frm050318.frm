VERSION 5.00
Begin VB.Form frm050318 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人/申請人名單"
   ClientHeight    =   1950
   ClientLeft      =   4755
   ClientTop       =   3525
   ClientWidth     =   4350
   ControlBox      =   0   'False
   LinkTopic       =   "Form17"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4350
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2100
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1470
      Width           =   420
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   3084
      MaxLength       =   9
      TabIndex        =   1
      Top             =   480
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1416
      MaxLength       =   9
      TabIndex        =   0
      Top             =   480
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3492
      TabIndex        =   8
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2700
      TabIndex        =   7
      Top             =   48
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   3096
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1152
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1416
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1152
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   3084
      MaxLength       =   9
      TabIndex        =   3
      Top             =   816
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1416
      MaxLength       =   9
      TabIndex        =   2
      Top             =   816
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "代理人編號："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   816
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "申請人編號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否含不寄雜誌的對象：           (Y : 是)"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   1500
      Width           =   3030
   End
   Begin VB.Line Line3 
      X1              =   2736
      X2              =   2976
      Y1              =   1272
      Y2              =   1272
   End
   Begin VB.Line Line2 
      X1              =   2736
      X2              =   2976
      Y1              =   936
      Y2              =   936
   End
   Begin VB.Line Line1 
      X1              =   2736
      X2              =   2976
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "國籍："
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   1155
      Width           =   540
   End
End
Attribute VB_Name = "frm050318"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String
Dim i As Integer, j As Integer, s As Integer
Dim strTemp(0 To 4) As String, PLeft(0 To 4) As Integer, iPrint As Integer, Page As Integer

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
      '900725 邱小姐說申請人或代理人或國籍必須最少輸入一樣
      '且申請人前6碼或代理人前6碼必須相同
      '  NICK   PEN
     If Option1(0).Value = True Then
        If (Len(txt1(0)) = 0 And Len(txt1(1)) = 0) And Len(txt1(5)) = 0 Then
            s = MsgBox("申請人編號區間 或 國籍編號區間不可空白 或 申請人前 6 碼必須相同!!", , "USER 輸入錯誤")
            txt1(0).SetFocus
            txt1_GotFocus (0)
            Exit Sub
        End If
        If Len(txt1(0)) <> 0 Or Len(txt1(1)) <> 0 Then
            If Mid(txt1(0), 1, 6) <> Mid(txt1(1), 1, 6) Then
               s = MsgBox("申請人編號區間 或 國籍編號區間不可空白 或 申請人前 6 碼必須相同!!", , "USER 輸入錯誤")
               txt1(0).SetFocus
               txt1_GotFocus (0)
               Exit Sub
            End If
        End If
     Else
        If (Len(txt1(2)) = 0 And Len(txt1(3)) = 0) And Len(txt1(5)) = 0 Then
            s = MsgBox("代理人編號區間 或 國籍編號區間不可空白 或 代理人前 6 碼必須相同!!", , "USER 輸入錯誤")
            txt1(2).SetFocus
            txt1_GotFocus (2)
            Exit Sub
        End If
        If Len(txt1(2)) <> 0 Or Len(txt1(3)) <> 0 Then
            If Mid(txt1(2), 1, 6) <> Mid(txt1(3), 1, 6) Then
               s = MsgBox("代理人編號區間 或 國籍編號區間不可空白 或 代理人前 6 碼必須相同!!", , "USER 輸入錯誤")
               txt1(2).SetFocus
               txt1_GotFocus (2)
               Exit Sub
            End If
        End If
     End If
     Screen.MousePointer = vbHourglass
     Me.Enabled = False
     ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/4 清除查詢印表記錄檔欄位
     Process
     Me.Enabled = True
     Screen.MousePointer = vbDefault
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Process()
Screen.MousePointer = vbHourglass
'cnnConnection.Execute "DELETE FROM R050318 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
If Option1(0).Value = True Then
    If Len(txt1(0)) <> 0 Then
        strSQL1 = strSQL1 + " AND CU01>='" & Mid(GetNewFagent(txt1(0)), 1, 8) & "' AND CU02>='" & Mid(GetNewFagent(txt1(0)), 9, 1) & "' "
    End If
    If Len(txt1(1)) <> 0 Then
        strSQL1 = strSQL1 + " AND CU01<='" & Mid(GetNewFagent(txt1(1)), 1, 8) & "' AND CU02<='" & Mid(GetNewFagent(txt1(1)), 9, 1) & "' "
    End If
    If Len(txt1(0)) <> 0 Or Len(txt1(1)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(0) & "-" & txt1(1) 'Add By Sindy 2010/10/4
    End If
Else
    If Len(txt1(2)) <> 0 Then
        strSQL1 = strSQL1 + " AND FA01>='" & Mid(GetNewFagent(txt1(2)), 1, 8) & "' AND FA02>='" & Mid(GetNewFagent(txt1(2)), 9, 1) & "' "
    End If
    If Len(txt1(3)) <> 0 Then
        strSQL1 = strSQL1 + " AND FA01<='" & Mid(GetNewFagent(txt1(3)), 1, 8) & "' AND FA02<='" & Mid(GetNewFagent(txt1(3)), 9, 1) & "' "
    End If
    If Len(txt1(2)) <> 0 Or Len(txt1(3)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/10/4
    End If
End If

'選擇列印申請人
If Option1(0).Value = True Then
   If Len(Trim(txt1(4))) <> 0 Then
      strSQL1 = strSQL1 + " AND cu10>='" & txt1(4) & "' "
   End If
   If Len(Trim(txt1(5))) <> 0 Then
      strSQL1 = strSQL1 & " AND cu10<='" & txt1(5) & "z' "
   End If
   If Len(txt1(4)) <> 0 Or Len(txt1(5)) <> 0 Then
      pub_QL05 = pub_QL05 & ";申請人" & Label1 & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/10/4
    End If
   'Add By Cheng 2002/05/02
   If Len(Trim(txt1(6))) <= 0 Then
      strSQL1 = strSQL1 & " AND CU32 Is Null "
   Else
      pub_QL05 = pub_QL05 & ";" & Left(Label2, 11) & "是" 'Add By Sindy 2010/10/4
   End If
   'Modify By Cheng 2002/05/02
'   strSQL = "SELECT NVL(NA03,CU10),CU01||CU02 AS A,CU03,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29)),NA01 FROM CUSTOMER,NATION WHERE cu10=na01(+) AND CU32 IS NULL " & strSQL1 & " ORDER BY NA01,A "
   '2005/6/22 modify by sonia
   'strSQL = "SELECT NVL(NA03,CU10),CU01||CU02 AS A,CU03,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29)),NA01 FROM CUSTOMER,NATION WHERE cu10=na01(+) " & strSQL1 & " ORDER BY NA01,A "
   strSql = "SELECT NVL(NA03,CU10),CU01||CU02 AS A,CU03,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NVL(CU23,NVL(CU24||CU25||CU26||CU27||CU28,CU29)),NA01 FROM CUSTOMER,NATION WHERE cu10=na01(+) " & strSQL1 & " ORDER BY substr(NA01,1,3),A "
   '2005/6/22 end
'選擇列印代理人
Else
   If Len(Trim(txt1(4))) <> 0 Then
      strSQL1 = strSQL1 + " AND fa10>='" & txt1(4) & "' "
   End If
   If Len(Trim(txt1(5))) <> 0 Then
      strSQL1 = strSQL1 & " AND fa10<='" & txt1(5) & "z' "
   End If
   If Len(txt1(4)) <> 0 Or Len(txt1(5)) <> 0 Then
      pub_QL05 = pub_QL05 & ";代理人" & Label1 & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/10/4
    End If
   'Add By Cheng 2002/05/02
   If Len(Trim(txt1(6))) <= 0 Then
      strSQL1 = strSQL1 & " AND FA24 Is Null "
   Else
      pub_QL05 = pub_QL05 & ";" & Left(Label2, 11) & "是" 'Add By Sindy 2010/10/4
   End If
   'Modify By Cheng 2002/05/02
'   strSQL = "SELECT NVL(NA03,FA10),FA01||FA02 AS A,FA03,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),NVL(FA17,NVL(FA18||FA19||FA20||FA21||FA22,FA23)),NA01 FROM FAGENT,NATION WHERE fa10=na01(+) and FA24 IS NULL " & strSQL1 & " ORDER BY NA01,A "
   '2005/6/22 modify by sonia
   'strSQL = "SELECT NVL(NA03,FA10),FA01||FA02 AS A,FA03,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),NVL(FA17,NVL(FA18||FA19||FA20||FA21||FA22,FA23)),NA01 FROM FAGENT,NATION WHERE fa10=na01(+) " & strSQL1 & " ORDER BY NA01,A "
   'Modify by Morgan 2007/1/24 加FA70
   strSql = "SELECT NVL(NA03,FA10),FA01||FA02 AS A,FA03,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),NVL(FA17,NVL(FA18||FA19||FA20||FA21||FA22||FA70,FA23)),NA01 FROM FAGENT,NATION WHERE fa10=na01(+) " & strSQL1 & " ORDER BY substr(NA01,1,3),A "
   '2005/6/22 end
End If
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/4
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 4
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(0) = StrConv(MidB(StrConv(strTemp(0), vbFromUnicode), 1, 18), vbUnicode)
            strTemp(3) = StrConv(MidB(StrConv(strTemp(3), vbFromUnicode), 1, 24), vbUnicode)
            strTemp(4) = StrConv(MidB(StrConv(strTemp(4), vbFromUnicode), 1, 56), vbUnicode)
            If iPrint > 10000 Then
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            PrintData
            .MoveNext
        Loop
    End With
Else
   InsertQueryLog (0)  'Add By Sindy 2010/10/4
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
Printer.EndDoc
CheckOC
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Private Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6600
Printer.CurrentY = iPrint
If Option1(0).Value Then
   Printer.Print "申請人名單"
Else
   Printer.Print "代理人名單"
End If
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6800
Printer.CurrentY = iPrint
Printer.Print "國籍：" & Format(txt1(4) & " ", "@@@@") & "－" & txt1(5)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "國  籍"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "編  號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "相對編號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "姓  名"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "地  址"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Private Sub PrintData()
For i = 0 To 4
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Private Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 2800
PLeft(2) = 4000
PLeft(3) = 5500
PLeft(4) = 8500
End Sub

Private Sub Form_Load()
MoveFormToCenter Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050318 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
     txt1(0).SetFocus
     txt1_GotFocus (0)
Case 1
     txt1(2).SetFocus
     txt1_GotFocus (2)
Case Else
End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
 KeyAscii = UpperCase(KeyAscii)
 Select Case Index
 Case 6 '是否含不寄雜誌的對象
   If KeyAscii <> 8 And KeyAscii <> 89 Then
      KeyAscii = 0
   End If
 End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 1
      If Len(txt1(Index - 1)) <> 0 Then
         If Left(txt1(Index - 1), 6) <> Left(txt1(Index), 6) Then
             s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
             txt1(Index - 1).SetFocus
             Exit Sub
         End If
      End If
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
      End If
Case 3
      If Len(txt1(Index - 1)) <> 0 Then
         If Left(txt1(Index - 1), 6) <> Left(txt1(Index), 6) Then
             s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
             txt1(Index - 1).SetFocus
             Exit Sub
         End If
      End If
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
      End If
Case 5
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case Else
End Select

End Sub
