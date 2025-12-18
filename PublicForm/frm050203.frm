VERSION 5.00
Begin VB.Form frm050203 
   BorderStyle     =   1  '單線固定
   Caption         =   "未請款明細查詢"
   ClientHeight    =   3975
   ClientLeft      =   735
   ClientTop       =   1680
   ClientWidth     =   4365
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4365
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   12
      Left            =   1200
      TabIndex        =   4
      Text            =   "Text1(12"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   13
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text1(13"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   11
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   12
      Text            =   "1"
      Top             =   2550
      Width           =   495
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   10
      Left            =   1185
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1050
      Width           =   975
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   9
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   8
      Left            =   2520
      TabIndex        =   11
      Text            =   "Text1(8"
      Top             =   2265
      Width           =   975
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   7
      Left            =   2520
      TabIndex        =   9
      Text            =   "Text1(7"
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   972
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   4104
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   6
      Left            =   1200
      TabIndex        =   10
      Text            =   "Text1(6"
      Top             =   2250
      Width           =   975
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   5
      Left            =   1200
      TabIndex        =   8
      Text            =   "Text1(5"
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   4
      Left            =   2520
      TabIndex        =   7
      Text            =   "Text1(4"
      Top             =   1605
      Width           =   975
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   3
      Left            =   1200
      TabIndex        =   6
      Text            =   "Text1(3"
      Top             =   1605
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "回前畫面&U)"
      Height          =   350
      Index           =   0
      Left            =   2556
      TabIndex        =   15
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   1
      Left            =   1776
      TabIndex        =   14
      Top             =   10
      Width           =   756
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   1
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "Text1(1"
      Top             =   768
      Width           =   975
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   2
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   2
      Text            =   "Text1(2"
      Top             =   768
      Width           =   975
   End
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1(0)"
      Top             =   444
      Width           =   2415
   End
   Begin VB.Label Label10 
      Caption         =   "業務區："
      Height          =   180
      Left            =   120
      TabIndex        =   28
      Top             =   1365
      Width           =   975
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   2280
      X2              =   2400
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label9 
      Caption         =   "      其他案件則查詢發文3個月以上案件,依發文日排"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   30
      TabIndex        =   27
      Top             =   3720
      Width           =   4245
   End
   Begin VB.Label Label6 
      Caption         =   "      專利處案件查詢發文3個月以上案件,依智權人員排"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   30
      TabIndex        =   26
      Top             =   3480
      Width           =   4245
   End
   Begin VB.Label Label5 
      Caption         =   "PS : FCP案件僅查詢發文2個月以上案件,依智權人員排"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   30
      TabIndex        =   25
      Top             =   3240
      Width           =   4245
   End
   Begin VB.Label Label4 
      Caption         =   "單據類別：                 (1:收據 / 請款單  2:帳單)"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   2595
      Width           =   4035
   End
   Begin VB.Label Label3 
      Caption         =   "案件性質："
      Height          =   180
      Left            =   120
      TabIndex        =   23
      Top             =   1095
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "查詢順序：                 (1:智權人員  2:發文日  3.承辦人)"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   2925
      Width           =   4215
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2280
      X2              =   2400
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2280
      X2              =   2400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2280
      X2              =   2400
      Y1              =   1725
      Y2              =   1725
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家："
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "申請人："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   2295
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "FC代理人："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   1965
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "發文日："
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   810
      Width           =   975
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2280
      X2              =   2400
      Y1              =   885
      Y2              =   885
   End
   Begin VB.Label Label7 
      Caption         =   "系統類別："
      Height          =   180
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frm050203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/09/22 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/2 日期欄已修改
Option Explicit

Dim bloKeyPreview As Boolean
Dim IntSysTotal As Integer
Dim ArySysName As Variant
Dim strTemp As Variant
Dim strTemp1 As Variant
Dim SysNums() As String
Dim IntSys As Integer
Dim i As Integer, j As Integer, s As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    If bloKeyPreview Then
        KeyAscii = UpperCase(KeyAscii)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm050203 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index))
    Select Case Index
    Case 5, 6
        bloKeyPreview = True
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
'查詢順序欄
If Index = 9 Then
   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End If
'單據類別欄
If Index = 11 Then
   If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim i As Integer
Dim j As Integer
Dim strTp As String 'Add by Amy 2013/11/20 記錄含or不含FMPFFP
    strTp = ""
    Select Case Index
    Case 0
        Unload Me
    Case 1
         If Len(Trim(Text1(0))) = 0 Then
            s = MsgBox("系統類別不可空白", , "USER 輸入錯誤")
            Text1(0).SetFocus
            Exit Sub
         Else
            'Add By Cheng 2002/03/18
            If PUB_CheckKeyInDate(Me.Text1(1)) = -1 Then
               Me.Text1(1).SetFocus
               Text1_GotFocus 1
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.Text1(2)) = -1 Then
               Me.Text1(2).SetFocus
               Text1_GotFocus 2
               Exit Sub
            End If
            If Len(Trim(Text1(2))) = 0 Then
                s = MsgBox("發文日區間不可空白", , "USER 輸入錯誤")
                Text1(1).SetFocus
                Text1_GotFocus (1)
                Exit Sub
            End If
        End If
        If Len(Text1(5)) <> 0 Then
            If Len(Text1(7)) <> 0 Then
                If Left(Text1(5), 6) <> Left(Text1(7), 6) Then
                    s = MsgBox("FC代理人代號前六碼必須相同", , "USER 輸入錯誤")
                    Text1(5).SetFocus
                    Text1(5).SelStart = 0
                    Text1(5).SelLength = Len(Text1(5))
                    Exit Sub
                End If
            Else
                s = MsgBox("FC代理人區間必須輸入", , "USER 輸入錯誤")
                Text1(7).SetFocus
                Exit Sub
            End If
        End If
        If Len(Text1(7)) <> 0 Then
            If Len(Text1(5)) <> 0 Then
                If Left(Text1(5), 6) <> Left(Text1(7), 6) Then
                    s = MsgBox("FC代理人代號前六碼必須相同", , "USER 輸入錯誤")
                    Text1(5).SetFocus
                    Text1(5).SelStart = 0
                    Text1(5).SelLength = Len(Text1(5))
                    Exit Sub
                End If
            Else
                s = MsgBox("FC代理人區間必須輸入", , "USER 輸入錯誤")
                Text1(5).SetFocus
                Exit Sub
            End If
        End If
        If Len(Text1(6)) <> 0 Then
            If Len(Text1(8)) <> 0 Then
                If Left(Text1(6), 6) <> Left(Text1(8), 6) Then
                    s = MsgBox("申請人代號前六碼必須相同", , "USER 輸入錯誤")
                    Text1(6).SetFocus
                    Text1(6).SelStart = 0
                    Text1(6).SelLength = Len(Text1(6))
                    Exit Sub
                End If
            Else
                s = MsgBox("申請人區間必須輸入", , "USER 輸入錯誤")
                Text1(8).SetFocus
                Exit Sub
            End If
        End If
        If Len(Text1(8)) <> 0 Then
            If Len(Text1(6)) <> 0 Then
                If Left(Text1(6), 6) <> Left(Text1(8), 6) Then
                    s = MsgBox("申請人代號前六碼必須相同", , "USER 輸入錯誤")
                    Text1(6).SetFocus
                    Text1(6).SelStart = 0
                    Text1(6).SelLength = Len(Text1(6))
                    Exit Sub
                End If
            Else
                s = MsgBox("申請人區間必須輸入", , "USER 輸入錯誤")
                Text1(6).SetFocus
                Exit Sub
            End If
        End If
        If Len(Me.Text1(11).Text) <= 0 Then
            s = MsgBox("單據類別必須輸入", , "USER 輸入錯誤")
            Text1(11).SetFocus
            Exit Sub
        End If
        If Len(Me.Text1(9).Text) <= 0 Then
            s = MsgBox("查詢順序必須輸入", , "USER 輸入錯誤")
            Text1(9).SetFocus
            Exit Sub
        End If
        'Add by Amy 2013/11/20
        'Modify by Amy 2018/02/22 +if 單據類別為1才需判斷
        If Val(Text1(11)) = 1 Then
            If InStr(1, Text1(0), "FCP") > 0 Then
                '系統類別含 FCP時詢問是否含FMP/FFP
                If MsgBox("是否含FMP、FFP？", vbYesNo + vbDefaultButton2) = vbYes Then
                    strTp = "ADD"
                End If
            ElseIf InStr(1, Text1(0) & ",", "P,") > 0 Or InStr(1, Text1(0) & ",", "CFP") > 0 Then
                '系統類別含 P or CFP時詢問是否剔除FMP/FFP
                If MsgBox("是否剔除FMP、FFP？", vbYesNo + vbDefaultButton2) = vbYes Then
                    strTp = "DEL"
                End If
            End If
        End If
        'end 2018/02/22
        'end 2013/11/20
        'StrMenu
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/9/28 清除查詢印表記錄檔欄位
        Screen.MousePointer = vbHourglass
        Me.Enabled = False
        Me.Hide
        frm050203a.Show
        frm050203a.Tag = strTp 'Add by Amy 2013/11/20
        frm050203a.StrMenu
        Do
        DoEvents
        If bolToEndByNick = True Then Unload Me: Exit Sub
        Loop Until Not frm050203a.Visible
        Unload frm050203a
        Me.Show
        Screen.MousePointer = vbDefault
        Me.Enabled = True
    Case Else
    End Select
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    Cleartxt
    'text1(0) = StrStartSystemByNick
    '89/10/4  邱小姐說要改
     Text1(0) = GetSystemKindByNick

End Sub

Private Sub Cleartxt()
Dim i As Integer
   'Mopdify By Cheng 2002/04/24
'    For i = 0 To 8
    'Modify by Amy 2016/12/07 加業務區
    For i = 0 To 13
        If i <> 11 Then Text1(i) = ""
    Next
    Text1(1) = 920101
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    Select Case Index
    Case 0
    Case 5, 6
        bloKeyPreview = False
    End Select
    Dim strA As String
Dim strB As String
Dim strName As String
    
    Select Case Index
    Case 0
        If Len(Trim(Text1(0))) <> 0 Then
            strTemp = Split(GetSystemKindByNick, ",")
            strTemp1 = Split(Text1(0), ",")
            For i = 0 To UBound(strTemp1)
                s = 0
                For j = 0 To UBound(strTemp)
                    If strTemp1(i) = strTemp(j) Then
                        s = 1
                    End If
                Next j
                If s = 0 Then
                    s = MsgBox(strUserNum + " 沒有 " + strTemp1(i) + " 的使用權限 ", , "USER 權限不足!!!")
                    Text1(0).SetFocus
                    Text1(0).SelStart = 0
                    Text1(0).SelLength = Len(Text1(0))
                    Exit Sub
                End If
            Next i
        End If
        '2013/5/14 add by sonia 電腦中心人員執行時預設發文止日,單據類別及查詢順序
        If UCase(GetStaffDepartment(strUserNum)) = "M51" Then
           Text1(11) = "1"
           'Modify by Amy 2016/12/06 原FCP, 查發文6個月以上
           If InStr(1, Text1(0), "FCP") > 0 Then
              Text1(2) = TransDate(CompDate(2, -1, CompDate(1, -2, Left(strSrvDate(1), 6) & "01")), 1)
              Text1(9) = "1"
           ElseIf InStr(1, Text1(0), "CFP,") > 0 Then
              Text1(2) = TransDate(CompDate(2, -1, CompDate(1, -3, Left(strSrvDate(1), 6) & "01")), 1)
              Text1(9) = "1"
           Else
              Text1(2) = TransDate(CompDate(2, -1, CompDate(1, -3, Left(strSrvDate(1), 6) & "01")), 1)
              Text1(9) = "2"
           End If
        End If
        '2013/5/14 end
    Case 1, 2
         'Modify By Cheng 2002/04/24
        If Len(Text1(Index).Text) > 0 Then
         If Not CheckIsTaiwanDate(Text1(Index)) Then
            s = MsgBox("日期輸入錯誤!!", , "錯誤!!")
            Text1(Index).SetFocus
            Text1_GotFocus (Index)
            Exit Sub
         End If
        End If
        If Index = 2 Then
            If Not nickChgRan(Text1(1), Text1(2), "發文日") Then
                Text1(1).SetFocus
                Text1_GotFocus (1)
               Exit Sub
            End If
        End If
    Case 4
        If Not nickChgRan(Text1(3), Text1(4), "申請國家") Then
                Text1(3).SetFocus
                Text1_GotFocus (3)
           Exit Sub
        End If
    Case 7
        If Not nickChgRan(Text1(5), Text1(7), "FC代理人") Then
                Text1(5).SetFocus
                Text1_GotFocus (5)
           Exit Sub
        End If
    Case 8
        If Not nickChgRan(Text1(6), Text1(8), "申請人") Then
                Text1(6).SetFocus
                Text1_GotFocus (6)
           Exit Sub
        End If
      
    Case Else
    End Select

End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub


