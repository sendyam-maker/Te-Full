VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090625 
   BorderStyle     =   1  '單線固定
   Caption         =   "工程師每週完稿明細"
   ClientHeight    =   2850
   ClientLeft      =   435
   ClientTop       =   405
   ClientWidth     =   4995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4995
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   1380
      MaxLength       =   6
      TabIndex        =   2
      Top             =   870
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   1380
      MaxLength       =   1
      TabIndex        =   1
      Top             =   570
      Width           =   450
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   750
      MaxLength       =   1
      TabIndex        =   0
      Top             =   570
      Width           =   450
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1380
      MaxLength       =   2
      TabIndex        =   10
      Top             =   2370
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   2460
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   11
      Top             =   2370
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1380
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2070
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2460
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2070
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1380
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1770
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2460
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1770
      Width           =   900
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3660
      TabIndex        =   13
      Top             =   45
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2880
      TabIndex        =   12
      Top             =   45
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2460
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1470
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "1"
      Top             =   1470
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1380
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1170
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Left            =   2370
      TabIndex        =   26
      Top             =   900
      Width           =   1815
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3201;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line5 
      X1              =   990
      X2              =   1650
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Label Label1 
      Caption         =   "員工編號："
      Height          =   180
      Index           =   11
      Left            =   180
      TabIndex        =   25
      Top             =   930
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "所別：                            (1:北所 2:中所 3:南所 4:高所 5:其他)"
      Height          =   180
      Index           =   10
      Left            =   180
      TabIndex        =   24
      Top             =   630
      Width           =   4635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作天"
      Height          =   180
      Index           =   9
      Left            =   3900
      TabIndex        =   23
      Top             =   2400
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作天"
      Height          =   180
      Index           =   8
      Left            =   3900
      TabIndex        =   22
      Top             =   2100
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作天"
      Height          =   180
      Index           =   7
      Left            =   3900
      TabIndex        =   21
      Top             =   1830
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作天"
      Height          =   180
      Index           =   6
      Left            =   3900
      TabIndex        =   20
      Top             =   1530
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作天"
      Height          =   180
      Index           =   5
      Left            =   3900
      TabIndex        =   19
      Top             =   1230
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "第四週日期："
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   18
      Top             =   2415
      Width           =   1185
   End
   Begin VB.Line Line4 
      X1              =   1815
      X2              =   2970
      Y1              =   2505
      Y2              =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "第三週日期："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   17
      Top             =   2115
      Width           =   1185
   End
   Begin VB.Line Line2 
      X1              =   1815
      X2              =   2970
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "第二週日期："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   16
      Top             =   1815
      Width           =   1185
   End
   Begin VB.Line Line1 
      X1              =   1815
      X2              =   2970
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Line Line3 
      X1              =   1815
      X2              =   2970
      Y1              =   1605
      Y2              =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "第一週日期："
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   15
      Top             =   1515
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "完稿月份：                             (Ex : 9206)"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   14
      Top             =   1230
      Width           =   3465
   End
End
Attribute VB_Name = "frm090625"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; Label1(12)=>lbl1
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, L(1 To 5, 1 To 12) As Single
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 17) As String, strTemp3 As String, IngL(1 To 3) As Single, tmpnickG As Integer, k As Integer
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, StrTemp99(0 To 7) As String, StrTemp7(0 To 9) As String
Dim allG As Integer, Hightmp As Single, Middletmp As Single, Lowtmp As Single, Avgtmp As Single, Sumtmp As Single, SumCount As Single
Dim RsTmpNick As New ADODB.Recordset

Private Sub cmdOK_Click(Index As Integer)
Dim ii As Integer

Select Case Index
Case 0
    If Me.txt1(10).Text <> "" Then
        If RunNick(txt1(9), txt1(10)) Then
            txt1(9).SetFocus
            txt1_GotFocus (9)
            Exit Sub
        End If
    End If
    If Me.txt1(11).Text <> "" Then
        Me.lbl1.Caption = GetStaffName(Me.txt1(11).Text, True)
        If Me.lbl1.Caption = "" Then
            MsgBox "員工編號輸入錯誤!!!", vbExclamation + vbOKOnly
            txt1(11).SetFocus
            txt1_GotFocus (11)
            Exit Sub
        End If
    Else
        Me.lbl1.Caption = ""
    End If
    If Len(txt1(0)) = 0 Then
        s = MsgBox("完稿月份不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        txt1_GotFocus 0
        Exit Sub
    End If
    If PUB_CheckKeyInYYMM(Me.txt1(0)) = -1 Then
       Me.txt1(0).SetFocus
       txt1_GotFocus 0
       Exit Sub
    End If
    For ii = 1 To 8
        If Me.txt1(ii).Text = "" Then
            MsgBox "請輸入日期!!!", vbExclamation + vbOKOnly
            Me.txt1(ii).SetFocus
            txt1_GotFocus ii
            Exit Sub
        End If
    Next ii
    Me.Hide: DoEvents
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/20 清除查詢印表記錄檔欄位
    frm090625_1.Show
Case 1 '回前畫面
    Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Me.txt1(0).Text = Left(strSrvDate(1), 6) - 191100
Me.txt1(8).Text = PUB_GetMonthDays(Left(Me.txt1(0).Text + 191100, 4), Val(Mid(Me.txt1(0).Text + 191100, 5, 2)))
SetDate
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090625 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
Select Case Index
Case 9, 10
    Select Case KeyAscii
    Case 49, 50, 51, 52, 53, 8
    Case Else
        KeyAscii = 0
    End Select
End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0 '考核月份
   If PUB_CheckKeyInYYMM(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
    If Me.txt1(0).Text <> "" Then
        Me.txt1(8).Text = PUB_GetMonthDays(Left(Me.txt1(0).Text + 191100, 4), Val(Mid(Me.txt1(0).Text + 191100, 5, 2)))
        SetDate
    End If
Case 2, 4, 6, 8 '日
    If Me.txt1(Index).Text <> "" Then
        If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
            Exit Sub
        End If
    End If
    If Me.txt1(Index - 1).Text <> "" And Me.txt1(Index).Text <> "" Then
        Me.Label1(5 + Index / 2).Caption = GetWorkDay(DBDATE(Val(Me.txt1(0).Text) & Format(Me.txt1(Index).Text, "00")), DBDATE(Val(Me.txt1(0).Text) & Format(Me.txt1(Index - 1).Text, "00")))
    End If
Case 10 '所別
    If Me.txt1(Index).Text <> "" Then
        If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
            Exit Sub
        End If
    End If
Case Else
End Select
End Sub

'預設期間
Private Sub SetDate()
Dim intWeekDay As Integer

intWeekDay = Weekday(ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00")))
'若1號小於星期三
If intWeekDay < 4 Then
    Me.txt1(2).Text = Day(DateAdd("d", ((7 - intWeekDay)), ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00"))))
'若1號大於等於星期三
Else
    Me.txt1(2).Text = Day(DateAdd("d", ((7 - intWeekDay) + 7), ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(1).Text, "00"))))
End If
Me.txt1(3).Text = Day(DateAdd("d", 1, ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(2).Text, "00"))))
Me.txt1(4).Text = Day(DateAdd("d", 6, ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(3).Text, "00"))))
Me.txt1(5).Text = Day(DateAdd("d", 1, ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(4).Text, "00"))))
Me.txt1(6).Text = Day(DateAdd("d", 6, ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(5).Text, "00"))))
Me.txt1(7).Text = Day(DateAdd("d", 1, ChangeTStringToWDateString(Me.txt1(0).Text & Format(Me.txt1(6).Text, "00"))))
txt1_LostFocus 2
txt1_LostFocus 4
txt1_LostFocus 6
txt1_LostFocus 8
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
        
    Select Case Index
    Case 11 '員工編號
        If Me.txt1(Index).Text <> "" Then
            Me.lbl1.Caption = GetStaffName(Me.txt1(Index).Text, True)
            If Me.lbl1.Caption = "" Then
                MsgBox "員工編號輸入錯誤!!!", vbExclamation + vbOKOnly
                Cancel = True
            End If
        Else
            Me.lbl1.Caption = ""
        End If
    End Select
    If Cancel = True Then txt1_GotFocus Index
End Sub
