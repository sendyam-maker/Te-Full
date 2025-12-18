VERSION 5.00
Begin VB.Form frm12040142 
   BorderStyle     =   1  '單線固定
   Caption         =   "收文未發文明細表"
   ClientHeight    =   5856
   ClientLeft      =   3168
   ClientTop       =   1200
   ClientWidth     =   6192
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5856
   ScaleWidth      =   6192
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   18
      Left            =   1680
      TabIndex        =   8
      Top             =   2320
      Width           =   4215
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   17
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1360
      Width           =   390
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   16
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2000
      Width           =   390
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   15
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1680
      Width           =   390
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   13
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   13
      Text            =   "111112"
      Top             =   3300
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   14
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   14
      Text            =   "111112"
      Top             =   3300
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   11
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   11
      Text            =   "111112"
      Top             =   3000
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   12
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   12
      Text            =   "111112"
      Top             =   3000
      Width           =   795
   End
   Begin VB.TextBox txtOver 
      Height          =   270
      Left            =   4905
      MaxLength       =   1
      TabIndex        =   19
      Top             =   3900
      Width           =   390
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   5325
      TabIndex        =   21
      Top             =   105
      Width           =   795
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4545
      TabIndex        =   20
      Top             =   105
      Width           =   750
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   10
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   18
      Top             =   3915
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   9
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   17
      Top             =   3915
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   8
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   16
      Top             =   3610
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   7
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   15
      Top             =   3610
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   6
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   10
      Top             =   2700
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   9
      Text            =   "850101"
      Top             =   2700
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1060
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   3
      Top             =   760
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   2
      Top             =   460
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   1
      Top             =   460
      Width           =   795
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   150
      Width           =   3270
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "自104/10/1商標改晚上批次跑"
      Height          =   180
      Left            =   3600
      TabIndex        =   50
      Top             =   3045
      Width           =   2250
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "E-MAIL副總收件人:"
      Height          =   180
      Left            =   0
      TabIndex        =   49
      Top             =   2362
      Width           =   1665
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "( 1: E-MAIL 2: 報表)"
      Height          =   180
      Left            =   1755
      TabIndex        =   48
      Top             =   1410
      Width           =   1530
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "輸出方式:"
      Height          =   180
      Left            =   360
      TabIndex        =   47
      Top             =   1405
      Width           =   765
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "定期跑報表暫不跑法務資料"
      Height          =   180
      Left            =   3600
      TabIndex        =   46
      Top             =   3345
      Width           =   2160
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "       　例：商標之第一期註冊費或專利之實體審查"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   330
      TabIndex        =   45
      Top             =   5580
      Width           =   3915
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "       6. 選特定區或特定人時，同時會印出期限未知但不可辦之案件"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   330
      TabIndex        =   44
      Top             =   5400
      Width           =   5175
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "       4. 專利8個月、商標法務5個月未處理交副總"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   330
      TabIndex        =   43
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "( 1: 智權人員 2 : 部門主管 3: 副總 4 : 自訂)"
      Height          =   180
      Left            =   1755
      TabIndex        =   42
      Top             =   2040
      Width           =   3270
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   " 收文管制時段:"
      Height          =   180
      Left            =   -30
      TabIndex        =   41
      Top             =   2045
      Width           =   1170
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "列印對象:"
      Height          =   180
      Left            =   360
      TabIndex        =   40
      Top             =   1725
      Width           =   765
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "( 1: 智權人員 2: 承辦人)"
      Height          =   180
      Left            =   1755
      TabIndex        =   39
      Top             =   1725
      Width           =   1830
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "法務收文日期:"
      Height          =   180
      Left            =   0
      TabIndex        =   38
      Top             =   3345
      Width           =   1125
   End
   Begin VB.Line Line6 
      X1              =   2160
      X2              =   2400
      Y1              =   3428
      Y2              =   3428
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "商標收文日期:"
      Height          =   180
      Left            =   0
      TabIndex        =   37
      Top             =   3045
      Width           =   1125
   End
   Begin VB.Line Line5 
      X1              =   2160
      X2              =   2400
      Y1              =   3128
      Y2              =   3128
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "       5. 月底先印給專業部過濾資料"
      ForeColor       =   &H0000C000&
      Height          =   180
      Left            =   330
      TabIndex        =   36
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "       3. 專利7個月、商標法務4個月未處理交各區主管"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   330
      TabIndex        =   35
      Top             =   4680
      Width           =   4095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "( N:不含)"
      Height          =   180
      Left            =   5445
      TabIndex        =   34
      Top             =   3930
      Width           =   690
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "是否含未到期資料:"
      Height          =   180
      Left            =   3360
      TabIndex        =   33
      Top             =   3930
      Width           =   1485
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "       2. 專利6個月、商標法務3個月未處理需銷案"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   330
      TabIndex        =   32
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "PS : 1. 不含 FC案件"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   330
      TabIndex        =   31
      Top             =   4200
      Width           =   1470
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   1
      Left            =   2160
      TabIndex        =   30
      Top             =   1102
      Width           =   1755
   End
   Begin VB.Line Line4 
      X1              =   2160
      X2              =   2400
      Y1              =   3738
      Y2              =   3738
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Left            =   360
      TabIndex        =   29
      Top             =   3915
      Width           =   765
   End
   Begin VB.Line Line3 
      X1              =   2160
      X2              =   2400
      Y1              =   4043
      Y2              =   4043
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   360
      TabIndex        =   28
      Top             =   3655
      Width           =   765
   End
   Begin VB.Line Line2 
      X1              =   2160
      X2              =   2400
      Y1              =   2828
      Y2              =   2828
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利收文日期:"
      Height          =   180
      Left            =   0
      TabIndex        =   27
      Top             =   2745
      Width           =   1125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   540
      TabIndex        =   26
      Top             =   1102
      Width           =   585
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   0
      Left            =   2145
      TabIndex        =   25
      Top             =   802
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   360
      TabIndex        =   24
      Top             =   802
      Width           =   765
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   2400
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "業務區:"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   540
      TabIndex        =   23
      Top             =   502
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別:"
      Height          =   180
      Left            =   360
      TabIndex        =   22
      Top             =   210
      Width           =   765
   End
End
Attribute VB_Name = "frm12040142"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
'2005/7/5整理
Option Explicit
Dim strSql As String, strTemp1 As Variant, strTemp2 As Variant, StrTest1 As String, StrTest2 As String, i As Integer, j As Integer, s As Integer
Dim PLeft(0 To 15) As Integer, k As Integer, TmpArea As String, iLine As Integer, Page As Integer
Dim strTemp3(0 To 15) As String, iPrint As Integer
Dim StrTest3 As String, Day1 As String, Day2 As String, StrTemp4 As String
Dim St As String, iK As Integer, iTatle As Integer
'Add By Cheng 2003/02/26
Dim iKK As Integer '業務區合計
'add by nick 2004/10/19
Dim StrTest4 As String, StrTest5 As String, StrTest6 As String, StrTest7 As String, StrTest31 As String
'Added by Lydia 2015/07/23 輸出方式:E-MAIL
Dim eFile As Integer, eFilename As String
Dim strPath As String
Dim strFileN As String
Dim mailFList As String, m_SpecMan As String
Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     If Len(Txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        Txt1(0).SetFocus
        Exit Sub
     Else
        'edit by nick 2004/10/19
        'If Len(txt1(6)) = 0  Then
        If Len(Txt1(6)) = 0 Or Len(Txt1(5)) = 0 Then
            'edit by nick 2004/10/19
            's = MsgBox("收文天數不可空白!!", , "USER 輸入錯誤")
            s = MsgBox("專利收文日期區間不可空白!!", , "USER 輸入錯誤")
            Txt1(5).SetFocus
            txt1_GotFocus (5)
            Exit Sub
        Else
            'add by nick 2004/10/19
            If Len(Txt1(11)) = 0 Or Len(Txt1(12)) = 0 Then
                s = MsgBox("商標收文日期區間不可空白!!", , "USER 輸入錯誤")
                Txt1(11).SetFocus
                txt1_GotFocus (11)
                Exit Sub
            Else
                If Len(Txt1(13)) = 0 Or Len(Txt1(14)) = 0 Then
                    s = MsgBox("法務收文日期區間不可空白!!", , "USER 輸入錯誤")
                    Txt1(13).SetFocus
                    txt1_GotFocus (13)
                    Exit Sub
                Else
                    If Len(Txt1(15)) = 0 Then
                        s = MsgBox("列印順序不可空白!!", , "USER 輸入錯誤")
                        Txt1(15).SetFocus
                        txt1_GotFocus (15)
                        Exit Sub
                    'Added by Lydia 2015/07/28 + txt1(17),txt1(18)
                    ElseIf Len(Txt1(17)) = 0 Then
                        s = MsgBox("輸出方式不可空白!!", , "USER 輸入錯誤")
                        Txt1(17).SetFocus
                        txt1_GotFocus (17)
                        Exit Sub
                    ElseIf Txt1(17) = "1" And Txt1(16) = "3" And Len(Txt1(18)) = 0 Then
                        s = MsgBox("E-MAIL副總收件人不可空白!!", , "USER 輸入錯誤")
                        Txt1(18).SetFocus
                        txt1_GotFocus (18)
                        Exit Sub
                    'end 2015/07/23
                    Else
                        Screen.MousePointer = vbHourglass
                        Me.Enabled = False
                        StrMenu1
                        Me.Enabled = True
                        Screen.MousePointer = vbDefault
                    End If
                End If
            End If
        End If
    End If
Case 1
     Unload Me
Case Else
End Select
End Sub
'Remove by Lydia 2015/07/23
'Add By Cheng 2003/01/09
'Sub StrPrintDocTotal()
'Dim strTempName As String '代理人名稱
'
'GetPrintLeft
'iLine = 1
'Page = 1
'StrPrintTital TmpArea, str(Page)
'iPrint = 2700
'iTatle = 0       ' 總數
'With adoRecordset
'    .MoveFirst
'    Do While .EOF = False
'        For j = 0 To 2
'            If Not IsNull(.Fields(j)) Then
'                strTemp3(j) = .Fields(j)
'            Else
'                strTemp3(j) = ""
'            End If
'        Next j
'        Printer.CurrentX = PLeft(0)
'        Printer.CurrentY = iPrint
'        Printer.Print strTemp3(1)
'        Printer.CurrentX = PLeft(2) + 750 - TextWidth(Format(strTemp3(2), "##0") & " 筆")
'        Printer.CurrentY = iPrint
'        Printer.Print Format(strTemp3(2), "##0") & " 筆"
'        iTatle = iTatle + Val(strTemp3(2))
'        .MoveNext
'        If (iLine Mod 26 = 0) Or iPrint >= 10000 Then
'            iPrint = iPrint + 300
'            Printer.NewPage
'            Page = Page + 1
'            StrPrintTital TmpArea, str(Page)
'            'iPrint = 2400
'            iPrint = 2100
'            iLine = 0
'        End If
'        iLine = iLine + 1
'        iPrint = iPrint + 300
'    Loop
'End With
''合計
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Print String(200, "-")
'iPrint = iPrint + 300
'Printer.CurrentX = PLeft(2) + 750 - TextWidth("合計：共 " & Format(iTatle, "##0") & " 筆")
'Printer.CurrentY = iPrint
'Printer.Print "合計：共 " & Format(iTatle, "##0") & " 筆"
'Printer.EndDoc
'ShowPrintOk
'CheckOC
'End Sub

Sub StrPrintTital(ByRef Area As String, ByRef Page As String)
GetPrintLeft
k = 200
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 7940 - (Printer.TextWidth("收文未發文明細表") / 2)
Printer.CurrentY = i
Printer.Print "收文未發文明細表"
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = 0
Printer.CurrentY = k + 500
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 7000
Printer.CurrentY = k + 500
Printer.Print "列印順序：" & IIf(Trim(Txt1(15)) = "1", "智權人員", "承辦人")
Printer.CurrentX = PLeft(2)
Printer.CurrentY = k + 500
Select Case Txt1(16)
   Case "1"
      '2012/9/27 modify by sonia
      'Printer.Print "管制時段：個人"
      'Remove by Lydia 2015/07/29 只顯示-個人
'      If txt1(15) = "1" Then
'         Printer.Print "管制時段：個人(請交個人核對資料)"
'      Else
         Printer.Print "管制時段：個人"
     ' End If
      '2012/9/27 end
   Case "2"
      'Modified by Lydia 2015/07/23
      'Printer.Print "管制時段：區主管"
      Printer.Print "管制時段：部門主管"
   Case "3"
      Printer.Print "管制時段：副總"
   Case "4"
      Printer.Print ""
End Select
Printer.CurrentX = 13000
Printer.CurrentY = k + 500
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.CurrentX = 0
Printer.CurrentY = k + 800
Printer.Print Area
Printer.CurrentX = 0
Printer.CurrentY = k + 1100
'modify by sonia 2015/10/1
'Printer.Print "專利收文日期：" & ChangeTStringToTDateString(txt1(5)) & "-" & ChangeTStringToTDateString(txt1(6)) & "  商標收文日期：" & ChangeTStringToTDateString(txt1(11)) & "-" & ChangeTStringToTDateString(txt1(12)) & "  法務收文日期：" & ChangeTStringToTDateString(txt1(13)) & "-" & ChangeTStringToTDateString(txt1(14))
Printer.Print "專利收文日期：" & ChangeTStringToTDateString(Txt1(5)) & "-" & ChangeTStringToTDateString(Txt1(6)) & IIf(Trim(Txt1(11)) = "111112", "", "  商標收文日期：" & ChangeTStringToTDateString(Txt1(11)) & "-" & ChangeTStringToTDateString(Txt1(12))) & IIf(Trim(Txt1(11)) = "111112", "", "  法務收文日期：" & ChangeTStringToTDateString(Txt1(13)) & "-" & ChangeTStringToTDateString(Txt1(14)))
'end 2015/10/1
Printer.CurrentX = 13000
Printer.CurrentY = k + 1100
Printer.Print "頁    次：" & Page
Printer.CurrentX = 0
Printer.CurrentY = k + 1400
Printer.Print String(200, "-")
Printer.Font.Underline = True
Printer.CurrentX = PLeft(0)
Printer.CurrentY = k + 1700
Printer.Print IIf(Trim(Txt1(15)) = "1", "智權人員", "承辦人")
Printer.CurrentX = PLeft(1)
Printer.CurrentY = k + 1700
Printer.Print "收文日"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = k + 1700
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = k + 1700
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = k + 1700
Printer.Print "申請人"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = k + 1700
Printer.Print "本所期限"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = k + 1700
Printer.Print "法定期限"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = k + 1700
Printer.Print "種類"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = k + 1700
Printer.Print "案件性質"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = k + 1700
Printer.Print IIf(Trim(Txt1(15)) = "1", "承辦人", "智權人員")
Printer.CurrentX = PLeft(10)
Printer.CurrentY = k + 1700
Printer.Print "申請國家"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = k + 1700
Printer.Print "未收金額"
'2010/4/1 ADD BY SONIA 加分所號
Printer.CurrentX = PLeft(12)
Printer.CurrentY = k + 1700
Printer.Print "分所號"
'2010/4/1 END
Printer.Font.Underline = False
Printer.CurrentX = 0
Printer.CurrentY = k + 2000
Printer.Print String(200, "-")
End Sub
'Remove by Lydia 2015/07/23
'Sub StrPrintEnd()
'End Sub
Sub GetPrintLeft()
   Erase PLeft
   PLeft(0) = 0 '智權人員
   PLeft(1) = 1500 - 500 '收文日
   PLeft(2) = 2400 '本所案號
   PLeft(3) = 4500 '案件名稱
   PLeft(4) = 5700 '申請人
   PLeft(5) = 7300 '本所期限
   PLeft(6) = 8500 '法定期限
   PLeft(7) = 9700 '種類
   PLeft(8) = 10900 '案件性質
   PLeft(9) = 12100 '承辦人
   PLeft(10) = 13300 '申請國家
   PLeft(11) = 14500  '未收金額
   PLeft(12) = 15800  '分所號   2010/4/1 ADD BY SONIA
End Sub
Private Sub Form_Load()
   MoveFormToCenter Me
   'modify by sonia 2015/10/1 自2015/10/1商標改晚上批次跑故改預設系統類別僅專利
   'strTemp1 = Split(UCase(GetSystemKindByNick), ",")
   'For i = 0 To UBound(strTemp1)
   '    If strTemp1(i) <> "FCP" And strTemp1(i) <> "FCT" And strTemp1(i) <> "FG" Then
   '        txt1(0) = txt1(0) + strTemp1(i) + ","
   '    End If
   'Next i
   Txt1(0) = "P,PS,CFP,CPS"
   'end 2015/10/1
   Txt1(6) = ChangeWDateStringToTString(DateAdd("m", -6, ChangeWStringToWDateString(strSrvDate(1))))
   'txt1(12) = ChangeWDateStringToTString(DateAdd("m", -3, ChangeWStringToWDateString(strSrvDate(1))))  'cancel by sonia 2015/10/1 商標改晚上批次跑
   'txt1(14) = ChangeWDateStringToTString(DateAdd("m", -6, ChangeWStringToWDateString(strSrvDate(1))))  '2013/5/14 cancel by sonia
   
   txtOver = "N"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040142 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   Txt1(Index).SelStart = 0
   Txt1(Index).SelLength = Len(Txt1(Index))
End Sub

'Add By Cheng 2003/02/24
'列印對象為智權人員
Sub StrMenu1()

'Add by Morgan 2003/12/08
'控制是否列印 cc1.cp01='T' AND cc1.cp10='715' 的資料
Dim bolNoSkip As Boolean, strSQLSkip As String
Dim strSQLSkipP As String
'add by nick 2004/12/28
Dim StrTest99 As String
Dim StrTest98 As String
bolNoSkip = False

Screen.MousePointer = vbHourglass
StrTest1 = ""       '專利
StrTest2 = ""       '商標
StrTest3 = ""       '法務
'add by nick 2004/10/19
StrTest31 = ""     '顧問
StrTest4 = ""       '服務專利
StrTest5 = " and decode(cc1.cp01,'S',sp09,'001') <> '000' "       '服務商標
StrTest6 = ""       '服務法務
StrTest7 = ""       '服務共用條件
'add by nick 2004/12/28
'modify by sonia 2016/9/13 cp27 IS NULL->cp158=0,cp57 IS NULL->cp159=0
'Modified by Lydia 2016/12/21 +排除D類收文 and substr(cc2.cp09,1,1) <> 'D'
StrTest99 = " and cc2.cp14 IS NOT NULL AND cc2.cp158=0 AND cc2.cp159=0 and substr(cc2.cp09,1,1) <> 'D' "
StrTest98 = " and cc2.cp14 IS NOT NULL AND cc2.cp158=0 AND cc2.cp159=0 and substr(cc2.cp09,1,1) <> 'D' "
If Len(Txt1(0)) <> 0 Then
   StrTest1 = StrTest1 & " AND cc1.cp01 IN (" & SQLGrpStr2(Txt1(0), 1) & ") "
   StrTest2 = StrTest2 & " AND cc1.cp01 IN (" & SQLGrpStr2(Txt1(0), 2) & ") "
   'edit by nick 2004/10/19
   'StrTest3 = StrTest3 & " AND cc1.cp01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   StrTest3 = StrTest3 & " AND cc1.cp01 IN (" & SQLGrpStr2(Txt1(0), 3) & ") "
   'add by nick 2004/10/19
   StrTest31 = StrTest31 & " AND cc1.cp01 IN (" & SQLGrpStr2(Txt1(0), 4) & ") "
   StrTest4 = StrTest4 & " AND cc1.cp01 IN (" & SQLGrpStr2(Txt1(0), 5) & ") "
   StrTest5 = StrTest5 & " AND cc1.cp01 IN (" & SQLGrpStr2(Txt1(0), 6) & ") "
   StrTest6 = StrTest6 & " AND cc1.cp01 IN (" & SQLGrpStr2(Txt1(0), 7) & ") "
End If
'92.4.2 MODIFY BY SONIA 林錦山, 楊文照己調區, 以前收文案件轉至新業務區列印
'故cc1.cp12改為S2.ST15
If Len(Txt1(1)) <> 0 Then
    StrTest1 = StrTest1 + " AND S2.ST15>='" & Txt1(1) & "' "
    StrTest2 = StrTest2 + " AND S2.ST15>='" & Txt1(1) & "' "
    StrTest3 = StrTest3 + " AND S2.ST15>='" & Txt1(1) & "' "
    '93.11.2 ADD BY SONIA
    StrTest31 = StrTest31 + " AND S2.ST15>='" & Txt1(1) & "' "
    'add by nick 25004/10/19
    StrTest7 = StrTest7 + " AND S2.ST15>='" & Txt1(1) & "' "
   'Add by Morgan 2003/12/08
   bolNoSkip = True
'2005/7/7 add by sonia
Else
   'Modified by Lydia 2015/07/23 改為輸出-報表和時段-副總
   'If txt1(16) >= "2" And txt1(16) <= "3" Then
   If Txt1(17) = "2" And Txt1(16) = "3" Then
      StrTest1 = StrTest1 + " AND S2.ST15>='S' "
      StrTest2 = StrTest2 + " AND S2.ST15>='S' "
      StrTest3 = StrTest3 + " AND S2.ST15>='S' "
      StrTest31 = StrTest31 + " AND S2.ST15>='S' "
      StrTest7 = StrTest7 + " AND S2.ST15>='S' "
   End If
'2005/7/7 END
End If
If Len(Txt1(2)) <> 0 Then
    StrTest1 = StrTest1 + " AND S2.ST15<='" & Txt1(2) & "' "
    StrTest2 = StrTest2 + " AND S2.ST15<='" & Txt1(2) & "' "
    StrTest3 = StrTest3 + " AND S2.ST15<='" & Txt1(2) & "' "
    '93.11.2 ADD BY SONIA
    StrTest31 = StrTest31 + " AND S2.ST15<='" & Txt1(2) & "' "
    'add by nick 200410/19
    StrTest7 = StrTest7 + " AND S2.ST15<='" & Txt1(2) & "' "
   'Add by Morgan 2003/12/08
   bolNoSkip = True
'2005/7/7 add by sonia
Else
   'Modified by Lydia 2015/07/23 改為輸出-報表和時段-副總
   'If txt1(16) >= "2" And txt1(16) <= "3" Then
   If Txt1(17) = "2" And Txt1(16) = "3" Then
      StrTest1 = StrTest1 + " AND S2.ST15<='S99' "
      StrTest2 = StrTest2 + " AND S2.ST15<='S99' "
      StrTest3 = StrTest3 + " AND S2.ST15<='S99' "
      StrTest31 = StrTest31 + " AND S2.ST15<='S99' "
      StrTest7 = StrTest7 + " AND S2.ST15<='S99' "
   End If
'2005/7/7 END
End If
If Len(Txt1(3)) <> 0 Then
    StrTest1 = StrTest1 + " AND cc1.cp13='" & Txt1(3) & "' "
    StrTest2 = StrTest2 + " AND cc1.cp13='" & Txt1(3) & "' "
    StrTest3 = StrTest3 + " AND cc1.cp13='" & Txt1(3) & "' "
    '93.11.2 ADD BY SONIA
    StrTest31 = StrTest31 + " AND cc1.cp13='" & Txt1(3) & "' "
    'add by nick 2004/10/19
    StrTest7 = StrTest7 + " AND cc1.cp13='" & Txt1(3) & "' "
   'Add by Morgan 2003/12/08
   bolNoSkip = True
   'add by nick 2004/12/28
   StrTest99 = StrTest99 & " and cc2.cp13='" & Txt1(3) & "' "
   StrTest98 = StrTest98 & " and cc2.cp13='" & Txt1(3) & "' "
End If
If Len(Txt1(4)) <> 0 Then
    StrTest1 = StrTest1 + " AND cc1.cp14='" & Txt1(4) & "' "
    StrTest2 = StrTest2 + " AND cc1.cp14='" & Txt1(4) & "' "
    StrTest3 = StrTest3 + " AND cc1.cp14='" & Txt1(4) & "' "
    '93.11.2 ADD BY SONIA
    StrTest31 = StrTest31 + " AND cc1.cp14='" & Txt1(4) & "' "
    'add by nick 2004/10/19
    StrTest7 = StrTest7 + " AND cc1.cp14='" & Txt1(4) & "' "
   'Add by Morgan 2003/12/08
   bolNoSkip = True
   'add by nick 2004/12/28
   StrTest99 = StrTest99 & " and cc2.cp14='" & Txt1(4) & "' "
   StrTest98 = StrTest98 & " and cc2.cp14='" & Txt1(4) & "' "
End If

If Len(Txt1(7)) <> 0 Then
    StrTest1 = StrTest1 + " AND cc1.cp10>='" & Txt1(7) & "' "
    StrTest2 = StrTest2 + " AND cc1.cp10>='" & Txt1(7) & "' "
    StrTest3 = StrTest3 + " AND cc1.cp10>='" & Txt1(7) & "' "
    '93.11.2 ADD BY SONIA
    StrTest31 = StrTest31 + " AND cc1.cp10>='" & Txt1(7) & "' "
    'add by nick 2004/10/19
    StrTest7 = StrTest7 + " AND cc1.cp10>='" & Txt1(7) & "' "
   'add by nick 2004/12/28
   StrTest99 = StrTest99 & " and cc2.cp10>='" & Txt1(7) & "' "
   StrTest98 = StrTest98 & " and cc2.cp10>='" & Txt1(7) & "' "
End If
If Len(Txt1(8)) <> 0 Then
    StrTest1 = StrTest1 + " AND cc1.cp10<='" & Txt1(8) & "' "
    StrTest2 = StrTest2 + " AND cc1.cp10<='" & Txt1(8) & "' "
    StrTest3 = StrTest3 + " AND cc1.cp10<='" & Txt1(8) & "' "
    '93.11.2 ADD BY SONIA
    StrTest31 = StrTest31 + " AND cc1.cp10<='" & Txt1(8) & "' "
    'add by nick 2004/10/19
    StrTest7 = StrTest7 + " AND cc1.cp10<='" & Txt1(8) & "' "
   'add by nick 2004/12/28
   StrTest99 = StrTest99 & " and cc2.cp10<='" & Txt1(8) & "' "
   StrTest98 = StrTest98 & " and cc2.cp10<='" & Txt1(8) & "' "
End If
If Len(Txt1(9)) <> 0 Then
    StrTest1 = StrTest1 + " AND PA09>='" & Txt1(9) & "' "
    StrTest2 = StrTest2 + " AND TM10>='" & Txt1(9) & "' "
    'edit by nick 2004/10/19
    'StrTest3 = StrTest3 + " AND SP09>='" & txt1(9) & "' "
    StrTest3 = StrTest3 + " AND LC15>='" & Txt1(9) & "' "
    '93.11.2 ADD BY SONIA
    If Txt1(9) = "000" Then StrTest31 = StrTest31 + " AND HC01='LA' "
    'add by nick 2004/10/19
    StrTest7 = StrTest7 + " AND SP09>='" & Txt1(9) & "' "
End If
If Len(Txt1(10)) <> 0 Then
    StrTest1 = StrTest1 + " AND PA09<='" & Txt1(10) & "' "
    StrTest2 = StrTest2 + " AND TM10<='" & Txt1(10) & "' "
    'edit by nick 2004/10/19
    'StrTest3 = StrTest3 + " AND SP09<='" & txt1(10) & "' "
    StrTest3 = StrTest3 + " AND LC15<='" & Txt1(10) & "' "
    '93.11.2 ADD BY SONIA
    If Txt1(10) = "000" Then StrTest31 = StrTest31 + " AND HC01='LA' "
    'add by nick 2004/10/19
    StrTest7 = StrTest7 + " AND SP09<='" & Txt1(10) & "' "
End If
'add by nick 2004/10/19  改成收文日期
If Len(Txt1(5)) <> 0 Then
    StrTest1 = StrTest1 + " AND cc1.cp05>='" & ChangeTStringToWString(Txt1(5)) & "' "
    StrTest4 = StrTest4 + " AND cc1.cp05>='" & ChangeTStringToWString(Txt1(5)) & "' "
   'add by nick 2004/12/28
   StrTest98 = StrTest98 & " AND cc2.cp05>='" & ChangeTStringToWString(Txt1(5)) & "' "
End If
If Len(Txt1(6)) <> 0 Then
    StrTest1 = StrTest1 + " AND cc1.cp05<='" & ChangeTStringToWString(Txt1(6)) & "' "
    StrTest4 = StrTest4 + " AND cc1.cp05<='" & ChangeTStringToWString(Txt1(6)) & "' "
   'add by nick 2004/12/28
   StrTest98 = StrTest98 & " AND cc2.cp05<='" & ChangeTStringToWString(Txt1(6)) & "' "
End If
If Len(Txt1(11)) <> 0 Then
    StrTest2 = StrTest2 + " AND cc1.cp05>='" & ChangeTStringToWString(Txt1(11)) & "' "
    StrTest5 = StrTest5 + " AND cc1.cp05>='" & ChangeTStringToWString(Txt1(11)) & "' "
   'add by nick 2004/12/28
   StrTest99 = StrTest99 & " AND cc2.cp05>='" & ChangeTStringToWString(Txt1(11)) & "' "
End If
If Len(Txt1(12)) <> 0 Then
    StrTest2 = StrTest2 + " AND cc1.cp05<='" & ChangeTStringToWString(Txt1(12)) & "' "
    StrTest5 = StrTest5 + " AND cc1.cp05<='" & ChangeTStringToWString(Txt1(12)) & "' "
   'add by nick 2004/12/28
   StrTest99 = StrTest99 & " AND cc2.cp05<='" & ChangeTStringToWString(Txt1(12)) & "' "
End If
If Len(Txt1(13)) <> 0 Then
    StrTest3 = StrTest3 + " AND cc1.cp05>='" & ChangeTStringToWString(Txt1(13)) & "' "
    StrTest31 = StrTest31 + " AND cc1.cp05>='" & ChangeTStringToWString(Txt1(13)) & "' "
    StrTest6 = StrTest6 + " AND cc1.cp05>='" & ChangeTStringToWString(Txt1(13)) & "' "
End If
If Len(Txt1(14)) <> 0 Then
    StrTest3 = StrTest3 + " AND cc1.cp05<='" & ChangeTStringToWString(Txt1(14)) & "' "
    StrTest31 = StrTest31 + " AND cc1.cp05<='" & ChangeTStringToWString(Txt1(14)) & "' "
    StrTest6 = StrTest6 + " AND cc1.cp05<='" & ChangeTStringToWString(Txt1(14)) & "' "
End If

'add by nick 2004/10/19 重新定義字串內容
StrTest4 = Mid(StrTest4, 5)
StrTest5 = Mid(StrTest5, 5)
StrTest6 = Mid(StrTest6, 5)

'Add By Cheng 2003/04/09
'若有本所期限者, 不能大於系統日
'Modify by Morgan 2004/4/7
'加未到期控制條件
If txtOver = "N" Then
   '2011/1/31 modify by sonia 未到期控制不限制新申請案及答辯(專利)案
   'StrTest1 = StrTest1 & " AND ( cc1.cp06 Is Null Or cc1.cp06<=" & ServerDate & " ) "
   '2015/08/31 未到期不出現原控制剔除新申請案及答辯(專利)案，請再加入非國外部收文(CP12第一碼非'F')者，即國外部收文者只要未到期都不出現
  ' StrTest1 = StrTest1 & " AND (((cc1.cp06 Is Null Or cc1.cp06<=" & ServerDate & ") AND cc1.cp10 NOT in (" & CaseMapIn & ") AND CC1.CP10<>'107') OR (cc1.cp10 in (" & CaseMapIn & ") OR CC1.CP10='107')) "
   StrTest1 = StrTest1 & " AND ((substr(cc1.cp12,1,1)='F' and (cc1.cp06 Is Null Or cc1.cp06<=" & ServerDate & ")) or (substr(cc1.cp12,1,1)<>'F' and ((cc1.cp06 Is Null Or cc1.cp06<=" & ServerDate & ") or (cc1.cp10 in (" & CaseMapIn & ") OR CC1.CP10='107'))))"
   '2011/1/31 end
   StrTest2 = StrTest2 & " AND ( cc1.cp06 Is Null Or cc1.cp06<=" & ServerDate & " ) "
   StrTest3 = StrTest3 & " AND ( cc1.cp06 Is Null Or cc1.cp06<=" & ServerDate & " ) "
    '93.11.2 ADD BY SONIA
   StrTest31 = StrTest31 & " AND ( cc1.cp06 Is Null Or cc1.cp06<=" & ServerDate & " ) "
   'add by nick 2004/10/19
   StrTest7 = StrTest7 & " AND ( cc1.cp06 Is Null Or cc1.cp06<=" & ServerDate & " ) "
   'add by nick 2004/12/28
   StrTest98 = StrTest98 & " AND ( cc2.cp06 Is Null Or cc2.cp06<=" & ServerDate & " ) "
   StrTest99 = StrTest99 & " AND ( cc2.cp06 Is Null Or cc2.cp06<=" & ServerDate & " ) "
End If

If bolNoSkip = False Then
   'edit by  nick 2004/12/21
   'strSQLSkip = " And Not( cc1.cp01='T' AND cc1.cp10 in ('715','716','717') and (tm16 is null or tm16='') and (cc1.cp06 is null or cc1.cp06='' )) "
   'edit by nickc 2006/09/08 加入 T & FCT 的 204、205
   'strSQLSkip = " And ((Not (cc1.cp01||cc1.cp10 in ('T715','T716','T717') and (cc1.cp06 is null or cc1.cp06='')) ) or (cc1.cp01 in (select cc2.cp01 from CASEPROGRESS cc2 where cc2.cp01=cc1.cp01 and cc2.cp02=cc1.cp02 and cc2.cp03=cc1.cp03 and cc2.cp04=cc1.cp04 and cc2.cp31='Y' " & StrTest99 & " ))) "
   strSQLSkip = " And ((Not (cc1.cp01||cc1.cp10 in ('T715','T716','T717','T204','T205','FCT204','FCT205') and (cc1.cp06 is null or cc1.cp06='')) ) or (cc1.cp01 in (select cc2.cp01 from CASEPROGRESS cc2 where cc2.cp01=cc1.cp01 and cc2.cp02=cc1.cp02 and cc2.cp03=cc1.cp03 and cc2.cp04=cc1.cp04 and cc2.cp31='Y' " & StrTest99 & " ))) "
   'add by nick 2004/12/21
   '2010/2/3 modify by sonia 加P111(P-087746)
   strSQLSkipP = " And ((Not (cc1.cp01||cc1.cp10 in ('CFP215','CFP416','CFP207','P421','P416','P211','P212','P111') and (cc1.cp06 is  null or cc1.cp06='')) ) or (cc1.cp01 in (select cc2.cp01 from CASEPROGRESS cc2 where cc2.cp01=cc1.cp01 and cc2.cp02=cc1.cp02 and cc2.cp03=cc1.cp03 and cc2.cp04=cc1.cp04 and cc2.cp31='Y'  " & StrTest98 & " ))) "
Else
   strSQLSkip = ""
   'add by nick 2004/12/21
   strSQLSkipP = ""
End If

'add by nickc 2005/04/18 加延遲過系統日或無延遲
StrTest1 = StrTest1 & " and (cc1.cp108 is null or cc1.cp108<=" & strSrvDate(1) & ") "
StrTest2 = StrTest2 & " and (cc1.cp108 is null or cc1.cp108<=" & strSrvDate(1) & ") "
StrTest3 = StrTest3 & " and (cc1.cp108 is null or cc1.cp108<=" & strSrvDate(1) & ") "
StrTest31 = StrTest31 & " and (cc1.cp108 is null or cc1.cp108<=" & strSrvDate(1) & ") "
StrTest4 = StrTest4 & " and (cc1.cp108 is null or cc1.cp108<=" & strSrvDate(1) & ") "
StrTest5 = StrTest5 & " and (cc1.cp108 is null or cc1.cp108<=" & strSrvDate(1) & ") "
StrTest6 = StrTest6 & " and (cc1.cp108 is null or cc1.cp108<=" & strSrvDate(1) & ") "
StrTest7 = StrTest7 & " and (cc1.cp108 is null or cc1.cp108<=" & strSrvDate(1) & ") "
StrTest99 = StrTest99 & " and (cc2.cp108 is null or cc2.cp108<=" & strSrvDate(1) & ") "
StrTest98 = StrTest98 & " and (cc2.cp108 is null or cc2.cp108<=" & strSrvDate(1) & ") "

'Added by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
   strExc(5) = ",(select cp01 v1c1,cp02 v1c2,cp03 v1c3,cp04 v1c4,cp06 v1c6,cp07 v1c7,cp12 v1c8 from casemap,caseprogress where cm10 in ('4','5') and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and ((cm10='4' and cp10='110') or (cm10='5' and cp10 in (" & CaseMapIn & "))) ) VT1 " & _
               ",(select cp01 v2c1,cp02 v2c2,cp03 v2c3,cp04 v2c4,cp06 v2c6,cp07 v2c7,cp12 v2c8 from divisioncase,caseprogress where dc01 in ('P','FCP') and dc01=cp01(+) and dc02=cp02(+) and dc03=cp03(+) and dc04=cp04(+) and cp10 = '307' ) VT2 "
   '判斷條件
   'strExc(6) = "and decode(v2c1,null,decode(v1c1,null,1,decode(v1c6,null,decode(substr(v1c8,1,1),'F',0,1))),decode(v2c6,null,decode(substr(v2c8,1,1),'F',0,1)))=1"
   strExc(6) = "and decode(v1c1||v2c1,null,1,decode(substr(v1c6||v1c8,1,1),'F',0,decode(substr(v2c6||v2c8,1,1),'F',0,1)))=1 "
   '專利案431(PPH)無期限不出現
   strExc(6) = strExc(6) & " and decode(cc1.cp10||cc1.cp06,'431',0,1)=1 "
'end 2015/09/09

If mdiMain.intPCaseKind = 專利 And mdiMain.intPWhere = 國外_CF Then
'edit by nick 2004/10/19
'   strSQL = "SELECT S2.ST01 AS A,cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',cc1.cp06,cc1.cp07," & _
'      "PTM03,cpm03,S1.ST02,NA03,PA26,PA27,PA28,PA29,PA30,PA75,cc1.cp44,S2.ST15 AS D,A0902, cc1.cp79 FROM CASEPROGRESS cc1,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
'      "PATENTTRADEMARKMAP,NATION,ACC090 WHERE (cc1.cp05 BETWEEN " & Day2 & " AND " & Day1 & ") AND cc1.cp05>=19960101 AND cc1.cp14 IS NOT NULL AND cc1.cp27 IS NULL AND " & _
'      "cc1.cp57 IS NULL  AND cc1.cp01=PA01(+) AND cc1.cp02=PA02(+) AND cc1.cp03=PA03(+) AND cc1.cp04=PA04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) " & _
'      "AND cc1.cp01=cpm01(+) AND cc1.cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) AND S2.ST15=A0901(+) " & StrTest1
   '2010/4/1 MODIFY BY SONIA 加分所號
   'Modified by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
'   strSql = "SELECT S2.ST01 AS " & IIf(Trim(txt1(15)) = "1", "A", "F") & ",cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',cc1.cp06,cc1.cp07," & _
      "DECODE(PA09,'000',PTM03,PTM04) ptm03,decode(pa09,'000',cpm03,cpm04) cpm03,cc1.cp14 as " & IIf(Trim(txt1(15)) = "1", "F", "A") & ",NA03,PA26,PA27,PA28,PA29,PA30,PA75,cc1.cp44,S2.ST15 AS " & IIf(Trim(txt1(15)) = "1", "D", "G") & ",AA.A0902 as " & IIf(Trim(txt1(15)) = "1", "E", "H") & ", cc1.cp79,s1.st03 as " & IIf(Trim(txt1(15)) = "1", "G", "D") & ",BB.A0902 as " & IIf(Trim(txt1(15)) = "1", "H", "E") & ",PA47 NO FROM CASEPROGRESS cc1,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
      "PATENTTRADEMARKMAP,NATION,ACC090 AA,Acc090 BB WHERE cc1.cp14 IS NOT NULL AND cc1.cp27 IS NULL AND " & _
      "cc1.cp57 IS NULL  AND cc1.cp01=PA01(+) AND cc1.cp02=PA02(+) AND cc1.cp03=PA03(+) AND cc1.cp04=PA04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) " & _
      "AND cc1.cp01=cpm01(+) AND cc1.cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) " & StrTest1
   'Modified by Lydia 2016/12/21 +排除D類收文 and substr(cc1.cp09,1,1) <> 'D'
   strSql = "SELECT S2.ST01 AS " & IIf(Trim(Txt1(15)) = "1", "A", "F") & ",cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',cc1.cp06,cc1.cp07," & _
      "DECODE(PA09,'000',PTM03,PTM04) ptm03,decode(pa09,'000',cpm03,cpm04) cpm03,cc1.cp14 as " & IIf(Trim(Txt1(15)) = "1", "F", "A") & ",NA03,PA26,PA27,PA28,PA29,PA30,PA75,cc1.cp44,S2.ST15 AS " & IIf(Trim(Txt1(15)) = "1", "D", "G") & ",AA.A0902 as " & IIf(Trim(Txt1(15)) = "1", "E", "H") & ", cc1.cp79,s1.st03 as " & IIf(Trim(Txt1(15)) = "1", "G", "D") & ",BB.A0902 as " & IIf(Trim(Txt1(15)) = "1", "H", "E") & ",PA47 NO " & _
      "FROM CASEPROGRESS cc1,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP,NATION,ACC090 AA,Acc090 BB " & _
      strExc(5) & "WHERE cc1.cp14 IS NOT NULL AND cc1.cp158=0 AND cc1.cp159=0 and substr(cc1.cp09,1,1) <> 'D' AND cc1.cp01=PA01(+) AND cc1.cp02=PA02(+) AND cc1.cp03=PA03(+) AND cc1.cp04=PA04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) " & _
      "AND cc1.cp01=cpm01(+) AND cc1.cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) " & _
      " and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+) " & _
      strExc(6) & StrTest1
Else
'edit by nick 2004/10/19
'   strSQL = "SELECT S2.ST01 AS A,cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',cc1.cp06,cc1.cp07," & _
'      "DECODE(PA09,'000',PTM03,PTM04),cpm03,S1.ST02,NA03,PA26,PA27,PA28,PA29,PA30,PA75,cc1.cp44,S2.ST15 AS D,A0902, cc1.cp79 FROM CASEPROGRESS cc1,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
'      "PATENTTRADEMARKMAP,NATION,ACC090 WHERE (cc1.cp05 BETWEEN " & Day2 & " AND " & Day1 & ") AND cc1.cp05>=19960101 AND cc1.cp14 IS NOT NULL AND cc1.cp27 IS NULL AND " & _
'      "cc1.cp57 IS NULL  AND cc1.cp01=PA01(+) AND cc1.cp02=PA02(+) AND cc1.cp03=PA03(+) AND cc1.cp04=PA04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) " & _
'      "AND cc1.cp01=cpm01(+) AND cc1.cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) AND S2.ST15=A0901(+) " & StrTest1
   '2010/4/1 MODIFY BY SONIA 加分所號
   'Modified by Lydia 2015/09/09 國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
'   strSql = "SELECT S2.ST01 AS " & IIf(Trim(txt1(15)) = "1", "A", "F") & ",cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',cc1.cp06,cc1.cp07," & _
      "DECODE(PA09,'000',PTM03,PTM04) ptm03,decode(pa09,'000',cpm03,cpm04) cpm03,cc1.cp14 as " & IIf(Trim(txt1(15)) = "1", "F", "A") & ",NA03,PA26,PA27,PA28,PA29,PA30,PA75,cc1.cp44,S2.ST15 AS " & IIf(Trim(txt1(15)) = "1", "D", "G") & ",AA.A0902 as " & IIf(Trim(txt1(15)) = "1", "E", "H") & ", cc1.cp79,s1.st03 as " & IIf(Trim(txt1(15)) = "1", "G", "D") & ",BB.A0902 as " & IIf(Trim(txt1(15)) = "1", "H", "E") & ",PA47 NO FROM CASEPROGRESS cc1,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
      "PATENTTRADEMARKMAP,NATION,ACC090 AA,Acc090 BB  WHERE  cc1.cp14 IS NOT NULL AND cc1.cp27 IS NULL AND " & _
      "cc1.cp57 IS NULL  AND cc1.cp01=PA01(+) AND cc1.cp02=PA02(+) AND cc1.cp03=PA03(+) AND cc1.cp04=PA04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) " & _
      "AND cc1.cp01=cpm01(+) AND cc1.cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) " & StrTest1
   'Modified by Lydia 2016/12/21 +排除D類收文 and substr(cc1.cp09,1,1) <> 'D'
   strSql = "SELECT S2.ST01 AS " & IIf(Trim(Txt1(15)) = "1", "A", "F") & ",cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(PA05,NVL(PA06,PA07)),'','',cc1.cp06,cc1.cp07," & _
      "DECODE(PA09,'000',PTM03,PTM04) ptm03,decode(pa09,'000',cpm03,cpm04) cpm03,cc1.cp14 as " & IIf(Trim(Txt1(15)) = "1", "F", "A") & ",NA03,PA26,PA27,PA28,PA29,PA30,PA75,cc1.cp44,S2.ST15 AS " & IIf(Trim(Txt1(15)) = "1", "D", "G") & ",AA.A0902 as " & IIf(Trim(Txt1(15)) = "1", "E", "H") & ", cc1.cp79,s1.st03 as " & IIf(Trim(Txt1(15)) = "1", "G", "D") & ",BB.A0902 as " & IIf(Trim(Txt1(15)) = "1", "H", "E") & ",PA47 NO " & _
      "FROM CASEPROGRESS cc1,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP,NATION,ACC090 AA,Acc090 BB " & _
      strExc(5) & "WHERE  cc1.cp14 IS NOT NULL AND cc1.cp158=0 AND cc1.cp159=0 and substr(cc1.cp09,1,1) <> 'D' AND cc1.cp01=PA01(+) AND cc1.cp02=PA02(+) AND cc1.cp03=PA03(+) AND cc1.cp04=PA04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) " & _
      "AND cc1.cp01=cpm01(+) AND cc1.cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) " & _
      " and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+) " & _
      strExc(6) & StrTest1
End If

'add by nick 2004/12/21
strSql = strSql & strSQLSkipP

'edit by nick 2004/10/19
'strSQL = strSQL + " union all select S2.ST01 AS A,cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(TM05,NVL(TM06,TM07))," & _
'   "'','',cc1.cp06,cc1.cp07,PTM03,cpm03,S1.ST02,NA03,TM23,'','','','',TM44,cc1.cp44,S2.ST15 AS D,A0902, cc1.cp79 FROM CASEPROGRESS cc1,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
'   "PATENTTRADEMARKMAP,NATION,ACC090 WHERE (cc1.cp05 BETWEEN " & Day2 & " AND " & Day1 & ") AND cc1.cp05>=19960101 AND cc1.cp14 IS NOT NULL AND cc1.cp27 IS NULL AND " & _
'   "cc1.cp57 IS NULL AND cc1.cp01=TM01(+) AND cc1.cp02=TM02(+) AND cc1.cp03=TM03(+) AND cc1.cp04=TM04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) AND " & _
'   "cc1.cp01=cpm01(+) AND cc1.cp10=cpm02(+) AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND (TM29<>'Y' OR TM29 IS NULL) AND S2.ST15=A0901(+) " & StrTest2
'2010/4/1 MODIFY BY SONIA 加分所號
'Modify By Sindy 2011/2/24 增加TM78,TM79,TM80,TM81
'Modified by Lydia 2016/12/21 +排除D類收文 and substr(cc1.cp09,1,1) <> 'D'
strSql = strSql + " union all select S2.ST01 AS " & IIf(Trim(Txt1(15)) = "1", "A", "F") & ",cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(TM05,NVL(TM06,TM07))," & _
   "'','',cc1.cp06,cc1.cp07,DECODE(tm10,'000',PTM03,PTM04) ptm03,decode(tm10,'000',cpm03,cpm04) cpm03,cc1.cp14 as " & IIf(Trim(Txt1(15)) = "1", "F", "A") & ",NA03,TM23,TM78,TM79,TM80,TM81,TM44,cc1.cp44,S2.ST15 AS " & IIf(Trim(Txt1(15)) = "1", "D", "G") & ",AA.A0902 as " & IIf(Trim(Txt1(15)) = "1", "E", "H") & ", cc1.cp79,s1.st03 as " & IIf(Trim(Txt1(15)) = "1", "G", "D") & ",BB.A0902 as " & IIf(Trim(Txt1(15)) = "1", "H", "E") & ",TM34 NO FROM CASEPROGRESS cc1,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
   "PATENTTRADEMARKMAP,NATION,ACC090 AA,Acc090 BB WHERE cc1.cp14 IS NOT NULL AND cc1.cp158=0 AND " & _
   "cc1.cp159=0 and substr(cc1.cp09,1,1) <> 'D' AND cc1.cp01=TM01(+) AND cc1.cp02=TM02(+) AND cc1.cp03=TM03(+) AND cc1.cp04=TM04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) AND " & _
   "cc1.cp01=cpm01(+) AND cc1.cp10=cpm02(+) AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND (TM29<>'Y' OR TM29 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) " & StrTest2
   
'Add by Morgan 2003/12/08
strSql = strSql & strSQLSkip

'add by nick 2004/10/19 法務
'2010/4/1 MODIFY BY SONIA 加分所號
'Modify By Sindy 2011/2/24 增加LC43,LC44,LC45,LC46
'Modified by Lydia 2016/12/21 +排除D類收文 and substr(cc1.cp09,1,1) <> 'D'
strSql = strSql + " union all select S2.ST01 AS " & IIf(Trim(Txt1(15)) = "1", "A", "F") & ",cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(lc05,NVL(lc06,lc07))," & _
   "'','',cc1.cp06,cc1.cp07,'',decode(lc15,'000',cpm03,cpm04) cpm03,cc1.cp14 as " & IIf(Trim(Txt1(15)) = "1", "F", "A") & ",NA03,lc11,LC43,LC44,LC45,LC46,lc22,cc1.cp44,S2.ST15 AS " & IIf(Trim(Txt1(15)) = "1", "D", "G") & ",AA.A0902 as " & IIf(Trim(Txt1(15)) = "1", "E", "H") & ", cc1.cp79,s1.st03 as " & IIf(Trim(Txt1(15)) = "1", "G", "D") & ",BB.A0902 as " & IIf(Trim(Txt1(15)) = "1", "H", "E") & ",LC16 NO FROM CASEPROGRESS cc1,Lawcase,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
   "NATION,ACC090 AA,Acc090 BB WHERE cc1.cp14 IS NOT NULL AND cc1.cp158=0 AND " & _
   "cc1.cp159=0 and substr(cc1.cp09,1,1) <> 'D' AND cc1.cp01=lc01(+) AND cc1.cp02=lc02(+) AND cc1.cp03=lc03(+) AND cc1.cp04=lc04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) AND " & _
   "cc1.cp01=cpm01(+) AND cc1.cp10=cpm02(+) AND  lc15=NA01(+) AND (lc34<>'Y' OR lc34 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) " & StrTest3
'add by nick 2004/11/01 顧問
'2006/8/23 modify by sonia 剔除顧問聘任
'strSQL = strSQL + " union all select S2.ST01 AS " & IIf(Trim(txt1(15)) = "1", "A", "F") & ",cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,hc06," & _
'   "'','',cc1.cp06,cc1.cp07,'',cpm03,cc1.cp14 as " & IIf(Trim(txt1(15)) = "1", "F", "A") & ",NA03,hc05,'','','','','',cc1.cp44,S2.ST15 AS " & IIf(Trim(txt1(15)) = "1", "D", "G") & ",AA.A0902 as " & IIf(Trim(txt1(15)) = "1", "E", "H") & ", cc1.cp79,s1.st03 as " & IIf(Trim(txt1(15)) = "1", "G", "D") & ",BB.A0902 as " & IIf(Trim(txt1(15)) = "1", "H", "E") & " FROM CASEPROGRESS cc1,hirecase,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
'   "NATION,ACC090 AA,Acc090 BB WHERE cc1.cp14 IS NOT NULL AND cc1.cp27 IS NULL AND " & _
'   "cc1.cp57 IS NULL AND cc1.cp01=hc01(+) AND cc1.cp02=hc02(+) AND cc1.cp03=hc03(+) AND cc1.cp04=hc04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) AND " & _
'   "cc1.cp01=cpm01(+) AND cc1.cp10=cpm02(+) AND  '000'=NA01(+) AND (hc09<>'Y' OR hc09 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) " & StrTest31
'2010/4/1 MODIFY BY SONIA 加分所號
'Modify By Sindy 2011/2/24 增加HC24,HC25,HC26,HC27
'Modified by Lydia 2016/12/21 +排除D類收文 and substr(cc1.cp09,1,1) <> 'D'
strSql = strSql + " union all select S2.ST01 AS " & IIf(Trim(Txt1(15)) = "1", "A", "F") & ",cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,hc06," & _
   "'','',cc1.cp06,cc1.cp07,'',cpm03,cc1.cp14 as " & IIf(Trim(Txt1(15)) = "1", "F", "A") & ",NA03,hc05,HC24,HC25,HC26,HC27,'',cc1.cp44,S2.ST15 AS " & IIf(Trim(Txt1(15)) = "1", "D", "G") & ",AA.A0902 as " & IIf(Trim(Txt1(15)) = "1", "E", "H") & ", cc1.cp79,s1.st03 as " & IIf(Trim(Txt1(15)) = "1", "G", "D") & ",BB.A0902 as " & IIf(Trim(Txt1(15)) = "1", "H", "E") & ",HC07 NO FROM CASEPROGRESS cc1,hirecase,CASEPROPERTYMAP,STAFF S1,STAFF S2," & _
   "NATION,ACC090 AA,Acc090 BB WHERE cc1.cp14 IS NOT NULL AND cc1.cp158=0 AND cc1.cp10<>'0' and " & _
   "cc1.cp159=0 and substr(cc1.cp09,1,1) <> 'D' AND cc1.cp01=hc01(+) AND cc1.cp02=hc02(+) AND cc1.cp03=hc03(+) AND cc1.cp04=hc04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) AND " & _
   "cc1.cp01=cpm01(+) AND cc1.cp10=cpm02(+) AND  '000'=NA01(+) AND (hc09<>'Y' OR hc09 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) " & StrTest31

'edit by nick 2004/12/08  加快
'strSQL = strSQL + " union all select S2.ST01 AS " & IIf(Trim(txt1(15)) = "1", "A", "F") & ",cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(SP05,NVL(SP06,SP07))," & _
   "'','',cc1.cp06,cc1.cp07,'',cpm03,cc1.cp14 as " & IIf(Trim(txt1(15)) = "1", "F", "A") & ",NA03,SP08,SP58,SP59,'','',SP26,cc1.cp44,S2.ST15 AS " & IIf(Trim(txt1(15)) = "1", "D", "G") & ",AA.A0902 as " & IIf(Trim(txt1(15)) = "1", "E", "H") & ", cc1.cp79,s1.st03 as " & IIf(Trim(txt1(15)) = "1", "G", "D") & ",BB.A0902 as " & IIf(Trim(txt1(15)) = "1", "H", "E") & " FROM CASEPROGRESS cc1,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1," & _
   "STAFF S2,NATION,ACC090 AA,Acc090 BB WHERE cc1.cp14 IS NOT NULL AND cc1.cp27 IS NULL AND cc1.cp57 IS NULL AND " & _
   "cc1.cp01=SP01(+) AND cc1.cp02=SP02(+) AND cc1.cp03=SP03(+) AND cc1.cp04=SP04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) AND cc1.cp01=cpm01(+) AND " & _
   "cc1.cp10=cpm02(+) AND SP09=NA01(+) AND (SP15<>'Y' OR SP15 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) and ((" & StrTest4 & ") or (" & StrTest5 & ") or (" & StrTest6 & ")) " & StrTest7
'2010/4/1 MODIFY BY SONIA 加分所號
'Modify By Sindy 2011/2/24 增加SP65,SP66
'Modified by Lydia 2016/12/21 +排除D類收文 and substr(cc1.cp09,1,1) <> 'D'
strSql = strSql + " union all select S2.ST01 AS " & IIf(Trim(Txt1(15)) = "1", "A", "F") & ",cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(SP05,NVL(SP06,SP07))," & _
   "'','',cc1.cp06,cc1.cp07,'',decode(sp09,'000',cpm03,cpm04) cpm03,cc1.cp14 as " & IIf(Trim(Txt1(15)) = "1", "F", "A") & ",NA03,SP08,SP58,SP59,SP65,SP66,SP26,cc1.cp44,S2.ST15 AS " & IIf(Trim(Txt1(15)) = "1", "D", "G") & ",AA.A0902 as " & IIf(Trim(Txt1(15)) = "1", "E", "H") & ", cc1.cp79,s1.st03 as " & IIf(Trim(Txt1(15)) = "1", "G", "D") & ",BB.A0902 as " & IIf(Trim(Txt1(15)) = "1", "H", "E") & ",SP28 NO FROM CASEPROGRESS cc1,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1," & _
   "STAFF S2,NATION,ACC090 AA,Acc090 BB WHERE cc1.cp14 IS NOT NULL AND cc1.cp158=0 AND cc1.cp159=0 and substr(cc1.cp09,1,1) <> 'D' AND " & _
   "cc1.cp01=SP01(+) AND cc1.cp02=SP02(+) AND cc1.cp03=SP03(+) AND cc1.cp04=SP04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) AND cc1.cp01=cpm01(+) AND " & _
   "cc1.cp10=cpm02(+) AND SP09=NA01(+) AND (SP15<>'Y' OR SP15 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) and (" & StrTest4 & ")  " & StrTest7
strSql = strSql + " union all select S2.ST01 AS " & IIf(Trim(Txt1(15)) = "1", "A", "F") & ",cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(SP05,NVL(SP06,SP07))," & _
   "'','',cc1.cp06,cc1.cp07,'',decode(sp09,'000',cpm03,cpm04) cpm03,cc1.cp14 as " & IIf(Trim(Txt1(15)) = "1", "F", "A") & ",NA03,SP08,SP58,SP59,SP65,SP66,SP26,cc1.cp44,S2.ST15 AS " & IIf(Trim(Txt1(15)) = "1", "D", "G") & ",AA.A0902 as " & IIf(Trim(Txt1(15)) = "1", "E", "H") & ", cc1.cp79,s1.st03 as " & IIf(Trim(Txt1(15)) = "1", "G", "D") & ",BB.A0902 as " & IIf(Trim(Txt1(15)) = "1", "H", "E") & ",SP28 NO FROM CASEPROGRESS cc1,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1," & _
   "STAFF S2,NATION,ACC090 AA,Acc090 BB WHERE cc1.cp14 IS NOT NULL AND cc1.cp158=0 AND cc1.cp159=0 and substr(cc1.cp09,1,1) <> 'D' AND " & _
   "cc1.cp01=SP01(+) AND cc1.cp02=SP02(+) AND cc1.cp03=SP03(+) AND cc1.cp04=SP04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) AND cc1.cp01=cpm01(+) AND " & _
   "cc1.cp10=cpm02(+) AND SP09=NA01(+) AND (SP15<>'Y' OR SP15 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) and (" & StrTest5 & ")  " & StrTest7
strSql = strSql + " union all select S2.ST01 AS " & IIf(Trim(Txt1(15)) = "1", "A", "F") & ",cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(SP05,NVL(SP06,SP07))," & _
   "'','',cc1.cp06,cc1.cp07,'',decode(sp09,'000',cpm03,cpm04) cpm03,cc1.cp14 as " & IIf(Trim(Txt1(15)) = "1", "F", "A") & ",NA03,SP08,SP58,SP59,SP65,SP66,SP26,cc1.cp44,S2.ST15 AS " & IIf(Trim(Txt1(15)) = "1", "D", "G") & ",AA.A0902 as " & IIf(Trim(Txt1(15)) = "1", "E", "H") & ", cc1.cp79,s1.st03 as " & IIf(Trim(Txt1(15)) = "1", "G", "D") & ",BB.A0902 as " & IIf(Trim(Txt1(15)) = "1", "H", "E") & ",SP28 NO FROM CASEPROGRESS cc1,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1," & _
   "STAFF S2,NATION,ACC090 AA,Acc090 BB WHERE cc1.cp14 IS NOT NULL AND cc1.cp158=0 AND cc1.cp159=0 and substr(cc1.cp09,1,1) <> 'D' AND " & _
   "cc1.cp01=SP01(+) AND cc1.cp02=SP02(+) AND cc1.cp03=SP03(+) AND cc1.cp04=SP04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) AND cc1.cp01=cpm01(+) AND " & _
   "cc1.cp10=cpm02(+) AND SP09=NA01(+) AND (SP15<>'Y' OR SP15 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) and  (" & StrTest6 & ") " & StrTest7

strSql = strSql + " ORDER BY D,A,B,C "

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly

If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
    adoRecordset.MoveNext
    Loop
    StrPrintDoc1       '列印主程式
    CheckOC
Else
    ShowNoData
    CheckOC
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Screen.MousePointer = vbDefault
End Sub

'Add By Cheng 2003/02/25
'智權人員明細
Sub StrPrintDoc1()
Dim strTempName As String '代理人名稱
Dim strSalesGrp As String '業務區
Dim strSalesGrpName As String '業務區名稱
'Added by Lydia 2015/07/23
Dim strSG As String, strSGN As String, strSName As String


If Txt1(17) = "1" Then 'Added by Lydia 2015/07/23
    strPath = App.path & "\textfile\"

    strFileN = "收文未發文明細表"
    eFilename = "": eFile = 0: mailFList = "": m_SpecMan = ""
    If Dir(strPath, vbDirectory) = "" Then
        MkDir strPath
    End If
    'Added by Lydia 2016/01/05 清除舊檔
    strExc(7) = Dir(strPath & "*" & strFileN & "*")
    Do While strExc(7) <> ""
       strExc(8) = Mid(strExc(7), 1, 7)
       If CheckIsTaiwanDate(strExc(8), False) Then
          '保留前一個月及當月的檔案
          strExc(8) = ChangeWStringToTString(ChangeWDateStringToWString(strExc(8)))
          If (Left(strExc(8), 5) = Left(strSrvDate(2), 5)) Or (Val(Left(strExc(8), 5)) = Val(Left(strSrvDate(2), 5)) - 1) Or (Val(Left(strExc(8), 5)) = Val(Left(strSrvDate(2), 5)) - 89) Then
             Exit Do
          Else
             Kill strPath & strExc(7)
          End If
       Else
          Exit Do
       End If
       strExc(7) = Dir(strPath & "*" & strFileN & "*")
    Loop
    '檔名前+日期
    strFileN = strSrvDate(2) & "_收文未發文明細表"
    'end 2016/01/05
Else
    GetPrintLeft
    iLine = 1
    Page = 1
End If
With adoRecordset
    .MoveFirst
    'edit by nick 2004/10/19
    'TmpArea = "業務區：" & .Fields("A0902").Value
    TmpArea = IIf(Trim(Txt1(15)) = "1", "業務區：", "部門：") & .Fields("E").Value
    'Modified by Lydia 2015/07/23
    'StrPrintTital TmpArea, str(Page)
    If Txt1(17) = "2" Then StrPrintTital TmpArea, str(Page)
    
    St = .Fields("A")
    'iPrint = 2700
    iPrint = 2400
    iTatle = 0       ' 總筆數
    iKK = 0           ' 合計
    iK = 0           ' 小計
    Do While .EOF = False
        For j = 0 To 11
            If Not IsNull(.Fields(j)) Then
                strTemp3(j) = .Fields(j)
            Else
                strTemp3(j) = ""
            End If
        Next j
    
        'add by nick 2004/10/19
        strTemp3(0) = "" & adoRecordset.Fields("A").Value '智權人
        strTemp3(10) = StrToStr(GetStaffName("" & adoRecordset.Fields("F").Value, True), 4) '承辦人-名稱
        If Not IsNull(adoRecordset.Fields(12)) Then '申請人1
            strTemp3(4) = adoRecordset.Fields(12)
        Else
            If Not IsNull(adoRecordset.Fields(13)) Then
                strTemp3(4) = adoRecordset.Fields(13)
            Else
                If Not IsNull(adoRecordset.Fields(14)) Then
                    strTemp3(4) = adoRecordset.Fields(14)
                Else
                    If Not IsNull(adoRecordset.Fields(15)) Then
                        strTemp3(4) = adoRecordset.Fields(15)
                    Else
                        If Not IsNull(adoRecordset.Fields(16)) Then
                            strTemp3(4) = adoRecordset.Fields(16)
                        End If
                    End If
                End If
            End If
        End If
        strTemp3(4) = GetPrjPeople1(strTemp3(4))
        If Not IsNull(adoRecordset.Fields(17)) Then 'FC代理人
            strTemp3(5) = adoRecordset.Fields(17)
        Else
            If Not IsNull(adoRecordset.Fields(18)) Then 'CP44代理人
                strTemp3(5) = adoRecordset.Fields(18)
            End If
        End If
      'Modify By Cheng 2002/07/05
      '若系統種類對照檔SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'        strTemp3(5) = GetPrjName1(strTemp3(5))
      If PUB_GetAgentName(SystemNumber(strTemp3(2), 1), strTemp3(5), strTempName) = True Then
           strTemp3(5) = strTempName
      Else
           strTemp3(5) = ""
      End If
      strTemp3(12) = "" & adoRecordset.Fields("NO").Value '2010/4/1 ADD BY SONIA 加分所號
        'Add By Cheng 2003/02/26
        '記錄業務區
        strSalesGrp = "" & .Fields("D").Value
        'edit by nick 2004/10/19
        'strSalesGrpName = "" & .Fields("A0902").Value
        strSalesGrpName = "" & .Fields("E").Value
        strSName = GetStaffName(strTemp3(0), True) 'Added by Lydia 2015/07/23
        iK = iK + 1
        'Add By Cheng 2003/02/26
        '記錄業務區筆數
        iKK = iKK + 1
        iTatle = iTatle + 1
        If Len(strTemp3(1)) > 7 Then
            strTemp3(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(1)))
        End If
        If Len(strTemp3(6)) > 7 Then
            strTemp3(6) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(6)))
        End If
        If Len(strTemp3(7)) > 7 Then
            strTemp3(7) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(7)))
        End If
      'Added by Lydia 2015/07/23
        If Txt1(17) = "1" Then
            '新增E-MAIL.TXT
            If St <> "" & .Fields("A") And eFile > 0 Then  '不同業務
               Print #eFile, String(135, "-")
               Print #eFile, String(10, " ") & "小計： " & iK - 1 & " 筆"
               Print #eFile, String(135, "-")
               iK = 1
            End If
            If eFilename <> "" And eFile > 0 And ((Txt1(16) = "1" And St <> "" & .Fields("A")) Or (InStr("2,3", Txt1(16)) > 0 And eFilename <> IIf(InStr("2,3", Txt1(16)) > 0, strFileN & "_" & .Fields("E").Value, strFileN & "_" & .Fields("E").Value & "_" & strSName))) Then
               If Txt1(16) <> "1" Then
                  Print #eFile, strSGN & "合計： " & iKK - 1 & " 筆"
                  Print #eFile, String(135, "-")
               End If
               iKK = 1
               Close eFile
               Call MailtoReceiver("1", strSG, St, eFilename, strSalesGrp)
               eFilename = ""
            End If
            If eFilename = "" Then
               eFilename = IIf(InStr("2,3", Txt1(16)) > 0, strFileN & "_" & .Fields("E").Value, strFileN & "_" & .Fields("E").Value & "_" & strSName)
               eFile = FreeFile
               If eFile > 0 Then Close #eFile
               eFile = FreeFile
               Open strPath & eFilename & ".txt" For Output As eFile
               Select Case Txt1(16)
                    Case "1"
                       strExc(3) = "管制時段：個人"
                    Case "2"
                       strExc(3) = "管制時段：部門主管"
                    Case "3"
                       strExc(3) = "管制時段：副總"
                    Case "4"
                       strExc(3) = ""
               End Select
  
               Print #eFile, "列印人：" & convForm(strUserName, 16) & convForm(strExc(3), 40) & "列印順序：" & convForm(IIf(Trim(Txt1(15)) = "1", "智權人員", "承辦人"), 22) & "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
               Print #eFile, TmpArea '業務區＆部門
               'modify by sonia 2015/10/1
               'Print #eFile, "專利收文日期：" & ChangeTStringToTDateString(txt1(5)) & "-" & ChangeTStringToTDateString(txt1(6)) & "  商標收文日期：" & ChangeTStringToTDateString(txt1(11)) & "-" & ChangeTStringToTDateString(txt1(12)) & "  法務收文日期：" & ChangeTStringToTDateString(txt1(13)) & "-" & ChangeTStringToTDateString(txt1(14))
               Print #eFile, "專利收文日期：" & ChangeTStringToTDateString(Txt1(5)) & "-" & ChangeTStringToTDateString(Txt1(6)) & IIf(Trim(Txt1(11)) = "111112", "", "  商標收文日期：" & ChangeTStringToTDateString(Txt1(11)) & "-" & ChangeTStringToTDateString(Txt1(12))) & IIf(Trim(Txt1(11)) = "111112", "", "  法務收文日期：" & ChangeTStringToTDateString(Txt1(13)) & "-" & ChangeTStringToTDateString(Txt1(14)))
               'end 2015/10/1
               Print #eFile, String(135, "-")
               If Txt1(15) = "1" Then '列印順序：智權人員
                 Print #eFile, "智權人員 收文日　  本所案號　　    案件名稱 申請人       本所期限  法定期限  種類     案件性質 承辦人   申請國家 未收金額 分所號　　  "
                 Print #eFile, "======== ========= =============== ======== ============ ========= ========= ======== ======== ======== ======== ======== ============"
               Else
                 Print #eFile, "承辦人   收文日　  本所案號　　    案件名稱 申請人       本所期限  法定期限  種類     案件性質 智權人員 申請國家 未收金額 分所號　　  "
                 Print #eFile, "======== ========= =============== ======== ============ ========= ========= ======== ======== ======== ======== ======== ============"
               End If
            End If
            
            '0~3
            strExc(1) = convForm(strSName, 8) & " " & convForm(strTemp3(1), 9) & " " & convForm(strTemp3(2), 15) & " " & convForm(strTemp3(3), 8)
            '4~8 ,strTemp3(5)=代理人
            strExc(1) = strExc(1) & " " & convForm(strTemp3(4), 12) & " " & convForm(strTemp3(6), 9) & " " & convForm(strTemp3(7), 9) & " " & convForm(strTemp3(8), 8)
            '9~12
            strExc(1) = strExc(1) & " " & convForm(strTemp3(9), 8) & " " & convForm(strTemp3(10), 8) & " " & convForm(strTemp3(11), 8) & " " & PUB_StrToStr(Format(.Fields(21).Value, "#,###"), 8, True, True) & " " & convForm(strTemp3(12), 12)
            Print #eFile, strExc(1)

            St = "" & .Fields("A"): strSG = strSalesGrp: strSGN = strSalesGrpName
        Else
            'Printer.Font.Name = "Arial"
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            If iK = 1 Then
                'Modify By Cheng 2003/02/24
                '取得承辦人姓名
    '           Printer.Print StrToStr(strTemp3(0), 4)
               Printer.Print StrToStr(GetStaffName(strTemp3(0), True), 4)
            Else
               Printer.Print ""
            End If
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print strTemp3(1) '收文日
            Printer.CurrentX = PLeft(2)
            Printer.CurrentY = iPrint
            Printer.Print strTemp3(2)  '本所案號
            Printer.CurrentX = PLeft(3)
            Printer.CurrentY = iPrint
            Printer.Print StrToStr(strTemp3(3), 4) '案件名稱
            Printer.CurrentX = PLeft(4)
            Printer.CurrentY = iPrint
            Printer.Print StrToStr(strTemp3(4), 6) '申請人
            Printer.CurrentX = PLeft(5)
            Printer.CurrentY = iPrint
            Printer.Print strTemp3(6) '所限
            Printer.CurrentX = PLeft(6)
            Printer.CurrentY = iPrint
            Printer.Print strTemp3(7) '法限
            Printer.CurrentX = PLeft(7)
            Printer.CurrentY = iPrint
            Printer.Print StrToStr(strTemp3(8), 4) '種類
            Printer.CurrentX = PLeft(8)
            Printer.CurrentY = iPrint
            Printer.Print StrToStr(strTemp3(9), 4) '案件性質
            Printer.CurrentX = PLeft(9)
            Printer.CurrentY = iPrint
            Printer.Print StrToStr(strTemp3(10), 4)
            Printer.CurrentX = PLeft(10)
            Printer.CurrentY = iPrint
            Printer.Print StrToStr(strTemp3(11), 4) '申請國家
            Printer.CurrentX = PLeft(11) + 1000 - Printer.TextWidth(Format(.Fields(21).Value, "#,###"))
            Printer.CurrentY = iPrint
            Printer.Print Format(.Fields(21).Value, "#,###") '未收金額
            '2010/4/1 ADD BY SONIA 加分所號
            Printer.CurrentX = PLeft(12)
            Printer.CurrentY = iPrint
            Printer.Print StrToStr(strTemp3(12), 4) '分所號
            '2010/4/1 END
        End If 'end 2015/07/23
      
        .MoveNext
        
'*******換頁判斷******************************
        If .EOF = False Then
'edit by nick 2004/10/19
'            If Not IsNull(.Fields(0)) Then
'                StrTest1 = .Fields(0)
            If Not IsNull(.Fields("A")) Then
                StrTest1 = .Fields("A")
            Else
                StrTest1 = ""
            End If
        End If

        If Txt1(17) = "2" Then 'Added by Lydia 2015/07/23
            If .EOF = False Then
                '若智權人員不同時
                If StrTest1 <> St Then
                   iPrint = iPrint + 300
                   Printer.CurrentX = 0
                   Printer.CurrentY = iPrint
                   Printer.Print String(200, "-")
                   iPrint = iPrint + 300
                   Printer.CurrentX = 1000
                   Printer.CurrentY = iPrint
                   Printer.Print "小計： " & Trim(str(iK)) & " 筆"
                   iK = 0
                   iPrint = iPrint + 300
                   Printer.CurrentX = 0
                   Printer.CurrentY = iPrint
                   Printer.Print String(200, "-")
                   '2012/9/26 MODIFY BY SONIA 列印對象為智權人員且管制時段為智權人員時依智權人員跳頁
                   'iLine = iLine + 3
                   If Txt1(15) = "1" And Txt1(16) = "1" Then
                      Printer.NewPage
                      Page = Page + 1
                      TmpArea = IIf(Trim(Txt1(15)) = "1", "業務區：", "部門：") & .Fields("E").Value
                      StrPrintTital TmpArea, str(Page)
                      iPrint = 2100
                      iLine = 0
                      iLine = iLine + 1
                      iPrint = iPrint + 300
                   Else
                      iLine = iLine + 3
                   End If
                    '2012/9/26 END
                    St = StrTest1
                End If
                If (iLine Mod 26 = 0) Or iPrint >= 10000 Then
                    iPrint = iPrint + 300
                    Printer.NewPage
                    Page = Page + 1
                    StrPrintTital TmpArea, str(Page)
                    'iPrint = 2400
                    iPrint = 2100
                    iLine = 0
                End If
                iLine = iLine + 1
                iPrint = iPrint + 300
                '若業務區不同時
                If strSalesGrp <> "" & .Fields("D").Value Then
                   '2012/9/26 MODIFY BY SONIA 列印對象為智權人員且管制時段為智權人員時依智權人員跳頁且不必印部門合計但仍要跳頁
                   If Txt1(15) = "1" And Txt1(16) = "1" Then
                   Else
                   '2012/9/26 END
                      Printer.CurrentX = PLeft(0)
                      Printer.CurrentY = iPrint
                      Printer.Print strSalesGrpName & "合計： " & Trim(str(iKK)) & " 筆"
                      iKK = 0
                      iPrint = iPrint + 300
                      Printer.CurrentX = 0
                      Printer.CurrentY = iPrint
                      Printer.Print String(200, "-")
                      iPrint = iPrint + 300
                      Printer.NewPage
                      Page = Page + 1
                      'edit by nick 2004/10/19
                      'TmpArea = "業務區：" & .Fields("A0902").Value
                      TmpArea = IIf(Trim(Txt1(15)) = "1", "業務區：", "部門：") & .Fields("E").Value
                      StrPrintTital TmpArea, str(Page)
                      'iPrint = 2400
                      iPrint = 2100
                      iLine = 0
                      iLine = iLine + 1
                      iPrint = iPrint + 300
                   End If   '2012/9/26 ADD BY SONIA
                End If
            End If
        End If  'Added by Lydia 2015/07/23
    Loop
End With

'報表表尾
      'Added by Lydia 2015/07/23
    If Txt1(17) = "1" Then
        '新增E-MAIL.TXT
        If St <> "" And eFile > 0 Then
           Print #eFile, String(135, "-")
           Print #eFile, String(10, " ") & "小計： " & iK & " 筆"
           Print #eFile, String(135, "-")
        End If
        If InStr("2,3", Txt1(16)) > 0 And eFile > 0 Then
           Print #eFile, strSGN & "合計： " & iKK & " 筆"
           Print #eFile, String(135, "-")
        End If
           Close eFile
           Call MailtoReceiver("2", strSG, St, eFilename, strSalesGrp)
           MsgBox "E-MAIL發送完成!!", vbOKOnly, "E-MAIL發送"
    Else
        'Add By Cheng 2003/02/26
        '智權人員小計
        iPrint = iPrint + 300
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        Printer.CurrentX = 1000
        Printer.CurrentY = iPrint
        Printer.Print "小計： " & Trim(str(iK)) & " 筆"
        iK = 0
        iPrint = iPrint + 300
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iLine = iLine + 3
        St = StrTest1
        If (iLine Mod 26 = 0) Or iPrint >= 10000 Then
            iPrint = iPrint + 300
            Printer.NewPage
            Page = Page + 1
            StrPrintTital TmpArea, str(Page)
            'iPrint = 2400
            iPrint = 2100
            iLine = 0
        End If
        iLine = iLine + 1
        iPrint = iPrint + 300
        '2012/9/26 ADD BY SONIA 列印對象為智權人員且管制時段為智權人員時依智權人員跳頁最後不必印部門合計
        If Txt1(15) = "1" And Txt1(16) = "1" Then
        Else
        '2012/9/26 END
           '業務區合計
           Printer.CurrentX = PLeft(0)
           Printer.CurrentY = iPrint
           Printer.Print strSalesGrpName & "合計： " & Trim(str(iKK)) & " 筆"
           iKK = 0
           iPrint = iPrint + 300
           Printer.CurrentX = 0
           Printer.CurrentY = iPrint
           Printer.Print String(200, "-")
           iPrint = iPrint + 300
        End If   '2012/9/26 ADD BY SONIA
        
        Printer.EndDoc
        ShowPrintOk
    End If
CheckOC

End Sub
'Remove by Lydia 2015/07/23
'Add By Cheng 2003/01/09
'Sub StrPrintDocTotal1()
'Dim strTempName As String '代理人名稱
'
'GetPrintLeft
'iLine = 1
'Page = 1
'StrPrintTital TmpArea, str(Page)
'iPrint = 2700
'iTatle = 0       ' 總數
'With adoRecordset
'    .MoveFirst
'    Do While .EOF = False
'        For j = 0 To 2
'            If Not IsNull(.Fields(j)) Then
'                strTemp3(j) = .Fields(j)
'            Else
'                strTemp3(j) = ""
'            End If
'        Next j
'        Printer.CurrentX = PLeft(0)
'        Printer.CurrentY = iPrint
'        Printer.Print strTemp3(1)
'        Printer.CurrentX = PLeft(2) + 750 - TextWidth(Format(strTemp3(2), "##0") & " 筆")
'        Printer.CurrentY = iPrint
'        Printer.Print Format(strTemp3(2), "##0") & " 筆"
'        iTatle = iTatle + Val(strTemp3(2))
'        .MoveNext
'        If (iLine Mod 26 = 0) Or iPrint >= 10000 Then
'            iPrint = iPrint + 300
'            Printer.NewPage
'            Page = Page + 1
'            StrPrintTital TmpArea, str(Page)
'            'iPrint = 2400
'            iPrint = 2100
'            iLine = 0
'        End If
'        iLine = iLine + 1
'        iPrint = iPrint + 300
'    Loop
'End With
''合計
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Print String(200, "-")
'iPrint = iPrint + 300
'Printer.CurrentX = PLeft(2) + 750 - TextWidth("合計：共 " & Format(iTatle, "##0") & " 筆")
'Printer.CurrentY = iPrint
'Printer.Print "合計：共 " & Format(iTatle, "##0") & " 筆"
'Printer.EndDoc
'ShowPrintOk
'CheckOC
'End Sub
'add by nick 2004/12/08
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      'Modified by Lydia 2015/07/23 +txt1(18)
      Case 0, 1, 2, 3, 4, 18
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Cancel = False
Dim bolUser As Boolean 'Added by Lydia 2015/07/23
Select Case Index
Case 0
     strTemp1 = Split(UCase(GetSystemKindByNick), ",")
     strTemp2 = Split(UCase(Txt1(0)), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            Txt1(0).SetFocus
            Txt1(0).SelStart = 0
            Txt1(0).SelLength = Len(Txt1(0))
            Cancel = True
            Exit Sub
        End If
    Next i
Case 3
     lbl1(0) = GetPrjSalesNM(Txt1(Index))
     If Len(Txt1(Index)) <> 0 Then
        If Len(lbl1(0).Caption) = 0 Then
            s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
            Txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Cancel = True
        End If
     End If
Case 4
     lbl1(1) = GetPrjSalesNM(Txt1(Index))
     If Len(Txt1(Index)) <> 0 Then
        If Len(lbl1(1).Caption) = 0 Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            Txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Cancel = True
        End If
     End If
Case 5, 11, 13
        Cancel = Not ChkDate(Txt1(Index).Text)
Case 2, 6, 8, 10, 12, 14
      If RunNick(Txt1(Index - 1), Txt1(Index)) Then
         Txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Cancel = True
         Exit Sub
      End If
      If Index = 6 Or Index = 12 Or Index = 14 Then
            Cancel = Not ChkDate(Txt1(Index).Text)
      End If
      
Case 15
    Select Case Txt1(Index)
    Case "1", "2"
    Case Else
         MsgBox "列印順序輸入錯誤！", , "輸入錯誤！"
         Cancel = True
    End Select
'2005/3/23 add by sonia
Case 16
   '2013/5/14 modify by sonia 月底跑預設少跑一個月,例102/5/25日跑專利則跑850101~1011130,而102/5/1跑至850101~1011031
   Select Case Txt1(Index)
      Case "1"
         If Val(Right(strSrvDate(1), 2)) > 20 Then
            Txt1(6) = TransDate(CompDate(2, -1, CompDate(1, -5, Left(strSrvDate(1), 6) & "01")), 1)
            'txt1(12) = TransDate(CompDate(2, -1, CompDate(1, -2, Left(strSrvDate(1), 6) & "01")), 1) 'cancel by sonia 2015/10/1 商標改晚上批次跑
            'txt1(14) = ChangeWDateStringToTString(DateAdd("m", -3, ChangeWStringToWDateString(Left(strSrvDate(1), 6) * 100 + 31)))   '2013/5/14 cancel by sonia
         Else
            Txt1(6) = TransDate(CompDate(2, -1, CompDate(1, -6, Left(strSrvDate(1), 6) & "01")), 1)
            'txt1(12) = TransDate(CompDate(2, -1, CompDate(1, -3, Left(strSrvDate(1), 6) & "01")), 1) 'cancel by sonia 2015/10/1 商標改晚上批次跑
            'txt1(14) = ChangeWDateStringToTString(DateAdd("m", -4, ChangeWStringToWDateString(Left(strSrvDate(1), 6) * 100 + 31)))    '2013/5/14 cancel by sonia
         End If
         Txt1(18).Text = "" 'Added by Lydia 2015/07/23
      Case "2"
         If Val(Right(strSrvDate(1), 2)) > 20 Then
            Txt1(6) = TransDate(CompDate(2, -1, CompDate(1, -6, Left(strSrvDate(1), 6) & "01")), 1)
            'txt1(12) = TransDate(CompDate(2, -1, CompDate(1, -3, Left(strSrvDate(1), 6) & "01")), 1) 'cancel by sonia 2015/10/1 商標改晚上批次跑
            'txt1(14) = ChangeWDateStringToTString(DateAdd("m", -4, ChangeWStringToWDateString(Left(strSrvDate(1), 6) * 100 + 31)))    '2013/5/14 cancel by sonia
         Else
            Txt1(6) = TransDate(CompDate(2, -1, CompDate(1, -7, Left(strSrvDate(1), 6) & "01")), 1)
            'txt1(12) = TransDate(CompDate(2, -1, CompDate(1, -4, Left(strSrvDate(1), 6) & "01")), 1) 'cancel by sonia 2015/10/1 商標改晚上批次跑
            'txt1(14) = ChangeWDateStringToTString(DateAdd("m", -5, ChangeWStringToWDateString(Left(strSrvDate(1), 6) * 100 + 31)))    '2013/5/14 cancel by sonia
         End If
         Txt1(18).Text = "" 'Added by Lydia 2015/07/23
      Case "3"
         If Val(Right(strSrvDate(1), 2)) > 20 Then
            Txt1(6) = TransDate(CompDate(2, -1, CompDate(1, -7, Left(strSrvDate(1), 6) & "01")), 1)
            'txt1(12) = TransDate(CompDate(2, -1, CompDate(1, -4, Left(strSrvDate(1), 6) & "01")), 1) 'cancel by sonia 2015/10/1 商標改晚上批次跑
            'txt1(14) = ChangeWDateStringToTString(DateAdd("m", -5, ChangeWStringToWDateString(Left(strSrvDate(1), 6) * 100 + 31)))    '2013/5/14 cancel by sonia
         Else
            Txt1(6) = TransDate(CompDate(2, -1, CompDate(1, -8, Left(strSrvDate(1), 6) & "01")), 1)
            'txt1(12) = TransDate(CompDate(2, -1, CompDate(1, -5, Left(strSrvDate(1), 6) & "01")), 1) 'cancel by sonia 2015/10/1 商標改晚上批次跑
            'txt1(14) = ChangeWDateStringToTString(DateAdd("m", -6, ChangeWStringToWDateString(Left(strSrvDate(1), 6) * 100 + 31)))    '2013/5/14 cancel by sonia
         End If
         'Added by Lydia 2015/07/23
         If Txt1(17).Text = "1" And Txt1(18).Text = "" Then
                'Modified by Morgan 2019/9/19
                'txt1(18).Text = "94007;68009"  '預設總經理和何主秘
                'modify by sonia 2020/1/9 +69005
                'modify by sonia 2020/3/4 -68006退休
                Txt1(18).Text = "94007;69005"  '預設總經理和杜主秘
         ElseIf Txt1(17).Text <> "1" Then
                Txt1(18).Text = ""
         End If
         'end 2015/07/23
         
      Case "4"
         Txt1(18).Text = "" 'Added by Lydia 2015/07/23
      Case Else
         MsgBox "收文管制時段輸入錯誤！", , "輸入錯誤！"
         Cancel = True
   End Select
'2005/3/23 END
'Added by Lydia 2015/07/23 +txt1(17)輸出方式,txt1(18)E-MAIL副總收件人
Case 17
    Select Case Txt1(Index)
    Case "1"
         'Modified by Morgan 2019/9/19
         'If txt1(16).Text = "3" And txt1(18).Text = "" Then txt1(18).Text = "94007;68009"
         'modify by sonia 2020/1/9 +69005
         'modify by sonia 2020/3/4 -68006退休
         If Txt1(16).Text = "3" And Txt1(18).Text = "" Then Txt1(18).Text = "94007;69005"
    Case "2"
         If Txt1(18).Text <> "" Then Txt1(18).Text = ""
    Case Else
         MsgBox "輸出方式輸入錯誤！", , "輸入錯誤！"
         Cancel = True
    End Select
Case 18
     If Len(Txt1(18)) > 0 Then
        strTemp1 = Split(UCase(Txt1(18)), ";")
        For i = 0 To UBound(strTemp1)
           If Len(strTemp1(i)) >= 5 Then
              strExc(1) = ""
              bolUser = ClsPDGetStaff(strTemp1(i), strExc(1))
              If bolUser = False Then
                 Cancel = True
              End If
           Else
              MsgBox "請輸入正確的員工編號!!", vbCritical, "E-MAIL副總收件人"
              Cancel = True
'              txt1(18).SetFocus
'              txt1_GotFocus (18)
'              Cancel = True
'              Exit Sub
           End If
           If Cancel = True Then
               Txt1(18).SetFocus
               txt1_GotFocus (18)
               Cancel = True
               Exit Sub
           End If
        Next i
     End If
End Select
'end 2015/07/23

End Sub

Private Sub txtOver_GotFocus()
   TextInverse txtOver
End Sub

Private Sub txtOver_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 89 And KeyAscii <> 32 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Sub
'Added by Lydia 2015/07/23 判斷收件人和發mail
'sKind=1 讀檔中;    sKind=2 最後一筆
'stArea=部門代號
'mailFList 多個附加檔寄出,區隔用*
Private Sub MailtoReceiver(sKind As String, stArea As String, stTO As String, fName As String, nArea As String)
Dim stCC As String, Str01 As String, Str02 As String
    
    stCC = ""
    Select Case Txt1(16)
        Case "1"
              Str01 = PUB_GetST03(stTO)
        Case "2"
              Str01 = GetDeptA09(stArea, "08")
              If Len(Str01) > 0 Then stTO = Str01
              If Left(stArea, 1) = "F" Then
                 stCC = Pub_GetSpecMan("O")
              End If
        Case "3"
             stTO = Trim(Txt1(18))
        Case Else
             stTO = ""
    End Select
    
    If Txt1(16) = "3" Or (Txt1(16) = "2" And stCC <> "") Then
        mailFList = mailFList & "*" & strPath & fName & ".txt" '多個附件的區隔用*
    End If
    
    If Txt1(16) = "1" Or Txt1(16) = "2" Then
        If fName <> "" And stTO <> "" And stCC = "" Then
           'modify by sonia 2016/4/1 收受者請假不彈訊息
           PUB_SendMail strUserNum, stTO, "", eFilename, vbCrLf & vbCrLf & "詳細內容請見附件", "", strPath & fName & ".txt", , , , , , , , , False
        End If
        'F開頭部門寄給系統特殊人員("O")大寫,但依部門之前二碼為部門檔案控制
        If stCC <> "" And (sKind = "2" Or (Left(stArea, 2) <> Left(nArea, 2))) Then
           'modify by sonia 2016/4/1 收受者請假不彈訊息
           PUB_SendMail strUserNum, stCC, "", eFilename, vbCrLf & vbCrLf & "詳細內容請見附件", "", IIf(Left(mailFList, 1) = "*", Mid(mailFList, 2, Len(mailFList) - 1), mailFList), , , , , , , , , False
           mailFList = ""
        End If
    ElseIf Txt1(16) = "3" And sKind = "2" Then '只發一封mail
           'modify by sonia 2016/4/1 收受者請假不彈訊息
           PUB_SendMail strUserNum, stTO, "", strFileN, vbCrLf & vbCrLf & "詳細內容請見附件", "", IIf(Left(mailFList, 1) = "*", Mid(mailFList, 2, Len(mailFList) - 1), mailFList), , , , , , , , , False
    End If
End Sub
