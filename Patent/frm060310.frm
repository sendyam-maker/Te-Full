VERSION 5.00
Begin VB.Form frm060310 
   BorderStyle     =   1  '單線固定
   Caption         =   "催審函/催審表"
   ClientHeight    =   3804
   ClientLeft      =   3132
   ClientTop       =   2136
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3804
   ScaleWidth      =   4800
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   16
      Left            =   2130
      MaxLength       =   4
      TabIndex        =   16
      Top             =   3450
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   15
      Left            =   1224
      MaxLength       =   4
      TabIndex        =   15
      Top             =   3450
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1224
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2550
      Width           =   435
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1224
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1680
      Width           =   1035
   End
   Begin VB.OptionButton opt 
      Caption         =   "催審期限："
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   1710
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   14
      Left            =   2340
      MaxLength       =   7
      TabIndex        =   6
      Top             =   2250
      Width           =   1035
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   13
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2250
      Width           =   1035
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3885
      TabIndex        =   18
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2970
      TabIndex        =   17
      Top             =   45
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   1224
      MaxLength       =   1
      TabIndex        =   1
      Top             =   780
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   2544
      MaxLength       =   9
      TabIndex        =   14
      Top             =   3150
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   1224
      MaxLength       =   9
      TabIndex        =   13
      Top             =   3150
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   2544
      MaxLength       =   9
      TabIndex        =   12
      Top             =   2850
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1224
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2850
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   2916
      MaxLength       =   2
      TabIndex        =   10
      Top             =   2550
      Width           =   360
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2550
      Width           =   210
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1728
      MaxLength       =   6
      TabIndex        =   8
      Top             =   2550
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2340
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1680
      Width           =   1035
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1230
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1095
      Width           =   285
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1230
      TabIndex        =   0
      Top             =   480
      Width           =   2145
   End
   Begin VB.OptionButton opt 
      Caption         =   "發文日期："
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   2280
      Width           =   1275
   End
   Begin VB.Line Line6 
      X1              =   1995
      X2              =   2235
      Y1              =   3570
      Y2              =   3570
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   35
      Top             =   3495
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   9
      Left            =   135
      TabIndex        =   34
      Top             =   2580
      Width           =   945
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2820
      X2              =   2940
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   3
      Left            =   3300
      TabIndex        =   33
      Top             =   2010
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   2
      Left            =   1980
      TabIndex        =   32
      Top             =   2010
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "上次列印發文日期"
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   31
      Top             =   2010
      Width           =   1440
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2820
      X2              =   2940
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   1
      Left            =   3300
      TabIndex        =   30
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   0
      Left            =   1980
      TabIndex        =   29
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "上次列印催審期限"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   28
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Line Line5 
      X1              =   1680
      X2              =   2820
      Y1              =   2385
      Y2              =   2385
   End
   Begin VB.Line Line4 
      X1              =   2235
      X2              =   3150
      Y1              =   3285
      Y2              =   3285
   End
   Begin VB.Line Line3 
      X1              =   2190
      X2              =   2850
      Y1              =   2955
      Y2              =   2955
   End
   Begin VB.Line Line2 
      X1              =   1350
      X2              =   3225
      Y1              =   2685
      Y2              =   2685
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1680
      X2              =   2820
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "(1.申請書 2.管制表)"
      Height          =   180
      Index           =   8
      Left            =   1620
      TabIndex        =   25
      Top             =   825
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(1.催實審 2.催審 3.申請案號 4.證書號數)"
      Height          =   180
      Index           =   7
      Left            =   1560
      TabIndex        =   24
      Top             =   1170
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   23
      Top             =   825
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   22
      Top             =   3195
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   2865
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "催審函性質："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   1155
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   540
      Width           =   900
   End
End
Attribute VB_Name = "frm060310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, strTemp3(0 To 9) As String, SavDay(0 To 1) As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 18) As String, StrTemp99(0 To 15) As String, PrintPage As Boolean
Dim PLeft(0 To 8) As Integer, strTemp1 As Variant, strTemp2 As Variant, Bol1 As Boolean, STRSTRING As String, SeekPrint As Integer, SeekPrintL As Integer
'Add By Cheng 2002/09/16
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
'add by nickc 2007/02/08
Dim StrSQL6 As String
Dim SavDay1 As String
Dim SavDay2 As String

'Add by Morgan 2005/1/12
Private Function TxtValidate() As Boolean
   '系統類別
   If Len(txt1(0)) = 0 Then
     s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
     txt1(0).SetFocus
     Exit Function
   End If
   '列印別
   If Len(txt1(12)) = 0 Then
      s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
      Me.txt1(12).SetFocus
      Exit Function
   End If
   
   '催審函性質
   If Len(txt1(1)) = 0 Then
      s = MsgBox("催審函性質不可空白!!", , "USER 輸入錯誤")
      txt1(1).SetFocus
      Exit Function
   End If
   
   '申請案號, 證書號數
   If txt1(12) = "2" And (txt1(1) = "3" Or txt1(1) = "4") Then
      s = MsgBox("催審函性質為'申請案號或證書號數'時,列印別不可為'管制表'", , "USER 輸入錯誤")
      txt1(1).SetFocus
      Exit Function
   End If
   
   '選擇催審期限
   If Me.opt(0).Value Then
      If txt1(1) = "1" Then
         s = MsgBox("催審函性質為'催實審'時,只可選發文日！", , "USER 輸入錯誤")
         txt1(1).SetFocus
         Exit Function
      End If
      If ChkDate(Me.txt1(2).Text) = False Then
         Me.txt1(2).SetFocus
         Exit Function
      End If
      If ChkDate(Me.txt1(3).Text) = False Then
         Me.txt1(3).SetFocus
         Exit Function
      End If
      If Val(Me.txt1(2).Text) > Val(Me.txt1(3).Text) Then
         MsgBox "催審期限區間輸入錯誤!!!", vbExclamation + vbOKOnly
         blnClkSure = True
         Me.txt1(2).SetFocus
         txt1_GotFocus 2
         Exit Function
      End If
   End If
   
   '選擇發文日期
   If Me.opt(1).Value Then
      If ChkDate(Me.txt1(13).Text) = False Then
         Me.txt1(13).SetFocus
         Exit Function
      End If
      If ChkDate(Me.txt1(14).Text) = False Then
         Me.txt1(14).SetFocus
         Exit Function
      End If
      If Val(Me.txt1(13).Text) > Val(Me.txt1(14).Text) Then
         MsgBox "發文日期區間輸入錯誤!!!", vbExclamation + vbOKOnly
         blnClkSure = True
         Me.txt1(13).SetFocus
         txt1_GotFocus 13
         Exit Function
      End If
   End If
   
   '申請人
   If Len(txt1(8)) <> 0 Or Len(txt1(9)) <> 0 Then
      If Mid(txt1(8), 1, 6) <> Mid(txt1(9), 1, 6) Then
         s = MsgBox("申請人前六碼必須相同!!", , "USER 輸入錯誤")
         blnClkSure = True
         txt1(8).SetFocus
         txt1_GotFocus (8)
         Exit Function
      End If
      If Me.txt1(8).Text > Me.txt1(9).Text Then
         MsgBox "申請人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         blnClkSure = True
         Me.txt1(8).SetFocus
         txt1_GotFocus 8
         Exit Function
      End If
   End If
   '代理人
   If Len(txt1(10)) <> 0 Or Len(txt1(11)) <> 0 Then
      If Mid(txt1(10), 1, 6) <> Mid(txt1(11), 1, 6) Then
         s = MsgBox("代理人前六碼必須相同!!", , "USER 輸入錯誤")
         blnClkSure = True
         txt1(10).SetFocus
         txt1_GotFocus (10)
         Exit Function
      End If
      If Me.txt1(10).Text > Me.txt1(11).Text Then
         MsgBox "代理人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         blnClkSure = True
         Me.txt1(10).SetFocus
         txt1_GotFocus 10
         Exit Function
      End If
   End If

   TxtValidate = True
End Function

'Add by Morgan 2005/1/12
Private Sub cmdok_Click(Index As Integer)

   Select Case Index
      Case 0
      
         If TxtValidate = False Then Exit Sub
         
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         
         Printer.Orientation = 2
         DoEvents
         
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/10 清除查詢印表記錄檔欄位
         If Val(txt1(12)) = 1 Then
            pub_QL05 = pub_QL05 & ";" & Label1(6) & "1.申請書" 'Add By Sindy 2010/12/10
         Else
            pub_QL05 = pub_QL05 & ";" & Label1(6) & "2.管制表" 'Add By Sindy 2010/12/10
         End If
         If txt1(1) = "1" Then
            pub_QL05 = pub_QL05 & ";" & Label1(1) & "1.催實審" 'Add By Sindy 2010/12/10
         ElseIf txt1(1) = "2" Then
            pub_QL05 = pub_QL05 & ";" & Label1(1) & "2.催審" 'Add By Sindy 2010/12/10
         ElseIf txt1(1) = "3" Then
            pub_QL05 = pub_QL05 & ";" & Label1(1) & "3.申請案號" 'Add By Sindy 2010/12/10
         ElseIf txt1(1) = "4" Then
            pub_QL05 = pub_QL05 & ";" & Label1(1) & "4.證書號數" 'Add By Sindy 2010/12/10
         End If
         
         '申請書 1
         If Val(txt1(12)) = 1 Then
            '催實審, 催審
            If txt1(1) = "1" Or txt1(1) = "2" Then
               ProcessToWord1
            '申請案號, 證書號數
            Else
               ProcessToWord
            End If
         '管制表 2
         Else
            '催實審, 催審
            If Val(txt1(1)) = 1 Or Val(txt1(1)) = 2 Then
              Process
              '紀錄(管制表)前次列印條件
              '101-->第一碼：催審函性質，第二碼：查詢條件，第三碼：日期起或迄
              If txt1(1) = "1" Or txt1(1) = "2" Then
                 '實審期限
                 If opt(0).Value = True Then
                     'Modified by Morgan 2013/5/21
                     'SaveSetting "TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "01", txt1(2).Text
                     'SaveSetting "TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "02", txt1(3).Text
                     'Label2(0).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "01", "")
                     'Label2(1).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "02", "")
                     PUB_SaveLastDate Me.Name, "DATE" & txt1(1) & "01", txt1(2).Text
                     PUB_SaveLastDate Me.Name, "DATE" & txt1(1) & "02", txt1(3).Text
                     Label2(0).Caption = PUB_GetLastDate(Me.Name, "DATE" & txt1(1) & "01")
                     Label2(1).Caption = PUB_GetLastDate(Me.Name, "DATE" & txt1(1) & "02")
                     'end 2013/5/21
                 '發文日期
                 Else
                    'Modified by Morgan 2013/5/21
                    'SaveSetting "TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "11", txt1(13).Text
                    'SaveSetting "TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "12", txt1(14).Text
                    'Label2(2).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "11", "")
                    'Label2(3).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "12", "")
                     PUB_SaveLastDate Me.Name, "DATE" & txt1(1) & "11", txt1(13).Text
                     PUB_SaveLastDate Me.Name, "DATE" & txt1(1) & "12", txt1(14).Text
                     Label2(2).Caption = PUB_GetLastDate(Me.Name, "DATE" & txt1(1) & "11")
                     Label2(3).Caption = PUB_GetLastDate(Me.Name, "DATE" & txt1(1) & "12")
                 End If
              End If
            End If
         End If
         Me.Enabled = True
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

'Remove by Morgan 2005/1/13 程式流程調整幅度大故改寫
'Private Sub cmdok_Click(Index As Integer)
'Select Case Index
'Case 0
'
'      Printer.Orientation = 2
'      DoEvents
'
''   'Add By Cheng 2002/09/16
''   blnClkSure = False
'
'     '系統類別
'     If Len(txt1(0)) = 0 Then
'        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
'        txt1(0).SetFocus
'        Exit Sub
'     Else
'        '催審函性質
'        If Len(txt1(1)) = 0 Then
'            s = MsgBox("催審函性質不可空白!!", , "USER 輸入錯誤")
'            txt1(1).SetFocus
'            Exit Sub
'        Else
'            'Add By Cheng 2002/08/06
'            '選擇催審期限
'            If Me.opt(0).Value Then
'               If ChkDate(Me.txt1(2).Text) = False Then
'                  Me.txt1(2).SetFocus
'                  Exit Sub
'               End If
'               If ChkDate(Me.txt1(3).Text) = False Then
'                  Me.txt1(3).SetFocus
'                  Exit Sub
'               End If
'               If Val(Me.txt1(2).Text) > Val(Me.txt1(3).Text) Then
'                  MsgBox "催審期限區間輸入錯誤!!!", vbExclamation + vbOKOnly
'                  blnClkSure = True
'                  Me.txt1(2).SetFocus
'                  txt1_GotFocus 2
'                  Exit Sub
'               End If
'            '選擇發文日期
'            ElseIf Me.opt(1).Value Then
'               If ChkDate(Me.txt1(13).Text) = False Then
'                  Me.txt1(13).SetFocus
'                  Exit Sub
'               End If
'               If ChkDate(Me.txt1(14).Text) = False Then
'                  Me.txt1(14).SetFocus
'                  Exit Sub
'               End If
'               If Val(Me.txt1(13).Text) > Val(Me.txt1(14).Text) Then
'                  MsgBox "發文日期區間輸入錯誤!!!", vbExclamation + vbOKOnly
'                  blnClkSure = True
'                  Me.txt1(13).SetFocus
'                  txt1_GotFocus 13
'                  Exit Sub
'               End If
''Remove by Morgan 2004/11/12 本所號改為一般條件
''            '選擇本所案號
''            ElseIf Me.opt(2).Value Then
''               If Me.txt1(4).Text = "" Then
''                  MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
''                  Me.txt1(4).SetFocus
''                  Exit Sub
''               End If
''               If Me.txt1(5).Text = "" Then
''                  MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
''                  Me.txt1(5).SetFocus
''                  Exit Sub
''               End If
''2004/11/12 end
'            End If
'
'
'
''            'Add By Cheng 2002/03/20
''            If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
''               Me.txt1(2).SetFocus
''               txt1_GotFocus 2
''               Exit Sub
''            End If
''            If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
''               Me.txt1(3).SetFocus
''               txt1_GotFocus 3
''               Exit Sub
''            End If
''            'If Len(txt1(3)) = 0 And txt1(1) = "2" Then 'modify by sonia 90.9.29
''            If Len(txt1(3)) = 0 And txt1(1) = "2" And Len(txt1(4)) = 0 Then
''                s = MsgBox("催審期限區間不可空白!!", , "USER 輸入錯誤")
''                If Len(txt1(2)) = 0 Then txt1(2).SetFocus
''                Exit Sub
''            Else
'
'
'                '列印別
'                If Len(txt1(12)) = 0 Then
'                    s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
'                    Me.txt1(12).SetFocus
'                    Exit Sub
'                Else
'
'                    If Len(txt1(8)) <> 0 Or Len(txt1(9)) <> 0 Then
'                        If Mid(txt1(8), 1, 6) <> Mid(txt1(9), 1, 6) Then
'                           s = MsgBox("申請人前六碼必須相同!!", , "USER 輸入錯誤")
'                           blnClkSure = True
'                           txt1(8).SetFocus
'                           txt1_GotFocus (8)
'                           Exit Sub
'                        End If
'                    End If
'                  'Add By Cheng 2002/09/16
'                  If Me.txt1(8).Text <> "" And Me.txt1(9).Text <> "" Then
'                     If Me.txt1(8).Text > Me.txt1(9).Text Then
'                        MsgBox "申請人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
'                        blnClkSure = True
'                        Me.txt1(8).SetFocus
'                        txt1_GotFocus 8
'                        Exit Sub
'                     End If
'                  End If
'
'                    If Len(txt1(10)) <> 0 Or Len(txt1(11)) <> 0 Then
'                        If Mid(txt1(10), 1, 6) <> Mid(txt1(11), 1, 6) Then
'                           s = MsgBox("代理人前六碼必須相同!!", , "USER 輸入錯誤")
'                           blnClkSure = True
'                           txt1(10).SetFocus
'                           txt1_GotFocus (10)
'                           Exit Sub
'                        End If
'                    End If
'                  'Add By Cheng 2002/09/16
'                  If Me.txt1(10).Text <> "" And Me.txt1(11).Text <> "" Then
'                     If Me.txt1(10).Text > Me.txt1(11).Text Then
'                        MsgBox "代理人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
'                        blnClkSure = True
'                        Me.txt1(10).SetFocus
'                        txt1_GotFocus 10
'                        Exit Sub
'                     End If
'                  End If
'
'                    Select Case Val(txt1(12))
'                    Case 1 '申請書
'
''Remove by Morgan 2004/11/12 本所號改為一般條件
''                         If Len(txt1(4)) = 0 Or Len(txt1(5)) = 0 Then
''                              s = MsgBox("本所案號不可空白！！", , "USER 輸入錯誤！！")
''                              If Me.opt(2).Value Then
''                                 txt1(4).SetFocus
''                                 txt1_GotFocus (4)
''                              End If
''                              Exit Sub
''                         Else
''2004/11/12 end
'                              Screen.MousePointer = vbHourglass
'                              '催實審, 催審
'                                'Modify By Cheng 2002/12/16
''                              If txt1(1) = "2" Then
'                              If txt1(1) = "1" Or txt1(1) = "2" Then
'                                 ProcessToWord1
'                              '申請案號, 證書號數
'                              Else
'                                 ProcessToWord
'                              End If
'                              Screen.MousePointer = vbDefault
'
''                         End If
'
'                    Case 2 '管制表
'                            'Modify By Cheng 2002/12/16
'                            '恢復催實審
''                         If Val(txt1(1)) = 2 Then
'                         '催實審, 催審
'                         If Val(txt1(1)) = 1 Or Val(txt1(1)) = 2 Then
'                            Screen.MousePointer = vbHourglass
'                            Me.Enabled = False
'                            Process
'
'                           'Add by Morgan 2004/8/9
'                           '紀錄(管制表)前次列印條件
'                           '101-->第一碼：催審函性質，第二碼：查詢條件，第三碼：日期起或迄
'                           If txt1(1) = "1" Or txt1(1) = "2" Then
'                              '實審期限
'                              If opt(0).Value = True Then
'                                 SaveSetting "TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "01", txt1(2).Text
'                                 SaveSetting "TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "02", txt1(3).Text
'                                 Label2(0).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "01", "")
'                                 Label2(1).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "02", "")
'                              '發文日期
'                              Else
'                                 SaveSetting "TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "11", txt1(13).Text
'                                 SaveSetting "TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "12", txt1(14).Text
'                                 Label2(2).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "11", "")
'                                 Label2(3).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "12", "")
'                              End If
'                           End If
'                           'Add end
'
'                            Me.Enabled = True
'                            Screen.MousePointer = vbDefault
'                        '申請案號, 證書號數
'                         Else
'                            s = MsgBox("催審函性質為'申請案號或證書號數'時,列印別不可為'管制表'", , "USER 輸入錯誤")
'                            Exit Sub
'                         End If
'                    Case Else
'                    End Select
'                End If
''            End If
'        End If
'     End If
'Case 1
'     Unload Me
'Case Else
'End Select
'End Sub

Sub ProcessToWord1()
cnnConnection.Execute "delete FROM R060310 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL6 = ""

If Len(Trim(txt1(8))) <> 0 And Len(Trim(txt1(9))) <> 0 Then
    strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(txt1(8)) & "' AND PA26<='" & GetNewFagent(txt1(9)) & "') OR (PA27>='" & GetNewFagent(txt1(8)) & "' AND PA27<='" & GetNewFagent(txt1(9)) & "') OR (PA28>='" & GetNewFagent(txt1(8)) & "' AND PA28<='" & GetNewFagent(txt1(9)) & "') OR (PA29>='" & GetNewFagent(txt1(8)) & "' AND PA29<='" & GetNewFagent(txt1(9)) & "') OR (PA30>='" & GetNewFagent(txt1(8)) & "' AND PA30<='" & GetNewFagent(txt1(9)) & "')) "
    strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(txt1(8)) & "' AND SP08<='" & GetNewFagent(txt1(9)) & "') OR (SP58<='" & GetNewFagent(txt1(8)) & "' AND SP58<='" & GetNewFagent(txt1(9)) & "') OR (SP59>='" & GetNewFagent(txt1(8)) & "' AND SP59<='" & GetNewFagent(txt1(9)) & "')) "
Else
    If Len(Trim(txt1(8))) <> 0 And Len(Trim(txt1(9))) = 0 Then
        strSQL1 = strSQL1 + " AND (PA26>='" & GetNewFagent(txt1(8)) & "' OR PA27>='" & GetNewFagent(txt1(8)) & "' OR PA28>='" & GetNewFagent(txt1(8)) & "' OR PA29>='" & GetNewFagent(txt1(8)) & "' OR PA30>='" & GetNewFagent(txt1(8)) & "') "
        strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(8)) & "' OR SP58>='" & GetNewFagent(txt1(8)) & "' OR SP59>='" & GetNewFagent(txt1(8)) & "') "
    Else
        If Len(Trim(txt1(8))) = 0 And Len(Trim(txt1(9))) <> 0 Then
            strSQL1 = strSQL1 + " AND (PA26<='" & GetNewFagent(txt1(9)) & "' OR PA27<='" & GetNewFagent(txt1(9)) & "' OR PA28<='" & GetNewFagent(txt1(9)) & "' OR PA29<='" & GetNewFagent(txt1(9)) & "' OR PA30<='" & GetNewFagent(txt1(9)) & "') "
            strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(9)) & "' OR SP58<='" & GetNewFagent(txt1(9)) & "' OR SP59<='" & GetNewFagent(txt1(9)) & "') "
        End If
    End If
End If
If Len(Trim(txt1(8))) <> 0 Or Len(Trim(txt1(9))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(8) & "-" & txt1(9) 'Add By Sindy 2010/12/10
End If
'代理人
If Len(Trim(txt1(10))) <> 0 And Len(Trim(txt1(11))) <> 0 Then
    strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(10)) & "' AND PA75<='" & GetNewFagent(txt1(11)) & "' "
    strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(10)) & "' AND SP26<='" & GetNewFagent(txt1(11)) & "' "
Else
    If Len(Trim(txt1(10))) <> 0 And Len(Trim(txt1(11))) = 0 Then
        strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(10)) & "' "
        strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(10)) & "' "
    Else
        If Len(Trim(txt1(10))) = 0 And Len(Trim(txt1(11))) <> 0 Then
            strSQL1 = strSQL1 + " AND PA75<='" & GetNewFagent(txt1(11)) & "' "
            strSQL2 = strSQL2 + " AND SP26<='" & GetNewFagent(txt1(11)) & "' "
        End If
    End If
End If
If Len(Trim(txt1(10))) <> 0 Or Len(Trim(txt1(11))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(10) & "-" & txt1(11) 'Add By Sindy 2010/12/10
End If
strSQL1 = strSQL1 + " AND PA57 IS NULL "
strSQL2 = strSQL2 + " AND SP15 IS NULL "
CheckOC

'選擇催審期限
If Me.opt(0).Value Then
   If Len(txt1(4)) <> 0 Then
       StrSQL6 = StrSQL6 & " AND NP02='" & txt1(4) & "' "
   End If
   If Len(txt1(5)) <> 0 Then
       StrSQL6 = StrSQL6 & " AND NP03='" & txt1(5) & "' "
   End If
   If Len(txt1(6)) <> 0 Then
       StrSQL6 = StrSQL6 & " AND NP04='" & txt1(6) & "' "
   End If
   If Len(txt1(7)) <> 0 Then
       StrSQL6 = StrSQL6 & " AND NP05='" & txt1(7) & "' "
   End If
   If Len(txt1(4)) <> 0 And Len(txt1(5)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(9) & txt1(4) & "-" & txt1(5) & "-" & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/10
   End If
   '2005/5/23 ADD BY SONIA
   '案件性質
   If txt1(15) <> "" Then
      StrSQL6 = StrSQL6 + " AND CP10>='" & txt1(15) & "'"
   End If
   If txt1(16) <> "" Then
      StrSQL6 = StrSQL6 + " AND CP10<='" & txt1(16) & "'"
   End If
   If txt1(15) <> "" Or txt1(16) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(10) & txt1(15) & "-" & txt1(16) 'Add By Sindy 2010/12/10
   End If
   '2005/5/23 end
   
   StrSQL6 = StrSQL6 + " AND CP22 IS NULL "
   
   If Len(txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " AND NP02 IN (" & SQLGrpStr(txt1(0), 1) & ") "
      strSQL2 = strSQL2 + " AND NP02 IN (" & SQLGrpStr(txt1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/10
   End If
   If Len(txt1(2)) <> 0 Then
      StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(txt1(2))) & " "
   End If
   StrSQL6 = StrSQL6 + " AND NP08<=" & Val(ChangeTStringToWString(txt1(3)))
   pub_QL05 = pub_QL05 & ";" & opt(0).Caption & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/10
   
   Select Case Val(txt1(1))
      Case 2
           'MODFIY BY SONIA 2013/7/25 FCP-028070
           'StrSQL6 = StrSQL6 & " AND NP07=411 AND (NP06 IS NULL OR NP06=' ') " '催審
           StrSQL6 = StrSQL6 & " AND NP07 in (411,1503) AND (NP06 IS NULL OR NP06=' ') " '催審
      Case Else
   End Select
   
   strSql = "SELECT CP27,NP08,PA11,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(PA05,NVL(PA06,PA07)),NVL(PTM03,PTM04),decode(pa09,'000',CPM03,CPM04),S1.ST02,S2.ST02,NP01,CP10,NP02,NP03,NP04,NP05,NP22,NP07,NP09,NP01 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE NP01=CP09(+) AND np10=S2.ST01(+) AND cp14=S1.ST01(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=cpm01(+) AND np07=to_number(CPM02(+)) AND '1'=PTM01(+) AND pa08=PTM02(+) " & strSQL1 & StrSQL6
   strSql = strSql + " UNION ALL SELECT CP27,NP08,SP11,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),'',decode(sp09,'000',CPM03,CPM04),S1.ST02,S2.ST02,NP01,CP10,NP02,NP03,NP04,NP05,NP22,NP07,NP09,NP01 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP WHERE NP01=CP09(+) AND np10=S2.ST01(+) AND cp14=S1.ST01(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP02=cpm01(+) AND np07=to_number(CPM02(+)) " & strSQL2 & StrSQL6
'2005/5/23 ADD BY SONIA
'選擇發文日期或本所案號((若案件國家檔的實查時間"CF05"為NULL, 則不管制)
Else
   If Len(txt1(4)) <> 0 Then
       StrSQL6 = StrSQL6 & " AND CP01='" & txt1(4) & "' "
   End If
   If Len(txt1(5)) <> 0 Then
       StrSQL6 = StrSQL6 & " AND CP02='" & txt1(5) & "' "
   End If
   If Len(txt1(6)) <> 0 Then
       StrSQL6 = StrSQL6 & " AND CP03='" & txt1(6) & "' "
   End If
   If Len(txt1(7)) <> 0 Then
       StrSQL6 = StrSQL6 & " AND CP04='" & txt1(7) & "' "
   End If
   If Len(txt1(4)) <> 0 And Len(txt1(5)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(9) & txt1(4) & "-" & txt1(5) & "-" & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/10
   End If
   '2005/5/23 ADD BY SONIA
   '案件性質
   If txt1(15) <> "" Then
      StrSQL6 = StrSQL6 + " AND CP10>='" & txt1(15) & "'"
   End If
   If txt1(16) <> "" Then
      StrSQL6 = StrSQL6 + " AND CP10<='" & txt1(16) & "'"
   End If
   If txt1(15) <> "" Or txt1(16) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(10) & txt1(15) & "-" & txt1(16) 'Add By Sindy 2010/12/10
   End If
   '2005/5/23 end
   StrSQL6 = StrSQL6 + " AND CP22 IS NULL "
   
   If Len(txt1(0)) <> 0 Then
      strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
      strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/10
   End If
   pub_QL05 = pub_QL05 & ";" & opt(1).Caption & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/12/10
   StrSQL6 = StrSQL6 & " AND CP27>=" & DBDATE(Me.txt1(13).Text) & " AND CP27<=" & DBDATE(Me.txt1(14).Text) & " "
   StrSQL6 = StrSQL6 & " AND CP27 IS NOT NULL AND CP24 IS NULL AND CP57 IS NULL AND CP09 <'C' "
   
   If txt1(1) <> "1" Then
      strSql = "SELECT CP27,'',PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),NVL(PTM03,PTM04),decode(pa09,'000',CPM03,CPM04),S1.ST02,S2.ST02,CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,PATENTTRADEMARKMAP,CASEFEE WHERE CP13=S2.ST01(+) AND cp14=S1.ST01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=cpm01(+) AND Cp10=to_number(CPM02(+)) AND '1'=PTM01(+) AND pa08=PTM02(+) AND PA01=CF01(+) AND PA09=CF02(+) AND CP10=CF03 AND CF05 IS NOT NULL " & strSQL1 & StrSQL6
      strSql = strSql + " UNION ALL SELECT CP27,'',SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),'',decode(sp09,'000',CPM03,CPM04),S1.ST02,S2.ST02,CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CASEFEE WHERE Cp13=S2.ST01(+) AND cp14=S1.ST01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=cpm01(+) AND Cp10=to_number(CPM02(+)) AND SP01=CF01(+) AND SP09=CF02(+) AND CP10=CF03 AND CF05 IS NOT NULL " & strSQL2 & StrSQL6
   Else
      strSql = "SELECT CP27,'',PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),NVL(PTM03,PTM04)" & _
         ",decode(pa09,'000',CPM03,CPM04),S1.ST02,S2.ST02,CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09" & _
         " FROM CASEPROGRESS CP1,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,PATENTTRADEMARKMAP,CASEFEE" & _
         " WHERE CP10 IN ('101','102','103','105','107','301','302','303','305','306','307','801','802','803','804')" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
         " AND S2.ST01(+)=CP13 AND S1.ST01(+)=cp14 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
         " AND PTM01(+)='1' AND PTM02(+)=pa08 AND CF01(+)=PA01 AND CF02(+)=PA09 AND CF03=CP10 AND CF05 IS NOT NULL" & _
         " AND NOT EXISTS(SELECT * FROM CASEPROGRESS CP2 WHERE CP2.CP43=CP1.CP09 AND CP2.CP10 IN ('1204','1217'))" & _
         " AND ( (CP10='101' AND EXISTS(SELECT * FROM CASEPROGRESS CP3 WHERE CP3.CP01=CP1.CP01 AND CP3.CP02=CP1.CP02" & _
         " AND CP3.CP03=CP1.CP03 AND CP3.CP04=CP1.CP04 AND CP3.CP10='416' AND CP3.CP27 IS NOT NULL AND CP3.CP57 IS NULL" & _
         " AND SYSDATE>ADD_MONTHS(TO_DATE(CP3.CP27,'YYYYMMDD'),9)))" & _
         " OR (CP10 IN ('102','103','105') AND SYSDATE>ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),9))" & _
         " OR (CP10 IN ('107','301','302','303','305','306','307','801','802','803','804')" & _
         " AND SYSDATE>ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),3))) " & strSQL1 & StrSQL6
   End If
End If
'2005/5/23 END
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/10
        .MoveFirst
        DoEvents
        Do While .EOF = False
            s = 0
            For i = 0 To 18
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Modify By Cheng 2002/12/16
            '恢復催實審
'            If Val(txt1(1)) = 2 Then
            If Val(txt1(1)) = 1 Or Val(txt1(1)) = 2 Then
                If strTemp(13) <> "0" Then
                    strSql = "SELECT PA16 FROM PATENT WHERE PA01='" & ChgSQL(strTemp(11)) & "' AND PA02='" & ChgSQL(strTemp(12)) & "' AND PA03='0' AND PA04='" & ChgSQL(strTemp(14)) & "' "
                    CheckOC2
                    adoRecordset1.CursorLocation = adUseClient
                    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                        If CheckStr(adoRecordset1.Fields(0)) = "1" Then
                        Else
                            s = 1
                        End If
                    End If
                    CheckOC2
                End If
                If s <> 1 Then
                     'Modify By Cheng 2002/04/12
'                    strSQL = "SELECT CP05 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(11)) & "' AND CP02='" & ChgSQL(strTemp(12)) & "' AND CP03='" & ChgSQL(strTemp(13)) & "' AND CP04='" & ChgSQL(strTemp(14)) & "' AND CP10='1201' AND SUBSTR(CP09,1,1)='C' ORDER BY CP05"
                    strSql = "SELECT CP05 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(11)) & "' AND CP02='" & ChgSQL(strTemp(12)) & "' AND CP03='" & ChgSQL(strTemp(13)) & "' AND CP04='" & ChgSQL(strTemp(14)) & "' AND CP10='1201' AND CP09>'C' ORDER BY CP05"
                    CheckOC2
                    adoRecordset1.CursorLocation = adUseClient
                    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                        adoRecordset1.MoveLast
                        'Modify By Cheng 2002/04/12
'                        strSQL = "SELECT CP09,CP05,CP27 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(11)) & "' AND CP02='" & ChgSQL(strTemp(12)) & "' AND CP03='" & ChgSQL(strTemp(13)) & "' AND CP04='" & ChgSQL(strTemp(14)) & "' AND CP10='204' AND (SUBSTR(CP09,1,1)='B' OR SUBSTR(CP09,1,1)='A')  ORDER BY CP05 "
                        strSql = "SELECT CP09,CP05,CP27 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(11)) & "' AND CP02='" & ChgSQL(strTemp(12)) & "' AND CP03='" & ChgSQL(strTemp(13)) & "' AND CP04='" & ChgSQL(strTemp(14)) & "' AND CP10='204' AND ( CP09<'C' )  ORDER BY CP05 "
                        CheckOC2
                        adoRecordset1.CursorLocation = adUseClient
                        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                            adoRecordset1.MoveLast
                            If IsNull(adoRecordset1.Fields(2)) Then
                                s = 1
                            End If
                        Else
                           s = 1
                        End If
                    
                    End If
                    CheckOC2
                End If
                If s <> 1 Then
                     'Modify By Cheng 2002/04/12
'                    strSQL = "SELECT CP09,CP05,CP27 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(11)) & "' AND CP02='" & ChgSQL(strTemp(12)) & "' AND CP03='" & ChgSQL(strTemp(13)) & "' AND CP04='" & ChgSQL(strTemp(14)) & "' AND CP10='1905' AND SUBSTR(CP09,1,1)='C' ORDER BY CP05 "
                    strSql = "SELECT CP09,CP05,CP27 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(11)) & "' AND CP02='" & ChgSQL(strTemp(12)) & "' AND CP03='" & ChgSQL(strTemp(13)) & "' AND CP04='" & ChgSQL(strTemp(14)) & "' AND CP10='1905' AND CP09>'C' ORDER BY CP05 "
                    CheckOC2
                    adoRecordset1.CursorLocation = adUseClient
                    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                        If DateDiff("M", ChangeWStringToWDateString(CheckStr(adoRecordset1.Fields(1))), ChangeWStringToWDateString(GetTodayDate)) < 3 Then
                            s = 1
                        End If
                    End If
                End If
                CheckOC2
            End If
            If s = 0 Then
                'strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                'strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                'cnnConnection.Execute "INSERT INTO R060310 VALUES('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "','" & chgsql(strTemp(2)) & "','" & chgsql(strTemp(3)) & "','" & chgsql(strTemp(4)) & "','" & chgsql(strTemp(5)) & "','" & chgsql(strTemp(6)) & "','" & chgsql(strTemp(7)) & "','" & chgsql(strTemp(8)) & "','" & strUserNum & "','" & chgsql(strTemp(18)) & "') "
                '11  NP02
                '12  NP03
                '13  NP04
                '14  NP05
                '16  NP07
                '17  NP01
                PrintLetter CheckStr(.Fields(16)), CheckStr(.Fields(11)), CheckStr(.Fields(2)), CheckStr(.Fields(0)), GetTodayDate, CheckStr(.Fields(18)), "", ""
                If txt1(12) = "1" Then
                  If Len(CheckStr(.Fields(1))) <> 8 Then
                      SavDay(0) = CheckStr(.Fields(1))
                  Else
                      SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                  End If
                  If Len(CheckStr(.Fields(17))) <> 8 Then
                      SavDay(1) = CheckStr(.Fields(17))
                  Else
                      SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(adoRecordset.Fields(17)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                  End If
                  'Modify By Cheng 2002/08/07
                  '取消更新下一程序的期限
'                  cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(18)) & "' AND NP07=" & Val(CheckStr(.Fields(16))) & " AND NP22=" & Val(CheckStr(.Fields(15)))
                End If
            End If
            .MoveNext
        Loop
    End With
Else
    InsertQueryLog (0) 'Add By Sindy 2010/12/10
    ShowNoData
    CheckOC
    Exit Sub
End If
'PrintData
ShowPrintOk
End Sub

Sub ProcessToWord()
strSQL1 = ""
If Len(txt1(4)) <> 0 Then
   strSQL1 = strSQL1 & " AND CP01='" & txt1(4) & "' "
End If
If Len(txt1(5)) <> 0 Then
   strSQL1 = strSQL1 & " AND CP02='" & txt1(5) & "' "
End If
If Len(txt1(6)) <> 0 Then
   strSQL1 = strSQL1 & " AND CP03='" & txt1(6) & "' "
End If
If Len(txt1(7)) <> 0 Then
   strSQL1 = strSQL1 & " AND CP04='" & txt1(7) & "'"
End If
If Len(txt1(4)) <> 0 And Len(txt1(5)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(9) & txt1(4) & "-" & txt1(5) & "-" & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/10
End If
strSQL1 = strSQL1 + " AND PA57 IS NULL  AND CP22 IS NULL "
strSql = ""
CheckOC
Dim StrCP26BY601 As String
With adoRecordset
   strSql = "SELECT CP09,CP10,CP27,PA11,PA22 FRoM CASEPROGRESS,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL1
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/10
      StrCP26BY601 = ""
      CheckOC2
      strSql = "SELECT CP27 FRoM CASEPROGRESS,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01 IN (" & SQLGrpStr("", 1) & ") AND CP10='601' " & strSQL1
      adoRecordset1.CursorLocation = adUseClient
      adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If adoRecordset1.RecordCount <> 0 Then
         StrCP26BY601 = CheckStr(adoRecordset1.Fields(0))
      End If
      CheckOC2
      PrintLetter CheckStr(.Fields(1)), txt1(4), CheckStr(.Fields(3)), CheckStr(.Fields(2)), GetTodayDate, CheckStr(.Fields(0)), CheckStr(.Fields(4)), StrCP26BY601
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/10
      ShowNoData
      Exit Sub
   End If
End With
ShowPrintOk

End Sub

Private Sub InsExpField(ByVal strCP10 As String, ByVal strCP01 As String, ByVal StrPA11 As String, ByVal strCP27 As String, ByVal strSysDate As String, ByVal strCP09 As String, ByVal strPA22 As String, ByVal str601CP27 As String)
   'Dim strSQL As String
   ' 下一程序
   'Dim StrCP10 As Strubg
   ' 系統別
   'Dim strCP01 As String
   ' 申請案號
   'Dim strPA11 As String
   ' 發文日
   'Dim strCP27 As String
   ' 系統日
   'Dim strSysDate As String
   ' 總收文號
   'Dim strCP09 As String
   '專利號數
   'DIM strCP22 AS String
   '領證發文日
   'DIM str601CP27 AS String
   Select Case txt1(1)
      '催實審, 催審
        'Modify By Cheng 2002/12/16
'      Case "2"
      '2005/5/23 MODIFY BY SONIA 將催實審, 催審分成二個選項
      'Case "1", "2"
      '      EndLetter "11", strCP09, "01", strUserNum
      '      'str((Val(Mid(strSysDate, 1, 4)) * 12 + Val(Mid(strSysDate, 5, 2))) - (Val(Mid(strCP27, 1, 4)) * 12 + Val(Mid(strCP27, 5, 2))))
      '      strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('11','" & strCP09 & "','01','" & strUserNum & "','已發文時間','" & str((Val(Mid(strSysDate, 1, 4)) * 12 + Val(Mid(strSysDate, 5, 2))) - (Val(Mid(strCP27, 1, 4)) * 12 + Val(Mid(strCP27, 5, 2)))) & "')"
      '      cnnConnection.Execute strSQL
      '催實審
      Case "1"
            EndLetter "11", strCP09, "04", strUserNum
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('11','" & strCP09 & "','04','" & strUserNum & "','已發文時間','" & str((Val(Mid(strSysDate, 1, 4)) * 12 + Val(Mid(strSysDate, 5, 2))) - (Val(Mid(strCP27, 1, 4)) * 12 + Val(Mid(strCP27, 5, 2)))) & "')"
            cnnConnection.Execute strSql
      '催審
      Case "2"
            EndLetter "11", strCP09, "01", strUserNum
            'str((Val(Mid(strSysDate, 1, 4)) * 12 + Val(Mid(strSysDate, 5, 2))) - (Val(Mid(strCP27, 1, 4)) * 12 + Val(Mid(strCP27, 5, 2))))
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('11','" & strCP09 & "','01','" & strUserNum & "','已發文時間','" & str((Val(Mid(strSysDate, 1, 4)) * 12 + Val(Mid(strSysDate, 5, 2))) - (Val(Mid(strCP27, 1, 4)) * 12 + Val(Mid(strCP27, 5, 2)))) & "')"
            cnnConnection.Execute strSql
      '2005/5/23 END
      'MODIFY BY SONIA 90.10.20 將申請案號,證書號數分開二種選項
      '申請案號/證書號數
      'Case "3"
      '     If Len(strPA11) = 0 Then
      '         EndLetter "11", strCP09, "02", strUserNum
      '     Else
      '         If Len(strPA22) = 0 Then
      '            EndLetter "11", strCP09, "03", strUserNum
      '            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('11','" & strCP09 & "','03','" & strUserNum & "','領證發文日','" & str601CP27 & "')"
      '            cnnConnection.Execute strSQL
      '         End If
      '     End If
      '申請案號
      Case "3"
            EndLetter "11", strCP09, "02", strUserNum
      '證書號數
      Case "4"
            EndLetter "11", strCP09, "03", strUserNum
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('11','" & strCP09 & "','03','" & strUserNum & "','領證發文日','" & str601CP27 & "')"
            cnnConnection.Execute strSql
      Case Else
   End Select
   '清除定稿例外欄位檔原有資料
   'EndLetter "10", strCP09, "05", strUserNum
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter(ByVal strCP10 As String, ByVal strCP01 As String, ByVal StrPA11 As String, ByVal strCP27 As String, ByVal strSysDate As String, ByVal strCP09 As String, ByVal strPA22 As String, ByVal str601CP27 As String)
   'Dim strSQL As String
   ' 下一程序
   'Dim StrCP10 As Strubg
   ' 系統別
   'Dim strCP01 As String
   ' 申請案號
   'Dim strPA11 As String
   ' 發文日
   'Dim strCP27 As String
   ' 系統日
   'Dim strSysDate As String
   ' 總收文號
   'Dim strCP09 As String
   '專利號數
   'DIM strCP22 AS String
   '領證發文日
   'DIM str601CP27 AS String
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField strCP10, strCP01, StrPA11, strCP27, strSysDate, strCP09, strPA22, str601CP27
   Select Case txt1(1)
      '催實審, 催審
        'Modify By Cheng 2002/12/16
'      Case "2"
      '2005/5/23 MODIFY BY SONIA 將催實審, 催審分成二個選項
      'Case "1", "2"
      '      NowPrint strCP09, "11", "01", False, strUserNum, 0 '
      '催實審
      Case "1"
            NowPrint strCP09, "11", "04", False, strUserNum, 0 '
      '催審
      Case "2"
            NowPrint strCP09, "11", "01", False, strUserNum, 0 '
      '2005/5/23 END
      'MODIFY BY SONIA 90.10.20 將申請案號,證書號數分成二個選項
      '申請案號/證書號數
      'Case "3"
      '     If Len(strPA11) = 0 Then
      '         NowPrint strCP09, "11", "02", False, strUserNum, "0"
      '     Else
      '         If Len(strPA22) = 0 Then
      '            NowPrint strCP09, "11", "03", False, strUserNum, 0 'str601CP27
      '         End If
      '     End If
      '申請案號
      Case "3"
            NowPrint strCP09, "11", "02", False, strUserNum, "0"
      '證書號數
      Case "4"
            NowPrint strCP09, "11", "03", False, strUserNum, 0 'str601CP27
      Case Else
   End Select
End Sub

Sub Process()
   'Add By Cheng 2002/08/07
   Dim blnUpdate As Boolean '是否要更新下一程序的期限
   'Add by Morgan 2009/5/25 變數改用區域宣告以免被覆蓋
   Dim adoProc As ADODB.Recordset, adoProc1 As ADODB.Recordset, intR As Integer
   
 '911106 nick transation
On Error GoTo CheckingErr
   
   cnnConnection.BeginTrans
   
   cnnConnection.Execute "delete FROM R060310 WHERE ID='" & strUserNum & "' "
   strSQL1 = ""
   strSQL2 = ""
   StrSQL6 = ""
   
'Remove by Morgan 2005/1/12 併到下方
'   '系統類別
'   If Len(txt1(0)) <> 0 Then
'      'Modify By Cheng 2002/08/22
'      If Me.opt(0).Value Then
'         strSQL1 = strSQL1 + " AND NP02 IN (" & SQLGrpStr(txt1(0), 1) & ") "
'         strSQL2 = strSQL2 + " AND NP02 IN (" & SQLGrpStr(txt1(0), 5) & ") "
'      Else
'         strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
'         strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
'      End If
'   End If
'
'   'Modify By Cheng 2002/08/22
'   '催審期限
'   If Me.opt(0).Value Then
'      StrSQL6 = " AND (NP06 IS NULL OR NP06=' ') "
'   End If
'
'   'Modify By Cheng 2002/08/06
'   If Me.opt(0).Value Then
'      If Len(txt1(2)) <> 0 Then
'         StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(txt1(2))) & " "
'      End If
'      StrSQL6 = StrSQL6 + " AND NP08<=" & Val(ChangeTStringToWString(txt1(3)))
'   End If
'   'Add By Cheng 2002/08/06
'   If Me.opt(1).Value Then
'      StrSQL6 = StrSQL6 & " AND CP27>=" & DBDATE(Me.txt1(13).Text) & " AND CP27<=" & DBDATE(Me.txt1(14).Text) & " "
'   End If
'2005/1/12 end

'Remove by Morgan 2005/1/12 併到下方
'   Select Case Val(txt1(1))
'      Case 1
'         'Modify By Cheng 2002/08/22
'         If Me.opt(0).Value Then
'           StrSQL6 = StrSQL6 & " AND NP07=1204 "
'         End If
'      Case 2
'         'Modify By Cheng 2002/08/22
'         If Me.opt(0).Value Then
'           StrSQL6 = StrSQL6 & " AND NP07=411 "
'         End If
'      Case Else
'   End Select
'2005/1/12 end

   
   'Modify By Cheng 2002/08/06
   'If Len(txt1(4)) <> 0 Then
   '    strsql6 = strsql6 & " AND NP02='" & txt1(4) & "' "
   'End If
   'If Len(txt1(5)) <> 0 Then
   '    strsql6 = strsql6 & " AND NP03='" & txt1(5) & "' "
   'End If
   'If Len(txt1(6)) <> 0 Then
   '    strsql6 = strsql6 & " AND NP04='" & txt1(6) & "' "
   'End If
   'If Len(txt1(7)) <> 0 Then
   '    strsql6 = strsql6 & " AND NP05='" & txt1(7) & "' "
   'End If
   If Len(Trim(txt1(8))) <> 0 And Len(Trim(txt1(9))) <> 0 Then
       strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(txt1(8)) & "' AND PA26<='" & GetNewFagent(txt1(9)) & "') OR (PA27>='" & GetNewFagent(txt1(8)) & "' AND PA27<='" & GetNewFagent(txt1(9)) & "') OR (PA28>='" & GetNewFagent(txt1(8)) & "' AND PA28<='" & GetNewFagent(txt1(9)) & "') OR (PA29>='" & GetNewFagent(txt1(8)) & "' AND PA29<='" & GetNewFagent(txt1(9)) & "') OR (PA30>='" & GetNewFagent(txt1(8)) & "' AND PA30<='" & GetNewFagent(txt1(9)) & "')) "
       strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(txt1(8)) & "' AND SP08<='" & GetNewFagent(txt1(9)) & "') OR (SP58<='" & GetNewFagent(txt1(8)) & "' AND SP58<='" & GetNewFagent(txt1(9)) & "') OR (SP59>='" & GetNewFagent(txt1(8)) & "' AND SP59<='" & GetNewFagent(txt1(9)) & "')) "
   Else
       If Len(Trim(txt1(8))) <> 0 And Len(Trim(txt1(9))) = 0 Then
           strSQL1 = strSQL1 + " AND (PA26>='" & GetNewFagent(txt1(8)) & "' OR PA27>='" & GetNewFagent(txt1(8)) & "' OR PA28>='" & GetNewFagent(txt1(8)) & "' OR PA29>='" & GetNewFagent(txt1(8)) & "' OR PA30>='" & GetNewFagent(txt1(8)) & "') "
           strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(8)) & "' OR SP58>='" & GetNewFagent(txt1(8)) & "' OR SP59>='" & GetNewFagent(txt1(8)) & "') "
       Else
           If Len(Trim(txt1(8))) = 0 And Len(Trim(txt1(9))) <> 0 Then
               strSQL1 = strSQL1 + " AND (PA26<='" & GetNewFagent(txt1(9)) & "' OR PA27<='" & GetNewFagent(txt1(9)) & "' OR PA28<='" & GetNewFagent(txt1(9)) & "' OR PA29<='" & GetNewFagent(txt1(9)) & "' OR PA30<='" & GetNewFagent(txt1(9)) & "') "
               strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(9)) & "' OR SP58<='" & GetNewFagent(txt1(9)) & "' OR SP59<='" & GetNewFagent(txt1(9)) & "') "
           End If
       End If
   End If
   If Len(Trim(txt1(8))) <> 0 Or Len(Trim(txt1(9))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(8) & "-" & txt1(9) 'Add By Sindy 2010/12/10
   End If
   '代理人
   If Len(Trim(txt1(10))) <> 0 And Len(Trim(txt1(11))) <> 0 Then
       strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(10)) & "' AND PA75<='" & GetNewFagent(txt1(11)) & "' "
       strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(10)) & "' AND SP26<='" & GetNewFagent(txt1(11)) & "' "
   Else
       If Len(Trim(txt1(10))) <> 0 And Len(Trim(txt1(11))) = 0 Then
           strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(10)) & "' "
           strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(10)) & "' "
       Else
           If Len(Trim(txt1(10))) = 0 And Len(Trim(txt1(11))) <> 0 Then
               strSQL1 = strSQL1 + " AND PA75<='" & GetNewFagent(txt1(11)) & "' "
               strSQL2 = strSQL2 + " AND SP26<='" & GetNewFagent(txt1(11)) & "' "
           End If
       End If
   End If
   If Len(Trim(txt1(10))) <> 0 Or Len(Trim(txt1(11))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(10) & "-" & txt1(11) 'Add By Sindy 2010/12/10
   End If
   strSQL1 = strSQL1 + " AND PA57 IS NULL "
   strSQL2 = strSQL2 + " AND SP15 IS NULL "
   
   '選擇催審期限
   If Me.opt(0).Value Then
      'Remove by Morgan 2004/11/12 本所號改為一般條件
      'If Me.opt(2).Value Then
      
         If Me.txt1(4).Text <> "" Then
            StrSQL6 = StrSQL6 & " AND NP02='" & Me.txt1(4).Text & "' "
         End If
         If Me.txt1(5).Text <> "" Then
            StrSQL6 = StrSQL6 & " AND NP03='" & Me.txt1(5).Text & "' "
         End If
         If Me.txt1(6).Text <> "" Then
            StrSQL6 = StrSQL6 & " AND NP04='" & Me.txt1(6).Text & "' "
         End If
         If Me.txt1(7).Text <> "" Then
            StrSQL6 = StrSQL6 & " AND NP05='" & Me.txt1(7).Text & "' "
         End If
         If Me.txt1(4).Text <> "" And Me.txt1(5).Text <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label1(9) & txt1(4) & "-" & txt1(5) & "-" & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/10
         End If
      'End If
      'Add by Morgan 2005/1/12
      '案件性質
      If txt1(15) <> "" Then
         StrSQL6 = StrSQL6 + " AND CP10>='" & txt1(15) & "'"
      End If
      If txt1(16) <> "" Then
         StrSQL6 = StrSQL6 + " AND CP10<='" & txt1(16) & "'"
      End If
      If txt1(15) <> "" Or txt1(16) <> "" Then
         pub_QL05 = pub_QL05 & ";" & Label1(10) & txt1(15) & "-" & txt1(16) 'Add By Sindy 2010/12/10
      End If
      '2005/1/12 end
   
      StrSQL6 = StrSQL6 + " AND CP22 IS NULL "
   
      '系統類別
      If Len(txt1(0)) <> 0 Then
         strSQL1 = strSQL1 + " AND NP02 IN (" & SQLGrpStr(txt1(0), 1) & ") "
         strSQL2 = strSQL2 + " AND NP02 IN (" & SQLGrpStr(txt1(0), 5) & ") "
         pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/10
      End If
      If Len(txt1(2)) <> 0 Then
         StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(txt1(2))) & " "
      End If
      StrSQL6 = StrSQL6 + " AND NP08<=" & Val(ChangeTStringToWString(txt1(3)))
      pub_QL05 = pub_QL05 & ";" & opt(0).Caption & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/10
      '2005/1/12 end
      
      If Val(txt1(1)) = 2 Then
         'MODFIY BY SONIA 2013/7/25 FCP-028070
         'StrSQL6 = StrSQL6 & " AND NP07=411 AND (NP06 IS NULL OR NP06=' ') "
         StrSQL6 = StrSQL6 & " AND NP07 in (411,1503) AND (NP06 IS NULL OR NP06=' ') "
      End If
      '抓下一程序檔資料
      'Modified by Lydia 2017/02/13 +FMP管制人
      If strSrvDate(1) < FMP管制人啟用日 Then
         'modify by sonia 2019/10/14 NP02=cpm01(+) AND np07=to_number(CPM02(+)改為CP01=cpm01(+) AND CP10=to_number(CPM02(+)
         strSql = "SELECT CP27,NP08,PA11,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(PA05,NVL(PA06,PA07)),NVL(PTM03,PTM04),decode(pa09,'000',CPM03,CPM04),S1.ST02,DECODE(PA75,'',S3.ST02,S2.ST02),NP01,CP10,NP02,NP03,NP04,NP05,NP22,NP07,NP09,NP01 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION N1,NATION N2,FAGENT,CUSTOMER WHERE NP01=CP09(+) AND cp14=S1.ST01(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND CP01=cpm01(+) AND CP10=to_number(CPM02(+)) AND '1'=PTM01(+) AND pa08=PTM02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=N1.NA01(+) AND N1.NA16=S2.ST01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND CU10=N2.NA01(+) AND N2.NA16=S3.ST01(+) " & strSQL1 & StrSQL6
         strSql = strSql + " UNION ALL SELECT CP27,NP08,SP11,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),'',decode(sp09,'000',CPM03,CPM04),S1.ST02,DECODE(SP26,'',S3.ST02,S2.ST02),NP01,CP10,NP02,NP03,NP04,NP05,NP22,NP07,NP09,NP01 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,NATION N1,NATION N2,FAGENT,CUSTOMER WHERE NP01=CP09(+) AND cp14=S1.ST01(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND CP01=cpm01(+) AND CP10=to_number(CPM02(+)) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND FA10=N1.NA01(+) AND N1.NA16=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CU10=N2.NA01(+) AND N2.NA16=S3.ST01(+) " & strSQL2 & StrSQL6
      Else
         'modify by sonia 2019/10/14 NP02=cpm01(+) AND np07=to_number(CPM02(+)改為CP01=cpm01(+) AND CP10=to_number(CPM02(+)
         strSql = "SELECT CP27 C01,NP08 C02,PA11 C03,NP02||'-'||NP03||'-'||NP04||'-'||NP05 C04,NVL(PA05,NVL(PA06,PA07)) C05,NVL(PTM03,PTM04) C06,decode(pa09,'000',CPM03,CPM04) C07,S1.ST02 C08,DECODE(PA75,'',DECODE(PA01,'P',NVL(N2.NA79,N2.NA16),N2.NA16),DECODE(PA01,'P',NVL(N1.NA79,N1.NA16),N1.NA16)) C09,NP01 C10,CP10 C11,NP02 C12,NP03 C13,NP04 C14,NP05 C15,NP22 C16,NP07 C17,NP09 C18,NP01 C19 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,STAFF S1,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION N1,NATION N2,FAGENT,CUSTOMER WHERE NP01=CP09(+) AND cp14=S1.ST01(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND CP01=cpm01(+) AND CP10=to_number(CPM02(+)) AND '1'=PTM01(+) AND pa08=PTM02(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=N1.NA01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND CU10=N2.NA01(+) " & strSQL1 & StrSQL6
         strSql = strSql + " UNION ALL SELECT CP27,NP08,SP11,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),'',decode(sp09,'000',CPM03,CPM04),S1.ST02,DECODE(SP26,'',DECODE(SP01,'PS',NVL(N2.NA79,N2.NA16),N2.NA16),DECODE(SP01,'PS',NVL(N1.NA79,N1.NA16),N1.NA16)) C09,NP01,CP10,NP02,NP03,NP04,NP05,NP22,NP07,NP09,NP01 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,CASEPROPERTYMAP,NATION N1,NATION N2,FAGENT,CUSTOMER WHERE NP01=CP09(+) AND cp14=S1.ST01(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND CP01=cpm01(+) AND CP10=to_number(CPM02(+)) AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND FA10=N1.NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CU10=N2.NA01(+) " & strSQL2 & StrSQL6
         strSql = "SELECT C01,C02,C03,C04,C05,C06,C07,C08,NVL(ST02,C09) C09,C10,C11,C12,C13,C14,C15,C16,C17,C18,C19 FROM (" + strSql + ") ,STAFF WHERE C09=ST01(+)"
      End If
      'end 2017/02/13
      
   '選擇發文日期或本所案號((若案件國家檔的實查時間"CF05"為NULL, 則不管制)
   Else
      If Len(txt1(4)) <> 0 Then
          StrSQL6 = StrSQL6 & " AND CP01='" & txt1(4) & "' "
      End If
      If Len(txt1(5)) <> 0 Then
          StrSQL6 = StrSQL6 & " AND CP02='" & txt1(5) & "' "
      End If
      If Len(txt1(6)) <> 0 Then
          StrSQL6 = StrSQL6 & " AND CP03='" & txt1(6) & "' "
      End If
      If Len(txt1(7)) <> 0 Then
          StrSQL6 = StrSQL6 & " AND CP04='" & txt1(7) & "' "
      End If
      If Len(txt1(4)) <> 0 And Len(txt1(5)) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label1(9) & txt1(4) & "-" & txt1(5) & "-" & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/10
      End If
      '2005/5/23 ADD BY SONIA
      '案件性質
      If txt1(15) <> "" Then
         StrSQL6 = StrSQL6 + " AND CP10>='" & txt1(15) & "'"
      End If
      If txt1(16) <> "" Then
         StrSQL6 = StrSQL6 + " AND CP10<='" & txt1(16) & "'"
      End If
      If txt1(15) <> "" Or txt1(16) <> "" Then
         pub_QL05 = pub_QL05 & ";" & Label1(10) & txt1(15) & "-" & txt1(16) 'Add By Sindy 2010/12/10
      End If
      '2005/5/23 end
      StrSQL6 = StrSQL6 + " AND CP22 IS NULL "
  
      'Add by Morgan 2005/1/12
      '發文日期
      '系統類別
      If Len(txt1(0)) <> 0 Then
         strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
         strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
         pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/10
      End If
      pub_QL05 = pub_QL05 & ";" & opt(1).Caption & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/12/10
      StrSQL6 = StrSQL6 & " AND CP27>=" & DBDATE(Me.txt1(13).Text) & " AND CP27<=" & DBDATE(Me.txt1(14).Text) & " "
      StrSQL6 = StrSQL6 & " AND CP27 IS NOT NULL AND CP24 IS NULL AND CP57 IS NULL AND CP09 <'C' "
      '2005/1/12 end
   
      'Modify by Morgan 2005/1/12
      '1.'101',有'416'且發文日+9個月<系統日
      '2.'102','103','105',發文日+9個月<系統日
      '3.'107','301','302','303','305','306','307','801','802','803','804',發文日+3個月<系統日
      
       '抓案件進度檔資料
'      'Modify By Cheng 2002/08/22
'   '   strSQL = "SELECT CP27,NP08,PA11,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(PA05,NVL(PA06,PA07)),NVL(PTM03,PTM04),decode(pa09,'000',CPM03,CPM04),S1.ST02,S2.ST02,NP01,CP10,NP02,NP03,NP04,NP05,NP22,NP07,NP09,NP01 FROM NEXTPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,PATENTTRADEMARKMAP,CASEFEE WHERE NP01=CP09(+) AND np10=S2.ST01(+) AND cp14=S1.ST01(+) AND NP02=PA01(+) AND NP03=PA02(+) AND NP04=PA03(+) AND NP05=PA04(+) AND NP02=cpm01(+) AND np07=to_number(CPM02(+)) AND '1'=PTM01(+) AND pa08=PTM02(+) AND CP01=CF01 AND PA09=CF02 AND CP10=CF03 AND CF05 IS NOT NULL " & strsql1 & strsql6
'   '   strSQL = strSQL + " UNION ALL SELECT CP27,NP08,SP11,NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(SP05,NVL(SP06,SP07)),'',decode(sp09,'000',CPM03,CPM04),S1.ST02,S2.ST02,NP01,CP10,NP02,NP03,NP04,NP05,NP22,NP07,NP09,NP01 FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CASEFEE WHERE NP01=CP09(+) AND np10=S2.ST01(+) AND cp14=S1.ST01(+) AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) AND NP02=cpm01(+) AND np07=to_number(CPM02(+)) AND CP01=CF01 AND SP09=CF02 AND CP10=CF03 AND CF05 IS NOT NULL " & strsql2 & strsql6
'   '92.04.02 nick add left join
'   '   strSQL = "SELECT CP27,'',PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),NVL(PTM03,PTM04),decode(pa09,'000',CPM03,CPM04),S1.ST02,S2.ST02,CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,PATENTTRADEMARKMAP,CASEFEE WHERE CP13=S2.ST01(+) AND cp14=S1.ST01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=cpm01(+) AND Cp10=to_number(CPM02(+)) AND '1'=PTM01(+) AND pa08=PTM02(+) AND CP01=CF01 AND PA09=CF02 AND CP10=CF03 AND CF05 IS NOT NULL " & strSQL1 & StrSQL6
'   '   strSQL = strSQL + " UNION ALL SELECT CP27,'',SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),'',decode(sp09,'000',CPM03,CPM04),S1.ST02,S2.ST02,CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CASEFEE WHERE Cp13=S2.ST01(+) AND cp14=S1.ST01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=cpm01(+) AND Cp10=to_number(CPM02(+)) AND CP01=CF01 AND SP09=CF02 AND CP10=CF03 AND CF05 IS NOT NULL " & strSQL2 & StrSQL6
'      StrSql = "SELECT CP27,'',PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),NVL(PTM03,PTM04),decode(pa09,'000',CPM03,CPM04),S1.ST02,S2.ST02,CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,PATENTTRADEMARKMAP,CASEFEE WHERE CP13=S2.ST01(+) AND cp14=S1.ST01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=cpm01(+) AND Cp10=to_number(CPM02(+)) AND '1'=PTM01(+) AND pa08=PTM02(+) AND CP01=CF01(+) AND PA09=CF02(+) AND CP10=CF03(+) AND CF05 IS NOT NULL " & strSQL1 & StrSQL6
'      StrSql = StrSql + " UNION ALL SELECT CP27,'',SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),'',decode(sp09,'000',CPM03,CPM04),S1.ST02,S2.ST02,CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CASEFEE WHERE Cp13=S2.ST01(+) AND cp14=S1.ST01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=cpm01(+) AND Cp10=to_number(CPM02(+)) AND CP01=CF01(+) AND SP09=CF02(+) AND CP10=CF03(+) AND CF05 IS NOT NULL " & strSQL2 & StrSQL6
      '2005/5/19 MODIFY BY SONIA
      'strSQL = "SELECT CP27,'',PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),NVL(PTM03,PTM04)" & _
      '   ",decode(pa09,'000',CPM03,CPM04),S1.ST02,S2.ST02,CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09" & _
      '   " FROM CASEPROGRESS CP1,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,PATENTTRADEMARKMAP,CASEFEE" & _
      '   " WHERE CP10 IN ('101','102','103','105','107','301','302','303','305','306','307','801','802','803','804')" & _
      '   " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
      '   " AND S2.ST01(+)=CP13 AND S1.ST01(+)=cp14 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      '   " AND PTM01(+)='1' AND PTM02(+)=pa08 AND CF01(+)=PA01 AND CF02(+)=PA09 AND CF03=CP10 AND CF05 IS NOT NULL" & _
      '   " AND NOT EXISTS(SELECT * FROM CASEPROGRESS CP2 WHERE CP2.CP09=CP1.CP43 AND CP2.CP10 IN ('1204','1217'))" & _
      '   " AND ( (CP10='101' AND EXISTS(SELECT * FROM CASEPROGRESS CP3 WHERE CP3.CP01=CP1.CP01 AND CP3.CP02=CP1.CP02" & _
      '   " AND CP3.CP03=CP1.CP03 AND CP3.CP04=CP1.CP04 AND CP3.CP10='416' AND CP3.CP27 IS NOT NULL AND CP3.CP57 IS NULL" & _
      '   " AND SYSDATE>ADD_MONTHS(TO_DATE(CP3.CP27,'YYYYMMDD'),9)))" & _
      '   " OR (CP10 IN ('102','103','105') AND SYSDATE>ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),9))" & _
      '   " OR (CP10 IN ('107','301','302','303','305','306','307','801','802','803','804')" & _
      '   " AND EXISTS(SELECT * FROM CASEPROGRESS CP3 WHERE CP3.CP01=CP1.CP01 AND CP3.CP02=CP1.CP02" & _
      '   " AND CP3.CP03=CP1.CP03 AND CP3.CP04=CP1.CP04 AND CP3.CP10='416' AND CP3.CP27 IS NOT NULL AND CP3.CP57 IS NULL" & _
      '   " AND SYSDATE>ADD_MONTHS(TO_DATE(CP3.CP27,'YYYYMMDD'),3)))) " & strSQL1 & StrSQL6
      If txt1(1) <> "1" Then
         'Modified by Lydia 2017/02/13 +FMP管制人
         If strSrvDate(1) < FMP管制人啟用日 Then
            strSql = "SELECT CP27,'',PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),NVL(PTM03,PTM04),decode(pa09,'000',CPM03,CPM04),S1.ST02,DECODE(PA75,'',S3.ST02,S2.ST02),CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09 FROM CASEPROGRESS,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,PATENTTRADEMARKMAP,CASEFEE,NATION N1,NATION N2,FAGENT,CUSTOMER WHERE cp14=S1.ST01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=cpm01(+) AND Cp10=to_number(CPM02(+)) AND '1'=PTM01(+) AND pa08=PTM02(+) AND PA01=CF01(+) AND PA09=CF02(+) AND CP10=CF03 AND CF05 IS NOT NULL AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=N1.NA01(+) AND N1.NA16=S2.ST01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND CU10=N2.NA01(+) AND N2.NA16=S3.ST01(+) " & strSQL1 & StrSQL6
            strSql = strSql + " UNION ALL SELECT CP27,'',SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),'',decode(sp09,'000',CPM03,CPM04),S1.ST02,DECODE(SP26,'',S3.ST02,S2.ST02),CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,CASEFEE,NATION N1,NATION N2,FAGENT,CUSTOMER WHERE cp14=S1.ST01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=cpm01(+) AND Cp10=to_number(CPM02(+)) AND SP01=CF01(+) AND SP09=CF02(+) AND CP10=CF03 AND CF05 IS NOT NULL AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND FA10=N1.NA01(+) AND N1.NA16=S2.ST01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CU10=N2.NA01(+) AND N2.NA16=S3.ST01(+) " & strSQL2 & StrSQL6
         Else
            strSql = "SELECT CP27 C01,'' C02,PA11 C03,CP01||'-'||CP02||'-'||CP03||'-'||CP04 C04,NVL(PA05,NVL(PA06,PA07)) C05,NVL(PTM03,PTM04) C06,decode(pa09,'000',CPM03,CPM04) C07,S1.ST02 C08,DECODE(PA75,'',DECODE(PA01,'P',NVL(N2.NA79,N2.NA16),N2.NA16),DECODE(PA01,'P',NVL(N1.NA79,N1.NA16),N1.NA16)) C09,CP09 C10,CP10 C11,CP01 C12,CP02 C13,CP03 C14,CP04 C15,'' C16,'' C17,'' C18,CP09 C19 FROM CASEPROGRESS,PATENT,STAFF S1,CASEPROPERTYMAP,PATENTTRADEMARKMAP,CASEFEE,NATION N1,NATION N2,FAGENT,CUSTOMER WHERE cp14=S1.ST01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=cpm01(+) AND Cp10=to_number(CPM02(+)) AND '1'=PTM01(+) AND pa08=PTM02(+) AND PA01=CF01(+) AND PA09=CF02(+) AND CP10=CF03 AND CF05 IS NOT NULL AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=N1.NA01(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND CU10=N2.NA01(+) " & strSQL1 & StrSQL6
            strSql = strSql + " UNION ALL SELECT CP27,'',SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),'',decode(sp09,'000',CPM03,CPM04),S1.ST02,DECODE(SP26,'',DECODE(SP01,'PS',NVL(N2.NA79,N2.NA16),N2.NA16),DECODE(SP01,'PS',NVL(N1.NA79,N1.NA16),N1.NA16)) C09,CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09 FROM CASEPROGRESS,SERVICEPRACTICE,STAFF S1,CASEPROPERTYMAP,CASEFEE,NATION N1,NATION N2,FAGENT,CUSTOMER WHERE cp14=S1.ST01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=cpm01(+) AND Cp10=to_number(CPM02(+)) AND SP01=CF01(+) AND SP09=CF02(+) AND CP10=CF03 AND CF05 IS NOT NULL AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND FA10=N1.NA01(+) AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND CU10=N2.NA01(+) " & strSQL2 & StrSQL6
            strSql = "SELECT C01,C02,C03,C04,C05,C06,C07,C08,NVL(ST02,C09) C09,C10,C11,C12,C13,C14,C15,C16,C17,C18,C19 FROM (" + strSql + ") ,STAFF WHERE C09=ST01(+)"
         End If
         'end 2017/02/13
      Else
         strSql = "SELECT CP27,'',PA11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),NVL(PTM03,PTM04)" & _
            ",decode(pa09,'000',CPM03,CPM04),S1.ST02,S2.ST02,CP09,CP10,CP01,CP02,CP03,CP04,'','','',CP09" & _
            " FROM CASEPROGRESS CP1,PATENT,STAFF S1,STAFF S2,CASEPROPERTYMAP,PATENTTRADEMARKMAP,CASEFEE" & _
            " WHERE CP10 IN ('101','102','103','105','107','301','302','303','305','306','307','801','802','803','804')" & _
            " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
            " AND S2.ST01(+)=CP13 AND S1.ST01(+)=cp14 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
            " AND PTM01(+)='1' AND PTM02(+)=pa08 AND CF01(+)=PA01 AND CF02(+)=PA09 AND CF03=CP10 AND CF05 IS NOT NULL" & _
            " AND NOT EXISTS(SELECT * FROM CASEPROGRESS CP2 WHERE CP2.CP43=CP1.CP09 AND CP2.CP10 IN ('1204','1217'))" & _
            " AND ( (CP10='101' AND EXISTS(SELECT * FROM CASEPROGRESS CP3 WHERE CP3.CP01=CP1.CP01 AND CP3.CP02=CP1.CP02" & _
            " AND CP3.CP03=CP1.CP03 AND CP3.CP04=CP1.CP04 AND CP3.CP10='416' AND CP3.CP27 IS NOT NULL AND CP3.CP57 IS NULL" & _
            " AND SYSDATE>ADD_MONTHS(TO_DATE(CP3.CP27,'YYYYMMDD'),9)))" & _
            " OR (CP10 IN ('102','103','105') AND SYSDATE>ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),9))" & _
            " OR (CP10 IN ('107','301','302','303','305','306','307','801','802','803','804')" & _
            " AND SYSDATE>ADD_MONTHS(TO_DATE(CP27,'YYYYMMDD'),3))) " & strSQL1 & StrSQL6
      End If
      '2005/5/19 END
   End If
   
   intR = 1
   Set adoProc = ClsLawReadRstMsg(intR, strSql)
   If intR = 1 Then
       With adoProc
            InsertQueryLog (adoProc.RecordCount) 'Add By Sindy 2010/12/10
            'Add By Cheng 2002/08/07
            '若選擇催審期限, 加詢問使用者是否要更新
            If Me.opt(0).Value Then
               If MsgBox("是否再管制三個月?", vbExclamation + vbYesNo) = vbYes Then
                  blnUpdate = True
               Else
                  blnUpdate = False
               End If
            '其他選項一律不更新
            Else
               blnUpdate = False
            End If
           
           .MoveFirst
           DoEvents
           Do While .EOF = False
               s = 0
               For i = 0 To 18
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               'Modify By Cheng 2002/08/22
               If Me.opt(0).Value Then
                   'Modify By Cheng 2002/12/16
                   '恢復催實審
   '               If Val(txt1(1)) = 2 Then
                  If Val(txt1(1)) = 1 Or Val(txt1(1)) = 2 Then
                      If strTemp(13) <> "0" Then
                        strSql = "SELECT PA16 FROM PATENT WHERE PA01='" & ChgSQL(strTemp(11)) & "' AND PA02='" & ChgSQL(strTemp(12)) & "' AND PA03='0' AND PA04='" & ChgSQL(strTemp(14)) & "' "
                        intR = 1
                        Set adoProc1 = ClsLawReadRstMsg(intR, strSql)
                        If intR = 1 Then
                           If CheckStr(adoProc1.Fields(0)) <> "1" Then
                              s = 1
                           End If
                        End If
                      End If
                      If s <> 1 Then
                        strSql = "SELECT CP05 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(11)) & "' AND CP02='" & ChgSQL(strTemp(12)) & "' AND CP03='" & ChgSQL(strTemp(13)) & "' AND CP04='" & ChgSQL(strTemp(14)) & "' AND CP10='1201' AND CP09>'C' ORDER BY CP05"
                        intR = 1
                        Set adoProc1 = ClsLawReadRstMsg(intR, strSql)
                        If intR = 1 Then
                           adoProc1.MoveLast
                           strSql = "SELECT CP09,CP05,CP27 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(11)) & "' AND CP02='" & ChgSQL(strTemp(12)) & "' AND CP03='" & ChgSQL(strTemp(13)) & "' AND CP04='" & ChgSQL(strTemp(14)) & "' AND CP10='204' AND ( CP09<'C' )  ORDER BY CP05 "
                           intR = 1
                           Set adoProc1 = ClsLawReadRstMsg(intR, strSql)
                           If intR = 1 Then
                               adoProc1.MoveLast
                               If IsNull(adoProc1.Fields(2)) Then
                                   s = 1
                               End If
                           Else
                              s = 1
                           End If
                          End If
                      End If
                      If s <> 1 Then
                           strSql = "SELECT CP09,CP05,CP27 FROM CASEPROGRESS WHERE CP01='" & ChgSQL(strTemp(11)) & "' AND CP02='" & ChgSQL(strTemp(12)) & "' AND CP03='" & ChgSQL(strTemp(13)) & "' AND CP04='" & ChgSQL(strTemp(14)) & "' AND CP10='1905' AND CP09>'C' ORDER BY CP05 "
                           intR = 1
                           Set adoProc1 = ClsLawReadRstMsg(intR, strSql)
                           If intR = 1 Then
                              'Modify by Morgan 2004/11/30 更正前後日期位置
                              If DateDiff("M", ChangeWStringToWDateString(CheckStr(adoProc1.Fields(1))), ChangeWStringToWDateString(GetTodayDate)) < 3 Then
                                  s = 1
                              End If
                          End If
                      End If
                  End If
               End If
               If s = 0 Then
                   strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
                   strTemp(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(1)))
                   cnnConnection.Execute "INSERT INTO R060310 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & strUserNum & "','" & ChgSQL(strTemp(18)) & "') "
                  '列印別為管制表時
                   If txt1(12) = "2" Then
                     If Len(CheckStr(.Fields(1))) <> 8 Then
                         SavDay(0) = CheckStr(.Fields(1))
                     Else
                         SavDay(0) = Format(DateAdd("M", 3, Format(Format(CheckStr(.Fields(1)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                     End If
                     If Len(CheckStr(.Fields(17))) <> 8 Then
                         SavDay(1) = CheckStr(.Fields(17))
                     Else
                         SavDay(1) = Format(DateAdd("M", 3, Format(Format(CheckStr(.Fields(17)), "####/##/##"), "yyyy/mm/dd")), "yyyymmdd")
                     End If
                     'Modify By Cheng 2002/08/07
                     If blnUpdate Then
                        cnnConnection.Execute "UPDATE NEXTPROGRESS SET NP08=" & Val(SavDay(0)) & ",NP09=" & Val(SavDay(1)) & " WHERE NP01='" & CheckStr(.Fields(18)) & "' AND NP07=" & Val(CheckStr(.Fields(16))) & " AND NP22=" & Val(CheckStr(.Fields(15)))
                     End If
                   End If
               End If
               .MoveNext
           Loop
       End With
   Else
       InsertQueryLog (0) 'Add By Sindy 2010/12/10
       cnnConnection.RollbackTrans
       ShowNoData
       CheckOC
       Exit Sub
   End If
   PrintData
   '911106 nick transation 邱小姐說在印完表後再更新資料，包括缺紙
   cnnConnection.CommitTrans
   GoTo ExitPort
   
CheckingErr:
   cnnConnection.RollbackTrans

ExitPort:
   Set adoProc = Nothing
   Set adoProc1 = Nothing
End Sub

Sub PrintData()
strSql = "SELECT * FROM R060310 WHERE ID='" & strUserNum & "' ORDER BY R043002,R043001,R043004 "
CheckOC
Page = 1
SavDay1 = " "
SavDay2 = " "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 8
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            If strTemp(1) = SavDay2 Then
                strTemp(1) = ""
                If strTemp(0) = SavDay1 Then
                    strTemp(0) = ""
                Else
                    SavDay1 = strTemp(0)
                End If
            Else
                SavDay2 = strTemp(1)
                SavDay1 = strTemp(0)
            End If
            strTemp(4) = StrToStr(strTemp(4), 14)
            strTemp(5) = StrToStr(strTemp(5), 4)
            strTemp(6) = StrToStr(strTemp(6), 4)
            strTemp(7) = StrToStr(strTemp(7), 8)
            strTemp(8) = StrToStr(strTemp(8), 3)
            PrintDatil
            If iPrint > 10000 Then
                Printer.CurrentX = 0
                Printer.CurrentY = iPrint
                Printer.Print String(250, "-")
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            .MoveNext
        Loop
    End With
Else
    ShowNoData
    Exit Sub
End If

Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
Printer.EndDoc
CheckOC
ShowPrintOk
End Sub

Sub PrintDatil()
For i = 0 To 8
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500 - 500
PLeft(1) = 1400 - 500 + 250
PLeft(2) = 2500 - 500 + 250
PLeft(3) = 4400 - 500
PLeft(4) = 6000
PLeft(5) = 9900
PLeft(6) = 11000
PLeft(7) = 12200
PLeft(8) = 13400
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 7500 - (Printer.TextWidth(GetTitleNick & "催 審 表") / 2)
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "催 審 表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
'Modify By Cheng 2002/08/22
'Printer.CurrentX = 7500 - (Printer.TextWidth("催審期限：" & Format(ChangeTStringToTDateString(txt1(2)), "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))) / 2)
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
'Modify By Cheng 2002/08/22
'Printer.Print "催審期限：" & Format(ChangeTStringToTDateString(txt1(2)), "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3))
Printer.Print IIf(Me.opt(0).Value, "催審期限：" & Format(ChangeTStringToTDateString(txt1(2)), "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(3)), _
               IIf(Me.opt(1).Value, "發文日期：" & Format(ChangeTStringToTDateString(txt1(13)), "@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(14)), _
               "本所案號：" & Me.txt1(4).Text & "-" & Left(Me.txt1(5).Text & "000000", 6) & "-" & Left(Me.txt1(6).Text & "0", 1) & "-" & Left(Me.txt1(7).Text & "00", 2)))
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人　：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
If txt1(1) = "1" Then
    Printer.Print "催審性質：催實審"
Else
    Printer.Print "催審性質：催審"
End If
Printer.CurrentX = 3000
Printer.CurrentY = iPrint
If Len(txt1(8)) <> 0 Then
    Printer.Print "申請人：" & GetPrjPeople1(txt1(8))
Else
    If Len(txt1(10)) <> 0 Then
        Printer.Print "代理人：" & GetPrjName1(txt1(10))
    End If
End If
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "催審期限"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "申請案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "專利種類"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "管制人員"
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick

'Add by Morgan 2004/8/9
Label2(0).Caption = ""
Label2(1).Caption = ""
Label2(2).Caption = ""
Label2(3).Caption = ""
'Add end

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm060310 = Nothing
End Sub

Private Sub opt_Click(Index As Integer)
'Add By Cheng 2002/08/06
Select Case Index
Case 0 '催審期限
   Me.opt(0).Value = True
   Me.txt1(2).Enabled = True
   Me.txt1(3).Enabled = True
   Me.txt1(2).SetFocus
   
   Me.opt(1).Value = False
   Me.txt1(13).Enabled = False
   Me.txt1(14).Enabled = False
   
'Remove by Morgan 2004/11/12
'   Me.opt(2).Value = False
'   Me.txt1(4).Enabled = False
'   Me.txt1(5).Enabled = False
'   Me.txt1(6).Enabled = False
'   Me.txt1(7).Enabled = False
'2004/11/12 end
   
Case 1 '發文日期
   Me.opt(0).Value = False
   Me.txt1(2).Enabled = False
   Me.txt1(3).Enabled = False
   
   Me.opt(1).Value = True
   Me.txt1(13).Enabled = True
   Me.txt1(14).Enabled = True
   Me.txt1(13).SetFocus
   
'Remove by Morgan 2004/11/12
'   Me.opt(2).Value = False
'   Me.txt1(4).Enabled = False
'   Me.txt1(5).Enabled = False
'   Me.txt1(6).Enabled = False
'   Me.txt1(7).Enabled = False
'
'Case 2 '本所案號
'   Me.opt(0).Value = False
'   Me.txt1(2).Enabled = False
'   Me.txt1(3).Enabled = False
'
'   Me.opt(1).Value = False
'   Me.txt1(13).Enabled = False
'   Me.txt1(14).Enabled = False
'
'
'   Me.opt(2).Value = True
'   Me.txt1(4).Enabled = True
'   Me.txt1(5).Enabled = True
'   Me.txt1(6).Enabled = True
'   Me.txt1(7).Enabled = True
'   Me.txt1(4).SetFocus
'2004/11/12 end

End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdOK(0).SetFocus
End If
End Sub
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/16
   Select Case Index
      Case 1
           'Modify By Cheng 2002/12/16
   '      If KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 8 Then
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
      Case 12
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
      'Add by Morgan 2005/1/12
      Case 13, 14, 15, 16
         If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
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
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
    Next i
Case 1
   If txt1(1) <> "" Then
     Select Case Val(txt1(1))
        'Modify By Cheng 2002/12/16
'     Case 2, 3, 4
     Case 1, 2, 3, 4
     Case Else
            'Modify By Cheng 2002/12/16
'          s = MsgBox("催審函性質只能輸入 2 或 3 或 4 !!", , "USER 輸入錯誤")
          s = MsgBox("催審函性質只能輸入 1 或 2 或 3 或 4 !!", , "USER 輸入錯誤")
          txt1(1).SetFocus
          txt1(1).SelStart = 0
          txt1(1).SelLength = Len(txt1(1))
          Exit Sub
     End Select
     
      'Add by Morgan 2004/8/9
      '管制表
      If txt1(12) = "2" Then
         '1催實審,2催審
         If txt1(1) = "1" Or txt1(1) = "2" Then
            'Modified by Morgan 2013/5/21
            'Label2(0).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "01", "")
            'Label2(1).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "02", "")
            'Label2(2).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "11", "")
            'Label2(3).Caption = GetSetting("TAIE", "FCP", Me.Name & "#DATE" & txt1(1) & "12", "")
            Label2(0).Caption = PUB_GetLastDate(Me.Name, "DATE" & txt1(1) & "01")
            Label2(1).Caption = PUB_GetLastDate(Me.Name, "DATE" & txt1(1) & "02")
            Label2(2).Caption = PUB_GetLastDate(Me.Name, "DATE" & txt1(1) & "11")
            Label2(3).Caption = PUB_GetLastDate(Me.Name, "DATE" & txt1(1) & "12")
            'end 2013/5/21
         End If
      End If
   End If
Case 3
   'Modify By Cheng 2002/09/16
   If blnClkSure = False Then
     If RunNick(txt1(2), txt1(3)) Then
         txt1(2).SetFocus
         txt1_GotFocus (2)
         Exit Sub
      End If
   Else
      blnClkSure = False
   End If
Case 4
     Select Case Trim(UCase(txt1(4)))
     Case "FG", "FCP", ""
     Case Else
          s = MsgBox("本所案號系統別只能 FG 或 FCP !!", , "USER 輸入錯誤")
          txt1(4).SetFocus
          txt1(4).SelStart = 0
          txt1(4).SelLength = Len(txt1(4))
          Exit Sub
     End Select
Case 9
   'Modify By Cheng 2002/09/16
   If blnClkSure = False Then
      If Len(txt1(Index - 1)) <> 0 Then
         If Left(txt1(Index - 1), 6) <> Left(txt1(Index), 6) Then
             s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
             txt1(Index - 1).SetFocus
             Exit Sub
         End If
      End If
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   Else
      blnClkSure = False
   End If
Case 11
   'Modify By Cheng 2002/09/16
   If blnClkSure = False Then
      If Len(txt1(Index - 1)) <> 0 Then
         If Left(txt1(Index - 1), 6) <> Left(txt1(Index), 6) Then
             s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
             txt1(Index - 1).SetFocus
             Exit Sub
         End If
      End If
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   Else
      blnClkSure = False
   End If
Case 12
   If txt1(Index) <> "" Then
     Select Case Val(txt1(12))
     Case 1, 2
     Case Else
          s = MsgBox("列印別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          txt1(12).SetFocus
          txt1(12).SelStart = 0
          txt1(12).SelLength = Len(txt1(12))
          Exit Sub
     End Select
   End If
'Add By Cheng 2002/08/06
Case 14
   'Modify By Cheng 2002/09/16
   If blnClkSure = False Then
     If RunNick(txt1(13), txt1(14)) Then
         txt1(13).SetFocus
         txt1_GotFocus (13)
         Exit Sub
      End If
   Else
      blnClkSure = False
   End If
Case Else
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 2, 3 '催審期限
         Cancel = Not ChkDate(txt1(Index))
         If Cancel Then TextInverse txt1(Index)
      'Add By Cheng 2002/08/06
      Case 13, 14 '發文日期
         Cancel = Not ChkDate(txt1(Index))
         If Cancel Then TextInverse txt1(Index)
   End Select
End Sub
