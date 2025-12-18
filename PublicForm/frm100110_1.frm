VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100110_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "爭議案件查詢"
   ClientHeight    =   4605
   ClientLeft      =   2655
   ClientTop       =   1725
   ClientWidth     =   5580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5580
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   15
      Left            =   2820
      MaxLength       =   7
      TabIndex        =   31
      Top             =   4200
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   14
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   30
      Top             =   4200
      Width           =   1092
   End
   Begin VB.OptionButton Option1 
      Caption         =   "法院案號："
      Height          =   180
      Index           =   4
      Left            =   96
      TabIndex        =   27
      Top             =   2592
      Width           =   1212
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   13
      Left            =   1380
      TabIndex        =   9
      Top             =   2532
      Width           =   2532
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   12
      Left            =   1380
      TabIndex        =   8
      Top             =   2200
      Width           =   2532
   End
   Begin VB.OptionButton Option1 
      Caption         =   "商品類別："
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   25
      Top             =   2260
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   10
      Left            =   2820
      MaxLength       =   9
      TabIndex        =   16
      Top             =   3860
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   1380
      MaxLength       =   9
      TabIndex        =   15
      Top             =   3860
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   2820
      MaxLength       =   9
      TabIndex        =   14
      Top             =   3528
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1380
      MaxLength       =   9
      TabIndex        =   13
      Top             =   3528
      Width           =   1092
   End
   Begin VB.CommandButton cmdGoInput 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4650
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   24
      Width           =   756
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3855
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1380
      TabIndex        =   1
      Top             =   480
      Width           =   2532
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1380
      TabIndex        =   5
      Top             =   1536
      Width           =   2532
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1380
      TabIndex        =   7
      Top             =   1868
      Width           =   2532
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   10
      Top             =   2864
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   2820
      MaxLength       =   7
      TabIndex        =   11
      Top             =   2864
      Width           =   1092
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   1380
      TabIndex        =   12
      Top             =   3196
      Width           =   2532
   End
   Begin VB.OptionButton Option1 
      Caption         =   "機關文號："
      Height          =   180
      Index           =   0
      Left            =   96
      TabIndex        =   0
      Top             =   540
      Value           =   -1  'True
      Width           =   1212
   End
   Begin VB.OptionButton Option1 
      Caption         =   "對造名稱："
      Height          =   180
      Index           =   1
      Left            =   96
      TabIndex        =   2
      Top             =   887
      Width           =   1212
   End
   Begin VB.OptionButton Option1 
      Caption         =   "條款："
      Height          =   180
      Index           =   2
      Left            =   96
      TabIndex        =   6
      Top             =   1928
      Width           =   915
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   330
      Index           =   1
      Left            =   1380
      TabIndex        =   4
      Top             =   1174
      Width           =   2535
      VariousPropertyBits=   679493659
      Size            =   "4471;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   330
      Index           =   0
      Left            =   1380
      TabIndex        =   3
      Top             =   812
      Width           =   2535
      VariousPropertyBits=   679493659
      Size            =   "4471;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2610
      X2              =   2730
      Y1              =   4350
      Y2              =   4350
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   29
      Top             =   4260
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "專利商標智商法院"
      Height          =   180
      Left            =   4080
      TabIndex        =   28
      Top             =   2592
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "機關文號欄可模糊比對, 約四分鐘才會有結果"
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   4080
      TabIndex        =   26
      Top             =   480
      Width           =   1350
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      Caption         =   "對造案件名稱："
      Height          =   180
      Left            =   48
      TabIndex        =   24
      Top             =   1249
      Width           =   1260
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2610
      X2              =   2730
      Y1              =   4005
      Y2              =   4005
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2610
      X2              =   2730
      Y1              =   3645
      Y2              =   3645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   23
      Top             =   3920
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   22
      Top             =   3588
      Width           =   720
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2610
      X2              =   2730
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "對造號數："
      Height          =   180
      Left            =   408
      TabIndex        =   21
      Top             =   1596
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "收文日期："
      Height          =   180
      Left            =   360
      TabIndex        =   20
      Top             =   2924
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   19
      Top             =   3256
      Width           =   900
   End
End
Attribute VB_Name = "frm100110_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/05 改成Form2.0 ; txt1(1)改成txtFM2(0)、txt1(11)改成txtFM2(1)
'Modified by Morgan 2021/8/12 Label2: 智財法院-->智商法院
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
Option Explicit

Dim s As Integer, i As Integer, j As Integer, strTemp As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

'92.04.16 nick
Public Sub PubShowNextData()

   Select Case cmdState
      Case 0 '確定
         cmdState = -1
         If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
            Me.txt1(4).SetFocus
            txt1_GotFocus 4
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(5)) = -1 Then
            Me.txt1(5).SetFocus
            txt1_GotFocus 5
            Exit Sub
         End If
         If Len(txt1(7)) <> 0 And Len(txt1(7)) > 6 Then
            If Len(txt1(8)) = 0 Or Len(txt1(8)) < 6 Then
                s = MsgBox("申請人代號須填入區間, 6 碼以上", , "USER 輸入錯誤")
                txt1(8).SetFocus
                txt1(8).SelStart = 0
                txt1(8).SelLength = Len(txt1(8))
                Exit Sub
            End If
            If Len(txt1(8)) >= 6 Then
               If Left(txt1(7), 6) <> Left(txt1(8), 6) Then
                   s = MsgBox("申請人代號前六碼須相同", , "USER 輸入錯誤")
                   txt1(8).SetFocus
                   txt1(8).SelStart = 0
                   txt1(8).SelLength = Len(txt1(8))
                   Exit Sub
                End If
             End If
         Else
             If Len(txt1(7)) < 6 And Len(txt1(8)) <> 0 Then
                s = MsgBox("申請人代號須填入區間, 6 碼以上", , "USER 輸入錯誤")
                txt1(7).SetFocus
                txt1(7).SelStart = 0
                txt1(7).SelLength = Len(txt1(7))
                Exit Sub
             End If
             If Len(txt1(7)) = 0 And Len(txt1(8)) <> 0 Then
                s = MsgBox("申請人代號須填入區間, 6 碼以上", , "USER 輸入錯誤")
                txt1(7).SetFocus
                txt1(7).SelStart = 0
                txt1(7).SelLength = Len(txt1(7))
                Exit Sub
             End If
         End If
         If Len(txt1(9)) <> 0 And Len(txt1(9)) > 6 Then
            If Len(txt1(10)) = 0 Or Len(txt1(10)) < 6 Then
                s = MsgBox("代理人代號須填入區間, 6 碼以上", , "USER 輸入錯誤")
                txt1(10).SetFocus
                txt1(10).SelStart = 0
                txt1(10).SelLength = Len(txt1(10))
                Exit Sub
            End If
            If Len(txt1(10)) >= 6 Then
               If Left(txt1(9), 6) <> Left(txt1(10), 6) Then
                   s = MsgBox("代理人代號前六碼須相同", , "USER 輸入錯誤")
                   txt1(10).SetFocus
                   txt1(10).SelStart = 0
                   txt1(10).SelLength = Len(txt1(10))
                   Exit Sub
                End If
             End If
         Else
             If Len(txt1(9)) < 6 And Len(txt1(10)) <> 0 Then
                s = MsgBox("代理人代號須填入區間, 6 碼以上", , "USER 輸入錯誤")
                txt1(9).SetFocus
                txt1(9).SelStart = 0
                txt1(9).SelLength = Len(txt1(9))
                Exit Sub
             End If
             If Len(txt1(9)) = 0 And Len(txt1(10)) <> 0 Then
                s = MsgBox("代理人代號須填入區間, 6 碼以上", , "USER 輸入錯誤")
                txt1(9).SetFocus
                txt1(9).SelStart = 0
                txt1(9).SelLength = Len(txt1(9))
                Exit Sub
             End If
         End If
         If Len(txt1(7)) >= 6 Then
            For i = 1 To 9 - Len(txt1(7))
                txt1(7) = txt1(7) + "0"
            Next i
         End If
         If Len(txt1(10)) >= 6 Then
            For i = 1 To 9 - Len(txt1(10))
               txt1(10) = txt1(10) + "0"
            Next i
         End If
         If Len(txt1(8)) >= 6 Then
            For i = 1 To 9 - Len(txt1(8))
               txt1(8) = txt1(8) + "0"
            Next i
         End If
         If Len(txt1(9)) >= 6 Then
            For i = 1 To 9 - Len(txt1(9))
               txt1(9) = txt1(9) + "0"
            Next i
         End If
       
         '選擇機關文號
         If Option1(0).Value = True Then
            If Len(Trim(txt1(0))) = 0 Then
               s = MsgBox("機關文號沒有輸入", , "USER 輸入錯誤")
               txt1(0).SetFocus
               txt1(0).SelStart = 0
               txt1(0).SelLength = Len(txt1(0))
               Exit Sub
            End If
            Me.Enabled = False
            If fnSaveParentForm(Me) = False Then
               Me.Enabled = True
               Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100110_2.Show
            frm100110_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         Else
             '選擇對造名稱
            If Option1(1).Value = True Then
               If Len(Trim(txtFM2(0))) = 0 And Len(Trim(txt1(2))) = 0 And Len(Trim(txtFM2(1))) = 0 Then
                  s = MsgBox("對造名稱, 案件名稱及號數皆沒有輸入", , "USER 輸入錯誤")
                  If Len(Trim(txtFM2(0))) = 0 Then
                     txtFM2(0).SetFocus
                     txtFM2(0).SelStart = 0
                     txtFM2(0).SelLength = Len(txtFM2(0))
                  ElseIf Len(Trim(txtFM2(1))) = 0 Then
                     txtFM2(1).SetFocus
                     txtFM2(1).SelStart = 0
                     txtFM2(1).SelLength = Len(txtFM2(1))
                  Else
                     txt1(2).SetFocus
                     txt1(2).SelStart = 0
                     txt1(2).SelLength = Len(txt1(2))
                  End If
                  Exit Sub
               End If
               Me.Enabled = False
               If fnSaveParentForm(Me) = False Then
                  Me.Enabled = True
                  Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               frm100110_3.Show
               frm100110_3.StrMenu
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            Else
               '選擇條款
               If Option1(2).Value = True Then
                  If Len(Trim(txt1(3))) = 0 Then
                     s = MsgBox("條款沒有輸入", , "USER 輸入錯誤")
                     txt1(3).SetFocus
                     txt1(3).SelStart = 0
                     txt1(3).SelLength = Len(txt1(3))
                     Exit Sub
                  End If
                  Me.Enabled = False
                  If fnSaveParentForm(Me) = False Then
                     Me.Enabled = True
                     Exit Sub
                  End If
                  Screen.MousePointer = vbHourglass
                  frm100110_4.Show
                  frm100110_4.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               Else
                  If Option1(3).Value = True Then
                     If Len(Trim(txt1(12))) = 0 Then
                        s = MsgBox("商品類別沒有輸入", , "USER 輸入錯誤")
                        txt1(12).SetFocus
                        txt1(12).SelStart = 0
                        txt1(12).SelLength = Len(txt1(12))
                        Exit Sub
                     End If
                     Me.Enabled = False
                     If fnSaveParentForm(Me) = False Then
                        Me.Enabled = True
                        Exit Sub
                     End If
                     Screen.MousePointer = vbHourglass
                     frm100110_4.Show
                     frm100110_4.StrMenu2
                     Screen.MousePointer = vbDefault
                     Me.Enabled = True
                     Exit Sub
                  '2013/8/19 add by sonia
                  Else
                     If Option1(4).Value = True Then
                        If Len(Trim(txt1(13))) = 0 Then
                           s = MsgBox("法院案號沒有輸入", , "USER 輸入錯誤")
                           txt1(13).SetFocus
                           txt1(13).SelStart = 0
                           txt1(13).SelLength = Len(txt1(13))
                           Exit Sub
                        End If
                        Me.Enabled = False
                        If fnSaveParentForm(Me) = False Then
                           Me.Enabled = True
                           Exit Sub
                        End If
                        Screen.MousePointer = vbHourglass
                        frm100110_2.Show
                        frm100110_2.StrMenu
                        Screen.MousePointer = vbDefault
                        Me.Enabled = True
                        Exit Sub
                     End If
                  '2013/8/19 end
                  End If
               End If
            End If
         End If
      Case 1
           fnCloseAllFrm100
      Case Else
   End Select
End Sub

Private Sub cmdGoInput_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
''92.04.16 nick 以下無效
'Select Case Index
'Case 0 '確定
'      'Add By Cheng 2002/03/18
'      If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
'         Me.txt1(4).SetFocus
'         txt1_GotFocus 4
'         Exit Sub
'      End If
'      If PUB_CheckKeyInDate(Me.txt1(5)) = -1 Then
'         Me.txt1(5).SetFocus
'         txt1_GotFocus 5
'         Exit Sub
'      End If
'
'      If Len(txt1(7)) <> 0 And Len(txt1(7)) > 6 Then
'         If Len(txt1(8)) = 0 Or Len(txt1(8)) < 6 Then
'             s = MsgBox("申請人代號須填入區間, 6 碼以上", , "USER 輸入錯誤")
'             txt1(8).SetFocus
'             txt1(8).SelStart = 0
'             txt1(8).SelLength = Len(txt1(8))
'             Exit Sub
'         End If
'         If Len(txt1(8)) >= 6 Then
'            If Left(txt1(7), 6) <> Left(txt1(8), 6) Then
'                s = MsgBox("申請人代號前六碼須相同", , "USER 輸入錯誤")
'                txt1(8).SetFocus
'                txt1(8).SelStart = 0
'                txt1(8).SelLength = Len(txt1(8))
'                Exit Sub
'             End If
'          End If
'       Else
'           If Len(txt1(7)) < 6 And Len(txt1(8)) <> 0 Then
'              s = MsgBox("申請人代號須填入區間, 6 碼以上", , "USER 輸入錯誤")
'              txt1(7).SetFocus
'              txt1(7).SelStart = 0
'              txt1(7).SelLength = Len(txt1(7))
'              Exit Sub
'           End If
'           If Len(txt1(7)) = 0 And Len(txt1(8)) <> 0 Then
'              s = MsgBox("申請人代號須填入區間, 6 碼以上", , "USER 輸入錯誤")
'              txt1(7).SetFocus
'              txt1(7).SelStart = 0
'              txt1(7).SelLength = Len(txt1(7))
'              Exit Sub
'           End If
'       End If
'       If Len(txt1(9)) <> 0 And Len(txt1(9)) > 6 Then
'         If Len(txt1(10)) = 0 Or Len(txt1(10)) < 6 Then
'             s = MsgBox("代理人代號須填入區間, 6 碼以上", , "USER 輸入錯誤")
'             txt1(10).SetFocus
'             txt1(10).SelStart = 0
'             txt1(10).SelLength = Len(txt1(10))
'             Exit Sub
'         End If
'         If Len(txt1(10)) >= 6 Then
'            If Left(txt1(9), 6) <> Left(txt1(10), 6) Then
'                s = MsgBox("代理人代號前六碼須相同", , "USER 輸入錯誤")
'                txt1(10).SetFocus
'                txt1(10).SelStart = 0
'                txt1(10).SelLength = Len(txt1(10))
'                Exit Sub
'             End If
'          End If
'       Else
'           If Len(txt1(9)) < 6 And Len(txt1(10)) <> 0 Then
'              s = MsgBox("代理人代號須填入區間, 6 碼以上", , "USER 輸入錯誤")
'              txt1(9).SetFocus
'              txt1(9).SelStart = 0
'              txt1(9).SelLength = Len(txt1(9))
'              Exit Sub
'           End If
'           If Len(txt1(9)) = 0 And Len(txt1(10)) <> 0 Then
'              s = MsgBox("代理人代號須填入區間, 6 碼以上", , "USER 輸入錯誤")
'              txt1(9).SetFocus
'              txt1(9).SelStart = 0
'              txt1(9).SelLength = Len(txt1(9))
'              Exit Sub
'           End If
'       End If
'       If Len(txt1(7)) >= 6 Then
'            For i = 1 To 9 - Len(txt1(7))
'                txt1(7) = txt1(7) + "0"
'            Next i
'       End If
'       If Len(txt1(10)) >= 6 Then
'            For i = 1 To 9 - Len(txt1(10))
'                txt1(10) = txt1(10) + "0"
'            Next i
'       End If
'       If Len(txt1(8)) >= 6 Then
'            For i = 1 To 9 - Len(txt1(8))
'                txt1(8) = txt1(8) + "0"
'            Next i
'       End If
'       If Len(txt1(9)) >= 6 Then
'            For i = 1 To 9 - Len(txt1(9))
'                txt1(9) = txt1(9) + "0"
'            Next i
'       End If
'
'    '選擇機關文號
'    If Option1(0).Value = True Then
'        If Len(Trim(txt1(0))) = 0 Then
'            s = MsgBox("機關文號沒有輸入", , "USER 輸入錯誤")
'            txt1(0).SetFocus
'            txt1(0).SelStart = 0
'            txt1(0).SelLength = Len(txt1(0))
'            Exit Sub
'        End If
'        Me.Enabled = False
'        Screen.MousePointer = vbHourglass
'        frm100110_2.Show
'        'frm100110_2.Hide
'        frm100110_2.StrMenu
'        Screen.MousePointer = vbDefault
'        Me.Hide
'        'frm100110_2.Show
'        Do
'        DoEvents
'        If bolToEndByNick = True Then Unload Me: Exit Sub
'        Loop Until Not frm100110_2.Visible
'        Unload frm100110_2
'        Me.Enabled = True
'        Me.Show
'    Else
'         '選擇對造名稱
'        If Option1(1).Value = True Then
'            'Modify By Cheng 2002/07/16
''            If Len(Trim(txtfm2(0))) = 0 And Len(Trim(txt1(2))) = 0 Then
'            If Len(Trim(txtfm2(0))) = 0 And Len(Trim(txt1(2))) = 0 And Len(Trim(txtfm2(1))) = 0 Then
'                 s = MsgBox("對造名稱, 案件名稱及號數皆沒有輸入", , "USER 輸入錯誤")
'                 If Len(Trim(txtfm2(0))) = 0 Then
'                      txtfm2(0).SetFocus
'                      txtfm2(0).SelStart = 0
'                      txtfm2(0).SelLength = Len(txtfm2(0))
'                 'Add By Cheng 2002/07/16
'                 ElseIf Len(Trim(txtfm2(1))) = 0 Then
'                      txtfm2(1).SetFocus
'                      txtfm2(1).SelStart = 0
'                      txtfm2(1).SelLength = Len(txtfm2(1))
'                 Else
'                      txt1(2).SetFocus
'                      txt1(2).SelStart = 0
'                      txt1(2).SelLength = Len(txt1(2))
'                 End If
'                 Exit Sub
'            End If
'            Me.Enabled = False
'            Screen.MousePointer = vbHourglass
'            frm100110_3.Show
'            'frm100110_3.Hide
'
'            frm100110_3.StrMenu
'            Screen.MousePointer = vbDefault
'            Me.Hide
'           ' frm100110_3.Show
'            Do
'            DoEvents
'            If bolToEndByNick = True Then Unload Me: Exit Sub
'            Loop Until Not frm100110_3.Visible
'            Unload frm100110_3
'            Me.Enabled = True
'            Me.Show
'        Else
'            '選擇條款
'            If Option1(2).Value = True Then
'                 If Len(Trim(txt1(3))) = 0 Then
'                     s = MsgBox("條款沒有輸入", , "USER 輸入錯誤")
'                     txt1(3).SetFocus
'                     txt1(3).SelStart = 0
'                     txt1(3).SelLength = Len(txt1(3))
'                     Exit Sub
'                 End If
'                 Me.Enabled = False
'                 Screen.MousePointer = vbHourglass
'                 frm100110_4.Show
'                 'frm100110_4.Hide
'
'                 frm100110_4.StrMenu
'                 Screen.MousePointer = vbDefault
'                 Me.Hide
'                 'frm100110_4.Show
'                 Do
'                 DoEvents
'                 If bolToEndByNick = True Then Unload Me: Exit Sub
'                 Loop Until Not frm100110_4.Visible
'                 Unload frm100110_4
'                 Me.Enabled = True
'                 Me.Show
'            Else
'                If Option1(3).Value = True Then
'                    If Len(Trim(txt1(12))) = 0 Then
'                        s = MsgBox("商品類別沒有輸入", , "USER 輸入錯誤")
'                        txt1(12).SetFocus
'                        txt1(12).SelStart = 0
'                        txt1(12).SelLength = Len(txt1(12))
'                        Exit Sub
'                    End If
'                    Me.Enabled = False
'                    Screen.MousePointer = vbHourglass
'                    frm100110_4.Show
'                    'frm100110_4.Hide
'
'                    frm100110_4.StrMenu2
'                    Screen.MousePointer = vbDefault
'                    Me.Hide
'                    'frm100110_4.Show
'                    Do
'                    DoEvents
'                    If bolToEndByNick = True Then Unload Me: Exit Sub
'                    Loop Until Not frm100110_4.Visible
'                    Unload frm100110_4
'                    Me.Enabled = True
'                    Me.Show
'                End If
'            End If
'        End If
'    End If
'Case 1
'     Unload Me
'Case Else
'End Select
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   bolToEndByNick = False
   Option1(1).Value = False
   Option1(2).Value = False
   
   'txtfm2(0).Enabled = False
   'txt1(2).Enabled = False
   'txt1(3).Enabled = False
   lbl1.Enabled = False
   lbl2.Enabled = False
   
   strTemp = Systemkind_g
   If bolFNation = False Then
       Label6(2).Visible = False
       txt1(9).Visible = False
       txt1(10).Visible = False
       Line1(2).Visible = False
   End If
   '92.04.16 nick
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100110_1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)

   Select Case Index
      Case 0 '以機關文號查詢
         If Option1(0).Value = True Then
            Option1(1).Value = False
            Option1(2).Value = False
            Option1(3).Value = False
            Option1(4).Value = False   '2013/8/19 add by sonia
            lbl1.Enabled = False
            lbl2.Enabled = False
            txt1(0).Enabled = True
            txt1(0).SetFocus
            txtFM2(0).Enabled = False
            txtFM2(1).Enabled = False
            txt1(2).Enabled = False
            txt1(3).Enabled = False
            txt1(12).Enabled = False
            txt1(13).Enabled = False  '2013/8/19 add by sonia
         End If
      Case 1 '以對造名稱查詢
         If Option1(1).Value = True Then
            Option1(0).Value = False
            Option1(2).Value = False
            Option1(3).Value = False
            Option1(4).Value = False   '2013/8/19 add by sonia
            lbl1.Enabled = True
            lbl2.Enabled = True
            txtFM2(0).Enabled = True
            txtFM2(0).SetFocus
            txtFM2(1).Enabled = True
            txt1(2).Enabled = True
            txt1(0).Enabled = False
            txt1(3).Enabled = False
            txt1(12).Enabled = False
            txt1(13).Enabled = False  '2013/8/19 add by sonia
          End If
      Case 2 '以條款查詢
         If Option1(2).Value = True Then
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(3).Value = False
            Option1(4).Value = False   '2013/8/19 add by sonia
            lbl1.Enabled = False
            lbl2.Enabled = False
            txt1(3).Enabled = True
            txt1(3).SetFocus
            txt1(0).Enabled = False
            txtFM2(0).Enabled = False
            txtFM2(1).Enabled = False
            txt1(2).Enabled = False
            txt1(12).Enabled = False
            txt1(13).Enabled = False  '2013/8/19 add by sonia
         End If
      Case 3 '以商品類別查詢
         If Option1(3).Value = True Then
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(2).Value = False
            Option1(4).Value = False   '2013/8/19 add by sonia
            lbl1.Enabled = False
            lbl2.Enabled = False
            txt1(12).Enabled = True
            txt1(12).SetFocus
            txt1(0).Enabled = False
            txtFM2(0).Enabled = False
            txtFM2(1).Enabled = False
            txt1(2).Enabled = False
            txt1(3).Enabled = False
            txt1(13).Enabled = False  '2013/8/19 add by sonia
         End If
      '2013/8/19 add by sonia
      Case 4 '以法院案號查詢
         If Option1(4).Value = True Then
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(2).Value = False
            Option1(3).Value = False
            lbl1.Enabled = False
            lbl2.Enabled = False
            txt1(13).Enabled = True
            txt1(13).SetFocus
            txt1(0).Enabled = False
            txtFM2(0).Enabled = False
            txtFM2(1).Enabled = False
            txt1(2).Enabled = False
            txt1(3).Enabled = False
            txt1(12).Enabled = False
            txt1(13).Text = Val(GetTaiwanThisYear) & "年度行專訴字第號"
            If Mid(GetStaffDepartment(strUserNum), 1, 2) = "P2" Then txt1(13).Text = Val(GetTaiwanThisYear) & "年度行商訴字第號"  'add by sonia 2018/4/19
         End If
      '2013/8/19 end
      Case Else
   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Dim intPos As Integer
      
   'add by sonia 2014/10/29
   Select Case Index
      Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 14, 15
         CloseIme
      Case Else
   End Select
   'end 2014/10/29
   
   '2013/8/19 add by sonia
   If Index = 13 Then
      With txt1(Index)
         If Len("" & .Text) > 0 Then
            intPos = InStr("" & .Text, "第")
            If intPos > 0 Then
               .SelStart = intPos
               .SelLength = 0
            End If
         End If
      End With
   Else
   '2013/8/19 end
      txt1(Index).SelStart = 0
      txt1(Index).SelLength = Len(txt1(Index))
   End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 4, 5
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
         If Index = 5 Then
             If RunNick(txt1(Index - 1), txt1(Index)) Then
                txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
             End If
         End If
      Case 8, 10
          If Mid(txt1(Index - 1), 1, 6) <> Mid(txt1(Index), 1, 6) Then
              s = MsgBox("前6碼必須相同！", , "錯誤！")
              txt1(Index - 1).SetFocus
              txt1_GotFocus (Index - 1)
              Exit Sub
          End If
             If RunNick(txt1(Index - 1), txt1(Index)) Then
                txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Exit Sub
             End If
      Case Else
   End Select
End Sub

Private Sub txt1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

   Select Case Index
      Case 0
          Option1(0).Value = True
      Case 1, 2
          Option1(1).Value = True
      Case 3
          Option1(2).Value = True
      Case Else
   End Select
   
End Sub

'Added by Lydia 2021/01/05
Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
End Sub
'end 2022/01/5
