VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090614 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人工作進度資料查詢"
   ClientHeight    =   3320
   ClientLeft      =   4080
   ClientTop       =   1730
   ClientWidth     =   4300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3320
   ScaleWidth      =   4300
   Begin VB.TextBox Txt1 
      Height          =   270
      Index           =   5
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2130
      Width           =   315
   End
   Begin VB.TextBox Txt1 
      Height          =   270
      Index           =   8
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2490
      Width           =   315
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   7
      Left            =   1815
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1790
      Width           =   480
   End
   Begin VB.TextBox Txt1 
      Height          =   270
      Index           =   6
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1790
      Width           =   480
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2160
      TabIndex        =   9
      Top             =   110
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3030
      TabIndex        =   10
      Top             =   110
      Width           =   1200
   End
   Begin VB.TextBox Txt1 
      Height          =   270
      Index           =   3
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1440
      Width           =   480
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   4
      Left            =   2085
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1430
      Width           =   480
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   0
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   0
      Top             =   720
      Width           =   270
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   1
      Left            =   1605
      MaxLength       =   1
      TabIndex        =   1
      Top             =   720
      Width           =   270
   End
   Begin VB.TextBox Txt1 
      Height          =   270
      Index           =   2
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "(1.螢幕 2.報表)"
      Height          =   180
      Index           =   21
      Left            =   1560
      TabIndex        =   23
      Top             =   2190
      Width           =   1310
   End
   Begin VB.Label Label1 
      Caption         =   "顯示方式："
      Height          =   180
      Index           =   22
      Left            =   210
      TabIndex        =   22
      Top             =   2190
      Width           =   1040
   End
   Begin VB.Label Label1 
      Caption         =   "螢幕顯示："
      Height          =   180
      Index           =   5
      Left            =   200
      TabIndex        =   21
      Top             =   2550
      Width           =   1040
   End
   Begin VB.Label Label1 
      Caption         =   "(N：不區分個人)"
      Height          =   180
      Index           =   3
      Left            =   1560
      TabIndex        =   20
      Top             =   2550
      Width           =   1550
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收文日大於發文日的資料，請用收文年月查詢!!!"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   210
      TabIndex        =   19
      Top             =   3090
      Width           =   3780
   End
   Begin VB.Line Line1 
      X1              =   1470
      X2              =   2055
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "部門別："
      Height          =   180
      Index           =   1
      Left            =   200
      TabIndex        =   18
      Top             =   1830
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "發文年月："
      Height          =   180
      Index           =   0
      Left            =   200
      TabIndex        =   17
      Top             =   1500
      Width           =   1040
   End
   Begin VB.Label Label1 
      Caption         =   "年"
      Height          =   180
      Index           =   2
      Left            =   1740
      TabIndex        =   16
      Top             =   1490
      Width           =   320
   End
   Begin VB.Label Label1 
      Caption         =   "月"
      Height          =   180
      Index           =   7
      Left            =   2700
      TabIndex        =   15
      Top             =   1500
      Width           =   320
   End
   Begin VB.Label Label1 
      Caption         =   "所別："
      Height          =   180
      Index           =   4
      Left            =   200
      TabIndex        =   14
      Top             =   770
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   8
      Left            =   200
      TabIndex        =   13
      Top             =   1130
      Width           =   1100
   End
   Begin VB.Label Label1 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      Height          =   180
      Index           =   24
      Left            =   1910
      TabIndex        =   12
      Top             =   800
      Width           =   2420
   End
   Begin VB.Line Line3 
      X1              =   1290
      X2              =   1740
      Y1              =   870
      Y2              =   870
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   0
      Left            =   2150
      TabIndex        =   11
      Top             =   1110
      Width           =   1790
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm090614"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/07 改成Form2.0 ; lbl1(0) ; Printer列印未改
'Memo By Sindy 2021/12/16 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Public TextOk As Boolean, ManaGrp As String
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 21) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String
Dim PLeft(0 To 21) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, Print1Ok As Boolean, k As Integer
Dim BolChkRs As Boolean
'Add By Cheng 2002/01/09
Dim m_bln_KeyinValid As Boolean
Dim mForm As Form 'Add By Sindy 2023/6/20
Public m_ProState As String 'P,T,ACS,FCP Add By Sindy 2024/2/23


'Modify By Sindy 2013/10/4
'Private Sub cmdOK_Click(Index As Integer)
Public Sub cmdok_Click(Index As Integer)
Dim ii As Integer

Select Case Index
Case 0 '確定
      'Add By Sindy 2025/3/3
      If PUB_ChkFormIsClose("frm090202_2", "電子承辦單簽辦作業") = False Then
         '若承辦人工作進度資料維護-歷程 開著,必須先處理完,才能操作此作業
         '  0.承辦人工作進度 3.繪圖人員工作進度
         If frm090202_2.intReceiveKind = 0 Or frm090202_2.intReceiveKind = 3 Then
            frm090202_2.Show
            Unload Me
            Exit Sub
         End If
      End If
      '2025/3/3 END
      
      'Modify By Sindy 2021/11/24
      '1. P1X人員預設P10~P11，但最大範圍只可為P10~P19且不可空白。
      '2. P2X人員預設P20~P21，但最大範圍只可為P20~P29且不可空白。
      '3. 其他人員不預設，但起迄欄之第一碼必須相同且不可空白。
      'Modify By Sindy 2024/2/23 + 外專承辦人工作進度管理那一支，林總進去要開放跟電腦中心人員的權限，所以請改為Pub_strUserST05為01或08者的權限要打開
      'If Pub_StrUserSt03 <> "M51" Then
      If Pub_StrUserSt03 <> "M51" And Pub_strUserST05 <> "01" And Pub_strUserST05 <> "08" Then
      '2024/2/23 END
         For ii = 6 To 7
            '林文雄A4023以P1X人員身份進入
            'Modify By Sindy 2024/3/5 排除此身份 InStr(Pub_GetSpecMan("協助機械組內專主管"), strUserNum) > 0 And m_ProState = "FCP"
            If (Mid(GetStaffDepartment(strUserNum), 1, 2) = "P1" Or _
               strUserNum = "A4023") And Not (InStr(Pub_GetSpecMan("協助機械組內專主管"), strUserNum) > 0 And m_ProState = "FCP") Then
               'Add By Sindy 2021/11/26 專利處程序會輸入外翻人員操作工作資料維護作業
               If txt1(2) <> "" And Left(txt1(2), 1) = "F" Then
                  If Left(PUB_GetST03(Left(PUB_GetST14(txt1(2)), 5)), 2) = "P1" Then
                     Exit For
                  End If
               End If
               '2021/11/26 END
               'Modify By Sindy 2021/12/3
               'If Left(txt1(ii), 2) <> Mid(GetStaffDepartment(strUserNum), 1, 2) Then
               If Left(txt1(ii), 2) <> "P1" Then
               '2021/12/3 END
                  s = MsgBox("最大範圍只可為P10~P19且不可空白!!", , "USER 輸入錯誤")
                  txt1(ii).SetFocus
                  txt1(ii).SelStart = 0
                  txt1(ii).SelLength = Len(txt1(5))
                  txt1(2).Text = "": lbl1(0).Caption = ""
                  Exit Sub
               End If
            ElseIf Mid(GetStaffDepartment(strUserNum), 1, 2) = "P2" Then
               If Left(txt1(ii), 2) <> "P2" Then
                  s = MsgBox("最大範圍只可為P20~P29且不可空白!!", , "USER 輸入錯誤")
                  txt1(ii).SetFocus
                  txt1(ii).SelStart = 0
                  txt1(ii).SelLength = Len(txt1(5))
                  txt1(2).Text = "": lbl1(0).Caption = ""
                  Exit Sub
               End If
            '江協理98020、林律師98003、沈佳穎96003
            ElseIf strUserNum = "98020" Or strUserNum = "98003" Or strUserNum = "96003" Then
               If Left(txt1(ii), 2) <> "P2" And Left(txt1(ii), 2) <> "F1" Then
                  s = MsgBox("最大範圍只可為F10~F19或P20~P29且不可空白!!", , "USER 輸入錯誤")
                  txt1(ii).SetFocus
                  txt1(ii).SelStart = 0
                  txt1(ii).SelLength = Len(txt1(5))
                  txt1(2).Text = "": lbl1(0).Caption = ""
                  Exit Sub
               End If
            'Add By Sindy 2023/6/21
            ElseIf Mid(GetStaffDepartment(strUserNum), 1, 2) = "F1" Then
               If Left(txt1(ii), 2) <> "F1" Then
                  s = MsgBox("最大範圍只可為F10~F19且不可空白!!", , "USER 輸入錯誤")
                  txt1(ii).SetFocus
                  txt1(ii).SelStart = 0
                  txt1(ii).SelLength = Len(txt1(5))
                  txt1(2).Text = "": lbl1(0).Caption = ""
                  Exit Sub
               End If
               '2023/6/21 END
            'Add By Sindy 2024/3/5 增加此身份 InStr(Pub_GetSpecMan("協助機械組內專主管"), strUserNum) > 0 And m_ProState = "FCP"
            ElseIf InStr(Pub_GetSpecMan("協助機械組內專主管"), strUserNum) > 0 And m_ProState = "FCP" Then
               If Left(txt1(ii), 2) <> "F2" Then
                  s = MsgBox("最大範圍只可為F20~F29且不可空白!!", , "USER 輸入錯誤")
                  txt1(ii).SetFocus
                  txt1(ii).SelStart = 0
                  txt1(ii).SelLength = Len(txt1(5))
                  txt1(2).Text = "": lbl1(0).Caption = ""
                  Exit Sub
               End If
            'Add By Sindy 2023/10/17
            ElseIf Mid(GetStaffDepartment(strUserNum), 1, 2) = "F2" Then
               If GetStaffDepartment(strUserNum) <> txt1(ii) Then
                  s = MsgBox("必須為同部門!!", , "USER 輸入錯誤")
                  txt1(ii).SetFocus
                  txt1(ii).SelStart = 0
                  txt1(ii).SelLength = Len(txt1(5))
                  txt1(2).Text = "": lbl1(0).Caption = ""
                  Exit Sub
               End If
               '組別要相同
               If GetStaffDepartment(strUserNum) = "F21" Then
                  If PUB_GetST05(strUserNum) <> "42" Then '42=外專工程師高級主管
                     If txt1(2) = "" Then
                        s = MsgBox("承辦人不可空白!!", , "USER 輸入錯誤")
                        txt1(2).SetFocus
                        Exit Sub
'                     '是否有權限
'                     ElseIf PUB_GetST52(Txt1(2), strUserNum) = False Then
'                        s = MsgBox("無查詢此承辦人的權限!!", , "USER 輸入錯誤")
'                        Txt1(2).SetFocus
'                        Exit Sub
                     End If
                  End If
                  'Modify By Sindy 2024/4/23 增加不同組的各級主管，可以單獨下員工編號查詢
                  If txt1(2) <> "" Then
                     '是否有權限
                     If PUB_GetST52(txt1(2), strUserNum) = False Then
                        s = MsgBox("無查詢此承辦人的權限!!", , "USER 輸入錯誤")
                        txt1(2).SetFocus
                        Exit Sub
                     End If
                  End If
'                  If Txt1(2) <> "" And PUB_GetStaffST16(Txt1(2)) <> PUB_GetStaffST16(strUserNum) Then
'                     s = MsgBox("必須為同組別!!", , "USER 輸入錯誤")
'                     Txt1(2).SetFocus
'                     Txt1(ii).SelStart = 0
'                     Txt1(ii).SelLength = Len(Txt1(5))
'                     Txt1(2).Text = "": lbl1(0).Caption = ""
'                     Exit Sub
'                  End If
                  '2024/4/23 END
               End If
               '2023/10/17 END
            Else
               If Left(txt1(ii), 1) <> Mid(GetStaffDepartment(strUserNum), 1, 1) Then
                  s = MsgBox("起迄欄之第一碼必須相同且不可空白!!", , "USER 輸入錯誤")
                  txt1(ii).SetFocus
                  txt1(ii).SelStart = 0
                  txt1(ii).SelLength = Len(txt1(5))
                  txt1(2).Text = "": lbl1(0).Caption = ""
                  Exit Sub
               End If
            End If
         Next ii
      End If
      
     txt1_Validate 2, False
     If m_bln_KeyinValid = False Then Exit Sub
     If Len(txt1(3)) = 0 Then
         s = MsgBox("發文年不可空白!!", , "USER 輸入錯誤")
         txt1(3).SetFocus
         Exit Sub
     Else
         If Len(txt1(4)) = 0 Then
             s = MsgBox("發文月不可空白!!", , "USER 輸入錯誤")
             txt1(4).SetFocus
             Exit Sub
         Else
             If Len(txt1(5)) = 0 Then
                 s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
                 txt1(5).SetFocus
                 Exit Sub
             Else
                 'Modify By Sindy 2012/5/15
'                 If Len(txt1(7)) = 0 And Len(txt1(2)) = 0 Then
'                     s = MsgBox("部門別或承辦人最少要輸入一個!!", , "USER 輸入錯誤")
'                     txt1(2).SetFocus
                 If Len(txt1(6)) = 0 Or Len(txt1(7)) = 0 Then
                     s = MsgBox("部門別不可空白!!", , "USER 輸入錯誤")
                     If Len(txt1(6)) = 0 Then txt1(6).SetFocus: Exit Sub
                     If Len(txt1(7)) = 0 Then txt1(7).SetFocus: Exit Sub
                 '2012/5/15 End
                 Else
                    'Add By Cheng 2003/06/03
                    If Me.txt1(6).Text <> "" And Me.txt1(7).Text <> "" Then
                        If Me.txt1(6).Text > Me.txt1(7).Text Then
                            MsgBox "部門別區間範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                            Me.txt1(6).SetFocus
                            txt1_GotFocus 6
                            Exit Sub
                        End If
                    End If
                     ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/17 清除查詢印表記錄檔欄位
                     '查詢
                     If Trim(txt1(5)) = "1" Then
                        pub_QL05 = pub_QL05 & ";" & Label1(22) & "1.螢幕" 'Add By Sindy 2010/12/17
                        Screen.MousePointer = vbHourglass
                        Me.Enabled = False
                        'Modify By Sindy 2012/5/15 專利及非專利拆開作業
                        'Modified by Morgan 2012/6/25 +F2部門
                        If (Left(Trim(Me.txt1(6).Text), 2) = "P1" Or Left(Trim(Me.txt1(6).Text), 2) = "F5") And _
                           (InStr(UCase(App.EXEName), "PROMOTER") > 0 Or _
                            (InStr(UCase(App.EXEName), "PATPRO") > 0 And InStr(UCase(App.EXEName), "PATPRO1") = 0)) Then
                           '專利處
'                           frm090201_2.Show
'                           If TextOk = False Then
'                              Unload frm090201_2
'                              Me.Show
'                           Else
'                              Me.Hide
'                           End If
                           'Modify By Sindy 2023/6/20
                           Set mForm = Forms(0).GetForm("frm090201_2")
                           mForm.Show
                           If TextOk = False Then
                              Unload mForm
                              Me.Show
                           Else
                              Me.Hide
                           End If
                           '2023/6/20 END
                        'Add By Sindy 2023/6/20
                        ElseIf Left(Trim(Me.txt1(6).Text), 2) = "F2" And _
                           (InStr(UCase(App.EXEName), "PROMOTER") > 0 Or InStr(UCase(App.EXEName), "PATPRO1") > 0) Then
                           '外專
                           Set mForm = Forms(0).GetForm("frm090909")
                           mForm.Show
                           If TextOk = False Then
                              Unload mForm
                              Me.Show
                           Else
                              Me.Hide
                           End If
                           '2023/6/20 END
                        'Modify By Sindy 2021/9/24 商標部
                        ElseIf Left(Trim(Me.txt1(6).Text), 2) = "P2" Or Left(Trim(Me.txt1(6).Text), 2) = "F1" Then
                           '內商
'                           frm090201_b.Show
'                           If TextOk = False Then
'                              Unload frm090201_b
'                              Me.Show
'                           Else
'                              Me.Hide
'                           End If
                           'Modify By Sindy 2023/6/20
                           Set mForm = Forms(0).GetForm("frm090201_b")
                           mForm.Show
                           If TextOk = False Then
                              Unload mForm
                              Me.Show
                           Else
                              Me.Hide
                           End If
                           '2023/6/20 END
                        ElseIf Left(Trim(Me.txt1(6).Text), 1) = "W" And _
                           (InStr(UCase(App.EXEName), "LAW") > 0 Or InStr(UCase(App.EXEName), "PROMOTER") > 0) Then
                           '非專利,商標
'                           frm090201_d.Show
'                           If TextOk = False Then
'                              Unload frm090201_d
'                              Me.Show
'                           Else
'                              Me.Hide
'                           End If
                           'Modify By Sindy 2023/6/20
                           Set mForm = Forms(0).GetForm("frm090201_d")
                           mForm.Show
                           If TextOk = False Then
                              Unload mForm
                              Me.Show
                           Else
                              Me.Hide
                           End If
                           '2023/6/20 END
                        '2021/9/24 END
                        Else
                           MsgBox "此系統無您( " & IIf(txt1(2).Text = "", Left(Trim(Me.txt1(6).Text), 2) & "單位", txt1(2).Text) & " )的「工作進度資料維護」可使用！"
                        End If
                        Me.Enabled = True
                        Screen.MousePointer = vbDefault
                     '印表
                     Else
                        pub_QL05 = pub_QL05 & ";" & Label1(22) & "2.報表" 'Add By Sindy 2010/12/17
                        TextOk = True
                        Screen.MousePointer = vbHourglass
                        Me.Enabled = False
                        Process
                        If TextOk = False Then
                           InsertQueryLog (0) 'Add By Sindy 2010/12/17
                           ShowNoData
                           'Added by Lydia 2016/12/16
                           Screen.MousePointer = vbDefault
                           Me.Enabled = True
                           'end 2016/12/16
                           Exit Sub
                        End If
                        PrintData
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

'專利處
Sub Process1()
Set mForm = Forms(0).GetForm("frm090201_2") 'Add By Sindy 2023/6/20
strSQL2 = ""
If Len(txt1(0)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST06>='" & txt1(0) & "' "
End If
If Len(txt1(1)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST06<='" & txt1(1) & "' "
End If
If Len(txt1(0)) <> 0 Or Len(txt1(1)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(0) & "-" & txt1(1) & Label1(24) 'Add By Sindy 2010/12/17
End If
If Len(txt1(2)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST01='" & txt1(2) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(2) & lbl1(0) 'Add By Sindy 2010/12/17
End If
If Len(Trim(txt1(6))) <> 0 Then
   strSQL2 = strSQL2 & " AND S1.ST03>='" & txt1(6) & "' "
End If
If Len(Trim(txt1(7))) <> 0 Then
   strSQL2 = strSQL2 & " AND S1.ST03<='" & txt1(7) & "' "
End If
If Len(Trim(txt1(6))) <> 0 Or Len(Trim(txt1(7))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/17
End If
pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(3) & txt1(4) 'Add By Sindy 2010/12/17
If txt1(8) = "N" Then
   pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(8) & "(N：不區分個人)" 'Add By Sindy 2010/12/17
End If
ManaGrp = ""
CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open "select distinct sg02 from staff_group ,staff where st01='" & strUserNum & "' and st11=sg01(+) ", cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
       .MoveFirst
       Do While .EOF = False
            ManaGrp = ManaGrp & CheckStr(.Fields(0))
            .MoveNext
            If .EOF = False Then
               ManaGrp = ManaGrp & ","
            End If
      Loop
   End If
End With
If Len(strSQL2) <> 0 Then
   CheckOC2
   With adoRecordset1
      .CursorLocation = adUseClient
      ' 91.08.19 邱小姐說不用鎖部門別
      '.Open "select s1.st01 from staff s1 where s1.st04='1' AND S1.ST03<>'P13' " & strsql2 & " order by 1 ", cnnConnection, adOpenStatic, adLockReadOnly
      '2012/8/3 modify by sonia 游經理說下承辦人時不考慮是否在職(93001賴健桓留職停薪)
      '.Open "select s1.st01||' ('||s1.st02||')' from staff s1 where s1.st04='1' " & strSQL2 & " order by 1 ", cnnConnection, adOpenStatic, adLockReadOnly
      .Open "select s1.st01||' ('||s1.st02||')' from staff s1 where " & IIf(Len(txt1(2)) <> 0, "s1.st04=s1.st04 ", "s1.st04='1' ") & strSQL2 & " order by 1 ", cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         .MoveFirst
         mForm.Combo1.Clear
         mForm.Combo1_String = ""              '92.6.26 ADD BY SONIA
         s = 0
         Do While .EOF = False
            mForm.Combo1.AddItem CheckStr(.Fields(0)), s
            s = s + 1
            '92.6.26 ADD BY SONIA
            If mForm.Combo1_String = "" Then
               mForm.Combo1_String = "'" & Trim(Left("" & .Fields(0), 6)) & "'"
            Else
               mForm.Combo1_String = mForm.Combo1_String + ",'" & Trim(Left("" & .Fields(0), 6)) & "'"
            End If
            '92.6.26 END
            .MoveNext
         Loop
         mForm.Combo1.Text = mForm.Combo1.List(0)
         TextOk = True
      Else
         TextOk = False
         ShowNoData
      End If
   End With
End If
End Sub

'Add By Sindy 2012/5/15 內商
Sub Process2()
Set mForm = Forms(0).GetForm("frm090201_b") 'Add By Sindy 2023/6/20
strSQL2 = ""
If Len(txt1(0)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST06>='" & txt1(0) & "' "
End If
If Len(txt1(1)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST06<='" & txt1(1) & "' "
End If
If Len(txt1(0)) <> 0 Or Len(txt1(1)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(0) & "-" & txt1(1) & Label1(24) 'Add By Sindy 2010/12/17
End If
If Len(txt1(2)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST01='" & txt1(2) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(2) & lbl1(0) 'Add By Sindy 2010/12/17
End If
If Len(Trim(txt1(6))) <> 0 Then
   strSQL2 = strSQL2 & " AND S1.ST03>='" & txt1(6) & "' "
End If
If Len(Trim(txt1(7))) <> 0 Then
   strSQL2 = strSQL2 & " AND S1.ST03<='" & txt1(7) & "' "
End If
If Len(Trim(txt1(6))) <> 0 Or Len(Trim(txt1(7))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/17
End If
pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(3) & txt1(4) 'Add By Sindy 2010/12/17
If txt1(8) = "N" Then
   pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(8) & "(N：不區分個人)" 'Add By Sindy 2010/12/17
End If
ManaGrp = ""
CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open "select distinct sg02 from staff_group ,staff where st01='" & strUserNum & "' and st11=sg01(+) ", cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
       .MoveFirst
       Do While .EOF = False
            ManaGrp = ManaGrp & CheckStr(.Fields(0))
            .MoveNext
            If .EOF = False Then
               ManaGrp = ManaGrp & ","
            End If
      Loop
   End If
End With
If Len(strSQL2) <> 0 Then
   CheckOC2
   With adoRecordset1
      .CursorLocation = adUseClient
      ' 91.08.19 邱小姐說不用鎖部門別
      '.Open "select s1.st01 from staff s1 where s1.st04='1' AND S1.ST03<>'P13' " & strsql2 & " order by 1 ", cnnConnection, adOpenStatic, adLockReadOnly
      '2012/8/3 modify by sonia 游經理說下承辦人時不考慮是否在職(93001賴健桓留職停薪)
      '.Open "select s1.st01||' ('||s1.st02||')' from staff s1 where s1.st04='1' " & strSQL2 & " order by 1 ", cnnConnection, adOpenStatic, adLockReadOnly
      'Modify By Sindy 2024/4/24 + and substr(s1.st01,1,1)<>'F'
      If Left(GetStaffDepartment(strUserNum), 2) = "F1" And Len(txt1(2)) = 0 Then
         strSQL2 = strSQL2 & " AND S1.ST16='" & PUB_GetStaffST16(strUserNum) & "' "
      End If
      .Open "select s1.st01||' ('||s1.st02||')' from staff s1 where " & IIf(Len(txt1(2)) <> 0, "s1.st04=s1.st04 ", "s1.st04='1' ") & strSQL2 & " and substr(s1.st01,1,1)<>'F' order by 1 ", cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         .MoveFirst
         mForm.Combo1.Clear
         mForm.Combo1_String = ""              '92.6.26 ADD BY SONIA
         s = 0
         Do While .EOF = False
            mForm.Combo1.AddItem CheckStr(.Fields(0)), s
            s = s + 1
            '92.6.26 ADD BY SONIA
            If mForm.Combo1_String = "" Then
               mForm.Combo1_String = "'" & Trim(Left("" & .Fields(0), 6)) & "'"
            Else
               mForm.Combo1_String = mForm.Combo1_String + ",'" & Trim(Left("" & .Fields(0), 6)) & "'"
            End If
            '92.6.26 END
            .MoveNext
         Loop
         mForm.Combo1.Text = mForm.Combo1.List(0)
         TextOk = True
      Else
         TextOk = False
         ShowNoData
      End If
   End With
End If
End Sub

'Add By Sindy 2021/9/24 法務,顧問作業
Sub Process3()
Set mForm = Forms(0).GetForm("frm090201_d") 'Add By Sindy 2023/6/20
strSQL2 = ""
If Len(txt1(0)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST06>='" & txt1(0) & "' "
End If
If Len(txt1(1)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST06<='" & txt1(1) & "' "
End If
If Len(txt1(0)) <> 0 Or Len(txt1(1)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(0) & "-" & txt1(1) & Label1(24)
End If
If Len(txt1(2)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST01='" & txt1(2) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(2) & lbl1(0)
End If
If Len(Trim(txt1(6))) <> 0 Then
   strSQL2 = strSQL2 & " AND S1.ST03>='" & txt1(6) & "' "
End If
If Len(Trim(txt1(7))) <> 0 Then
   strSQL2 = strSQL2 & " AND S1.ST03<='" & txt1(7) & "' "
End If
If Len(Trim(txt1(6))) <> 0 Or Len(Trim(txt1(7))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(6) & "-" & txt1(7)
End If
pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(3) & txt1(4)
If txt1(8) = "N" Then
   pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(8) & "(N：不區分個人)"
End If
ManaGrp = ""
CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open "select distinct sg02 from staff_group ,staff where st01='" & strUserNum & "' and st11=sg01(+) ", cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
       .MoveFirst
       Do While .EOF = False
            ManaGrp = ManaGrp & CheckStr(.Fields(0))
            .MoveNext
            If .EOF = False Then
               ManaGrp = ManaGrp & ","
            End If
      Loop
   End If
End With
If Len(strSQL2) <> 0 Then
   CheckOC2
   With adoRecordset1
      .CursorLocation = adUseClient
      ' 91.08.19 邱小姐說不用鎖部門別
      '.Open "select s1.st01 from staff s1 where s1.st04='1' AND S1.ST03<>'P13' " & strsql2 & " order by 1 ", cnnConnection, adOpenStatic, adLockReadOnly
      '2012/8/3 modify by sonia 游經理說下承辦人時不考慮是否在職(93001賴健桓留職停薪)
      '.Open "select s1.st01||' ('||s1.st02||')' from staff s1 where s1.st04='1' " & strSQL2 & " order by 1 ", cnnConnection, adOpenStatic, adLockReadOnly
      'Modify By Sindy 2024/4/24 + and substr(s1.st01,1,1)<>'F'
      .Open "select s1.st01||' ('||s1.st02||')' from staff s1 where " & IIf(Len(txt1(2)) <> 0, "s1.st04=s1.st04 ", "s1.st04='1' ") & strSQL2 & " and substr(s1.st01,1,1)<>'F' order by 1 ", cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         .MoveFirst
         mForm.Combo1.Clear
         mForm.Combo1_String = ""              '92.6.26 ADD BY SONIA
         s = 0
         Do While .EOF = False
            mForm.Combo1.AddItem CheckStr(.Fields(0)), s
            s = s + 1
            '92.6.26 ADD BY SONIA
            If mForm.Combo1_String = "" Then
               mForm.Combo1_String = "'" & Trim(Left("" & .Fields(0), 6)) & "'"
            Else
               mForm.Combo1_String = mForm.Combo1_String + ",'" & Trim(Left("" & .Fields(0), 6)) & "'"
            End If
            '92.6.26 END
            .MoveNext
         Loop
         mForm.Combo1.Text = mForm.Combo1.List(0)
         TextOk = True
      Else
         TextOk = False
         ShowNoData
      End If
   End With
End If
End Sub

'Add By Sindy 2023/10/17
'外專
Sub Process4()
Set mForm = Forms(0).GetForm("frm090909")
strSQL2 = ""
If Len(txt1(0)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST06>='" & txt1(0) & "' "
End If
If Len(txt1(1)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST06<='" & txt1(1) & "' "
End If
If Len(txt1(0)) <> 0 Or Len(txt1(1)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(0) & "-" & txt1(1) & Label1(24) 'Add By Sindy 2010/12/17
End If
If Len(txt1(2)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST01='" & txt1(2) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(2) & lbl1(0) 'Add By Sindy 2010/12/17
End If
If Len(Trim(txt1(6))) <> 0 Then
   strSQL2 = strSQL2 & " AND S1.ST03>='" & txt1(6) & "' "
End If
If Len(Trim(txt1(7))) <> 0 Then
   strSQL2 = strSQL2 & " AND S1.ST03<='" & txt1(7) & "' "
End If
If Len(Trim(txt1(6))) <> 0 Or Len(Trim(txt1(7))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/17
End If

'Modify By Sindy 2024/3/5
If InStr(Pub_GetSpecMan("協助機械組內專主管"), strUserNum) > 0 And m_ProState = "FCP" Then
   strSQL2 = strSQL2 & " AND substr(S1.ST01,4,1)='9' AND S1.ST01<>'94099' "
   pub_QL05 = pub_QL05 & ";協助機械組內專主管=" & strUserNum
Else
'2024/3/5 END
   If GetStaffDepartment(strUserNum) = "F21" And Len(txt1(2)) = 0 Then
      strSQL2 = strSQL2 & " AND S1.ST16='" & PUB_GetStaffST16(strUserNum) & "' "
   End If
End If

pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(3) & txt1(4) 'Add By Sindy 2010/12/17
If txt1(8) = "N" Then
   pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(8) & "(N：不區分個人)" 'Add By Sindy 2010/12/17
End If
ManaGrp = "P,PS,CFP,CPS,"
CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open "select distinct sg02 from staff_group ,staff where st01='" & strUserNum & "' and st11=sg01(+) ", cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
       .MoveFirst
       Do While .EOF = False
            ManaGrp = ManaGrp & CheckStr(.Fields(0))
            .MoveNext
            If .EOF = False Then
               ManaGrp = ManaGrp & ","
            End If
      Loop
   End If
End With
If ManaGrp <> "" And Right(ManaGrp, 1) = "," Then ManaGrp = Left(ManaGrp, Len(ManaGrp) - 1)
If Len(strSQL2) <> 0 Then
   CheckOC2
   With adoRecordset1
      .CursorLocation = adUseClient
      ' 91.08.19 邱小姐說不用鎖部門別
      '.Open "select s1.st01 from staff s1 where s1.st04='1' AND S1.ST03<>'P13' " & strsql2 & " order by 1 ", cnnConnection, adOpenStatic, adLockReadOnly
      '2012/8/3 modify by sonia 游經理說下承辦人時不考慮是否在職(93001賴健桓留職停薪)
      '.Open "select s1.st01||' ('||s1.st02||')' from staff s1 where s1.st04='1' " & strSQL2 & " order by 1 ", cnnConnection, adOpenStatic, adLockReadOnly
      'Modify By Sindy 2024/4/24 + and substr(s1.st01,1,1)<>'F'
      .Open "select s1.st01||' ('||s1.st02||')',s1.st02 from staff s1 where " & IIf(Len(txt1(2)) <> 0, "s1.st04=s1.st04 ", "s1.st04='1' ") & strSQL2 & " and substr(s1.st01,1,1)<>'F' order by 1 ", cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         .MoveFirst
         mForm.Combo1.Clear
         mForm.Combo1_String = "" '92.6.26 ADD BY SONIA
         mForm.Combo1_Name = "" 'Add By Sindy 2024/3/5
         s = 0
         Do While .EOF = False
            mForm.Combo1.AddItem CheckStr(.Fields(0)), s
            s = s + 1
            '92.6.26 ADD BY SONIA
            If mForm.Combo1_String = "" Then
               mForm.Combo1_String = "'" & Trim(Left("" & .Fields(0), 6)) & "'"
               mForm.Combo1_Name = Trim(.Fields(1)) 'Add By Sindy 2024/3/5
            Else
               mForm.Combo1_String = mForm.Combo1_String + ",'" & Trim(Left("" & .Fields(0), 6)) & "'"
               mForm.Combo1_Name = mForm.Combo1_Name + "," & Trim(.Fields(1)) 'Add By Sindy 2024/3/5
            End If
            '92.6.26 END
            .MoveNext
         Loop
         If mForm.Combo1.ListIndex >= 0 Then mForm.Combo1.Text = mForm.Combo1.List(0)
         TextOk = True
      Else
         TextOk = False
         ShowNoData
      End If
   End With
End If
End Sub

Sub Process()
'Modify By Cheng 2003/05/08
'cnnConnection.Execute "DELETE FROM R090614 WHERE ID='" & strUserNum & "' "
adoEng.Execute "DELETE FROM R090614 WHERE ID='" & strUserNum & "' "
    'add by nickc 2007/12/17
    adoEng.Execute "drop table R090614 "
    adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text,R110006 text,R110007 text,R110008 text,R110009 text,R110010 text,R110011 text,R110012 text,R110013 text,R110014 text,R110015 text,R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo,R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text)"

    'adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text,R110006 text,R110007 text,R110008 text,R110009 text,R110010 text,R110011 text,R110012 text,R110013 text,R110014 text,R110015 text,R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo,R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text,R110026 double,R110027 double,R110028 double,R110029 text,R110030 text)"

StrSQL6 = ""
strSQL1 = ""
strSQL2 = ""
BolChkRs = False    '檢查是否曾經讀過資料庫
If Len(txt1(0)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST06>='" & txt1(0) & "' "
End If
If Len(txt1(1)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST06<='" & txt1(1) & "' "
End If
If Len(txt1(0)) <> 0 Or Len(txt1(1)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(0) & "-" & txt1(1) & Label1(24) 'Add By Sindy 2010/12/17
End If
If Len(txt1(2)) <> 0 Then
    strSQL2 = strSQL2 + " AND S1.ST01='" & txt1(2) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(2) & lbl1(0) 'Add By Sindy 2010/12/17
End If
If Len(Trim(txt1(6))) <> 0 Then
   strSQL2 = strSQL2 & " AND S1.ST03>='" & txt1(6) & "' "
End If
If Len(Trim(txt1(7))) <> 0 Then
   strSQL2 = strSQL2 & " AND S1.ST03<='" & txt1(7) & "' "
End If
If Len(Trim(txt1(6))) <> 0 Or Len(Trim(txt1(7))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/17
End If
pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(3) & txt1(4) 'Add By Sindy 2010/12/17
ManaGrp = ""
CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open "select distinct sg02 from staff_group ,staff where st01='" & strUserNum & "' and st11=sg01(+) ", cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
       .MoveFirst
       Do While .EOF = False
            ManaGrp = ManaGrp & CheckStr(.Fields(0))
            .MoveNext
            If .EOF = False Then
               ManaGrp = ManaGrp & ","
            End If
      Loop
   End If
End With
If Len(strSQL2) <> 0 Then
   CheckOC2
   With adoRecordset1
      .CursorLocation = adUseClient
        'Modify By Cheng 2003/05/08
'      .Open "select s1.st01 from staff s1 where st04='1' " & strSQL2, cnnConnection, adOpenStatic, adLockReadOnly
      '2012/8/3 modify by sonia 游經理說下承辦人時不考慮是否在職(93001賴健桓留職停薪)
      '.Open "select s1.st01||' ('||S1.ST02||')' from staff s1 where st04='1' " & strSQL2, cnnConnection, adOpenStatic, adLockReadOnly
      .Open "select s1.st01||' ('||S1.ST02||')' from staff s1 where " & IIf(Len(txt1(2)) <> 0, "s1.st04=s1.st04 ", "s1.st04='1' ") & strSQL2, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         .MoveFirst
         Do While .EOF = False
            'StrSQL6 = StrSQL6 + " and CP14='" & Trim(Combo1.Text) & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null) or (CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp27 is null))) and cp05>=19980101"
            'StrSQL1 = " and s1.st01='" & CheckStr(.Fields(0)) & "' AND ((SUBSTR(CP27,1,6)=" & Trim(str(Val(Trim(txt1(3)) & Trim(Right(ChgNumByNick(txt1(4)), 2)) & "01 & " or SUBSTR(CP57,1,6)=" & Trim(str(Val(Trim(txt1(3)) & Trim(Right("0" & Trim(txt1(4)), 2))) + 191100)) & ") and cp05>=19980101"
            'StrSQL6 = " and s1.st01='" & CheckStr(.Fields(0)) & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim(str(Val(Trim(Txt1(3) + 1911) & Trim(Right("0" & Trim(Txt1(4)), 2))) + 191100)) & "01 and cp27<=" & Trim((Val(Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(Txt1(4)), 2)) & "31 and cp57 is null) or (CP57>=" & Trim((Val(Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(Txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(Txt1(4)), 2)) & "31 and cp27 is null))) and cp05>=19980101 "
            
            'StrGrp090201 = frm090614.ManaGrp
            StrSQL6 = ""
            strSQL1 = ""
            strSQL2 = ""
            'Add By Cheng 2002/04/22
            '為了要與承辦人管理的查詢條件一致
            StrSQL6 = " and cp05<=" & Trim((Val(Me.txt1(3).Text) + 1911)) & Trim(Right(ChgNumByNick(Me.txt1(4).Text), 2)) & "31 "
            'Modify By Cheng 2003/05/08
'            StrSQL6 = StrSQL6 + " and CP14='" & CheckStr(.Fields(0)) & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null) or (CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp27 is null))) and cp05>=19980101"
'            strSQL1 = strSQL1 & " and CP14='" & CheckStr(.Fields(0)) & "' AND ((CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or (CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 )) and cp05>=19980101"
            'Modify By Cheng 2003/07/18
            '不限制發文日止日及取消收文日止日
'            StrSQL6 = StrSQL6 + " and CP14='" & Trim(Left("" & .Fields(0), 6)) & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 and cp57 is null) or (CP57>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 and cp27 is null))) and cp05>=19980101"
'            strSQL1 = strSQL1 & " and CP14='" & Trim(Left("" & .Fields(0), 6)) & "' AND ((CP27>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31) or (CP57>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 )) and cp05>=19980101"
            'edit by nickc 2005/05/13
            'StrSQL6 = StrSQL6 + " and CP14='" & Trim(Left("" & .Fields(0), 6)) & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp57 is null) or (CP57>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp27 is null))) and cp05>=19980101"

'Modify By Sindy 2024/7/30 CP14 改判斷 EP05
            StrSQL6 = StrSQL6 + " and EP05='" & Trim(Left("" & .Fields(0), 6)) & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 ) or (CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27 is null))) and cp05>=19980101"
            strSQL1 = strSQL1 & " and EP05='" & Trim(Left("" & .Fields(0), 6)) & "' AND ((CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 ) or (CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 )) and cp05>=19980101"
            
            'Modified by Lydia 2016/12/16 + 閉卷符號
'            strSql = "SELECT CP14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 1) & ") "
'            strSql = strSql + " UNION all  SELECT CP14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 2) & ") "
'            strSql = strSql + " UNION all  SELECT CP14,ep01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 3) & ") "
'            strSql = strSql + " UNION all  SELECT CP14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,CP26,'',nvl(CPM03,cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 4) & ") "
'            strSql = strSql + " UNION all  SELECT CP14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 5) & ") "
            
'Modify By Sindy 2024/7/30 CP14 改判斷 EP05
            strSql = "SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 1) & ") "
            strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 2) & ") "
            strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 3) & ") "
            strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 4) & ") "
            strSql = strSql + " UNION all  SELECT EP05,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 5) & ") "
            strSql = strSql + " ORDER BY 1,4 "
         CheckOC
         Print1Ok = False
          adoRecordset.CursorLocation = adUseClient
          'Debug.Print Timer
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
          'Debug.Print Timer
          If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              adoRecordset.MoveFirst
              k = 0
              DoEvents
              '判斷等級是否屬於專利
              'modify by sonia 2014/4/29 加cp14=94007
              
'Modify By Sindy 2024/7/30 CP14 改判斷 EP05
              If (Val(CheckStr(adoRecordset.Fields(22))) >= 31 And Val(CheckStr(adoRecordset.Fields(22))) <= 39) Or (Val(CheckStr(adoRecordset.Fields(22))) >= 71 And Val(CheckStr(adoRecordset.Fields(22))) <= 89) Or CheckStr(adoRecordset.Fields("EP05")) = "94007" Then
                  Print1Ok = True
              End If
              Do While adoRecordset.EOF = False
                  If BolChkRs = False Then
                     BolChkRs = True
                  End If
                  For i = 0 To 21
                      strTemp(i) = CheckStr(adoRecordset.Fields(i))
                  Next i
                  '計算承辦天數
                  If Len(strTemp(14)) <> 0 And Len(strTemp(12)) <> 0 Then
                       strTemp(18) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(14))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(12))))))
                  Else
                      If Len(strTemp(13)) <> 0 And Len(strTemp(12)) <> 0 Then
                          strTemp(18) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(13))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(12))))))
                      End If
                  End If
                  'Modify By Cheng 2002/04/18
'                  strSQL = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "') "
                  strSql = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','','','') "
                    'Modify By Cheng 2003/05/08
'                  cnnConnection.Execute strSQL
                    adoEng.Execute strSql
                  adoRecordset.MoveNext
                  DoEvents
              Loop
          End If
          CheckOC
         If BolChkRs = True Then
            TextOk = True
         Else
            TextOk = False
         End If
'            strSQL = " SELECT S1.ST01,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND S1.ST01=EP05(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 1) & ") "
'            strSQL = strSQL + " UNION all  SELECT S1.ST01,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND S1.ST01=EP05(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 2) & ") "
'            strSQL = strSQL + " UNION all  SELECT S1.ST01,ep01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND S1.ST01=EP05(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 3) & ") "
'            strSQL = strSQL + " UNION all  SELECT S1.ST01,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,CP26,'',nvl(CPM03,cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND S1.ST01=EP05(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 4) & ") "
'            strSQL = strSQL + " UNION all  SELECT S1.ST01,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND S1.ST01=EP05(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 5) & ") "
            'strSQL = strSQL + " UNION all  SELECT S1.ST01,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND S1.ST01=EP05(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL1 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 1) & ") "
            'strSQL = strSQL + " UNION all  SELECT S1.ST01,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND S1.ST01=EP05(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) " & StrSQL1 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 2) & ") "
            'strSQL = strSQL + " UNION all  SELECT S1.ST01,ep01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND S1.ST01=EP05(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL1 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 3) & ") "
            'strSQL = strSQL + " UNION all  SELECT S1.ST01,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,HC06,CP26,'',nvl(CPM03,cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND S1.ST01=EP05(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL1 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 4) & ") "
            'strSQL = strSQL + " UNION all  SELECT S1.ST01,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,nvl(S2.ST02,cp13),CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND S1.ST01=EP05(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL1 & " AND CP01 IN (" & SQLGrpStr(ManaGrp, 5) & ") "
'            strSQL = strSQL + " ORDER BY 1,4 "
'            CheckOC
'            adoRecordset.CursorLocation = adUseClient
'            adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'            If adoRecordset.RecordCount <> 0 Then
'                  adoRecordset.MoveFirst
'                  DoEvents
'                  '判斷等級是否屬於專利
'                  If (Val(CheckStr(adoRecordset.Fields(22))) >= 31 And Val(CheckStr(adoRecordset.Fields(22))) <= 39) Or (Val(CheckStr(adoRecordset.Fields(22))) >= 71 And Val(CheckStr(adoRecordset.Fields(22))) <= 89) Then
 '                     Print1Ok = True
'                  End If
'                  Do While adoRecordset.EOF = False
'                      For i = 0 To 21
'                          strTemp(i) = CheckStr(adoRecordset.Fields(i))
 '                     Next i
'                      '計算承辦天數
'                      If Len(strTemp(14)) <> 0 And Len(strTemp(12)) <> 0 Then
'                           strTemp(18) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(14))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(12))))))
'                      Else
'                          If Len(strTemp(13)) <> 0 And Len(strTemp(12)) <> 0 Then
'                              strTemp(18) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(13))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(12))))))
'                          End If
'                      End If
'                      strSQL = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "') "
'                      cnnConnection.Execute strSQL
'                      adoRecordset.MoveNext
'                      DoEvents
'                  Loop
'            End If
            .MoveNext
         Loop
'
      End If
   End With
End If
    CALCUTE_090201 IIf(Len(txt1(2)) <> 0, txt1(2), ""), Trim(str(Val(Trim(txt1(3)) & Trim(Right("0" & Trim(txt1(4)), 2))) + 191100))
End Sub

Sub PrintData()
If Print1Ok = True Then
    PrintData2   '專利
Else
    PrintData1   '一般
End If
End Sub

Sub PrintData1()
strSql = "SELECT DISTINCT R110001 FROM R090614 WHERE ID='" & strUserNum & "' "
CheckOC2
Page = 1
adoRecordset1.CursorLocation = adUseClient
'Modify By Cheng 2003/05/08
'adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset1.Open strSql, adoEng, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
   InsertQueryLog (adoRecordset1.RecordCount) 'Add By Sindy 2010/12/17
   adoRecordset1.MoveFirst
   Do While adoRecordset1.EOF = False
      strTemp3 = CheckStr(adoRecordset1.Fields(0))
      PrintData1_1 (CheckStr(adoRecordset1.Fields(0)))
      PrintEnd1_1 (CheckStr(adoRecordset1.Fields(0)))
      Page = Page + 1
      Printer.NewPage
      adoRecordset1.MoveNext
   Loop
Else
   InsertQueryLog (0) 'Add By Sindy 2010/12/17
End If
CheckOC2
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintData1_1(Strindex As String)
If Len(Strindex) = 0 Then
    strSql = "SELECT r110001,r110002,r110003,r110004,r110005,r110006,r110007,r110008,r110009,r110010,r110011,r110012,r110013,r110014,r110015,r110016,r110017,r110018,r110019,r110020,r110021,r110022,id FROM R090614 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='') order by r110002 "
Else
    strSql = "SELECT r110001,r110002,r110003,r110004,r110005,r110006,r110007,r110008,r110009,r110010,r110011,r110012,r110013,r110014,r110015,r110016,r110017,r110018,r110019,r110020,r110021,r110022,id FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "' order by r110002"
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    'Modify By Cheng 2003/05/08
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle_1
        Do While .EOF = False
            For i = 0 To 20
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(5) = StrToStr(strTemp(5), 10)
            strTemp(7) = StrToStr(strTemp(7), 3)
            strTemp(8) = StrToStr(strTemp(8), 4)
            strTemp(15) = StrToStr(strTemp(15), 3)
            strTemp(19) = StrToStr(strTemp(19), 5)
            strTemp(20) = StrToStr(strTemp(20), 3)
            PrintDatil_1
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle_1
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC

End Sub

Sub PrintData2_1(Strindex As String)
If Len(Strindex) = 0 Then
    strSql = "SELECT r110001,r110002,r110003,r110004,r110005,r110006,r110007,r110008,r110009,r110010,r110011,r110012,r110013,r110014,r110015,r110016,r110017,r110018,r110019,r110020,r110021,r110022,id FROM R090614 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='') order by r110002 "
Else
    strSql = "SELECT r110001,r110002,r110003,r110004,r110005,r110006,r110007,r110008,r110009,r110010,r110011,r110012,r110013,r110014,r110015,r110016,r110017,r110018,r110019,r110020,r110021,r110022,id FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "' order by r110002 "
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    'Modify By Cheng 2003/05/08
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle_2
        Do While .EOF = False
            For i = 0 To 20
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(5) = StrToStr(strTemp(5), 10)
            strTemp(7) = StrToStr(strTemp(7), 3)
            strTemp(8) = StrToStr(strTemp(8), 4)
            strTemp(15) = StrToStr(strTemp(15), 3)
            strTemp(19) = StrToStr(strTemp(19), 5)
            strTemp(20) = StrToStr(strTemp(20), 3)
            PrintDatil_2
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle_2
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC
End Sub

Sub PrintData2()
strSql = "SELECT DISTINCT R110001 FROM R090614 WHERE ID='" & strUserNum & "' "
CheckOC2
Page = 1
adoRecordset1.CursorLocation = adUseClient
'Modify By Cheng 2003/05/08
'adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset1.Open strSql, adoEng, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
   InsertQueryLog (adoRecordset1.RecordCount) 'Add By Sindy 2010/12/17
   adoRecordset1.MoveFirst
   Do While adoRecordset1.EOF = False
      strTemp3 = CheckStr(adoRecordset1.Fields(0))
      PrintData2_1 (CheckStr(adoRecordset1.Fields(0)))
      PrintEnd2_1 (CheckStr(adoRecordset1.Fields(0)))
      Page = Page + 1
      Printer.NewPage
      adoRecordset1.MoveNext
   Loop
Else
   InsertQueryLog (0) 'Add By Sindy 2010/12/17
End If
CheckOC2
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd1_1(Strindex As String)
'列印結尾
ShowLine
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_1
End If
If Len(Strindex) = 0 Then
    'Modify By Cheng 2003/05/08
'    strSQL = "SELECT SUM(DECODE(R111003,0,0,R111003)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)),SUM(DECODE(R111006,0,0,R111006)),SUM(DECODE(R111007,0,0,R111007)),SUM(DECODE(R111008,0,0,R111008)),SUM(DECODE(R111009,0,0,R111009)) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='1' "
    'edit by nickc 2005/05/04
    'strSQL = "SELECT SUM(R111003),SUM(R111004),SUM(R111005),SUM(R111006),SUM(R111007),SUM(R111008),SUM(R111009) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='1' "
    strSql = "SELECT SUM(R111003),SUM(R111004),SUM(R111005),SUM(R111006),SUM(R111007),SUM(R111008),SUM(R111009),SUM(R111014),SUM(R111015),SUM(R111016),SUM(R111017),SUM(R111018),SUM(R111019),SUM(R111020) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='1' "
Else
    'Modify By Cheng 2003/05/08
'    strSQL = "SELECT SUM(DECODE(R111003,0,0,R111003)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)),SUM(DECODE(R111006,0,0,R111006)),SUM(DECODE(R111007,0,0,R111007)),SUM(DECODE(R111008,0,0,R111008)),SUM(DECODE(R111009,0,0,R111009)) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='1' "
    'edit by nickc 2005/05/04
    'strSQL = "SELECT SUM(R111003),SUM(R111004),SUM(R111005),SUM(R111006),SUM(R111007),SUM(R111008),SUM(R111009) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='1' "
    strSql = "SELECT SUM(R111003),SUM(R111004),SUM(R111005),SUM(R111006),SUM(R111007),SUM(R111008),SUM(R111009),SUM(R111014),SUM(R111015),SUM(R111016),SUM(R111017),SUM(R111018),SUM(R111019),SUM(R111020) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='1' "
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    'Modify By Cheng 203/05/05
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "本月收文件數：" & Format("0" & CheckStr(.Fields(0)), "###,###,###,###,##0") & " 件"
        Printer.Print "本月收文件數：" & Format("0" & CheckStr(.Fields(0)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(7)), "###,###,###,###,##0.00") & ") 件"
        iPrint = iPrint + 300
        If iPrint >= 11000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_1
        End If
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "本月發文件數：" & Format("0" & CheckStr(.Fields(1)), "###,###,###,###,##0") & " 件, "
        Printer.Print "本月發文件數：" & Format("0" & CheckStr(.Fields(1)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(8)), "###,###,###,###,##0.00") & ") 件, "
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "點數：" & Format("0" & CheckStr(.Fields(2)), "###,###,###,###,##0.00") & " 點"
        Printer.Print "點數：" & Format("0" & CheckStr(.Fields(2)), "###,###,###,###,##0.00") & "(" & Format("0" & CheckStr(.Fields(9)), "###,###,###,###,##0.00") & ") 點"
        iPrint = iPrint + 300
        If iPrint >= 11000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_1
        End If
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "目前未完稿的件數：" & Format("0" & CheckStr(.Fields(3)), "###,###,###,###,##0") & " 件"
        Printer.Print "目前未完稿的件數：" & Format("0" & CheckStr(.Fields(3)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(10)), "###,###,###,###,##0.00") & ") 件"
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "會稿中的件數：" & Format("0" & CheckStr(.Fields(4)), "###,###,###,###,##0") & " 件"
        Printer.Print "會稿中的件數：" & Format("0" & CheckStr(.Fields(4)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(11)), "###,###,###,###,##0.00") & ") 件"
        iPrint = iPrint + 300
        If iPrint >= 11000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_1
        End If
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "超過承辦期限之件數：" & Format("0" & CheckStr(.Fields(5)), "###,###,###,###,##0") & " 件"
        Printer.Print "超過承辦期限之件數：" & Format("0" & CheckStr(.Fields(5)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(12)), "###,###,###,###,##0.00") & ") 件"
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "當日法定期限之件數：" & Format("0" & CheckStr(.Fields(6)), "###,###,###,###,##0") & " 件"
        Printer.Print "當日法定期限之件數：" & Format("0" & CheckStr(.Fields(6)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(13)), "###,###,###,###,##0.00") & ") 件"
        iPrint = iPrint + 300
        If iPrint >= 11000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_1
        End If
        ShowLine
    End If
End With
CheckOC
End Sub

Sub PrintEnd2_1(Strindex As String)
Dim blnThisMonth As Boolean '發文年月與系統年月是否相符

'Add By Cheng 2003/07/18
If (Val(Me.txt1(3).Text) + 1911) & Format(Val(Me.txt1(4).Text), "00") = Left(strSrvDate(1), 6) Then
    blnThisMonth = True
Else
    blnThisMonth = False
End If
'列印結尾
ShowLine
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
End If
If Len(Strindex) = 0 Then
    'Modify By Cheng 2003/05/08
'    strSQL = "SELECT SUM(DECODE(R111010,0,0,R111010)),SUM(DECODE(R111011,0,0,R111011)),SUM(DECODE(R111012,0,0,R111012)),SUM(DECODE(R111013,0,0,R111013)),SUM(DECODE(R111009,0,0,R111009)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)),SUM(DECODE(R111008,0,0,R111008)) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='2' "
    'edit by nickc 2005/05/04
    'strSQL = "SELECT SUM(R111010),SUM(R111011),SUM(R111012),SUM(R111013),SUM(R111009),SUM(R111004),SUM(R111005),SUM(R111008) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='2' "
    strSql = "SELECT SUM(R111010),SUM(R111011),SUM(R111012),SUM(R111013),SUM(R111009),SUM(R111004),SUM(R111005),SUM(R111008),SUM(R111021),SUM(R111022),SUM(R111023),SUM(R111024),SUM(R111020),SUM(R111015),SUM(R111016),SUM(R111019) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='2' "
Else
    'Modify By Cheng 2003/05/08
'    strSQL = "SELECT SUM(DECODE(R111010,0,0,R111010)),SUM(DECODE(R111011,0,0,R111011)),SUM(DECODE(R111012,0,0,R111012)),SUM(DECODE(R111013,0,0,R111013)),SUM(DECODE(R111009,0,0,R111009)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)),SUM(DECODE(R111008,0,0,R111008)) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='2' "
    'edit by nickc 2005/05/04
    'strSQL = "SELECT SUM(R111010),SUM(R111011),SUM(R111012),SUM(R111013),SUM(R111009),SUM(R111004),SUM(R111005),SUM(R111008) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='2' "
    strSql = "SELECT SUM(R111010),SUM(R111011),SUM(R111012),SUM(R111013),SUM(R111009),SUM(R111004),SUM(R111005),SUM(R111008),SUM(R111021),SUM(R111022),SUM(R111023),SUM(R111024),SUM(R111020),SUM(R111015),SUM(R111016),SUM(R111019) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='2' "
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    'Modify By Cheng 2003/05/08
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        If blnThisMonth = False Then
            'edit by nickc 2005/05/04
            'Printer.Print "可辦非設計案件：" & Format("0" & CheckStr(.Fields(0)), "###,###,###,###,##0") & " 件"
            Printer.Print "可辦非設計案件：" & Format("0" & CheckStr(.Fields(0)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(8)), "###,###,###,###,##0.00") & ") 件"
        End If
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        If blnThisMonth = False Then
            'edit by nickc 2005/05/04
            'Printer.Print "可辦設計案件：" & Format("0" & CheckStr(.Fields(1)), "###,###,###,###,##0") & " 件"
            Printer.Print "可辦設計案件：" & Format("0" & CheckStr(.Fields(1)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(9)), "###,###,###,###,##0.00") & ") 件"
        End If
        iPrint = iPrint + 300
        If iPrint >= 11000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_2
        End If
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "本月已完稿非設計件數：" & Format("0" & CheckStr(.Fields(2)), "###,###,###,###,##0") & " 件"
        Printer.Print "本月已完稿非設計件數：" & Format("0" & CheckStr(.Fields(2)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(10)), "###,###,###,###,##0.00") & ") 件"
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "本月已完稿設計件數：" & Format("0" & CheckStr(.Fields(3)), "###,###,###,###,##0") & " 件"
        Printer.Print "本月已完稿設計件數：" & Format("0" & CheckStr(.Fields(3)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(11)), "###,###,###,###,##0.00") & ") 件"
        iPrint = iPrint + 300
        If iPrint >= 11000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_2
        End If
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        If blnThisMonth = False Then
            'edit by nickc 2005/05/04
            'Printer.Print "當日法定期限之件數：" & Format("0" & CheckStr(.Fields(4)), "###,###,###,###,##0") & " 件"
            Printer.Print "當日法定期限之件數：" & Format("0" & CheckStr(.Fields(4)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(12)), "###,###,###,###,##0.00") & ") 件"
        End If
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "本月發文件數：" & Format("0" & CheckStr(.Fields(5)), "###,###,###,###,##0") & " 件"
        Printer.Print "本月發文件數：" & Format("0" & CheckStr(.Fields(5)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(13)), "###,###,###,###,##0.00") & ") 件"
        iPrint = iPrint + 300
        If iPrint >= 11000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_2
        End If
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "本月發文點數：" & Format("0" & CheckStr(.Fields(6)), "###,###,###,###,##0.00") & " 點"
        Printer.Print "本月發文點數：" & Format("0" & CheckStr(.Fields(6)), "###,###,###,###,##0.00") & "(" & Format("0" & CheckStr(.Fields(14)), "###,###,###,###,##0.00") & ") 點"
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        If blnThisMonth = False Then
            'edit by nickc 2005/05/04
            'Printer.Print "超過承辦期限之件數：" & Format("0" & CheckStr(.Fields(7)), "###,###,###,###,##0") & " 件"
            Printer.Print "超過承辦期限之件數：" & Format("0" & CheckStr(.Fields(7)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(15)), "###,###,###,###,##0.00") & ") 件"
        End If
        iPrint = iPrint + 300
        If iPrint >= 11000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_2
        End If
        ShowLine
    End If
End With
CheckOC
End Sub

Sub PrintTitle_1() '列印抬頭
Printer.Orientation = 2
iPrint = 0
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "工作進度資料表(一般)"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "發文年月：" & txt1(3) & "/" & txt1(4)
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & GetPrjSalesNM(strTemp3)
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
ShowLine
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_1
    Exit Sub
End If
GetPleft_1
Printer.Font.Size = 9
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "目次"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "收文"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "Y/N"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "承辦"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "本所"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "法定"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "齊備日"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "會稿日"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "核稿人"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "會稿"
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "承辦"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "備註"
Printer.CurrentX = PLeft(20)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
iPrint = iPrint + 300
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_1
    Exit Sub
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "類別"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "完成日"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "天數"
iPrint = iPrint + 300
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_1
    Exit Sub
End If
ShowLine
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_1
    Exit Sub
End If
Printer.Font.Size = 9
End Sub

Sub PrintTitle_2() '列印抬頭
Printer.Orientation = 2
iPrint = 0
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "工作進度資料表(專利)"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "發文年月：" & txt1(3) & "/" & txt1(4)
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & GetPrjSalesNM(strTemp3)
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
ShowLine
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
    Exit Sub
End If
GetPleft_2
Printer.Font.Size = 9
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "目次"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "收文"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "Y/N"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "承辦"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "本所"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "法定"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "齊備日"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "會稿日"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "核稿人"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "會稿"
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "承辦"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "備註"
Printer.CurrentX = PLeft(20)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
iPrint = iPrint + 300
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
    Exit Sub
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "類別"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "完成日"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "天數"
iPrint = iPrint + 300
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
    Exit Sub
End If
ShowLine
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
    Exit Sub
End If
Printer.Font.Size = 9
End Sub

Sub PrintDatil_1() '列印資料

For i = 1 To 20
    If i = 1 Or i = 18 Then
        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp(i), "####0")
    Else
        Printer.CurrentX = PLeft(i)
        Printer.CurrentY = iPrint
        Printer.Print strTemp(i)
    End If
Next i
iPrint = iPrint + 300
End Sub

Sub PrintDatil_2() '列印資料

For i = 1 To 20
    If i = 1 Or i = 18 Then
        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp(i), "####0")
    Else
        Printer.CurrentX = PLeft(i)
        Printer.CurrentY = iPrint
        Printer.Print strTemp(i)
    End If
Next i
iPrint = iPrint + 300
End Sub


Sub GetPleft_1()
'定陣列
'字 SIZE = 9
'1 WORD = 180 PIX
'0.5 WORD = 90 PIX
'SPACE = 90 PIX
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
PLeft(2) = PLeft(1) + (2.5 * 180)
PLeft(3) = PLeft(2) + (2.5 * 180)
'Modified by Lydia 2016/12/16 + 50
PLeft(4) = PLeft(3) + (4.5 * 180) + 50
'Modified by Lydia 2016/12/16
'PLeft(5) = PLeft(4) + (8 * 180)
PLeft(5) = PLeft(4) + (9 * 180)
PLeft(6) = PLeft(5) + (10.5 * 180)
PLeft(7) = PLeft(6) + (2 * 180)
PLeft(8) = PLeft(7) + (3.5 * 180)
PLeft(9) = PLeft(8) + (4.5 * 180)
PLeft(10) = PLeft(9) + (4.5 * 180)
PLeft(11) = PLeft(10) + (4.5 * 180)
PLeft(12) = PLeft(11) + (4.5 * 180)
PLeft(13) = PLeft(12) + (4.5 * 180)
PLeft(14) = PLeft(13) + (4.5 * 180)
PLeft(15) = PLeft(14) + (4.5 * 180)
PLeft(16) = PLeft(15) + (4.5 * 180)
PLeft(17) = PLeft(16) + (4.5 * 180)
PLeft(18) = PLeft(17) + (4.5 * 180)
PLeft(19) = PLeft(18) + (2.5 * 180)
PLeft(20) = PLeft(19) + (5.5 * 180)
End Sub

Sub GetPleft_2()
'定陣列
'字 SIZE = 9
'1 WORD = 180 PIX
'0.5 WORD = 90 PIX
'SPACE = 90 PIX
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
PLeft(2) = PLeft(1) + (2.5 * 180)
PLeft(3) = PLeft(2) + (2.5 * 180)
'Modified by Lydia 2016/12/16 + 50
PLeft(4) = PLeft(3) + (4.5 * 180) + 50
'Modified by Lydia 2016/12/16
'PLeft(5) = PLeft(4) + (8 * 180)
PLeft(5) = PLeft(4) + (9 * 180)
PLeft(6) = PLeft(5) + (10.5 * 180)
PLeft(7) = PLeft(6) + (2 * 180)
PLeft(8) = PLeft(7) + (3.5 * 180)
PLeft(9) = PLeft(8) + (4.5 * 180)
PLeft(10) = PLeft(9) + (4.5 * 180)
PLeft(11) = PLeft(10) + (4.5 * 180)
PLeft(12) = PLeft(11) + (4.5 * 180)
PLeft(13) = PLeft(12) + (4.5 * 180)
PLeft(14) = PLeft(13) + (4.5 * 180)
PLeft(15) = PLeft(14) + (4.5 * 180)
PLeft(16) = PLeft(15) + (4.5 * 180)
PLeft(17) = PLeft(16) + (4.5 * 180)
PLeft(18) = PLeft(17) + (4.5 * 180)
PLeft(19) = PLeft(18) + (2.5 * 180)
PLeft(20) = PLeft(19) + (5.5 * 180)
End Sub

Sub ShowLine()
'Modified by Lydia 2016/12/16
'Printer.Line (0, iPrint + 150)-(16500, iPrint + 150)
Printer.Line (0, iPrint + 150)-(16700, iPrint + 150)
iPrint = iPrint + 300
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txt1(5) = "1"
   '2005/8/30 ADD BY SONIA
   txt1(3) = Val(Left(ServerDate, 4)) - 1911
   txt1(4) = Val(Mid(ServerDate, 5, 2))
   '2005/8/30 END
   
   'Add By Sindy 2024/2/23 因幫總經理預設單位,所以只能在 mdiMain 傳入應該操作那一個單位值
   'modify by sonia 2024/3/7
   'If m_ProState = "P" And Left(Pub_StrUserSt03, 1) <> "P" Then
   If m_ProState = "P" Then
      txt1(6) = "P10"
      txt1(7) = "P11"
   'modify by sonia 2024/3/7
   'ElseIf m_ProState = "T" And Left(Pub_StrUserSt03, 1) <> "P" Then
   ElseIf m_ProState = "T" Then
      txt1(6) = "P20"
      txt1(7) = "P21"
   ElseIf m_ProState = "FCP" And Left(Pub_StrUserSt03, 1) <> "F" Then
      txt1(6) = "F20"
      txt1(7) = "F29"
   ElseIf m_ProState = "ACS" And Left(Pub_StrUserSt03, 1) <> "W" Then
      txt1(6) = "W20"
      txt1(7) = "W20"
   Else
      txt1(6) = GetStaffDepartment(strUserNum)
      txt1(7) = GetStaffDepartment(strUserNum)
   End If
   '2024/2/23 END
'   '2007/5/15 ADD BY SONIA
'   'Modify By Sindy 2021/11/24
'   '1.P1X人員預設P10~P11，但最大範圍只可為P10~P19且不可空白。
'   '2.P2X人員預設P20~P21，但最大範圍只可為P20~P29且不可空白。
'   '3.其他人員不預設，但起迄欄之第一碼必須相同且不可空白。
'   '  林文雄A4023以P1X人員身份進入
'   If Mid(GetStaffDepartment(strUserNum), 1, 2) = "P1" Or strUserNum = "A4023" Then
'      Txt1(6) = "P10"
'      Txt1(7) = "P11"
'   'Modify By Sindy 2021/11/24 江協理98020、林律師98003進入改為以P2X人員身份進入。
'   'Modify By Sindy 2021/11/26
'   '  沈佳穎96003以P2X人員身份進入
'   ElseIf Mid(GetStaffDepartment(strUserNum), 1, 2) = "P2" Or _
'      strUserNum = "98020" Or strUserNum = "98003" Or strUserNum = "96003" Then
'      Txt1(6) = "P20"
'      Txt1(7) = "P21"
'   Else
'      'Modify By Sindy 2024/1/3
'      'Txt1(6) = ""
'      'Txt1(7) = ""
'      Txt1(6) = GetStaffDepartment(strUserNum)
'      Txt1(7) = GetStaffDepartment(strUserNum)
'      '2024/1/3 END
'   End If
   
   lbl1(0).Caption = "" 'Add By Sindy 2021/12/16
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mForm = Nothing 'Add By Sindy 2023/6/20
Set frm090614 = Nothing
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
'Add By Cheng 2003/06/02
Select Case Index
   Case 5 '顯示方式
      If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
          KeyAscii = 0
      End If
   Case 8 '螢幕顯示
      If KeyAscii <> 8 And KeyAscii <> 78 Then
         KeyAscii = 0
      End If
End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     Select Case Trim(txt1(0))
     Case "1", "2", "", "3", "4", "5"
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(0).SetFocus
          txt1(0).SelStart = 0
          txt1(0).SelLength = Len(txt1(0))
          Exit Sub
     End Select
Case 1
     Select Case Trim(txt1(1))
     Case "1", "2", "", "3", "4", "5"
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(1).SetFocus
          txt1(1).SelStart = 0
          txt1(1).SelLength = Len(txt1(1))
          Exit Sub
     End Select
     If RunNick(txt1(0), txt1(1)) Then
         txt1(0).SetFocus
         txt1_GotFocus (0)
         Exit Sub
      End If
'承辦人
Case 2
      'Modify By Cheng 2002/01/09
'     lbl1(0).Caption = GetPrjSales(txt1(2))
Case 3
     If IsNumeric(txt1(Index)) = False And Len(txt1(Index)) <> 0 Then
            s = MsgBox("年輸入錯誤!!", , "USER 輸入錯誤")
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
     End If
Case 4
        If IsNumeric(txt1(Index)) = False And Len(txt1(Index)) <> 0 Then
            s = MsgBox("月輸入錯誤!!", , "USER 輸入錯誤")
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
        Else
            If (Val(txt1(Index)) < 1 Or Val(txt1(Index)) > 12) And txt1(Index) <> "" Then
                s = MsgBox("月輸入錯誤!!", , "USER 輸入錯誤")
                txt1(Index).SetFocus
                txt1(Index).SelStart = 0
                txt1(Index).SelLength = Len(txt1(Index))
                Exit Sub
            End If
         End If
Case 5
     Select Case Trim(txt1(5))
     Case "1", "2", ""
     Case Else
          s = MsgBox("顯示方式只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          txt1(5).SetFocus
          txt1(5).SelStart = 0
          txt1(5).SelLength = Len(txt1(5))
          Exit Sub
     End Select
Case 7 '部門別(迄)
      If RunNick(txt1(6), txt1(7)) Then
         txt1(6).SetFocus
         txt1_GotFocus (6)
         Exit Sub
      End If
Case Else
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Cancel = False
m_bln_KeyinValid = True
Select Case Index
Case 2 '承辦人
     'Add By Cheng 2002/01/09
     lbl1(0).Caption = GetPrjSales(txt1(2))
     '92.6.27 ADD BY SONIA
     If lbl1(0).Caption <> "" Then
         txt1(8) = ""
         'Add By Sindy 2012/5/31 輸入承辦人時,一併帶入此人員的部門別至部門別起迄欄位中
         txt1(6) = GetStaffDepartment(txt1(Index))
         txt1(7) = GetStaffDepartment(txt1(Index))
         '2012/5/31 End
     End If
     '92.6.27 END
     If Me.lbl1(0).Caption <> "" And Me.lbl1(0).Caption = Me.txt1(2).Text Then
       Me.lbl1(0).Caption = ""
       Me.txt1(2).SetFocus
       txt1_GotFocus 2
       Cancel = True
       m_bln_KeyinValid = False
     End If
End Select
End Sub
