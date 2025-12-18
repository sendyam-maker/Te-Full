VERSION 5.00
Begin VB.Form frm12040143 
   BorderStyle     =   1  '單線固定
   Caption         =   "逾期未處理案件明細表"
   ClientHeight    =   3780
   ClientLeft      =   3480
   ClientTop       =   3195
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3750
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2268
      TabIndex        =   11
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1476
      TabIndex        =   10
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   2730
      MaxLength       =   9
      TabIndex        =   9
      Top             =   2490
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1410
      MaxLength       =   9
      TabIndex        =   8
      Top             =   2490
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1410
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2190
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1605
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1890
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1365
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1590
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2685
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1290
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1365
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1290
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2685
      MaxLength       =   7
      TabIndex        =   2
      Top             =   990
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1365
      MaxLength       =   7
      TabIndex        =   1
      Top             =   990
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1410
      TabIndex        =   0
      Top             =   690
      Width           =   2100
   End
   Begin VB.OptionButton Option1 
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   1
      Left            =   285
      TabIndex        =   14
      Top             =   1290
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所期限:"
      Height          =   180
      Index           =   0
      Left            =   285
      TabIndex        =   13
      Top             =   990
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "       2. 管制智權人員者只印未收文者"
      Height          =   180
      Left            =   360
      TabIndex        =   23
      Top             =   3120
      Width           =   2835
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "PS : 1. 管制承辦人者只印收文未發文者"
      Height          =   180
      Left            =   360
      TabIndex        =   22
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Line Line3 
      X1              =   2370
      X2              =   2610
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "申請人:"
      Height          =   180
      Left            =   330
      TabIndex        =   21
      Top             =   2505
      Width           =   585
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Height          =   180
      Index           =   1
      Left            =   2250
      TabIndex        =   20
      Top             =   2220
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   330
      TabIndex        =   19
      Top             =   2220
      Width           =   765
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Height          =   180
      Index           =   0
      Left            =   2505
      TabIndex        =   18
      Top             =   1935
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "智權人員/承辦人:"
      Height          =   180
      Left            =   285
      TabIndex        =   17
      Top             =   1935
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1. 承辦人 2.智權人員)"
      Height          =   180
      Left            =   1845
      TabIndex        =   16
      Top             =   1590
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "管制對象:"
      Height          =   180
      Left            =   285
      TabIndex        =   15
      Top             =   1605
      Width           =   765
   End
   Begin VB.Line Line2 
      X1              =   2325
      X2              =   2565
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Line Line1 
      X1              =   2325
      X2              =   2565
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "系統類別:"
      Height          =   180
      Left            =   330
      TabIndex        =   12
      Top             =   690
      Width           =   765
   End
End
Attribute VB_Name = "frm12040143"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer, iLine As Integer, k As Integer
Dim strSql As String, StrTest As String
Dim strTemp As Variant, strTemp1 As Variant, StrTempP As Variant, StrTempP2 As Variant
Dim Page As Integer, iPrint As Integer, St As String, TmpArea As String
Dim PLeft1(0 To 7) As Integer, Pleft2(0 To 10) As Integer, PLeft3(0 To 8) As Integer
Dim strSQL2 As String, strSQL1 As String, StrSQL3 As String, StrSQL6 As String
Dim STRSTRING As String
Dim StrR001001 As String
Dim StrR001002 As String
Dim StrR001003 As String
Dim StrR001004 As String
Dim StrR001005 As String
Dim StrR001006 As String
Dim StrR001007 As String
Dim StrR001008 As String
Dim StrR001009 As String
Dim StrR001010 As String
Dim StrR001011 As String
Dim StrR001012 As String
Dim StrR001013 As String
Dim StrR001014 As String
Dim StrR001015 As String
Dim StrR001016 As String
Dim StrR001017 As String
Dim StrR001018 As String
Dim StrR001019 As String

Dim StrR002001 As String
Dim StrR002002 As String
Dim StrR002003 As String
Dim StrR002004 As String
Dim StrR002005 As String
Dim StrR002006 As String
Dim StrR002007 As String
Dim StrR002008 As String
Dim StrR002009 As String
Dim StrR002010 As String
Dim StrR002011 As String
Dim StrR002012 As String
Dim StrR002013 As String
Dim StrR002014 As String
Dim StrR002015 As String

Dim StrR003001 As String
Dim StrR003002 As String
Dim StrR003003 As String
Dim StrR003004 As String
Dim StrR003005 As String
Dim StrR003006 As String
Dim StrR003007 As String
Dim StrR003008 As String
Dim StrR003009 As String
Dim StrR003010 As String
Dim StrR003011 As String
Dim StrR003012 As String
Dim StrR003013 As String
Dim StrR003014 As String
Dim StrR003015 As String
Dim StrR003016 As String
Dim StrR003017 As String
Dim StrR003018 As String
Dim StrR003019 As String
Dim StrR003020 As String
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0 '確定
    blnClkSure = False
    If Len(Txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白", , "USER 輸入錯誤")
         Txt1(0).SetFocus
         Exit Sub
    Else
         If Option1(0).Value = True Then
            If PUB_CheckKeyInDate(Me.Txt1(1)) = -1 Then
               Me.Txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.Txt1(2)) = -1 Then
               Me.Txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            If Me.Txt1(1).Text <> "" And Me.Txt1(2).Text <> "" Then
               If Val(Me.Txt1(1).Text) > Val(Me.Txt1(2).Text) Then
                  MsgBox "本所期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.Txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
                  
            If Len(Trim(Txt1(2))) = 0 Then
                s = MsgBox("本所期限不可空白", , "USER 輸入錯誤")
                Txt1(1).SetFocus
                txt1_GotFocus (1)
                Exit Sub
            End If
         Else
            If Option1(1).Value = True Then
               If PUB_CheckKeyInDate(Me.Txt1(3)) = -1 Then
                  Me.Txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
               If PUB_CheckKeyInDate(Me.Txt1(4)) = -1 Then
                  Me.Txt1(4).SetFocus
                  txt1_GotFocus 4
                  Exit Sub
               End If
               If Me.Txt1(3).Text <> "" And Me.Txt1(4).Text <> "" Then
                  If Val(Me.Txt1(3).Text) > Val(Me.Txt1(4).Text) Then
                     MsgBox "法定期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.Txt1(3).SetFocus
                     txt1_GotFocus 3
                     Exit Sub
                  End If
               End If
                               
                If Len(Trim(Txt1(4))) = 0 Then
                    s = MsgBox("法定期限不可空白", , "USER 輸入錯誤")
                    Txt1(3).SetFocus
                    txt1_GotFocus (3)
                    Exit Sub
                End If
            End If
         End If
         If Len(Trim(Txt1(5))) = 0 Then
            s = MsgBox("管制對象不可空白", , "USER 輸入錯誤")
            Txt1(5).SetFocus
            Exit Sub
         End If
     End If
     If Txt1(6) <> "" Then
         'edit by nickc 2007/02/09 不用 dll 了
         'If objPublicData.GetStaff(txt1(6), strExc(0)) Then
         If ClsPDGetStaff(Txt1(6), strExc(0)) Then
            LBL1(0) = strExc(0)
         Else
            LBL1(0) = ""
            Me.Txt1(6).SetFocus
            txt1_GotFocus 6
            Exit Sub
         End If
      End If
      If Len(Txt1(7)) <> 0 Then
         LBL1(1) = GetPrjState6HM("P", Txt1(7))
         If LBL1(1) = "" Then
            MsgBox "案件性質錯誤，請重新輸入 !", vbCritical
            Me.Txt1(7).SetFocus
            txt1_GotFocus 7
            Exit Sub
         End If
      End If
     
     If Len(Trim(Txt1(8))) <> 0 Or Len(Trim(Txt1(9))) <> 0 Then
        If Left(Txt1(8), 6) <> Left(Txt1(9), 6) Then
            s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
            blnClkSure = True
             Txt1(8).SetFocus
             txt1_GotFocus (8)
            Exit Sub
        End If
     End If
      If Me.Txt1(8).Text <> "" And Me.Txt1(9).Text <> "" Then
         If Me.Txt1(8).Text > Me.Txt1(9).Text Then
            MsgBox "申請人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            blnClkSure = True
            Me.Txt1(8).SetFocus
            txt1_GotFocus 8
            Exit Sub
         End If
      End If
     
     Screen.MousePointer = vbHourglass
     Me.Enabled = False
     StrMenu
     Me.Enabled = True
     Screen.MousePointer = vbDefault
Case 1 '結束
    Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Txt1(0) = GetSystemKindByNick
Option1(1).Value = False
Txt1(3).Enabled = False
Txt1(4).Enabled = False
End Sub

Sub StrMenu()
Screen.MousePointer = vbHourglass
Me.Enabled = False
If Txt1(5) = "1" And Option1(0).Value = True Then
    StrMenu1 '本所期限+承辦人(管制對象)
Else
    If Txt1(5) = "1" And Option1(1).Value = True Then
        StrMenu2 '法定期限+承辦人(管制對象)
    Else
        If Txt1(5) = "2" And Option1(0).Value = True Then
            StrMenu3 '本所期限+智權人員(管制對象)
        Else
            StrMenu2 '法定期限+智權人員(管制對象)
        End If
    End If
End If
Me.Enabled = True
Screen.MousePointer = vbDefault
End Sub

Sub StrMenu1()              '處理主程式
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050302_1 WHERE ID='" & strUserNum & "' "
StrSQL3 = ""
If Len(Trim(Txt1(6))) <> 0 Then
   If Trim(Txt1(5)) = "1" Then
      StrSQL3 = StrSQL3 & " AND CP14='" & Txt1(6) & "' "
   Else
      StrSQL3 = StrSQL3 & " AND CP13='" & Txt1(6) & "' "
   End If
End If
If Len(Trim(Txt1(7))) <> 0 Then
    StrSQL3 = StrSQL3 & " AND CP10='" & Txt1(7) & "' "
End If
If Len(Txt1(0)) <> 0 Then
    StrSQL3 = StrSQL3 & " AND CP01 IN (" & GetAddStr(Txt1(0)) & ") "
End If
'Modified by Lydia 2016/12/21 排除D類收文
'strSql = "SELECT CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP44 FROM CASEPROGRESS WHERE CP06>=" & Val(ChangeTStringToWString(txt1(1))) & " AND CP06<=" & Val(ChangeTStringToWString(txt1(2))) & " AND CP27 IS NULL AND CP57 IS NULL " & StrSQL3
strSql = "SELECT /*+ INDEX(CASEPROGRESS IDXCP15815905011014) */ CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP44 FROM CASEPROGRESS " & _
         "WHERE CP06>=" & Val(ChangeTStringToWString(Txt1(1))) & " AND CP06<=" & Val(ChangeTStringToWString(Txt1(2))) & " AND CP158=0 AND CP159 = 0 AND SUBSTR(CP09,1,1) <> 'D' " & StrSQL3
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
If adoRecordset.RecordCount <> 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        DoEvents
        StrR002001 = ""
        StrR002002 = ""
        StrR002003 = ""
        StrR002004 = ""
        StrR002005 = ""
        StrR002006 = ""
        StrR002007 = ""
        StrR002008 = ""
        StrR002009 = ""
        StrR002010 = ""
        StrR002011 = ""
        StrR002012 = ""
        StrR002013 = ""
        StrR002014 = ""
        StrR002015 = ""
        If Not IsNull(adoRecordset.Fields(0)) Then
            StrR002001 = adoRecordset.Fields(0)
        Else
            StrR002001 = ""
        End If
        If Not IsNull(adoRecordset.Fields(1)) Then
            StrR002002 = adoRecordset.Fields(1)
        Else
            StrR002002 = ""
        End If
        If Not IsNull(adoRecordset.Fields(2)) Then
            StrR002003 = adoRecordset.Fields(2)
        Else
            StrR002003 = ""
        End If
        If Not IsNull(adoRecordset.Fields(3)) Then
            StrR002004 = adoRecordset.Fields(3)
        Else
            StrR002004 = ""
        End If
        If Not IsNull(adoRecordset.Fields(4)) Then
            StrR002014 = adoRecordset.Fields(4)
        Else
            StrR002014 = ""
        End If
        If Not IsNull(adoRecordset.Fields(5)) Then
            StrR002015 = ChangeTStringToTDateString(ChangeWStringToTString(adoRecordset.Fields(5)))
        Else
            StrR002015 = ""
        End If
        If Not IsNull(adoRecordset.Fields(6)) Then
            StrR002007 = adoRecordset.Fields(6)
        Else
            StrR002007 = ""
        End If
        '智權人員代號
        If Not IsNull(adoRecordset.Fields(7)) Then
            StrR002005 = adoRecordset.Fields(7)
        Else
            StrR002005 = ""
        End If
        If Not IsNull(adoRecordset.Fields(8)) Then
            StrR002006 = adoRecordset.Fields(8)
        Else
            StrR002006 = ""
        End If
        If Not IsNull(adoRecordset.Fields(9)) Then
            StrR002008 = adoRecordset.Fields(9)
        Else
            StrR002008 = ""
        End If
        DoEvents
         
         'Modify By Cheng 2002/01/29
        '若專利基本檔或服務業務基本檔已閉卷, 則不列印
'        DoPaAndSp1
        If DoPaAndSp1_1 <> 0 Then GoTo NextRecord
        
        s = 1
        If Len(Trim(Txt1(8))) <> 0 And s <> 0 Then
            If (IIf(StrR002009 = "", False, StrR002009 >= GetNewFagent(Txt1(8))) Or IIf(StrR002010 = "", False, StrR002010 >= GetNewFagent(Txt1(8))) Or IIf(StrR002011 = "", False, StrR002011 >= GetNewFagent(Txt1(8))) Or IIf(StrR002012 = "", False, StrR002012 >= GetNewFagent(Txt1(8))) Or IIf(StrR002013 = "", False, StrR002013 >= GetNewFagent(Txt1(8)))) Then
                s = 1
            Else
                s = 0
            End If
        Else
            s = 1
        End If
        '911023 nick
        '***** start
        If s <> 0 Then
            If Len(Trim(Txt1(9))) <> 0 Then
                If (IIf(StrR002009 = "", False, StrR002009 <= GetNewFagent(Txt1(9))) Or IIf(StrR002010 = "", False, StrR002010 <= GetNewFagent(Txt1(9))) Or IIf(StrR002011 = "", False, StrR002011 <= GetNewFagent(Txt1(9))) Or IIf(StrR002012 = "", False, StrR002012 <= GetNewFagent(Txt1(9))) Or IIf(StrR002013 = "", False, StrR002013 <= GetNewFagent(Txt1(9)))) Then
                    s = 1
                Else
                    s = 0
                End If
            End If
        End If
        '***** end
        StrR002008 = ""
        DoEvents
        If s = 0 Then
            adoRecordset.Delete
        Else
            StrR002005 = GetPrjSalesNM(StrR002005)
            StrR002006 = StrR002006
            StrR002007 = GetPrjState4(StrR002001 + "-" + StrR002002 + "-" + StrR002003 + "-" + StrR002004, StrR002007)
            StrR002008 = GetPrjName1(StrR002008)
            StrR002009 = GetPrjPeople1(StrR002009)
            StrR002010 = GetPrjPeople1(StrR002010)
            StrR002011 = GetPrjPeople1(StrR002011)
            StrR002012 = GetPrjPeople1(StrR002012)
            StrR002013 = GetPrjPeople1(StrR002013)
            If Val(StrR002014) < Val(GetTodayDate) Then
                StrR002014 = "*" + ChangeTStringToTDateString(ChangeWStringToTString(StrR002014))
            Else
                If Val(StrR002014) = Val(GetTodayDate) Then
                    StrR002014 = "V" + ChangeTStringToTDateString(ChangeWStringToTString(StrR002014))
                Else
                    StrR002014 = ChangeTStringToTDateString(ChangeWStringToTString(StrR002014))
                End If
            End If
            'Modified by Lydia 2016/12/21 字串中斷問題
'            If LenB(StrR002005) > 8 Then
'                StrR002005 = StrToStr(StrR002005, 4)
'            End If
'            If LenB(StrR002006) > 8 Then
'                StrR002006 = StrToStr(StrR002006, 4)
'            End If
'            If LenB(StrR002007) > 12 Then
'                StrR002007 = StrToStr(StrR002007, 6)
'            End If
'            If LenB(StrR002008) > 9 Then
'                StrR002008 = StrToStr(StrR002008, 4)
'            End If
'            If LenB(StrR002009) > 12 Then
'                StrR002009 = StrToStr(StrR002009, 6)
'            End If
'            If LenB(StrR002010) > 12 Then
'                StrR002010 = StrToStr(StrR002010, 6)
'            End If
'            If LenB(StrR002011) > 12 Then
'                StrR002011 = StrToStr(StrR002011, 6)
'            End If
'            If LenB(StrR002012) > 12 Then
'                StrR002012 = StrToStr(StrR002012, 6)
'            End If
'            If LenB(StrR002013) > 12 Then
'                StrR002013 = StrToStr(StrR002013, 6)
'            End If
             StrR002005 = StrToStr(StrR002005, 12)
             StrR002006 = StrToStr(StrR002006, 12)
             StrR002007 = StrToStr(StrR002007, 140)
             StrR002008 = StrToStr(StrR002008, 80)
             StrR002009 = StrToStr(StrR002009, 120)
             StrR002010 = StrToStr(StrR002010, 120)
             StrR002011 = StrToStr(StrR002011, 120)
             StrR002012 = StrToStr(StrR002012, 120)
             StrR002013 = StrToStr(StrR002013, 120)
             'end 2016/12/21
            strSql = "INSERT INTO R050302_1 VALUES ('" & ChgSQL(StrR002001) & "','" & ChgSQL(StrR002002) & "','" & ChgSQL(StrR002003) & "','" & ChgSQL(StrR002004) & "','" & ChgSQL(StrR002005) & "','" & ChgSQL(StrR002006) & "','" & ChgSQL(StrR002007) & "','" & ChgSQL(StrR002008) & "','" & ChgSQL(StrR002009) & "','" & ChgSQL(StrR002010) & "','" & ChgSQL(StrR002011) & "','" & ChgSQL(StrR002012) & "','" & ChgSQL(StrR002013) & "','" & ChgSQL(StrR002014) & "','" & ChgSQL(StrR002015) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
        End If

'Add By Cheng 2002/01/29
NextRecord:

        adoRecordset.MoveNext
        DoEvents
    Loop
    If adoRecordset.RecordCount = 0 Then
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
Else
    ShowNoData
    Screen.MousePointer = vbDefault
    Exit Sub
End If
PriMenu1
End Sub

Sub StrMenu2()              '處理主程式
Dim strWhereSql As String

Screen.MousePointer = vbHourglass
cnnConnection.Execute "delete from R050302_2 WHERE ID='" & strUserNum & "' "
strWhereSql = ""

''強迫寫入"1"表示資料來自案件進度檔
'strSql = "SELECT CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP44,CP05,CP09,CP64,'1' FROM CASEPROGRESS WHERE CP07>=" & Val(ChangeTStringToWString(txt1(3))) & " AND CP07<=" & Val(ChangeTStringToWString(txt1(4))) & " AND CP27 IS NULL AND CP57 IS NULL "
'If txt1(5) = "1" Then
'   If Len(txt1(6)) <> 0 Then
'      strSql = strSql & " AND CP14='" & txt1(6) & "' "
'   End If
'Else
'   If Len(txt1(6)) <> 0 Then
'      strSql = strSql & " AND CP13='" & txt1(6) & "' "
'   End If
'End If
'If Len(txt1(7)) <> 0 Then
'   strSql = strSql & " AND CP10='" & Val(txt1(7)) & "' "
'End If
'CheckOC
'If Len(txt1(0)) <> 0 Then
'    strSql = strSql & " AND CP01 IN (" & GetAddStr(txt1(0)) & ") "
'End If
''強迫寫入"2"表示資料來自下一程序檔

If Len(Txt1(6)) <> 0 And Txt1(5) = "2" Then
   strWhereSql = strWhereSql & " AND NP10='" & Txt1(6) & "' "
End If
If Len(Txt1(7)) <> 0 Then
   strWhereSql = strWhereSql & " AND NP07=" & Val(Txt1(7)) & " "
End If
If Len(Txt1(0)) <> 0 Then
    strWhereSql = strWhereSql & " AND NP02 IN (" & GetAddStr(Txt1(0)) & ") "
End If

'2006/4/6 MODIFY BY SONIA 管制智權人員只抓未收文資料,
'2010/3/22 MODIFY BY SONIA 程序管制案件性質不印改以strNpSqlOfNoSalesDuty控制
'Modify By Sindy 2011/3/15 FCT,T,TF延展(102)和第二期(716)專用權須存在(TM17=Y)
'strSql = "Select NP02,NP03,NP04,NP05,NP08,NP09,NP07,NP10,CP14,CP44,CP05,CP09,NP15,'2' FROM NEXTPROGRESS,CASEPROGRESS WHERE NP09>=" & Val(ChangeTStringToWString(txt1(3))) & " AND NP09<=" & Val(ChangeTStringToWString(txt1(4))) & " AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') AND np01=cp09(+) AND CP27 IS NOT NULL" & strNpSqlOfNoSalesDuty
'Modified by Lydia 2016/12/21 +排除D類收文 AND SUBSTR(CP09,1,1) <> 'D'
strSql = "Select NP02,NP03,NP04,NP05,NP08,NP09,NP07,NP10,CP14,CP44,CP05,CP09,NP15,'2' FROM NEXTPROGRESS,CASEPROGRESS,trademark WHERE NP09>=" & Val(ChangeTStringToWString(Txt1(3))) & " AND NP09<=" & Val(ChangeTStringToWString(Txt1(4))) & " AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') AND np01=cp09(+) AND CP27 IS NOT NULL AND SUBSTR(CP09,1,1) <> 'D' and np02=tm01 and np03=tm02 and np04=tm03 and np05=tm04 and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' " & strNpSqlOfNoSalesDuty & strWhereSql & _
             " union Select NP02,NP03,NP04,NP05,NP08,NP09,NP07,NP10,CP14,CP44,CP05,CP09,NP15,'2' FROM NEXTPROGRESS,CASEPROGRESS,patent WHERE NP09>=" & Val(ChangeTStringToWString(Txt1(3))) & " AND NP09<=" & Val(ChangeTStringToWString(Txt1(4))) & " AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') AND np01=cp09(+) AND CP27 IS NOT NULL AND SUBSTR(CP09,1,1) <> 'D' and np02=pa01 and np03=pa02 and np04=pa03 and np05=pa04 " & strNpSqlOfNoSalesDuty & strWhereSql & _
             " union Select NP02,NP03,NP04,NP05,NP08,NP09,NP07,NP10,CP14,CP44,CP05,CP09,NP15,'2' FROM NEXTPROGRESS,CASEPROGRESS,servicepractice WHERE NP09>=" & Val(ChangeTStringToWString(Txt1(3))) & " AND NP09<=" & Val(ChangeTStringToWString(Txt1(4))) & " AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') AND np01=cp09(+) AND CP27 IS NOT NULL AND SUBSTR(CP09,1,1) <> 'D' and np02=sp01 and np03=sp02 and np04=sp03 and np05=sp04 " & strNpSqlOfNoSalesDuty & strWhereSql & _
             " union Select NP02,NP03,NP04,NP05,NP08,NP09,NP07,NP10,CP14,CP44,CP05,CP09,NP15,'2' FROM NEXTPROGRESS,CASEPROGRESS,lawcase WHERE NP09>=" & Val(ChangeTStringToWString(Txt1(3))) & " AND NP09<=" & Val(ChangeTStringToWString(Txt1(4))) & " AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') AND np01=cp09(+) AND CP27 IS NOT NULL AND SUBSTR(CP09,1,1) <> 'D' and np02=lc01 and np03=lc02 and np04=lc03 and np05=lc04 " & strNpSqlOfNoSalesDuty & strWhereSql & _
             " union Select NP02,NP03,NP04,NP05,NP08,NP09,NP07,NP10,CP14,CP44,CP05,CP09,NP15,'2' FROM NEXTPROGRESS,CASEPROGRESS,hirecase WHERE NP09>=" & Val(ChangeTStringToWString(Txt1(3))) & " AND NP09<=" & Val(ChangeTStringToWString(Txt1(4))) & " AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') AND np01=cp09(+) AND CP27 IS NOT NULL AND SUBSTR(CP09,1,1) <> 'D' and np02=hc01 and np03=hc02 and np04=hc03 and np05=hc04 " & strNpSqlOfNoSalesDuty & strWhereSql
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
If adoRecordset.RecordCount <> 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        DoEvents
        StrR003001 = ""
        StrR003002 = ""
        StrR003003 = ""
        StrR003004 = ""
        StrR003005 = ""
        StrR003006 = ""
        StrR003007 = ""
        StrR003008 = ""
        StrR003009 = ""
        StrR003010 = ""
        StrR003011 = ""
        StrR003012 = ""
        StrR003013 = ""
        StrR003014 = ""
        StrR003015 = ""
        StrR003016 = ""
        StrR003017 = ""
        StrR003018 = ""
        StrR003019 = ""
        StrR003020 = ""
        
        If Not IsNull(adoRecordset.Fields(0)) Then
            StrR003001 = adoRecordset.Fields(0)
        Else
            StrR003001 = ""
        End If
        If Not IsNull(adoRecordset.Fields(1)) Then
            StrR003002 = adoRecordset.Fields(1)
        Else
            StrR003002 = ""
        End If
        If Not IsNull(adoRecordset.Fields(2)) Then
            StrR003003 = adoRecordset.Fields(2)
        Else
            StrR003003 = ""
        End If
        If Not IsNull(adoRecordset.Fields(3)) Then
            StrR003004 = adoRecordset.Fields(3)
        Else
            StrR003004 = ""
        End If
        If Not IsNull(adoRecordset.Fields(4)) Then
            StrR003015 = ChangeTStringToTDateString(ChangeWStringToTString(adoRecordset.Fields(4)))
        Else
            StrR003015 = ""
        End If
        If Not IsNull(adoRecordset.Fields(5)) Then
            StrR003016 = ChangeTStringToTDateString(ChangeWStringToTString(adoRecordset.Fields(5)))
        Else
            StrR003016 = ""
        End If
        If Not IsNull(adoRecordset.Fields(6)) Then
            StrR003007 = adoRecordset.Fields(6)
        Else
            StrR003007 = ""
        End If
        '智權人員代號
        If Not IsNull(adoRecordset.Fields(7)) Then
            StrR003005 = adoRecordset.Fields(7)
        Else
            StrR003005 = ""
        End If
        If Not IsNull(adoRecordset.Fields(8)) Then
            StrR003006 = adoRecordset.Fields(8)
        Else
            StrR003006 = ""
        End If
        If Not IsNull(adoRecordset.Fields(10)) Then
            StrR003017 = ChangeTStringToTDateString(ChangeWStringToTString(adoRecordset.Fields(10)))
        Else
            StrR003017 = ""
        End If
        If Not IsNull(adoRecordset.Fields(11)) Then
            StrR003018 = Left(adoRecordset.Fields(11), 1)
        Else
            StrR003018 = ""
        End If
        If Not IsNull(adoRecordset.Fields(12)) Then
            StrR003019 = LeftB(adoRecordset.Fields(12), 20)
        Else
            StrR003019 = ""
        End If
        If Not IsNull(adoRecordset.Fields(13)) Then
            StrR003020 = adoRecordset.Fields(13)
        Else
            StrR003020 = ""
        End If
        
        DoEvents
        '若專利基本檔或服務業務基本檔已閉卷, 則不列印
        If DoPaAndSp2_1 <> 0 Then GoTo NextRecord
        
        s = 0
        If Len(Trim(Txt1(8))) <> 0 Then
            If (IIf(StrR003010 = "", False, StrR003010 >= GetNewFagent(Txt1(8))) Or IIf(StrR003011 = "", False, StrR003011 >= GetNewFagent(Txt1(8))) Or IIf(StrR003012 = "", False, StrR003012 >= GetNewFagent(Txt1(8))) Or IIf(StrR003013 = "", False, StrR003013 >= GetNewFagent(Txt1(8))) Or IIf(StrR003014 = "", False, StrR003014 >= GetNewFagent(Txt1(8)))) Then

                s = 1
            Else
                s = 0
            End If
        Else
            s = 1
        End If
        If s <> 0 Then
            If Len(Trim(Txt1(9))) <> 0 Then
                If (IIf(StrR003010 = "", False, StrR003010 <= GetNewFagent(Txt1(9))) Or IIf(StrR003011 = "", False, StrR003011 <= GetNewFagent(Txt1(9))) Or IIf(StrR003012 = "", False, StrR003012 <= GetNewFagent(Txt1(9))) Or IIf(StrR003013 = "", False, StrR003013 <= GetNewFagent(Txt1(9))) Or IIf(StrR003014 = "", False, StrR003014 <= GetNewFagent(Txt1(9)))) Then
                    s = 1
                Else
                    s = 0
                End If
            End If
        End If
        StrR002008 = ""
        DoEvents
        If s = 0 Then
            adoRecordset.Delete
        Else
            StrR003005 = StrR003005
            StrR003006 = GetPrjSalesNM(StrR003006)
            StrR003007 = GetPrjState4(StrR003001 + "-" + StrR003002 + "-" + StrR003003 + "-" + StrR003004, StrR003007)
            StrR003008 = GetPrjName1(StrR003008)
            StrR003009 = GetPrjName1(StrR003009)
            StrR003010 = GetPrjPeople1(StrR003010)
            StrR003011 = GetPrjPeople1(StrR003011)
            StrR003012 = GetPrjPeople1(StrR003012)
            StrR003013 = GetPrjPeople1(StrR003013)
            StrR003014 = GetPrjPeople1(StrR003014)
            If LenB(StrR003006) > 8 Then
                StrR003006 = LeftB(StrR003006, 8)
            End If
            If LenB(StrR003007) > 12 Then
                StrR003007 = LeftB(StrR003007, 12)
            End If
            If LenB(StrR003008) > 9 Then
                StrR003008 = LeftB(StrR003008, 8)
            End If
            If LenB(StrR003010) > 12 Then
                StrR003010 = LeftB(StrR003010, 12)
            End If
            If LenB(StrR003011) > 12 Then
                StrR003011 = LeftB(StrR003011, 12)
            End If
            If LenB(StrR003012) > 12 Then
                StrR003012 = LeftB(StrR003012, 12)
            End If
            If LenB(StrR003013) > 12 Then
                StrR003013 = LeftB(StrR003013, 12)
            End If
            If LenB(StrR003014) > 12 Then
                StrR003014 = LeftB(StrR003014, 12)
            End If
            If LenB(StrR003009) > 9 Then
                StrR003009 = LeftB(StrR003009, 8)
            End If
            If LenB(StrR003019) > 20 Then
                StrR003019 = LeftB(StrR003019, 20)
            End If
            strSql = "INSERT INTO R050302_2 VALUES ('" & ChgSQL(StrR003001) & "','" & ChgSQL(StrR003002) & "','" & ChgSQL(StrR003003) & "','" & ChgSQL(StrR003004) & "','" & ChgSQL(StrR003005) & "','" & ChgSQL(StrR003006) & "','" & ChgSQL(StrR003007) & "','" & ChgSQL(StrR003008) & "','" & ChgSQL(StrR003009) & "','" & ChgSQL(StrR003010) & "','" & ChgSQL(StrR003011) & "','" & ChgSQL(StrR003012) & "','" & ChgSQL(StrR003013) & "','" & ChgSQL(StrR003014) & "','" & ChgSQL(StrR003015) & "','" & ChgSQL(StrR003016) & "','" & ChgSQL(StrR003017) & "','" & ChgSQL(StrR003018) & "','" & ChgSQL(StrR003019) & "','" & strUserNum & "','" & StrR003020 & "') "
            cnnConnection.Execute strSql
        End If
        
NextRecord:

        adoRecordset.MoveNext
        DoEvents
    Loop
    If adoRecordset.RecordCount = 0 Then
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
Else
    ShowNoData
    Screen.MousePointer = vbDefault
    Exit Sub
End If
PriMenu2
Screen.MousePointer = vbDefault
End Sub

Sub StrMenu3()         '處理主程式
Dim strWhereSql As String

Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050302_3 WHERE ID='" & strUserNum & "' "
strWhereSql = ""

If Len(Txt1(6)) <> 0 Then
   strWhereSql = strWhereSql & " AND NP10='" & Txt1(6) & "' "
End If
If Len(Txt1(7)) <> 0 Then
   strWhereSql = strWhereSql & " AND NP07=" & Val(Txt1(7)) & " "
End If
If Len(Txt1(0)) <> 0 Then
    strWhereSql = strWhereSql & " AND NP02 IN (" & GetAddStr(Txt1(0)) & ") "
End If

'2006/4/6 MODIFY BY SONIA 管制智權人員只抓未收文資料
'2010/3/22 MODIFY BY SONIA 程序管制案件性質不印改以strNpSqlOfNoSalesDuty控制
'Modify By Sindy 2011/3/15 FCT,T,TF延展(102)和第二期(716)專用權須存在(TM17=Y)
'strSql = "SELECT NP01,NP02,NP03,NP04,NP05,NP10,NP08,NP09,NP07,NP10 FROM NEXTPROGRESS WHERE NP08>=" & Val(ChangeTStringToWString(txt1(1))) & " AND NP08<=" & Val(ChangeTStringToWString(txt1(2))) & " AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') " & strNpSqlOfNoSalesDuty
strSql = "SELECT NP01,NP02,NP03,NP04,NP05,NP10,NP08,NP09,NP07,NP10 FROM NEXTPROGRESS,trademark WHERE NP08>=" & Val(ChangeTStringToWString(Txt1(1))) & " AND NP08<=" & Val(ChangeTStringToWString(Txt1(2))) & " AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') and np02=tm01 and np03=tm02 and np04=tm03 and np05=tm04 and decode(np02||np07,'T716',tm17,'T102',tm17,'FCT716',tm17,'FCT102',tm17,'TF716',tm17,'TF102',tm17,'Y')='Y' " & strNpSqlOfNoSalesDuty & strWhereSql & _
             " union SELECT NP01,NP02,NP03,NP04,NP05,NP10,NP08,NP09,NP07,NP10 FROM NEXTPROGRESS,patent WHERE NP08>=" & Val(ChangeTStringToWString(Txt1(1))) & " AND NP08<=" & Val(ChangeTStringToWString(Txt1(2))) & " AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') and np02=pa01 and np03=pa02 and np04=pa03 and np05=pa04 " & strNpSqlOfNoSalesDuty & strWhereSql & _
             " union SELECT NP01,NP02,NP03,NP04,NP05,NP10,NP08,NP09,NP07,NP10 FROM NEXTPROGRESS,servicepractice WHERE NP08>=" & Val(ChangeTStringToWString(Txt1(1))) & " AND NP08<=" & Val(ChangeTStringToWString(Txt1(2))) & " AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') and np02=sp01 and np03=sp02 and np04=sp03 and np05=sp04 " & strNpSqlOfNoSalesDuty & strWhereSql & _
             " union SELECT NP01,NP02,NP03,NP04,NP05,NP10,NP08,NP09,NP07,NP10 FROM NEXTPROGRESS,lawcase WHERE NP08>=" & Val(ChangeTStringToWString(Txt1(1))) & " AND NP08<=" & Val(ChangeTStringToWString(Txt1(2))) & " AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') and np02=lc01 and np03=lc02 and np04=lc03 and np05=lc04 " & strNpSqlOfNoSalesDuty & strWhereSql & _
             " union SELECT NP01,NP02,NP03,NP04,NP05,NP10,NP08,NP09,NP07,NP10 FROM NEXTPROGRESS,hirecase WHERE NP08>=" & Val(ChangeTStringToWString(Txt1(1))) & " AND NP08<=" & Val(ChangeTStringToWString(Txt1(2))) & " AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') and np02=hc01 and np03=hc02 and np04=hc03 and np05=hc04 " & strNpSqlOfNoSalesDuty & strWhereSql
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
If adoRecordset.RecordCount <> 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        DoEvents
        StrR001001 = ""
        StrR001002 = ""
        StrR001003 = ""
        StrR001004 = ""
        StrR001005 = ""
        StrR001006 = ""
        StrR001007 = ""
        StrR001008 = ""
        StrR001009 = ""
        StrR001010 = ""
        StrR001011 = ""
        StrR001012 = ""
        StrR001013 = ""
        StrR001014 = ""
        StrR001015 = ""
        StrR001016 = ""
        StrR001017 = ""
        StrR001018 = ""
        StrR001019 = ""
        If Not IsNull(adoRecordset.Fields(0)) Then
            StrR001001 = adoRecordset.Fields(0)
        Else
            StrR001001 = ""
        End If
        If Not IsNull(adoRecordset.Fields(1)) Then
            StrR001002 = adoRecordset.Fields(1)
        Else
            StrR001002 = ""
        End If
        If Not IsNull(adoRecordset.Fields(2)) Then
            StrR001003 = adoRecordset.Fields(2)
        Else
            StrR001003 = ""
        End If
        If Not IsNull(adoRecordset.Fields(3)) Then
            StrR001004 = adoRecordset.Fields(3)
        Else
            StrR001004 = ""
        End If
        If Not IsNull(adoRecordset.Fields(4)) Then
            StrR001005 = adoRecordset.Fields(4)
        Else
            StrR001005 = ""
        End If
        '智權人員代號
        If Not IsNull(adoRecordset.Fields(5)) Then
            StrR001006 = adoRecordset.Fields(5)
        Else
            StrR001006 = ""
        End If
        If Not IsNull(adoRecordset.Fields(6)) Then
            StrR001017 = str(adoRecordset.Fields(6))
        Else
            StrR001017 = ""
        End If
        If Not IsNull(adoRecordset.Fields(7)) Then
            StrR001019 = ChangeTStringToTDateString(ChangeWStringToTString(adoRecordset.Fields(7)))
        Else
            StrR001019 = ""
        End If
        DoCaseProgress
        DoEvents
        
        '若專利基本檔或服務業務基本檔已閉卷, 則不列印
        If DoPaAndSp_1 <> 0 Then GoTo NextRecord
        
        StrR001008 = CheckStr(adoRecordset.Fields(8))
        s = 0
        DoStaff
        If Len(Trim(Txt1(8))) <> 0 Then
            If (IIf(StrR001010 = "", False, StrR001010 >= GetNewFagent(Txt1(8))) Or IIf(StrR001011 = "", False, StrR001011 >= GetNewFagent(Txt1(8))) Or IIf(StrR001012 = "", False, StrR001012 >= GetNewFagent(Txt1(8))) Or IIf(StrR001013 = "", False, StrR001013 >= GetNewFagent(Txt1(8))) Or IIf(StrR001014 = "", False, StrR001014 >= GetNewFagent(Txt1(8)))) Then
                s = 1
            Else
                s = 0
            End If
        Else
            s = 1
        End If
        If s <> 0 Then
            If Len(Trim(Txt1(9))) <> 0 Then
                If (IIf(StrR001010 = "", False, StrR001010 <= GetNewFagent(Txt1(9))) Or IIf(StrR001011 = "", False, StrR001011 <= GetNewFagent(Txt1(9))) Or IIf(StrR001012 = "", False, StrR001012 <= GetNewFagent(Txt1(9))) Or IIf(StrR001013 = "", False, StrR001013 <= GetNewFagent(Txt1(9))) Or IIf(StrR001014 = "", False, StrR001014 <= GetNewFagent(Txt1(9)))) Then
                    s = 1
                Else
                    s = 0
                End If
            Else
                s = 1
            End If
        End If
        DoEvents
        If s = 0 Then
            adoRecordset.Delete
        Else
            StrR001007 = GetPrjSalesNM(StrR001007)
            StrR001008 = GetPrjState4(StrR001002 + "-" + StrR001003 + "-" + StrR001004 + "-" + StrR001005, StrR001008)
            StrR001009 = GetPrjName1(StrR001009)
            StrR001010 = GetPrjPeople1(StrR001010)
            StrR001011 = GetPrjPeople1(StrR001011)
            StrR001012 = GetPrjPeople1(StrR001012)
            StrR001013 = GetPrjPeople1(StrR001013)
            StrR001014 = GetPrjPeople1(StrR001014)
            StrR001015 = GetPrjName1(StrR001015)
            StrR001016 = StrR001016
            If Val(StrR001017) < Val(GetTodayDate) Then
                StrR001017 = "*" + ChangeTStringToTDateString(ChangeWStringToTString(StrR001017))
            Else
                If Val(StrR001017) = Val(GetTodayDate) Then
                    StrR001017 = "V" + ChangeTStringToTDateString(ChangeWStringToTString(StrR001017))
                Else
                    If UCase(Left(StrR001001, 1)) = "C" And StrR001018 = "Y" Then
                        StrR001017 = "#" + ChangeTStringToTDateString(ChangeWStringToTString(StrR001017))
                    Else
                        StrR001017 = ChangeTStringToTDateString(ChangeWStringToTString(StrR001017))
                    End If
                End If
            End If
            If LenB(StrR001007) > 8 Then
                StrR001007 = LeftB(StrR001007, 8)
            End If
            If LenB(StrR001008) > 12 Then
                StrR001008 = LeftB(StrR001008, 12)
            End If
            If LenB(StrR001009) > 9 Then
                StrR001009 = LeftB(StrR001009, 8)
            End If
            If LenB(StrR001010) > 12 Then
                StrR001010 = LeftB(StrR001010, 12)
            End If
            If LenB(StrR001011) > 12 Then
                StrR001011 = LeftB(StrR001011, 12)
            End If
            If LenB(StrR001012) > 12 Then
                StrR001012 = LeftB(StrR001012, 12)
            End If
            If LenB(StrR001013) > 12 Then
                StrR001013 = LeftB(StrR001013, 12)
            End If
            If LenB(StrR001014) > 12 Then
                StrR001014 = LeftB(StrR001014, 12)
            End If
            If LenB(StrR001015) > 9 Then
                StrR001015 = LeftB(StrR001015, 8)
            End If
            If LenB(StrR001016) > 20 Then
                StrR001016 = LeftB(StrR001016, 20)
            End If
            strSql = "INSERT INTO R050302_3 VALUES ('" & ChgSQL(StrR001001) & "','" & ChgSQL(StrR001002) & "','" & ChgSQL(StrR001003) & "','" & ChgSQL(StrR001004) & "','" & ChgSQL(StrR001005) & "','" & ChgSQL(StrR001006) & "','" & ChgSQL(StrR001007) & "','" & ChgSQL(StrR001008) & "','" & ChgSQL(StrR001009) & "','" & ChgSQL(StrR001010) & "','" & ChgSQL(StrR001011) & "','" & ChgSQL(StrR001012) & "','" & ChgSQL(StrR001013) & "','" & ChgSQL(StrR001014) & "','" & ChgSQL(StrR001015) & "','" & ChgSQL(StrR001016) & "','" & ChgSQL(StrR001017) & "','" & ChgSQL(StrR001018) & "','" & ChgSQL(StrR001019) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
        End If

NextRecord:

        adoRecordset.MoveNext
        DoEvents
    Loop
    If adoRecordset.RecordCount = 0 Then
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
Else
    ShowNoData
    Screen.MousePointer = vbDefault
    Exit Sub
End If
PriMenu3
Screen.MousePointer = vbDefault
End Sub

Sub DoStaff()
strSql = "SELECT ST15 FROM STAFF WHERE ST01='" & ChgSQL(StrR001006) & "'"
CheckOC2
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 Then
    If Not IsNull(adoRecordset1.Fields(0)) Then
        StrR001016 = adoRecordset1.Fields(0)
    Else
        StrR001016 = ""
    End If
Else
    StrR001016 = ""
End If
CheckOC2
End Sub

Function DoPaAndSp2_1() As Integer

DoPaAndSp2_1 = -1
'2006/4/7 MODIFY BY SONIA 只抓申請人1
strSql = "SELECT PA26,'','','','',PA75 FROM PATENT WHERE PA01='" & ChgSQL(StrR003001) & "' AND PA02='" & ChgSQL(StrR003002) & "' AND PA03='" & ChgSQL(StrR003003) & "' AND PA04='" & ChgSQL(StrR003004) & "' AND (PA57<>'Y' Or PA57 Is Null) "
strSql = strSql + " union all select TM23,'','','','',TM29 FROM Trademark WHERE TM01='" & ChgSQL(StrR003001) & "' AND TM02='" & ChgSQL(StrR003002) & "' AND TM03='" & ChgSQL(StrR003003) & "' AND TM04='" & ChgSQL(StrR003004) & "' AND (TM29<>'Y' or TM29 Is Null) "
strSql = strSql + " union all select LC11,'','','','',LC08 FROM Lawcase WHERE LC01='" & ChgSQL(StrR003001) & "' AND LC02='" & ChgSQL(StrR003002) & "' AND LC03='" & ChgSQL(StrR003003) & "' AND LC04='" & ChgSQL(StrR003004) & "' AND (LC08<>'Y' Or LC08 Is Null) "
strSql = strSql + " union all select HC05,'','','','',HC09 FROM Hirecase WHERE HC01='" & ChgSQL(StrR003001) & "' AND HC02='" & ChgSQL(StrR003002) & "' AND HC03='" & ChgSQL(StrR003003) & "' AND HC04='" & ChgSQL(StrR003004) & "' AND (HC09<>'Y' Or HC09 Is Null) "
strSql = strSql + " union all select SP08,'','','','',SP26 FROM SERVICEPRACTICE WHERE SP01='" & ChgSQL(StrR003001) & "' AND SP02='" & ChgSQL(StrR003002) & "' AND SP03='" & ChgSQL(StrR003003) & "' AND SP04='" & ChgSQL(StrR003004) & "' AND (SP15<>'Y' Or SP15 Is Null) "
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 Then
   DoPaAndSp2_1 = 0
    If Not IsNull(adoRecordset1.Fields(0)) Then
        StrR003010 = adoRecordset1.Fields(0)
    Else
        StrR003010 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(1)) Then
        StrR003011 = adoRecordset1.Fields(1)
    Else
        StrR003011 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(2)) Then
        StrR003012 = adoRecordset1.Fields(2)
    Else
        StrR003012 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(3)) Then
        StrR003013 = adoRecordset1.Fields(3)
    Else
        StrR003013 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(4)) Then
        StrR003014 = adoRecordset1.Fields(4)
    Else
        StrR003014 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(5)) Then
        StrR003009 = adoRecordset1.Fields(5)
    Else
        StrR003009 = ""
    End If
Else
    StrR001010 = ""
    StrR001011 = ""
    StrR001012 = ""
    StrR001013 = ""
    StrR001014 = ""
    StrR001009 = ""
End If
CheckOC2

End Function

Function DoPaAndSp1_1() As Integer

DoPaAndSp1_1 = -1
'2006/4/7 MODIFY BY SONIA 只抓申請人1
strSql = "SELECT PA26,'','','','' FROM PATENT WHERE PA01='" & ChgSQL(StrR002001) & "' AND PA02='" & ChgSQL(StrR002002) & "' AND PA03='" & ChgSQL(StrR002003) & "' AND PA04='" & ChgSQL(StrR002004) & "' AND (PA57<>'Y' Or PA57 Is Null) "
strSql = strSql + " union all select TM23,'','','','' FROM Trademark WHERE TM01='" & ChgSQL(StrR002001) & "' AND TM02='" & ChgSQL(StrR002002) & "' AND TM03='" & ChgSQL(StrR002003) & "' AND TM04='" & ChgSQL(StrR002004) & "' AND (TM29<>'Y' Or TM29 Is Null) "
strSql = strSql + " union all select LC11,'','','','' FROM Lawcase WHERE LC01='" & ChgSQL(StrR002001) & "' AND LC02='" & ChgSQL(StrR002002) & "' AND LC03='" & ChgSQL(StrR002003) & "' AND LC04='" & ChgSQL(StrR002004) & "' AND (LC08<>'Y' Or LC08 Is Null) "
strSql = strSql + " union all select HC05,'','','','' FROM Hirecase WHERE HC01='" & ChgSQL(StrR002001) & "' AND HC02='" & ChgSQL(StrR002002) & "' AND HC03='" & ChgSQL(StrR002003) & "' AND HC04='" & ChgSQL(StrR002004) & "' AND (HC09<>'Y' Or HC09 Is Null) "
strSql = strSql + " union all select SP08,'','','','' FROM SERVICEPRACTICE WHERE SP01='" & ChgSQL(StrR002001) & "' AND SP02='" & ChgSQL(StrR002002) & "' AND SP03='" & ChgSQL(StrR002003) & "' AND SP04='" & ChgSQL(StrR002004) & "' AND (SP15<>'Y' Or SP15 Is Null) "
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 Then
   DoPaAndSp1_1 = 0
    If Not IsNull(adoRecordset1.Fields(0)) Then
        StrR002009 = adoRecordset1.Fields(0)
    Else
        StrR002009 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(1)) Then
        StrR002010 = adoRecordset1.Fields(1)
    Else
        StrR002010 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(2)) Then
        StrR002011 = adoRecordset1.Fields(2)
    Else
        StrR002011 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(3)) Then
        StrR002012 = adoRecordset1.Fields(3)
    Else
        StrR002012 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(4)) Then
        StrR002013 = adoRecordset1.Fields(4)
    Else
        StrR002013 = ""
    End If
Else
    StrR001009 = ""
    StrR001010 = ""
    StrR001011 = ""
    StrR001012 = ""
    StrR001013 = ""
End If
CheckOC2

End Function

Function DoPaAndSp_1() As Integer

DoPaAndSp_1 = -1
'2006/4/7 MODIFY BY SONIA 只抓申請人1
strSql = "SELECT PA26,'','','','',PA75 FROM PATENT WHERE PA01='" & ChgSQL(StrR001002) & "' AND PA02='" & ChgSQL(StrR001003) & "' AND PA03='" & ChgSQL(StrR001004) & "' AND PA04='" & ChgSQL(StrR001005) & "' AND (PA57<>'Y' Or PA57 Is Null) "
strSql = strSql + " union all select TM23,'','','','',TM29 FROM Trademark WHERE TM01='" & ChgSQL(StrR001002) & "' AND TM02='" & ChgSQL(StrR001003) & "' AND TM03='" & ChgSQL(StrR001004) & "' AND TM04='" & ChgSQL(StrR001005) & "' AND (TM29<>'Y' Or TM29 Is Null) "
strSql = strSql + " union all select LC11,'','','','',LC08 FROM Lawcase WHERE LC01='" & ChgSQL(StrR001002) & "' AND LC02='" & ChgSQL(StrR001003) & "' AND LC03='" & ChgSQL(StrR001004) & "' AND LC04='" & ChgSQL(StrR001005) & "' AND (LC08<>'Y' Or LC08 Is Null) "
strSql = strSql + " union all select HC05,'','','','',HC09 FROM Hirecase WHERE HC01='" & ChgSQL(StrR001002) & "' AND HC02='" & ChgSQL(StrR001003) & "' AND HC03='" & ChgSQL(StrR001004) & "' AND HC04='" & ChgSQL(StrR001005) & "' AND (HC09<>'Y' Or HC09 Is Null) "
strSql = strSql + " union all select SP08,'','','','',SP26 FROM SERVICEPRACTICE WHERE SP01='" & ChgSQL(StrR001002) & "' AND SP02='" & ChgSQL(StrR001003) & "' AND SP03='" & ChgSQL(StrR001004) & "' AND SP04='" & ChgSQL(StrR001005) & "' AND (SP15<>'Y' Or SP15 Is Null) "
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 Then
    DoPaAndSp_1 = 0
    If Not IsNull(adoRecordset1.Fields(0)) Then
        StrR001010 = adoRecordset1.Fields(0)
    Else
        StrR001010 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(1)) Then
        StrR001011 = adoRecordset1.Fields(1)
    Else
        StrR001011 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(2)) Then
        StrR001012 = adoRecordset1.Fields(2)
    Else
        StrR001012 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(3)) Then
        StrR001013 = adoRecordset1.Fields(3)
    Else
        StrR001013 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(4)) Then
        StrR001014 = adoRecordset1.Fields(4)
    Else
        StrR001014 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(5)) Then
        StrR001015 = adoRecordset1.Fields(5)
    Else
        StrR001015 = ""
    End If
Else
    StrR001010 = ""
    StrR001011 = ""
    StrR001012 = ""
    StrR001013 = ""
    StrR001014 = ""
    StrR001015 = ""
End If
CheckOC2
End Function

Sub DoCaseProgress()
strSql = "SELECT CP14,CP10,CP44,CP27 FROM CASEPROGRESS WHERE CP09='" & ChgSQL(StrR001001) & "'"
CheckOC2
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 Then
    If Not IsNull(adoRecordset1.Fields(0)) Then
        StrR001007 = adoRecordset1.Fields(0)
    Else
        StrR001007 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(1)) Then
        StrR001008 = adoRecordset1.Fields(1)
    Else
        StrR001008 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(2)) Then
        StrR001009 = adoRecordset1.Fields(2)
    Else
        StrR001009 = ""
    End If
    If Not IsNull(adoRecordset1.Fields(3)) Then
        StrR001018 = "N"
    Else
        StrR001018 = "Y"
    End If
Else
    StrR001018 = ""
    StrR001007 = ""
    StrR001008 = ""
    StrR001009 = ""
End If
CheckOC2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm12040143 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
      If Option1(0).Value = True Then
         Txt1(1).Enabled = True
         Txt1(2).Enabled = True
         Txt1(3).Enabled = False
         Txt1(4).Enabled = False
         Option1(1).Value = False
      End If
Case 1
      If Option1(1).Value = True Then
         Txt1(3).Enabled = True
         Txt1(4).Enabled = True
         Txt1(1).Enabled = False
         Txt1(2).Enabled = False
         Option1(0).Value = False
      End If
Case Else
End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    Select Case Index
    Case 5 '管制對象
        If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     If Len(Txt1(0)) <> 0 Then
        STRSTRING = ""
        StrTempP = Split(Replace(Txt1(0), ",,", ""), ",")
        StrTempP2 = Split(Replace(GetSystemKindByNick, ",,", ""), ",")
        For i = 0 To UBound(StrTempP)
            s = 0
            For j = 0 To UBound(StrTempP2)
                If StrTempP(i) = StrTempP2(j) Then
                    s = 1
                End If
            Next j
            If s = 0 Then
                STRSTRING = STRSTRING + StrTempP(i) + " "
            End If
        Next i
        If Len(STRSTRING) <> 0 Then
            s = MsgBox(STRSTRING + " 不是 " + strUserNum + " 的權限!!", , "警告!!!")
            Txt1(0).SetFocus
            Exit Sub
        End If
        If Len(Txt1(7)) <> 0 Then
            LBL1(1).Caption = GetPrjState4(StrTempP(0) + "---", Txt1(7))
        End If
      End If
Case 2, 4
   If blnClkSure = False Then
      If RunNick(Txt1(Index - 1), Txt1(Index)) Then
         Txt1(Index - 1).SetFocus
      End If
   Else
      blnClkSure = False
   End If
Case 9
   If blnClkSure = False Then
      If Len(Txt1(Index - 1)) <> 0 Then
         If Left(Txt1(Index - 1), 6) <> Left(Txt1(Index), 6) Then
             s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
             Txt1(Index - 1).SetFocus
             Exit Sub
         End If
      End If
      If RunNick(Txt1(Index - 1), Txt1(Index)) Then
         Txt1(Index - 1).SetFocus
      End If
   Else
      blnClkSure = False
   End If
End Select

End Sub

Sub PriMenu1()                     '印表主程式
Dim StrR050302_1(0 To 14) As String
Dim strSaleName As String '智權人員名稱
strSaleName = ""
'日期排序不能用符號
strSql = "SELECT r002001,r002002,r002003,r002004,r002005,st02,r002007,r002008,r002009,r002010,r002011,r002012,r002013,r002014,r002015,r002006 FROM R050302_1,staff WHERE r002006=st01(+) and ID='" & strUserNum & "' ORDER BY R002006,decode(substr(R002014,1,1),'*',substr(r002014,2,10),'#',substr(r002014,2,10),'V',substr(r002014,2,10),r002014),R002001,R002002,R002003,R002004 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount = 0 Then
    Exit Sub
End If
iLine = 1
Page = 1
If Not IsNull(adoRecordset.Fields(5)) Then
    TmpArea = adoRecordset.Fields(5)
Else
    TmpArea = ""
End If
PriTiTle1 TmpArea, 1
iPrint = 2700

With adoRecordset
    .MoveFirst
    Do While .EOF = False
        For j = 0 To 14
        If Not IsNull(.Fields(j)) Then
            StrR050302_1(j) = .Fields(j)
        Else
            StrR050302_1(j) = ""
        End If
        Next j
        Printer.CurrentX = PLeft1(0)
        Printer.CurrentY = iPrint
        Printer.Print StrR050302_1(13)
        Printer.CurrentX = PLeft1(1)
        Printer.CurrentY = iPrint
        Printer.Print StrR050302_1(14)
        Printer.CurrentX = PLeft1(2)
        Printer.CurrentY = iPrint
        Printer.Print StrR050302_1(0) + "-" + StrR050302_1(1) + "-" + StrR050302_1(2) + "-" + StrR050302_1(3)
        Printer.CurrentX = PLeft1(3)
        Printer.CurrentY = iPrint
        Printer.Print LeftB(Format(GetPrjName(StrR050302_1(0) + "-" + StrR050302_1(1) + "-" + StrR050302_1(2) + "-" + StrR050302_1(3)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"), 34)
        Printer.CurrentX = PLeft1(4)
        Printer.CurrentY = iPrint
        Printer.Print LeftB(Format(GetPrjNation(StrR050302_1(0) + "-" + StrR050302_1(1) + "-" + StrR050302_1(2) + "-" + StrR050302_1(3)), "!@@@@@@@@@@@@"), 12)
        Printer.CurrentX = PLeft1(5)
        Printer.CurrentY = iPrint
        Printer.Print LeftB(Format(StrR050302_1(6), "!@@@@@@@@@@@@"), 12)
        Printer.CurrentX = PLeft1(7)
        Printer.CurrentY = iPrint
         
         '列印智權人員姓名(與上筆相同不列印)
         If strSaleName <> StrR050302_1(4) Then
            Printer.Print LeftB(Format(StrR050302_1(4), "!@@@@@@@@"), 8)
            strSaleName = StrR050302_1(4)
         Else
            Printer.Print ""
         End If
         
        For j = 1 To 5
            If Len(Trim(StrR050302_1(7 + j))) <> 0 Then
                Printer.CurrentX = PLeft1(6)
                Printer.CurrentY = iPrint
                Printer.Print LeftB(Format(StrR050302_1(7 + j), "!@@@@@@@@@@@@"), 12)
                If Not .EOF Then
                    St = StrR050302_1(5)
                Else
                    St = ""
                End If
                If (iLine Mod 25 = 0) Then
                    PriTiEnd1
                    Printer.NewPage
                    Page = Page + 1
                    PriTiTle1 St, str(Page)
                    iPrint = 2400
                    iLine = 0
                End If
                If j >= 1 And j <= 4 Then
                    If Len(Trim(StrR050302_1(8 + j))) <> 0 Then
                        iLine = iLine + 1
                        iPrint = iPrint + 300
                    End If
                End If
            End If
        Next j
        If Len(Trim(StrR050302_1(5))) <> 0 Then
            TmpArea = StrR050302_1(5)
        Else
            TmpArea = ""
        End If
        .MoveNext
        If .EOF = False Then
            If Not IsNull(.Fields(5)) Then
                St = .Fields(5)
            Else
                St = ""
            End If
            If ((iLine Mod 25 = 0) Or (TmpArea <> St)) And (iLine <> 0) Then
                PriTiEnd1
                Printer.NewPage
                Page = Page + 1
                PriTiTle1 St, str(Page)
                iPrint = 2400
                iLine = 0
            End If
            iLine = iLine + 1
            iPrint = iPrint + 300
        End If
    Loop
End With
PriTiEnd1
Printer.EndDoc
ShowPrintOk
CheckOC

End Sub

Sub PriMenu2()                 '印表主程式
Dim StrR050302_2(0 To 18) As String
Dim strDate As String '法定期限
Dim strServerDate As String '系統日期
Dim strSaleName As String '智權人員名稱
strSaleName = ""
strSql = "SELECT r003001,r003002,r003003,r003004,st02,r003006,r003007,r003008,r003009,r003010,r003011,r003012,r003013,r003014,r003015,r003016,r003017,r003018,r003019,r003005,r003020 FROM staff,R050302_2 WHERE r003005=st01(+) and ID='" & strUserNum & "' ORDER BY R003016,R003006,R003001,R003002,R003003,R003004 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount = 0 Then
    Exit Sub
End If
iLine = 1
Page = 1
PriTiTle2 1
iPrint = 2700
strServerDate = ServerDate
With adoRecordset
    .MoveFirst
    Do While .EOF = False
        For j = 0 To 18
        If Not IsNull(.Fields(j)) Then
            StrR050302_2(j) = .Fields(j)
        Else
            StrR050302_2(j) = ""
        End If
        Next j
         '列印法定期限
        Printer.CurrentX = Pleft2(0)
        Printer.CurrentY = iPrint
        'Add By Cheng 2002/03/15
        strDate = Val(Replace(StrR050302_2(15), "/", "")) + 19110000
        '若法定日期小於系統日期, 則在法定日期前加"*"號
        '2010/9/15 MODIFY BY SONIA
        'If strDate < strServerDate Then
        If Val(strDate) < Val(strServerDate) Then
           strDate = "*"
        '若法定日期等於系統日期, 則在法定日期前加"V"號
        ElseIf strDate = strServerDate Then
           strDate = "V"
        Else
           '發文日為NULL時, 在法定日期前加"#"
           If .Fields(20).Value = "1" Then
              strDate = "#"
           Else
              strDate = ""
           End If
        End If
        Printer.Print strDate & StrR050302_2(15)
        
        Printer.CurrentX = Pleft2(1)
        Printer.CurrentY = iPrint
        Printer.Print StrR050302_2(14)
        '列印承辦人
        Printer.CurrentX = Pleft2(2)
        Printer.CurrentY = iPrint
        If .Fields(20).Value = "1" Then
           Printer.Print "*" & StrToStr(StrR050302_2(5), 4)
        Else
           Printer.Print StrToStr(StrR050302_2(5), 4)
        End If
        
        Printer.CurrentX = Pleft2(3)
        Printer.CurrentY = iPrint
        Printer.Print StrR050302_2(0) + "-" + StrR050302_2(1) + "-" + StrR050302_2(2) + "-" + StrR050302_2(3)
        Printer.CurrentX = Pleft2(4)
        Printer.CurrentY = iPrint
        Printer.Print LeftB(Format(GetPrjName(StrR050302_2(0) + "-" + StrR050302_2(1) + "-" + StrR050302_2(2) + "-" + StrR050302_2(3)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"), 24)
        Printer.CurrentX = Pleft2(5)
        Printer.CurrentY = iPrint
        Printer.Print StrR050302_2(16)
        Printer.CurrentX = Pleft2(6)
        Printer.CurrentY = iPrint
        Printer.Print StrR050302_2(17)
        Printer.CurrentX = Pleft2(8)
        Printer.CurrentY = iPrint
        Printer.Print StrToStr(StrR050302_2(6), 4)
        Printer.CurrentX = Pleft2(9)
        Printer.CurrentY = iPrint
         '不論智權人員是否與上筆相同皆列印出來
          If .Fields(20).Value = "2" Then
              Printer.Print "*" & LeftB(Format(StrR050302_2(4), "!@@@@@@@@"), 8)
          Else
              Printer.Print LeftB(Format(StrR050302_2(4), "!@@@@@@@@"), 8)
          End If
        Printer.CurrentX = Pleft2(10)
        Printer.CurrentY = iPrint
        Printer.Print LeftB(Format(StrR050302_2(18), "!@@@@@@@@@@@@@@@@@@@@"), 18)
        For j = 1 To 5
            If Len(Trim(StrR050302_2(8 + j))) <> 0 Then
                Printer.CurrentX = Pleft2(7)
                Printer.CurrentY = iPrint
                Printer.Print StrToStr(StrR050302_2(8 + j), 3)
                If (iLine Mod 25 = 0) Then
                    PriTiEnd2
                    Printer.NewPage
                    Page = Page + 1
                    PriTiTle2 str(Page)
                    iPrint = 2400
                    iLine = 0
                End If
                If j >= 1 And j <= 4 Then
                    If Len(Trim(StrR050302_2(10 + j))) <> 0 Then
                        iLine = iLine + 1
                        iPrint = iPrint + 300
                    End If
                End If
            End If
        Next j
        .MoveNext
        If .EOF = False Then
            If (iLine Mod 25 = 0) And (iLine <> 0) Then
                PriTiEnd2
                Printer.NewPage
                Page = Page + 1
                PriTiTle2 str(Page)
                iPrint = 2400
                iLine = 0
            End If
            iLine = iLine + 1
            iPrint = iPrint + 300
        End If
    Loop
End With
PriTiEnd2
Printer.EndDoc
ShowPrintOk
CheckOC

End Sub

Sub PriMenu3()                 '印表主程式
Dim StrR050302_3(0 To 18) As String
Dim strSaleName As String '智權人員名稱
strSaleName = ""
'日期排序不能用符號
strSql = "SELECT r001001,r001002,r001003,r001004,r001005,st02,r001007,r001008,r001009,r001010,r001011,r001012,r001013,r001014,r001015,a0902,r001017,r001018,r001019,r001016,r001006 FROM staff,R050302_3,acc090 WHERE r001016=a0901 and ID='" & strUserNum & "' and r001006=st01(+) ORDER BY R001016,R001006,decode(substr(R001017,1,1),'*',substr(R001017,2,10),'#',substr(R001017,2,10),'V',substr(R001017,2,10),r001017),R001002,R001003,R001004,R001005 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount = 0 Then
    Exit Sub
End If
iLine = 1
Page = 1
If Not IsNull(adoRecordset.Fields(15)) Then
    TmpArea = adoRecordset.Fields(15)
Else
    TmpArea = ""
End If
PriTiTle3 TmpArea, 1
iPrint = 2700

With adoRecordset
    .MoveFirst
    Do While .EOF = False
        For j = 0 To 18
        If Not IsNull(.Fields(j)) Then
            StrR050302_3(j) = .Fields(j)
        Else
            StrR050302_3(j) = ""
        End If
        Next j
        Printer.CurrentX = PLeft3(0)
        Printer.CurrentY = iPrint
        '列印智權人員名稱
        If strSaleName <> StrR050302_3(5) Then
           Printer.Print LeftB(Format(StrR050302_3(5), "!@@@@@@@@"), 8)
           strSaleName = StrR050302_3(5)
        Else
           Printer.Print ""
        End If
        
        Printer.CurrentX = PLeft3(1)
        Printer.CurrentY = iPrint
        Printer.Print StrR050302_3(16)
        Printer.CurrentX = PLeft3(2)
        Printer.CurrentY = iPrint
        Printer.Print StrR050302_3(18)
        Printer.CurrentX = PLeft3(3)
        Printer.CurrentY = iPrint
        Printer.Print StrR050302_3(1) + "-" + StrR050302_3(2) + "-" + StrR050302_3(3) + "-" + StrR050302_3(4)
        Printer.CurrentX = PLeft3(4)
        Printer.CurrentY = iPrint
        Printer.Print LeftB(Format(GetPrjName(StrR050302_3(1) + "-" + StrR050302_3(2) + "-" + StrR050302_3(3) + "-" + StrR050302_3(4)), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"), 34)
        Printer.CurrentX = PLeft3(5)
        Printer.CurrentY = iPrint
        Printer.Print LeftB(Format(GetPrjNation(StrR050302_3(1) + "-" + StrR050302_3(2) + "-" + StrR050302_3(3) + "-" + StrR050302_3(4)), "!@@@@@@@@@@@@"), 12)
        Printer.CurrentX = PLeft3(6)
        Printer.CurrentY = iPrint
        Printer.Print LeftB(Format(StrR050302_3(7), "!@@@@@@@@@@@@"), 12)
        Printer.CurrentX = PLeft3(8)
        Printer.CurrentY = iPrint
        Printer.Print LeftB(Format(StrR050302_3(6), "!@@@@@@@@@@@@"), 12)
        For j = 1 To 5
            If Len(Trim(StrR050302_3(8 + j))) <> 0 Then
                Printer.CurrentX = PLeft3(7)
                Printer.CurrentY = iPrint
                Printer.Print LeftB(Format(StrR050302_3(8 + j), "!@@@@@@@@@@@@"), 12)
                If Not .EOF Then
                    St = StrR050302_3(15)
                Else
                    St = ""
                End If
                If (iLine Mod 25 = 0) Then
                    PriTiEnd3
                    Printer.NewPage
                    Page = Page + 1
                    PriTiTle3 St, str(Page)
                    iPrint = 2400
                    iLine = 0
                End If
                If j >= 1 And j <= 4 Then
                    If Len(Trim(StrR050302_3(9 + j))) <> 0 Then
                        iLine = iLine + 1
                        iPrint = iPrint + 300
                    End If
                End If
            End If
        Next j
        If Len(Trim(StrR050302_3(15))) <> 0 Then
            TmpArea = StrR050302_3(15)
        Else
            TmpArea = ""
        End If
        .MoveNext
        If .EOF = False Then
            If Not IsNull(.Fields(15)) Then
                St = .Fields(15)
            Else
                St = ""
            End If
            If ((iLine Mod 25 = 0) Or (TmpArea <> St)) And (iLine <> 0) Then
                PriTiEnd3
                Printer.NewPage
                Page = Page + 1
                PriTiTle3 St, str(Page)
                iPrint = 2400
                iLine = 0
            End If
            iLine = iLine + 1
            iPrint = iPrint + 300
        End If
    Loop
End With
PriTiEnd3
Printer.EndDoc
ShowPrintOk
CheckOC
End Sub

Sub PriTiTle3(ByRef Area As String, ByRef Page As String)             '印表頭
GetPrintLeft3
k = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = i
Printer.Print "逾期未處理案件明細表"
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = 6000
Printer.CurrentY = k + 500
Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(Txt1(1)) & " ", "@@@@@@@@@") & "-" & ChangeTStringToTDateString(Txt1(2))
Printer.Font.Bold = False
Printer.CurrentX = 500
Printer.CurrentY = k + 800
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = k + 800
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.CurrentX = 500
Printer.CurrentY = k + 1100
Printer.Print "業務區：" & Area
Printer.CurrentX = 13000
Printer.CurrentY = k + 1100
Printer.Print "頁次：" & Page
Printer.CurrentX = 500
Printer.CurrentY = k + 1400
Printer.Print String(200, "-")

Printer.Font.Underline = True
Printer.CurrentX = PLeft3(0)
Printer.CurrentY = k + 1700
Printer.Print "智權人員"
Printer.CurrentX = PLeft3(1)
Printer.CurrentY = k + 1700
Printer.Print "本所期限"
Printer.CurrentX = PLeft3(2)
Printer.CurrentY = k + 1700
Printer.Print "法定期限"
Printer.CurrentX = PLeft3(3)
Printer.CurrentY = k + 1700
Printer.Print "本所案號"
Printer.CurrentX = PLeft3(4)
Printer.CurrentY = k + 1700
Printer.Print "案件名稱"
Printer.CurrentX = PLeft3(5)
Printer.CurrentY = k + 1700
Printer.Print "申請國家"
Printer.CurrentX = PLeft3(6)
Printer.CurrentY = k + 1700
Printer.Print "案件性質"
Printer.CurrentX = PLeft3(7)
Printer.CurrentY = k + 1700
Printer.Print "申請人"
Printer.CurrentX = PLeft3(8)
Printer.CurrentY = k + 1700
Printer.Print "承辦人"
Printer.Font.Underline = False
Printer.CurrentX = 500
Printer.CurrentY = k + 2000
Printer.Print String(200, "-")
End Sub

Sub PriTiTle2(ByRef Page As String)              '印表頭
GetPrintLeft2
k = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = i
Printer.Print "逾期未處理案件明細表"
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = 6000
Printer.CurrentY = k + 500
Printer.Print "法定期限：" & Format(ChangeTStringToTDateString(Txt1(3)) & " ", "@@@@@@@@@") & "-" & ChangeTStringToTDateString(Txt1(4))
Printer.Font.Bold = False
Printer.CurrentX = 0
Printer.CurrentY = k + 800
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = k + 800
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.CurrentX = 13000
Printer.CurrentY = k + 1100
Printer.Print "頁次：" & Page
Printer.CurrentX = 0
Printer.CurrentY = k + 1400
Printer.Print String(200, "-")
Printer.Font.Underline = True
Printer.CurrentX = Pleft2(0)
Printer.CurrentY = k + 1700
Printer.Print "法定期限"
Printer.CurrentX = Pleft2(1)
Printer.CurrentY = k + 1700
Printer.Print "本所期限"
Printer.CurrentX = Pleft2(2)
Printer.CurrentY = k + 1700
Printer.Print "承辦人"
Printer.CurrentX = Pleft2(3)
Printer.CurrentY = k + 1700
Printer.Print "本所案號"
Printer.CurrentX = Pleft2(4)
Printer.CurrentY = k + 1700
Printer.Print "案件名稱"
Printer.CurrentX = Pleft2(5)
Printer.CurrentY = k + 1700
Printer.Print "收文日"
Printer.CurrentX = Pleft2(6)
Printer.CurrentY = k + 1700
Printer.Print "收文種類"
Printer.CurrentX = Pleft2(7)
Printer.CurrentY = k + 1700
Printer.Print "申請人"
Printer.CurrentX = Pleft2(8)
Printer.CurrentY = k + 1700
Printer.Print "案件性質"
Printer.CurrentX = Pleft2(9)
Printer.CurrentY = k + 1700
Printer.Print "智權人員"
Printer.CurrentX = Pleft2(10)
Printer.CurrentY = k + 1700
Printer.Print "備註"
Printer.Font.Underline = False
Printer.CurrentX = 0
Printer.CurrentY = k + 2000
Printer.Print String(200, "-")

End Sub

Sub PriTiTle1(ByRef Area As String, ByRef Page As String)             '印表頭
GetPrintLeft1
k = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = i
Printer.Print "逾期未處理案件明細表"
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = 6000
Printer.CurrentY = k + 500
Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(Txt1(1)) & " ", "@@@@@@@@@") & "-" & ChangeTStringToTDateString(Txt1(2))
Printer.Font.Bold = False
Printer.CurrentX = 500
Printer.CurrentY = k + 800
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = k + 800
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.CurrentX = 500
Printer.CurrentY = k + 1100
Printer.Print "承辦人：" & Area
Printer.CurrentX = 13000
Printer.CurrentY = k + 1100
Printer.Print "頁次：" & Page
Printer.CurrentX = 500
Printer.CurrentY = k + 1400
Printer.Print String(200, "-")

Printer.Font.Underline = True
Printer.CurrentX = PLeft1(0)
Printer.CurrentY = k + 1700
Printer.Print "本所期限"
Printer.CurrentX = PLeft1(1)
Printer.CurrentY = k + 1700
Printer.Print "法定期限"
Printer.CurrentX = PLeft1(2)
Printer.CurrentY = k + 1700
Printer.Print "本所案號"
Printer.CurrentX = PLeft1(3)
Printer.CurrentY = k + 1700
Printer.Print "案件名稱"
Printer.CurrentX = PLeft1(4)
Printer.CurrentY = k + 1700
Printer.Print "申請國家"
Printer.CurrentX = PLeft1(5)
Printer.CurrentY = k + 1700
Printer.Print "案件性質"
Printer.CurrentX = PLeft1(6)
Printer.CurrentY = k + 1700
Printer.Print "申請人"
Printer.CurrentX = PLeft1(7)
Printer.CurrentY = k + 1700
Printer.Print "智權人員"
Printer.Font.Underline = False
Printer.CurrentX = 500
Printer.CurrentY = k + 2000
Printer.Print String(200, "-")
End Sub

Sub GetPrintLeft3()                            '設定定位點
    Erase PLeft3
    PLeft3(0) = 500
    PLeft3(1) = 1500
    PLeft3(2) = 2700
    PLeft3(3) = 3800
    PLeft3(4) = 5700
    PLeft3(5) = 10000
    PLeft3(6) = 11200
    PLeft3(7) = 13200
    PLeft3(8) = 15200
End Sub
Sub GetPrintLeft2()              '設定定位點
    Erase Pleft2
    Pleft2(0) = 0
    Pleft2(1) = 1100 + 100
    Pleft2(2) = 2200 + 100
    Pleft2(3) = 3200 + 100 + 100
    Pleft2(4) = 5100 + 100 + 100
    Pleft2(5) = 8100 + 100 + 100
    Pleft2(6) = 9400 + 100 + 100
    Pleft2(7) = 10400 + 100 + 100
    Pleft2(8) = 11200 + 100 + 100
    Pleft2(9) = 12200 + 100 + 100
    Pleft2(10) = 13400 + 100 + 100
End Sub
Sub GetPrintLeft1()          '設定定位點
    Erase PLeft1
    PLeft1(0) = 500
    PLeft1(1) = 1600 + 100
    PLeft1(2) = 2700 + 100
    PLeft1(3) = 4600 + 100
    PLeft1(4) = 8900 + 100
    PLeft1(5) = 10100 + 100
    PLeft1(6) = 11500 + 100
    PLeft1(7) = 13100 + 100
End Sub
Sub PriTiEnd3()          '印尾巴
End Sub
Sub PriTiEnd2()          '印尾巴
End Sub
Sub PriTiEnd1()          '印尾巴
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2, 3, 4 '本所期限起, 迄, 法定期限起, 迄
   If PUB_CheckKeyInDate(Me.Txt1(Index)) = -1 Then
      Cancel = True
   End If

Case 5
     If Txt1(Index) <> "1" And Txt1(Index) <> "2" Then
        s = MsgBox("管制對象必須輸入 1 或 2 !!", , "USER 輸入錯誤")
        Cancel = True
     End If
Case 6
   If Txt1(Index) <> "" Then
      'edit by nickc 2007/02/09 不用 dll 了
      'If objPublicData.GetStaff(txt1(Index), strExc(0)) Then
      If ClsPDGetStaff(Txt1(Index), strExc(0)) Then
         LBL1(0) = strExc(0)
      Else
         LBL1(0) = ""
         Cancel = True
      End If
   End If
Case 7
     If Len(Txt1(7)) <> 0 Then
     '   StrTempP = Split(Replace(txt1(0), ",,", ""), ",")
      LBL1(1) = GetPrjState6HM("P", Txt1(Index))
      If LBL1(1) = "" Then
         MsgBox "案件性質錯誤，請重新輸入 !", vbCritical
         Cancel = True
      End If
     End If
Case Else

End Select
If Cancel Then TextInverse Txt1(Index)
End Sub
