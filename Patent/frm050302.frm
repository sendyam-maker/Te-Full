VERSION 5.00
Begin VB.Form frm050302 
   BorderStyle     =   1  '單線固定
   Caption         =   "期限管制表"
   ClientHeight    =   3795
   ClientLeft      =   3480
   ClientTop       =   3195
   ClientWidth     =   4410
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4410
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   10
      Left            =   1155
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1920
      Width           =   330
   End
   Begin VB.TextBox txtPA46 
      Height          =   285
      Left            =   1785
      TabIndex        =   11
      Top             =   3135
      Width           =   330
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3420
      TabIndex        =   13
      Top             =   150
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2625
      TabIndex        =   12
      Top             =   150
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   2610
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2805
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1155
      MaxLength       =   9
      TabIndex        =   9
      Top             =   2805
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1155
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2505
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1575
      MaxLength       =   6
      TabIndex        =   7
      Top             =   2205
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   5
      Left            =   1155
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1605
      Width           =   330
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2565
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1320
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1320
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2565
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1005
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1005
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1155
      TabIndex        =   0
      Top             =   705
      Width           =   2325
   End
   Begin VB.OptionButton Option1 
      Caption         =   "法定期限："
      Height          =   180
      Index           =   1
      Left            =   255
      TabIndex        =   15
      Top             =   1350
      Width           =   1260
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   255
      TabIndex        =   14
      Top             =   1050
      Value           =   -1  'True
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   255
      TabIndex        =   24
      Top             =   750
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "列印對象：　   （1.非智權部同仁 2.全部）"
      Height          =   180
      Index           =   10
      Left            =   255
      TabIndex        =   23
      Top             =   1965
      Width           =   3330
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   180
      Index           =   5
      Left            =   255
      TabIndex        =   22
      Top             =   2850
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員/承辦人："
      Height          =   180
      Index           =   3
      Left            =   130
      TabIndex        =   21
      Top             =   2250
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   4
      Left            =   255
      TabIndex        =   20
      Top             =   2550
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "管制對象：　     (1.承辦人 2.智權人員)"
      Height          =   180
      Index           =   1
      Left            =   255
      TabIndex        =   19
      Top             =   1650
      Width           =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PCT進入國家階段：　 （Y：國家階段）"
      Height          =   180
      Index           =   7
      Left            =   255
      TabIndex        =   18
      Top             =   3180
      Width           =   3180
   End
   Begin VB.Line Line3 
      X1              =   2250
      X2              =   2490
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   1
      Left            =   1965
      TabIndex        =   17
      Top             =   2550
      Width           =   1635
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   0
      Left            =   2385
      TabIndex        =   16
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   2205
      X2              =   2445
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   2205
      X2              =   2445
      Y1              =   1140
      Y2              =   1140
   End
End
Attribute VB_Name = "frm050302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
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
'Add By Cheng 2002/03/15
Dim StrR003020 As String
'Add By Cheng 2002/09/12
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
      'Add By Cheng 2002/09/12
      blnClkSure = False
     
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         If Option1(0).Value = True Then
            'Add By Cheng 2002/03/19
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
            'Add By Cheng 2002/09/12
            If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
               If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
                  MsgBox "本所期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
                  
            If Len(Trim(txt1(2))) = 0 Then
                s = MsgBox("本所期限不可空白", , "USER 輸入錯誤")
                txt1(1).SetFocus
                txt1_GotFocus (1)
                Exit Sub
            End If
         Else
            If Option1(1).Value = True Then
               'Add By Cheng 2002/03/19
               If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
               If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
                  Me.txt1(4).SetFocus
                  txt1_GotFocus 4
                  Exit Sub
               End If
               'Add By Cheng 2002/09/12
               If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
                  If Val(Me.txt1(3).Text) > Val(Me.txt1(4).Text) Then
                     MsgBox "法定期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(3).SetFocus
                     txt1_GotFocus 3
                     Exit Sub
                  End If
               End If
                               
                If Len(Trim(txt1(4))) = 0 Then
                    s = MsgBox("法定期限不可空白", , "USER 輸入錯誤")
                    txt1(3).SetFocus
                    txt1_GotFocus (3)
                    Exit Sub
                End If
            End If
         End If
         If Len(Trim(txt1(5))) = 0 Then
            s = MsgBox("管制對象不可空白", , "USER 輸入錯誤")
            txt1(5).SetFocus
            Exit Sub
         End If
     End If
     If txt1(6) <> "" Then
         'Add By Cheng 2002/09/12
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(txt1(6), strExc(0)) Then
         If ClsPDGetStaffN(txt1(6), strExc(0)) Then
            lbl1(0) = strExc(0)
         Else
            lbl1(0) = ""
            Me.txt1(6).SetFocus
            txt1_GotFocus 6
            Exit Sub
         End If
      End If
      If Len(txt1(7)) <> 0 Then
         '2009/12/21 modify by sonia 以第一個系統類別抓,否則CFP的607會錯誤
         'lbl1(1) = GetPrjState6HM("P", txt1(7))
         StrTempP = Split(Replace(txt1(0), ",,", ""), ",")
         StrSQL6 = StrTempP(0)
         lbl1(1) = GetPrjState6HM(StrSQL6, txt1(7))
         '2009/12/21 end
         If lbl1(1) = "" Then
            MsgBox "案件性質錯誤，請重新輸入 !", vbCritical
            Me.txt1(7).SetFocus
            txt1_GotFocus 7
            Exit Sub
         End If
      End If
     
     If Len(Trim(txt1(8))) <> 0 Or Len(Trim(txt1(9))) <> 0 Then
        If Left(txt1(8), 6) <> Left(txt1(9), 6) Then
            s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
            blnClkSure = True
             txt1(8).SetFocus
             txt1_GotFocus (8)
            Exit Sub
        End If
     End If
      'Add By Cheng 2002/09/12
      If Me.txt1(8).Text <> "" And Me.txt1(9).Text <> "" Then
         If Me.txt1(8).Text > Me.txt1(9).Text Then
            MsgBox "申請人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            blnClkSure = True
            Me.txt1(8).SetFocus
            txt1_GotFocus 8
            Exit Sub
         End If
      End If
     
     Screen.MousePointer = vbHourglass
     Me.Enabled = False
     StrMenu
     Me.Enabled = True
     Screen.MousePointer = vbDefault
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Activate()
MoveFormToCenter Me
End Sub

Private Sub Form_Load()
txt1(0) = GetSystemKindByNick
Option1(1).Value = False
txt1(3).Enabled = False
txt1(4).Enabled = False
End Sub

Sub StrMenu()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   ClearQueryLog (Me.Name) '2009/12/17 ADD BY SONIA 清除查詢印表記錄檔欄位
   
   If txt1(5) = "1" And Option1(0).Value = True Then
       StrMenu1 '本所期限+承辦人(管制對象)
   Else
       If txt1(5) = "1" And Option1(1).Value = True Then
           StrMenu2 '法定期限+承辦人(管制對象)
       Else
           If txt1(5) = "2" And Option1(0).Value = True Then
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
   If Len(Trim(txt1(6))) <> 0 Then
      If Trim(txt1(5)) = "1" Then
         StrSQL3 = StrSQL3 & " AND CP14='" & txt1(6) & "' "
      Else
         StrSQL3 = StrSQL3 & " AND CP13='" & txt1(6) & "' "
      End If
      pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(6) & lbl1(0)        '2009/12/17 add by sonia
   End If
   If Len(Trim(txt1(7))) <> 0 Then
      StrSQL3 = StrSQL3 & " AND CP10='" & txt1(7) & "' "
      pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(7) & lbl1(1)        '2009/12/17 add by sonia
   End If
   If Len(txt1(0)) <> 0 Then
      StrSQL3 = StrSQL3 & " AND CP01 IN (" & GetAddStr(txt1(0)) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0)                  '2009/12/17 add by sonia
   End If
   strSql = "SELECT CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP44 FROM CASEPROGRESS WHERE CP06>=" & Val(ChangeTStringToWString(txt1(1))) & " AND CP06<=" & Val(ChangeTStringToWString(txt1(2))) & " AND CP01 NOT IN ('P','FCP','T','CFT','FCT','TF','LA','CFL','FCL','L') AND CP27 IS NULL AND CP57 IS NULL " & StrSQL3
   pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(1) & "-" & txt1(2) & ";管制對象：承辦人"    '2009/12/17 add by sonia
   
   'Add by Morgan 2005/2/14
   If txtPA46 = "Y" Then
      strSql = strSql & " And Exists( select * from patent where pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and PA09<>'056' AND PA46='Y') "
      pub_QL05 = pub_QL05 & ";" & Left(Label1(7), 9) '2009/12/17 add by sonia
   End If
   
   '2009/12/17 add by sonia
   If Len(Trim(txt1(8))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(8) & "-" & txt1(9)  '2009/12/17 add by sonia
   End If
   '2009/12/17 end
   
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
         'DoPaAndSp1
         If DoPaAndSp1_1 <> 0 Then GoTo NextRecord
        
         s = 1
'2011/11/16 modify by sonia起迄要同時檢查,否則下X43727時第二申請人<=X43727的都會出來
'         If Len(Trim(txt1(8))) <> 0 And s <> 0 Then
'            If (IIf(StrR002009 = "", False, StrR002009 >= GetNewFagent(txt1(8))) Or IIf(StrR002010 = "", False, StrR002010 >= GetNewFagent(txt1(8))) Or IIf(StrR002011 = "", False, StrR002011 >= GetNewFagent(txt1(8))) Or IIf(StrR002012 = "", False, StrR002012 >= GetNewFagent(txt1(8))) Or IIf(StrR002013 = "", False, StrR002013 >= GetNewFagent(txt1(8)))) Then
'               s = 1
'            Else
'               s = 0
'            End If
'         Else
'            s = 1
'         End If
'         '911023 nick
'         '***** start
'         If s <> 0 Then
'            If Len(Trim(txt1(9))) <> 0 Then
'               If (IIf(StrR002009 = "", False, StrR002009 <= GetNewFagent(txt1(9))) Or IIf(StrR002010 = "", False, StrR002010 <= GetNewFagent(txt1(9))) Or IIf(StrR002011 = "", False, StrR002011 <= GetNewFagent(txt1(9))) Or IIf(StrR002012 = "", False, StrR002012 <= GetNewFagent(txt1(9))) Or IIf(StrR002013 = "", False, StrR002013 <= GetNewFagent(txt1(9)))) Then
'                  s = 1
'               Else
'                  s = 0
'               End If
'            End If
'         End If
'         '***** end
         If Len(Trim(txt1(8))) <> 0 Then
            If Len(Trim(StrR002009)) <> 0 And StrR002009 >= GetNewFagent(txt1(8)) And StrR002009 <= GetNewFagent(txt1(9)) Then
               s = 1
            Else
               s = 0
            End If
            If s = 0 Then
               If Len(Trim(StrR002010)) <> 0 And StrR002010 >= GetNewFagent(txt1(8)) And StrR002010 <= GetNewFagent(txt1(9)) Then
                  s = 1
               Else
                  s = 0
               End If
            End If
            If s = 0 Then
               If Len(Trim(StrR002011)) <> 0 And StrR002011 >= GetNewFagent(txt1(8)) And StrR002011 <= GetNewFagent(txt1(9)) Then
                  s = 1
               Else
                  s = 0
               End If
            End If
            If s = 0 Then
               If Len(Trim(StrR002012)) <> 0 And StrR002012 >= GetNewFagent(txt1(8)) And StrR002012 <= GetNewFagent(txt1(9)) Then
                  s = 1
               Else
                  s = 0
               End If
            End If
            If s = 0 Then
               If Len(Trim(StrR002013)) <> 0 And StrR002013 >= GetNewFagent(txt1(8)) And StrR002013 <= GetNewFagent(txt1(9)) Then
                  s = 1
               Else
                  s = 0
               End If
            End If
         Else
            s = 1
         End If
'2011/11/16 END
         
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
            If Val(StrR002014) < Val(strSrvDate(1)) Then
               StrR002014 = "*" + ChangeTStringToTDateString(ChangeWStringToTString(StrR002014))
            Else
               If Val(StrR002014) = Val(GetTodayDate) Then
                  StrR002014 = "V" + ChangeTStringToTDateString(ChangeWStringToTString(StrR002014))
               Else
                  StrR002014 = ChangeTStringToTDateString(ChangeWStringToTString(StrR002014))
               End If
            End If
            If LenB(StrR002005) > 8 Then
               StrR002005 = StrToStr(StrR002005, 4)
            End If
            If LenB(StrR002006) > 8 Then
               StrR002006 = StrToStr(StrR002006, 4)
            End If
            If LenB(StrR002007) > 12 Then
               StrR002007 = StrToStr(StrR002007, 6)
            End If
            If LenB(StrR002008) > 9 Then
               StrR002008 = StrToStr(StrR002008, 4)
            End If
            If LenB(StrR002009) > 12 Then
               StrR002009 = StrToStr(StrR002009, 6)
            End If
            If LenB(StrR002010) > 12 Then
               StrR002010 = StrToStr(StrR002010, 6)
            End If
            If LenB(StrR002011) > 12 Then
               StrR002011 = StrToStr(StrR002011, 6)
            End If
            If LenB(StrR002012) > 12 Then
               StrR002012 = StrToStr(StrR002012, 6)
            End If
            If LenB(StrR002013) > 12 Then
               StrR002013 = StrToStr(StrR002013, 6)
            End If
            strSql = "INSERT INTO R050302_1 VALUES ('" & ChgSQL(StrR002001) & "','" & ChgSQL(StrR002002) & "','" & ChgSQL(StrR002003) & "','" & ChgSQL(StrR002004) & "','" & ChgSQL(StrR002005) & "','" & ChgSQL(StrR002006) & "','" & ChgSQL(StrR002007) & "','" & ChgSQL(StrR002008) & "','" & ChgSQL(StrR002009) & "','" & ChgSQL(StrR002010) & "','" & ChgSQL(StrR002011) & "','" & ChgSQL(StrR002012) & "','" & ChgSQL(StrR002013) & "','" & ChgSQL(StrR002014) & "','" & ChgSQL(StrR002015) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
         End If
   
   'Add By Cheng 2002/01/29
NextRecord:
   
         adoRecordset.MoveNext
         DoEvents
      Loop
      If adoRecordset.RecordCount = 0 Then
         InsertQueryLog (0)  '2010/1/7 ADD BY SONIA
         ShowNoData
         Screen.MousePointer = vbDefault
         Exit Sub
      '2009/12/17 add by sonia
      Else
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         strSql = "select * from R050302_1 where id='" & strUserNum & "' "
         adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
         If adoRecordset.RecordCount <> 0 Then
            InsertQueryLog (adoRecordset.RecordCount)
         End If
      '2009/12/17 end
      End If
   Else
      InsertQueryLog (0)  '2010/1/7 ADD BY SONIA
      ShowNoData
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   PriMenu1
End Sub
Sub StrMenu2()              '處理主程式
   Screen.MousePointer = vbHourglass
   cnnConnection.Execute "delete from R050302_2 WHERE ID='" & strUserNum & "' "
   'Modify By Cheng 2002/03/15
   '強迫寫入"1"表示資料來自案件進度檔
   'strSQL = "SELECT CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP44,CP05,CP09,CP64 FROM CASEPROGRESS WHERE CP07>=" & Val(ChangeTStringToWString(txt1(3))) & " AND CP07<=" & Val(ChangeTStringToWString(txt1(4))) & " AND (CP01<>'T' AND CP01<>'CFT' AND CP01<>'FCT' AND CP01<>'TF' AND CP01<>'LA' AND CP01<>'CFL' AND CP01<>'FCL' AND CP01<>'L' AND CP01<>'P' ) AND CP27 IS NULL AND CP57 IS NULL "
   strSql = "SELECT CP01,CP02,CP03,CP04,CP06,CP07,CP10,CP13,CP14,CP44,CP05,CP09,CP64,'1' FROM CASEPROGRESS WHERE CP07>=" & Val(ChangeTStringToWString(txt1(3))) & " AND CP07<=" & Val(ChangeTStringToWString(txt1(4))) & " AND (CP01<>'T' AND CP01<>'CFT' AND CP01<>'FCT' AND CP01<>'TF' AND CP01<>'LA' AND CP01<>'CFL' AND CP01<>'FCL' AND CP01<>'L' AND CP01<>'P' ) AND CP27 IS NULL AND CP57 IS NULL "
   pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & txt1(3) & "-" & txt1(4)     '2009/12/17 add by sonia
   If txt1(5) = "1" Then
      pub_QL05 = pub_QL05 & ";管制對象：承辦人"                           '2009/12/17 add by sonia
      If Len(txt1(6)) <> 0 Then
         strSql = strSql & " AND CP14='" & txt1(6) & "' "
         pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(6) & lbl1(0)        '2009/12/17 add by sonia
      End If
   Else
      pub_QL05 = pub_QL05 & ";管制對象：智權人員"                           '2009/12/17 add by sonia
      If Len(txt1(6)) <> 0 Then
         StrSQL3 = StrSQL3 & " AND CP13='" & txt1(6) & "' "
         pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(6) & lbl1(0)        '2009/12/17 add by sonia
      End If
   End If
   If Len(txt1(7)) <> 0 Then
      strSql = strSql & " AND CP10='" & Val(txt1(7)) & "' "
      pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(7) & lbl1(1)        '2009/12/17 add by sonia
   End If
   
   CheckOC
   If Len(txt1(0)) <> 0 Then
       'strTemp1 = Split(Replace(txt1(0), ",,", ""), ",")
       strSql = strSql & " AND CP01 IN (" & GetAddStr(txt1(0)) & ") "
       pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0)                  '2009/12/17 add by sonia
   End If
   
   'Add by Morgan 2005/2/14
   If txtPA46 = "Y" Then
      strSql = strSql & " And Exists( select * from patent where pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and PA09<>'056' AND PA46='Y') "
      pub_QL05 = pub_QL05 & ";" & Left(Label1(7), 9) '2009/12/17 add by sonia
   End If
   
   'Add by Morgan 2006/5/30
   If txt1(5) = "2" Then
      If txt1(10) = "1" Then
         strSql = strSql & " And exists(select * from staff where st01=cp13 and  SUBSTR(ST15,1,1)<>'S')"
         pub_QL05 = pub_QL05 & ";" & Left(Label1(10), 5) & "非智權部同仁" '2009/12/17 add by sonia
      End If
   End If
   
   'Modify By Cheng 2002/03/15
   '強迫寫入"2"表示資料來自下一程序檔
   'strSQL = strSQL + " union all select NP02,NP03,NP04,NP05,NP08,NP09,NP07,NP10,CP14,CP44,CP05,CP09,NP15 FROM NEXTPROGRESS,CASEPROGRESS WHERE NP09>=" & Val(ChangeTStringToWString(txt1(3))) & " AND NP09<=" & Val(ChangeTStringToWString(txt1(4))) & " AND (NP02<>'T' AND NP02<>'CFT' AND NP02<>'FCT' AND NP02<>'TF' AND NP02<>'LA' AND NP02<>'CFL' AND NP02<>'FCL' AND CP01<>'P' AND NP02<>'L') AND NP06 IS NULL AND CP09=NP01 AND CP27 IS NOT NULL"
   '91.11.21 MODIFY BY SONIA
   'strSQL = strSQL + " union all select NP02,NP03,NP04,NP05,NP08,NP09,NP07,NP10,CP14,CP44,CP05,CP09,NP15,'2' FROM NEXTPROGRESS,CASEPROGRESS WHERE NP09>=" & Val(ChangeTStringToWString(txt1(3))) & " AND NP09<=" & Val(ChangeTStringToWString(txt1(4))) & " AND (NP02<>'T' AND NP02<>'CFT' AND NP02<>'FCT' AND NP02<>'TF' AND NP02<>'LA' AND NP02<>'CFL' AND NP02<>'FCL' AND CP01<>'P' AND NP02<>'L') AND NP06 IS NULL AND CP09=NP01 AND CP27 IS NOT NULL"
   '92.03.27
   'strSQL = strSQL + " union all select NP02,NP03,NP04,NP05,NP08,NP09,NP07,NP10,CP14,CP44,CP05,CP09,NP15,'2' FROM NEXTPROGRESS,CASEPROGRESS WHERE NP09>=" & Val(ChangeTStringToWString(txt1(3))) & " AND NP09<=" & Val(ChangeTStringToWString(txt1(4))) & " AND (NP02<>'T' AND NP02<>'CFT' AND NP02<>'FCT' AND NP02<>'TF' AND NP02<>'LA' AND NP02<>'CFL' AND NP02<>'FCL' AND CP01<>'P' AND NP02<>'L') AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') AND CP09=NP01 AND CP27 IS NOT NULL"
   'Modify by Morgan 2009/7/13 +995,996
   'strSQL = strSQL + " union all select NP02,NP03,NP04,NP05,NP08,NP09,NP07,NP10,CP14,CP44,CP05,CP09,NP15,'2' FROM NEXTPROGRESS,CASEPROGRESS WHERE NP09>=" & Val(ChangeTStringToWString(txt1(3))) & " AND NP09<=" & Val(ChangeTStringToWString(txt1(4))) & " AND (NP02<>'T' AND NP02<>'CFT' AND NP02<>'FCT' AND NP02<>'TF' AND NP02<>'LA' AND NP02<>'CFL' AND NP02<>'FCL' AND CP01<>'P' AND NP02<>'L') AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') AND np01=cp09(+) AND CP27 IS NOT NULL"
   strSql = strSql + " union all select NP02,NP03,NP04,NP05,NP08,NP09,NP07,NP10,CP14,CP44,CP05,CP09,NP15,'2' FROM NEXTPROGRESS,CASEPROGRESS WHERE NP09>=" & Val(ChangeTStringToWString(txt1(3))) & " AND NP09<=" & Val(ChangeTStringToWString(txt1(4))) & " AND (NP02<>'T' AND NP02<>'CFT' AND NP02<>'FCT' AND NP02<>'TF' AND NP02<>'LA' AND NP02<>'CFL' AND NP02<>'FCL' AND CP01<>'P' AND NP02<>'L') AND NP06 IS NULL AND NP07 NOT IN ('411','997','998','995','996') AND np01=cp09(+) AND CP27 IS NOT NULL"
   '91.11.21 END
   If Len(txt1(6)) <> 0 And txt1(5) = "2" Then
      strSql = strSql & " AND NP10='" & txt1(6) & "' "
   End If
   If Len(txt1(7)) <> 0 Then
      strSql = strSql & " AND NP07=" & Val(txt1(7)) & " "
   End If
   If Len(txt1(0)) <> 0 Then
       'strTemp1 = Split(Replace(txt1(0), ",,", ""), ",")
       strSql = strSql & " AND NP02 IN (" & GetAddStr(txt1(0)) & ") "
   End If
   
   'Add by Morgan 2005/2/14
   If txtPA46 = "Y" Then
      strSql = strSql & " And Exists( select * from patent where pa01=np02 and pa02=np03 and pa03=np04 and pa04=np05 and PA09<>'056' AND PA46='Y') "
   End If
   
   'Add by Morgan 2006/5/30
   If txt1(5) = "2" Then
      If txt1(10) = "1" Then
         strSql = strSql & " And exists(select * from staff where st01=np10 and  SUBSTR(ST15,1,1)<>'S')"
      End If
   '2009/12/21 add by sonia
   Else
      If Len(txt1(6)) <> 0 Then
         strSql = strSql & " AND CP14='" & txt1(6) & "' "
      End If
   '2009/12/21 end
   End If
   
   '2009/12/17 add by sonia
   If Len(Trim(txt1(8))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(8) & "-" & txt1(9)  '2009/12/17 add by sonia
   End If
   '2009/12/17 end
      
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
            'Add By Cheng 2002/03/15
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
            'Modify By Cheng 2002/03/15
   '        If Not IsNull(adoRecordset.Fields(9)) Then
   '            StrR003007 = adoRecordset.Fields(9)
   '        Else
   '            StrR003007 = ""
   '        End If
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
           'Add By Cheng 2002/03/15
           If Not IsNull(adoRecordset.Fields(13)) Then
               StrR003020 = adoRecordset.Fields(13)
           Else
               StrR003020 = ""
           End If
           
           DoEvents
            'Modify By Cheng 2002/01/29
           '若專利基本檔或服務業務基本檔已閉卷, 則不列印
   '        DoPaAndSp2
           If DoPaAndSp2_1 <> 0 Then GoTo NextRecord
           
           s = 0
'2011/11/16 modify by sonia起迄要同時檢查,否則下X43727時第二申請人<=X43727的都會出來
'           If Len(Trim(txt1(8))) <> 0 Then
'               If (IIf(StrR003010 = "", False, StrR003010 >= GetNewFagent(txt1(8))) Or IIf(StrR003011 = "", False, StrR003011 >= GetNewFagent(txt1(8))) Or IIf(StrR003012 = "", False, StrR003012 >= GetNewFagent(txt1(8))) Or IIf(StrR003013 = "", False, StrR003013 >= GetNewFagent(txt1(8))) Or IIf(StrR003014 = "", False, StrR003014 >= GetNewFagent(txt1(8)))) Then
'                   s = 1
'               Else
'                   s = 0
'               End If
'           Else
'               s = 1
'           End If
'           '911023 nick
'           '***** start
'           If s <> 0 Then
'               If Len(Trim(txt1(9))) <> 0 Then
'                   If (IIf(StrR003010 = "", False, StrR003010 <= GetNewFagent(txt1(9))) Or IIf(StrR003011 = "", False, StrR003011 <= GetNewFagent(txt1(9))) Or IIf(StrR003012 = "", False, StrR003012 <= GetNewFagent(txt1(9))) Or IIf(StrR003013 = "", False, StrR003013 <= GetNewFagent(txt1(9))) Or IIf(StrR003014 = "", False, StrR003014 <= GetNewFagent(txt1(9)))) Then
'                       s = 1
'                   Else
'                       s = 0
'                   End If
'               End If
'           End If
'           '***** end
         If Len(Trim(txt1(8))) <> 0 Then
            If Len(Trim(StrR003010)) <> 0 And StrR003010 >= GetNewFagent(txt1(8)) And StrR003010 <= GetNewFagent(txt1(9)) Then
               s = 1
            Else
               s = 0
            End If
            If s = 0 Then
               If Len(Trim(StrR003011)) <> 0 And StrR003011 >= GetNewFagent(txt1(8)) And StrR003011 <= GetNewFagent(txt1(9)) Then
                  s = 1
               Else
                  s = 0
               End If
            End If
            If s = 0 Then
               If Len(Trim(StrR003012)) <> 0 And StrR003012 >= GetNewFagent(txt1(8)) And StrR003012 <= GetNewFagent(txt1(9)) Then
                  s = 1
               Else
                  s = 0
               End If
            End If
            If s = 0 Then
               If Len(Trim(StrR003013)) <> 0 And StrR003013 >= GetNewFagent(txt1(8)) And StrR003013 <= GetNewFagent(txt1(9)) Then
                  s = 1
               Else
                  s = 0
               End If
            End If
            If s = 0 Then
               If Len(Trim(StrR003014)) <> 0 And StrR003014 >= GetNewFagent(txt1(8)) And StrR003014 <= GetNewFagent(txt1(9)) Then
                  s = 1
               Else
                  s = 0
               End If
            End If
         Else
            s = 1
         End If
'2011/11/16 END
           'If Len(Trim(txt1(10))) <> 0 And S <> 0 Then
           '    If StrR003008 >= txt1(10) Or StrR003009 >= txt1(10) Then
           '        S = 1
           '    Else
           '        S = 0
           '    End If
           'End If
           'If Len(Trim(txt1(11))) <> 0 And S <> 0 Then
           '    If StrR003008 <= txt1(11) Or StrR001009 <= txt1(11) Then
           '        S = 1
           '    Else
           '        S = 0
           '    End If
           'End If
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
               'If LenB(StrR003005) > 8 Then
               '    StrR003005 = LeftB(StrR003005, 8)
               'End If
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
               'Modify By Cheng 2002/03/15
   '            strSQL = "INSERT INTO R050302_2 VALUES ('" & ChgSQL(StrR003001) & "','" & ChgSQL(StrR003002) & "','" & ChgSQL(StrR003003) & "','" & ChgSQL(StrR003004) & "','" & ChgSQL(StrR003005) & "','" & ChgSQL(StrR003006) & "','" & ChgSQL(StrR003007) & "','" & ChgSQL(StrR003008) & "','" & ChgSQL(StrR003009) & "','" & ChgSQL(StrR003010) & "','" & ChgSQL(StrR003011) & "','" & ChgSQL(StrR003012) & "','" & ChgSQL(StrR003013) & "','" & ChgSQL(StrR003014) & "','" & ChgSQL(StrR003015) & "','" & ChgSQL(StrR003016) & "','" & ChgSQL(StrR003017) & "','" & ChgSQL(StrR003018) & "','" & ChgSQL(StrR003019) & "','" & strUserNum & "') "
               strSql = "INSERT INTO R050302_2 VALUES ('" & ChgSQL(StrR003001) & "','" & ChgSQL(StrR003002) & "','" & ChgSQL(StrR003003) & "','" & ChgSQL(StrR003004) & "','" & ChgSQL(StrR003005) & "','" & ChgSQL(StrR003006) & "','" & ChgSQL(StrR003007) & "','" & ChgSQL(StrR003008) & "','" & ChgSQL(StrR003009) & "','" & ChgSQL(StrR003010) & "','" & ChgSQL(StrR003011) & "','" & ChgSQL(StrR003012) & "','" & ChgSQL(StrR003013) & "','" & ChgSQL(StrR003014) & "','" & ChgSQL(StrR003015) & "','" & ChgSQL(StrR003016) & "','" & ChgSQL(StrR003017) & "','" & ChgSQL(StrR003018) & "','" & ChgSQL(StrR003019) & "','" & strUserNum & "','" & StrR003020 & "') "
               cnnConnection.Execute strSql
           End If
           
   'Add By Cheng 2002/01/29
NextRecord:
   
           adoRecordset.MoveNext
           DoEvents
       Loop
       If adoRecordset.RecordCount = 0 Then
           InsertQueryLog (0)  '2010/1/7 ADD BY SONIA
           ShowNoData
           Screen.MousePointer = vbDefault
           Exit Sub
       '2009/12/17 add by sonia
       Else
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         strSql = "select * from R050302_2 where id='" & strUserNum & "' "
         adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
         If adoRecordset.RecordCount <> 0 Then
            InsertQueryLog (adoRecordset.RecordCount)
         End If
       '2009/12/17 end
       End If
   Else
       InsertQueryLog (0)  '2010/1/7 ADD BY SONIA
       ShowNoData
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   PriMenu2
   Screen.MousePointer = vbDefault
End Sub
Sub StrMenu3()         '處理主程式
   Screen.MousePointer = vbHourglass
   cnnConnection.Execute "DELETE FROM R050302_3 WHERE ID='" & strUserNum & "' "
   '91.11.21 modify by sonia
   'strSQL = "SELECT NP01,NP02,NP03,NP04,NP05,NP10,NP08,NP09,NP07,NP10 FROM NEXTPROGRESS WHERE NP08>=" & Val(ChangeTStringToWString(txt1(1))) & " AND NP08<=" & Val(ChangeTStringToWString(txt1(2))) & " AND NP02 NOT IN ('T','CFT','FCT','TF','LA','CFL','FCL','L') AND NP06 IS NULL "
   'Modify by Morgan 2009/7/13 +995,996
   'strSQL = "SELECT NP01,NP02,NP03,NP04,NP05,NP10,NP08,NP09,NP07,NP10 FROM NEXTPROGRESS WHERE NP08>=" & Val(ChangeTStringToWString(txt1(1))) & " AND NP08<=" & Val(ChangeTStringToWString(txt1(2))) & " AND NP02 NOT IN ('T','CFT','FCT','TF','LA','CFL','FCL','L') AND NP06 IS NULL AND NP07 NOT IN ('411','997','998') "
   strSql = "SELECT NP01,NP02,NP03,NP04,NP05,NP10,NP08,NP09,NP07,NP10 FROM NEXTPROGRESS WHERE NP08>=" & Val(ChangeTStringToWString(txt1(1))) & " AND NP08<=" & Val(ChangeTStringToWString(txt1(2))) & " AND NP02 NOT IN ('T','CFT','FCT','TF','LA','CFL','FCL','L') AND NP06 IS NULL AND NP07 NOT IN ('411','997','998','995','996') "
   pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txt1(1) & "-" & txt1(2) & ";管制對象：智權人員"     '2009/12/17 add by sonia
   '91.11.21 end
   If Len(txt1(6)) <> 0 Then
      strSql = strSql & " AND NP10='" & txt1(6) & "' "
      pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(6) & lbl1(0)        '2009/12/17 add by sonia
   End If
   If Len(txt1(7)) <> 0 Then
      strSql = strSql & " AND NP07=" & Val(txt1(7)) & " "
      pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(7) & lbl1(1)        '2009/12/17 add by sonia
   End If
   CheckOC
   If Len(txt1(0)) <> 0 Then
       'strTemp1 = Split(Replace(txt1(0), ",,", ""), ",")
       strSql = strSql & " AND NP02 IN (" & GetAddStr(txt1(0)) & ") "
       pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0)                  '2009/12/17 add by sonia
   End If
   
   'Add by Morgan 2005/2/14
   If txtPA46 = "Y" Then
      strSql = strSql & " And Exists( select * from patent where pa01=np02 and pa02=np03 and pa03=np04 and pa04=np05 and PA09<>'056' AND PA46='Y') "
      pub_QL05 = pub_QL05 & ";" & Left(Label1(7), 9) '2009/12/17 add by sonia
   End If
   
   'Add by Morgan 2006/5/30
   If txt1(5) = "2" Then
      If txt1(10) = "1" Then
         strSql = strSql & " And exists(select * from staff where st01=np10 and  SUBSTR(ST15,1,1)<>'S')"
         pub_QL05 = pub_QL05 & ";" & Left(Label1(10), 5) & "非智權部同仁"                '2009/12/17 add by sonia
      End If
   End If
   
   '2009/12/17 add by sonia
   If Len(Trim(txt1(8))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(8) & "-" & txt1(9)  '2009/12/17 add by sonia
   End If
   '2009/12/17 end
      
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
           
           'Modify By Cheng 2002/01/29
           '若專利基本檔或服務業務基本檔已閉卷, 則不列印
   '        DoPaAndSp
           If DoPaAndSp_1 <> 0 Then GoTo NextRecord
           
           'StrR001007 = CheckStr(adoRecordset.Fields(8))
           StrR001008 = CheckStr(adoRecordset.Fields(8))
           s = 0
           DoStaff
'2011/11/16 modify by sonia起迄要同時檢查,否則下X43727時第二申請人<=X43727的都會出來
'           If Len(Trim(txt1(8))) <> 0 Then
'               If (IIf(StrR001010 = "", False, StrR001010 >= GetNewFagent(txt1(8))) Or IIf(StrR001011 = "", False, StrR001011 >= GetNewFagent(txt1(8))) Or IIf(StrR001012 = "", False, StrR001012 >= GetNewFagent(txt1(8))) Or IIf(StrR001013 = "", False, StrR001013 >= GetNewFagent(txt1(8))) Or IIf(StrR001014 = "", False, StrR001014 >= GetNewFagent(txt1(8)))) Then
'                   s = 1
'               Else
'                   s = 0
'               End If
'           Else
'               s = 1
'           End If
'           '911023 nick
'           '***** start
'           If s <> 0 Then
'               If Len(Trim(txt1(9))) <> 0 Then
'                   If (IIf(StrR001010 = "", False, StrR001010 <= GetNewFagent(txt1(9))) Or IIf(StrR001011 = "", False, StrR001011 <= GetNewFagent(txt1(9))) Or IIf(StrR001012 = "", False, StrR001012 <= GetNewFagent(txt1(9))) Or IIf(StrR001013 = "", False, StrR001013 <= GetNewFagent(txt1(9))) Or IIf(StrR001014 = "", False, StrR001014 <= GetNewFagent(txt1(9)))) Then
'                       s = 1
'                   Else
'                       s = 0
'                   End If
'               Else
'                   s = 1
'               End If
'           End If
'           '***** end
         If Len(Trim(txt1(8))) <> 0 Then
            If Len(Trim(StrR001010)) <> 0 And StrR001010 >= GetNewFagent(txt1(8)) And StrR001010 <= GetNewFagent(txt1(9)) Then
               s = 1
            Else
               s = 0
            End If
            If s = 0 Then
               If Len(Trim(StrR001011)) <> 0 And StrR001011 >= GetNewFagent(txt1(8)) And StrR001011 <= GetNewFagent(txt1(9)) Then
                  s = 1
               Else
                  s = 0
               End If
            End If
            If s = 0 Then
               If Len(Trim(StrR001012)) <> 0 And StrR001012 >= GetNewFagent(txt1(8)) And StrR001012 <= GetNewFagent(txt1(9)) Then
                  s = 1
               Else
                  s = 0
               End If
            End If
            If s = 0 Then
               If Len(Trim(StrR001013)) <> 0 And StrR001013 >= GetNewFagent(txt1(8)) And StrR001013 <= GetNewFagent(txt1(9)) Then
                  s = 1
               Else
                  s = 0
               End If
            End If
            If s = 0 Then
               If Len(Trim(StrR001014)) <> 0 And StrR001014 >= GetNewFagent(txt1(8)) And StrR001014 <= GetNewFagent(txt1(9)) Then
                  s = 1
               Else
                  s = 0
               End If
            End If
         Else
            s = 1
         End If
'2011/11/16 END
           'If Len(Trim(txt1(10))) <> 0 And S <> 0 Then
           '    If StrR001009 >= txt1(10) Or StrR001015 >= txt1(10) Then
           '        S = 1
           '    Else
           '        S = 0
           '    End If
           'Else
           '    S = 1
           'End If
           'If Len(Trim(txt1(11))) <> 0 And S <> 0 Then
           '    If StrR001009 <= txt1(11) Or StrR001015 <= txt1(11) Then
           '        S = 1
           '    Else
           '        S = 0
           '    End If
           'Else
           '    S = 1
           'End If
           DoEvents
           If s = 0 Then
               adoRecordset.Delete
           Else
               'StrR001006 = StrR001006
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
               If Val(StrR001017) < Val(strSrvDate(1)) Then
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
   
   'Add By Cheng 2002/01/29
NextRecord:
   
           adoRecordset.MoveNext
           DoEvents
       Loop
       If adoRecordset.RecordCount = 0 Then
           InsertQueryLog (0)  '2010/1/7 ADD BY SONIA
           ShowNoData
           Screen.MousePointer = vbDefault
           Exit Sub
       '2009/12/17 add by sonia
       Else
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         strSql = "select * from R050302_3 where id='" & strUserNum & "' "
         adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
         If adoRecordset.RecordCount <> 0 Then
            InsertQueryLog (adoRecordset.RecordCount)
         End If
       '2009/12/17 end
       End If
   Else
       InsertQueryLog (0)  '2010/1/7 ADD BY SONIA
       ShowNoData
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   PriMenu3
   Screen.MousePointer = vbDefault
End Sub
Sub DoStaff()
'91.11.21 MODIFY BY SONIA
'strSQL = "SELECT ST03 FROM STAFF WHERE ST01='" & ChgSQL(StrR001006) & "'"
strSql = "SELECT ST15 FROM STAFF WHERE ST01='" & ChgSQL(StrR001006) & "'"
'91.11.21 END
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
Sub DoPaAndSp2()
strSql = "SELECT PA26,PA27,PA28,PA29,PA30,PA75 FROM PATENT WHERE PA01='" & ChgSQL(StrR003001) & "' AND PA02='" & ChgSQL(StrR003002) & "' AND PA03='" & ChgSQL(StrR003003) & "' AND PA04='" & ChgSQL(StrR003004) & "' AND (PA57<>'Y' or pa57 is null) "
strSql = strSql + " union all select SP08,SP58,SP59,'','',SP26 FROM SERVICEPRACTICE WHERE SP01='" & ChgSQL(StrR003001) & "' AND SP02='" & ChgSQL(StrR003002) & "' AND SP03='" & ChgSQL(StrR003003) & "' AND SP04='" & ChgSQL(StrR003004) & "' AND (SP15<>'Y' or sp15 is null) "
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 Then
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

End Sub

'Add By Cheng 2002/01/29
'取代 Sub DoPaAndSp2
Function DoPaAndSp2_1() As Integer

DoPaAndSp2_1 = -1
strSql = "SELECT PA26,PA27,PA28,PA29,PA30,PA75 FROM PATENT WHERE PA01='" & ChgSQL(StrR003001) & "' AND PA02='" & ChgSQL(StrR003002) & "' AND PA03='" & ChgSQL(StrR003003) & "' AND PA04='" & ChgSQL(StrR003004) & "' AND (PA57<>'Y' or pa57 is null) "
strSql = strSql + " union all select SP08,SP58,SP59,'','',SP26 FROM SERVICEPRACTICE WHERE SP01='" & ChgSQL(StrR003001) & "' AND SP02='" & ChgSQL(StrR003002) & "' AND SP03='" & ChgSQL(StrR003003) & "' AND SP04='" & ChgSQL(StrR003004) & "' AND (SP15<>'Y' or sp15 is null) "
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

Sub DoPaAndSp1()
strSql = "SELECT PA26,PA27,PA28,PA29,PA30 FROM PATENT WHERE PA01='" & ChgSQL(StrR002001) & "' AND PA02='" & ChgSQL(StrR002002) & "' AND PA03='" & ChgSQL(StrR002003) & "' AND PA04='" & ChgSQL(StrR002004) & "' AND (PA57<>'Y' or pa57 is null) "
strSql = strSql + " union all select SP08,SP58,SP59,'','' FROM SERVICEPRACTICE WHERE SP01='" & ChgSQL(StrR002001) & "' AND SP02='" & ChgSQL(StrR002002) & "' AND SP03='" & ChgSQL(StrR002003) & "' AND SP04='" & ChgSQL(StrR002004) & "' AND (SP15<>'Y' or sp15 is null) "
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 Then
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

End Sub
'Add By Cheng 2002/01/29
'取代 Sub DoPaAndSp1()
Function DoPaAndSp1_1() As Integer

DoPaAndSp1_1 = -1
strSql = "SELECT PA26,PA27,PA28,PA29,PA30 FROM PATENT WHERE PA01='" & ChgSQL(StrR002001) & "' AND PA02='" & ChgSQL(StrR002002) & "' AND PA03='" & ChgSQL(StrR002003) & "' AND PA04='" & ChgSQL(StrR002004) & "' AND (PA57<>'Y' or pa57 is null) "
strSql = strSql + " union all select SP08,SP58,SP59,'','' FROM SERVICEPRACTICE WHERE SP01='" & ChgSQL(StrR002001) & "' AND SP02='" & ChgSQL(StrR002002) & "' AND SP03='" & ChgSQL(StrR002003) & "' AND SP04='" & ChgSQL(StrR002004) & "' AND (SP15<>'Y' or sp15 is null) "
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

Sub DoPaAndSp()
strSql = "SELECT PA26,PA27,PA28,PA29,PA30,PA75 FROM PATENT WHERE PA01='" & ChgSQL(StrR001002) & "' AND PA02='" & ChgSQL(StrR001003) & "' AND PA03='" & ChgSQL(StrR001004) & "' AND PA04='" & ChgSQL(StrR001005) & "' AND (PA57<>'Y' or pa57 is null) "
strSql = strSql + " union all select SP08,SP58,SP59,'','',SP26 FROM SERVICEPRACTICE WHERE SP01='" & ChgSQL(StrR001002) & "' AND SP02='" & ChgSQL(StrR001003) & "' AND SP03='" & ChgSQL(StrR001004) & "' AND SP04='" & ChgSQL(StrR001005) & "' AND (SP15<>'Y' or sp15 is null) "
CheckOC2
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 Then
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
End Sub

'Add By Cheng 2002/01/29
'取代 Sub DoPaAndSp_1
Function DoPaAndSp_1() As Integer

DoPaAndSp_1 = -1
strSql = "SELECT PA26,PA27,PA28,PA29,PA30,PA75 FROM PATENT WHERE PA01='" & ChgSQL(StrR001002) & "' AND PA02='" & ChgSQL(StrR001003) & "' AND PA03='" & ChgSQL(StrR001004) & "' AND PA04='" & ChgSQL(StrR001005) & "' AND (PA57<>'Y' or pa57 is null) "
strSql = strSql + " union all select SP08,SP58,SP59,'','',SP26 FROM SERVICEPRACTICE WHERE SP01='" & ChgSQL(StrR001002) & "' AND SP02='" & ChgSQL(StrR001003) & "' AND SP03='" & ChgSQL(StrR001004) & "' AND SP04='" & ChgSQL(StrR001005) & "' AND (SP15<>'Y' or sp15 is null) "
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
Set frm050302 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
      If Option1(0).Value = True Then
         txt1(1).Enabled = True
         txt1(2).Enabled = True
         txt1(3).Enabled = False
         txt1(4).Enabled = False
         Option1(1).Value = False
      End If
Case 1
      If Option1(1).Value = True Then
         txt1(3).Enabled = True
         txt1(4).Enabled = True
         txt1(1).Enabled = False
         txt1(2).Enabled = False
         Option1(0).Value = False
      End If
Case Else
End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme   '2009/12/17 add by sonia
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/12
   Select Case Index
   Case 5 '管制對象
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
      'Add by Morgan 2006/5/30
      If KeyAscii = Asc("2") Then
         txt1(10).Enabled = True: txt1(10) = "1"
      Else
         txt1(10) = "": txt1(10).Enabled = False
      End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) <> 0 Then
        STRSTRING = ""
        StrTempP = Split(Replace(txt1(0), ",,", ""), ",")
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
            txt1(0).SetFocus
            Exit Sub
        End If
        If Len(txt1(7)) <> 0 Then
            lbl1(1).Caption = GetPrjState4(StrTempP(0) + "---", txt1(7))
        End If
      End If
Case 2, 4
   'Modify By Cheng 2002/09/12
   If blnClkSure = False Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
      End If
   Else
      blnClkSure = False
   End If
Case 9
   'Modify By Cheng 2002/09/12
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
      End If
   Else
      blnClkSure = False
   End If
End Select

End Sub
Sub PriMenu1()                     '印表主程式
Dim StrR050302_1(0 To 14) As String

'Add By Cheng 2002/01/29
Dim strSaleName As String '智權人員名稱
strSaleName = ""
'91/03/12  日期排序不能用符號
'nick
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
         
         'Modify By Cheng 2002/01/29
         '列印智權人員姓名(與上筆相同不列印)
'        Printer.Print LeftB(Format(StrR050302_1(4), "!@@@@@@@@"), 8)
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
                'Iline = Iline + 1
                'iPrint = iPrint + 300
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
'Add By Cheng 2002/03/15
Dim strDate As String '法定期限
Dim strServerDate As String '系統日期
'Add By Cheng 2002/01/29
Dim strSaleName As String '智權人員名稱
strSaleName = ""

'Modify By Cheng 2002/03/15
'strSQL = "SELECT r003001,r003002,r003003,r003004,st02,r003006,r003007,r003008,r003009,r003010,r003011,r003012,r003013,r003014,r003015,r003016,r003017,r003018,r003019,r003005 FROM staff,R050302_2 WHERE r003005=st01(+) and ID='" & strUserNum & "' ORDER BY R003016,R003006,R003001,R003002,R003003,R003004 "
strSql = "SELECT r003001,r003002,r003003,r003004,st02,r003006,r003007,r003008,r003009,r003010,r003011,r003012,r003013,r003014,r003015,r003016,r003017,r003018,r003019,r003005,r003020 FROM staff,R050302_2 WHERE r003005=st01(+) and ID='" & strUserNum & "' ORDER BY R003016,R003006,R003001,R003002,R003003,R003004 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount = 0 Then
    Exit Sub
End If
iLine = 1
Page = 1
'If Not IsNull(adoRecordset.Fields(15)) Then
'    TmpArea = adoRecordset.Fields(15)
'Else
'    TmpArea = ""
'End If
PriTiTle2 1
iPrint = 2700

'Add By Cheng 2002/03/15
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
        If strDate < strSrvDate(1) Then
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
'        Printer.Print StrR050302_2(15)
        Printer.Print strDate & StrR050302_2(15)
        
        Printer.CurrentX = Pleft2(1)
        Printer.CurrentY = iPrint
        Printer.Print StrR050302_2(14)
        '列印承辦人
        Printer.CurrentX = Pleft2(2)
        Printer.CurrentY = iPrint
'        Printer.Print StrToStr(StrR050302_2(5), 4)
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
         'Modify By Cheng 2002/03/15
         '不論智權人員是否與上筆相同皆列印出來
'         'Modify By Cheng 2002/01/29
'         '列印智權人員姓名(與上筆相同不列印)
'         If strSaleName <> StrR050302_2(4) Then
            
         'Modify By Cheng 2002/03/15
'            Printer.Print LeftB(Format(StrR050302_2(4), "!@@@@@@@@"), 8)
          If .Fields(20).Value = "2" Then
              Printer.Print "*" & LeftB(Format(StrR050302_2(4), "!@@@@@@@@"), 8)
          Else
              Printer.Print LeftB(Format(StrR050302_2(4), "!@@@@@@@@"), 8)
          End If
'            strSaleName = StrR050302_2(4)
'         Else
'            Printer.Print ""
'         End If
        Printer.CurrentX = Pleft2(10)
        Printer.CurrentY = iPrint
        Printer.Print LeftB(Format(StrR050302_2(18), "!@@@@@@@@@@@@@@@@@@@@"), 18)
        For j = 1 To 5
            If Len(Trim(StrR050302_2(8 + j))) <> 0 Then
                Printer.CurrentX = Pleft2(7)
                Printer.CurrentY = iPrint
                Printer.Print StrToStr(StrR050302_2(8 + j), 3)
                'Iline = Iline + 1
                'iPrint = iPrint + 300
                'If Not .EOF Then
                '    St = StrR050302_2(15)
                'Else
                '    St = ""
                'End If
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
        'If Len(Trim(StrR050302_2(15))) <> 0 Then
        '    TmpArea = StrR050302_2(15)
        'Else
        '    TmpArea = ""
        'End If
        .MoveNext
        If .EOF = False Then
            'If Not IsNull(.Fields(15)) Then
            '    St = .Fields(15)
            'Else
            '    St = ""
            'End If
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
       ' If Page = 2 And Iline = 22 Then
       '     PriTiEnd3
       '     Printer.EndDoc
       '     Exit Sub
       ' End If
        
    Loop

End With
PriTiEnd2
Printer.EndDoc
ShowPrintOk
CheckOC

End Sub
Sub PriMenu3()                 '印表主程式
Dim StrR050302_3(0 To 18) As String

'Add By Cheng 2002/01/29
Dim strSaleName As String '智權人員名稱
Dim strSaleName2 As String '20140304START Modify By eric

strSaleName = ""
'91/03/12 日期排序不能用符號
'nick
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
    
    strSaleName2 = .Fields(5)           '20140304ADD By eric
    
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
        
        'Modify By Cheng 2002/01/29
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
                'Iline = Iline + 1
                'iPrint = iPrint + 300
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
                strSaleName2 = .Fields(5)               '20140304ADD eric
            Else
                St = ""
            End If
            
            '20140304START Modify By eric  增加 非智權同仁每人一張表
            'If ((iLine Mod 25 = 0) Or (TmpArea <> St)) And (iLine <> 0) Then
            '    PriTiEnd3
            '    Printer.NewPage
            '    Page = Page + 1
            '    PriTiTle3 St, str(Page)
            '    iPrint = 2400
            '    iLine = 0
            'End If
            If ((iLine Mod 25 = 0) Or (TmpArea <> St) Or ((St = TmpArea) And (strSaleName2 <> StrR050302_3(5)))) And (iLine <> 0) Then
                If txt1(10) = "1" Then
                  strSaleName2 = .Fields(5)
                End If
                PriTiEnd3
                Printer.NewPage
                Page = Page + 1
                PriTiTle3 St, str(Page)
                iPrint = 2400
                iLine = 0
            End If
            '20140304END
       
            iLine = iLine + 1
            iPrint = iPrint + 300
        End If
       ' If Page = 2 And Iline = 22 Then
       '     PriTiEnd3
       '     Printer.EndDoc
       '     Exit Sub
       ' End If
        
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
'Modify By Cheng 2002/03/15
'Printer.Print "本所期限管制表"
Printer.Print IIf(intPWhere = 國外_CF, "CFP ", "") & "本所期限管制表"
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = 6000
Printer.CurrentY = k + 500
Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@") & "-" & ChangeTStringToTDateString(txt1(2))
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

'Add by Morgan 2004/2/27
'加印特殊標示說明
Printer.CurrentX = 5000
Printer.CurrentY = k + 1100
Printer.Print " *:本所期限逾期   #:C 類來函未發文   V:當日本所期限"
'Add end 2004/2/27

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
Printer.Print "CFP 法定期限管制表"
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = 6000
Printer.CurrentY = k + 500
Printer.Print "法定期限：" & Format(ChangeTStringToTDateString(txt1(3)) & " ", "@@@@@@@@@") & "-" & ChangeTStringToTDateString(txt1(4))
Printer.Font.Bold = False
Printer.CurrentX = 0
Printer.CurrentY = k + 800
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = k + 800
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
'Printer.CurrentX = 500
'Printer.CurrentY = k + 1100
'Printer.Print "業務區：" & Area

'Add by Morgan 2004/2/27
'加印特殊標示說明
Printer.CurrentX = 5000
Printer.CurrentY = k + 1100
Printer.Print " *:本所期限逾期   #:C 類來函未發文   V:當日本所期限"
'Add end 2004/2/27

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
'Modify By Cheng 2002/03/15
'Printer.Print "本所期限管制表"
Printer.Print IIf(intPWhere = 國外_CF, "CFP ", "") & "本所期限管制表"
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = 6000
Printer.CurrentY = k + 500
Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@") & "-" & ChangeTStringToTDateString(txt1(2))
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

'Add by Morgan 2004/2/27
'加印特殊標示說明
Printer.CurrentX = 5000
Printer.CurrentY = k + 1100
Printer.Print " *:本所期限逾期   #:C 類來函未發文   V:當日本所期限"
'Add end 2004/2/27

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
    PLeft3(3) = 3850
    PLeft3(4) = 5750
    PLeft3(5) = 10050
    PLeft3(6) = 11250
    PLeft3(7) = 13250
    PLeft3(8) = 15250
End Sub
Sub GetPrintLeft2()              '設定定位點
    Erase Pleft2
    Pleft2(0) = 0
    'Modify By Cheng 2002/03/15
'    PLeft2(1) = 1100
'    PLeft2(2) = 2200
'    PLeft2(3) = 3200
'    PLeft2(4) = 5100
'    PLeft2(5) = 8100
'    PLeft2(6) = 9400
'    PLeft2(7) = 10400
'    PLeft2(8) = 11200
'    PLeft2(9) = 12200
'    PLeft2(10) = 13400
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
   'Modify By Cheng 2002/03/15
'    PLeft1(1) = 1600
'    PLeft1(2) = 2700
'    PLeft1(3) = 4600
'    PLeft1(4) = 8900
'    PLeft1(5) = 10100
'    PLeft1(6) = 11500
'    PLeft1(7) = 13100
    PLeft1(1) = 1600 + 100
    PLeft1(2) = 2750 + 100
    PLeft1(3) = 4650 + 100
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
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
   End If

Case 5
     If txt1(Index) <> "1" And txt1(Index) <> "2" Then
        s = MsgBox("管制對象必須輸入 1 或 2 !!", , "USER 輸入錯誤")
        Cancel = True
     End If
Case 6
   If txt1(Index) <> "" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetStaff(txt1(Index), strExc(0)) Then
      If ClsPDGetStaffN(txt1(Index), strExc(0)) Then
         lbl1(0) = strExc(0)
      Else
         lbl1(0) = ""
         Cancel = True
      End If
   End If
Case 7
     If Len(txt1(7)) <> 0 Then
      '2009/12/21 modify by sonia 以第一個系統類別抓,否則CFP的607會錯誤
      'lbl1(1) = GetPrjState6HM("P", txt1(Index))
      StrTempP = Split(Replace(txt1(0), ",,", ""), ",")
      StrSQL6 = StrTempP(0)
      lbl1(1) = GetPrjState6HM(StrSQL6, txt1(Index))
      '2009/12/21 end
      If lbl1(1) = "" Then
         MsgBox "案件性質錯誤，請重新輸入 !", vbCritical
         Cancel = True
      End If
     End If
'Add by Morgan 2006/5/30
Case 10
      Select Case Val(txt1(Index))
      Case 1, 2
      Case Else
         s = MsgBox("列印對象只能輸入 1,2 !!", , "USER 輸入錯誤")
         Cancel = True
      End Select
Case Else

End Select
If Cancel Then TextInverse txt1(Index)
End Sub
'Add by Morgan 2005/2/14 加PCT進入國家階段條件
Private Sub txtPA46_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtPA46.IMEMode = 2
   CloseIme
   TextInverse txtPA46
End Sub

Private Sub txtPA46_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Chr(KeyAscii) <> "Y" Then
      KeyAscii = 0
      Beep
   End If
End Sub
