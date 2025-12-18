VERSION 5.00
Begin VB.Form frm040207 
   BorderStyle     =   1  '單線固定
   Caption         =   "延展前無效商標管制表"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4260
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   20
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1890
      Width           =   285
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   0
      Left            =   1140
      TabIndex        =   0
      Top             =   450
      Width           =   2340
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   3
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   1
      Top             =   765
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   4
      Left            =   2295
      MaxLength       =   7
      TabIndex        =   2
      Top             =   765
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   5
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1035
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   6
      Left            =   2295
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1035
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   7
      Left            =   1140
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1320
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   8
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1605
      Width           =   285
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   9
      Left            =   1140
      MaxLength       =   9
      TabIndex        =   8
      Top             =   2175
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   10
      Left            =   2280
      MaxLength       =   9
      TabIndex        =   9
      Top             =   2175
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   11
      Left            =   1140
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2460
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   12
      Left            =   2280
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2460
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   13
      Left            =   1140
      MaxLength       =   4
      TabIndex        =   12
      Top             =   2745
      Width           =   1005
   End
   Begin VB.TextBox TXT1 
      Height          =   264
      Index           =   14
      Left            =   2295
      MaxLength       =   4
      TabIndex        =   13
      Top             =   2745
      Width           =   1005
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2655
      TabIndex        =   14
      Top             =   30
      Width           =   756
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3465
      TabIndex        =   15
      Top             =   30
      Width           =   756
   End
   Begin VB.Label Label1 
      Caption         =   "查詢內容："
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   27
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(1.爭議無效  2.未繳第二期  3.全部)"
      Height          =   180
      Index           =   1
      Left            =   1515
      TabIndex        =   26
      Top             =   1920
      Width           =   2685
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   25
      Top             =   495
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "本所期限："
      Height          =   180
      Index           =   3
      Left            =   60
      TabIndex        =   24
      Top             =   780
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   23
      Top             =   1065
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   5
      Left            =   60
      TabIndex        =   22
      Top             =   1335
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "查詢種類："
      Height          =   180
      Index           =   6
      Left            =   60
      TabIndex        =   21
      Top             =   1635
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   180
      Index           =   7
      Left            =   60
      TabIndex        =   20
      Top             =   2220
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      Height          =   180
      Index           =   8
      Left            =   60
      TabIndex        =   19
      Top             =   2490
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   9
      Left            =   60
      TabIndex        =   18
      Top             =   2775
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "(1.查詢  2.報表)"
      Height          =   180
      Index           =   11
      Left            =   1515
      TabIndex        =   17
      Top             =   1635
      Width           =   1620
   End
   Begin VB.Label LBL1 
      Height          =   180
      Left            =   2310
      TabIndex        =   16
      Top             =   1335
      Width           =   1125
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1710
      X2              =   2850
      Y1              =   915
      Y2              =   915
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1635
      X2              =   2775
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1605
      X2              =   2745
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   1710
      X2              =   2850
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   1515
      X2              =   2655
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "frm040207"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 4) As String, strTemp3 As String
Dim PLeft(0 To 4) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String
Dim StrSQL3 As String  '2010/3/2 add by sonia

Private Sub cmdok_Click(Index As Integer)
Dim ii As Integer
On Error GoTo ErrorHandler

Select Case Index
Case 0 '確定
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
         strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
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
         If Len(txt1(4)) = 0 Then
             s = MsgBox("本所期限區間不可空白!!", , "USER 輸入錯誤")
             txt1(3).SetFocus
             txt1_GotFocus (3)
             Exit Sub
         Else
             If Len(txt1(8)) = 0 Then
                 s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
                 txt1(8).SetFocus
                 Exit Sub
             Else
                '列印管制表
                If Mid(txt1(9), 1, 6) <> Mid(txt1(10), 1, 6) Then
                    s = MsgBox("申請人代號前六碼必須相同!!", , "USER 輸入錯誤")
                    txt1(9).SetFocus
                    txt1(9).SelStart = 0
                    txt1(9).SelLength = Len(txt1(9))
                    Exit Sub
                End If
                If Mid(txt1(11), 1, 6) <> Mid(txt1(12), 1, 6) Then
                    s = MsgBox("代理人代號前六碼必須相同!!", , "USER 輸入錯誤")
                    txt1(11).SetFocus
                    txt1(11).SelStart = 0
                    txt1(11).SelLength = Len(txt1(11))
                    Exit Sub
                End If
                ClearQueryLog (Me.Name) 'Add By Sindy 2010/9/30 清除查詢印表記錄檔欄位
                Screen.MousePointer = vbHourglass
                Me.Enabled = False
                Process
                Me.Enabled = True
                Screen.MousePointer = vbDefault
             End If
         End If
     End If
Case 1 '結束
    Me.Enabled = False
    Unload Me
Case Else
End Select
Exit Sub
ErrorHandler:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    MsgBox "(" & Err.Number & ")" & Err.Description

End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040207 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
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
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
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
Case 4
      If Not nickChgRan(txt1(3), txt1(4), "本所期限") Then
         txt1(3).SetFocus
         txt1_GotFocus (3)
         Exit Sub
      End If
Case 6
     If Not nickChgRan(txt1(5), txt1(6), "業務區") Then

         txt1(5).SetFocus
         txt1_GotFocus (5)
         Exit Sub
     End If
Case 7
     If txt1(7) <> "" Then
      LBL1 = GetPrjSalesNM(txt1(7))
      If LBL1.Caption = "" Then
           s = MsgBox("智權人員錯誤！", , "錯誤！")
           txt1(7).SetFocus
           txt1_GotFocus (7)
           Exit Sub
       End If
      End If
Case 8
     Select Case txt1(8)
     Case "1", "2", "", " "
     Case Else
          s = MsgBox("查詢種類只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          txt1(8).SetFocus
          txt1(8).SelStart = 0
          txt1(8).SelLength = Len(txt1(8))
          Exit Sub
     End Select
Case 10
     If Not nickChgRan(txt1(9), txt1(10), "申請人") Then
         txt1(9).SetFocus
         txt1_GotFocus (9)
         Exit Sub
     End If
Case 12
     If Not nickChgRan(txt1(11), txt1(12), "代理人") Then
         txt1(11).SetFocus
         txt1_GotFocus (11)
         Exit Sub
      End If
Case 14
      If Not nickChgRan(txt1(13), txt1(14), "申請國家") Then
          txt1(13).SetFocus
          txt1_GotFocus (13)
          Exit Sub
       End If
Case 20
     Select Case txt1(20)
     Case "1", "2", "3", "", " "
     Case Else
          s = MsgBox("查詢內容只能輸入 1 、2 或 3 !!", , "USER 輸入錯誤")
          txt1(16).SetFocus
          txt1(16).SelStart = 0
          txt1(16).SelLength = Len(txt1(16))
          Exit Sub
     End Select
Case Else
End Select
End Sub
Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case 3
        If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Cancel = True
        End If
    Case 4
        If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Cancel = True
        End If
    End Select
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
' 暫存   id+1 = 爭議無效 ；id+2 =未繳第二期
cnnConnection.Execute "DELETE FROM R040207 WHERE ID='" & strUserNum & "1" & "' or id='" & strUserNum & "2" & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
If txt1(8) = "1" Then pub_QL05 = pub_QL05 & ";" & Label1(6) & "查詢" 'Add By Sindy 2010/9/30
If txt1(8) = "2" Then pub_QL05 = pub_QL05 & ";" & Label1(6) & "報表" 'Add By Sindy 2010/9/30
'系統類別
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " and NP02 in (" & SQLGrpStr(txt1(0), 2) & ") "
   strSQL2 = strSQL2 + " and NP02 in (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0)  'Add By Sindy 2010/9/30
End If
'申請人
If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) <> 0 Then
    strSQL1 = strSQL1 & " AND (TM23>='" & GetNewFagent(txt1(9)) & "' AND TM23<='" & GetNewFagent(txt1(10)) & "') "
    strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(txt1(9)) & "' AND SP08<='" & GetNewFagent(txt1(10)) & "') OR (SP58<='" & GetNewFagent(txt1(9)) & "' AND SP58<='" & GetNewFagent(txt1(10)) & "') OR (SP59>='" & GetNewFagent(txt1(9)) & "' AND SP59<='" & GetNewFagent(txt1(10)) & "')) "
Else
    If Len(Trim(txt1(9))) <> 0 And Len(Trim(txt1(10))) = 0 Then
        strSQL1 = strSQL1 & " AND (TM23>='" & GetNewFagent(txt1(9)) & "' ) "
        strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(9)) & "' OR SP58>='" & GetNewFagent(txt1(9)) & "' OR SP59>='" & GetNewFagent(txt1(9)) & "') "
    Else
        If Len(Trim(txt1(9))) = 0 And Len(Trim(txt1(10))) <> 0 Then
            strSQL1 = strSQL1 & " AND (TM23<='" & GetNewFagent(txt1(10)) & "') "
            strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(10)) & "' OR SP58<='" & GetNewFagent(txt1(10)) & "' OR SP59<='" & GetNewFagent(txt1(10)) & "') "
        End If
    End If
End If
If Len(Trim(txt1(9))) <> 0 Or Len(Trim(txt1(10))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(7) & Trim(txt1(9)) & "-" & Trim(txt1(10)) 'Add By Sindy 2010/9/30
End If
'代理人
If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) <> 0 Then
    strSQL1 = strSQL1 + " AND (TM44>='" & GetNewFagent(txt1(11)) & "' AND TM44<='" & GetNewFagent(txt1(12)) & "') "
    strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(txt1(11)) & "' AND SP26<='" & GetNewFagent(txt1(12)) & "') "
Else
    If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) = 0 Then
        strSQL1 = strSQL1 + " AND (TM44>='" & GetNewFagent(txt1(11)) & "' ) "
        strSQL2 = strSQL2 + " AND (SP26>='" & GetNewFagent(txt1(11)) & "' ) "
    Else
        If Len(Trim(txt1(11))) = 0 And Len(Trim(txt1(12))) <> 0 Then
            strSQL1 = strSQL1 + " AND (TM44<='" & GetNewFagent(txt1(12)) & "' ) "
            strSQL2 = strSQL2 + " AND (SP26<='" & GetNewFagent(txt1(12)) & "' ) "
        End If
    End If
End If
If Len(Trim(txt1(11))) <> 0 Or Len(Trim(txt1(12))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(8) & Trim(txt1(11)) & "-" & Trim(txt1(12)) 'Add By Sindy 2010/9/30
End If
'申請國家
If Len(txt1(13)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10>='" & txt1(13) & "' "
    strSQL2 = strSQL2 + " AND SP09>='" & txt1(13) & "' "
End If
If Len(txt1(14)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10<='" & txt1(14) & "' "
    strSQL2 = strSQL2 + " AND SP09<='" & txt1(14) & "' "
End If
If Len(Trim(txt1(13))) <> 0 Or Len(Trim(txt1(14))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(9) & Trim(txt1(13)) & "-" & Trim(txt1(14)) 'Add By Sindy 2010/9/30
End If
strSQL1 = strSQL1 + " AND (TM29<>'Y' OR TM29 IS NULL) "
strSQL2 = strSQL2 + " AND (SP15<>'Y' OR SP15 IS NULL) "

'add by nickc 2006/05/30 延展和第二期專用權須存在
strSQL1 = strSQL1 & " and decode(np07,102,tm17,'Y')='Y' "
'2010/3/2 modify by sonia 加被異議續展不必判斷專用權是否存在,所以加StrSQL3
'2010/3/19 cancel by sonia 取消被異議續展109改在期限管制表做,故還原上一句
'StrSQL3 = StrSQL3 & " and decode(np07,102,tm17,'Y')='Y' "

If Process2 = False Then
    InsertQueryLog (0) 'Add By Sindy 2010/9/30
    ShowNoData
    Exit Sub
End If
Select Case Trim(txt1(8))
Case "1"
         Me.Hide
         frm040207a.Show
Case "2"
        PrintData
Case Else
End Select
End Sub

Function Process2() As Boolean          '延展,刊登廣告,第一期註冊費,第二期註冊費,繳年費     只有 102,702,715,716,708
StrSQL6 = ""

'2010/3/2 cancel by sonia
'2010/3/19 取消被異議續展109故還原 by sonia
StrSQL6 = " And NP07=102 "

If Len(txt1(3)) <> 0 Then
   StrSQL6 = StrSQL6 + " AND NP08>=" & Val(ChangeTStringToWString(txt1(3))) & " "
End If
If Len(txt1(4)) <> 0 Then
   StrSQL6 = StrSQL6 + " AND NP08<=" & Val(ChangeTStringToWString(txt1(4))) & " "
End If
If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/9/30
End If
StrSQL6 = StrSQL6 & " AND (NP06 IS NULL OR NP06='') "
If Len(txt1(5)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND s1.ST15>='" & txt1(5) & "' "
End If
If Len(txt1(6)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND s1.ST15<='" & txt1(6) & "' "
End If
If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & "-" & txt1(6)  'Add By Sindy 2010/9/30
End If
If Len(txt1(7)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND NP10='" & txt1(7) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(5) & txt1(7) & LBL1 'Add By Sindy 2010/9/30
End If
strSql = ""
If txt1(20) = "1" Then pub_QL05 = pub_QL05 & ";" & Label1(2) & "爭議無效" 'Add By Sindy 2010/9/30
If txt1(20) = "2" Then pub_QL05 = pub_QL05 & ";" & Label1(2) & "未繳第二期" 'Add By Sindy 2010/9/30
If txt1(20) = "3" Then pub_QL05 = pub_QL05 & ";" & Label1(2) & "全部" 'Add By Sindy 2010/9/30
If txt1(20) = "1" Or txt1(20) = "3" Then
    '2010/3/2 modify by sonia 被異議續展不必判斷專用權是否存在,所以拆成二句
    strSql = "SELECT NP02,nP03,NP04,NP05,TM15,tm12,NVL(TM05,NVL(TM06,TM07)),TM23,'" & strUserNum & "1' " & _
             " FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & _
             " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND " & _
             "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND " & _
             "decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 and tm19='Y' " & StrSQL6 & strSQL1
    '2010/3/19 cancel by sonia 取消被異議續展109改在期限管制表做,故還原上一句
    'strSql = "SELECT NP02,nP03,NP04,NP05,TM15,tm12,NVL(TM05,NVL(TM06,TM07)),TM23,'" & strUserNum & "1' " & _
             " FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & _
             " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND " & _
             "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND " & _
             "decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 and tm19='Y' and np07=102 " & StrSQL6 & strSQL1 & StrSQL3 & " union " & _
             "SELECT NP02,nP03,NP04,NP05,TM15,tm12,NVL(TM05,NVL(TM06,TM07)),TM23,'" & strUserNum & "1' " & _
             " FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & _
             " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND " & _
             "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND " & _
             "decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 and tm19='Y' and np07=109 " & StrSQL6 & strSQL1 & StrSQL3
End If
If txt1(20) = "2" Or txt1(20) = "3" Then
    If strSql <> "" Then
        strSql = strSql & " union all "
    End If
    strSql = strSql & "SELECT NP02,nP03,NP04,NP05,TM15,tm12,NVL(TM05,NVL(TM06,TM07)),TM23,'" & strUserNum & "2' " & _
             " FROM Staff_Group, NEXTPROGRESS,TRADEMARK,CASEPROGRESS,NATION,STAFF s1,STAFF S2,CASEPROPERTYMAP,CUSTOMER " & _
             " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=s1.ST01(+) AND " & _
             "NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND CP14=S2.ST01(+) AND SUBSTR(TM23,1,8)=CU01(+) AND " & _
             "decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND TM10=NA01(+) And '" & strGroup & "'=SG01 And SG02=NP02 And SG03=NP07 " & StrSQL6 & strSQL1
End If
Dim iCount
cnnConnection.Execute "insert into r040207 " & strSql, iCount
If iCount = 0 Then
    ShowNoData
    CheckOC
    Process2 = False
    Exit Function
End If
CheckOC
With adoRecordset
    If txt1(20) = "1" Or txt1(20) = "3" Then
        .CursorLocation = adUseClient
        .Open "select * from r040207 where id='" & strUserNum & "1' ", cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    '檢查 602、604、606、1601-1606 發文日最大且有准駁ㄉ
                    CheckOC3
                    AdoRecordSet3.CursorLocation = adUseClient
                    'modify by sonia 2017/9/1 +624,1619,1620
                    AdoRecordSet3.Open "select * from caseprogress where to_char(cp27)||cp09 in (select max(to_char(cp27)||cp09) from caseprogress where cp01='" & CheckStr(.Fields(0)) & "' and cp02='" & CheckStr(.Fields(1)) & "' and cp03='" & CheckStr(.Fields(2)) & "' and cp04='" & CheckStr(.Fields(3)) & "' and cp10 in ('602','604','606','624','1601','1602','1603','1604','1605','1606','1619','1620') and cp24 is not null) and cp01='" & CheckStr(.Fields(0)) & "' and cp02='" & CheckStr(.Fields(1)) & "' and cp03='" & CheckStr(.Fields(2)) & "' and cp04='" & CheckStr(.Fields(3)) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
                    If Not AdoRecordSet3.EOF And Not AdoRecordSet3.BOF Then
                        If CheckStr(AdoRecordSet3.Fields("cp24")) = "2" Then
                            Process2 = True
                        Else
                            '不是駁的，不管
                            cnnConnection.Execute "delete from r040207 where id='" & strUserNum & "1' and r116001='" & CheckStr(.Fields(0)) & "' and r116002='" & CheckStr(.Fields(1)) & "' and r116003='" & CheckStr(.Fields(2)) & "' and r116004='" & CheckStr(.Fields(3)) & "' "
                        End If
                    Else
                        '若都沒有准駁，不管
                        cnnConnection.Execute "delete from r040207 where id='" & strUserNum & "1' and r116001='" & CheckStr(.Fields(0)) & "' and r116002='" & CheckStr(.Fields(1)) & "' and r116003='" & CheckStr(.Fields(2)) & "' and r116004='" & CheckStr(.Fields(3)) & "' "
                    End If
                    .MoveNext
                Loop
       End If
    End If
    If txt1(20) = "2" Or txt1(20) = "3" Then
        CheckOC
        .CursorLocation = adUseClient
        .Open "select * from r040207 where id='" & strUserNum & "2' ", cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    '檢查 716 法定期限最大
                    CheckOC3
                    AdoRecordSet3.CursorLocation = adUseClient
                    AdoRecordSet3.Open "select * from nextprogress where to_char(np09)||np01 in (select max(to_char(np09)||np01) from nextprogress where np02='" & CheckStr(.Fields(0)) & "' and np03='" & CheckStr(.Fields(1)) & "' and np04='" & CheckStr(.Fields(2)) & "' and np05='" & CheckStr(.Fields(3)) & "' and np07=716 ) and np02='" & CheckStr(.Fields(0)) & "' and np03='" & CheckStr(.Fields(1)) & "' and np04='" & CheckStr(.Fields(2)) & "' and np05='" & CheckStr(.Fields(3)) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
                    If Not AdoRecordSet3.EOF And Not AdoRecordSet3.BOF Then
                        If CheckStr(AdoRecordSet3.Fields("np06")) = "N" Then
                            Process2 = True
                        Else
                            '續辦ㄉ，不管
                            '2011/9/30 MODIFY BY SONIA
                            'cnnConnection.Execute "delete from r040207 where id='" & strUserNum & "2' and r116001='" & CheckStr(.Fields(0)) & "' and r116002='" & CheckStr(.Fields(1)) & "' and r116003='" & CheckStr(.Fields(2)) & "' and r116004='" & CheckStr(.Fields(3)) & "' "
                            If CheckStr(AdoRecordSet3.Fields("np06")) = "Y" Then
                               cnnConnection.Execute "delete from r040207 where id='" & strUserNum & "2' and r116001='" & CheckStr(.Fields(0)) & "' and r116002='" & CheckStr(.Fields(1)) & "' and r116003='" & CheckStr(.Fields(2)) & "' and r116004='" & CheckStr(.Fields(3)) & "' "
                            Else
                               Process2 = True
                            End If
                            '2011/9/30 END
                        End If
                    Else
                        '若都沒有716，不管
                        cnnConnection.Execute "delete from r040207 where id='" & strUserNum & "2' and r116001='" & CheckStr(.Fields(0)) & "' and r116002='" & CheckStr(.Fields(1)) & "' and r116003='" & CheckStr(.Fields(2)) & "' and r116004='" & CheckStr(.Fields(3)) & "' "
                    End If
                    .MoveNext
                Loop
       End If
    End If
End With
CheckOC
CheckOC3
End Function

Private Sub PrintData()
Dim ii As Integer
Dim jj As Integer

For jj = IIf(txt1(20) = "3", 1, Val(txt1(20))) To IIf(txt1(20) = "3", 2, Val(txt1(20)))
    CheckOC
    With adoRecordset
            strSql = "select r116001||'-'||r116002||'-'||r116003||'-'||r116004 as 本所案號,r116005 as 註冊號,r116006 as 申請案號,r116007 as 商標名稱,r116008||' '||NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) as 申請人 from r040207,customer where id='" & strUserNum & Trim(jj) & "'  and substr(r116008,1,8)=cu01(+) and decode(SUBSTR(r116008,9,1),'','0',SUBSTR(r116008,9,1))=CU02(+) order by 1,2 "
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount <> 0 And .RecordCount > 0 Then
                    InsertQueryLog (.RecordCount) 'Add By Sindy 2010/9/30
                    Page = 1
                    PrintTitle jj
                    .MoveFirst
                    Do While Not .EOF
                        '本所案號
                        Printer.CurrentX = PLeft(0)
                        Printer.CurrentY = iPrint
                        Printer.Print CheckStr(.Fields(0))
                        '註冊號
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = iPrint
                        Printer.Print StrToStr(CheckStr(.Fields(1)), 5)
                        '申請案號
                        Printer.CurrentX = PLeft(2)
                        Printer.CurrentY = iPrint
                        Printer.Print StrToStr(CheckStr(.Fields(2)), 5)
                        '商標名稱
                        Printer.CurrentX = PLeft(3)
                        Printer.CurrentY = iPrint
                        Printer.Print StrToStr(CheckStr(.Fields(3)), 10)
                        '申請人
                        Printer.CurrentX = PLeft(4)
                        Printer.CurrentY = iPrint
                        Printer.Print StrToStr(CheckStr(.Fields(4)), 17)
                        iPrint = iPrint + 300
                        If iPrint > 15000 Then
                            Printer.CurrentX = 0
                            Printer.CurrentY = iPrint
                            Printer.Print String(200, "-")
                            Printer.NewPage
                            Page = Page + 1
                            PrintTitle jj
                        End If
                        .MoveNext
                Loop
                Printer.CurrentX = 0
                Printer.CurrentY = iPrint
                Printer.Print String(200, "-")
                Printer.EndDoc
            Else
                InsertQueryLog (0) 'Add By Sindy 2010/9/30
            End If
    End With
Next jj
ShowPrintOk
End Sub

Sub PrintTitle(oKind As Integer)
GetPleft
iPrint = 500
Printer.Orientation = 1
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("延展前無效商標管制表") / 2
Printer.CurrentY = iPrint
Printer.Print "延展前無效商標管制表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("本所期限：" & Format(ChangeTStringToTDateString(Me.txt1(3).Text) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Me.txt1(4).Text)) / 2
Printer.CurrentY = iPrint
Printer.Print "本所期限：" & Format(ChangeTStringToTDateString(Me.txt1(3).Text) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Me.txt1(4).Text)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 9300
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印種類：" & IIf(oKind = 1, "爭議無效", "第二期註冊費未委託本所繳費")
Printer.CurrentX = 9300
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "註冊號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "申請案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "商標名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "申請人"
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1750
PLeft(2) = 3100
PLeft(3) = 4450
PLeft(4) = 7000
End Sub
