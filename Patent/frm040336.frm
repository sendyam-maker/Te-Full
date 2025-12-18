VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm040336 
   BorderStyle     =   1  '單線固定
   Caption         =   "資策會專利案件季報表"
   ClientHeight    =   3300
   ClientLeft      =   3696
   ClientTop       =   1560
   ClientWidth     =   7404
   ControlBox      =   0   'False
   LinkTopic       =   "Form16"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7404
   Begin VB.CheckBox Check1 
      Caption         =   "顯示代表圖"
      Height          =   210
      Left            =   165
      TabIndex        =   31
      Top             =   210
      Width           =   1300
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   0
      Top             =   450
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   14
      Left            =   6285
      MaxLength       =   7
      TabIndex        =   4
      Top             =   780
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   13
      Left            =   4950
      MaxLength       =   7
      TabIndex        =   3
      Top             =   780
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   810
      Index           =   12
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   18
      Top             =   2340
      Width           =   6075
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   5715
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1140
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1110
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1140
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1455
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2475
      MaxLength       =   7
      TabIndex        =   2
      Top             =   765
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1140
      MaxLength       =   7
      TabIndex        =   1
      Top             =   765
      Width           =   800
   End
   Begin VB.CheckBox Check2 
      Caption         =   "不含"
      Height          =   210
      Left            =   210
      TabIndex        =   17
      Top             =   2610
      Width           =   675
   End
   Begin VB.CheckBox Check5 
      Caption         =   "專用期仍有效的"
      Height          =   210
      Left            =   5640
      TabIndex        =   11
      Top             =   1485
      Width           =   1560
   End
   Begin VB.CheckBox ChkClose 
      Caption         =   "含閉卷"
      Height          =   210
      Left            =   3390
      TabIndex        =   9
      Top             =   1485
      Value           =   1  '核取
      Width           =   1110
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2475
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1455
      Width           =   800
   End
   Begin VB.CheckBox Check6 
      Caption         =   "含銷卷"
      Height          =   210
      Left            =   4515
      TabIndex        =   10
      Top             =   1485
      Value           =   1  '核取
      Width           =   1110
   End
   Begin VB.CheckBox Check7 
      Caption         =   "只印申請案"
      Height          =   210
      Left            =   3555
      TabIndex        =   14
      Top             =   1770
      Width           =   1200
   End
   Begin VB.CheckBox Check8 
      Caption         =   "只印已核准"
      Height          =   210
      Left            =   4815
      TabIndex        =   15
      Top             =   1770
      Width           =   1200
   End
   Begin VB.CheckBox Check9 
      Caption         =   "只印未核准"
      Height          =   210
      Left            =   6075
      TabIndex        =   16
      Top             =   1770
      Width           =   1200
   End
   Begin VB.CheckBox Check11 
      Caption         =   "已閉卷或銷卷"
      Height          =   210
      Left            =   2115
      TabIndex        =   13
      Top             =   1770
      Width           =   1380
   End
   Begin VB.CheckBox Check12 
      Caption         =   "已閉卷或銷卷或核駁"
      Height          =   210
      Left            =   135
      TabIndex        =   12
      Top             =   1770
      Width           =   1920
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   7860
      Top             =   96
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   6540
      TabIndex        =   20
      Top             =   30
      Width           =   756
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5760
      TabIndex        =   19
      Top             =   30
      Width           =   756
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "程式執行中請勿開啟Excel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2280
      TabIndex        =   33
      Top             =   90
      Width           =   2895
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "本所案號請勿輸入「-」 若多個案號請以「,] 區隔　ex:P012345,CFP012345100"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   1110
      TabIndex        =   32
      Top             =   2130
      Width           =   6030
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請日期:"
      Height          =   180
      Index           =   4
      Left            =   3960
      TabIndex        =   30
      Top             =   810
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收文日期："
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   29
      Top             =   780
      Width           =   900
   End
   Begin VB.Line Line4 
      X1              =   5925
      X2              =   6165
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   2385
      Width           =   900
   End
   Begin VB.Line Line3 
      X1              =   2115
      X2              =   2355
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Line Line2 
      X1              =   2115
      X2              =   2355
      Y1              =   885
      Y2              =   885
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "(1.本所案號 2.案件名稱 3.申請國家+本所案號 4.申請國家+案件名稱)"
      Height          =   180
      Left            =   1665
      TabIndex        =   27
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "輸出順序："
      Height          =   180
      Index           =   1
      Left            =   165
      TabIndex        =   26
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "(Y:分開)"
      Height          =   180
      Left            =   6195
      TabIndex        =   25
      Top             =   1170
      Width           =   645
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "國內外是否分開列印:"
      Height          =   180
      Left            =   3960
      TabIndex        =   24
      Top             =   1170
      Width           =   1665
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "(N:不含)"
      Height          =   180
      Index           =   0
      Left            =   1650
      TabIndex        =   23
      Top             =   1155
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "是否含核駁："
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   22
      Top             =   1155
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   150
      TabIndex        =   21
      Top             =   1485
      Width           =   900
   End
End
Attribute VB_Name = "frm040336"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2021/09/10
Option Explicit

'Const msoTrue = -1 'Mark by Amy 2023/07/27 PutXlsImg函數改至basUpdate
Dim RsQ As New ADODB.Recordset
Dim arrCaseNo, strField, intWidth
Dim i As Integer, ii As Integer, j As Integer, intQ As Integer, intRow As Integer, intField As Integer, intTitleR As Integer
Dim bolOpenXls As Boolean '已開啟Excel
Dim strQ As String, strFileN As String, strAllF As String, strWidth As String
Dim strCusNo_S As String, strCusNo_E As String, strSalesNo As String, m_IsOrNot As String '客戶編號起/迄/業務員/是否含指定案號

Private Sub cmdOK_Click(Index As Integer)
    Select Case Index
        Case 0 '確定
            If FormCheck = False Then Exit Sub
            Screen.MousePointer = vbHourglass
            Call SetQueryLog
            Call ReadData
            Screen.MousePointer = vbDefault
        Case 1 '結束
            Unload Me
    End Select
End Sub

Private Function FormCheck() As Boolean
    Dim bCancel As Boolean
    
    FormCheck = False
    '輸入順序
    If Trim(txt1(9)) = MsgText(601) Then
        MsgBox Replace(Label1(1).Caption, ":", "") & "不可為空"
        Exit Function
    End If
    '收文日
    If Trim(txt1(3)) <> MsgText(601) Then
        Call txt1_Validate(3, bCancel)
        If bCancel = True Then
            Exit Function
        End If
    End If
    If Trim(txt1(4)) <> MsgText(601) Then
        Call txt1_Validate(4, bCancel)
        If bCancel = True Then
            Exit Function
        End If
    End If
    If Trim(txt1(3)) <> MsgText(601) And Trim(txt1(4)) <> MsgText(601) Then
        If Val(txt1(3)) > Val(txt1(4)) Then
            MsgBox Replace(Label2(3).Caption, ":", "") & "迄日不可大於止日"
            Exit Function
        End If
    End If
    '申請日
    If Trim(txt1(13)) <> MsgText(601) Then
        Call txt1_Validate(13, bCancel)
        If bCancel = True Then
            Exit Function
        End If
    End If
    If Trim(txt1(14)) <> MsgText(601) Then
        Call txt1_Validate(14, bCancel)
        If bCancel = True Then
            Exit Function
        End If
    End If
    If Trim(txt1(13)) <> MsgText(601) And Trim(txt1(14)) <> MsgText(601) Then
        If Val(txt1(13)) > Val(txt1(14)) Then
            MsgBox Replace(Label2(4).Caption, ":", "") & "迄日不可大於止日"
            Exit Function
        End If
    End If
    '本所案號
    If Trim(txt1(12)) <> MsgText(601) Then
        If InStr(txt1(12), "-") > 0 Then
            MsgBox Replace(Label2(2).Caption, ":", "") & "不可輸入「-」符號"
            Exit Function
        End If
    End If
    
    FormCheck = True
End Function

Private Sub SetQueryLog()
    ClearQueryLog (Me.Name) '清除查詢印表記錄檔欄位
    '顯示代表圖
    If Check1.Value = vbChecked Then pub_QL05 = pub_QL05 & ";" & Check1.Caption
    '輸出順序
    If Len(Trim(txt1(9))) <> 0 Then pub_QL05 = pub_QL05 & ";" & Label1(1) & Trim(txt1(9)) & Label11
    '收文日期
    If Len(Trim(txt1(3))) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
        pub_QL05 = pub_QL05 & ";" & Label2(3) & Trim(txt1(3)) & "-" & Trim(txt1(4))
    End If
    '申請日期
    If Len(Trim(txt1(13))) <> 0 Or Len(Trim(txt1(14))) <> 0 Then
        pub_QL05 = pub_QL05 & ";" & Label2(4) & Trim(txt1(13)) & "-" & Trim(txt1(14))
    End If
    '是否含核駁
    If Len(Trim(txt1(7))) <> 0 Then pub_QL05 = pub_QL05 & ";" & Label6(0) & Trim(txt1(7)) & Label7(0)
    '國內外是否分開列印
    If Len(Trim(txt1(8))) <> 0 Then pub_QL05 = pub_QL05 & ";" & Label8 & Trim(txt1(8)) & Label9
    '申請國家
    If Len(Trim(txt1(5))) <> 0 Or Len(Trim(txt1(6))) <> 0 Then
        pub_QL05 = pub_QL05 & ";" & Label4 & Trim(txt1(5)) & "-" & Trim(txt1(6))
    End If
    '含閉卷
    If ChkClose.Value = vbChecked Then pub_QL05 = pub_QL05 & ";" & ChkClose.Caption
    '含銷卷
    If Check6.Value = vbChecked Then pub_QL05 = pub_QL05 & ";" & Check6.Caption
    '專用期仍有效的
    If Check5.Value = vbChecked Then pub_QL05 = pub_QL05 & ";" & Check5.Caption
    '已閉卷或銷卷或核駁
    If Check12.Value = vbChecked Then pub_QL05 = pub_QL05 & ";" & Check12.Caption
    '已閉卷或銷卷
    If Check11.Value = vbChecked Then pub_QL05 = pub_QL05 & ";" & Check11.Caption
    '只印申請案
    If Check7.Value = vbChecked Then pub_QL05 = pub_QL05 & ";" & Check7.Caption
    '只印已核准
    If Check8.Value = vbChecked Then pub_QL05 = pub_QL05 & ";" & Check8.Caption
    '只印未核淮
    If Check9.Value = vbChecked Then pub_QL05 = pub_QL05 & ";" & Check9.Caption
    '本所案號
    If Len(Trim(txt1(12))) <> 0 Then pub_QL05 = pub_QL05 & ";" & Label2(2) & txt1(12)
    '不含
    If Check2.Value = vbChecked Then pub_QL05 = pub_QL05 & ";" & "不含指定案號"
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    strCusNo_S = "X38805030"
    strCusNo_E = "X38805030"
    'Mark by Amy 2023/07/27 因2023/07/10 改時秀玲休假,和秀玲確認後不需判斷智權人員
'    'Modify by Amy 2023/07/10 11204月修改智權為P1004,也要能看W2001資料
'    strExc(1) = ""
'    strExc(1) = Pub_GetField("Customer", "cu01||cu02='" & strCusNo_S & "'", "cu13")
'    strSalesNo = "W2001," & stTP
'    'end 2023/07/10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm040336 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    txt1(Index).SelStart = 0
    txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    Select Case Index
        '是否含核駁
        Case 7
            If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
                KeyAscii = 0
                Beep
            End If
        '國內外是否分開列印
        Case 8
            If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
                KeyAscii = 0
                Beep
            End If
        '輸出順序
        Case 9
            If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> Asc("4") Then
                KeyAscii = 0
                Beep
            End If
    End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
    If txt1(Index).Text = MsgText(601) Then Exit Sub
    
    Select Case Index
        '收文日/申請日
        Case 3, 4, 13, 14
            If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
                Me.txt1(Index).SetFocus
                txt1_GotFocus Index
                Exit Sub
            End If
    End Select
End Sub

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strField)
        If UCase(strField(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function

Private Sub ReadData()
    Dim intR As Integer
    Dim strCmd As String, strOrder As String, strWhere(2) As String
    
On Error GoTo ErrHnd
    strQ = ""
    
    '收文日
    If Len(txt1(3)) <> 0 Then
        strWhere(1) = strWhere(1) & " And cp05>=" & Val(ChangeTStringToWString(txt1(3))) & " "
    End If
    If Len(txt1(4)) <> 0 Then
        strWhere(1) = strWhere(1) & " And cp05<=" & Val(ChangeTStringToWString(txt1(4))) & " "
    Else
        If Len(txt1(3)) <> 0 Then
            strWhere(1) = strWhere(1) & " And cp05<=" & Val(ChangeTStringToWString(ServerDate - 19110000)) & " "
        End If
    End If
    
    '申請日
    If Me.txt1(13) <> "" And Me.txt1(14) = "" Then
        strWhere(1) = strWhere(1) & " And ((PA10>=" & Val(ChangeTStringToWString(txt1(13))) & " And PA10<=" & strSrvDate(1) & " )  ) "
    End If
    If txt1(13) = "" And txt1(14) <> "" Then
        strWhere(1) = strWhere(1) & " And (PA10<=" & Val(ChangeTStringToWString(txt1(14))) & "  ) "
    End If
    If txt1(13) <> "" And txt1(14) <> "" Then
        strWhere(1) = strWhere(1) & " And ((PA10>=" & Val(ChangeTStringToWString(txt1(13))) & " And PA10<=" & Val(ChangeTStringToWString(txt1(14))) & " )  ) "
    End If
    
    '是否含核駁
    If Me.txt1(7).Text <> "" Then
        If txt1(15) = "3" Then
            strWhere(1) = strWhere(1) + " And ( CP24<>'2' Or CP24 Is Null ) "
        Else
            strWhere(1) = strWhere(1) + " And ( PA16<>'2' Or PA16 Is Null ) "
        End If
    End If

    '申請國家
    If Len(txt1(5)) <> 0 Then
        strWhere(1) = strWhere(1) + " And pa09>='" & txt1(5) & "' "
    End If
    If Len(txt1(6)) <> 0 Then
        strWhere(1) = strWhere(1) + " And pa09<='" & txt1(6) & "' "
    End If
    
    '閉卷
    If ChkClose.Value <> vbChecked Then
        strWhere(1) = strWhere(1) + " And (PA57<>'Y' Or pa57 is null) "
    End If
    
    '銷卷
    If Check6.Value <> vbChecked Then
       strWhere(1) = strWhere(1) + " And pa108 Is Null "
    End If
    
    '專用期仍有效的
    If Check5.Value = 1 Then
        strWhere(1) = strWhere(1) + " And ( PA25>=" & strSrvDate(1) & " Or PA25 Is Null ) "
    End If

    '已閉卷或銷卷或核駁
    If Check12.Value = vbChecked Then
       strWhere(1) = strWhere(1) + " And (PA57||PA108 Is Not Null Or PA16='2') "
    End If

    '已閉卷或銷卷
    If Check11.Value = vbChecked Then
       strWhere(1) = strWhere(1) + " And PA57||PA108 Is Not Null "
    End If
    
    '只印申請案
    If Check7.Value = vbChecked Then
       strWhere(1) = strWhere(1) + " And PA23='1' "
    End If

    '只印已核准
    If Check8.Value = vbChecked Then
       strWhere(1) = strWhere(1) + " And PA16='1' "
    End If

    '只印未核准(核駁或申請中)
    If Check9.Value = vbChecked Then
       strWhere(1) = strWhere(1) + " And Nvl(PA16,'2')='2' "
    End If

    If UCase(txt1(7)) = "N" Then
        'C類核駁的資料不出現
        strWhere(1) = strWhere(1) + " And '1002'<>CP10(+) "
    End If
    
    '勾選「不含」 本所案號
    If Check2.Value = 1 Then
        m_IsOrNot = " Not "
    Else
        m_IsOrNot = ""
    End If
    '若有設定本所案號
    If Trim(Me.txt1(12).Text) <> "" Then
        strWhere(1) = strWhere(1) & " And " & m_IsOrNot & " ( "
        
        arrCaseNo = Split(Replace(txt1(12), vbCrLf, ""), ",")
        For ii = LBound(arrCaseNo) To UBound(arrCaseNo)
            strWhere(1) = strWhere(1) & " ( " & ChgPatent(arrCaseNo(ii)) & " ) "
            If ii < UBound(arrCaseNo) Then
                strWhere(1) = strWhere(1) & " Or "
            End If
        Next ii
        strWhere(1) = strWhere(1) & " ) "
    End If
    
    '智權人員
    If strSalesNo <> MsgText(601) Then
        'Modify by Amy 2023/07/10 也要抓W2001資料
        If InStr(strSalesNo, ",") > 0 Then
            strWhere(1) = strWhere(1) & " And CU13 in ('" & Replace(strSalesNo, ",", "','") & "')"
        Else
            strWhere(1) = strWhere(1) & " And CU13='" & strSalesNo & "' "
        End If
    End If
    
    strWhere(1) = strWhere(1) & " And (CP57 Is Null" & _
      " Or (cp09=(Select MIN(B.cp09) From CaseProgress B Where B.cp01=A.cp01 And B.cp02=A.cp02 And B.cp03=A.cp03 And B.cp04=A.cp04)" & _
      " And Not Exists(Select * From CaseProgress C Where C.cp01=A.cp01 And C.cp02=A.cp02 And C.cp03=A.cp03 And C.cp04=A.cp04 And C.CP57 Is Null And C.cp09<'B')" & _
      "))"
    
    '申請人
    strWhere(0) = " And (pa26>='" & GetNewFagent(strCusNo_S) & "' And pa26<='" & GetNewFagent(strCusNo_E) & "') " & _
                           " And SubStr(pa26,1,8)=cu01(+) And Decode(SubStr(pa26,9,1),Null,'0',SubStr(pa26,9,1))=cu02(+) "
      
    For i = 1 To 5
        Select Case i
            Case 1 '申請人1
                strWhere(2) = strWhere(0)
            Case 2 '申請人2
                strWhere(2) = Replace(strWhere(0), "pa26", "pa27")
            Case 3 '申請人3
                strWhere(2) = Replace(strWhere(0), "pa26", "pa28")
            Case 4 '申請人4
                strWhere(2) = Replace(strWhere(0), "pa26", "pa29")
            Case 5 '申請人5
                strWhere(2) = Replace(strWhere(0), "pa26", "pa30")
        End Select
        If i >= 2 Then strQ = strQ & " Union "
        strQ = strQ & "Select '" & strUserNum & "',pa01,pa02,pa03,pa04,Nvl(pa05,Nvl(pa06,pa07)) as CaseN,pa08,pa09,pa10,pa11,pa48,cu01,cu02,pa57 " & _
                    "From Patent,CaseProgress A,Customer " & _
                    "Where pa01=cp01(+) And pa02=cp02(+) And pa03=cp03(+) And pa04=cp04(+) And 'C'>cp09(+) And pa04='00' " & strWhere(1) & strWhere(2)
    Next i
    
    '刪除暫存檔
    strCmd = "Delete From R040336 Where ID='" & strUserNum & "' "
    cnnConnection.Execute strCmd, intR
    '資料寫入暫存檔
    strCmd = "Insert Into R040336 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012,R017) " & _
                    strQ
    cnnConnection.Execute strCmd, intR
    '有資料
    If intR > 0 Then
        '更新 新案委任日-CP31=Y之 CP05
        strCmd = "Update R040336 Set R013=(Select CP05 From CaseProgress Where R001=cp01 And R002=cp02 And R003=cp03 And R004=cp04 And CP31='Y') " & _
                        "Where ID='" & strUserNum & "' "
        cnnConnection.Execute strCmd, intR
        '更新 新案撰稿完成日-CP31=Y 之 EP08
        'Modify by Amy 2021/10/04 改抓EP28 (預定會稿日)
         strCmd = "Update R040336 Set R014=(Select EP28 From CaseProgress,EngineerProgress Where R001=cp01 And R002=cp02 And R003=cp03 And R004=cp04 And CP31='Y' And cp09=ep02) " & _
                        "Where ID='" & strUserNum & "' "
        cnnConnection.Execute strCmd, intR
        '更新 最近主管機關通知日-C類來函最大收文日(剔除1908-代理人請款/1909-已提申,且收文日為111111)
        strQ = "Select Max(CP05) From CaseProgress Where R001=cp01 And R002=cp02 And R003=cp03 And R004=cp04 " & _
                   "And SubStr(CP09,1,1)='C' And CP10<>'1908' And CP10<>'1909' And CP05<>19221111 "
        strCmd = "Update R040336 Set R015=(" & strQ & ") " & _
                        "Where ID='" & strUserNum & "' "
        cnnConnection.Execute strCmd, intR
        '更新 陳述意見/答辯/申覆完成日-最後一道CP10為107(再審)或205(申復)之EP08
        'Modify by Amy 2021/10/04 改抓A類最後一道CP10為107(再審)或205(申復)之EP28 (預定會稿日)
        strQ = "Select Max(cp09) From CaseProgress Where R001=cp01 And R002=cp02 And R003=cp03 And R004=cp04 And CP10 In('107','205') And SubStr(CP09,1,1)='A' "
        strCmd = "Update R040336 Set R016=(Select EP28 From EngineerProgress Where ep02=(" & strQ & ") ) " & _
                        "Where ID='" & strUserNum & "' "
        cnnConnection.Execute strCmd, intR
        'Add by Amy 2021/10/04 陳述意見/答辯/申覆完成日(預計完成日期)欄有值時,則新案撰稿完成日期(預計完成日期)欄清空
        strCmd = "Update R040336 Set R014=Null Where ID='" & strUserNum & "' And R016 is not null "
        cnnConnection.Execute strCmd, intR
        
        '讀取資料
        strOrder = "Order by R011,R012"
        '國內外分開列印 (不同系統類別分開列印)
        If txt1(8) = "Y" Then
            strOrder = strOrder & ",R001"
        End If
        '排序方式
        Select Case Val(txt1(9))
            '本所案號
            Case 1
                If InStr(strOrder, ",R001") = 0 Then
                    strOrder = strOrder & ",R001,R002,R003,R004"
                Else
                    strOrder = strOrder & ",R002,R003,R004"
                End If
            '案件名稱
            Case 2
                strOrder = strOrder & ",R005"
            '申請國家+本所案號
            Case 3
                strOrder = strOrder & ",R007"
                If InStr(strOrder, ",R001") = 0 Then
                    strOrder = strOrder & ",R001,R002,R003,R004"
                Else
                    strOrder = strOrder & ",R002,R003,R004"
                End If
            '申請國家+案件名稱
            Case 4
                strOrder = strOrder & ",R007,R005"
        End Select
    
        strAllF = "": strWidth = "": strQ = ""
        '顯示代表圖
        If Check1.Value = 1 Then
            strAllF = "圖樣,"
            strWidth = "16,"
            strQ = "'' as Photo,"
        End If
        'Modify by Amy 2021/10/04 +最新進度
        strAllF = strAllF & "本所案號,客戶案件案號,專利名稱,申請日,申請國家" & _
                                    ",申請案號,種類,新案委任<br>日期,新案撰稿<br>完成日期<br>(預計完成日期),最近主管機關<br>通知日期" & _
                                    ",最新進度,陳述意見/答辯/<br>申覆完成日<br>(預計完成日期),建議事項"
        strWidth = strWidth & "13,13,20,9,10" & _
                                        ",13,8,9,14,13" & _
                                        ",15,14,15"
        'Modify by Amy 2021/10/04 +NewCaseProgress(最新進度)/R007(申請國家)
        strQ = "Select " & strQ & "Decode(R017,'Y','*','')||R001||'-'||R002||Decode(R003||R004,'000','','-'||R003||'-'||R004) as CaseName,R010 as CusCaseNo,R005 as PatentName,SqlDateT(R008) as ApplyDate,Na03" & _
                        ",R009 as ApplyCaseNo,Decode(R007,'020',ptm04,ptm03) as PatentKind,SqlDateT(R013) as AppointDate,SqlDateT(R014) as NewCaseEP08,SqlDateT(R015) as CClassCP05" & _
                        ",'' as NewCaseProgress,SqlDateT(R016) as ExpectDate,'' as Suggest,R011||R012 as CusNo,R001 as SysKind,R001,R002,R003,R004,R007 " & _
                  "From R040336,Nation,PatentTrademarkMap,Patent Where ID='" & strUserNum & "' And R007=na01(+) And '1'=ptm01(+) And R006=ptm02(+) " & _
                  "And R001=pa01(+) And R002=pa02(+) And R003=pa03(+) And R004=pa04(+) " & strOrder
        
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, strQ)
        If intQ = 1 Then
            bolOpenXls = False
            If SaveExcel = True Then
                MsgBox "檔案已產生於" & vbCrLf & _
                              " [" & strExcelPath & strFileN & "] "
            End If
        End If
    Else
        MsgBox "無資料產生 ！"
    End If
    Exit Sub

ErrHnd:
    If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Function SaveExcel() As Boolean
    Dim Xls As New Excel.Application
    Dim Wks As New Worksheet
    Dim strWkName As String, strFormat As String
    Dim strOldCusNo As String, StrCusName As String, strOldSysKind As String, strCaseNo(1 To 4) As String
    Dim intAlign As Integer, intPage As Integer
    Dim strTmp(1) As String
    
On Error GoTo ErrHnd1
     
    SaveExcel = False
    intAlign = 0 '0-置中/1-靠左/2-靠右
    intField = 65:  intRow = 1: intTitleR = 1: intPage = 1
    strFileN = strCusNo_S & "案件總簿" & ServerDate & ServerTime & MsgText(43)
    If Dir(strExcelPath & strFileN) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileN
    End If
   
    Xls.SheetsInNewWorkbook = 3
    Xls.Workbooks.add
    '工作表名稱改為中文
    If strWkName = MsgText(601) Then strWkName = Left(Xls.Worksheets(1).Name, Len(Xls.Worksheets(1).Name) - 1)
    Set Wks = Xls.Worksheets(strWkName & intPage)
    Wks.Activate
    'Xls.Visible = True
    bolOpenXls = True
    
    strField = Split(strAllF, ",")
    intWidth = Split(strWidth, ",")
    
    RsQ.MoveFirst
    Do While RsQ.EOF = False
        '申請人編號不同換頁
        If strOldCusNo <> RsQ.Fields("CusNo") Then
            If strOldCusNo = MsgText(601) Then
                StrCusName = GetCustomerName(RsQ.Fields("CusNo"))
                StrCusName = RsQ.Fields("CusNo") & " " & StrCusName
            Else
                Call SetXlsEnd(Xls, Wks)
                If strCusNo_S <> strCusNo_E Then
                    Wks.Name = strOldCusNo & strOldSysKind
                ElseIf txt1(8) = "Y" Then
                    Wks.Name = strOldSysKind
                End If
                intPage = intPage + 1
                If intPage > 3 Then
                    Xls.Worksheets.add
                End If
                Set Wks = Xls.Worksheets(strWkName & intPage)
                Wks.Activate
                
                StrCusName = GetCustomerName(strOldCusNo)
                StrCusName = strOldCusNo & " " & StrCusName
            End If
            intRow = 1: strOldSysKind = "": intTitleR = 1
            Call SetTitle(Wks, StrCusName)
        '系統別不同換頁
        ElseIf txt1(8) = "Y" And strOldSysKind <> RsQ.Fields("SysKind") And strOldSysKind <> MsgText(601) Then
            Call SetXlsEnd(Xls, Wks)
            If strCusNo_S <> strCusNo_E Then
                Wks.Name = strOldCusNo & strOldSysKind
            ElseIf txt1(8) = "Y" Then
                Wks.Name = strOldSysKind
            End If
            intPage = intPage + 1
            If intPage > 3 Then
                Xls.Worksheets.add
            End If
            Set Wks = Xls.Worksheets(strWkName & intPage)
            Wks.Activate
            
            StrCusName = GetCustomerName(RsQ.Fields("CusNo"))
            StrCusName = RsQ.Fields("CusNo") & " " & StrCusName
            intRow = 1: intTitleR = 1
            Call SetTitle(Wks, StrCusName)
        End If
        
        strCaseNo(1) = "" & RsQ.Fields("R001")
        strCaseNo(2) = "" & RsQ.Fields("R002")
        strCaseNo(3) = "" & RsQ.Fields("R003")
        strCaseNo(4) = "" & RsQ.Fields("R004")
        
        For j = LBound(strField) To UBound(strField)
            strFormat = "": intAlign = 0: strTmp(1) = ""
            strTmp(0) = Replace(strField(j), "<br>", "")
            
            Select Case strTmp(0)
                Case "圖樣"
                    strTmp(1) = ""
                    'Modify by Amy 2023/07/27 +strFileN 函數改寫至basUpdate
                    Call PutXlsImg(Me.Name, Wks, Chr(intField + j), intRow, strCaseNo(1), strCaseNo(2), strCaseNo(3), strCaseNo(4))
                Case "本所案號"
                    intAlign = 1
                    strTmp(1) = "" & RsQ.Fields("CaseName")
                Case "客戶案件案號"
                    strFormat = "@"
                    strTmp(1) = "" & RsQ.Fields("CusCaseNo")
                Case "專利名稱"
                    intAlign = 1
                    strTmp(1) = "" & RsQ.Fields("PatentName")
                Case "申請日"
                    strTmp(1) = "" & RsQ.Fields("ApplyDate")
                Case "申請國家"
                    strTmp(1) = "" & RsQ.Fields("Na03")
                Case "申請案號"
                    strFormat = "@"
                    strTmp(1) = "" & RsQ.Fields("ApplyCaseNo")
                Case "種類"
                    strTmp(1) = "" & RsQ.Fields("PatentKind")
                Case "新案委任日期"
                    strTmp(1) = "" & RsQ.Fields("AppointDate")
                Case "新案撰稿完成日期(預計完成日期)"
                    strTmp(1) = "" & RsQ.Fields("NewCaseEP08")
                Case "最近主管機關通知日期"
                    strTmp(1) = "" & RsQ.Fields("CClassCP05")
                'Add by Amy 2021/10/04
                Case "最新進度"
                    intAlign = 1
                    strTmp(1) = GetCaseProgress_NewA("" & RsQ.Fields("R007"), RsQ.Fields("R001"), RsQ.Fields("R002"), RsQ.Fields("R003"), RsQ.Fields("R004"))
                Case "陳述意見/答辯/申覆完成日(預計完成日期)"
                    strTmp(1) = "" & RsQ.Fields("ExpectDate")
                Case "建議事項"
                    intAlign = 1
                    strTmp(1) = GetAllEmpElectronMemo(RsQ.Fields("R001"), RsQ.Fields("R002"), RsQ.Fields("R003"), RsQ.Fields("R004"), "" & RsQ.Fields("R007"))
            End Select
            '設定儲存格格式
            If strFormat <> MsgText(601) Then
                Wks.Range(Chr(intField + j) & intRow).NumberFormatLocal = strFormat
            End If
            Wks.Range(Chr(intField + j) & intRow).Value = strTmp(1)
            Select Case intAlign
                Case 0 '置中
                    Wks.Range(Chr(intField + j) & intRow).HorizontalAlignment = xlCenter
                Case 1 '靠左
                    Wks.Range(Chr(intField + j) & intRow).HorizontalAlignment = xlLeft
                Case 2 '靠右
                    Wks.Range(Chr(intField + j) & intRow).HorizontalAlignment = xlRight
            End Select
            'Modify by Amy 2021/10/04 +最新進度/建議事項
            If j = GetValue("專利名稱") Or j = GetValue("最新進度") Or j = GetValue("建議事項") Then
                Wks.Range(Chr(intField + j) & intRow).WrapText = True
            End If
        Next j
        
        strOldCusNo = "" & RsQ.Fields("CusNo")
        strOldSysKind = "" & RsQ.Fields("SysKind")
        intRow = intRow + 1
        RsQ.MoveNext
    Loop
    Wks.Range(Chr(intField) & intTitleR + 1 & ":" & Chr(UBound(strField) + intField) & intRow).Font.Size = 11
  
    Call SetXlsEnd(Xls, Wks)
    If strCusNo_S <> strCusNo_E Then
        Wks.Name = strOldCusNo & strOldSysKind
    ElseIf txt1(8) = "Y" Then
        Wks.Name = strOldSysKind
    End If
    
    '判斷若版本2007以上改變存格式
    If Val(Xls.Version) < 12 Then
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
    Else
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
    End If
    Xls.Workbooks.Close
    Xls.Quit
    SaveExcel = True
    Exit Function
    
ErrHnd1:
    If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
    If bolOpenXls = True Then
        '判斷若版本2007以上改變存格式
        If Val(Xls.Version) < 12 Then
            Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
        Else
            Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
        End If
        Xls.Workbooks.Close
        Xls.Quit
    End If

End Function

Private Sub SetTitle(ByRef Wks As Worksheet, ByVal stCusomer As String)
    With Wks
        '***表頭設定***
        .Range(Chr(intField) & intRow).Value = "客戶案件總簿 (專利)"
        .Range(Chr(intField) & intRow).Font.Size = 18
        .Range(Chr(intField) & intRow).Font.Bold = True
        .Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).MergeCells = True
        .Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).HorizontalAlignment = xlCenter
        .Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).VerticalAlignment = xlCenter
        intRow = intRow + 1
        
        .Range(Chr(intField) & intRow).Value = "收件人：" & stCusomer
        .Range(Chr(intField) & intRow).Font.Size = 12
        .Range(Chr(intField) & intRow).HorizontalAlignment = xlLeft
        .Range(Chr(intField + UBound(strField) - 1) & intRow).Font.Size = 12
        .Range(Chr(intField + UBound(strField) - 1) & intRow).Value = "列印日期:" & CFDate(ACDate(ServerDate))
        intRow = intRow + 1
        
        For i = LBound(strField) To UBound(strField)
            .Columns(Chr(intField + i) & ":" & Chr(intField + i)).ColumnWidth = intWidth(i)
            .Range(Chr(intField + i) & intRow).Value = Replace(strField(i), "<br>", vbCrLf)
            .Range(Chr(intField + i) & intRow).HorizontalAlignment = xlCenter
        Next i
        intTitleR = intRow
        intRow = intRow + 1
    End With
End Sub

Private Sub SetXlsEnd(ByRef Xls As Excel.Application, ByRef Wks As Worksheet)
    With Wks.Range(Chr(intField) & intTitleR & ":" & Chr(intField + UBound(strField)) & intRow - 1)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin '細線
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    '設定
    Wks.PageSetup.PaperSize = 9 'A4
    Wks.PageSetup.PrintTitleRows = "$1:$" & intTitleR
    Wks.PageSetup.Orientation = xlLandscape '橫印
    If Check1.Value = 1 Then
        Wks.PageSetup.Zoom = 75
    Else
        Wks.PageSetup.Zoom = 80
    End If
    Wks.PageSetup.LeftMargin = 0 '邊界
    Wks.PageSetup.RightMargin = 0
    Wks.PageSetup.TopMargin = Xls.InchesToPoints(0.4)
    Wks.PageSetup.BottomMargin = Xls.InchesToPoints(0.4)
    Wks.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
End Sub

Private Function GetAllEmpElectronMemo(ByVal stCP01 As String, ByVal stCP02 As String, ByVal stCP03 As String, ByVal stCP04 As String, ByVal stNation As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim intQ As Integer, strQ As String, stTP As String
    
    GetAllEmpElectronMemo = ""
    strQ = "Select Nvl(Nvl(Decode('" & stNation & "','000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10) as CaseProperty,EED05 " & _
              "From CaseProgress,EmpElectronData,CasePropertyMap " & _
              "Where cp09=EED01(+) And EED01 is not null And cp01=cpm01(+) And cp10=cpm02(+) And EED05 is not null " & _
              "And cp01='" & stCP01 & "' And cp02='" & stCP02 & "' And cp03='" & stCP03 & "' And cp04='" & stCP04 & "' " & _
              "Order by cp05,cp66,cp67 "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            stTP = stTP & "@@@" & RsQ.Fields("CaseProperty") & "：" & RsQ.Fields("EED05")
            RsQ.MoveNext
        Loop
        If stTP <> MsgText(601) Then
            stTP = Mid(stTP, 4)
            GetAllEmpElectronMemo = Replace(stTP, "@@@", vbCrLf)
        End If
    End If
    Set RsQ = Nothing
End Function
