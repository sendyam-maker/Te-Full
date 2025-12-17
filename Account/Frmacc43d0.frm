VERSION 5.00
Begin VB.Form Frmacc43d0 
   AutoRedraw      =   -1  'True
   Caption         =   "取消過帳或月(年)結"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6900
   Begin VB.TextBox TxtB 
      Height          =   285
      Index           =   6
      Left            =   3360
      MaxLength       =   3
      TabIndex        =   11
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox TxtB 
      Height          =   285
      Index           =   5
      Left            =   3360
      MaxLength       =   3
      TabIndex        =   9
      Top             =   2700
      Width           =   495
   End
   Begin VB.TextBox TxtB 
      Height          =   285
      Index           =   4
      Left            =   4320
      MaxLength       =   6
      TabIndex        =   7
      Top             =   2325
      Width           =   800
   End
   Begin VB.TextBox TxtB 
      Height          =   285
      Index           =   3
      Left            =   3360
      MaxLength       =   6
      TabIndex        =   6
      Top             =   2325
      Width           =   800
   End
   Begin VB.TextBox TxtB 
      Height          =   285
      Index           =   2
      Left            =   3360
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1935
      Width           =   800
   End
   Begin VB.TextBox TxtB 
      Height          =   285
      Index           =   1
      Left            =   3360
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1560
      Width           =   800
   End
   Begin VB.OptionButton Option1 
      Caption         =   "取消12月過帳及年結"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "取消年結"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   8
      Top             =   2730
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "取消過帳與月結"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Top             =   2340
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      Caption         =   "取消月結"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   1950
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      Caption         =   "取消過帳"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdProc 
      Caption         =   "取消處理"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox TxtB 
      Height          =   285
      Index           =   0
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   0
      Top             =   645
      Width           =   510
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   4440
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "已過帳日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "已月結日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   20
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "已年結日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4725
      TabIndex        =   19
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "L3(1)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1380
      TabIndex        =   18
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "L3(2)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3675
      TabIndex        =   17
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "L3(3)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   16
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "L3(0)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   15
      Top             =   690
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   690
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   912
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc43d0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 (無需修改)
'Create by Lydia 2016/01/22 取消過帳或月(年)結
Option Explicit
Dim oOpt As OptionButton
Dim oText As TextBox

Private Sub cmdExit_Click()
    If TxtB(0).Locked = False Then
       Unload Me
    End If
End Sub

Private Sub CmdProc_Click()
Dim idx As Integer
Dim strTmp As String

On Error GoTo ErrHandle1

   If Option1(0).Value = True Then
      idx = 1
      'add by sonia 2021/10/6
      If TxtB(1) <= TxtB(2) Then
         MsgBox "取消過帳年月不可＜＝月結年月！", vbCritical
         Exit Sub
      End If
      'end 2021/10/6
   ElseIf Option1(1).Value = True Then
      idx = 2
      'add by sonia 2021/10/6
      If TxtB(2) < TxtB(1) Then
         MsgBox "取消月結年月不可＜過帳年月！", vbCritical
         Exit Sub
      End If
      'end 2021/10/6
   ElseIf Option1(2).Value = True Then
      idx = 3
      'add by sonia 2021/10/6
      If TxtB(3) < TxtB(1) Then
         MsgBox "取消月結年月不可＜過帳年月！", vbCritical
         Exit Sub
      End If
      'end 2021/10/6
   ElseIf Option1(3).Value = True Then
      idx = 4
   ElseIf Option1(4).Value = True Then
      idx = 5
   Else
      idx = 0
      MsgBox "請選擇取消處理的項目!", vbCritical
      Exit Sub
   End If
   
   If CheckTxtB(0, True) = False Then
      TxtB(0).SetFocus
      TxtB_GotFocus 0
      Exit Sub
   End If
   
   If Option1(2).Value = True Then
      If TxtB(3) > TxtB(4) Then
         MsgBox "起始年月不可大於終止年月!", vbCritical
         Exit Sub
      ElseIf Val(Right(TxtB(3), 3)) <> Val(Right(TxtB(4), 3)) Then
         MsgBox "不可跨年度!", vbCritical
         Exit Sub
      End If
   End If
   
   strTmp = "select a0b10 from acc0b0 where a0b04=" & CNULL(TxtB(0))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strTmp)
   If intI = 1 Then
      If IsNull(RsTemp(0)) Then
          adoTaie.BeginTrans
             strTmp = "update acc0b0 set a0b10='10' where a0b04=" & CNULL(TxtB(0))
             adoTaie.Execute strTmp, intI
          adoTaie.CommitTrans
      Else
          MsgBox "目前已有其他人在作業,不可執行取消處理!", vbCritical
          Exit Sub
      End If
   End If
   
   Screen.MousePointer = vbHourglass 'Added by Lydia 2016/05/13
   CmdProc.Enabled = False
   cmdExit.Enabled = False
   If idx > 0 Then
      Call Process(idx)
   End If
   
   '釋放作業
   adoTaie.BeginTrans
      strTmp = "update acc0b0 set a0b10=null where a0b04=" & CNULL(TxtB(0))
      adoTaie.Execute strTmp, intI
   adoTaie.CommitTrans
           
   CmdProc.Enabled = True
   cmdExit.Enabled = True
   Screen.MousePointer = vbDefault 'Added by Lydia 2016/05/13
   
   If GetA0b01(TxtB(0)) = True Then
      Call CheckValidate
   End If
   
   Exit Sub
   
ErrHandle1:
    If Err.Number <> 0 Then
       adoTaie.RollbackTrans
       Screen.MousePointer = vbDefault 'Added by Lydia 2016/05/13
       MsgBox Err.Description, , MsgText(5)
       CmdProc.Enabled = True
       cmdExit.Enabled = True
    End If
End Sub

Private Sub Process(ByVal inX As Integer)
Dim strDate1 As String
Dim strDate2 As String
Dim strCon1 As String '刪除月結傳票
Dim strCon2 As String '更新月結日期
Dim strCon3 As String '重整傳票號碼
Dim strCon4 As String '更新自動編號檔
Dim stRtn As String
Dim strKind As String
Dim rsAD As New ADODB.Recordset
Dim strTmp As String
Dim intL As Long 'intI 可能不敷使用
On Error GoTo ErrHandle

    Select Case inX
        Case 1 '取消過帳
            strDate1 = Replace(TxtB(1), "/", ""): strDate2 = Replace(TxtB(1), "/", "")
            strCon1 = "": strCon3 = "": strCon4 = ""
            '更新acc0b0日期:a0b01=取消年月之上月底日期
            strExc(1) = ACDate(GetLastDay(CompDate(2, -1, DBDATE(strDate1 & "01"))))
            strCon2 = "a0b01=" & CNULL(strExc(1), True)
        Case 2 '取消月結
            strDate1 = Replace(TxtB(2), "/", ""): strDate2 = Replace(TxtB(2), "/", "")
            '刪除月結傳票 (取消年月起月之次月)
            strExc(3) = IIf(Right(strDate1, 2) = "12", Format(Val(Left(strDate1, 3)) + 1, "000") & "01", Left(strDate1, 3) & Format(Val(Right(strDate1, 2)) + 1, "00"))
            strCon1 = " and a0205>=" & CNULL(strExc(3) & "01", True) & " and a0205<=" & CNULL(strExc(3) & "31", True)
            
            '更新acc0b0日期:a0b02=取消年月之上月底日期
            strExc(1) = ACDate(GetLastDay(CompDate(2, -1, DBDATE(strDate1 & "01"))))
            strCon2 = "a0b02=" & CNULL(strExc(1), True)
            '重整該年度起月之次月的傳票號碼
            strCon3 = "A0201=" & CNULL(TxtB(0)) & " and a0205>=" & CNULL(strExc(3) & "01", True) & " and a0205<=" & CNULL(strExc(3) & "31", True)
            '更新自動編號檔  a1r02=西元年度 and a1r03=起月之次月;
            strCon4 = "and a1r02=" & Val(DBYEAR(strExc(3) & "01")) & " and a1r03=" & Val(DBMONTH(strExc(3) & "01"))
        Case 3 '取消過帳及月結
            strDate1 = Replace(TxtB(3), "/", ""): strDate2 = Replace(TxtB(4), "/", "")
            '刪除月結傳票 (a0205>=取消年月起月之次月1日 and a0205<=取消年月迄月之次月31日)
            strExc(3) = IIf(Right(strDate1, 2) = "12", Format(Val(Left(strDate1, 3)) + 1, "000") & "01", Left(strDate1, 3) & Format(Val(Right(strDate1, 2)) + 1, "00"))
            strExc(4) = IIf(Right(strDate2, 2) = "12", Format(Val(Left(strDate2, 3)) + 1, "000") & "01", Left(strDate2, 3) & Format(Val(Right(strDate2, 2)) + 1, "00"))
            strCon1 = " and a0205>=" & CNULL(strExc(3) & "01", True) & " and a0205<=" & CNULL(strExc(4) & "31", True)
            
            '更新acc0b0日期:a0b01=取消年月起月之上月底日期,a0b02=取消年月起月之上月底日期
            strExc(1) = ACDate(GetLastDay(CompDate(2, -1, DBDATE(strDate1 & "01"))))
            strCon2 = "a0b01=" & CNULL(strExc(1), True) & ", a0b02=" & CNULL(strExc(1), True)
            '重整該年度起月之次月至迄月之次月的傳票號碼
            strCon3 = "A0201=" & CNULL(TxtB(0)) & " and a0205>=" & CNULL(strExc(3) & "01", True) & " and a0205<=" & CNULL(strExc(4) & "31", True)
            '更新自動編號檔  a1r02=西元年度 and a1r03>=起月之次月 and a1r03<=迄月之次月
            strCon4 = "and a1r02=" & Val(DBYEAR(strExc(3) & "01")) & " and a1r03>=" & Val(DBMONTH(strExc(3) & "01")) & " and a1r03<=" & Val(DBMONTH(strExc(3) & "01"))
        Case 4 '取消年結
            strDate1 = TxtB(5) & "12": strDate2 = TxtB(5) & "12"
            '刪除月結傳票 (and a0205>=取消年12月1日 and a0205<=取消年12月迄31日)
            'Modifed by Lydia 2016/05/09 排除11月月結傳票
            'strCon1 = " and a0205>=" & CNULL(strDate1 & "01", True) & " and a0205<=" & CNULL(strDate1 & "31", True)
            strCon1 = " and a0205>=" & CNULL(strDate1 & "01", True) & " and a0205<=" & CNULL(strDate1 & "31", True) & " and instr(ax212,'前月損益') = 0"
            
            '更新acc0b0日期:a0b03=取消年前一年之12/31
            'Modified by Lydia 2016/05/13 單純取消年結時,還要再更新a0b02=取消年之11/30
            strExc(1) = Format(Val(TxtB(5)) - 1, "000") & "1231"
            strExc(2) = TxtB(5) & "1130"
            strCon2 = "a0b03=" & CNULL(strExc(1), True) & ", a0b02=" & CNULL(strExc(2), True)
            ' 重整該年度12月的傳票號碼
            strCon3 = "A0201=" & CNULL(TxtB(0)) & " and a0205>=" & CNULL(strDate1 & "01", True) & " and a0205<=" & CNULL(strDate1 & "31", True)
             '更新自動編號檔 a1r02=西元年度 and a1r03=12
            strCon4 = "and a1r02=" & Val(DBYEAR(strDate1 & "01")) & " and a1r03=12"
           
        Case 5 '取消12月過帳及年結
            strDate1 = TxtB(6) & "12": strDate2 = TxtB(6) & "12"
            '刪除月結傳票 (and a0205>=取消年12月1日 and a0205<=取消年12月迄31日)
            'Modifed by Lydia 2016/05/09 排除11月月結傳票
            'strCon1 = " and a0205>=" & CNULL(strDate1 & "01", True) & " and a0205<=" & CNULL(strDate1 & "31", True)
            strCon1 = " and a0205>=" & CNULL(strDate1 & "01", True) & " and a0205<=" & CNULL(strDate1 & "31", True) & " and instr(ax212,'前月損益') = 0"
            
            '更新acc0b0日期:a0b01=取消年之11/30,a0b02=取消年之11/30, a0b03=取消年前一年之12/31
            strExc(1) = Format(Val(TxtB(6)) - 1, "000") & "1231"
            strExc(2) = TxtB(6) & "1130"
            strCon2 = "a0b01=" & CNULL(strExc(2), True) & ", a0b02=" & CNULL(strExc(2), True) & ", a0b03=" & CNULL(strExc(1), True)
            '重整該年度12月的傳票號碼
            strCon3 = "A0201=" & CNULL(TxtB(0)) & " and a0205>=" & CNULL(strDate1 & "01", True) & " and a0205<=" & CNULL(strDate1 & "31", True)
             '更新自動編號檔 a1r02=西元年度 and a1r03=12
            strCon4 = "and a1r02=" & Val(DBYEAR(strDate1 & "01")) & " and a1r03=12"
    End Select
    strKind = ""
    If TxtB(0) = "1" Then
       strKind = MsgText(801) 'D
    ElseIf TxtB(0) = "J" Then
        strKind = MsgText(819) 'JD
    'Add by Amy 2020/04/16 +L
    ElseIf TxtB(0) = "L" Then
         strKind = MsgText(820) 'LD
    End If
    
    'Added by Lydia 2017/03/06 檢查是否有月結傳票
    If inX > 1 Then
       strTmp = "select distinct a0201 a01,a0202 a02 from acc020,acc021 where a0201='" & TxtB(0) & "' " & strCon1 & " and a0201=ax201(+) and a0202=ax202(+) and ax205='3222' "
       intI = 1
       Set rsAD = ClsLawReadRstMsg(intI, strTmp)
       If intI = 1 Then
          If inX < 4 Then
             'Modified by Lydia 2018/01/08 +,再做取消處理
             MsgBox "請先將月結傳票移做他用,再做取消處理!!"
          Else
             'Modified by Lydia 2018/01/08 +,再做取消處理
             MsgBox "請先將12月的月結傳票移做他用,再做取消處理!!"
          End If
          Exit Sub
       End If
    End If
    'end 2017/03/06
    
    adoTaie.BeginTrans
        '取消過帳(日記帳)
        'Added by Lydia 2016/05/13 單純取消月結或單純取消年結時,不必做取消過帳的資料更新
        If inX <> 2 And inX <> 4 Then
            strTmp = "update acc040 set A0406=0,A0407=0,A0408=0 where a0401=" & Val(Left(strDate1, 3)) & _
                    " and a0402>=" & Val(Right(strDate1, 2)) & " and a0402<=" & Val(Right(strDate2, 2)) & " and a0403=" & CNULL(TxtB(0))
                adoTaie.Execute strTmp, intL
                
            strTmp = "update acc021 set ax210=null where (ax201,ax202) in (select a0201,a0202 from acc020 where a0201='" & TxtB(0) & "' and a0205>='" & strDate1 & "01" & "' and a0205<='" & strDate2 & "31" & "') "
                adoTaie.Execute strTmp, intL
        End If
        'end 2016/05/13
        '刪除月結傳票
        'Remove by Lydia 2017/03/06
'        If strCon1 <> "" Then
'            strTmp = "select  ''''||a02||'''' from (select distinct a0201 a01,a0202 a02 from acc020,acc021 where a0201='" & TxtB(0) & "' " & strCon1 & " and a0201=ax201(+) and a0202=ax202(+) and ax205='3222') "
'            intI = 1
'            Set rsAD = ClsLawReadRstMsg(intI, strTmp)
'            If intI = 1 Then
'               stRtn = rsAD.GetString(adClipString, , , ",")
'               stRtn = Left(stRtn, Len(stRtn) - 1)
'               strTmp = " delete from acc020 where a0202 in (" & stRtn & ") and a0201=" & CNULL(TxtB(0))
'               adoTaie.Execute strTmp, intL
'               strTmp = " delete from acc021 where ax202 in (" & stRtn & ") and ax201=" & CNULL(TxtB(0))
'               adoTaie.Execute strTmp, intL
'            End If
'        End If
        'end 2017/03/06
        
        '更新acc0b0日期
        If strCon2 <> "" Then
            strTmp = "update acc0b0 set " & strCon2 & " where a0b04='" & TxtB(0) & "' "
            adoTaie.Execute strTmp, intL
        End If
        'Remove by Lydia 2017/03/06 因為無法確認月結傳票,所以不做重整
        '重整傳票號碼
'        If strCon3 <> "" Then
'           strTmp = " DECLARE" & _
'                    " V_DATA ACC020%ROWTYPE;V_NUM NUMBER:=1000000000;" & _
'                    " V_NEW_A0202 ACC020.A0202%TYPE; V_LST_A0201 ACC020.A0201%TYPE:='X';" & _
'                    " V_LST_A0205 ACC020.A0205%TYPE:=0;V_NOW_A0205 ACC020.A0205%TYPE:=0;" & _
'                    " V_A1R01 ACC1R0.A1R01%TYPE:='X';V_A1R02 ACC1R0.A1R02%TYPE:=0;V_A1R03 ACC1R0.A1R03%TYPE:=0;" & _
'                    " CURSOR C_ACC020 IS" & _
'                    " SELECT * FROM ACC020 WHERE " & strCon3 & _
'                    " ORDER BY A0201,A0205, A0202;" & _
'                    " BEGIN" & _
'                    "   OPEN C_ACC020;" & _
'                    " Loop" & _
'                    " FETCH C_ACC020 INTO V_DATA;" & _
'                    "      EXIT WHEN C_ACC020%NOTFOUND;" & _
'                    "      V_NOW_A0205:=TRUNC(V_DATA.A0205/100);" & _
'                    "      If (V_LST_A0201 <> V_DATA.A0201 Or V_LST_A0205 <> V_NOW_A0205) Then" & _
'                    "         UPDATE ACC1R0 SET A1R04=MOD(V_NUM,10000) WHERE A1R01=V_A1R01 AND A1R02=V_A1R02 AND A1R03=V_A1R03;" & _
'                    "         V_NUM:=1000000000+10000*V_NOW_A0205;V_LST_A0201:=V_DATA.A0201;V_LST_A0205:=V_NOW_A0205;" & _
'                    "         V_A1R01:=V_DATA.A0201;V_A1R02:=TRUNC(V_NOW_A0205/100)+1911;V_A1R03:=MOD(V_NOW_A0205,100);" & _
'                    "      END IF;" & _
'                    "      V_NUM:=V_NUM+1;V_NEW_A0202:='X'||SUBSTR(TO_CHAR(V_NUM),2);" & _
'                    "      UPDATE ACC020 SET A0202=V_NEW_A0202 WHERE A0201=V_DATA.A0201 AND A0202=V_DATA.A0202;UPDATE ACC021 SET AX202=V_NEW_A0202 WHERE AX201=V_DATA.A0201 AND AX202=V_DATA.A0202;" & _
'                    " END LOOP;" & _
'                    "   UPDATE ACC1R0 SET A1R04=MOD(V_NUM,10000) WHERE A1R01=V_A1R01 AND A1R02=V_A1R02 AND A1R03=V_A1R03;" & _
'                    "   UPDATE ACC020 SET A0202='D'||SUBSTR(A0202,2) WHERE A0202 LIKE 'X%';UPDATE ACC021 SET AX202='D'||SUBSTR(AX202,2) WHERE AX202 LIKE 'X%';" & _
'                    " END;"
'            adoTaie.Execute strTmp
'        End If
'        '更新自動編號檔
'        If strCon4 <> "" Then
'            strTmp = "update acc1r0 set a1r04=(select count(*) from acc020 where a0201='" & TxtB(0) & "' and a0205>=(a1r02-1911)||lpad(a1r03,2,'0')||'01' and a0205<=(a1r02-1911)||lpad(a1r03,2,'0')||'31') where a1r01=" & CNULL(strKind) & strCon4
'            adoTaie.Execute strTmp, intL
'        End If
        'end Remove by Lydia 2017/03/06 因為無法確認月結傳票,所以不做重整
    adoTaie.CommitTrans
    
    MsgBox "作業完成!", vbInformation + vbOKOnly
    
    Exit Sub
ErrHandle:
    If Err.Number <> 0 Then
       adoTaie.RollbackTrans
       MsgBox Err.Description, , MsgText(5)
    End If
        
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 7020
   Me.Height = 4200
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath3)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next

   TxtB(0) = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc43d0 = Nothing
End Sub

Private Sub TxtB_Change(Index As Integer)
   If Index = 0 Then
      If TxtB(Index).Locked = False Then
        If CheckTxtB(Index) Then
           Call CheckValidate
        Else
           TxtB(Index).SetFocus
           TxtB_GotFocus Index
        End If
      End If
   End If
End Sub
Private Sub CheckValidate()
   If Label3(0) <> "" Then
        For Each oOpt In Option1
           oOpt.Enabled = False
        Next
        For Each oText In TxtB
            If oText.Index > 0 Then
               oText.Locked = True
            End If
        Next
        '若a0b03=a0b02=a0b01則只可點選'取消年結'或'取消12月過帳及年結'(預設在取消12月過帳及年結)
        If Label3(3) = Label3(2) And Label3(2) = Label3(1) Then
           Option1(3).Enabled = True: Option1(4).Enabled = True
           TxtB(5).Locked = False: TxtB(6).Locked = False
           Option1(4).Value = 1
        '若a0b02=a0b01則只可點選'取消月結'或'取消月結及過帳'(預設在取消月結及過帳)
        ElseIf Label3(2) = Label3(1) Then
           Option1(1).Enabled = True: Option1(2).Enabled = True
           TxtB(2).Locked = False: TxtB(3).Locked = False: TxtB(4).Locked = False
           Option1(2).Value = 1
        Else
           '其他-取消過帳
           Option1(0).Enabled = True
           TxtB(1).Locked = False
           Option1(0).Value = 1
        End If
   End If
End Sub

Private Sub TxtB_GotFocus(Index As Integer)
    TextInverse TxtB(Index)
End Sub

Private Sub TxtB_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 0 Then
      KeyAscii = UpperCase(KeyAscii)
  Else
     If KeyAscii <> Asc("/") Then
        KeyAscii = Pub_NumAscii(KeyAscii)
     End If
  End If
End Sub

Private Sub TxtB_LostFocus(Index As Integer)
  If Index > 0 Then
    If CheckTxtB(Index) = False Then
       TxtB(Index).SetFocus
       TxtB_GotFocus Index
    End If
  End If
End Sub
Private Function CheckTxtB(ByVal idR As Integer, Optional ByVal BolA As Boolean = False) As Boolean

   CheckTxtB = False
   Select Case idR
       Case 0
           If TxtB(idR) = "" Then
              MsgBox "公司別不可空白!", vbCritical
              Exit Function
           'Modify by Amy 2020/04/16 改抓作帳公司
           'ElseIf TxtB(idR) <> "1" And TxtB(idR) <> "J" Then
           ElseIf InStr(GetBookKeepCmp, TxtB(idR)) = 0 Then
              MsgBox Label2(0) & MsgText(63), vbCritical
           'end 2020/04/16
              Exit Function
           ElseIf BolA = False Then
              If GetA0b01(TxtB(idR).Text) = False Then
                 Exit Function
              End If
           End If
       Case 1, 2, 3, 4
           If Len(TxtB(idR)) = 5 And InStr(TxtB(idR), "/") = 0 Then
              TxtB(idR) = Left(TxtB(idR), 3) & "/" & Right(TxtB(idR), 2)
           End If
           If Len(TxtB(idR)) <> 6 Then
              MsgBox "請輸入年月!", vbCritical
              Exit Function
           ElseIf Val(Replace(TxtB(idR).Text, "/", "")) > Val(Replace(TxtB(idR).Tag, "/", "")) Then
              MsgBox "取消年月超出範圍!", vbCritical
              Exit Function
           End If
       Case 5, 6
           If Len(TxtB(idR)) < 3 Then
              MsgBox "請輸入年度!", vbCritical
              Exit Function
           ElseIf Val(TxtB(idR).Text) > Val(TxtB(idR).Tag) Then
              MsgBox "取消年度超出範圍!", vbCritical
              Exit Function
           End If
   End Select
   
   CheckTxtB = True
   
End Function
Private Function GetA0b01(ByVal strB As String) As Boolean

   GetA0b01 = False
   Label3(0) = "": Label3(1) = "": Label3(2) = ""

   If strB <> "" Then
      'Modify by Amy 2020/04/16 1公司顯示2公司名稱
      strExc(0) = "Select Decode(a0801,'1',AName,a0802) a0802,Acc0b0.* From acc080,acc0b0 " & _
                        ",(Select '1' ANo,a0802 AName,a0820 BName From Acc080 Where a0801='2' )" & _
                        "Where a0801 = a0b04 And a0801=ANo(+) And a0801 = " & CNULL(strB)
      intI = 0
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         TxtB(0) = strB
         Label3(0) = "" & RsTemp.Fields("A0802")
         Label3(1) = CFDate(RsTemp.Fields("A0B01"))
         Label3(2) = CFDate(RsTemp.Fields("A0B02"))
         Label3(3) = CFDate(RsTemp.Fields("A0B03"))
         TxtB(1) = Mid(Label3(1), 1, 6)
         TxtB(2) = Mid(Label3(2), 1, 6)
         TxtB(3) = Mid(Label3(2), 1, 6)
         TxtB(4) = Mid(Label3(2), 1, 6)
         TxtB(5) = Mid(Label3(3), 1, 3)
         TxtB(6) = Mid(Label3(3), 1, 3)
         For Each oText In TxtB
             oText.Tag = oText.Text
         Next
      Else
         Exit Function
      End If
   End If
   GetA0b01 = True
End Function
