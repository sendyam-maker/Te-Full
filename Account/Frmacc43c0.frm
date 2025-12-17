VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc43c0 
   AutoRedraw      =   -1  'True
   Caption         =   "每月業績開放/關閉輸入"
   ClientHeight    =   5940
   ClientLeft      =   48
   ClientTop       =   576
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   8460
   Begin VB.CommandButton Cmd_AddSalesPoint 
      BackColor       =   &H00C0FFC0&
      Caption         =   "      關閉後又加           安全基金撥補     請按此鈕"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   3936
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   390
      Width           =   1880
   End
   Begin VB.CommandButton cmdMail 
      Caption         =   "4.通知主管寫報告"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6000
      TabIndex        =   8
      Top             =   810
      Width           =   2040
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2310
      TabIndex        =   6
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "2.轉撥檢查"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "3.關閉"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   6000
      TabIndex        =   4
      Top             =   465
      Width           =   2040
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "1.開放"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   6000
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.CheckBox Check1 
      Caption         =   "發E-mail通知"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1230
      TabIndex        =   1
      Top             =   30
      Width           =   1000
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   2508
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   8004
      _ExtentX        =   14118
      _ExtentY        =   4424
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   8
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "綠色:表已確認後又加人員且未確認"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   192
      Left            =   240
      TabIndex        =   10
      Top             =   3156
      Width           =   3492
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   204
      Left            =   12
      TabIndex        =   9
      Top             =   1140
      Width           =   648
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "資料年月："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   1125
   End
End
Attribute VB_Name = "Frmacc43c0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/11/01 Form2.0已修改 grdDataList
'Create by Amy 2016/01/11
Option Explicit

Dim strColN(), intWidth()
Dim i As Integer
Dim strA0b01 As String, strA0b05 As String 'Add by Amy 2020/06/18
Dim strAxb17(0) As String 'Add byAmy 2023/04/17

'Add by Amy 2022/02/17 安全基金撥補會於關閉後才加傳票,故SalesPoint 不會有值
Private Sub Cmd_AddSalesPoint_Click()
    Dim RsQ As New ADODB.Recordset, rsA As New ADODB.Recordset
    Dim strQ As String, intQ As Integer, strA As String, intA As Integer
    Dim strDate As String, strTime As String
    
    '判斷是否已有傳票
    strDate = Left(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##"))), 6)
    strQ = "Select st15,Acc021.* From Acc020,Acc021,Staff Where a0201<>'L' " & _
                "And a0205>=" & Val(strDate) - 191100 & "01 And a0205<=" & Val(strDate) - 191100 & "31 " & _
                "And ax209='M0109' And a0201=ax201(+) And a0202=ax202(+) And SubStr(ax205,1,1)='4' " & _
                "And ax209=st01(+) "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        strA = "Select * From SalesPoint Where sp01=" & strDate & " And sp02='M0109' "
        intA = 1
        Set rsA = ClsLawReadRstMsg(intA, strA)
        If intA = 1 Then
            MsgBox Val(strDate) - 191100 & "月已有安全基金！"
        Else
            strTime = Now
            strQ = "Insert Into SalesPoint (SP01,SP02,SP48,SP15,SP16,SP17,SP18,SP36,SP37,SP38,SP39) Values " & _
                       "(" & strDate & ",'" & RsQ.Fields("ax209") & "','" & RsQ.Fields("st15") & "'" & _
                       ",0,'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'HH24MISS')" & _
                       ",0,'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'HH24MISS'))"
            cnnConnection.Execute strQ, intQ
            If intQ = 1 Then
                MsgBox Val(strDate) - 191100 & "月安全基金已新增！"
            End If
        End If
        Set rsA = Nothing
    Else
        MsgBox Val(strDate) - 191100 & "月未有安全基金撥補之傳票！"
    End If
    Set RsQ = Nothing
End Sub

Private Sub cmdCheck_Click()
    Dim bCancel As Boolean
    If Trim(Text1) = MsgText(601) Then MsgBox "資料年月不可空白！", , MsgText(5): Exit Sub
    Text1_Validate bCancel
    If bCancel = True Then Exit Sub
    
    tool3_enabled
    Frmacc43c0_1.m_YearMon = Val(Text1)
    Frmacc43c0_1.Show
    Me.Enabled = False
End Sub

'開放主管寫報告鈕
Private Sub cmdMail_Click()
    Dim RsQ As New ADODB.Recordset
    Dim stSQL As String, stTO As String, stSubject As String
    
    If MsgBox("確認要寄E-mail通知區主管，要繼續？", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    If (Val(Left(GetA0b01(strExc(0), "1"), 5)) < Val(Text1)) Or (Val(Left(GetA0b01(strExc(0), "J"), 5)) < Val(Text1)) Then
        MsgBox "當月資料尚未過帳,暫不開放！", , MsgText(5)
        Exit Sub
    End If
    
    '讀取業務各區主管(只有S部門)
    'Memo by Amy 2021/01/07 瑞婷過帳後才發現有69005結餘轉撥給10051沒輸隱藏版,此不需改,因部門主管員編不會有<'6' 或 >'F'
    'Modify by Amy 2024/09/06 拿掉sp48 及Order by sp48 ,因1130906 杜燕文協理出差,因目前為3區確認主管,導致組信件內容時[出差]字樣組了3次
    stSQL = "Select Distinct a0908 From SalesPoint,Acc090,Staff Where sp48=a0901(+) And sp02=st01(+) And sp01=" & Val(Text1) + 191100 & _
                " And SubStr(sp48,1,1)='S' And Decode(st04,2,'F0000',sp02)>='6' And Decode(st04,2,'F0000',sp02)<'F' Order by a0908 "
    RsQ.CursorLocation = adUseClient
    RsQ.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            stTO = stTO & ";" & RsQ.Fields("a0908")
            RsQ.MoveNext
        Loop
        stSubject = Left(Val(Text1), 3) & "年 " & Val(Right(Val(Text1), 2)) & "月點數已結算完成，您可逕行下載該月工作報告！"
    End If
    If stTO <> MsgText(601) Then
        'Modify by Amy 2016/02/15 +瑞婷無法知道是否已處理,故加發給瑞婷
        stTO = Mid(stTO, 2) & ";" & Pub_GetSpecMan("財務處總帳人員")
        PUB_SendMail strUserNum, stTO, "", stSubject, "如主旨"
        MsgBox "已發E-mail通知區主管！"
    End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
    Dim strCmd As String, intRec As Integer 'Add by Amy 2021/05/28
    Dim stMsg As String 'Add by Amy 2022/06/09
    Dim stAxb1(0) As String 'Add by Amy 2022/06/10 結餘是否有修改
    
    'Modify by Amy 2022/05/09 原判斷改至FormCheck
    stMsg = IIf(Index = 0, "開放", "關閉")
    If FormCheck("cmdok", stMsg) = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Select Case Index
        Case 0 '開放
            'Modify by Amy  2021/05/28 +if 關閉後開放
            'Modify by Amy 2022/06/09 取前5個字
            If Left(grdDataList.TextMatrix(1, 1), 5) = "關閉後開放" Then
                If grdDataList.TextMatrix(1, 0) = MsgText(601) Then
                
                Else
                    If Right(strA0b05, 2) = "01" Then
                        strCmd = Val(strA0b05) + 191011
                    Else
                         strCmd = Val(strA0b05) + 191099
                    End If
                    strCmd = Val(strCmd) - 191100
                    strCmd = "Update Acc0b0 Set a0b05=" & strCmd
                    cnnConnection.Execute strCmd, intRec
                    If intRec = 0 Then
                        MsgBox "關閉後開放有誤,請洽電腦中心！"
                    Else
                        MsgBox "「關閉後開放」已執行完成！"
                        Call doQuery(False)
                    End If
                End If
            '其他部門開放
            Else
                Call OpenSalesPoint
            End If
        Case 1 '關閉
            'Add by Amy 2022/06/10 結餘資料有修改(axb16=Y),智權期末結餘保留資料需刪除
            Call bolAcc0b1(8, Text1, stAxb1())
            If stAxb1(0) = "Y" Then
                Screen.MousePointer = vbDefault
                If TranNoLock("Frmacc41k0", "Frmacc41g0") = False Then
                   Exit Sub
                End If
                Frmacc41k0.Show vbModal
                Call bolAcc0b1(8, Text1, stAxb1())
                If stAxb1(0) = "Y" Then
                    MsgBox "「結餘」資料有修改" & vbCrLf & _
                                  "結餘保留「分配」資料已產生(SalesBalance)" & vbCrLf & _
                                  "需刪除才能關閉！"
                Else
                    Screen.MousePointer = vbHourglass
                    Call UpdateA0b05
                End If
            Else
                Call UpdateA0b05
            End If
            
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub OpenSalesPoint()
    Dim RsQ As New ADODB.Recordset
    Dim stSQL As String, stTO As String, strSubject As String
    Dim stST15 As String, stA0908 As String, strMsg As String
    Dim intRec As Integer, j As Integer
    Dim bolF11Dept As Boolean, stToSpecID As String 'Add by Amy 2021/07/16 有F11部門資料/特殊發信人員
    
On Error GoTo ErrHand1:
    
    If MsgBox("確認要做「開放」處理" & IIf(Check1.Value = 1, vbCrLf & "此動作會同時寄E-mail，要繼續？", "？") _
        , vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    '第一次開放
    If IsSPExist(Val(Text1) + 191100) = False Then
        If SaveData(1, intRec) = True Then
            If intRec > 0 Then strMsg = Left(Text1, 3) & "年 " & Right(Text1, 2) & "月開放資料已產生！"
            If Check1.Value = 1 And intRec > 0 Then
                '讀取需寄mail通知的人員(sp02>'6' And sp02<'F' ex:10051不需發mail)
                'modify by sonia 2016/1/26 只通知在職人員,故加and st04='1'
                'Modify by Amy 2019/08/05 開放F4102王文安操作,F4103陳鳳英操作,發mail時抓F4102/03部門之區主管
                'Modify by Amy 2020/05/22 +W1001/W2001
                'Modify by Amy 2020/06/03 +20091發mail 通知林柄佑操作
                'Memo by Amy 2021/01/07 員編10051(XX區)...不需發mail
                'Modify by Amy 2021/01/18 原非智權抓'F4102','F4103','W1001','W2001','20091' 改智權點數實績與結餘特殊員編
                'Modify by Amy 2021/07/16 F1(外商)改抓st14,F4106 洪琬姿/F4107 葉易雲 輸
                'Modify by Amy 2021/12/08 P2005 改由沈佳穎輸,江郁仁確認,其餘非智權部都改抓st14(避免P20沒資料,故抓Staff)
'                "Union Select Distinct a0908 as sp02,sp48 as st15 From SalesPoint,Acc090 Where sp01=" & Val(Text1) + 191100 & _
'                            " And sp02 In ('" & Replace(智權點數實績與結餘特殊員編, ";", "','") & "') And sp48=a0901(+) And SubStr(sp48,1,2)<>'F1' " & _
'                 "Union Select Distinct st14 as sp02,sp48 as st15 From SalesPoint,Staff Where sp01=" & Val(Text1) + 191100 & _
'                            " And sp02 In ('" & Replace(智權點數實績與結餘特殊員編, ";", "','") & "') And sp48=st15(+) And SubStr(sp48,1,2)='F1' And sp02=st01(+) "
                stSQL = "Select sp02,st15 From SalesPoint,Staff Where sp02=st01(+) And sp01=" & Val(Text1) + 191100 & _
                            " And st04='1' And sp02>'6' And sp02<'F' And SubStr(st15,1,1)='S' " & _
                "Union Select Distinct a0908 as sp02,st15 From Staff,Acc090 Where st15='P20' And st15=a0901(+) " & _
                "Union Select Distinct st14 as sp02,sp48 as st15 From SalesPoint,Staff Where sp01=" & Val(Text1) + 191100 & _
                            " And sp02 In ('" & Replace(智權點數實績與結餘特殊員編, ";", "','") & "') And sp48=st15(+) And sp02=st01(+) And st14<>'99997' " & _
                "Order by st15,sp02"
                
                RsQ.CursorLocation = adUseClient
                RsQ.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
                If RsQ.RecordCount > 0 Then
                    RsQ.MoveFirst
                    Do While RsQ.EOF = False
                        'Modify by Amy 2021/06/21 +if 陳鳳英退休改由江郁仁(98020)操作
                        'Modify by Amy 2021/07/16 原判斷有F11 部門mail給 98020,改判斷是否有F11部門資料要輸(80030輸F4106/78011輸F4107)
                        If "" & RsQ.Fields("st15") = "F11" Then
                            bolF11Dept = True
                        End If
                        stTO = stTO & ";" & RsQ.Fields("sp02")
                        'end 2021/07/16
                        RsQ.MoveNext
                    Loop
                    'Add by Amy 2016/03/03 +通知淑芳-瑞婷
                    stTO = stTO & ";87027"
                    'Add by Amy 2021/07/16 有F11 部門也需通知江郁仁(98020)
                    If bolF11Dept = True Then
                        '判斷江郁仁是否請假一整天
                        Call Set98020Ag(stToSpecID)
                        If stToSpecID = MsgText(601) Then
                            stTO = stTO & ";98020"
                        End If
                    End If
                    'Modified by Lydia 2019/07/03 更名
                    'strSubject = Left(Text1, 3) & "年 " & Right(Text1, 2) & "月財務收款已輸入完畢，請同仁至 智權部->日常工作->智權點數實績與結餘輸入 功能輸入您本月欲做點數 !"
                    strSubject = Left(Text1, 3) & "年 " & Right(Text1, 2) & "月財務收款已輸入完畢，請同仁至 智權部->財務作業->每月點數查詢／輸入 功能輸入您本月欲做點數 !"
                End If
            End If
        Else
            Exit Sub
        End If
    'SalesPoint已有資料
    Else
        With grdDataList
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, 0)) = "V" Then
                    .TextMatrix(i, 0) = ""
                    stST15 = stST15 & ",'" & .TextMatrix(i, GetValue("ST15")) & "'"
                    'Modify by Amy 2021/06/21 +if 陳鳳英退休改由江郁仁(98020)操作
                    If InStr(stST15, "F11") > 0 Then
                        'Modify by Amy 2021/07/16 再開放 F11 部門通知江郁仁(98020),若請假通知特殊職代-秀玲:同智權部
                        '判斷江郁仁是否請假一整天
                        Call Set98020Ag(stToSpecID)
                        If stToSpecID = MsgText(601) Then
                            stA0908 = "98020"
                        End If
                    Else
                        stA0908 = GetDeptMan(.TextMatrix(i, GetValue("ST15"))) '部門主管
                    End If
                    'Modify by Amy 2021/07/16 +if stToAgID = MsgText(601),因江郁仁請假mail 需另外發
                    If stToSpecID = MsgText(601) Then
                        If stA0908 = .TextMatrix(i, GetValue("ST01")) Then
                            stTO = stTO & ";" & .TextMatrix(i, GetValue("ST01"))     '確認主管ID(可能為職代)
                        Else
                            '若為職代確認,則原區主管也要發
                            stTO = stTO & ";" & .TextMatrix(i, GetValue("ST01")) & ";" & stA0908
                        End If
                    End If
                   
                    For j = 0 To .Cols - 1
                        .col = j
                        .CellBackColor = &HFFC0C0
                    Next j
                End If
            Next i
        End With
        'Modified by Lydia 2019/07/03 更名
        'strSubject = Left(Text1, 3) & "年 " & Right(Text1, 2) & "月業績輸入已開放，請同仁至 智權部->日常工作->智權點數實績與結餘輸入 功能修改 !"
        strSubject = Left(Text1, 3) & "年 " & Right(Text1, 2) & "月業績輸入已開放，請同仁至 智權部->財務作業->每月點數查詢／輸入 功能修改 !"
        If stST15 <> MsgText(601) Then
            If SaveData(2, intRec, stST15) = True Then
                If intRec > 0 Then strMsg = "勾選的部門資料已開放！"
            Else
                Exit Sub
            End If
        Else
            strMsg = "請勾選欲開放的業務區！"
        End If
    End If
    
    ' 勾選發mail才發
    If intRec > 0 Then
        If stTO <> MsgText(601) And Check1.Value = 1 Then
            'Modify by Amy 2016/02/15 +瑞婷無法知道是否已處理,故加發給瑞婷
            stTO = Mid(stTO, 2) & ";" & Pub_GetSpecMan("財務處總帳人員")
            PUB_SendMail strUserNum, stTO, "", strSubject, "如主旨"
            strMsg = strMsg & vbCrLf & "通知的E-mail已發送！"
        End If
        'Add by Amy 2021/07/16 江郁仁請假,因不發江郁仁人事職代,故另外發
        If stToSpecID <> MsgText(601) And Check1.Value = 1 Then
            PUB_SendMail strUserNum, stToSpecID, "", strSubject, "因收件人江郁仁，請副本收件人處理此郵件。" & vbCrLf & vbCrLf, , , , , , , , , , True
            If strMsg = MsgText(601) Then strMsg = "通知的E-mail已發送！"
        End If
        'end 2021/07/16
        Call doQuery(False)
    End If
    If strMsg <> MsgText(601) Then MsgBox strMsg
    Exit Sub

ErrHand1:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub UpdateA0b05()
    Dim strMsg As String
    Dim intRec As Integer
    
On Error GoTo ErrHand2:
    If CheckAccept(strMsg) = False Then
        MsgBox strMsg & vbCrLf & "區主管尚未確認，業績輸入不可關閉！"
        Exit Sub
    End If
    'Mark by Amy 2023/06/05 不可做在此,否則開放前有資料,會無法關閉(因會一直彈要先產生「期末實績」的傳票
'    'Add by Amy 2023/04/17 確認 ACS分潤當月合計是否與期末金額相符
'    If ChkACSIncomeAndEndAmt(Me.Name, Text1, , strMsg) = False Then
'        MsgBox strMsg
'        Exit Sub
'    End If
    If CheckBalance = False Then
        MsgBox "轉撥金額不平衡，請再確認...！"
        Exit Sub
    End If
    
    If MsgBox("確認要做「關閉」處理，要繼續？", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
        
    If SaveData(3, intRec) = True Then
        Call doQuery(False) 'Add byAmy 2021/05/28 重讀資料,因加「關閉後開放」
        MsgBox "業績輸入已關閉！"
    Else
        Exit Sub
    End If
    Exit Sub
    
ErrHand2:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdSearch_Click()
    
    'Modify by Amy 2022/06/09 原檢查改至FormCheck
    If FormCheck("cmdSearch") = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    doQuery
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    Dim intX As Integer
    Dim intY As Integer
    Dim sglWidth As Single
    Dim sglHeight As Single

    strFormName = Name
    'Modify by Amy 2023/06/14 +提醒文字,改長及寬
    Me.Width = 8670
    Me.Height = 6500
    sglWidth = (lngWidth - Me.Width) / 2
    sglHeight = (lngHeight - Me.Height) / 2
    If sglHeight < 0 Then sglHeight = 10
    Me.Move sglWidth, sglHeight
    
    'Add by Amy 2016/11/03 +提示文字-瑞婷
    Label2 = "1.每月點數請先產生報表並檢查有無應轉撥點數" & vbCrLf & _
                   "2.若手動調整轉撥請記得傳票及電腦實績與結餘都要輸入" & vbCrLf & _
                   "3.包含各區應轉撥點數應再確認有無做傳票及輸入" & vbCrLf & _
                   "4.各區點數關閉後請重印智權點數與實績分析表 (請一定要重新產生)"
'Add by Amy 2023/06/14 +提示文字-辜
    Label2 = Label2 & vbCrLf & "5.[結餘] or [結餘轉撥] 相關輸入產生傳票後又修改" & vbCrLf & _
                     "　關閉->智權期未結餘保留資料刪除->智權期未結餘保留傳票產生 按「更正傳票」" & vbCrLf & _
                     "6.已有[智權期未結餘保留]傳票又於[隱藏版] 新增 or 修改" & vbCrLf & _
                     "　開放->隱藏版修正後->[沒有]傳票號, 按「產生傳票」,[有]傳票號, 按「更正傳票」" & vbCrLf & _
                     "　關閉->智權期未結餘保留資料刪除->進 智權期未結餘保留傳票產生 按「更正傳票」"
    '業績年月預設系統日前一個月
    Text1 = Val((CompDate(1, -1, strSrvDate(2)) - 19110000)) \ 100
    SetDataListWidth
    Check1.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    Set Frmacc43c0 = Nothing
End Sub

Private Sub SetDataListWidth()
    Dim iCol As Integer
   
    'Modify by Amy 2021/05/28 +Sort
    ReDim strColN(0 To 7)
    ReDim intWidth(0 To 7)
    
    strColN = Array("V", "業務區", "確認主管", "日期", "時間", "ST15", "ST01", "Sort")
    intWidth = Array(300, 1800, 1000, 1000, 1000, 0, 0, 0) 'Modify by Amy 業務區加寬(+已過帳)
    'end 2021/05/28
    
    With grdDataList
        .Visible = False

        For iCol = 0 To UBound(strColN)
            .ColWidth(iCol) = intWidth(iCol)
            .TextMatrix(0, iCol) = strColN(iCol)
        Next
        .Visible = True
    End With
End Sub

Private Sub doQuery(Optional ByVal bolShowMsg As Boolean = True)
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strMsg As String
    Dim strField As String, strWhere As String 'Add byAmy 2019/10/24
    Dim strQ1 As String 'Add by Amy 2021/05/28
    Dim strState As String 'Add by Amy 2022/06/09
    
    'Modify by Amy 2016/09/05 9/2台南杜經理(部門S31)請假由79053(部門為M71)代為操作,因確認主管存79053,抓部門時抓到M71,故無法開放S31
    '    strQ = "Select Distinct '' as V,A0902,st02,SqlDatet(sp46) as SP46,Decode(length(sp47),6,Substr(sp47,1,2)||':'||Substr(sp47,3,2),'0'||Substr(sp47,1,1)||':'||Substr(sp47,2,2)) as SP47 " & _
    '                ",st15,st01 " & _
    '                "From SalesPoint,Staff,Acc090 Where SP01='" & Val(Text1) + 191100 & "' And  SP45=ST01(+) And ST15=A0901(+) And SP45 is not null " & _
    '                "Order by st15,st01"
        'Modify by Amy 2019/08/05 開放F4102王文安操作,F4103陳鳳英操作,F部門顯示2個字部門名稱 原:A0902
        'Modify by Amy 2019/09/03 王文安掛多個部門,導致顯示多個
        'Modify by Amy 2019/10/24 開放W部門操作,且W/F部門可能由財務處做「主管確認」, 若為財務做確認不需再開放(總經理欄位有值區主管不可再改)可直接改資料
    '    strQ = "Select '' as V,Decode(SubStr(A0901,1,1),'F',SubStr(A0902,1,2),A0902) AS A0902,ST02, SP46, SP47 , ST15, ST01 " & _
    '               "From (Select Distinct Decode(SP45,'88003','F21',Decode(a0901, 'F10','F11',a0901)) AS DepNo,st02,SqlDatet(sp46) as SP46,Decode(length(sp47),6,SubStr(sp47,1,2)||':'||SubStr(sp47,3,2),'0'||SubStr(sp47,1,1)||':'||SubStr(sp47,2,2)) as SP47 " & _
    '                                                ",Decode(a0901,'F10','F11',Decode(SP45,'88003','F21',a0901)) as ST15,SP45 as ST01 " & _
    '                         "From SalesPoint ,Acc090,Staff Where SP01=" & Val(Text1) + 191100 & " And SP45 is not null And InStr(a0908||a0914,SP45)>0 And SP45=ST01(+) ),Acc090 " & _
    '               "Where DepNo=a0901(+) Order by st15,st01"
    'Modify by Amy 2020/06/18 無法得知需確認部門,開放 柄佑 輸20091(S29部門)及確認且舊資料有些不需確認(ex:40011),故只有當月未關閉前用新語法
    'Modify by Amy 2022/06/09 重抓strA0b01/05
    'strA0b01 = GetA0b01(strA0b05)
    strState = ""
    Call FormCheck("Read", strState)
    'end 2022/06/09
    'Modify by Amy 2021/06/21 簡化-抓sp45當下確認人員,因江郁仁(L01)輸外商點數
    '點數輸入[未]關閉
    If Val(Text1) > Val(strA0b05) Then
        strWhere = " And sp01=" & Val(Text1) + 191100
        strWhere = strWhere & " And (SubStr(sp48,1,1)='S' Or (SubStr(sp48,1,1)<>'S' And sp02 in('" & Replace(智權點數實績與結餘特殊員編, ";", "','") & "') )) " 'Modify by Amy 2021/05/28 開放林純真輸P2005
        'Modify by Amy 2021/01/07 原:SubStr(sp48,1,1) in ('S','F','W')
        'Modify by Amy 2021/05/28 與下方語法欄位一致,否則會Error
        strQ1 = "Select Distinct sp48 as Dept,SP46,SP47,SP45 From SalesPoint Where SP45 is not null " & strWhere
 
        strQ = "Select '' as V,Decode(SubStr(A0901,1,1),'F',SubStr(A0902,1,2),A0902) AS A0902,st02" & _
                             ",SqlDatet(sp46) as sp46,Decode(sp47,null,'',Decode(length(sp47),6,SubStr(sp47,1,2)||':'||SubStr(sp47,3,2),'0'||SubStr(sp47,1,1)||':'||SubStr(sp47,2,2)) ) as SP47,sp48,st01,'' as Sort " & _
            "From (Select Distinct sp48 From SalesPoint Where SubStr(sp48,1,1) in ('" & Replace(智權點數實績與結餘輸入部門, ",", "','") & "') " & strWhere & ")," & _
                     "(" & strQ1 & "),Staff,Acc090 Where sp48=Dept(+) and sp45=st01(+) And Sp48=a0901(+) Order by sp48,st01"
                     
    '點數輸入[已]關閉
    Else
        'Modify by Amy 2021/06/21 簡化-抓sp45當下確認人員
        'Memo by Amy 2020/06/18 InStr 條件+sp45 查舊資料 10502 S13部門會出不來,因其主管已離職
        'Modify by Amy 2021/05/28 所有需輸入之部門主管已確認,過帳前若需再開放則多顯示「 關閉後開放」-瑞婷
        strField = ",SqlDatet(sp46) as SP46,Decode(length(sp47),6,SubStr(sp47,1,2)||':'||SubStr(sp47,3,2),'0'||SubStr(sp47,1,1)||':'||SubStr(sp47,2,2)) as SP47,SP45 "
        strWhere = "And SP01=" & Val(Text1) + 191100 & " "
        strQ = "": strQ1 = ""
        '上個月且所有主管已確認
        'Modify by Amy 2021/01/05 bug 系統日-1個月,寫法有問題 ex:11101月操作11012月
        'Modify by Amy 2022/06/09 拿掉 And Left(Val(strA0b01) + 19110000, 6) < Val(Text1) + 191100 ,若當月已過帳不會出現「關閉後開放」
        If Val(Left(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##"))), 6)) - 191100 And CheckAccept = True Then
            '判斷業績輸入關閉年月為系統日-1個月則顯示「關閉後開放」
            'Modify by Amy 2022/06/09 +stState
            strQ = "Select '' as V ,'關閉後開放" & strState & "' AS A0902, '' as ST02, '' as SP46, '' as  SP47 , '' as SP48, '' as ST01,1 as Sort From Acc0b0 Where a0b05+191100=" & Val(Left(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##"))), 6))
            strQ = strQ & " Union "
        End If
        strQ1 = strQ1 & "Select Distinct SP48" & strField & "From SalesPoint Where SP45 is not null " & strWhere
                                  
        strQ = strQ & "Select '' as V,Decode(SubStr(A0901,1,1),'F',SubStr(A0902,1,2),A0902) AS A0902,ST02, SP46, SP47 , SP48, SP45,2 as Sort From (" & _
                                strQ1 & "),Acc090,Staff Where Sp48=a0901(+) And Sp45=St01(+) Order by Sort,sp48,st01 "
        'end 2021/05/28
    End If
    'end 2021/06/21
    'end 2020/06/18
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    grdDataList.Clear
    If RsQ.RecordCount > 0 Then
        Set grdDataList.Recordset = RsQ
        'Add by Amy 2021/01/07 瑞婷過帳後才發現有69005結餘轉撥給10051沒輸隱藏版,補輸後10051會沒確認
        '                                         應顯示不同才可通知主管確認
        Call SetGridListColor
    Else
        grdDataList.Rows = 2
        strMsg = "資料庫中搜尋不到符合資料!!"
    End If
    SetDataListWidth
    If strMsg <> MsgText(601) And bolShowMsg = True Then MsgBox strMsg, , "沒有資料"
    RsQ.Close
End Sub

Private Sub GrdDataList_Click()
    Dim strMsg As String 'Add by Amy 2021/05/28
    
    With grdDataList
        .Visible = False
        .col = 0
        'Modify by Amy 2021/05/28 if 關閉後開放
        If .TextMatrix(1, 1) = "關閉後開放" Then
            If .row = 1 Then
                If .Text = "V" Then
                    .Text = ""
                    For i = 0 To .Cols - 1
                        .col = i
                        .CellBackColor = vbRed
                    Next i
                Else
                    .Text = "V"
                    For i = 0 To .Cols - 1
                        .col = i
                        .CellBackColor = &HFFC0C0
                    Next i
                End If
            Else
                strMsg = "請先執行「關閉後開放」才能再開放其他部門！"
            End If
        '其他部門
        Else
            If .row <> 0 Then
                If .Text = "V" Then
                    .Text = ""
                    For i = 0 To .Cols - 1
                        .col = i
                        .CellBackColor = QBColor(15)
                    Next i
                Else
                    .Text = "V"
                    For i = 0 To .Cols - 1
                        .col = i
                        .CellBackColor = &HFFC0C0
                    Next i
                End If
            End If
        End If
        .Visible = True
    End With
    'Add by Amy 2021/05/28
    If strMsg <> MsgText(601) Then
        MsgBox strMsg
    End If
End Sub

Private Sub Text1_GotFocus()
    TextInverse Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Trim(Text1) = MsgText(601) Then Exit Sub
    
    If ChkDate(Val(Text1) & "01") = False Then
        Text1_GotFocus
        Cancel = True
    ElseIf Val(Text1) >= Val(Left(strSrvDate(2), 5)) Then
        MsgBox "資料年月有誤，請確認！", , MsgText(5)
        Text1_GotFocus
        Cancel = True
    End If
    
End Sub

'檢查記錄於SalesPoint是否已經存在
Private Function IsSPExist(ByVal strKEY01 As String, Optional ByRef stGetData As String = "*") As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select " & stGetData & " From SalesPoint Where SP01=" & Val(strKEY01)
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    ' 檢查讀取的資料筆數
    If rsTmp.RecordCount > 0 Then
        IsSPExist = True
        If stGetData = "*" Then
            stGetData = ""
        Else
            stGetData = "" & rsTmp.Fields(0)
        End If
    Else
        IsSPExist = False
    End If
    rsTmp.Close
    Set rsTmp = Nothing
End Function

'確認所有S部門是否區主管已確認
Private Function CheckAccept(Optional ByRef stSt15List As String) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim j As Integer
    
    stSt15List = "": j = 1
    
    'Modify by Amy 2021/01/07 10051/20091...st06小於6也可能要確認,因瑞婷過帳後才發現有69005結餘轉撥給10051沒輸隱藏版
    '                                             frm210152 2019/10/16 有調整,但此未調到,怕不一致,故改至共用Function
'    strSql = "Select Distinct a0902,b.st02,a.st15 From SalesPoint,Staff a,Acc090,Staff b " & _
'                "Where sp01=" & Val(Text1) + 191100 & " And sp02=a.st01(+) " & _
'                "And SubStr(a.st15,1,1)='S' And a.st15=a0901(+) And sp45 is null " & _
'                "And Decode(a.st04,2,'F0000',a.st01)>='6' And Decode(a.st04,2,'F0000',a.st01)<'F' " & _
'                "And a0908=b.st01(+) Order by st15"
    strSql = ChkPointAcceptSql(Val(Text1) + 191100, Me.Name)
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount = 0 Then
        CheckAccept = True
    Else
        rsTmp.MoveFirst
        Do While rsTmp.EOF = False
            stSt15List = stSt15List & " / " & rsTmp.Fields("a0902") & "-" & _
            IIf(rsTmp.Fields("sp48") = "S00", "財務處", rsTmp.Fields("st02")) & IIf(j Mod 4 = 0, vbCrLf, "")
            j = j + 1
            rsTmp.MoveNext
        Loop
        If stSt15List <> MsgText(601) Then stSt15List = Mid(stSt15List, 4)
        CheckAccept = False
    End If
End Function

'確認轉撥欄位是否合計為0
Private Function CheckBalance() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strQ As String
    
    CheckBalance = False
    strQ = "Select Sum(sp19) as sp19,Sum(sp40) as sp40 From SalesPoint " & _
                "Where sp01=" & Val(Text1) + 191100
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        If Val("" & rsTmp.Fields("sp19")) = 0 And Val("" & rsTmp.Fields("sp40")) = 0 Then
            CheckBalance = True
        End If
    End If
    rsTmp.Close
    Set rsTmp = Nothing
End Function

Private Function GetValue(pRowN As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(strColN)
       If UCase(strColN(jj)) = UCase(pRowN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Private Function SaveData(intChoose As Integer, ByRef intRec As Integer, Optional ByVal strST15 As String = "", Optional ByVal strA0b05 As String) As Boolean
    Dim strUpd As String, strErr As String
    Dim strSPData As String, strACSAmt As String 'Add by Amy 2023/04/17 SalesPointACS 資料/ACS 期初+當月收入

On Error GoTo ErrTran:
    intRec = 0: SaveData = False
    cnnConnection.BeginTrans
    Select Case intChoose
        '第一次開放
        Case 1
            strErr = Val(Text1) + 191100 & " 第一次開放有誤！" & vbCrLf
'*** Memo by Amy 2024/07/04 此處抓取人員有修改要確認下方,[關閉]人員是否也要改 ***
            'Modfiy by Amy 2016/04/07 避免P1001抓不到,故函數傳入參數將st01<'F'拿掉
            strUpd = "Insert Into SalesPoint (sp01,sp02,sp48) " & _
                                GetPoint(2, Text1, Text1, , , , , Me.Name, True)
            cnnConnection.Execute strUpd, intRec
'*** End Memo by Amy 2024/07/04 此處抓取人員有修改要確認下方,[關閉]人員是否也要改 ***

            'Add by Amy 2018/02/06 M0100 frmacc44j0報表抓傳票拆各部門,故不需新增至SalesPoint
            'Modify by Amy 2021/05/21 +M0109 不需輸 智權點數實績與結餘輸入 改共用變數
            strUpd = "Delete From SalesPoint Where SP01=" & Val(Text1) + 191100 & " And SP02 In ('" & Replace(不需新增SalesPoint人員, ",", "','") & "') "
            cnnConnection.Execute strUpd
            
            '更新業務區
            'Modfiy by Amy 2016/03/25 +預設上個月勾選
            'Modify by Amy 2016/11/24 +文雄員編部門設北四
            strUpd = Val(Mid(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(DBDATE(Left(Val(Text1) + 191100, 6) & "01")))), 1, 6))
            strUpd = "Update SalesPoint a Set sp48=(Select Decode(st01,'A4023','S14',st15) From Staff Where sp02=st01) " & _
                            ",sp49=(Select SP49 From SalesPoint b Where sp01=" & strUpd & " And a.sp02=b.sp02)" & _
                            ",sp50=(Select SP50 From SalesPoint b Where sp01=" & strUpd & " And a.sp02=b.sp02)" & _
                            ",sp51=(Select SP51 From SalesPoint b Where sp01=" & strUpd & " And a.sp02=b.sp02)" & _
                           "Where sp01=" & Val(Text1) + 191100
            cnnConnection.Execute strUpd
        '畫面勾選開放區主管已確認之業務區
        Case 2
            strErr = "業務區開放有誤！" & vbCrLf
            strUpd = "Update SalesPoint Set sp45=null,sp46=null,sp47=null Where sp01=" & Val(Text1) + 191100 & " And sp02 in (" & _
                    "Select sp02 From SalesPoint,Staff Where sp01=" & Val(Text1) + 191100 & " And sp02=st01(+) And st15 in(" & Mid(strST15, 2) & ") " & _
                    ")"
            cnnConnection.Execute strUpd, intRec
        '關閉
        Case 3
            strUpd = "Update Acc0b0 Set a0b05=" & Val(Text1)
            cnnConnection.Execute strUpd
'*** Memo by Amy 2024/07/04 此處抓取人員有修改要確認上方,[第一次開放]人員是否也要改 ***
            'Add by Amy 2017/05/15 for 自動產生傳票
            '新增當月有傳票SalesPoint 沒資料
            'Modify by Amy 2018/02/06 SalesPoint Trigger 修改會自動更新人員/日期/時間,且M0100 不新增(frmacc44j0 報表抓傳票拆各部門)
            'Modify by Amy 2019/10/16 開放W部門使用原只抓會計科目為41字頭
            'Modify by Amy 2021/05/28 排除M0109 改抓共用變數
            'Modify by Amy 2024/07/04 改為[同開放],因11306月做A3002 謝俊民(中一區) 轉專業收入傳票 (借:420101 貸:411102/413101...),關閉後此人員寫入SalesPoint
            '                                                     關閉後又開放,會出現中一區有未確認之人員,故此人員(收入借貸=0之調整)不應該寫入SalesPoint
'            strUpd = "Insert Into SalesPoint (SP48,SP01,SP02) " & _
'                        "Select Distinct ST15," & Val(Text1) + 191100 & ",ax209" & _
'                        " From Acc021, Acc020,Staff Where ax201(+) = a0201 And ax202(+) = a0202 And ax209 is not null " & _
'                            "And (SubStr(ax205, 1, 1) = '4' Or ax205='7121') And A0205 >= " & Val(Text1) & "01 And A0205 <= " & Val(Text1) & "31 " & _
'                            "And ax209=st01(+) And  ax209 not in (Select SP02 From SalesPoint Where  SP01=" & Val(Text1) + 191100 & ") And ax209 not in ('" & Replace(不需新增SalesPoint人員, ",", "','") & "')"
            strUpd = Replace(UCase(GetPoint(2, Text1, Text1, , , , , Me.Name, True)), "ST01,ST02,ST04,ST05", "ST01,ST02,ST15")
            'Modify by Amy 2024/11/07 原UCase("Distinct 202406,ST01,'ZZZ'")-bug
            strUpd = Replace(strUpd, UCase("Distinct " & Val(Text1) + 191100 & ",ST01,'ZZZ'"), "DISTINCT ST15," & Val(Text1) + 191100 & ",ST01")
            strUpd = "Insert Into SalesPoint (SP48,SP01,SP02) " & strUpd & " And  st01 not in (Select SP02 From SalesPoint Where  SP01=" & Val(Text1) + 191100 & ") " & _
                              "And st01 not in ('" & Replace(不需新增SalesPoint人員, ",", "','") & "')"
            cnnConnection.Execute strUpd
'*** End Memo by Amy 2024/07/04 此處抓取人員有修改要確認上方,[第一次開放]人員是否也要改 ***
            
            'Add by Amy 2023/04/17 畫面年月若當月ACS需分潤的案子有收款,需增加M0101 期末實績保留=期末實績保留+當月實績
            strSPData = Pub_GetField("SalesPoint", "SP01||SP02='" & Val(Text1) + 191100 & "M0101'", "Nvl(SP15,0)", True)
            'ACS 期初+當月
            strExc(2) = "Select Sum(V11)/1000 From(" & GetPoint(1.1, Text1, Text1, , , "M0101", , Me.Name, True) & " Union All " & _
                                                                                               GetPoint(1.3, Text1, Text1, , , "M0101", , Me.Name, True) & " )"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(2))
            If intI = 1 Then
                strUpd = ""
                strACSAmt = "" & RsTemp.Fields(0)
                'SalesPoint 無 M0101資料
                If strSPData = "NoData" Then
                    strUpd = "Insert Into SalsePoint (sp01,sp02,sp48,sp15,sp36) Values(" & Text1 & ",'M0101'," & GetST15("M0101") & "," & strACSAmt & ",0)"
                'SalesPoint 有 M0101資料,傳票資料與SalesPoint不同,更新為期初+當月
                ElseIf Val(strSPData) <> Val(strACSAmt) Then
                    strUpd = "Update SalesPoint Set SP15=(" & strACSAmt & "),SP36=0 Where SP01=" & Val(Text1) + 191100 & " And SP02='M0101' "
                End If
                If strUpd <> MsgText(601) Then
                    cnnConnection.Execute strUpd
                End If
            End If
            'end 2023/04/17
            
            '更新未輸入期末值=0(全部欄位都未輸)
            'Modify by Amy 2018/02/06 Tigger 修改會自動更新人員/日期/時間,故將SP16='QPGMR',SP17=" & strSrvDate(1) & ",SP18=" & ServerTime & " ...拿掉
            strUpd = "Update SalesPoint Set SP15=0,SP36=0 Where SP01=" & Val(Text1) + 191100 & " And SP02 in " & _
                        "(Select SP02 From SalesPoint Where sp01=" & Val(Text1) + 191100 & " And sp03||sp07||sp11||sp15||sp19||sp20||sp24||sp28||sp32||sp36||sp40||sp41 is Null)"
            cnnConnection.Execute strUpd
            'end 2017/05/15
    End Select
    cnnConnection.CommitTrans
    
    SaveData = True
    Exit Function
    
ErrTran:
    cnnConnection.RollbackTrans
    MsgBox strErr & Err.Description, vbCritical
End Function

'Add by Amy 2021/01/07 同部門有主管已確認及未確認的資料(瑞婷過帳後才發現有69005結餘轉撥給10051沒輸隱藏版,後補資料應顯示有未確認)
Private Sub SetGridListColor()
    Dim rsA As New ADODB.Recordset
    Dim strA As String, intA As Integer, ii As Integer, jj As Integer
    
    'Modify by Amy 2021/05/28 調整加P2005後顯示之資料
    strA = "Select Distinct sp48 From SalesPoint Where sp01=" & Val(Text1) + 191100 & " And SubStr(sp48,1,1) ='S' And sp45 is null "
    strA = strA & " Union " & _
                "Select Distinct sp48 From SalesPoint Where sp01=" & Val(Text1) + 191100 & " And sp02 in('" & Replace(Replace(智權點數實績與結餘特殊員編, "20091;", ""), ";", "','") & "') And sp45 is null " & _
                " Order by sp48"
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, strA)
    If intA = 1 Then
        With grdDataList
            For ii = 1 To .Rows - 1
                If rsA.EOF = True Then Exit For
                If Val(Text1) > Val(strA0b05) Then
                    '已確認(未確認=只有部門名稱,其他欄空白)
                    If .TextMatrix(ii, 2) <> "" And .TextMatrix(ii, 5) = "" & rsA.Fields("sp48") Then
                        For jj = 0 To .Cols - 1
                            .row = ii: .col = jj
                            grdDataList.CellBackColor = vbGreen
                        Next jj
                        If rsA.EOF = False Then rsA.MoveNext
                    End If
                End If
            Next ii
        End With
    End If
    Set rsA = Nothing
    'Modify by Amy 2021/05/28 關閉後開放 (有此列就顯示紅色)
    'Modify by Amy 2022/06/09 取前5個字
    If Left(grdDataList.TextMatrix(1, 1), 5) = "關閉後開放" Then
        For jj = 0 To grdDataList.Cols - 1
            grdDataList.row = 1: grdDataList.col = jj
            grdDataList.CellBackColor = vbRed
        Next jj
    End If
End Sub

'Add by Amy 2021/07/16 江郁仁是否請假一整天,職代發葉易雲(78011)
Private Sub Set98020Ag(ByRef stToSpecNo As String)
    Dim bolIsRest As Boolean, bolRest1Day As Boolean '是請假/是請假一整天
    
    stToSpecNo = ""
    bolIsRest = CheckIsPersonRest("98020", strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolRest1Day)
    If bolIsRest = True And bolRest1Day = True Then
        stToSpecNo = "98020;78011"
    End If
End Sub

'Add by Amy 2022/06/09 避免cmdOK及cmdSearch有未檢查到
Private Function FormCheck(ByVal stCmdN As String, Optional ByRef stMsg As String) As Boolean
    Dim bCancel As Boolean
    
    FormCheck = False
    If UCase(stCmdN) <> UCase("Read") Then
        If Trim(Text1) = MsgText(601) Then MsgBox "資料年月不可空白！", , MsgText(5): Exit Function
        
        Text1_Validate bCancel
        If bCancel = True Then Exit Function
    End If
    'Add by Amy 2022/12/15 +畫面年月未有SalesPoint 資料不可按「關閉」
    '因關閉會只會Insert 傳票資料,有目標者不會 Insert導致智權的各區/所業務工作報告表目標會出不來
    If UCase(stCmdN) = UCase("cmdOK") And stMsg = "關閉" Then
        If IsSPExist(Val(Text1) + 191100) = False Then
            MsgBox Text1 & "月尚無點數資料不可按「關閉」！", , MsgText(5): Exit Function
        End If
    End If
    
    strA0b01 = GetA0b01(strA0b05)
    'modify by sonia 2016/1/26 先取消 Or Val(Text1) <> Val(strA0b05)
    'If Val(Text1) <= Val(Left(strA0b01, 5)) Or Val(Text1) <> Val(strA0b05) Then
    If Val(Text1) <= Val(Left(strA0b01, 5)) Then
        If UCase(stCmdN) = UCase("cmdOK") Then
            'Modify by Amy 2018/06/13 改訊息原:業績輸入已關閉不可再…
            MsgBox Text1 & "月傳票已過帳不可再" & stMsg & "！", , MsgText(5)
            Exit Function
        ElseIf UCase(stCmdN) = UCase("Read") Then
            stMsg = "(已過帳)"
        End If
    End If
    
    FormCheck = True
End Function


