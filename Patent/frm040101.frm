VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040101 
   BorderStyle     =   1  '單線固定
   Caption         =   "內專分案"
   ClientHeight    =   6360
   ClientLeft      =   -230
   ClientTop       =   2230
   ClientWidth     =   9340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9340
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4515
      Left            =   90
      TabIndex        =   22
      Top             =   1800
      Width           =   9135
      _ExtentX        =   16104
      _ExtentY        =   7955
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   14
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   120
      TabIndex        =   31
      Top             =   5520
      Width           =   9075
      Begin VB.CommandButton cmdFile 
         Caption         =   "檢視接洽單"
         CausesValidation=   0   'False
         Height          =   360
         Left            =   390
         TabIndex        =   35
         Top             =   345
         Width           =   1065
      End
      Begin VB.CommandButton Command1 
         Caption         =   "補件完成(&O)"
         Height          =   400
         Left            =   7500
         TabIndex        =   34
         Top             =   300
         Width           =   1125
      End
      Begin VB.TextBox txtCustNo 
         Height          =   705
         Index           =   1
         Left            =   2595
         MaxLength       =   6
         TabIndex        =   32
         Top             =   0
         Width           =   4845
      End
      Begin VB.Label Label1 
         Caption         =   "呈報內容："
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   33
         Top             =   75
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdNoneReceive 
      Caption         =   "未註記(&N)"
      Height          =   400
      Left            =   2880
      TabIndex        =   30
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "取消註記(&U)"
      Height          =   400
      Index           =   1
      Left            =   1665
      TabIndex        =   29
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "收到註記(&R)"
      Height          =   400
      Index           =   0
      Left            =   450
      TabIndex        =   28
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   6030
      MaxLength       =   4
      TabIndex        =   12
      Top             =   1485
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   6000
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1110
      Width           =   300
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   6720
      MaxLength       =   1
      TabIndex        =   11
      Top             =   1110
      Width           =   300
   End
   Begin VB.CommandButton ComSure 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7560
      TabIndex        =   16
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton ComBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8388
      TabIndex        =   17
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton ComAllData 
      Caption         =   "所有資料(&L)"
      Height          =   400
      Left            =   6336
      TabIndex        =   15
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton ComUCase 
      Caption         =   "未分案(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5325
      TabIndex        =   14
      Top             =   70
      Width           =   990
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "全部選取(&A)"
      Height          =   400
      Left            =   4110
      TabIndex        =   13
      Top             =   70
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   540
      Width           =   4692
      Begin VB.OptionButton Option1 
         Caption         =   "電子收文未分案:"
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   9
         Top             =   870
         Width           =   1752
      End
      Begin VB.TextBox txtGDate1 
         Height          =   270
         Index           =   1
         Left            =   2880
         TabIndex        =   1
         Top             =   180
         Width           =   1092
      End
      Begin VB.OptionButton Option1 
         Caption         =   "以前未分案 :"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   8
         Top             =   870
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號："
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   7
         Top             =   555
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "收文日期："
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtGDate1 
         Height          =   270
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   180
         Width           =   1092
      End
      Begin VB.TextBox txtcp01 
         Height          =   270
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   2
         Top             =   510
         Width           =   495
      End
      Begin VB.TextBox txtcp02 
         Height          =   270
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   3
         Top             =   510
         Width           =   1095
      End
      Begin VB.TextBox txtcp03 
         Height          =   270
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   4
         Top             =   510
         Width           =   375
      End
      Begin VB.TextBox txtcp04 
         Height          =   270
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   5
         Top             =   510
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   2640
         X2              =   2760
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.Frame Frame3 
      Height          =   525
      Left            =   4944
      TabIndex        =   18
      Top             =   540
      Width           =   4272
      Begin VB.OptionButton Option6 
         Caption         =   "接洽及內部收文單"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   120
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option7 
         Caption         =   "主管機關來函"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   180
         Width           =   1575
      End
   End
   Begin VB.Label Label5 
      Height          =   240
      Left            =   8160
      TabIndex        =   27
      Top             =   1485
      Width           =   975
   End
   Begin MSForms.Label Label4 
      Height          =   240
      Left            =   6720
      TabIndex        =   26
      Top             =   1485
      Width           =   1335
      VariousPropertyBits=   27
      Size            =   "2355;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "案件性質："
      Height          =   240
      Left            =   5040
      TabIndex        =   25
      Top             =   1485
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "收文所別："
      Height          =   240
      Index           =   0
      Left            =   5040
      TabIndex        =   24
      Top             =   1155
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   6360
      X2              =   6600
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "(1.北  2.中  3.南  4.高)"
      Height          =   240
      Left            =   7200
      TabIndex        =   23
      Top             =   1155
      Width           =   1695
   End
End
Attribute VB_Name = "frm040101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/3 改成Form2.0(Label4..)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
Option Explicit

Dim intLastRow As Integer, blnOKtoShow As Boolean
Dim intClick As Boolean 'True ComAllData ,False ComUCase
Public strSave As String
' 搜尋的方式 1:所有資料 2:未分案 3:未註記
Dim m_QueryType As Integer
Dim intOrderQty As Integer  'Add by Amy 2014/11/19 接洽單案件性質數量
Dim lngX As Long, lngY As Long 'Add by Amy 2014/12/04
Dim stDefArea1 As String, stDefArea2 As String 'Add by Amy 2022/11/15


Private Sub cmdNoneReceive_Click()
   m_QueryType = 3
   Screen.MousePointer = vbHourglass
   If CheckChoese(3) Then
      PutDataInGrid
   End If
   GridHead
   Screen.MousePointer = vbDefault
   intClick = True
End Sub

Private Sub cmdReceive_Click(Index As Integer)
   Dim i As Integer, bolGoNext As Boolean
   Dim strCP156 As String 'Add by Amy 2014/11/19 更新CP156
   With MSHFlexGrid1
      'Modify by Amy 2014/11/19 改以GetValue(避免插入欄位不好改)
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "v" Then
            'Modify by Morgan 2010/6/23
            'strSql = "Update Caseprogress Set CP86=" & IIf(Index = 0, "'Y'", "NULL") & " Where CP09='" & .TextMatrix(i, 1) & "'"
            'cnnConnection.Execute strSql
            '.TextMatrix(i, 18) = IIf(Index = 0, "Y", "")
            strCP156 = "" 'Add by Amy 2014/12/04
            If Index = 0 Then
               If GetValue(i, "註記") = "" Then '.TextMatrix(i, 18) = ""
                  'Modify by Amy 2014/11/19 +更新CP156
                  'Modify by Morgan 2016/6/22 非臺灣案電子化
                  'If Pub_StrUserSt03 = "P12" And Left(GetValue(i, "本所案號"), 2) = "P-" And GetValue(i, "申請國家") = 台灣國家代號 _
                     And Val(Replace(GetValue(i, "數量"), "-", "")) > 0 Then strCP156 = ",CP156=" & Val(GetValue(i, "數量"))
                  If Pub_StrUserSt03 = "P12" And Val(GetValue(i, "數量")) > 0 Then strCP156 = ",CP156=" & Val(GetValue(i, "數量"))
                  'end 2016/6/22
                  strSql = "Update Caseprogress Set CP86='Y'" & strCP156 & " Where CP09='" & GetValue(i, "收文號") & "'" '.TextMatrix(i, 1)
                  cnnConnection.Execute strSql
                  .TextMatrix(i, Val(GetValue(0, "註記"))) = "Y"
               End If
            Else
               If GetValue(i, "註記") = "Y" Then '.TextMatrix(i, 18) = "Y"
                  'Modify by Amy 2014/11/19 +取消CP156
                  'Modify by Morgan 2016/6/22 非臺灣案電子化
                   'If Pub_StrUserSt03 = "P12" And Left(GetValue(i, "本所案號"), 2) = "P-" And GetValue(i, "申請國家") = 台灣國家代號 _
                      And Val(Replace(GetValue(i, "數量"), "-", "")) > 0 Then strCP156 = ",CP156=null"
                  If Pub_StrUserSt03 = "P12" And Val(GetValue(i, "數量")) > 0 Then strCP156 = ",CP156=null"
                  'end 2016/6/22
                  strSql = "Update Caseprogress Set CP86=NULL" & strCP156 & " Where CP09='" & GetValue(i, "收文號") & "'"
                  cnnConnection.Execute strSql
                  .TextMatrix(i, Val(GetValue(0, "註記"))) = ""
                  .TextMatrix(i, 0) = ""
                  If strCP156 <> MsgText(601) Then .TextMatrix(i, Val(GetValue(0, "數量"))) = ""
                  'end 2014/11/19
               End If
            End If
            'end 2010/6/23
            
            'Add by Morgan 2010/6/23
            If Index = 0 Then
               '取消勾選
               .TextMatrix(i, 0) = ""
               '有承辦人
               If GetValue(i, "承辦人") <> "" Then '.TextMatrix(i, 5)
                  '大陸新案未建國內案
                  If InStr(CaseMapIn, GetValue(i, "CP10")) > 0 Then '.TextMatrix(i, 19)
                     strExc(0) = "select 1 from caseprogress where cp09='" & GetValue(i, "收文號") & "'" & _
                        " and exists(select * from patent where pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and pa09='020')" & _
                        " and not exists(select * from casemap where cm01=cp01 and cm02=cp02 and cm03=cp03 and cm04=cp04 and cm10='0')"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        .TextMatrix(i, 0) = "v"
                        bolGoNext = True
                     End If
                  End If
               End If
            End If
            'end 2010/6/23
         End If
      Next
      'end 2014/11/19
      If bolGoNext Then ComSure_Click 'Add by Morgan 2010/6/23
      
   End With
   
End Sub

Private Sub cmdSearch_Click()
 Dim i As Integer
   Screen.MousePointer = vbHourglass
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = "v"
      Next
   End With
   Screen.MousePointer = vbDefault
   'Modify By Cheng 2002/04/23
'   ComSure.SetFocus
   If Me.Visible Then ComSure.SetFocus
End Sub

Private Sub ComAllData_Click()
   m_QueryType = 1
   Screen.MousePointer = vbHourglass
   If CheckChoese(2) Then
      PutDataInGrid
   End If
   GridHead
   Screen.MousePointer = vbDefault
   intClick = True
End Sub

Private Sub ComBack_Click()
   blnIsFormBack = False
   Unload Me
End Sub

Private Sub ComSure_Click()
   Dim i As Integer, j As Integer
   Dim nPos As Integer
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim bFind As Boolean
   'Add By Cheng 2001/12/25
   Dim ii As Integer '回圈流水號
 
   ' 90.07.05 modify by louis
   If MSHFlexGrid1.Rows < 2 Then
      strTit = "檢核資料"
      strMsg = "請先選取資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         .Visible = False
         If .TextMatrix(i, 0) = "v" Then
            'Add by Amy 2014/12/04
            'Modify by Amy 2015/01/22 +判斷收文日大於P台灣案電子化啟用日
            'Modify by Morgan 2016/6/22 非臺灣案電子化
            'If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And P台灣案電子化啟用日 <= Val(DBDATE(GetValue(i, "收文日"))) Then
            '    If Pub_StrUserSt03 = "P12" And Left(GetValue(i, "本所案號"), 2) = "P-" And GetValue(i, "申請國家") = 台灣國家代號 And GetValue(i, "註記") = MsgText(601) And GetValue(.row, "數量") <> "-" Then
            If Pub_StrUserSt03 = "P12" Then
               If (內專全面電子化啟用日 <= Val(DBDATE(GetValue(.row, "收文日"))) Or (P台灣案電子化啟用日 <= Val(DBDATE(GetValue(.row, "收文日")) And GetValue(.row, "申請國家") = 台灣國家代號))) Then
                  If GetValue(.row, "註記") = MsgText(601) And GetValue(.row, "數量") <> "-" Then
            'end 2016/6/22
                     .Visible = True
                     MsgBox "請輸入接洽單案件性質數量！"
                     .TextMatrix(i, 0) = ""
                     Exit Sub
                  End If 'Added by Morgan 2016/6/22
                End If
            End If
            'end 2014/12/04
            Exit For
         Else
            If i = .Rows - 1 Then
               .Visible = True
               MsgBox "請點選欲分案資料"
               Exit Sub
             End If
         End If
      Next
      .Visible = True
   End With
   frm040101.Tag = ""
   strSave = ""
   
   'Added by Morgan 2021/12/9
   '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
   If PUB_CheckFormExist("frm040101_1") = False Then
      Set frm040101_1 = Nothing
   End If
   'end 2021/12/9
   
   frm040101_1.Show
   'Add by Morgan 2004/4/20
   '若為主管機關來函時，轉本所案號不可輸入
   If Option7.Value = True Then
      frm040101_1.textPA1.Enabled = False
      frm040101_1.textPA2.Enabled = False
      frm040101_1.textPA3.Enabled = False
      frm040101_1.textPA4.Enabled = False
   End If
    
    'Add By Cheng 2003/12/12
   Me.Hide
    'End
   'Add By Cheng 2001/12/25
   DoEvents
   For ii = 0 To Forms.Count - 1
      '專利案件基本資料維護(frm050701)
      If Forms(ii).Name = "frm050701" Then
         frm040101_1.ZOrder 1
         'Add by Amy 2022/12/26 直接開啟接洽單-玲玲
         If PUB_CheckFormExist("frm090801_Q") = True Then
            'Modify by Amy 2023/03/31
            frm050701.intOpen090801 = 1 '不重開
         End If
            'Modify By Cheng 2003/08/18
            'Begin
'         'Add By Cheng 2002/01/03
'         frm050701.SelectToolbarButtom
            'End
         Exit For
      End If
   Next ii
   'Add by Amy 2022/12/23 直接開啟接洽單-玲玲
   If PUB_CheckFormExist("frm090801_Q") = True Then
        frm090801_Q.ZOrder 1
        frm090801_Q.SetFocus
   End If
    'Modify By Cheng 2003/12/12
    '下段程式上移
'   Me.Hide
    'End
EXITSUB:
End Sub

Private Sub ComUCase_Click()
   m_QueryType = 2
   Screen.MousePointer = vbHourglass
   If CheckChoese(1) Then
        PutDataInGrid
    'Add By Cheng 2003/12/12
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'End
   GridHead
   Screen.MousePointer = vbDefault
   intClick = False
   'Add By Cheng 2002/04/23
   '若只搜尋到一筆資料, 則直接進入下一畫面
   'Modify by Morgan 2003/12/23
   'If Me.MSHFlexGrid1.Rows = 2 Then
   If Me.MSHFlexGrid1.Rows = 2 And Me.Visible = True Then
   'Modify end 2003/12/23
      cmdSearch_Click
      ComSure_Click
   End If
End Sub

Private Sub Form_Activate()
 Dim i As Integer
 ' 90.06.28 modify by louis
 Exit Sub
   'True ComAllData ,False ComUCase
   If intClick Then
      ComAllData_Click
   Else
      ComUCase_Click
   End If
   ' 91.01.22 modify by louis (更新狀態時要變更游標型態)
   Screen.MousePointer = vbHourglass
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         'Modify by Amy 2014/11/19 改GetValue 原:.TextMatrix(i, 1)
         If InStr(Me.Tag, GetValue(i, "收文號")) > 0 Then
            .TextMatrix(i, 0) = "v"
         End If
         If InStr(strSave, GetValue(i, "收文號")) > 0 Then
            .TextMatrix(i, 0) = "*"
         End If
      Next
   End With
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtGDate1(0).Text = strSrvDate(2)
   txtGDate1(1) = txtGDate1(0)
   Option1_Click 0
   InitGrid 17, MSHFlexGrid1
   GridHead
   '93.6.27 add by sonia
   Text1 = PUB_GetST06(strUserNum)
  ' Text2 = PUB_GetST06(strUserNum)
   Text2 = Text1
   Text3 = ""
   '93.6.27 END
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
   'add by sonia 2020/8/17 專利國內部人員進入時改預設值(中所管理部及外專不改)
   If Left(Pub_StrUserSt03, 2) = "P1" Then
      txtGDate1(0).Text = TransDate(CompWorkDay(-2, strSrvDate(1), 1), 1)
      Text1 = "1"
      Text2 = "4"
   End If
   'end 2020/8/17
   'Add by Amy 2022/11/15 記錄預設收文所別
   stDefArea1 = Text1
   stDefArea2 = Text2
End Sub

Private Sub GridHead()
 Dim i As Integer
 Dim nowCol As Integer 'Add by Amy 2014/11/19
On Error Resume Next
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
'Modify by Amy 2014/11/19 修改寫法並加數量及申請國家
      nowCol = 1
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
        .ColWidth(nowCol) = 450
      Else
      .ColWidth(nowCol) = 0
      End If
      .col = nowCol: .Text = "數量"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignCenterCenter
      nowCol = nowCol + 1
      
      .col = nowCol: .ColWidth(nowCol) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      nowCol = nowCol + 1
   
      .col = nowCol: .ColWidth(nowCol) = 800: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      nowCol = nowCol + 1
    
      .col = nowCol: .ColWidth(nowCol) = 1100: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      nowCol = nowCol + 1
     
      .col = nowCol: .ColWidth(nowCol) = 1300: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      nowCol = nowCol + 1
   
      .col = nowCol: .ColWidth(nowCol) = 700: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      nowCol = nowCol + 1
     
      .col = nowCol: .ColWidth(nowCol) = 400: .Text = "目次" 'Add By Cheng 2003/06/09
      .CellAlignment = flexAlignCenterCenter
      nowCol = nowCol + 1
  
      .col = nowCol: .ColWidth(nowCol) = 1600: .Text = "案件名稱 "
      
      For i = 9 To 18
         .col = i: .ColWidth(i) = 0
      Next
      nowCol = 19
      .col = nowCol: .ColWidth(nowCol) = 400: .Text = "註記"
      .ColAlignment(.col) = flexAlignCenterCenter
      nowCol = nowCol + 1
      
      'Modify by Amy 2022/10/17
      .col = nowCol: .ColWidth(nowCol) = 0: .Text = "目前表單狀態"
      If strSrvDate(1) >= 接洽單電子收文啟用日 Then
        .ColWidth(nowCol) = 800
      End If
      nowCol = nowCol + 1
      
      '21
      .ColWidth(nowCol) = 0 'Add by Morgan 2010/6/23
      
    'Add by Lydia 2014/11/17
      For i = 22 To 25
      'end 2022/10/17
         .col = i: .ColWidth(i) = 0
      Next
      
      'Add by Amy 2022/11/09 +CP122
      nowCol = i
      .col = nowCol: .ColWidth(nowCol) = 0: .Text = "CP122"
      .ColWidth(nowCol) = 0
      'end 2022/11/09
      
      .Visible = True
'end 2014/11/19
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
'Add By Cheng 2002/07/18
Set frm040101 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   'Modify by Amy 2014/11/19 原:使用共用函數
   'GridClick MSHFlexGrid1, intLastRow, 0, 1
   GridClick1 intLastRow, 0, 1
   ComSure.SetFocus
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
    'Modify by Amy 2014/11/19 原:使用共用函數
    'GridClick MSHFlexGrid1, intLastRow, 0, 1
    GridClick1 intLastRow, 0, 1
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lngX = x
    lngY = y
End Sub

Private Sub Option1_Click(Index As Integer)
 On Error Resume Next
    
   txtGDate1(0).Enabled = False
   txtGDate1(1).Enabled = False
   txtcp01.Enabled = False
   txtcp02.Enabled = False
   txtcp03.Enabled = False
   txtcp04.Enabled = False
   'Add by Amy 2022/11/15 不是選 以前未分案,改回預設所別
   If Index <> 2 Then
     Text1 = stDefArea1
     Text2 = stDefArea2
   End If
   'end 2022/11/15
   Select Case Index
      Case 0
         txtGDate1(0).Enabled = True
         txtGDate1(1).Enabled = True
         txtGDate1(0).SetFocus
      Case 1
         txtcp01.Enabled = True
         txtcp02.Enabled = True
         txtcp03.Enabled = True
         txtcp04.Enabled = True
         txtcp01.SetFocus
      Case 2
         'Add by Amy 2022/11/15 避免資料量太多,造成溢位,選 以前未分案 帶User所別
         Text1 = PUB_GetST06(strUserNum)
         Text2 = Text1
   End Select
End Sub
'93.6.29 ADD BY SONIA
Private Sub Text3_Validate(Cancel As Boolean)
Dim strTempName As String, bolChk As Boolean
   If Text3 <> "" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'bolChk = objPublicData.GetCaseProperty("P", Text3, strTempName, bolChk)
      bolChk = ClsPDGetCaseProperty("P", Text3, strTempName, bolChk)
      Cancel = Not bolChk
      Label4 = strTempName
   Else
      Label4 = ""
   End If
   If Cancel Then TextInverse Text3

End Sub
'93.6.29 END
Private Sub txtcp01_GotFocus()
   TextInverse txtcp01
   CloseIme
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
   If txtcp01 <> "" Then
      txtcp01 = UCase(txtcp01)
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
      If FMP2open = True And FMP2openSQL <> "" Then
         If txtcp01 <> "P" And txtcp01 <> "PS" Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
            Cancel = True
         End If
      ElseIf ChkSysName(txtcp01) = True Then
         If txtcp01 <> "P" And txtcp01 <> "PS" Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtcp01
End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtcp02
End Sub

Private Sub PutDataInGrid()
 Dim i As Integer, strPropertyName As String, strTempName As String, j As Integer
   With MSHFlexGrid1
      .Visible = False
      'Modify by Amy 2014/11/19 TextMatrix改GetValue
      For i = 1 To .Rows - 1
         .row = i
         For j = Val(GetValue(0, "PA23")) To Val(GetValue(0, "CP39"))  '原:For j = 8 + 1 To 12 + 1
            strExc(j - Val(GetValue(0, "PA23"))) = .TextMatrix(i, j) '原:strExc(j - (8 + 1)) = .TextMatrix(i, j)
         Next
         .col = 6 + 1
         '非申請案
         If strExc(0) <> "1" Then
            '新案
            If strExc(1) = "Y" Then
               '對造案件中文名
               If strExc(2) <> "" Then
                  .Text = strExc(2)
               Else
                  '對造案件英文名
                  If strExc(3) <> "" Then
                     .Text = strExc(3)
                  '對造案件日文名
                  ElseIf strExc(4) <> "" Then
                     .Text = strExc(4)
                  End If
               End If
            End If
         End If
         'Add by Lydia 2014/11/17 設FMP寰華案與自動發文一樣為綠色
         'Modify by Amy If (Left(.TextMatrix(i, 3), 2) = "P-" Or Left(.TextMatrix(i, 3), 2) = "PS") And Left(.TextMatrix(i, 21), 1) = "F" And Left(Pub_StrUserSt03, 1) <> "F" Then
         If (Left(GetValue(i, "本所案號"), 2) = "P-" Or Left(GetValue(i, "本所案號"), 2) = "PS") And Left(GetValue(i, "CP12"), 1) = "F" And Left(Pub_StrUserSt03, 1) <> "F" Then
            '取案號
              Dim mDis As Integer
            If Left(GetValue(i, "CaseNo1"), 2) = "PS" Then mDis = 1 '原: .TextMatrix(i, 8)
                strExc(1) = Mid(GetValue(i, "CaseNo1"), 1, 1 + mDis)
                strExc(2) = Mid(GetValue(i, "CaseNo1"), 2 + mDis, 6)
                strExc(3) = Mid(GetValue(i, "CaseNo1"), 8 + mDis, 1)
                strExc(4) = Mid(GetValue(i, "CaseNo1"), 9 + mDis, 2)
            If PUB_FMPtoCheck(1, 2, Pub_strUserST05, strExc(1), strExc(2), strExc(3), strExc(4)) = True Then
               '2014/11/19 FMP寰華案不需輸數量
                .TextMatrix(i, GetValue(0, "數量")) = "-"
               For j = 0 To .Cols - 1
                  .col = j
                  .CellBackColor = &HFF7F& '綠色
               Next
            End If
         End If
'         If .TextMatrix(i, 13 + 1) <> "" And .TextMatrix(i, 13 + 1) <= Format(ChangeWStringToWDateString(GetTodayDate), "YYYYMMDD") And .TextMatrix(i, 16 + 1) = "" Then
        '若有本所期限
        'Modif by Amy 2022/11/09 +CP122=Y 顯示紅色
         If GetValue(i, "CP06") <> "" Or GetValue(i, "CP122") = "Y" Then '原:.TextMatrix(i, 13 + 1)
            '若本所期限小於等於系統日或本所期限為假日且未發文
            'Modify by Amy If (.TextMatrix(i, 13 + 1) <= strSrvDate(1) Or WorkDayCheck(.TextMatrix(i, 14)) = True) And .TextMatrix(i, 16 + 1) = "" Then
            If ((GetValue(i, "CP06") <= strSrvDate(1) Or WorkDayCheck(GetValue(i, "CP06")) = True) And GetValue(i, "CP27") = "") Or GetValue(i, "CP122") = "Y" Then
                'Modify by Morgan 2004/9/13
                'For j = 0 To 16 + 1
                For j = 0 To .Cols - 1
                   .col = j
                   .CellBackColor = &H8080FF '紅色
                Next
                
            'Add by Morgan 2011/1/18 自動收文
            'Mark by Lydia 2023/01/16 因為內專已全面電子化收文，所以只有FMP寰華案變綠色
            'ElseIf GetValue(i, "註記") = "N" Then '原:.TextMatrix(i, 18)
            '   For j = 0 To .Cols - 1
            '      .col = j
            '      .CellBackColor = &HFF7F& '綠色
            '   Next
            'end 2023/01/16
            End If
            '若已閉卷
         ElseIf GetValue(i, "閉卷") = "Y" Then '原:.TextMatrix(i, 14 + 1)
            'Modify by Morgan 2004/9/13
            'For j = 0 To 16 + 1
            For j = 0 To .Cols - 1
               .col = j
               .CellBackColor = &HFFFF& '黃色
            Next
            '若已取消收文
         ElseIf GetValue(i, "CP57") <> "" Then '原:.TextMatrix(i, 15 + 1)
            'Modify by Morgan 2004/9/13
            'For j = 0 To 16 + 1
            For j = 0 To .Cols - 1
               .col = j
               .CellBackColor = &HE0E0E0 '灰色
            Next
         End If
         'Add by Amy 2014/11/19 非A類收文不需輸數量
         If GetValue(i, "收文號") >= "B" Then
             .TextMatrix(i, GetValue(0, "數量")) = "-"
         End If
         'Add by Amy 2022/12/23 需補件者,案件性質顯示 粉紅色
         If GetValue(i, "目前表單狀態") = "程序補件" Then
            .col = GetValue(0, "案件性質")
            .CellBackColor = &HFF80FF     '粉紅色
         End If
      Next i
      .Visible = True
   End With
End Sub

Private Function CheckChoese(ByRef i As Integer) As Boolean
 Dim LcTmp As String
 Dim strField As String 'Add by Amy 2015/01/22
 Dim strWhere As String 'Add by Amy 2022/11/15

   'Add by Amy 2015/01/22 操作人員為北所且北所分案日期有值才顯示承辦人
   'Modify by Amy 2022/11/03 接洽單電子收文上線後直接顯示(cp14=cra09)
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
        strField = " S1.ST02 as 承辦人, "
        If Option1(0).Value = False Then strField = Replace(strField, " S1.ST02", "ST02")
   Else
        If Option1(0).Value Then
             If pub_strUserOffice = "1" Then
                 strField = " Decode(Nvl(cp157,0),0,'',S1.ST02) as 承辦人, "
             Else
                 strField = " S1.ST02 as 承辦人, "
             End If
        Else
             If pub_strUserOffice = "1" Then
                  strField = " Decode(Nvl(cp157,0),0,'',ST02) as 承辦人, "
             Else
                  strField = " ST02 as 承辦人, "
             End If
        End If
   End If
   'Modify By Sindy 2023/6/8 電子收文未分案
   If Option1(3).Value = True Then
      strWhere = "And F0308='A7' And F0309='" & Flow_處理中 & "' And F0301 IS NOT NULL "
   Else
   '2023/6/8 END
      'Add by Amy 2022/11/15 +未分案條件
      If i = 1 Then
         strWhere = "And (F0308= 'A7' Or F0309='" & Flow_已分案 & "' Or F0301 IS NULL) "
      End If
   End If
   
   'Modify by Amy 2015/01/22 北所需顯示北所分案日沒值的資料
   'Modify By Sindy 2023/6/8
   If Option1(3).Value = True Then
      strExc(0) = ""
   Else
   '2023/6/8 END
      If i = 1 And Option1(2).Value = False Then
         If pub_strUserOffice = "1" Then
               strExc(0) = " (CP14 IS NULL or CP157 is null) AND CP10 <> '907' AND CP10<>'913' AND "
           Else
               strExc(0) = " CP14 IS NULL AND CP10 <> '907' AND CP10<>'913' AND "
           End If
      Else
         strExc(0) = " CP10 <> '907' AND CP10<>'913' AND "
         If Option1(2).Value = True Then
            '選 以前未分案
            If pub_strUserOffice = "1" Then
               strExc(0) = strExc(0) & " (CP14 IS NULL or CP157 is null) AND "
            Else
               strExc(0) = strExc(0) & " CP14 IS NULL AND "
            End If
         End If
      End If
      'end2015/01/22
   End If
   
   If Option6.Value Then
      'Modify By Cheng 2002/04/12
'      strExc(0) = strExc(0) & " (substr(cp09,1,1)='A' or substr(cp09,1,1)='B') AND"
      strExc(0) = strExc(0) & " ( cp09<'C' ) AND"
   ElseIf Option7.Value Then
      'Modify By Cheng 2002/04/12
'      strExc(0) = strExc(0) & " substr(cp09,1,1)='C' AND"
      'modify by sonia 2021/2/19 剔除1605通知年費逾期,1913通知期限,990副本信函
      strExc(0) = strExc(0) & " cp09>'C' AND CP10 <> '1605' AND CP10<>'1913' AND CP10<>'990' AND "
   End If
   
   'Add by Morgan 2008/2/12 未註記的
   If i = 3 Then
      strExc(0) = strExc(0) & " CP86 is null AND"
   End If
   
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   If FMP2open = True And FMP2openSQL <> "" Then
      If UCase(Right(strExc(0), 3)) = "AND" Then strExc(0) = Mid(strExc(0), 1, Len(strExc(0)) - 3)
         strExc(0) = strExc(0) & FMP2openSQL & " AND"
   End If
   
   '收文日期
   If Option1(0).Value Then
      If txtGDate1(1) = "" Then MsgBox "請輸入日期 !", vbInformation: Exit Function
      If txtGDate1(0) = "" Then
         strExc(0) = strExc(0) & " cp05<=" & TransDate(txtGDate1(1), 2) & " AND"
      Else
         strExc(0) = strExc(0) & " cp05 between " & TransDate(txtGDate1(0), 2) & _
            " and " + TransDate(txtGDate1(1), 2) & " AND"
      End If
      '93.6.29 add by sonia 加收文所別及案件性質
      If Text1 <> "1" And Text2 <> "1" Then
         strExc(0) = strExc(0) & " S2.ST06 >= '" & Text1 & "' AND S2.ST06 <= '" & Text2 & "' AND "
      Else
         strExc(0) = strExc(0) & " ((S2.ST06 >= '" & Text1 & "' AND S2.ST06 <= '" & Text2 & "') OR S2.ST06='5') AND "
      End If
      If Text3 <> "" Then
         strExc(0) = strExc(0) & " cp10 = '" & Text3 & "' AND "
      End If
      '93.6.29 END
      'Add by Morgan 2004/7/21
      '收文日+未分案排除已取消收文
      If m_QueryType = 2 Then
         strExc(0) = strExc(0) & " cp57 is null AND "
      End If
      
        'Modify By Cheng 2003/06/09
        '加ENG目次
'      strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & "," & _
'         "DECODE(PA09," & 台灣國家代號 & ",CPM03,CPM04),ST02,NVL(PA05,NVL(PA06,PA07)),CP01||CP02||CP03||CP04,PA23," & _
'         "CP31,CP37,CP38,CP39,CP06,PA57,CP57,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF WHERE " & _
'         "CP01 IN ('P','') AND" & strExc(0) & " CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & _
'         "CP01=CPM01(+) and CP10=CPM02(+) AND CP14=ST01(+) UNION " & _
'         "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & "," & _
'         "DECODE(SP09," & 台灣國家代號 & ",CPM03,CPM04),ST02,NVL(SP05,NVL(SP06,SP07)),CP01||CP02||CP03||CP04,1,'','','',''," & _
'         "CP06,SP15,CP57,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF WHERE " & _
'         "CP01 IN ('PS','') AND" & strExc(0) & " CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & _
'         "CP14=ST01(+) AND CP01=CPM01(+) and CP10=CPM02(+)"
      '93.6.29 MODIFY BY SONIA
      'strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & "," & _
      '   "DECODE(PA09," & 台灣國家代號 & ",CPM03,CPM04),ST02, Nvl(EP01,0), NVL(PA05,NVL(PA06,PA07)),CP01||CP02||CP03||CP04,PA23," & _
      '   "CP31,CP37,CP38,CP39,CP06,PA57,CP57,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF, EngineerProgress WHERE " & _
      '   "CP01 IN ('P','') AND" & strExc(0) & " CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & _
      '   "CP01=CPM01(+) and CP10=CPM02(+) AND CP14=ST01(+) And CP09=EP02(+) UNION " & _
      '   "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & "," & _
      '   "DECODE(SP09," & 台灣國家代號 & ",CPM03,CPM04),ST02, Nvl(EP01,0), NVL(SP05,NVL(SP06,SP07)),CP01||CP02||CP03||CP04,1,'','','',''," & _
      '   "CP06,SP15,CP57,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF, EngineerProgress WHERE " & _
      '   "CP01 IN ('PS','') AND" & strExc(0) & " CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & _
      '   "CP14=ST01(+) AND CP01=CPM01(+) and CP10=CPM02(+) And CP09=EP02(+) "
      
      'Modify by Morgan 2004/9/13 加 CP86
      'Add by Lydia 2014/10/31 設別名f0,+FMP2openSQL+,CP44,CP12
      'Modify by Amy 2014/11/19 +每個欄位別名及增加CP156及申請國家
      'Modify by Amy 2015/01/22 操作人員為北所且北所分案日期有值才顯示辦人
      'Modify by Morgan 2016/6/22 非臺灣案電子化
      'strExc(0) = "SELECT '' as V,Decode(pa09||substr(cp09,1,1)||CP140,'000A',''||CP156,'-') as 數量,CP09 as 收文號," & SQLDate("CP05") & " as 收文日," & ChgCaseprogress("", 1) & " as 本所案號,"
      '"UNION SELECT '' as V,Decode(''||CP140,'',''||CP156,'-') as 數量,CP09 as 收文號," & SQLDate("CP05") & " as 收文日," & ChgCaseprogress("", 1) & " as 本所案號,"
      'Modified by Morgan 2016/6/30 FMP案都不必檢查接洽單(不必輸入數量)--玲玲
      'Modified by Morgan 2018/11/5 改FMP案也要檢查接洽單--玲玲
      'Modify by Amy 2022/10/28 +目前表單狀態,cp140
      'Modify by Amy 2022/11/09 +CP122
      'Modify by Amy 2022/11/15 +strWhere
      strExc(0) = "Select V,數量,收文號,收文日,本所案號,案件性質,承辦人,目次,案件名稱,CaseNo1,PA23,CP31,CP37,CP38,CP39,CP06,閉卷,CP57,CP27,註記,Decode(F0309,'" & Flow_處理中 & "',Decode(Decode(Decode(cp157,null,0,1),1,'已分案',F0309),'已分案','已分案'," & ShowFlow表單狀態中文 & ")," & ShowFlow表單狀態中文 & ") as 目前表單狀態,CP10,CP44,CP12,申請國家,CP140,CP122 From Flow003,(" & _
         "SELECT '' as V,decode( CP140||substr(cp09,1,1),'A',''||CP156,'-') as 數量,CP09 as 收文號," & SQLDate("CP05") & " as 收文日," & ChgCaseprogress("", 1) & " as 本所案號," & _
         "DECODE(PA09," & 台灣國家代號 & ",CPM03,CPM04) as 案件性質," & strField & " Nvl(EP01,0) as 目次, NVL(PA05,NVL(PA06,PA07)) as 案件名稱,CP01||CP02||CP03||CP04 as CaseNo1,PA23," & _
         "CP31,CP37,CP38,CP39,CP06,PA57 as 閉卷,CP57,CP27,CP86 as 註記,CP10,CP44,CP12,PA09 as 申請國家,CP140,CP157,CP122 FROM CASEPROGRESS f0,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2, EngineerProgress WHERE " & _
         "CP01 IN ('P','') AND" & strExc(0) & " CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & _
         "CP01=CPM01(+) and CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And CP09=EP02(+) " & _
         "UNION SELECT '' as V,Decode( CP140||substr(cp09,1,1),'A',''||CP156,'-') as 數量,CP09 as 收文號," & SQLDate("CP05") & " as 收文日," & ChgCaseprogress("", 1) & " as 本所案號," & _
         "DECODE(SP09," & 台灣國家代號 & ",CPM03,CPM04) as 案件性質," & strField & " Nvl(EP01,0) as 目次, NVL(SP05,NVL(SP06,SP07)) as 案件名稱,CP01||CP02||CP03||CP04 as CaseNo1,1 as PA23," & _
         "'' as CP31,'' as CP37,'' as CP38,'' as CP39,CP06,SP15,CP57 as 閉卷,CP27,CP86 as 註記,CP10,CP44,CP12,SP09 as 申請國家,CP140,CP157,CP122 FROM CASEPROGRESS f0,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2, EngineerProgress WHERE " & _
         "CP01 IN ('PS','') AND" & strExc(0) & " CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & _
         "CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND CP01=CPM01(+) and CP10=CPM02(+) And CP09=EP02(+) " & _
         ") Where CP140=F0301(+)  " & strWhere & "  order by 收文號"
      '93.6.29 END
      
   '本所案號
   ElseIf Option1(1).Value Then
      If txtcp01 = "" Or txtcp02 = "" Then MsgBox "本所案號錯誤，請重新輸入 !", vbCritical: Exit Function
      If txtcp03.Text = "" Then txtcp03 = "0"
      If txtcp04.Text = "" Then txtcp04.Text = "00"
      LcTmp = ChgCaseprogress(txtcp01 & txtcp02 & txtcp03 & txtcp04)
      If txtcp01 = "P" Then
            'Modify By Cheng 2003/06/09
            '加ENG目次
'         strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & "," & _
'            "DECODE(PA09," & 台灣國家代號 & ",CPM03,CPM04),ST02,NVL(PA05,NVL(PA06,PA07)),CP01||CP02||CP03||CP04,PA23," & _
'            "CP31,CP37,CP38,CP39,CP06,PA57,CP57,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF WHERE " & _
'            "CP01 IN ('P','') AND " & LcTmp & " AND" & strExc(0) & " CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND " & _
'            "CP04=PA04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)"
         
         'Modify by Morgan 2004/9/13 加 CP86
         'Add by Lydia 2014/10/31 設別名f0,+FMP2openSQL+,CP44,CP12
         'Modify by Amy 2014/11/19 +每個欄位別名及增加CP156及申請國家
         'Modify by Amy 2015/01/22 操作人員為北所且北所分案日期有值才顯示辦人
         'Modify by Morgan 2016/6/22 非臺灣案電子化
         'strExc(0) = "SELECT '' as V,Decode(pa09||substr(cp09,1,1)||CP140,'000A',''||CP156,'-') as 數量,CP09 as 收文號," & SQLDate("CP05") & " as 收文日," & ChgCaseprogress("", 1) & " as 本所案號,"
         'Modified by Morgan 2016/6/30 FMP案都不必檢查接洽單(不必輸入數量)--玲玲
         'Modified by Morgan 2018/11/5 改FMP案也要檢查接洽單--玲玲
         'Modify by Amy 2022/10/28 +目前表單狀態,cp140
         'Modify by Amy 2022/11/14 +CP122
         'Modify by Amy 2022/11/15 +strWhere
         strExc(0) = "SELECT '' as V,decode( CP140||substr(cp09,1,1),'A',''||CP156,'-') as 數量,CP09 as 收文號," & SQLDate("CP05") & " as 收文日," & ChgCaseprogress("", 1) & " as 本所案號," & _
            "DECODE(PA09," & 台灣國家代號 & ",CPM03,CPM04) as 案件性質," & strField & " Nvl(EP01,0) as 目次, NVL(PA05,NVL(PA06,PA07)) as 案件名稱,CP01||CP02||CP03||CP04 as CaseNo1,PA23," & _
            "CP31,CP37,CP38,CP39,CP06,PA57 as 閉卷,CP57,CP27,CP86 as 註記,Decode(F0309,'" & Flow_處理中 & "',Decode(Decode(Decode(cp157,null,0,1),1,'已分案',F0309),'已分案','已分案'," & ShowFlow表單狀態中文 & ")," & ShowFlow表單狀態中文 & ") as 目前表單狀態,CP10,CP44,CP12,PA09 as 申請國家,cp140,CP122 " & _
            "FROM CASEPROGRESS f0,PATENT,CASEPROPERTYMAP,STAFF,EngineerProgress,Flow003 WHERE " & _
            "CP01 IN ('P','') AND " & LcTmp & " AND" & strExc(0) & " CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND " & _
            "CP04=PA04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) And CP09=EP02(+) And CP140=F0301(+)  " & strWhere
         '2007/2/17 ADD BY SONIA 加排序條件
         strExc(0) = strExc(0) & " ORDER BY CP09 "
         '2007/2/17 END
      ElseIf txtcp01 = "PS" Then
            'Modify By Cheng 2003/06/09
            '加ENG目次
'         strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & "," & _
'            "DECODE(SP09," & 台灣國家代號 & ",CPM03,CPM04),ST02,NVL(SP05,NVL(SP06,SP07)),CP01||CP02||CP03||CP04,1,'','','',''," & _
'            "CP06,SP15,CP57,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF WHERE " & _
'            "CP01 IN ('PS','') AND " & LcTmp & " AND" & strExc(0) & " CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND " & _
'            "CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)"
         
         'Modify by Morgan 2004/9/13 加 CP86
         'Add by Lydia 2014/10/31 設別名f0,+FMP2openSQL+,CP44,CP12
         'Modify by Amy 2014/11/19 +每個欄位別名及增加CP156及申請國家
         'Modify by Amy 2015/01/22 操作人員為北所且北所分案日期有值才顯示辦人
         'Modify by Morgan 2016/6/22 非臺灣案電子化
         'strExc(0) = "SELECT '' as V,Decode(sp09||substr(cp09,1,1)||CP140,'000A',''||CP156,'-') as 數量,CP09 as 收文號," & SQLDate("CP05") & " as 收文日," & ChgCaseprogress("", 1) & " as 本所案號,"
         'Modified by Morgan 2016/6/30 FMP案都不必檢查接洽單(不必輸入數量)--玲玲
         'Modified by Morgan 2018/11/5 改FMP案也要檢查接洽單--玲玲
         'Modify by Amy 2022/10/28 +目前表單狀態,cp140
         'Modify by Amy 2022/11/14 +CP122
         'Modify by Amy 2022/11/15 +strWhere
         strExc(0) = "SELECT '' as V,decode( CP140||substr(cp09,1,1),'A',''||CP156,'-') as 數量,CP09 as 收文號," & SQLDate("CP05") & " as 收文日," & ChgCaseprogress("", 1) & " as 本所案號," & _
            "DECODE(SP09," & 台灣國家代號 & ",CPM03,CPM04) as 案件性質," & strField & " Nvl(EP01,0) as 目次, NVL(SP05,NVL(SP06,SP07)) as 案件名稱,CP01||CP02||CP03||CP04 as CaseNo1,1 as PA23," & _
            "'' as CP31,'' as CP37,'' as CP38,'' as CP39,CP06,SP15 as 閉卷,CP57,CP27,CP86 as 註記,Decode(F0309,'" & Flow_處理中 & "',Decode(Decode(Decode(cp157,null,0,1),1,'已分案',F0309),'已分案','已分案'," & ShowFlow表單狀態中文 & ")," & ShowFlow表單狀態中文 & ") as 目前表單狀態,CP10,CP44,CP12,SP09 as 申請國家,cp140,CP122 " & _
            "FROM CASEPROGRESS f0,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF, EngineerProgress,Flow003 WHERE " & _
            "CP01 IN ('PS','') AND " & LcTmp & " AND" & strExc(0) & " CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND " & _
            "CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) And CP09=EP02(+) And CP140=F0301(+)  " & strWhere
         '2007/2/17 ADD BY SONIA 加排序條件
         strExc(0) = strExc(0) & " ORDER BY CP09 "
         '2007/2/17 END
      End If
   
   '以前未分案
   'Modify By Sindy 2023/6/8 + Or Option1(3).Value 電子收文未分案
   ElseIf Option1(2).Value Or Option1(3).Value Then
        'Modify By Cheng 2003/06/09
        '加ENG目次
'      strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & "," & _
'         "DECODE(PA09," & 台灣國家代號 & ",CPM03,CPM04),ST02,NVL(PA05,NVL(PA06,PA07)),CP01||CP02||CP03||CP04,PA23," & _
'         "CP31,CP37,CP38,CP39,CP06,PA15,CP57,CP27 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF WHERE " & _
'         "CP14 IS NULL AND CP01 IN ('P','') AND" & strExc(0) & " CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND " & _
'         "PA04=CP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) UNION " & _
'         "SELECT '',CP09," & SQLDate("CP05") & "," & ChgCaseprogress("", 1) & _
'         ",DECODE(SP09," & 台灣國家代號 & ",CPM03,CPM04),ST02,NVL(SP05,NVL(SP06,SP07)),CP01||CP02||CP03||CP04,1,'','','',''," & _
'         "CP06,SP15,CP57,CP27 FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF WHERE " & _
'         "CP14 IS NULL AND CP01 IN ('PS','') AND" & strExc(0) & " CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND " & _
'         "CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+)"
      
      'Modify by Morgan 2004/9/13 加 CP86
      'Add by Lydia 2014/10/31 設別名f0,+FMP2openSQL+,CP44,CP12
      'Modify by Amy 2014/11/19 +每個欄位別名及增加CP156及申請國家
      'Modify by Amy 2015/01/22 操作人員為北所且北所分案日期有值才顯示辦人
      'Modify by Morgan 2016/6/22 非臺灣案電子化
      'strExc(0) = "SELECT '' as V,Decode(pa09||substr(cp09,1,1)||CP140,'000A',''||CP156,'-') as 數量,CP09 as 收文號," & SQLDate("CP05") & " as 收文日," & ChgCaseprogress("", 1) & " as 本所案號,"
      '"UNION SELECT '' as V,Decode(''||CP140,'',''||CP156,'-') as 數量,CP09 as 收文號," & SQLDate("CP05") & " as 收文日," & ChgCaseprogress("", 1) & " as 本所案號,"
      'Modified by Morgan 2016/6/30 FMP案都不必檢查接洽單(不必輸入數量)--玲玲
      'Modified by Morgan 2018/11/5 改FMP案也要檢查接洽單--玲玲
      'Modify by Amy 2022/10/28 +目前表單狀態,cp140
      'Modify by Amy 2022/11/09 +CP122
      'Modify by Amy 2022/11/15 +strWhere
      strExc(0) = "Select V,數量,收文號,收文日,本所案號,案件性質,承辦人,目次,案件名稱,CaseNo1,PA23,CP31,CP37,CP38,CP39,CP06,閉卷,CP57,CP27,註記,Decode(F0309,'" & Flow_處理中 & "',Decode(Decode(Decode(cp157,null,0,1),1,'已分案',F0309),'已分案','已分案'," & ShowFlow表單狀態中文 & ")," & ShowFlow表單狀態中文 & ") as 目前表單狀態,CP10,CP44,CP12,申請國家,CP140,CP122 From Flow003,(" & _
        "SELECT '' as V,decode( CP140||substr(cp09,1,1),'A',''||CP156,'-') as 數量,CP09 as 收文號," & SQLDate("CP05") & " as 收文日," & ChgCaseprogress("", 1) & " as 本所案號," & _
         "DECODE(PA09," & 台灣國家代號 & ",CPM03,CPM04) as 案件性質," & strField & " Nvl(EP01,0) as 目次, NVL(PA05,NVL(PA06,PA07)) as 案件名稱,CP01||CP02||CP03||CP04 as CaseNo1,PA23," & _
         "CP31,CP37,CP38,CP39,CP06,PA57 as 閉卷,CP57,CP27,CP86 as 註記,CP10,CP44,CP12,PA09 as 申請國家,cp140,CP157,CP122 FROM CASEPROGRESS f0,PATENT,CASEPROPERTYMAP,STAFF, EngineerProgress WHERE " & _
         "CP01 IN ('P','') AND" & strExc(0) & " CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND " & _
         "CP04=PA04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) And CP09=EP02(+) " & _
         "UNION SELECT '' as V,Decode(substr(cp09,1,1)||CP140,'A',''||CP156,'-') as 數量,CP09 as 收文號," & SQLDate("CP05") & " as 收文日," & ChgCaseprogress("", 1) & " as 本所案號," & _
         "DECODE(SP09," & 台灣國家代號 & ",CPM03,CPM04) as 案件性質," & strField & " Nvl(EP01,0) as 目次, NVL(SP05,NVL(SP06,SP07)) as 案件名稱,CP01||CP02||CP03||CP04 as CaseNo1,1 as PA23," & _
         "'' as CP31,'' as CP37,'' as CP38,'' as CP39,CP06,SP15 as 閉卷,CP57,CP27,CP86 as 註記,CP10,CP44,CP12,SP09 as 申請國家,cp140,CP157,CP122 FROM CASEPROGRESS f0,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF, EngineerProgress WHERE " & _
         "CP01 IN ('PS','') AND" & strExc(0) & " CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND " & _
         "CP04=SP04(+) AND CP14=ST01(+) AND CP01=CPM01(+) AND CP10=CPM02(+) And CP09=EP02(+) " & _
         ") Where CP140=F0301(+)  " & strWhere
   End If
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   If intI = 1 Then
      CheckChoese = True
      '93.9.29 add by sonia
      Label5 = "共 " & MSHFlexGrid1.Rows - 1 & " 筆"
      '93.9.29 end
   Else
      CheckChoese = False
   End If
   ' 91.04.04 modify by louis (游標)
   'Screen.MousePointer = vbDefault
End Function

Private Sub txtcp03_GotFocus()
   TextInverse txtcp03
End Sub

Private Sub txtcp03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp04_GotFocus()
   TextInverse txtcp04
End Sub

Private Sub txtGDate1_GotFocus(Index As Integer)
   TextInverse txtGDate1(Index)
End Sub

Private Sub txtGDate1_Validate(Index As Integer, Cancel As Boolean)
   If txtGDate1(Index).Text <> "" Then
      If Not ChkDate(txtGDate1(Index)) Then
         Cancel = True
      Else
         If Index = 1 And txtGDate1(0) <> "" And txtGDate1(1) <> "" Then
            If Not ChkRange(txtGDate1(0), txtGDate1(1), "收文") Then Cancel = True
         End If
      End If
      If Cancel Then TextInverse txtGDate1(Index)
   End If
End Sub
'93.6.27 ADD BY SONIA
' 收文所別(起)
Private Sub Text1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(Text1) = True Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "收文所別(起)不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      If Text1 < "1" Or Text1 > "4" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文所別(起)只可為 '1'~'4' "
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text1_GotFocus
      End If
   End If
End Sub
' 收文所別(止)
Private Sub Text2_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(Text2) = True Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "收文所別(止)不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      If Text2 < "1" Or Text2 > "4" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文所別(止)只可為 '1'~'4' "
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text2_GotFocus
      Else
         If Text2 < Text1 Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "收文所別範圍錯誤 "
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Text1_GotFocus
         End If
      End If
   End If
End Sub
'93.6.27 END

' 90.07.06 modify by louis (回到該畫面以原有條件再重新查詢一次)
Public Sub RefreshData()
   Select Case m_QueryType
      Case 1:
         ComAllData_Click
      Case 2:
         ComUCase_Click
      'Add by Morgan 2008/2/12
      Case 3:
         cmdNoneReceive_Click
      Case Else:
   End Select
End Sub

'Add By Cheng 2003/07/22
'檢查本所期限是否為假日期限
Private Function WorkDayCheck(strDate As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

WorkDayCheck = False
If strDate = "" Then Exit Function
StrSQLa = "Select * From Workday Where WD01>" & strSrvDate(1) & " Order By 1 "
rsA.CursorLocation = adUseClient
'Add by Morgan 2003/12/31
rsA.MaxRecords = 1

rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    If Val(strDate) >= Val(strSrvDate(1)) And Val(strDate) < Val("" & rsA.Fields(0).Value) Then
        WorkDayCheck = True
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function
'93.6.27 ADD BY SONIA
Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text3_GotFocus()
   InverseTextBox Text3
End Sub
'93.6.27 END

'Add by Amy 2014/11/19
Public Function GetValue(pRow As Integer, pFieldName As String) As String
   Dim ii As Integer
   With Me.MSHFlexGrid1
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pFieldName) Then
         If pRow = 0 Then
            '回傳第幾欄
            GetValue = ii
         Else
            '回傳欄位內容
            GetValue = .TextMatrix(pRow, ii)
         End If
         Exit For
      End If
   Next
   End With
End Function

'Copy 共用函數GridClick修改
'加入當程序(P12)勾選台灣A類收文非電子送件且註記為空需輸入接洽單案件性質數量
Private Sub GridClick1(intRow As Integer, intCheck As Integer, Optional ByVal iSitu As Integer = 0)
Dim i As Integer, j As Integer
Dim strHeight As String

   With MSHFlexGrid1
      If .row = 0 Then Exit Sub
      
      If iSitu = 0 Then
         If intRow = .row Then
            .col = 0
            If .CellBackColor = &HFFC0C0 Then
               For i = 0 To .Cols - 1
                  .col = i
                  .CellBackColor = .BackColor
               Next
               If .Cols >= intCheck Then .col = intCheck: .Text = ""
            Else
               For i = 0 To .Cols - 1
                  .col = i
                  .CellBackColor = &HFFC0C0
               Next
               If .Cols >= intCheck Then .col = intCheck: .Text = "v"
               'Modify by Amy 2015/01/22 +判斷P台灣案電子化啟用日大於收文日
               'Modify by Morgan 2016/6/22 非臺灣案電子化
               'If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And P台灣案電子化啟用日 <= Val(DBDATE(GetValue(.row, "收文日"))) Then
               ' If Pub_StrUserSt03 = "P12" And Left(GetValue(.row, "本所案號"), 2) = "P-" And GetValue(.row, "申請國家") = 台灣國家代號 And GetValue(.row, "註記") = MsgText(601) And GetValue(.row, "數量") <> "-" Then
               If Pub_StrUserSt03 = "P12" Then
                  If (內專全面電子化啟用日 <= Val(DBDATE(GetValue(.row, "收文日"))) Or (P台灣案電子化啟用日 <= Val(DBDATE(GetValue(.row, "收文日")) And GetValue(.row, "申請國家") = 台灣國家代號))) Then
                     If GetValue(.row, "註記") = MsgText(601) And GetValue(.row, "數量") <> "-" Then
               'end 2016/6/22
                        'Add by Amy 2014/12/04 +彈出表單位置控制
                         frm040101_3.Label3.Caption = GetValue(.row, "本所案號") & " (" & GetValue(.row, "收文號") & ")"
                         strHeight = mdiMain.Top + Me.Top + .Top + lngY + (mdiMain.Height - mdiMain.ScaleHeight) + (Me.Height - Me.ScaleHeight)
                         If Val(strHeight) + frm040101_3.Height > Val(Val(mdiMain.Top + mdiMain.Height)) Then
                             strHeight = Val(strHeight) - frm040101_3.Height - Val(MSHFlexGrid1.RowHeight(1))
                         End If
                         frm040101_3.Move mdiMain.Left + Me.Left + .Left + lngX, Val(strHeight)
                         frm040101_3.Show vbModal
                         frm040101_3.Move mdiMain.Left + Me.Left + .Left + lngX, mdiMain.Top + Me.Top + .Top + lngY + (mdiMain.Height - mdiMain.ScaleHeight) + (Me.Height - Me.ScaleHeight)
                         frm040101_3.Label3.Caption = GetValue(.row, "本所案號") & " (" & GetValue(.row, "收文號") & ")"
                         frm040101_3.Show vbModal
                         intOrderQty = Val(strPublicTemp)
                         strPublicTemp = ""
                         If intOrderQty = 0 Then
                             .Text = ""
                             For i = 0 To .Cols - 1
                                 .col = i
                                 .CellBackColor = .BackColor
                             Next
                         Else
                             .TextMatrix(.row, GetValue(0, "數量")) = intOrderQty
                         End If
                      End If 'Added by Morgan 2016/6/22
                  End If
               End If
            End If
         Else
            intRow = .row
            .Visible = False
            For i = 1 To .Rows - 1
               .row = i
               .col = 0
               If .CellBackColor <> .BackColor Then
                  For j = 0 To .Cols - 1
                     .col = j
                     If j = intCheck Then .Text = ""
                     .CellBackColor = .BackColor
                  Next
               End If
            Next
            .col = 0
            .row = intRow
            For i = 0 To .Cols - 1
               .col = i
               .CellBackColor = &HFFC0C0
            Next
            If .Cols >= intCheck Then .col = intCheck: .Text = "v"
             'Modify by Amy 2015/01/22 +判斷P台灣案電子化啟用日大於收文日
             'Modify by Morgan 2016/6/22 非臺灣案電子化
             'If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And P台灣案電子化啟用日 <= Val(DBDATE(GetValue(.row, "收文日"))) Then
             '  If Pub_StrUserSt03 = "P12" And Left(GetValue(.row, "本所案號"), 2) = "P-" And GetValue(.row, "申請國家") = 台灣國家代號 And GetValue(.row, "註記") = MsgText(601) And GetValue(.row, "數量") <> "-" Then
               If Pub_StrUserSt03 = "P12" Then
                  If (內專全面電子化啟用日 <= Val(DBDATE(GetValue(.row, "收文日"))) Or (P台灣案電子化啟用日 <= Val(DBDATE(GetValue(.row, "收文日")) And GetValue(.row, "申請國家") = 台灣國家代號))) Then
                     If GetValue(.row, "註記") = MsgText(601) And GetValue(.row, "數量") <> "-" Then
             'end 2016/6/22
                       'Add by Amy 2014/12/04 +彈出表單位置控制
                       frm040101_3.Label3.Caption = GetValue(.row, "本所案號") & " (" & GetValue(.row, "收文號") & ")"
                        strHeight = mdiMain.Top + Me.Top + .Top + lngY + (mdiMain.Height - mdiMain.ScaleHeight) + (Me.Height - Me.ScaleHeight)
                        If Val(strHeight) + frm040101_3.Height > Val(Val(mdiMain.Top + mdiMain.Height)) Then
                            strHeight = Val(strHeight) - frm040101_3.Height - Val(MSHFlexGrid1.RowHeight(1))
                        End If
                        frm040101_3.Move mdiMain.Left + Me.Left + .Left + lngX, Val(strHeight)
                        frm040101_3.Show vbModal
                      intOrderQty = Val(strPublicTemp)
                     strPublicTemp = ""
                     If intOrderQty = 0 Then
                         .Text = ""
                        For i = 0 To .Cols - 1
                            .col = i
                            .CellBackColor = .BackColor
                        Next
                     Else
                        .TextMatrix(.row, GetValue(0, "數量")) = intOrderQty
                     End If
                  End If 'Added by Morgan 2016/6/22
               End If
             End If
            
            .Visible = True
         End If
      Else
         If .Cols >= intCheck Then
            .col = intCheck
            If .Text = "" Then
               .Text = "v"
               'Modify by Amy 2015/01/22 +判斷P台灣案電子化啟用日大於收文日
               'Modify by Morgan 2016/6/22 非臺灣案電子化
               'If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And P台灣案電子化啟用日 <= Val(DBDATE(GetValue(.row, "收文日"))) Then
               ' If Pub_StrUserSt03 = "P12" And Left(GetValue(.row, "本所案號"), 2) = "P-" And GetValue(.row, "申請國家") = 台灣國家代號 And GetValue(.row, "註記") = MsgText(601) And GetValue(.row, "數量") <> "-" Then
               If Pub_StrUserSt03 = "P12" Then
                  If (內專全面電子化啟用日 <= Val(DBDATE(GetValue(.row, "收文日"))) Or (P台灣案電子化啟用日 <= Val(DBDATE(GetValue(.row, "收文日")) And GetValue(.row, "申請國家") = 台灣國家代號))) Then
                     If GetValue(.row, "註記") = MsgText(601) And GetValue(.row, "數量") <> "-" Then
               'end 2016/6/22
                        'Add by Amy 2014/12/04 +彈出表單位置控制
                        frm040101_3.Label3.Caption = GetValue(.row, "本所案號") & " (" & GetValue(.row, "收文號") & ")"
                        strHeight = mdiMain.Top + Me.Top + .Top + lngY + (mdiMain.Height - mdiMain.ScaleHeight) + (Me.Height - Me.ScaleHeight)
                        If Val(strHeight) + frm040101_3.Height > Val(mdiMain.Top + mdiMain.Height) Then
                            strHeight = Val(strHeight) - frm040101_3.Height - Val(MSHFlexGrid1.RowHeight(1))
                        End If
                        frm040101_3.Move mdiMain.Left + Me.Left + .Left + lngX, Val(strHeight)
                        frm040101_3.Show vbModal
                        intOrderQty = Val(strPublicTemp)
                        strPublicTemp = ""
                        If intOrderQty = 0 Then
                           .Text = ""
                           For i = 0 To .Cols - 1
                               .col = i
                               .CellBackColor = .BackColor
                           Next
                        Else
                           .TextMatrix(.row, GetValue(0, "數量")) = intOrderQty
                        End If
                    End If 'Added by Morgan 2016/6/22
                  End If
               End If
            Else
               .Text = ""
            End If
         End If
      End If
      .Refresh
   End With
End Sub
