VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880002 
   BorderStyle     =   1  '單線固定
   Caption         =   "優先權日及國家輸入"
   ClientHeight    =   5748
   ClientLeft      =   504
   ClientTop       =   1020
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   8952
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   345
      Left            =   4530
      TabIndex        =   28
      Top             =   60
      Visible         =   0   'False
      Width           =   4365
      Begin VB.TextBox txt2 
         Height          =   270
         Index           =   0
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   29
         Top             =   0
         Width           =   405
      End
      Begin VB.TextBox txt2 
         Height          =   270
         Index           =   1
         Left            =   2895
         MaxLength       =   6
         TabIndex        =   30
         Top             =   0
         Width           =   645
      End
      Begin VB.TextBox txt2 
         Height          =   270
         Index           =   2
         Left            =   3645
         MaxLength       =   1
         TabIndex        =   31
         Top             =   0
         Width           =   180
      End
      Begin VB.TextBox txt2 
         Height          =   270
         Index           =   3
         Left            =   3930
         MaxLength       =   2
         TabIndex        =   32
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton cmd2 
         Caption         =   "從右邊本所案號複製過來"
         Height          =   285
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   2385
      End
      Begin VB.Line Line1 
         X1              =   4170
         X2              =   2640
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.TextBox txtPriority 
      Height          =   264
      Index           =   5
      Left            =   4065
      MaxLength       =   395
      TabIndex        =   2
      Top             =   990
      Width           =   3585
   End
   Begin VB.TextBox txtPriority 
      Height          =   264
      Index           =   4
      Left            =   5370
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1290
      Width           =   732
   End
   Begin VB.TextBox txtPriority 
      Height          =   264
      Index           =   3
      Left            =   1605
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1590
      Width           =   465
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   285
      Left            =   1605
      MaxLength       =   12
      TabIndex        =   6
      Top             =   1890
      Width           =   1365
   End
   Begin VB.TextBox txtPriority 
      Height          =   264
      Index           =   2
      Left            =   1605
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1290
      Width           =   2892
   End
   Begin VB.TextBox txtPriority 
      Height          =   264
      Index           =   1
      Left            =   1605
      MaxLength       =   7
      TabIndex        =   1
      Top             =   990
      Width           =   1092
   End
   Begin VB.TextBox txtPriority 
      Height          =   264
      Index           =   0
      Left            =   1605
      MaxLength       =   3
      TabIndex        =   0
      Top             =   690
      Width           =   732
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "清除(&C)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7740
      TabIndex        =   10
      Top             =   2910
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "刪除(&D)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6780
      TabIndex        =   9
      Top             =   2910
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "加入(&A)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5805
      TabIndex        =   8
      Top             =   2910
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   6810
      TabIndex        =   11
      Top             =   495
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7770
      TabIndex        =   12
      Top             =   495
      Width           =   1110
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   2385
      Left            =   60
      TabIndex        =   26
      Top             =   3330
      Width           =   8865
      _ExtentX        =   15642
      _ExtentY        =   4212
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      HighLight       =   0
      AllowUserResizing=   3
      FormatString    =   "代|優先權國家|種類|本所案號|存取碼|優先權日|優先權號|商品類別"
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
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1605
      TabIndex        =   7
      Top             =   2220
      Width           =   4860
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "8572;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      Caption         =   "註：在資料列上雙擊，會將資料帶至上方畫面的欄位中"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   270
      TabIndex        =   27
      Top             =   3090
      Width           =   4395
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "商品類別："
      Height          =   180
      Left            =   3120
      TabIndex        =   25
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label15 
      Caption         =   "注意：當""存取碼""輸入同於""優先權國家代碼""時,代表是以電子交換檢送"
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   4590
      TabIndex        =   24
      Top             =   1590
      Width           =   1995
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "存取碼："
      Height          =   180
      Left            =   4650
      TabIndex        =   23
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label12 
      Caption         =   $"frm880002.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   60
      TabIndex        =   22
      Top             =   15
      Width           =   4395
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   2115
      TabIndex        =   21
      Top             =   1590
      Width           =   2340
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "專利/商標種類："
      Height          =   180
      Left            =   240
      TabIndex        =   20
      Top             =   1620
      Width           =   1305
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   240
      TabIndex        =   19
      Top             =   2250
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   240
      TabIndex        =   18
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "注意：不知優先權號時請輸入空白，若一國　　　主張多次時空白數也要輸入多個！"
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   270
      TabIndex        =   17
      Top             =   2610
      Width           =   3435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "優先權號："
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label lblNationName 
      Height          =   255
      Left            =   2385
      TabIndex        =   15
      Top             =   690
      Width           =   3960
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "優先權國家："
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "優先權日："
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   1020
      Width           =   900
   End
End
Attribute VB_Name = "frm880002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/11/30 改成Form2.0 ;cboCaseName、Grd1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit

Public strPriority1 As String, strPriority2 As String, strPriority3 As String, strPriority4 As String, strPriority6 As String
Public strPD09 As String 'Add by Amy 2014/04/07
Public m_blnAddNew As Boolean
'Add by Morgan 2004/6/9 台灣專利種類
Public m_stPA08 As String
'Add by Morgan 2005/10/24
Public m_stPA09 As String '申請國家
Public m_bolDblCheck As Boolean '是否檢查與原優先權資料一致
Public m_bolAppCheck As Boolean '申請人是否一致檢查
Public m_strCaseNo As String '本所案號
Dim strPriority1old As String, strPriority2old As String, strPriority3old As String, strPriority4old As String, strPriority6old As String '原優先權資料
Public strPD09old As String 'Add by Amy 2014/04/07 原存取碼
Dim strCaseNo(1 To 4) As String
Dim strSPSign As String '分隔符號
Dim bolMsg As Boolean 'Added by Lydia 2016/10/18
Dim dblPrevRow As Double 'Add By Sindy 2017/10/11
Dim m_strSysKind 'Add By Sindy 2019/1/23 系統別為那個系統:T.商標 或 P.專利
Public frmParent As Form 'Add by Amy 2023/01/05 上一層表單


Private Sub FormClear()
   txtPriority(0) = ""
   txtPriority(1) = ""
   txtPriority(2) = ""
   txtPriority(3) = ""
   txtPriority(4) = "" 'Add by Amy 2014/04/07 存取碼
   txtPriority(5) = "" 'Add by Sindy 2017/9/29 商品類別
   txtCaseNo = ""
   cboCaseName.Clear
End Sub

'Add by Morgan 2004/6/9
Private Function CheckRule(ByVal iRule As Integer, Optional ByRef p_strCaseNo As String) As Boolean

   CheckRule = True
   
   Dim stPA09 As String, stPA10 As String, stPA11 As String
   
   stPA09 = Trim(txtPriority(0))
   stPA10 = DBDATE(txtPriority(1))
   stPA11 = Trim(txtPriority(2))
   
   'Add by Morgan 2007/3/9 台灣案的基本檔申請號第一碼沒有存 '0'
   If stPA09 = "000" Then
      'Add by Morgan 2010/8/23
      '改9碼格式後一致
      If Not bolNewAppNoFormat Then
      'end 2010/8/23
         If Left(stPA11, 1) = "0" Then
            stPA11 = Mid(stPA11, 2)
         End If
      End If
   End If
   'end 2007/3/9
   
   Select Case iRule
      '自先申請案申請日起已逾12個月
      Case 1
         stPA10 = Format(DateAdd("M", 12, ChangeWStringToWDateString(stPA10)), "YYYYMMDD")
         If strSrvDate(1) > stPA10 Then CheckRule = False
      
      Case 2, 3, 4, 5
         '先申請案中所記載之發明或創作已經主張優先權
         If iRule = 2 Then
            strSql = "SELECT PA11 FROM PATENT, PRIDATE" & _
            " WHERE PA09='000' AND PA10=" & stPA10 & " AND PA11='" & stPA11 & "'" & _
            " AND PD01=PA01 AND PD02=PA02 AND PD03=PA03 AND PD04=PA04"
            
         '先申請案為分割案或改請案
         ElseIf iRule = 3 Then
            strSql = "SELECT PA11 FROM PATENT" & _
            " WHERE PA09='000' AND PA10=" & stPA10 & " AND PA11='" & stPA11 & "'" & _
            " AND EXISTS (SELECT * FROM CASEPROGRESS WHERE CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04" & _
            " AND CP10 LIKE '3%' AND CP57 IS NULL)"
            
         '先申請案已經審定或處分
         ElseIf iRule = 4 Then
            'Modified by Morgan 2013/1/11 改已核駁或已公告
            'strSql = "SELECT PA11 FROM PATENT" & _
            " WHERE PA09='000' AND PA10=" & stPA10 & " AND PA11='" & stPA11 & "'" & _
            " AND PA16 IS NOT NULL"
            'Modified by Morgan 2025/7/17 若有再審或訴願審查中的除外(有收文且無結果)--玲玲
            'strSql = "SELECT PA11 FROM PATENT" & _
            " WHERE PA09='000' AND PA10=" & stPA10 & " AND PA11='" & stPA11 & "'" & _
            " AND (PA16='2' or pa14>0)"
            strSql = "SELECT PA11 FROM PATENT" & _
            " WHERE PA09='000' AND PA10=" & stPA10 & " AND PA11='" & stPA11 & "'" & _
            " and (pa14>0 or (pa16='2' and not EXISTS(select * from caseprogress where CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 AND CP10 in ('107','501')  and cp57||cp24 is null)))"
            
            'end 2025/7/17
         'Add by Morgan 2005/11/7
         ElseIf iRule = 5 Then
            strSql = "SELECT PA11 FROM PRIDATE, PATENT" & _
            " WHERE PD05=" & stPA10 & " AND PD06='" & stPA11 & "' AND PD07='" & stPA09 & "'" & _
            " AND PA01(+)=PD01 AND PA02(+)=PD02 AND PA03(+)=PD03 AND PA04(+)=PD04" & _
            " AND PA09=PD07 AND PD01||PD02||PD03||PD04<>'" & p_strCaseNo & "'"
         End If
         
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount > 0 Then
            CheckRule = False
         End If
         CheckOC
   End Select
End Function

'Add By Sindy 2019/1/21 從右邊本所案號複製過來
Private Sub cmd2_Click()
Dim arrTmp(3) As String
Dim i As Integer
   
   If txt2(0) <> "" And txt2(1) <> "" Then
      If Len(txt2(1)) < 6 Then txt2(1) = Format(txt2(1), "000000")
      If txt2(2) = "" Then txt2(2) = "0"
      If txt2(3) = "" Then txt2(3) = "00"
   Else
      MsgBox "請輸入欲複製的本所案號！", vbExclamation
      txt2(0).SetFocus
      Exit Sub
   End If
   
   '系統別要同系統:商標只能複製商標,反之,專利亦同
   strExc(0) = "SELECT * FROM systemkind WHERE sk01='" & txt2(0) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not (m_strSysKind = "T" And RsTemp.Fields("sk02") = "2") Or _
         (m_strSysKind = "P" And RsTemp.Fields("sk02") = "1") Then
         MsgBox "系統別不同性質(商標/專利)，不可複製", vbExclamation
         txt2(0).SetFocus
         Exit Sub
      End If
   Else
      MsgBox "無此系統別！", vbExclamation
      txt2(0).SetFocus
      Exit Sub
   End If
   
   '複製資料
   strExc(0) = "SELECT PriDate.*,na03 FROM PriDate,nation" & _
               " WHERE pd01='" & txt2(0) & "' and pd02='" & txt2(1) & "' and pd03='" & txt2(2) & "' and pd04='" & txt2(3) & "'" & _
               " and pd07=na01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GRD1.Clear
      Call GRIDHEAND
      i = 0
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         i = i + 1
         If i > 1 Then GRD1.AddItem "" '新增
         
         strExc(1) = "R"
         If ConfirmAppNo("" & RsTemp.Fields("pd06"), "", "", "" & RsTemp.Fields("pd07"), "" & RsTemp.Fields("pd05"), strExc(1)) Then
         End If
         strExc(1) = convForm(strExc(1), 12)
         arrTmp(3) = Space(6)
         If "" & RsTemp.Fields("pd08") <> "" Then
            arrTmp(3) = GetKindName(IIf(strCaseNo(1) = "", "FCP", strCaseNo(1)), "" & RsTemp.Fields("pd08"), intI)
            arrTmp(3) = Left(arrTmp(3), 3) + Space(6 - GetTextLength(arrTmp(3)))
         End If
         GRD1.TextMatrix(i, 0) = "" & RsTemp.Fields("pd07") '國代碼
         GRD1.TextMatrix(i, 1) = "" & RsTemp.Fields("na03") '優先權國家
         GRD1.TextMatrix(i, 2) = "" & RsTemp.Fields("pd08") & " " & arrTmp(3) '種類
         GRD1.TextMatrix(i, 3) = strExc(1) '本所案號
         GRD1.TextMatrix(i, 4) = "" & RsTemp.Fields("pd09") '存取碼
         GRD1.TextMatrix(i, 5) = "" & RsTemp.Fields("pd05") '優先權日
         GRD1.TextMatrix(i, 6) = "" & RsTemp.Fields("pd06") '優先權號
         GRD1.TextMatrix(i, 7) = "" & RsTemp.Fields("pd10") '商品類別
         
         RsTemp.MoveNext
      Loop
      
      '若有資料游標停在第一筆
      GRD1.Visible = False
      GRD1.col = 0
      GRD1.row = 1
      dblPrevRow = GRD1.row
      If GRD1.Rows - 1 > 0 Then
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
      GRD1.Visible = True
   Else
      MsgBox "查無資料！", vbExclamation
      Exit Sub
   End If
End Sub

Private Sub cmdMove_Click(Index As Integer)
   Dim i As Integer, ii As Integer
   Dim varTemp As Variant
   Dim arrTmp(1 To 4) As String
   Dim bolSelect As Boolean
      
On Error GoTo ErrorHandle 'Added by Lydia 2018/06/27

If Index = 0 Then '加入
   'Added by Morgan 2022/6/20 有輸入優先權號時,要去掉前後的空白及跳行符號
   If Trim(txtPriority(2)) <> "" Then
       txtPriority(2) = Trim(txtPriority(2))
       txtPriority(2) = Replace(txtPriority(2), vbCrLf, "")
   End If
   'end 2022/6/20
   
   bolMsg = True 'Added by Lydia 2016/10/18
   
   For i = 0 To 5 'Modify by Amy 2014/04/07 +存取碼 Sindy 2017/9/29 +商品類別
      If CheckKeyIn(i) <> 1 Then
         If txtPriority(i).Enabled = True Then
            txtPriority(i).SetFocus
            txtPriority_GotFocus i
         End If
         Exit For
      End If
   Next
   bolMsg = False 'Added by Lydia 2016/10/18
   'If i < 5 Then Exit Sub 'Modify by Amy 2014/03/25 +存取碼
   If i < 6 Then Exit Sub 'Modify by Sindy 2017/9/29 +商品類別
   
   'Remove by Morgan 2008/8/29 前面已有控制
   'If txtPriority(0).Text = "" Or txtPriority(1).Text = "" Then Exit Sub
   
   'Add by Morgan 2006/5/15
   'Modify By Sindy 2019/1/23 專利才檢查
   If m_strSysKind = "P" Then
   '2019/1/23 END
      If m_strCaseNo <> "" Then
         'Add By Sindy 2019/1/24
         strExc(1) = Left(m_strCaseNo, Len(m_strCaseNo) - 9) 'Added by Morgan 2019/5/9
         If strExc(1) = "P" Or strExc(1) = "CFP" Then
         '2019/1/24 END
            If CheckCountryMatch(m_strCaseNo, txtPriority(0).Text) = False Then
               Exit Sub
            End If
         End If
         'Add by Morgan 2010/8/23
         '檢查優先權號格式
         If bolNewAppNoFormat Then
            strExc(1) = Left(m_strCaseNo, Len(m_strCaseNo) - 9)
            If strExc(1) = "P" Or strExc(1) = "CFP" Or strExc(1) = "FCP" Then
               If (txtPriority(0) = "000" Or txtPriority(0) = "020") And Val(txtPriority(1)) > 0 And Trim(txtPriority(2)) <> "" And txtPriority(3) <> "" Then
                  If Not ChkAppNo(txtPriority(2), txtPriority(3), IIf(txtPriority(0) = "000", "0", "1")) Then
                     txtPriority(2).SetFocus
                     Exit Sub
                  End If
               End If
            End If
         End If
      End If
      
      'Add by Morgan 2007/4/25
      'Modified by Morgan 2019/7/26 美國案除外--甄妮
      If m_stPA08 = "3" And txtPriority(3) = "1" And m_stPA09 <> "101" Then
         MsgBox "設計不得主張發明優先權！", vbCritical
         Exit Sub
      End If
      'end 2007/4/25
      
      'Add by Morgan 2004/6/7
      '申請國為台灣才會傳專利種類
      'Modify by Morgan 2007/3/5 改判斷申請國家
      'If Val(m_stPA08) > 0 Then
      If m_stPA09 = "000" Then
         If txtPriority(0).Text = "000" Then
            '台灣設計不得主張國內優先權
            If Val(m_stPA08) = 3 Then
               MsgBox "設計不得主張國內優先權！"
               Exit Sub
            '檢查專利法29條第1項4款
            Else
               If CheckRule(1) = False Then
                  MsgBox "自先申請案申請日起已逾12個月，不得主張優先權！", vbCritical
                  Exit Sub
               ElseIf CheckRule(2) = False Then
                  MsgBox "先申請案中所記載之發明或創作已經主張優先權，不得主張優先權！", vbCritical
                  Exit Sub
               ElseIf CheckRule(3) = False Then
                  MsgBox "先申請案為分割案或改請案，不得主張優先權！", vbCritical
                  Exit Sub
               ElseIf CheckRule(4) = False Then
                  'Modified by Morgan 2013/1/11
                  'MsgBox "先申請案已經審定或處分，不得主張優先權！", vbCritical
                  MsgBox "先申請案已核駁審定或處分,或核准公告,不得主張國內優先權！", vbCritical
                  Exit Sub
               End If
            End If
         End If
      End If
   End If
   
   'Add by Morgan 2005/11/7
   '國內優先權(所有國家都要)
   If m_stPA09 = Trim(txtPriority(0)) Then
      '2011/6/1 modify by sonia P-098703郭雅娟說林柄佑說已無此規定
      'If CheckRule(5, m_strCaseNo) = False Then
      'Removed by Morgan 2013/5/31 取消--郭
      'If strExc(1) <> "P" And CheckRule(5, m_strCaseNo) = False Then
      '   MsgBox "先申請案已被主張國內優先權，不得重複主張！", vbCritical
      '   Exit Sub
      'End If
      'end 2013/5/31
      
      'Add by Morgan 2005/11/14
      If m_bolAppCheck = True Then
         If CheckApplicant = False Then
            If MsgBox("先申請案之申請人與本案不同是否要繼續！", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
      End If
   End If
   
   strExc(2) = "" 'Added by Morgan 2021/1/11 要先清否則會殘留變數之前的內容
   'Add by Morgan 2005/3/22 若優先權號有對應的本所案但申請日期或國家不同時要確認！
   If Trim(txtPriority(2)) <> "" Then
      'Modified by Lydia 2016/10/13 抓本所案號
      'If ConfirmAppNo(txtPriority(2), txtPriority(1), txtPriority(0)) = False Then
      strExc(2) = ""
      If ConfirmAppNo(txtPriority(2), txtPriority(1), txtPriority(0), , , strExc(2)) = False Then
         Exit Sub
      End If
   End If
   
   For i = 1 To GRD1.Rows - 1
      If Trim(GRD1.TextMatrix(i, 0)) = "" Then
         Exit For
      Else
         '檢查是否有申請國&優先權號相同
         'Modified by Morgan 2018/7/17
         'If txtPriority(0) = Trim(grd1.TextMatrix(i, 0)) And txtPriority(2) = Trim(grd1.TextMatrix(i, 6)) Then
         If txtPriority(0) = Trim(GRD1.TextMatrix(i, 0)) And txtPriority(2) = GRD1.TextMatrix(i, 6) Then
         'end 2018/7/17
            'Modify By Sindy 2019/1/18 Mark, 重覆代表修改
'            ShowMsg MsgText(9200)
'            Exit Sub
            Exit For
            '2019/1/18 END
         End If
      End If
   Next
   '沒有重複時新增
   If i > GRD1.Rows - 1 Then
      GRD1.AddItem "" '新增, 反之就是修改
   End If
   GRD1.TextMatrix(i, 0) = CStr(txtPriority(0)) '國代碼
   GRD1.TextMatrix(i, 1) = lblNationName.Caption '優先權國家
   GRD1.TextMatrix(i, 2) = " " & txtPriority(3) & " " & Label7 '種類
   GRD1.TextMatrix(i, 3) = convForm(strExc(2), 12) '本所案號
   GRD1.TextMatrix(i, 4) = " " & txtPriority(4) '存取碼
   GRD1.TextMatrix(i, 5) = " " & txtPriority(1) '優先權日
   'Modified by Morgan 2018/7/17 不可加空白否則會與尚無優先權號的空白混淆
   'grd1.TextMatrix(i, 6) = " " & txtPriority(2) '優先權號
   GRD1.TextMatrix(i, 6) = txtPriority(2)   '優先權號
   'end 2018/7/17
   GRD1.TextMatrix(i, 7) = " " & txtPriority(5) '商品類別
   FormClear
   
ElseIf Index = 1 Then '刪除
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If GRD1.CellBackColor = &HFFC0C0 Then
         bolSelect = True
         Exit For
      End If
   Next i
   If bolSelect = False Then
      ShowMsg MsgText(8006)
   Else
      dblPrevRow = i
      'Added by Morgan 2018/4/11
      'grd1.RemoveItem i
      If GRD1.Rows = 2 Then
         For ii = 0 To GRD1.Cols - 1
            GRD1.TextMatrix(i, ii) = ""
         Next
      Else
         GRD1.RemoveItem i
      End If
      'end 2018/4/11
      If GRD1.Rows - 1 <> 0 Then
         If dblPrevRow = GRD1.Rows Then
            dblPrevRow = GRD1.Rows - 1
         End If
      End If
   End If
   
Else '清除,是清除全部資料
   For i = GRD1.Rows - 1 To 1 Step -1
      GRD1.RemoveItem i
   Next i
   'Add By Sindy 2019/1/18
   '第一筆無法用 RemoveItem 刪除, 因此須再呼叫刪除鍵, 清除grd1欄位值
   Call cmdMove_Click(1)
   '2019/1/18 END
End If

txtPriority(0).SetFocus

'Added by Lydia 2018/06/27
Exit Sub

ErrorHandle:
     Resume Next
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim i As Integer, varTemp As Variant, strTmp As String

   If Index = 0 Then '確定
      strPriority1 = ""
      strPriority2 = ""
      strPriority3 = ""
      strPriority4 = ""
      strPD09 = "" 'Add by Amy 2014/04/07
      strPriority6 = "" 'Add by Sindy 2017/10/11
      If GRD1.Rows - 1 > 0 Then
         For i = 1 To GRD1.Rows - 2
            '國家代碼
            strPriority1 = strPriority1 + Trim(GRD1.TextMatrix(i, 0)) + "，"
            '優先權日:統一用西元年 8 碼
            'Modified by Lydia 2016/10/13 temp(5) => temp(6)
            strPriority2 = strPriority2 + Trim(GRD1.TextMatrix(i, 5)) + "，"
            '若無優先權號, 則以空白替代
            'Modified by Lydia 2016/10/13 temp(6) => temp(7)
            'strPriority3 = strPriority3 + IIf(varTemp(7) = "", " ", varTemp(7)) + "，"
            'Modified by Lydia 2018/07/13 無優先權號, 則以空白替代
            'strPriority3 = strPriority3 + Trim(Grd1.TextMatrix(i, 6)) + "，"
            
            strPriority3 = strPriority3 + GRD1.TextMatrix(i, 6) + "，"
            'Add by Morgan 2007/4/24
            '專利/商標種類
            If Trim(GRD1.TextMatrix(i, 2)) = "" Then
               strPriority4 = strPriority4 + "，"
            Else
               strPriority4 = strPriority4 + Left(Trim(GRD1.TextMatrix(i, 2)), 1) + "，"
            End If
            'Add by Amy 2014/04/07 存取碼
            'Modified by Lydia 2016/10/13 temp(4) => temp(5)
            strPD09 = strPD09 + Trim(GRD1.TextMatrix(i, 4)) + "，"
            'Add by Sindy 2017/10/11 商品類別
            strPriority6 = strPriority6 + Trim(GRD1.TextMatrix(i, 7)) + "，"
         Next
         '國家代碼
         strPriority1 = strPriority1 + Trim(GRD1.TextMatrix(i, 0))
         '優先權日:統一用西元年 8 碼
         'Modified by Lydia 2016/10/13 temp(5) => temp(6)
         strPriority2 = strPriority2 + Trim(GRD1.TextMatrix(i, 5))
         '優先權號
         'Modified by Lydia 2018/07/13 無優先權號, 則以空白替代
         'strPriority3 = strPriority3 + Trim(Grd1.TextMatrix(i, 6))
         strPriority3 = strPriority3 + GRD1.TextMatrix(i, 6)
         'Add by Morgan 2007/4/24
         '專利/商標種類
         If Trim(GRD1.TextMatrix(i, 2)) <> "" Then
            strPriority4 = strPriority4 + Left(Trim(GRD1.TextMatrix(i, 2)), 1)
         End If
         'Add by Amy 2014/04/07 存取碼
         'Modified by Lydia 2016/10/13 temp(4) => temp(5)
         strPD09 = strPD09 + Trim(GRD1.TextMatrix(i, 4))
         'Add by Sindy 2017/10/11 商品類別
         strPriority6 = strPriority6 + Trim(GRD1.TextMatrix(i, 7))
      End If
      'Add by Morgan 2005/10/24 檢查一致性
      If m_bolDblCheck = True And Not (strPriority1 = "" And strPriority1old = "") Then
         If CheckIdentical = False Then
            If MsgBox("本次輸入優先權資料與前次資料不一致，是否要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
               Exit Sub
            End If
         End If
      End If
   Else
      'Add by Morgan 2005/11/7 控制一致性檢查
      m_bolDblCheck = False
   End If
   Unload Me
End Sub

Private Sub ChkColText()
   'Add by Morgan 2007/4/23 拆本所案號
   If m_strCaseNo <> "" Then
      'Modify By Sindy 2017/11/9 已開放不需管控
      'txtPriority(3).Enabled = True
      ChgCaseNo m_strCaseNo, strCaseNo
   'Modify By Sindy 2017/11/9 FCP也要輸種類,因此開放此欄位均可輸入
'   Else
'      txtPriority(3).Enabled = False
   End If
   
   'Add By Sindy 2019/1/23
   m_strSysKind = ""
   If strCaseNo(1) <> "" Then
      Frame1.Visible = True
      '系統別為那個系統:商標 或 專利
      strExc(0) = "SELECT * FROM systemkind WHERE sk01='" & strCaseNo(1) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_strSysKind = IIf(RsTemp.Fields("sk02") = "2", "T", "P")
      End If
   Else
      Frame1.Visible = False
   End If
   '2019/1/23 END
   'Add By Sindy 2024/4/17 請判斷若為專利案件要將商品類別欄鎖住；而專利案件才開放存取碼欄可輸入。
   txtPriority(5).Enabled = False
   txtPriority(4).Enabled = False
   If m_strSysKind = "T" Then
      txtPriority(5).Enabled = True
   ElseIf m_strSysKind = "P" Then
      txtPriority(4).Enabled = True
   Else
      txtPriority(5).Enabled = True
      txtPriority(4).Enabled = True
   End If
   '2024/4/17 END
End Sub

Private Sub Form_Load()
   'Modify by Amy 2014/03/25 +varPriorityTemp5
   'Modify by Sindy 2017/10/6 +varPriorityTemp6
   Dim i As Integer, varPriorityTemp1, varPriorityTemp2, varPriorityTemp3, varPriorityTemp4, varPriorityTemp5, varPriorityTemp6
   Dim strTemp As String
   Dim arrTmp(1 To 6) As String 'Modify by Amy 2014/04/07
   Dim intRow As Integer 'Add By Sindy 2017/10/6
   
   MoveFormToCenter Me
   
   strSPSign = " " & vbVerticalTab
   '統一用西元年 8 碼
   txtPriority(1).MaxLength = 8
   
   'Modify By Sindy 2024/4/18
   Call ChkColText
'   'Add By Sindy 2019/1/23
'   m_strSysKind = ""
'   If strCaseNo(1) <> "" Then
'      Frame1.Visible = True
'      '系統別為那個系統:商標 或 專利
'      strExc(0) = "SELECT * FROM systemkind WHERE sk01='" & strCaseNo(1) & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         m_strSysKind = IIf(RsTemp.Fields("sk02") = "2", "T", "P")
'      End If
'   Else
'      Frame1.Visible = False
'   End If
'   '2019/1/23 END
   '2024/4/18 END
   
   Call GRIDHEAND 'Add By Sindy 2017/10/6
   'Modify by Morgan 2005/10/24 若要檢查與原優先權資料一致時要重輸
   'If strPriority1 <> "" Then
   If m_bolDblCheck = True Then
      strPriority1old = strPriority1
      strPriority2old = strPriority2
      strPriority3old = strPriority3
      strPriority4old = strPriority4 'Add by Morgan 2007/4/23
      strPD09old = strPD09 'Add by Amy 2014/04/07
      strPriority6old = strPriority6 'Add by Sindy 2017/9/29
   ElseIf strPriority1 <> "" Then
      varPriorityTemp1 = Split(strPriority1, "，")
      'Add by Morgan 2008/8/29 優先權日可空白故後面補一空白以免無陣列可讀
      If strPriority2 = "" Then strPriority2 = " "
      
      varPriorityTemp2 = Split(strPriority2, "，")
      varPriorityTemp3 = Split(strPriority3, "，")
      varPriorityTemp4 = Split(IIf(strPriority4 = "", " ", strPriority4), "，")
      varPriorityTemp5 = Split(IIf(strPD09 = "", " ", strPD09), "，") 'Add by Amy 2014/04/07
      varPriorityTemp6 = Split(IIf(strPriority6 = "", " ", strPriority6), "，") 'Add by Sindy 2017/10/6
      For i = 0 To UBound(varPriorityTemp1)
         If ClsPDGetNation(CStr(varPriorityTemp1(i)), strTemp) Then
            intRow = intRow + 1
            'Modified by Morgan 2018/4/11
            'grd1.AddItem ""
            If intRow > GRD1.Rows - 1 Then GRD1.AddItem ""
            'end 2018/4/11
            GRD1.TextMatrix(intRow, 0) = CStr(varPriorityTemp1(i)) '國代碼
            GRD1.TextMatrix(intRow, 1) = strTemp '優先權國家
            
            '統一用西元年 8 碼
            'Modify by Morgan 2007/4/23 加專利/商標種類
            '申請國家
'            'Modify by Amy 2014/04/11 優先權國家原顯示8個字改6個字
'            arrTmp(1) = Left(strTemp, 6)
'            arrTmp(1) = arrTmp(1) + Space(12 - GetTextLength(arrTmp(1)))   '原:Space(16 - GetTextLength(ArrTmp(1)))
'            'end 2014/04/11
            arrTmp(2) = Space(1)
            arrTmp(3) = Space(6)
            If strPriority4 <> "" Then
               arrTmp(2) = Right(Space(1) + varPriorityTemp4(i), 1)
'               If strCaseNo(1) <> "" Then
                  arrTmp(3) = GetKindName(IIf(strCaseNo(1) = "", "FCP", strCaseNo(1)), arrTmp(2), intI)
                  arrTmp(3) = Left(arrTmp(3), 3) + Space(6 - GetTextLength(arrTmp(3)))
'               End If
            End If
            '應該不用判斷,原來有先保留
            arrTmp(4) = Empty
            If strPriority3 <> "" Then
               arrTmp(4) = varPriorityTemp3(i)
            End If
            'Add by Amy 2014/03/25 +存取碼
            arrTmp(5) = Space(6)
            If strPD09 <> "" Then
               arrTmp(5) = varPriorityTemp5(i)
               arrTmp(5) = arrTmp(5) + Space(6 - GetTextLength(arrTmp(5)))
            End If
            
            'Modify by Amy 2014/04/07 +存取碼(增加欄位最好加於優先權日之前,因優先權號為key且不確定長度 null時Insert 空白)
            'lstPriority.AddItem varPriorityTemp1(i) + strSPSign + ArrTmp(1) + strSPSign + ArrTmp(2) + strSPSign + ArrTmp(3) + strSPSign & Space(8 - Len(varPriorityTemp2(i))) & varPriorityTemp2(i) + strSPSign + ArrTmp(4)
            'Modified by Lydia 2016/10/13 本所案號加在存取碼前面
            'lstPriority.AddItem varPriorityTemp1(i) + strSPSign + arrTmp(1) + strSPSign + arrTmp(2) + strSPSign + arrTmp(3) + strSPSign + arrTmp(5) + strSPSign & Space(8 - Len(varPriorityTemp2(i))) & varPriorityTemp2(i) + strSPSign + arrTmp(4)
            strExc(1) = "R"
            If ConfirmAppNo(arrTmp(4), "", "", CStr(varPriorityTemp1(i)), varPriorityTemp2(i), strExc(1)) Then
            End If
            strExc(1) = convForm(strExc(1), 12)
'            lstPriority.AddItem varPriorityTemp1(i) + strSPSign + arrTmp(1) + strSPSign + arrTmp(2) + strSPSign + arrTmp(3) + strSPSign + strExc(1) + strSPSign + arrTmp(5) + strSPSign & Space(8 - Len(varPriorityTemp2(i))) & varPriorityTemp2(i) + strSPSign + arrTmp(4)
            'end 2014/04/07
            
            GRD1.TextMatrix(intRow, 2) = " " & varPriorityTemp4(i) & " " & arrTmp(3) '種類
            GRD1.TextMatrix(intRow, 3) = strExc(1) '本所案號
            GRD1.TextMatrix(intRow, 4) = " " & varPriorityTemp5(i) '存取碼
            GRD1.TextMatrix(intRow, 5) = " " & varPriorityTemp2(i) '優先權日
            'Modified by Morgan 2018/7/17 不可加空白否則會與尚無優先權號的空白混淆
            'Grd1.TextMatrix(intRow, 6) = " " & varPriorityTemp3(i) '優先權號
            GRD1.TextMatrix(intRow, 6) = varPriorityTemp3(i)  '優先權號
            'end 2018/7/17
            GRD1.TextMatrix(intRow, 7) = " " & varPriorityTemp6(i) '商品類別
         End If
      Next
      '若有資料游標停在第一筆
      GRD1.Visible = False
      GRD1.col = 0
      GRD1.row = 1
      dblPrevRow = GRD1.row
      If GRD1.Rows - 1 > 0 Then
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
      GRD1.Visible = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/06/05
If Me.GRD1.Rows - 1 > 0 Then
   Me.m_blnAddNew = True
End If

'Add by Amy 2023/01/05 不是強制表單需於此回傳值
If TypeName(frmParent) <> "Nothing" Then
    Select Case UCase(frmParent.Name)
        Case UCase("frm040101_1") 'P分案
            frmParent.strPrity1 = strPriority1
            frmParent.strPrity2 = strPriority2
            frmParent.strPrity3 = strPriority3
            frmParent.strPrity4 = strPriority4
            frmParent.strPrity5 = strPD09
        Case UCase("frm050101_2") 'CFP分案
            frmParent.strPriority1 = strPriority1
            frmParent.strPriority2 = strPriority2
            frmParent.strPriority3 = strPriority3
            frmParent.strPriority4 = strPriority4
            frmParent.strPriority5 = strPD09
        Case UCase("frm020101_02") 'T分案
            frmParent.m_Priority1 = strPriority1
            frmParent.m_Priority2 = strPriority2
            frmParent.m_Priority3 = strPriority3
            frmParent.m_Priority4 = strPriority4
            frmParent.m_Priority5 = strPD09
            frmParent.m_Priority6 = strPriority6
    End Select
    Forms(0).Enabled = True 'Modify by Amy 2023/01/07 Casher 會錯,原:mdiMain.Enabled = True
    frmParent.Enabled = True
    m_bolDblCheck = False
    strPriority1 = "": strPriority2 = "": strPriority3 = "": strPriority4 = "": strPD09 = "": strPriority6 = ""
    Set frm880002 = Nothing
End If
'end 2023/01/05

'Memo 強制表單於ModifyPriority 函數把此表單Set Nothing
'Add By Cheng 2002/07/18
'Set frm880002 = Nothing
End Sub

Private Sub GRD1_DblClick()
   If GRD1.Rows - 1 > 0 Then
      GRD1.col = 0
      GRD1.row = dblPrevRow
      If GRD1.CellBackColor <> &HFFC0C0 Then
         ShowMsg "請點選資料列!"
         Exit Sub
      End If
      txtPriority(0) = Trim(GRD1.TextMatrix(dblPrevRow, 0))
      lblNationName = Trim(GRD1.TextMatrix(dblPrevRow, 1))
      txtPriority(3) = Left(Trim(GRD1.TextMatrix(dblPrevRow, 2)), 1)
      Label7 = Trim(Mid(Trim(GRD1.TextMatrix(dblPrevRow, 2)), 2))
      'Added by Lydia 2016/10/18 +本所案號
      txtCaseNo.Text = Trim(GRD1.TextMatrix(dblPrevRow, 3))
      If txtCaseNo.Text <> "" Then
         GetCaseData txtCaseNo
      Else
         cboCaseName.Clear
      End If
      'end 2016/10/18
      
      'Modify by Amy 2014/04/07 +存取碼
      'Modified by Lydia 2016/10/18 index + 1
      txtPriority(4) = Trim(GRD1.TextMatrix(dblPrevRow, 4))
      txtPriority(1) = Trim(GRD1.TextMatrix(dblPrevRow, 5))
      'Modified by Morgan2018/7/17 優先權號可能是空白
      'txtPriority(2) = Trim(grd1.TextMatrix(dblPrevRow, 6))
      txtPriority(2) = GRD1.TextMatrix(dblPrevRow, 6)
      'end 2018/7/17
      'end 2014/04/07
      txtPriority(5) = Trim(GRD1.TextMatrix(dblPrevRow, 7)) 'Add By Sindy 2017/9/29
   End If
End Sub

Private Sub grd1_SelChange()
Dim i As Integer

GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      GRD1.col = 0
      GRD1.row = dblPrevRow
      'Grd1.Text = ""
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   dblPrevRow = GRD1.row
'   If grd1.Text = "V" Then
'      grd1.Text = ""
'      For i = 0 To grd1.Cols - 1
'         grd1.col = i
'         grd1.CellBackColor = QBColor(15)
'      Next i
'   Else
      If Trim(GRD1.TextMatrix(GRD1.row, 1)) <> "" Then
         'Grd1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
'   End If
End If
GRD1.Visible = True
End Sub

'Add By Sindy 2019/1/21
Private Sub txt2_GotFocus(Index As Integer)
   InverseTextBox txt2(Index)
   CloseIme
End Sub
Private Sub txt2_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 0
      KeyAscii = UpperCase(KeyAscii)
   Case Else
      '開放可以打T
      If Index = 2 Then
         KeyAscii = UpperCase(KeyAscii)
         Select Case KeyAscii
         Case 48 To 57, 8, 65 To 90
         Case Else
            KeyAscii = 0
         End Select
      Else
         Select Case KeyAscii
         Case 48 To 57, 8
         Case Else
            KeyAscii = 0
         End Select
      End If
End Select
End Sub

Private Sub txtCaseNo_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtCaseNo.IMEMode = 2
   CloseIme
End Sub

Private Sub txtCaseNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub GetCaseData(strCaseNo As String)
   Dim stCaseNo(1 To 4) As String
   ChgCaseNo strCaseNo, stCaseNo
   Select Case CheckSys(stCaseNo(1))
      Case "1" '專利
         strSql = "select pa09,pa10,pa11,pa05,pa06,pa07,pa08 from patent where pa01='" & stCaseNo(1) & "' and pa02='" & stCaseNo(2) & "' and pa03='" & stCaseNo(3) & "' and pa04='" & stCaseNo(4) & "'"
      Case "2" '商標
         'Modified by Lydia 2016/10/18 tm12=> nvl(tm12,tm15)
         strSql = "select tm10,tm11,nvl(tm12,tm15),tm05,tm06,tm07,tm08 from trademark where tm01='" & stCaseNo(1) & "' and tm02='" & stCaseNo(2) & "' and tm03='" & stCaseNo(3) & "' and tm04='" & stCaseNo(4) & "'"
   End Select
   
On Error GoTo ErrHnd
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         txtPriority(0) = "" & .Fields(0)
         txtPriority(1) = "" & .Fields(1)
         'Modify by Morgan 2010/12/28 申請案號改碼數台灣案不必再補零
         'If txtPriority(0) = "000" Then
         '   txtPriority(2) = "" & .Fields(2)
         'Else
            txtPriority(2) = "" & .Fields(2)
         'End If
         'end 2010/12/28
         'Add by Morgan 2007/4/25
         txtPriority(3) = "" & .Fields(6)
         'Added by Lydia 2016/11/01 預設代入專利種類
         If txtCaseNo.Text <> "" Then CheckKeyIn 3
         
         cboCaseName.Clear
         'Modified by Lydia 2016/10/18
         'cboCaseName.AddItem "" & .Fields(3)
         cboCaseName.AddItem "中: " & .Fields(3)
         cboCaseName.AddItem "英: " & .Fields(4)
         cboCaseName.AddItem "日: " & .Fields(5)
         'end 2016/10/18
         cboCaseName.ListIndex = 0
      End If
   End With
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description
End Sub

Private Sub txtCaseNo_Validate(Cancel As Boolean)
   GetCaseData txtCaseNo
End Sub

Private Sub txtPriority_Change(Index As Integer)
   Select Case Index
      Case 0
         lblNationName = ""
      Case 3
         Label7 = ""
   End Select
End Sub

Private Sub txtPriority_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 4 Then Exit Sub
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPriority_Validate(Index As Integer, Cancel As Boolean)
   'Remove by Morgan 2005/7/4 只要確定時檢查
   'If CheckKeyIn(Index) = -1 Then Cancel = True
   If Cancel Then
      txtPriority_GotFocus Index
   'Add by Morgan 2005/6/6
   '加控制台灣優先權號預設0開頭
   Else
      If Index = 0 Then
         CheckKeyIn Index
         If txtPriority(0) = "000" Then
            If txtPriority(2) = "" Then txtPriority(2) = "0"
         Else
            If txtPriority(2) = "0" Then txtPriority(2) = ""
         End If
      'Added by Lydia 2016/10/18
      ElseIf Index = 2 Then
         CheckKeyIn Index
      ElseIf Index = 3 Then
         CheckKeyIn Index
      'end 2016/10/18
      ElseIf Index = 5 Then
         CheckKeyIn Index
      End If
   End If
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String
Dim nCount As Integer
Dim nIndex As Integer
   
   CheckKeyIn = -1
   Select Case intIndex
      Case 0 '優先權國家
         If txtPriority(intIndex) = "" Then
            CheckKeyIn = 0
         'edit by nickc 2007/02/02 不用 dll 了
         'ElseIf objPublicData.GetNation(txtPriority(intIndex).Text, strTemp) Then
         ElseIf ClsPDGetNation(txtPriority(intIndex).Text, strTemp) Then
            lblNationName.Caption = strTemp
            CheckKeyIn = 1
         End If
      Case 1 '優先權日
         If txtPriority(intIndex) = "" Then
            'Add by Morgan 2008/8/28 優先權號為本所案號的例外
            'Modify by Morgan 2010/8/9 也可以輸CFP案
            'If InStr("000,020,044", txtPriority(0)) > 0 And Left(txtPriority(2), 1) = "P" Then
            'Modified by Morgan 2025/9/10 規則要和優先權號檢查一樣否則本所案號若有輸入"-"會沒檢查到
            'If Left(txtPriority(2), 1) = "P" Or Left(txtPriority(2), 3) = "CFP" Then
            If (Left(txtPriority(2), 1) = "P" And IsNumeric(Mid(txtPriority(2), 2, 1))) Or (Left(txtPriority(2), 3) = "CFP" And IsNumeric(Mid(txtPriority(2), 4, 1))) Then
            'end 2025/9/10
               CheckKeyIn = 1
            Else
               MsgBox "優先權日不可空白 !", vbCritical
            End If
         Else
            If CheckIsDate(txtPriority(intIndex).Text) Then
               '不可大於系統日
               If TransDate(txtPriority(intIndex).Text, 2) > strSrvDate(1) Then
                  MsgBox "優先權日不可大於系統日！", vbExclamation
               Else
                  CheckKeyIn = 1
               End If
            End If
         End If
         
      Case 2 '優先權號
         If txtPriority(2) = "" Then
            MsgBox "請輸入優先權號或空白字元!", vbCritical
         Else
            'Add by Morgan 2008/8/28 優先權號為國內本所案號的需檢查
            'Modify by Morgan 2010/8/9 也可以輸CFP案,主張PCT案的申請號也是P開頭(PCTXXX)
            'If InStr("000,020,044", txtPriority(0)) > 0 And Left(txtPriority(2), 1) = "P" Then
            'Modified by Morgan 2019/5/9 CFP案號應該要檢查第4碼為數字
            'If (Left(txtPriority(2), 1) = "P" Or Left(txtPriority(2), 3) = "CFP") And IsNumeric(Mid(txtPriority(2), 2, 1)) Then
            If (Left(txtPriority(2), 1) = "P" And IsNumeric(Mid(txtPriority(2), 2, 1))) Or (Left(txtPriority(2), 3) = "CFP" And IsNumeric(Mid(txtPriority(2), 4, 1))) Then
               strExc(0) = "SELECT PA11,PA08,PA09,PA01||PA02||PA03||PA04 CNO FROM PATENT WHERE " & ChgPatent(txtPriority(2))
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If "" & RsTemp("PA11") <> "" Then
                     MsgBox "輸入之本所案號已有申請案號，請改輸申請案號！"
                  ElseIf txtPriority(1) <> "" Then
                     MsgBox "優先權號為本所案號時不可輸入優先權日！"
                  ElseIf txtPriority(0) <> "" And "" & RsTemp("PA09") <> txtPriority(0) Then
                     MsgBox "優先權國家與本所案號之申請國家不同！"
                  ElseIf txtPriority(3) <> "" And "" & RsTemp("PA08") <> txtPriority(3) Then
                     MsgBox "專利種類與本所案號之申請專利種類不同！"
                  Else
                     txtPriority(2) = "" & RsTemp("CNO")
                     CheckKeyIn = 1
                  End If
               'Add by Morgan 2010/8/12 若有申請日則不管
               ElseIf txtPriority(1) <> "" Then
                  CheckKeyIn = 1
               Else
                  MsgBox "該優先權國家之國內案本所案號不存在！"
               End If
            Else
               'Added by Lydia 2016/10/18 輸入優先權號,存在於基本檔且卷宗性質為'申請'時, 檢查申請日欄與畫面輸入之優先權日是否相同, 同時帶出專利/商標種類及本所案號, 案件名稱(中英日).
               If m_strCaseNo <> "" And Trim(txtPriority(2)) <> "" And bolMsg = False Then
                  If Trim(txtPriority(0)) = "" Then
                     MsgBox "請輸入優先權國家!", vbCritical
                     txtPriority(0).SetFocus
                  Else
                     strExc(0) = "select sk02 from systemkind where sk01='" & strCaseNo(1) & "' "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        If RsTemp.Fields(0) = "1" Then
                           strExc(0) = "select pa01||pa02||pa03||pa04 as CaseNo,pa10,pa23 from patent where pa11='" & Trim(txtPriority(2)) & "' and pa09='" & txtPriority(0) & "' "
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              If Trim(txtPriority(1)) <> "" And RsTemp.Fields("pa23") = "1" And DBDATE(txtPriority(1)) <> "" & RsTemp.Fields("pa10") Then
                                 MsgBox "輸入的優先權日與申請日不同!", vbCritical
                              Else
                                 CheckKeyIn = 1
                              End If
                              If txtCaseNo.Text <> "" & RsTemp.Fields("CaseNo") Then
                                 txtCaseNo.Text = "" & RsTemp.Fields("CaseNo")
                                 GetCaseData txtCaseNo
                              End If
                           End If
                        End If
                     End If
                  End If
               Else
               'end 2016/10/18
                  CheckKeyIn = 1
               End If
            End If
         End If
      'Add by Morgan 2007/4/23
      Case 3 '專利/商標種類
         If txtPriority(3) = "" Then
            'Add By Sindy 2017/11/9 FCP:非電子交換時一定要輸入專利種類
            If strCaseNo(1) = "P" Or strCaseNo(1) = "CFP" Or _
               (txtPriority(0) <> txtPriority(4) And txtPriority(4) <> "") Then
               MsgBox "請輸入專利種類!", vbCritical
            Else
               CheckKeyIn = 1
            End If
         Else
'            If strCaseNo(1) <> "" Then
               'Modify By Sindy 2017/11/9
               'Label7 = GetKindName(strCaseNo(1), txtPriority(3), intI)
               Label7 = GetKindName(IIf(strCaseNo(1) = "", "FCP", strCaseNo(1)), txtPriority(3), intI)
               '2017/11/9 END
               If intI = 1 Then
                  CheckKeyIn = 1
               Else
                  MsgBox "專利/商標種類輸入錯誤!", vbCritical
               End If
'            Else
'               CheckKeyIn = 1
'            End If
         End If
      'Add by Amy  2014/04/07
      Case 4 '存取碼
         'Modify By Sindy 2016/1/12 開放存取碼可以輸入為優先權國家代碼
         'If Trim(txtPriority(4)) <> "" And Len(Trim(txtPriority(4))) <> "4" Then
         If Trim(txtPriority(4)) <> "" And Not (Len(Trim(txtPriority(4))) = "4" Or Trim(txtPriority(4)) = Trim(txtPriority(0))) Then
         '2016/1/12 END
            MsgBox Label13 & "輸入錯誤!", vbCritical
         Else
            CheckKeyIn = 1
         End If
      'Add By Sindy 2017/10/11
      Case 5 '商品類別
         nCount = GetSubStringCount(txtPriority(5))
         For nIndex = 1 To nCount
            strTemp = GetSubString(txtPriority(5), nIndex)
            For nCount = 1 To nCount
               If nIndex <> nCount Then
                  If strTemp = GetSubString(txtPriority(5), nCount) Then
                     MsgBox "商品類別<" & strTemp & ">不可重覆", vbOKOnly, "檢核資料"
                     Call txtPriority_GotFocus(5)
                     Exit Function
                  End If
               End If
            Next nCount
         Next nIndex
         CheckKeyIn = 1
         txtPriority(5) = Replace(txtPriority(5), " ", "")
   End Select
End Function

Private Sub txtPriority_GotFocus(Index As Integer)
   TextInverse txtPriority(Index)
   'Add by Morgan 2005/6/6
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtPriority(Index).IMEMode = 2
   CloseIme
   'Remove by Morgan 2007/9/7 不需切換輸入法
   'If Index = 2 Then
   '   'edit by nickc 2007/06/06 切換輸入法改用API
   '   'txtPriority(Index).SelStart = 1
   '   OpenIme
   'End If
   'end 2007/9/7
End Sub

'Modified by Lydia 2016/10/13 +判斷國別c_PA09、申請日c_PA10、讀取本所案號 p_CaseNo
'Private Function ConfirmAppNo(ByVal p_AppNo As String, p_AppDate As String, p_AppCountry As String) As Boolean
Private Function ConfirmAppNo(ByVal p_AppNo As String, p_AppDate As String, p_AppCountry As String, Optional ByVal c_PA09 As String, Optional ByVal c_PA10 As String, Optional ByRef p_CaseNo As String) As Boolean
   Dim strMsg As String, bolNoMatch As Boolean
   strSql = "SELECT PA01,PA02,PA03,PA04,PA09,PA10 FROM PATENT WHERE PA11='" & p_AppNo & "'"
   'Added by Lydia 2016/10/13
   'Modified by Lydia 2017/03/16 判斷非空白trim
   If Trim(c_PA09) <> "" Then
      strSql = strSql & " AND PA09='" & c_PA09 & "'"
   End If
   'Modified by Lydia 2017/03/16 判斷非空白trim
   If Trim(c_PA10) <> "" Then
      strSql = strSql & " AND PA10='" & c_PA10 & "'"
   End If
   'end 2016/10/13
   
On Error GoTo ErrHnd
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         'Added by Lydia 2016/10/13 讀案號,不彈訊息
         If p_CaseNo <> "" Then
            bolNoMatch = False
            p_CaseNo = Trim("" & .Fields("PA01") & .Fields("PA02") & .Fields("PA03") & .Fields("PA04"))
         Else
         'end 2016/10/13
            strMsg = strMsg & "本所案件 " & .Fields("PA01") & "-" & .Fields("PA02") & "-" & .Fields("PA03") & "-" & .Fields("PA04") & _
               " 與該筆優先權資料申請案號相同但"
            If "" & .Fields("PA09") <> p_AppCountry Then
               strMsg = strMsg & "申請國家(" & .Fields("PA09") & ")不同"
               bolNoMatch = True
            End If
            If "" & .Fields("PA10") <> p_AppDate Then
               If bolNoMatch = True Then
                  strMsg = strMsg & "且申請日(" & .Fields("PA10") & ")也不同"
               Else
                  strMsg = strMsg & "申請日(" & .Fields("PA10") & ")不同"
               End If
               bolNoMatch = True
            End If
            strMsg = strMsg & "，是否確定要繼續？"
            p_CaseNo = Trim("" & .Fields("PA01") & .Fields("PA02") & .Fields("PA03") & .Fields("PA04")) 'Added by Lydia 2016/10/13
         End If
      End If
      
      'Added by Lydia 2016/10/13 讀案號,不彈訊息
       If p_CaseNo = "R" Then
           bolNoMatch = False
           p_CaseNo = ""
       End If
       
   End With
   
   If bolNoMatch = True Then
      If MsgBox(strMsg, vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
         ConfirmAppNo = True
      End If
   Else
      ConfirmAppNo = True
   End If
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Function CheckIdentical() As Boolean
   Dim varPriorityTemp1, varPriorityTemp2, varPriorityTemp3, varPriorityTemp4
   Dim varPriorityTemp1old, varPriorityTemp2old, varPriorityTemp3old, varPriorityTemp4old
   Dim varPriorityPD09, varPriorityPD09old  'Add by Amy 2014/04/07
   Dim ii As Integer, jj As Integer, bolFound As Boolean
   Dim varPriorityTemp6, varPriorityTemp6old 'Add By Sindy 2017/9/29
   
   varPriorityTemp1 = Split(strPriority1, "，")
   varPriorityTemp2 = Split(strPriority2, "，")
   varPriorityTemp3 = Split(strPriority3, "，")
   varPriorityTemp4 = Split(strPriority4, "，")
   varPriorityPD09 = Split(strPD09, "，") 'Add by Amy 2014/04/07
   varPriorityTemp6 = Split(strPriority6, "，") 'Add By Sindy 2017/10/11
   
   varPriorityTemp1old = Split(strPriority1old, "，")
   varPriorityTemp2old = Split(strPriority2old, "，")
   varPriorityTemp3old = Split(strPriority3old, "，")
   varPriorityTemp4old = Split(strPriority4old, "，")
   varPriorityPD09old = Split(strPD09old, "，") 'Add by Amy 20014/04/07
   varPriorityTemp6old = Split(strPriority6old, "，") 'Add By Sindy 2017/9/29
   
   If UBound(varPriorityTemp1) <> UBound(varPriorityTemp1old) Then
      Exit Function
   End If
   If UBound(varPriorityTemp2) <> UBound(varPriorityTemp2old) Then
      Exit Function
   End If
   If UBound(varPriorityTemp3) <> UBound(varPriorityTemp3old) Then
      Exit Function
   End If
   'Add by Morgan 2007/4/25
   If strPriority4old <> "" And UBound(varPriorityTemp4) <> UBound(varPriorityTemp4old) Then
      Exit Function
   End If
   'Add by Amy 2014/04/07
   If (strPD09 <> "" Or strPD09old <> "") And UBound(varPriorityPD09) <> UBound(varPriorityPD09old) Then
      Exit Function
   End If
   'Add by Sindy 2017/10/11
   If (strPriority6 <> "" Or strPriority6old <> "") And UBound(varPriorityTemp6) <> UBound(varPriorityTemp6old) Then
      Exit Function
   End If
   For ii = LBound(varPriorityTemp1) To UBound(varPriorityTemp1)
      bolFound = False
      For jj = LBound(varPriorityTemp1old) To UBound(varPriorityTemp1old)
         If varPriorityTemp1(ii) = varPriorityTemp1old(jj) Then
            If varPriorityTemp2(ii) = varPriorityTemp2old(jj) Then
               If varPriorityTemp3(ii) = varPriorityTemp3old(jj) Then
                   'Modify by Amy 2014/04/07
'                  'Add by Morgan 2007/4/25
'                  If strPriority4Old = "" Then
'                     bolFound = True
'                     Exit For
'                  'end 2007/4/25
'                  ElseIf varPriorityTemp4(ii) = varPriorityTemp4Old(jj) Then
'                     bolFound = True
'                     Exit For
'                  End If
                  If strPriority4old = "" Then
                     If strPD09 = "" Then
                        bolFound = True
                        Exit For
                     ElseIf varPriorityPD09(ii) = varPriorityPD09old(jj) Then
                        bolFound = True
                        Exit For
                     End If
                  ElseIf varPriorityTemp4(ii) = varPriorityTemp4old(jj) Then
                     If strPD09 = "" Then
                        bolFound = True
                        Exit For
                     ElseIf varPriorityPD09(ii) = varPriorityPD09old(jj) Then
                        bolFound = True
                        Exit For
                     End If
                  End If
                  'end 2014/04/07
               End If
            End If
         End If
      Next
      If bolFound = False Then Exit Function
   Next
   
   CheckIdentical = True
End Function

'Add by Morgan 2005/11/7 檢查申請人是否一致
Private Function CheckApplicant() As Boolean
   Dim App1(1 To 5) As String, App2(1 To 5) As String, idx As Integer, intI As Integer, intJ As Integer
   Dim stPA09 As String, stPA10 As String, stPA11 As String
   
   stPA09 = Trim(txtPriority(0))
   stPA10 = DBDATE(txtPriority(1))
   stPA11 = Trim(txtPriority(2))
   
   strSql = " select pa26,pa27,pa28,pa29,pa30 from patent where " & ChgPatent(m_strCaseNo)
   strSql = strSql & " union all"
   strSql = strSql & " select pa26,pa27,pa28,pa29,pa30 from patent where pa09='" & stPA09 & "' and pa10=" & stPA10 & " and pa11='" & stPA11 & "'"
   
On Error GoTo ErrHnd

   CheckOC3
   CheckApplicant = True
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount = 2 Then
         intI = 0: intJ = 0
         For idx = 1 To 5
            If Not IsNull(.Fields("PA" & Format(25 + idx))) Then
               App1(idx) = .Fields("PA" & Format(25 + idx))
               intI = intI + 1
            End If
         Next
         .MoveNext
         For idx = 1 To 5
            If Not IsNull(.Fields("PA" & Format(25 + idx))) Then
               App2(idx) = .Fields("PA" & Format(25 + idx))
               intJ = intJ + 1
            End If
         Next
         '申請人數不一樣
         If intI <> intJ Then
            CheckApplicant = False
         Else
            For intI = 1 To 5
               If App1(intI) <> "" Then
                  For intJ = 1 To 5
                    If App2(intJ) = App1(intI) Then
                     Exit For
                    End If
                  Next
                  '申請人不一樣
                  If intJ = 6 Then
                     CheckApplicant = False
                     Exit For
                  End If
               End If
            Next
         End If
      End If
   End With
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function

'Add by Morgan 2006/5/15 檢查主張國內優先權的優先權國家必須與申請國相同(國際優先權則須不同)
Private Function CheckCountryMatch(ByVal p_CP1234 As String, ByVal p_PD07 As String) As Boolean
   Dim stMsg As String
   
   'Modified by Morgan 2012/12/21 +124
   strSql = "select CP10,PA09 FROM CASEPROGRESS,PATENT WHERE " & ChgCaseprogress(p_CP1234) & " AND CP10 IN ('121','106','124') AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         'Modified by Morgan 2014/4/3
         'If "" & .Fields(0) = "121" And "" & .Fields(1) <> p_PD07 Then
         '   MsgBox "主張國內優先權時，優先權國家必須與申請國相同！", vbExclamation
         'ElseIf "" & .Fields(0) = "106" And "" & .Fields(1) = p_PD07 Then
         '   MsgBox "主張國際優先權時，優先權國家不可與申請國相同！", vbExclamation
         'Else
         '   CheckCountryMatch = True
         'End If
         
         .MoveFirst
         .Find "CP10='124'"
         If Not .EOF Then
            CheckCountryMatch = True
         Else
            .MoveFirst
            If .Fields(1) = p_PD07 Then
               .Find "CP10='121'"
               stMsg = "優先權國家與申請國相同時必須有收文主張國內優先權！"
            Else
               .Find "CP10='106'"
               stMsg = "優先權國家與申請國不相同時必須有收文主張國際優先權！"
            End If
            If Not .EOF Then
               stMsg = ""
            End If
         End If
         'end 2014/4/3
      Else
         stMsg = "本案主張優先權尚未收文！"
         
         'Added by Morgan 2019/7/16
         '寰華案若是有收文"935案件轉至本所"表示中間來所補資料可不必收文--敏莉
         If Pub_StrUserSt03 = "F22" Then
            cnnConnection.Execute "update CASEPROGRESS set cp26=cp26 WHERE " & ChgCaseprogress(p_CP1234) & " and cp10='935'", intI
            If intI > 0 Then
               stMsg = ""
            End If
         End If
         
      End If
      
      If stMsg <> "" Then
         MsgBox stMsg, vbExclamation
      Else
         CheckCountryMatch = True
      End If
   End With
End Function

'Add by Morgan 2007/4/23 抓專利/商標種類
Private Function GetKindName(p_SysCode As String, p_KindCode As String, ByRef p_iRet As Integer) As String
   strExc(0) = "select ptm03 from systemkind,patenttrademarkmap where sk01='" & p_SysCode & "' and ptm01(+)=sk02 and ptm02='" & p_KindCode & "'"
   p_iRet = 1
   Set RsTemp = ClsLawReadRstMsg(p_iRet, strExc(0))
   If p_iRet = 1 Then
      GetKindName = "" & RsTemp.Fields(0)
   End If
End Function

'Add By Sindy 2017/10/6
Private Function GRIDHEAND()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   '                        0         1             2       3           4         5           6           7
   arrGridHeadText = Array("國代碼", "優先權國家", "種類", "本所案號", "存取碼", "優先權日", "優先權號", "商品類別")
   arrGridHeadWidth = Array(400, 1000, 1000, 1000, 1000, 850, 1500, 2000)
   GRD1.Cols = UBound(arrGridHeadText) + 1
   'Added by Morgan 2018/4/11 要有固定列否則欄寬無法調整
   GRD1.Rows = 2
   GRD1.FixedRows = 1
   'end 2018/4/11
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      'Modified by Morgan 2018/4/11
      GRD1.CellAlignment = flexAlignLeftCenter
      'Grd1.ColAlignmentFixed(iRow) = flexAlignLeftCenter
      'end 2018/4/11
      GRD1.ColAlignment(iRow) = flexAlignLeftCenter 'Added by Morgan 2018/7/17
   Next
End Function
