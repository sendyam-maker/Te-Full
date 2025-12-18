VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050102_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文"
   ClientHeight    =   5604
   ClientLeft      =   5376
   ClientTop       =   2088
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5604
   ScaleWidth      =   9300
   Begin VB.CommandButton cmdOK 
      Caption         =   "延期(&D)"
      Height          =   400
      Index           =   3
      Left            =   5610
      TabIndex        =   21
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3732
      Left            =   144
      TabIndex        =   12
      Top             =   1680
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   6583
      _Version        =   393216
      Cols            =   20
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
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
      _Band(0).Cols   =   20
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發文資料(&F)"
      Default         =   -1  'True
      Height          =   400
      Index           =   2
      Left            =   6420
      TabIndex        =   13
      Top             =   70
      Width           =   1100
   End
   Begin VB.Frame fraChoose 
      BorderStyle     =   0  '沒有框線
      Height          =   252
      Index           =   0
      Left            =   1380
      TabIndex        =   19
      Top             =   990
      Width           =   1692
      Begin VB.TextBox txtReceiveCode 
         Height          =   264
         Left            =   0
         MaxLength       =   9
         TabIndex        =   10
         Top             =   0
         Width           =   1452
      End
   End
   Begin VB.Frame fraChoose 
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   372
      Index           =   1
      Left            =   1380
      TabIndex        =   16
      Top             =   660
      Width           =   3492
      Begin VB.Frame fraElse 
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   840
         TabIndex        =   17
         Top             =   0
         Width           =   2652
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   2
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   4
            Top             =   0
            Width           =   492
         End
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   1
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   3
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   0
            Left            =   0
            MaxLength       =   6
            TabIndex        =   2
            Top             =   0
            Width           =   1212
         End
      End
      Begin VB.TextBox txtSystem 
         Height          =   288
         Left            =   0
         MaxLength       =   3
         TabIndex        =   1
         Top             =   0
         Width           =   732
      End
      Begin VB.Frame fraTF 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   840
         TabIndex        =   18
         Top             =   0
         Width           =   2652
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   3
            Left            =   2040
            TabIndex        =   8
            Top             =   0
            Width           =   492
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   2
            Left            =   1560
            TabIndex        =   7
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   1
            Left            =   1080
            TabIndex        =   6
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   972
         End
      End
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "本所案號："
      CausesValidation=   0   'False
      Height          =   252
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   660
      Value           =   -1  'True
      Width           =   1212
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "收文號："
      CausesValidation=   0   'False
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Top             =   990
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7530
      TabIndex        =   14
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8376
      TabIndex        =   15
      Top             =   70
      Width           =   800
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1380
      TabIndex        =   11
      Top             =   1320
      Width           =   7815
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "13785;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label24 
      Caption         =   "案件名稱："
      Height          =   252
      Left            =   180
      TabIndex        =   20
      Top             =   1320
      Width           =   972
   End
End
Attribute VB_Name = "frm050102_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/4 改成Form2.0 (grdDataList,cboCaseName)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Copy from frm020102_1 by Morgan 2009/7/24
Option Explicit

Dim intLastRow As Integer, blnOKtoShow As Boolean
Dim intOpt As Integer
Dim m_blnExcClear As Boolean '是否執行Clear段
Public bolIsEMPFlow As Boolean 'Add By Sindy 2013/5/20 是否為電子承辦簽核
'Add By Sindy 2018/1/8
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String, m_AppNo As String
Dim m_Done As Boolean
'2018/1/8 END


'Modify By Sindy 2013/5/17
'Private Sub cmdOK_Click(index As Integer)
Public Sub cmdok_Click(Index As Integer)
'2013/5/17 End
Dim i As Integer, bolRt As Integer

   Select Case Index
      Case 0 '確定
         If Me.grdDataList.TextMatrix(1, 0) = "" Then Exit Sub
         '已閉卷案件不可發文
         If PUB_CaseClosedCP09(Me.grdDataList.TextMatrix(Me.grdDataList.row, 5)) Then Exit Sub
         '未輸入承辦人不可發文
         If PUB_ChkCP14IsNull(Me.grdDataList.TextMatrix(Me.grdDataList.row, 5)) = True Then Exit Sub
         'Add by Amy 2015/01/22 北所未分案不可發文
         If Me.grdDataList.TextMatrix(grdDataList.row, 5) < "B" And Me.grdDataList.TextMatrix(grdDataList.row, 11) = "" Then MsgBox "北所尚未分案，不可發文!!": Exit Sub
         '2009/4/20 ADD BY SONIA '美專母案領證發文需檢查CIP,CA或分割或CPA(但限設計)案未發文或已發文未提申則母案不可發文
         If ChkChild(Me.grdDataList.TextMatrix(Me.grdDataList.row, 5)) = True Then Exit Sub
         '2009/4/20 END
         '2009/10/23 add by sonia 新加坡發明實審若未提檢索報告則提醒操作者,不可發文
         If grdDataList.TextMatrix(grdDataList.row, 8) = "014" And grdDataList.TextMatrix(grdDataList.row, 7) = "416" Then
            If Chk014416(Me.grdDataList.TextMatrix(Me.grdDataList.row, 5)) = True Then Exit Sub
         End If
         '2009/10/23 end
         '2011/5/18 add by sonia 比利時209及西班牙211發明申請若未收文申請檢索報告則提醒操作者,不可發文
         'Modified by Morgan 2012/10/29 比利時209取消--甄妮
         'If (grdDataList.TextMatrix(grdDataList.row, 8) = "209" Or grdDataList.TextMatrix(grdDataList.row, 8) = "211") And grdDataList.TextMatrix(grdDataList.row, 7) = "101" Then
         If grdDataList.TextMatrix(grdDataList.row, 8) = "211" And grdDataList.TextMatrix(grdDataList.row, 7) = "101" Then
            If Chk421rec(Me.grdDataList.TextMatrix(Me.grdDataList.row, 5)) = True Then Exit Sub
         End If
         '2011/5/18 end
         
         'Added by Morgan 2023/2/14 檢查美國申請案(發明、設計、CIP、CPA)有申請人非個人且未收文讓渡,不可發文。(分割、暫時申請不用)
         'Modified by Morgan 2023/2/20 改條件:申請人非個人-->申請人與發明人不完全相同
         If grdDataList.TextMatrix(grdDataList.row, 8) = "101" And InStr("101,103,113,114", grdDataList.TextMatrix(grdDataList.row, 7)) > 0 Then
            If ChkUsNeed701(Me.grdDataList.TextMatrix(Me.grdDataList.row, 5)) = True Then
               'MsgBox "美國申請案(發明、設計、CIP、CPA)有申請人非個人且未收文讓渡,不可發文!!", vbExclamation
               MsgBox "美國申請案(發明、設計、CIP、CPA)未收文讓渡且申請人與發明人不完全相同,不可發文!!", vbExclamation
               Exit Sub
            End If
         End If
         'end 2023/2/14
         
         If PUB_ChkCP141IsSend(grdDataList.TextMatrix(Me.grdDataList.row, 5), True) = False Then Exit Sub 'Added by Morgan 2024/1/26
         
         If Where020102ToGo = False Then Exit Sub
         
         cmdOK(2).SetFocus
         Me.Hide
            
      Case 1 '結束
         Unload Me
         
      Case 2 '發文資料
         '選擇收文號
         If intOpt = 0 Then
            If CheckKeyIn3 Then
               'Add By Cheng 2003/03/26
               '取得本所案號
               GetOurCaseNo Me.txtReceiveCode.Text
               GetSendCaseData
            Else
               TextInverse txtReceiveCode
            End If
         '選擇本所案號
         'Add by Morgan 2004/2/17
         '控制只能是'CFP'和'CPS'的案件
         ElseIf txtSystem <> "CFP" And txtSystem <> "CPS" Then
             MsgBox "系統類別必須為 'CFP' 或 'CPS'！！", vbCritical
         Else
            If txtSystem = 馬德里案 Then
               bolRt = CheckKeyIn1(3)
            Else
               bolRt = CheckKeyIn2(2)
            End If
            If bolRt = 1 Then
               GetSendCaseData
            Else
               txtSystem.SetFocus
            End If
         End If
         
      Case 3 '延期
         If Me.grdDataList.TextMatrix(1, 0) = "" Then Exit Sub
         '所點選的案件性質不可為"延期"
         If PUB_CPKindDelay(Me.grdDataList.TextMatrix(Me.grdDataList.row, 5), "P") Then
            Exit Sub
         End If
         '已閉卷案件不可發文
         If PUB_CaseClosedCP09(Me.grdDataList.TextMatrix(Me.grdDataList.row, 5)) Then
            Exit Sub
         End If
         
         'Add By Sindy 2013/11/14
         '檢查是否有承辦歷程是否有產生承辦單可以發文
         If PUB_IsEmpFlowIsSend(Me.grdDataList.TextMatrix(Me.grdDataList.row, 5)) = False Then
            Exit Sub
         End If
         
         'Add By Sindy 2018/1/8
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtSystem & txtCode(0) & txtCode(1) & txtCode(2) Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
               Exit Sub
            End If
         End If
         '2018/1/8 END
         '延期記錄資料來源為案件進度檔
         frm050102_2.m_str_DL05 = "1"
         '此按鈕只有在外專才有
         frm050102_2.intWhereComeFrom = 1
         frm050102_2.Show
         Me.Hide
         
   End Select
End Sub

'2009/10/23 add by sonia 新加坡發明實審若未提檢索報告則提醒操作者,不可發文
Public Function Chk014416(strCP09 As String) As Boolean
Dim StrSQLa As String
Dim rsA  As New ADODB.Recordset
   
   Chk014416 = False
   'Modified by Morgan 2012/11/7 +PCT案除外(PCT已有檢索報告)
   StrSQLa = "Select c2.cp09,c2.cp57,c2.cp27 From caseProgress c1,caseprogress c2,patent WHERE c1.cp09='" & strCP09 & "' " & _
            " and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and c2.cp10(+)='421' " & _
            " and c1.cp01=pa01 and c1.cp02=pa02 and c1.cp03=pa03 and c1.cp04=pa04 and '1'=pa08 and pa46 is null"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If Not IsNull(rsA.Fields(0)) Then
         If IsNull(rsA.Fields(2)) Then
            Chk014416 = True
         End If
      Else
         Chk014416 = True
      End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   If Chk014416 = True Then
      'modify by sonia 2017/11/27 改為提醒仍可選擇繼續發文CFP-028046(禧佩:新加坡是可提其他國家的檢索報告然後在新加坡提實審,檢索不一定要在新加坡提)
      'MsgBox "請提醒智權同仁：新加坡發明必須先提檢索報告才可提實審！", vbExclamation + vbOKOnly
      If MsgBox("新加坡發明必須先提檢索報告才可提實審！本案應不可發文！是否仍要繼續發文？", vbYesNo + vbDefaultButton2) = vbNo Then
         Me.Show
         Exit Function
      End If
      Chk014416 = False
      'end 2017/11/27
   End If
End Function
'2009/10/23 end

Private Sub GetSendCaseData(Optional bolBlank As Boolean = False)
'TF為馬德里案，另外判斷
   If bolBlank = True Then
      GetSendData "", "", "", "", ""
   Else
      If txtSystem = 馬德里案 Then
         GetSendData txtReceiveCode, txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
            IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3))
      Else
         GetSendData txtReceiveCode, txtSystem, txtCode(0), _
            IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
      End If
   End If
End Sub

Private Sub GetSendData(ByRef strReceiveCode As String, ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String)
Dim varSaveCursor
Dim ii As Integer  'ADD BY SONIA 2014/5/13
   
   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   'Modify by Morgan 2004/2/17
   '將C類收文併入
   'Set grdDataList.Recordset = objPublicData.ReadSendCaseRst(intOpt, intPCaseKind, intPWhere, strGroup, strReceiveCode, strCode1, strCode2, strCode3, strCode4)
   Set grdDataList.Recordset = ReadSendCaseRst(intOpt, intPCaseKind, intPWhere, strGroup, strReceiveCode, strCode1, strCode2, strCode3, strCode4)
   'ADD BY SONIA 2014/5/13 加相關總收號案件性質
   For ii = 1 To Me.grdDataList.Rows - 1
       Me.grdDataList.TextMatrix(ii, 1) = Me.grdDataList.TextMatrix(ii, 1) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 5), "1")
   Next ii
   'END 2014/5/13
   SetDataListVision grdDataList, True
   intLastRow = 0
   If grdDataList.Rows > 1 Then
      ShowBar grdDataList, intLastRow, 10
      cmdOK(0).Enabled = True
      cmdOK(3).Enabled = True
      cmdOK(0).Default = True
   Else
      cmdOK(0).Enabled = False
      cmdOK(3).Enabled = False
      cmdOK(2).Default = True
      MsgBox "無符合條件資料 !", vbCritical
   End If
   Screen.MousePointer = varSaveCursor
End Sub

'讀取發文資料,intCaseKind系統分類
'Copy from clsPublicData by Morgan 2004/2/17
Private Function ReadSendCaseRst(ByRef intOpt As Integer, ByRef intCaseKind As Integer, ByRef intWhere As Integer, ByRef strGroup As String, ByRef strReceiveCode As String, ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String) As ADODB.Recordset
   Dim strSql As String, strSQL1 As String, rsRecordset As New ADODB.Recordset
   Dim strDateLine As String
        
    'Modify by Amy 2015/01/22 +CP157
    If intWhere <> 國外_CF Then strDateLine = "-1911"
    strSql = "select " & SQLDate("cp05") & " s01,"
    strSQL1 = "select " & SQLDate("cp05") & " s01,decode(sp09," + CNULL(大陸國家代號) + ",cpm04,cpm03) s02,staff.st02 s03,staff1.st02 s04,cp64 s05,cp09,cp79,cp10,sp09 s06,sk02,staff1.st01 s07,cp157 from caseprogress,servicepractice,casepropertymap,staff,staff staff1,systemkind where sk01=sp01 and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and cp01=cpm01(+) and cp10=cpm02 and cp14=staff.st01(+) and cp13=staff1.st01(+) "
    strSql = strSql + "decode(pa09," + CNULL(大陸國家代號) + ",cpm04,cpm03) s02,staff.st02 s03,staff1.st02 s04,cp64 s05,cp09,cp79,cp10,pa09 s06,sk02,staff1.st01 s07,cp157 from caseprogress,patent,casepropertymap,staff,staff staff1,systemkind where sk01=cp01 and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04"
    strSql = strSql + " and cp01=cpm01(+) and cp10=cpm02 and cp14=staff.st01(+) and cp13=staff1.st01(+)"
    If intOpt = 1 Then
       strSql = strSql + " and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4)
       strSQL1 = strSQL1 + " and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4)
    Else
       strSql = strSql + " and cp09=" + CNULL(strReceiveCode)
       strSQL1 = strSQL1 + " and cp09=" + CNULL(strReceiveCode)
    End If
    strSql = strSql + " and cp27 is null and cp57 is null"
    strSQL1 = strSQL1 + " and cp27 is null and cp57 is null"
    'Modify by Morgan 2010/8/11 百年蟲
    'strSql = "select s01 收文日,s02 案件性質,s03 承辦人,s04 智權人員,s05 進度備註,cp09 總收文號,cp79 未收金額,cp10 案件性質,s06 申請國家,sk02 系統種類,s07 智權人員代號 from (" + strSql + " union " + strSQL1 + ") order by s01"
    strSql = "select substrb(' '||s01,-9) 收文日,s02 案件性質,s03 承辦人,s04 智權人員,s05 進度備註,cp09 總收文號,cp79 未收金額,cp10 案件性質,s06 申請國家,sk02 系統種類,s07 智權人員代號,cp157 from (" + strSql + " union " + strSQL1 + ") order by 1"
    'end 2015/01/22
    
    'edit by nickc 2007/02/02 不用 dll 了
    'Set ReadSendCaseRst = objPublicData.ReadRst(StrSql)
    Set ReadSendCaseRst = ClsPDReadRst(strSql)
End Function

Private Sub SetDataListWidth()
Dim varGridWidth() As Variant
'Add by Amy 2015/01/22 +cp157
varGridWidth = Array(900, 1500, 900, 900, 4150, 1000, 1500, 0, 0, 0, 0, 0)
SetGridDataListWidth grdDataList, varGridWidth()
blnOKtoShow = True
End Sub

'Add By Sindy 2019/7/26
Private Sub GridHead()
Dim i As Integer
 
   blnOKtoShow = True
   FixGrid grdDataList
   With grdDataList
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 900: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1500: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 900: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 900: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 4150: .Text = "進度備註"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1000: .Text = "總收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1500: .Text = "未收金額"
      .col = 7: .ColWidth(7) = 0: .Text = "案件性質"
      .col = 8: .ColWidth(8) = 0: .Text = "申請國家"
      .col = 9: .ColWidth(9) = 0: .Text = "系統種類"
      .col = 10: .ColWidth(10) = 0: .Text = "智權人員代號"
      For i = 7 To .Cols - 1
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Public Sub QueryData()
   GetSendCaseData True
   If intPCaseKind = 專利 Then
      cmdOK(3).Visible = True
   Else
      cmdOK(3).Visible = False
   End If
   'Modify By Sindy 2010/8/17 比對自動編號年度
   'If optChoose(0).Value Then txtReceiveCode = 接洽記錄單 + GetTaiwanThisYear
   If optChoose(0).Value Then txtReceiveCode = 接洽記錄單 + CompAutoNumberYear(GetTaiwanThisYear)
End Sub

Private Sub Form_Activate()
   'Modify By Cheng 2002/03/08
   '若執行了發文後, 回到此畫面時, 將游標預設於本所案號欄位
   'Add By Cheng 2002/01/09
   If Me.Visible Then
      If m_blnExcClear = False Then
         If Me.optChoose(0).Value Then
            optChoose_Click 0
         Else
            optChoose_Click 1
         End If
      Else
         Me.optChoose(1).Value = True
         optChoose_Click 1
         m_blnExcClear = False
      End If
   End If
   cmdOK(2).Default = True
   txtSystem = "CFP"
   
   'Added by Sindy 2018/1/8
   If m_strIR01 <> "" And m_Done = False Then
      txtSystem.Text = m_strCP01
      txtCode(0).Text = m_strCP02
      txtCode(1).Text = m_strCP03
      txtCode(2).Text = m_strCP04
      cmdOK(2).Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2018/1/8 END
End Sub

'Private Sub Form_Activate()
   'GetSendCaseData True
   'If intPCaseKind = 專利 Then
   '   cmdOK(3).Visible = True
   'Else
   '   cmdOK(3).Visible = False
   'End If
   'If intOpt = 0 Then
   '   txtReceiveCode.SetFocus
   'Else
   '   txtSystem.SetFocus
   'End If
   'If optChoose(0).Value Then txtReceiveCode = 接洽記錄單 + GetTaiwanThisYear
'End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'SetDataListWidth
   GridHead
   
   intOpt = 0
   m_blnExcClear = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'edit by nickc 2007/02/06 不用 dll 了
   'Set obj002 = Nothing
   
   bolIsEMPFlow = False 'Add By Sindy 2013/5/20
   'Add By Cheng 2002/07/18
   Set frm050102_1 = Nothing
End Sub

Private Sub grdDataList_DblClick()
   cmdok_Click 0
End Sub

Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 1 And grdDataList.Rows > 1 Then
      cmdOK(0).Default = True
   End If
End Sub

Private Sub optChoose_Click(Index As Integer)
   intOpt = Index
   Select Case Index
             Case 0 '收文號
                        fraChoose(0).Enabled = True
                        fraChoose(1).Enabled = False
                        If txtReceiveCode.Visible = True And txtReceiveCode.Enabled = True Then 'Add By Sindy 2014/5/20 +if 因為在待送件區呼叫此Form時,會出現”程式執行階段5...錯誤訊息”
                           txtReceiveCode.SetFocus
                        End If
             Case 1 '本所案號
                        fraChoose(0).Enabled = False
                        fraChoose(1).Enabled = True
                        'modify by sonia 2014/5/13
                        'txtSystem.SetFocus
                        If txtCode(0).Visible = True And txtCode(0).Enabled = True Then 'Add By Sindy 2014/5/20 +if 因為在待送件區呼叫此Form時,會出現”程式執行階段5...錯誤訊息”
                           txtCode(0).SetFocus
                        End If
                        'end 2014/5/13
   End Select
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_LostFocus(Index As Integer)
   If CheckKeyIn2(Index) = -1 Then
      txtCode(0).SetFocus
   End If
End Sub

Private Sub txtReceiveCode_Change()
   'If cboCaseName.ListCount > 0 Then cboCaseName.Clear
   'If grdDataList.Rows > 1 Then GetSendCaseData True
End Sub

Private Sub txtReceiveCode_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtReceiveCode_LostFocus()
   If txtReceiveCode <> "" Then
      If CheckKeyIn3 = False Then
         txtReceiveCode.SetFocus
      End If
   End If
End Sub

Private Sub txtSystem_Change()
   If txtSystem.Text = 馬德里案 Then
      fraTF.Visible = True
      fraElse.Visible = False
   Else
      fraTF.Visible = False
      fraElse.Visible = True
   End If
   If cboCaseName.ListCount > 0 Then cboCaseName.Clear
   'If grdDataList.Rows > 1 Then GetSendCaseData True
End Sub

Private Sub txtSystem_GotFocus()
   txtSystem.SelStart = 0
   txtSystem.SelLength = Len(txtSystem.Text)
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetGroupCase(txtSystem, strGroup) = False Then
   If ClsPDGetGroupCase(txtSystem, strGroup) = False Then
   
      ShowMsg MsgText(9171)
      Cancel = True
      txtSystem_GotFocus
   End If
End Sub

Private Function CheckKeyIn3() As Boolean
Dim strCode1 As String, strCode2 As String, strCode3 As String, strCode4 As String
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim intCaseKind As Integer, intWhere As Integer

On Error GoTo ErrHand
   'Modify by Morgan 2004/2/17
   '將C類收文併入
   'If objPublicData.CheckRecieveCode(txtReceiveCode, strCode1, strCode2, strCode3, strCode4) <> 0 Then
   If oCheckRecieveCode(txtReceiveCode, strCode1, strCode2, strCode3, strCode4) <> 0 Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetGroupCase(strCode1, strGroup) = False Then
      If ClsPDGetGroupCase(strCode1, strGroup) = False Then
      
         ShowMsg MsgText(9171)
       'Add by Morgan 2004/2/17
       '控制只能是'CFP'和'CPS'的案件
       ElseIf strCode1 <> "CFP" And strCode1 <> "CPS" Then
            MsgBox "系統類別必須為 'CFP' 或 'CPS'！！", vbCritical
       
      Else
       'Modify by Morgan 2004/2/17
       '將C類收文併入
         'If objPublicData.CheckCaseCodeIsExist(strCode1, strCode2, strCode3, strCode4, strCaseName1, strCaseName2, strCaseName3, , , , , False) Then
         'Modified by Morgan 2020/9/3
         'If oCheckCaseCodeIsExist(strCode1, strCode2, strCode3, strCode4, strCaseName1, strCaseName2, strCaseName3, , , , , False) Then
         If ClsPDCheckCaseCodeIsExist(strCode1, strCode2, strCode3, strCode4, strCaseName1, strCaseName2, strCaseName3, , , , , False) Then
         'end 2020/9/3
            SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
            CheckKeyIn3 = True
         Else
            ShowMsg MsgText(9176)
         End If
      End If
   End If
   Exit Function
ErrHand:
   ErrorMsg
End Function

'判斷收文號是否存在
'Copy from clsPublicData by Morgan 2004/2/17
Private Function oCheckRecieveCode(ByRef strRecieve As String, ByRef strCode0 As String, ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String) As Integer
Dim strSql As String, rsRecordset As New ADODB.Recordset

On Error GoTo ErrHand

   strSql = "select cp01,cp02,cp03,cp04,cp10 from caseprogress where cp09=" + CNULL(strRecieve)
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      strCode0 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
      strCode1 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
      strCode2 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
      strCode3 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
      If Left(strRecieve, 1) = 接洽記錄單 Or Left(strRecieve, 1) = 內部收文 Then
         If rsRecordset.Fields(4) = 顧問聘任 Then
            oCheckRecieveCode = 2
         Else
            oCheckRecieveCode = 1
         End If
      Else
         oCheckRecieveCode = 1
      End If
   Else
      ShowMsg MsgText(9161)
   End If
   rsRecordset.Close
   Exit Function
ErrHand:
   ErrorMsg
End Function

'檢查本所案號是否存在於案件基本檔
'Copy from clsPublicData by Morgan 2004/2/17
Private Function oCheckCaseCodeIsExist(ByRef cp01 As String, ByRef cp02 As String, ByRef cp03 As String, ByRef cp04 As String, Optional strCaseName1 As String, Optional strCaseName2 As String, Optional strCaseName3 As String, Optional strCustomer As String, Optional strNation As String, Optional strNumber1 As String, Optional strNumber2 As String, Optional bolMsgShow As Boolean = True) As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset, intCaseKind As Integer

On Error GoTo ErrHand
    strSql = "select pa05,pa06,pa07," & _
       "NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),pa09," & _
       "pa22,pa11 from patent,customer where pa01=" & CNULL(cp01) & " and " & _
       "pa02=" & CNULL(cp02) & " and pa03=" & CNULL(cp03) & " and " & _
       "pa04=" & CNULL(cp04) & " and substr(pa26,1,8)=cu01(+) and " & _
       "substr(pa26,9,1)=cu02(+)"
      
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      If IsNull(rsRecordset.Fields(0)) Then
         strCaseName1 = ""
      Else
         strCaseName1 = rsRecordset.Fields(0)
      End If
      If IsNull(rsRecordset.Fields(1)) Then
         strCaseName2 = ""
      Else
         strCaseName2 = rsRecordset.Fields(1)
      End If
      If IsNull(rsRecordset.Fields(2)) Then
         strCaseName3 = ""
      Else
         strCaseName3 = rsRecordset.Fields(2)
      End If
      
      strCustomer = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
      strNation = IIf(IsNull(rsRecordset.Fields(4)), "", rsRecordset.Fields(4))
      strNumber1 = IIf(IsNull(rsRecordset.Fields(5)), "", rsRecordset.Fields(5))
      strNumber2 = IIf(IsNull(rsRecordset.Fields(6)), "", rsRecordset.Fields(6))
      oCheckCaseCodeIsExist = True
   Else
      If bolMsgShow Then ShowMsg MsgText(9141)
   End If
   rsRecordset.Close

Exit Function
ErrHand:
   ErrorMsg
End Function

Private Sub txtCode_Change(Index As Integer)
   'If cboCaseName.ListCount > 0 Then cboCaseName.Clear
   'If grdDataList.Rows > 1 Then GetSendCaseData True
End Sub

Private Sub txtTFCode_Change(Index As Integer)
   'If cboCaseName.ListCount > 0 Then cboCaseName.Clear
   'If grdDataList.Rows > 1 Then GetSendCaseData True
End Sub

Private Sub txtTFCode_GotFocus(Index As Integer)
   txtTFCode(Index).SelStart = 0
   txtTFCode(Index).SelLength = Len(txtTFCode(Index).Text)
End Sub

Private Sub txtTFCode_LostFocus(Index As Integer)
   If CheckKeyIn1(Index) = -1 Then
      txtTFCode(0).SetFocus
   End If
End Sub

Private Function CheckKeyIn1(ByRef intIndex As Integer) As Integer
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String

   CheckKeyIn1 = -1
   If Len(txtTFCode(intIndex)) > 0 And Len(txtTFCode(intIndex)) < txtTFCode(intIndex).MaxLength Then
      ShowMsg MsgText(1509)
   ElseIf intIndex = 3 Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
            IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3) Then
      If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
            IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3) Then
            
         SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
         CheckKeyIn1 = 1
      End If
   Else
      CheckKeyIn1 = 1
   End If
End Function

Private Sub txtCode_GotFocus(Index As Integer)
   txtCode(Index).SelStart = 0
   txtCode(Index).SelLength = Len(txtCode(Index).Text)
End Sub

Private Function CheckKeyIn2(ByRef intIndex As Integer) As Integer
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String

   CheckKeyIn2 = -1
   If Len(txtCode(intIndex)) > 0 And Len(txtCode(intIndex)) < txtCode(intIndex).MaxLength Then
      ShowMsg MsgText(1509)
   ElseIf intIndex = 2 Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
           IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3) Then
      If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
           IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3) Then
           
         SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
         CheckKeyIn2 = 1
      End If
   Else
      CheckKeyIn2 = 1
   End If
End Function

Private Sub txtReceiveCode_GotFocus()
   txtReceiveCode.SelStart = 0
   txtReceiveCode.SelLength = Len(txtReceiveCode.Text)
End Sub

Private Sub grdDataList_GotFocus()
   GridGotFocus grdDataList
End Sub

Private Sub grdDataList_LostFocus()
   GridLostFocus grdDataList
End Sub

Private Sub grdDataList_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then grdDataList_DblClick
End Sub

Private Sub grdDataList_RowColChange()
   If intLastRow <> grdDataList.row Then
      If blnOKtoShow Then
         blnOKtoShow = False
         ShowBar grdDataList, intLastRow, 6
         blnOKtoShow = True
      End If
   End If
End Sub

'Remove by Morgan 2007/10/16 不再使用
''儲存下一程序資料
'Public Sub InsertNextProgress()
'Dim adoRecord As ADODB.Recordset, strCounter As String
'Dim varConfirm As Variant, lngDate As Long
'
'   'edit by nickc 2007/02/02 不用 dll 了
'   'If objPublicData.ReadCaseFee("*", txtSystem, grdDataList.TextMatrix(grdDataList.Row, 8), grdDataList.TextMatrix(grdDataList.Row, 7), adoRecord) = True Then
'   If ClsPDReadCaseFee("*", txtSystem, grdDataList.TextMatrix(grdDataList.Row, 8), grdDataList.TextMatrix(grdDataList.Row, 7), adoRecord) = True Then
'      If IsNull(adoRecord.Fields("cf23").Value) = False Then
'         lngDate = Val(Format(CDate(grdDataList.TextMatrix(grdDataList.Row, 0)) + Val(adoRecord.Fields("cf23").Value), "YYYYMMDD"))
'         'edit by nickc 2007/02/02 不用 dll 了
'         'strCounter = objPublicData.GetNextProgressNo
'         'varConfirm = objPublicData.SaveNextProgress("'" & grdDataList.TextMatrix(grdDataList.Row, 5) & "', '" & txtSystem & "', '" & txtCode(0) & "', '" & txtCode(1) & "', '" & txtCode(2) & "', null, " & 收達 & ", " & lngDate & ", " & lngDate & ", '" & grdDataList.TextMatrix(grdDataList.Row, 10) & "', null, null, null, null, null, null, null, null, null, null, null, '" & strCounter & "'", 1)
'         strCounter = GetNextProgressNo
'         varConfirm = ClsPDSaveNextProgress("'" & grdDataList.TextMatrix(grdDataList.Row, 5) & "', '" & txtSystem & "', '" & txtCode(0) & "', '" & txtCode(1) & "', '" & txtCode(2) & "', null, " & 收達 & ", " & lngDate & ", " & lngDate & ", '" & grdDataList.TextMatrix(grdDataList.Row, 10) & "', null, null, null, null, null, null, null, null, null, null, null, '" & strCounter & "'", 1)
'      End If
'   End If
'End Sub

' 90.07.12 modify by louis (重新查詢)
Public Sub ReQuery()
   'Add By Cheng 2003/03/26
   If Me.optChoose(0).Value = True Then
       Me.optChoose(0).Value = False
       Me.optChoose(1).Value = True
   End If
   cmdok_Click 2
End Sub

' 90.07.12 modify by louis (清除畫面)
Public Sub Clear()
   txtReceiveCode = Empty
   txtSystem = Empty
   txtCode(0) = Empty
   txtCode(1) = Empty
   txtCode(2) = Empty
   cboCaseName.Clear
   ' 不在此做 (非Active Form在執行時會觸發 optChoose_Click 事件, 其中SetFocus會失敗
   'optChoose(0).Value = True
   'optChoose(1).Value = False
   grdDataList.Clear
   'Add By Cheng 2002/03/08
   '當執行完發文後, 回到此畫面時, 將本所案號設為預設值
   m_blnExcClear = True
   GridHead 'Add By Sindy 2019/7/26
End Sub

'Add By Cheng 2003/03/26
Private Sub GetOurCaseNo(strCP09 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
   StrSQLa = "Select * From CaseProgress Where CP09='" & strCP09 & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       If rsA("CP01").Value = 馬德里案 Then
           Me.txtSystem.Text = "" & rsA("CP01").Value
           Me.txtCode(0).Text = "" & rsA("CP02").Value
           Me.txtCode(1).Text = "" & Mid(rsA("CP03").Value, 1, 5)
           Me.txtCode(2).Text = "" & Mid(rsA("CP03").Value, 6, 1)
           Me.txtCode(3).Text = "" & rsA("CP04").Value
       Else
           Me.txtSystem.Text = "" & rsA("CP01").Value
           Me.txtCode(0).Text = "" & rsA("CP02").Value
           Me.txtCode(1).Text = "" & rsA("CP03").Value
           Me.txtCode(2).Text = "" & rsA("CP04").Value
       End If
   Else
       Me.txtSystem.Text = ""
       Me.txtCode(0).Text = ""
       Me.txtCode(1).Text = ""
       Me.txtCode(2).Text = ""
       Me.txtCode(3).Text = ""
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Sub

'2009/4/20 add by sonia
'美專母案領證發文需檢查CIP,CA或分割或CPA(但限設計)案未發文或已發文未提申則母案不可發文
Public Function ChkChild(strCP09 As String) As Boolean
Dim StrSQLa As String
Dim rsA  As New ADODB.Recordset

   StrSQLa = "SELECT C2.CP27,C2.CP47 FROM CASEPROGRESS C1,CASEPROGRESS C2,PATENT WHERE C1.CP09='" & strCP09 & "' AND C1.CP10='601' " & _
            "AND C1.CP01=PA01 AND C1.CP02=PA02 AND C1.CP03=PA03 AND C1.CP04=PA04 AND '101'=PA09 " & _
            "AND C1.CP01=C2.CP01 AND C1.CP02=C2.CP02 AND C2.CP03<>'0' AND C2.CP10 IN ('113','122') AND C2.CP57 IS NULL UNION " & _
            "SELECT C2.CP27,C2.CP47 FROM CASEPROGRESS C1,CASEPROGRESS C2,PATENT P1,PATENT P2 WHERE C1.CP09='" & strCP09 & "' AND C1.CP10='601' " & _
            "AND C1.CP01=P1.PA01 AND C1.CP02=P1.PA02 AND C1.CP03=P1.PA03 AND C1.CP04=P1.PA04 AND '101'=P1.PA09 " & _
            "AND C1.CP01=C2.CP01 AND C1.CP02=C2.CP02 AND C2.CP03<>'0' AND C2.CP10='114' AND C2.CP57 IS NULL " & _
            "AND C2.CP01=P2.PA01 AND C2.CP02=P2.PA02 AND C2.CP03=P2.PA03 AND C2.CP04=P2.PA04 AND '3'=P2.PA08 UNION " & _
            "SELECT C2.CP27,C2.CP47 FROM CASEPROGRESS C1,CASEPROGRESS C2,PATENT,DivisionCase WHERE C1.CP09='" & strCP09 & "' AND C1.CP10='601' " & _
            "AND C1.CP01=PA01 AND C1.CP02=PA02 AND C1.CP03=PA03 AND C1.CP04=PA04 AND '101'=PA09 " & _
            "AND C1.CP01=DC05 AND C1.CP02=DC06 AND C1.CP03=DC07 AND C1.CP04=DC08 " & _
            "AND DC01=C2.CP01 AND DC02=C2.CP02 AND DC03=C2.CP03 AND DC04=C2.CP04 AND C2.CP10='307' AND C2.CP57 IS NULL "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If IsNull(rsA.Fields(0)) Or IsNull(rsA.Fields(1)) Then
         ChkChild = True
      Else
         ChkChild = False
      End If
   Else
      ChkChild = False
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   If ChkChild = True Then
       MsgBox "本案之接續案或分割案尚未發文或未提申，母案領證不可發文!!!", vbExclamation + vbOKOnly
   End If
End Function
'2009/4/20 end

'Copy from bas088 by Morgan 2009/7/24
'intComeFrom=0為要秀出EMail    =1時直接進入發文
Public Function Where020102ToGo(Optional intComeFrom As Integer = 0) As Boolean
Dim strCP10 As String, strCP09 As String, cp(4) As String
Dim strPA09 As String
   
   strCP09 = grdDataList.TextMatrix(Me.grdDataList.row, 5)
   strCP10 = grdDataList.TextMatrix(grdDataList.row, 7)
   cp(1) = txtSystem
   cp(2) = txtCode(0)
   cp(3) = Right("0" & txtCode(1), 1)
   cp(4) = Right("00" & txtCode(1), 2)
   
   'Add By Sindy 2013/11/14
   '檢查是否有承辦歷程是否有產生承辦單可以發文
   'Modify By Sindy 2023/2/3 柏翰:CFP請改回跟P一樣發文時要檢查歷程
   If PUB_IsEmpFlowIsSend(strCP09) = False Then
      Where020102ToGo = False
      Me.Show
      Exit Function
   End If
   
   'Add by Morgan 2010/3/17
   If strCP10 = "605" Or strCP10 = "606" Or strCP10 = "607" Then
      If PUB_ChkNPExist(cp, strCP10) Then
         ClsPDGetCaseProperty cp(1), strCP10, strExc(1)
         MsgBox "本案下一程序有<" & strExc(1) & ">期限，不可發文!!!", vbExclamation + vbOKOnly
         Where020102ToGo = False
         Me.Show
         Exit Function
      End If
   End If
   
   'Add By Sindy 2018/1/8
   If m_strIR01 <> "" Then
      If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtSystem & txtCode(0) & txtCode(1) & txtCode(2) Then
         MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
         Exit Function
      End If
   End If
   '2018/1/8 END
            
   Where020102ToGo = True
'Remove by Lydia 2018/08/30  (應收帳款管控)取消預定收款日,改成付款週期=>不發email
'   If intComeFrom = 0 Then
'      If grdDataList.TextMatrix(grdDataList.row, 6) <> "" And grdDataList.TextMatrix(grdDataList.row, 6) <> "0" Then
'         If PUB_ChkPaidByCP09(strCP09) = False Then    'Added by Morgan 2016/8/23 出納繳款確認後就可送件
'            'frm020102_K.Show
'            frm050102_b.Show
'            Exit Function
'         End If 'Added by Morgan 2016/8/22
'      End If
'   End If
'end 2018/08/30

   'Add by Morgan 2011/4/20
   'EPC實審指定費(回覆檢索報告收文檢查)
   strPA09 = grdDataList.TextMatrix(Me.grdDataList.row, 8)
   If strPA09 = "221" And (strCP10 = "416" Or strCP10 = "215") Then
      strExc(0) = "select cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "'" & _
         " and cp03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP10='218'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         If MsgBox("本案尚未收到回覆檢索報告，" & grdDataList.TextMatrix(Me.grdDataList.row, 1) & "應不可發文！是否仍要繼續發文？", vbYesNo + vbDefaultButton2) = vbNo Then
            Me.Show
            Exit Function
         End If
      End If
   End If
   
   'Add by Morgan 2009/7/22 歐盟設計發文需控制其他多國都已有申請日
   If PUB_Chk103in239(Me.grdDataList.TextMatrix(Me.grdDataList.row, 5)) = True Then
      Where020102ToGo = False
      Me.Show
      Exit Function
   End If
         
   If grdDataList.TextMatrix(grdDataList.row, 9) = 專利 Then
      Select Case strCP10
         Case 延期
            '延期記錄資料來源為下一程序檔
            frm050102_2.m_str_DL05 = "2"
            frm050102_2.intWhereComeFrom = 2
            frm050102_2.Show
         'Add by Morgan 2006/8/14 加122CA申請
         'Modify by Morgan 2009/7/29 +109 PCT申請
         'Modified by Morgan 2025/3/5 +125 衍生設計申請
         Case 發明申請, 新型申請, 設計申請, 聯合申請, CIP申請, CPA申請, 再發行, 美國暫時申請, 分割, "122", "109", "125"
            'Added by Morgan 2021/12/10
            '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
            If PUB_CheckFormExist("frm050102_3") = False Then
               Set frm050102_3 = Nothing
            End If
            'end 2021/12/10
            frm050102_3.Show
         Case 變更
            'Added by Morgan 2021/12/10
            '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
            If PUB_CheckFormExist("frm050102_4") = False Then
               Set frm050102_4 = Nothing
            End If
            'end 2021/12/10
            frm050102_4.Show
         'Modify by Morgan 2007/7/27 加"繼承"
         'Modified by Morgan 2016/3/3 +126 期末拋棄
         'Modified by Lydia 2016/08/26 +438 再考量試行計畫(AFCP 2.0)
         Case 實體審查, 答辯, "126", 修正, 主動修正, 提供前案資料, 選取, 讓與, "214", "427", 繼承, "438"
            'Added by Morgan 2021/12/10
            '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
            If PUB_CheckFormExist("frm050102_5") = False Then
               Set frm050102_5 = Nothing
            End If
            'end 2021/12/10
            frm050102_5.Show
         Case 補文件
            'Added by Morgan 2021/12/10
            '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
            If PUB_CheckFormExist("frm050102_6") = False Then
               Set frm050102_6 = Nothing
            End If
            'end 2021/12/10
            frm050102_6.Show
         Case 申請優先權證明
            'Added by Morgan 2021/12/10
            '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
            If PUB_CheckFormExist("frm050102_7") = False Then
               Set frm050102_7 = Nothing
            End If
            'end 2021/12/10
            frm050102_7.Show
         Case 領證及繳年費
            'Added by Morgan 2021/12/10
            '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
            If PUB_CheckFormExist("frm050102_8") = False Then
               Set frm050102_8 = Nothing
            End If
            'end 2021/12/10
            frm050102_8.Show
         'Modify by Amy 2018/03/20 +612 年費移作次年
         Case 年費, 維持費, 延展費, "612"
            'Added by Morgan 2021/12/10
            '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
            If PUB_CheckFormExist("frm050102_9") = False Then
               Set frm050102_9 = Nothing
            End If
            'end 2021/12/10
            frm050102_9.Show
         Case 授權
            'Added by Morgan 2021/12/10
            '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
            If PUB_CheckFormExist("frm050102_a") = False Then
               Set frm050102_a = Nothing
            End If
            'end 2021/12/10
            frm050102_a.Show
         Case Else
            'Added by Morgan 2021/12/10
            '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
            If PUB_CheckFormExist("frm050102_6") = False Then
               Set frm050102_6 = Nothing
            End If
            'end 2021/12/10
            frm050102_6.Show
      End Select
   Else
      'Added by Morgan 2022/12/29
      '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
      If PUB_CheckFormExist("frm050102_6") = False Then
         Set frm050102_6 = Nothing
      End If
      frm050102_6.Show
      
   End If
End Function

'2011/5/18 add by sonia 比利時209及西班牙211發明申請若未收文申請檢索報告則提醒操作者,不可發文
Public Function Chk421rec(strCP09 As String) As Boolean
Dim StrSQLa As String
Dim rsA  As New ADODB.Recordset
   
   Chk421rec = False
   StrSQLa = "Select c2.cp09,c2.cp57 From caseProgress c1,caseprogress c2,patent WHERE c1.cp09='" & strCP09 & "' " & _
            " and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and c2.cp10(+)='421' " & _
            " and c1.cp01=pa01 and c1.cp02=pa02 and c1.cp03=pa03 and c1.cp04=pa04"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If Not IsNull(rsA.Fields(0)) Then
         If Not IsNull(rsA.Fields(1)) Then
            Chk421rec = True
         End If
      Else
         Chk421rec = True
      End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   If Chk421rec = True Then
       MsgBox "請提醒智權同仁：比利時或西班牙發明申請必須收文申請檢索報告才可發文！", vbExclamation + vbOKOnly
   End If
End Function
'2011/5/18 end

'Added by Morgan 2023/2/14 檢查美國申請案(發明、設計、CIP、CPA)有申請人非個人且未收文讓渡,不可發文。(分割、暫時申請不用)
'Modified by Morgan 2023/2/20 改條件:申請人非個人-->申請人與發明人不完全相同
Private Function ChkUsNeed701(ByVal pCP09 As String) As Boolean
   Dim stSQL As String, intQ As Integer, stName As String, intApps As Integer
   Dim RsQ  As ADODB.Recordset
   
   'stSQL = "select * From caseprogress a,patent" & _
      " where cp09='" & pCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and exists(select * from customer where instr(pa26||pa27||pa28||pa29||pa30,cu01||cu02)>0 and cu15<>'0')" & _
      " and not exists(select * from caseprogress b where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='701')"
   stSQL = "select pa26,pa27,pa28,pa29,pa30,pi06,in04,c1.cu04 pa26c,c2.cu04 pa27c,c3.cu04 pa28c,c4.cu04 pa29c,c5.cu04 pa30c" & _
      " From caseprogress a,patent,customer c1,customer c2,customer c3,customer c4,customer c5,patentinventor,inventor" & _
      " where cp09='" & pCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and not exists(select * from caseprogress b where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='701')" & _
      " and c1.cu01(+)=substr(pa26,1,8) and c1.cu02(+)=substr(pa26,9)" & _
      " and c2.cu01(+)=substr(pa27,1,8) and c2.cu02(+)=substr(pa27,9)" & _
      " and c3.cu01(+)=substr(pa28,1,8) and c3.cu02(+)=substr(pa28,9)" & _
      " and c4.cu01(+)=substr(pa29,1,8) and c4.cu02(+)=substr(pa29,9)" & _
      " and c5.cu01(+)=substr(pa30,1,8) and c5.cu02(+)=substr(pa30,9)" & _
      " and pi01(+)=pa01 and pi02(+)=pa02 and pi03(+)=pa03 and pi04(+)=pa04" & _
      " and in01(+)=substr(pi06,1,8) and in02(+)=substr(pi06,9) order by pi05"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      intApps = 0
      For intQ = 1 To 5
         stName = "" & RsQ.Fields("pa" & (25 + intQ) & "c")
         If stName <> "" Then
            intApps = intApps + 1
            RsQ.Find "in04='" & RsQ.Fields("pa" & (25 + intQ) & "c") & "'", , , 1
            If RsQ.EOF Then '申請人非發明人
               ChkUsNeed701 = True
               Exit For
            End If
         Else
            Exit For
         End If
      Next
   End If
   If intApps <> RsQ.RecordCount Then '人數不同
      ChkUsNeed701 = True
   End If
   Set RsQ = Nothing
End Function
