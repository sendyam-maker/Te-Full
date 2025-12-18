VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090222 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件催審延緩維護"
   ClientHeight    =   4600
   ClientLeft      =   900
   ClientTop       =   1060
   ClientWidth     =   8420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4600
   ScaleWidth      =   8420
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   5160
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1268
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "更新催審期限"
      Height          =   285
      Left            =   6150
      TabIndex        =   7
      Top             =   1268
      Width           =   1440
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "案件進度(&C)"
      Height          =   380
      Index           =   1
      Left            =   5640
      TabIndex        =   9
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "尋找(&S)"
      Default         =   -1  'True
      Height          =   285
      Index           =   0
      Left            =   3140
      TabIndex        =   4
      Top             =   570
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   380
      Index           =   2
      Left            =   6870
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox textCP 
      Height          =   285
      Index           =   4
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   3
      Top             =   570
      Width           =   375
   End
   Begin VB.TextBox textCP 
      Height          =   285
      Index           =   3
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   2
      Top             =   570
      Width           =   240
   End
   Begin VB.TextBox textCP 
      Height          =   285
      Index           =   2
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   1
      Top             =   570
      Width           =   855
   End
   Begin VB.TextBox textCP 
      Height          =   285
      Index           =   1
      Left            =   1215
      MaxLength       =   3
      TabIndex        =   0
      Top             =   570
      Width           =   480
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm090222.frx":0000
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   8180
      _ExtentX        =   14429
      _ExtentY        =   3193
      _Version        =   393216
      FixedCols       =   0
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
      _Band(0).Cols   =   2
   End
   Begin MSForms.TextBox Text1 
      Height          =   330
      Left            =   1200
      TabIndex        =   6
      Top             =   2010
      Width           =   5985
      VariousPropertyBits=   671105051
      MaxLength       =   200
      Size            =   "10557;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblT 
      Caption         =   "延緩原因："
      Height          =   210
      Left            =   240
      TabIndex        =   22
      Top             =   2040
      Width           =   975
   End
   Begin MSForms.Label lblC 
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   21
      Top             =   1680
      Width           =   1740
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3069;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblC 
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1085;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblC 
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   19
      Top             =   1320
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblC 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1085;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblC 
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   17
      Top             =   960
      Width           =   6630
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "11695;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "下次催審日期："
      Height          =   210
      Index           =   3
      Left            =   3840
      TabIndex        =   16
      Top             =   1305
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "催審期限："
      Height          =   210
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   2360
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "申請人："
      Height          =   210
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   920
   End
   Begin VB.Label Label3 
      Caption         =   "申請國家："
      Height          =   210
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   920
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   210
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   920
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   210
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frm090222"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; lblC(index)、Text1、GRD1改字型=新細明體-ExtB
'Create by Lydia 2016/01/07 案件催審延緩維護
Option Explicit

Dim m_PrevForm As Form '前一畫面
'本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
'-------------
Dim colNP01 As Integer
Dim colNP22 As Integer
Dim colNo As Integer
Dim dblPrevRow As Integer

'Added by Lydia 2016/04/01 提供給案件資料及案件進度查詢(frm100101_2)的下一筆呼叫
Public Sub PubShowNextData()
End Sub
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0 '尋找
         QueryData
      Case 1 '案件進度
         If textCP(1) & textCP(2) <> "" Then
            If textCP(3) = "" Then textCP(3) = "0"
            If textCP(4) = "" Then textCP(4) = "00"
            Me.Enabled = False
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            frm100101_2.Tag = Pub_RplStr(textCP(1) & "-" & IIf(textCP(2) = "", "000000", textCP(2)) & "-" & textCP(3) & "-" & textCP(4))
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
         End If
      Case 2 '結束
         Unload Me
   End Select
End Sub

Private Sub cmdUpdate_Click()
Dim oText As TextBox
Dim Cancel As Boolean

   Cancel = False
   For Each oText In textCP
       If oText = "" Then
          MsgBox "本所案號不可空白!", vbCritical
          Exit Sub
       End If
       textCP_Validate oText.Index, Cancel
       If Cancel Then
          Exit Sub
       End If
   Next
   
   '下次催審日期
   If txtDate = "" Then
       MsgBox "下次催審日期不可空白!", vbCritical
       txtDate.SetFocus
       Exit Sub
   End If
   txtDate_Validate Cancel
   If Cancel Then Exit Sub
   
   If dblPrevRow = 0 Then
      MsgBox "未選取催審期限記錄!", vbCritical
      Exit Sub
   ElseIf GRD1.TextMatrix(dblPrevRow, colNo) = "" Or GRD1.TextMatrix(dblPrevRow, 0) <> "V" Then
      MsgBox "未選取催審期限記錄!", vbCritical
      Exit Sub
   ElseIf GRD1.TextMatrix(dblPrevRow, colNo) <> textCP(1) & textCP(2) & textCP(3) & textCP(4) Then
      MsgBox "本所案號與催審期限記錄不一致!", vbCritical
      Exit Sub
   'Added by Lydia 2016/04/14 增加原因
   ElseIf Len(Trim(Text1.Text)) < 2 Then
      MsgBox "請輸入延緩原因!", vbCritical
      Text1.SetFocus
      Exit Sub
   Else
      'Added by Lydia 2021/12/21 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me, , True, "TextBox") = False Then
          Exit Sub
      End If
      'end 2021/12/21
      If OnSaveData Then
         'Added by Lydia 2016/04/14 通知主管
         strExc(1) = textCP(1) & "-" & textCP(2) & IIf(textCP(3) & textCP(4) <> "000", "-" & textCP(3) & "-" & textCP(4), "") & " 延緩催審"
         strExc(2) = "本所案號：" & textCP(1) & "-" & textCP(2) & "-" & textCP(3) & "-" & textCP(4) & vbCrLf & _
                    "案件名稱：" & lblC(0).Caption & vbCrLf & _
                    "申請國家：" & lblC(2).Caption & vbCrLf & _
                    "總收文號：" & GRD1.TextMatrix(dblPrevRow, 1) & vbCrLf & _
                    "案件性質：" & GRD1.TextMatrix(dblPrevRow, 3) & vbCrLf & _
                    "下次催審期限：" & ChangeTStringToTDateString(txtDate) & vbCrLf & _
                    "延緩原因：" & Trim(Text1)
         'Added by Lydia 2023/09/04 加入FCT商爭C類的來函性質;參考basQuery.PUB_GetTMdebate
         strExc(0) = ""
         If lblC(1) <= "010" Then
            'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
            strSql = "select sg03 from staff_group where sg02='FCT' and sg01='C1'" & _
                      " and sg03 not in('204','205','207','303','305','306','307','310','614','615','706','722','1005','1202','1203','1204','1401','1607','1611','1614','1615'," & FCT_NotTMdebate & ")"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strExc(0) = RsTemp.GetString(adClipString, , , ",")
            End If
         End If
         'end 2023/09/04
         
         '台灣案非商申(=商爭)或大陸案，通知林經理和承慧(86048)
         'Modified by Lydia 2023/09/04 加入FCT商爭C類的來函性質
         'If lblC(1) > "010" Or (lblC(1) <= "010" And InStr(TMdebate, grd1.TextMatrix(dblPrevRow, 2))) > 0 Then
         If lblC(1) > "010" Or (lblC(1) <= "010" And InStr(strExc(0) & ",", GRD1.TextMatrix(dblPrevRow, 2))) > 0 Then
            'Modified by Lydia 2021/11/12 取消69008林經理不增加人
            PUB_SendMail strUserNum, "86048", "", strExc(1), strExc(2)
         '台灣案商申，通知林經理(69008)和嘉雯(84027)
         Else
            PUB_SendMail strUserNum, "84027", "", strExc(1), strExc(2)
         End If
         'end 2016/04/14
         
         QueryData
      Else
         MsgBox "更新失敗", vbCritical
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   FormClear
   SetGrd
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090222 = Nothing
   If TypeName(m_PrevForm) <> "Nothing" Then
      If m_PrevForm.textCP(1) <> "" And m_PrevForm.textCP(2) <> "" Then
         m_PrevForm.QueryData
      End If
      m_PrevForm.Show
      Set m_PrevForm = Nothing
   End If
End Sub

Public Sub QueryData()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim strTemp As String
   
   If textCP(1) = "" Or textCP(2) = "" Then
      MsgBox "請輸入案號!!!", vbExclamation + vbOKOnly
      If textCP(1) = "" Then
         Me.textCP(1).SetFocus
      ElseIf textCP(2) = "" Then
         Me.textCP(2).SetFocus
      End If
      Exit Sub
   End If

   If textCP(3) = "" Then textCP(3) = "0"
   If textCP(4) = "" Then textCP(4) = "00"
   dblPrevRow = 0
   
   m_TM01 = Trim(textCP(1))
   m_TM02 = Trim(textCP(2))
   m_TM03 = Trim(textCP(3))
   m_TM04 = Trim(textCP(4))
   If ClsPDCheckCaseCodeIsExist(m_TM01, m_TM02, m_TM03, m_TM04, strExc(1), strExc(2), strExc(3), strExc(4), strExc(5)) Then
      lblC(0).Caption = IIf(strExc(1) <> "", strExc(1), IIf(strExc(2) <> "", strExc(2), strExc(3)))
      lblC(4).Caption = strExc(4)
      lblC(1).Caption = strExc(5)
      If ClsPDGetNation(strExc(5), strExc(6)) Then
         lblC(2).Caption = strExc(6)
      End If
   Else
      FormClear
      Exit Sub
   End If
   'Added by Lydia 2016/04/14 +NP15
   strSql = "SELECT '' V,NP01,CP10,NVL(C2.CPM03,CP10) AS CP10M,NVL(S1.ST02,NP10) AS NP10,NVL(CP27 - 19110000, NULL) AS CP27," & _
            "NVL(NP08 - 19110000, NULL) AS NP08,NP15,NP22,NVL(S2.ST02,CP14) AS CP14,CP01||CP02||CP03||CP04 CASENO " & _
            "FROM NEXTPROGRESS, CASEPROGRESS C1, CASEPROPERTYMAP C2, STAFF S1, STAFF S2 " & _
            "WHERE NP02='" & m_TM01 & "' AND NP03='" & m_TM02 & "' AND NP04='" & m_TM03 & "' AND NP05='" & m_TM04 & "' " & _
            "AND NP06 IS NULL AND NP07='305' AND NP01 = C1.CP09(+) AND NP10 = S1.ST01(+) AND CP14 = S2.ST01(+) AND CP01 = C2.CPM01(+) AND CP10 = C2.CPM02(+) " & _
            "ORDER BY NP08,NP01,CP27 "
   '非台灣
   If lblC(1) > "010" Then
      strSql = Replace(strSql, "C2.CPM03", "C2.CPM04")
   End If
   
   Call SetGrd 'Added by Lydia 2018/02/12
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      'Modified by Lydia 2018/02/12
      'SetGrd (rsTmp.RecordCount + 1)
      Call SetGrd(False)
      
      'Added by Lydia 2016/04/14 只有一筆,預設勾選
      If rsTmp.RecordCount = 1 Then
         GRD1.TextMatrix(1, 0) = "V"
         dblPrevRow = 1
      End If
      'txtDate.SetFocus
      Text1.SetFocus
   Else
      MsgBox "無下一程序催審的記錄!!!", vbExclamation + vbOKOnly
      Me.textCP(2).SetFocus
      rsTmp.Close
      Set rsTmp = Nothing
      'SetGrd 'Remove by Lydia 2018/02/12
      Exit Sub
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

'Modified by Lydia 2018/02/12
'Private Sub SetGrd(Optional ByVal iR As Integer = 2)
Private Sub SetGrd(Optional ByVal pReset As Boolean = True)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

   'v, NP01,CP10,CP10M,NP10,CP27,NP08,NP15,NP22,CP14,CASENO
   arrGridHeadText = Array("V", "總收文號", "CP10", "案件性質", "NP10", "發文日", "催審期限", "備註", "NP22", "CP14", "CASENO")
   'Modified by Lydia 2016/05/20 加寬備註2200->4500
   arrGridHeadWidth = Array(200, 1100, 0, 1200, 0, 900, 900, 4500, 0, 0, 0)

   GRD1.Visible = False
   With GRD1
        .Cols = UBound(arrGridHeadText) + 1
        'Modified by Lydia 2018/02/12 預設先清空
        '.Rows = iR
        If pReset = True Then
             GRD1.Clear
             GRD1.Rows = 2
        End If
        'end 2018/02/12
        For iRow = 0 To .Cols - 1
           .row = 0
           .col = iRow
           .Text = arrGridHeadText(iRow)
           .ColWidth(iRow) = arrGridHeadWidth(iRow)
        Next
   End With
   
   If colNo = 0 Then colNo = PUB_MGridGetId("CASENO", GRD1)
   If colNP01 = 0 Then colNP01 = PUB_MGridGetId("總收文號", GRD1)
   If colNP22 = 0 Then colNP22 = PUB_MGridGetId("NP22", GRD1)

   GRD1.Visible = True

End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol >= 0 Then GRD1.col = nCol
   If nRow >= 0 Then GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim TmpRow As Integer
Dim jj As Integer

   TmpRow = GRD1.MouseRow
   
   GRD1.Visible = False
   If TmpRow > 0 Then
      '上一筆資料列清除反白
      If dblPrevRow > 0 Then
         GRD1.col = 0
         GRD1.row = dblPrevRow
         GRD1.Text = ""
         For jj = 0 To GRD1.Cols - 1
            GRD1.col = jj
            GRD1.CellBackColor = QBColor(15)
         Next jj
      End If
      '目前資料列反白
      GRD1.col = 0
      GRD1.row = TmpRow
      dblPrevRow = GRD1.row
       If GRD1.TextMatrix(GRD1.row, 1) <> "" Then
          GRD1.Text = "V"
          For jj = 0 To GRD1.Cols - 1
             GRD1.col = jj
             GRD1.CellBackColor = &HFFC0C0
          Next jj
       End If
   End If
   GRD1.Visible = True
   'txtDate.SetFocus   'add by sonia 2016/3/31
   Text1.SetFocus
End Sub

Private Sub FormClear()
Dim oLabel As Object
Dim oText As TextBox

   For Each oText In textCP
      oText.Text = ""
   Next
   
   txtDate = ""
   dblPrevRow = 0
   For Each oLabel In lblC
      oLabel.Caption = ""
   Next
   Text1.Text = ""
   
End Sub

Private Function OnSaveData() As Boolean
Dim strTmp As String
Dim m_NP01 As String, m_NP22 As String
   
On Error GoTo ErrorHandler

   OnSaveData = False
   
   m_NP01 = GRD1.TextMatrix(dblPrevRow, colNP01)
   m_NP22 = GRD1.TextMatrix(dblPrevRow, colNP22)
   'Added by Lydia 2016/04/14 + 原因
   strTmp = "UPDATE NEXTPROGRESS SET NP08=" & CNULL(DBDATE(txtDate), True) & ", NP09=" & CNULL(DBDATE(txtDate), True) & _
            " ,NP15=NP15||'" & ChangeTStringToTDateString(strSrvDate(2)) & " " & strUserName & " 延緩催審，原因:" & Trim(Text1.Text) & ";" & "' WHERE np01='" & m_NP01 & "' and np22=" & CNULL(m_NP22, True)
            
   cnnConnection.BeginTrans
      cnnConnection.Execute strTmp, intI
   cnnConnection.CommitTrans
    
   OnSaveData = True
   Exit Function
   
ErrorHandler:
    cnnConnection.RollbackTrans
    
End Function

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub textCP_GotFocus(Index As Integer)
   TextInverse textCP(Index)
   CloseIme
End Sub

Private Sub textCP_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP_Validate(Index As Integer, Cancel As Boolean)
   
   Select Case Index
       Case 1
           If textCP(Index).Text <> "" And InStr(textCP(Index), "T") = 0 Then
               MsgBox "限定使用商標案!!", vbCritical
               textCP(Index).SetFocus
               Cancel = True
           End If
   End Select
End Sub

Private Sub txtDate_GotFocus()
    TextInverse txtDate
    CloseIme
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
Dim strTmp As String

   If txtDate <> "" Then
      strTmp = txtDate
      If CheckIsTaiwanDate(strTmp) = False Then
         GoTo JumpCancel
      Else
         strTmp = ChangeWDateStringToWString(strTmp)
           If strTmp <= strSrvDate(1) Then
                MsgBox "下次催審日期必須大於系統日!", vbCritical
                GoTo JumpCancel
           '若輸入非工作日時自動改為前一工作日
           ElseIf ChkWorkDay(strTmp) = False Then
                  strTmp = CompWorkDay(-2, strTmp, 1)
                  txtDate.Text = ChangeWStringToTString(strTmp)
           End If
      End If
   End If
       
   Exit Sub
    
JumpCancel:
   txtDate.SetFocus
   Cancel = True
End Sub
