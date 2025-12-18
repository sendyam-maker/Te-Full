VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm110103_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "閉卷"
   ClientHeight    =   4128
   ClientLeft      =   1248
   ClientTop       =   792
   ClientWidth     =   8280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4128
   ScaleWidth      =   8280
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   1
      Left            =   4140
      MaxLength       =   1
      TabIndex        =   1
      Top             =   2670
      Width           =   405
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   0
      Left            =   1050
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2664
      Width           =   1092
   End
   Begin VB.ComboBox cboReason 
      Height          =   300
      Left            =   1050
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   3060
      Width           =   7125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&Q)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6150
      TabIndex        =   5
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5310
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7290
      TabIndex        =   6
      Top             =   120
      Width           =   800
   End
   Begin MSForms.ComboBox cboNote 
      Height          =   300
      Left            =   1050
      TabIndex        =   3
      Top             =   3480
      Width           =   7125
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12568;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件備註欄：不可銷卷案請加註 ""不銷卷"" 字樣！  與他案合併計算結餘請註明""與某案號合併計算結餘""！"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   115
      Left            =   30
      TabIndex        =   37
      Top             =   3840
      Width           =   8220
   End
   Begin VB.Label Label1 
      Caption         =   "後續准駁簡單報告：               (Y：核准以及C類來函簡單報告)"
      Height          =   180
      Index           =   0
      Left            =   2490
      TabIndex        =   36
      Top             =   2670
      Width           =   5000
   End
   Begin VB.Label Label21 
      Caption         =   "案件備註："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   3540
      Width           =   975
   End
   Begin MSForms.Label lblName 
      Height          =   180
      Left            =   1896
      TabIndex        =   20
      Top             =   696
      Width           =   3252
      VariousPropertyBits=   27
      Size            =   "3625;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNo 
      Height          =   180
      Left            =   936
      TabIndex        =   19
      Top             =   696
      Width           =   852
   End
   Begin VB.Label lblNumber 
      Height          =   180
      Index           =   10
      Left            =   1656
      TabIndex        =   34
      Top             =   2328
      Width           =   1092
   End
   Begin VB.Label lblNumber 
      Height          =   180
      Index           =   9
      Left            =   1656
      TabIndex        =   33
      Top             =   2004
      Width           =   1092
   End
   Begin VB.Label lblNumber 
      Height          =   180
      Index           =   8
      Left            =   5856
      TabIndex        =   32
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label lblNumber 
      Height          =   180
      Index           =   7
      Left            =   3696
      TabIndex        =   31
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label lblNumber 
      Height          =   180
      Index           =   6
      Left            =   1656
      TabIndex        =   30
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label lblNumber 
      Height          =   180
      Index           =   5
      Left            =   5856
      TabIndex        =   29
      Top             =   1368
      Width           =   1092
   End
   Begin VB.Label lblNumber 
      Height          =   180
      Index           =   4
      Left            =   3696
      TabIndex        =   28
      Top             =   1368
      Width           =   1092
   End
   Begin VB.Label lblNumber 
      Height          =   180
      Index           =   3
      Left            =   1656
      TabIndex        =   27
      Top             =   1368
      Width           =   1092
   End
   Begin VB.Label lblNumber 
      Height          =   180
      Index           =   2
      Left            =   5856
      TabIndex        =   26
      Top             =   1020
      Width           =   1092
   End
   Begin VB.Label lblNumber 
      Height          =   180
      Index           =   1
      Left            =   3696
      TabIndex        =   25
      Top             =   1020
      Width           =   1092
   End
   Begin VB.Label lblNumber 
      Height          =   180
      Index           =   0
      Left            =   1656
      TabIndex        =   24
      Top             =   1020
      Width           =   1092
   End
   Begin VB.Label Label5 
      Caption         =   "閉卷日期："
      Height          =   180
      Left            =   120
      TabIndex        =   23
      Top             =   2664
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "閉卷原因："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   3096
      Width           =   972
   End
   Begin VB.Label lblTitle 
      Caption         =   "申請人："
      Height          =   180
      Left            =   120
      TabIndex        =   21
      Top             =   696
      Width           =   972
   End
   Begin VB.Label Label16 
      Caption         =   "Total："
      Height          =   180
      Left            =   1032
      TabIndex        =   18
      Top             =   2328
      Width           =   612
   End
   Begin VB.Label Label15 
      Caption         =   "其他："
      Height          =   180
      Left            =   1032
      TabIndex        =   17
      Top             =   2004
      Width           =   612
   End
   Begin VB.Label Label14 
      Caption         =   "案件數："
      Height          =   180
      Left            =   120
      TabIndex        =   16
      Top             =   1020
      Width           =   852
   End
   Begin VB.Label Label13 
      Caption         =   "FCL："
      Height          =   180
      Left            =   5256
      TabIndex        =   15
      Top             =   1368
      Width           =   612
   End
   Begin VB.Label Label12 
      Caption         =   "L："
      Height          =   180
      Left            =   5256
      TabIndex        =   14
      Top             =   1680
      Width           =   612
   End
   Begin VB.Label Label11 
      Caption         =   "CFL："
      Height          =   180
      Left            =   5256
      TabIndex        =   13
      Top             =   1020
      Width           =   612
   End
   Begin VB.Label Label10 
      Caption         =   "FCT："
      Height          =   180
      Left            =   3096
      TabIndex        =   12
      Top             =   1368
      Width           =   612
   End
   Begin VB.Label Label8 
      Caption         =   "T："
      Height          =   180
      Left            =   3096
      TabIndex        =   11
      Top             =   1680
      Width           =   612
   End
   Begin VB.Label Label7 
      Caption         =   "CFT："
      Height          =   180
      Left            =   3096
      TabIndex        =   10
      Top             =   1020
      Width           =   612
   End
   Begin VB.Label Label4 
      Caption         =   "FCP："
      Height          =   180
      Left            =   1080
      TabIndex        =   9
      Top             =   1368
      Width           =   612
   End
   Begin VB.Label Label3 
      Caption         =   "P："
      Height          =   180
      Left            =   1080
      TabIndex        =   8
      Top             =   1680
      Width           =   612
   End
   Begin VB.Label Label2 
      Caption         =   "CFP："
      Height          =   180
      Left            =   1080
      TabIndex        =   7
      Top             =   1020
      Width           =   612
   End
End
Attribute VB_Name = "frm110103_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/5/10 改成Form2.0(lblName,cboNote)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'2010/8/3 日期欄已修改 by sonia
Option Explicit
'bolLeave判斷離開時，是否要彈出詢問視窗
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim bolLeave As Boolean, intLeaveKind As Integer
'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
'intWhere 國內,國外_CF,國外_FC
Dim intCaseKind As Integer, intWhere As Integer
'儲存解除期限原因的編號
Dim strReasonNo() As String
'strCaseCode上一畫面frm110103_2勾選的本所案號
'intTotalCaseCode上一畫面frm110103_2勾選的本所案號總數
Dim strCaseCode() As String, intTotalCaseCode As Integer
'edit by nickc 2007/02/05 不用 dll 了
'Dim obj011 As prjTaieDll011.cls011
Dim strSql As String, cp(1 To 4) As String, SCp(1 To 79) As String

Private Function SaveDatabase() As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String, strReceiveCode As String, i As Integer, varSaveCursor

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
On Error GoTo ErrHand
For i = 0 To intTotalCaseCode - 1
'edit by nickc 2007/02/02 不用 dll 了
'       If objPublicData.GetSystemKind(strCaseCode(0, i), intCaseKind, , intWhere) = False Then GoTo Err1
'       If objPublicData.GetReceiveCode(strCaseCode(0, i), strCaseCode(1, i), strCaseCode(2, i), strCaseCode(3, i), strReceiveCode) = False Then GoTo Err1
'       If objPublicData.ReadAllData(strReceiveCode, cp(), field(), intCaseKind, intWhere) = False Then GoTo Err1
       If ClsPDGetSystemKind(strCaseCode(0, i), intCaseKind, , intWhere) = False Then GoTo err1
       If ClsPDGetReceiveCode(strCaseCode(0, i), strCaseCode(1, i), strCaseCode(2, i), strCaseCode(3, i), strReceiveCode) = False Then GoTo err1
        ReDim cp(TF_CP) As String
        cp(9) = strReceiveCode
        If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) = False Then GoTo err1
        
       Select Case intCaseKind
                    Case 專利
                               field(57) = "Y"
                               field(58) = txtCaseField(0)
                               field(59) = strReasonNo(cboReason.ListIndex)
                    Case 商標
                               field(29) = "Y"
                               field(30) = txtCaseField(0)
                               field(31) = strReasonNo(cboReason.ListIndex)
                    Case 法務
                               field(8) = "Y"
                               field(9) = txtCaseField(0)
                               field(10) = strReasonNo(cboReason.ListIndex)
                    Case 顧問
                               field(9) = "Y"
                               field(10) = txtCaseField(0)
                               field(11) = strReasonNo(cboReason.ListIndex)
                    Case Else
                               field(15) = "Y"
                               field(16) = txtCaseField(0)
                               field(17) = strReasonNo(cboReason.ListIndex)
    End Select
    'edit by nickc 2007/02/05 不用 dll 了
    'If obj011.SaveCloseCaseData(intCaseKind, intWhere, cp(), field()) = False Then
    If Cls011SaveCloseCaseData(intCaseKind, intWhere, cp(), field()) = False Then
       Exit For
    End If
Next
If i = intTotalCaseCode Then
   SaveDatabase = True
Else
err1:
   ShowMsg MsgText(9004)
End If
Screen.MousePointer = varSaveCursor
Exit Function
ErrHand:
Screen.MousePointer = varSaveCursor
ErrorMsg
End Function
Private Sub cmdok_Click(Index As Integer)
Dim i As Integer, varSaveCursor, j As Integer
Dim bolDelCM As Boolean 'Added by Lydia 2016/10/19 新案(101,102)銷案時,取消一案兩請關聯
Dim bolFMP As Boolean  'Added by Lydia 2016/10/19 是否為FMP案
Dim bolFMP2 As Boolean 'Added by Lydia 2023/06/09 是否為寰華案
Dim m_PA177 As String 'Added by Lydia 2023/07/28 FCP專利連結通知

Select Case Index
             Case 0
                        varSaveCursor = Screen.MousePointer
                        Screen.MousePointer = vbHourglass
                        For i = 0 To 0
                               If txtCaseField(i).Enabled Then
                                  If CheckKeyIn(i) <> 1 Then
                                     txtCaseField(i).SetFocus
                                     txtCaseField_GotFocus (i)
                                    'Modify By Cheng 2002/05/29
'                                     Exit For
                                      Screen.MousePointer = varSaveCursor
                                      Exit Sub
                                  End If
                               End If
                        Next
                        'If i = 1 Then
                        '    If SaveDatabase Then
                        '       bolLeave = True
                        '       Unload Me
                        '    End If
                        'End If
                        '**************************************************
                        ' nick 900803 改
                        For j = 1 To intTotalCaseCode
                           If Trim(UCase(frm110103_2.grdDataList.TextMatrix(j, 2))) <> "Y" Then
                              cp(1) = frm110103_2.grdDataList.TextMatrix(j, 6)
                              cp(2) = frm110103_2.grdDataList.TextMatrix(j, 7)
                              cp(3) = frm110103_2.grdDataList.TextMatrix(j, 8)
                              cp(4) = frm110103_2.grdDataList.TextMatrix(j, 9)
                                
                              'Added by Lydia 2016/10/19 判斷是否為FMP案
                              'Modified by Morgan 2021/2/2
                              'bolFMP = False
                              'If cp(1) = "P" Then
                              '   strExc(0) = "select cp09 from caseprogress,patent where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' " & _
                              '               "and cp31='Y' and substr(cp12,1,1)='F' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09<>'000'"
                              '   intI = 1
                              '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              '   If intI = 1 Then
                              '      bolFMP = True
                              '   End If
                              'End If
                              bolFMP = PUB_ChkIsFMP(cp(1), cp(2), cp(3), cp(4))
                              'end 2021/2/2
                              'Added by Lydia 2023/06/09 判斷寰華案
                              bolFMP2 = False
                              If bolFMP = True Then
                                 bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, cp(1), cp(2), cp(3), cp(4))
                              End If
                              'end 2023/06/09
                              'Added by Lydia 2023/07/28 FCP專利連結通知
                              If cp(1) = "FCP" Then
                                 strExc(0) = "select pa177 from patent where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & cp(3) & "' and pa04='" & cp(4) & "' "
                                 intI = 1
                                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                                 If intI = 1 Then
                                    m_PA177 = "" & RsTemp.Fields("pa177")
                                 End If
                              Else
                                 m_PA177 = ""
                              End If
                              'end 2023/07/28
                              
                              'Added by Lydia 2016/10/12 新案(101,102)銷案時,取消一案兩請關聯
                              bolDelCM = False
                              strExc(8) = frm110103_2.grdDataList.TextMatrix(j, 10)
                              If cp(1) <> "FCP" And bolFMP = False And (strExc(8) = "1" Or strExc(8) = "2") Then
                                If PUB_DualCaseRelationExist(cp) Then
                                   If PUB_ChkCPExist(cp, IIf(strExc(8) = "1", "101", "102"), 1) Then '判斷未發文的新案才取消關聯
                                      bolDelCM = True
                                   End If
                                End If
                              End If
                              'end 2016/10/12
                              
                              'UPDATE 基本檔是否閉卷,閉卷日期,閉卷原因
                              'Modify By Cheng 2002/01/29
                              '更新各基本檔的備註進度
                              Select Case Val(CheckSys(cp(1)))
                              Case 1
                                    'Modify By Cheng 2002/01/29
'                                   strSQL = "UPDATE PATENT SET PA57='Y',PA58=" & ChangeTStringToWString(txtCaseField(0)) & ",PA59='" & strReasonNo(cboReason.ListIndex) & "' WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "' "
                                   strSql = "UPDATE PATENT SET PA57='Y',PA58=" & ChangeTStringToWString(txtCaseField(0)) & ",PA59='" & strReasonNo(cboReason.ListIndex) & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,PA91='" & Me.cboNote.Text & "' ") & " WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "' "
                                   cnnConnection.Execute strSql
                                    'Add By Cheng 2002/05/29
                                   strSql = "UPDATE PATENT SET PA89='" & Me.txtCaseField(1).Text & "' WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "' "
                                   cnnConnection.Execute strSql
                                   
                              Case 2
                                    'Modify By Cheng 2002/01/29
'                                   strSQL = "UPDATE TRADEMARK SET TM29='Y',TM30=" & ChangeTStringToWString(txtCaseField(0)) & ",TM31='" & strReasonNo(cboReason.ListIndex) & "' WHERE TM01='" & cp(1) & "' AND TM02='" & cp(2) & "' AND TM03='" & cp(3) & "' AND TM04='" & cp(4) & "' "
                                   strSql = "UPDATE TRADEMARK SET TM29='Y',TM30=" & ChangeTStringToWString(txtCaseField(0)) & ",TM31='" & strReasonNo(cboReason.ListIndex) & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,TM58='" & Me.cboNote.Text & "' ") & " WHERE TM01='" & cp(1) & "' AND TM02='" & cp(2) & "' AND TM03='" & cp(3) & "' AND TM04='" & cp(4) & "' "
                                   cnnConnection.Execute strSql
                              Case 3
                                    'Modify By Cheng 2002/01/29
'                                   strSQL = "UPDATE LAWCASE SET LC08='Y',LC09=" & ChangeTStringToWString(txtCaseField(0)) & ",LC10='" & strReasonNo(cboReason.ListIndex) & "' WHERE LC01='" & cp(1) & "' AND LC02='" & cp(2) & "' AND LC03='" & cp(3) & "' AND LC04='" & cp(4) & "' "
                                   strSql = "UPDATE LAWCASE SET LC08='Y',LC09=" & ChangeTStringToWString(txtCaseField(0)) & ",LC10='" & strReasonNo(cboReason.ListIndex) & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,LC27='" & Me.cboNote.Text & "' ") & " WHERE LC01='" & cp(1) & "' AND LC02='" & cp(2) & "' AND LC03='" & cp(3) & "' AND LC04='" & cp(4) & "' "
                                   cnnConnection.Execute strSql
                              Case 4
                                    'Modify By Cheng 2002/01/29
'                                   strSQL = "UPDATE HIRECASE SET HC09='Y',HC10=" & ChangeTStringToWString(txtCaseField(0)) & ",HC11='" & strReasonNo(cboReason.ListIndex) & "' WHERE HC01='" & cp(1) & "' AND HC02='" & cp(2) & "' AND HC03='" & cp(3) & "' AND HC04='" & cp(4) & "' "
                                   strSql = "UPDATE HIRECASE SET HC09='Y',HC10=" & ChangeTStringToWString(txtCaseField(0)) & ",HC11='" & strReasonNo(cboReason.ListIndex) & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,HC12='" & Me.cboNote.Text & "' ") & " WHERE HC01='" & cp(1) & "' AND HC02='" & cp(2) & "' AND HC03='" & cp(3) & "' AND HC04='" & cp(4) & "' "
                                   cnnConnection.Execute strSql
                              Case 5, 6, 7, 8
                                    'Modify By Cheng 2002/01/29
'                                   strSQL = "UPDATE SERVICEPRACTICE SET SP15='Y',SP16=" & ChangeTStringToWString(txtCaseField(0)) & ",SP17='" & strReasonNo(cboReason.ListIndex) & "' WHERE SP01='" & cp(1) & "' AND SP02='" & cp(2) & "' AND SP03='" & cp(3) & "' AND SP04='" & cp(4) & "' "
                                   strSql = "UPDATE SERVICEPRACTICE SET SP15='Y',SP16=" & ChangeTStringToWString(txtCaseField(0)) & ",SP17='" & strReasonNo(cboReason.ListIndex) & "' " & IIf(Len(Me.cboNote.Text) <= 0, "", " ,SP18='" & Me.cboNote.Text & "' ") & " WHERE SP01='" & cp(1) & "' AND SP02='" & cp(2) & "' AND SP03='" & cp(3) & "' AND SP04='" & cp(4) & "' "
                                   cnnConnection.Execute strSql
                              Case Else
                              End Select
                              'UPDATE 進度檔,取消收文日期,取消收文原因
                              '93.10.5 MODIFY BY SONIA
                              'strSQL = "UPDATE CASEPROGRESS SET CP57=" & ChangeTStringToWString(txtCaseField(0)) & ",CP58='" & strReasonNo(cboReason.ListIndex) & "' WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP57 IS NULL AND CP27 IS NULL "
                              strSql = "UPDATE CASEPROGRESS SET CP26='N',CP57=" & ChangeTStringToWString(txtCaseField(0)) & ",CP58='" & strReasonNo(cboReason.ListIndex) & "' WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP57 IS NULL AND CP27 IS NULL "
                              '93.10.5 END
                               'Added by Lydia 2016/01/29 排除FCP案的代辦退費(實審,再審和再審延期)
                              If cp(1) = "FCP" Then
                                  strSql = strSql & "and cp09 not in (select a.cp09 from caseprogress a,caseprogress b where a.cp01='" & cp(1) & "' and a.cp02='" & cp(2) & "' and a.cp03='" & cp(3) & "' and a.cp04='" & cp(4) & "' and a.cp10='" & 退費 & "' and a.cp27||a.cp57 is null and b.cp09(+)=a.cp43 and b.cp10 in ('416','107') " & _
                                           "union select a.cp09 from  caseprogress a,caseprogress b,nextprogress where a.cp01='" & cp(1) & "' and a.cp02='" & cp(2) & "' and a.cp03='" & cp(3) & "' and a.cp04='" & cp(4) & "' and a.cp10='" & 退費 & "' and a.cp27||a.cp57 is null and b.cp09(+)=a.cp43 and b.cp10='404' and np01(+)=b.cp43 and np07='107' " & _
                                           "union select a.cp09 from  caseprogress a,caseprogress b,caseprogress c where a.cp01='" & cp(1) & "' and a.cp02='" & cp(2) & "' and a.cp03='" & cp(3) & "' and a.cp04='" & cp(4) & "' and a.cp10='" & 退費 & "' and a.cp27||a.cp57 is null and b.cp09(+)=a.cp43 and b.cp10='404' and c.cp09(+)=b.cp43 and c.cp10='107') "
                              End If
                              'end 2016/01/29
                              cnnConnection.Execute strSql
                                                                     
                              'Add By Cheng 2002/01/22
                              '更新案件進度檔時, 當無發文日(CP27 Is Null)資料時, 才更新是否算案件數CP26為N
                              '93.10.5 CANCEL BY SONIA 改在前一句
                              'strSQL = "UPDATE CASEPROGRESS SET CP26='N' WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP27 IS NULL "
                              'cnnConnection.Execute strSQL
                              '93.10.5 END
                                       
                                       'UPDATE CASEPROGRESS SET CP57=20010802,CP58='01' WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS WHERE CP01='FCP' AND CP02='022626')
                              'UPDATE 下一程序檔解除期限日期,解除期限原因
                              '93.10.5 MODIFY BY SONIA 只更新 是否續辦為 NULL 者
                              'strSQL = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & ChangeTStringToWString(txtCaseField(0)) & ",NP12='" & strReasonNo(cboReason.ListIndex) & "' WHERE NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP05='" & cp(4) & "' AND NP11 IS NULL "
                              strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & ChangeTStringToWString(txtCaseField(0)) & ",NP12='" & strReasonNo(cboReason.ListIndex) & "' WHERE NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP05='" & cp(4) & "' AND NP11 IS NULL AND NP06 IS NULL"
                              '93.10.5 END
                              cnnConnection.Execute strSql
                             
                              ' ADD 到案件進度檔
                                 Dim strAutoNum As String
                                 'Modify By Cheng 2002/10/01
'                                 If objPublicData.GetAutoNumber("B", strAutoNum, True, False) Then
                                 'edit by nickc 2007/02/02 不用 dll 了
                                 'If objPublicData.GetAutoNumber("B", strAutoNum, True, True) Then
                                 If ClsPDGetAutoNumber("B", strAutoNum, True, True) Then
                                      CheckOC
                                      strSql = "select au01||(au02-1911) from autonumber where au01='B'"
                                      adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                                      If Not adoRecordset.BOF Then adoRecordset.MoveFirst
                                      If adoRecordset.BOF And adoRecordset.EOF Then MsgBox "自動編號錯誤", vbInformation: Exit Sub
                                      'Modify By Sindy 2010/8/18 比對自動編號年度
                                      'strAutoNum = CheckStr(adoRecordset.Fields(0).Value) & strAutoNum
                                      strAutoNum = "B" + CompAutoNumberYear(CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911)) & strAutoNum
                                      CheckOC
                                      strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp12,cp13,cp14,cp15,cp16,cp17,cp18,cp19,cp20,cp21,cp22,cp23,cp24,cp25,cp26,cp27,cp28,cp29,cp30,cp31,cp32,cp33,cp34,cp35,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp43,cp44,cp45,cp46,cp47,cp48,cp49,cp50,cp51,cp52,cp53,cp54,cp55,cp56,cp57,cp58,cp59,cp60,cp61,cp62,cp63,cp64,cp71,cp72,cp73,cp74,cp75,cp76,cp77,cp78,cp79) values "
                                      'Set SCp() = cp()
                                      For i = 1 To 79
                                         Select Case i
                                         '文字null
                                          'Modify By Cheng 2002/05/29
'                                         Case 8, 11, 12, 13, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 49, 50, 51, 52, 55, 56, 58, 59, 60, 61, 62, 63, 64
                                         '92.1.25 MODIFY BY SONIA 取消收文日及原因要存
                                         'Case 8, 11, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 49, 50, 51, 52, 55, 56, 58, 59, 60, 61, 62, 63, 64
                                         Case 8, 11, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 49, 50, 51, 52, 55, 56, 59, 60, 61, 62, 63, 64
                                         '92.1.25 END
                                              SCp(i) = "null "
                                         Case 13 '智權人員
                                              SCp(i) = "'" & frm110103_1.txtCaseField(4).Text & "'"
                                         Case 12 '業務區
                                              SCp(i) = "'" & GetST15(frm110103_1.txtCaseField(4).Text) & "'"
                                         '文字畫面上
                                         Case 14
                                              SCp(i) = "'" & strUserNum & "'"
                                         Case 1, 2, 3, 4
                                              SCp(i) = "'" & Trim(ChgSQL(cp(i))) & "'"
                                         Case 5, 27
                                              SCp(i) = GetTodayDate
                                         Case 9
                                              SCp(i) = "'" & strAutoNum & "'"
                                         '91.12.6 modify by sonia
                                         'Case 26, 20, 32
                                         '     SCp(i) = "'N'"
                                         Case 20
                                            If intWhere <> "2" Then
                                               SCp(i) = "'N'"
                                               '2013/8/13 add by sonia FMT要請款
                                               If cp(1) = "T" And Left(GetST15(frm110103_1.txtCaseField(4).Text), 1) = "F" Then
                                                  SCp(i) = "null "
                                               End If
                                               '2013/8/13 end
                                            Else
                                               SCp(i) = "null "
                                            End If
                                         Case 26, 32
                                               SCp(i) = "'N'"
                                         '91.12.6 end
                                         'Case 43
                                         '     SCp(i) = "'" & frm110102_1.grdDataList.TextMatrix(frm110102_1.grdDataList.Row, 0) & "'"
                                         Case 10
                                               Select Case Val(CheckSys(cp(1)))
                                               Case 1, 5         'patent
                                                  SCp(i) = "'913'"
                                               Case 2, 6         'trademark
                                                  SCp(i) = "'704'"
                                               Case 3, 4, 7, 8   'lawcase & hirecase
                                                  SCp(i) = "'993'" 'Modify By Sindy 2011/10/26 999=>993.閉卷
                                               Case Else
                                               End Select
                                         Case 65, 66, 67, 68, 69, 70
                                              SCp(i) = ""
                                         '92.1.25 ADD BY SONIA
                                         Case 57
                                              SCp(i) = ChangeTStringToWString(txtCaseField(0))
                                         Case 58
                                                'Modify By Cheng 2004/04/15
'                                              SCp(i) = strReasonNo(cboReason.ListIndex)
                                                SCp(i) = IIf(strReasonNo(cboReason.ListIndex) = "", "Null", CNULL(strReasonNo(cboReason.ListIndex)))
                                                'End
                                          '92.1.25 END
                                         '數字
                                         Case Else
                                              SCp(i) = "null "
                                         End Select
                                      Next i
                                      strSql = strSql & " ("
                                      For i = 1 To 79
                                          Select Case i
                                          Case 65, 66, 67, 68, 69, 70
                                          Case Else
                                               strSql = strSql & SCp(i)
                                               If i <> 79 Then
                                                  strSql = strSql & ","
                                               End If
                                          End Select
                                      Next i
                                      strSql = strSql & ") "
                                      cnnConnection.Execute strSql
                                      
                                      'Add by Sindy 2013/04/12 更新c類的代理人及彼所案號，要在新增c類之後
                                      Pub_UpdateFromMaxCP27 cp(1), cp(2), cp(3), cp(4)
                                      
                                      'Added by Lydia 2016/10/19 新案(101,102)銷案時,取消一案兩請關聯
                                      If bolDelCM = True Then
                                        strExc(0) = cp(1): strExc(1) = cp(2): strExc(2) = cp(3): strExc(3) = cp(4)
                                        strExc(4) = "": strExc(5) = "": strExc(6) = "": strExc(7) = ""
                                        If PUB_DeleteCaseRelation(strExc, 3) Then
                                        End If
                                      End If
                                      'end 2016/10/19
                                      'Added by Lydia 2023/06/09 當寰華案在key閉卷按確認時，請判斷是否有相關香港案及澳門案未不續辦/閉卷，若有則發mail
                                      If bolFMP2 = True And cp(1) = "P" Then
                                         Call ClsPDGetCaseNation(1, cp(1), cp(2), cp(3), cp(4), strExc(0))
                                         If strExc(0) = "020" Then
                                            'Modified by Lydia 2023/06/28 傳入案件性質SCp(10)
                                            'Modified by Lydia 2025/04/02 去掉案件性質SCp(10)的單引號Replace
                                            Call PUB_CloseMailto013044("1", cp(1), cp(2), cp(3), cp(4), Replace(SCp(10), "'", ""))
                                         End If
                                      End If
                                      'end 2023/06/09
                                      'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：輸入閉卷913自動收文「通知資訊變更961」,發一封Email給承辦工程師
                                      If cp(1) = "FCP" And m_PA177 = "Y" Then
                                         'Memo by Lydia 2025/04/02 模組內已去掉SCp(10)的單引號Replace
                                         If PUB_GetFCPlinkMC("6", TransDate(txtCaseField(0), 2), cp, strAutoNum, SCp(10)) = True Then
                                         End If
                                      End If
                                      'end 2023/07/28
                                      
                                      intLeaveKind = 1
                                      Me.Hide
                                 Else
                                     MsgBox ("自動給號錯誤")
                                     Me.Hide
                                 End If
                           End If
                           frm110103_2.ReChoose j, strCaseCode()
                        Next j
                        bolLeave = True
                        Call PUB_SendMailCache 'Added by Lydia 2023/06/09
                        Unload Me
                        Screen.MousePointer = vbDefault
                        
                        '**************************************************************************************************

                        Screen.MousePointer = varSaveCursor
             Case 1, 2
                        If Index = 2 Then
                           intLeaveKind = 0
                        Else
                           intLeaveKind = 1
                        End If
                        bolLeave = False
                        Unload Me
End Select
End Sub
Private Sub Form_Activate()
Dim strReasonName() As String, i As Integer

Select Case ReadReasonOfRelief(strReasonNo(), strReasonName())
             Case 1
                        For i = 0 To UBound(strReasonNo)
                              cboReason.AddItem strReasonName(i)
                        Next
                        cboReason.ListIndex = 0
             Case -1
                        Unload Me
End Select
End Sub
Private Sub Form_Load()
Dim i As Integer

'Memo by Amy 2025/08/06  不續辦但准通知 改為 後續准駁簡單報告
MoveFormToCenter Me
intTotalCaseCode = frm110103_2.grdDataList.Rows - 1
ReDim Preserve strCaseCode(3, intTotalCaseCode - 1)
For i = 1 To intTotalCaseCode
       strCaseCode(0, i - 1) = frm110103_2.grdDataList.TextMatrix(i, 6)
       strCaseCode(1, i - 1) = frm110103_2.grdDataList.TextMatrix(i, 7)
       strCaseCode(2, i - 1) = frm110103_2.grdDataList.TextMatrix(i, 8)
       strCaseCode(3, i - 1) = frm110103_2.grdDataList.TextMatrix(i, 9)
       Select Case strCaseCode(0, i - 1)
                    Case "CFP"
                               lblNumber(0) = Val(lblNumber(0)) + 1
                    Case "CFT"
                               lblNumber(1) = Val(lblNumber(1)) + 1
                    Case "CFL"
                               lblNumber(2) = Val(lblNumber(2)) + 1
                    Case "FCP"
                               lblNumber(3) = Val(lblNumber(3)) + 1
                    Case "FCT"
                               lblNumber(4) = Val(lblNumber(4)) + 1
                    Case "FCL"
                               lblNumber(5) = Val(lblNumber(5)) + 1
                    Case "P"
                               lblNumber(6) = Val(lblNumber(6)) + 1
                    Case "T"
                               lblNumber(7) = Val(lblNumber(7)) + 1
                    Case "L"
                               lblNumber(8) = Val(lblNumber(8)) + 1
                    Case Else
                               lblNumber(9) = Val(lblNumber(9)) + 1
       End Select
       lblNumber(10) = Val(lblNumber(10)) + 1
Next
bolLeave = False
intLeaveKind = 1
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If intLeaveKind = 1 Then
   frm110103_2.Show
Else
   Unload frm110103_2
End If
   'Add By Cheng 2002/07/18
   Set frm110103_4 = Nothing
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = -1 Then
   Cancel = True
   txtCaseField_GotFocus (Index)
End If
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
CheckKeyIn = -1
Select Case intIndex
             Case 0
                        If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                           '2010/8/3 加val
                           If Val(txtCaseField(intIndex)) <= Val(GetTaiwanTodayDate) Then
                              CheckKeyIn = 1
                           Else
                              ShowMsg MsgText(8003)
                           End If
                         End If
             Case Else
                        CheckKeyIn = 1
End Select
End Function
Private Sub txtCaseField_GotFocus(Index As Integer)
txtCaseField(Index).SelStart = 0
txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
End Sub
Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
             Case 0
                        KeyAscii = UpperCase(KeyAscii)
             'Add By Cheng 2002/05/29
             Case 1
                        KeyAscii = UpperCase(KeyAscii)
                        If KeyAscii <> 89 And KeyAscii <> 8 Then
                           KeyAscii = 0
                        End If
End Select
End Sub

'讀取解除期限原因
Private Function ReadReasonOfRelief(ByRef strReasonNo() As String, ByRef strReasonName() As String) As Integer
Dim strSql As String, rsRecordset As New ADODB.Recordset, i As Integer

On Error GoTo ErrHand
strSql = "select ror01,ror02 from reasonofrelief"
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSql, cnnConnection
If rsRecordset.RecordCount > 0 Then
   Do While Not rsRecordset.EOF
         ReDim Preserve strReasonNo(i) As String
         ReDim Preserve strReasonName(i) As String
         strReasonNo(i) = rsRecordset.Fields(0)
         strReasonName(i) = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
         i = i + 1
         rsRecordset.MoveNext
   Loop
   ReadReasonOfRelief = 1
Else
   ReadReasonOfRelief = 0
End If
Exit Function
ErrHand:
ShowMsg MsgText(8001)
ReadReasonOfRelief = -1
End Function

