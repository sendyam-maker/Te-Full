VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071005 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文"
   ClientHeight    =   5760
   ClientLeft      =   636
   ClientTop       =   4428
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9324
   Begin VB.CommandButton Command1 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7968
      TabIndex        =   1
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7140
      TabIndex        =   0
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4092
      Left            =   192
      TabIndex        =   5
      Top             =   1560
      Width           =   8952
      _ExtentX        =   15790
      _ExtentY        =   7218
      _Version        =   393216
      Cols            =   8
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   1110
      Width           =   7245
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12779;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   285
      Left            =   2370
      TabIndex        =   8
      Top             =   780
      Width           =   6435
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "11351;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   420
      Width           =   972
   End
   Begin VB.Label lbeCaseNum 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   420
      Width           =   2292
   End
   Begin VB.Label lbeCustomer 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   780
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "當  事  人："
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   780
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   1140
      Width           =   972
   End
End
Attribute VB_Name = "frm071005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; cboCaseName、lbeCusName、MSHFlexGrid1
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim intLastRow As Integer, blnOKtoShow As Boolean, intCols As Integer
Dim intClkRow As Integer

Private Sub Command1_Click(Index As Integer)
 Dim i As Integer, n As Integer
Select Case Index
Case 0
    With MSHFlexGrid1
    n = 0
    .Visible = False
    For i = 1 To .Rows - 1
        .row = i
        .col = 0
        If .Text = "v" Then
           Exit For
        Else
           If i = .Rows - 1 Then
             MsgBox "請點選欲發文資料"
             .Visible = True
             
             Exit Sub
           End If
        End If
    Next
          .Visible = True
    End With
    
   'Add By Sindy 2021/10/8
   '檢查是否有承辦歷程是否有產生承辦單可以發文
   If PUB_IsEmpFlowIsSend(MSHFlexGrid1.TextMatrix(i, 2)) = False Then '收文號
      Exit Sub
   End If
   '2021/10/8 END
    
    frm071006.Show
    If IsNoExistData Then
       IsNoExistData = False
       Unload frm071006
    Else
       Me.Hide
    End If
Case 1
    frm071004.Show
    Unload Me
End Select
    intForm = 5
End Sub

Private Sub Form_Activate()
   GridHead
End Sub

Private Sub Form_Load()
 Dim strDNum As String
   MoveFormToCenter Me
   lbeCaseNum = GiveSymbol(frm071004.txtcp01, frm071004.txtcp02, frm071004.txtcp03, frm071004.txtcp04, strDNum)
   GetCaseNum (strDNum)
End Sub

Private Sub GridHead()
 Dim i As Integer
   With MSHFlexGrid1
      blnOKtoShow = False
      .Visible = False
      .row = 0
      .col = 0
      .MergeCells = flexMergeRestrictRows
      .MergeRow(0) = True
      .ColWidth(0) = 200: .Text = "v"
      .col = 1: .ColWidth(1) = 900: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1200: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 900: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 900: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      'Modified by Lydia 2015/10/05
      '.col = 6: .ColWidth(6) = 900: .Text = "法務人員"
      .col = 6: .ColWidth(6) = 900: .Text = "協辦人員"
      .CellAlignment = flexAlignCenterCenter
      'Add by Amy 2018/08/15
      .col = 7: .ColWidth(7) = 2000: .Text = "進度備註"
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
      intLastRow = 0
      blnOKtoShow = True
      '判斷是否有資料
      If .Rows > 1 Then
         '將第一列反白
         .row = 1
      End If
      .Visible = True
   End With
End Sub

Private Sub GetCaseNum(txt As String)
   '2009/9/9 MODIFY BY SONIA 取消cp09<'C'條件
   'Modify by Amy 2018/08/15 +cp64
   If frm071004.txtcp01 <> "LA" Then
      strExc(1) = "select CP05, cp09, CP10,cp13,cp14,cp29,cp64,cp01,cp02,cp03,cp04,lc05,lc06,lc07,lc11 from lawcase,caseprogress" + _
         " where " & ChgCaseprogress(txt) + " and CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 and cp27 is null and cp57 is null"
   Else
      '2011/6/20 MODIFY BY SONIA 加入CP10<>'0'條件
      strExc(1) = "select CP05, cp09 ,CP10,cp13,cp14,cp29,cp64,cp01,cp02,cp03,cp04,hc06,hc05 from hirecase,caseprogress" + _
         " where " & ChgCaseprogress(txt) + " and CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 and cp27 is null and cp57 is null and cp10<>'0'"
   End If
   'end 2018/08/15
   strExc(1) = strExc(1) & " ORDER BY CP05 DESC,CP09"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      GridHead
      PutDataInObj
   Else
     IsNoExistData = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm071005 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
 Dim i As Integer
   With MSHFlexGrid1
   intCols = MSHFlexGrid1.Cols - 1
   ShowBar MSHFlexGrid1, intLastRow, intCols

      intClkRow = .row
      .col = 0
      ClearGrid
      .row = intLastRow
      If .Text = "v" Then
         .Text = ""
      Else
         .Text = "v"
      End If
   End With

End Sub

Private Sub PutDataInObj()
Dim i As Integer, strTempName As String, strSys As String, strTemp As String
Dim strCPM As String 'Added by Lydia 2023/12/27

   strSys = GetCaseNumSysKind(lbeCaseNum)
   'Modify By Sindy 2009/07/24 增加LIN系統類別
   'modify by sonia 2019/7/29 +ACS系統類別
   If strSys = "L" Or strSys = "FCL" Or strSys = "CFL" Or strSys = "LIN" Or strSys = "ACS" Then
      If Not IsNull(RsTemp.Fields!LC11) Then
         strTemp = RsTemp.Fields!LC11
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCustomer(strTemp, strTempName) Then lbeCustomer = strTemp: lbeCusName = strTempName
         If ClsPDGetCustomer(strTemp, strTempName) Then lbeCustomer = strTemp: lbeCusName = strTempName
      End If
      If Not IsNull(RsTemp.Fields!lc05) Then cboCaseName.AddItem "中:" + RsTemp.Fields!lc05
      If Not IsNull(RsTemp.Fields!lc06) Then cboCaseName.AddItem "英:" + RsTemp.Fields!lc06
      If Not IsNull(RsTemp.Fields!lc07) Then cboCaseName.AddItem "日:" + RsTemp.Fields!lc07
   ElseIf strSys = "LA" Then
      If Not IsNull(RsTemp.Fields!hc05) Then
         strTemp = RsTemp.Fields!hc05
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCustomer(strTemp, strTempName) Then lbeCustomer = strTemp: lbeCusName = strTempName
         If ClsPDGetCustomer(strTemp, strTempName) Then lbeCustomer = strTemp: lbeCusName = strTempName
      End If
      If Not IsNull(RsTemp.Fields!hc06) Then cboCaseName.AddItem "中:" + RsTemp.Fields!hc06
   End If
   If cboCaseName.ListCount <> 0 Then
      cboCaseName.ListIndex = 0
   End If
   
   With MSHFlexGrid1
      If Not (RsTemp.EOF And RsTemp.BOF) Then
         RsTemp.MoveFirst
         i = 1
         Do
             If .Rows = .row + 1 Then .Rows = .Rows + 1
             .row = i
             .col = 1
             If Not IsNull(RsTemp.Fields!cp05) Then
                .Text = ChangeTStringToTDateString(ChangeWStringToTString(RsTemp.Fields!cp05))
             End If
             .col = 2
             .Text = IIf(IsNull(RsTemp.Fields!CP09), "", RsTemp.Fields!CP09)
             strCPM = "" 'Added by Lydia 2023/12/27
             .col = 3
                 'edit by nickc 2007/02/07 不用 dll 了
                 'If objPublicData.GetCaseProperty(RsTemp.Fields!CP01, RsTemp.Fields!CP10, strTempName, False) Then
                 If ClsPDGetCaseProperty(RsTemp.Fields!cp01, RsTemp.Fields!CP10, strTempName, False) Then
                  .Text = strTempName
                  strCPM = strTempName
                 End If
             .col = 4
                 If Not IsNull(RsTemp.Fields!cp13) Then
                     'edit by nickc 2007/02/07 不用 dll 了
                     'If objPublicData.GetStaff(RsTemp.Fields!cp13, strTempName) Then .Text = strTempName
                     'Modified by Lydia 2023/12/27 修改智權人員欄或CP14承辦人欄己離職，A：仍要帶出姓名、B：彈出的訊息請帶出是智權人員或是承辦人或是協辦人員。
                     'If ClsPDGetStaff(RsTemp.Fields!cp13, strTempName) Then .Text = strTempName
                     strTempName = GetStaffName("" & RsTemp.Fields("cp13"), True, , , strTemp)
                     If strTemp <> "1" Then
                         MsgBox "收文號：" & RsTemp.Fields("CP09") & strCPM & vbCrLf & "智權人員已離職！"
                     End If
                     .Text = strTempName
                 End If
             .col = 5
                 If Not IsNull(RsTemp.Fields!cp14) Then
                     'edit by nickc 2007/02/07 不用 dll 了
                     'If objPublicData.GetStaff(RsTemp.Fields!cp14, strTempName) Then .Text = strTempName
                     'Modified by Lydia 2023/12/27 修改智權人員欄或CP14承辦人欄己離職，A：仍要帶出姓名、B：彈出的訊息請帶出是智權人員或是承辦人或是協辦人員。
                     'If ClsPDGetStaff(RsTemp.Fields!cp14, strTempName) Then .Text = strTempName
                     strTempName = GetStaffName("" & RsTemp.Fields("cp14"), True, , , strTemp)
                     If strTemp <> "1" Then
                         MsgBox "收文號：" & RsTemp.Fields("CP09") & strCPM & vbCrLf & "承辦人員已離職！"
                     End If
                     .Text = strTempName
                 End If
             .col = 6
                 If Not IsNull(RsTemp.Fields!cp29) Then
                     'edit by nickc 2007/02/07 不用 dll 了
                     'If objPublicData.GetStaff(RsTemp.Fields!cp29, strTempName) Then .Text = strTempName
                     'Modified by Lydia 2023/12/27 修改智權人員欄或CP14承辦人欄己離職，A：仍要帶出姓名、B：彈出的訊息請帶出是智權人員或是承辦人或是協辦人員。
                     'If ClsPDGetStaff(RsTemp.Fields!cp29, strTempName) Then .Text = strTempName
                     strTempName = GetStaffName("" & RsTemp.Fields("cp29"), True, , , strTemp)
                     If strTemp <> "1" Then
                         MsgBox "收文號：" & RsTemp.Fields("CP09") & strCPM & vbCrLf & "協辦人員已離職！"
                     End If
                     .Text = strTempName
                 End If
             'Add by Amy 2018/08/15
             .col = 7
             .Text = "" & RsTemp.Fields!CP64
             .CellAlignment = flexAlignLeftCenter
             'end 2018/08/15
             RsTemp.MoveNext
             i = i + 1
         Loop Until RsTemp.EOF
      End If
   End With
End Sub

Private Sub ClearGrid()
Dim i As Integer
   
   With MSHFlexGrid1
      .Visible = False
      For i = 1 To .Rows - 1
         .col = 0
         .row = i
         .Text = ""
      Next
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then MSHFlexGrid1_Click

End Sub
