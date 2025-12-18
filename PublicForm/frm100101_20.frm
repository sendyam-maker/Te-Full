VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_20 
   BorderStyle     =   1  '單線固定
   Caption         =   "以往來對象查詢國內往來記錄"
   ClientHeight    =   5580
   ClientLeft      =   1815
   ClientTop       =   2595
   ClientWidth     =   9465
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9465
   Begin VB.ListBox lstUsers 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      ItemData        =   "frm100101_20.frx":0000
      Left            =   8400
      List            =   "frm100101_20.frx":0007
      MultiSelect     =   1  '簡易多重選取
      Sorted          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4950
      Width           =   1035
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Bindings        =   "frm100101_20.frx":0015
      Height          =   4950
      Left            =   90
      TabIndex        =   3
      Top             =   570
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   8731
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
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
      _Band(0).Cols   =   18
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆(&N)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   7845
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   75
      Width           =   870
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "往來記錄(&C)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   2
      Left            =   6750
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   75
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   8715
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   75
      Width           =   660
   End
   Begin MSForms.Label Label3 
      Height          =   300
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   5130
      VariousPropertyBits=   27
      Size            =   "9049;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "往來對象："
      Height          =   180
      Left            =   135
      TabIndex        =   4
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "frm100101_20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、Label3
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
Option Explicit
Public CRdateF As String, CRdateT As String
Public cmdState As Integer

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   cmdState = -1
   lstUsers(0).Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100101_20 = Nothing
End Sub

Public Sub PubShowNextData()
   Dim i As Integer, j As Integer
   Select Case cmdState
      Case 0 '結束
         fnCloseAllFrm100
      Case 1 '下一筆
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 2
         Me.Enabled = False
         With grdDataList
            For i = 1 To .Rows - 1
               .col = 0
               .row = i
               If Trim(.Text) = "V" Then
                  .col = 0
                  .Text = ""
                  For j = 0 To .Cols - 1
                      If j <> 1 Then
                          .col = j
                          .CellBackColor = QBColor(15)
                      End If
                  Next j
                  .col = 1
                  If Not IsNull(.Text) Then
                     If fnSaveParentForm(Me) = False Then
                        Me.Enabled = True
                        Exit Sub
                     End If
                     Screen.MousePointer = vbHourglass
                     frm100101_19.Show
                     frm100101_19.Tag = Pub_RplStr(.Text)
                     frm100101_19.StrMenu
                     Screen.MousePointer = vbDefault
                     Me.Enabled = True
                     Exit Sub
                  End If
               End If
            Next i
         End With
         Me.Enabled = True
   End Select
End Sub

Public Sub StrMenu()
Dim strKey As String, strCon As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer, k As Integer
Dim strCU13 As String 'Add by Amy 2017/07/17

   'Add By Sindy 2011/01/03 檢查國內外權限
   If CheckSR12(Me.Tag) = False Then
      Screen.MousePointer = vbDefault
      Me.Enabled = True
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
   
   strKey = Left(Me.Tag, 8)
   
   '往來對象資料
   strExc(0) = "select N1,N2, N3, NO1 from( "
   'Modified by Lydia 2019/10/17 +CU12
   strExc(0) = strExc(0) & " select NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) N1,CU01 NO1,CU13 N2,CU10 N3 from customer where cu01='" & strKey & "' and cu02='0' "
   'strExc(0) = strExc(0) & " union all select NVL(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)) N1,FA01 NO1,FA10 N3 from fagent where fa01='" & strKey & "' and fa02='0' "
   strExc(0) = strExc(0) & " union all select NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)) N1,PCU01 NO1,PCU38 N2,PCU09 N3 from potcustomer where pcu01='" & strKey & "' and pcu02='0' "
   strExc(0) = strExc(0) & " union all select NVL(POC03,DECODE(POC23,NULL,POC27,POC23||' '||POC24||' '||POC25||' '||POC26)) N1,POC01 NO1,POC13 N2,POC04 N3 from potcustomer1 where poc01='" & strKey & "' and poc02='0' "
   strExc(0) = strExc(0) & " ) A "
   
   lstUsers(0).Clear
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Label3 = strKey & " " & RsTemp.Fields("N1")
      'Modify by Amy 2017/07/17 X54363010 因為智權為MCTF開頭 run PUB_Id2Num 會錯
      If InStr(Left("" & RsTemp.Fields("N2"), 4), "MCTF") > 0 Then
        strCU13 = Replace(Pub_GetSpecMan(RsTemp.Fields("N2"), False), ";", ",")
      Else
        strCU13 = "" & RsTemp.Fields("N2")
      End If
      'SetlstUsers 0, "" & RsTemp.Fields("N2")
      SetlstUsers 0, strCU13
      'end 2017/07/17
      'Modify By Sindy 2020/5/21
      'If CheckModifyLimit() = False Then
      'Modified by Lydia 2020/11/03 建檔人改成必傳(COR06或ADD新增用)
      'If PUB_CheckModifyLimit_frm100101_19(strCU13, Me.Tag, "") = False Then
      ''2020/5/21 END
      If PUB_CheckModifyLimit_frm100101_19(strCU13, Me.Tag, "ADD", True) = False Then
      
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         Exit Sub
      End If
   End If
   
   '往來日期
   If Len(CRdateF) <> 0 Then
       strCon = strCon & " AND COR02>=" & Val(ChangeTStringToWString(CRdateF))
   End If
   If Len(CRdateT) <> 0 Then
       strCon = strCon & " AND COR02<=" & Val(ChangeTStringToWString(CRdateT))
   End If
   
   strExc(0) = "select ' ' AS V,COR01 AS 往來記錄編號," & SQLDate("COR02") & " 往來日期,COR03 往來對象,COR04 主旨,COR05 內容" & _
      " from contactrecord1 where SUBSTR(cor03,1,8)='" & strKey & "'" & strCon & _
      " order by cor01"
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      
      SetDataListWidth
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         
         grdDataList.Rows = grdDataList.Rows + 1
         grdDataList.row = grdDataList.Rows - 2
         
         If Not IsNull(adoRecordset.Fields(1)) Then
            grdDataList.TextMatrix(grdDataList.row, 1) = adoRecordset.Fields(1)
         End If
         If Not IsNull(adoRecordset.Fields(2)) Then
            grdDataList.TextMatrix(grdDataList.row, 2) = adoRecordset.Fields(2)
         End If
         If Not IsNull(adoRecordset.Fields(3)) Then
            grdDataList.TextMatrix(grdDataList.row, 3) = adoRecordset.Fields(3)
         End If
         If Not IsNull(adoRecordset.Fields(4)) Then
            grdDataList.TextMatrix(grdDataList.row, 4) = adoRecordset.Fields(4)
         End If
         If Not IsNull(adoRecordset.Fields(5)) Then
            grdDataList.TextMatrix(grdDataList.row, 5) = adoRecordset.Fields(5)
         End If
         
NextRecord:
         'Added by Lydia 2018/12/22 統一靠左
         For i = 0 To grdDataList.Cols - 1
            grdDataList.col = i
            grdDataList.CellAlignment = flexAlignLeftCenter
         Next i
         'end 2018/12/22
         adoRecordset.MoveNext
      Loop
      grdDataList.Rows = grdDataList.Rows - 1
      
      If grdDataList.row = 0 Then
         ShowNoData
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      End If
   Else
      ShowNoData
      Screen.MousePointer = vbDefault
      Me.Enabled = True
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
   End If
End Sub

Private Sub SetDataListWidth()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer
   
   arrGridHeadText = Array("V", "往來記錄編號", "往來日期", "往來對象", "主旨", "內容")
   
   arrGridHeadWidth = Array(200, 1000, 800, 1500, 2000, 2000)
   
   grdDataList.Clear
   grdDataList.Rows = 2
   grdDataList.Cols = UBound(arrGridHeadText) + 1
   With grdDataList
      .row = 0
      For iCol = 0 To UBound(arrGridHeadText)
         .col = iCol
         .Text = arrGridHeadText(iCol)
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .CellAlignment = flexAlignLeftCenter 'flexAlignCenterCenter
      Next iCol
   End With
End Sub

Private Sub GrdDataList_Click()
   Dim i As Integer
   With grdDataList
      .Visible = False
      .row = .MouseRow
      .col = 0
      If .row <> 0 Then
         If .Text = "V" Then
            .Text = ""
            For i = 0 To .Cols - 1
               If i <> 1 Then
                  .col = i
                  .CellBackColor = QBColor(15)
               End If
            Next i
         Else
            .Text = "V"
            For i = 0 To .Cols - 1
               If i <> 1 Then
                  .col = i
                  .CellBackColor = &HFFC0C0
               End If
            Next i
         End If
      End If
      .Visible = True
   End With
End Sub

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   
   lstUsers(p_idx).Clear
   If p_stNums <> "" Then
      strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         arrID = Split(p_stNums, ",")
         With RsTemp
         '照原順序排
         For intI = UBound(arrID) To LBound(arrID) Step -1
            .MoveFirst
            Do While Not .EOF
               If .Fields("st01") = arrID(intI) Then
                  lstUsers(p_idx).AddItem "" & .Fields(1), 0
                  lstUsers(p_idx).ITEMDATA(0) = PUB_Id2Num(.Fields(0)) '員工編號
                  .MoveLast
               End If
               .MoveNext
            Loop
         Next
         End With
      End If
   End If
End Sub

''檢查維護權限
'Private Function CheckModifyLimit() As Boolean
'Dim idx As Integer
'
'   If lstUsers(0).ListCount = 0 Then
'      CheckModifyLimit = True
'      Exit Function
'   End If
'
'   '2009/5/14 add by sonia 開放M51權限
'   'modify by sonia 2015/6/5 開放01,09等級權限
'   If Pub_StrUserSt03 = "M51" Or Pub_strUserST05 = "01" Or Pub_strUserST05 = "09" Then
'      CheckModifyLimit = True
'      Exit Function
'   End If
'   '2009/5/14 end
'
'   'LoginUser須為開發者或其案件主管, 方可維護此筆資料
'   For idx = 0 To lstUsers(0).ListCount - 1
'      If strUserNum = PUB_Num2Id(lstUsers(0).ItemData(idx)) Then
'         CheckModifyLimit = True
'         Exit Function
'      Else
'         'modify by sonia 2017/8/10 A0909->A0908
'         strExc(0) = "SELECT A0908 FROM STAFF,ACC090 " & _
'                     "WHERE ST03=A0901(+) and ST01= '" & PUB_Num2Id(lstUsers(0).ItemData(idx)) & "' "
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If strUserNum = RsTemp(0) Then
'               CheckModifyLimit = True
'               Exit Function
'            End If
'         End If
'         'add by sonia 2017/10/17 帶人主管也可以看 82026可看X37109
'         strExc(0) = "SELECT ST52,ST53,ST54,ST55 FROM STAFF " & _
'                     "WHERE ST01= '" & PUB_Num2Id(lstUsers(0).ItemData(idx)) & "' "
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If strUserNum = "" & RsTemp(0) Then
'               CheckModifyLimit = True
'               Exit Function
'            ElseIf strUserNum = "" & RsTemp(1) Then
'               CheckModifyLimit = True
'               Exit Function
'            ElseIf strUserNum = "" & RsTemp(2) Then
'               CheckModifyLimit = True
'               Exit Function
'            ElseIf strUserNum = "" & RsTemp(3) Then
'               CheckModifyLimit = True
'               Exit Function
'            End If
'         End If
'         'end 2017/10/17
'      End If
'   Next
'
'   'Added by Lydia 2019/10/18 若客戶存在於待活化客戶檔，開放同一所別所有人都可查詢往來記錄
'   If Left(Me.Tag, 8) <> "" And Left(Me.Tag, 1) = "X" Then
'        strExc(0) = "select ocu01,st06 from oldcustomer,customer,staff " & _
'                         "where ocu01='" & Left(Me.Tag, 8) & "' and ocu01=cu01 and cu02='0' and nvl(ocu03,0)=0 and cu13=st01(+) "
'        intI = 1
'        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'        If intI = 1 Then
'           If pub_strUserOffice = "" & RsTemp("st06") Then
'               CheckModifyLimit = True
'               Exit Function
'           End If
'        End If
'   End If
'
'   CheckModifyLimit = False
'   MsgBox "無查詢權限 !!!", vbInformation
'End Function
