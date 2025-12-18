VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040118 
   BorderStyle     =   1  '單線固定
   Caption         =   "結案單審核作業"
   ClientHeight    =   5724
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5724
   ScaleWidth      =   8952
   Begin VB.CommandButton cmdChoose 
      Caption         =   "全選"
      Height          =   360
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   600
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3585
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   120
      Width           =   1400
   End
   Begin VB.CommandButton Command5 
      Caption         =   "畫面更新(&Q)"
      Default         =   -1  'True
      Height          =   345
      Left            =   5670
      TabIndex        =   3
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "明細資料(&D)"
      Height          =   345
      Index           =   2
      Left            =   6840
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   0
      Left            =   7995
      TabIndex        =   5
      Top             =   120
      Width           =   870
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4935
      Left            =   135
      TabIndex        =   7
      Top             =   690
      Width           =   8685
      _ExtentX        =   15325
      _ExtentY        =   8700
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "V|員工姓名|表單編號|表單類別|本所案號|總收文號|案件性質|本所期限|法定期限"
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
      _Band(0).Cols   =   9
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1500
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2646;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "0 / 0"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   300
      TabIndex        =   9
      Top             =   480
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "表單類別："
      Height          =   240
      Left            =   2640
      TabIndex        =   8
      Top             =   180
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "補看人員："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   930
   End
End
Attribute VB_Name = "frm040118"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/27 Form2.0已修改
'Created by Sindy 2015/1/20 參考frm040117
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Public cmdState As Integer
Dim i As Integer, j As Integer
Dim lTotRows As Long, lSelRows As Long


Public Sub PubShowNextData()
   Dim intFCState As String, strST15 As String, strSysKind As String, strNation As String 'Add by Amy 2025/04/10
   Dim strCCM18 As String 'Add by Amy 2025/06/26
   
   Select Case cmdState
   Case 2 '明細資料
      Me.Enabled = False
      For i = 1 To MSHFlexGrid1.Rows - 1
         MSHFlexGrid1.col = 0
         MSHFlexGrid1.row = i
         If Trim(MSHFlexGrid1.Text) = "V" Then
            MSHFlexGrid1.col = 0
            MSHFlexGrid1.Text = ""
            lSelRows = lSelRows - 1
            For j = 0 To MSHFlexGrid1.Cols - 1
               MSHFlexGrid1.col = j
               MSHFlexGrid1.CellBackColor = QBColor(15)
            Next j
            MSHFlexGrid1.col = 4
            If Not IsNull(MSHFlexGrid1.Text) Then
               Screen.MousePointer = vbHourglass
               Me.Hide
               'Add by Amy 2025/04/10 +FC結案單
               intFCState = 0 '非FC結案單
               strST15 = PUB_GetStaffST15(MSHFlexGrid1.TextMatrix(i, PUB_MGridGetId("F0316", MSHFlexGrid1)), 1)
               strSysKind = MSHFlexGrid1.TextMatrix(i, PUB_MGridGetId("SYSKIND", MSHFlexGrid1))
               strNation = GetPrjNation(Replace(MSHFlexGrid1.TextMatrix(i, PUB_MGridGetId("本所案號", MSHFlexGrid1)), "-", ""))
               If strSrvDate(1) >= FCP結案單電子化啟用日 Then
                  'Modify by Amy 2025/06/26 發現舊資料會頁籤判斷會有問題FCP-065275 由P案轉入,1080730閉卷-年費時,由林士堯結案
                  '       ex:FCP-065275 由P案轉入,1080730閉卷-年費時,由林士堯結案 / 外商承辦使用國內結案單操作結案 ex:T-242111(結案單號11203939)
                  strCCM18 = Pub_GetField("CloseCaseMain", "CCM01='" & MSHFlexGrid1.TextMatrix(i, PUB_MGridGetId("表單編號", MSHFlexGrid1)) & "'", "CCM18")
                  If strCCM18 = "F" Then
                     If strSysKind = "FCP" Or strSysKind = "FG" Or strSysKind = "P" Or strSysKind = "CFP" Then
                        intFCState = 2
                     Else
                        intFCState = 1
                     End If
                  End If
                  'end 2025/06/26
               End If
               frm210148_1.intFCState = intFCState
               frm210148_1.m_NP07 = MSHFlexGrid1.TextMatrix(i, PUB_MGridGetId("CP10", MSHFlexGrid1))
               'end 2025/04/10
                                 
               Call frm210148_1.SetParent(Me)
               frm210148_1.Hide
               frm210148_1.m_SignFlowEmp = Left(Combo1, 5)
               frm210148_1.txtF0301 = MSHFlexGrid1.TextMatrix(i, 1) '表單編號
               frm210148_1.Command1(1).Visible = True '進度
               frm210148_1.Command1(3).Visible = True '完整卷宗
               frm210148_1.cmdFile.Visible = False '檢覆回覆單
               frm210148_1.Show
               frm210148_1.QueryData
               
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
            lTotRows = MSHFlexGrid1.Rows - 1
            lblCount = lSelRows & " / " & lTotRows
         End If
      Next i
      Me.Enabled = True
      Call QueryData
   End Select
End Sub

'Add by Amy 2018/10/08
Private Sub cmdChoose_Click()
    MSHFlexGrid1.Visible = False
    If MSHFlexGrid1.Rows > 1 Then
        If MSHFlexGrid1.TextMatrix(1, 1) <> "" Then
            For j = 1 To MSHFlexGrid1.Rows - 1
                MSHFlexGrid1.col = 0
                MSHFlexGrid1.row = j
                MSHFlexGrid1.Text = "V"
                For i = 0 To MSHFlexGrid1.Cols - 1
                    MSHFlexGrid1.col = i
                    MSHFlexGrid1.CellBackColor = &HFFC0C0
                Next i
            Next j
        End If
    End If
    MSHFlexGrid1.Visible = True
End Sub
'end 2018/10/08

Private Sub Combo1_Click()
   If Combo1.Tag <> Combo1 Then
      Command5.Value = True
      Combo1.Tag = Combo1
   End If
End Sub

Private Sub Command1_Click(Index As Integer)
   Select Case Index
   Case 0
      Unload Me
   Case 1, 2, 3
      cmdState = Index
      PubShowNextData
      Exit Sub
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Modify by Amy 2023/02/13 + Me.Name
   Call Flow_SetF0302Combo(Combo2, , , Me.Name)
   SetCombo1
   'Call SetEmpDutyCombo(Combo1) 'Modify by Amy 2018/08/27 Sindy-改用原來的
   lTotRows = 0: lSelRows = 0
End Sub

Private Sub Command5_Click()
   SetMouseBusy
   If QueryData = False Then
      If Combo1.Tag <> "" Then
         ShowNoData
      End If
   End If
   SetMouseReady
End Sub

Private Sub SetGrd(Optional pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   'Modify by Amy 2025/04/10 +F0316,SYSKIND,CP10
   '                                                0           1                     2                  3                  4                   5                 6                   7                  8
   arrGridHeadText = Array("V", "表單編號", "表單類別", "智權人員", "本所案號", "總收文號", "案件性質", "本所期限", "法定期限" _
                                       , "F0316", "SYSKIND", "CP10")
   arrGridHeadWidth = Array(200, 800, 1000, 850, 1200, 1000, 1500, 800, 800, 0, 0, 0)
   'end 2025/04/10
   MSHFlexGrid1.Visible = False
   MSHFlexGrid1.Cols = UBound(arrGridHeadText) + 1
   MSHFlexGrid1.Rows = 2
   For iRow = 0 To MSHFlexGrid1.Cols - 1
      MSHFlexGrid1.row = 0
      MSHFlexGrid1.col = iRow
      MSHFlexGrid1.Text = arrGridHeadText(iRow)
      MSHFlexGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
   Next
   MSHFlexGrid1.Visible = True
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String
Dim strCTB As String, strWhrBase As String, strWhrCP As String, strWhrNp As String 'Add by Amy 2025/04/10
   
   m_blnColOrderAsc = True
   QueryData = True
   MSHFlexGrid1.Clear
   SetGrd
   
   strCon = "F0308='" & Left(Combo1, 5) & "' and F0309 in('" & Flow_已完成 & "','" & Flow_判發重送 & "')"
   If InStr(Combo2.Text, "全部") = 0 Then
      strCon = strCon & " and F0302='" & Trim(Left(Combo2.Text, 2)) & "'"
   End If
   'Add by Amy 2025/04/10 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
   If strSrvDate(1) >= FCP結案單電子化啟用日 Then
      strCTB = ",CloseCaseMain"
      strWhrCP = " And F0301=CCM01 and CCM03 is null and length(CCM02)=9 and CCM02=cp09 "
      strWhrNp = " And F0301=CCM01 and CCM03 is not null and CCM02=NP01(+) and CCM03=NP22(+) "
      strWhrBase = " And F0301=CCM01 and length(CCM02)>=10 " & _
                              " and substr(CCM02,1,length(CCM02)-9)=PA01(+) and substr(CCM02,length(CCM02)-8,6)=PA02(+) " & _
                              " and substr(CCM02,length(CCM02)-2,1)=PA03(+) and substr(CCM02,length(CCM02)-1,2)=PA04(+) "
   Else
      strWhrCP = " and F0304 is null and length(F0303)=9 and F0303=cp09 "
      strWhrNp = " and F0304 is not null and F0303=NP01(+) and F0304=NP22(+) "
      strWhrBase = " and length(F0303)>=10 and substr(F0303,1,length(F0303)-9)=PA01(+) and substr(F0303,length(F0303)-8,6)=PA02(+) and substr(F0303,length(F0303)-2,1)=PA03(+) and substr(F0303,length(F0303)-1,2)=PA04(+) "
   End If
   
   Screen.MousePointer = vbHourglass
   'Modify by Amy 2018/06/19 非P 案結案電子化,加入其他基本檔
   'Modify by Amy 2025/04/10 +FC結案單,f0316,cp01,cp10
   strSql = "select '' V,F0301 表單編號,decode(F0302," & ShowFlow表單類別中文 & ") 表單類別,st02 智權人員,CP01||'-'||CP02||'-'||CP03||'-'||CP04 本所案號,CP09 總收文號,DECODE(PA09,'000',CPM03,Decode(PA09,'1',Nvl(CPM03,CPM04),CPM04)) 案件性質" & _
            ",sqldatet(cp06) 本所期限,sqldatet(cp07) 法定期限,F0316,CP01 as SYSKIND,CP10" & _
            " from (" & _
            "select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,pa09 from flow003,caseprogress,patent " & strCTB & " Where " & strCon & strWhrCP & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) And PA01 is not null" & _
            " Union select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,TM10 as pa09 from flow003,caseprogress,TradeMark " & strCTB & " Where " & strCon & strWhrCP & " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) And TM01 is not null" & _
            " Union select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,LC15 as pa09 from flow003,caseprogress,LawCase " & strCTB & " Where " & strCon & strWhrCP & " and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) And LC01 is not null" & _
            " Union select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,'1' as pa09 from flow003,caseprogress,HireCase " & strCTB & " Where " & strCon & strWhrCP & " and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) And HC01 is not null" & _
            " Union select flow003.*,cp01,cp02,cp03,cp04,cp09,cp10,cp06,cp07,SP09 as pa09 from flow003,caseprogress,ServicePractice " & strCTB & " Where " & strCon & strWhrCP & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) And SP01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,pa09 from flow003,nextprogress,patent " & strCTB & " Where " & strCon & strWhrNp & " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) And Pa01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,TM10 as pa09 from flow003,nextprogress,TradeMark " & strCTB & " Where " & strCon & strWhrNp & " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) And TM01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,LC15 as pa09 from flow003,nextprogress,LawCase " & strCTB & " Where " & strCon & strWhrNp & " and np02=lc01(+) and np03=lc02(+) and np04=lc03(+) and np05=lc04(+) And LC01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,'1' as pa09 from flow003,nextprogress,HireCase " & strCTB & " Where " & strCon & strWhrNp & " and np02=hc01(+) and np03=hc02(+) and np04=hc03(+) and np05=hc04(+) And HC01 is not null" & _
            " Union select flow003.*,np02,np03,np04,np05,np01,np07,np08,np09,SP09 as pa09 from flow003,nextprogress,ServicePractice " & strCTB & " Where " & strCon & strWhrNp & " and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) And SP01 is not null" & _
            " Union select flow003.*,PA01,PA02,PA03,PA04,'','',0,0,pa09 from flow003,patent " & strCTB & " Where " & strCon & strWhrBase & " And PA01 is not null" & _
            " Union select flow003.*,TM01,TM02,TM03,TM04,'','',0,0,TM10 as pa09 from flow003,TradeMark " & strCTB & " Where " & strCon & Replace(UCase(strWhrBase), "PA", "TM") & " And TM01 is not null" & _
            " Union select flow003.*,LC01,LC02,LC03,LC04,'','',0,0,LC15 as pa09 from flow003,LawCase " & strCTB & " Where " & strCon & Replace(UCase(strWhrBase), "PA", "LC") & " And LC01 is not null" & _
            " Union select flow003.*,HC01,HC02,HC03,HC04,'','',0,0,'1' as pa09 from flow003,HireCase " & strCTB & " Where " & strCon & Replace(UCase(strWhrBase), "PA", "HC") & " And HC01 is not null" & _
            " Union select flow003.*,SP01,SP02,SP03,SP04,'','',0,0,SP09 as pa09 from flow003,ServicePractice " & strCTB & " Where " & strCon & Replace(UCase(strWhrBase), "PA", "SP") & " And SP01 is not null" & _
            "),CASEPROPERTYMAP,staff" & _
            " where cp01=cpm01(+) and cp10=cpm02(+) and f0316=st01(+)" & _
            " order by F0301 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   lSelRows = 0
   If rsTmp.RecordCount > 0 Then
      Set MSHFlexGrid1.Recordset = rsTmp
      lTotRows = rsTmp.RecordCount
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      'ShowNoData
      Exit Function
   End If
   lblCount = lSelRows & " / " & lTotRows
   
   '若有資料游標停在第一筆
   MSHFlexGrid1.Visible = False
   MSHFlexGrid1.col = 0
   MSHFlexGrid1.row = 1
   'Modify By Sindy 2015/10/16 游經理說不要預設勾選第一筆
'   If rsTmp.RecordCount > 0 Then
'      MSHFlexGrid1.Text = "V"
'      For i = 0 To MSHFlexGrid1.Cols - 1
'         MSHFlexGrid1.col = i
'         MSHFlexGrid1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   MSHFlexGrid1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'Private Function GetFieldId(pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As Integer
'   Dim iRow As Integer
'   With FlexGrid
'   For iRow = 0 To .Cols - 1
'      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
'         GetFieldId = iRow
'         Exit For
'      End If
'   Next
'   End With
'End Function
'
'Private Function GetValue(pRow As Integer, pFieldName As String) As String
'   Dim iRow As Integer
'   With MSHFlexGrid1
'   For iRow = 0 To .Cols - 1
'      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
'         GetValue = .TextMatrix(pRow, iRow)
'         Exit For
'      End If
'   Next
'   End With
'End Function
'
'Private Function SetValue(pRow As Integer, pFieldName As String, pValue As String) As Boolean
'   Dim iRow As Integer
'   With MSHFlexGrid1
'   For iRow = 0 To .Cols - 1
'      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
'         .TextMatrix(pRow, iRow) = pValue
'         SetValue = True
'         Exit Function
'      End If
'   Next
'   End With
'End Function

Private Sub SetCombo1()
   Combo1.Clear
'   If Pub_StrUserSt03 = "M51" Then
'      Combo1.AddItem "      " & "全部"
'   End If
   Combo1.AddItem strUserNum & " " & strUserName
   'Add By Sindy 2022/12/6
   If strUserNum = "71011" Then
      Combo1.AddItem "99050 " & GetPrjSalesNM("99050")
   End If
   
   '檢查當時是否需要為他人職代
   Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040118 = Nothing
End Sub

'Add by Amy 2018/10/03 原:MSHFlexGrid1_SelChange 第一次點選會無效
Private Sub MSHFlexGrid1_Click()
MSHFlexGrid1.Visible = False
If MSHFlexGrid1.MouseRow <> 0 Then
   MSHFlexGrid1.col = 0
   MSHFlexGrid1.row = MSHFlexGrid1.MouseRow
   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, 1) <> "" Then
      If MSHFlexGrid1.Text = "V" Then
         MSHFlexGrid1.Text = ""
         For i = 0 To MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.col = i
            MSHFlexGrid1.CellBackColor = QBColor(15)
         Next i
         lSelRows = lSelRows - 1
      Else
         MSHFlexGrid1.Text = "V"
         For i = 0 To MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.col = i
            MSHFlexGrid1.CellBackColor = &HFFC0C0
         Next i
         lSelRows = lSelRows + 1
      End If
   End If
End If
lTotRows = MSHFlexGrid1.Rows - 1
lblCount = lSelRows & " / " & lTotRows
MSHFlexGrid1.Visible = True

End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow MSHFlexGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MSHFlexGrid1.col = nCol
   MSHFlexGrid1.row = nRow
   If Me.MSHFlexGrid1.row < 1 And Me.MSHFlexGrid1.Text <> "V" Then
      If Me.MSHFlexGrid1.Text = "表單編號" Then
         If m_blnColOrderAsc = True Then
            Me.MSHFlexGrid1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.MSHFlexGrid1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.MSHFlexGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MSHFlexGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub SetMouseBusy()
   Screen.MousePointer = vbHourglass
   MSHFlexGrid1.MousePointer = vbHourglass
End Sub

Private Sub SetMouseReady()
   Screen.MousePointer = vbDefault
   MSHFlexGrid1.MousePointer = vbDefault
End Sub
