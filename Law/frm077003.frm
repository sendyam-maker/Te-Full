VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm077003 
   BorderStyle     =   1  '單線固定
   Caption         =   "介紹案源管理"
   ClientHeight    =   4620
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9396
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   9396
   Begin VB.CommandButton cmdOK 
      Caption         =   "法律案源接洽單"
      Height          =   405
      Index           =   3
      Left            =   120
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   120
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案源卷宗區(&C)"
      Height          =   405
      Index           =   2
      Left            =   4860
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000C0C0&
      Caption         =   "放棄案件(&Q)"
      Height          =   405
      Index           =   0
      Left            =   7410
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   120
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000C0C0&
      Caption         =   "接洽單(&P)"
      Height          =   405
      Index           =   1
      Left            =   6240
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   1140
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   8580
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "重新整理(&O)"
      Default         =   -1  'True
      Height          =   405
      Left            =   3570
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3645
      Left            =   30
      TabIndex        =   0
      Top             =   660
      Width           =   9285
      _ExtentX        =   16383
      _ExtentY        =   6414
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   12
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "介紹日|管制日期|業務區|介紹人|法務人員|介紹客戶|智慧所案號|案件性質|承辦人|總收文號|接洽單列印日期"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS：接洽單要等櫃檯收文後才不再列出！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   4380
      Width           =   3525
   End
   Begin VB.Label lblCnt 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7590
      TabIndex        =   6
      Top             =   4380
      Width           =   1710
   End
End
Attribute VB_Name = "frm077003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/11 改成Form2.0 ; grdDataList改字型=新細明體-ExtB ; Printer列印未改
'Created by Morgan 2020/4/22
Option Explicit

Dim m_adoRst As ADODB.Recordset
Dim intLastRow As Integer '上一次反白的Row
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

'放棄案源回傳欄位
Public iReturn As Integer, strAbortReason As String

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim strKey As String, StrKey2 As String, strKey3 As String, strFlg As String
   Dim frmTmp As Form
   
   If intLastRow < 1 Then
      MsgBox "尚未點選資料！", vbExclamation
      Exit Sub
   End If
   
   strKey = GetValue(intLastRow, "案源單號", grdDataList)
   strFlg = GetValue(intLastRow, "Flg", grdDataList)
   '先檢查是否已收文或已放棄(其他人也作業，畫面未重整)
   If strFlg = "1" Then
      strExc(0) = "select los06,los07 from lawofficesource where los15='" & strKey & "' and (los06||los07 is not null)"
   Else
      strExc(0) = "select LOS21 as los06,los07 from lawofficesource where los15='" & strKey & "' and (los21||los07 is not null)"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp("los06")) Then
         MsgBox "該案源已收文，案源資料將重新整理！", vbExclamation
      ElseIf Not IsNull(RsTemp("los07")) Then
         MsgBox "該案源已放棄，案源資料將重新整理！", vbExclamation
      End If
      CmdSearch.Value = True
      Exit Sub
   End If
   
   Select Case Index
   Case 0
      If strKey <> "" Then
         frm077003_1.strLOS15 = strKey
         frm077003_1.Show vbModal
         If iReturn = 1 Then
            If grdDataList.Rows = 2 Then
               CmdSearch.Value = True
            Else
               grdDataList.RemoveItem intLastRow
               grdDataList.row = 0
               intLastRow = 0
            End If
         End If
      End If
      
   Case 1 '接洽單
      StrKey2 = GetValue(intLastRow, "los17", grdDataList)
'      'Modify By Sindy 2022/10/3
'      If GetValue(intLastRow, "los12", grdDataList) >= 接洽單電子收文啟用日 Then
'         frm090801_New.SetParent Me
'         frm090801_New.m_SignFlowEmp = strUserNum
'         frm090801_New.m_blnCallPrint = True
'         frm090801_New.Text5 = StrKey2
'         Call frm090801_New.cmdOK_Click(4)
'      Else
'      '2022/10/3 END
         'Modified by Morgan 2020/7/23
         'If PUB_CheckFormExist("frm090801") Then
         '    MsgBox "請先關閉接洽單畫面！"
         '    Exit Sub
         'End If
         'Set frmTmp = Forms(0).GetForm("frm090801")
         Set frmTmp = New frm090801
         'end 2020/7/23
         'With frm090801
         With frmTmp
         .SetParent Me
         '列印
         .Load4Print strKey, StrKey2, True, IIf(strFlg = "2", True, False)
         .m_blnCallPrint_CRL119 = True 'Added by Morgan 2022/4/25 若有特殊收據也要列印
         .Show
         End With
         Set frmTmp = Nothing
'      End If
      
   Case 2 '卷宗區
      StrKey2 = GetValue(intLastRow, "總收文號", grdDataList)
      If StrKey2 <> "" Then
         If PUB_CheckFormExist("frm100101_L") Then
             MsgBox "請先關閉共同查詢〔卷宗區〕畫面！"
             Exit Sub
         End If
         With frm100101_L
         .m_strKey = StrKey2
         .SetParent Me
         If .QueryData = True Then
            .Show
            Me.Hide
         End If
         End With
      End If
   
   'Added by Morgan 2023/2/6
   Case 3 '法律案源接洽單
      StrKey2 = GetValue(intLastRow, "los17", grdDataList)
      strKey3 = DBDATE(GetValue(intLastRow, "介紹日", grdDataList))
      Call PUB_Queryfrm090801(StrKey2, strKey3, Me)
      'Me.Hide 'Modify By Sindy 2023/5/9 Mark
   'end 2023/2/6
   End Select
End Sub

'Add By Sindy 2022/10/3
'Private Sub cmdSearch_Click()
Public Sub cmdSearch_Click()
'2022/10/3 END
   Screen.MousePointer = vbHourglass
   grdDataList.MousePointer = flexHourglass
   Me.Enabled = False
   doQuery
   Me.Enabled = True
   grdDataList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   Static bolActivated As Boolean
   If bolActivated = False Then
      bolActivated = True
      doQuery
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_AddExcuteLog Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm077003 = Nothing
End Sub

Private Sub doQuery()
   Dim ii As Integer, jj As Integer
   Dim stCon As String
   
   'Modified by Morgan 2023/2/6 取消
   'stCon = ""
   'If Check1.Value = vbUnchecked Then
   '   stCon = stCon & " and LOS01 Is Not Null"
   'End If
   stCon = " and LOS01 Is Not Null"
   'end 2023/2/6
   
   '有案源總收文號LOS01(分案確認),無法律所總收文號LOS06(未收文),無放棄日期LOS07(未放棄)
   'Modify By Sindy 2022/10/3 +簽核檔
   strSql = "select sqldatet(los12) 介紹日,sqldatet(los16) 管制日期,a0902 業務區" & _
      ",s1.st02 介紹人,los02 類型,s2.st02 法務人員,NVL(CRA07,CRA08) 介紹客戶,CRL57 介紹內容" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 智慧所案號,c1.cpm03 案件性質" & _
      ",s3.st02 承辦人,s4.st02||' '||sqldatet(to_char(los14,'yyyymmdd'))||to_char(los14,' hh24:mi:ss') 接洽單列印人員時間" & _
      ",cp09 總收文號,NVL(CRL20,0)+NVL(CRL25,0)+NVL(CRL30,0)+NVL(CRL35,0) 法務費用,NVL(CRL21,0)+NVL(CRL26,0)+NVL(CRL31,0)+NVL(CRL36,0) 法務規費" & _
      ",NVL(CRL22,0)+NVL(CRL27,0)+NVL(CRL32,0)+NVL(CRL37,0) 法務點數,CRL06 新案" & _
      ",CRL07||decode(CRL08,'','','-'||CRL08||'-'||CRL09||'-'||CRL10) 法律所案號" & _
      ",c2.cpm03||decode(CRL24,'','','...') 法律所案件性質,los15 案源單號,los02,los12,'' Srt1,cp12,cp13,los17,los18,crl01,'1' Flg" & _
      " from lawofficesource,caseprogress,acc090,ConsultRecordList,ConsultRecApp,casepropertymap c1,casepropertymap c2,staff s1,staff s2,staff s3,staff s4,flow003" & _
      " Where LOS06 Is Null And LOS07 is null" & stCon & _
      " and cp09(+)=los01 and a0901(+)=cp12 and s1.st01(+)=cp13 and s2.st01(+)=los03 and s3.st01(+)=cp14" & _
      " and crl01(+)=los17 and cra01(+)=crl01 and cra02(+)='1' and c1.cpm01(+)=cp01 and c1.cpm02(+)=cp10" & _
      " and c2.cpm01(+)=CRL07 and c2.cpm02(+)=substr(CRL19,1,instr(CRL19,' ')-1) and s4.st01(+)=LOS24" & _
      " and f0301(+)=CRL01 and (f0301 is null or f0308='A4' or f0309='" & Flow_已完成 & "')"
      
   'Modify By Sindy 2022/10/3 +簽核檔
   'Modified by Morgan 2025/4/18 修正法律案接洽單號2相關欄位
   strSql = strSql & " union select sqldatet(los12) 介紹日,sqldatet(los23) 管制日期,a0902 業務區" & _
      ",s1.st02 介紹人,los02 類型,s2.st02 法務人員,NVL(CRA07,CRA08) 介紹客戶,CRL57 介紹內容" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 智慧所案號,c1.cpm03 案件性質" & _
      ",s3.st02 承辦人,s4.st02||' '||sqldatet(to_char(los25,'yyyymmdd'))||to_char(los25,' hh24:mi:ss') 接洽單列印人員時間" & _
      ",cp09 總收文號,NVL(CRL20,0)+NVL(CRL25,0)+NVL(CRL30,0)+NVL(CRL35,0) 法務費用,NVL(CRL21,0)+NVL(CRL26,0)+NVL(CRL31,0)+NVL(CRL36,0) 法務規費" & _
      ",NVL(CRL22,0)+NVL(CRL27,0)+NVL(CRL32,0)+NVL(CRL37,0) 法務點數,CRL06 新案" & _
      ",CRL07||decode(CRL08,'','','-'||CRL08||'-'||CRL09||'-'||CRL10) 法律所案號" & _
      ",c2.cpm03||decode(CRL24,'','','...') 法律所案件性質,los15 案源單號,los02,los12,'' Srt1,cp12,cp13,los20 los17,los18,crl01,'2' Flg" & _
      " from lawofficesource,caseprogress,acc090,ConsultRecordList,ConsultRecApp,casepropertymap c1,casepropertymap c2,staff s1,staff s2,staff s3,staff s4,flow003" & _
      " Where LOS06 Is not Null And LOS07 is null and LOS20 is not null and LOS21 is null" & stCon & _
      " and cp09(+)=los01 and a0901(+)=cp12 and s1.st01(+)=cp13 and s2.st01(+)=los22 and s3.st01(+)=cp14" & _
      " and crl01(+)=los20 and cra01(+)=crl01 and cra02(+)='1' and c1.cpm01(+)=cp01 and c1.cpm02(+)=cp10" & _
      " and c2.cpm01(+)=CRL07 and c2.cpm02(+)=substr(CRL19,1,instr(CRL19,' ')-1) and s4.st01(+)=LOS26" & _
      " and f0301(+)=CRL01 and (f0301 is null or f0308='A4' or f0309='" & Flow_已完成 & "')"
      
   strSql = strSql & " order by los12,cp12,cp13,crl01"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   Set m_adoRst = RsTemp.Clone
   SetRst2Grid
   If RsTemp.RecordCount = 0 Then
      MsgBox "查無資料！", vbInformation
   Else
      With grdDataList
      If .Rows > 1 Then
         .Visible = False
         For ii = 1 To .Rows - 1
            .row = ii
            '有管制期限變黃
            SetRowColor
         Next
         .Visible = True
      End If
      End With
   End If
   
   grdDataList.row = 0
   intLastRow = 0
   SetButton
End Sub

'有管制期限變黃
Private Sub SetRowColor()
   Dim jj As Integer
   With grdDataList
   If .TextMatrix(.row, 1) <> "" Then
      For jj = 0 To .Cols - 1
         .col = jj
         '黃
         .CellBackColor = &HFFFF&
      Next
   End If
   End With
End Sub

Private Sub SetGrid()
   Dim iUnitWidth As Integer
   With grdDataList
      .Visible = False
      .FontFixed.Size = 8
      .Font.Size = 9
      '               0      1        2      3      4    ５       6        7        8　　      9　      10     11　　　　  　         12　　　 13       14       15       16   17         18             19
      .FormatString = "介紹日|管制日期|業務區|介紹人|類型|法務人員|介紹客戶|介紹內容|智慧所案號|案件性質|承辦人|接洽單列印人員時間|總收文號|法務費用|法務規費|法務點數|新案|法律所案號|法律所案件性質|案源單號"
      '設欄寬
      iUnitWidth = .ColWidth(0) / 3
      .ColWidth(0) = iUnitWidth * 4
      .ColWidth(6) = iUnitWidth * 5 '介紹客戶
      .ColWidth(7) = iUnitWidth * 6 '介紹內容
      .ColWidth(11) = iUnitWidth * 10 '接洽單列印人員時間
      .ColWidth(12) = iUnitWidth * 5 '總收文號
      .ColWidth(13) = iUnitWidth * 5 '案源單號
      .ColWidth(17) = iUnitWidth * 5 '法律所案號
      For intI = 0 To .Cols - 1
         If intI < .FixedCols Then
            .ColAlignmentFixed(intI) = 0
         End If
         If intI = 13 Or intI = 14 Or intI = 15 Then
            .ColAlignment(intI) = 7
         Else
            .ColAlignment(intI) = 0
         End If
         '非電腦中心不顯示類型
         If Pub_StrUserSt03 <> "M51" Then
            .ColWidth(4) = 0
         End If
         If intI > 19 Then
            .ColWidth(intI) = 0
         End If
      Next
      .Visible = True
   End With
End Sub

Private Sub SetRst2Grid()
   Dim ii As Integer, jj As Integer
   
   grdDataList.FixedCols = 0
   If m_adoRst.RecordCount > 0 Then
      Set grdDataList.Recordset = m_adoRst
      grdDataList.FixedCols = 4
      LblCnt.Caption = "共 " & m_adoRst.RecordCount & " 筆"
   Else
      grdDataList.Rows = 2
      grdDataList.Clear
      LblCnt.Caption = "共 0 筆"
   End If
   grdDataList.FixedCols = 4
   With grdDataList
   For ii = 1 To .Rows - 1
      .RowHeight(ii) = 255
      .row = ii
      '固定欄位變回底色
      For jj = 0 To .FixedCols - 1
         .col = jj
         .CellBackColor = .BackColor
         .CellFontSize = 9
      Next
   Next
   End With
   SetGrid
End Sub

Private Sub grdDataList_DblClick()
   If m_adoRst.RecordCount = 0 Then Exit Sub
   
   If grdDataList.MouseRow > 0 Then
      If cmdOK(1).Enabled = True Then cmdOK(1).Value = True
   End If
End Sub

Private Function GetValue(pRow As Integer, pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As String
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, iRow)
         Exit For
      End If
   Next
   End With
End Function

Private Function GetFieldId(pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As Integer
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetFieldId = iRow
         Exit For
      End If
   Next
   End With
End Function

'呼叫接洽單畫面用
Public Sub PubShowNextData()
   '更新接洽單列印人員時間
   If intLastRow > 0 Then
      strExc(0) = "select st02||' '||sqldatet(to_char(nvl(los25,los14),'yyyymmdd'))||to_char(los14,' hh24:mi:ss') from lawofficesource,staff where los15='" & GetValue(intLastRow, "案源單號", grdDataList) & "' and st01(+)=nvl(los26,los24)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         intI = GetFieldId("接洽單列印人員時間", grdDataList)
         If intI > 0 Then
            grdDataList.TextMatrix(intLastRow, intI) = "" & RsTemp(0)
         End If
      End If
   End If
End Sub

Private Sub SetButton()
   If intLastRow > 0 Then
      strExc(0) = GetValue(intLastRow, "los02", grdDataList)
      'Modified by Morgan 2020/8/13 改A4也可放棄(經與秀玲確認無PT案者應皆可放棄)
      'If Left(strExc(0), 1) = "A" And strExc(0) <> "A4" Then
      'Modified by Morgan 2022/6/9 +B1類也可放棄
      'If Left(strExc(0), 1) = "A" Then
      If (Left(strExc(0), 1) = "A" Or strExc(0) = "B1") Then
      'end 2022/6/9
         'Modified by Morgan 2022/9/26 +L-888888不可放棄
         strExc(0) = GetValue(intLastRow, "法律所案號", grdDataList)
         If Left(strExc(0), 8) = "L-888888" Then
            cmdOK(0).Enabled = False
         Else
            cmdOK(0).Enabled = True
         End If
         'end 2022/9/26
      Else
         cmdOK(0).Enabled = False
      End If
      cmdOK(1).Enabled = True
      cmdOK(2).Enabled = True
   Else
      cmdOK(0).Enabled = False
      cmdOK(1).Enabled = False
      cmdOK(2).Enabled = False
   End If
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If m_adoRst.RecordCount = 0 Then Exit Sub
   
   Dim nCol As Integer, nRow As Integer, iRow As Integer
   Dim strLOS15 As String, iSrt As Integer
         
   With grdDataList
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow = 0 Then
      '紀錄前次點選的接洽單號
      If intLastRow > 0 Then
         strLOS15 = GetValue(intLastRow, "案源單號", grdDataList)
      End If
      
      '點選欄位更新到拿來排序用的欄位
      iSrt = GetFieldId("Srt1", grdDataList)
      If iSrt > 0 Then
         For iRow = 1 To .Rows - 1
            .TextMatrix(iRow, iSrt) = .TextMatrix(iRow, nCol)
         Next
      End If
      .col = iSrt + 3
      .ColSel = iSrt
      
      If m_blnColOrderAsc = False Then '字串降冪
         .Sort = 5 '字串昇冪
         m_blnColOrderAsc = True
      Else
         .Sort = 6 '字串降冪
         m_blnColOrderAsc = False
      End If
      .col = nCol
      '重設排序後前次點選的位置
      If intLastRow > 0 Then
         For iRow = 1 To .Rows - 1
            If strLOS15 = GetValue(iRow, "案源單號", grdDataList) Then
               intLastRow = iRow
               Exit For
            End If
         Next
      End If
   End If
   End With
End Sub

Private Sub grdDataList_SelChange()
   If m_adoRst.RecordCount = 0 Then Exit Sub
   
   Dim nRow As Integer, iRow As Integer
   
   With grdDataList
   nRow = .MouseRow
   If nRow > 0 Then
      If intLastRow <> nRow Then
         iRow = intLastRow
         ShowBar grdDataList, intLastRow, grdDataList.Cols - 1
         grdDataList.row = intLastRow
         If iRow > 0 Then
            grdDataList.row = iRow
            SetRowColor
            grdDataList.row = intLastRow
         End If
         SetButton
      End If
   End If
   End With
End Sub
