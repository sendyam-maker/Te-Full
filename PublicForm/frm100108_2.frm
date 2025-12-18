VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100108_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "關聯案件資料及正聯商標查詢"
   ClientHeight    =   5710
   ClientLeft      =   4440
   ClientTop       =   3710
   ClientWidth     =   9320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5710
   ScaleWidth      =   9320
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   6564
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   10
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7320
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8520
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   10
      Width           =   756
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找"
      Default         =   -1  'True
      Height          =   400
      Left            =   5808
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   10
      Width           =   756
   End
   Begin VB.OptionButton Option1 
      Caption         =   "審定號數/證書號數"
      Height          =   180
      Index           =   1
      Left            =   3780
      TabIndex        =   2
      Top             =   495
      Width           =   1812
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   450
      Width           =   2412
   End
   Begin VB.TextBox Txt1 
      Height          =   300
      Index           =   1
      Left            =   5640
      TabIndex        =   3
      Top             =   450
      Width           =   2292
   End
   Begin VB.OptionButton Option1 
      Caption         =   "申請案號"
      Height          =   180
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   495
      Width           =   1212
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4860
      Left            =   30
      TabIndex        =   8
      Top             =   795
      Width           =   9255
      _ExtentX        =   16334
      _ExtentY        =   8573
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
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
      _Band(0).Cols   =   12
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "●代表銷卷＊代表閉卷"
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3810
      TabIndex        =   9
      Top             =   210
      Width           =   1830
   End
End
Attribute VB_Name = "frm100108_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
Option Explicit

Dim strSQL1 As String, strSQL2 As String, strSQL5 As String, StrSQL6 As String, StrSQL7 As String, strSQL8 As String
Dim Str1 As String, Str2 As String, intK As Integer, strTemp As String, StrSQL3 As String
Dim strSql As String, i As Integer, j As Integer, s As Integer, StrTag As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Added by Lydia 2019/11/01 利益衝突案件
Dim mESeqNo As String '暫存檔的序號
Dim rsAD As New ADODB.Recordset
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件
Dim colCaseNo As Integer '本所案號欄位
'利益衝突案件：於後面增加欄位
Dim SeColPA As String
Dim SeColTM As String
Dim SeColSP As String
Dim SeColLC As String
Dim SeColHC As String
Dim StrTest1 As String


Private Sub SetDataListWidth()
'Added by Lydia 2019/11/01
Dim intField As Integer
intField = 18
grdDataList.Cols = intField
'end 2019/11/01

grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "本所案號"
grdDataList.ColWidth(1) = 1550
grdDataList.CellAlignment = flexAlignCenterCenter
Dim iDep As String
iDep = PUB_GetST06(strUserNum)
grdDataList.col = 2: grdDataList.Text = "分所號"
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList.ColWidth(2) = 0
Else
    grdDataList.ColWidth(2) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(3) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "申請人"
grdDataList.ColWidth(4) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "相關人"
grdDataList.ColWidth(5) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "申請國家"
grdDataList.ColWidth(6) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "專利商標種類"
grdDataList.ColWidth(7) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "目前准駁"
grdDataList.ColWidth(8) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 9: grdDataList.Text = "專用權是否存在"
grdDataList.ColWidth(9) = 1400
grdDataList.CellAlignment = flexAlignCenterCenter

grdDataList.col = 10: grdDataList.Text = "審定號/專利號數"
grdDataList.ColWidth(10) = 1400
grdDataList.CellAlignment = flexAlignLeftCenter

'Modified by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
'grdDataList.col = 11: grdDataList.Text = ""
'grdDataList.ColWidth(11) = 0
'grdDataList.CellAlignment = flexAlignCenterCenter
For intI = 11 To intField - 1
     grdDataList.col = intI
     grdDataList.ColWidth(intI) = 0
Next intI
If colCaseNo = 0 Then
    colCaseNo = PUB_MGridGetId("本所案號", grdDataList)
End If
'end 2019/11/01
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
     Me.Enabled = False
     Screen.MousePointer = vbHourglass
     StrTag = ""
     For i = 1 To grdDataList.Rows - 1
     grdDataList.col = 0
     grdDataList.row = i
     If Trim(grdDataList.Text) = "V" Then
        grdDataList.col = 0
        grdDataList.Text = ""
        For j = 0 To grdDataList.Cols - 1
             grdDataList.col = j
             grdDataList.CellBackColor = QBColor(15)
        Next j
         'Modified by Lydia 2019/11/01 改成變數
         'grdDataList.col = 10
         'strTemp = SystemNumber(grdDataList.Text, 1)
         'If UCase(Left(strTemp, 1)) = "N" Then
         '   strTemp = Right(strTemp, Len(strTemp) - 1)
         'End If
         grdDataList.col = colCaseNo
         StrTag = Pub_RplStr(grdDataList.Text)
         strTemp = SystemNumber(StrTag, 1)
         'end 2019/11/01
         
         'edit by nick 2004/07/30 加入分割案
         'If ((frm100108_1.Txt1(7) = "2" Or frm100108_1.Txt1(7) = "3") And (strTemp = "CFT" Or strTemp = "FCT" Or strTemp = "T" Or strTemp = "TF")) Or (frm100108_1.Txt1(7) = "1") Then
         If (frm100108_1.Txt1(7) = "1") Then
            'Modified by Lydia 2019/11/01 GrdDataList.Text改成變數StrTag
            If Not IsNull(StrTag) Then
                If fnSaveParentForm(Me) = False Then
                    Me.Enabled = True
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                frm100108_3.Show
                'Modified by Lydia 2019/11/01 GrdDataList.Text改成變數StrTag
                frm100108_3.Tag = Pub_RplStr(StrTag)
                frm100108_3.StrMenu
                Screen.MousePointer = vbDefault
                Me.Enabled = True
                Exit Sub
            End If
          Else
                'add by nick 2004/07/30 加入分割案
                'edit by nick 2004/09/14
                'If frm100108_1.Txt1(7).Text = "4" Or ((frm100108_1.Txt1(7) = "2" Or frm100108_1.Txt1(7) = "3") And (strTemp = "CFT" Or strTemp = "FCT" Or strTemp = "T" Or strTemp = "TF")) Then
                If frm100108_1.Txt1(7).Text = "3" Or ((frm100108_1.Txt1(7) = "2") And (strTemp = "CFT" Or strTemp = "FCT" Or strTemp = "T" Or strTemp = "TF")) Then
                        'Modified by Lydia 2019/11/01 GrdDataList.Text改成變數StrTag
                        If Not IsNull(StrTag) Then
                            If fnSaveParentForm(Me) = False Then
                                Me.Enabled = True
                                Exit Sub
                            End If
                            Screen.MousePointer = vbHourglass
                            frm100108_4.Show
                            frm100108_4.frm100108_txt_7 = frm100108_1.Txt1(7).Text
                            frm100108_4.SetDataListWidth
                            'Modified by Lydia 2019/11/01 GrdDataList.Text改成變數StrTag
                            frm100108_4.Tag = Pub_RplStr(StrTag)
                            frm100108_4.StrMenu
                            Screen.MousePointer = vbDefault
                            Me.Enabled = True
                            Exit Sub
                        End If
                Else
                    'Modified by Lydia 2019/11/01 GrdDataList.Text改成變數StrTag
                    s = MsgBox("此本所案號沒有正聯商標, 無法查詢!!" & StrTag & "  ", , "錯誤")
                    Screen.MousePointer = vbDefault
                    Me.Enabled = True
                    Exit Sub
                End If
          End If
     End If
     Next i
     Me.Enabled = True
     Screen.MousePointer = vbDefault
Case 1
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 2
     fnCloseAllFrm100
Case Else
End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub

End Sub

Private Sub cmdSearch_Click()
grdDataList.Clear
grdDataList.Rows = 2
If Option1(0).Value = True And Len(Trim(Txt1(0))) <> 0 Then
    pub_QL05 = ";申請案號：" & Txt1(0) 'Add By Sindy 2025/9/4
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    StrMenu1 (Txt1(0))
     If grdDataList.Rows = 2 Then
         grdDataList.row = 1
         grdDataList.col = 1
         If Len(grdDataList.Text) <> 0 Then
            grdDataList.col = 0
            grdDataList.Text = "V"
         End If
         Screen.MousePointer = vbDefault
         cmdok_Click (0)
      End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
Else
    If Option1(1).Value = True And Len(Trim(Txt1(1))) <> 0 Then
        pub_QL05 = ";審定號數/證書號數：" & Txt1(1) 'Add By Sindy 2025/9/4
        Screen.MousePointer = vbHourglass
        Me.Enabled = False
        StrMenu2 (Txt1(1))
         If grdDataList.Rows = 2 Then
            grdDataList.row = 1
            grdDataList.col = 1
            If Len(grdDataList.Text) <> 0 Then
               grdDataList.col = 0
               grdDataList.Text = "V"
            End If
            Screen.MousePointer = vbDefault
            cmdok_Click (0)
         End If
        Me.Enabled = True
        Screen.MousePointer = vbDefault
    End If
End If
End Sub

Private Sub Form_Load()
bolToEndByNick = False
   MoveFormToCenter Me
SetDataListWidth
Option1(0).Value = True
'92.04.16 nick
cmdState = -1

 'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
 SeColTM = " ,tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno "
 SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
 SeColSP = " ,sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno "
 'end 2019/11/01
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100108_2 = Nothing
End Sub

Private Sub grdDataList_SelChange()
grdDataList.Visible = False
grdDataList.row = grdDataList.MouseRow
grdDataList.col = 0
If grdDataList.row <> 0 Then
If grdDataList.Text = "V" Then
     grdDataList.Text = ""
     For i = 0 To grdDataList.Cols - 1
          grdDataList.col = i
          grdDataList.CellBackColor = QBColor(15)
    Next i
Else
     grdDataList.Text = "V"
     For i = 0 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = &HFFC0C0
     Next i
End If
End If
grdDataList.Visible = True
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
     If Option1(0).Value = True Then
        'Option1(1).Value = False
        If Me.Enabled = True And Me.Visible = True Then
         Txt1(0).SetFocus
         txt1_GotFocus (0)
        End If
     End If
Case 1
     If Option1(1).Value = True Then
        Option1(0).Value = False
        If Me.Enabled = True And Me.Visible = True Then
         Txt1(1).SetFocus
         txt1_GotFocus (1)
        End If
     End If
Case Else
End Select
End Sub

Sub StrMenu()        '判讀 前畫面所呼叫用
If frm100108_1.Option1(1).Value = True Then  '申請案號
   Me.Enabled = False

   DoEvents
   StrMenu1 (frm100108_1.Txt1(4))
      If grdDataList.Rows = 2 Then
         grdDataList.row = 1
         grdDataList.col = 1
         If Len(grdDataList.Text) <> 0 Then
            grdDataList.col = 0
            grdDataList.Text = "V"
         End If

         cmdok_Click (0)
      End If

   Me.Enabled = True
Else
   '以審定號/證書號查詢
   If frm100108_1.Option1(2).Value = True Then
      Me.Enabled = False

      DoEvents
      'Modify By Cheng 2002/02/25
      If frm100108_1.Txt1(7).Text = 3 Then
         StrMenu3 (frm100108_1.Txt1(5))
      Else
         StrMenu2 (frm100108_1.Txt1(5))
      End If
      If grdDataList.Rows = 2 Then
         grdDataList.row = 1
         grdDataList.col = 1
         If Len(grdDataList.Text) <> 0 Then
            grdDataList.col = 0
            grdDataList.Text = "V"
         End If

         cmdok_Click (0)
      End If

      Me.Enabled = True
   End If
End If
End Sub

Sub StrMenu1(Str1 As String)         '申請案號
Dim dblRow As Double 'Add By Sindy 2025/9/3

Me.Enabled = False
strSQL1 = ""
strSQL2 = ""
strSQL5 = ""
StrSQL6 = ""
StrSQL7 = ""
strSQL8 = ""
'Added by Lydia 2019/11/01 利益衝突案件
m_AllSys = IIf(frm100108_1.Txt1(6) <> "ALL", frm100108_1.Txt1(6), GetAllSysKind(frm100108_1.Txt1(6)))
intCufaCnt = 0
'end 2019/11/01

If Len(Trim(frm100108_1.Txt1(6))) <> 0 Then
   'Modified by Lydia 2019/11/01
'   strSQL1 = strSQL1 & " AND PA01 IN (" & SQLGrpStr(frm100108_1.Txt1(6), 1) & ") "
'   strSQL2 = strSQL2 & " AND TM01 IN (" & SQLGrpStr(frm100108_1.Txt1(6), 2) & ") "
'   StrSQL3 = StrSQL3 & " AND SP01 IN (" & SQLGrpStr(frm100108_1.Txt1(6), 5) & ") "
'   StrSQL6 = StrSQL6 & " AND CP01 IN (" & SQLGrpStr(frm100108_1.Txt1(6), 1) & ") "
'   StrSQL7 = StrSQL7 & " AND CP01 IN (" & SQLGrpStr(frm100108_1.Txt1(6), 2) & ") "
'   strSQL8 = strSQL8 & " AND CP01 IN (" & SQLGrpStr(frm100108_1.Txt1(6), 5) & ") "
   strSQL1 = strSQL1 & " AND PA01 IN (" & SQLGrpStr(m_AllSys, 1) & ") "
   strSQL2 = strSQL2 & " AND TM01 IN (" & SQLGrpStr(m_AllSys, 2) & ") "
   StrSQL3 = StrSQL3 & " AND SP01 IN (" & SQLGrpStr(m_AllSys, 5) & ") "
   StrSQL6 = StrSQL6 & " AND CP01 IN (" & SQLGrpStr(m_AllSys, 1) & ") "
   StrSQL7 = StrSQL7 & " AND CP01 IN (" & SQLGrpStr(m_AllSys, 2) & ") "
   strSQL8 = strSQL8 & " AND CP01 IN (" & SQLGrpStr(m_AllSys, 5) & ") "
   'end 2019/11/01
End If

'Modify By Cheng 2002/02/25
'加審定號/證書號欄位
'strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS N,'" & strUserNum & "' FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM12='" & Str1 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
'strSQL = strSQL + " union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,DeCODE(Pa09,'000',PTM03,PTM04) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS N,'" & strUserNum & "' FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA11='" & Str1 & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1
'strSQL = strSQL + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS N,'" & strUserNum & "' FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP11='" & Str1 & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & StrSQL5
'
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "'  FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str1 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,pa09) AS 申請國家,DeCODE(PA09,'000',PTM03,PTM04) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "'  FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str1 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "'  FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP36='" & Str1 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & StrSQL8
'
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP30='" & Str1 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,DeCODE(PA09,'000',PTM03,PTM04) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "' FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP30='" & Str1 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP30='" & Str1 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & StrSQL8 & "  "

'Modify By Cheng 2002/04/25
'若已閉卷, 則在本所案號後加"*"號
'edit by nickc 2005/05/13
'strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS N,'" & strUserNum & "',TM15 AS 審定專利號數 FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM12='" & Str1 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
'strSQL = strSQL + " union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(Pa09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS N,'" & strUserNum & "',PA22 AS 審定專利號數 FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA11='" & Str1 & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1
'strSQL = strSQL + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS N,'" & strUserNum & "',SP14 AS 審定專利號數 FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP11='" & Str1 & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL5
'
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "',TM15 AS 審定專利號數 FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str1 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "',PA22 AS 審定專利號數 FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str1 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "',SP14 AS 審定專利號數 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP36='" & Str1 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL8
'
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "',TM15 AS 審定專利號數 FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP30='" & Str1 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "',PA22 AS 審定專利號數 FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP30='" & Str1 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "',SP14 AS 審定專利號數 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP30='" & Str1 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL8 & "  "

'Modified by Lydia 2019/11/01 增加欄位SeColTM, SeColPA, SeColSP, 並且改用Rdatafactory暫存檔
'strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS N,TM15 AS 審定專利號數,'" & strUserNum & "' FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM12='" & Str1 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
'strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(Pa09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS N,PA22 AS 審定專利號數,'" & strUserNum & "' FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA11='" & Str1 & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1
'strSql = strSql + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS N,SP14 AS 審定專利號數,'" & strUserNum & "' FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP11='" & Str1 & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL5
'
'strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,TM15 AS 審定專利號數,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str1 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
'strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,PA22 AS 審定專利號數,'" & strUserNum & "' FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str1 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
'strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,SP14 AS 審定專利號數,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP36='" & Str1 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL8
'
'strSql = strSql + " union all select ' ' AS V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,TM15 AS 審定專利號數,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP30='" & Str1 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
'strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,PA22 AS 審定專利號數,'" & strUserNum & "' FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP30='" & Str1 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
'strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,SP14 AS 審定專利號數,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP30='" & Str1 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL8 & "  "
'
'cnnConnection.Execute "DELETE FROM R100108 WHERE ID='" & strUserNum & "' "
'cnnConnection.Execute "INSERT INTO R100108 " & strSql
'cnnConnection.Execute "DELETE FROM R100108 WHERE R001002 IN (SELECT DISTINCT 'N'||R001002 FROM R100108 WHERE SUBSTR(R001002,1,1)<>'N' and id='" & strUserNum & "' ) and id='" & strUserNum & "' "
''Modify By Cheng 2002/02/25
''strSQL = "SELECT R001001,R001002,R001003,R001004,R001005,R001006,R001007,R001008,R001009,R001010 FROM R100108 WHERE ID='" & strUserNum & "' order by r001010 "
'strSql = "SELECT R001001,R001002,R001003,R001004,R001005,R001006,R001007,R001008,R001009,R001010,R001012,ID FROM R100108 WHERE ID='" & strUserNum & "' order by r001011 "
'-----申請案號
strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM15 AS 審定專利號數,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS CaseNo" & SeColTM & _
                " FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM12='" & Str1 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(Pa09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,PA22 AS 審定專利號數,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS CaseNo" & SeColPA & _
                " FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA11='" & Str1 & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1
strSql = strSql + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,SP14 AS 審定專利號數,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS CaseNo" & SeColSP & _
                " FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP11='" & Str1 & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL5
'-----對造
strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM15 AS 審定專利號數,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CaseNo" & SeColTM & _
                " FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str1 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,PA22 AS 審定專利號數,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CaseNo" & SeColPA & _
                " FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str1 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,SP14 AS 審定專利號數,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CaseNo" & SeColSP & _
                " FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP36='" & Str1 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL8
'-----對方案件號數
strSql = strSql + " union all select ' ' AS V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM15 AS 審定專利號數,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CaseNo" & SeColTM & _
                " FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP30='" & Str1 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,PA22 AS 審定專利號數,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CaseNo" & SeColPA & _
                " FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP30='" & Str1 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,SP14 AS 審定專利號數,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CaseNo" & SeColSP & _
                " FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP30='" & Str1 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL8 & "  "
intK = 1
Set RsTemp = ClsLawReadRstMsg(intK, strSql)
If intK = 1 Then
    '丟到暫存檔
    Set rsAD = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
    StrTest1 = "delete from rdatafactory where formname='" & Me.Name & "' and id = '" & strUserNum & "' and seqno = '" & mESeqNo & "' "
    StrTest1 = StrTest1 & "and r002 in (select distinct 'N'||r002 from rdatafactory where formname='" & Me.Name & "' and id = '" & strUserNum & "' and seqno = '" & mESeqNo & "' and substr(r002,1,1)<>'N') "
    cnnConnection.Execute StrTest1, intK
    strSql = "SELECT R001 as V,R002 as 本所案號,R003 as 分所號,R004 as 案件名稱,R005 as 申請人,R006 as 相關人,R007 as 申請國家," & _
                "R008 as 專利商標種類,R009 as 目前准駁,R010 as 專用權是否存在,R011 as 審定專利號,R012 as CaseNo," & _
                "R013 as cust01,R014 as cust02,R015 as cust03,R016 as cust04,R017 as cust05,R018 as fcno " & _
                "FROM RDATAFACTORY where formname='" & Me.Name & "' and id = '" & strUserNum & "' and seqno = '" & mESeqNo & "' "
    Set rsAD = Nothing
Else
    If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/9/4
    GoTo JumpToNoData
End If
'end 2019/11/01

CheckOC
j = 0
adoRecordset.CursorLocation = adUseClient
'Modified by Lydia 2019/11/01 改變型態
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic

If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
    'Added by Lydia 2019/11/01 逐案號判斷
    If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            '利益衝突案件：逐案號判斷
            If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                intCufaCnt = intCufaCnt + 1
                adoRecordset.Delete
            End If
            adoRecordset.MoveNext
        Loop
        '利益衝突案件：限閱案件
        If intCufaCnt > 0 Then
            pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
        End If
        InsertQueryLog (dblRow) 'Add By Sindy 2010/11/3
        If adoRecordset.RecordCount = 0 Then
              GoTo JumpToNoData
        End If
    Else
        InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/3
    End If
    'end 2019/11/01
   
Else
   InsertQueryLog (0) 'Add By Sindy 2010/11/3
JumpToNoData:   'Added by Lydia 2019/11/01
   Me.Enabled = True
   cmdOK(0).Enabled = False
   ShowNoData
   Screen.MousePointer = vbDefault
   '92.04.18 nick
   'Me.Hide
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
Set grdDataList.Recordset = adoRecordset

SetDataListWidth
CheckOC
Screen.MousePointer = vbDefault
Me.Enabled = True
End Sub

Sub StrMenu2(Str2 As String)       '審定號數查正聯商標
Dim dblRow As Double 'Add By Sindy 2025/9/3

Me.Enabled = False
strSQL1 = ""
strSQL2 = ""
strSQL5 = ""
StrSQL6 = ""
StrSQL7 = ""
strSQL8 = ""
'Added by Lydia 2019/11/01 利益衝突案件
m_AllSys = IIf(frm100108_1.Txt1(6) <> "ALL", frm100108_1.Txt1(6), GetAllSysKind(frm100108_1.Txt1(6)))
intCufaCnt = 0
'end 2019/11/01

If Len(Trim(frm100108_1.Txt1(6))) <> 0 Then
   'Modified by Lydia 2019/11/01
'   strSQL1 = strSQL1 & " AND PA01 IN (" & SQLGrpStr(frm100108_1.txt1(6), 1) & ") "
'   strSQL2 = strSQL2 & " AND TM01 IN (" & SQLGrpStr(frm100108_1.txt1(6), 2) & ") "
'   StrSQL3 = StrSQL3 & " AND SP01 IN (" & SQLGrpStr(frm100108_1.txt1(6), 5) & ") "
'   StrSQL6 = StrSQL6 & " AND CP01 IN (" & SQLGrpStr(frm100108_1.txt1(6), 1) & ") "
'   StrSQL7 = StrSQL7 & " AND CP01 IN (" & SQLGrpStr(frm100108_1.txt1(6), 2) & ") "
'   strSQL8 = strSQL8 & " AND CP01 IN (" & SQLGrpStr(frm100108_1.txt1(6), 5) & ") "
   strSQL1 = strSQL1 & " AND PA01 IN (" & SQLGrpStr(m_AllSys, 1) & ") "
   strSQL2 = strSQL2 & " AND TM01 IN (" & SQLGrpStr(m_AllSys, 2) & ") "
   StrSQL3 = StrSQL3 & " AND SP01 IN (" & SQLGrpStr(m_AllSys, 5) & ") "
   StrSQL6 = StrSQL6 & " AND CP01 IN (" & SQLGrpStr(m_AllSys, 1) & ") "
   StrSQL7 = StrSQL7 & " AND CP01 IN (" & SQLGrpStr(m_AllSys, 2) & ") "
   strSQL8 = strSQL8 & " AND CP01 IN (" & SQLGrpStr(m_AllSys, 5) & ") "
   'end 2019/11/01
End If

'Modify By Cheng 2002/02/25
'加審定號/專利號數欄
'strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS N,'" & strUserNum & "' FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM15='" & Str2 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
'strSQL = strSQL + " union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,DeCODE(PA09,'000',PTM03,PTM04) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS N,'" & strUserNum & "' FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA22='" & Str2 & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1
'strSQL = strSQL + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS N,'" & strUserNum & "' FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP14='" & Str2 & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & StrSQL5
'
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "'  FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str2 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,pa09) AS 申請國家,DeCODE(PA09,'000',PTM03,PTM04) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "'  FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str2 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "'  FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP36='" & Str2 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & StrSQL8
'Modify By Cheng 2002/04/25
'若已閉卷, 則本所案號後加"＊"號
'edit by nickc 2005/05/13
'strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS N,'" & strUserNum & "',TM15 AS 審定專利號數 FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM15='" & Str2 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
'strSQL = strSQL + " union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS N,'" & strUserNum & "',PA22 AS 審定專利號數 FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA22='" & Str2 & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1
'strSQL = strSQL + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS N,'" & strUserNum & "',SP14 AS 審定專利號數 FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP14='" & Str2 & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL5
'
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "',TM15 AS 審定專利號數 FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str2 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "',PA22 AS 審定專利號數 FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str2 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
'strSQL = strSQL + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,'" & strUserNum & "',SP14 AS 審定專利號數 FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP36='" & Str2 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL8
'Modified by Lydia 2019/11/01 增加欄位SeColTM, SeColPA, SeColSP, 並且改用Rdatafactory暫存檔
'strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS N,TM15 AS 審定專利號數,'" & strUserNum & "' FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM15='" & Str2 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
'strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS N,PA22 AS 審定專利號數,'" & strUserNum & "' FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA22='" & Str2 & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1
'strSql = strSql + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS N,SP14 AS 審定專利號數,'" & strUserNum & "' FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP14='" & Str2 & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL5
'
'strSql = strSql + " union all select ' ' AS V,decode(tm28,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,TM15 AS 審定專利號數,'" & strUserNum & "' FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str2 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
'strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,PA22 AS 審定專利號數,'" & strUserNum & "' FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str2 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
'strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS N,SP14 AS 審定專利號數,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP36='" & Str2 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL8
'
'cnnConnection.Execute "DELETE FROM R100108 WHERE ID='" & strUserNum & "' "
'cnnConnection.Execute "INSERT INTO R100108 " & strSql
'cnnConnection.Execute "DELETE FROM R100108 WHERE R001002 IN (SELECT DISTINCT 'N'||R001002 FROM R100108 WHERE SUBSTR(R001002,1,1)<>'N' and id='" & strUserNum & "' ) and id='" & strUserNum & "' "
'
''Modify By Cheng 2002/02/25
''strSQL = "SELECT R001001,R001002,R001003,R001004,R001005,R001006,R001007,R001008,R001009,R001010 FROM R100108 WHERE ID='" & strUserNum & "' order by r001010 "
'strSql = "SELECT R001001,R001002,R001003,R001004,R001005,R001006,R001007,R001008,R001009,R001010,R001012,ID FROM R100108 WHERE ID='" & strUserNum & "' order by r001011 "
'-----審定號/證書號
strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM15 AS 審定專利號數,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS CaseNo" & SeColTM & _
                " FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM15='" & Str2 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
strSql = strSql + " union all select ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(Pa09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,PA22 AS 審定專利號數,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS CaseNo" & SeColPA & _
                " FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA22='" & Str2 & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1
strSql = strSql + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,SP14 AS 審定專利號數,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS CaseNo" & SeColSP & _
                " FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP14='" & Str2 & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL5
'-----對造
strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM15 AS 審定專利號數,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CaseNo" & SeColTM & _
                " FROM CASEPROGRESS,TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str2 & "' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & StrSQL7
strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,pa09) AS 申請國家,decode(pa01,'CFP',ptm03,DeCODE(PA09,'000',PTM03,PTM04)) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,PA22 AS 審定專利號數,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CaseNo" & SeColPA & _
                " FROM CASEPROGRESS,PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE CP36='" & Str2 & "' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & StrSQL6
strSql = strSql + " union all select ' ' AS V,'N'||CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(CP37,NVL(CP38,CP39)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP08) AS 申請人,NVL(CP40,NVL(CP41,CP42)) AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,SP14 AS 審定專利號數,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CaseNo" & SeColSP & _
                " FROM CASEPROGRESS,SERVICEPRACTICE,CUSTOMER,NATION WHERE CP36='" & Str2 & "' AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & strSQL8
intK = 1
Set RsTemp = ClsLawReadRstMsg(intK, strSql)
If intK = 1 Then
    '丟到暫存檔
    Set rsAD = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
    StrTest1 = "delete from rdatafactory where formname='" & Me.Name & "' and id = '" & strUserNum & "' and seqno = '" & mESeqNo & "' "
    StrTest1 = StrTest1 & "and r002 in (select distinct 'N'||r002 from rdatafactory where formname='" & Me.Name & "' and id = '" & strUserNum & "' and seqno = '" & mESeqNo & "' and substr(r002,1,1)<>'N') "
    cnnConnection.Execute StrTest1, intK
    strSql = "SELECT R001 as V,R002 as 本所案號,R003 as 分所號,R004 as 案件名稱,R005 as 申請人,R006 as 相關人,R007 as 申請國家," & _
                "R008 as 專利商標種類,R009 as 目前准駁,R010 as 專用權是否存在,R011 as 審定專利號,R012 as CaseNo," & _
                "R013 as cust01,R014 as cust02,R015 as cust03,R016 as cust04,R017 as cust05,R018 as fcno " & _
                "FROM RDATAFACTORY where formname='" & Me.Name & "' and id = '" & strUserNum & "' and seqno = '" & mESeqNo & "' "
    Set rsAD = Nothing
Else
    If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/9/4
    GoTo JumpToNoData
End If
'end 2019/11/01

CheckOC
j = 0
adoRecordset.CursorLocation = adUseClient
'Modified by Lydia 2019/11/01 改變型態
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic

If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3

    'Added by Lydia 2019/11/01 逐案號判斷
    If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            '利益衝突案件：逐案號判斷
            If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                intCufaCnt = intCufaCnt + 1
                adoRecordset.Delete
            End If
            adoRecordset.MoveNext
        Loop
        '利益衝突案件：限閱案件
        If intCufaCnt > 0 Then
            pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
        End If
        InsertQueryLog (dblRow) 'Add By Sindy 2010/11/3
        If adoRecordset.RecordCount = 0 Then
              GoTo JumpToNoData
        End If
    Else
        InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/3
    End If
    'end 2019/11/01
    
    cmdOK(0).Enabled = True
Else
    InsertQueryLog (0) 'Add By Sindy 2010/11/3
JumpToNoData:   'Added by Lydia 2019/11/01
    cmdOK(0).Enabled = False
    Me.Enabled = True
    ShowNoData
    Screen.MousePointer = vbDefault
    '92.04.18 nick
    'Me.Hide
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Set grdDataList.Recordset = adoRecordset
SetDataListWidth
CheckOC
Me.Enabled = True
End Sub

Sub StrMenu3(Str2 As String)       '審定號數/證書號查分割案
Dim dblRow As Double 'Add By Sindy 2025/9/3

Me.Enabled = False
strSQL1 = ""
strSQL2 = ""
strSQL5 = ""
StrSQL6 = ""
StrSQL7 = ""
strSQL8 = ""
'Added by Lydia 2019/11/01 利益衝突案件
m_AllSys = IIf(frm100108_1.Txt1(6) <> "ALL", frm100108_1.Txt1(6), GetAllSysKind(frm100108_1.Txt1(6)))
intCufaCnt = 0
'end 2019/11/01

If Len(Trim(frm100108_1.Txt1(6))) <> 0 Then
   'Modified by Lydia 2019/11/01
'   strSQL1 = strSQL1 & " AND PA01 IN (" & SQLGrpStr(frm100108_1.Txt1(6), 1) & ") "
'   strSQL2 = strSQL2 & " AND TM01 IN (" & SQLGrpStr(frm100108_1.Txt1(6), 2) & ") "
'   StrSQL3 = StrSQL3 & " AND SP01 IN (" & SQLGrpStr(frm100108_1.Txt1(6), 5) & ") "
'   StrSQL6 = StrSQL6 & " AND CP01 IN (" & SQLGrpStr(frm100108_1.Txt1(6), 1) & ") "
'   StrSQL7 = StrSQL7 & " AND CP01 IN (" & SQLGrpStr(frm100108_1.Txt1(6), 2) & ") "
'   strSQL8 = strSQL8 & " AND CP01 IN (" & SQLGrpStr(frm100108_1.Txt1(6), 5) & ") "
   strSQL1 = strSQL1 & " AND PA01 IN (" & SQLGrpStr(m_AllSys, 1) & ") "
   strSQL2 = strSQL2 & " AND TM01 IN (" & SQLGrpStr(m_AllSys, 2) & ") "
   StrSQL3 = StrSQL3 & " AND SP01 IN (" & SQLGrpStr(m_AllSys, 5) & ") "
   StrSQL6 = StrSQL6 & " AND CP01 IN (" & SQLGrpStr(m_AllSys, 1) & ") "
   StrSQL7 = StrSQL7 & " AND CP01 IN (" & SQLGrpStr(m_AllSys, 2) & ") "
   strSQL8 = strSQL8 & " AND CP01 IN (" & SQLGrpStr(m_AllSys, 5) & ") "
   'end 2019/11/01
End If

'strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS N,'" & strUserNum & "',TM15 AS 審定專利號數 FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM15='" & Str2 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
'strSQL = strSQL + " union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),PA26) AS 申請人,'' AS 相關人,nvl(NA03,pa09) AS 申請國家,DeCODE(PA09,'000',PTM03,PTM04) AS 專利商標種類,DECODE(PA16,'1','准','2','駁','') AS 目前准駁,DECODE(PA17,'Y','是','N','否','') AS 專用權是否存在,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS N,'" & strUserNum & "',PA22 AS 審定專利號數 FROM PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA22='" & Str2 & "' AND " & SQLNewFag("PA26", "CU") & " AND PA09=NA01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strSQL1
'strSQL = strSQL + " union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,'' AS 相關人,nvl(NA03,sp09) AS 申請國家,'' AS 專利商標種類,'' AS 目前准駁,'' AS 專用權是否存在,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS N,'" & strUserNum & "',SP14 AS 審定專利號數 FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE SP14='" & Str2 & "' AND " & SQLNewFag("SP08", "CU") & " AND SP09=NA01(+) " & StrSQL5

'Modify By Cheng 2002/04/23
'strSQL = "SELECT '' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,'' AS 相關人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS N,'" & strUserNum & "',TM15 AS 審定專利號數 " & _
'         " FROM TRADEMARK,nation,customer,PATENTTRADEMARKMAP " & _
'         " WHERE TM27='" & Str2 & "' and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
'Modify By Cheng 2002/04/25
'edit by  nickc 2005/05/13
'strSQL = "SELECT '' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,'' AS 相關人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS N,'" & strUserNum & "',TM15 AS 審定專利號數 " & _
         " FROM TRADEMARK,nation,customer,PATENTTRADEMARKMAP " & _
         " WHERE TM15='" & Str2 & "' and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
'Modified by Lydia 2019/11/01 增加欄位SeColTM, 並且改用Rdatafactory暫存檔
'strSql = "SELECT '' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,'' AS 相關人,nvl(na03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS N,TM15 AS 審定專利號數,'" & strUserNum & "' " & _
'         " FROM TRADEMARK,nation,customer,PATENTTRADEMARKMAP " & _
'         " WHERE TM15='" & Str2 & "' and tm10=na01(+) and '2'=ptm01(+) and tm08=ptm02(+) and " & SQLNewFag("tm23", "cu") & " " & strSQL2
'
'cnnConnection.Execute "DELETE FROM R100108 WHERE ID='" & strUserNum & "' "
'cnnConnection.Execute "INSERT INTO R100108 " & strSql
'cnnConnection.Execute "DELETE FROM R100108 WHERE R001002 IN (SELECT DISTINCT 'N'||R001002 FROM R100108 WHERE SUBSTR(R001002,1,1)<>'N' and id='" & strUserNum & "' ) and id='" & strUserNum & "' "
'
''Modify By Cheng 2002/02/25
''strSQL = "SELECT R001001,R001002,R001003,R001004,R001005,R001006,R001007,R001008,R001009,R001010 FROM R100108 WHERE ID='" & strUserNum & "' order by r001010 "
'strSql = "SELECT R001001,R001002,R001003,R001004,R001005,R001006,R001007,R001008,R001009,R001010,R001012,ID FROM R100108 WHERE ID='" & strUserNum & "' order by r001011 "
'-----審定號
strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,'' AS 相關人,nvl(NA03,tm10) AS 申請國家,DeCODE(TM10,'000',PTM03,PTM04) AS 專利商標種類,DECODE(TM16,'1','准','2','駁','') AS 目前准駁,DECODE(TM17,'Y','是','N','否','') AS 專用權是否存在,TM15 AS 審定專利號數,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS CaseNo" & SeColTM & _
                " FROM TRADEMARK,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE TM15='" & Str2 & "' AND " & SQLNewFag("TM23", "CU") & " AND TM10=NA01(+) AND '2'=PTM01(+) AND TM08=PTM02(+) " & strSQL2
intK = 1
Set RsTemp = ClsLawReadRstMsg(intK, strSql)
If intK = 1 Then
    '丟到暫存檔
    Set rsAD = PUB_CreateRecordset(RsTemp, , , , Me.Name, mESeqNo)
    StrTest1 = "delete from rdatafactory where formname='" & Me.Name & "' and id = '" & strUserNum & "' and seqno = '" & mESeqNo & "' "
    StrTest1 = StrTest1 & "and r002 in (select distinct 'N'||r002 from rdatafactory where formname='" & Me.Name & "' and id = '" & strUserNum & "' and seqno = '" & mESeqNo & "' and substr(r002,1,1)<>'N') "
    cnnConnection.Execute StrTest1, intK
    strSql = "SELECT R001 as V,R002 as 本所案號,R003 as 分所號,R004 as 案件名稱,R005 as 申請人,R006 as 相關人,R007 as 申請國家," & _
                "R008 as 專利商標種類,R009 as 目前准駁,R010 as 專用權是否存在,R011 as 審定專利號,R012 as CaseNo," & _
                "R013 as cust01,R014 as cust02,R015 as cust03,R016 as cust04,R017 as cust05,R018 as fcno " & _
                "FROM RDATAFACTORY where formname='" & Me.Name & "' and id = '" & strUserNum & "' and seqno = '" & mESeqNo & "' "
    Set rsAD = Nothing
Else
    GoTo JumpToNoData
End If
'end 2019/11/01

CheckOC
j = 0
adoRecordset.CursorLocation = adUseClient
'Modified by Lydia 2019/11/01 改變型態
'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic

If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3

    'Added by Lydia 2019/11/01 逐案號判斷
    If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
        adoRecordset.MoveFirst
        Do While adoRecordset.EOF = False
            '利益衝突案件：逐案號判斷
            If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("cust01") & "," & adoRecordset.Fields("cust02") & "," & adoRecordset.Fields("cust03") & "," & adoRecordset.Fields("cust04") & "," & adoRecordset.Fields("cust05"), "" & adoRecordset.Fields("fcno")) = False Then
                intCufaCnt = intCufaCnt + 1
                adoRecordset.Delete
            End If
            adoRecordset.MoveNext
        Loop
        '利益衝突案件：限閱案件
        If intCufaCnt > 0 Then
            pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
        End If
        InsertQueryLog (dblRow) 'Add By Sindy 2010/11/3
        If adoRecordset.RecordCount = 0 Then
              GoTo JumpToNoData
        End If
    Else
        InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/3
    End If
    'end 2019/11/01
    
    cmdOK(0).Enabled = True
Else
   InsertQueryLog (0) 'Add By Sindy 2010/11/3
JumpToNoData:   'Added by Lydia 2019/11/01
    cmdOK(0).Enabled = False
    Me.Enabled = True
    ShowNoData
    Screen.MousePointer = vbDefault
    '92.04.18 nick
    'Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Set grdDataList.Recordset = adoRecordset
SetDataListWidth
CheckOC
Me.Enabled = True
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))

End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
Case 0
    Option1(0).Value = True
Case 1
    Option1(1).Value = True
Case Else
End Select

End Sub
