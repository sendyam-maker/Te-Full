VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100120_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "以發明人查詢"
   ClientHeight    =   5750
   ClientLeft      =   240
   ClientTop       =   990
   ClientWidth     =   9320
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   9320
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   4260
      Left            =   30
      TabIndex        =   7
      Top             =   1440
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   7514
      _Version        =   393216
      Cols            =   12
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
      _Band(0).Cols   =   12
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   5
      Left            =   8496
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   70
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料"
      Height          =   400
      Index           =   0
      Left            =   4824
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   70
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度"
      Height          =   400
      Index           =   1
      Left            =   6348
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   4
      Left            =   7572
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   70
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   1590
      TabIndex        =   12
      Top             =   1140
      Width           =   7635
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13467;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   1590
      TabIndex        =   11
      Top             =   870
      Width           =   7635
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13467;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1590
      TabIndex        =   10
      Top             =   600
      Width           =   7635
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13467;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "發明人日文名稱："
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1140
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "發明人英文名稱："
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   870
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "發明人中文名稱："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1440
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   330
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發明人編號："
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   330
      Width           =   1080
   End
End
Attribute VB_Name = "frm100120_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/29 改成Form2.0 ; GrdDataList改字型=新細明體-ExtB、lbl1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Dim strSql  As String, strSQL1 As String, StrSQL6 As String
Dim StrTag As String, i As Integer, j As Integer, s As Integer
Dim Str02 As String, Str03 As String, Str04 As String, Str05 As String, Str06 As String, Str07 As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件


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
'2008/9/16 ADD BY SONIA 畫面加分所號時此處未加
grdDataList.col = 2: grdDataList.Text = "分所號"
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And PUB_GetST06(strUserNum) = "1" Then
    grdDataList.ColWidth(2) = 0
Else
    grdDataList.ColWidth(2) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter
'2008/9/16 END
grdDataList.col = 3: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(3) = 1600
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "申請國家"
grdDataList.ColWidth(4) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "申請案號"
grdDataList.ColWidth(5) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "申請人1"
grdDataList.ColWidth(6) = 1100
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "申請人2"
grdDataList.ColWidth(7) = 1100
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "申請人3"
grdDataList.ColWidth(8) = 1100
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 9: grdDataList.Text = "申請人4"
grdDataList.ColWidth(9) = 1100
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 10: grdDataList.Text = "申請人5"
grdDataList.ColWidth(10) = 1100
grdDataList.CellAlignment = flexAlignCenterCenter
'add by nickc 2005/05/13
grdDataList.col = 11: grdDataList.Text = ""
grdDataList.ColWidth(11) = 0
grdDataList.CellAlignment = flexAlignCenterCenter
'Added by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
For intI = 12 To intField - 1
     grdDataList.col = intI
     grdDataList.ColWidth(intI) = 0
Next intI
'end 2019/11/01
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
      Me.Enabled = False
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
        Dim Str01 As String
        grdDataList.col = 1
        Str01 = SystemNumber(grdDataList, 1)
        If Mid(UCase(Str01), 1, 1) = "N" Then
            Str01 = Mid(Str01, 2, 3)
        End If
        If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Select Case Pub_RplStr(Str01)
            Case "CFP", "FCP", "P"   '專利
                  Screen.MousePointer = vbHourglass
                  frm100101_3.Show
                  frm100101_3.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_3.StrMenu
                  Screen.MousePointer = vbDefault
            Case "CFT", "FCT", "T", "TF"   '商標
                  Screen.MousePointer = vbHourglass
                  frm100101_4.Show
                  frm100101_4.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_4.StrMenu
                  Screen.MousePointer = vbDefault
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            'modify by sonia 2019/7/29 +ACS系統類別
            Case "CFL", "FCL", "L", "LIN", "ACS"   '法務
                  Screen.MousePointer = vbHourglass
                  frm100101_5.Show
                  frm100101_5.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_5.StrMenu
                  Screen.MousePointer = vbDefault
            Case "LA"            '顧問
                  Screen.MousePointer = vbHourglass
                  frm100101_6.Show
                  frm100101_6.Tag = Pub_RplStr(grdDataList.Text)
                  frm100101_6.StrMenu
                  Screen.MousePointer = vbDefault
            Case Else                  '服務
                 Select Case Pub_RplStr(Str01)
                     Case "TB"    '條碼
                         Screen.MousePointer = vbHourglass
                         frm100101_7.Show
                         frm100101_7.Tag = Pub_RplStr(grdDataList.Text)
                         frm100101_7.StrMenu
                         Screen.MousePointer = vbDefault
                     Case "TM"
                         Screen.MousePointer = vbHourglass
                         frm100101_8.Show
                         frm100101_8.Tag = Pub_RplStr(grdDataList.Text)
                         frm100101_8.StrMenu
                         Screen.MousePointer = vbDefault
                     Case "TD"
                         Screen.MousePointer = vbHourglass
                         frm100101_9.Show
                         frm100101_9.Tag = Pub_RplStr(grdDataList.Text)
                         frm100101_9.StrMenu
                         Screen.MousePointer = vbDefault
                     Case "TC", "CFC"
                         Screen.MousePointer = vbHourglass
                         frm100101_A.Show
                         frm100101_A.Tag = Pub_RplStr(grdDataList.Text)
                         frm100101_A.StrMenu
                         Screen.MousePointer = vbDefault
                     Case Else
                         Screen.MousePointer = vbHourglass
                         frm100101_B.Show
                         frm100101_B.Tag = Pub_RplStr(grdDataList.Text)
                         frm100101_B.StrMenu
                         Screen.MousePointer = vbDefault
                  End Select
            End Select
             Me.Enabled = True
             Exit Sub
        End If
        
     End If
     Next i
     Me.Enabled = True
Case 1
     Me.Enabled = False
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
        grdDataList.col = 1
         If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
                Screen.MousePointer = vbHourglass
                frm100101_2.Show
                frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
                frm100101_2.StrMenu
                frm100101_2.cmdOK(0).Enabled = False
                frm100101_2.cmdOK(1).Enabled = False
                Screen.MousePointer = vbDefault
                StrTag = StrTag + grdDataList.Text
            Me.Enabled = True
            Exit Sub
         End If
         
     End If
     Next i
     Me.Enabled = True
Case 4
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 5
     fnCloseAllFrm100
Case Else
End Select
End Sub


Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData

End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
SetDataListWidth
'Call StrMenu
'92.04.16 nick
cmdState = -1
End Sub

Sub StrMenu()
Dim dblRow As Double 'Add By Sindy 2025/9/3

Screen.MousePointer = vbHourglass
grdDataList.Visible = False
Dim Str01 As String, strTemp As Variant
Dim strArr(62) As String, StrOk(32) As String, StrOkTxt(12) As String
Str01 = ""    '發明人編號
Str02 = ""    '系統類別
Str03 = ""    '收文日期(起)
Str04 = ""    '收文日期(迄)
Str05 = ""    '案件性質(起)
Str06 = ""    '案件性質(迄)
Str07 = ""    '是否含來函資料

Me.Tag = Mid(Me.Tag, 1, 8) & Mid(Me.Tag, 10, 2)  '2008/9/2 ADD BY SONIA 取消畫面發明人編號之'-'
Str01 = Me.Tag
pub_QL05 = pub_QL05 & ";發明人編號：" & Str01 & "(案件)" 'Add By Sindy 2025/8/13
'Modify By Cheng 2002/03/14
'Str02 = frm100120_1.Text3
Str02 = IIf(frm100120_1.Text3.Text <> "ALL", frm100120_1.Text3.Text, GetAllSysKind(frm100120_1.Text3))
Str03 = frm100120_1.Text4
'Modify By Cheng 2002/03/18
'Str04 = frm100120_1.Text5
Str04 = IIf(Len(frm100120_1.Text4.Text) > 0 And Len(frm100120_1.Text5.Text) <= 0, ServerDate - 19110000, frm100120_1.Text5.Text)
Str05 = frm100120_1.Text6
Str06 = frm100120_1.Text7
Str07 = frm100120_1.Text8
StrSQL6 = ""
'Added by Lydia 2019/11/01
m_AllSys = Str02
intCufaCnt = 0
'end 2019/11/01

If Len(Str02) <> 0 Then
   StrSQL6 = StrSQL6 & " AND PA01 IN (" & SQLGrpStr(Str02, 1) & ") "
   pub_QL05 = pub_QL05 & ";系統類別：" & Str02 'Add By Sindy 2025/8/13
End If
If Len(Str03) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP05>=" & ChangeTStringToWString(Str03) & " "
End If
If Len(Str04) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP05<=" & ChangeTStringToWString(Str04) & " "
'Add By Cheng 2002/03/18
Else
   If Len(Str03) > 0 Then
      StrSQL6 = StrSQL6 & " AND CP05<=" & ChangeTStringToWString(ServerDate - 19110000) & " "
   End If
End If
'Add By Sindy 2025/8/13
If Len(Str03) <> 0 Or Len(Str04) <> 0 Then
   pub_QL05 = pub_QL05 & ";收文日期：" & Str03 & "-" & Str04
End If
'2025/8/13 END
If Len(Str05) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP10>='" & Str05 & "' "
End If
If Len(Str06) <> 0 Then
   StrSQL6 = StrSQL6 & " AND CP10<='" & Str06 & "' "
End If
'Add By Sindy 2025/8/13
If Len(Str05) <> 0 Or Len(Str06) <> 0 Then
   pub_QL05 = pub_QL05 & ";案件性質：" & Str05 & "-" & Str06
End If
'2025/8/13 END
If UCase(Str07) = "N" Then
    StrSQL6 = StrSQL6 + " and cp09 < 'C' "
    pub_QL05 = pub_QL05 & ";是否含來函資料：不含" 'Add By Sindy 2025/8/13
End If
StrSQL6 = StrSQL6 & " AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) "
'顯示表單上面的值
Label3.Caption = Me.Tag
strSql = "SELECT IN04,IN05,IN06 FROM INVENTOR WHERE IN01='" & Mid(GetNewFagent2(Me.Tag), 1, 8) & "' AND IN02='" & Mid(GetNewFagent2(Me.Tag), 9, 2) & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If IsNull(adoRecordset.Fields(0)) Then
        LBL1(0).Caption = ""
    Else
        LBL1(0).Caption = adoRecordset.Fields(0)
    End If
    If IsNull(adoRecordset.Fields(1)) Then
        LBL1(1).Caption = ""
    Else
        LBL1(1).Caption = adoRecordset.Fields(1)
    End If
    If IsNull(adoRecordset.Fields(2)) Then
        LBL1(2).Caption = ""
    Else
        LBL1(2).Caption = adoRecordset.Fields(2)
    End If
End If
CheckOC
'欲搜尋的SQL字串
'Modify By Cheng 2002/04/26
'若己閉卷, 則在本所案號後加"*"號
'Memo by Lydia 2021/08/17 刪除舊程式碼：專利發明人在專利基本檔60~69
'Modified by Lydia 2019/11/01 利益衝突案件：於Fsort後面，增加申請人1~5,FC代理人
'strSql = " SELECT distinct ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(na03,PA09) AS 申請國家,PA11 AS 申請案號,nvl(NVL(cu1.CU04,DECODE(cu1.cu05,null,cu1.CU06,cu1.cu05||' '||cu1.cu88||' '||cu1.cu89||' '||cu1.cu90)),PA26) AS 申請人1,nvl(NVL(cu2.CU04,DECODE(cu2.cu05,null,cu2.CU06,cu2.cu05||' '||cu2.cu88||' '||cu2.cu89||' '||cu2.cu90)),PA27) AS 申請人2,nvl(NVL(cu3.CU04,DECODE(cu3.cu05,null,cu3.CU06,cu3.cu05||' '||cu3.cu88||' '||cu3.cu89||' '||cu3.cu90)),PA28) AS 申請人3,nvl(NVL(cu4.CU04,DECODE(cu4.cu05,null,cu4.CU06,cu4.cu05||' '||cu4.cu88||' '||cu4.cu89||' '||cu4.cu90)),PA29) AS 申請人4,nvl(NVL(cu5.CU04,DECODE(cu5.cu05,null,cu5.CU06,cu5.cu05||' '||cu5.cu88||' '||cu5.cu89||' '||cu5.cu90)),PA30) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04 as FSort " & _
'         " FROM PATENT,customer cu1,customer cu2,customer cu3,customer cu4,customer cu5,nation,CASEPROGRESS,PATENTInventor WHERE pi06='" & Me.Tag & "' and pi01=pa01 and pi02=pa02 and pi03=pa03 and pi04=pa04 and " & SQLNewFag("pa26", "cU1.cu") & " and " & SQLNewFag("pa27", "cU2.cu") & " and " & SQLNewFag("pa28", "cU3.cu") & " and " & SQLNewFag("pa29", "cU4.cu") & " and " & SQLNewFag("pa30", "cU5.cu") & " and pa09=na01(+)  " & StrSQL6
'2014/11/6 END
strSql = " SELECT distinct ' ' AS V,decode(pa23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●')  AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,nvl(na03,PA09) AS 申請國家,PA11 AS 申請案號" & _
            ",nvl(NVL(cu1.CU04,DECODE(cu1.cu05,null,cu1.CU06,cu1.cu05||' '||cu1.cu88||' '||cu1.cu89||' '||cu1.cu90)),PA26) AS 申請人1,nvl(NVL(cu2.CU04,DECODE(cu2.cu05,null,cu2.CU06,cu2.cu05||' '||cu2.cu88||' '||cu2.cu89||' '||cu2.cu90)),PA27) AS 申請人2,nvl(NVL(cu3.CU04,DECODE(cu3.cu05,null,cu3.CU06,cu3.cu05||' '||cu3.cu88||' '||cu3.cu89||' '||cu3.cu90)),PA28) AS 申請人3,nvl(NVL(cu4.CU04,DECODE(cu4.cu05,null,cu4.CU06,cu4.cu05||' '||cu4.cu88||' '||cu4.cu89||' '||cu4.cu90)),PA29) AS 申請人4,nvl(NVL(cu5.CU04,DECODE(cu5.cu05,null,cu5.CU06,cu5.cu05||' '||cu5.cu88||' '||cu5.cu89||' '||cu5.cu90)),PA30) AS 申請人5" & _
            ",PA01||'-'||PA02||'-'||PA03||'-'||PA04 as FSort,PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO " & _
            " FROM PATENT,customer cu1,customer cu2,customer cu3,customer cu4,customer cu5,nation,CASEPROGRESS,PATENTInventor" & _
            " WHERE pi06='" & Me.Tag & "' and pi01=pa01 and pi02=pa02 and pi03=pa03 and pi04=pa04 and " & SQLNewFag("pa26", "cu1.cu") & " and " & SQLNewFag("pa27", "cu2.cu") & " and " & SQLNewFag("pa28", "cu3.cu") & " and " & SQLNewFag("pa29", "cu4.cu") & " and " & SQLNewFag("pa30", "cu5.cu") & " and pa09=na01(+)  " & StrSQL6
'end 2019/11/01

CheckOC
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
       If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/8/13
       If adoRecordset.RecordCount = 0 Then
             GoTo JumpToNoData
       End If
   Else
       If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/8/13
   End If
   'end 2019/11/01
Else
   If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
JumpToNoData:   'Added by Lydia 2019/11/01
   ShowNoData
   Screen.MousePointer = vbDefault
   Me.Enabled = True
'   Me.Hide
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
Set grdDataList.Recordset = adoRecordset
CheckOC
grdDataList.Visible = True
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100120_2 = Nothing
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

'Sub StrMenu1()
'GrdDataList.Visible = False
'Dim Str01 As String, strTemp As Variant
'Dim StrArr(62) As String, StrOk(32) As String, StrOkTxt(12) As String
'Str01 = ""    '申請人編號
'Str02 = ""    '系統類別
'Str03 = ""    '收文日期(起)
'Str04 = ""    '收文日期(迄)
'Str05 = ""    '案件性質(起)
'Str06 = ""    '案件性質(迄)
'Str07 = ""    '是否含來函資料
'Str01 = Me.Tag
'Str02 = frm100102_1.Text3
'Str03 = frm100102_1.Text4
'Str04 = frm100102_1.Text5
'Str05 = frm100102_1.Text6
'Str06 = frm100102_1.Text7
'Str07 = frm100102_1.Text8
'顯示表單上面的值
'Label3.Caption = Me.Tag

   'If Len(Trim(Me.Tag)) = 9 Then
   '   StrSQL = "SELECT NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),CU13,ST02 FROM CUSTOMER,STAFF WHERE CU01='" & Left$(Me.Tag, 6) & "'"
   'Else
'   strSQL = "SELECT NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),CU13,ST02 FROM CUSTOMER,STAFF WHERE instr(CU01,'" & Left$(Me.Tag, 6) & "')=1 and cu13=st01(+)  "
   'End If
'CheckOC
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    If IsNull(adoRecordset.Fields(0)) Then
'        lbl1(0).Caption = ""
'    Else
'        lbl1(0).Caption = adoRecordset.Fields(0)
'    End If
'    If IsNull(adoRecordset.Fields(1)) Then
'        lbl1(1).Caption = ""
'    Else
'        lbl1(1).Caption = adoRecordset.Fields(1)
'    End If
'    If IsNull(adoRecordset.Fields(2)) Then
'        lbl1(2).Caption = ""
'    Else
'        lbl1(2).Caption = adoRecordset.Fields(2)
'    End If
'End If
'CheckOC
''欲搜尋的SQL字串
'strSQL = "SELECT ' ' AS V,TM01||'-'||TM02||'-'||TM03||'-'||TM04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,TM10 AS 申請國家,TM12 AS 申請案號,TM23 AS 申請人1,' ' AS 申請人2,' ' AS 申請人3,' ' AS 申請人4,' ' AS 申請人5 FROM TRADEMARK WHERE TM23='" & Me.Tag & "' "
'strSQL = strSQL + "union all select ' ' AS V,PA01||'-'||PA02||'-'||PA03||'-'||PA04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,PA09 AS 申請國家,PA11 AS 申請案號,PA26 AS 申請人1,PA27 AS 申請人2,PA28 AS 申請人3,PA29 AS 申請人4,PA30 AS 申請人5 FROM PATENT WHERE PA26='" & Me.Tag & "' OR PA27='" & Me.Tag & "' OR PA28='" & Me.Tag & "' OR PA29='" & Me.Tag & "' OR PA30='" & Me.Tag & "' "
'strSQL = strSQL + "union all select ' ' AS V,SP01||'-'||SP02||'-'||SP03||'-'||SP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,SP09 AS 申請國家,SP11 AS 申請案號,SP08 AS 申請人1,SP58 AS 申請人2,SP59 AS 申請人3,' ' AS 申請人4,' ' AS 申請人5 FROM SERVICEPRACTICE WHERE SP08='" & Me.Tag & "' OR SP58='" & Me.Tag & "' OR SP59='" & Me.Tag & "' "
'strSQL = strSQL + "union all select ' ' AS V,LC01||'-'||LC02||'-'||LC03||'-'||LC04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,LC15 AS 申請國家,' ' AS 申請案號,LC11 AS 申請人1,' ' AS 申請人2,' ' AS 申請人3,' ' AS 申請人4,' ' AS 申請人5 FROM LAWCASE WHERE LC11='" & Me.Tag & "' "
'strSQL = strSQL + "union all select ' ' AS V,HC01||'-'||HC02||'-'||HC03||'-'||HC04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06 AS 案件名稱,' ' AS 申請國家,' ' AS 申請案號,HC05 AS 申請人1,' ' AS 申請人2,' ' AS 申請人3,' ' AS 申請人4,' ' AS 申請人5 FROM HIRECASE WHERE HC05='" & Me.Tag & "' ORDER BY 本所案號"
'CheckOC
'adoRecordset.CursorLocation = adUseClient
'adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    If Len(Trim(Str02)) <> 0 Then
'        strTemp = Split(Str02, ",")
'    End If
'    adoRecordset.MoveFirst
'    Dim StrTest2 As String, StrTest4 As String, S As Integer
 '   Do While adoRecordset.EOF = False
'        StrTest2 = adoRecordset.Fields(1)           '##############     尋找是否符合權限
'        If Len(Trim(Str02)) <> 0 Then
'            StrTest4 = SystemNumber(StrTest2, 1)
'            S = 0
'            If Len(Trim(Str02)) <> 0 Then
'                For i = 0 To UBound(strTemp)
'                    If StrTest4 = strTemp(i) Then
'                        S = 1
'                    End If
'                Next i
'                If S = 0 Then
'                    adoRecordset.Delete
'                End If
'            End If
'        End If
'        adoRecordset.MoveNext
'    Loop
'Else
'   Exit Sub
'End If
'Set GrdDataList.Recordset = adoRecordset
'CheckOC


'For i = 1 To GrdDataList.Rows - 1
'GrdDataList.Row = i
'GrdDataList.Col = 3
''在   GRID   轉換申請國家
'If Not IsNull(GrdDataList.Text) Then
'   strSQL = "SELECT NA03 FROM NATION WHERE NA01='" & GrdDataList.Text & "'"
'   CheckOC
'   adoRecordset.CursorLocation = adUseClient
'   adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'        If Not IsNull(adoRecordset.Fields(0)) Then
'            GrdDataList.Text = adoRecordset.Fields(0)
'        Else
'            GrdDataList.Text = ""
'        End If
'    Else
'        GrdDataList.Text = ""
'    End If
'    CheckOC
'End If
''在   GRID   轉換申請人名稱
'For j = 5 To 9
'GrdDataList.Col = j
'If Len(Trim(GrdDataList.Text)) <> 0 Then
'    If Len(Trim(GrdDataList.Text)) = 9 Then
'        strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06 FROM CUSTOMER WHERE CU01='" & Left$(GrdDataList.Text, 8) & "' AND CU02='" & Right$(GrdDataList.Text, 1) & "'"
'    Else
'        strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06 FROM CUSTOMER WHERE CU01='" & Left$(GrdDataList.Text, 8) & "' AND CU02='0'"
'    End If
'    CheckOC
'    adoRecordset.CursorLocation = adUseClient
'    adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'       If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(2)) Then
'                    GrdDataList.Text = ""
'                Else
'                    GrdDataList.Text = adoRecordset.Fields(2)
'                End If
'            Else
'                GrdDataList.Text = adoRecordset.Fields(1)
'            End If
'        Else
'            GrdDataList.Text = adoRecordset.Fields(0)
'        End If
'    Else
'        GrdDataList.Text = ""
'    End If
'    CheckOC
'End If
'Next j
'Next i
'GrdDataList.Visible = True
'End Sub
