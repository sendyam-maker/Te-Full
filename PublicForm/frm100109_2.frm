VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100109_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "以收文日查詢來函"
   ClientHeight    =   5740
   ClientLeft      =   60
   ClientTop       =   2280
   ClientWidth     =   9320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5740
   ScaleWidth      =   9320
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7308
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度(&C)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   6084
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   0
      Left            =   4560
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   10
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   8532
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   10
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5250
      Left            =   90
      TabIndex        =   4
      Top             =   450
      Width           =   9135
      _ExtentX        =   16104
      _ExtentY        =   9243
      _Version        =   393216
      Cols            =   13
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
      _Band(0).Cols   =   13
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：＊閉卷●銷卷"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   10
      Left            =   2550
      TabIndex        =   5
      Top             =   210
      Width           =   1980
   End
End
Attribute VB_Name = "frm100109_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/20 改成Form2.0(grdDataList改Fonts)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer, s As Integer, intK As Integer
Dim strSql As String, strTemp As Variant, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件

Private Sub SetDataListWidth()
'Modified by Lydia 2019/11/01
'GrdDataList.Cols = 15
Dim intField As Integer
intField = 21
grdDataList.Cols = intField
'end 2019/11/01

grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "收文日"
grdDataList.ColWidth(1) = 810
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "本所案號"
'Modified by Lydia 2015/10/14
'grdDataList.ColWidth(2) = 1550
grdDataList.ColWidth(2) = 1450
grdDataList.CellAlignment = flexAlignCenterCenter
Dim iDep As String
iDep = PUB_GetST06(strUserNum)
grdDataList.col = 3: grdDataList.Text = "分所號"
'電腦中心，跟分所才秀
If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
    grdDataList.ColWidth(3) = 0
Else
    grdDataList.ColWidth(3) = 620
End If
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(4) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "來函性質"
grdDataList.ColWidth(5) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "申請人"
grdDataList.ColWidth(6) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "承辦人"
'Modified by Lydia 2015/10/14
'grdDataList.ColWidth(7) = 800
grdDataList.ColWidth(7) = 720
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "智權人員"
'Modified by Lydia 2015/10/14
'grdDataList.ColWidth(8) = 800
grdDataList.ColWidth(8) = 720
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 9: grdDataList.Text = "下一程序"
grdDataList.ColWidth(9) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 10: grdDataList.Text = "本所期限"
grdDataList.ColWidth(10) = 810
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 11: grdDataList.Text = "法定期限"
grdDataList.ColWidth(11) = 810
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 12: grdDataList.Text = "申請國家"
grdDataList.ColWidth(12) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
'Add By Sindy 2011/3/18
grdDataList.col = 13: grdDataList.Text = "相關收文號發文日"
grdDataList.ColWidth(13) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
'Add By Cheng 2003/08/15
grdDataList.col = 14: grdDataList.Text = "CP09"
grdDataList.ColWidth(14) = 0
grdDataList.CellAlignment = flexAlignCenterCenter

'Added by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
For intI = 15 To intField - 1
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
        grdDataList.col = 2
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
            Case "CFL", "FCL", "L", "LIN", "ACS"    '法務
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
         grdDataList.col = 2
         If Not IsNull(grdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            frm100101_2.Show
            frm100101_2.Tag = Pub_RplStr(grdDataList.Text)
            frm100101_2.StrMenu
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
         End If
     End If
     Next i
     Me.Enabled = True
Case 2
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 3
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

Private Sub Form_Load()
bolToEndByNick = False
   MoveFormToCenter Me
SetDataListWidth
'92.04.16 nick
cmdState = -1
End Sub

Sub StrMenu()
'Add By Cheng 2002/01/23
Dim strSQL11 As String
Dim strSQL21 As String
Dim strSQL31 As String
Dim strSQL41 As String
Dim strSQL51 As String
Dim ii As Integer
Dim strCon1 As String, strCon2 As String 'Added by Lydia 2023/07/25
Dim dblRow As Double 'Add By Sindy 2025/9/3

Me.Enabled = False
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
'收文日期
If Len(Trim(frm100109_1.txt1(0))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(frm100109_1.txt1(0))) & " "
End If
If Len(Trim(frm100109_1.txt1(1))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP05<=" & Val(ChangeTStringToWString(frm100109_1.txt1(1))) & " "
End If
If Len(Trim(frm100109_1.txt1(0))) <> 0 Or Len(Trim(frm100109_1.txt1(1))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100109_1.Label1(0) & frm100109_1.txt1(0) & "-" & frm100109_1.txt1(1) 'Add By Sindy 2010/11/3
End If
'承辦人
If Len(Trim(frm100109_1.txt1(2))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP14='" & frm100109_1.txt1(2) & "' "
    pub_QL05 = pub_QL05 & ";" & frm100109_1.Label1(2) & frm100109_1.txt1(2) & frm100109_1.LBL1(0) 'Add By Sindy 2010/11/3
End If
'智權人員
If Len(Trim(frm100109_1.txt1(3))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP13='" & frm100109_1.txt1(3) & "' "
    pub_QL05 = pub_QL05 & ";" & frm100109_1.Label1(3) & frm100109_1.txt1(3) & frm100109_1.LBL1(1) 'Add By Sindy 2010/11/3
End If
'案件性質
If Len(Trim(frm100109_1.txt1(5))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10>='" & frm100109_1.txt1(5) & "'  "
End If
If Len(Trim(frm100109_1.txt1(6))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10<='" & frm100109_1.txt1(6) & "'  "
End If
If Len(Trim(frm100109_1.txt1(5))) <> 0 Or Len(Trim(frm100109_1.txt1(6))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100109_1.Label1(5) & frm100109_1.txt1(5) & "-" & frm100109_1.txt1(6) 'Add By Sindy 2010/11/3
End If
strCon1 = strSQL1 'Added by Lydia 2023/07/25 專利案子查詢

'Add By Cheng 2002/01/23
'申請人國籍
If Len(Trim(frm100109_1.txt1(9))) <> 0 Then
    strSQL1 = strSQL1 + " AND CU10>='" & frm100109_1.txt1(9) & "'  "
End If
If Len(Trim(frm100109_1.txt1(10))) <> 0 Then
    strSQL1 = strSQL1 + " AND CU10<='" & frm100109_1.txt1(10) & "z'  "
End If
If Len(Trim(frm100109_1.txt1(9))) <> 0 Or Len(Trim(frm100109_1.txt1(10))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100109_1.Label1(6) & frm100109_1.txt1(9) & "-" & frm100109_1.txt1(10) 'Add By Sindy 2010/11/3
End If

strSQL2 = strSQL1
StrSQL3 = strSQL1
StrSQL4 = strSQL1
strSQL5 = strSQL1
strCon2 = strCon1 'Added by Lydia 2023/07/25 商標案子查詢
If Len(Trim(frm100109_1.txt1(4))) <> 0 Then
   'Modify By Cheng 2002/03/14
'   strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(frm100109_1.txt1(4), 1) & ") "
'   strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(frm100109_1.txt1(4), 2) & ") "
'   StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(frm100109_1.txt1(4), 3) & ") "
'   StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(frm100109_1.txt1(4), 4) & ") "
'   StrSQL5 = StrSQL5 & " AND CP01 IN (" & SQLGrpStr(frm100109_1.txt1(4), 5) & ") "
   strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100109_1.txt1(4).Text <> "ALL", frm100109_1.txt1(4).Text, GetAllSysKind(frm100109_1.txt1(4))), 1) & ") "
   strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100109_1.txt1(4).Text <> "ALL", frm100109_1.txt1(4).Text, GetAllSysKind(frm100109_1.txt1(4))), 2) & ") "
   StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100109_1.txt1(4).Text <> "ALL", frm100109_1.txt1(4).Text, GetAllSysKind(frm100109_1.txt1(4))), 3) & ") "
   StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100109_1.txt1(4).Text <> "ALL", frm100109_1.txt1(4).Text, GetAllSysKind(frm100109_1.txt1(4))), 4) & ") "
   strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100109_1.txt1(4).Text <> "ALL", frm100109_1.txt1(4).Text, GetAllSysKind(frm100109_1.txt1(4))), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Left(frm100109_1.Label1(4), 5) & frm100109_1.txt1(4) 'Add By Sindy 2010/11/3
   'Added by Lydia 2023/07/25
   strCon1 = strCon1 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100109_1.txt1(4).Text <> "ALL", frm100109_1.txt1(4).Text, GetAllSysKind(frm100109_1.txt1(4))), 1) & ") "
   strCon2 = strCon2 & " AND CP01 IN (" & SQLGrpStr(IIf(frm100109_1.txt1(4).Text <> "ALL", frm100109_1.txt1(4).Text, GetAllSysKind(frm100109_1.txt1(4))), 2) & ") "
   'end 2023/07/25
End If

'Add By Cheng 2002/01/23
strSQL11 = "": strSQL21 = "": strSQL31 = "": strSQL41 = "": strSQL51 = ""
'申請國家
If Len(Trim(frm100109_1.txt1(7))) <> 0 Then
    strSQL11 = " AND PA09>='" & frm100109_1.txt1(7) & "'  "
    strSQL21 = " AND TM10>='" & frm100109_1.txt1(7) & "'  "
    strSQL31 = " AND LC15>='" & frm100109_1.txt1(7) & "'  "
    strSQL41 = " AND '000'>='" & frm100109_1.txt1(7) & "'  "
    strSQL51 = " AND SP09>='" & frm100109_1.txt1(7) & "'  "
End If
If Len(Trim(frm100109_1.txt1(8))) <> 0 Then
    strSQL11 = strSQL11 + " AND PA09<='" & frm100109_1.txt1(8) & "'  "
    strSQL21 = strSQL21 + " AND TM10<='" & frm100109_1.txt1(8) & "'  "
    strSQL31 = strSQL31 + " AND LC15<='" & frm100109_1.txt1(8) & "'  "
    strSQL41 = strSQL41 + " AND '000'<='" & frm100109_1.txt1(8) & "'  "
    strSQL51 = strSQL51 + " AND SP09<='" & frm100109_1.txt1(8) & "'  "
End If
If Len(Trim(frm100109_1.txt1(7))) <> 0 Or Len(Trim(frm100109_1.txt1(8))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100109_1.Label1(1) & frm100109_1.txt1(7) & "-" & frm100109_1.txt1(8) 'Add By Sindy 2010/11/3
End If

'Add By Cheng 2002/03/27
'FCP管制人之控制
'If Len(Trim(frm100109_1.txt1(11).Text)) > 0 Then
'    strSQL11 = strSQL11 + " AND NA16>='" & frm100109_1.txt1(11).Text & "'  "
'    strSQL21 = strSQL21 + " AND NA16>='" & frm100109_1.txt1(11).Text & "'  "
'    strSQL31 = strSQL31 + " AND NA16>='" & frm100109_1.txt1(11).Text & "'  "
'    strSQL41 = strSQL41 + " AND NA16>='" & frm100109_1.txt1(11).Text & "'  "
'    strSQL51 = strSQL51 + " AND NA16>='" & frm100109_1.txt1(11).Text & "'  "
'End If
'If Len(Trim(frm100109_1.txt1(12).Text)) > 0 Then
'    strSQL11 = strSQL11 + " AND NA16<='" & frm100109_1.txt1(12).Text & "'  "
'    strSQL21 = strSQL21 + " AND NA16<='" & frm100109_1.txt1(12).Text & "'  "
'    strSQL31 = strSQL31 + " AND NA16<='" & frm100109_1.txt1(12).Text & "'  "
'    strSQL41 = strSQL41 + " AND NA16<='" & frm100109_1.txt1(12).Text & "'  "
'    strSQL51 = strSQL51 + " AND NA16<='" & frm100109_1.txt1(12).Text & "'  "
'End If
'FCP管制人(抓代理人國籍的FCP管制人, 若無則抓申請人國籍的FCP管制人 )
If Len(Trim(frm100109_1.txt1(11).Text)) <> 0 Then
    'Modified by Lydia 2017/02/13 +FMP管制人
    If strSrvDate(1) < FMP管制人啟用日 Then
        strSQL11 = strSQL11 & " AND DECODE(PA75,NULL,N2.NA16,N3.NA16) >='" & frm100109_1.txt1(11).Text & "' "
        strSQL51 = strSQL51 & " AND DECODE(SP26,NULL,N2.NA16,N3.NA16) >='" & frm100109_1.txt1(11).Text & "' "
    Else
        strSQL11 = strSQL11 & " AND DECODE(PA01,'P',DECODE(PA75,NULL,NVL(N2.NA79,N2.NA16),NVL(N3.NA79,N3.NA16)),DECODE(PA75,NULL,N2.NA16,N3.NA16)) >='" & frm100109_1.txt1(11).Text & "' "
        strSQL51 = strSQL51 & " AND DECODE(SP01,'PS',DECODE(SP26,NULL,NVL(N2.NA79,N2.NA16),NVL(N3.NA79,N3.NA16)),DECODE(SP26,NULL,N2.NA16,N3.NA16)) >='" & frm100109_1.txt1(11).Text & "' "
    End If
    'end 2017/02/13
End If
If Len(Trim(frm100109_1.txt1(12).Text)) <> 0 Then
    'Modified by Lydia 2017/02/13 +FMP管制人
    If strSrvDate(1) < FMP管制人啟用日 Then
        strSQL11 = strSQL11 + " AND DECODE(PA75,NULL,N2.NA16,N3.NA16) <='" & frm100109_1.txt1(12).Text & "' "
        strSQL51 = strSQL51 + " AND DECODE(SP26,NULL,N2.NA16,N3.NA16) <='" & frm100109_1.txt1(12).Text & "' "
    Else
        strSQL11 = strSQL11 & " AND DECODE(PA01,'P',DECODE(PA75,NULL,NVL(N2.NA79,N2.NA16),NVL(N3.NA79,N3.NA16)),DECODE(PA75,NULL,N2.NA16,N3.NA16)) <='" & frm100109_1.txt1(12).Text & "' "
        strSQL51 = strSQL51 & " AND DECODE(SP01,'PS',DECODE(SP26,NULL,NVL(N2.NA79,N2.NA16),NVL(N3.NA79,N3.NA16)),DECODE(SP26,NULL,N2.NA16,N3.NA16)) <='" & frm100109_1.txt1(12).Text & "' "
    End If
    'end 2017/02/13
End If
If Len(Trim(frm100109_1.txt1(11).Text)) <> 0 Or Len(Trim(frm100109_1.txt1(12).Text)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & frm100109_1.Label1(7) & frm100109_1.txt1(11) & "-" & frm100109_1.txt1(12) 'Add By Sindy 2010/11/3
End If
'Modify By Cheng 2002/01/23
'加搜尋欄位--申請國家
'                strSQL = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),pa26) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(PA09,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限 FROM CASEPROGRESS,NEXTPROGRESS,PATENT,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2 WHERE CP09>'C' AND CP09=NP01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("PA26", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL1
'strSQL = strSQL & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(TM10,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限 FROM CASEPROGRESS,NEXTPROGRESS,TRADEMARK,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2 WHERE CP09>'C' AND CP09=NP01(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("TM23", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & strSQL2
'strSQL = strSQL & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(DECODE(LC15,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),LC11) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(LC15,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限 FROM CASEPROGRESS,NEXTPROGRESS,LAWCASE,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2 WHERE CP09>'C' AND CP09=NP01(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("LC11", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL3
'strSQL = strSQL & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,NVL(C1.CPM03,CP10)                             AS 來函性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),HC05) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(C2.CPM03,TO_CHAR(NP07))                             AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限 FROM CASEPROGRESS,NEXTPROGRESS,HIRECASE,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2 WHERE CP09>'C' AND CP09=NP01(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("HC05", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL4
'strSQL = strSQL & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(DECODE(SP09,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(SP09,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限 FROM CASEPROGRESS,NEXTPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2 WHERE CP09>'C' AND CP09=NP01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("SP08", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) " & StrSQL5
'strSQL = strSQL + " ORDER BY 收文日,本所案號 "
'Modify By Cheng 2002/04/25
'若已閉卷, 則在本所案號後加"*"號
'Modify By Cheng 2003/08/15
'                strSQL = "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),pa26) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(PA09,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,nvl(N2.NA03,N2.NA04) As 申請國家 FROM CASEPROGRESS,NEXTPROGRESS,PATENT,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,Nation N2,NATION N3,FAGENT " & _
'                                " WHERE CP09>'C' AND CP09=NP01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("PA26", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And PA09=N2.NA01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=N3.NA01(+) " & strSQL11 & strSQL1
'strSQL = strSQL & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),TM23) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(TM10,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,nvl(N2.NA03,N2.NA04) As 申請國家 FROM CASEPROGRESS,NEXTPROGRESS,TRADEMARK,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,Nation N2 " & _
'                " WHERE CP09>'C' AND CP09=NP01(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("TM23", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And TM10=N2.NA01(+) " & strSQL21 & strSQL2
'strSQL = strSQL & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,NVL(LC05,NVL(LC06,LC07)) AS 案件名稱,NVL(DECODE(LC15,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),LC11) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(LC15,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,nvl(N2.NA03,N2.NA04) As 申請國家 FROM CASEPROGRESS,NEXTPROGRESS,LAWCASE,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,Nation N2 " & _
'                " WHERE CP09>'C' AND CP09=NP01(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("LC11", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And LC15=N2.NA01(+) " & strSQL31 & StrSQL3
'strSQL = strSQL & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,HC06                     AS 案件名稱,NVL(C1.CPM03,CP10)                             AS 來函性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),HC05) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(C2.CPM03,TO_CHAR(NP07))                             AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,nvl(N2.NA03,N2.NA04) As 申請國家 FROM CASEPROGRESS,NEXTPROGRESS,HIRECASE,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,Nation N2 " & _
'                " WHERE CP09>'C' AND CP09=NP01(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("HC05", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And N2.NA01='000' " & strSQL41 & StrSQL4
'strSQL = strSQL & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(DECODE(SP09,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,NVL(CU05||' '||CU88||' '||CU89||' '||CU90,CU06)),SP08) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(SP09,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,nvl(N2.NA03,N2.NA04) As 申請國家 FROM CASEPROGRESS,NEXTPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,Nation N2,NATION N3,FAGENT " & _
'                " WHERE CP09>'C' AND CP09=NP01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("SP08", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And SP09=N2.NA01(+)  AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND FA10=N3.NA01(+)  " & strSQL51 & StrSQL5
'strSQL = strSQL + " ORDER BY 收文日,本所案號 "
'2010/9/14 MODIFY BY SONIA 日期欄改百年日期排序問題

'Added by Lydia 2015/10/14 +排除特定來函
If frm100109_1.Check1.Value = 1 Then
    'Modified by Lydia 2015/10/20 專利案+1217,1229
    'Modified by Lydia 2017/02/21 專利案+1917 通知告准
    'Modified by Morgan 2020/2/6 專利案取消 1204 通知實審日,1207 通知即將公開--何淑華
    'Modified by Morgan 2020/2/6 通知期限1913(專利),1725(商標)為 D 類，已控制剔除不必再列
    'Modified by Morgan 2020/2/6 專利案取消 1217 通知形式審查
    strSQL1 = strSQL1 & " AND CP10 NOT IN (1101,1605,1604,1229,1917) "
    strSQL2 = strSQL2 & " AND CP10 NOT IN (1101,1704,1722,1723) "
    pub_QL05 = pub_QL05 & ";" & frm100109_1.Check1.Caption & frm100109_1.lblCheck1.Caption & frm100109_1.lblCheck2.Caption
End If
'end 2015/10/14

'Added by Lydia 2019/11/01 利益衝突案件
m_AllSys = IIf(frm100109_1.txt1(4).Text <> "ALL", frm100109_1.txt1(4).Text, GetAllSysKind(, frm100109_1.txt1(4).Text))
intCufaCnt = 0
'end 2019/11/01

'strSql = "select V,收文日,本所案號,分所號,案件名稱,來函性質,申請人,承辦人,智權人員,下一程序,本所期限,法定期限,申請國家,SUBSTR(' '||sqldatet(C2.CP27),-9) AS 相關收文號發文日,a.CP09 from ("
strSql = "select V,收文日,本所案號,分所號,案件名稱,來函性質,申請人,承辦人,智權人員,下一程序,本所期限,法定期限" & _
            ",申請國家,SUBSTR(' '||sqldatet(C2.CP27),-9) AS 相關收文號發文日,a.CP09 " & _
            ",cust01,cust02,cust03,cust04,cust05,fcno from ("
'end 2019/11/01

'Modified by Lydia 2018/05/24 排除D類收文 and CP09 < 'D'
'Modified by Lydia 2019/11/01 增加欄位:申請人1~5(cust01~cust05),FC代理人
'Modified by Lydia 2023/07/25 (原2023/06/28)專利案件性質1004延期受理，或商標的案件性質1005延期受理，Grid中之本所期限及法定期限，請帶該進度之本所期限及法定期限。
                              '=>改成以延期受理之CP43抓出延期A，再以延期A的CP43判斷，若為C類則抓下一程序的案件性質，若為A或B類則抓進度檔的案件性質。
'strSql = strSql & "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),pa26) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(PA09,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,nvl(N2.NA03,N2.NA04) As 申請國家,CP43,CP09" & _
                    ",pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno" & _
                    " FROM CASEPROGRESS,NEXTPROGRESS,PATENT,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,Nation N2,NATION N3,FAGENT " & _
                    " WHERE CP09>'C' and CP09 < 'D' AND CP09=NP01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("PA26", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And PA09=N2.NA01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=N3.NA01(+) " & strSQL11 & strSQL1
'strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),TM23) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(TM10,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,nvl(N2.NA03,N2.NA04) As 申請國家,CP43,CP09" & _
                ",tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno" & _
                " FROM CASEPROGRESS,NEXTPROGRESS,TRADEMARK,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,Nation N2 " & _
                " WHERE CP09>'C' and CP09 < 'D' AND CP09=NP01(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("TM23", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And TM10=N2.NA01(+) " & strSQL21 & strSQL2
'專利
strSql = strSql & "SELECT '' AS V,SUBSTR(' '||SQLDATET(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(LENGTH(NVL(PA108,'')),NULL,'','●') AS 本所案號,DECODE(LENGTH(NVL(PA136,'')),NULL,'','●')||PA47 AS 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(PA09,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||SQLDATET(NP08),-9) AS 本所期限,SUBSTR(' '||SQLDATET(NP09),-9) AS 法定期限,NVL(N2.NA03,N2.NA04) AS 申請國家,CP43,CP09" & _
                    ",PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO" & _
                    " FROM CASEPROGRESS,NEXTPROGRESS,PATENT,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,NATION N2,NATION N3,FAGENT " & _
                    " WHERE CP09>'C' AND CP09 < 'D' AND CP10 <> '1004' AND CP09=NP01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("PA26", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND PA09=N2.NA01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=N3.NA01(+) " & strSQL11 & strSQL1
  '---以延期受理之CP43抓出延期A，再以延期A的CP43判斷，若為C類則抓下一程序的案件性質，若為A或B類則抓進度檔的案件性質。
  strExc(1) = "select c1.cp01 as nextcp01,c1.cp02 as nextcp02,c1.cp03 as nextcp03,c1.cp04 as nextcp04,c1.cp09 as nextcp09,c2.cp09 as nextcp09_2,c2.cp43 as nextcp43 " & _
              ",decode(substr(c2.cp43,1,1),'C',np07,c3.cp10) as nextpty ,decode(substr(c2.cp43,1,1),'C',np08,c3.cp06) as nextdate1,decode(substr(c2.cp43,1,1),'C',np09,c3.cp07) as nextdate2 " & _
              "from caseprogress c1, caseprogress c2 ,caseprogress c3, nextprogress where c1.cp10='1004' " & Replace(Replace(UCase(strCon1), "CP0", "C1.CP0"), "CP1", "C1.CP1") & " and c1.cp43=c2.cp09(+) " & _
              "and c2.cp43=c3.cp09(+) and c2.cp43=np01(+) and c2.cp01=np02(+) " & _
              "group by c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp09,c2.cp09,c2.cp43,decode(substr(c2.cp43,1,1),'C',np07,c3.cp10),decode(substr(c2.cp43,1,1),'C',np08,c3.cp06),decode(substr(c2.cp43,1,1),'C',np09,c3.cp07) "
strSql = strSql & "UNION ALL SELECT '' AS V,SUBSTR(' '||SQLDATET(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(LENGTH(NVL(PA108,'')),NULL,'','●') AS 本所案號,DECODE(LENGTH(NVL(PA136,'')),NULL,'','●')||PA47 AS 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA26) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(PA09,'000',C2.CPM03,C2.CPM04),TO_CHAR(NEXTPTY)) AS 下一程序,SUBSTR(' '||SQLDATET(NEXTDATE1),-9) AS 本所期限,SUBSTR(' '||SQLDATET(NEXTDATE2),-9) AS 法定期限,NVL(N2.NA03,N2.NA04) AS 申請國家,CP43,CP09" & _
                    ",PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO" & _
                    " FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,NATION N2,NATION N3,FAGENT,( " & strExc(1) & ") VTB1 " & _
                    " WHERE CP09>'C' AND CP09 < 'D' AND CP10 = '1004' AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND CP09=NEXTCP09 AND C2.CPM01=NEXTCP01(+) AND C2.CPM02=NEXTPTY(+) AND " & SQLNewFag("PA26", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND PA09=N2.NA01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=N3.NA01(+) " & strSQL11 & strSQL1
'商標
strSql = strSql & " UNION ALL SELECT '' AS V,SUBSTR(' '||SQLDATET(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(LENGTH(NVL(TM57,'')),NULL,'','●') AS 本所案號,DECODE(LENGTH(NVL(TM73,'')),NULL,'','●')||TM34 AS 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(TM10,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||SQLDATET(NP08),-9) AS 本所期限,SUBSTR(' '||SQLDATET(NP09),-9) AS 法定期限,NVL(N2.NA03,N2.NA04) AS 申請國家,CP43,CP09" & _
                ",TM23 AS CUST01,TM78 AS CUST02,TM79 AS CUST03,TM80 AS CUST04,TM81 AS CUST05,TM44 AS FCNO" & _
                " FROM CASEPROGRESS,NEXTPROGRESS,TRADEMARK,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,NATION N2 " & _
                " WHERE CP09>'C' AND CP09 < 'D' AND CP10 <> '1005' AND CP09=NP01(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("TM23", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND TM10=N2.NA01(+) " & strSQL21 & strSQL2
  '---以延期受理之CP43抓出延期A，再以延期A的CP43判斷，若為C類則抓下一程序的案件性質，若為A或B類則抓進度檔的案件性質。
  strExc(2) = "select c1.cp01 as nextcp01,c1.cp02 as nextcp02,c1.cp03 as nextcp03,c1.cp04 as nextcp04,c1.cp09 as nextcp09,c2.cp09 as nextcp09_2,c2.cp43 as nextcp43 " & _
              ",decode(substr(c2.cp43,1,1),'C',np07,c3.cp10) as nextpty ,decode(substr(c2.cp43,1,1),'C',np08,c3.cp06) as nextdate1,decode(substr(c2.cp43,1,1),'C',np09,c3.cp07) as nextdate2 " & _
              "from caseprogress c1, caseprogress c2 ,caseprogress c3, nextprogress where c1.cp10='1005' " & Replace(Replace(UCase(strCon2), "CP0", "C1.CP0"), "CP1", "C1.CP1") & " and c1.cp43=c2.cp09(+) " & _
              "and c2.cp43=c3.cp09(+) and c2.cp43=np01(+) and c2.cp01=np02(+) " & _
              "group by c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp09,c2.cp09,c2.cp43,decode(substr(c2.cp43,1,1),'C',np07,c3.cp10),decode(substr(c2.cp43,1,1),'C',np08,c3.cp06),decode(substr(c2.cp43,1,1),'C',np09,c3.cp07) "
strSql = strSql & " UNION ALL SELECT '' AS V,SUBSTR(' '||SQLDATET(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊','')||DECODE(LENGTH(NVL(TM57,'')),NULL,'','●') AS 本所案號,DECODE(LENGTH(NVL(TM73,'')),NULL,'','●')||TM34 AS 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM23) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(TM10,'000',C2.CPM03,C2.CPM04),TO_CHAR(NEXTPTY)) AS 下一程序,SUBSTR(' '||SQLDATET(NEXTDATE1),-9) AS 本所期限,SUBSTR(' '||SQLDATET(NEXTDATE2),-9) AS 法定期限,NVL(N2.NA03,N2.NA04) AS 申請國家,CP43,CP09" & _
                ",TM23 AS CUST01,TM78 AS CUST02,TM79 AS CUST03,TM80 AS CUST04,TM81 AS CUST05,TM44 AS FCNO" & _
                " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,NATION N2,( " & strExc(2) & ") VTB1 " & _
                " WHERE CP09>'C' AND CP09 < 'D' AND CP10 = '1005' AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND CP09=NEXTCP09 AND C2.CPM01=NEXTCP01(+) AND C2.CPM02=NEXTPTY(+) AND " & SQLNewFag("TM23", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND TM10=N2.NA01(+) " & strSQL21 & strSQL2
                
'Modified by Lydia 2023/06/28 專利案件性質1004延期受理，或商標的案件性質1005延期受理，Grid中之本所期限及法定期限，請帶該進度之本所期限及法定期限。

'''strSql = strSql & "SELECT '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(pa108,'')),null,'','●') AS 本所案號,DECODE(length(nvl(pa136,'')),null,'','●')||pa47 as 分所號,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,NVL(DECODE(PA09,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),pa26) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(PA09,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序," & _
'''                    "DECODE(INSTR('P1004,CFP1004,FCP1004',CP01||CP10),0,SUBSTR(' '||sqldatet(NP08),-9),SUBSTR(' '||sqldatet(CP06),-9)) AS 本所期限,DECODE(INSTR('P1004,CFP1004,FCP1004',CP01||CP10), 0 , SUBSTR(' '||sqldatet(NP09),-9), SUBSTR(' '||sqldatet(CP07),-9)) AS 法定期限,nvl(N2.NA03,N2.NA04) As 申請國家,CP43,CP09" & _
'''                    ",pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno" & _
'''                    " FROM CASEPROGRESS,NEXTPROGRESS,PATENT,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,Nation N2,NATION N3,FAGENT " & _
'''                    " WHERE CP09>'C' and CP09 < 'D' AND CP09=NP01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("PA26", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And PA09=N2.NA01(+) AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=N3.NA01(+) " & strSQL11 & strSQL1
'''strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(tm29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●') AS 本所案號,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 as 分所號,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,NVL(DECODE(TM10,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),TM23) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(TM10,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序," & _
'''                "DECODE(INSTR('T1005,CFT1005,FCT1005',CP01||CP10),0,SUBSTR(' '||sqldatet(NP08),-9),SUBSTR(' '||sqldatet(CP06),-9)) AS 本所期限,DECODE(INSTR('T1005,CFT1005,FCT1005',CP01||CP10), 0 , SUBSTR(' '||sqldatet(NP09),-9), SUBSTR(' '||sqldatet(CP07),-9)) AS 法定期限,nvl(N2.NA03,N2.NA04) As 申請國家,CP43,CP09" & _
'''                ",tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno" & _
'''                " FROM CASEPROGRESS,NEXTPROGRESS,TRADEMARK,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,Nation N2 " & _
'''                " WHERE CP09>'C' and CP09 < 'D' AND CP09=NP01(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("TM23", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And TM10=N2.NA01(+) " & strSQL21 & strSQL2
''''end 2023/06/28
strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(lc08,'Y','＊','')||DECODE(length(nvl(lc34,'')),null,'','●') AS 本所案號,DECODE(length(nvl(lc36,'')),null,'','●')||lc16 as 分所號,nvl(lc05,nvl(lc06,lc07)) AS 案件名稱,NVL(DECODE(LC15,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),LC11) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(LC15,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,nvl(N2.NA03,N2.NA04) As 申請國家,CP43,CP09" & _
                ",lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno" & _
                " FROM CASEPROGRESS,NEXTPROGRESS,LAWCASE,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,Nation N2 " & _
                " WHERE CP09>'C' and CP09 < 'D' AND CP09=NP01(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("LC11", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And LC15=N2.NA01(+) " & strSQL31 & StrSQL3
strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(hc09,'Y','＊','')||DECODE(length(nvl(hc19,'')),null,'','●') AS 本所案號,DECODE(length(nvl(hc20,'')),null,'','●')||hc07 as 分所號,hc06                     AS 案件名稱,NVL(C1.CPM03,CP10)                             AS 來函性質,NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),HC05) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(C2.CPM03,TO_CHAR(NP07))                             AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,nvl(N2.NA03,N2.NA04) As 申請國家,CP43,CP09" & _
                ",hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno" & _
                " FROM CASEPROGRESS,NEXTPROGRESS,HIRECASE,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,Nation N2 " & _
                " WHERE CP09>'C' and CP09 < 'D' AND CP09=NP01(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("HC05", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And N2.NA01='000' " & strSQL41 & StrSQL4
strSql = strSql & " union all select '' AS V,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(sp15,'Y','＊','')||DECODE(length(nvl(sp61,'')),null,'','●') AS 本所案號,DECODE(length(nvl(sp68,'')),null,'','●')||sp28 as 分所號,NVL(SP05,NVL(SP06,SP07)) AS 案件名稱,NVL(DECODE(SP09,'000',C1.CPM03,C1.CPM04),CP10) AS 來函性質,NVL(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),SP08) AS 申請人,NVL(S1.ST02,CP14) AS 承辦人,NVL(S2.ST02,CP13) AS 智權人員,NVL(DECODE(SP09,'000',C2.CPM03,C2.CPM04),TO_CHAR(NP07)) AS 下一程序,SUBSTR(' '||sqldatet(NP08),-9) AS 本所期限,SUBSTR(' '||sqldatet(NP09),-9) AS 法定期限,nvl(N2.NA03,N2.NA04) As 申請國家,CP43,CP09" & _
                ",sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno" & _
                " FROM CASEPROGRESS,NEXTPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP C1,CASEPROPERTYMAP C2,CUSTOMER,STAFF S1,STAFF S2,Nation N2,NATION N3,FAGENT " & _
                " WHERE CP09>'C' and CP09 < 'D' AND CP09=NP01(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=C1.CPM01(+) AND CP10=C1.CPM02(+) AND NP02=C2.CPM01(+) AND TO_CHAR(NP07)=C2.CPM02(+) AND " & SQLNewFag("SP08", "CU") & " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) And SP09=N2.NA01(+)  AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+) AND FA10=N3.NA01(+)  " & strSQL51 & strSQL5
'end 2019/11/01
strSql = strSql & ") a,CASEPROGRESS C2 where a.CP43=C2.CP09(+)"
strSql = strSql + " ORDER BY 收文日,本所案號 "

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
        InsertQueryLog (dblRow) 'Add By Sindy 2010/11/3
        If adoRecordset.RecordCount = 0 Then
              GoTo JumpToNoData
        End If
     Else
        InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/3
     End If
    'end 2019/11/01
    
    cmdOK(0).Enabled = True
    cmdOK(1).Enabled = True
Else
    InsertQueryLog (0) 'Add By Sindy 2010/11/3
JumpToNoData:   'Added by Lydia 2019/11/01
    cmdOK(0).Enabled = False
    cmdOK(1).Enabled = False
    Me.Enabled = True
    ShowNoData
    Screen.MousePointer = vbDefault
    '92.04.18 nick
    'Me.Hide
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Exit Sub
End If
Me.grdDataList.Visible = False
Set grdDataList.Recordset = adoRecordset
SetDataListWidth
For ii = 1 To Me.grdDataList.Rows - 1
    'Modify by Morgan 2011/7/20
    'Me.grdDataList.TextMatrix(ii, 5) = Me.grdDataList.TextMatrix(ii, 5) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 13), "1")
    Me.grdDataList.TextMatrix(ii, 5) = Me.grdDataList.TextMatrix(ii, 5) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(ii, 14), "1")
Next ii
Me.grdDataList.Visible = True
CheckOC
Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100109_2 = Nothing
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
