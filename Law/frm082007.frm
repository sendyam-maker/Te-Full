VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm082007 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外部法務處期限通知"
   ClientHeight    =   5360
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5360
   ScaleWidth      =   9390
   Begin VB.CommandButton cmdHelp 
      Caption         =   "事件說明(&H)"
      Height          =   400
      Left            =   2745
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   60
      Width           =   1110
   End
   Begin VB.TextBox txtUsernum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1125
      MaxLength       =   6
      TabIndex        =   0
      Top             =   510
      Width           =   915
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "隱藏白色(&H)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   45
      TabIndex        =   9
      Top             =   60
      Width           =   1380
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "所有未發文(&A)"
      Height          =   400
      Index           =   1
      Left            =   1440
      TabIndex        =   8
      Top             =   60
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   400
      Index           =   3
      Left            =   7740
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   2
      Left            =   6210
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   60
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm082007.frx":0000
      Left            =   6150
      List            =   "frm082007.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   510
      Width           =   3195
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   5400
      TabIndex        =   3
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8550
      TabIndex        =   1
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4275
      Left            =   45
      TabIndex        =   2
      Top             =   840
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   7549
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V|本所期限|法定期限|承辦期限|核稿期限|管制人|承辦人|事件　|本所案號　　　|案件性質|備註　　　　|案件名稱　　　"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
   End
   Begin VB.Label lblCnt 
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7620
      TabIndex        =   13
      Top             =   5160
      Width           =   1710
   End
   Begin VB.Label lblUserName 
      Height          =   180
      Left            =   2070
      TabIndex        =   11
      Top             =   570
      Width           =   1710
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   135
      TabIndex        =   10
      Top             =   555
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "顏色符號說明："
      Height          =   180
      Left            =   4860
      TabIndex        =   5
      Top             =   570
      Width           =   1260
   End
End
Attribute VB_Name = "frm082007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/11 Form2.0已檢查 (無需修改的物件); 已改用frm072005
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim bolBarShow As Boolean
Public cmdState As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_stSort As String '排序方式
Dim m_adoRst As ADODB.Recordset
Dim bLvl4 As Boolean, bLvl5 As Boolean
Dim stNumList1(1 To 5) As String


Private Sub cmdHelp_Click()
   strExc(0) = "" & vbCrLf
   strExc(0) = strExc(0) & "１達本所：本所期限＜＝系統日＋２個工作天之未發文案件。" & vbCrLf
   strExc(0) = strExc(0) & "　　　　或本所期限＜＝系統日＋３個工作天之通知開庭案件。" & vbCrLf
   strExc(0) = strExc(0) & vbCrLf
   strExc(0) = strExc(0) & "２未發文：已收文未發文且無期限或期限未達管制日期之案件。（按所有未發文按鈕才會顯示）" & vbCrLf
   strExc(0) = strExc(0) & vbCrLf
   strExc(0) = strExc(0) & "３達承辦：承辦期限＜＝系統日＋２個工作天之未發文案件。" & vbCrLf
   strExc(0) = strExc(0) & vbCrLf
   strExc(0) = strExc(0) & "４未收文：本所期限＜＝系統日＋５個工作天之未收文案件。" & vbCrLf
   strExc(0) = strExc(0) & vbCrLf
   strExc(0) = strExc(0) & "５未回執：發文日＋６天(日曆天)＜＝系統日之未回執案件。" & vbCrLf
   
   MsgBox strExc(0), vbOKOnly, "事件說明"
End Sub

Private Sub cmdHide_Click()
'   SetRst2Grid
   SetColor cmdHide.Tag
End Sub

'Private Sub SetRst2Grid()
'   grdDataList.FixedCols = 0
'   Set grdDataList.Recordset = m_adoRst
'   grdDataList.FixedCols = 3
'End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData(Optional bolRefresh As Boolean = False)
   Dim i As Integer, StrTag As String, lngColor As Long, ii As Integer
On Error GoTo ErrorHandler
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
         bolRefresh = False
         grdDataList.col = 0
         grdDataList.Text = ""
         grdDataList.CellBackColor = grdDataList.BackColor
         grdDataList.col = 3
         lngColor = grdDataList.CellBackColor
         For ii = 1 To 2
            grdDataList.col = ii
            grdDataList.CellBackColor = lngColor
         Next
         
         Dim Str01 As String
         
         StrTag = grdDataList.TextMatrix(i, 9)
         If Left(Right(StrTag, 7), 1) = "-" Then
            StrTag = StrTag & "-0-00"
         ElseIf Left(Right(StrTag, 2), 1) = "-" Then
            StrTag = StrTag & "-00"
         End If
         
         If Left(StrTag, 1) < "A" Or Left(StrTag, 1) > "Z" Then
            StrTag = Right(StrTag, Len(StrTag) - 1)
         End If
         Str01 = SystemNumber(StrTag, 1)
         If fnSaveParentForm(Me) = False Then
            Exit For
         End If
         Me.Show
         Select Case cmdState
            Case 2 '案件基本資料
               Select Case Pub_RplStr(Str01)
                  Case "CFL", "FCL", "LIN", "L" '法務
                      Screen.MousePointer = vbHourglass
                      frm100101_5.Show
                      frm100101_5.Tag = StrTag
                      frm100101_5.StrMenu
                      Screen.MousePointer = vbDefault
               End Select
               
            Case 3 '案件進度
               frm100101_2.Show
               frm100101_2.Tag = StrTag
               frm100101_2.StrMenu
         End Select
         Exit For
      End If
   Next i
   
   If bolRefresh = True Then
      cmdQuery_Click 0
   End If
   
ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "(" & Err.Number & ")" & Err.Description
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Public Sub cmdQuery_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   If pub_bolInformCheck = False Then
      If MsgBox("是否確定要查詢？", vbYesNo + vbDefaultButton2) = vbNo Then
         GoTo SubOut
      End If
   End If
   
   Me.Enabled = False
   doQuery Index
   Me.Enabled = True
   
SubOut:
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   bolBarShow = Forms(0).StatusBar1.Visible
   Forms(0).StatusBar1.Visible = True
End Sub

Private Sub Form_Deactivate()
   Forms(0).StatusBar1.Visible = bolBarShow
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function GetNumList(p_UID As String, Optional iLevel As Integer) As String
   Dim stRtn As String, stSQL As String
   
   stSQL = "select ''''||st01||'''' from staff"
   Select Case iLevel
      Case 2
         stSQL = stSQL & " where ST52='" & p_UID & "'"
      Case 3
         stSQL = stSQL & " where ST53='" & p_UID & "'"
      Case 4
         stSQL = stSQL & " where ST54='" & p_UID & "'"
      Case 5
         stSQL = stSQL & " where ST55='" & p_UID & "'"
      Case Else
         stSQL = stSQL & " where ST52='" & p_UID & "'"
   End Select
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      stRtn = RsTemp.GetString(adClipString, , , ",")
      stRtn = Left(stRtn, Len(stRtn) - 1)
   End If
   GetNumList = stRtn
End Function

'語法內有用組合欄位為條件以控制使用特定index(避掉某些不適當的)
Private Sub doQuery(idx As Integer)
   Dim stVTB As String, stDate1 As String, stDate2 As String, stDate3 As String, stDate4 As String, stDate5 As String
   Dim stConCP06 As String, stConCP48 As String, stConCP0603 As String
   Dim stNumList As String, stDept As String
   Dim ii As Integer, stIdList
   Dim stUserID As String
   Dim stCP01 As String
   Dim strOtherUser As String
   Dim rsTmp As New ADODB.Recordset
   
   'Modify By Sindy 2011/10/20 +L
   stCP01 = " and cp01 in ('FCL','CFL','LIN','L')"
   
   If lblUserName = "" Then
      MsgBox "員工編號錯誤！"
      If txtUsernum.Enabled = True Then
         txtUsernum_GotFocus
         txtUsernum.SetFocus
      End If
      Exit Sub
   End If
   
   stUserID = txtUsernum
   '使用者收文智權人員所屬部門
   If stUserID = strUserNum Then
      stDept = Pub_StrUserSt15
   Else
      stDept = GetST15(stUserID)
   End If
   
   '抓員工外譯對照資料
   stNumList = PUB_GetMapID(stUserID, 0)
   If stNumList <> "" Then
      stNumList = "'" & stNumList & "','" & stUserID & "'"
   Else
      stNumList = "'" & stUserID & "'"
   End If
   stNumList1(1) = stNumList
   
   '期限管制人
   For ii = 2 To 5
      stNumList1(ii) = GetNumList(stUserID, ii)
      If stNumList1(ii) <> "" Then
         stNumList = stNumList & "," & stNumList1(ii)
      End If
   Next
   
   'CompWorkDay計算工作天時,為何要原工作天加2,是因為當天不算+1,再來是因為期限比較時使用<計算所以再+1
   stDate1 = strSrvDate(1) - 10000
   stDate2 = CompWorkDay(4, strSrvDate(1)) '+2個工作天
   stDate3 = CompWorkDay(7, strSrvDate(1)) '+5個工作天
   stDate4 = CompWorkDay(5, strSrvDate(1)) 'Add By Sindy 2011/6/27 系統日+3個工作天
   stDate5 = DBDATE(DateAdd("d", -6, ChangeWStringToWDateString(strSrvDate(1)))) 'Add By Sindy 2011/6/27 系統日-6天
   
   stConCP06 = " AND CP06>=" & stDate1 & " AND CP06< " & stDate2
   stConCP48 = " AND CP48>=" & stDate1 & " AND CP48< " & stDate2
   stConCP0603 = " AND CP06>=" & stDate1 & " AND CP06< " & stDate4 'Add By Sindy 2011/6/27
   
   '特殊權限
   bLvl4 = CheckLevel(stUserID, "U") '第四級管制人
   bLvl5 = CheckLevel(stUserID, "O") '第五級管制人
   
   '代碼1:A=達本所,B=達承辦,C=達核稿,D=未收文,E=未發文,F=未請款,G=未交稿,H=未回執
   '代碼2:0=未分案,1=承辦人,2=管制人,3=智權人員,4=核稿人,8=無核稿人 9=未交稿
   
   '已收文未發文,2個工作天後達本所期限者(不含當日) --承辦人-A1(本所期限,承辦人)
   'Modified by Lydia 2016/10/20 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
   strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,CP06,CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
            "From CASEPROGRESS, LawCase" & _
            " WHERE CP05>20030000" & _
            " AND (CP14 IN(" & stNumList & ") or CP29 IN(" & stNumList & ")) AND CP158=0 AND CP159=0" & stConCP06 & stCP01 & _
            " AND CP10<>'9001'" & _
            " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
            " AND LC08 IS NULL AND LC34 is null" & _
            " AND LC01 IS NOT NULL"
   cnnConnection.Execute strSql, intI
   'Add By Sindy 2011/6/27 案件性質為9001者改為本所期限<=系統日+3個工作天
   'Modified by Lydia 2016/10/20 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
   strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,CP06,CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
            "From CASEPROGRESS, LawCase" & _
            " WHERE CP05>20030000" & _
            " AND (CP14 IN(" & stNumList & ") or CP29 IN(" & stNumList & ")) AND CP158=0 AND CP159=0" & stConCP0603 & stCP01 & _
            " AND CP10='9001'" & _
            " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
            " AND LC08 IS NULL AND LC34 is null" & _
            " AND LC01 IS NOT NULL"
   cnnConnection.Execute strSql, intI
   '已收文未發文,2個工作天後達承辦期限者(不含當日) --承辦人-B1(承辦期限,承辦人)
   'Modified by Lydia 2016/10/20 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
   strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,CP06,CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
            "From CASEPROGRESS, LawCase" & _
            " WHERE CP05>20030000" & _
            " AND (CP14 IN(" & stNumList & ") or CP29 IN(" & stNumList & ")) AND CP158=0 AND CP159=0" & stConCP48 & stCP01 & _
            " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
            " AND LC08 IS NULL AND LC34 is null" & _
            " AND LC01 IS NOT NULL"
   cnnConnection.Execute strSql, intI
   '已收文已發文,發文日期+6天未回執者 --承辦人-H1(發文日期,承辦人)
   strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','H' EV1,'1' EV2,CP09,to_number(to_char(to_date(cp27,'YYYYMMDD')+6,'YYYYMMDD')),CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
            "From CASEPROGRESS, LawCase" & _
            " WHERE cp27>=20110701 and cp50 is not null" & _
            " AND (CP14 IN(" & stNumList & ") or CP29 IN(" & stNumList & ")) AND CP57||cp46 is null AND CP27 is not null" & _
            " AND CP27<=" & stDate5 & stCP01 & _
            " AND LC01=CP01 AND LC02=CP02 AND LC03=CP03 AND LC04=CP04" & _
            " AND LC08 IS NULL AND LC34 is null" & _
            " AND LC01 IS NOT NULL"
   cnnConnection.Execute strSql, intI
   '未收文且 5個工作天 後達本所期限者(不含當日) --智權人員-D3(未收文,智權人員)
   strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
            " SELECT '" & strUserNum & "','D' EV1,'3' EV2,NP01,NP08,NP09,0+null CP48,0+NULL EP08,'' NA16,NULL CP14,NP10,NP22 " & _
            "From NEXTPROGRESS, LawCase" & _
            " WHERE NP02||NP06 in ('FCL','LIN')" & _
            " AND NP10 IN(" & stNumList & ")" & _
            " AND NP08>=" & stDate1 & " AND NP08< " & stDate3 & _
            " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05" & _
            " AND LC08 IS NULL AND LC34 is null"
   cnnConnection.Execute strSql, intI
   '所有未發文--承辦人-E(未發文)
   If idx = 1 Then
      'Modified by Lydia 2016/10/20 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
      'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (cp06||cp48 is null or (cp06>=" & stDate2 & " and cp48 is null) or (cp06 is null and cp48>=" & stDate2 & ") or (cp06>=" & stDate2 & " and cp48>=" & stDate2 & "))"
      strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','E' EV1,'1' EV2,CP09,CP06,CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
               " From CASEPROGRESS, LawCase" & _
               " WHERE CP05>20030000" & _
               " AND (CP14 IN(" & stNumList & ") or CP29 IN(" & stNumList & ")) AND CP158=0 AND CP159=0" & stCP01 & _
               " AND LC01=CP01 AND LC02=CP02 AND LC03=CP03 AND LC04=CP04" & _
               " AND LC08 IS NULL AND LC34 is null" & _
               " AND LC01 IS NOT NULL"
      cnnConnection.Execute strSql, intI
   End If
   
   '未分案-0
   If bLvl4 = True Or bLvl5 = True Then
      '承辦人為非法務處人員
      'Modified by Lydia 2016/10/20 cp27 is null and cp57 is null => CP158=0 AND CP159=0
      strSql = "select distinct cp14 from caseprogress " & _
               "where cp01 in ('CFL','FCL','LIN','L') " & _
               "and CP158=0 AND CP159=0 " & _
               "and cp06 is not null " & _
               "and exists (select * from staff where st01=cp14 and st11 not in ('F1','F2','D4')) "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      strOtherUser = ""
      If intI = 1 Then
         With RsTemp
            .MoveFirst
            Do While Not .EOF
               If Trim(.Fields(0)) <> "" Then
                  If strOtherUser <> "" Then strOtherUser = strOtherUser & ","
                  strOtherUser = strOtherUser & "'" & Trim(.Fields(0)) & "'"
               End If
               .MoveNext
            Loop
         End With
      End If
      '非法務處人員承辦案件(st11 非 F1,F2,D4 字頭的)
      If strOtherUser <> "" Then
         '已收文未發文,2個工作天後達本所期限者(不含當日) --承辦人-A1(本所期限,承辦人)
         'Modified by Lydia 2016/10/20 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
         strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,CP06,CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
                  "From CASEPROGRESS, LawCase" & _
                  " WHERE CP05>20030000" & _
                  " AND (CP14 IN(" & strOtherUser & ") or CP29 IN(" & strOtherUser & ")) AND CP158=0 AND CP159=0" & stConCP06 & stCP01 & _
                  " AND CP10<>'9001'" & _
                  " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                  " AND LC08 IS NULL AND LC34 is null" & _
                  " AND LC01 IS NOT NULL"
         cnnConnection.Execute strSql, intI
         'Add By Sindy 2011/6/27 案件性質為9001者改為本所期限<=系統日+3個工作天
         'Modified by Lydia 2016/10/20 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
         strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,CP06,CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
                  "From CASEPROGRESS, LawCase" & _
                  " WHERE CP05>20030000" & _
                  " AND (CP14 IN(" & strOtherUser & ") or CP29 IN(" & strOtherUser & ")) AND CP158=0 AND CP159=0" & stConCP0603 & stCP01 & _
                  " AND CP10='9001'" & _
                  " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                  " AND LC08 IS NULL AND LC34 is null" & _
                  " AND LC01 IS NOT NULL"
         cnnConnection.Execute strSql, intI
         '已收文未發文,2個工作天後達承辦期限者(不含當日) --承辦人-B1(承辦期限,承辦人)
         'Modified by Lydia 2016/10/20 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
         strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','B' EV1,'1' EV2,CP09,CP06,CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
                  "From CASEPROGRESS, LawCase" & _
                  " WHERE CP05>20030000" & _
                  " AND (CP14 IN(" & strOtherUser & ") or CP29 IN(" & strOtherUser & ")) AND CP158=0 AND CP159=0" & stConCP48 & stCP01 & _
                  " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
                  " AND LC08 IS NULL AND LC34 is null" & _
                  " AND LC01 IS NOT NULL"
         cnnConnection.Execute strSql, intI
         '已收文已發文,發文日期+6天未回執者 --承辦人-H1(發文日期,承辦人)
         strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','H' EV1,'1' EV2,CP09,to_number(to_char(to_date(cp27,'YYYYMMDD')+6,'YYYYMMDD')),CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
                  "From CASEPROGRESS, LawCase" & _
                  " WHERE cp27>=20110701 and cp50 is not null" & _
                  " AND (CP14 IN(" & strOtherUser & ") or CP29 IN(" & strOtherUser & ")) AND CP57||cp46 is null AND CP27 is not null" & _
                  " AND CP27<=" & stDate5 & stCP01 & _
                  " AND LC01=CP01 AND LC02=CP02 AND LC03=CP03 AND LC04=CP04" & _
                  " AND LC08 IS NULL AND LC34 is null" & _
                  " AND LC01 IS NOT NULL"
         cnnConnection.Execute strSql, intI
         '所有未發文--承辦人-E(未發文)
         If idx = 1 Then
            'Modified by Lydia 2016/10/20 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
            'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (cp06||cp48 is null or (cp06>=" & stDate2 & " and cp48 is null) or (cp06 is null and cp48>=" & stDate2 & ") or (cp06>=" & stDate2 & " and cp48>=" & stDate2 & "))"
            strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                     " SELECT '" & strUserNum & "','E' EV1,'1' EV2,CP09,CP06,CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
                     "From CASEPROGRESS, LawCase" & _
                     " WHERE CP05>20030000" & _
                     " AND (CP14 IN(" & strOtherUser & ") or CP29 IN(" & strOtherUser & ")) AND CP158=0 AND CP159=0" & stCP01 & _
                     " AND LC01=CP01 AND LC02=CP02 AND LC03=CP03 AND LC04=CP04" & _
                     " AND LC08 IS NULL AND LC34 is null" & _
                     " AND LC01 IS NOT NULL"
            cnnConnection.Execute strSql, intI
         End If
      End If
      
      '[未分案]
      '已收文未發文,2個工作天後達本所期限者(不含當日)-A0(本所期限,未分案)
      'Modified by Lydia 2016/10/20 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
      strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CP09,CP06,CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
               "From CASEPROGRESS, LawCase" & _
               " WHERE CP05>20030000" & _
               " AND CP14||CP29 IS NULL AND CP158=0 AND CP159=0" & stConCP06 & stCP01 & _
               " AND CP10<>'9001'" & _
               " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
               " AND LC08 IS NULL AND LC34 is null"
      cnnConnection.Execute strSql, intI
      'Add By Sindy 2011/6/27 案件性質為9001者改為本所期限<=系統日+3個工作天
      'Modified by Lydia 2016/10/20 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
      strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','A' EV1,'0' EV2,CP09,CP06,CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
               "From CASEPROGRESS, LawCase" & _
               " WHERE CP05>20030000" & _
               " AND CP14||CP29 IS NULL AND CP158=0 AND CP159=0" & stConCP0603 & stCP01 & _
               " AND CP10='9001'" & _
               " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
               " AND LC08 IS NULL AND LC34 is null"
      cnnConnection.Execute strSql, intI
      '已收文未發文,2個工作天後達承辦期限者(不含當日)-B0(承辦期限,未分案)
      'Modified by Lydia 2016/10/20 CP57 is null AND CP27 is null => CP158=0 AND CP159=0
      strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','B' EV1,'0' EV2,CP09,CP06,CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
               "From CASEPROGRESS, LawCase" & _
               " WHERE CP05>20030000" & _
               " AND CP14||CP29 IS NULL AND CP158=0 AND CP159=0" & stConCP48 & stCP01 & _
               " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
               " AND LC08 IS NULL AND LC34 is null"
      cnnConnection.Execute strSql, intI
      '已收文已發文,發文日期+6天未回執者 --承辦人-H1(發文日期,承辦人)
      strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
               " SELECT '" & strUserNum & "','H' EV1,'0' EV2,CP09,to_number(to_char(to_date(cp27,'YYYYMMDD')+6,'YYYYMMDD')),CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
               "From CASEPROGRESS, LawCase" & _
               " WHERE cp27>=20110701 and cp50 is not null" & _
               " AND CP14||CP29 IS NULL AND CP57||cp46 is null AND CP27 is not null" & _
               " AND CP27<=" & stDate5 & stCP01 & _
               " AND LC01=CP01 AND LC02=CP02 AND LC03=CP03 AND LC04=CP04" & _
               " AND LC08 IS NULL AND LC34 is null" & _
               " AND LC01 IS NOT NULL"
      cnnConnection.Execute strSql, intI
      '未分案,所有未發文-E(未發文)
      If idx = 1 Then
        'Modified by Lydia 2016/10/20 CP27 is null AND CP57 is null => CP158=0 AND CP159=0
        'Modify By Sindy 2017/6/22 Mark : 未發文不管期限 " AND (cp06||cp48 is null or (cp06>=" & stDate2 & " and cp48 is null) or (cp06 is null and cp48>=" & stDate2 & ") or (cp06>=" & stDate2 & " and cp48>=" & stDate2 & "))"
         strSql = "INSERT INTO R082007(ID,EV1,EV2,CP09,CP06,CP07,CP48,EP08,NA16,CP14,CP13,NP22)" & _
                  " SELECT '" & strUserNum & "','E' EV1,'0' EV2,CP09,CP06,CP07,CP48,0+NULL EP08,'' NA16,CP14,CP13,0 " & _
                  "From CASEPROGRESS, LawCase" & _
                  " WHERE CP05>20030000" & _
                  " AND CP14||CP29 IS NULL AND CP158=0 AND CP159=0" & stCP01 & _
                  " AND LC01=CP01 AND LC02=CP02 AND LC03=CP03 AND LC04=CP04" & _
                  " AND LC08 IS NULL AND LC34 is null"
         cnnConnection.Execute strSql, intI
      End If
   End If
   
   '若同案有 'A達本所'期限,其他的就不要再顯示
   strSql = "delete from R082007 R1 where R1.ID='" & strUserNum & "' and R1.EV1<>'A'" & _
            " AND exists (select * from R030301 R2 where R2.ID='" & strUserNum & "' and R2.EV1='A' and R2.CP09=R1.CP09 and R2.np22=R1.np22)"
   cnnConnection.Execute strSql, intI
   
   'Modify By Sindy 2015/5/27 承辦人若無資料,則抓法務協辦人員
   strExc(0) = "SELECT distinct '' V," & _
      "NVL(lpad(SQLDateT(CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(CP07),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "nvl(S2.ST02,S4.ST02) 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','D','未收文','E','未發文','H','未回執') 事件," & _
      "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04) 本所案號," & _
      "DECODE(INSTR('020,013',LC15),0,CPM03,CPM04) 案件性質," & _
      "CP64 案件備註," & _
      "LC05 案件名稱," & _
      "FA10 代理人國籍," & _
      "EV1,EV2,NA16,CP14,CP13,CP01,CP02,CP03,CP04,'' 未收款,CP10,CP27,CP09,CP29" & _
      " FROM R030301,Lawcase,caseprogress,STAFF S1,STAFF S2,STAFF S3,STAFF S4,CASEPROPERTYMAP" & _
      " WHERE ID='" & strUserNum & "' AND caseprogress.CP09(+)=R030301.CP09 AND NP22=0" & _
      " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16" & _
      " AND S2.ST01(+)=CP14" & _
      " AND S3.ST01(+)=CP13" & _
      " AND S4.ST01(+)=CP29" & _
      " AND CPM01(+)=CP01" & _
      " AND CPM02(+)=CP10"
   'Add By Sindy 2018/11/27
   '下一程序
   strExc(0) = strExc(0) & " UNION " & _
      "SELECT distinct '' V," & _
      "NVL(lpad(SQLDateT(CP06),9,' '),'2') 本所期限," & _
      "NVL(lpad(SQLDateT(CP07),9,' '),'2') 法定期限," & _
      "NVL(lpad(SQLDateT(CP48),10,' '),'2') 承辦期限," & _
      "NVL(lpad(SQLDateT(EP08),10,' '),'2') 核稿期限," & _
      "S1.ST02 管制人," & _
      "nvl(S2.ST02,S4.ST02) 承辦人," & _
      "S3.ST02 智權人員," & _
      "DECODE(EV1,'A','達本所','B','達承辦','D','未收文','E','未發文','H','未回執') 事件," & _
      "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04) 本所案號," & _
      "DECODE(INSTR('020,013',LC15),0,CPM03,CPM04) 案件性質," & _
      "CP64 案件備註," & _
      "LC05 案件名稱," & _
      "FA10 代理人國籍," & _
      "EV1,EV2,NA16,CP14,CP13,CP01,CP02,CP03,CP04,'' 未收款,CP10,CP27,CP09,CP29" & _
      " FROM R030301,Lawcase,NEXTPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,CASEPROPERTYMAP" & _
      " WHERE ID='" & strUserNum & "' AND NP01(+)=R030301.CP09 AND NEXTPROGRESS.NP22(+)=R030301.NP22 AND R030301.NP22>0" & _
      " AND LC01(+)=NP02 AND LC02(+)=NP03 AND LC03(+)=NP04 AND LC04(+)=NP05 AND LC01 IS NOT NULL" & _
      " AND S1.ST01(+)=NA16" & _
      " AND S2.ST01(+)=CP14" & _
      " AND S3.ST01(+)=CP13" & _
      " AND S4.ST01(+)=CP29" & _
      " AND CPM01(+)=CP01" & _
      " AND CPM02(+)=CP10"
   'Modify By Sindy 2015/1/30
   strExc(0) = strExc(0) & " order by 本所期限 asc,承辦人 asc,本所案號 asc"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If RsTemp Is Nothing Then Exit Sub
'   If RsTemp.RecordCount = 0 Then
'      Set m_adoRst = RsTemp.Clone
'      SetRst2Grid
'      MsgBox "查無資料！", vbInformation
'      cmdHide.Enabled = False
'      lblCnt.Caption = "共 0 筆"
'   Else
'      'Modify by Amy 2014/06/09 +FormName 改暫存TB
'      'Set m_adoRst = PUB_CreateRecordset(RsTemp, , , 300)
'      Set m_adoRst = PUB_CreateRecordset(RsTemp, , , 300, Me.Name)
'      m_stSort = "本所期限 asc,承辦人 asc,本所案號 asc"
'      m_adoRst.Sort = m_stSort
'      SetRst2Grid
'      SetGrid
'      RecordShow
'
'      SetColor
'      cmdHide.Enabled = True
'      m_blnColOrderAsc = True
'   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grdDataList.Recordset = rsTmp
      SetGrid
      RecordShow
      
      SetColor
      cmdHide.Enabled = True
      m_blnColOrderAsc = True
   Else
      Screen.MousePointer = vbDefault
      MsgBox "查無資料！", vbInformation
      rsTmp.Close
      Set rsTmp = Nothing
      cmdHide.Enabled = False
      LblCnt.Caption = "共 0 筆"
      Exit Sub 'Add By Sindy 2014/9/17
   End If
   rsTmp.Close
End Sub

Private Sub SetGrid()
   With grdDataList
      .Visible = False
      .FontFixed.Size = 8
      .Font.Size = 9
      .FormatString = "V|本所期限 |法定期限 |承辦期限 |核稿期限 |管制人 |承辦人 |智權人員 |事件　 |本所案號　　|案件性質 |備註　　　　|案件名稱　　　"
      For intI = 0 To .Cols - 1
         .ColAlignment(intI) = 0
         'If (intI > 12 And intI < 23) Or intI > 23 Then
         If (intI > 3 And intI < 6) Or intI > 12 Then
            .ColWidth(intI) = 0
         End If
      Next
'      .ColWidth(23) = 700
'      .ColAlignment(23) = flexAlignRightTop
      .ColAlignment(1) = flexAlignRightTop
      .ColAlignment(2) = flexAlignRightTop
      .ColAlignment(3) = flexAlignRightTop
      .ColAlignment(4) = flexAlignRightTop
      .Visible = True
   End With
End Sub

Private Sub SetColor(Optional sHide As String = "N")
   Dim lngToday As Long, lngCP06 As Long, lngCP48 As Long, lngEP08 As Long, stType As String
   Dim ii As Integer, jj As Integer, dblCnt As Double
   
   dblCnt = 0
   With grdDataList
   If .Rows > 1 Then
      .Visible = False
      ChgEmptyDate False
      lngToday = Val(strSrvDate(2))
      For ii = 1 To .Rows - 1
         lngCP06 = Val(Replace(.TextMatrix(ii, 1), "/", ""))
         lngCP48 = Val(Replace(.TextMatrix(ii, 3), "/", ""))
         lngEP08 = Val(Replace(.TextMatrix(ii, 4), "/", ""))
         stType = .TextMatrix(ii, 14)
         .row = ii
         '固定欄位變回白色
         For jj = 0 To 2
            .col = jj
            .CellBackColor = &HFFFFFF
            .CellAlignment = flexAlignRightTop
            .CellFontSize = 9
         Next
         '逾管控期限
         If ((stType = "A" Or stType = "D" Or stType = "H") And lngCP06 > 0 And lngCP06 < lngToday) Or _
            ((stType = "B") And lngCP48 > 0 And lngCP48 < lngToday) Then
            .TextMatrix(ii, 9) = "*" & .TextMatrix(ii, 9)
            For jj = 1 To .Cols - 1
               .col = jj
               '紅
               .CellBackColor = &HFF&
            Next
         '當日期限
         ElseIf ((stType = "A" Or stType = "D" Or stType = "H") And lngCP06 = lngToday) Or _
            ((stType = "B") And lngCP48 = lngToday) Then
            .TextMatrix(ii, 9) = "v" & .TextMatrix(ii, 9)
            For jj = 1 To .Cols - 1
               .col = jj
               '綠
               .CellBackColor = &HC000&
            Next
         '未分案
         ElseIf .TextMatrix(ii, 15) = "0" Then
            '第五級不看
            If bLvl5 = True Then
               .RowHeight(ii) = 0
            Else
               For jj = 1 To .Cols - 1
                  .col = jj
                  '黃
                  .CellBackColor = &HFFFF&
               Next
            End If
         ElseIf sHide <> "N" Then
            .RowHeight(ii) = 0
         Else
            strExc(1) = .TextMatrix(ii, 15)
            Select Case strExc(1)
               '承辦人,核稿人,法務協辦人員
               Case "1", "4"
                  strExc(2) = .TextMatrix(ii, 17)
               Case "2" '管制人
                  strExc(2) = .TextMatrix(ii, 16)
               Case "3" '智權人員
                  strExc(2) = .TextMatrix(ii, 18)
               Case Else
                  strExc(2) = ""
            End Select
            
            If strExc(2) <> "" Then
               '本人或第二級才看
               'If InStr(stNumList1(1) & "," & stNumList1(2), strExc(2)) = 0 Then
               If InStr(stNumList1(1) & "," & stNumList1(2), strExc(2)) = 0 And _
                  InStr(stNumList1(1) & "," & stNumList1(2), .TextMatrix(ii, 27)) = 0 Then
                  .RowHeight(ii) = 0
               End If
            End If
         End If
         If .RowHeight(ii) > 0 Then
            dblCnt = dblCnt + 1
         End If
      Next
      .TopRow = 1
      .Visible = True
   End If
   End With
   If sHide = "N" Then
      cmdHide.Tag = "Y"
      cmdHide.Caption = "隱藏白色(&H)"
   Else
      cmdHide.Tag = "N"
      cmdHide.Caption = "顯示白色(&S)"
   End If
   LblCnt.Caption = "共 " & dblCnt & " 筆"
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Combo1.Clear
   Combo1.AddItem "紅色(*)：表示逾管控期限"
   Combo1.AddItem "綠色(v)：表示當日期限"
   Combo1.AddItem "黃色(#)：表示未分案"
   Combo1.AddItem "藍色：表示點選資料"
   Combo1.ListIndex = 0
   txtUsernum = strUserNum
   If Pub_StrUserSt03 = "M51" Then
      txtUsernum.Enabled = True
   End If
   'Modify by Morgan 2011/4/21 從 Unload 移來(因畫面沒離開時沒寫Log會造成逾時重新登入後重複執行)
   PUB_AddExcuteLog Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If Not bolUnloading Then 'Add by Sindy 2016/7/22
      Dim strSql As String, bolRun As Boolean
      
      '電腦中心除外
      If Pub_StrUserSt03 <> "M51" Then
         '(內)法務
         bolRun = False
         'Modified by Lydia 2016/10/20 cp27||cp57 is null => CP158=0 AND CP159=0
         strSql = "select count(*) from caseprogress " & _
                        "where cp01 in ('L') " & _
                        "and CP158=0 AND CP159=0 and cp06>0 " & _
                        "and cp14='" & strUserNum & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               bolRun = True
            End If
         End If
         If strGroup = "F1" Or strGroup = "F2" Or strGroup = "D4" Or strGroup = "G1" Or bolRun = True Then
            If CheckUse("frm072005", strExec, False) = True Or bolRun = True Then
               strSql = "select * from executelog where el01='frm072005' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI <> 1 Then
                  pub_bolInformCheck = True
                  Load frm072005
                  frm072005.cmdQuery(0).Value = True
                  Exit Sub
               End If
            End If
         End If
         '專利
         bolRun = False
         'Modified by Lydia 2016/09/14 cp27||cp57 is null => NVL(CP27,0)=0 AND NVL(CP57,0)=0
         strSql = "select sum(aa) from ( " & _
                        "select count(*) as aa from caseprogress " & _
                        "where cp01 in ('FCP','FG') " & _
                        "and NVL(CP27,0)=0 AND NVL(CP57,0)=0 " & _
                        "and (cp06>0 or cp48>0) " & _
                        "and cp13='" & strUserNum & "' " & _
                        "Union All " & _
                        "select count(*) as aa from caseprogress " & _
                        "where cp01 in ('FCP','FG') " & _
                        "and NVL(CP27,0)=0 AND NVL(CP57,0)=0 " & _
                        "and (cp06>0 or cp48>0) " & _
                        "and cp14='" & strUserNum & "')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               bolRun = True
            End If
         End If
         If CheckUse("frm060204", strExec, False) = True Or bolRun = True Then
            strSql = "select * from executelog where el01='frm060204' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI <> 1 Then
               pub_bolInformCheck = True 'Add By Sindy 2009/09/21
               Load frm060204
               frm060204.cmdQuery(0).Value = True
               Exit Sub
            End If
         End If
         '商標
         bolRun = False
         'Modified by Lydia 2016/09/14 cp27||cp57 is null => NVL(CP27,0)=0 AND NVL(CP57,0)=0
         strSql = "select sum(aa) from ( " & _
                        "select count(*) as aa from caseprogress " & _
                        "where cp01 in ('FCT','CFT','CFC','S') " & _
                        "and NVL(CP27,0)=0 AND NVL(CP57,0)=0 " & _
                        "and (cp06>0 or cp48>0) " & _
                        "and cp13='" & strUserNum & "' " & _
                        "Union All " & _
                        "select count(*) as aa from caseprogress " & _
                        "where cp01 in ('FCT','CFT','CFC','S') " & _
                        "and NVL(CP27,0)=0 AND NVL(CP57,0)=0 " & _
                        "and (cp06>0 or cp48>0) " & _
                        "and cp14='" & strUserNum & "')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               bolRun = True
            End If
         End If
         If CheckUse("frm030301", strExec, False) = True Or bolRun = True Then
            'Modify By Sindy 2022/12/14 外商系統的期限彈跳提醒自動執行功能，請改為早上及下午各一次
            'strSql = "select * from executelog where el01='frm030301' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
            'Modify By Sindy 2025/9/3 琬姿副理在反應期限會一直啟動,故再調整一下判斷
            If ServerTime >= 130000 Then
               strSql = "select * from executelog where el01='frm030301' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
            Else
               strSql = "select * from executelog where el01='frm030301' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
            End If
            '2025/9/3 END
            'strSQL = "select * from executelog where el01='frm030301' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
            'strSql = "select * from executelog where el01='frm030301' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI <> 1 Then
               pub_bolInformCheck = True 'Add By Sindy 2009/09/21
               Load frm030301
               frm030301.cmdQuery(0).Value = True
               Exit Sub
            End If
         End If
      End If
      
      MenuEnabled
   End If
   
   Set frm082007 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub

'Modify By Sindy 2015/1/30
Private Sub GrdDataList_Click()
Dim nCol As Integer, nRow As Integer
      
   With grdDataList
      .Visible = False
      nCol = .MouseCol
      nRow = .MouseRow
      If nRow = 0 Then
         .col = nCol
         If m_blnColOrderAsc = False Then '字串降冪
            .Sort = 5 '字串昇冪
            m_blnColOrderAsc = True
         Else
            .Sort = 6 '字串降冪
            m_blnColOrderAsc = False
         End If
      End If
      .Visible = True
   End With
End Sub

'Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim iCol As Integer
'    iCol = grdDataList.MouseCol
'    If grdDataList.MouseRow < 1 Then
'      grdDataList.Visible = False
'      ChgEmptyDate True
'      Set grdDataList.Recordset = Nothing
'      If m_blnColOrderAsc = True Then
'         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " desc," & m_stSort
'         m_blnColOrderAsc = False
'      Else
'         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " asc," & m_stSort
'         m_blnColOrderAsc = True
'      End If
'      SetRst2Grid
'      SetGrid
'      SetColor
'      grdDataList.Visible = True
'    End If
'End Sub

Private Sub ChgEmptyDate(Optional p_bolBeforeSort As Boolean)
   Dim ii As Integer, jj As Integer
   With grdDataList
   If .Rows > 1 Then
      For ii = 1 To .Rows - 1
         For jj = 1 To 4
            If p_bolBeforeSort Then
               If .TextMatrix(ii, jj) = "" Then
                  .TextMatrix(ii, jj) = "2"
               End If
            Else
               If .TextMatrix(ii, jj) = "2" Then
                  .TextMatrix(ii, jj) = ""
               End If
            End If
         Next
      Next
   End If
   End With
End Sub

Private Sub grdDataList_SelChange()
   Dim ii As Integer, lngColor As Long
   With grdDataList
      If .MouseRow > 0 Then
         .Visible = False
         .row = .MouseRow
         .col = 0
         If .Text = "V" Then
            .Text = ""
            .col = 0
            .CellBackColor = .BackColor
            .col = 3
            lngColor = .CellBackColor
            For ii = 1 To 2
               .col = ii
               .CellBackColor = lngColor
            Next
         Else
            .Text = "V"
            For ii = 0 To 2
               .col = ii
               .CellBackColor = &HFFC0C0
            Next
         End If
         .Visible = True
      End If
   End With
End Sub

Private Sub txtUsernum_Change()
   If Len(txtUsernum) >= 5 Then
      lblUserName = GetStaffName(txtUsernum, True)
   Else
      lblUserName = ""
   End If
End Sub

Private Sub txtUsernum_GotFocus()
   TextInverse txtUsernum
End Sub

'Add By Sindy 2010/11/26
Private Sub txtUsernum_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
