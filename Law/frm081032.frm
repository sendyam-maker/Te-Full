VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm081032 
   BorderStyle     =   1  '單線固定
   Caption         =   "ACS案件期限通知"
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
      TabIndex        =   11
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
      ItemData        =   "frm081032.frx":0000
      Left            =   6150
      List            =   "frm081032.frx":0002
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
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V|本所期限|法定期限|承辦期限|核稿期限|管制人|承辦人|智權人員|事件　|本所案號　　　|案件性質|備註　　　　|案件名稱　　　"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
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
      _Band(0).Cols   =   13
   End
   Begin MSForms.Label lblUserName 
      Height          =   255
      Left            =   2100
      TabIndex        =   13
      Top             =   525
      Width           =   1710
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "3016;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCnt 
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7620
      TabIndex        =   12
      Top             =   5160
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
Attribute VB_Name = "frm081032"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/23 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblUserName
'Create by Lydia 2020/12/09 ACS案件期限通知
Option Explicit

Dim bolBarShow As Boolean
Public cmdState As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_stSort As String '排序方式
Dim m_adoRst As ADODB.Recordset
Dim stNumList1(1 To 5) As String
Dim colEV1 As Integer, colEV2 As Integer '欄位位置：代碼1、代碼2、本所案號(加註記)、本所案號、承辦人、智權人員
Dim colCno As Integer, colCP14 As Integer, colCP13 As Integer  '欄位位置：本所案號(加註記)、承辦人、智權人員

Private Sub cmdHelp_Click()
   strExc(0) = "" & vbCrLf
   strExc(0) = strExc(0) & "１達本所：本所期限＜＝系統日＋５個工作天（含當天）之未發文未取消收文案件。" & vbCrLf
   strExc(0) = strExc(0) & vbCrLf
   strExc(0) = strExc(0) & "２未發文：已收文未發文之案件。（按所有未發文按鈕才會顯示）" & vbCrLf
   MsgBox strExc(0), vbOKOnly, "事件說明"
End Sub

Private Sub cmdHide_Click()
   SetRst2Grid
   SetColor cmdHide.Tag
End Sub

Private Sub SetRst2Grid()
   grdDataList.FixedCols = 0
   Set grdDataList.Recordset = m_adoRst
   grdDataList.FixedCols = 3
End Sub

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData(Optional bolRefresh As Boolean = False)
   Dim i As Integer, lngColor As Long, ii As Integer
   Dim Str01 As String, StrTag As String
   
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
                  
         StrTag = grdDataList.TextMatrix(i, colCno)
         StrTag = Pub_RplStr(StrTag) '清掉符號
         If InStr(StrTag, "-") > 0 Then
             If InStrRev(StrTag, "-") < 6 Then StrTag = StrTag & "-0-00"
         End If
           
         Str01 = SystemNumber(StrTag, 1)
         If fnSaveParentForm(Me) = False Then
            Exit For
         End If
         Me.Show
         Select Case cmdState
            Case 2 '案件基本資料
               Select Case Pub_RplStr(Str01)
                  Case "L", "ACS" '法務
                      Screen.MousePointer = vbHourglass
                      frm100101_5.Show
                      frm100101_5.Tag = StrTag
                      frm100101_5.StrMenu
                      Screen.MousePointer = vbDefault
                  Case "LA" '顧問
                      Screen.MousePointer = vbHourglass
                      frm100101_6.Show
                      frm100101_6.Tag = StrTag
                      frm100101_6.StrMenu
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

Private Sub doQuery(idx As Integer)
Dim stVTB As String, stDate1 As String, stDate2 As String, stDate5 As String, stDate3 As String, stDate6 As String
Dim stConCP06 As String, stConCP48 As String, stConCP0603 As String
Dim stNumList As String, stCP01 As String, ii As Integer
Dim stUserID As String, strOtherUser As String
Dim rsTmp As New ADODB.Recordset
Dim bolCaseMan As Boolean 'ACS分案人員

   If lblUserName = "" Then
      MsgBox "員工編號錯誤！"
      If txtUsernum.Enabled = True Then
         txtUsernum_GotFocus
         txtUsernum.SetFocus
      End If
      Exit Sub
   End If
   
   stCP01 = " and cp01 ='ACS' "
   stUserID = txtUsernum
  
   stNumList = "'" & stUserID & "'"
   stNumList1(1) = stNumList  '1級=自己
   
   '期限管制人
   For ii = 2 To 5
      stNumList1(ii) = GetNumList(stUserID, ii)
      If stNumList1(ii) <> "" Then
         stNumList = stNumList & "," & stNumList1(ii)
      End If
   Next
   
   stDate1 = strSrvDate(1) - 10000
   stDate2 = CompWorkDay(3, strSrvDate(1)) '+2個工作天
   stDate5 = CompWorkDay(6, strSrvDate(1)) '+5個工作天(含當天)
   
   stConCP06 = " AND CP06>=" & stDate1 & " AND CP06< " & stDate5
   
   '特殊權限
   bolCaseMan = False
   If InStr(Pub_GetSpecMan("ACS分案人員") & ",", stUserID) > 0 Then
      bolCaseMan = True
   End If
   
   Call SetGrid(True) '清空
   '清除暫存檔
   strSql = "delete R081032 where ID='" & strUserNum & "'"
   cnnConnection.Execute strSql, intI
   
   '代碼1:A=達本所,B=達承辦,C=達核稿,D=未收文,E=未發文,F=未請款,G=未交稿,H=未回執
   '代碼2:0=未分案,1=承辦人,2=管制人,3=智權人員,4=核稿人,8=無核稿人 9=未交稿
   
   '已收文未發文,5個工作天後達本所期限者(不含當日) --承辦人-A1(本所期限,承辦人)
   strSql = "INSERT INTO R081032(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13) " & _
       "SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,CP06,CP07,CP48,CP14,CP13 " & _
         "From CASEPROGRESS, LawCase" & _
       " WHERE CP05>20030000" & _
         " AND CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & stCP01 & stConCP06 & _
         " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
         " AND LC08 IS NULL AND LC34 is null" & _
         " AND LC01 IS NOT NULL" & _
         " AND not exists (select * from R081032 where ID='" & strUserNum & "' and EV1='A' and EV2='1' )"
   cnnConnection.Execute strSql, intI
         
   '所有未發文--承辦人-E(未發文)
   If idx = 1 Then
      strSql = "INSERT INTO R081032(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13) " & _
       "SELECT '" & strUserNum & "','E' EV1,'1' EV2,CP09,CP06,CP07,CP48,CP14,CP13 " & _
        " From CASEPROGRESS, LawCase" & _
       " WHERE CP05>20030000" & _
         " AND CP14 IN(" & stNumList & ") AND CP158=0 AND CP159=0" & stCP01 & _
         " AND (CP06 IS NULL OR CP06>=" & stDate2 & ") AND LC01=CP01 AND LC02=CP02 AND LC03=CP03 AND LC04=CP04" & _
         " AND LC08 IS NULL AND LC34 is null" & _
         " AND LC01 IS NOT NULL" & _
         " AND not exists (select * from R081032 where ID='" & strUserNum & "' and EV1='E' and EV2='1' )"
      cnnConnection.Execute strSql, intI
   End If
   
   '未分案-0
   If bolCaseMan = True Then
      '承辦人為非創新業務部
      'modify by sonia 2023/12/22 st03 not like 'W%' 改為 st03<>'W20'
      strSql = "select distinct cp14 from caseprogress " & _
                     "where cp01 in ('ACS') " & _
                     "and CP158=0 AND CP159=0 " & _
                     "and cp06 is not null " & _
                     "and exists (select * from staff where st01=cp14 and st03<>'W20' ) "
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
      If strOtherUser <> "" Then
         '已收文未發文,5個工作天後達本所期限者(不含當日) --承辦人-A1(本所期限,承辦人)
         strSql = "INSERT INTO R081032(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13) " & _
            "SELECT '" & strUserNum & "','A' EV1,'1' EV2,CP09,CP06,CP07,CP48,CP14,CP13 " & _
              "From CASEPROGRESS, LawCase" & _
            " WHERE CP05>20030000" & _
              " AND CP14 IN(" & strOtherUser & ") AND CP158=0 AND CP159=0" & stCP01 & stConCP06 & _
              " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
              " AND LC08 IS NULL AND LC34 is null" & _
              " AND LC01 IS NOT NULL" & _
              " AND not exists (select * from R081032 where ID='" & strUserNum & "' and EV1='A' and EV2='1' )"
         cnnConnection.Execute strSql, intI
         
         '所有未發文--承辦人-E(未發文)
         If idx = 1 Then
            strSql = "INSERT INTO R081032(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13) " & _
             "SELECT '" & strUserNum & "','E' EV1,'1' EV2,CP09,CP06,CP07,CP48,CP14,CP13 " & _
               "From CASEPROGRESS, LawCase" & _
             " WHERE CP05>20030000" & _
               " AND CP14 IN(" & strOtherUser & ") AND CP158=0 AND CP159=0" & stCP01 & _
               " AND (CP06 IS NULL OR CP06>=" & stDate2 & ") AND LC01=CP01 AND LC02=CP02 AND LC03=CP03 AND LC04=CP04" & _
               " AND LC08 IS NULL AND LC34 is null" & _
               " AND LC01 IS NOT NULL" & _
               " AND not exists (select * from R081032 where ID='" & strUserNum & "' and EV1='E' and EV2='1' )"
            cnnConnection.Execute strSql, intI
         End If
      End If
      
      '[未分案]
      '已收文未發文,5個工作天後達本所期限者(不含當日)-A0(本所期限,未分案)
      strSql = "INSERT INTO R081032(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13) " & _
         "SELECT '" & strUserNum & "','A' EV1,'0' EV2,CP09,CP06,CP07,CP48,CP14,CP13 " & _
           "From CASEPROGRESS, LawCase" & _
         " WHERE CP05>20030000" & _
           " AND CP14 IS NULL AND CP158=0 AND CP159=0" & stCP01 & stConCP06 & _
           " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04" & _
           " AND LC08 IS NULL AND LC34 is null" & _
           " AND not exists (select * from R081032 where ID='" & strUserNum & "' and EV1='A' and EV2='0' )"
      cnnConnection.Execute strSql, intI

      '未分案,所有未發文-E(未發文)
      If idx = 1 Then
         strSql = "INSERT INTO R081032(ID,EV1,EV2,CP09,CP06,CP07,CP48,CP14,CP13) " & _
           "SELECT '" & strUserNum & "','E' EV1,'0' EV2,CP09,CP06,CP07,CP48,CP14,CP13 " & _
             "From CASEPROGRESS, LawCase" & _
           " WHERE CP05>20030000" & _
             " AND CP14 IS NULL AND CP158=0 AND CP159=0" & stCP01 & _
             " AND (CP06 IS NULL OR CP06>=" & stDate2 & ") AND LC01=CP01 AND LC02=CP02 AND LC03=CP03 AND LC04=CP04" & _
             " AND LC08 IS NULL AND LC34 is null" & _
             " AND not exists (select * from R081032 where ID='" & strUserNum & "' and EV1='E' and EV2='0')"
         cnnConnection.Execute strSql, intI
      End If
   End If
   
   '若同案有 'A達本所'期限,其他的就不要再顯示
   strSql = "delete from R081032 R1 where R1.ID='" & strUserNum & "' and R1.EV1<>'A'" & _
            " AND exists (select * from R081032 R2 where R2.ID='" & strUserNum & "' and R2.EV1='A' and R2.CP09=R1.CP09) "
   cnnConnection.Execute strSql, intI
   
   'SORT: 無所限CP06排最下面=>DECODE(VT1.CP06,2,9,NULL,9,1)
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',LC15),0,CPM03,CPM04) => DECODE(LC15,'000',CPM03,CPM04)
   strExc(0) = "SELECT DISTINCT '' V,NVL(LPAD(SQLDATET(VT1.CP06),9,' '),'2') 本所期限," & _
       " NVL(LPAD(SQLDATET(VT1.CP07),9,' '),'2') 法定期限," & _
       " NVL(LPAD(SQLDATET(VT1.CP48),10,' '),'2') 承辦期限," & _
       " S2.ST02 承辦人,S3.ST02 智權人員," & _
       " DECODE(EV1,'A','達本所','B','達承辦','D','未收文','E','未發文','H','未回執') 事件," & _
       " C1.CP01||'-'||C1.CP02||DECODE(C1.CP03,'0','','-'||C1.CP03)||DECODE(C1.CP04,'00','','-'||C1.CP04) 本所案號," & _
       " DECODE(LC15,'000',CPM03,CPM04) 案件性質,C1.CP64 進度備註," & _
       " NVL(LC05,NVL(LC06,LC07)) 案件名稱, NVL(CU04,NVL(CU05,CU06)) 申請人," & _
       " EV1,EV2,VT1.CP14,VT1.CP13,LC01,LC02,LC03,LC04,C1.CP10,C1.CP09," & _
       " DECODE(VT1.CP06,2,9,NULL,9,1) SORT,LC01||'-'||LC02||'-'||LC03||'-'||LC04 AS CASENO" & _
       " FROM R081032 VT1,CASEPROGRESS C1,LAWCASE,STAFF S2,STAFF S3,CASEPROPERTYMAP,CUSTOMER" & _
       " WHERE ID='" & strUserNum & "' AND C1.CP09(+)=VT1.CP09" & _
       " AND LC01(+)=C1.CP01 AND LC02(+)=C1.CP02 AND LC03(+)=C1.CP03 AND LC04(+)=C1.CP04" & _
       " AND LC01 IS NOT NULL AND S2.ST01(+)=VT1.CP14 AND S3.ST01(+)=VT1.CP13" & _
       " AND CPM01(+)=C1.CP01 AND CPM02(+)=C1.CP10 AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+)"
   strExc(0) = strExc(0) & " order by sort asc, 本所期限 asc,承辦人 asc,本所案號 asc"
   
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
   If rsTmp Is Nothing Then Exit Sub
   If rsTmp.RecordCount = 0 Then
      Set m_adoRst = rsTmp.Clone
      SetRst2Grid
      MsgBox "查無資料！", vbInformation
      cmdHide.Enabled = False
      lblCnt.Caption = "共 0 筆"
   Else
      Set m_adoRst = PUB_CreateRecordset(rsTmp, , , 300, Me.Name)
      SetRst2Grid
      SetGrid
      RecordShow
      
      SetColor
      cmdHide.Enabled = True
      m_blnColOrderAsc = True
   End If
   
   Set rsTmp = Nothing
End Sub

Private Sub SetGrid(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer, iR As Integer
    
                                          '0       1                  2                  3                   4              5                  6           7                  8                  9                  10                 11
   arrGridHeadText = Array("V", "本所期限", "法定期限", "承辦期限", "承辦人", "智權人員", "事件", "本所案號", "案件性質", "進度備註", "案件名稱", "申請人", _
                                         "EV1", "EV2", "CP14", "CP13", "LC01", "LC02", "LC03", "LC04", "CP09", "CP10", "SORT", "CASENO")
                                         '12         13         14          15        16          17         18          19        20           21        22           23
   arrGridHeadWidth = Array(300, 900, 900, 0, 900, 900, 900, 1200, 1100, 1100, 1100, 1100)
                                         
   grdDataList.Visible = False
   grdDataList.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
         grdDataList.Clear
         grdDataList.Rows = 2
   End If
       
    For iRow = 0 To grdDataList.Cols - 1
       grdDataList.row = 0
       grdDataList.col = iRow
       grdDataList.Text = arrGridHeadText(iRow)
       If iRow <= UBound(arrGridHeadWidth) Then
            grdDataList.ColWidth(iRow) = arrGridHeadWidth(iRow)
       Else '申請人以後的欄位
            grdDataList.ColWidth(iRow) = 0
       End If
    Next
    
    If colEV1 = 0 Then
        colEV1 = PUB_MGridGetId("EV1", grdDataList)
        colEV2 = PUB_MGridGetId("EV2", grdDataList)
        colCP14 = PUB_MGridGetId("CP14", grdDataList)
        colCP13 = PUB_MGridGetId("CP13", grdDataList)
        colCno = PUB_MGridGetId("本所案號", grdDataList)
    End If
    
    grdDataList.Visible = True
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
         stType = "" & .TextMatrix(ii, colEV1)
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
            .TextMatrix(ii, colCno) = "*" & .TextMatrix(ii, colCno)
            For jj = 1 To .Cols - 1
               .col = jj
               '紅
               .CellBackColor = &HFF&
            Next
         '當日期限
         ElseIf ((stType = "A" Or stType = "D" Or stType = "H") And lngCP06 = lngToday) Or _
            ((stType = "B") And lngCP48 = lngToday) Then
            .TextMatrix(ii, colCno) = "v" & .TextMatrix(ii, colCno)
            For jj = 1 To .Cols - 1
               .col = jj
               '綠
               .CellBackColor = &HC000&
            Next
         '未分案
         ElseIf .TextMatrix(ii, colEV2) = "0" Then
            For jj = 1 To .Cols - 1
               .col = jj
               '黃
               .CellBackColor = &HFFFF&
            Next
         ElseIf sHide <> "N" Then
            .RowHeight(ii) = 0
         Else
            strExc(1) = .TextMatrix(ii, colEV2)
            Select Case strExc(1)
               '承辦人,核稿人,法務協辦人員
               Case "1", "4"
                  strExc(2) = .TextMatrix(ii, colCP14)
               Case "3" '智權人員
                  strExc(2) = .TextMatrix(ii, colCP13)
               Case Else
                  strExc(2) = ""
            End Select
            
            If strExc(2) <> "" Then
               '本人或第二級才看
               If InStr(stNumList1(1) & "," & stNumList1(2), strExc(2)) = 0 Then
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
   lblCnt.Caption = "共 " & dblCnt & " 筆"
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
   'Added by Lydia 2023/05/10 開放輸入「員工編號」欄：總經理
   If InStr("01,08,", Pub_strUserST05 & ",") > 0 Then
      txtUsernum.Enabled = True
   End If
   'end 2023/05/10
   
   PUB_AddExcuteLog Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Memo by Lydia 2019/11/04   國外部自動通知順序: FMP案frm060206=>
                                           '國外部期限frm060204=>
                                           '外商frm030301=>
                                           '外法frm072005=>
                                           '國外部行事曆frm060209
                                           'ACS案frm081032=> 'Added by Lydia 2020/12/10 在mdiMain(Law)
                                          
   If Not bolUnloading Then '
      Dim strSql As String, bolRun As Boolean
      
      '電腦中心除外
      If Pub_StrUserSt03 <> "M51" Then
         '專利
         bolRun = False
         strSql = "select sum(aa) from ( " & _
                        "select count(*) as aa from caseprogress " & _
                        "where cp01 in ('FCP','FG') " & _
                        "and cp158=0 AND cp159=0 " & _
                        "and (cp06>0 or cp48>0) " & _
                        "and cp13='" & strUserNum & "' " & _
                        "Union All " & _
                        "select count(*) as aa from caseprogress " & _
                        "where cp01 in ('FCP','FG') " & _
                        "and cp158=0 AND cp159=0 " & _
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
               pub_bolInformCheck = True
               Load frm060204
               frm060204.cmdQuery(0).Value = True
               Exit Sub
            End If
         End If
         '商標
         bolRun = False
         strSql = "select sum(aa) from ( " & _
                        "select count(*) as aa from caseprogress " & _
                        "where cp01 in ('FCT','CFT','CFC','S') " & _
                        "and cp158=0 AND cp159=0 " & _
                        "and (cp06>0 or cp48>0) " & _
                        "and cp13='" & strUserNum & "' " & _
                        "Union All " & _
                        "select count(*) as aa from caseprogress " & _
                        "where cp01 in ('FCT','CFT','CFC','S') " & _
                        "and cp158=0 AND cp159=0 " & _
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
            strSql = "select * from executelog where el01='frm030301' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
            'strSql = "select * from executelog where el01='frm030301' and el02='" & strUserNum & "' and el03=" & strSrvDate(1)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI <> 1 Then
               pub_bolInformCheck = True
               Load frm030301
               frm030301.cmdQuery(0).Value = True
               Exit Sub
            End If
         End If
         '行事曆
         If Left(Pub_StrUserSt03, 2) = "F2" Then
            If CheckUse("frm060209", strExec, False) = True Then
                strSql = "select * from executelog where el01='frm060209' and el02='" & strUserNum & "' and el03=" & strSrvDate(1) & " and el04>=decode(sign(to_char(sysdate,'hh24')-12),1,130000,0)"
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI <> 1 Then
                   pub_bolInformCheck = True
                   Load frm060209
                   pub_bolInformCheck = False
                   Exit Sub
                End If
            End If
         End If
         
      End If
      
      MenuEnabled
   End If
   
   Set frm081032 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow grdDataList, x, y, nCol, nRow
   grdDataList.col = nCol
   grdDataList.row = nRow
End Sub

Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iCol As Integer
    iCol = grdDataList.col
    If grdDataList.row < 1 Then
      grdDataList.Visible = False
      ChgEmptyDate True
      Set grdDataList.Recordset = Nothing
      If m_blnColOrderAsc = True Then
         m_adoRst.Sort = "sort asc, " & m_adoRst.Fields(iCol).Name & " desc"
         m_blnColOrderAsc = False
      Else
         m_adoRst.Sort = "sort asc, " & m_adoRst.Fields(iCol).Name & " asc"
         m_blnColOrderAsc = True
      End If
      SetRst2Grid
      SetGrid
      SetColor
      grdDataList.Visible = True
    End If
End Sub

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
      If .row > 0 Then
         .Visible = False
         .row = .row
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

Private Sub txtUsernum_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
