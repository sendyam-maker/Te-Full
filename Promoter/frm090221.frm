VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090221 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標委任書正本案號維護"
   ClientHeight    =   5730
   ClientLeft      =   4005
   ClientTop       =   2445
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdok 
      Caption         =   "儲存委任書註記(&S)"
      Height          =   400
      Index           =   2
      Left            =   5400
      TabIndex        =   2
      Top             =   30
      Width           =   1725
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "刪除註記(&D)"
      Height          =   400
      Index           =   3
      Left            =   7170
      TabIndex        =   3
      Top             =   30
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1170
      MaxLength       =   9
      TabIndex        =   0
      Top             =   120
      Width           =   1365
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   4500
      TabIndex        =   1
      Top             =   30
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   4665
      Left            =   60
      TabIndex        =   5
      Top             =   1020
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   8229
      _Version        =   393216
      Cols            =   13
      FixedCols       =   4
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
   Begin VB.CommandButton cmdok 
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
      Left            =   8370
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   30
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "符號說明：＊閉卷;△非申請人案;●銷卷;◎商標舊委任狀;□商標新委任狀"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   3420
      TabIndex        =   11
      Top             =   780
      Width           =   5760
   End
   Begin MSForms.Label Label7 
      Height          =   210
      Left            =   2070
      TabIndex        =   10
      Top             =   765
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "2302;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      Caption         =   " "
      Height          =   210
      Left            =   1170
      TabIndex        =   9
      Top             =   765
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   765
      Width           =   900
   End
   Begin MSForms.Label Label4 
      Height          =   240
      Left            =   1170
      TabIndex        =   7
      Top             =   495
      Width           =   5640
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "9948;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人編號："
      Height          =   180
      Left            =   60
      TabIndex        =   6
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "frm090221"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/22 改成Form2.0 (Label4,Label7)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Create By Sindy 2012/2/9
Option Explicit

Dim i As Integer, j As Integer
'紀錄作用按鍵
Public cmdState As Integer
Dim Str02 As String
Dim m_TM01 As String, m_TM02 As String, m_TM03 As String, m_TM04 As String
Public strSysID As String 'Add By Sindy 2014/7/2
Private Const cntLstPayYearSQL As String = " decode( substr(lpad(pa72,200,' '),200,1),' ',' ',decode( substr(lpad(pa72,200,' '),199,1),' ',substr(lpad(pa72,200,' '),200,1),',',substr(lpad(pa72,200,' '),200,1) ,decode( substr(lpad(pa72,200,' '),198,1),',',substr(lpad(pa72,200,' '),199,2) ,decode( substr(lpad(pa72,200,' '),197,1),',',substr(lpad(pa72,200,' '),198,3) ,decode( substr(lpad(pa72,200,' '),196,1),',',substr(lpad(pa72,200,' '),197,4) ) ) ) ) )"


'Modify By Sindy 2014/7/2 +strSysID
Private Sub SetDataListWidth(strSysID As String)
Dim arrGridHeadText, arrGridHeadWidth, iDep As String
Dim iCol As Integer

iDep = PUB_GetST06(strUserNum)

'Modify By Sindy 2014/7/2
If strSysID = "P" Then
   arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "申請國家" _
                  , "申請案號", "申請日", "專利號數", "准駁", "申請人1" _
                  , "專用期間", "公告號", "最後已繳年度", "申請人2" _
                  , "申請人3", "申請人4", "申請人5", "", "")
   '電腦中心，跟分所才秀
   If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
      arrGridHeadWidth = Array(200, 1600, 0, 1600, 800 _
                        , 1200, 800, 1150, 450, 1000 _
                        , 1800, 1000, 1200, 1100 _
                        , 1100, 1100, 1100, 0, 0)
   Else
      arrGridHeadWidth = Array(200, 1600, 620, 1600, 800 _
                        , 1200, 800, 1150, 450, 1000 _
                        , 1800, 1000, 1200, 1100 _
                        , 1100, 1100, 1100, 0, 0)
   End If
Else
'2014/7/2 END
   arrGridHeadText = Array("V", "本所案號", "分所號", "案件名稱", "申請國家" _
                  , "申請案號", "申請日", "審定號", "准駁", "申請人1" _
                  , "商品類別", "專用期間", "申請人2" _
                  , "申請人3", "申請人4", "申請人5", "", "")
   '電腦中心，跟分所才秀
   If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
      arrGridHeadWidth = Array(200, 1600, 0, 1600, 800 _
                        , 1200, 800, 1150, 450, 1000 _
                        , 800, 1800, 1000 _
                        , 1100, 1100, 1100, 0, 0)
   Else
      arrGridHeadWidth = Array(200, 1600, 620, 1600, 800 _
                        , 1200, 800, 1150, 450, 1000 _
                        , 800, 1800, 1000 _
                        , 1100, 1100, 1100, 0, 0)
   End If
End If

GrdDataList.Cols = UBound(arrGridHeadText) + 1
For iCol = 0 To GrdDataList.Cols - 1
   GrdDataList.row = 0
   GrdDataList.col = iCol
   GrdDataList.Text = arrGridHeadText(iCol)
   GrdDataList.ColWidth(iCol) = arrGridHeadWidth(iCol)
   GrdDataList.CellAlignment = flexAlignCenterCenter
Next iCol
End Sub

Public Sub PubShowNextData()
Dim bolSave As Boolean

On Error GoTo ErrHand

bolSave = False
Select Case cmdState
   Case 0 '結束
      Unload Me
      
   Case 1 '查詢
      Call StrMenu
      
   Case 2 '儲存委任書註記
      Me.Enabled = False
      Screen.MousePointer = vbHourglass
      cnnConnection.BeginTrans
      For i = 1 To GrdDataList.Rows - 1
         GrdDataList.col = 0
         GrdDataList.row = i
         If Trim(GrdDataList.Text) = "V" Then
            GrdDataList.col = 0
            GrdDataList.Text = ""
            '固定欄位不變色
            For j = 4 To GrdDataList.Cols - 1
               GrdDataList.col = j
               GrdDataList.CellBackColor = QBColor(15)
            Next j
            GrdDataList.col = 1
            If Not IsNull(GrdDataList.Text) Then
               strSql = Pub_RplStr(GrdDataList.Text)
               m_TM01 = SystemNumber(strSql, 1)
               m_TM02 = SystemNumber(strSql, 2)
               m_TM03 = SystemNumber(strSql, 3)
               m_TM04 = SystemNumber(strSql, 4)
               'Add By Sindy 2014/7/2
               If strSysID = "P" Then
                  strSql = "update patent set pa165='Y'" & _
                           " where pa01='" & m_TM01 & "' and pa02='" & m_TM02 & "' and pa03='" & m_TM03 & "' and pa04='" & m_TM04 & "'"
                  Pub_SeekTbLog strSql 'Added by Morgan 2022/9/2
                  cnnConnection.Execute strSql
               Else
               '2014/7/2 END
                  'Modify By Sindy 2012/5/29 2012年6月1日開始為A.新委任書
                  If strSrvDate(1) >= TMdebateStarDT Then
                     strSql = "update trademark set tm128='A'" & _
                              " where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "'"
                  Else
                  '2012/5/29 End
                     strSql = "update trademark set tm128='Y'" & _
                              " where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "'"
                  End If
                  Pub_SeekTbLog strSql 'Added by Morgan 2022/9/2
                  cnnConnection.Execute strSql
               End If
               bolSave = True
            End If
         End If
      Next i
      cnnConnection.CommitTrans
      If bolSave = True Then
         Call StrMenu
         MsgBox "儲存完畢！"
      End If
      Screen.MousePointer = vbDefault
      Me.Enabled = True
      
   Case 3 '刪除註記
      Me.Enabled = False
      Screen.MousePointer = vbHourglass
      cnnConnection.BeginTrans
      For i = 1 To GrdDataList.Rows - 1
         GrdDataList.col = 0
         GrdDataList.row = i
         If Trim(GrdDataList.Text) = "V" Then
            GrdDataList.col = 0
            GrdDataList.Text = ""
            '固定欄位不變色
            For j = 4 To GrdDataList.Cols - 1
               GrdDataList.col = j
               GrdDataList.CellBackColor = QBColor(15)
            Next j
            GrdDataList.col = 1
            If Not IsNull(GrdDataList.Text) Then
               strSql = Pub_RplStr(GrdDataList.Text)
               m_TM01 = SystemNumber(strSql, 1)
               m_TM02 = SystemNumber(strSql, 2)
               m_TM03 = SystemNumber(strSql, 3)
               m_TM04 = SystemNumber(strSql, 4)
               'Add By Sindy 2014/7/2
               If strSysID = "P" Then
                  strSql = "update patent set pa165=null" & _
                           " where pa01='" & m_TM01 & "' and pa02='" & m_TM02 & "' and pa03='" & m_TM03 & "' and pa04='" & m_TM04 & "'"
                  Pub_SeekTbLog strSql 'Added by Morgan 2022/9/2
                  cnnConnection.Execute strSql
               Else
               '2014/7/2 END
                  strSql = "update trademark set tm128=null" & _
                           " where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "'"
                  Pub_SeekTbLog strSql 'Added by Morgan 2022/9/2
                  cnnConnection.Execute strSql
               End If
               bolSave = True
            End If
         End If
      Next i
      cnnConnection.CommitTrans
      If bolSave = True Then
         Call StrMenu
         MsgBox "刪除完畢！"
      End If
      Screen.MousePointer = vbDefault
      Me.Enabled = True
   Case Else
End Select

Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Sub

Public Sub cmdOK_Click(Index As Integer)
'紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   'Modify By Sindy 2014/7/2
   If GetStaffDepartment(strUserNum) <> "M51" Then
      If Left(Pub_StrUserSt03, 2) = "P1" Then
         strSysID = "P"
      Else
         strSysID = "T"
      End If
   End If
   Call SetDataListWidth(strSysID)
   If strSysID = "P" Then
      Me.Caption = "專利總委任書正本案號維護"
      Label2.Caption = "符號說明：＊閉卷;△非申請人案;●銷卷;＃台灣專利總委任書"
   Else
      Me.Caption = "商標委任書正本案號維護"
      Label2.Caption = "符號說明：＊閉卷;△非申請人案;●銷卷;◎商標舊委任狀;□商標新委任狀"
   End If
   '2014/7/2 END
   cmdState = -1
End Sub

Sub StrMenu()
Dim strAppNo As String

Me.Enabled = False

Text1 = ChangeCustomerL(Text1)
strAppNo = Left(Text1, 8)
Str02 = GetAllSysKind(, "ALL")

'檢查國內外權限
If CheckSR12(Text1) = False Then
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If

'組字串
strSql = "SELECT NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),CU13,ST02,cu111,'' PCC05" & _
   " FROM CUSTOMER,STAFF WHERE CU01='" & Left$(Text1, 8) & "' AND CU02='" & Right$(Text1, 1) & "' AND CU13=ST01(+) "

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If IsNull(adoRecordset.Fields(0)) Then
        Label4.Caption = ""
    Else
        Label4.Caption = adoRecordset.Fields(0)
    End If
    If IsNull(adoRecordset.Fields(1)) Then
        Label6.Caption = ""
    Else
        Label6.Caption = adoRecordset.Fields(1)
    End If
    If IsNull(adoRecordset.Fields(2)) Then
        Label7.Caption = ""
    Else
        Label7.Caption = adoRecordset.Fields(2)
    End If
End If
CheckOC

'隱藏申請人1
GrdDataList.ColWidth(9) = 0

'Add By Sindy 2014/7/2
If strSysID = "P" Then
   strSql = "SELECT ' ' AS V,decode(PA23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(PA108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(PA136,'')),null,'','●')||PA47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
            ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(PA108,'')),null,'','●')||DECODE(PA165,'Y','＃','') as FSort,DECODE(PA149,NULL,C1.CU01||C1.CU127,C1.CU01||PA149) CNT FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5" & _
            " WHERE PA09=na01(+) and substr(PA26,1,8)='" & strAppNo & "' and PA01 in (" & SQLGrpStr(Str02, 1) & ") and PA09='000'" & _
            " and SUBSTR(PA26,1,8)=c1.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=c1.CU02(+)" & _
            " and substr(PA27,1,8)=c2.cu01(+) and decode(substr(PA27,9,1),null,'0',substr(PA27,9,1))=c2.cu02(+)" & _
            " and substr(PA28,1,8)=c3.cu01(+) and decode(substr(PA28,9,1),null,'0',substr(PA28,9,1))=c3.cu02(+)" & _
            " and substr(PA29,1,8)=c4.cu01(+) and decode(substr(PA29,9,1),null,'0',substr(PA29,9,1))=c4.cu02(+)" & _
            " and substr(PA30,1,8)=c5.cu01(+) and decode(substr(PA30,9,1),null,'0',substr(PA30,9,1))=c5.cu02(+)"
   strSql = strSql + " union SELECT ' ' AS V,decode(PA23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(PA108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(PA136,'')),null,'','●')||PA47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
            ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(PA108,'')),null,'','●')||DECODE(PA165,'Y','＃','') as FSort,DECODE(PA149,NULL,C1.CU01||C1.CU127,C1.CU01||PA149) CNT FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5" & _
            " WHERE PA09=na01(+) and substr(PA27,1,8)='" & strAppNo & "' and PA01 in (" & SQLGrpStr(Str02, 1) & ") and PA09='000'" & _
            " and SUBSTR(PA26,1,8)=c1.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=c1.CU02(+)" & _
            " and substr(PA27,1,8)=c2.cu01(+) and decode(substr(PA27,9,1),null,'0',substr(PA27,9,1))=c2.cu02(+)" & _
            " and substr(PA28,1,8)=c3.cu01(+) and decode(substr(PA28,9,1),null,'0',substr(PA28,9,1))=c3.cu02(+)" & _
            " and substr(PA29,1,8)=c4.cu01(+) and decode(substr(PA29,9,1),null,'0',substr(PA29,9,1))=c4.cu02(+)" & _
            " and substr(PA30,1,8)=c5.cu01(+) and decode(substr(PA30,9,1),null,'0',substr(PA30,9,1))=c5.cu02(+)"
   strSql = strSql + " union SELECT ' ' AS V,decode(PA23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(PA108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(PA136,'')),null,'','●')||PA47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度度" & _
            ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(PA108,'')),null,'','●')||DECODE(PA165,'Y','＃','') as FSort,DECODE(PA149,NULL,C1.CU01||C1.CU127,C1.CU01||PA149) CNT FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5" & _
            " WHERE PA09=na01(+) and substr(PA28,1,8)='" & strAppNo & "' and PA01 in (" & SQLGrpStr(Str02, 1) & ") and PA09='000'" & _
            " and SUBSTR(PA26,1,8)=c1.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=c1.CU02(+)" & _
            " and substr(PA27,1,8)=c2.cu01(+) and decode(substr(PA27,9,1),null,'0',substr(PA27,9,1))=c2.cu02(+)" & _
            " and substr(PA28,1,8)=c3.cu01(+) and decode(substr(PA28,9,1),null,'0',substr(PA28,9,1))=c3.cu02(+)" & _
            " and substr(PA29,1,8)=c4.cu01(+) and decode(substr(PA29,9,1),null,'0',substr(PA29,9,1))=c4.cu02(+)" & _
            " and substr(PA30,1,8)=c5.cu01(+) and decode(substr(PA30,9,1),null,'0',substr(PA30,9,1))=c5.cu02(+)"
   strSql = strSql + " union SELECT ' ' AS V,decode(PA23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(PA108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(PA136,'')),null,'','●')||PA47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
            ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(PA108,'')),null,'','●')||DECODE(PA165,'Y','＃','') as FSort,DECODE(PA149,NULL,C1.CU01||C1.CU127,C1.CU01||PA149) CNT FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5" & _
            " WHERE PA09=na01(+) and substr(PA29,1,8)='" & strAppNo & "' and PA01 in (" & SQLGrpStr(Str02, 1) & ") and PA09='000'" & _
            " and SUBSTR(PA26,1,8)=c1.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=c1.CU02(+)" & _
            " and substr(PA27,1,8)=c2.cu01(+) and decode(substr(PA27,9,1),null,'0',substr(PA27,9,1))=c2.cu02(+)" & _
            " and substr(PA28,1,8)=c3.cu01(+) and decode(substr(PA28,9,1),null,'0',substr(PA28,9,1))=c3.cu02(+)" & _
            " and substr(PA29,1,8)=c4.cu01(+) and decode(substr(PA29,9,1),null,'0',substr(PA29,9,1))=c4.cu02(+)" & _
            " and substr(PA30,1,8)=c5.cu01(+) and decode(substr(PA30,9,1),null,'0',substr(PA30,9,1))=c5.cu02(+)"
   strSql = strSql + " union SELECT ' ' AS V,decode(PA23,'1','','N')||PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(PA108,'')),null,'','●')||DECODE(PA165,'Y','＃','') AS 本所案號 ,DECODE(length(nvl(PA136,'')),null,'','●')||PA47 AS 分所號 ,NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,na03 AS 申請國家,PA11 AS 申請案號 ," & SQLDate("PA10") & " AS 申請日 ,PA22 AS 專利號數,DECODE(PA16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,DECODE(PA24,NULL,'','','',(SUBSTR(PA24,1,4)||'/'||SUBSTR(PA24,5,2)||'/'||SUBSTR(PA24,7,2)))||'-'||DECODE(PA25,NULL,'','','',(SUBSTR(PA25,1,4)||'/'||SUBSTR(PA25,5,2)||'/'||SUBSTR(PA25,7,2))) AS 專用期間,NVL(PA15,PA13) AS 專利公告號," & cntLstPayYearSQL & " AS 最近已繳年度" & _
            ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,PA01||'-'||PA02||'-'||PA03||'-'||PA04||DECODE(PA57,'Y','＊','')||DECODE(length(nvl(PA108,'')),null,'','●')||DECODE(PA165,'Y','＃','') as FSort,DECODE(PA149,NULL,C1.CU01||C1.CU127,C1.CU01||PA149) CNT FROM PATENT,nation,customer c1,customer c2,customer c3,customer c4,customer c5" & _
            " WHERE PA09=na01(+) and substr(PA30,1,8)='" & strAppNo & "' and PA01 in (" & SQLGrpStr(Str02, 1) & ") and PA09='000'" & _
            " and SUBSTR(PA26,1,8)=c1.CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=c1.CU02(+)" & _
            " and substr(PA27,1,8)=c2.cu01(+) and decode(substr(PA27,9,1),null,'0',substr(PA27,9,1))=c2.cu02(+)" & _
            " and substr(PA28,1,8)=c3.cu01(+) and decode(substr(PA28,9,1),null,'0',substr(PA28,9,1))=c3.cu02(+)" & _
            " and substr(PA29,1,8)=c4.cu01(+) and decode(substr(PA29,9,1),null,'0',substr(PA29,9,1))=c4.cu02(+)" & _
            " and substr(PA30,1,8)=c5.cu01(+) and decode(substr(PA30,9,1),null,'0',substr(PA30,9,1))=c5.cu02(+)"
Else
'2014/7/2 END
   'Modify By Sindy 2012/5/29 DECODE(TM128,'Y','◎','')==>DECODE(TM128,'Y','◎','A','□','')
   strSql = "SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定號,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間" & _
            ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,DECODE(TM123,NULL,C1.CU01||C1.CU127,C1.CU01||TM123) CNT FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5" & _
            " WHERE tm10=na01(+) and substr(TM23,1,8)='" & strAppNo & "' and tm01 in (" & SQLGrpStr(Str02, 2) & ") AND TM04='00'" & _
            " and SUBSTR(TM23,1,8)=c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1))=c1.CU02(+)" & _
            " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+)" & _
            " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+)" & _
            " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+)" & _
            " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+)"
   strSql = strSql + " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定號,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間" & _
            ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,DECODE(TM123,NULL,C2.CU01||C2.CU127,C1.CU01||TM123) CNT FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5" & _
            " WHERE tm10=na01(+) and substr(TM78,1,8)='" & strAppNo & "' and tm01 in (" & SQLGrpStr(Str02, 2) & ") AND TM04='00'" & _
            " and SUBSTR(TM23,1,8)=c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1))=c1.CU02(+)" & _
            " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+)" & _
            " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+)" & _
            " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+)" & _
            " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+)"
   strSql = strSql + " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定號,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間" & _
            ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,DECODE(TM123,NULL,C3.CU01||C3.CU127,C1.CU01||TM123) CNT FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5" & _
            " WHERE tm10=na01(+) and substr(TM79,1,8)='" & strAppNo & "' and tm01 in (" & SQLGrpStr(Str02, 2) & ") AND TM04='00'" & _
            " and SUBSTR(TM23,1,8)=c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1))=c1.CU02(+)" & _
            " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+)" & _
            " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+)" & _
            " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+)" & _
            " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+)"
   strSql = strSql + " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定號,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間" & _
            ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,DECODE(TM123,NULL,C4.CU01||C4.CU127,C1.CU01||TM123) CNT FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5" & _
            " WHERE tm10=na01(+) and substr(TM80,1,8)='" & strAppNo & "' and tm01 in (" & SQLGrpStr(Str02, 2) & ") AND TM04='00'" & _
            " and SUBSTR(TM23,1,8)=c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1))=c1.CU02(+)" & _
            " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+)" & _
            " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+)" & _
            " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+)" & _
            " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+)"
   strSql = strSql + " union SELECT ' ' AS V,decode(tm28,'1','','N')||TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') AS 本所案號 ,DECODE(length(nvl(tm73,'')),null,'','●')||tm34 AS 分所號 ,NVL(TM05,NVL(TM06,TM07)) AS 案件名稱,na03 AS 申請國家,TM12 AS 申請案號 ," & SQLDate("TM11") & " AS 申請日 ,tm15 AS 審定號,DECODE(TM16,'1','准','2','駁',' ') AS 准駁 , NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1 ,NVL(TM09,' ') AS 商品類別,DECODE(TM21,NULL,'','','',(SUBSTR(TM21,1,4)||'/'||SUBSTR(TM21,5,2)||'/'||SUBSTR(TM21,7,2)))||'-'||DECODE(TM22,NULL,'','','',(SUBSTR(TM22,1,4)||'/'||SUBSTR(TM22,5,2)||'/'||SUBSTR(TM22,7,2))) AS 專用期間" & _
            ",NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3,NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5,TM01||'-'||TM02||'-'||TM03||'-'||TM04||DECODE(TM29,'Y','＊','')||DECODE(length(nvl(tm57,'')),null,'','●')||DECODE(TM128,'Y','◎','A','□','') as FSort,DECODE(TM123,NULL,C5.CU01||C5.CU127,C1.CU01||TM123) CNT FROM TRADEMARK,nation,customer c1,customer c2,customer c3,customer c4,customer c5" & _
            " WHERE tm10=na01(+) and substr(TM81,1,8)='" & strAppNo & "' and tm01 in (" & SQLGrpStr(Str02, 2) & ") AND TM04='00'" & _
            " and SUBSTR(TM23,1,8)=c1.CU01(+) AND DECODE(SUBSTR(TM23,9,1),NULL,'0',SUBSTR(TM23,9,1))=c1.CU02(+)" & _
            " and substr(tm78,1,8)=c2.cu01(+) and decode(substr(tm78,9,1),null,'0',substr(tm78,9,1))=c2.cu02(+)" & _
            " and substr(tm79,1,8)=c3.cu01(+) and decode(substr(tm79,9,1),null,'0',substr(tm79,9,1))=c3.cu02(+)" & _
            " and substr(tm80,1,8)=c4.cu01(+) and decode(substr(tm80,9,1),null,'0',substr(tm80,9,1))=c4.cu02(+)" & _
            " and substr(tm81,1,8)=c5.cu01(+) and decode(substr(tm81,9,1),null,'0',substr(tm81,9,1))=c5.cu02(+)"
   strSql = strSql & " ORDER BY FSort,本所案號"
End If
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
GrdDataList.FixedCols = 0
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   adoRecordset.MoveFirst
End If
Set GrdDataList.Recordset = adoRecordset
Call SetDataListWidth(strSysID)
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   GrdDataList.FixedCols = 4
Else
   GrdDataList.AddItem ""
   ShowNoData
End If
Screen.MousePointer = vbDefault
Me.Enabled = True
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090221 = Nothing
End Sub

Private Sub grdDataList_SelChange()
GrdDataList.Visible = False
GrdDataList.row = GrdDataList.MouseRow
GrdDataList.col = 0
If GrdDataList.row <> 0 Then
If GrdDataList.Text = "V" Then
     GrdDataList.Text = ""
     For i = 4 To GrdDataList.Cols - 1
          GrdDataList.col = i
          GrdDataList.CellBackColor = QBColor(15)
    Next i
Else
     GrdDataList.Text = "V"
     For i = 4 To GrdDataList.Cols - 1
         GrdDataList.col = i
         GrdDataList.CellBackColor = &HFFC0C0
     Next i
End If
End If
GrdDataList.Visible = True
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1)
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
