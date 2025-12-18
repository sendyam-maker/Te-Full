VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100127_1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "優先權資料查詢"
   ClientHeight    =   5750
   ClientLeft      =   110
   ClientTop       =   990
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8950
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1485
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "1"
      Top             =   1050
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "優先權號："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   810
      Width           =   1245
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   0
      Left            =   2265
      MaxLength       =   6
      TabIndex        =   2
      Top             =   450
      Width           =   1212
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   1
      Left            =   3540
      MaxLength       =   1
      TabIndex        =   3
      Top             =   450
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   2
      Left            =   3990
      MaxLength       =   2
      TabIndex        =   4
      Top             =   450
      Width           =   492
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3345
      Left            =   60
      TabIndex        =   10
      Top             =   2370
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   5891
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   1485
      MaxLength       =   20
      TabIndex        =   6
      Top             =   750
      Width           =   2292
   End
   Begin VB.TextBox txtSystem 
      Height          =   264
      Left            =   1485
      MaxLength       =   3
      TabIndex        =   1
      Top             =   450
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7470
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6630
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   30
      Width           =   800
   End
   Begin MSForms.Label Label12 
      Height          =   300
      Index           =   3
      Left            =   4470
      TabIndex        =   20
      Top             =   2100
      Width           =   1410
      VariousPropertyBits=   27
      Size            =   "2487;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   300
      Index           =   2
      Left            =   930
      TabIndex        =   19
      Top             =   2100
      Width           =   1410
      VariousPropertyBits=   27
      Size            =   "2487;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   300
      Index           =   1
      Left            =   930
      TabIndex        =   18
      Top             =   1770
      Width           =   7860
      VariousPropertyBits=   27
      Size            =   "13864;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   300
      Index           =   0
      Left            =   1140
      TabIndex        =   17
      Top             =   1440
      Width           =   7650
      VariousPropertyBits=   27
      Size            =   "13494;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "專用期間："
      Height          =   180
      Index           =   0
      Left            =   3450
      TabIndex        =   16
      Top             =   2100
      Width           =   975
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "申請日："
      Height          =   180
      Left            =   180
      TabIndex        =   15
      Top             =   2100
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   180
      TabIndex        =   14
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   13
      Top             =   1440
      Width           =   945
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   90
      X2              =   8910
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "查詢條件："
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Top             =   1110
      Width           =   1245
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "(1.本案主張之優先權 2.主張本案之案件)"
      Height          =   180
      Left            =   2160
      TabIndex        =   11
      Top             =   1110
      Width           =   3135
   End
End
Attribute VB_Name = "frm100127_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/07 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、Label2(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Dim intTemp As Boolean
Dim i As Integer, j As Integer, intK As Double, bolSelData As Boolean
Public cmdState As Integer
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件

Private Sub SetDataListWidth()
'Modified by Lydia 2019/11/01
'grdDataList.Cols = 7
Dim intField As Integer
intField = 19
grdDataList.Cols = intField
'end 2019/11/01

grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "本所案號"
grdDataList.ColWidth(0) = 1500
grdDataList.CellAlignment = flexAlignLeftCenter
grdDataList.col = 1: grdDataList.Text = "申請案號或審定號"
grdDataList.ColWidth(1) = 1600
grdDataList.CellAlignment = flexAlignLeftCenter
grdDataList.col = 2: grdDataList.Text = "申請國家"
grdDataList.ColWidth(2) = 1000
grdDataList.CellAlignment = flexAlignLeftCenter
grdDataList.col = 3: grdDataList.Text = "主張優先權號"
grdDataList.ColWidth(3) = 1400
grdDataList.CellAlignment = flexAlignLeftCenter
grdDataList.col = 4: grdDataList.Text = "優先權日"
grdDataList.ColWidth(4) = 900
grdDataList.CellAlignment = flexAlignLeftCenter
grdDataList.col = 5: grdDataList.Text = "優先權國家"
grdDataList.ColWidth(5) = 1000
grdDataList.CellAlignment = flexAlignLeftCenter
grdDataList.col = 6: grdDataList.Text = "本所案號"
grdDataList.ColWidth(6) = 1500
grdDataList.CellAlignment = flexAlignLeftCenter 'flexAlignCenterCenter

'Added by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
For intI = 7 To intField - 1
     grdDataList.col = intI
     grdDataList.ColWidth(intI) = 0
Next intI
'end 2019/11/01
End Sub

Public Sub PubShowNextData()
Select Case cmdState
Case 2 '結束
     fnCloseAllFrm100
Case Else
End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub cmdSearch_Click()
Dim s As Integer
Dim strCode1 As String, strCode2 As String, strCode3 As String, strCode4 As String
Dim strInNo As String, strSQLP As String, strSQLT As String
'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
Dim SeColPA As String, SeColPA_2 As String
Dim SeColTM As String, SeColTM_2 As String
Dim dblRow As Double 'Add By Sindy 2025/9/3

bolSelData = False

If Option1(0).Value = True Then
   If Len(Trim(txtSystem)) = 0 Or Len(Trim(txtCode(0))) = 0 Then
      s = MsgBox("輸入條件不可空白", , "User 輸入錯誤")
      If Len(Trim(txtSystem)) = 0 Then txtSystem.SetFocus
      Exit Sub
   End If
Else
   If Option1(1).Value = True Then
      If Len(Trim(Text1)) = 0 Then
         s = MsgBox("輸入條件不可空白", , "User 輸入錯誤")
         If Len(Trim(Text1)) = 0 Then Text1.SetFocus
         Exit Sub
      End If
   End If
End If
If Len(Trim(Text2)) = 0 Then
   s = MsgBox("輸入條件不可空白", , "User 輸入錯誤")
   If Len(Trim(Text2)) = 0 Then Text2.SetFocus
   Exit Sub
End If

'Added by Lydia 2019/11/01 清空
Label12(0) = ""  '案件名稱
Label12(1) = ""  '申請人
Label12(2) = ""  '申請日
Label12(3) = ""  '專用期間
   
'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
SeColTM = " ,t1.tm23 as custa1,t1.tm78 as custa2,t1.tm79 as custa3,t1.tm80 as custa4,t1.tm81 as custa5,t1.tm44 as fcnoa " & _
                  " ,'' as custb1,'' as custb2,'' as custb3,'' as custb4,'' as custb5,'' as fcnob "
SeColPA = " ,p1.pa26 as custa1,p1.pa27 as custa2,p1.pa28 as custa3,p1.pa29 as custa4,p1.pa30 as custa5,p1.pa75 as fcnoa " & _
                  " ,'' as custb1,'' as custb2,'' as custb3,'' as custb4,'' as custb5,'' as fcnob "
SeColTM_2 = " ,t1.tm23 as custa1,t1.tm78 as custa2,t1.tm79 as custa3,t1.tm80 as custa4,t1.tm81 as custa5,t1.tm44 as fcnoa " & _
                  " ,t2.tm23 as custb1,t2.tm78 as custb2,t2.tm79 as custb3,t2.tm80 as custb4,t2.tm81 as custb5,t2.tm44 as fcnob "
SeColPA_2 = " ,p1.pa26 as custa1,p1.pa27 as custa2,p1.pa28 as custa3,p1.pa29 as custa4,p1.pa30 as custa5,p1.pa75 as fcnoa " & _
                " ,p2.pa26 as custb1,p2.pa27 as custb2,p2.pa28 as custb3,p2.pa29 as custb4,p2.pa30 as custb5,p2.pa75 as fcnob "
intCufaCnt = 0
'end 2019/11/01
ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/16 清除查詢印表記錄檔欄位
'本所案號
If Option1(0).Value = True Then
   strCode1 = Trim(txtSystem)
   strCode2 = Trim(txtCode(0))
   strCode3 = Trim(txtCode(1))
   If strCode3 = "" Then strCode3 = "0"
   strCode4 = Trim(txtCode(2))
   If strCode4 = "" Then strCode4 = "00"
   pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & strCode1 & "-" & strCode2 & "-" & strCode3 & "-" & strCode4 'Add By Sindy 2010/11/16
   '1.本案主張之優先權
   If Trim(Text2) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label4 & "1.本案主張之優先權" 'Add By Sindy 2010/11/16
      'Modified by Lydia 2019/11/01 +增加欄位 SeColPA
      strSql = " select P1.pa01||'-'||P1.pa02||'-'||P1.pa03||'-'||P1.pa04 AS 本所案號,P1.pa11 AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                     SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,'　　　　' AS 本所案號_1" & SeColPA & _
                     " from pridate,patent P1,nation n1,nation n2" & _
                     " where pd01='" & strCode1 & "' and pd02='" & strCode2 & "' and pd03='" & strCode3 & "' and pd04='" & strCode4 & "'" & _
                     " and pd01=P1.pa01 and pd02=P1.pa02 and pd03=P1.pa03 and pd04=P1.pa04" & _
                     " and P1.pa09=n1.na01(+)" & _
                     " and pd07=n2.na01(+)"
      'Modified by Lydia 2019/11/01 +增加欄位 SeColTM
      strSql = strSql & " Union All" & _
                     " select T1.tm01||'-'||T1.tm02||'-'||T1.tm03||'-'||T1.tm04 AS 本所案號,decode(T1.tm15,null,T1.tm12,T1.tm15) AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                     SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,'　　　　' AS 本所案號_1" & SeColTM & _
                     " from pridate,trademark T1,nation n1,nation n2" & _
                     " where pd01='" & strCode1 & "' and pd02='" & strCode2 & "' and pd03='" & strCode3 & "' and pd04='" & strCode4 & "'" & _
                     " and pd01=T1.tm01 and pd02=T1.tm02 and pd03=T1.tm03 and pd04=T1.tm04" & _
                     " and T1.tm10=n1.na01(+)" & _
                     " and pd07=n2.na01(+)" & _
                     " order by 主張優先權號 asc"
   '2.主張本案之案件
   ElseIf Trim(Text2) = "2" Then
      pub_QL05 = pub_QL05 & ";" & Label4 & "2.主張本案之案件" 'Add By Sindy 2010/11/16
      'Modified by Lydia 2019/11/01 +增加欄位 SeColPA_2
      strSql = " select pd01||'-'||pd02||'-'||pd03||'-'||pd04 AS 本所案號,P2.pa11 AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                     SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,P1.pa01||'-'||P1.pa02||'-'||P1.pa03||'-'||P1.pa04 AS 本所案號_1" & SeColPA_2 & _
                     " from pridate,(select * from patent" & _
                     " where pa01='" & strCode1 & "' and pa02='" & strCode2 & "' and pa03='" & strCode3 & "' and pa04='" & strCode4 & "') P1,nation n1,nation n2,patent P2" & _
                     " Where pd06 = P1.pa11" & _
                     " and pd01=P2.pa01 and pd02=P2.pa02 and pd03=P2.pa03 and pd04=P2.pa04" & _
                     " and P2.pa09=n1.na01(+)" & _
                     " and pd07=n2.na01(+)"
      'Modified by Morgan 2016/6/22
      '1.優先權號加抓本所案號,2.另申請號前面加'0'要判斷有申請號的案子否則未提申的案號會固定抓到優先權號輸'0'的資料
      'Modified by Lydia 2019/11/01 +增加欄位 SeColPA_2
      strSql = strSql & " Union All" & _
                     " select pd01||'-'||pd02||'-'||pd03||'-'||pd04 AS 本所案號,P2.pa11 AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                     SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,P1.pa01||'-'||P1.pa02||'-'||P1.pa03||'-'||P1.pa04 AS 本所案號_1" & SeColPA_2 & _
                     " from pridate,(select * from patent" & _
                     " where pa01='" & strCode1 & "' and pa02='" & strCode2 & "' and pa03='" & strCode3 & "' and pa04='" & strCode4 & "') P1,nation n1,nation n2,patent P2" & _
                     " where pd06='0'||P1.pa11 and P1.pa11 is not null" & _
                     " and pd01=P2.pa01 and pd02=P2.pa02 and pd03=P2.pa03 and pd04=P2.pa04" & _
                     " and P2.pa09=n1.na01(+)" & _
                     " and pd07=n2.na01(+)"
      'Modified by Lydia 2019/11/01 +增加欄位 SeColPA_2
      strSql = strSql & " Union All" & _
                     " select pd01||'-'||pd02||'-'||pd03||'-'||pd04 AS 本所案號,P2.pa11 AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                     SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,P1.pa01||'-'||P1.pa02||'-'||P1.pa03||'-'||P1.pa04 AS 本所案號_1" & SeColPA_2 & _
                     " from pridate,(select * from patent" & _
                     " where pa01='" & strCode1 & "' and pa02='" & strCode2 & "' and pa03='" & strCode3 & "' and pa04='" & strCode4 & "') P1,nation n1,nation n2,patent P2" & _
                     " where pd06=P1.pa01||P1.pa02||P1.pa03||P1.pa04" & _
                     " and pd01=P2.pa01 and pd02=P2.pa02 and pd03=P2.pa03 and pd04=P2.pa04" & _
                     " and P2.pa09=n1.na01(+)" & _
                     " and pd07=n2.na01(+)"
      'end 2016/6/22
      'Modified by Lydia 2019/11/01 +增加欄位 SeColPA_2
      strSql = strSql & " Union All" & _
                     " select pd01||'-'||pd02||'-'||pd03||'-'||pd04 AS 本所案號,P2.pa11 AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                     SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,P1.pa01||'-'||P1.pa02||'-'||P1.pa03||'-'||P1.pa04 AS 本所案號_1" & SeColPA_2 & _
                     " from pridate,(select * from patent" & _
                     " where pa01='" & strCode1 & "' and pa02='" & strCode2 & "' and pa03='" & strCode3 & "' and pa04='" & strCode4 & "') P1,nation n1,nation n2,patent P2" & _
                     " Where pd06 = P1.pa22" & _
                     " and pd01=P2.pa01 and pd02=P2.pa02 and pd03=P2.pa03 and pd04=P2.pa04" & _
                     " and P2.pa09=n1.na01(+)" & _
                     " and pd07=n2.na01(+)"
      'Modified by Lydia 2019/11/01 +增加欄位 SeColTM_2
      strSql = strSql & " Union All" & _
                     " select pd01||'-'||pd02||'-'||pd03||'-'||pd04 AS 本所案號,decode(T2.tm15,null,T2.tm12,T2.tm15) AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                     SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,T1.tm01||'-'||T1.tm02||'-'||T1.tm03||'-'||T1.tm04 AS 本所案號_1" & SeColTM_2 & _
                     " from pridate,(select * from trademark" & _
                     " where tm01='" & strCode1 & "' and tm02='" & strCode2 & "' and tm03='" & strCode3 & "' and tm04='" & strCode4 & "') T1,nation n1,nation n2,trademark T2" & _
                     " Where pd06 = t1.tm12" & _
                     " and pd01=T2.tm01 and pd02=T2.tm02 and pd03=T2.tm03 and pd04=T2.tm04" & _
                     " and T2.tm10=n1.na01(+)" & _
                     " and pd07=n2.na01(+)"
      'Modified by Morgan 2016/6/22
      '1.優先權號加抓本所案號,2.另申請號前面加'0'要判斷有申請號的案子否則未提申的案號會固定抓到優先權號輸'0'的資料
      'Modified by Lydia 2019/11/01 +增加欄位 SeColTM_2
      strSql = strSql & " Union All" & _
                     " select pd01||'-'||pd02||'-'||pd03||'-'||pd04 AS 本所案號,decode(T2.tm15,null,T2.tm12,T2.tm15) AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                     SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,T1.tm01||'-'||T1.tm02||'-'||T1.tm03||'-'||T1.tm04 AS 本所案號_1" & SeColTM_2 & _
                     " from pridate,(select * from trademark" & _
                     " where tm01='" & strCode1 & "' and tm02='" & strCode2 & "' and tm03='" & strCode3 & "' and tm04='" & strCode4 & "') T1,nation n1,nation n2,trademark T2" & _
                     " where pd06='0'||T1.tm12 and T1.tm12 is not null" & _
                     " and pd01=T2.tm01 and pd02=T2.tm02 and pd03=T2.tm03 and pd04=T2.tm04" & _
                     " and T2.tm10=n1.na01(+)" & _
                     " and pd07=n2.na01(+)"
      'Modified by Lydia 2019/11/01 +增加欄位 SeColTM_2
      strSql = strSql & " Union All" & _
                     " select pd01||'-'||pd02||'-'||pd03||'-'||pd04 AS 本所案號,decode(T2.tm15,null,T2.tm12,T2.tm15) AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                     SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,T1.tm01||'-'||T1.tm02||'-'||T1.tm03||'-'||T1.tm04 AS 本所案號_1" & SeColTM_2 & _
                     " from pridate,(select * from trademark" & _
                     " where tm01='" & strCode1 & "' and tm02='" & strCode2 & "' and tm03='" & strCode3 & "' and tm04='" & strCode4 & "') T1,nation n1,nation n2,trademark T2" & _
                     " where pd06=T1.tm01||T1.tm02||T1.tm03||T1.tm04" & _
                     " and pd01=T2.tm01 and pd02=T2.tm02 and pd03=T2.tm03 and pd04=T2.tm04" & _
                     " and T2.tm10=n1.na01(+)" & _
                     " and pd07=n2.na01(+)"
      'end 2016/6/22
      'Modified by Lydia 2019/11/01 +增加欄位 SeColTM_2
      strSql = strSql & " Union All" & _
                     " select pd01||'-'||pd02||'-'||pd03||'-'||pd04 AS 本所案號,decode(T2.tm15,null,T2.tm12,T2.tm15) AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                     SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,T1.tm01||'-'||T1.tm02||'-'||T1.tm03||'-'||T1.tm04 AS 本所案號_1" & SeColTM_2 & _
                     " from pridate,(select * from trademark" & _
                     " where tm01='" & strCode1 & "' and tm02='" & strCode2 & "' and tm03='" & strCode3 & "' and tm04='" & strCode4 & "') T1,nation n1,nation n2,trademark T2" & _
                     " Where pd06 = t1.TM15" & _
                     " and pd01=T2.tm01 and pd02=T2.tm02 and pd03=T2.tm03 and pd04=T2.tm04" & _
                     " and T2.tm10=n1.na01(+)" & _
                     " and pd07=n2.na01(+)"
   End If
   
'優先權號
ElseIf Option1(1).Value = True Then
   strInNo = "'" & Trim(Text1) & "'"
   If Left(Trim(Text1), 1) = "0" And Len(Trim(Text1)) = 9 Then
      '當第1碼為0時,增加一組去0號數
      strInNo = strInNo & ",'" & Mid(Trim(Text1), 2, Len(Trim(Text1))) & "'"
   ElseIf Left(Trim(Text1), 1) <> "0" And Len(Trim(Text1)) = 8 Then
      '當第1碼非0時,增加一組+0號數
      strInNo = strInNo & ",'0" & Trim(Text1) & "'"
   End If
   pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Trim(Text1) 'Add By Sindy 2010/11/16
   '1.本案主張之優先權
   If Trim(Text2) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label4 & "1.本案主張之優先權" 'Add By Sindy 2010/11/16
      'Modified by Lydia 2019/11/01 +增加欄位 SeColPA
      strSql = " select P1.pa01||'-'||P1.pa02||'-'||P1.pa03||'-'||P1.pa04 AS 本所案號,P1.pa11 AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                    SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,'　　　　' AS 本所案號_1" & SeColPA & _
                     " from pridate,(select * from patent where pa11 in(" & strInNo & ") Union select * from patent where pa22 in(" & strInNo & ")) P1,nation n1,nation n2" & _
                     " Where pd01 = P1.PA01 And pd02 = P1.pa02 And pd03 = P1.pa03 And pd04 = P1.pa04" & _
                     " and P1.pa09=n1.na01(+)" & _
                     " and pd07=n2.na01(+)"
      'Modified by Lydia 2019/11/01 +增加欄位 SeColTM
      strSql = strSql & " Union All" & _
                     " select T1.tm01||'-'||T1.tm02||'-'||T1.tm03||'-'||T1.tm04 AS 本所案號,decode(T1.tm15,null,T1.tm12,T1.tm15) AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                     SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,'　　　　' AS 本所案號_1" & SeColTM & _
                     " from pridate,(select * from trademark where tm12 in(" & strInNo & ") Union select * from trademark where tm15 in(" & strInNo & ")) T1,nation n1,nation n2" & _
                     " Where pd01 = t1.tm01 And pd02 = t1.tm02 And pd03 = t1.tm03 And pd04 = t1.tm04" & _
                     " and T1.tm10=n1.na01(+)" & _
                     " and pd07=n2.na01(+)" & _
                     " order by 主張優先權號 asc"
   '2.主張本案之案件
   ElseIf Trim(Text2) = "2" Then
      pub_QL05 = pub_QL05 & ";" & Label4 & "2.主張本案之案件" 'Add By Sindy 2010/11/16
      'Modified by Lydia 2019/11/01 +增加欄位 SeColPA
      strSql = " select P1.pa01||'-'||P1.pa02||'-'||P1.pa03||'-'||P1.pa04 AS 本所案號,P1.pa11 AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                    SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,'　　　　' AS 本所案號_1" & SeColPA & _
                     " from pridate,patent P1,nation n1,nation n2" & _
                     " where pd06 in(" & strInNo & ")" & _
                     " and pd01=P1.pa01(+) and pd02=P1.pa02(+) and pd03=P1.pa03(+) and pd04=P1.pa04(+)" & _
                     " and P1.pa01 is not null" & _
                     " and P1.pa09=n1.na01(+)" & _
                     " and pd07=n2.na01(+)"
      'Modified by Lydia 2019/11/01 +增加欄位 SeColTM
      strSql = strSql & " Union All" & _
                     " select T1.tm01||'-'||T1.tm02||'-'||T1.tm03||'-'||T1.tm04 AS 本所案號,decode(T1.tm15,null,T1.tm12,T1.tm15) AS 申請案號或審定號,n1.na03 AS 申請國家,pd06 AS 主張優先權號," & _
                     SQLDate("pd05", False) & " AS 優先權日,n2.na03 AS 優先權國家,'　　　　' AS 本所案號_1" & SeColTM & _
                     " from pridate,trademark T1,nation n1,nation n2" & _
                     " where pd06 in(" & strInNo & ")" & _
                     " and pd01=T1.tm01(+) and pd02=T1.tm02(+) and pd03=T1.tm03(+) and pd04=T1.tm04(+)" & _
                     " and T1.tm01 is not null" & _
                     " and T1.tm10=n1.na01(+)" & _
                     " and pd07=n2.na01(+)"
   End If
End If

grdDataList.Clear
grdDataList.Rows = 2
SetDataListWidth
Screen.MousePointer = vbHourglass
Me.Enabled = False
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
If adoRecordset.RecordCount <> 0 Then
   dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
   'Added by Lydia 2019/11/01 逐案號判斷
   If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
       adoRecordset.MoveFirst
       Do While adoRecordset.EOF = False
           '利益衝突案件：逐案號判斷
           If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號"), "" & adoRecordset.Fields("custa1") & "," & adoRecordset.Fields("custa2") & "," & adoRecordset.Fields("custa3") & "," & adoRecordset.Fields("custa4") & "," & adoRecordset.Fields("custa5"), "" & adoRecordset.Fields("fcnoa")) = False Then
               intCufaCnt = intCufaCnt + 1
               adoRecordset.Delete
           Else
               '優先權抓到的本所案號
               If "" & adoRecordset.Fields("custb1") <> "" Then
                   If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields("本所案號_1"), "" & adoRecordset.Fields("custb1") & "," & adoRecordset.Fields("custb2") & "," & adoRecordset.Fields("custb3") & "," & adoRecordset.Fields("custb4") & "," & adoRecordset.Fields("custb5"), "" & adoRecordset.Fields("fcnob")) = False Then
                       intCufaCnt = intCufaCnt + 1
                       adoRecordset.Delete
                   End If
               End If
           End If
           adoRecordset.MoveNext
       Loop
       '利益衝突案件：限閱案件
       If intCufaCnt > 0 Then
          pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
          MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
       End If
       InsertQueryLog (dblRow) 'Add By Sindy 2010/11/16
       If adoRecordset.RecordCount = 0 Then
             GoTo JumpToNoData
       End If
   Else
       InsertQueryLog (dblRow) 'Add By Sindy 2010/11/16
   End If
   'end 2019/11/01
   
   grdDataList.Rows = adoRecordset.RecordCount + 1
Else
   InsertQueryLog (0) 'Add By Sindy 2010/11/16
JumpToNoData:   'Added by Lydia 2019/11/01
   CheckOC
   ShowNoData
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   Exit Sub
End If

'把資料放進 GRID
If adoRecordset.RecordCount <> 0 Then
   Set grdDataList.Recordset = adoRecordset
   intK = adoRecordset.RecordCount
End If
CheckOC

'查詢案件資料
If Option1(0).Value = True Then
   '上面程式段已取得畫面上所輸入的本所案號
ElseIf Option1(1).Value = True Then
   strSql = "select pa01,pa02,pa03,pa04 from patent" & _
                  " where pa11 in(" & strInNo & ") or pa22 in(" & strInNo & ")" & _
                  " Union All" & _
                  " select tm01,tm02,tm03,tm04 from trademark" & _
                  " where tm12 in(" & strInNo & ") or tm15 in(" & strInNo & ")"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCode1 = "" & RsTemp.Fields(0)
      strCode2 = "" & RsTemp.Fields(1)
      strCode3 = "" & RsTemp.Fields(2)
      strCode4 = "" & RsTemp.Fields(3)
   End If
End If
strSql = " select NVL(NVL(PA05,PA06),PA07),NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90))," & SQLDate("pa10", True) & "," & SQLDate("pa24", False) & "||'-'||" & SQLDate("pa25", False) & _
               " from patent,CUSTOMER" & _
               " where pa01='" & strCode1 & "' and pa02='" & strCode2 & "' and pa03='" & strCode3 & "' and pa04='" & strCode4 & "'" & _
               " AND SUBSTR(PA26,1,8)=CU01(+)" & _
               " AND decode(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+)" & _
               " Union All" & _
               " select NVL(NVL(TM05,TM06),TM07),NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90))," & SQLDate("tm11", True) & "," & SQLDate("tm21", False) & "||'-'||" & SQLDate("tm22", False) & _
               " from trademark,CUSTOMER" & _
               " where tm01='" & strCode1 & "' and tm02='" & strCode2 & "' and tm03='" & strCode3 & "' and tm04='" & strCode4 & "'" & _
               " AND SUBSTR(TM23,1,8)=CU01(+)" & _
               " AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+)"
intI = 1
Set RsTemp = ClsLawReadRstMsg(intI, strSql)
If intI = 1 Then
   Label12(0) = "" & RsTemp.Fields(0) '案件名稱
   Label12(1) = "" & RsTemp.Fields(1) '申請人
   Label12(2) = "" & RsTemp.Fields(2) '申請日
   Label12(3) = "" & RsTemp.Fields(3) '專用期間
End If

intCufaCnt = 0 'Added by Lydia 2019/11/01

'逐筆檢查是否有本所案號, 若無, 則依主張優先權號取得本所案號
For i = 1 To grdDataList.Rows - 1
   If Trim(grdDataList.TextMatrix(i, 6)) = "" Then
      strInNo = "'" & Trim(grdDataList.TextMatrix(i, 3)) & "'"
      If Left(Trim(grdDataList.TextMatrix(i, 3)), 1) = "0" And Len(Trim(grdDataList.TextMatrix(i, 3))) = 9 Then
         '當第1碼為0時,增加一組去0號數
         strInNo = strInNo & ",'" & Mid(Trim(grdDataList.TextMatrix(i, 3)), 2, Len(Trim(grdDataList.TextMatrix(i, 3)))) & "'"
      ElseIf Left(Trim(grdDataList.TextMatrix(i, 3)), 1) <> "0" And Len(Trim(grdDataList.TextMatrix(i, 3))) = 8 Then
         '當第1碼非0時,增加一組+0號數
         strInNo = strInNo & ",'0" & Trim(grdDataList.TextMatrix(i, 3)) & "'"
      End If
      
      'Modified by Lydia 2019/11/01 增加欄位 SeColPA, SeColTM
      'strSql = "select pa01||'-'||pa02||'-'||pa03||'-'||pa04 from patent" & _
                     " where pa11 in(" & strInNo & ") or pa22 in(" & strInNo & ")" & _
                     " Union All" & _
                     " select tm01||'-'||tm02||'-'||tm03||'-'||tm04 from trademark" & _
                     " where tm12 in(" & strInNo & ") or tm15 in(" & strInNo & ")"
      strSql = "select P1.pa01||'-'||P1.pa02||'-'||P1.pa03||'-'||P1.pa04 as 本所案號 " & SeColPA & " from patent P1 " & _
                     " where P1.pa11 in(" & strInNo & ") or P1.pa22 in(" & strInNo & ")" & _
                     " Union All" & _
                     " select T1.tm01||'-'||T1.tm02||'-'||T1.tm03||'-'||T1.tm04 as 本所案號 " & SeColTM & " from trademark T1 " & _
                     " where T1.tm12 in(" & strInNo & ") or T1.tm15 in(" & strInNo & ")"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'Added by Lydia 2019/11/01 逐案號判斷
         If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & RsTemp.Fields("本所案號"), "" & RsTemp.Fields("custa1") & "," & RsTemp.Fields("custa2") & "," & RsTemp.Fields("custa3") & "," & RsTemp.Fields("custa4") & "," & RsTemp.Fields("custa5"), "" & RsTemp.Fields("fcnoa")) = False Then
                 '不顯示優先權抓到的本所案號
                 grdDataList.RowHeight(i) = 0
                 intCufaCnt = intCufaCnt + 1
            Else
                 grdDataList.TextMatrix(i, 6) = "" & RsTemp.Fields(0)
            End If
         Else
         'end 2019/11/01
            grdDataList.TextMatrix(i, 6) = "" & RsTemp.Fields(0)
         End If 'end 2019/11/01
      End If
   End If
Next i

'Added by Lydia 2019/11/01 利益衝突案件：限閱案件 ;  優先權抓到的本所案號
If intCufaCnt > 0 Then
    MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
End If
'end 2019/11/01

If Option1(0).Value = True Then
   Call txtCode_GotFocus(0)
ElseIf Option1(1).Value = True Then
   Call Text1_GotFocus
End If

Me.Enabled = True
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
If Option1(0).Value = True Then
   txtSystem.SetFocus
   txtSystem.SelStart = 0
   txtSystem.SelLength = Len(txtSystem)
End If
End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
grdDataList.Clear
SetDataListWidth
bolToEndByNick = False
intTemp = False
bolSelData = False
cmdState = -1
m_AllSys = GetAllSysKind(, "ALL") 'Added by Lydia 2019/11/01 利益衝突案件：系統別
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100127_1 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
         If Option1(0).Value = True Then
            Option1(0).Value = True
            Option1(1).Value = False
            If intTemp = False Then
               txtSystem.SetFocus
            End If
            intTemp = False
         End If
      Case 1
         If Option1(1).Value = True Then
            Option1(1).Value = True
            Option1(0).Value = False
            Text1.SetFocus
         End If
      Case Else
   End Select
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1)
   '加輸入法控制
   CloseIme
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Option1(1).Value = True
End Sub

Private Sub Text2_GotFocus()
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2)
   '加輸入法控制
   CloseIme
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   intTemp = True
   txtCode(Index).SelStart = 0
   txtCode(Index).SelLength = Len(txtCode(Index))
   '加輸入法控制
   CloseIme
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Option1(0).Value = True
End Sub

Private Sub txtSystem_GotFocus()
   If Option1(0).Value = True Then
      txtSystem.SetFocus
      txtSystem.SelStart = 0
      txtSystem.SelLength = Len(txtSystem)
   End If
   If Option1(1).Value = True Then
      Text1.SetFocus
      Text1.SelStart = 0
      Text1.SelLength = Len(Text1)
   End If
   '加輸入法控制
   CloseIme
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
