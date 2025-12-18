VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm081020_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "E-Mail資料已存在，是否要繼續？"
   ClientHeight    =   5964
   ClientLeft      =   156
   ClientTop       =   996
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5964
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdok 
      Caption         =   "否(&N)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3870
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   45
      Width           =   870
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "是(&Y)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   2940
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   45
      Width           =   870
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5250
      Left            =   30
      TabIndex        =   2
      Top             =   660
      Width           =   4725
      _ExtentX        =   8340
      _ExtentY        =   9250
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   555
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   2745
   End
End
Attribute VB_Name = "frm081020_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer

Private Sub SetDataListWidth()
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "編號"
grdDataList.ColWidth(0) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "名稱 "
grdDataList.ColWidth(1) = 3000
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 1
     fnCloseAllFrm100
Case Else
End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
'CmdState = Index
'PubShowNextData
'Exit Sub
'92.04.16 nick 以下無效
Select Case Index
Case 0
   'tmpBol = fnCancelNowFormAndShowParentForm(Me)
   frm081020.txtEmailSameCnt = "Y"
   Me.Hide
   frm081020.Show
   frm081020.OnAction vbKeyF9   'SONIA
Case 1
   'fnCloseAllFrm100
   frm081020.txtEmailSameCnt = "N"
   Me.Hide
   frm081020.Show
'     bolToEndByNick = True
'     Unload Me
'     Exit Sub
Case Else
End Select
End Sub

Private Sub Form_Load()
   'bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth
   '92.04.16 nick
   'CmdState = -1
End Sub

Sub StrMenu()
Dim m_i As Integer
Dim rsTmp As New ADODB.Recordset
Dim strSR04 As String

Me.Enabled = False

strSql = "SELECT * FROM ("
'客戶檔
'Modified by Lydia 2024/09/18 +財務副本信箱(CU200)
strSql = strSql & "SELECT CU01||CU02 as 編號,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as 名稱 FROM Customer " & _
                  "Where (instr(NLS_Upper(CU20),'" & UCase(ChgSQL(frm081020.textECD13)) & "')>0 " & _
                  "or instr(NLS_Upper(CU115),'" & UCase(ChgSQL(frm081020.textECD13)) & "')>0 " & _
                  "or instr(NLS_Upper(CU116),'" & UCase(ChgSQL(frm081020.textECD13)) & "')>0 " & _
                  "or instr(NLS_Upper(CU117),'" & UCase(ChgSQL(frm081020.textECD13)) & "')>0 " & _
                  "or instr(NLS_Upper(CU118),'" & UCase(ChgSQL(frm081020.textECD13)) & "')>0 " & _
                  "or instr(NLS_Upper(CU200),'" & UCase(ChgSQL(frm081020.textECD13)) & "')>0) "
'國外代理人檔
'Modified by Lydia 2018/07/20 +FA105 財務信箱(CF)
'Modified by Lydia 2024/09/18 +財務副本信箱(FA134)
strSql = strSql & " union all SELECT FA01||FA02 as 編號,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) as 名稱 FROM Fagent " & _
                  "Where (instr(NLS_Upper(fa16),'" & UCase(ChgSQL(frm081020.textECD13)) & "')> 0 " & _
                  "or instr(NLS_Upper(fa79),'" & UCase(ChgSQL(frm081020.textECD13)) & "')> 0 " & _
                  "Or InStr(NLS_Upper(fa105),'" & UCase(ChgSQL(frm081020.textECD13)) & "') > 0 " & _
                  "or instr(NLS_Upper(fa80),'" & UCase(ChgSQL(frm081020.textECD13)) & "')> 0 " & _
                  "or instr(NLS_Upper(fa81),'" & UCase(ChgSQL(frm081020.textECD13)) & "') > 0 " & _
                  "Or InStr(NLS_Upper(fa82),'" & UCase(ChgSQL(frm081020.textECD13)) & "') > 0 " & _
                  "Or InStr(NLS_Upper(FA134),'" & UCase(ChgSQL(frm081020.textECD13)) & "') > 0) "
'潛在客戶檔
strSql = strSql & " union all SELECT PCU01||PCU02 as 編號,NVL(PCU08,NVL(RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06),PCU07)) as 名稱 FROM potcustomer " & _
                  "Where (instr(NLS_Upper(pcu18),'" & UCase(ChgSQL(frm081020.textECD13)) & "')> 0) "
'國內潛在客戶檔
strSql = strSql & " union all SELECT POC01||POC02 as 編號,NVL(POC03,NVL(RTRIM(POC23||' '||POC24||' '||POC25||' '||POC26),POC27)) as 名稱 FROM potcustomer1 " & _
                  "Where (instr(NLS_Upper(poc09),'" & UCase(ChgSQL(frm081020.textECD13)) & "')> 0) "
'外法開拓客戶檔
strSql = strSql & " union all SELECT ECD02||'-'||ECD01 as 編號,NVL(ecd03,'')||NVL(ecd04,'') as 名稱 FROM expandcusdetail " & _
                  "Where (instr(NLS_Upper(ecd13),'" & UCase(ChgSQL(frm081020.textECD13)) & "')> 0) "
strSql = strSql & ") X order by 編號"

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
Else
'   ShowNoData
'   Me.Enabled = True
'   Screen.MousePointer = vbDefault
'   tmpBol = fnCancelNowFormAndShowParentForm(Me)
'   Exit Sub
   Me.Hide
End If
Set grdDataList.Recordset = adoRecordset
frm081020.txtEmailSameCnt = adoRecordset.RecordCount
Label1.Caption = "E-Mail：" & frm081020.textECD13
CheckOC
Me.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm081020_1 = Nothing
End Sub
