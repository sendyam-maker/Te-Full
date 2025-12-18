VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210128_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "與　為關係企業？"
   ClientHeight    =   5745
   ClientLeft      =   150
   ClientTop       =   990
   ClientWidth     =   8460
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8460
   Begin VB.CommandButton cmdok 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7500
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   90
      Width           =   870
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&C)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   6570
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   90
      Width           =   870
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4800
      Left            =   30
      TabIndex        =   1
      Top             =   930
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   8467
      _Version        =   393216
      Cols            =   9
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
      _Band(0).Cols   =   9
   End
   Begin VB.Label Label2 
      Caption         =   "備註：X.申請人"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   690
      Width           =   6075
   End
   Begin MSForms.Label Label1 
      Height          =   495
      Left            =   390
      TabIndex        =   2
      Top             =   90
      Width           =   5895
      VariousPropertyBits=   27
      Caption         =   "Label1"
      Size            =   "10398;873"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm210128_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/14 Form2.0已修改 label1/grdDataList
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim G_strIDNO As String
Dim m_Type As String


Private Sub SetDataListWidth()
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "編號"
grdDataList.ColWidth(0) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "名稱"
grdDataList.ColWidth(1) = 2000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "國籍"
grdDataList.ColWidth(2) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "智權人員"
grdDataList.ColWidth(3) = 1000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "地址"
grdDataList.ColWidth(4) = 3000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "電話"
grdDataList.ColWidth(5) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "傳真"
grdDataList.ColWidth(6) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "狀態"
grdDataList.ColWidth(7) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 8: grdDataList.Text = "備註"
grdDataList.ColWidth(8) = 2000
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub cmdOK_Click(Index As Integer)
'Add By Sindy 2009/06/23
If m_Type = "0" Then '國外
   Select Case Index
      Case 0
         frm140402.txtSameCnt = "Y"
         frm140402.txtPCU(47) = G_strIDNO
      Case 1
         frm140402.txtSameCnt = "E"
      Case Else
   End Select
'   Call frm140402.txtPCU47N_Validate(False)
'   Me.Hide
'   frm140402.Show
'2009/06/23 End
Else
   Select Case Index
      Case 0
         frm210128.txtSameCnt = "Y"
         frm210128.txtPOC(16) = G_strIDNO
      Case 1
         frm210128.txtSameCnt = "E"
      Case Else
   End Select
'   Call frm210128.txtPOC16N_Validate(False)
'   Me.Hide
'   frm210128.Show
End If
Unload frm210128_2
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
'   SetDataListWidth
End Sub

'Modify By Sindy 2009/06/23 增加strType : 0.國外 1.國內 ;strName
'Sub StrMenu()
'Modify By Sindy 2014/2/27
Public Function StrMenu(strType As String, strName As String) As Boolean
Dim m_i As Integer
Dim rsTmp As New ADODB.Recordset
Dim strSR04 As String
Dim StrSQLa As String

StrMenu = True

Label1.Caption = "客戶名稱：" & strName
G_strIDNO = ""
'Add By Sindy 2009/06/23
m_Type = strType
If m_Type = "0" Then '國外
   frm140402.txtPCU(47) = ""
'2009/06/23 End
Else
   frm210128.txtPOC(16) = ""
End If
grdDataList.Clear

'若為國內智權人員或國內工程師, 不可查代理人資料
'Modify By Sindy 2011/01/04 取消
'If bolFNation = False Then
'    StrSQLa = " And FA01<'Y' "
'End If

Screen.MousePointer = vbHourglass

strSql = "SELECT * FROM ("
'客戶檔
strSql = strSql & "SELECT CU01||CU02 AS 編號,CU04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU31,null,CU23,CU31) AS 地址,Decode(CU16,null,CU17,CU16) AS 電話,Decode(CU18,null,CU19,CU18) AS 傳真,CU80 AS 狀態,CU79 AS 備註 FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1 From Customer Where CU04='" & ChgSQL(strName) & "') A WHERE CU10=NA01(+) AND CU01=A.A1 AND CU13=ST01(+)"
'Add By Sindy 2009/06/23
If m_Type = "0" Then
   '國外代理人
   'Modify by Morgan 2011/5/26 +FA70 並清除空白
   strSql = strSql & " union all SELECT FA01||FA02 AS 編號,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) AS 名稱,NA03 AS 國籍,'' AS 智權人員,NVL(FA17,DECODE(FA18,NULL,FA23,rtrim(FA18||' '||FA19||' '||FA20||' '||FA21||' '||FA22||' '||fa70))) AS 地址,Decode(FA12,null,FA13,FA12) AS 電話,Decode(FA14,null,FA15,FA14) AS 傳真,FA69 AS 狀態,FA29 AS 備註 FROM FAGENT,NATION, (Select Distinct FA01 As A1 From Fagent Where NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65))='" & ChgSQL(strName) & "') A WHERE FA10=NA01(+) AND FA01=A.A1 "
End If
'2009/06/23 End
strSql = strSql & ") X order by 編號"

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   If adoRecordset.RecordCount = 1 Then
      'Add By Sindy 2009/06/23
      If m_Type = "0" Then '國外
         frm140402.txtPCU(47) = adoRecordset.Fields(0)
      '2009/06/23 End
      Else
         frm210128.txtPOC(16) = adoRecordset.Fields(0)
      End If
      Me.Hide
   End If
Else
   Screen.MousePointer = vbDefault
   StrMenu = False
   MsgBox "無此客戶資料!", vbExclamation + vbOKOnly
   'Me.Hide
   Exit Function
End If
Set grdDataList.Recordset = adoRecordset
'Add By Sindy 2009/06/23
If m_Type = "0" Then '國外
   frm140402.txtSameCnt = adoRecordset.RecordCount
   Label2.Caption = "備註：X.申請人、Y.代理人"
'2009/06/23 End
Else
   frm210128.txtSameCnt = adoRecordset.RecordCount
   Label2.Caption = "備註：X.申請人"
End If
CheckOC
Screen.MousePointer = vbDefault
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm210128_2 = Nothing
End Sub

'Private Sub grdDataList_Click()
'Dim i As Integer
'grdDataList.Visible = False
'grdDataList.row = grdDataList.MouseRow
'grdDataList.col = 0
'If grdDataList.row <> 0 Then
'    If grdDataList.Text = "V" Then
'         grdDataList.Text = ""
'         For i = 0 To grdDataList.Cols - 1
'            If i <> 1 Then
'                grdDataList.col = i
'                grdDataList.CellBackColor = QBColor(15)
'            End If
'        Next i
'    Else
'         grdDataList.Text = "V"
'         For i = 0 To grdDataList.Cols - 1
'            If i <> 1 Then
'                grdDataList.col = i
'                grdDataList.CellBackColor = &HFFC0C0
'            End If
'         Next i
'    End If
'End If
'grdDataList.Visible = True
'End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow grdDataList, x, y, nCol, nRow
grdDataList.col = nCol
grdDataList.row = nRow
Call grdDataList_SelChange
End Sub

Private Sub grdDataList_SelChange()
Dim tmpMouseRow
Dim i, j

grdDataList.Visible = False
tmpMouseRow = grdDataList.row
grdDataList.Visible = True
If tmpMouseRow <> 0 Then
    grdDataList.row = tmpMouseRow
    grdDataList.col = 0
    If grdDataList.CellBackColor <> &HFFC0C0 Then
         grdDataList.Visible = False
         For j = 1 To grdDataList.Rows - 1
             grdDataList.row = j
             For i = 0 To grdDataList.Cols - 1
                  grdDataList.col = i
                  grdDataList.CellBackColor = QBColor(15)
             Next i
        Next j
        grdDataList.row = tmpMouseRow
        For i = 0 To grdDataList.Cols - 1
             grdDataList.col = i
             grdDataList.CellBackColor = &HFFC0C0
        Next i
        G_strIDNO = grdDataList.TextMatrix(tmpMouseRow, 0)
        grdDataList.Visible = True
    End If
End If
End Sub
