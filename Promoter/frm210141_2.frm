VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210141_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "繳款輸入-繳款記錄刪除"
   ClientHeight    =   3192
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9384
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3192
   ScaleWidth      =   9384
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox txtSales 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   4
      Top             =   240
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "明細(&T)"
      Height          =   400
      Index           =   2
      Left            =   7065
      TabIndex        =   3
      Top             =   90
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "刪除(&D)"
      Height          =   400
      Index           =   0
      Left            =   5760
      TabIndex        =   1
      Top             =   90
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   8040
      TabIndex        =   0
      Top             =   90
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm210141_2.frx":0000
      Height          =   2385
      Left            =   90
      TabIndex        =   2
      Top             =   570
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   4212
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "Rdate"
         Caption         =   "繳款日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "RTime"
         Caption         =   "繳款時間"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "A4404"
         Caption         =   "票據號碼"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "A4405"
         Caption         =   "票據金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "A4406"
         Caption         =   "北所電匯金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "A4407"
         Caption         =   "分所電匯金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "A4408"
         Caption         =   "現金"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "A4409"
         Caption         =   "抵暫收款"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "A4430"
         Caption         =   "其他"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "A4410"
         Caption         =   "溢收款"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "A4411"
         Caption         =   "手續費"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "A4422"
         Caption         =   "補扣繳"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "A4426"
         Caption         =   "外幣"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "A4431"
         Caption         =   "其他備註"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "A4412"
         Caption         =   "智權人員備註"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   912.189
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   912.189
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   924.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   947.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   708.095
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   792
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   792
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column13 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column14 
            Locked          =   -1  'True
            ColumnWidth     =   1980.284
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4455
      Top             =   180
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   572
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSForms.Label lblSalesName 
      Height          =   285
      Left            =   2070
      TabIndex        =   6
      Top             =   240
      Width           =   1710
      VariousPropertyBits=   27
      Size            =   "3016;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   285
      Width           =   900
   End
End
Attribute VB_Name = "frm210141_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、lblSalesName
'Memo by Lydia 2019/07/01 表單名稱:智權人員繳款資料輸入=>繳款輸入
'Created by Morgan 2013/12/3
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
   Case 0
      If MsgBox("是否確定要刪除本次繳款記錄？" & vbCrLf & vbCrLf & "繳款日期：" & DataGrid1.Columns(0) & vbCrLf & "繳款時間：" & DataGrid1.Columns(1), vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
         If FormDelete = False Then
            MsgBox "刪除失敗，請洽系統管理員 !", vbCritical
         Else
            Adodc1.Recordset.Delete
            Adodc1.Recordset.UPDATE
            MsgBox "繳款記錄已刪除!", vbInformation
            If Adodc1.Recordset.RecordCount = 0 Then
               Unload Me
            End If
         End If
      End If
   Case 2
      doQuery
   Case 1
      Unload Me
   End Select
End Sub

Private Sub DataGrid1_DblClick()
   cmdOK(0).Value = True
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210141_2 = Nothing
End Sub

Private Function FormDelete() As Boolean
   cnnConnection.BeginTrans
On Error GoTo ErrHnd

   strSql = "delete acc441 where AXD01='" & Adodc1.Recordset("A4401") & "' AND AXD02=" & Adodc1.Recordset("A4402") & " AND AXD03=" & Adodc1.Recordset("A4403")
   cnnConnection.Execute strSql, intI
   
   If Not IsNull(Adodc1.Recordset.Fields("A4421")) Then
      strSql = "UPDATE acc230 SET A2308=NULL,A2309=NULL,A2321=NULL where A2301='" & Adodc1.Recordset.Fields("A4421") & "'"
      cnnConnection.Execute strSql, intI
      'Added by Morgan 2014/1/22
      If Not IsNull(Adodc1.Recordset.Fields("A4427")) Then
         strSql = "UPDATE acc230 SET A2308=NULL,A2309=NULL,A2321=NULL where A2301 in ('" & Replace(Adodc1.Recordset.Fields("A4427"), ";", "','") & "')"
         cnnConnection.Execute strSql, intI
      End If
      'end 2014/1/22
   End If
   
   strSql = "delete acc440 where a4401='" & Adodc1.Recordset("A4401") & "' AND A4402=" & Adodc1.Recordset("A4402") & " AND A4403=" & Adodc1.Recordset("A4403")
   cnnConnection.Execute strSql, intI
   
   cnnConnection.CommitTrans
   FormDelete = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

Private Sub doQuery()
   Dim adoquery As ADODB.Recordset
   Dim dblVal(3) As Double
   
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modified by Morgan 2021/11/10 公司別統一改用簡稱 a0k11-->a0820
   'Modified by Lydia 2023/11/13 開立INVOICE，不列印收據;decode(nvl(a0k19,0),0,'◎')=> decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))
   strExc(0) = "select sqldatet(a0k02) 單據日期" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",decode(a0j04,'000',CPM03,CPM04) 案件性質" & _
      ",na03 國別,axd06 服務費,axd07 規費,axd08 扣繳金額,a0k01||decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))||decode(AXC01,null,'','＊') 收據編號" & _
      ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱,a0k03,a0k04,a0820 公司別" & _
      " from ACC441,ACC0J0,acc0k0,acc080,acc431,caseprogress,casepropertymap,nation" & _
      ",trademark,patent,lawcase,servicepractice,hirecase" & _
      " where AXD01='" & Adodc1.Recordset("A4401") & "' and AXD02=" & Adodc1.Recordset("A4402") & _
      " and AXD03=" & Adodc1.Recordset("A4403") & " and A0J01(+)=AXD05 AND A0J13(+)=AXD04" & _
      " and a0k01(+)=a0j13 and a0801(+)=a0k11 and axc02(+)=a0j13 and cp09(+)=a0j01" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and na01(+)=a0j04" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04" & _
      " order by a0k02,a0j13,a0j01"
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
   If adoquery.RecordCount > 0 Then
      With frm210141_3
      'Modify by Amy 2014/06/16 +FormName 改暫存TB
      Set .Adodc1.Recordset = PUB_CreateRecordset(adoquery, , , , .Name)
         With adoquery
         .MoveFirst
         Do While Not .EOF
            dblVal(1) = dblVal(1) + Val("" & .Fields("服務費"))
            dblVal(2) = dblVal(2) + Val("" & .Fields("規費"))
            dblVal(3) = dblVal(3) + Val("" & .Fields("扣繳金額"))
            .MoveNext
         Loop
         End With
         
         .txtTot(2) = Format(dblVal(1), "#,##0")
         .txtTot(3) = Format(dblVal(2), "#,##0")
         .txtTot(4) = Format(dblVal(3), "#,##0")
         .txtTot(5) = Format(dblVal(1) + dblVal(2) - dblVal(3), "#,##0")
         .Show vbModal
      End With
   End If
   Set adoquery = Nothing
End Sub

