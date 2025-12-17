VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frmacc7150 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "分所出納之智權人員繳款確認"
   ClientHeight    =   4920
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   9384
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9384
   Begin VB.CommandButton Command1 
      Caption         =   "繳款明細(&D)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2640
      TabIndex        =   5
      Top             =   90
      Width           =   1425
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消確認(&C)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   3
      Left            =   4152
      TabIndex        =   4
      Top             =   90
      Width           =   1600
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "重新查詢(&Q)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   2
      Left            =   6840
      TabIndex        =   3
      Top             =   90
      Width           =   1425
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   8310
      TabIndex        =   1
      Top             =   90
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確認(&Y)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   5840
      TabIndex        =   0
      Top             =   90
      Width           =   960
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc7150.frx":0000
      Height          =   4215
      Left            =   90
      TabIndex        =   2
      Top             =   570
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   7430
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   24
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "IDate"
         Caption         =   "簽收通知日"
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
         DataField       =   "Sales"
         Caption         =   "智權人員"
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
         DataField       =   "JComp"
         Caption         =   "J"
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
      BeginProperty Column03 
         DataField       =   "a0k03"
         Caption         =   "客戶代號"
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
      BeginProperty Column04 
         DataField       =   "a0k04"
         Caption         =   "收據抬頭"
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
      BeginProperty Column05 
         DataField       =   "RDate"
         Caption         =   "繳款日期時間"
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
      BeginProperty Column06 
         DataField       =   "Amount"
         Caption         =   "總金額"
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
         DataField       =   "Type"
         Caption         =   "類別"
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
      BeginProperty Column08 
         DataField       =   "CDate"
         Caption         =   "出納確認日期時間"
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
      BeginProperty Column09 
         DataField       =   "A4429"
         Caption         =   "留分所"
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
         DataField       =   "Memo"
         Caption         =   "出納備註"
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
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   912.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   192.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   912.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1548.284
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1607.811
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   912.189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   671.811
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
            ColumnWidth     =   1788.095
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column10 
            Locked          =   -1  'True
            ColumnWidth     =   2352.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   180
      Top             =   90
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   550
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
End
Attribute VB_Name = "Frmacc7150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/28 Form2.0已修改(DataGrid1改Fonts)
'Created by Morgan 2013/12/12
Option Explicit

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If Not IsNull(Adodc1.Recordset.Fields("a4413")) Then
      cmdOK(3).Enabled = True
   Else
      cmdOK(3).Enabled = False
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
   Case 0
      ReadData
   Case 1
      Unload Me
   Case 2
      QueryData
   'Added by Morgan 2014/2/27
   Case 3
      If MsgBox("是否確定要取消確認？" & vbCrLf & vbCrLf & "智權人員：" & Adodc1.Recordset.Fields("Sales") & vbCrLf & "繳款時間：" & Adodc1.Recordset.Fields("RDate") & vbCrLf & "總金額：　" & Adodc1.Recordset.Fields("Amount"), vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
         CancelRec
      End If
   End Select
End Sub

Private Sub CancelRec()
On Error GoTo ErrHnd
   'Add by Lydia 2015/01/14 +票號A4428,留分所A4429
   strSql = "update acc440 set a4413=null,a4414=null,a4415=null,a4423=null,a4428=null,a4429=null " & _
      " where a4401='" & Adodc1.Recordset.Fields("a4401") & "' and a4402=" & Adodc1.Recordset.Fields("a4402") & " And A4403 = " & Adodc1.Recordset.Fields("a4403") & " and a4416 is null"
   cnnConnection.Execute strSql, intI
   
   If intI = 1 Then
      Adodc1.Recordset.Fields("CDate") = Null
      Adodc1.Recordset.Fields("Memo") = Null
      Adodc1.Recordset.Fields("a4413") = Null
      Adodc1.Recordset.UpdateBatch
      MsgBox "已取消!!", vbInformation
   Else
      CheckStatus Adodc1.Recordset.Fields("a4401"), Adodc1.Recordset.Fields("a4402"), Adodc1.Recordset.Fields("a4403"), , "取消確認失敗!!"
   End If
   Exit Sub
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Sub


Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
   QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   MenuEnabled
   Set Frmacc7150 = Nothing
End Sub

Private Sub QueryData()
   Dim stCon As String
   
   'Modified by Morgan 2014/6/2 開放北所財務人員可確認分所繳款資料
   If pub_strUserOffice = "1" Then
      stCon = " and s1.st06<>'" & pub_strUserOffice & "'"
   Else
      stCon = " and s1.st06='" & pub_strUserOffice & "'"
   End If
   'Add by Lydia 2015/01/14 +A4429
   'Modified by Morgan 2015/7/15 +外幣,其他
   'Modified by Morgan 2017/12/4 +確認中的也帶出
   'Modified by Morgan 2017/12/18 電匯+所別
   'Mofified by Morgan 2022/2/16 DataGrid1的第1欄改為簽收通知日(原為業務區),語法本來就有抓不用改
   strExc(0) = "select a0902 Zone,s1.st02 Sales,JComp,sqldatet(a4402)||' '||sqltime(a4403) RDate" & _
      ",nvl(a4405,0)+nvl(a4406,0)+nvl(a4407,0)+nvl(a4408,0)+nvl(a4409,0)-nvl(a4410,0)+nvl(a4411,0)+nvl(a4422,0)+nvl(a4426,0)+nvl(a4430,0) Amount" & _
      ",substr(decode(sign(a4405),1,'支票')||decode(sign(nvl(a4406,0)+nvl(a4407,0)),1,'電匯')||decode(sign(a4408),1,'現金')||'其他',1,2)||decode(sign(a4405),0,decode(sign(nvl(a4406,0)),1,'-北',decode(sign(nvl(a4407,0)),1,'-分'))) Type" & _
      ",sqldatet(to_char(A2324,'yyyymmdd')) IDate" & _
      ",sqldatet(a4413)||' '||sqltime(a4423) Cdate,a4415 Memo,A4429,A4401,A4402,A4403,s2.ST02 CUser,s1.st15" & _
      ",a0k03,a0k04,a4413 from acc440,staff s1,acc090,staff s2" & _
      ",(select axd01,axd02,axd03,min(axd04) axd04,max(decode(a0k11,'J',a0k11,'L',a0k11)) JComp from acc440,acc441,acc0k0 where axd01(+)=a4401 and axd02(+)=a4402 and axd03(+)=a4403" & _
      " and NVL(a4416,'X')='X' and a0k01(+)=axd04 group by axd01,axd02,axd03) x,acc0k0,acc230" & _
      " where NVL(a4416,'X')='X' and s1.st01(+)=a4401 and a0901(+)=s1.st15 and s2.st01(+)=a4414" & stCon & _
      " and axd01(+)=a4401 and axd02(+)=a4402 and axd03(+)=a4403 and a0k01(+)=axd04 and a2301(+)=a4421" & _
      " order by s1.st15,A4401,A4402,A4403"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/07/01 +FormName 改暫存TB
   Set Adodc1.Recordset = PUB_CreateRecordset(RsTemp, , , , Me.Name)
End Sub

Private Function ReadData() As Boolean
   Dim A4401 As String, A4402 As String, A4403 As String
   Dim iTopRow As Integer
   
   iTopRow = DataGrid1.FirstRow
   
   A4401 = Adodc1.Recordset.Fields("a4401")
   A4402 = Adodc1.Recordset.Fields("a4402")
   A4403 = Adodc1.Recordset.Fields("a4403")
   
   strSql = "update acc440 set a4416='X'" & _
      " where a4401='" & A4401 & "' and a4402=" & A4402 & " and a4403=" & A4403 & _
      " and a4416 is null"
   cnnConnection.Execute strSql, intI
   If intI = 1 Then
      With Frmacc7151
      .m_A4401 = A4401
      .m_A4402 = A4402
      .m_A4403 = A4403
      .lblSales = "" & Adodc1.Recordset.Fields("Sales")
      .lblRDate = "" & Adodc1.Recordset.Fields("Rdate")
      .lblCUser = "" & Adodc1.Recordset.Fields("CUser")
      .lblCDate = "" & Adodc1.Recordset.Fields("Cdate")
      .Show vbModal
      strFormName = Me.Name
      End With
      
      strSql = "update acc440 set a4416=Null" & _
         " where a4401='" & A4401 & "' and a4402=" & A4402 & " and a4403=" & A4403 & _
         " and a4416='X'"
      cnnConnection.Execute strSql, intI
   Else
      CheckStatus A4401, A4402, A4403
   End If
   QueryData
   If Adodc1.Recordset.RecordCount > 0 Then
      DataGrid1.Visible = False
      With Adodc1.Recordset
      Do While Not .EOF
         If .Fields("A4401") = A4401 And .Fields("A4402") = A4402 And .Fields("A4403") = A4403 Then
            Exit Do
         End If
         .MoveNext
      Loop
      If .EOF Then .MoveFirst
      End With
      DataGrid1.Scroll 0, iTopRow - DataGrid1.FirstRow
      DataGrid1.Visible = True
   End If
End Function

'檢查繳款記錄目前狀態
Private Function CheckStatus(pA4401 As String, pA4402 As String, pA4403 As String, Optional pNoMsg As Boolean = False, Optional pAddMsg As String) As Integer
   Dim stSQL As String, intR As Integer, stMsg As String
   Dim adoquery As ADODB.Recordset
   Dim stMsg2 As String
   
   
   stMsg2 = vbCrLf & vbCrLf & "若狀態有異常請通知電腦中心處理！"
   
   stSQL = "select A4416,st02,sqldatet(a4402) RDate,sqltime(a4403) RTime from acc440,staff" & _
      " where a4401='" & pA4401 & "' and a4402=" & pA4402 & " and a4403=" & pA4403 & _
      " and st01(+)=a4401"
   intR = 1
   Set adoquery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With adoquery
      stMsg = vbCrLf & vbCrLf & "智權人員:" & .Fields("st02") & _
               vbCrLf & "繳款日期:" & .Fields("RDate") & _
               vbCrLf & "繳款時間:" & .Fields("RTime")
               
      If .Fields("A4416") = "X" Then
         If pNoMsg = False Then
            'Modify by Amy 2022/04/18 +M31-Morgan:開放財務操作
            If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31" Then
               If MsgBox("目前狀態：繳款記錄出納確認中...." & pAddMsg & stMsg & vbCrLf & vbCrLf & "是否要解除鎖定？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                  stSQL = "update acc440 set A4416='' where a4401='" & pA4401 & "' and a4402=" & pA4402 & " and a4403=" & pA4403
                  cnnConnection.Execute stSQL, intR
                  If intR = 1 Then
                     MsgBox "已解除！", vbExclamation
                  Else
                     MsgBox "解除失敗！", vbCritical
                  End If
               End If
            Else
               MsgBox "目前狀態：繳款記錄出納確認中...." & pAddMsg & stMsg & stMsg2, vbExclamation
            End If
         End If
         CheckStatus = 1
      ElseIf .Fields("A4416") = "Y" Then
         If pNoMsg = False Then
            MsgBox "目前狀態：繳款記錄收款中...." & pAddMsg & stMsg & stMsg2, vbExclamation
         End If
         CheckStatus = 2
      ElseIf Not IsNull(.Fields("A4416")) Then
         If pNoMsg = False Then
            MsgBox "目前狀態：繳款記錄已收款!!" & pAddMsg & stMsg & stMsg2, vbExclamation
         End If
         CheckStatus = 3
      End If
      End With
   End If
   Set adoquery = Nothing
End Function

'Added by Lydia 2016/09/12
Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount = 0 Then
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If

   tool3_enabled

Call cmdDetail_Click
End Sub

Private Sub cmdDetail_Click()
'copy from Promoter.frm210142
Dim stVTB11 As String, stVTB22 As String
Dim iCol    As Integer
Dim rtNo    As String
Dim strCon  As String
Dim Role    As String
Dim PayDate As Long
Dim PayTime As Long
Dim dblVal(3) As Double
Dim adoquery As New ADODB.Recordset

   Role = Adodc1.Recordset("A4401")                '智權人員
   PayDate = Adodc1.Recordset("A4402")             '繳款日期
   PayTime = Adodc1.Recordset("A4403")             '繳款時間
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   'Modified by Lydia 2023/11/13 開立INVOICE，不列印收據;decode(nvl(a0k19,0),0,'◎')=> decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))
   strExc(0) = "select sqldatet(a0k02) 單據日期" & _
       ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
       ",decode(a0j04,'000',cpm03,cpm04) 案件性質" & _
       ",na03 國別,axd06 服務費,axd07 規費,axd08 扣繳金額,a0k01||decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))||decode(AXC01,null,'','＊') 收據編號" & _
       ",nvl(tm05,nvl(pa05,nvl(lc05,nvl(sp05,hc06)))) 案件名稱,a0k03,a0k04" & _
       " from ACC441,ACC0J0,acc0k0,acc431,caseprogress,casepropertymap,nation" & _
       ",trademark,patent,lawcase,servicepractice,hirecase" & _
       " where A0J01(+)=AXD05 AND A0J13(+)=AXD04" & _
       " and axd01='" & Role & "' and axd02='" & PayDate & "' and axd03='" & PayTime & "'" & _
       " and a0k01(+)=a0j13 and axc02(+)=a0j13 and cp09(+)=a0j01" & _
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
      .Caption = "智權人員繳款確認-繳款資料明細"
      .cmdOK(0).Visible = cmdOK(0).Enabled 'Added by Morgan 2017/12/18
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
         'Added by Morgan 2017/12/18
         If .m_Return = 0 Then
            cmdOK(0).Value = True
         End If
         'end 2017/12/18
      End With
   End If
   
End Sub
'end 2016/09/12


