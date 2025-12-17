VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc1123 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收據開立作業-整批"
   ClientHeight    =   5112
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   9372
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5112
   ScaleWidth      =   9372
   Begin VB.CheckBox Check1 
      Caption         =   "不顯示收費 0 或負點數案件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2928
      TabIndex        =   6
      Top             =   96
      Width           =   3108
   End
   Begin VB.TextBox txtDate 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   5
      Top             =   30
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重新整理"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6165
      TabIndex        =   2
      Top             =   90
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開立收據"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7695
      TabIndex        =   1
      Top             =   90
      Width           =   1485
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   675
      Top             =   2340
      Visible         =   0   'False
      Width           =   960
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1123.frx":0000
      Height          =   4380
      Left            =   90
      TabIndex        =   0
      Top             =   660
      Width           =   9195
      _ExtentX        =   16235
      _ExtentY        =   7726
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "Status"
         Caption         =   "狀態"
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
         DataField       =   "ST02"
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
         DataField       =   "C01"
         Caption         =   "本所案號"
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
         DataField       =   "C02"
         Caption         =   "案件性質"
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
         DataField       =   "C03"
         Caption         =   "申請國家"
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
         DataField       =   "C10"
         Caption         =   "出"
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
         DataField       =   "CRL119"
         Caption         =   "特"
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
      BeginProperty Column07 
         DataField       =   "C04"
         Caption         =   "服務費"
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
         DataField       =   "C05"
         Caption         =   "規費"
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
         DataField       =   "C06"
         Caption         =   "年費年度"
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
      BeginProperty Column10 
         DataField       =   "C07"
         Caption         =   "收文日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "C08"
         Caption         =   "所別"
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
      BeginProperty Column12 
         DataField       =   "CP09"
         Caption         =   "收文號"
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
         Size            =   275
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   432
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   852.095
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1235.906
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   815.811
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   288
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   264.189
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   648
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   852.095
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column12 
            Locked          =   -1  'True
            ColumnWidth     =   984.189
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "作業日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   90
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PS：移轉讓與案因收據客戶編號不固定，故不在整批開立！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   420
      Width           =   5085
   End
End
Attribute VB_Name = "Frmacc1123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Create by Morgan 2010/12/3
Option Explicit

Dim m_CP09 As String
Public m_Rtn As String
Public m_Continue As Boolean
'Dim WithEvents eventConn As ADODB.Connection 'Add By Sindy 2023/11/28


Private Sub Check1_Click()
   Command2.Value = True
End Sub

Private Sub Command1_Click()
   doBatch
End Sub

Public Sub doBatch()
   strFormName = Me.Name
   tool3_enabled
   MenuDisabled
   
   m_Continue = False
   With Adodc1.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         m_CP09 = .Fields("cp09")
         If CheckStatus Then
            Frmacc1124.Label2(0) = "" & .Fields("c00") & " " & .Fields("st02")
            Frmacc1124.Label2(1) = "" & .Fields("c01")
            Frmacc1124.Label2(2) = "" & .Fields("c02")
            Frmacc1124.Label2(3) = "" & .Fields("c03")
            Frmacc1124.Label2(4) = Format("" & .Fields("c04"), DDollar)
            Frmacc1124.Label2(5) = Format("" & .Fields("c05"), DDollar)
            Frmacc1124.Label2(6) = "" & .Fields("c06")
            Frmacc1124.Label2(7) = "" & .Fields("c07")
            Frmacc1124.Label2(8) = "" & .Fields("c08")
            Frmacc1124.Label2(9) = "" & .Fields("cp09")
            Frmacc1124.Label2(10) = "" & .Fields("app") & " " & .Fields("cu04")
            Frmacc1124.Show vbModal
            strFormName = Me.Name
            Select Case m_Rtn
               Case 0 '繼續
                  'Modify By Sindy 2023/1/12
                  'If Val("" & .Fields("c04")) <= 0 Then '費用0,開啟接洽單; Sindy 2023/5/2 取消:And Val("" & .Fields("c05")) = 0
                  'cp16-nvl(cp17,0)=c04; cp17=c05
                  If Val("" & .Fields("c04")) <= 0 And _
                     Not (Val("" & .Fields("c04")) = 0 And Val("" & .Fields("c05")) > 0) Then 'Modify by Sindy 2023/7/12 調整判斷if ex:CFP-31208(超項費)
                     strSql = "select cp140 from caseprogress where cp09='" & "" & m_CP09 & "' and cp140 is not null"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        Call PUB_Queryfrm090801(RsTemp.Fields("cp140"), "", Me, True)
                        InsertSkipRec '略過
                     End If
                  Else
                  '2023/1/12 END
                     If Not IsNull(.Fields("app")) Then
                        Frmacc1120.m_AutoProcess = True
                        Frmacc1120.m_CustNo = "" & .Fields("app")
                        Frmacc1120.m_CP09 = .Fields("cp09")
                        Frmacc1120.Caption = Frmacc1120.Caption & "-批次"
                        Frmacc1120.Show
                        Me.Hide
                     End If
                  End If
                  EraseNowRec
                  Exit Do
               Case 1 '略過
                  InsertSkipRec
                  EraseNowRec
               Case 2 '返回
                  Exit Do
            End Select
         Else
            EraseNowRec
         End If
         .MoveNext
      Loop
      If .EOF Then
         MsgBox "已無待開收據資料！"
      End If
   Else
      MsgBox "無待開收據資料！"
   End If
   End With
End Sub
'檢查是否已開收據
Private Function CheckStatus() As Boolean
   strSql = "select * from caseprogress where cp09='" & m_CP09 & "' and cp60 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      CheckStatus = True
   End If
End Function
'清除當前收文號
Private Sub EraseNowRec()
   With Adodc1.Recordset
   If .Fields("cp09") = m_CP09 Then
      .Delete
      .UpdateBatch
   End If
   End With
End Sub
'刪除舊資料
Private Sub DeleteOldSkipRec()
   strSql = "delete acc270 where a2702<to_char(sysdate-10,'yyyymmdd')"
   cnnConnection.Execute strSql, intI
End Sub
'新增略過資料
Private Sub InsertSkipRec()
   strSql = "insert into acc270 (a2701,a2702) values('" & m_CP09 & "'," & strSrvDate(1) & ")"
   cnnConnection.Execute strSql, intI
End Sub

Private Sub Command2_Click()
   'Modified by Morgan 2023/1/4 +作業日期
   If txtDate = "" Then
      MsgBox "請輸入作業日期！", vbExclamation
   ElseIf CheckIsTaiwanDate(txtDate) Then
      OpenTable
   End If
End Sub

Private Sub Form_Activate()
   strFormName = Me.Name
   tool3_enabled
   MenuDisabled
   If m_Continue Then
      doBatch
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(133) & ", " & MsgText(134)
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, 9500, 5500
   txtDate = strSrvDate(2)
   
'   'Add By Sindy 2023/11/28
'   Set eventConn = cnnConnection
'   KillCmdLog
'   '2023/11/28 END

   OpenTable
End Sub

'無FC代理人,非移轉讓與案,非C類來函,非國外部收文
'Modifed by Morgan 2013/5/9 排除待確認的電子送件程序
'Modify By Sindy 2013/12/25 增加特殊出名公司,出:c10
'Modify By Sindy 2014/2/10 增加特殊收據,特:CRL119
'Modify By Sindy 2014/4/9 增加狀態
Private Sub OpenTable()
   Dim strCP149 As String
   Dim strDate1 As String, strDate2 As String
  
   strDate2 = DBDATE(txtDate)
   strCP149 = CompWorkDay(3, strDate2, 1)
   strDate1 = PUB_GetWorkDay1(CompDate(2, -1, strDate2), True)
   
   '北所(前一工作日(不含)以後收文,含例假日加班)
   'Modified by Morgan 2015/10/30 +分所自動收文案件
   '專利
   'Modified by Morgan 2013/7/2 相同案號有待確認電子送件程序的也不能開
   'Modified by Morgan 2015/9/9 自動收文狀態放"自" ,decode(cp01||cp10||pa09||cp140,'P601000','紙本','P605000','紙本','') --> ,decode(cp140,'','','自')
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '1. P及CFP的605、606、607，北所未分案時不出現。
   '2. 狀態欄取消原來的'自'、'分案'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。
   'strSql = " select decode(length(cp31||cp149),9,'分案',decode(cp140,'','','自')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01"
   'Modify By Sindy 2023/1/12 接洽記錄單 金額"0" 也要顯示出來
   '                          cp16>0 => (cp16>0 or (cp16=0 and cp140 is not null))
   '                          decode(s1.st06,'1','9',s1.st06) sort1 => decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1
   'Modified by Morgan 2024/8/21 同一張接洽單有領證或年費未分案時也不要顯示
   strSql = " select decode(instr(cp01,'L'),0,'',decode(cp31,'Y','新紙','紙')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01" & _
      ",decode(na01,'000',cpm03,cpm04) c02,na03 c03,cp16-nvl(cp17,0) c04,cp17 c05" & _
      ",decode(instr(',601,605,606,607,',','||cp10||','),0,'',CP53||'-'||CP54) c06" & _
      ",substrb(' '||sqldatet(cp05),-9) c07,decode(s1.st06,'1','北','2','中','3','南','4','高','他') c08" & _
      ",decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1,decode(cp140,null,1,2) sort2,'' sort3,'' sort4,cp09,pa26 app,s2.st02,cu04,pa161 c10,CRL119" & _
      " from caseprogress a,staff s1,staff s2,patent,nation,casepropertymap,customer,consultrecordlist" & _
      " where (cp16>0 or (cp16=0 and cp140 is not null)) and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C' and cp60 is null and nvl(cp118,'N')<>'W'" & _
      " and cp05>" & strDate1 & " and cp05<=" & strDate2 & "" & _
      " and not exists(select * from acc270 where a2701=cp09)" & _
      " and s1.st01(+)=cp65 and (s1.st06='1' or cp140 is not null) and cp01 in ('P','CFP','FCP')" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and (substr(cp12,1,1)='S' or pa75 is null)" & _
      " and not exists (select * from caseprogress b" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05" & _
      " and b.cp55||b.cp93||b.cp94||b.cp95||b.cp96<>b.cp56||b.cp89||b.cp90||b.cp91||b.cp92)" & _
      " and na01(+)=pa09 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and s2.st01(+)=cp13" & _
      " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')" & _
      " and cp140=crl01(+) and not(cp01 in ('P','CFP') and cp10 in ('605','606','607') and nvl(cp157,0)=0)" & _
      " and not exists(select * from caseprogress b where cp140=a.cp140 and cp01 in ('P','CFP') and cp10 in ('601','605','606','607') and nvl(cp157,0)=0)"
   
   'Added by Morgan 2013/5/27 分案改用電子送件的案件(含分所)
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '1. P及CFP的605、606、607，北所未分案時不出現。
   '2. 狀態欄取消原來的'自'、'分案'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。
   'strSql = strSql & " union select decode(length(cp31||cp149),9,'分案',decode(cp140,'','','自')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01"
   'Modify By Sindy 2023/1/12 接洽記錄單 金額"0" 也要顯示出來
   '                          cp16>0 => (cp16>0 or (cp16=0 and cp140 is not null))
   '                          decode(s1.st06,'1','9',s1.st06) sort1 => decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1
   'Modified by Morgan 2024/8/21 同一張接洽單有領證或年費未分案時也不要顯示
   strSql = strSql & " union select decode(instr(cp01,'L'),0,'',decode(cp31,'Y','新紙','紙')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01" & _
      ",decode(na01,'000',cpm03,cpm04) c02,na03 c03,cp16-nvl(cp17,0) c04,cp17 c05" & _
      ",decode(instr(',601,605,606,607,',','||cp10||','),0,'',CP53||'-'||CP54) c06" & _
      ",substrb(' '||sqldatet(cp05),-9) c07,decode(s1.st06,'1','北','2','中','3','南','4','高','他') c08" & _
      ",decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1,decode(cp140,null,1,2) sort2,'' sort3,'' sort4,cp09,pa26 app,s2.st02,cu04,pa161 c10,CRL119" & _
      " from caseprogress a,staff s1,staff s2,patent,nation,casepropertymap,customer,consultrecordlist" & _
      " where (cp16>0 or (cp16=0 and cp140 is not null)) and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C' and cp60 is null and cp118='Y'" & _
      " and cp149=" & strCP149 & _
      " and not exists(select * from acc270 where a2701=cp09)" & _
      " and s1.st01(+)=cp65 and cp01 in ('P','CFP','FCP')" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and (substr(cp12,1,1)='S' or pa75 is null)" & _
      " and not exists (select * from caseprogress b" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05" & _
      " and b.cp55||b.cp93||b.cp94||b.cp95||b.cp96<>b.cp56||b.cp89||b.cp90||b.cp91||b.cp92)" & _
      " and na01(+)=pa09 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and s2.st01(+)=cp13" & _
      " and cp140=crl01(+) and not(cp01 in ('P','CFP') and cp10 in ('605','606','607') and nvl(cp157,0)=0)" & _
      " and not exists(select * from caseprogress b where cp140=a.cp140 and cp01 in ('P','CFP') and cp10 in ('601','605','606','607') and nvl(cp157,0)=0)"
      
   '商標
   'Modify By Sindy 2015/8/26 decode(length(cp31||cp149),9,'分案','') Status ==> decode(length(cp31||cp149),9,'分案',decode(cp01||cp10||tm10||cp140,'T303000','紙本','T201000','紙本','T211000','紙本','T206000','紙本','')) Status
   'Modified by Morgan 2015/9/9 自動收文狀態放"自"
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '2. 狀態欄取消原來的'自'、'分案'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。
   'strSql = strSql & " union select decode(length(cp31||cp149),9,'分案',decode(cp140,'','','自')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01"
   'Modify By Sindy 2023/1/12 接洽記錄單 金額"0" 也要顯示出來
   '                          cp16>0 => (cp16>0 or (cp16=0 and cp140 is not null))
   '                          decode(s1.st06,'1','9',s1.st06) sort1 => decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1
   strSql = strSql & " union select decode(instr(cp01,'L'),0,'',decode(cp31,'Y','新紙','紙')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01" & _
      ",decode(na01,'000',cpm03,cpm04) c02,na03 c03,cp16-nvl(cp17,0) c04,cp17 c05" & _
      ",'' c06,substrb(' '||sqldatet(cp05),-9) c07,decode(s1.st06,'1','北','2','中','3','南','4','高','他') c08" & _
      ",decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1,decode(cp140,null,1,2) sort2,'' sort3,'' sort4,cp09,tm23 app,s2.st02,cu04,tm130 c10,CRL119" & _
      " from caseprogress a,staff s1,staff s2,trademark,nation,casepropertymap,customer,consultrecordlist" & _
      " where (cp16>0 or (cp16=0 and cp140 is not null)) and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C' and cp60 is null and nvl(cp118,'N')<>'W'" & _
      " and cp05>" & strDate1 & " and cp05<=" & strDate2 & "" & _
      " and s1.st01(+)=cp65 and (s1.st06='1' or cp140 is not null) and cp01 in ('T','TF','CFT','FCT')" & _
      " and not exists(select * from acc270 where a2701=cp09)" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and (substr(cp12,1,1)='S' or tm44 is null)" & _
      " and not exists (select * from caseprogress b" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05" & _
      " and b.cp55||b.cp93||b.cp94||b.cp95||b.cp96<>b.cp56||b.cp89||b.cp90||b.cp91||b.cp92)" & _
      " and na01(+)=tm10 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) and s2.st01(+)=cp13" & _
      " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')" & _
      " and cp140=crl01(+)"
      
   '法務  'modify by sonia 2019/7/24 +ACS系統類別
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '2. 狀態欄取消原來的'自'、'分案'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。
   'strSql = strSql & " union select decode(length(cp31||cp149),9,'分案',decode(cp140,'','','自')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01"
   'Modify By Sindy 2023/1/12 接洽記錄單 金額"0" 也要顯示出來
   '                          cp16>0 => (cp16>0 or (cp16=0 and cp140 is not null))
   '                          decode(s1.st06,'1','9',s1.st06) sort1 => decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1
   strSql = strSql & " union select decode(instr(cp01,'L'),0,'',decode(cp31,'Y','新紙','紙')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01" & _
      ",decode(na01,'000',cpm03,cpm04) c02,na03 c03,cp16-nvl(cp17,0) c04,cp17 c05" & _
      ",'' c06,substrb(' '||sqldatet(cp05),-9) c07,decode(s1.st06,'1','北','2','中','3','南','4','高','他') c08" & _
      ",decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1,decode(cp140,null,1,2) sort2,'' sort3,'' sort4,cp09,lc11 app,s2.st02,cu04,lc48 c10,CRL119" & _
      " from caseprogress a,staff s1,staff s2,lawcase,nation,casepropertymap,customer,consultrecordlist" & _
      " where (cp16>0 or (cp16=0 and cp140 is not null)) and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C' and cp60 is null and nvl(cp118,'N')<>'W'" & _
      " and cp05>" & strDate1 & " and cp05<=" & strDate2 & "" & _
      " and s1.st01(+)=cp65 and (s1.st06='1' or cp140 is not null) and cp01 in ('L','CFL','FCL','LIN','ACS')" & _
      " and not exists(select * from acc270 where a2701=cp09)" & _
      " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
      " and (substr(cp12,1,1)='S' or lc22 is null)" & _
      " and not exists (select * from caseprogress b" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05" & _
      " and b.cp55||b.cp93||b.cp94||b.cp95||b.cp96<>b.cp56||b.cp89||b.cp90||b.cp91||b.cp92)" & _
      " and na01(+)=lc15 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(lc11,1,8) and cu02(+)=substr(lc11,9) and s2.st01(+)=cp13" & _
      " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')" & _
      " and cp140=crl01(+)"
      
   '服務   'modify by sonia 2019/7/24 +排除ACS系統類別
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '2. 狀態欄取消原來的'自'、'分案'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。
   'strSql = strSql & " union select decode(length(cp31||cp149),9,'分案',decode(cp140,'','','自')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01"
   'Modify By Sindy 2023/1/12 接洽記錄單 金額"0" 也要顯示出來
   '                          cp16>0 => (cp16>0 or (cp16=0 and cp140 is not null))
   '                          decode(s1.st06,'1','9',s1.st06) sort1 => decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1
   strSql = strSql & " union select decode(instr(cp01,'L'),0,'',decode(cp31,'Y','新紙','紙')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01" & _
      ",decode(na01,'000',cpm03,cpm04) c02,na03 c03,cp16-nvl(cp17,0) c04,cp17 c05" & _
      ",'' c06,substrb(' '||sqldatet(cp05),-9) c07,decode(s1.st06,'1','北','2','中','3','南','4','高','他') c08" & _
      ",decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1,decode(cp140,null,1,2) sort2,'' sort3,'' sort4,cp09,sp08 app,s2.st02,cu04,sp85 c10,CRL119" & _
      " from caseprogress a,staff s1,staff s2,servicepractice,nation,casepropertymap,customer,consultrecordlist" & _
      " where (cp16>0 or (cp16=0 and cp140 is not null)) and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C' and cp60 is null and nvl(cp118,'N')<>'W'" & _
      " and cp05>" & strDate1 & " and cp05<=" & strDate2 & "" & _
      " and s1.st01(+)=cp65 and (s1.st06='1' or cp140 is not null) and cp01 not in ('P','CFP','FCP','T','TF','CFT','FCT','L','CFL','FCL','LIN','ACS','LA')" & _
      " and not exists(select * from acc270 where a2701=cp09)" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and (substr(cp12,1,1)='S' or sp26 is null)" & _
      " and not exists (select * from caseprogress b" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05" & _
      " and b.cp55||b.cp93||b.cp94||b.cp95||b.cp96<>b.cp56||b.cp89||b.cp90||b.cp91||b.cp92)" & _
      " and na01(+)=sp09 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and s2.st01(+)=cp13" & _
      " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')" & _
      " and cp140=crl01(+)"
      
   '顧問
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '2. 狀態欄取消原來的'自'、'分案'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。
   'strSql = strSql & " union select decode(length(cp31||cp149),9,'分案',decode(cp140,'','','自')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01"
   'Modify By Sindy 2023/1/12 接洽記錄單 金額"0" 也要顯示出來
   '                          cp16>0 => (cp16>0 or (cp16=0 and cp140 is not null))
   '                          decode(s1.st06,'1','9',s1.st06) sort1 => decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1
   strSql = strSql & " union select decode(instr(cp01,'L'),0,'',decode(cp31,'Y','新紙','紙')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01" & _
      ",decode(na01,'000',cpm03,cpm04) c02,na03 c03,cp16-nvl(cp17,0) c04,cp17 c05" & _
      ",'' c06,substrb(' '||sqldatet(cp05),-9) c07,decode(s1.st06,'1','北','2','中','3','南','4','高','他') c08" & _
      ",decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1,decode(cp140,null,1,2) sort2,'' sort3,'' sort4,cp09,hc05 app,s2.st02,cu04,'' c10,CRL119" & _
      " from caseprogress a,staff s1,staff s2,hirecase,nation,casepropertymap,customer,consultrecordlist" & _
      " where (cp16>0 or (cp16=0 and cp140 is not null)) and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C' and cp60 is null and nvl(cp118,'N')<>'W'" & _
      " and cp05>" & strDate1 & " and cp05<=" & strDate2 & "" & _
      " and s1.st01(+)=cp65 and (s1.st06='1' or cp140 is not null) and cp01='LA'" & _
      " and not exists(select * from acc270 where a2701=cp09)" & _
      " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04" & _
      " and not exists (select * from caseprogress b" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05" & _
      " and b.cp55||b.cp93||b.cp94||b.cp95||b.cp96<>b.cp56||b.cp89||b.cp90||b.cp91||b.cp92)" & _
      " and na01(+)='000' and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(hc05,1,8) and cu02(+)=substr(hc05,9) and s2.st01(+)=cp13" & _
      " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')" & _
      " and cp140=crl01(+)"
   
   '分所(前一工作日至前一日曆日收文,含例假日加班)
   '專利
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '1. P及CFP的605、606、607，北所未分案時不出現。
   '2. 狀態欄取消原來的'自'、'分案'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。
   'strSql = strSql & " union select decode(length(cp31||cp149),9,'分案',decode(cp140,'','','自')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01"
   'Modify By Sindy 2023/1/12 接洽記錄單 金額"0" 也要顯示出來
   '                          cp16>0 => (cp16>0 or (cp16=0 and cp140 is not null))
   '                          decode(s1.st06,'1','9',s1.st06) sort1 => decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1
   strSql = strSql & " union select decode(instr(cp01,'L'),0,'',decode(cp31,'Y','新紙','紙')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01" & _
      ",decode(na01,'000',cpm03,cpm04) c02,na03 c03,cp16-nvl(cp17,0) c04,cp17 c05" & _
      ",decode(instr(',601,605,606,607,',','||cp10||','),0,'',CP53||'-'||CP54) c06" & _
      ",substrb(' '||sqldatet(cp05),-9) c07,decode(s1.st06,'1','北','2','中','3','南','4','高','他') c08" & _
      ",decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1,decode(cp140,null,1,2) sort2,decode(cp01,'P','1','CFP','2','29') sort3,cp09 sort4,cp09,pa26 app,s2.st02,cu04,pa161 c10,CRL119" & _
      " from caseprogress a,staff s1,staff s2,patent,nation,casepropertymap,customer,consultrecordlist" & _
      " where (cp16>0 or (cp16=0 and cp140 is not null)) and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C' and cp60 is null and nvl(cp118,'N')<>'W'" & _
      " and cp05>=" & strDate1 & " and cp05<" & strDate2 & "" & _
      " and s1.st01(+)=cp65 and s1.st06<>'1' and cp140 is null and cp01 in ('P','CFP','FCP')" & _
      " and not exists(select * from acc270 where a2701=cp09)" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and (substr(cp12,1,1)='S' or pa75 is null)" & _
      " and not exists (select * from caseprogress b" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05" & _
      " and b.cp55||b.cp93||b.cp94||b.cp95||b.cp96<>b.cp56||b.cp89||b.cp90||b.cp91||b.cp92)" & _
      " and na01(+)=pa09 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and s2.st01(+)=cp13" & _
      " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')" & _
      " and cp140=crl01(+) and not(cp01 in ('P','CFP') and cp10 in ('605','606','607') and nvl(cp157,0)=0)"
      
   '商標
   'Modify By Sindy 2015/8/26 decode(length(cp31||cp149),9,'分案','') Status ==> decode(length(cp31||cp149),9,'分案',decode(cp01||cp10||tm10||cp140,'T303000','紙本','T201000','紙本','T211000','紙本','T206000','紙本','')) Status
   'Modified by Morgan 2015/9/9 自動收文狀態放"自"
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '2. 狀態欄取消原來的'自'、'分案'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。
   'strSql = strSql & " union select decode(instr(cp01,'L'),0,'',decode(cp31,'Y','新紙','紙')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01"
   'Modify By Sindy 2023/1/12 接洽記錄單 金額"0" 也要顯示出來
   '                          cp16>0 => (cp16>0 or (cp16=0 and cp140 is not null))
   '                          decode(s1.st06,'1','9',s1.st06) sort1 => decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1
   strSql = strSql & " union select decode(length(cp31||cp149),9,'分案',decode(cp140,'','','自')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01" & _
      ",decode(na01,'000',cpm03,cpm04) c02,na03 c03,cp16-nvl(cp17,0) c04,cp17 c05" & _
      ",'' c06,substrb(' '||sqldatet(cp05),-9) c07,decode(s1.st06,'1','北','2','中','3','南','4','高','他') c08" & _
      ",decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1,decode(cp140,null,1,2) sort2,decode(cp01,'T','3','CFT','4','49') sort3,cp09 sort4,cp09,tm23 app,s2.st02,cu04,tm130 c10,CRL119" & _
      " from caseprogress a,staff s1,staff s2,trademark,nation,casepropertymap,customer,consultrecordlist" & _
      " where (cp16>0 or (cp16=0 and cp140 is not null)) and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C' and cp60 is null and nvl(cp118,'N')<>'W'" & _
      " and cp05>=" & strDate1 & " and cp05<" & strDate2 & "" & _
      " and s1.st01(+)=cp65 and s1.st06<>'1' and cp140 is null and cp01 in ('T','TF','CFT','FCT')" & _
      " and not exists(select * from acc270 where a2701=cp09)" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and (substr(cp12,1,1)='S' or tm44 is null)" & _
      " and not exists (select * from caseprogress b" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05" & _
      " and b.cp55||b.cp93||b.cp94||b.cp95||b.cp96<>b.cp56||b.cp89||b.cp90||b.cp91||b.cp92)" & _
      " and na01(+)=tm10 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) and s2.st01(+)=cp13" & _
      " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')" & _
      " and cp140=crl01(+)"
      
   '法務   'modify by sonia 2019/7/24 +ACS系統類別
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '2. 狀態欄取消原來的'自'、'分案'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。
   'strSql = strSql & " union select decode(length(cp31||cp149),9,'分案',decode(cp140,'','','自')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01"
   'Modify By Sindy 2023/1/12 接洽記錄單 金額"0" 也要顯示出來
   '                          cp16>0 => (cp16>0 or (cp16=0 and cp140 is not null))
   '                          decode(s1.st06,'1','9',s1.st06) sort1 => decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1
   strSql = strSql & " union select decode(instr(cp01,'L'),0,'',decode(cp31,'Y','新紙','紙')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01" & _
      ",decode(na01,'000',cpm03,cpm04) c02,na03 c03,cp16-nvl(cp17,0) c04,cp17 c05" & _
      ",'' c06,substrb(' '||sqldatet(cp05),-9) c07,decode(s1.st06,'1','北','2','中','3','南','4','高','他') c08" & _
      ",decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1,decode(cp140,null,1,2) sort2,decode(cp01,'L','5','CFL','7','FCL','71','59') sort3,cp09 sort4,cp09,lc11 app,s2.st02,cu04,lc48 c10,CRL119" & _
      " from caseprogress a,staff s1,staff s2,lawcase,nation,casepropertymap,customer,consultrecordlist" & _
      " where (cp16>0 or (cp16=0 and cp140 is not null)) and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C' and cp60 is null and nvl(cp118,'N')<>'W'" & _
      " and cp05>=" & strDate1 & " and cp05<" & strDate2 & "" & _
      " and s1.st01(+)=cp65 and s1.st06<>'1' and cp140 is null and cp01 in ('L','CFL','FCL','LIN','ACS')" & _
      " and not exists(select * from acc270 where a2701=cp09)" & _
      " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
      " and (substr(cp12,1,1)='S' or lc22 is null)" & _
      " and not exists (select * from caseprogress b" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05" & _
      " and b.cp55||b.cp93||b.cp94||b.cp95||b.cp96<>b.cp56||b.cp89||b.cp90||b.cp91||b.cp92)" & _
      " and na01(+)=lc15 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(lc11,1,8) and cu02(+)=substr(lc11,9) and s2.st01(+)=cp13" & _
      " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')" & _
      " and cp140=crl01(+)"
      
   '服務    'modify by sonia 2019/7/24 +排除ACS系統類別
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '2. 狀態欄取消原來的'自'、'分案'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。
   'strSql = strSql & " union select decode(length(cp31||cp149),9,'分案',decode(cp140,'','','自')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01"
   'Modify By Sindy 2023/1/12 接洽記錄單 金額"0" 也要顯示出來
   '                          cp16>0 => (cp16>0 or (cp16=0 and cp140 is not null))
   '                          decode(s1.st06,'1','9',s1.st06) sort1 => decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1
   'Modified by Morgan 2025/6/11 排除TT999999否則會重複，下面會單獨抓且不限制收文日期
   strSql = strSql & " union select decode(instr(cp01,'L'),0,'',decode(cp31,'Y','新紙','紙')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01" & _
      ",decode(na01,'000',cpm03,cpm04) c02,na03 c03,cp16-nvl(cp17,0) c04,cp17 c05" & _
      ",'' c06,substrb(' '||sqldatet(cp05),-9) c07,decode(s1.st06,'1','北','2','中','3','南','4','高','他') c08" & _
      ",decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1,decode(cp140,null,1,2) sort2,decode(cp01,'PS','11','CPS','21','CFC','41','S','42','31') sort3,cp09 sort4,cp09,sp08 app,s2.st02,cu04,sp85 c10,CRL119" & _
      " from caseprogress a,staff s1,staff s2,servicepractice,nation,casepropertymap,customer,consultrecordlist" & _
      " where (cp16>0 or (cp16=0 and cp140 is not null)) and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C' and cp60 is null and nvl(cp118,'N')<>'W'" & _
      " and cp05>=" & strDate1 & " and cp05<" & strDate2 & "" & _
      " and s1.st01(+)=cp65 and s1.st06<>'1' and cp140 is null and cp01 not in ('P','CFP','FCP','T','TF','CFT','FCT','L','CFL','FCL','LIN','ACS','LA') and not (cp01='TT' and cp02='999999')" & _
      " and not exists(select * from acc270 where a2701=cp09)" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and (substr(cp12,1,1)='S' or sp26 is null)" & _
      " and not exists (select * from caseprogress b" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05" & _
      " and b.cp55||b.cp93||b.cp94||b.cp95||b.cp96<>b.cp56||b.cp89||b.cp90||b.cp91||b.cp92)" & _
      " and na01(+)=sp09 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and s2.st01(+)=cp13" & _
      " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')" & _
      " and cp140=crl01(+)"
      
   '顧問
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '2. 狀態欄取消原來的'自'、'分案'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。
   'strSql = strSql & " union select decode(length(cp31||cp149),9,'分案',decode(cp140,'','','自')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01"
   'Modify By Sindy 2023/1/12 接洽記錄單 金額"0" 也要顯示出來
   '                          cp16>0 => (cp16>0 or (cp16=0 and cp140 is not null))
   '                          decode(s1.st06,'1','9',s1.st06) sort1 => decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1
   strSql = strSql & " union select decode(instr(cp01,'L'),0,'',decode(cp31,'Y','新紙','紙')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01" & _
      ",decode(na01,'000',cpm03,cpm04) c02,na03 c03,cp16-nvl(cp17,0) c04,cp17 c05" & _
      ",'' c06,substrb(' '||sqldatet(cp05),-9) c07,decode(s1.st06,'1','北','2','中','3','南','4','高','他') c08" & _
      ",decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1,decode(cp140,null,1,2) sort2,'6' sort3,cp09 sort4,cp09,hc05 app,s2.st02,cu04,'' c10,CRL119" & _
      " from caseprogress a,staff s1,staff s2,hirecase,nation,casepropertymap,customer,consultrecordlist" & _
      " where (cp16>0 or (cp16=0 and cp140 is not null)) and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C' and cp60 is null and nvl(cp118,'N')<>'W'" & _
      " and cp05>=" & strDate1 & " and cp05<" & strDate2 & "" & _
      " and s1.st01(+)=cp65 and s1.st06<>'1' and cp140 is null and cp01='LA'" & _
      " and not exists(select * from acc270 where a2701=cp09)" & _
      " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04" & _
      " and not exists (select * from caseprogress b" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05" & _
      " and b.cp55||b.cp93||b.cp94||b.cp95||b.cp96<>b.cp56||b.cp89||b.cp90||b.cp91||b.cp92)" & _
      " and na01(+)='000' and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(hc05,1,8) and cu02(+)=substr(hc05,9) and s2.st01(+)=cp13" & _
      " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')" & _
      " and cp140=crl01(+)"
      
   'Add by Amy 2020/05/14 +TT-999999不限制收文日期
   'Modified by Morgan 2023/1/7 --辜,秀玲
   '2. 狀態欄取消原來的'自'、'分案'，改為法律所案件出現'紙'，若法律所新案號出現'新紙'。
   'strSql = strSql & " union select decode(length(cp31||cp149),9,'分案',decode(cp140,'','','自')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01"
   'Modify By Sindy 2023/1/12 接洽記錄單 金額"0" 也要顯示出來
   '                          cp16>0 => (cp16>0 or (cp16=0 and cp140 is not null))
   '                          decode(s1.st06,'1','9',s1.st06) sort1 => decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1
   strSql = strSql & " union select decode(instr(cp01,'L'),0,'',decode(cp31,'Y','新紙','紙')) Status,cp13 c00,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) c01" & _
      ",decode(na01,'000',cpm03,cpm04) c02,na03 c03,cp16-nvl(cp17,0) c04,cp17 c05" & _
      ",'' c06,substrb(' '||sqldatet(cp05),-9) c07,decode(s1.st06,'1','北','2','中','3','南','4','高','他') c08" & _
      ",decode(nvl(cp16,0),0,'Z',decode(s1.st06,'1','9',s1.st06)) sort1,decode(cp140,null,1,2) sort2,decode(cp01,'PS','11','CPS','21','CFC','41','S','42','31') sort3,cp09 sort4,cp09,sp08 app,s2.st02,cu04,sp85 c10,CRL119" & _
      " from caseprogress a,staff s1,staff s2,servicepractice,nation,casepropertymap,customer,consultrecordlist" & _
      " where (cp16>0 or (cp16=0 and cp140 is not null)) and cp57 is null and cp20 is null and cp09<'C' and cp60 is null " & _
      " and s1.st01(+)=cp65 and cp01='TT' and cp02='999999'" & _
      " and not exists(select * from acc270 where a2701=cp09)" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and (substr(cp12,1,1)='S' or sp26 is null)" & _
      " and not exists (select * from caseprogress b" & _
      " where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp05=a.cp05" & _
      " and b.cp55||b.cp93||b.cp94||b.cp95||b.cp96<>b.cp56||b.cp89||b.cp90||b.cp91||b.cp92)" & _
      " and na01(+)=sp09 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and s2.st01(+)=cp13" & _
      " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp118='W')" & _
      " and cp140=crl01(+)"
   'end 2020/05/14
   'Modify by Amy 2020/03/26 +系統種類對照表,SK02為法務案、顧問案公司別顯示L
   'Modify by Amy 2020/04/10 原判斷SK02改為系統別,顯示L
   If strSrvDate(1) >= 智慧所更名日 Then
      strSql = "Select Status,C00,C01,C02,C03,C04,C05,C06,C07,C08,Sort1,Sort2,Sort3,Sort4,CP09,App,St02,Cu04,Decode(InStr(SubStr(C01,1,instr(C01,'-')-1),'L'),0,C10,'L') C10,Crl119 From (" & strSql & "),SystemKind Where SubStr(C01,1,instr(C01,'-')-1)=Sk01(+) "
      'Added by Morgan 2024/7/12 不顯示收費 0 或負點數案件
      If Check1.Value = vbChecked Then
         strSql = strSql & " and (c04+nvl(c05,0)>0 and c04>=0)"
      End If
      'end 2024/7/12
   End If
   'Modified by Morgan 2017/10/19  排序取消系統別,收文號改由小到大--瑞婷
   'strSql = strSql & " order by sort1,sort2,sort3,sort4 desc,cp09"
   strSql = strSql & " order by sort1 asc,sort2 asc,cp09 asc"
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   WCmdLog "ClsLawReadRstMsg 結束" 'Add By Sindy 2023/11/28
   
   'Modify by Amy 2014/06/24 +FormName 改暫存TB
   Set Adodc1.Recordset = PUB_CreateRecordset(RsTemp, , , 100, Me.Name)
'   WCmdLog "PUB_CreateRecordset 結束" 'Add By Sindy 2023/11/28
End Sub

''Add By Sindy 2023/11/28
'Private Sub eventConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
'   WCmdLog pCommand.CommandText
'End Sub
'Function WCmdLog(oStrLog As String)
'On Error GoTo ErrHnd
'
'Dim ffa As Integer
'ffa = FreeFile
'Open App.path & "\$$cmdlog_" & Me.Name & ".log" For Append As ffa
'Print #ffa, Trim(Now) & "  ==>  " & oStrLog
'Close ffa
'
'ErrHnd:
'End Function
'Private Sub KillCmdLog()
'On Error GoTo ErrHnd
'   Kill App.path & "\$$cmdlog_" & Me.Name & ".log"
'ErrHnd:
'End Sub
''2023/11/28 END

Private Sub Form_Unload(Cancel As Integer)
'   Set eventConn = Nothing 'Add By Sindy 2023/11/28
   StatusClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Call PUB_GetLock("", "Frmacc1120") 'Added by Morgan 2024/5/24
   Call PUB_GetLock("", Me.Name)
   Set Frmacc1123 = Nothing
End Sub

Private Sub txtDate_GotFocus()
TextInverse txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
   KeyAscii = 0
   Beep
End If
End Sub
