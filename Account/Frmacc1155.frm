VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc1155 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收款作業-整批"
   ClientHeight    =   5160
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9504
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9504
   Begin VB.CommandButton Command3 
      Caption         =   "明細"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6660
      TabIndex        =   8
      Top             =   90
      Width           =   1050
   End
   Begin VB.TextBox TxtComp 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3735
      MaxLength       =   1
      TabIndex        =   1
      Top             =   120
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "收款"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7740
      TabIndex        =   3
      Top             =   90
      Width           =   1455
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
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   90
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8010
      Top             =   4740
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1155.frx":0000
      Height          =   4305
      Left            =   90
      TabIndex        =   4
      Top             =   510
      Width           =   9195
      _ExtentX        =   16214
      _ExtentY        =   7599
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Amount"
         Caption         =   "繳款總金額"
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "SMemo"
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
      BeginProperty Column09 
         DataField       =   "Zone"
         Caption         =   "業務區"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1008
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   288
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   552.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   552.189
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1068.094
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   1535.811
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
            ColumnWidth     =   1368
         EndProperty
         BeginProperty Column09 
            Locked          =   -1  'True
            ColumnWidth     =   708.095
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1260.284
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "主要公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2520
      TabIndex        =   7
      Top             =   165
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收款日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   6
      Top             =   172
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "狀態：◎表示分所出納未確認。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   4860
      Width           =   2730
   End
End
Attribute VB_Name = "Frmacc1155"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/08 Form2.0已修改
'Created by Morgan 2013/12/12
Option Explicit

Public m_iReturn As Integer


Private Sub Command1_Click()
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox "請輸入收款日期！", vbExclamation, "收款檢查"
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label1 & MsgText(63), , MsgText(5)
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   
   If txtComp = "" Then
      MsgBox "請輸入主要公司別！", vbExclamation, "收款檢查"
      txtComp.SetFocus
      Exit Sub
   'Modify by Amy 2020/04/10
   'ElseIf txtComp <> "1" And txtComp <> "J" Then
   ElseIf InStr(GetBookKeepCmp, txtComp) = 0 Then
      MsgBox "主要公司別輸入錯誤！", vbExclamation, "收款檢查"
      txtComp.SetFocus
      Exit Sub
   End If
   
   If Adodc1.Recordset.Fields("Status") = "◎" Then
      MsgBox "分所出納尚未確認，不可收款！", vbExclamation, "收款檢查"
      Exit Sub
   End If
   
   ReadData
End Sub

Private Sub Command2_Click()
   OpenTable
End Sub

Private Sub Command3_Click()
   doQuery
End Sub

Private Sub Form_Activate()
   strFormName = Me.Name
   tool3_enabled
   MenuDisabled
   EraseNowRec
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
   MaskEdBox1.Mask = DFormat
   OpenTable
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   MenuEnabled
   Set Frmacc1155 = Nothing
End Sub

Private Sub OpenTable()
   'Modified by Morgan 2014/8/25 +外幣金額
   'Modified by Morgan 2015/7/15 +其他
   'Modify by Amy 2020/04/09 +公司別L
   'Modified by Morgan 2023/4/6 收款中也要顯示(前次整批收款當掉的資料)
   strExc(0) = "select s1.st02 Sales,JComp,a0k04" & _
      ",nvl(a4405,0)+nvl(a4406,0)+nvl(a4407,0)+nvl(a4408,0)+nvl(a4409,0)-nvl(a4410,0)+nvl(a4411,0)+nvl(a4422,0)+nvl(a4426,0)+nvl(a4430,0) Amount" & _
      ",decode(s1.st06,'1','',decode(a4413,null,'◎')) Status" & _
      ",substr(decode(sign(a4405),1,'支票')||decode(sign(nvl(a4406,0)+nvl(a4407,0)),1,'電匯')||decode(sign(a4408),1,'現金')||'其他',1,2) Type" & _
      ",sqldatet(to_char(A2324,'yyyymmdd')) IDate" & _
      ",sqldatet(a4402)||' '||sqltime(a4403) RDate" & _
      ",a4415 Memo,A0902 Zone,a0k03" & _
      ",A4401,A4402,A4403,s1.st06,A4413,A4423 from acc440,staff s1,acc090" & _
      ",(select axd01,axd02,axd03,min(axd04) axd04,max(decode(a0k11,'J',a0k11,'L',a0k11)) JComp from acc440,acc441,acc0k0 where axd01(+)=a4401 and axd02(+)=a4402 and axd03(+)=a4403" & _
      " and nvl(a4416,'Y')='Y' and a0k01(+)=axd04 group by axd01,axd02,axd03) x,acc0k0,acc230" & _
      " where nvl(a4416,'Y')='Y' and s1.st01(+)=a4401 and a0901(+)=s1.st15" & _
      " and axd01(+)=a4401 and axd02(+)=a4402 and axd03(+)=a4403 and a0k01(+)=axd04 and a2301(+)=a4421" & _
      " order by A4402,s1.st06,A4413,A4423,s1.st15,A4401,A4403"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/06/24 +FormName 改暫存TB
   Set Adodc1.Recordset = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   If RsTemp.RecordCount > 0 Then
      Command1.Enabled = True
      Command3.Enabled = True
   Else
      Command1.Enabled = False
      Command3.Enabled = False
   End If
End Sub

Private Function ReadData() As Boolean
   Dim A4401 As String, A4402 As String, A4403 As String
   
   A4401 = Adodc1.Recordset.Fields("a4401")
   A4402 = Adodc1.Recordset.Fields("a4402")
   A4403 = Adodc1.Recordset.Fields("a4403")
   
   'Added by Morgan 2023/4/13
   strExc(0) = "select distinct axd04 from acc440,acc441 where a4401='" & A4401 & "' and a4402=" & A4402 & " and a4403=" & A4403 & _
      " and axd01(+)=a4401 and axd02(+)=a4402 and axd03(+)=a4403 " & _
      " and not exists(select * from acc0j0 where a0j13=axd04)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(1) = RsTemp.GetString
      MsgBox "下列收據已經作廢或刪除，請通知智權同仁刪除繳款記錄，重新操作繳款作業！" & vbCrLf & vbCrLf & strExc(1), vbCritical
      Exit Function
   End If
   'end 2023/4/13
   
   'Added by Morgan 2025/4/22
   'ACS案若有記錄但是無ATR08智權人員比例時，不可收款並提醒
   strExc(0) = "select distinct axd04 from acc440,acc441 where a4401='" & A4401 & "' and a4402=" & A4402 & " and a4403=" & A4403 & _
      " and axd01(+)=a4401 and axd02(+)=a4402 and axd03(+)=a4403 " & _
      " and exists(select * from acc0j0,caseprogress,Acs_Tips_Rate where a0j13=axd04 and cp09(+)=a0j01 and atr01(+)=cp01 and atr02(+)=cp02 and atr03(+)=cp03 and atr04(+)=cp04 and atr05='1' and atr08 is null)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(1) = RsTemp.GetString
      MsgBox "繳款資料有ACS案應分配點數但顧服組尚未輸入比例，不可先行收款！", vbCritical
      Exit Function
   End If
   'end 2025/4/22
   
   
   'Modified by Morgan 2014/2/6
   'strExc(0) = "select * from acc441,acc0k0,acc431 where a0k01(+)=axd04 and a0k11='J' and axc02(+)=a0k01 and axc01 is null"
   strExc(0) = "select * from acc440,acc441,acc0k0,acc431 where a4401='" & A4401 & "' and a4402=" & A4402 & " and a4403=" & A4403 & " and axd01(+)=a4401 and axd02(+)=a4402 and axd03(+)=a4403 and a0k01(+)=axd04 and a0k11='J' and axc02(+)=a0k01 and axc01 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(1) = GetInvDate
      If Val(ChangeTDateStringToTString(MaskEdBox1)) < Val(strExc(1)) Then
         MsgBox "繳款資料有含J公司未開發票請款單，收款日期不可早於最後發票日【" & ChangeTStringToTDateString(strExc(1)) & "】!!", vbExclamation, "收款檢查"
         MaskEdBox1.SetFocus
         Exit Function
      'Added by Morgan 2014/9/18
      '未開發票請款單收款日期不可晚於下一工作日
      Else
         strExc(1) = PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))
         If Val(ChangeTDateStringToTString(MaskEdBox1)) > Val(strExc(1)) Then
            MsgBox "繳款資料有含J公司未開發票請款單，收款日期不可晚於下一工作日【" & ChangeTStringToTDateString(strExc(1)) & "】!!", vbExclamation, "收款檢查"
            MaskEdBox1.SetFocus
            Exit Function
         End If
      'end 2014/9/18
      End If
      'Added by Morgan 2014/3/4
      strExc(1) = Val(ChangeTDateStringToTString(MaskEdBox1)) \ 100
      strExc(0) = "select * from acc410 where a4101<=" & strExc(1) & " and a4102>=" & strExc(1)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI <> 1 Then
         MsgBox Left(MaskEdBox1, 6) & "月份發票資料尚未建立，無法開立發票!!", vbExclamation, "收款檢查"
         MaskEdBox1.SetFocus
         Exit Function
      End If
      'end 2014/3/4
   End If
   
   strSql = "update acc440 set a4416='Y'" & _
      " where a4401='" & A4401 & "' and a4402=" & A4402 & " and a4403=" & A4403 & _
      " and a4416 is null"
   cnnConnection.Execute strSql, intI
   If intI = 1 Then
      Me.MousePointer = vbHourglass
      tool1_enabled
      MenuDisabled
      m_iReturn = 0
      
      With Frmacc1150
      .Show
      .MaskEdBox1 = MaskEdBox1
      .Text21 = txtComp
      .m_A4401 = A4401
      .m_A4402 = A4402
      .m_A4403 = A4403
      .AutoProcess
      End With
      
      Me.Hide
      Me.MousePointer = vbDefault
   Else
      CheckStatus A4401, A4402, A4403
   End If
End Function

'檢查繳款記錄目前狀態
Private Function CheckStatus(pA4401 As String, pA4402 As String, pA4403 As String, Optional pNoMsg As Boolean = False) As Integer
   Dim stSQL As String, intR As Integer, stMsg As String
   Dim adoquery As ADODB.Recordset
   
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
            MsgBox "繳款記錄出納確認中...." & stMsg, vbInformation
         End If
         CheckStatus = 1
      ElseIf .Fields("A4416") = "Y" Then
         If pNoMsg = False Then
            MsgBox "繳款記錄收款中...." & stMsg, vbInformation
         End If
         CheckStatus = 2
      ElseIf Not IsNull(.Fields("A4416")) Then
         If pNoMsg = False Then
            MsgBox "繳款記錄已收款!!" & stMsg, vbInformation
         End If
         CheckStatus = 3
      End If
      End With
   End If
   Set adoquery = Nothing
End Function

'清除當前收文號
Private Sub EraseNowRec()
   Dim A4401 As String, A4402 As String, A4403 As String
   
   If m_iReturn <> 0 Then
      '取消收款
      If m_iReturn = -1 Then
         A4401 = Adodc1.Recordset.Fields("a4401")
         A4402 = Adodc1.Recordset.Fields("a4402")
         A4403 = Adodc1.Recordset.Fields("a4403")
         
         strSql = "update acc440 set a4416=''" & _
            " where a4401='" & A4401 & "' and a4402=" & A4402 & " and a4403=" & A4403 & _
            " and a4416='Y'"
         cnnConnection.Execute strSql, intI
      '確定收款
      Else
         With Adodc1.Recordset
         .Delete
         .UpdateBatch
         If .RecordCount = 0 Then
            Command3.Enabled = False
         End If
         End With
      End If
   End If
   m_iReturn = 0
End Sub

Private Sub txtComp_GotFocus()
   TextInverse txtComp
   CloseIme
End Sub

Private Sub txtComp_KeyPress(KeyAscii As Integer)
   Dim i As Integer
   Dim stTmp
   
   'Modify by Amy 2020/04/10
   stTmp = Split(GetBookKeepCmp, ",")
   KeyAscii = UpperCase(KeyAscii)
   For i = LBound(stTmp) To UBound(stTmp)
        'If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("J") Then
        If KeyAscii = Asc(stTmp(i)) Then
           Exit Sub
        End If
   Next i
   If KeyAscii <> 8 Then
        KeyAscii = 0
   End If
End Sub


Private Sub doQuery()
   Dim dblVal(3) As Double
   
   'Modified by Morgan 2021/11/10 公司別統一改用簡稱 a0k11-->a0820
   'Modified by Lydia 2023/11/13 開立INVOICE，不列印收據;decode(nvl(a0k19,0),0,'◎')=> decode(a0k32,'Z','',decode(nvl(a0k19,0),0,'◎'))
   strExc(0) = "select sqldatet(a0k02) 單據日期" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",decode(a0j04,'000',cpm03,cpm04) 案件性質" & _
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
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If RsTemp.RecordCount > 0 Then
      With frm210141_3
      Set .Adodc1.Recordset = RsTemp
         With RsTemp
         .MoveFirst
         Do While Not .EOF
            dblVal(1) = dblVal(1) + Val("" & .Fields("服務費"))
            dblVal(2) = dblVal(2) + Val("" & .Fields("規費"))
            dblVal(3) = dblVal(3) + Val("" & .Fields("扣繳金額"))
            .MoveNext
         Loop
         End With
         .Caption = Me.Caption & "-明細"
         .txtTot(2) = Format(dblVal(1), "#,##0")
         .txtTot(3) = Format(dblVal(2), "#,##0")
         .txtTot(4) = Format(dblVal(3), "#,##0")
         .txtTot(5) = Format(dblVal(1) + dblVal(2) - dblVal(3), "#,##0")
         .Show vbModal
      End With
   End If
End Sub

