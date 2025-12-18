VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm1104_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "轉案號關聯確認"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   990
   ClientWidth     =   6270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6270
   Begin VB.CommandButton cmdOK 
      Caption         =   "否(&N)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5400
      TabIndex        =   0
      Top             =   480
      Width           =   795
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4815
      Top             =   3270
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "是(&Y)"
      Height          =   400
      Index           =   1
      Left            =   4590
      TabIndex        =   1
      Top             =   480
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frm1104_1.frx":0000
      Height          =   2730
      Left            =   135
      TabIndex        =   2
      Top             =   960
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   4815
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "C00"
         Caption         =   "關聯種類"
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
         DataField       =   "C01"
         Caption         =   "狀態"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "C02"
         Caption         =   "本所案號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "C04"
         Caption         =   "收文日"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "####/##/##"
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
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1530.142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1665.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1094.74
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   135
      TabIndex        =   3
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frm1104_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/11 改成Form2.0 (DataGrid2)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/8/18 日期欄已修改
Option Explicit

Public m_CaseNoBefore As String '轉號前本所案號
Public m_CaseNoAfter As String '轉號後本所案號
Public m_form As Form '呼叫的表單 回傳 Me.tag:0=取消, 1=轉關聯 , 2=不轉關聯
Public m_CP09 As String '收文號
Dim varMouseCursor
Dim bolTransfer As Boolean

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         m_form.Tag = Index
      Case 1
         If bolTransfer = True Then
            m_form.Tag = 1
         Else
            m_form.Tag = 2
         End If
   End Select
   Unload Me
End Sub

Private Sub Form_Load()
   varMouseCursor = Screen.MousePointer
   Screen.MousePointer = vbDefault
   MoveFormToCenter Me
   InitGrid
End Sub

Private Sub SetMsg()
   Dim CNo
   CNo = Split(m_CaseNoBefore, "-")
   strExc(0) = "select * from caseprogress where cp01='" & CNo(0) & "' and cp02='" & CNo(1) & "' and cp03='" & CNo(2) & "' and cp04='" & CNo(3) & "' and cp09<>'" & m_CP09 & "' and cp10 in (" & CaseMapOut & ")"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Label1.Caption = "案號 " & m_CaseNoBefore & " 現有關聯案如下表；" & vbCrLf & "因本案尚有其他新案案件性質的程序，" & vbCrLf & "故該案的所有關聯將仍維持不變。" & vbCrLf & "是否確定要繼續？"
      bolTransfer = False
   Else
      Label1.Caption = "案號 " & m_CaseNoBefore & " 現有關聯案如下表；" & vbCrLf & "本作業會將該案的所有關聯更改為案號 " & m_CaseNoAfter & "，" & vbCrLf & "是否確定要繼續？"
      bolTransfer = True
   End If
End Sub

Private Sub InitGrid()
   'Modify by Amy 2014/06/10 +FormName 改暫存TB
   'Set Adodc2.Recordset = PUB_CreateRecordset(, 5)
   Set Adodc2.Recordset = PUB_CreateRecordset(, 5, , , Me.Name)
   Set DataGrid2.DataSource = Adodc2
End Sub

Public Function GetRelation() As Boolean
   Dim CNo
   CNo = Split(m_CaseNoBefore, "-")
   strExc(0) = "select decode(cm10,'0','國內外','1','美國IDS','2','大陸發明','3','一案兩請','4','大陸香港','5','多國','6','相關','7','分割',cm10) C00" & _
      ", decode(status,'1','國外案','2','國內案','3','母案','4','分案',status) C01" & _
      ", cm01||'-'||cm02||'-'||cm03||'-'||cm04 C02, na03 C03, cp05 C04" & _
      " from (select cm10,'1' status,cm01,cm02,cm03,cm04" & _
      " from casemap where cm05='" & CNo(0) & "' and cm06='" & CNo(1) & "' and cm07='" & CNo(2) & "' and cm08='" & CNo(3) & "'" & _
      " Union select cm10,'2',cm05,cm06,cm07,cm08" & _
      " from casemap where cm01='" & CNo(0) & "' and cm02='" & CNo(1) & "' and cm03='" & CNo(2) & "' and cm04='" & CNo(3) & "'" & _
      " Union select '5',null,CR05,CR06,CR07,CR08" & _
      " from caserelation where CR01='" & CNo(0) & "' AND CR02='" & CNo(1) & "' AND CR03='" & CNo(2) & "' AND CR04='" & CNo(3) & "'" & _
      " Union select '6',null,CR05,CR06,CR07,CR08" & _
      " from caserelation1 where CR01='" & CNo(0) & "' AND CR02='" & CNo(1) & "' AND CR03='" & CNo(2) & "' AND CR04='" & CNo(3) & "'" & _
      " Union SELECT '7','3',DC05,DC06,DC07,DC08" & _
      " FROM DIVISIONCASE WHERE DC01='" & CNo(0) & "' AND DC02='" & CNo(1) & "' AND DC03='" & CNo(2) & "' AND DC04='" & CNo(3) & "'" & _
      " Union SELECT '7','4',DC01,DC02,DC03,DC04" & _
      " FROM DIVISIONCASE WHERE DC05='" & CNo(0) & "' AND DC06='" & CNo(1) & "' AND DC07='" & CNo(2) & "' AND DC08='" & CNo(3) & "'" & _
      " ) X, patent, nation, caseprogress where pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04" & _
      " and na01(+)=pa09" & _
      " and cp01(+)=cm01 and cp02(+)=cm02 and cp03(+)=cm03 and cp04(+)=cm04 and cp31(+)='Y'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Set Adodc2.Recordset = RsTemp.Clone
      SetMsg
      GetRelation = True
   Else
      GetRelation = False
   End If
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   Screen.MousePointer = varMouseCursor
   Set frm1104_1 = Nothing
End Sub
