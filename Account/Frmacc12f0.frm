VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc12f0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收文金額異常檢查"
   ClientHeight    =   4560
   ClientLeft      =   40
   ClientTop       =   320
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   9270
   Begin VB.CommandButton Command4 
      Caption         =   "檢視接洽單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4944
      TabIndex        =   5
      Top             =   192
      Width           =   1410
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc12f0.frx":0000
      Height          =   3948
      Left            =   48
      TabIndex        =   2
      Top             =   516
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   6967
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "收文金額異常檢查"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "C01"
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
         DataField       =   "C02"
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
      BeginProperty Column02 
         DataField       =   "C03"
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
      BeginProperty Column03 
         DataField       =   "C04"
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
         DataField       =   "C05"
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
      BeginProperty Column05 
         DataField       =   "C06"
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
      BeginProperty Column06 
         DataField       =   "C07"
         Caption         =   "收文日"
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
         DataField       =   "C08"
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
         BeginProperty Column00 
            ColumnWidth     =   860.032
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1489.89
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1159.937
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   980.221
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   980.221
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   909.921
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   6528
      Top             =   168
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   564
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1275
      TabIndex        =   0
      Top             =   157
      Width           =   1575
      _ExtentX        =   2769
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3195
      TabIndex        =   1
      Top             =   157
      Width           =   1575
      _ExtentX        =   2769
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2955
      TabIndex        =   4
      Top             =   180
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收文日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   3
      Top             =   180
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc12f0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2024/8/12
Option Explicit

Private Sub Command4_Click()
   If DataGrid1.row >= 0 Then
      Call PUB_Queryfrm090801("" & Adodc1.Recordset.Fields("cp140"), "", Me)
   End If
End Sub

Private Sub DataGrid1_DblClick()
   Command4.Value = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height, strBackPicPath2
   
   FormClear
   
   MaskEdBox1 = CFDate(ChangeWStringToTString(CompDate(1, -1, (Left(strSrvDate(1), 6) & "01"))))
   MaskEdBox2 = CFDate(ChangeWStringToTString(strSrvDate(1)))
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc12f0 = Nothing
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
   Screen.MousePointer = vbHourglass

   Dim stCon As String, stVTable As String, stCon1 As String
   Dim stConPA As String, stConTM As String, stConSP As String, stConLC As String
   
On Error GoTo Checking

   stCon = "": stConPA = "": stConTM = "": stConSP = "": stConLC = ""
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      stCon = stCon & " and cp05 >= " & Val(CADate(FCDate(MaskEdBox1.Text))) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      stCon = stCon & " and cp05 <= " & Val(CADate(FCDate(MaskEdBox2.Text))) & ""
   End If
   
   'Modified by Morgan 2024/11/29 改已開收據也要顯示--瑞婷
   'modify by sonia 2025/4/11 只要是CP05符合CP18<0都要出現，也先不管是否有CP60，不管其他CP條件--瑞婷：P-134622(113/11/7內部收文其他)
   '專利
   strSql = "select s2.st02 C01,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) C02,decode(na01,'000',cpm03,cpm04) C03" & _
      ",na03 C04,nvl(cp16,0)-nvl(cp17,0) C05,cp17 C06,sqldatet(cp05) C07,cp09 C08,cp140,cp05,cp09" & _
      " from caseprogress a,staff s1,staff s2,patent,nation,casepropertymap" & _
      " Where (cp18<0 or (cp140 Is Not Null and (cp16 = 0)" & _
      " and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C')" & _
      " and (substr(cp12,1,1)='S' or pa75 is null))" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stCon & _
      " and s1.st01(+)=cp65 and s2.st01(+)=cp13 and na01(+)=pa09 and cpm01(+)=cp01 and cpm02(+)=cp10"
   '商標
   strSql = strSql & " union select s2.st02 C01,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) C02,decode(na01,'000',cpm03,cpm04) C03" & _
      ",na03 C04,nvl(cp16,0)-nvl(cp17,0) C05,cp17 C06,sqldatet(cp05) C07,cp09 C08,cp140,cp05,cp09" & _
      " from caseprogress a,staff s1,staff s2,trademark,nation,casepropertymap" & _
      " Where (cp18<0 or (cp140 Is Not Null and (cp16 = 0)" & _
      " and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C')" & _
      " and (substr(cp12,1,1)='S' or tm44 is null))" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & stCon & _
      " and s1.st01(+)=cp65 and s2.st01(+)=cp13 and na01(+)=tm10 and cpm01(+)=cp01 and cpm02(+)=cp10"
   '服務
   strSql = strSql & " union select s2.st02 C01,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) C02,decode(na01,'000',cpm03,cpm04) C03" & _
      ",na03 C04,nvl(cp16,0)-nvl(cp17,0) C05,cp17 C06,sqldatet(cp05) C07,cp09 C08,cp140,cp05,cp09" & _
      " from caseprogress a,servicepractice,staff s1,staff s2,nation,casepropertymap" & _
      " Where (cp18<0 or (cp140 Is Not Null and (cp16 = 0)" & _
      " and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C')" & _
      " and (substr(cp12,1,1)='S' or sp26 is null))" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stCon & _
      " and s1.st01(+)=cp65 and s2.st01(+)=cp13 and na01(+)=sp09 and cpm01(+)=cp01 and cpm02(+)=cp10"
   '法務
   strSql = strSql & " union select s2.st02 C01,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) C02,decode(na01,'000',cpm03,cpm04) C03" & _
      ",na03 C04,nvl(cp16,0)-nvl(cp17,0) C05,cp17 C06,sqldatet(cp05) C07,cp09 C08,cp140,cp05,cp09" & _
      " from caseprogress a,staff s1,staff s2,lawcase,nation,casepropertymap" & _
      " Where (cp18<0 or (cp140 Is Not Null and (cp16 = 0)" & _
      " and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C')" & _
      " and (substr(cp12,1,1)='S' or lc22 is null))" & _
      " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04 and lc01 is not null" & stCon & _
      " and s1.st01(+)=cp65 and s2.st01(+)=cp13 and na01(+)=lc15 and cpm01(+)=cp01 and cpm02(+)=cp10"
   '顧問
   strSql = strSql & " union select s2.st02 C01,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) C02,decode(na01,'000',cpm03,cpm04) C03" & _
      ",na03 C04,nvl(cp16,0)-nvl(cp17,0) C05,cp17 C06,sqldatet(cp05) C07,cp09 C08,cp140,cp05,cp09" & _
      " from caseprogress a,staff s1,staff s2,hirecase,nation,casepropertymap" & _
      " Where (cp18<0 or (cp140 Is Not Null and (cp16 = 0)" & _
      " and cp57 is null and cp20 is null and substr(cp12, 1, 1) <> 'F' and cp09<'C'))" & _
      " and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04 and hc01 is not null" & stCon & _
      " and s1.st01(+)=cp65 and s2.st01(+)=cp13 and na01(+)='000' and cpm01(+)=cp01 and cpm02(+)=cp10"
      
   strSql = strSql & " order by cp05 asc,cp09 asc"
            
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoRecordset
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
   End If
   
Checking:
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then MsgBox Err.Description, , MsgText(5)
   
End Sub
'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
End Sub
'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function
'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
            Screen.MousePointer = vbHourglass
            AdodcRefresh
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub
