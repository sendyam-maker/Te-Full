VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2153 
   AutoRedraw      =   -1  'True
   Caption         =   "帳單審核作業"
   ClientHeight    =   6660
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   12720
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "點我展開"
      Height          =   345
      Left            =   8784
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   36
      Width           =   4515
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5412
      Left            =   8784
      TabIndex        =   8
      Top             =   372
      Width           =   4512
      ExtentX         =   7964
      ExtentY         =   9551
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2153.frx":0000
      Height          =   3180
      Left            =   225
      TabIndex        =   5
      Top             =   375
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   5609
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "帳單資料"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "a1521"
         Caption         =   "可否結匯"
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
         DataField       =   "flg"
         Caption         =   "退"
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
         DataField       =   "a1525"
         Caption         =   "紙本會簽"
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
         DataField       =   "a1526"
         Caption         =   "水單請款"
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
         DataField       =   "a1501"
         Caption         =   "帳單編號"
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
         DataField       =   "a1502"
         Caption         =   "帳單日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "a1503"
         Caption         =   "代理人"
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
         DataField       =   "a1504"
         Caption         =   "代理人D/N編號"
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
         DataField       =   "a1505"
         Caption         =   "幣別"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "a1506"
         Caption         =   "帳單金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "a1509"
         Caption         =   "備註"
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
      BeginProperty Column11 
         DataField       =   "a1507"
         Caption         =   "作廢日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###/##/##"
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
            ColumnWidth     =   984.189
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   252.283
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   288
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   299.906
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1332.284
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1535.811
         EndProperty
         BeginProperty Column10 
            Locked          =   -1  'True
            ColumnWidth     =   4524.095
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1235.906
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1035
      TabIndex        =   0
      Top             =   30
      Width           =   3795
   End
   Begin VB.CommandButton Command3 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7740
      TabIndex        =   3
      Top             =   60
      Width           =   768
   End
   Begin VB.CommandButton Command1 
      Caption         =   "顯示已審核"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6330
      TabIndex        =   2
      Top             =   60
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "顯示未審核"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4920
      TabIndex        =   1
      Top             =   60
      Width           =   1320
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   252
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   2138
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "退回意見"
      Height          =   888
      Left            =   1656
      TabIndex        =   9
      Top             =   4932
      Visible         =   0   'False
      Width           =   5784
      Begin MSForms.TextBox txtA1524 
         Height          =   675
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Width           =   5610
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "9895;1191"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   240
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FFFF&
      Caption         =   "退回"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   216
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   5040
      Width           =   684
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FFFF&
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   936
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   684
   End
   Begin VB.CommandButton Command5 
      Caption         =   "符合寬度"
      Height          =   372
      Left            =   7560
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   984
   End
   Begin MSForms.TextBox Text1 
      Height          =   765
      Left            =   1620
      TabIndex        =   15
      Top             =   5850
      Width           =   6915
      VariousPropertyBits=   -1467989985
      ScrollBars      =   2
      Size            =   "12197;1349"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   240
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "帳單備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   255
      TabIndex        =   14
      Top             =   5850
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "系統別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   255
      TabIndex        =   4
      Top             =   45
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   12
      Top             =   3096
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc2153"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/12 改成Form2.0 (DataGrid1,GrdDataList,txtA1524,Text1)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoadodc1 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset
Dim intSelect As Integer
Dim strTemp1 As Variant
Dim strTemp2 As Variant
Dim i As Integer
Dim j As Integer
Dim s As Integer
Dim strSystemKind As String
Dim strSql As String
Dim mstrGroup As String 'Add by Lydia 設群組權限
'Added by Morgan 2018/1/26
Dim bolOpenPdf As Boolean, bolClick As Boolean, bolDblClick As Boolean, bolPaint As Boolean, bolSenKey As Boolean
Dim m_AttachPath As String, strReadList As String, strPreClickNo As String


'Added by Morgan 2018/1/26 從DataGrid1_Click事件移來(Click事件會抓到前一筆，因觸發時間早於資料移動)
Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If bolClick Then
      Adodc2Refresh
      bolClick = False
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   If Index = 1 Then
      If cmdok(1).Caption = "確定" Then
         If txtA1524 = "" Or txtA1524 = "請輸入退回意見!!" Then
            MsgBox "請輸入退回意見!!", vbExclamation
            txtA1524.SetFocus
            Exit Sub
         End If
         FormSave
      ElseIf InStr(strReadList, Adodc1.Recordset("a1501")) = 0 Then
         MsgBox "尚未讀取帳單內容不可退回！" & vbCrLf & vbCrLf & "(雙擊該列帳單可載入PDF檔)", vbInformation
      Else
         SetA1524 True
      End If
   '取消
   ElseIf Index = 2 Then
      SetA1524
   End If
End Sub

Private Sub Command1_Click()
   If MsgBox("顯示已審核資料需要執行很長的時間，請問還要繼續嗎？", vbOKCancel) = vbCancel Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   intSelect = 2
   AdodcRefresh
   Adodc2Refresh
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
   Screen.MousePointer = vbHourglass
   intSelect = 1
   AdodcRefresh
   Adodc2Refresh
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()

   strExc(1) = ""
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   Else
      'Add By Sindy 2010/8/19
      Adodc1.Recordset.MoveFirst
      Do While Adodc1.Recordset.EOF = False
         If Adodc1.Recordset.Fields("a1521") = "" Or IsNull(Adodc1.Recordset.Fields("a1521")) Then
            MsgBox "可否結匯欄位不可空白!!!"
            Exit Sub
         End If
         'Added by Lydia 2021/06/22 78011葉易雲在操作按確定時，要逐筆發MAIL給江協理98020
         If Adodc1.Recordset.Fields("a1521") = "Y" And strUserNum = "78011" Then
            strExc(1) = strExc(1) & ",'" & Adodc1.Recordset.Fields("a1501") & "' "
         End If
         'end 2021/06/22
         Adodc1.Recordset.MoveNext
      Loop
      '2010/8/19 End
   End If
   'Added by Lydia 2021/06/22 78011葉易雲在操作按確定時，要逐筆發MAIL給江協理98020
   If strExc(1) <> "" Then
      strExc(1) = Mid(strExc(1), 2)
      strExc(0) = "select axf03,a1501,a1502,a1503||' '||nvl(fa05,nvl(fa04,fa06)) as faname,a1504,a1505,a1506,axf14,a1509 " & _
                       "From Acc150, Acc151, Fagent where a1501 in (" & strExc(1) & ") and a1501=axf01(+) and substr(a1503,1,8)=fa01(+) and substr(a1503,9,1)=fa02(+) " & _
                       "order by a1501 asc "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
          strExc(2) = Format(ServerTime, "000000")
          RsTemp.MoveFirst
          Do While Not RsTemp.EOF
               strExc(3) = "本所案號：" & RsTemp.Fields("axf03") & vbCrLf & _
                                "帳單編號：" & RsTemp.Fields("a1501") & vbCrLf & _
                                "帳單日期：" & ChangeTStringToTDateString(RsTemp.Fields("a1502")) & vbCrLf & _
                                "代理人：" & RsTemp.Fields("faname") & vbCrLf & _
                                "代理人D/N編號：" & RsTemp.Fields("a1504") & vbCrLf & _
                                "幣別：" & RsTemp.Fields("a1505") & vbCrLf & _
                                "帳單金額：" & RsTemp.Fields("a1506") & vbCrLf & _
                                "盈虧：" & RsTemp.Fields("axf14") & vbCrLf & _
                                "備註：" & RsTemp.Fields("a1509")
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                   " values( '" & strUserNum & "','98020', '" & strSrvDate(1) & "', '" & strExc(2) & "','CFT帳單主管己審核確認通知！','" & strExc(3) & "',null)"
               cnnConnection.Execute strSql
               strExc(2) = Format(Val(strExc(2)) + 1, "000000")
               RsTemp.MoveNext
          Loop
      End If
   End If
   'end 2021/06/22
   Screen.MousePointer = vbHourglass
   Adodc1.Recordset.UpdateBatch
   PUB_SendMailCache 'Added by Morgan 2018/2/6
   AdodcRefresh
   Adodc2Refresh
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command4_Click()
   If Command4.Caption = "點我展開" Then
      RePosForm True
   Else
      RePosForm False
   End If
End Sub

Private Sub Command5_Click()
   WebBrowser1.SetFocus
   DoEvents
   SendKeys "^2"
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   'If IsNull(Adodc1.Recordset.Fields("a1521").Value) Or Adodc1.Recordset.Fields("a1521").Value <> MsgText(602) Then
   '   Adodc1.Recordset.Fields("a1521").Value = MsgText(603)
   'End If
   Adodc1.Recordset.UPDATE
Checking:
   Exit Sub
End Sub

Private Sub DataGrid1_Click()
  
   'Added by Morgan 2018/1/25
   'DblClick會觸發2次Clcik,控制第2次不動作
   If bolDblClick Then
      bolDblClick = False
      Exit Sub
   End If
   'end 2018/1/25
   
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   If DataGrid1.col = 0 Then
      'Modified by Morgan 2018/2/1 此時讀到值的是前次資料非目前點選列，改在 KeyPress 事件控制
      'If DataGrid1.Columns(0).Text = MsgText(603) Then
      '   SendKeys "{BACKSPACE}"
      '   SendKeys "{Y}"
      '   SendKeys "{ENTER}"
      'Else
      '   SendKeys "{BACKSPACE}"
      '   SendKeys "{N}"
      '   SendKeys "{ENTER}"
      'End If
      bolSenKey = True
      strPreClickNo = Adodc1.Recordset("a1501")
      'Modified by Morgan 2019/4/17 郭DblClick後再Click會發生KeyDown事件不觸發情形,此時欄位內會是空白,故改為送Enter
      'SendKeys "{BACKSPACE}"
      SendKeys "{ENTER}"
      'end 2019/4/17
      'end 2018/2/1
         
      'Pub_WriteSysLog "(" & strUserNum & ") -> GC: col=" & DataGrid1.col 'debug

   End If

   'Modified by Morgan 2018/2/1 移到 Adodc1_MoveComplete
   'Adodc2Refresh
   bolClick = True
   'end 2018/2/1
   
Checking:
   Exit Sub
End Sub
'Added by Morgan 2018/1/25
Private Sub DataGrid1_DblClick()
   bolDblClick = True
   If bolOpenPdf Then OpenPdf
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'Removed by Morgan 2019/4/17 改KeyUp
'   'Modified by Morgan 2019/4/17 修正是否可結匯變空白後無法點選為Y/N問題
'   Pub_WriteSysLog "(" & strUserNum & ") -> KD1: col=" & DataGrid1.col & ", KeyCode=" & KeyCode & ",SendKey=" & bolSenKey
'   If DataGrid1.col = 0 Then
'      If bolSenKey Then
'         bolSenKey = False
'
'         If strPreClickNo = Adodc1.Recordset("a1501") Then
'            If DataGrid1.Text = "N" Then
'               DataGrid1.Text = "Y"
'
'            ElseIf DataGrid1.Text = "Y" Then
'               DataGrid1.Text = "N"
'
'            End If
'         End If
'
'      ElseIf KeyCode = 78 Then
'         DataGrid1.Text = "N"
'
'      ElseIf KeyCode = 89 Then
'         DataGrid1.Text = "Y"
'
'      End If
'   End If
'   KeyCode = 0
'   Pub_WriteSysLog "(" & strUserNum & ") -> KD2: col=" & DataGrid1.col & ", KeyCode=" & KeyCode

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   'Modified by Morgan 2019/4/17
   'Pub_WriteSysLog "(" & strUserNum & ") -> KU1: col=" & DataGrid1.col & ", KeyCode=" & KeyCode & ",SendKey=" & bolSenKey
   If DataGrid1.col = 0 Then
      If bolSenKey Then
         bolSenKey = False
         
         If strPreClickNo = Adodc1.Recordset("a1501") Then
            'Modified by Morgan 2019/6/25 R不可直接上Y(程序處理完會變N)
            'If DataGrid1.Text = "N" Or DataGrid1.Text = "R" Then
            If DataGrid1.Text = "N" Then
               DataGrid1.Text = "Y"
               
            ElseIf DataGrid1.Text = "Y" Then
               'Modified by Morgan 2019/6/25 R不可直接上Y(程序處理完會變N)
               'If txtA1524 <> "" And Frame2.Visible = True Then
               '   DataGrid1.Text = "R"
               'Else
                  DataGrid1.Text = "N"
               'End If
               'end 2019/6/25
            End If
         End If
         
      ElseIf KeyCode = 78 Then
         DataGrid1.Text = "N"
      
      ElseIf KeyCode = 89 Then
         DataGrid1.Text = "Y"
         
      End If
   End If
   KeyCode = 0
   'Pub_WriteSysLog "(" & strUserNum & ") -> KU2: col=" & DataGrid1.col & ", KeyCode=" & KeyCode & ",SendKey=" & bolSenKey
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   Adodc2Refresh
End Sub

Private Sub Form_Activate()
   If Screen.ActiveForm.Name <> Me.Name Then Exit Sub 'Added by Morgan 2025/8/12

'edit by nickc 2007/02/08
'   '93.3.16 ADD BY SONIA
'   If IsObject(mdiMain) Then
'      mdiMain.toolshow
'   End If
'   '93.3.16 END
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
   
   'Modify By Sindy 2021/1/18 + Or Pub_StrUserSt03 = "P20"
   If Pub_StrUserSt03 = "P12" Or Pub_StrUserSt03 = "M51" Or _
      Pub_StrUserSt03 = "P20" Then
      bolOpenPdf = True
      Me.WindowState = 2
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub PaintForm()
   Dim intX As Integer
   Dim intY As Integer
   Dim sglWidth As Single
   Dim sglHeight As Single
      
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
End Sub

Private Sub Form_Load()
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 5400
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   
   'Added by Morgan 2018/1/25
   bolPaint = True
   PaintForm '舊程式寫成 Sub 共用
   m_AttachPath = App.path & "\" & strUserNum
   KillTemp
   'end 2018/1/25
   
   intSelect = 1
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
   If FMP2open = True And (UCase(App.EXEName) = "PATPRO" Or UCase(App.EXEName) = "TEPATPRO") Then
     Text12 = "P,PS,"
   Else
     Text12 = GetSystemKindByNick
   End If
   mstrGroup = Text12
   'add by sonia 2022/9/5 有CFT權限者，帳單審核再加開CFC，否則審核主管沒有CFC操作權限就不能審核
   If InStr(mstrGroup, "CFT") > 0 And InStr(mstrGroup, "CFC") = 0 Then
      mstrGroup = mstrGroup & "CFC,"
      Text12 = mstrGroup
   End If
   'end 2022/9/5
   'strTemp1 = Split(UCase(GetSystemKindByNick), ",")
    strTemp1 = Split(UCase(mstrGroup), ",")
   strTemp2 = Split(UCase(Text12), ",")
   AdodcRefresh
   Adodc2Refresh
'cancel by sonia 2017/8/22 桂所長可按確定,其他等級01的主管不行,故開桂所長個人權限,01等級由form之fo04設定不可使用
'   '2015/10/21 add by sonia 加入總經理權限(等級01),可使用所有程式,但維護程式只有查詢功能,不可新增刪除修改)
'  If PUB_GetST05(strUserNum) = "01" Then
'      Command3.Enabled = False
'   End If
'   '2015/10/21 end
'end 2017/8/22
   WebBrowser1.Navigate "about:blank"
   
   'Added by Morgan 2019/3/12
   'Modify By Sindy 2021/1/18 + Or Pub_StrUserSt03 = "P20"
   If Not (Pub_StrUserSt03 = "P12" Or Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "P20") Then
      Me.DataGrid1.Columns(1).Width = 0
      Me.DataGrid1.Columns(2).Width = 0
   End If
   'end 2019/3/12
End Sub

Private Sub Form_Resize()
   If bolPaint Then '控制圖片載入後
      PaintForm
   End If
   
   If Me.WindowState = 0 Or Me.WindowState = 2 Then
      If Command4.Caption = "點我展開" Then
         RePosForm False
      Else
         RePosForm True
      End If
   End If
End Sub

Private Sub RePosForm(pFull As Boolean)
   Static lngLeft As Long
   Dim a1 As Integer
   If Forms(0).WindowState <> 1 Then
      If lngLeft = 0 Then lngLeft = Command4.Left
      If pFull = True Then
         WebBrowser1.Left = 0
         WebBrowser1.Width = Me.Width - 90
         WebBrowser1.Height = Me.Height - Command4.Height - 390
         Command4.Caption = "點我還原"
      Else
         WebBrowser1.Left = lngLeft
         Command4.Caption = "點我展開"
      End If
      If Me.Width > 90 + WebBrowser1.Left Then
         WebBrowser1.Width = Me.Width - 90 - WebBrowser1.Left
      End If
      If Me.Height > Command4.Height + 390 Then
         WebBrowser1.Height = Me.Height - Command4.Height - 390
      End If
      Command4.Left = WebBrowser1.Left
      Command4.Width = WebBrowser1.Width
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc2153 = Nothing
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  資料表更新(主檔)
'
'*************************************************
Private Sub AdodcRefresh(Optional pNo As String)
   strReadList = "" 'Added by Morgan 2018/2/1
   strSql = ""
   strSystemKind = ""
   For i = 0 To UBound(strTemp2)
      strSystemKind = strSystemKind & "'" & strTemp2(i) & "',"
   Next i
   If strSystemKind <> "" Then
      strSystemKind = Mid(strSystemKind, 1, Len(strSystemKind) - 1)
      'strSQL = strSQL & " and substr(axf03, 1, length(axf03) - 9) in (" & strSystemKind & ")"
      strSql = strSql & " and cp01 in (" & strSystemKind & ")"
   Else
      strSql = strSql & " and cp01 = 'Z'"
   End If
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
'   Select Case intSelect
'      Case 2
'         '93.12.9 modify by sonia 作廢不顯示
'         '2010/8/4 modify by sonia A1521改為upper(A1521)
'         '2011/8/17 MODIFY BY SONIA 已結匯或已抵帳不再出現 '602=Y
'         adoadodc1.Open "select * from acc150 where a1512 is null and (a1520 is null or a1520=0) and a1507 is null and upper(a1521) = '" & MsgText(602) & "' and a1501 in (select distinct axf01 from acc150, caseprogress, acc151 where a1501 = axf01 (+) and axf02 = cp09 (+) and upper(A1521) = '" & MsgText(602) & "' " & strSql & ") order by a1501 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'         '93.12.9 end
'      Case Else
'         '93.12.9 modify by sonia 作廢不顯示
'         '2011/8/17 MODIFY BY SONIA 已結匯或已抵帳不再出現 '603=N
'         adoadodc1.Open "select * from acc150 where a1512 is null and (a1520 is null or a1520=0) and a1507 is null and upper(a1521) = '" & MsgText(603) & "' and a1501 in (select distinct axf01 from acc150, caseprogress, acc151 where a1501 = axf01 (+) and axf02 = cp09 (+) and upper(A1521) = '" & MsgText(603) & "' " & strSql & ") order by a1501 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'         '93.12.9 end
'   End Select

'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
'設別名f0,+FMP2openSQL
   Dim midSql As String
   
   'Modified by Morgan 2016/7/29
   '因增加待審核(W)改判斷不是已審核(Y)
   'If intSelect = 2 Then
   '   strExc(0) = MsgText(602)
   'Else
   '   strExc(0) = MsgText(603)
   'End If
   'strExc(1) = " select * from acc150 where a1512 is null and (a1520 is null or a1520=0) and a1507 is null " & _
               " and upper(a1521)='" & strExc(0) & "' and a1501 in " & _
               "(select distinct axf01 from acc150, caseprogress f0, acc151 where a1501=axf01(+) and axf02=f0.cp09(+) " & _
               " and upper(A1521)='" & strExc(0) & "' " & strSql & FMP2openSQL & ") order by a1501 asc "
   If intSelect = 2 Then
      strExc(0) = " and a1521='Y' "
   Else
      strExc(0) = " and a1521<>'Y' "
   End If
   'Modified by Morgan 2023/3/8 郭雅娟不需審核寰華案，改以帳單輸入人員過濾
   'strExc(1) = " select a.*,b.*,decode(a1524,'','','退') flg from acc150 a, acc152 b where a1512 is null and (a1520 is null or a1520=0) and a1507 is null " & strExc(0) & _
               " and a1501 in (select distinct axf01 from acc150, caseprogress f0, acc151 where a1501=axf01(+) and axf02=f0.cp09(+) " & _
               strExc(0) & strSql & FMP2openSQL & ") and ayf01(+)=a1501 order by a1501 asc "
   'Modified by Morgan 2023/8/30 st01=a1519->st01=nvl(a1519,a1516)
   If FMP2open Then
      strExc(2) = " and exists(select * from staff where st01=nvl(a1519,a1516) and st03 like 'F2%')"
   Else
      strExc(2) = " and not exists(select * from staff where st01=nvl(a1519,a1516) and st03 like 'F2%')"
   End If
   strExc(1) = " select a.*,b.*,decode(a1524,'','','退') flg from acc150 a, acc152 b" & _
      " where a1512 is null and (a1520 is null or a1520=0) and a1507 is null " & strExc(0) & _
      " and ayf01(+)=a1501 " & strExc(2)
   If strSql <> "" Then
      strExc(1) = strExc(1) & " and exists(select * from acc151,caseprogress where axf01=a1501 and cp09(+)=axf02" & strSql & ")"
   End If
   strExc(1) = strExc(1) & " order by a1501 asc "
   'end 2023/3/8
   'end 2016/7/29
   
   adoadodc1.Open strExc(1), adoTaie, adOpenDynamic, adLockBatchOptimistic
   Set Adodc1.Recordset = adoadodc1
   Adodc1.Recordset.ReQuery
   
   'Added by Morgan 2018/1/26
   WebBrowser1.Navigate "about:blank"
   DoEvents
   DataGrid1.Caption = "帳單資料 (" & Adodc1.Recordset.RecordCount & ")"
   If Adodc1.Recordset.RecordCount > 0 And pNo <> "" Then
      Adodc1.Recordset.Find "a1501='" & pNo & "'"
      If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
   End If
   SetA1524
   'end 2018/1/26
   
End Sub

'*************************************************
'  資料表更新(明細檔)
'
'*************************************************
Private Sub Adodc2Refresh()
   Dim m_ProcessProfit As String, strTWAmount As Double
   cmdok(1).Visible = False 'Added by Morgan 2018/2/1
   Frame2.Visible = False 'Added by Morgan 2018/2/1
   'Screen.MousePointer = vbHourglass
   grdDataList.Rows = 2
   grdDataList.Clear
   SetDataListWidth
   If Adodc1.Recordset.RecordCount = 0 Then
'      Set Adodc2.Recordset = adoadodc1
'      Adodc2.Recordset.ReQuery
      Exit Sub
   End If
   
   Text1 = "" & Adodc1.Recordset.Fields("a1509").Value 'Added by Morgan 2019/3/15
   
'   If adoadodc2.State = adStateOpen Then
'      adoadodc2.Close
'   End If
'   adoadodc2.CursorLocation = adUseClient
'   adoadodc2.Open "select * from acc151, caseprogress, casepropertymap where axf02 = cp09 and cp01 = cpm01 and cp10 = cpm02 and axf01 = '" & Adodc1.Recordset.Fields("a1501").Value & "' order by axf02 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Set Adodc2.Recordset = adoadodc2
'   Adodc2.Recordset.ReQuery
   'Modify By Sindy 2010/8/19
   'Modified by Morgan 2022/3/23
   'strSql = "select axf03 as 本所案號,axf02 as 總收文號,cpm03 as 案件性質,axf04 as 帳單金額,axf14 as 盈虧,' ' as 收文號盈虧,axf12 as 案件名稱,axf13 as 收據抬頭,CP01 系統別 from acc151, caseprogress, casepropertymap where axf02 = cp09 and cp01 = cpm01 and cp10 = cpm02 and axf01 = '" & Adodc1.Recordset.Fields("a1501").Value & "' order by axf02 asc"
   strSql = "select axf03 as 本所案號,axf02 as 總收文號,GETCP10DESCCaseNO(cp10,cp01,cp02,cp03,cp04) as 案件性質,axf04 as 帳單金額,axf14 as 盈虧,' ' as 收文號盈虧,axf12 as 案件名稱,axf13 as 收據抬頭,CP01 系統別 from acc151, caseprogress where axf02 = cp09 and axf01 = '" & Adodc1.Recordset.Fields("a1501").Value & "' order by axf02 asc"
   'end 2022/3/23
   CheckOC
   'Screen.MousePointer = vbHourglass
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   If adoRecordset.RecordCount <> 0 Then
      Set grdDataList.Recordset = adoRecordset
      For i = 1 To grdDataList.Rows - 1
         If CalculateProcessProfit(grdDataList.TextMatrix(i, 1), m_ProcessProfit, 0, strTWAmount, "") Then
            grdDataList.TextMatrix(i, 5) = m_ProcessProfit
         End If
      Next i
      'Added by Morgan 2018/2/1
      'Modified by Morgan 2018/10/18 +CFP,CPS
      'Modified by Sindy 2021/1/19 + Left(adoRecordset("系統別"), 1) = "T"
      If adoRecordset("系統別") = "P" Or adoRecordset("系統別") = "PS" Or adoRecordset("系統別") = "CFP" Or adoRecordset("系統別") = "CPS" Or _
         Left(adoRecordset("系統別"), 1) = "T" Then
         Command5.Visible = True
         SetA1524
      Else
         Command5.Visible = False
      End If
      'end 2018/2/1
   Else
'       CheckOC
'       ShowNoData
   End If
   'Screen.MousePointer = vbDefault
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
     'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件(P)，但非此類案件時外專程序人員不可操作。
    'strTemp1 = Split(UCase(GetSystemKindByNick), ",")
     strTemp1 = Split(UCase(mstrGroup), ",")
     strTemp2 = Split(UCase(Text12), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            Cancel = True
            Text12.SetFocus
            Exit Sub
        End If
     Next i
End Sub

'Add By Sindy 2010/8/19
Private Sub SetDataListWidth()
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "本所案號 "
grdDataList.ColWidth(0) = 1400
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "總收文號 "
grdDataList.ColWidth(1) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "案件性質 "
grdDataList.ColWidth(2) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "帳單金額"
grdDataList.ColWidth(3) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "盈虧"
grdDataList.ColWidth(4) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "收文號盈虧"
grdDataList.ColWidth(5) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(6) = 2000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "收據抬頭"
grdDataList.ColWidth(7) = 2000
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Private Function OpenPdf() As Boolean
   Dim stFileName As String
   Dim stA1501 As String, stAyf02 As String
   Dim stMsg As String
   
   WebBrowser1.Navigate "about:blank"
   DoEvents
   stA1501 = "" & Adodc1.Recordset("a1501")
   stAyf02 = "" & Adodc1.Recordset("ayf02")
   If stAyf02 <> "" Then
      If PUB_GetAttachFile_Invoice(stA1501, stAyf02, m_AttachPath, stFileName) = True Then
         WebBrowser1.Navigate m_AttachPath & "\" & stFileName
         strReadList = strReadList & stA1501 & ";"
         OpenPdf = True
         
         'Added by Morgan 2019/3/12
         stMsg = ""
         If Adodc1.Recordset("a1525") = "Y" Then
            stMsg = "【紙本會簽】"
         End If
         'Added by Morgan 2019/6/25
         If Adodc1.Recordset("a1526") = "Y" Then
            stMsg = stMsg & "【憑水單請款】"
         End If
         'end 2019/6/25
         If stMsg <> "" Then
            MsgBox "【" & stA1501 & "】為" & stMsg & "！", vbExclamation
         End If
         'end 2019/3/12
      End If
   Else
      MsgBox "【" & stA1501 & "】電子檔尚未上傳無法預覽！", vbCritical
   End If
   
End Function

Private Sub KillTemp()
On Error GoTo ErrHnd
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
   Exit Sub
   
ErrHnd:
   Resume Next
End Sub

Private Sub FormSave()
   Screen.MousePointer = vbHourglass
   'strSql = "update acc150 set a1521='R',a1524='" & ChgSQL(txtA1524) & "' where a1501='" & Adodc1.Recordset("a1501") & "'"
   'cnnConnection.Execute strSql, intI
   'AdodcRefresh Adodc1.Recordset("a1501")
   'Adodc2Refresh
   Adodc1.Recordset("a1521") = "R"
   Adodc1.Recordset("a1524") = txtA1524
   SetA1524
   Screen.MousePointer = vbDefault
End Sub

Private Sub txtA1524_GotFocus()
   OpenIme
End Sub

Private Sub txtA1524_LostFocus()
   CloseIme
End Sub

'帶出退回意見
Private Sub SetA1524(Optional bolReturn As Boolean = False)
   
   If Adodc1.Recordset.RecordCount = 0 Then
      Frame2.Visible = False
      cmdok(2).Visible = False
      cmdok(1).Visible = False
      Exit Sub
   ElseIf Not bolReturn Then
      txtA1524 = "" & Adodc1.Recordset("a1524")
   End If
   
   If bolReturn Then
      Frame2.Visible = True
      cmdok(1).Caption = "確定"
      If txtA1524 = "" Then
         txtA1524 = "請輸入退回意見!!"
      End If
      txtA1524.SetFocus
      txtA1524.SelStart = 0
      txtA1524.SelLength = Len(txtA1524)
      txtA1524.SetFocus
      txtA1524.Locked = False
      cmdok(2).Visible = True
      DataGrid1.Enabled = False
   Else
      If txtA1524 <> "" Then
         Frame2.Visible = True
      Else
         Frame2.Visible = False
      End If
      cmdok(1).Caption = "退回"
      txtA1524.Locked = True
      cmdok(2).Visible = False
      DataGrid1.Enabled = True
      If Adodc1.Recordset("a1521") = "N" Then
         cmdok(1).Visible = True
      Else
         cmdok(1).Visible = False
      End If
   End If
End Sub
