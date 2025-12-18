VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm04010508_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人通知修正"
   ClientHeight    =   5340
   ClientLeft      =   -3132
   ClientTop       =   1656
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9348
   Begin VB.Frame Frame1 
      Height          =   684
      Left            =   120
      TabIndex        =   10
      Top             =   552
      Width           =   7695
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(F)"
         Default         =   -1  'True
         Height          =   324
         Left            =   6480
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   1520
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "P"
         Top             =   252
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   5016
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號："
         Height          =   255
         Left            =   192
         TabIndex        =   11
         Top             =   252
         Value           =   -1  'True
         Width           =   1236
      End
      Begin VB.OptionButton Option2 
         Caption         =   "申請案號："
         Height          =   255
         Left            =   3864
         TabIndex        =   4
         Top             =   252
         Width           =   1236
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   1896
         MaxLength       =   6
         TabIndex        =   1
         Top             =   252
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   2
         Left            =   2736
         MaxLength       =   1
         TabIndex        =   2
         Top             =   252
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   3
         Left            =   3096
         MaxLength       =   2
         TabIndex        =   3
         Top             =   252
         Width           =   375
      End
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   8388
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7560
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm04010508_1.frx":0000
      Height          =   3900
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   6879
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
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
      Caption         =   " "
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "NAME0"
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
      BeginProperty Column01 
         DataField       =   "NAME1"
         Caption         =   "案件名稱"
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
         DataField       =   "NAME2"
         Caption         =   "專利種類"
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
         DataField       =   "NA03"
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
         DataField       =   "CU04"
         Caption         =   "申請人"
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
         BeginProperty Column00 
            ColumnWidth     =   1152
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1391.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1152
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   912.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2052.283
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   336
      Left            =   144
      Top             =   1632
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
End
Attribute VB_Name = "frm04010508_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (DataGrid1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim pmain As New ADODB.Recordset
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String, m_AppNo As String
Dim m_Done As Boolean
Dim m_PrevForm As Form
'2016/10/7 END


'Add By Sindy 2022/7/1
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmkok_Click(Index As Integer)
   Select Case Index
      Case 0
        If pmain.BOF And pmain.EOF Then Exit Sub
         strKey1 = pmain.Fields(5).Value
         StrKey2 = pmain.Fields(6).Value
         strKey3 = pmain.Fields(7).Value
         strKey4 = pmain.Fields(8).Value
         'Add By Sindy 2017/12/27
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> strKey1 & StrKey2 & strKey3 & strKey4 Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
               Exit Sub
            End If
            If Not m_PrevForm Is Nothing Then
               Call frm04010508_2.SetParent(m_PrevForm)
            End If
         End If
         '2017/12/27 END
         frm04010508_1.Hide
         'Add By Sindy 2016/10/7
         frm04010508_2.m_strIR01 = m_strIR01
         frm04010508_2.m_strIR02 = m_strIR02
         frm04010508_2.m_strIR03 = m_strIR03
         frm04010508_2.m_strIR04 = m_strIR04
         '2016/10/7 END
         frm04010508_2.Show
         frm04010508_2.QueryData
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Command1_Click()
   If Option1.Value = True Then
      If Text1(2).Text = "" Then Text1(2).Text = "0"
      If Text1(3).Text = "" Then Text1(3).Text = "00"
      
      If Text1(0).Text = "PS" Then
         strExc(0) = "SELECT SP01||SP02||SP03||SP04 FROM SERVICEPRACTICE WHERE SP01='" & Text1(0) & "' AND SP02='" & Text1(1) & "' AND SP03='" & Text1(2) & "' AND SP04='" & Text1(3) & "' "
      ElseIf Text1(0).Text = "P" Then
         strExc(0) = "SELECT PA01||PA02||PA03||PA04 FROM PATENT WHERE PA01='" & Text1(0) & "' AND PA02='" & Text1(1) & "' AND PA03='" & Text1(2) & "' AND PA04='" & Text1(3) & "' "
      End If
      If pmain.State = adStateOpen Then pmain.Close
      pmain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      If pmain.BOF Or pmain.EOF Then
         MsgBox "此案號不存於基本檔中", vbInformation
         Text1(0).SetFocus
         Exit Sub
      End If
            
       'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
        If FMP2open = True Then
           If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text1(0), Text1(1), Text1(2), Text1(3)) = False Then
             Text1(1).SetFocus
             Exit Sub
           End If
         End If
         
      If Text1(0).Text = "PS" Then
         strExc(0) = "SELECT SP01||SP02||SP03||SP04 AS NAME0,NVL(SP05,NVL(SP06,SP07)) AS NAME1," & _
            "'' AS NAME2,NA03,CU04,SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE,CUSTOMER,NATION WHERE " & _
            "SP01='" & Text1(0) & "' AND SP02='" & Text1(1) & "' AND SP03='" & Text1(2) & "' AND SP04='" & Text1(3) & "' AND " & _
            "SP09<>'000' AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02 AND " & _
            "SP09=NA01"
      ElseIf Text1(0).Text = "P" Then
         strExc(0) = "SELECT PA01||PA02||PA03||PA04 AS NAME0,NVL(PA05,NVL(PA06,PA07)) AS NAME1," & _
            "DECODE(PA09,'000',PTM03,PTM04) AS NAME2,NA03,CU04,PA01,PA02,PA03,PA04 FROM " & _
            "PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE " & _
            "PA01='" & Text1(0) & "' AND PA02='" & Text1(1) & "' AND PA03='" & Text1(2) & "' AND PA04='" & Text1(3) & "' AND " & _
            "PA09<>'000' AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02 AND " & _
            "PA09=NA01 AND '1'=PTM01 AND PA08=PTM02"
      End If
      Screen.MousePointer = vbHourglass
      If pmain.State = adStateOpen Then pmain.Close
      pmain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      Screen.MousePointer = vbDefault
      If pmain.EOF Or pmain.BOF Then
         MsgBox "資料庫內無資料", vbInformation
         Text1(0).SetFocus
         Exit Sub
      End If
      cmkok(0).Default = True
      If pmain.RecordCount = 1 Then
         strKey1 = pmain.Fields(5).Value
         StrKey2 = pmain.Fields(6).Value
         strKey3 = pmain.Fields(7).Value
         strKey4 = pmain.Fields(8).Value
         'Add By Sindy 2017/12/27
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> strKey1 & StrKey2 & strKey3 & strKey4 Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
               Exit Sub
            End If
         End If
         '2017/12/27 END
         Me.Hide
         'Add By Sindy 2016/10/7
         frm04010508_2.m_strIR01 = m_strIR01
         frm04010508_2.m_strIR02 = m_strIR02
         frm04010508_2.m_strIR03 = m_strIR03
         frm04010508_2.m_strIR04 = m_strIR04
         '2016/10/7 END
         frm04010508_2.Show
         frm04010508_2.QueryData
      End If
      Set Adodc1.Recordset = pmain
      Adodc1.Recordset.ReQuery
      
   Else
      If pmain.State = adStateOpen Then pmain.Close
      If Text2.Text = "" Then MsgBox "申請案號不可為空值 !", vbInformation: Exit Sub
      strExc(0) = "SELECT PA01||PA02||PA03||PA04 AS NAME0,NVL(PA05,NVL(PA06,PA07)) AS NAME1," & _
         "DECODE(PA09,'000',PTM03,PTM04)AS NAME2,NA03,CU04,PA01,PA02,PA03,PA04 FROM " & _
         "PATENT,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE PA11='" & Text2 & "' AND " & _
         "PA09<>'000' AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02 AND NA01=PA09 AND " & _
         "PTM01='1' AND PTM02=PA08 union all " & _
         "select SP01||SP02||SP03||SP04 AS NAME0,NVL(SP05,NVL(SP06,SP07)) AS NAME1," & _
         "'' AS NAME2,NA03,CU04,SP01,SP02,SP03,SP04 FROM " & _
         "SERVICEPRACTICE,CUSTOMER,NATION,PATENTTRADEMARKMAP WHERE SP11='" & Text2 & "' AND " & _
         "SP09<>'000' AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02 AND NA01=SP09"
      pmain.Open strExc(0), cnnConnection
      If Not pmain.BOF Then pmain.MoveFirst
      If pmain.EOF Or pmain.BOF Then
         MsgBox "資料庫內無此本所案號之資料", vbInformation: Exit Sub
      End If
      cmkok(0).Default = True
      If pmain.RecordCount = 1 Then
       'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
         If FMP2open = True Then
           If PUB_FMPtoCheck(0, 1, Pub_strUserST05, pmain.Fields(5).Value, pmain.Fields(6).Value, pmain.Fields(7).Value, pmain.Fields(8).Value) = False Then
             Text2.SetFocus
             Exit Sub
           End If
         End If
         
         strKey1 = pmain.Fields(5).Value
         StrKey2 = pmain.Fields(6).Value
         strKey3 = pmain.Fields(7).Value
         strKey4 = pmain.Fields(8).Value
         'Add By Sindy 2017/12/27
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> strKey1 & StrKey2 & strKey3 & strKey4 Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
               Exit Sub
            End If
         End If
         '2017/12/27 END
         Me.Hide
         'Add By Sindy 2016/10/7
         frm04010508_2.m_strIR01 = m_strIR01
         frm04010508_2.m_strIR02 = m_strIR02
         frm04010508_2.m_strIR03 = m_strIR03
         frm04010508_2.m_strIR04 = m_strIR04
         '2016/10/7 END
         frm04010508_2.Show
         frm04010508_2.QueryData
      End If
      Set Adodc1.Recordset = pmain
      Adodc1.Recordset.ReQuery
   End If
End Sub

Private Sub Form_Activate()
   Command1.Default = True
   If strKey1 = "1" Then
      'Modify By Cheng 2002/07/19
'      Text1(0).Text = ""
      Text1(0).Text = "P"
      Text1(1).Text = ""
      Text1(2).Text = ""
      Text1(3).Text = ""
      Text2.Text = ""
      If pmain.State = adStateOpen Then pmain.Close
      strExc(0) = "SELECT PA01||PA02||PA03||PA04 FROM PATENT WHERE PA01='" & Text1(0) & "' AND PA02='" & Text1(1) & "' AND PA03='" & Text1(2) & "' AND PA04='" & Text1(3) & "' "
      pmain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
      Set Adodc1.Recordset = pmain
      Adodc1.Recordset.ReQuery
      pmain.Close
   End If
   'Added by Sindy 2016/10/7
   If m_strIR01 <> "" And m_Done = False Then
      If Text2 <> "" Then
         Text2.Text = m_AppNo
         Option2.Value = True
      Else
         Text1(0).Text = m_strCP01
         Text1(1).Text = m_strCP02
         Text1(2).Text = m_strCP03
         Text1(3).Text = m_strCP04
         Option1.Value = True
      End If
      Command1.Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2016/10/7 END
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   pmain.CursorLocation = adUseClient
   Text2.Enabled = False
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2022/7/1
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2022/7/1 END
   
   Set frm04010508_1 = Nothing
End Sub

Private Sub Option1_Click()
   Text1(0).Enabled = True
   Text1(1).Enabled = True
   Text1(2).Enabled = True
   Text1(3).Enabled = True
   Text2.Enabled = False
   Text1(0).SetFocus
End Sub

Private Sub Option2_Click()
   Text2.Enabled = True
   Text1(0).Enabled = False
   Text1(1).Enabled = False
   Text1(2).Enabled = False
   Text1(3).Enabled = False
   Text2.SetFocus
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
        If Text1(0).Text <> "P" And Text1(0).Text <> "PS" Then
           MsgBox "只可為P及PS案件", vbInformation
           TextInverse Text1(Index)
           Cancel = True
        End If
   End Select
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
