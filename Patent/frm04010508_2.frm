VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010508_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人通知修正"
   ClientHeight    =   5748
   ClientLeft      =   60
   ClientTop       =   972
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9336
   Begin VB.CommandButton cmkok 
      Caption         =   "結束(&X)"
      Default         =   -1  'True
      Height          =   405
      Index           =   2
      Left            =   8436
      TabIndex        =   13
      Top             =   70
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm04010508_2.frx":0000
      Height          =   3615
      Left            =   180
      TabIndex        =   11
      Top             =   1980
      Width           =   9015
      _ExtentX        =   15896
      _ExtentY        =   6371
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "cp09"
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
      BeginProperty Column01 
         DataField       =   "name1"
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
      BeginProperty Column02 
         DataField       =   "cpm04"
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
         DataField       =   "name2"
         Caption         =   "發文日"
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
         DataField       =   "name3"
         Caption         =   "代理人收達日"
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
         DataField       =   "cp24"
         Caption         =   "結果"
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
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1152
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1188.284
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1091.906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmkok 
      Caption         =   "回前畫面(&U)"
      Height          =   405
      Index           =   1
      Left            =   7212
      TabIndex        =   10
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmkok 
      Caption         =   " 確定(&O)"
      Height          =   405
      Index           =   0
      Left            =   6384
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   180
      Top             =   1920
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
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1380
      TabIndex        =   12
      Top             =   1260
      Width           =   4695
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "8281;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      Caption         =   "專利號數："
      Height          =   255
      Left            =   3900
      TabIndex        =   8
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   4980
      TabIndex        =   7
      Top             =   660
      Width           =   2535
   End
   Begin MSForms.Label Label8 
      Height          =   255
      Left            =   1140
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Caption         =   "申請人："
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "案件名稱 ："
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   1260
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號："
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1260
      TabIndex        =   1
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   660
      Width           =   975
   End
End
Attribute VB_Name = "frm04010508_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Combo1,Label8)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim pmain As New ADODB.Recordset
Dim pmain1 As New ADODB.Recordset
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2016/10/7 END


'Add By Sindy 2022/7/1
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmkok_Click(Index As Integer)
   Select Case Index
      Case 0
         strKey5 = pmain.Fields(0).Value
         strKey6 = "" & pmain1.Fields(2).Value 'Modified by Morgan 2015/12/22 有可能沒有中文案件名稱
         If IsNull(pmain1.Fields(3).Value) Then
         strKey7 = ""
         Else
         strKey7 = pmain1.Fields(3).Value
         End If
         If IsNull(pmain1.Fields(4).Value) Then
         strKey8 = ""
         Else
         strKey8 = pmain1.Fields(4).Value
         End If
         
         'Added by Morgan 2021/12/20
         '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
         If PUB_CheckFormExist("frm04010508_3") = False Then
            Set frm04010508_3 = Nothing
         End If
         'end 2021/12/20
         
         If m_strIR01 <> "" Then
            If Not m_PrevForm Is Nothing Then
               Call frm04010508_3.SetParent(m_PrevForm)
            End If
            'Add By Sindy 2016/10/7
            frm04010508_3.m_strIR01 = m_strIR01
            frm04010508_3.m_strIR02 = m_strIR02
            frm04010508_3.m_strIR03 = m_strIR03
            frm04010508_3.m_strIR04 = m_strIR04
            '2016/10/7 END
         End If
         frm04010508_3.Show
         frm04010508_2.Hide
      Case 1
         frm04010508_1.Show
         Unload Me
      Case 2
         Unload frm04010508_1
         Unload Me
End Select
End Sub

Public Sub QueryData()
   Combo1.Clear
   If strKey1 = "P" Then
      strExc(1) = "select pa22,pa11,nvl(pa05,''),nvl(pa06,''),nvl(pa07,''),nvl(cu04,''),pa09 from patent,customer where pa01='" & strKey1 & "' and pa02='" & StrKey2 & "' and pa03='" & strKey3 & "' and pa04='" & strKey4 & "' and SUBSTR(pa26,1,8)=cu01(+) AND SUBSTR(PA26,9,1)=cu02"
   ElseIf strKey1 = "PS" Then
      strExc(1) = "select sp14,sp11,nvl(sp05,''),nvl(sp06,''),nvl(sp07,''),nvl(cu04,''),sp09 from servicepractice,customer where sp01='" & strKey1 & "' and sp02='" & StrKey2 & "' and sp03='" & strKey3 & "' and sp04='" & strKey4 & "' and SUBSTR(SP08,1,8)=cu01(+) AND SUBSTR(SP08,9,1)=cu02"
   End If
   If pmain1.State = adStateOpen Then pmain1.Close
   pmain1.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
   If pmain1.EOF Or pmain1.BOF Then
      MsgBox "沒有符合條件的資料", vbOKOnly + vbCritical, "查詢資料"
      frm04010508_1.Show
      Unload Me
      GoTo EXITSUB
   End If
   
   If Not IsNull(pmain1.Fields(0).Value) Then Label10.Caption = pmain1.Fields(0).Value
   If Not IsNull(pmain1.Fields(1).Value) Then Label4.Caption = pmain1.Fields(1).Value
   If Not IsNull(pmain1.Fields(5).Value) Then Label8.Caption = pmain1.Fields(5).Value
   If IsNull(pmain1.Fields(2).Value) Then
      Combo1.AddItem "中: ", 0
      Combo1.Text = "中: "
      Else
      Combo1.AddItem "中: " + pmain1.Fields(2).Value, 0
      Combo1.Text = "中: " + pmain1.Fields(2).Value
      End If
      If IsNull(pmain1.Fields(3).Value) Then
      Combo1.AddItem "英: ", 1
      Else
      Combo1.AddItem "英: " + pmain1.Fields(3).Value, 1
      End If
      If IsNull(pmain1.Fields(4).Value) Then
      Combo1.AddItem "日: ", 2
      Else
      Combo1.AddItem "日: " + pmain1.Fields(4).Value, 2
      End If
   Label2.Caption = strKey1 + "-" + StrKey2 + "-" + strKey3 + "-" + strKey4
   'strExc(0) = "select cp09,cp05-19110000 as name1,cpm04,cp27-19110000 as name2,cp46-19110000 as name3,cp24 from caseprogress,casepropertymap where cp01='" & strKey1 & "' and cp02='" & strKey2 & "' and cp03='" & strKey3 & "' and cp04='" & strKey4 & "' and (cp27 is not null or cp27<>0) and cp01=cpm01 and cp10=cpm02 and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B')"
   'Modify By Cheng 2002/04/15
'   strExc(0) = "select cp09,nvl(substr(cp05-19110000,1,2) || '/' || substr(cp05-19110000,3,2) || '/' || substr(cp05-19110000,5,2),null) as name1,cpm04,cp27-19110000 as name2,cp46-19110000 as name3,decode(cp24,'1','准勝','2','駁敗') as cp24 from caseprogress,casepropertymap where cp01='" & strKey1 & "' and cp02='" & strKey2 & "' and cp03='" & strKey3 & "' and cp04='" & strKey4 & "' and (cp27 is not null or cp27<>0) and cp01=cpm01(+) and cp10=cpm02(+) and (substr(cp09,1,1)='A' or substr(cp09,1,1)='B') " & _
'               "order by cp05, cp09 desc "
' 91.09.13 modify by louis
'   strExc(0) = "select cp09,nvl(substr(cp05-19110000,1,2) || '/' || substr(cp05-19110000,3,2) || '/' || substr(cp05-19110000,5,2),null) as name1,cpm04,cp27-19110000 as name2,cp46-19110000 as name3,decode(cp24,'1','准勝','2','駁敗') as cp24 from caseprogress,casepropertymap where cp01='" & strKey1 & "' and cp02='" & strKey2 & "' and cp03='" & strKey3 & "' and cp04='" & strKey4 & "' and (cp27 is not null or cp27<>0) and cp01=cpm01(+) and cp10=cpm02(+) and ( cp09<'C' ) " & _
'               "order by cp05, cp09 desc "
 '  strExc(0) = "select cp09,nvl(substr(cp05-19110000,1,2) || '/' || substr(cp05-19110000,3,2) || '/' || substr(cp05-19110000,5,2),null) as name1,cpm04,cp27-19110000 as name2,cp46-19110000 as name3,decode(cp24,'1','准勝','2','駁敗') as cp24,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap " & _
               "where cp01='" & strKey1 & "' and cp02='" & strKey2 & "' and cp03='" & strKey3 & "' and cp04='" & strKey4 & "' and (cp27 is not null or cp27<>0) and cp01=cpm01(+) and cp10=cpm02(+) and ( cp09<'C' ) " & _
               "order by SORTFIELD desc "
'Modified by Lydia 2014/10/21 民國百年問題
   strExc(0) = "select cp09,nvl(substr(cp05,1,4)-1911||'/' || substr(cp05,5,2) || '/' || substr(cp05,7,2) ,null)as name1,cpm04,cp27-19110000 as name2,cp46-19110000 as name3,decode(cp24,'1','准勝','2','駁敗') as cp24,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD from caseprogress,casepropertymap " & _
               "where cp01='" & strKey1 & "' and cp02='" & StrKey2 & "' and cp03='" & strKey3 & "' and cp04='" & strKey4 & "' and (cp27 is not null or cp27<>0) and cp01=cpm01(+) and cp10=cpm02(+) and ( cp09<'C' ) " & _
               "order by SORTFIELD desc "
   If pmain.State = adStateOpen Then pmain.Close
   pmain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If pmain.EOF Or pmain.BOF Then
      MsgBox "沒有符合條件的資料", vbOKOnly + vbCritical, "查詢資料"
      frm04010508_1.Show
      Unload Me
      GoTo EXITSUB
   End If
   Set Adodc1.Recordset = pmain
   Adodc1.Recordset.ReQuery
   
   ' 90.10.09 modify by louis
   If pmain.RecordCount = 1 Then
      cmkok_Click 0
   End If
   
EXITSUB:
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   pmain.CursorLocation = adUseClient
   pmain1.CursorLocation = adUseClient
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010508_1.m_strIR01
   m_strIR02 = frm04010508_1.m_strIR02
   m_strIR03 = frm04010508_1.m_strIR03
   m_strIR04 = frm04010508_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2022/7/1
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2022/7/1 END
   
   Unload frm04010508_3
   Set frm04010508_2 = Nothing
End Sub
