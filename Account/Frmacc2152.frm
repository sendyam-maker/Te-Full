VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2152 
   AutoRedraw      =   -1  'True
   Caption         =   "帳單資料選取"
   ClientHeight    =   5060
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5060
   ScaleWidth      =   8720
   Begin VB.TextBox Text13 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   4416
      TabIndex        =   19
      Top             =   1332
      Width           =   1300
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   2316
      TabIndex        =   17
      Top             =   1368
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
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
      Left            =   7872
      TabIndex        =   9
      Top             =   252
      Width           =   816
   End
   Begin VB.CommandButton Command2 
      Caption         =   "確定"
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
      Left            =   6924
      TabIndex        =   8
      Top             =   252
      Width           =   816
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   7092
      TabIndex        =   16
      Top             =   1368
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1128
      TabIndex        =   14
      Top             =   984
      Width           =   1416
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1128
      MaxLength       =   3
      TabIndex        =   0
      Top             =   228
      Width           =   504
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4428
      MaxLength       =   14
      TabIndex        =   5
      Top             =   228
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1632
      MaxLength       =   6
      TabIndex        =   1
      Top             =   228
      Width           =   780
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      TabIndex        =   2
      Top             =   228
      Width           =   240
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2652
      TabIndex        =   3
      Top             =   228
      Width           =   348
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   4
      Top             =   228
      Visible         =   0   'False
      Width           =   396
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2152.frx":0000
      Height          =   3072
      Left            =   216
      TabIndex        =   7
      Top             =   1740
      Width           =   8412
      _ExtentX        =   14834
      _ExtentY        =   5415
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
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
         DataField       =   "PropertyName"
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
      BeginProperty Column02 
         DataField       =   "cp05"
         Caption         =   "收文日"
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
      BeginProperty Column03 
         DataField       =   "cp27"
         Caption         =   "發文日"
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
      BeginProperty Column04 
         DataField       =   "CP14N"
         Caption         =   "承辦人"
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
         DataField       =   "RecAmount"
         Caption         =   "應收金額"
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
      BeginProperty Column06 
         DataField       =   "PayCount"
         Caption         =   "帳單次數"
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
      BeginProperty Column07 
         DataField       =   "FagentName"
         Caption         =   "代理人名稱"
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
         DataField       =   "cp44"
         Caption         =   "代理人編號"
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
         Size            =   284
         BeginProperty Column00 
            ColumnWidth     =   1179.78
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1399.748
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   980.221
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   980.221
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1039.748
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   4360.252
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1269.921
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   204
      Top             =   1656
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
   Begin MSForms.ComboBox Combo1 
      Height          =   345
      Left            =   1128
      TabIndex        =   6
      Top             =   600
      Width           =   7065
      VariousPropertyBits=   679495707
      BackColor       =   16777215
      DisplayStyle    =   3
      Size            =   "12462;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "財務其他支出"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   4392
      TabIndex        =   20
      Top             =   1056
      Width           =   1368
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "舊系統台幣帳單金額"
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
      Left            =   180
      TabIndex        =   18
      Top             =   1392
      Width           =   2112
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -12
      Top             =   4704
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "目前盈虧"
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
      Left            =   6108
      TabIndex        =   15
      Top             =   1416
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "申請國家"
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
      Left            =   168
      TabIndex        =   13
      Top             =   1032
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "案件名稱"
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
      Left            =   168
      TabIndex        =   12
      Top             =   624
      Width           =   972
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
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
      Left            =   156
      TabIndex        =   11
      Top             =   228
      Width           =   972
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "帳單金額"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc2152"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/14 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Combo1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'2005/7/7整理
Option Explicit
Public adocase As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoloop As New ADODB.Recordset
Public adocal As New ADODB.Recordset
Dim bolYes As Boolean
Dim douTWAmount As Double
Dim m_CP09 As String        '2005/7/7 ADD BY SONIA   第一筆收文之總收文號
Dim bolYesMSG As String     '2009/9/16 ADD BY SONIA  需主管審核原因
Dim m_PayAmount As Double   '2009/9/17 ADD BY SONIA  點選收文號
Dim m_bolFMP As Boolean     'add by sonia 2017/8/11  是否 FMP 案
Dim m_bolCFPC As Boolean    'add by sonia 2017/10/11 是否 CFP 案之 C 類收文號
Dim m_Country As String 'Added by Morgan 2019/5/16
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String 'Added by Morgan 2022/6/23

Private Sub Combo1_GotFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
OpenIme
End Sub

Private Sub Combo1_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Command1_Click()
   strCon9 = ""
   strCon10 = ""
   KeyEnter vbKeyEscape
End Sub

Private Sub Command2_Click()
Dim strAXF16 As String
Dim m_Amount As Double    '2009/9/16 ADD BY SONIA
   
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   If IsNull(Adodc1.Recordset.Fields("PayCount").Value) = False Then
      '2007/4/12 MODIFY BY SONIA 2->4
      'If Val(Adodc1.Recordset.Fields("PayCount").Value) > 2 Then
      If Val(Adodc1.Recordset.Fields("PayCount").Value) > 4 Then
         MsgBox MsgText(205), , MsgText(5)
         Exit Sub
      End If
   End If
   
   'Added by Morgan 2021/4/13
   If Pub_StrUserSt03 = "M31" Then
      If InStr("" & Adodc1.Recordset.Fields("cp01").Value, "P") > 0 And "" & Adodc1.Recordset.Fields("cp10").Value = "201" And "" & Adodc1.Recordset.Fields("cp14D").Value <> "F51" Then
         MsgBox "新案翻譯承辦人非外翻人員！", vbCritical
         Exit Sub
      End If
   End If
   'end 2021/4/13
   
   'ADD BY SONIA 2014/9/16 未發文不可輸帳單 S-003761
   'MODFIY BY SONIA 2014/9/18 CFP維持費除外 CFP-021263
   'Modify by Amy 2014/10/07 改以cp10判斷並加 222/208/1002/1006/1201/1209
   'modify by sonia 2014/11/25 再排除CFP的核准1001
   'modify by sonia 2014/12/30 再排除CFP的通知證書號數1602,專利證書1603
   'Modified by Morgan 2015/7/27 +排除CFP的通知要求選取1206
   'If Val(Adodc1.Recordset.Fields("CP27").Value) = 0 And Text10 & Adodc1.Recordset.Fields("PropertyName").Value <> "CFP維持費" Then
   'Modified by Lydia 2015/10/05 + 1008
   'Modified by Morgan 2018/7/17 + CFP通知面詢1401--禧佩 Ex:CFP-27406
   'Modify by sonia 2019/9/10 + 1811
   If Val(Adodc1.Recordset.Fields("CP27").Value) = 0 And Text10 = "CFP" And Not (Adodc1.Recordset.Fields("CP10").Value = "208" Or Adodc1.Recordset.Fields("CP10").Value = "222" Or Adodc1.Recordset.Fields("CP10").Value = "606" _
        Or Adodc1.Recordset.Fields("CP10").Value = "1002" Or Adodc1.Recordset.Fields("CP10").Value = "1006" Or Adodc1.Recordset.Fields("CP10").Value = "1201" Or Adodc1.Recordset.Fields("CP10").Value = "1209" _
        Or Adodc1.Recordset.Fields("CP10").Value = "1001" Or Adodc1.Recordset.Fields("CP10").Value = "1602" Or Adodc1.Recordset.Fields("CP10").Value = "1603" Or Adodc1.Recordset.Fields("CP10").Value = "1206" _
        Or Adodc1.Recordset.Fields("CP10").Value = "1008" Or Adodc1.Recordset.Fields("CP10").Value = "1401" Or Adodc1.Recordset.Fields("CP10").Value = "1811") Then
      MsgBox "此收文號尚未發文，不可輸入帳單！", , MsgText(5)
      Exit Sub
   End If
   '2014/9/16 END
   '2009/9/23 add by sonia
   If IsNull(Adodc1.Recordset.Fields("CP64").Value) = False And (Text10 = "CFP" Or Text10 = "CPS") Then
      MsgBox "該筆進度備註：" & Adodc1.Recordset.Fields("CP64").Value, , MsgText(5)
   End If
   '2009/9/23 end
   
   If adoaccsum.State = adStateOpen Then
      adoaccsum.Close
   End If
   adoaccsum.CursorLocation = adUseClient
   '2005/8/26 MODIFY BY SONIA
   'adoaccsum.Open "select a2103 from acc210 where a2102 = '" & strTitle & "' and a2101 = (select max(a2101) from acc210 where a2102 = '" & strTitle & "' and a2101 <= " & Val(ACDate(ServerDate)) & ")", adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select a2103 from acc210 where a2102 = '" & Frmacc2150.Combo1 & "' and a2101 = (select max(a2101) from acc210 where a2102 = '" & strTitle & "' and a2101 <= " & strSrvDate(2) & ")", adoTaie, adOpenStatic, adLockReadOnly
   '2005/8/26 END
   If adoaccsum.RecordCount <> 0 Then
      douTWAmount = (Val(Text11) * Val(adoaccsum.Fields("a2103").Value))
   Else
      douTWAmount = Val(Text11)
   End If
      
   'Added by Morgan 2015/9/23
   '檢查翻譯費支出是否超過比例
   If PUB_ChkTranslationFee(Adodc1.Recordset.Fields("CP09").Value, douTWAmount, True, strCon8) = False Then Exit Sub
   'end 2015/9/23
   
'2009/9/16 CANCEL BY SONIA 應該下一句就可以
'   If Adodc1.Recordset.Bookmark = Adodc1.Recordset.RecordCount Then
'      bolYes = False
'      bolYesMSG = ""             '2009/9/16 add by sonia
'   End If
'2009/9/16 END
   '判斷若該案第一筆收文未輸過帳單,點選第一筆時不出現審核訊息
   '2005/7/7 MODIFY BY SONIA 因改變排序方式,不判斷點選第一筆改判斷點選之收文號
   'If Adodc1.Recordset.Bookmark = 1 Then
   If Adodc1.Recordset.Fields("cp09").Value = m_CP09 Then
   '2005/7/7 END
      bolYes = False
      bolYesMSG = ""             '2009/9/16 add by sonia
   End If
   'Ken 92/08/22 判斷若舊系統有金額表示已開過帳單
   If Val(Text8) > 0 Then
      bolYes = False
      bolYesMSG = ""             '2009/9/16 add by sonia
   End If
   If Adodc1.Recordset.Fields("PayCount").Value > 0 Then  '此收文號已輸過帳單
      bolYes = True
      If bolYesMSG = "" Then
         bolYesMSG = "此收文號已輸過帳單"
      Else
         bolYesMSG = bolYesMSG & ",此收文號已輸過帳單"
      End If
   End If
   'add by sonia 2017/8/11 判斷是否FMP案
   If Mid(Adodc1.Recordset.Fields("cp12").Value, 1, 1) = "F" And InStr(Text10, "P") > 0 Then
      m_bolFMP = True
   Else
      m_bolFMP = False
   End If
   Frmacc2150.strFMP = m_bolFMP  'add by sonia 2017/9/13
   'end 2017/8/11
   
   'add by sonia 2017/10/11 CFP之C類不檢查收文號及案號之虧損
   m_bolCFPC = False
   'Modified by Morgan 2018/8/1 +判斷有費用的還是要檢查 Ex:CFP-27740(專利證書)
   If Text10 = "CFP" And Left(Adodc1.Recordset.Fields("CP09").Value, 1) = "C" And Val("" & Adodc1.Recordset.Fields("CP16").Value) = 0 Then
      m_bolCFPC = True
   End If
   Frmacc2150.strCFPC = m_bolCFPC
   'end 2017/10/11
   
   'modify by sonia 2017/10/11 CFP之C類不檢查收文號及案號之虧損
   If Not m_bolFMP And Not m_bolCFPC Then  'add by sonia 2017/8/11 FMP案不判斷虧損
      If Val(Text2) < 0 Then                                '此案號目前有虧損,另最小收文號尚未輸過帳單在FORM_LOAD處理
         bolYes = True
         If bolYesMSG = "" Then
            bolYesMSG = "此案號目前有虧損"
         Else
            bolYesMSG = bolYesMSG & ",此案號目前有虧損"
         End If
      End If
   End If

   '2009/9/17 ADD BY SONIA 此收文號其他帳單金額CFP-019921
   m_PayAmount = 0
   If adoaccsum.State = adStateOpen Then
      adoaccsum.Close
   End If
   adoaccsum.CursorLocation = adUseClient
   '2014/3/7 modify by sonia 加入抓抵帳匯率ACC1G0(CFP-019761的U10302128)
   'adoaccsum.Open "select sum(decode(a1507, null, nvl(NVL(AXF04*decode(A1906,0,null,A1906),axf15), 0), 0)) as PayAmount from acc151, acc150, ACC190 where axf01 = a1501 AND AXF01=A1902(+) and axf02 = '" & Adodc1.Recordset.Fields("cp09").Value & "' ", adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select sum(decode(a1507, null, nvl(NVL(AXF04*decode(A1906,0,NULL,A1906),NVL(AXF04*decode(A1G03,0,null,A1G03),AXF15)), 0), 0)) as PayAmount from acc151, acc150, ACC190, ACC1G0 where axf01 = a1501 AND AXF01=A1902(+) AND A1512=A1G01(+) and axf02 = '" & Adodc1.Recordset.Fields("cp09").Value & "' ", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields("PayAmount").Value) = False Then m_PayAmount = Val(adoaccsum.Fields("PayAmount").Value)
   End If
   '2009/9/17 end
   '2009/9/16 加判斷該收文號之損益,帳單金額>收入-點數*1000-此收文號其他帳單金額-安全基金
   If Val(Adodc1.Recordset.Fields("RecAmount")) > 0 Then   'RecAmount=0者不可再減點數,否則負負為正CFP-022018答辯
      m_Amount = Val(Adodc1.Recordset.Fields("RecAmount")) - (Val(Adodc1.Recordset.Fields("cp18").Value) * 1000) - m_PayAmount
   End If
   '新案件才扣安全基金
   If Adodc1.Recordset.Fields("cp31") = "Y" Then
      If Text10 = "TF" Then
         m_Amount = m_Amount - GetFloatPrepareCase(Text10.Text, Text5.Text & Text7.Text, Text9.Text, Text12.Text)
      Else
         m_Amount = m_Amount - GetFloatPrepareCase(Text10.Text, Text5.Text, Text7.Text, Text9.Text)
      End If
   End If
   'modify by sonia 2017/10/11 CFP之C類不檢查收文號及案號之虧損
   'modify by sonia 2021/4/22 FF案件請款單非主要請款案件性質不管收文號有無虧損U11002516
   'If Not m_bolFMP And Not m_bolCFPC Then
   If Not m_bolFMP And Not m_bolCFPC And Not (Adodc1.Recordset.Fields("RecAmount") = 0 And "" & Adodc1.Recordset.Fields("cp60") > "X") Then 'add by sonia 2017/8/11 FMP案不判斷虧損
      If Val(douTWAmount) > m_Amount Then                    '此收文號有虧損
         bolYes = True
         If bolYesMSG = "" Then
            bolYesMSG = "此收文號有虧損"
         Else
            bolYesMSG = bolYesMSG & ",此收文號有虧損"
         End If
      End If
   End If
   '2009/9/16 END
   adoaccsum.Close
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   'Modify by Morgan 2008/2/14 +Ex1,判斷美國發明的領證費帳單是否含公開費
   adoquery.Open "select cp01||cp02||cp03||cp04 as CaseNo, pa26 as CustomerNo, a0k04, cp61,cp01||pa09||pa08||cp10 Ex1 from caseprogress, patent, acc0k0 where cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp60 = a0k01 (+) and cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' union " & _
                 "select cp01||cp02||cp03||cp04 as CaseNo, tm23 as CustomerNo, a0k04, cp61,'' Ex1 from caseprogress, trademark, acc0k0 where cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp60 = a0k01 (+) and cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' union " & _
                 "select cp01||cp02||cp03||cp04 as CaseNo, lc11 as CustomerNo, a0k04, cp61,'' Ex1 from caseprogress, lawcase, acc0k0 where cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp60 = a0k01 (+) and cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' union " & _
                 "select cp01||cp02||cp03||cp04 as CaseNo, hc05 as CustomerNo, a0k04, cp61,'' Ex1 from caseprogress, hirecase, acc0k0 where cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and cp60 = a0k01 (+) and cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' union " & _
                 "select cp01||cp02||cp03||cp04 as CaseNo, sp08 as CustomerNo, a0k04, cp61,'' Ex1 from caseprogress, servicepractice, acc0k0 where cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp60 = a0k01 (+) and cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
   
      'Add by Morgan 2008/2/14
      strAXF16 = ""
      If "" & adoquery.Fields("Ex1") = "CFP1011601" Or "" & adoquery.Fields("Ex1") = "CFP1011217" Then
         strExc(0) = "select 1 from acc151 where axf02='" & Adodc1.Recordset.Fields("cp09").Value & "' and axf16='Y'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            If MsgBox("本帳單是否含公開費？", vbYesNo + vbDefaultButton1, "詢問") = vbYes Then
               strAXF16 = "Y"
            End If
         End If
      End If
      'end 2008/2/14
      
      'Modify by Morgan 2008/2/14 +axf16
      If IsNull(adoquery.Fields("CustomerNo").Value) = False Then
        'Modify By Cheng 2004/01/05
'         strCon9 = "insert into acc151 (axf01, axf02, axf03, axf04, axf05, axf06, axf07, axf08, axf12, axf13, axf14, axf15) values ('" & strCon8 & "', '" & Adodc1.Recordset.Fields("cp09").Value & "', '" & adoquery.Fields("CaseNo").Value & "', " & Val(Text11) & ", '" & adoquery.Fields("CustomerNo").Value & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', '" & Mid(Combo1, 4, Len(Combo1)) & "', '" & adoquery.Fields("a0k04").Value & "', " & Val(Text2) & ", " & douTWAmount & ")"
         strCon9 = "insert into acc151 (axf01, axf02, axf03, axf04, axf05, axf06, axf07, axf08, axf12, axf13, axf14, axf15,axf16) values ('" & strCon8 & "', '" & Adodc1.Recordset.Fields("cp09").Value & "', '" & adoquery.Fields("CaseNo").Value & "', " & Val(Text11) & ", '" & adoquery.Fields("CustomerNo").Value & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & ChgSQL(Mid(Combo1, 4, Len(Combo1))) & "', '" & ChgSQL(IIf(IsNull(adoquery.Fields("a0k04").Value), "", adoquery.Fields("a0k04").Value)) & "', " & Val(Text2) & ", " & douTWAmount & "," & CNULL(strAXF16) & ")"
        'End
      Else
        'Modify By Cheng 2004/01/05
'         strCon9 = "insert into acc151 (axf01, axf02, axf03, axf04, axf05, axf06, axf07, axf08, axf12, axf13, axf14, axf15) values ('" & strCon8 & "', '" & Adodc1.Recordset.Fields("cp09").Value & "', '" & adoquery.Fields("CaseNo").Value & "', " & Val(Text11) & ", null, " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', '" & Mid(Combo1, 4, Len(Combo1)) & "', '" & adoquery.Fields("a0k04").Value & "', " & Val(Text2) & ", " & douTWAmount & ")"
         strCon9 = "insert into acc151 (axf01, axf02, axf03, axf04, axf05, axf06, axf07, axf08, axf12, axf13, axf14, axf15,axf16) values ('" & strCon8 & "', '" & Adodc1.Recordset.Fields("cp09").Value & "', '" & adoquery.Fields("CaseNo").Value & "', " & Val(Text11) & ", null, " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & ChgSQL(Mid(Combo1, 4, Len(Combo1))) & "', '" & ChgSQL(IIf(IsNull(adoquery.Fields("a0k04").Value), "", adoquery.Fields("a0k04").Value)) & "', " & Val(Text2) & ", " & douTWAmount & "," & CNULL(strAXF16) & ")"
        'End
      End If
      '2007/4/12 MODIFY BY SONIA
      'strCon10 = "begin update caseprogress set cp61=nvl(cp61, '" & strCon8 & "') where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'; " & _
      '           "update caseprogress set cp62=decode(cp61, '" & strCon8 & "', null, nvl(cp62, '" & strCon8 & "')) where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'; " & _
      '           "update caseprogress set cp63=decode(cp61, '" & strCon8 & "', null, decode(cp62, '" & strCon8 & "', null, nvl(cp63, '" & strCon8 & "'))) where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'; end;"
      
      'Modified by Morgan 2017/1/16 修正修改第一張帳單時後面的都會被清除問題
      'strCon10 = "begin update caseprogress set cp61=nvl(cp61, '" & strCon8 & "') where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'; " & _
                 "update caseprogress set cp62=decode(cp61, '" & strCon8 & "', null, nvl(cp62, '" & strCon8 & "')) where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'; " & _
                 "update caseprogress set cp63=decode(cp61, '" & strCon8 & "', null, decode(cp62, '" & strCon8 & "', null, nvl(cp63, '" & strCon8 & "'))) where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'; " & _
                 "update caseprogress set cp87=decode(cp61, '" & strCon8 & "', null, decode(cp62, '" & strCon8 & "', null, decode(cp63, '" & strCon8 & "',null, nvl(cp87, '" & strCon8 & "')))) where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'; " & _
                 "update caseprogress set cp88=decode(cp61, '" & strCon8 & "', null, decode(cp62, '" & strCon8 & "', null, decode(cp63, '" & strCon8 & "',null, decode(cp87, '" & strCon8 & "',null, nvl(cp88, '" & strCon8 & "'))))) where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'; end;"
      strCon10 = "begin update caseprogress set cp61='" & strCon8 & "' where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' and cp61 is null; " & _
                 "update caseprogress set cp62='" & strCon8 & "' where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' and cp62 is null and instr(cp61,'" & strCon8 & "')=0 ; " & _
                 "update caseprogress set cp63='" & strCon8 & "' where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' and cp63 is null and instr(cp61||cp62,'" & strCon8 & "')=0 ; " & _
                 "update caseprogress set cp87='" & strCon8 & "' where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' and cp87 is null and instr(cp61||cp62||cp63,'" & strCon8 & "')=0 ; " & _
                 "update caseprogress set cp88='" & strCon8 & "' where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' and cp88 is null and instr(cp61||cp62||cp63||cp87,'" & strCon8 & "')=0 ; end;"
      'end 2017/1/16
      '2007/4/12 END
   Else
      strCon9 = ""
      strCon10 = ""
   End If
   adoquery.Close
   If IsNull(Adodc1.Recordset.Fields("cp44").Value) = False Then
      strCustNo = Adodc1.Recordset.Fields("cp44").Value
   End If
   If bolYes Then
      MsgBox bolYesMSG & "！" & MsgText(187), , MsgText(5)
      Frmacc2150.strYes = MsgText(603)
      'Added by Morgan 2019/3/15
      Frmacc2150.Text8 = Replace(Frmacc2150.Text8, bolYesMSG & ";", "")
      Frmacc2150.Text8 = bolYesMSG & ";" & Frmacc2150.Text8
      'end 2019/3/15
   Else
      Frmacc2150.strYes = MsgText(601)
   End If
   
   'Added by Morgan 2012/4/25
   '有集體案件要提醒
   If Text10 = "CFP" And Text7 = "0" Then
      strExc(0) = "select * from caseprogress where cp01='" & Text10 & "' and cp02='" & Text5 & "' and cp03<>'0' and cp57 is null and cp27>0 and cp10='105'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox "本案有集體案件!!", vbExclamation
      End If
   End If
   'end 2012/4/25
   '2013/1/25 ADD BY SONIA 點選程序若只有服務費無規費者提醒使用者注意若為憑帳單請款,要記得輸入代理人請款之來函向客戶請款
   strExc(0) = "select * from caseprogress where cp09='" & Adodc1.Recordset.Fields("cp09").Value & "' AND NVL(CP18,0)>0 AND NVL(CP17,0)=0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox "本程序只有服務費無規費，請注意若為憑帳單請款，要記得輸入代理人請款之來函向客戶請款!!", vbExclamation
   End If
   '2013/1/25 END
   
   'Added by Morgan 2022/6/23
   ' PS及CPS案件可計結餘明細 --郭雅娟
   If Text10 = "PS" Or Text10 = "CPS" Then
      Pub_EndModCashMsg m_Country, Text10.Text, Text5.Text, Text7.Text, Text9.Text
      Pub_UpdateEndModCash Text10.Text, Text5.Text, Text7.Text, Text9.Text
   End If
   'end 2022/6/23
   'add by sonia 2024/12/20 P,CFP之809提第三方意見及CFP之606維持費、607延展費、第五年年費605不必詢問直接上可結餘日
   If (Text10 = "P" Or Text10 = "CFP") And Adodc1.Recordset.Fields("cp10").Value = "809" Then
      bolEndModCash = True
      Pub_UpdateEndModCash Text10.Text, Text5.Text, Text7.Text, Text9.Text
   ElseIf (Text10 = "P" Or Text10 = "CFP") And (Adodc1.Recordset.Fields("cp10").Value = "606" Or Adodc1.Recordset.Fields("cp10").Value = "607") Then
      bolEndModCash = True
      Pub_UpdateEndModCash Text10.Text, Text5.Text, Text7.Text, Text9.Text
   ElseIf (Text10 = "P" Or Text10 = "CFP") And (Adodc1.Recordset.Fields("cp10").Value = "605" And Adodc1.Recordset.Fields("cp53").Value <= 5 And Adodc1.Recordset.Fields("cp54").Value >= 5) Then
      bolEndModCash = True
      Pub_UpdateEndModCash Text10.Text, Text5.Text, Text7.Text, Text9.Text
   'add by sonia 2025/3/31 再加第10年，第15年年費
   ElseIf (Text10 = "P" Or Text10 = "CFP") And (Adodc1.Recordset.Fields("cp10").Value = "605" And Adodc1.Recordset.Fields("cp53").Value <= 10 And Adodc1.Recordset.Fields("cp54").Value >= 10) Then
      bolEndModCash = True
      Pub_UpdateEndModCash Text10.Text, Text5.Text, Text7.Text, Text9.Text
   ElseIf (Text10 = "P" Or Text10 = "CFP") And (Adodc1.Recordset.Fields("cp10").Value = "605" And Adodc1.Recordset.Fields("cp53").Value <= 15 And Adodc1.Recordset.Fields("cp54").Value >= 15) Then
      bolEndModCash = True
      Pub_UpdateEndModCash Text10.Text, Text5.Text, Text7.Text, Text9.Text
   'end 2025/3/31
   End If
   'end 2024/12/20
   
   Unload Me
   tool2_enabled
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   Calculate
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   m_CP09 = ""      '2005/7/7 ADD BY SONIA
   
  'Modified by Lydia 2021/12/14 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 8850
'   Me.Height = 5500
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   PUB_InitForm Me, 8850, 5800, strBackPicPath1
   'end 2021/12/07
   
   Text10 = strCon2
   If Text10 = "TF" Then
      Text12.Visible = True
   Else
      Text12.Visible = False
   End If
   Text5 = strCon3
   Text7 = strCon4
   Text9 = strCon5
   Text12 = strCon6
   Text11 = strCon7
   FormShow
   AdodcRefresh
   Calculate
   bolYes = False
   bolYesMSG = "" '2009/9/16 add by sonia
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   
'modify by sonia 2021/4/29 應改用adocal.Clone的資料來檢查U11000603
'   'Modified by Morgan 2019/5/16 台灣案不必檢查--秀玲
'   'If Adodc1.Recordset.RecordCount > 0 Then
'   If Adodc1.Recordset.RecordCount > 0 And m_Country <> "000" Then
'   'end 2019/5/16
'      Adodc1.Recordset.MoveLast
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select cp61 from caseprogress where cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' and cp61 is not null", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount = 0 Then
'         bolYes = True                                       '最小收文號尚未輸過帳單在FORM_LOAD處理
'         m_CP09 = Adodc1.Recordset.Fields("cp09").Value      '2005/7/7 ADD BY SONIA
'         bolYesMSG = "最小收文號尚未輸過帳單"                '2009/9/16 add by sonia
'      End If
'      adoquery.Close
'      'add by sonia 2017/8/11
'      If bolYesMSG = "" Then
'         Adodc1.Recordset.MoveFirst
'      End If
'      'end 2017/8/11
'   End If
   Set adoloop = adocal.Clone
   If adoloop.RecordCount > 0 And m_Country <> "000" Then
      adoloop.MoveLast
      adoquery.CursorLocation = adUseClient
      adoquery.Open "select cp61 from caseprogress where cp09 = '" & adoloop.Fields("cp09").Value & "' and cp61 is not null", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount = 0 Then
         bolYes = True                                       '最小收文號尚未輸過帳單在FORM_LOAD處理
         m_CP09 = adoloop.Fields("cp09").Value
         bolYesMSG = "最小收文號尚未輸過帳單"
      End If
      adoquery.Close
      If bolYesMSG = "" Then
         adoloop.MoveFirst
      End If
   End If
   
   'Added by Morgan 2021/4/13
   If Pub_StrUserSt03 = "M31" Then
      DataGrid1.Columns(5).Width = 0
   Else
      DataGrid1.Columns(4).Width = 0
   End If
   'end 2021/4/13
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Removed by Morgan 2019/8/2
   '改在 Frmacc2150.Form_Unload(此處寄信後回前畫面會無法帶出資料且會無法操作,可能跟在transaction內有關)
   'PUB_SendMailCache 'Added by Lydia 2019/07/03
   'end 2019/8/2
   strItemNo = ""
   tool3_enabled
   Frmacc2150.Enabled = True
   Frmacc2150.Show
   Set Frmacc2152 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   If adocase.State = adStateOpen Then
      adocase.Close
   End If
   If adocal.State = adStateOpen Then
      adocal.Close
   End If
   adocase.CursorLocation = adUseClient
   adocal.CursorLocation = adUseClient
   '2009/9/11 modify by sonia 帳單申請國家一定非台灣故不抓CPM03改抓CPM04(把decode(cpm03, '（無）', cpm04, cpm03)改為CPM04),U09807576大陸復審答辯406會出現參加訴願
   '2009/9/22 modify by sonia 加進度備註
   '2009/10/8 modify by sonia FCP之舜禹翻譯帳單的案件性質會出現（無）U09808512
   'Modified by Morgan 2018/8/1 +CP16
   'Modified by Morgan 2019/3/27 改先判斷編號再名稱，否則若名稱有特殊符號如 "&"時會找不到資料 Ex:U10802622
   'Modified by Morgan 2021/4/13 +,cp01,st03 CP14D,st02 CP14N
   Select Case Text10
      Case "TF"
         If strCustNo <> "" Then
            '93.7.16 modify by sonia: cp16-cp77 改成 DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))
            'adocase.Open "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and fa01 = '" & Mid(strCustNo, 1, 8) & "' and fa02 = '" & Mid(strCustNo, 9, 1) & "' union " & _
            '             "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(a1k30, 0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent, acc1k0 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and fa01 = '" & Mid(strCustNo, 1, 8) & "' and fa02 = '" & Mid(strCustNo, 9, 1) & _
            '             "' order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            'adocal.Open "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' union " & _
            '            "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(a1k30, 0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent, acc1k0 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            '2009/9/16 MODIFY BY SONIA 加入 CP31
            'modify by sonia +cp12
            'modify by sonia 2021/4/21 解決FF案件請款單重覆計算問題U11002495(CFP-032254),+cp60以判斷FF案件請款單非主要請款案件性質不管收文號有無虧損
            'adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and fa01 = '" & Mid(strCustNo, 1, 8) & "' and fa02 = '" & Mid(strCustNo, 9, 1) & "' and st01(+)=cp14 union " & _
                         "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, NVL(nvl(a1k30, A1K11),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, acc1k0, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and fa01 = '" & Mid(strCustNo, 1, 8) & "' and fa02 = '" & Mid(strCustNo, 9, 1) & "' and st01(+)=cp14 " & _
                         " order by CP05 DESC,cp09 DEsc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2024/12/20 +cp53,cp54
            adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,cp60,cp53,cp54 from caseprogress, casepropertymap, fagent, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and fa01 = '" & Mid(strCustNo, 1, 8) & "' and fa02 = '" & Mid(strCustNo, 9, 1) & "' and st01(+)=cp14 union " & _
                         "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,cp60,cp53,cp54 from caseprogress, casepropertymap, fagent, acc1k0, staff, " & _
                         " (SELECT cp60 z1,MIN(cp09) z2 FROM caseprogress WHERE 1=1 AND cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 GROUP BY cp60) z " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and fa01 = '" & Mid(strCustNo, 1, 8) & "' and fa02 = '" & Mid(strCustNo, 9, 1) & "' and st01(+)=cp14 and z1(+)=cp60 " & _
                         " order by CP05 DESC,cp09 DEsc", adoTaie, adOpenStatic, adLockReadOnly
            'end 2021/4/21
            'modify by sonia 2021/4/21 解決FF案件請款單重覆計算問題U11002495(CFP-032254),同時取消CP61~CP63,CP87,CP88條件，否則計算目前盈虧會少算
            'adocal.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                        "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, NVL(nvl(a1k30, A1K11),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, acc1k0, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2021/4/29 加CP66,CP67改排序原為CP09 ASC改為CP66 DESC,CP67 DESC,CP09 DESC,U11000603(CFP-031978會抓到AA9041239而非AA9041238)造成最小收文號尚未輸過帳單
            adocal.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,CP66,CP67 from caseprogress, casepropertymap, fagent, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                        "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,CP66,CP67 from caseprogress, casepropertymap, fagent, acc1k0, staff, " & _
                         " (SELECT cp60 z1,MIN(cp09) z2 FROM caseprogress WHERE 1=1 AND cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 GROUP BY cp60) z " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 and z1(+)=cp60 order by CP05 DESC,CP66 DESC,CP67 DESC,CP09 DESC", adoTaie, adOpenStatic, adLockReadOnly
            'end 2021/4/21
            '93.7.16 END
         ElseIf strCon1 <> "" Then
            '93.7.16 modify by sonia: cp16-cp77 改成 DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))
            'adocase.Open "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and fa05||fa63||fa64||fa65 like '" & Replace(strCon1, "'", "''") & "%' union " & _
            '             "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(a1k30, 0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent, acc1k0 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and fa05||fa63||fa64||fa65 like '" & Replace(strCon1, "'", "''") & _
            '             "%' order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            'adocal.Open "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' union " & _
            '            "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(a1k30, 0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent, acc1k0 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            '94.1.19 MODIFY BY SONIA 改 CP09 之排序,原為 ASC 改為 DESC
            '2006/4/28 MODIFY BY SONIA fa05||fa63||fa64||fa65-->UPPER(fa05||fa63||fa64||fa65)
            '2006/10/23 MODIFY BY SONIA UPPER(fa05||fa63||fa64||fa65)-->UPPER(NVL(fa05||fa63||fa64||fa65,FA04))
            '2007/4/12 MODIFY BY SONIA 加入 CP87,CP88
            '2009/9/16 MODIFY BY SONIA 加入 CP31
            '2013/5/31 modify by sonia 剔除假收文資料T-152202(U10203421),故加入 and cp05<>19221111
            'Modify by Amy 2014/10/07 +cp10
            'modify by sonia +cp12
            '2021/4/21 modify by sonia 解決FF案件請款單重覆計算問題U11002495(CFP-032254),+cp60以判斷FF案件請款單非主要請款案件性質不管收文號有無虧損
            'adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, staff  " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and cp05<>19221111 and UPPER(NVL(fa05||fa63||fa64||fa65,FA04)) like UPPER('" & Replace(strCon1, "'", "''") & _
                         "%') and st01(+)=cp14 union " & _
                         "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, NVL(nvl(a1k30, A1K11),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, acc1k0, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and UPPER(NVL(fa05||fa63||fa64||fa65,FA04)) like UPPER('" & Replace(strCon1, "'", "''") & _
                         "%') and st01(+)=cp14 order by CP05 DESC,cp09 DEsc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2024/12/20 +cp53,cp54
            adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,cp60,cp53,cp54 from caseprogress, casepropertymap, fagent, staff  " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and cp05<>19221111 and UPPER(NVL(fa05||fa63||fa64||fa65,FA04)) like UPPER('" & Replace(strCon1, "'", "''") & _
                         "%') and st01(+)=cp14 union " & _
                         "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,cp60,cp53,cp54 from caseprogress, casepropertymap, fagent, acc1k0, staff, " & _
                         " (SELECT cp60 z1,MIN(cp09) z2 FROM caseprogress WHERE 1=1 AND cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 GROUP BY cp60) z " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and UPPER(NVL(fa05||fa63||fa64||fa65,FA04)) like UPPER('" & Replace(strCon1, "'", "''") & _
                         "%') and st01(+)=cp14 and z1(+)=cp60 order by CP05 DESC,cp09 DEsc", adoTaie, adOpenStatic, adLockReadOnly
            'end 2021/4/21
            'modify by sonia 2021/4/21 解決FF案件請款單重覆計算問題U11002495(CFP-032254),同時取消CP61~CP63,CP87,CP88條件，否則計算目前盈虧會少算
            'adocal.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10 ,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                        "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, NVL(nvl(a1k30, A1K11),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, acc1k0, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2021/4/29 加CP66,CP67改排序原為CP09 ASC改為CP66 DESC,CP67 DESC,CP09 DESC,U11000603(CFP-031978會抓到AA9041239而非AA9041238)造成最小收文號尚未輸過帳單
            adocal.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,CP66,CP67 from caseprogress, casepropertymap, fagent, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                        "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,CP66,CP67 from caseprogress, casepropertymap, fagent, acc1k0, staff, " & _
                        " (SELECT cp60 z1,MIN(cp09) z2 FROM caseprogress WHERE 1=1 AND cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 GROUP BY cp60) z " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 and z1(+)=cp60 order by CP05 DESC,CP66 DESC,CP67 DESC,CP09 DESC", adoTaie, adOpenStatic, adLockReadOnly
            'end 2021/4/21
            '93.7.16 END
         Else
            '93.7.16 modify by sonia: cp16-cp77 改成 DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))
            'adocase.Open "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' union " & _
            '             "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(a1k30, 0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent, acc1k0 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            'adocal.Open "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' union " & _
            '            "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(a1k30, 0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent, acc1k0 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            '2009/9/16 MODIFY BY SONIA 加入 CP31
            'modify by sonia +cp12
            '2021/4/21 modify by sonia 解決FF案件請款單重覆計算問題U11002495(CFP-032254),+cp60以判斷FF案件請款單非主要請款案件性質不管收文號有無虧損
            'adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                         "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, NVL(nvl(a1k30, A1K11),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, acc1k0, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 order by CP05 DESC,cp09 DEsc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2024/12/20 +cp53,cp54
            adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,cp60,cp53,cp54 from caseprogress, casepropertymap, fagent, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                         "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,cp60,cp53,cp54 from caseprogress, casepropertymap, fagent, acc1k0, staff, " & _
                         " (SELECT cp60 z1,MIN(cp09) z2 FROM caseprogress WHERE 1=1 AND cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 GROUP BY cp60) z " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 and z1(+)=cp60 order by CP05 DESC,cp09 DEsc", adoTaie, adOpenStatic, adLockReadOnly
            'end 2021/4/21
            'modify by sonia 2021/4/21 解決FF案件請款單重覆計算問題U11002495(CFP-032254),同時取消CP61~CP63,CP87,CP88條件，否則計算目前盈虧會少算
            'adocal.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                        "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, NVL(nvl(a1k30, A1K11),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, acc1k0, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2021/4/29 加CP66,CP67改排序原為CP09 ASC改為CP66 DESC,CP67 DESC,CP09 DESC,U11000603(CFP-031978會抓到AA9041239而非AA9041238)造成最小收文號尚未輸過帳單
            adocal.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,CP66,CP67 from caseprogress, casepropertymap, fagent, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                        "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,CP66,CP67 from caseprogress, casepropertymap, fagent, acc1k0, staff, " & _
                        " (SELECT cp60 z1,MIN(cp09) z2 FROM caseprogress WHERE 1=1 AND cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 GROUP BY cp60) z " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp05<>19221111 and st01(+)=cp14 and z1(+)=cp60 order by CP05 DESC,CP66 DESC,CP67 DESC,CP09 DESC", adoTaie, adOpenStatic, adLockReadOnly
            'end 2021/4/21
            '93.7.16 END
         End If

      Case Else
         
         If strCustNo <> "" Then
            '93.7.16 modify by sonia: cp16-cp77 改成 DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))
            'adocase.Open "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and fa01 = '" & Mid(strCustNo, 1, 8) & "' and fa02 = '" & Mid(strCustNo, 9, 1) & "' union " & _
            '             "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(a1k30, 0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent, acc1k0 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and fa01 = '" & Mid(strCustNo, 1, 8) & "' and fa02 = '" & Mid(strCustNo, 9, 1) & _
            '             "' order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            'adocal.Open "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' union " & _
            '            "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(a1k30, 0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent, acc1k0 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            '2009/9/16 MODIFY BY SONIA 加入 CP31
            'modify by sonia +cp12
            '2021/4/21 modify by sonia 解決FF案件請款單重覆計算問題U11002495(CFP-032254),+cp60以判斷FF案件請款單非主要請款案件性質不管收文號有無虧損
            'adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and fa01 = '" & Mid(strCustNo, 1, 8) & "' and fa02 = '" & Mid(strCustNo, 9, 1) & "' and st01(+)=cp14 union " & _
                         "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, NVL(nvl(a1k30, A1K11),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, acc1k0, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and fa01 = '" & Mid(strCustNo, 1, 8) & "' and fa02 = '" & Mid(strCustNo, 9, 1) & _
                         "' and st01(+)=cp14 order by CP05 DESC,cp09 DEsc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2024/12/20 +cp53,cp54
            adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,cp60,cp53,cp54 from caseprogress, casepropertymap, fagent, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and fa01 = '" & Mid(strCustNo, 1, 8) & "' and fa02 = '" & Mid(strCustNo, 9, 1) & "' and st01(+)=cp14 union " & _
                         "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0) AS RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0) AS cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,cp60,cp53,cp54 from caseprogress, casepropertymap, fagent, acc1k0, staff, " & _
                         " (SELECT cp60 z1,MIN(cp09) z2 FROM caseprogress WHERE 1=1 AND cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 GROUP BY cp60) z " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and fa01 = '" & Mid(strCustNo, 1, 8) & "' and fa02 = '" & Mid(strCustNo, 9, 1) & _
                         "' and st01(+)=cp14 and z1(+)=cp60 order by CP05 DESC,cp09 DEsc", adoTaie, adOpenStatic, adLockReadOnly
            'end 2021/4/21
            'modify by sonia 2021/4/21 解決FF案件請款單重覆計算問題U11002495(CFP-032254),同時取消CP61~CP63,CP87,CP88條件，否則計算目前盈虧會少算
            'adocal.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                        "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, NVL(nvl(a1k30, A1K11),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, acc1k0, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2021/4/29 加CP66,CP67改排序原為CP09 ASC改為CP66 DESC,CP67 DESC,CP09 DESC,U11000603(CFP-031978會抓到AA9041239而非AA9041238)造成最小收文號尚未輸過帳單
            adocal.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,CP66,CP67 from caseprogress, casepropertymap, fagent, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                        "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0) AS RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0) AS cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,CP66,CP67 from caseprogress, casepropertymap, fagent, acc1k0, staff, " & _
                        " (SELECT cp60 z1,MIN(cp09) z2 FROM caseprogress WHERE 1=1 AND cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 GROUP BY cp60) z " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 and z1(+)=cp60 order by CP05 DESC,CP66 DESC,CP67 DESC,CP09 DESC", adoTaie, adOpenStatic, adLockReadOnly
            'end 2021/4/21
            '93.7.16 END
            
         ElseIf strCon1 <> "" Then
            '93.7.16 modify by sonia: cp16-cp77 改成 DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))
            'adocase.Open "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and fa05||fa63||fa64||fa65 like '" & Replace(strCon1, "'", "''") & "%' union " & _
            '             "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(a1k30, 0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent, acc1k0 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and fa05||fa63||fa64||fa65 like '" & Replace(strCon1, "'", "''") & _
            '             "%' order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            'adocal.Open "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' union " & _
            '            "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(a1k30, 0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent, acc1k0 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            '2006/4/28 MODIFY BY SONIA fa05||fa63||fa64||fa65-->UPPER(fa05||fa63||fa64||fa65)
            '2006/10/23 MODIFY BY SONIA UPPER(fa05||fa63||fa64||fa65)-->UPPER(NVL(fa05||fa63||fa64||fa65,FA04))
            '2009/9/16 MODIFY BY SONIA 加入 CP31
            'modify by sonia +cp12
            '2021/4/21 modify by sonia 解決FF案件請款單重覆計算問題U11002495(CFP-032254),+cp60以判斷FF案件請款單非主要請款案件性質不管收文號有無虧損
            'adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and UPPER(NVL(fa05||fa63||fa64||fa65,FA04)) like UPPER('" & Replace(strCon1, "'", "''") & _
                         "%') and st01(+)=cp14 union " & _
                         "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, NVL(nvl(a1k30, A1K11),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, acc1k0, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and UPPER(NVL(fa05||fa63||fa64||fa65,FA04)) like UPPER('" & Replace(strCon1, "'", "''") & _
                         "%') and st01(+)=cp14 order by CP05 DESC,cp09 DEsc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2024/12/20 +cp53,cp54
            adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,cp60,cp53,cp54 from caseprogress, casepropertymap, fagent, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and UPPER(NVL(fa05||fa63||fa64||fa65,FA04)) like UPPER('" & Replace(strCon1, "'", "''") & _
                         "%') and st01(+)=cp14 union " & _
                         "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,cp60,cp53,cp54 from caseprogress, casepropertymap, fagent, acc1k0, staff, " & _
                         " (SELECT cp60 z1,MIN(cp09) z2 FROM caseprogress WHERE 1=1 AND cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 GROUP BY cp60) z " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and UPPER(NVL(fa05||fa63||fa64||fa65,FA04)) like UPPER('" & Replace(strCon1, "'", "''") & _
                         "%') and st01(+)=cp14 and z1(+)=cp60 order by CP05 DESC,cp09 DEsc", adoTaie, adOpenStatic, adLockReadOnly
            'end 2021/4/21
            'modify by sonia 2021/4/21 解決FF案件請款單重覆計算問題U11002495(CFP-032254),同時取消CP61~CP63,CP87,CP88條件，否則計算目前盈虧會少算
            'adocal.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                        "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, NVL(nvl(a1k30, A1K11),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, acc1k0, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2021/4/29 加CP66,CP67改排序原為CP09 ASC改為CP66 DESC,CP67 DESC,CP09 DESC,U11000603(CFP-031978會抓到AA9041239而非AA9041238)造成最小收文號尚未輸過帳單
            adocal.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,CP66,CP67 from caseprogress, casepropertymap, fagent, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                        "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,CP66,CP67 from caseprogress, casepropertymap, fagent, acc1k0, staff, " & _
                        " (SELECT cp60 z1,MIN(cp09) z2 FROM caseprogress WHERE 1=1 AND cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 GROUP BY cp60) z " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 and z1(+)=cp60 order by CP05 DESC,CP66 DESC,CP67 DESC,CP09 DESC", adoTaie, adOpenStatic, adLockReadOnly
            'end 2021/4/21
            '93.7.16 END
         Else
            '93.7.16 modify by sonia: cp16-cp77 改成 DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))
            'adocase.Open "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' union " & _
            '             "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(a1k30, 0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent, acc1k0 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            'adocal.Open "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' union " & _
            '            "select cp09, CPM04 as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(a1k30, 0) as RecAmount, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), ((nvl(cp16,0) - nvl(cp77,0)) / 1000)) as cp18 from caseprogress, casepropertymap, fagent, acc1k0 where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            '2009/9/16 MODIFY BY SONIA 加入 CP31
            'modify by sonia +cp12
            '2021/4/21 modify by sonia 解決FF案件請款單重覆計算問題U11002495(CFP-032254),+cp60以判斷FF案件請款單非主要請款案件性質不管收文號有無虧損
            'adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, staff " & _
                          " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                          "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, NVL(nvl(a1k30, A1K11),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, acc1k0, staff " & _
                          " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 order by CP05 DESC,cp09 DEsc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2024/12/20 +cp53,cp54
            adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,cp60,cp53,cp54 from caseprogress, casepropertymap, fagent, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                         "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,cp60,cp53,cp54 from caseprogress, casepropertymap, fagent, acc1k0, staff, " & _
                         " (SELECT cp60 z1,MIN(cp09) z2 FROM caseprogress WHERE 1=1 AND cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 GROUP BY cp60) z " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 and z1(+)=cp60 order by CP05 DESC,cp09 DEsc", adoTaie, adOpenStatic, adLockReadOnly
            'end 2021/4/21
            'modify by sonia 2021/4/21 解決FF案件請款單重覆計算問題U11002495(CFP-032254),同時取消CP61~CP63,CP87,CP88條件，否則計算目前盈虧會少算
            'adocal.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                         "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, NVL(nvl(a1k30, A1K11),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N from caseprogress, casepropertymap, fagent, acc1k0, staff " & _
                         " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and ((cp61 <> '" & strCon8 & "' or cp61 is null) and (cp62 <> '" & strCon8 & "' or cp62 is null) and (cp63 <> '" & strCon8 & "' or cp63 is null) and (cp87 <> '" & strCon8 & "' or cp87 is null) and (cp88 <> '" & strCon8 & "' or cp88 is null)) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 order by CP05 DESC,cp09 asc", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2021/4/29 加CP66,CP67改排序原為CP09 ASC改為CP66 DESC,CP67 DESC,CP09 DESC,U11000603(CFP-031978會抓到AA9041239而非AA9041238)造成最小收文號尚未輸過帳單
            adocal.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, nvl(cp16,0) - nvl(cp77,0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp77, 0, nvl(cp18, 0), null, nvl(cp18, 0), DECODE((CP16-CP77),0,0, (CP18-(nvl(cp77,0) / 1000)))) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,CP66,CP67 from caseprogress, casepropertymap, fagent, staff " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and (substr(cp60, 1, 1) = 'E' or cp60 is null) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 union " & _
                        "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, decode(cp09,z2,DECODE(A1K29,'Y',nvl(a1k30,0), nvl(A1K11,0)),0) as RecAmount, decode(cp88, null, decode(cp87, null, decode(cp63, null, decode(cp62, null, decode(cp61, null, 0, 1), 2), 3), 4), 5) as PayCount, decode(cp09,z2,(nvl(a1k11, 0)-nvl(a1k09,0))/1000,0) as cp18,CP31,CP64,cp10,cp12,cp16,cp01,st03 CP14D,st02 CP14N,CP66,CP67 from caseprogress, casepropertymap, fagent, acc1k0, staff, " & _
                        " (SELECT cp60 z1,MIN(cp09) z2 FROM caseprogress WHERE 1=1 AND cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 GROUP BY cp60) z " & _
                        " where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp60 = a1k01 (+) and (substr(cp60, 1, 1) = 'X') and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp05<>19221111 and st01(+)=cp14 and z1(+)=cp60 order by CP05 DESC,CP66 DESC,CP67 DESC,CP09 DESC", adoTaie, adOpenStatic, adLockReadOnly
            'end 2021/4/21
            '93.7.16 END
         End If
   End Select
   Set Adodc1.Recordset = adocase
   
Checking:
   If Err.NUMBER = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  顯示畫面
'
'*************************************************
Public Sub FormShow()
   Combo1.Clear
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   '2012/10/3 MODIFY BY SONIA CFP案加申請人PA26
   'MODIFY BY SONIA 2014/3/21 加特殊出名公司
   'MODIFY BY SONIA 2014/4/1 商標案也加申請人TM23
   'Modified by Morgan 2019/5/16 +申請國家 PA09,
   adoquery.Open "select pa05 as Name1, pa06 as Name2, pa07 as Name3, nvl(na03, na04) as NationName,PA26,PA161 AS COMP,NA01 from patent, nation where pa09 = na01 (+) and pa01 = '" & Text10 & "' and pa02 = '" & Text5 & "' and pa03 = '" & Text7 & "' and pa04 = '" & Text9 & "' union " & _
                 "select tm05 as Name1, tm06 as Name2, tm07 as Name3, nvl(na03, na04) as NationName,TM23 as PA26,TM130 AS COMP,NA01 from trademark, nation where tm10 = na01 (+) and tm01 = '" & Text10 & "' and tm02 = '" & Text5 & "' and tm03 = '" & Text7 & "' and tm04 = '" & Text9 & "' and tm01 <> 'TF' union " & _
                 "select tm05 as Name1, tm06 as Name2, tm07 as Name3, nvl(na03, na04) as NationName,TM23 as PA26,TM130 AS COMP,NA01 from trademark, nation where tm10 = na01 (+) and tm01 = '" & Text10 & "' and tm02 = '" & Text5 & Text7 & "' and tm03 = '" & Text9 & "' and tm04 = '" & Text12 & "' and tm01 = 'TF' union " & _
                 "select lc05 as Name1, lc06 as Name2, lc07 as Name3, nvl(na03, na04) as NationName,'' as PA26,LC48 AS COMP,NA01 from lawcase, nation where lc15 = na01 (+) and lc01 = '" & Text10 & "' and lc02 = '" & Text5 & "' and lc03 = '" & Text7 & "' and lc04 = '" & Text9 & "' union " & _
                 "select hc06 as Name1, '' as Name2, '' as Name3, '' as NationName,'' as PA26,'' AS COMP,'000' NA01  from hirecase where hc01 = '" & Text10 & "' and hc02 = '" & Text5 & "' and hc03 = '" & Text7 & "' and hc04 = '" & Text9 & "' union " & _
                 "select sp05 as Name1, sp06 as Name2, sp07 as Name3, nvl(na03, na04) as NationName,'' as PA26,SP85 AS COMP,NA01 from servicepractice, nation where sp09 = na01 (+) and sp01 = '" & Text10 & "' and sp02 = '" & Text5 & "' and sp03 = '" & Text7 & "' and sp04 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("Name1").Value) = False Then
         Combo1 = "中--" & adoquery.Fields("Name1").Value
         Combo1.AddItem "中--" & adoquery.Fields("Name1").Value
      End If
      If IsNull(adoquery.Fields("Name2").Value) = False Then
         Combo1.AddItem "英--" & adoquery.Fields("Name2").Value
      End If
      If IsNull(adoquery.Fields("Name3").Value) = False Then
         Combo1.AddItem "日--" & adoquery.Fields("Name3").Value
      End If
      If IsNull(adoquery.Fields("NationName").Value) = False Then
         Text1 = adoquery.Fields("NationName").Value
      Else
         Text1 = MsgText(601)
      End If
      
      m_Country = "" & adoquery.Fields("NA01").Value 'Added by Morgan 2019/5/16
      
'******* 此處的所有提醒功能,若和P案有關,則basQuery的PUB_AddNewFBillData也要加*********
      
      '2012/10/3 ADD BY SONIA X63219國立中正大學 及 X43988060國立虎尾科技大學 的CFP案要彈訊息提醒操作人員
      If IsNull(adoquery.Fields("PA26").Value) = False Then
         'MODIFY BY SONIA 2014/4/1 改為所有國外案都要且訊息統一,顏永堅3/25郵件所列大學及其關係企業都要加
         'If Text10 = "CFP" Then
         '   Select Case Left(adoquery.Fields("PA26").Value, 8)
         '      '2013/3/12 MODIFY BY SONIA 加入X6383801中國醫藥大學
         '      Case "X6321900", "X6383801"
         '         MsgBox "客戶為 " & adoquery.Fields("PA26").Value & GetCustomerName(adoquery.Fields("PA26").Value) & " , 憑水單,帳單請款且每一程序只能請款一次 !", , MsgText(5)
         '      Case "X4398806"
         '         MsgBox "客戶為 X43988060國立虎尾科技大學, 憑帳單請款 !", , MsgText(5)
         '   End Select
         'End If
         'modify by sonia 2017/4/10 婉莘要求加X69534彩豐精技  2017/5/15陳德發及郭雅娟要求取消
         'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要,'modify by sonia 2017/10/11 張詠翔要求取消X6014900
         Select Case Left(adoquery.Fields("PA26").Value, 6)
            Case "X44551", "X62079", "X43988", "X63219", "X60498", "X62319", "X62702", "X63838"
               MsgBox "客戶為 " & adoquery.Fields("PA26").Value & GetCustomerName(adoquery.Fields("PA26").Value) & " , 憑水單,帳單請款且每一程序只能請款一次 !", , MsgText(5)
            Case Else
         End Select
         '2014/4/1 END
         'add by sonia 2017/4/14 單獨編號而關係企業不要的, 例:王副總的華碩客戶(X69011010)
         Select Case Left(adoquery.Fields("PA26").Value, 8)
            'modify by sonia 2017/4/27 高國碩要求X60149改為只要母號,關係企業不要
            'modify by sonia 2017/10/11 張詠翔要求取消X6014900
            'modify by sonia 2017/10/26 郭雅娟要求加X60738010國立清華大學
            'modify by sonia 2020/7/20 茹曣加X79919000,且X60738010國立清華大學二者都改訊息
            'Case "X6901101", "X6073801"
            '   MsgBox "客戶為 " & adoquery.Fields("PA26").Value & GetCustomerName(adoquery.Fields("PA26").Value) & " , 憑水單,帳單請款且每一程序只能請款一次 !", , MsgText(5)
            'modify by sonia 2021/3/17 顧服組黃教威加5編號X69365010(長庚醫療財團法人嘉義長庚紀念醫院),X54243070(國立台灣大學),X80847020(財團法人國家實驗研究院),X83983000(盧彥蓓),X83984000(林致廷)
            'Modified by Morgan 2023/2/22 +X69365000、X69365020、X69365050、X69365060、X69365070、X69365080--黃教威
            'Modified by Morgan 2024/5/16 +X82532010、X82504030、X82504040、X69365110、X69365090、X87287010、X87287000;另X38805030從下面移來此處--茹曣
            Case "X6901101", "X6936501", "X5424307", "X8084702", "X8398300", "X8398400", "X6936500", "X6936502", "X6936505", "X6936506", "X6936507", "X6936508", "X8253201", "X8250403", "X8250404", "X6936511", "X6936509", "X8728701", "X8728700", "X3880503"
               MsgBox "客戶為 " & adoquery.Fields("PA26").Value & GetCustomerName(adoquery.Fields("PA26").Value) & " , 憑水單,帳單請款且每一程序只能請款一次 !", , MsgText(5)
            'modify by sonia 2020/11/24 魏裕仁加X43714050  '2021/1/5取消   '2021/4/19再次加入
            'modify by sonia 2021/9/28 顧服組的客戶X38805030資策會--郭雅娟
            'Modified by Morgan 2024/5/16 X38805030移到上面--茹曣
            Case "X6073801", "X7991900", "X4371405"
               MsgBox "客戶為 " & adoquery.Fields("PA26").Value & GetCustomerName(adoquery.Fields("PA26").Value) & " , 憑代理人帳單請款且每一程序只能請款一次，帳單請列印交智權同仁(請程序提醒智權同仁匯率僅可選擇代理人請款日或本所收據日其中一者的當天匯率) !", , MsgText(5)
            'end 2020/7/20
            Case Else
         End Select
         'end 2017/4/14
      End If
      '2012/10/3 END
      'ADD BY SONIA 2014/3/21 特殊出名公司提醒
      If "" & adoquery.Fields("COMP").Value = "J" Then
         MsgBox "此案為智權公司出名案件, 請注意代理人帳單是否有開台一智權公司 !", , MsgText(5)
      End If
      '2014/3/21 END
   Else
      Text1 = MsgText(601)
   End If
   adoquery.Close
End Sub

'*************************************************
'  計算並顯示盈虧
'
'*************************************************
Public Sub Calculate()
Dim strSQL1 As String
Dim strSQL2, StrSQL3 As String  'add by sonia 2021/4/21
 
   strSQL1 = ""
   Select Case Text10
      Case "TF"
         strSQL1 = strSQL1 & " and axf03 = '" & Text10 & Text5 & Text7 & Text9 & Text12 & "'"
         'add by sonia 2021/4/21
         strSQL2 = " instr(ax214,'" & Text10 & Text5 & Text7 & "')=1 "
         StrSQL3 = " instr(a1p17,'" & Text10 & Text5 & Text7 & "')=1 "
         'end 2021/4/21
      Case Else
         strSQL1 = strSQL1 & " and axf03 = '" & Text10 & Text5 & Text7 & Text9 & "'"
         'add by sonia 2021/4/21
         strSQL2 = " ax214 = '" & Text10 & Text5 & Text7 & Text9 & "' "
         StrSQL3 = " a1p17 = '" & Text10 & Text5 & Text7 & Text9 & "' "
         'end 2021/4/21
   End Select
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select sum(nvl(a1520, 0)) from acc151, acc150 where axf01 = a1501 and axf02 = '000000000'" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) Then
         Text8 = "0"
      Else
         Text8 = adoquery.Fields(0).Value
      End If
   Else
      Text8 = "0"
   End If
   adoquery.Close
   Text2 = MsgText(601)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   If adoaccsum.State = adStateOpen Then
      adoaccsum.Close
   End If
   adoaccsum.CursorLocation = adUseClient
   '2005/8/26 MODIFY BY SONIA
   'adoaccsum.Open "select a2103 from acc210 where a2102 = '" & strTitle & "' and a2101 = (select max(a2101) from acc210 where a2102 = '" & strTitle & "' and a2101 <= " & Val(ACDate(ServerDate)) & ")", adoTaie, adOpenStatic, adLockReadOnly
   adoaccsum.Open "select a2103 from acc210 where a2102 = '" & Frmacc2150.Combo1 & "' and a2101 = (select max(a2101) from acc210 where a2102 = '" & strTitle & "' and a2101 <= " & strSrvDate(2) & ")", adoTaie, adOpenStatic, adLockReadOnly
   '2005/8/26 END
   If adoaccsum.RecordCount <> 0 Then
      Text2 = Val(Text2) - (Val(Text11) * Val(adoaccsum.Fields("a2103").Value))
   Else
      Text2 = Val(Text2) - Val(Text11)
   End If
   adoaccsum.Close
   Set adoloop = adocal.Clone
   adoloop.MoveFirst
   Do While adoloop.EOF = False
'      If adoloop.Fields("cp09").Value = Adodc1.Recordset.Fields("cp09").Value Then
         '2009/9/17 MODIFY BY SONIA RecAmount=0者不可再減點數,否則負負為正CFP-022018答辯
         If Val(adoloop.Fields("RecAmount").Value) = 0 Then
            Text2 = Val(Text2)
         Else
            Text2 = Val(Text2) + adoloop.Fields("RecAmount").Value
            '2009/9/17 自外移入
            If IsNull(adoloop.Fields("cp18").Value) = False Then
               Text2 = Val(Text2) - (Val(adoloop.Fields("cp18").Value) * 1000)
            End If
         End If
         If adoquery.State = adStateOpen Then
            adoquery.Close
         End If
         adoquery.CursorLocation = adUseClient
         '2005/12/2 MODIFY BY SONIA已付改抓AXF04*ACC190之A1906實際付款匯率,未付才抓AXF15
         'adoquery.Open "select a1505, a1502, sum(decode(a1507, null, nvl(axf15, 0), 0)) as PayAmount from acc151, acc150 where axf01 = a1501 and axf02 = '" & adoloop.Fields("cp09").Value & "' group by a1505, a1502", adoTaie, adOpenStatic, adLockReadOnly
         'Modify by Morgan 2007/
         'Modify by Morgan 2007/3/13 當未結匯(a1906=0)時也抓axf15
         '2014/3/7 modify by sonia 加入抓抵帳匯率ACC1G0(CFP-019761的U10302128)
         'adoquery.Open "select a1505, a1502, sum(decode(a1507, null, nvl(NVL(AXF04*decode(A1906,0,null,A1906),axf15), 0), 0)) as PayAmount from acc151, acc150, ACC190 where axf01 = a1501 AND AXF01=A1902(+) and axf02 = '" & adoloop.Fields("cp09").Value & "' group by a1505, a1502", adoTaie, adOpenStatic, adLockReadOnly
         'modify by sonia 2022/7/27 抵帳單要加回T-124366
         'adoquery.Open "select a1505, a1502, sum(decode(a1507, null, nvl(NVL(AXF04*decode(A1906,0,NULL,A1906),NVL(AXF04*decode(A1G03,0,null,A1G03),AXF15)), 0), 0)) as PayAmount from acc151, acc150, ACC190, ACC1G0 " & _
                     "where axf01 = a1501 AND AXF01=A1902(+) AND A1512=A1G01(+) and axf02 = '" & adoloop.Fields("cp09").Value & "' group by a1505, a1502", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "select a1505,a1502,AXF04,a.A1906,A1G03,NVL(AXF04*decode(a.A1906,0,null,a.A1906),NVL(AXF04*decode(A1G03,0,null,A1G03),0)) as PayAmount,nvl(axf04, 0) as FpayAmount, AXF02 from acc151, acc150,ACC190 a,ACC1G0 " & _
                    "where axf02<>'000000000' and axf01 = a1501(+) and a1507 is null AND AXF01=a.A1902(+) AND A1512=A1G01(+) and axf02 = '" & adoloop.Fields("cp09").Value & "' AND nvl(axf04, 0)>0 " & _
                    "union select a1605,a1602,AXG04,nvl(A1906,a1g03),0,nvl(AXg04*nvl(A1906,a1g03),0)*-1 as PayAmount,nvl(axg04, 0)*-1 as FpayAmount, AXg02 from acc160,acc161,acc190,acc1i0 c,acc1i0 d,acc1g0 " & _
                    "where axg01 = a1601(+) and axg01=a1902(+) and axg02 = '" & adoloop.Fields("cp09").Value & "' AND nvl(axg04, 0)>0 and a1607=c.a1i03(+) and a1605=c.a1i05(+) and a1607=d.a1i03(+) and nvl(c.a1i01,d.a1i01)=a1g01(+)", adoTaie, adOpenStatic, adLockReadOnly
         'end 2022/7/27
         '2005/12/2 END
         Do While adoquery.EOF = False
            If IsNull(adoquery.Fields("PayAmount").Value) = False Then
'               If adoaccsum.State = adStateOpen Then
'                  adoaccsum.Close
'               End If
'               adoaccsum.CursorLocation = adUseClient
'               adoaccsum.Open "select a2103 from acc210 where a2102 = '" & adoquery.Fields("a1505").Value & "' and a2101 = (select max(a2101) from acc210 where a2102 = '" & adoquery.Fields("a1505").Value & "' and a2101 <= " & adoquery.Fields("a1502").Value & ")", adoTaie, adOpenStatic, adLockReadOnly
'               If adoaccsum.RecordCount <> 0 Then
'                  Text2 = Val(Text2) - (Val(adoquery.Fields("PayAmount").Value) * Val(adoaccsum.Fields("a2103").Value))
'               Else
                  Text2 = Val(Text2) - Val(adoquery.Fields("PayAmount").Value)
'               End If
'               adoaccsum.Close
            End If
            adoquery.MoveNext
         Loop
         adoquery.Close
'      End If
      adoloop.MoveNext
   Loop
   'edit by nickc 2005/11/02
'   adoquery.CursorLocation = adUseClient
   '新案件才扣安全基金
   'cancel by sonia 2021/4/21 每一案件都要算一次安全基金故取消
   'If Adodc1.Recordset.Fields("cp31") = "Y" Then   'add by sonia 2016/11/21 CFP-029039 拆二個收文號輸會重覆扣安全基金
      Select Case Text10
         '92.3.3 加PA08
         Case "TF"
   'edit by nickc 2005/11/02
   '         adoquery.Open "select pa09 as NationNo, pa01 as SystemKind,PA08 from patent where pa01 = '" & Text10 & "' and pa02 = '" & Text5 & Text7 & "' and pa03 = '" & Text9 & "' and pa04 = '" & Text12 & "' union " & _
                          "select tm10 as NationNo, tm01 as SystemKind,'' AS PA08 from trademark where tm01 = '" & Text10 & "' and tm02 = '" & Text5 & Text7 & "' and tm03 = '" & Text9 & "' and tm04 = '" & Text12 & "' union " & _
                          "select lc15 as NationNo, lc01 as SystemKind,'' AS PA08 from lawcase where lc01 = '" & Text10 & "' and lc02 = '" & Text5 & Text7 & "' and lc03 = '" & Text9 & "' and lc04 = '" & Text12 & "' union " & _
                          "select '' as NationNo, hc01 as SystemKind,'' AS PA08 from hirecase where hc01 = '" & Text10 & "' and hc02 = '" & Text5 & Text7 & "' and hc03 = '" & Text9 & "' and hc04 = '" & Text12 & "' union " & _
                          "select sp09 as NationNo, sp01 as SystemKind,'' AS PA08 from servicepractice where sp01 = '" & Text10 & "' and sp02 = '" & Text5 & Text7 & "' and sp03 = '" & Text9 & "' and sp04 = '" & Text12 & "'", adoTaie, adOpenStatic, adLockReadOnly
               Text2 = Val(Text2) - GetFloatPrepareCase(Text10.Text, Text5.Text & Text7.Text, Text9.Text, Text12.Text)
         Case Else
   'edit by nickc 2005/11/02
   '         adoquery.Open "select pa09 as NationNo, pa01 as SystemKind,PA08 from patent where pa01 = '" & Text10 & "' and pa02 = '" & Text5 & "' and pa03 = '" & Text7 & "' and pa04 = '" & Text9 & "' union " & _
                          "select tm10 as NationNo, tm01 as SystemKind,'' AS PA08 from trademark where tm01 = '" & Text10 & "' and tm02 = '" & Text5 & "' and tm03 = '" & Text7 & "' and tm04 = '" & Text9 & "' union " & _
                          "select lc15 as NationNo, lc01 as SystemKind,'' AS PA08 from lawcase where lc01 = '" & Text10 & "' and lc02 = '" & Text5 & "' and lc03 = '" & Text7 & "' and lc04 = '" & Text9 & "' union " & _
                          "select '' as NationNo, hc01 as SystemKind,'' AS PA08 from hirecase where hc01 = '" & Text10 & "' and hc02 = '" & Text5 & "' and hc03 = '" & Text7 & "' and hc04 = '" & Text9 & "' union " & _
                          "select sp09 as NationNo, sp01 as SystemKind,'' AS PA08 from servicepractice where sp01 = '" & Text10 & "' and sp02 = '" & Text5 & "' and sp03 = '" & Text7 & "' and sp04 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
               Text2 = Val(Text2) - GetFloatPrepareCase(Text10.Text, Text5.Text, Text7.Text, Text9.Text)
      End Select
      'edit by nickc 2005/11/02
   'End If    'end 2016/11/21     'cancel by sonia 2021/4/21 每一案件都要算一次安全基金故取消

'   If adoquery.RecordCount <> 0 Then
'      Select Case adoquery.Fields("SystemKind").Value
'         Case "CFP"
'            If IsNull(adoquery.Fields("NationNo").Value) Then
'               Text2 = Val(Text2) - 5000
'            Else
'               If Mid(adoquery.Fields("NationNo").Value, 1, 3) = "101" Then
'                  Text2 = Val(Text2) - 3000
'               Else
'                  If adoquery.Fields("PA08").Value = "2" And (Mid(adoquery.Fields("NationNo").Value, 1, 3) = "231" Or Mid(adoquery.Fields("NationNo").Value, 1, 3) = "011") Then
'                     Text2 = Val(Text2) - 3000
'                  Else
'                     Text2 = Val(Text2) - 5000
'                  End If
'               End If
'            End If
'         Case "TS", "S", "LA"  '92.6.19 ADD BY SONIA  93.5.10 加入 LA
'         '93.1.2 ADD BY SONIA
'         Case "T"
'            If Mid(adoquery.Fields("NationNo").Value, 1, 3) = "020" Then
'               Text2 = Val(Text2) - 2000
'            Else
'               Text2 = Val(Text2) - 3000
'            End If
'         '93.1.2 END
'         Case Else
'            Text2 = Val(Text2) - 3000
'      End Select
'   Else
'      Text2 = Val(Text2) - 3000
'   End If
   
   'add by sonia 2021/4/21 加財務其他支出Text13(直接由傳票輸入之規費),判斷AX212有無結餘二字若放在語法中當沒資料時會有點慢instr(ax212,'結餘')=0
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "SELECT ax201,ax202,ax203,ax206-ax207 AMT,ax212 FROM acc021,(SELECT * FROM acc1p0 WHERE " & StrSQL3 & " AND a1p02<>'L') WHERE " & strSQL2 & " AND ax201=a1p01(+) AND ax202=a1p22(+) and ax205=a1p05(+) and ax214=a1p17(+) and ax206=a1p07(+) and ax207=a1p08(+) AND a1p04 IS NULL AND ax205 LIKE '2201%' ", adoTaie, adOpenStatic, adLockReadOnly
   Text13 = 0
   Do While adoaccsum.EOF = False
      If InStr("" & adoaccsum.Fields("ax212"), "結餘") = 0 Then
         Text13 = Text13 + adoaccsum.Fields("AMT").Value
      End If
      adoaccsum.MoveNext
   Loop
   adoaccsum.Close
   'end 2021/4/21
   
   'modify by sonia 2021/4/21 再減財務其他支出Text13
   'Text2 = Val(Text2) - Val(Text8)
   Text2 = Val(Text2) - Val(Text8) - Val(Text13)
   'end 2021/4/21
   Text2 = Format(Text2, FAmount)
'edit by nickc 2005/11/02
'   adoquery.Close
End Sub


