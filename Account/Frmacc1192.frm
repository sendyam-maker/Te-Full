VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc1192 
   AutoRedraw      =   -1  'True
   Caption         =   "本所案號退費/銷帳資料輸入"
   ClientHeight    =   5004
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9408
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5004
   ScaleWidth      =   9408
   Begin VB.CommandButton cmdCancel 
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
      Height          =   396
      Left            =   8040
      TabIndex        =   9
      Top             =   3912
      Width           =   1224
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
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
      Left            =   4920
      TabIndex        =   7
      Top             =   3900
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
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
      Left            =   1704
      TabIndex        =   5
      Top             =   4248
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
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
      Left            =   1704
      TabIndex        =   3
      Top             =   3888
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
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
      Left            =   1704
      TabIndex        =   0
      Top             =   4608
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc1192.frx":0000
      Height          =   3504
      Left            =   24
      TabIndex        =   1
      Top             =   216
      Width           =   9348
      _ExtentX        =   16489
      _ExtentY        =   6181
      _Version        =   393216
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
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
         DataField       =   "case"
         Caption         =   "本所案號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "a1u03"
         Caption         =   "總收文號"
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
         DataField       =   "cp16"
         Caption         =   "應收金額"
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
      BeginProperty Column03 
         DataField       =   "cp17"
         Caption         =   "應收規費"
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
      BeginProperty Column04 
         DataField       =   "cp73"
         Caption         =   "已收服務費"
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
      BeginProperty Column05 
         DataField       =   "cp74"
         Caption         =   "已收規費"
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
      BeginProperty Column06 
         DataField       =   "a1u08"
         Caption         =   "退服務費金額"
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
         DataField       =   "a1u10"
         Caption         =   "退規費金額"
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
      BeginProperty Column08 
         DataField       =   "a1u07"
         Caption         =   "銷帳服務費"
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
         DataField       =   "a1u09"
         Caption         =   "銷帳規費"
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
         DataField       =   "a1u06"
         Caption         =   "扣繳金額退費"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   315
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   947.906
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   924.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1128.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   924.095
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1091.906
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1188.284
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   1235.906
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnWidth     =   1284.095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   336
      Left            =   24
      Top             =   96
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   614
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "本畫面會鎖住表列的收文資料，請勿停留太久！"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   4272
      TabIndex        =   10
      Top             =   4608
      Width           =   5040
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "扣繳金額退費"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   3900
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "銷帳金額"
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
      Left            =   264
      TabIndex        =   6
      Top             =   4608
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "退規費金額"
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
      Left            =   264
      TabIndex        =   4
      Top             =   4248
      Width           =   1452
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   24
      Top             =   3648
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "退服務費金額"
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
      Left            =   264
      TabIndex        =   2
      Top             =   3888
      Width           =   1452
   End
End
Attribute VB_Name = "Frmacc1192"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adocheck As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim douSAmount As Double
Dim douTAmount As Double
Dim douCAmount As Double
Dim douPAmount As Double
Dim mSeqNo As String 'Added by Lydia 2021/08/27 暫存檔之序號
Public frmCall As Form 'Added by Morgan 2024/11/12

Private Sub cmdCancel_Click()
   frmCall.Tag = "N"
   Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   'Added by Morgan 2024/11/13
   If KeyCode = vbKeyEscape Then
      Unload Me
   Else
   'end 2024/11/13
      KeyEnter KeyCode
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Morgan 2024/11/12
   'Me.Icon = LoadPicture(strIcoPath)
   'strFormName = Name
   'Me.Width = 9500
   'Me.Height = 5500
   'Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   'Image1 = LoadPicture(strBackPicPath2)
   'sglWidth = Image1.Width
   'sglHeight = Image1.Height
   'For intX = 0 To Int(ScaleWidth / sglWidth)
   '    For intY = 0 To Int(ScaleHeight / sglHeight)
   '        PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
   '    Next
   'Next
   PUB_InitForm Me, 9500, 5500, strBackPicPath2
   'end 2024/11/12
   
   Text1 = strCon8
   Text2 = strCon9
   Text3 = strCon10
   Text4 = Val(strCon5) * (-1)
   OpenTable
   StatusClear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim douService As Double
Dim douTax As Double
Dim strMan As String
Dim strCust As String
Dim strRemark As String
Dim strSerialNo As String
Dim strSalesNo As String
Dim strAccNo As String
Dim strYes As String
Dim strDept As String   '93.11.25 ADD BY SONIA
   
   'Added by Morgan 2024/11/12
   If frmCall.Tag = "N" Then
      GoTo ExitPoint
   End If
   'end 2024/11/12
   
   'Added by Morgan 2016/1/14
   '扣繳金額退費檢查
   With Adodc1.Recordset
   .MoveFirst
   Do While Not .EOF
      If Val(DataGrid1.Columns(10).Value) <> 0 Then
         If CheckValue(10) = False Then
            Cancel = True
            Exit Do
         End If
      End If
      
      'Added by Morgan 2025/7/8
      If Frmacc1190.adoacc0k0.Fields("a0k11") = "J" And Left(DataGrid1.Columns(0).Value, 3) = "ACS" Then
         If Val(DataGrid1.Columns(8).Value) > 0 Then
            If Val(DataGrid1.Columns(9).Value) = 0 Then
               MsgBox "智權公司ACS案件在銷帳時,不能沒有輸銷帳規費(稅)!!", vbCritical, "智權公司ACS案銷帳規費(稅)檢查"
               Cancel = True
               Exit Do
            Else
               intI = Round(Val(DataGrid1.Columns(8).Value) * 0.05)
               If Val(DataGrid1.Columns(9).Value) <> intI Then
                  If MsgBox("銷帳規費(稅)錯誤!!應該為 " & intI & " (" & Val(DataGrid1.Columns(8).Value) & "x0.05)。" & vbCrLf & vbCrLf & "是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2, "智權公司ACS案銷帳規費(稅)檢查") = vbNo Then
                     Cancel = True
                     Exit Do
                  End If
               End If
            End If
         End If
      End If
      'end 2025/7/8
      .MoveNext
   Loop
   .MoveFirst 'Added by Morgan 2024/11/12
   End With
   If Cancel = CInt(True) Then
      'tool3_enabled 'Removed by Morgan 2024/11/12
      DataGrid1.SetFocus
      'strExitControl = MsgText(602) 'Removed by Morgan 2024/11/12
      Exit Sub
   End If
   'end 2016/1/14
   
   If adocheck.State = 1 Then adocheck.Close
   adocheck.CursorLocation = adUseClient
   adocheck.Open "select sum(a1u08), sum(a1u10), sum(a1u07), sum(a1u09), sum(a1u06) from acc1u0 where a1u01 = '" & strCon6 & "' and a1u02 = '" & strCon7 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      If IsNull(adocheck.Fields(0).Value) Then
         douSAmount = 0
      Else
         douSAmount = adocheck.Fields(0).Value
      End If
      If IsNull(adocheck.Fields(1).Value) Then
         douTAmount = 0
      Else
         douTAmount = adocheck.Fields(1).Value
      End If
      If IsNull(adocheck.Fields(2).Value) Then
         douCAmount = 0
      Else
         douCAmount = adocheck.Fields(2).Value
      End If
      If IsNull(adocheck.Fields(3).Value) Then
         douCAmount = douCAmount
      Else
         douCAmount = douCAmount + Val(adocheck.Fields(3).Value)
      End If
      If IsNull(adocheck.Fields(4).Value) Then
         douPAmount = 0
      Else
         '93.1.6 MODIFY BY SONIA
         'douAmount = Val(adocheck.Fields(4).Value)
         douPAmount = Val(adocheck.Fields(4).Value)
         '93.1.6 END
      End If
   Else
      douSAmount = 0
      douTAmount = 0
      douCAmount = 0
      douPAmount = 0
   End If
   adocheck.Close
   If Val(Text1) <> douSAmount Then
      MsgBox MsgText(117), , MsgText(5)
      'tool3_enabled 'Removed by Morgan 2024/11/12
      Cancel = True
      DataGrid1.SetFocus
      'strExitControl = MsgText(602) 'Removed by Morgan 2024/11/12
      Exit Sub
   End If
   If Val(Text2) <> douTAmount Then
      MsgBox MsgText(118), , MsgText(5)
      'tool3_enabled 'Removed by Morgan 2024/11/12
      Cancel = True
      DataGrid1.SetFocus
      'strExitControl = MsgText(602) 'Removed by Morgan 2024/11/12
      Exit Sub
   End If
   If Val(Text3) <> douCAmount Then
      MsgBox MsgText(119), , MsgText(5)
      'tool3_enabled 'Removed by Morgan 2024/11/12
      Cancel = True
      DataGrid1.SetFocus
      'strExitControl = MsgText(602) 'Removed by Morgan 2024/11/12
      Exit Sub
   End If
   If Val(Text4) <> douPAmount Then
      MsgBox MsgText(217), , MsgText(5)
      'tool3_enabled 'Removed by Morgan 2024/11/12
      Cancel = True
      DataGrid1.SetFocus
      'strExitControl = MsgText(602) 'Removed by Morgan 2024/11/12
      Exit Sub
   End If
   
   'Added by Morgan 2014/1/20
   If FormSave = False Then
      'tool3_enabled 'Removed by Morgan 2024/11/12
      Cancel = True
      DataGrid1.SetFocus
      'strExitControl = MsgText(602) 'Removed by Morgan 2024/11/12
      Exit Sub
   End If
   'end 2014/1/20
   
   strItemNo = MsgText(601)
   'tool1_enabled
   
   'Modified by Morgan 2014/8/7
   'Frmacc1190.Enabled = True
   'modify by sonia 2017/6/8 J公司只要銷退都要顯示Frmacc1194,不管有沒有開發票E10613585
   'If Frmacc1190.Text4 <> "" Then
   If Frmacc1190.adoacc0k0.Fields("a0k11") = "J" And Frmacc1190.Text3 = "3" Then
      'Modified by Morgan 2024/11/12
      'Frmacc1194.Show
      frmCall.Tag = "F"
      'end 2024/11/12
   'Added by Morgan 2022/6/21
   '案源都要顯示分錄畫面
   ElseIf Frmacc1190.m_LOS02 <> "" And Frmacc1190.Text3 = "3" Then
      'Modified by Morgan 2024/11/12
      'Frmacc1194.Show
      frmCall.Tag = "F"
      'end 2024/11/12
   'end 2022/6/21
   Else
      Frmacc1190.Enabled = True
      'Added by Lydia 2020/04/10 若有工作點數分配(ACC1N0且A1N02='3')的資料必須強制進入工作點數分配frm071021且帶出原分配資料，且必須點數合計正確才能離開；若有一筆以上，則連續進行工作點數分配。
      'Mark by Lydia 2021/08/24 (保留)是在主畫面存檔時若法務案有分配工作點數，檢查剩餘點數與工作點數不符時進入工作點數分配畫面
      'strExc(0) = " select distinct(a1n01) a1n01 from acc1u0,acc1n0 where a1u01='" & strCon6 & "' and a1u02='" & strCon7 & "' and a1u03=a1n01(+) and a1n02='3' "
      'intI = 1
      'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      'If intI = 1 Then
      '     strExc(0) = RsTemp.GetString(adClipString, , , ",")
      '     Set frm071021.m_PrevForm = Frmacc1190
      '     frm071021.m_bolPrev = True
      '     frm071021.m_KeyList = strExc(0)
      '     Frmacc1190.Enabled = False
      '     frm071021.Show
      'End If
      'end 2020/04/10
      'end 2021/08/24
   End If
   'end 2014/8/7
   
ExitPoint:

   Set Frmacc1192 = Nothing
  
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
Dim intQ As Integer, rsQuery As New ADODB.Recordset  'Added by Lydia 2021/08/27

On Error GoTo Checking
   adoTaie.Execute "delete from acc1u0 where a1u01 = '" & strCon6 & "' and a1u02 = '" & strCon7 & "'"
   'Modified by Morgan 2011/10/12 考慮拆收據改抓 acc0j0
   'adoTaie.Execute "insert into acc1u0 select '" & strCon6 & "', '" & strCon7 & "', cp09, 0, 0, 0, 0, 0, 0, 0 from caseprogress where cp60 = '" & strCon7 & "'"
   strSql = "insert into acc1u0 select '" & strCon6 & "', '" & strCon7 & "', a0j01, 0, 0, 0, 0, 0, 0, 0" & _
      " from acc0j0 where a0j13 = '" & strCon7 & "' and a0j02='" & Frmacc1190.cboCaseNo & "'"
   adoTaie.Execute strSql, intI
   
   'Add by Morgan 2011/10/5 考慮拆收據已收金要改抓1U0資料,又配合 DATAGRID 更新故要先寫暫存
   adoTaie.Execute "delete ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   adoTaie.Execute "insert into ACCTMP08(T01,T02,T03,T04,T05,T14) SELECT A1U02 T01,A1U03 T02,sum(a1u04) T03,sum(a1u05) T04,'" & Me.Name & "' T05,'" & strUserNum & "' T14 from acc1u0 where a1u01 <>'" & strCon6 & "' and a1u02 = '" & strCon7 & "' group by A1U02,a1u03"
   
   If adoadodc1.State = adStateOpen Then adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   'Modified by Morgan 2011/10/12 考慮拆收據改抓 acc0j0
   'adoadodc1.Open "select cp01||cp02||cp03||cp04 as case, a1u03, nvl(cp16, 0) as cp16, nvl(cp17, 0) as cp17, nvl(cp73, 0) as cp73, nvl(cp74, 0) as cp74, a1u07, a1u08, a1u09, a1u10, a1u01, a1u02, a1u06 from caseprogress, acc1u0 where cp09 = a1u03 and a1u01 = '" & strCon6 & "' and a1u02 = '" & strCon7 & "' order by cp01||cp02||cp03||cp04 asc, a1u03 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   'Modified by Lydia 2021/08/27 配合O12；改成先丟暫存檔
   'adoadodc1.Open "select cp01||cp02||cp03||cp04 as case, a1u03, nvl(a0j09, 0)+nvl(a0j10,0) as cp16, nvl(a0j10, 0) as cp17, nvl(T03,0) as cp73, nvl(T04,0) as cp74, a1u07, a1u08, a1u09, a1u10, a1u01, a1u02, a1u06 from acc1u0, caseprogress, acc0j0, ACCTMP08 where a1u01 = '" & strCon6 & "' and a1u02 = '" & strCon7 & "' and cp09(+) = a1u03 and a0j01(+)=a1u03 and a0j13(+)=a1u02 and t01(+)=a1u02 and t02(+)=a1u03 and t05(+)='" & Me.Name & "' and T14(+)='" & strUserNum & "' order by cp01||cp02||cp03||cp04 asc, cp09 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   strSql = "select cp01||cp02||cp03||cp04 as case, a1u03, nvl(a0j09, 0)+nvl(a0j10,0) as cp16, nvl(a0j10, 0) as cp17, nvl(T03,0) as cp73, nvl(T04,0) as cp74, a1u07, a1u08, a1u09, a1u10, a1u01, a1u02, a1u06 from acc1u0, caseprogress, acc0j0, ACCTMP08 where a1u01 = '" & strCon6 & "' and a1u02 = '" & strCon7 & "' and cp09(+) = a1u03 and a0j01(+)=a1u03 and a0j13(+)=a1u02 and t01(+)=a1u02 and t02(+)=a1u03 and t05(+)='" & Me.Name & "' and T14(+)='" & strUserNum & "' order by cp01||cp02||cp03||cp04 asc, cp09 asc"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
        Set adoadodc1 = PUB_CreateRecordset(rsQuery, , , , Me.Name, mSeqNo)
        '暫存檔欄位=原本欄位
        'R001=Case, R002=a1u03, R003=CP16, R004=CP17, r005=cp73, r006=cp74, r007=a1u07, r008=a1u08,
        'r009=a1u09, r010=a1u10, r011=a1u01, r012=a1u02, r013=a1u06
        'Datagrid1可編輯欄位
        '6=a1u08, 7=a1u10, 8=a1u07, 9=a1u09, 10=a1u06
   End If
   Set rsQuery = Nothing
   'end 2021/08/27
   
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
On Error GoTo Checking
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   With DataGrid1
      If CheckValue(ColIndex) = False Then
         Exit Sub
      End If
      'Ken 91/11/01 小計改為銷帳規費
      'Adodc1.Recordset.Fields("a1u09").Value = (Val(.Columns(6).Value) + Val(.Columns(7).Value) - Val(.Columns(8).Value)) * (-1)
      Adodc1.Recordset.UpdateBatch
      'Added by Lydia 2021/08/27 從Datagrid1回寫到acc1u0
      'Datagrid1可編輯欄位 6=a1u08, 7=a1u10, 8=a1u07, 9=a1u09, 10=a1u06
      strSql = "update acc1u0 set a1u08=" & CNULL(Val(.Columns(6).Value), True) & ", a1u10=" & CNULL(Val(.Columns(7).Value), True) & _
                  " ,a1u07=" & CNULL(Val(.Columns(8).Value), True) & ", a1u09=" & CNULL(Val(.Columns(9).Value), True) & ", a1u06=" & CNULL(Val(.Columns(10).Value), True) & _
                  " where a1u02='" & strCon7 & "' and a1u01='" & strCon6 & "' and a1u03='" & DataGrid1.Columns(1).Value & "' "
      adoTaie.Execute strSql, intI
      'end 2021/08/27
   End With
   
   Exit Sub
Checking:
   MsgBox Err.Description, , MsgText(5)
   SendKeys "{0}"
End Sub

Private Sub DataGrid1_GotFocus()
   Dim intCounter As Integer
   
   'Added by Morgan 2024/11/12
   DataGrid1.col = 4: SendKeys "{RIGHT}"
   Exit Sub
   'end 2024/11/12
   
   DataGrid1.col = 0
   For intCounter = 1 To 6
      SendKeys "{RIGHT}"
   Next intCounter
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Select Case DataGrid1.col
            Case 6
               SendKeys "{RIGHT}"
            Case 7
               SendKeys "{RIGHT}"
            Case 8
               SendKeys "{RIGHT}"
            Case 9
               SendKeys "{RIGHT}"
            Case 10
               SendKeys "{DOWN}"
               SendKeys "{LEFT}"
               SendKeys "{LEFT}"
               SendKeys "{LEFT}"
               SendKeys "{LEFT}"
         End Select
   End Select
End Sub
'Add by Morgan 2006/6/12
'檢查輸入金額是否正確
Private Function CheckValue(p_iColIndex As Integer) As Boolean
   'Modified by Morgan 2016/1/14 扣繳也要檢查
   'Modified by Lydia 2021/08/27 配合O12：抓暫存檔
   'strSql = "select sum(a1u07) u07, sum(a1u08) u08, sum(a1u09) u09, sum(a1u10) u10, sum(a1u06) u06,max(a0j07) Tx from acc1u0,acc0j0 where a1u02='" & strCon7 & "' and a1u01<> '" & strCon6 & "' and a1u03='" & DataGrid1.Columns(1).Value & "' and a0j01(+)=a1u03 and a0j13(+)=a1u02"
   'Modified by Lydia 2021/10/13 因為DataGrid1_AfterColUpdate已更新acc1u0
   'strSql = "select sum(to_number(r007)) u07, sum(to_number(r008)) u08, sum(to_number(r009)) u09, sum(to_number(r010)) u10, sum(to_number(r013)) u06, " & _
               "max(a0j07) tx from  rdatafactory,acc0j0 where Formname='" & Me.Name & "' And Id='" & strUserNum & "' And Seqno='" & mSeqNo & "' " & _
               "and r012='" & strCon7 & "' and r011<> '" & strCon6 & "' and r002='" & DataGrid1.Columns(1).Value & "' and a0j01(+)=r002 and a0j13(+)=r012 "
   strSql = "select sum(a1u07) u07, sum(a1u08) u08, sum(a1u09) u09, sum(a1u10) u10, sum(a1u06) u06,max(a0j07) Tx from acc1u0,acc0j0 where a1u02='" & strCon7 & "' and a1u01<> '" & strCon6 & "' and a1u03='" & DataGrid1.Columns(1).Value & "' and a0j01(+)=a1u03 and a0j13(+)=a1u02"
   
   With RsTemp
      If .State = adStateOpen Then .Close
      .CursorLocation = adUseClient
      .Open strSql, adoTaie, adOpenForwardOnly, adLockReadOnly
      If Not (.EOF And .BOF) Then
         Select Case p_iColIndex
            Case 6 '退費服務費
               If Val(DataGrid1.Columns(p_iColIndex).Value) > Val(DataGrid1.Columns(4).Value) - Val("" & .Fields("u08")) Then
                  MsgBox "退費服務費不可大於已收服務費(扣除已退費)", , MsgText(5)
               Else
                  CheckValue = True
               End If
            Case 7 '退費規費
               If Val(DataGrid1.Columns(p_iColIndex).Value) > Val(DataGrid1.Columns(5).Value) - Val("" & .Fields("u10")) Then
                  MsgBox "退費規費不可大於已收規費(扣除已退費)", , MsgText(5)
               Else
                  CheckValue = True
               End If
            Case 8 '銷帳服務費
               If Val(DataGrid1.Columns(p_iColIndex).Value) > Val(DataGrid1.Columns(2).Value) - DataGrid1.Columns(3).Value - Val("" & .Fields("u07")) Then
                  MsgBox "銷帳服務費不可大於應收服務費(扣除已銷帳)", , MsgText(5)
               Else
                  CheckValue = True
               End If
            Case 9 '銷帳規費
               If Val(DataGrid1.Columns(p_iColIndex).Value) > Val(DataGrid1.Columns(3).Value) - Val("" & .Fields("u09")) Then
                  MsgBox "銷帳規費不可大於應收規費(扣除已銷帳)", , MsgText(5)
               Else
                  CheckValue = True
               End If
            'Added by Morgan 2016/1/14
            Case 10 '扣繳金額退費
               'Modified by Morgan 2021/11/2 小數點要四捨五入 Ex:I11000599
               strExc(1) = Round(0.1 * (Val(DataGrid1.Columns(6).Value) + IIf(.Fields("Tx") = "Y", Val(DataGrid1.Columns(7).Value), 0)))
               'Modified by Morgan 2024/11/12
               'If Val(DataGrid1.Columns(p_iColIndex).Value) > 0 Or (-1 * Val(DataGrid1.Columns(p_iColIndex).Value) > Val(strExc(1)) And -1 * Val(DataGrid1.Columns(p_iColIndex).Value) <> Val("" & .Fields("u06"))) Then
               If Val(DataGrid1.Columns(p_iColIndex).Value) > 0 Then
                  MsgBox "退費時,稅額應先扣抵,請輸入負值!!", , MsgText(5)
               ElseIf (-1 * Val(DataGrid1.Columns(p_iColIndex).Value) > Val(strExc(1)) And -1 * Val(DataGrid1.Columns(p_iColIndex).Value) <> Val("" & .Fields("u06"))) Then
               'end 2024/11/12
                  MsgBox "扣繳金額退費輸入錯誤!!", , MsgText(5)
               Else
                  CheckValue = True
               End If
            
            'Add by Morgan 2006/9/11
            Case Else
               CheckValue = True
               
         End Select
      End If
   End With
End Function

'Added by Morgan 2014/1/20
'從 Form_Unload 抽來
Private Function FormSave() As Boolean
Dim douService As Double
Dim douTax As Double
Dim strMan As String
Dim strCust As String
Dim strRemark As String
Dim strSerialNo As String
Dim strSalesNo As String
Dim strAccNo As String
Dim strYes As String
Dim strDept As String   '93.11.25 ADD BY SONIA
'Added by Morgan 2014/1/2
Dim strCompNo As String '公司別
Dim strMaxFee As String '最大規費
Dim strMaxFeeNo As String '最大規費科目項次

'Added by Morgan 2014/1/20
   'adoTaie.BeginTrans
On Error GoTo ErrHnd
'end 2014/1/20

   'Added by Morgan 2014/1/21
   strCompNo = "" & Frmacc1190.adoacc0k0("a0k11")
   'Modified by Morgan 2021/11/2 加L公司改寫法
   'If strCompNo <> "J" Then strCompNo = "1"
   If strCompNo < "A" Then strCompNo = "1"
   'end 2021/11/2
   'end 2014/1/21
   
   'Add by Morgan 2011/10/20 考慮拆收據情形改批次更新
   strSql = "update caseprogress set (cp77,cp78)=(select nvl(sum(a1u07),0)+nvl(sum(a1u09),0)" & _
      ",nvl(sum(a1u08),0)+nvl(sum(a1u10),0) from acc1u0 where a1u03=cp09)" & _
      " where cp09 in (select a1u03 from acc1u0 where a1u01='" & strCon6 & "')"
   adoTaie.Execute strSql, intI
   
   strSql = "update caseprogress set cp79 = nvl(cp16, 0) - nvl(cp75, 0) - nvl(cp77, 0) + nvl(cp78, 0)" & _
      " where cp09 in (select a1u03 from acc1u0 where a1u01='" & strCon6 & "')"
   adoTaie.Execute strSql, intI
   
   'Added by Morgan 2012/4/19 更新進度檔已扣繳金額
   strSql = "update caseprogress set cp76=(select nvl(sum(a1u06),0) from acc1u0 where a1u03=cp09) " & _
      " where cp09 in (select a1u03 from acc1u0 where a1u01='" & strCon6 & "')"
   adoTaie.Execute strSql, intI
   'end 2012/4/19

If Frmacc1190.adoacc0k0.Fields("a0k11") <> "J" Then 'Added by Morgan 2014/1/3 排除J公司
      '更新 acc1v0 資料
      strExc(0) = "select a1u02,a1u03,a1v01,a1v02 from acc1u0 u1,acc1v0 where a1u01='" & strCon6 & "'" & _
         " and a1v01(+)=a1u03 and a1v02(+)=a1u02" & _
         " and exists(select * from acc1u0 u2 where u2.a1u02(+)=u1.a1u02 and u2.a1u03=u1.a1u03 and substr(u2.a1u01,1,1)='F')"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            '沒有 1v0 資料 Insert(應該不需要新增,因為銷部份時1v0不會刪除)
            If IsNull(.Fields("a1v01")) Then
               'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
               strSql = "insert into acc1v0(a1v01,a1v02,a1v03,a1v04,a1v05,a1v06,a1v07,a1v09,a1v12,a1v13,a1v17,a1v18)" & _
                  " select a0j01 a1v01,a0j13 a1v02,a0k11 a1v03" & _
                  ",0.1*(nvl(a0j09,0)-nvl(x2,0)+decode(a0j07,'Y',nvl(a0j10,0)-nvl(x3,0),0)) a1v04" & _
                  ",nvl(a0k13,'N') a1v05,nvl(x1,0) a1v06" & _
                  ",0.1*(nvl(a0j09,0)-nvl(x2,0)+decode(a0j07,'Y',nvl(a0j10,0)-nvl(x3,0),0))-nvl(x1,0) a1v07" & _
                  ",a0k16 a1v09,getcp10desc(cp01,cp10,a0j04) a1v12,na03 a1v13,y1 a1v17,decode(sign(x1),1,'1') a1v18" & _
                  " From acc0j0,acc0k0,(select a1u02,a1u03,sum(a1u06) x1,sum(a1u07) x2,sum(a1u09) x3 from acc1u0" & _
                  " where a1u02='" & .Fields("a1u02") & "' and a1u03='" & .Fields("a1u03") & "'" & _
                  " group by a1u02,a1u03) x,(select a0m02,max(a0m03) y1 from acc0m0 where a0m02='" & .Fields("a1u02") & "'" & _
                  " group by a0m02) y,caseprogress,nation" & _
                  " where  a0j01='" & .Fields("a1u03") & "' and a0j13='" & .Fields("a1u02") & "'" & _
                  " and a0k01(+)=a0j13 and a1u03(+)=a0j01 and a1u02(+)=a0j13 and a0m02(+)=a0j13" & _
                  " and cp09(+)=a0j01 and na01(+)=a0j04"
                  
            '有 1v0 資料 Update
            Else
               strSql = "update acc1v0 set (a1v04,a1v06,a1v07)=(" & _
                  " select 0.1*(nvl(max(a0j09),0)-nvl(sum(a1u07),0)+decode(max(a0j07),'Y',nvl(max(a0j10),0)-nvl(sum(a1u09),0),0)) a1v04" & _
                  ",nvl(sum(a1u06),0) a1v06" & _
                  ",0.1*(nvl(max(a0j09),0)-nvl(sum(a1u07),0)+decode(max(a0j07),'Y',nvl(max(a0j10),0)-nvl(sum(a1u09),0),0))-nvl(sum(a1u06),0) a1v07" & _
                  " from acc0j0,acc1u0 where a0j01=a1v01 and a0j13=a1v02 and a1u02(+)=a1v02 and a1u03(+)=a1v01)" & _
                  " where a1v01='" & .Fields("a1u03") & "' and a1v02='" & .Fields("a1u02") & "'"
                  
            End If
            adoTaie.Execute strSql, intI
            .MoveNext
         Loop
         End With
      End If
      'end 2011/10/20
End If 'Added by Morgan 2014/1/3
   
'Remove by Morgan 2011/10/20 移到上面改批次更新
'--舊程式已刪除--


   'Modify by Morgan 2011/10/18 呼叫本畫面時有設定 strItemNo 為傳票號此處不必重抓,但若要重抓則也應該重設傳票號才會一致
'--舊程式已刪除--

   If strItemNo <> "" And UCase(strItemNo) <> "NULL" Then
      strYes = "'" & MsgText(602) & "'"
   Else
      strYes = "null"
   End If
   'end 2011/10/18
   
   'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
   strSql = "delete from acc1p0 where a1p02 = 'Z' and a1p04 = '" & strCon6 & "' and (a1p07 <> 0 or (a1p07 = 0 and a1p08 = 0))"
   If Frmacc1190.m_KeepItem <> "" Then strSql = strSql & " and instr('" & Frmacc1190.m_KeepItem & "',a1p03)=0" 'Added by Morgan 2015/8/11 保留銷項稅額相關科目
   adoTaie.Execute strSql
   
   If adoquery.State = 1 Then adoquery.Close
   adoquery.CursorLocation = adUseClient
   'Modified by Morgan 2011/10/14 考慮拆收據情形
   'adoquery.Open "select * from caseprogress, acc0k0, acc0j0, acc1u0 where cp60 = a0k01 and cp09 = a0j01 (+) and cp09 = a1u03 and a1u01 = '" & strCon6 & "'", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Morgan 2011/12/27 取消 a0j20
   adoquery.Open "select a.*,b.*,c.*,d.*,getcp10desc(cp01,cp10,a0j04) cp10N" & _
      " from acc1u0 a,caseprogress b, acc0k0 c, acc0j0 d where a1u01='" & strCon6 & "'" & _
      " and cp09(+)=a1u03 and a0k01(+)=a1u02 and a0j13(+)=a1u02 and a0j01(+)=a1u03", adoTaie, adOpenStatic, adLockReadOnly
   douService = 0
   douTax = 0
   Do While adoquery.EOF = False
      If IsNull(adoquery.Fields("a0k20").Value) = False Then
         If adoaccsum.State = 1 Then adoaccsum.Close
         adoaccsum.CursorLocation = adUseClient
         adoaccsum.Open "select sn01 from salesno where sn02 = '" & adoquery.Fields("a0k20").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoaccsum.RecordCount <> 0 Then
            If IsNull(adoaccsum.Fields("sn01").Value) Then
               strMan = ""
            Else
               strMan = adoaccsum.Fields("sn01").Value
            End If
         End If
         adoaccsum.Close
      Else
         strMan = ""
      End If
      If IsNull(adoquery.Fields("a0k04").Value) = False Then
         strCust = MidB(adoquery.Fields("a0k04").Value, 1, 16)
      Else
         strCust = ""
      End If
      'Modify by Morgan 2004/4/7
      'strRemark = strMan & "/" & strCust & "/" & strCon6
      '借方摘要
      'Modified by Morgan 2011/12/27 取消 a0j20
      'strRemark = strMan & "/" & Left(strCust, 4) & "/" & "" & adoquery.Fields("a0j20").Value & "/" & strCon6
      strRemark = strMan & "/" & Left(strCust, 4) & "/" & "" & adoquery.Fields("cp10N").Value & "/" & strCon6
      If IsNull(adoquery.Fields("a1u08").Value) = False And adoquery.Fields("a1u08").Value <> 0 Then
         'Modify by Morgan 2006/9/12 項次借貸方都要算才不會重複
         'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & strCon6 & "'", 3)
         If adoaccsum.State = 1 Then adoaccsum.Close
         adoaccsum.CursorLocation = adUseClient
         adoaccsum.Open "select cpm11 from casepropertymap where cpm01 = '" & adoquery.Fields("cp01").Value & "' and cpm02 = '" & adoquery.Fields("cp10").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoaccsum.RecordCount <> 0 Then
            If IsNull(adoaccsum.Fields("cpm11").Value) = False Then
               'modify by sonia 2021/1/29 加傳本所案號以判別FCP,FCT英日文組
               'If AccNoToSalesNo(adoaccsum.Fields("cpm11").Value) = "" Then
               If AccNoToSalesNo(adoaccsum.Fields("cpm11").Value, adoquery.Fields("a0j02").Value) = "" Then
                  strSalesNo = IIf(IsNull(adoquery.Fields("a0k20").Value), "", adoquery.Fields("a0k20").Value)
               Else
                  'modify by sonia 2021/1/29 加傳本所案號以判別FCP,FCT英日文組
                  'strSalesNo = AccNoToSalesNo(adoaccsum.Fields("cpm11").Value)
                  strSalesNo = AccNoToSalesNo(adoaccsum.Fields("cpm11").Value, adoquery.Fields("a0j02").Value)
               End If
            Else
               strSalesNo = ""
            End If
            If IsNull(adoquery.Fields("a1u08").Value) = False And adoquery.Fields("a1u08").Value <> 0 Then
               douService = Val(adoquery.Fields("a1u08").Value)
            Else
               douService = 0
            End If
            '2007/6/11 MODIFY BY SONIA 改判斷非台灣
            'If adoquery.Fields("a0j04").Value = "020" Then
            If adoquery.Fields("a0j04").Value <> "000" And (Mid(adoquery.Fields("cp01").Value, 1, 1) = "P" Or Mid(adoquery.Fields("cp01").Value, 1, 1) = "T") Then
               If Mid(adoquery.Fields("cp01").Value, 1, 1) = "P" Then
                  strAccNo = "411103"
               Else
                  strAccNo = "410103"
               End If
            Else
               If IsNull(adoaccsum.Fields("cpm11").Value) Then
                  strAccNo = "XXX"
               Else
                  strAccNo = adoaccsum.Fields("cpm11").Value
               End If
            End If
            '93.11.25 ADD BY SONIA
            If IsNull(adoquery.Fields("cp01").Value) Then
               strDept = "null"
            Else
               'MODIFY BY SONIA 2016/1/5
               'Select Case Mid(strAccNo, 1, 4)
               '   Case "4101", "4151"
               '      strDept = "T"
               '   Case "4111"
               '      strDept = "P"
               '   Case "4121"
               '      strDept = "CFT"
               '   Case "4172"
               '      If adoaccsum.Fields("cpm11").Value = "417202" Then
               '         strDept = "T"
               '      Else
               '         strDept = "FCT"
               '      End If
               '   Case "4131"
               '      strDept = "CFP"
               '   Case "4141"
               '      strDept = "L"
               '   Case "4171"
               '      strDept = "FCP"
               '   Case "4181"
               '      strDept = "L"
               '   Case "4161"
               '      strDept = "FCL"
               '   Case Else
               '      strDept = "TOT"
               'End Select
               If Left(strAccNo, 1) = "4" Then
                  strDept = PUB_GETAccNODept(strAccNo, strDept)
               Else
                  strDept = "TOT"
               End If
               'END 2016/1/5
            End If
            '93.11.25 END
            If douService <> 0 Then
               '93.11.25 MODIFY BY SONIA
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('1', 'Z', '" & strSerialNo & "', '" & strCon6 & "', '" & strAccNo & "', '" & IIf(IsNull(adoquery.Fields("cp01").Value), "", adoquery.Fields("cp01").Value) & "', " & douService & ", 0, null, null, null, null, null, '" & strRemark & "', '" & IIf(IsNull(adoquery.Fields("a0k03").Value), "", adoquery.Fields("a0k03").Value) & "', '" & strSalesNo & "', '" & adoquery.Fields("cp01").Value & adoquery.Fields("cp02").Value & adoquery.Fields("cp03").Value & adoquery.Fields("cp04").Value & "', " & Val(FCDate(strTitle)) & ", null, null, 0, " & strItemNo & ", null, null, 0, null, " & strYes & ")"
               'Modified by Morgan 2014/8/7 會有J公司 a1p01 改用變數
               'ADD BY SONIA 2016/1/5 105年起法務收入改其他部門收入(傳CP09以判斷案件性質及收文人員)
               If Val(FCDate(strTitle)) >= 1050101 And (Left(strAccNo, 4) = "4141" Or Left(strAccNo, 4) = "4161" Or Left(strAccNo, 4) = "4181") Then
                  'Modified by Morgan 2016/2/1 傳票號要去掉單引號
                  InsertLawACC1P0 strCompNo, "Z", strSerialNo, strCon6, strAccNo, strDept, Val(douService), 0, "", "", "", "", "", ChgSQL(strRemark), IIf(IsNull(adoquery.Fields("a0k03").Value), "", adoquery.Fields("a0k03").Value), strSalesNo, adoquery.Fields("cp01").Value & adoquery.Fields("cp02").Value & adoquery.Fields("cp03").Value & adoquery.Fields("cp04").Value, Val(FCDate(strTitle)), "", "", 0, Replace(strItemNo, "'", ""), "", "", 0, "", Replace(strYes, "'", ""), "", "", adoquery.Fields("a0j01")
               Else
               'END 2016/1/5
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & strCon6 & "', '" & strAccNo & "', '" & strDept & "', " & douService & ", 0, null, null, null, null, null, '" & strRemark & "', '" & IIf(IsNull(adoquery.Fields("a0k03").Value), "", adoquery.Fields("a0k03").Value) & "', '" & strSalesNo & "', '" & adoquery.Fields("cp01").Value & adoquery.Fields("cp02").Value & adoquery.Fields("cp03").Value & adoquery.Fields("cp04").Value & "', " & Val(FCDate(strTitle)) & ", null, null, 0, " & strItemNo & ", null, null, 0, null, " & strYes & ")"
               End If  'add by sonia 2016/1/5
               '93.11.25 END
               'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
               adoTaie.Execute "update acc1p0 set a1p15 = '" & IIf(IsNull(adoquery.Fields("a0k03").Value), "", adoquery.Fields("a0k03").Value) & "', a1p16 = '" & IIf(IsNull(adoquery.Fields("a0k20").Value), "", adoquery.Fields("a0k20").Value) & "', a1p17 = '" & adoquery.Fields("cp01").Value & adoquery.Fields("cp02").Value & adoquery.Fields("cp03").Value & adoquery.Fields("cp04").Value & "' where a1p02 = 'Z' and a1p04 = '" & strCon6 & "' and a1p08 <> 0"
            End If
         End If
         adoaccsum.Close
      End If
      
      If IsNull(adoquery.Fields("a1u10").Value) = False And adoquery.Fields("a1u10").Value <> 0 Then
         'Modify by Morgan 2006/9/12 項次借貸方都要算才不會重複
         'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & strCon6 & "'", 3)
         If adoaccsum.State = 1 Then adoaccsum.Close
         adoaccsum.CursorLocation = adUseClient
         adoaccsum.Open "select cpm12 from casepropertymap where cpm01 = '" & adoquery.Fields("cp01").Value & "' and cpm02 = '" & adoquery.Fields("cp10").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoaccsum.RecordCount <> 0 Then
            If IsNull(adoquery.Fields("a1u10").Value) = False And adoquery.Fields("a1u10").Value <> 0 Then
               douTax = Val(adoquery.Fields("a1u10").Value)
            Else
               douTax = 0
            End If
            '2007/6/11 MODIFY BY SONIA 改判斷非台灣
            'If adoquery.Fields("a0j04").Value = "020" Then
            If adoquery.Fields("a0j04").Value <> "000" And (Mid(adoquery.Fields("cp01").Value, 1, 1) = "P" Or Mid(adoquery.Fields("cp01").Value, 1, 1) = "T") Then
               If Mid(adoquery.Fields("cp01").Value, 1, 1) = "P" Then
                  strAccNo = "220112"
               Else
                  strAccNo = "220111"
               End If
            Else
               If IsNull(adoaccsum.Fields("cpm12").Value) Then
                  strAccNo = "XXX"
               Else
                  strAccNo = adoaccsum.Fields("cpm12").Value
               End If
            End If
            '93.11.25 ADD BY SONIA
            If IsNull(adoquery.Fields("cp01").Value) Then
               strDept = "null"
            Else
               'MODIFY BY SONIA 2016/1/5
               'Select Case Mid(strAccNo, 1, 4)
               '   Case "4101", "4151"
               '      strDept = "T"
               '   Case "4111"
               '      strDept = "P"
               '   Case "4121"
               '      strDept = "CFT"
               '   Case "4172"
               '      If adoaccsum.Fields("cpm11").Value = "417202" Then
               '         strDept = "T"
               '      Else
               '         strDept = "FCT"
               '      End If
               '   Case "4131"
               '      strDept = "CFP"
               '   Case "4141"
               '      strDept = "L"
               '   Case "4171"
               '      strDept = "FCP"
               '   Case "4181"
               '      strDept = "L"
               '   Case "4161"
               '      strDept = "FCL"
               '   Case Else
               '      strDept = "TOT"
               'End Select
               If Left(strAccNo, 1) = "4" Then
                  strDept = PUB_GETAccNODept(strAccNo, strDept)
               Else
                  strDept = "TOT"
               End If
               'END 2016/1/5
            End If
            '93.11.25 END
            If douTax <> 0 Then
               '93.11.25 MODIFY BY SONIA
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('1', 'Z', '" & strSerialNo & "', '" & strCon6 & "', '" & strAccNo & "', '" & MsgText(55) & "', " & douTax & ", 0, null, null, null, null, null, '" & strRemark & "', '" & IIf(IsNull(adoquery.Fields("a0k03").Value), "", adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(adoquery.Fields("a0k20").Value), "", adoquery.Fields("a0k20").Value) & "', '" & adoquery.Fields("cp01").Value & adoquery.Fields("cp02").Value & adoquery.Fields("cp03").Value & adoquery.Fields("cp04").Value & "',  " & Val(FCDate(strTitle)) & ", null, null, 0, " & strItemNo & ", null, null, 0, null, " & strYes & ")"
               'Modified by Morgan 2014/8/7 會有J公司 a1p01 改用變數
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & strCon6 & "', '" & strAccNo & "', '" & IIf(strAccNo = "610103", strDept, MsgText(55)) & "', " & douTax & ", 0, null, null, null, null, null, '" & strRemark & "', '" & IIf(IsNull(adoquery.Fields("a0k03").Value), "", adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(adoquery.Fields("a0k20").Value), "", adoquery.Fields("a0k20").Value) & "', '" & adoquery.Fields("cp01").Value & adoquery.Fields("cp02").Value & adoquery.Fields("cp03").Value & adoquery.Fields("cp04").Value & "',  " & Val(FCDate(strTitle)) & ", null, null, 0, " & strItemNo & ", null, null, 0, null, " & strYes & ")"
               '93.11.25 END
               'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
               adoTaie.Execute "update acc1p0 set a1p15 = '" & IIf(IsNull(adoquery.Fields("a0k03").Value), "", adoquery.Fields("a0k03").Value) & "', a1p16 = '" & IIf(IsNull(adoquery.Fields("a0k20").Value), "", adoquery.Fields("a0k20").Value) & "', a1p17 = '" & adoquery.Fields("cp01").Value & adoquery.Fields("cp02").Value & adoquery.Fields("cp03").Value & adoquery.Fields("cp04").Value & "' where a1p02 = 'Z' and a1p04 = '" & strCon6 & "' and a1p08 <> 0"
            End If
         End If
         adoaccsum.Close
      End If
      adoquery.MoveNext
   Loop
   adoTaie.Execute "delete from acc1u0 where a1u01 = '" & strCon6 & "' and a1u02 = '" & strCon7 & "' and a1u04 = 0 and a1u05 = 0 and a1u07 = 0 and a1u08 = 0 and a1u09 = 0 and a1u10 = 0"
   
   If Val(Text1) <> 0 Or Val(Text2) <> 0 Or Val(Text4) <> 0 Then 'Added by Morgan 2025/7/23 有退費才要跑
   
      'Added by Morgan 2025/7/8
      '有銷項稅額 2119時用最大規費扣(參照frmacc1190)
      intI = 1
      strSql = "select a1p07 from acc1p0 where a1p02 = 'Z' and a1p04 = '" & strCon6 & "' and a1p05='2119' and a1p07>0"
      Set adoquery = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         douTax = adoquery("a1p07")
         strSql = "select a1p03 from acc1p0 where a1p02 = 'Z' and a1p04 = '" & strCon6 & "' and a1p05 like '2201%' and a1p07>" & douTax & " order by a1p07 desc,a1p03"
         Set adoquery = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strSql = "update acc1p0 set a1p07=a1p07-" & douTax & _
            " where a1p02 = 'Z' and a1p04 = '" & strCon6 & "' and a1p03='" & adoquery("a1p03") & "'"
            adoTaie.Execute strSql, intI
         End If
         douTax = 0
      End If
      'end 2025/7/8
      
   
      'Add by Morgan 2006/9/29 重排項次
      Dim iTotItem As Integer '總項次數
      Dim iEmptItem As Integer '前面跳過的項次數
      Dim iDebItem As Integer '貸方項次數
      '抓總項目數-->iTotItem
      'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
      'Modified by Morgan 2025/7/8
      'strSql = "update acc1p0 set a1p03=a1p03 where a1p02 = 'Z' and a1p04 = '" & strCon6 & "'"
      'adoTaie.Execute strSql, iTotItem
      intI = 1
      strSql = "select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & strCon6 & "'"
      Set adoquery = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         iTotItem = Val("" & adoquery(0))
      End If
      'end 2025/7/8
   
      '抓前面跳過的項次數-->iTotItem
      'Removed by Morgan 2025/7/8
      'strSql = "update acc1p0 set a1p03=a1p03 where a1p02 = 'Z' and a1p04 = '" & strCon6 & "' and a1p03+0<=" & iTotItem
      'adoTaie.Execute strSql, intI
      'iEmptItem = iTotItem - intI
      'end 2025/7/8
   
      '把貸方搬到後面,並抓貸方項目數-->iDebItem
      strSql = "update acc1p0 set a1p03=a1p03+" & iTotItem & " where a1p02 = 'Z' and a1p04 = '" & strCon6 & "' and a1p08>0"
      adoTaie.Execute strSql, iDebItem
   
   
      '更正項次為從1開始
      'Modified by Morgan 2025/7/8
      'iEmptItem = iEmptItem + iDebItem
      'strSql = "update acc1p0 set a1p03=lpad(a1p03-" & iEmptItem & ",3,'0') where a1p02 = 'Z' and a1p04 = '" & strCon6 & "'"
      strSql = "update acc1p0 a set a1p03=(select lpad(count(*),3,'0') from acc1p0 b where a1p02=a.a1p02 and a1p04=a.a1p04 and a1p03+0<=a.a1p03+0) where a1p02 = 'Z' and a1p04 = '" & strCon6 & "'"
      'end 2025/7/8
      adoTaie.Execute strSql, intI
      'end 2006/9/19
   End If

   'Add By Sindy 2010/6/18
   Dim dblsumA0K06 As Double, dblsumA0K07 As Double
   Dim dblsumA1U07 As Double, dblsumA1U09 As Double
   dblsumA0K06 = 0: dblsumA0K07 = 0
   dblsumA1U07 = 0: dblsumA1U09 = 0
   
   If adoaccsum.State = 1 Then adoaccsum.Close
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select nvl(a0k06,0),nvl(a0k07,0) from acc0k0 where a0k01 = '" & strCon7 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0)) = False Then
         dblsumA0K06 = adoaccsum.Fields(0)
      End If
      If IsNull(adoaccsum.Fields(1)) = False Then
         dblsumA0K07 = adoaccsum.Fields(1)
      End If
   End If
   
   If adoaccsum.State = 1 Then adoaccsum.Close
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(nvl(a1u07,0)),sum(nvl(a1u09,0)) from acc1u0 where a1u02 = '" & strCon7 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0)) = False Then
         dblsumA1U07 = adoaccsum.Fields(0)
      End If
      If IsNull(adoaccsum.Fields(1)) = False Then
         dblsumA1U09 = adoaccsum.Fields(1)
      End If
   End If
   adoaccsum.Close
   If (dblsumA0K06 = dblsumA1U07) And (dblsumA0K07 = dblsumA1U09) Then
      'Modified by Lydia 2023/12/12 排除Z.確定不印
      'adoTaie.Execute "update acc0k0 set a0k32=null where a0k01 = '" & strCon7 & "'"
      adoTaie.Execute "update acc0k0 set a0k32=null where a0k01 = '" & strCon7 & "' and nvl(a0k32,'Y') <>'Z' "
   'Added by Morgan 2013/9/13
   '若銷部份帳款且已取消收文時,上可列印
   Else
      'Modified by Lydia 2023/12/12 排除Z.確定不印 +and a0k32<>'Z'
      strSql = "update acc0k0 set a0k32=null where a0k01 = '" & strCon7 & "' and a0k32 is not null and a0k32<>'Z' and not exists(select * from caseprogress where cp60=a0k01 and cp57 is null)"
      adoTaie.Execute strSql, intI
   'end 2013/9/13
   End If
   '2010/6/18 End
   
'Added by Morgan 2014/1/20
'Modified by Morgan 2015/5/7 規則要一致改用共用函數
'   strSql = "update acc0k0 set  a0k36=NULL, a0k37=NULL where a0k01='" & strCon7 & "'"
'   adoTaie.Execute strSql, intI
   PUB_UpdateReceiptStatus strCon7
'end 2015/5/7
   'adoTaie.CommitTrans 'Removed by Morgan 2024/11/12
   FormSave = True
   Exit Function

ErrHnd:
   'adoTaie.RollbackTrans 'Removed by Morgan 2024/11/12
   MsgBox Err.Description
End Function
