VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071021 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "法務工作點數分配"
   ClientHeight    =   4836
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8784
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4836
   ScaleWidth      =   8784
   Begin VB.CommandButton CmdNext 
      Caption         =   "下一筆"
      Height          =   345
      Left            =   5100
      TabIndex        =   26
      Top             =   570
      Width           =   765
   End
   Begin VB.CommandButton CmdPrev 
      Caption         =   "上一筆"
      Height          =   345
      Left            =   4320
      TabIndex        =   25
      Top             =   570
      Width           =   765
   End
   Begin VB.CommandButton Command10 
      Height          =   300
      Left            =   2300
      Picture         =   "frm071021.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CommandButton Command9 
      Caption         =   "點數重新分配"
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
      Left            =   6930
      TabIndex        =   24
      Top             =   540
      Width           =   1770
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4155
      Left            =   90
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   630
      Width           =   8610
      _ExtentX        =   15177
      _ExtentY        =   7324
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "工作點數"
      TabPicture(0)   =   "frm071021.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Shape1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtST02"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DataGrid1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Adodc1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtSum"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtA1N03"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtA1N04"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtA1N05"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7125
         Picture         =   "frm071021.frx":011E
         Style           =   1  '圖片外觀
         TabIndex        =   6
         ToolTipText     =   "清除畫面"
         Top             =   3315
         Width           =   550
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7725
         Picture         =   "frm071021.frx":09E8
         Style           =   1  '圖片外觀
         TabIndex        =   7
         ToolTipText     =   "取消"
         Top             =   3315
         Width           =   550
      End
      Begin VB.TextBox txtA1N05 
         Alignment       =   1  '靠右對齊
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
         Left            =   2670
         MaxLength       =   8
         TabIndex        =   3
         Top             =   3495
         Width           =   945
      End
      Begin VB.TextBox txtA1N04 
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
         Left            =   270
         MaxLength       =   6
         TabIndex        =   2
         Top             =   3495
         Width           =   972
      End
      Begin VB.TextBox txtA1N03 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3495
         Width           =   1125
      End
      Begin VB.CommandButton Command5 
         Height          =   300
         Left            =   4845
         Picture         =   "frm071021.frx":1052
         Style           =   1  '圖片外觀
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3510
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   2520
         TabIndex        =   18
         Top             =   2730
         Width           =   1005
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   315
         Left            =   360
         Top             =   1590
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
         Height          =   2145
         Left            =   180
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   450
         Width           =   8295
         _ExtentX        =   14626
         _ExtentY        =   3789
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         BackColor       =   -2147483624
         HeadLines       =   1
         RowHeight       =   16
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
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "a1n04"
            Caption         =   "承辦人代號"
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
            DataField       =   "st02"
            Caption         =   "姓名"
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
            DataField       =   "a1n05"
            Caption         =   "點數"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "a1n03"
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
         BeginProperty Column04 
            DataField       =   "cpm03"
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
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1175.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1031.811
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1272.189
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin MSForms.TextBox txtST02 
         Height          =   315
         Left            =   1260
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3495
         Width           =   1335
         VariousPropertyBits=   671107099
         BackColor       =   14737632
         Size            =   "2355;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   915
         Left            =   135
         Top             =   3120
         Width           =   8340
      End
      Begin VB.Label Label14 
         Alignment       =   2  '置中對齊
         BackStyle       =   0  '透明
         Caption         =   "點數"
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
         Left            =   2670
         TabIndex        =   22
         Top             =   3240
         Width           =   945
      End
      Begin VB.Label Label13 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "承辦人"
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
         Left            =   1065
         TabIndex        =   21
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label6 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "收文號"
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
         Left            =   3975
         TabIndex        =   20
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "點數合計"
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
         Left            =   1485
         TabIndex        =   19
         Top             =   2745
         Width           =   900
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "離開"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7605
      TabIndex        =   23
      Top             =   90
      Width           =   1095
   End
   Begin VB.TextBox txtPts 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   6300
      TabIndex        =   14
      Top             =   120
      Width           =   1212
   End
   Begin VB.TextBox txtCP09 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1125
      MaxLength       =   9
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
   Begin VB.TextBox txtCP01 
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   3690
      TabIndex        =   11
      Top             =   120
      Width           =   492
   End
   Begin VB.TextBox txtCP02 
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   4170
      TabIndex        =   10
      Top             =   120
      Width           =   852
   End
   Begin VB.TextBox txtCP03 
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   5010
      TabIndex        =   9
      Top             =   120
      Width           =   252
   End
   Begin VB.TextBox txtCP04 
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   5250
      TabIndex        =   8
      Top             =   120
      Width           =   372
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "點數"
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
      Left            =   5820
      TabIndex        =   15
      Top             =   150
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收 文 號"
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
      Left            =   165
      TabIndex        =   13
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
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
      Left            =   2730
      TabIndex        =   12
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "frm071021"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/31 法務系統的工作點數分配功能先上線(110/9/1)
'Memo by Lydia 2021/08/24 Form2.0已修改; txtST02、DataGrid1改字型=新細明體-ExtB
'Memo by Lydia 2020/04/10 取消內外法之檔案維護新增之法務工作點數分配功能；改在案件進度檔維護第一畫面若instr(cp01,'L')>0則增加工作點數分配按鈕，進入frm071021。
'Create by Lydia 2015/06/01 法務工作點數分配
Option Explicit

Dim adoacc1n0 As ADODB.Recordset
Dim rsTmp1 As New ADODB.Recordset
Dim strTmpSQL As String
Dim intR As Integer

Dim m_bolAddNew As Boolean
Dim m_A1N03_CPM03 As String '收文號的案件性質
Dim cmdState As Integer

Dim m_bolQuery As Boolean '是否為查詢
Dim m_EditMode As String '是否編輯

Dim bolActive As Boolean '只執行一次

Public m_PrevForm As Form  '前一畫面
Public m_bolPrev As Boolean '是否為外部呼叫
Public m_KeyList As String '收文號: 用,區隔
Dim ArrCP09 As Variant  '連續輸入工作點數的收文號
Dim nMax As Integer, nPos As Integer

'下一筆
Private Sub cmdNext_Click()
   If nPos = nMax Then
      MsgBox MsgText(8), , MsgText(5)
   Else
      If ChkInputPoint() = True Then
         nPos = nPos + 1
         txtCP09.Text = ArrCP09(nPos)
         Call Command10_Click
      End If
   End If
End Sub

'上一筆
Private Sub CmdPrev_Click()
   If nPos = 0 Then
      MsgBox MsgText(7), , MsgText(5)
   Else
      If ChkInputPoint() = True Then
         nPos = nPos - 1
         txtCP09.Text = ArrCP09(nPos)
         Call Command10_Click
      End If
   End If
End Sub

Private Sub Command1_Click()
   AdodcDelete adoacc1n0
   AdodcClear
   DataGrid1.Refresh
   SumShow
End Sub

Private Sub Command2_Click()
   AdodcClear
   txtA1N04.SetFocus
End Sub

'放大鏡：查詢收文號 (已不顯示)
Private Sub Command5_Click()
   Dim bCancel As Boolean
   
   'Remove by Lydia 2020/04/20
'   strTmpSQL = "select '',cp09,sqldatet(cp05) cp05,decode(cpm03,'（無）',cpm04,cpm03) cpm03,st02,cp13" & _
'      " from caseprogress,casepropertymap,staff where cp09='" & txtCP09 & "'" & _
'      " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp14"
'   intR = 1
'   Set rsTmp1 = ClsLawReadRstMsg(intR, strTmpSQL)
'   If intR = 1 Then
'      Set Frmacc21h4.grdDataList.Recordset = rsTmp1
'      Set Frmacc21h4.fmParent = Me
'      Frmacc21h4.Show vbModal
'      strFormName = Me.Name
'      If Me.Tag <> "" Then
'         txtA1N03 = Me.Tag
'      End If
'      txtA1N03.SetFocus
'   End If
   'end 2020/04/20
End Sub

'離開
Private Sub Command8_Click()
   If nPos < nMax Then '若不是最後一筆，自動往下一筆
       Call cmdNext_Click
   Else
       If m_bolQuery Or (Val(txtSum) = 0 And txtSum.Tag = txtSum.Text) Then
          Unload Me
       Else
          If ChkInputPoint = True Then
               If FormSave Then
                   Unload Me
               End If
          End If
       End If
   End If
End Sub

Private Function ChkInputPoint() As Boolean
  ChkInputPoint = False
  If m_bolQuery Or (Val(txtSum) = 0 And txtSum.Tag = txtSum.Text) Then
  Else
      If Val(txtPts) > 0 And (Val(txtPts) <> Val(txtSum)) Then
           '保留: 輸入點數與預計點數不符時，不可離開。
           'If MsgBox("輸入點數與收文點數不符，是否要繼續輸入？", vbYesNo + vbDefaultButton2) = vbYes Then
           '   SSTab1.Tab = 0
           'End If
           MsgBox "輸入點數與收文點數不符！", vbCritical, "資料檢核"
           Exit Function
      Else
           ChkInputPoint = True
      End If
  End If
  ChkInputPoint = True
End Function

'點數重新分配
Private Sub Command9_Click()
   If txtCP09.Tag <> txtCP09.Text Then
      Call Command10_Click
   End If
   
   If Val(txtPts) = 0 Then
        MsgBox "無點數可供分配!", vbExclamation
   Else
        If MsgBox("系統將清除目前分配並依照規則重新分配，是否確定要繼續？", vbYesNo + vbDefaultButton2) = vbYes Then
           'Modified by Lydia 2020/04/10 改成共用模組
           'Get_PointAutoassign (txtCP09)
           If PUB_GetLawPointAuto(txtCP09, False, True) = True Then
           End If
           OpenTable '畫面重整只能放在彈訊息的後面
           'end 2020/04/10
        End If
   End If
End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   AdodcShow
End Sub

Private Sub Form_Activate()
  
   If m_KeyList <> "" And bolActive = False Then
      strTmpSQL = Replace(GetAddStr(m_KeyList), "'", "") '去掉多餘的資料
      ArrCP09 = Empty
      ArrCP09 = Split(strTmpSQL, ",")
      Me.txtCP09.Text = ArrCP09(0)
      nPos = 0
      nMax = UBound(ArrCP09)
      If nMax > 0 Then  '顯示上一筆/下一筆
          CmdPrev.Visible = True
          CmdNext.Visible = True
      Else
          CmdPrev.Visible = False
          CmdNext.Visible = False
      End If
      Call Command10_Click
      bolActive = True
   End If
   
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If Not m_bolQuery Then
      KeyDefine KeyCode
   End If
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
   
   Screen.MousePointer = vbDefault
   tool4_enabled
   SSTab1.Tab = 0

End Sub

Private Sub SetFormEnable(bolEnabled As Boolean)
   Dim oControl As Control
   For Each oControl In Me.Controls
      If TypeName(oControl) = "CommandButton" Then
         oControl.Enabled = bolEnabled
      
      ElseIf TypeName(oControl) = "TextBox" Then
         oControl.Locked = Not bolEnabled
      End If
   Next
   
   '一律開放單據查詢和離開鈕
   Command8.Enabled = True
   Command10.Enabled = True
   txtCP09.Enabled = True
   txtA1N03.Locked = True
   '外部呼叫,只限該收文號
   If m_bolPrev = True Then
      txtCP09.Locked = True
      Command10.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601) 'Added by Morgan 2024/1/3
   Set rsTmp1 = Nothing 'Added by Lydia 2020/04/10
   Set frm071021 = Nothing
   
   If m_bolPrev = True Then
        tool1_enabled
        If TypeName(m_PrevForm) = "Frmacc1190" Then
            m_PrevForm.Enabled = True
        Else
            m_PrevForm.Visible = True
        End If
   End If
End Sub

Private Function FormSave() As Boolean

On Error GoTo ErrHnd
   cnnConnection.BeginTrans
        strSql = "delete acc1n0 where a1n01='" & txtCP09.Tag & "'"
        cnnConnection.Execute strSql, intI
        With adoacc1n0
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              strSql = "insert into acc1n0(a1n01,a1n02,a1n03,a1n04,a1n05)" & _
                 " values('" & txtCP09.Tag & "','3','" & .Fields("a1n03") & "'" & _
                 ",'" & .Fields("a1n04") & "'," & .Fields("a1n05") & ")"
              cnnConnection.Execute strSql, intI
              .MoveNext
           Loop
        End If
        End With
   cnnConnection.CommitTrans
   
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
       MsgBox "工作點數存檔失敗：" & vbCrLf & Err.Description
   End If
End Function

'放大鏡：查詢
Public Sub Command10_Click()
If m_bolQuery = False Then
   If ChkInputPoint = True Then
      If Val(txtPts) > 0 And Val(txtSum) > 0 Then
         FormSave
      End If
   End If
End If
txtCP09.Tag = txtCP09.Text
OpenTable '收據號碼(重新查詢,載入基本資料)

End Sub
'*************************************************
'  開啟資料表
'*************************************************
Private Function OpenTable() As Boolean
Dim amt1 As Double
On Error GoTo Checking
   
   'Modified by Lydia 2020/04/16
   'strTmpSQL = "select cp01,cp02,cp03,cp04,cp13,cp18,nvl(a1u07,0)/1000 a1u07 from caseprogress, acc1u0 " & _
               "where cp09='" & txtCP09 & "' and cp09=a1u03(+)"
   strTmpSQL = "select cp01,cp02,cp03,cp04,cp13,cp18,sum(nvl(a1u07,0)/1000) a1u07 from caseprogress, acc1u0 " & _
               "where cp09='" & txtCP09 & "' and cp09=a1u03(+) and cp60=a1u02(+) " & _
               "group by cp01,cp02,cp03,cp04,cp13,cp18 "
   intR = 1
   Set rsTmp1 = ClsLawReadRstMsg(intR, strTmpSQL)
   If intR = 1 Then
      rsTmp1.MoveFirst
      txtCP01.Text = "" & rsTmp1.Fields("cp01")
      txtCP02.Text = "" & rsTmp1.Fields("cp02")
      txtCP03.Text = "" & rsTmp1.Fields("cp03")
      txtCP04.Text = "" & rsTmp1.Fields("cp04")
      amt1 = Val("" & rsTmp1.Fields("cp18")) - rsTmp1.Fields("a1u07")
      '減收據有財務處銷帳或銷退後的點數
      Do While Not rsTmp1.EOF
         If rsTmp1.AbsolutePosition > 1 Then
            amt1 = amt1 - rsTmp1.Fields("a1u07")
         End If
         rsTmp1.MoveNext
      Loop
      txtPts = Format(amt1, "###0.000")
   Else
      MsgBox "資料庫查無資料!!"
      Exit Function
   End If
   '3:(原)法務工作點數
   strTmpSQL = "select st02,a1n02,a1n03,decode(decode(lc01,null,'000',lc15),'000',cpm03,cpm04) cpm03,a1n04,a1n05,a1n06" & _
      " from acc1n0,staff,caseprogress,casepropertymap,lawcase,hirecase" & _
      " where a1n01='" & txtCP09 & "' and a1n02='3'" & _
      " and cp09(+)=a1n03 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and st01(+)=a1n04 and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+)"
   strTmpSQL = strTmpSQL & " order by a1n04,a1n03"
   intR = 1
   Set rsTmp1 = ClsLawReadRstMsg(intR, strTmpSQL)
   '改暫存TB
   Set adoacc1n0 = PUB_CreateRecordset(rsTmp1, , , , Me.Name)
   Set rsTmp1 = Nothing
   Set Adodc1.Recordset = adoacc1n0
   
   Set DataGrid1.DataSource = Adodc1
   DataGrid1.Refresh
   DataGrid1.col = 0
   DataGrid1.CurrentCellVisible = True
   SumShow
   AdodcClear
   
   txtSum.Tag = txtSum.Text '記錄已分配點數
     
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
   
End Function
'*************************************************
'  顯示 Adodc 之資料
'
'*************************************************
Private Sub AdodcShow()
   With adoacc1n0
   txtA1N04 = .Fields("a1n04").Value
   txtA1N04.Tag = txtA1N04.Text 'Added by Lydia 2020/04/10
   txtA1N05 = Round(.Fields("a1n05").Value, 3)
   txtA1N03 = "" & .Fields("a1n03").Value
   txtST02 = "" & .Fields("st02").Value
   End With
   m_bolAddNew = False
End Sub


Private Sub txtCP09_GotFocus()
   TextInverse txtCP09
End Sub

Private Sub txtCP09_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA1N03_GotFocus()
   TextInverse txtA1N03
End Sub

Private Sub txtA1N03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA1N03_Validate(Cancel As Boolean)
   m_A1N03_CPM03 = ""
   If Trim(txtA1N03) <> "" Then
      strTmpSQL = "select decode(decode(lc01,null,'000',lc15),'000',cpm03,cpm04) cpm03" & _
                           " from caseprogress,casepropertymap,lawcase,hirecase " & _
                           " where cp09='" & txtA1N03 & "' and cpm01(+)=cp01 and cpm02(+)=cp10" & _
                           " and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+)"
      intR = 1
      Set rsTmp1 = ClsLawReadRstMsg(intR, strTmpSQL)
      If intR = 1 Then
         m_A1N03_CPM03 = "" & rsTmp1.Fields(0)
      Else
         MsgBox "收文號輸入錯誤！"
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtA1N04_GotFocus()
   TextInverse txtA1N04
End Sub

Private Sub txtA1N04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA1N04_Validate(Cancel As Boolean)
   'Modified by Lydia 2020/04/10 輸入人員若為離職人員要提醒但仍可輸入
   'txtST02 = ""
   'If txtA1N04 <> "" Then
      'txtST02 = StaffQuery(txtA1N04)
      'If txtST02 = "" Then
      '   MsgBox "承辦人輸入錯誤！"
      '   Cancel = True
      '   txtA1N04_GotFocus
      '   Exit Sub
      'End If
   If txtA1N04 <> "" And txtA1N04.Tag <> txtA1N04.Text Then
      txtST02 = ""
      strExc(1) = GetStaffName(txtA1N04.Text, True, , , strExc(2))
      If strExc(2) = "2" Then
          MsgBox MsgText(9149), vbExclamation, "資料檢核"
         Cancel = True
         txtA1N04_GotFocus
         Exit Sub
      End If
      If strExc(1) = "" Then
         MsgBox MsgText(9150), vbExclamation, "資料檢核"
         Cancel = True
         txtA1N04_GotFocus
         Exit Sub
      End If
      txtST02 = strExc(1)
      txtA1N04.Tag = txtA1N04.Text
   'end 2020/04/10
      strExc(1) = PUB_GetStaffST15(txtA1N04, 1)
      If strExc(1) = "F51" Or strExc(1) = "F52" Then
         MsgBox "不可輸入外翻編號！"
         Cancel = True
         txtA1N04_GotFocus
         Exit Sub
      End If
      
      strExc(2) = GetDeptA09(strExc(1), "10")
      If Len(strExc(2)) = 0 Then
         MsgBox "承辦人作帳部門未設定或無法讀取！"
         Cancel = True
         txtA1N04_GotFocus
         Exit Sub
      End If
      
      '收文號只有一個時預設
      If Cancel = False And Trim(txtA1N03) = "" Then
         txtA1N03 = txtCP09
      End If
   End If
End Sub

Private Sub txtA1N05_GotFocus()
   TextInverse txtA1N05
End Sub

Private Sub SumShow()
Dim ii As Integer

   Set rsTmp1 = Nothing 'Added by Lydia 2020/04/10
   txtSum = 0
   Set rsTmp1 = Adodc1.Recordset.Clone
   With rsTmp1
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         txtSum = Val(txtSum) + Val("" & .Fields("a1n05"))
         .MoveNext
      Loop
   End If
   End With
End Sub

'*************************************************
'  刪除 Adodc 之資料
'
'*************************************************
Private Sub AdodcDelete(adoRst As ADODB.Recordset)
   With adoRst
   If Not (.EOF Or .BOF) Then
      .Delete
      .UPDATE
   End If
   End With
End Sub

'*************************************************
'  清除查詢顯示
'
'*************************************************
Public Sub AdodcClear()
   txtA1N04 = ""
   txtA1N04.Tag = "" 'Added by Lydia 2020/04/10
   txtST02 = ""
   txtA1N05 = ""
   txtA1N03 = " " '預設空白
   m_bolAddNew = True
   If txtA1N04.Enabled And txtA1N04.Visible Then txtA1N04.SetFocus
End Sub

Private Sub AdodcAdd()
   Dim bolAdd As Boolean
   bolAdd = True
   
   With adoacc1n0
   
   If .RecordCount > 0 Then
      .Sort = "a1n04,a1n03"
      .MoveFirst
      .Find "a1n04='" & txtA1N04 & "'"
      If Not .EOF Then
         .Find "a1n03='" & txtA1N03 & "'"
         If Not .EOF Then
            If txtA1N04 = .Fields("a1n04") Then
               bolAdd = False
               If MsgBox("資料已存在，是否要更新！", vbYesNo + vbDefaultButton2) = vbNo Then
                  GoTo EXITSUB
               End If
            End If
         End If
      End If
   End If
   If bolAdd Then .AddNew
   .Fields("a1n04").Value = txtA1N04
   .Fields("a1n05").Value = Val(txtA1N05)
   .Fields("a1n03").Value = txtA1N03
   .Fields("st02").Value = txtST02
   .Fields("cpm03").Value = m_A1N03_CPM03
   .UPDATE
   .Sort = "a1n04,a1n03"
   AdodcClear
   SumShow
   
EXITSUB:
   End With
   
End Sub

Private Sub AdodcUpdate()
   Dim iPos As Integer
   
   With adoacc1n0
   iPos = .AbsolutePosition
   .MoveFirst
   .Find "a1n04='" & txtA1N04 & "'"
   If Not .EOF Then
      .Find "a1n03='" & txtA1N03 & "'"
      If Not .EOF Then
         If txtA1N04 = .Fields("a1n04") And iPos <> .AbsolutePosition Then
            MsgBox "承辦人+收文號資料重複，請重新輸入！"
            If iPos = 1 Then
               .MoveFirst
            Else
               .Move iPos - 1, 1
            End If
            GoTo EXITSUB
         End If
      End If
   End If
   If iPos = 1 Then
      .MoveFirst
   Else
      .Move iPos - 1, 1
   End If
   .Fields("a1n04").Value = txtA1N04
   .Fields("a1n05").Value = Val(txtA1N05)
   .Fields("a1n03").Value = txtA1N03
   .Fields("st02").Value = txtST02
   .Fields("cpm03").Value = m_A1N03_CPM03
   .UPDATE
   AdodcClear
   SumShow
   
EXITSUB:
   End With
   
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyInsert '新增記錄按Insert鍵
         If SSTab1.Tab = 0 Then
            If TxtValidate Then
               If m_bolAddNew Then
                  AdodcAdd
               Else
                  AdodcUpdate
               End If
            End If
         End If
   End Select
   'KeyEnter KeyCode 'Remove by Lydia 2020/04/20 不使用共用模組
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   If txtA1N04 = "" Then
      MsgBox "承辦人不可空白！"
      txtA1N04.SetFocus
      Exit Function
   Else
      txtA1N04_Validate bCancel
      If bCancel Then
         txtA1N04_GotFocus
         txtA1N04.SetFocus
         Exit Function
      End If
   End If
   
   If txtA1N04 < "F" And Trim(txtA1N03) = "" Then
      MsgBox "承辦人非部門時收文號不可空白！"
      txtA1N03_GotFocus
      txtA1N03.SetFocus
      Exit Function
      
   ElseIf txtA1N04 > "F" And Trim(txtA1N03) <> "" Then
      MsgBox "承辦人為部門時不可輸入收文號！"
      txtA1N03_GotFocus
      txtA1N03.SetFocus
      Exit Function
      
   End If
   
   If (txtCP01 = "FCL" Or txtCP01 = "LIN" Or txtCP01 = "CFL") And txtA1N04 = "97009" Then
      MsgBox "FCL,LIN,CFL案件承辦人不可用 97009 編號！"
      txtA1N04_GotFocus
      txtA1N04.SetFocus
      Exit Function
   End If
   
   txtA1N03_Validate bCancel
   If bCancel Then
      txtA1N03_GotFocus
      txtA1N03.SetFocus
      Exit Function
   End If

   If Val(txtA1N05) = 0 Then
      MsgBox "點數必須大於 0 ！", vbExclamation
      txtA1N05.SetFocus
      Exit Function
   End If
   
   TxtValidate = True
End Function

'依照預設規則分配點數
'Modified by Lydia 2020/04/10 改為共用模組PUB_GetLawPointAuto
'Private Function Get_PointAutoassign(strCP09 As String) As Boolean
'   Dim stSQL As String, intR As Integer
'   Dim adoRst As ADODB.Recordset
'   Dim douPtSum As Double '總點數
'
'On Error GoTo ErrHnd
'
'       stSQL = "select cp09,cp14,cp18,nvl(a1u07,0) damt2 from caseprogress,acc1u0 " & _
'              "where cp09='" & strCP09 & "' and cp18>0 and cp09=a1u03(+) and length(cp14) > 0 "
'      intR = 1
'      Set adoRst = ClsLawReadRstMsg(intR, stSQL)
'      If intR = 1 Then 'strCP09
'        cnnConnection.BeginTrans
'          cnnConnection.Execute "delete acc1n0 where a1n01='" & strCP09 & "'"
'         '算到小數點
'         douPtSum = adoRst.Fields("cp18") - adoRst.Fields("damt2") / 1000 '減收據有財務處銷帳或銷退後的點數
'         If douPtSum > 0 Then
'            stSQL = "insert into ACC1N0(a1n01,a1n02,a1n03,a1n04,a1n05)" & _
'                  " values ('" & strCP09 & "','3','" & adoRst.Fields("cp09") & "','" & "" & adoRst.Fields("cp14") & "'," & douPtSum & ")"
'            cnnConnection.Execute stSQL, intR
'
'         End If
'        cnnConnection.CommitTrans
'
'        OpenTable '先帶入畫面
'
'        MsgBox "重新分配完畢，若有特殊分配請再人工調整！"
'      Else
'        MsgBox "收文號無承辦人!", vbExclamation
'      End If 'strCP09
'
'   GoTo EXITSUB
'
'ErrHnd:
'      cnnConnection.RollbackTrans
'      MsgBox Err.Description
'
'EXITSUB:
'   Set adoRst = Nothing
'
'End Function

Private Sub txtcp01_Change()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse

   m_bolQuery = True
   If IsEmptyText(txtCP01) = False Then
      ' 檢查系統類別
      If IsCorrectSysKind(txtCP01) = False Then
         strTit = "資料檢核"
         strMsg = "本所案號中的系統別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtCP09_GotFocus
         GoTo FrmChange
      End If
      ' 檢查使用者權限
      If IsUserHasRightOfSystem(strUserNum, txtCP01) = False Then
         strTit = "資料檢核"
         strMsg = "您沒有使用該系統類別的權限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtCP09_GotFocus
         GoTo FrmChange
      End If
   End If
   
   m_bolQuery = False
   
FrmChange:
   If m_bolQuery Then
      SetFormEnable False
   Else
      SetFormEnable True
   End If
   
End Sub

