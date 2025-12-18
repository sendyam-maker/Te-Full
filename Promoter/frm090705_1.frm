VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090705_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "未齊備、未完稿、未發文查詢"
   ClientHeight    =   5715
   ClientLeft      =   -1710
   ClientTop       =   1515
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   8040
      TabIndex        =   7
      Top             =   48
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5136
      Left            =   36
      TabIndex        =   0
      Top             =   528
      Width           =   9264
      _ExtentX        =   16351
      _ExtentY        =   9049
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      MouseIcon       =   "frm090705_1.frx":0000
      TabCaption(0)   =   "已收文草圖未齊備"
      TabPicture(0)   =   "frm090705_1.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grd1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "已草圖完成墨圖未齊備"
      TabPicture(1)   =   "frm090705_1.frx":0038
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grd1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "墨圖已完成未發文"
      TabPicture(2)   =   "frm090705_1.frx":0054
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl1(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "grd1(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   4488
         Index           =   0
         Left            =   72
         TabIndex        =   2
         Top             =   612
         Width           =   9132
         _ExtentX        =   16113
         _ExtentY        =   7911
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   4488
         Index           =   1
         Left            =   -74928
         TabIndex        =   4
         Top             =   612
         Width           =   9132
         _ExtentX        =   16113
         _ExtentY        =   7911
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   4488
         Index           =   2
         Left            =   -74928
         TabIndex        =   6
         Top             =   612
         Width           =   9132
         _ExtentX        =   16113
         _ExtentY        =   7911
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
      Begin VB.Label lbl1 
         Height          =   180
         Index           =   2
         Left            =   -74928
         TabIndex        =   5
         Top             =   360
         Width           =   9108
      End
      Begin VB.Label lbl1 
         Height          =   180
         Index           =   1
         Left            =   -74928
         TabIndex        =   3
         Top             =   360
         Width           =   9108
      End
      Begin VB.Label lbl1 
         Height          =   180
         Index           =   0
         Left            =   72
         TabIndex        =   1
         Top             =   360
         Width           =   9108
      End
   End
End
Attribute VB_Name = "frm090705_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/07 改成Form2.0 ; grd1(index)改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Private Sub cmdOK_Click()
Me.Hide
frm090705.Show
Unload Me
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
StrMenu
SetGrd1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090705_1 = Nothing
End Sub

Sub StrMenu()
strSql = "select r107001,r107002,r107003,r107004,r107005,r107006,r107007,r107008,r107013 from r090705 where id='" & strUserNum & "' and r107014='1' order by 1,2,3 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd1(0).Recordset = adoRecordset
        lbl1(0).Caption = "已收文未草圖齊備之期限 " & frm090705.Txt1(9) & " 天,共 " & Trim(str(.RecordCount)) & " 件 "
    Else
        lbl1(0).Caption = "已收文未草圖齊備之期限 " & frm090705.Txt1(9) & " 天,共   件 "
    End If
    CheckOC
    strSql = "select r107001,r107002,r107003,r107004,r107005,r107006,r107009,r107010,r107013 from r090705 where id='" & strUserNum & "' and r107014='2' order by 1,2,3 "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd1(1).Recordset = adoRecordset
        lbl1(1).Caption = "已草圖完成未墨圖齊備之期限 " & frm090705.Txt1(10) & " 天,共 " & Trim(str(.RecordCount)) & " 件 "
    Else
        lbl1(1).Caption = "已草圖完成未墨圖齊備之期限 " & frm090705.Txt1(10) & " 天,共   件 "
    End If
    CheckOC
    strSql = "select r107001,r107002,r107003,r107004,r107005,r107006,r107011,r107012,r107013 from r090705 where id='" & strUserNum & "' and r107014='3' order by 1,2,3 "
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd1(2).Recordset = adoRecordset
        lbl1(2).Caption = "已墨圖完成未發文之期限 " & frm090705.Txt1(11) & " 天,共 " & Trim(str(.RecordCount)) & " 件 "
    Else
        lbl1(2).Caption = "已墨圖完成未發文之期限 " & frm090705.Txt1(11) & " 天,共   件 "
    End If
    CheckOC
End With
End Sub

Private Sub SetGrd1()
'設定grid
With grd1(0)
    .Cols = 9
    .row = 0
    .col = 0:   .Text = "所別"
    .ColWidth(0) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 1
    If frm090705.Txt1(5) = "1" Then
        .Text = "繪圖人員"
    Else
        .Text = "承辦人"
    End If
    .ColWidth(1) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 2: .Text = "本所案號"
    .ColWidth(2) = 1550
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = "案件名稱"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = "種類"
    .ColWidth(4) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 5: .Text = "案件性質"
    .ColWidth(5) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 6: .Text = "收文日"
    .ColWidth(6) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = "收文天數"
    .ColWidth(7) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 8
    If frm090705.Txt1(5) = "1" Then
        .Text = "承辦人"
    Else
        .Text = "繪圖人員"
    End If
    .ColWidth(8) = 1500
    .CellAlignment = flexAlignCenterCenter
End With
With grd1(1)
    .Cols = 9
    .row = 0
    .col = 0:   .Text = "所別"
    .ColWidth(0) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 1
    If frm090705.Txt1(5) = "1" Then
        .Text = "繪圖人員"
    Else
        .Text = "承辦人"
    End If
    .ColWidth(1) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 2: .Text = "本所案號"
    .ColWidth(2) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = "案件名稱"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = "種類"
    .ColWidth(4) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 5: .Text = "案件性質"
    .ColWidth(5) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 6: .Text = "草圖完稿日"
    .ColWidth(6) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = "草圖完稿天數"
    .ColWidth(7) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 8
    If frm090705.Txt1(5) = "1" Then
        .Text = "承辦人"
    Else
        .Text = "繪圖人員"
    End If
    .ColWidth(8) = 1500
    .CellAlignment = flexAlignCenterCenter
End With
With grd1(2)
    .Cols = 9
    .row = 0
    .col = 0:   .Text = "所別"
    .ColWidth(0) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 1
    If frm090705.Txt1(5) = "1" Then
        .Text = "繪圖人員"
    Else
        .Text = "承辦人"
    End If
    .ColWidth(1) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 2: .Text = "本所案號"
    .ColWidth(2) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = "案件名稱"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = "種類"
    .ColWidth(4) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 5: .Text = "案件性質"
    .ColWidth(5) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 6: .Text = "墨圖完稿日"
    .ColWidth(6) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = "墨圖完稿天數"
    .ColWidth(7) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 8
    If frm090705.Txt1(5) = "1" Then
        .Text = "承辦人"
    Else
        .Text = "繪圖人員"
    End If
    .ColWidth(8) = 1500
    .CellAlignment = flexAlignCenterCenter
End With
End Sub

