VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090612_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "未齊備,未完稿,未發文查詢"
   ClientHeight    =   5712
   ClientLeft      =   -2256
   ClientTop       =   1596
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5712
   ScaleWidth      =   9312
   Begin VB.CommandButton cmdOK 
      Caption         =   "未完稿/未會稿案件(&Word)"
      CausesValidation=   0   'False
      Height          =   330
      Index           =   6
      Left            =   48
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   70
      Visible         =   0   'False
      Width           =   2184
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "繪圖進度(&E)"
      Height          =   330
      Index           =   5
      Left            =   6072
      TabIndex        =   12
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "承辦進度(&D)"
      Height          =   330
      Index           =   4
      Left            =   4848
      TabIndex        =   11
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度(&C)"
      Height          =   330
      Index           =   3
      Left            =   3624
      TabIndex        =   10
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料(&B)"
      Height          =   330
      Index           =   2
      Left            =   2400
      TabIndex        =   9
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      Height          =   330
      Index           =   1
      Left            =   8520
      TabIndex        =   6
      Top             =   70
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   7296
      TabIndex        =   5
      Top             =   70
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4890
      Left            =   0
      TabIndex        =   2
      Top             =   795
      Width           =   12210
      _ExtentX        =   21527
      _ExtentY        =   8615
      _Version        =   393216
      TabHeight       =   520
      TabMaxWidth     =   5292
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "已收文未齊備 / 已齊備未完稿"
      TabPicture(0)   =   "frm090612_1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "已完稿未會稿"
      TabPicture(1)   =   "frm090612_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(4)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "已會稿未會完 / 已會完未發文"
      TabPicture(2)   =   "frm090612_1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(3)"
      Tab(2).Control(1)=   "Frame1(2)"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   2250
         Index           =   4
         Left            =   -74940
         TabIndex        =   17
         Top             =   330
         Width           =   9192
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
            Height          =   2040
            Index           =   4
            Left            =   48
            TabIndex        =   18
            Top             =   180
            Width           =   9108
            _ExtentX        =   16066
            _ExtentY        =   3598
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            HighLight       =   2
            AllowUserResizing=   3
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   2250
         Index           =   2
         Left            =   -74970
         TabIndex        =   15
         Top             =   330
         Width           =   9192
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
            Height          =   2040
            Index           =   2
            Left            =   48
            TabIndex        =   16
            Top             =   192
            Width           =   9108
            _ExtentX        =   16066
            _ExtentY        =   3598
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            HighLight       =   2
            AllowUserResizing=   3
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   2250
         Index           =   3
         Left            =   -74970
         TabIndex        =   13
         Top             =   2595
         Width           =   9192
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
            Height          =   2040
            Index           =   3
            Left            =   48
            TabIndex        =   14
            Top             =   180
            Width           =   9108
            _ExtentX        =   16066
            _ExtentY        =   3598
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            HighLight       =   2
            AllowUserResizing=   3
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   2250
         Index           =   1
         Left            =   36
         TabIndex        =   4
         Top             =   2592
         Width           =   9192
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
            Height          =   2040
            Index           =   1
            Left            =   36
            TabIndex        =   8
            Top             =   180
            Width           =   9108
            _ExtentX        =   16066
            _ExtentY        =   3598
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            HighLight       =   2
            AllowUserResizing=   3
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   2250
         Index           =   0
         Left            =   36
         TabIndex        =   3
         Top             =   324
         Width           =   9192
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
            Height          =   2040
            Index           =   0
            Left            =   36
            TabIndex        =   7
            Top             =   180
            Width           =   9108
            _ExtentX        =   16066
            _ExtentY        =   3598
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            HighLight       =   2
            AllowUserResizing=   3
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
      End
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   420
      Width           =   1815
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "3201;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   144
      TabIndex        =   0
      Top             =   504
      Width           =   1716
   End
End
Attribute VB_Name = "frm090612_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/23 改成Form2.0 ; grd1(index)改字型=新細明體-ExtB、Combo1
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim i As Integer, j As Integer
Public cmdState As Integer 'Added by Lydia 2017/05/05 紀錄作用按鍵
Dim StrTag As String 'Added by Lydia 2017/05/05
Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Me.Hide
     frm090612.Show
     Unload Me
Case 1
     Me.Hide
     Unload frm090612
     Unload Me
'Added by Lydia 2017/05/05
Case 2, 3, 4, 5 '基本資料,案件進度,承辦進度,繪圖進度
     cmdState = Index
     PubShowNextData
'end 2017/05/05

'Added by Morgan 2024/3/18
Case 6 '未完稿/未會稿案件(&Word)
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   ExportWord
   Me.Enabled = True
   Screen.MousePointer = vbDefault
Case Else
End Select
End Sub

'Modified by Lydia 2022/02/07 Form2.0點選同一人不會觸發Click事件，改用DropButtonClick事件但要控制第2次才執行
'Private Sub Combo1_Click()
Private Sub Combo1_DropButtonClick()
   Static bClick As Boolean
   If bClick = False Then
      bClick = True
      Exit Sub
   End If
   bClick = False
'end 2022/02/07

Process
SetGrd
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
If Val(frm090612.Txt1(5)) = 1 Then
    lbl1.Caption = "承辦人："
Else
    lbl1.Caption = "智權人員："
End If

SSTab1.Tab = 0 'Added by Lydia 2017/05/11
StrMenu
Process
SetGrd
End Sub

Sub Process()

For i = 0 To 4 'Modified by Morgan 2021/8/4 3->4
    'Set grd1(I).Recordset = Nothing
    grd1(i).Clear
    grd1(i).Rows = 2
    
Next i
'第一個
If frm090612.Check1(0).Value = 1 Then
   If Len(Trim(Combo1.Text)) = 0 Then
       'Modified by Lydia 2017/05/05
       'strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107007,R107008,R107016 FROM R090612 WHERE R107017='1' AND ID='" & strUserNum & "' AND (R107001 IS NULL OR R107001='') ORDER BY 1"
       strSql = "SELECT ' ' V,R107002,R107003,R107004,R107005,R107006,R107007,R107008,R107016,R107019 FROM R090612 WHERE R107017='1' AND ID='" & strUserNum & "' AND (R107001 IS NULL OR R107001='') ORDER BY R107002"
   Else
       'Modified by Lydia 2017/05/05
       'strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107007,R107008,R107016 FROM R090612 WHERE R107017='1' AND ID='" & strUserNum & "' AND R107001='" & Combo1.Text & "' ORDER BY 1"
       strSql = "SELECT ' ' V,R107002,R107003,R107004,R107005,R107006,R107007,R107008,R107016,R107019 FROM R090612 WHERE R107017='1' AND ID='" & strUserNum & "' AND R107001='" & Combo1.Text & "' ORDER BY R107002"
   End If
   CheckOC
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           Set grd1(0).Recordset = adoRecordset
       End If
   End With
   CheckOC
   i = 0
   j = 0
   If Len(Trim(Combo1.Text)) = 0 Then
       strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='1' AND (R107001='' OR R107001 IS NULL) AND R107018='#' "
   Else
       strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='1' AND R107001='" & Combo1.Text & "' AND R107018='#' "
   End If
   CheckOC
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           i = Val(CheckStr(.Fields(0)))
       End If
   End With
   CheckOC
   If Len(Trim(Combo1.Text)) = 0 Then
       strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='1' AND (R107001='' OR R107001 IS NULL) "
   Else
       strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='1' AND R107001='" & Combo1.Text & "' "
   End If
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           j = Val(CheckStr(.Fields(0)))
       End If
   End With
   CheckOC
   Frame1(0).Caption = "超過已收文未齊備之時限申請案 " & frm090612.Txt1(11).Text & " 天, 非申請案 " & frm090612.Txt1(11).Text & " 天, 申請案共 " & str(i) & " 件, 非申請案共 " & str(j - i) & " 件 "
Else
    Frame1(0).Caption = "超過已收文未齊備之時限申請案  天, 非申請案  天, 申請案共  件, 非申請案共  件 "
End If
If frm090612.Check1(1).Value = 1 Then
   '第2個
   If Len(Trim(Combo1.Text)) = 0 Then
       'Modified by Lydia 2017/05/05
       'strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107009,R107010,R107016 FROM R090612 WHERE R107017='2' AND ID='" & strUserNum & "' AND (R107001='' OR R107001 IS NULL) ORDER BY 1"
       strSql = "SELECT ' ' V,R107002,R107003,R107004,R107005,R107006,R107009,R107010,R107016,R107019 FROM R090612 WHERE R107017='2' AND ID='" & strUserNum & "' AND (R107001='' OR R107001 IS NULL) ORDER BY R107002"
   Else
       'Modified by Lydia 2017/05/05
       'strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107009,R107010,R107016 FROM R090612 WHERE R107017='2' AND ID='" & strUserNum & "' AND R107001='" & Combo1.Text & "' ORDER BY 1"
       strSql = "SELECT ' ' V,R107002,R107003,R107004,R107005,R107006,R107009,R107010,R107016,R107019 FROM R090612 WHERE R107017='2' AND ID='" & strUserNum & "' AND R107001='" & Combo1.Text & "' ORDER BY R107002"
   End If
   CheckOC
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           Set grd1(1).Recordset = adoRecordset
       End If
   End With
   CheckOC
   j = 0
   If Len(Trim(Combo1.Text)) = 0 Then
       strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='2' AND (R107001='' OR R107001 IS NULL) "
   Else
       strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='2' AND R107001='" & Combo1.Text & "' "
   End If
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           j = Val(CheckStr(.Fields(0)))
       End If
   End With
   CheckOC
   
   Frame1(1).Caption = "超過已齊備未完稿之時限 " & frm090612.Txt1(12).Text & " 天, 共 " & str(j) & " 件 "
Else
    Frame1(1).Caption = "超過已齊備未完稿之時限  天, 共  件 "
End If
If frm090612.Check1(2).Value = 1 Then
   '第3個
   If Len(Trim(Combo1.Text)) = 0 Then
       'Modified by Lydia 2017/05/05
       'strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107013,R107016 FROM R090612 WHERE R107017='3' AND ID='" & strUserNum & "' AND (R107001='' OR R107001 IS NULL) ORDER BY 1"
       strSql = "SELECT ' ' V,R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107013,R107016,R107019 FROM R090612 WHERE R107017='3' AND ID='" & strUserNum & "' AND (R107001='' OR R107001 IS NULL) ORDER BY R107002"
   Else
       'Modified by Lydia 2017/05/05
       'strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107013,R107016 FROM R090612 WHERE R107017='3' AND ID='" & strUserNum & "' AND R107001='" & Combo1.Text & "' ORDER BY 1"
       strSql = "SELECT ' ' V,R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107013,R107016,R107019 FROM R090612 WHERE R107017='3' AND ID='" & strUserNum & "' AND R107001='" & Combo1.Text & "' ORDER BY R107002"
   End If
   CheckOC
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           Set grd1(2).Recordset = adoRecordset
       End If
   End With
   CheckOC
   j = 0
   If Len(Trim(Combo1.Text)) = 0 Then
       strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='3' AND (R107001='' OR R107001 IS NULL) "
   Else
       strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='3' AND R107001='" & Combo1.Text & "' "
   End If
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           j = Val(CheckStr(.Fields(0)))
       End If
   End With
   CheckOC
   
   Frame1(2).Caption = "超過已會稿未會稿完成之時限 " & frm090612.Txt1(13).Text & " 天, 共 " & str(j) & " 件 "
Else
    Frame1(2).Caption = "超過已會稿未會稿完成之時限  天, 共  件 "
End If
If frm090612.Check1(3).Value = 1 Then
   '第4個
   If Len(Trim(Combo1.Text)) = 0 Then
       'Modified by Lydia 2017/05/05
       'strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107014,R107015,R107016 FROM R090612 WHERE R107017='4' AND ID='" & strUserNum & "' AND (R107001='' OR R107001 IS NULL) ORDER BY 1"
       strSql = "SELECT ' ' V,R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107014,R107015,R107016,R107019 FROM R090612 WHERE R107017='4' AND ID='" & strUserNum & "' AND (R107001='' OR R107001 IS NULL) ORDER BY R107002"
   Else
       'Modified by Lydia 2017/05/05
       'strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107014,R107015,R107016 FROM R090612 WHERE R107017='4' AND ID='" & strUserNum & "' AND R107001='" & Combo1.Text & "' ORDER BY 1"
       strSql = "SELECT ' ' V,R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107014,R107015,R107016,R107019 FROM R090612 WHERE R107017='4' AND ID='" & strUserNum & "' AND R107001='" & Combo1.Text & "' ORDER BY R107002"
   End If
   CheckOC
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           Set grd1(3).Recordset = adoRecordset
       End If
   End With
   CheckOC
   j = 0
   If Len(Trim(Combo1.Text)) = 0 Then
       strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='4' AND (R107001='' OR R107001 IS NULL) "
   Else
       strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='4' AND R107001='" & Combo1.Text & "' "
   End If
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           j = Val(CheckStr(.Fields(0)))
       End If
   End With
   CheckOC
   
   Frame1(3).Caption = "超過已會稿完成未發文之時限 " & frm090612.Txt1(14).Text & " 天, 共 " & str(j) & " 件 "
Else
    Frame1(3).Caption = "超過已會稿完成未發文之時限  天, 共  件 "
End If

'Added by Morgan 2021/8/4 已完稿未會稿
If frm090612.Check1(4).Value = 1 Then
   '第5個
   If Len(Trim(Combo1.Text)) = 0 Then
       strSql = "SELECT ' ' V,R107002,R107003,R107004,R107005,R107006,R107011,R107013,R107016,R107019 FROM R090612 WHERE R107017='5' AND ID='" & strUserNum & "' AND (R107001='' OR R107001 IS NULL) ORDER BY R107002"
   Else
       strSql = "SELECT ' ' V,R107002,R107003,R107004,R107005,R107006,R107011,R107013,R107016,R107019 FROM R090612 WHERE R107017='5' AND ID='" & strUserNum & "' AND R107001='" & Combo1.Text & "' ORDER BY R107002"
   End If
   CheckOC
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           Set grd1(4).Recordset = adoRecordset
       End If
   End With
   CheckOC
   j = 0
   If Len(Trim(Combo1.Text)) = 0 Then
       strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='5' AND (R107001='' OR R107001 IS NULL) "
   Else
       strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='5' AND R107001='" & Combo1.Text & "' "
   End If
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           j = Val(CheckStr(.Fields(0)))
       End If
   End With
   CheckOC
   
   Frame1(4).Caption = "超過已完稿未會稿之時限 " & frm090612.Txt1(17).Text & " 天, 共 " & str(j) & " 件 "
Else
    Frame1(4).Caption = "超過已完稿未會稿之時限  天, 共  件 "
End If
'end 2021/8/4
End Sub

Sub StrMenu()
strSql = "SELECT DISTINCT R107001 FROM R090612 WHERE ID='" & strUserNum & "' "
CheckOC
j = 0
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            Combo1.AddItem CheckStr(.Fields(0)), j
            j = j + 1
            .MoveNext
        Loop
    End If
End With
CheckOC

If Combo1.ListCount > 0 Then  'Added by Lydia 2021/12/28 若無List清單，Form 2.0不接受
    Combo1.Text = Combo1.List(0)
End If

End Sub

Private Sub SetGrd()
For i = 0 To 4 'Modified by Morgan 2021/8/4 3 -> 4
    grd1(i).Visible = False
Next i

With grd1(0)
    'Modified by Lydia 2017/05/05
    '.Cols = 8
    .Cols = 10
    .row = 0
    'Added by Lydia 2017/05/05 +勾選項
    .col = 0: .Text = "V"
    .ColWidth(0) = 270
    .CellAlignment = flexAlignCenterCenter
    'end 2017/05/05
    'Modified by Lydia .col + 1 ; .ColWidth + 1
    .col = 1: .Text = "目次"
    .ColWidth(1) = 500
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
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = "收文天數"
    .ColWidth(7) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 8
    If Val(frm090612.Txt1(5)) = 1 Then
        .Text = "智權人員"
    Else
        .Text = "承辦人"
    End If
    .ColWidth(8) = 800
    'end 2017/05/05
    'Added by Lydia 2017/05/05
    .col = 9: .Text = "收文號"
    .ColWidth(9) = 0
    'end 2017/05/05
    .CellAlignment = flexAlignCenterCenter
End With

With grd1(1)
    'Modified by Lydia 2017/05/05
    '.Cols = 8
    .Cols = 10
    .row = 0
    'Added by Lydia 2017/05/05 +勾選項
    .col = 0: .Text = "V"
    .ColWidth(0) = 270
    .CellAlignment = flexAlignCenterCenter
    'end 2017/05/05
    'Modified by Lydia .col + 1 ; .ColWidth + 1
    .col = 1: .Text = "目次"
    .ColWidth(1) = 500
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
    .col = 6: .Text = "齊備日"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = "齊備天數"
    .ColWidth(7) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 8
    If Val(frm090612.Txt1(5)) = 1 Then
        .Text = "智權人員"
    Else
        .Text = "承辦人"
    End If
    .ColWidth(8) = 800
    'end 2017/05/05
    'Added by Lydia 2017/05/05
    .col = 9: .Text = "收文號"
    .ColWidth(9) = 0
    'end 2017/05/05
    .CellAlignment = flexAlignCenterCenter
End With

With grd1(2)
    'Modified by Lydia 2017/05/05
    '.Cols = 9
    .Cols = 11
    .row = 0
    'Added by Lydia 2017/05/05 +勾選項
    .col = 0: .Text = "V"
    .ColWidth(0) = 270
    .CellAlignment = flexAlignCenterCenter
    'end 2017/05/05
    'Modified by Lydia .col + 1 ; .ColWidth + 1
    .col = 1: .Text = "目次"
    .ColWidth(1) = 500
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
    .col = 6: .Text = "完稿日"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = "會稿日"
    .ColWidth(7) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 8: .Text = "會稿天數"
    .ColWidth(8) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 9
    If Val(frm090612.Txt1(5)) = 1 Then
        .Text = "智權人員"
    Else
        .Text = "承辦人"
    End If
    .ColWidth(9) = 800
    'end 2017/05/05
    'Added by Lydia 2017/05/05
    .col = 10: .Text = "收文號"
    .ColWidth(10) = 0
    'end 2017/05/05
    .CellAlignment = flexAlignCenterCenter
End With

With grd1(3)
    'Modified by Lydia 2017/05/05
    '.Cols = 10
    .Cols = 12
    .row = 0
    'Added by Lydia 2017/05/05 +勾選項
    .col = 0: .Text = "V"
    .ColWidth(0) = 270
    .CellAlignment = flexAlignCenterCenter
    'end 2017/05/05
    'Modified by Lydia .col + 1 ; .ColWidth + 1
    .col = 1: .Text = "目次"
    .ColWidth(1) = 500
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
    .col = 6: .Text = "完稿日"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = "會稿日"
    .ColWidth(7) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 8: .Text = "會稿完成日"
    .ColWidth(8) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 9: .Text = "會完天數"
    .ColWidth(9) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 10
    If Val(frm090612.Txt1(5)) = 1 Then
        .Text = "智權人員"
    Else
        .Text = "承辦人"
    End If
    .ColWidth(10) = 800
    'end 2017/05/05
    'Added by Lydia 2017/05/05
    .col = 11: .Text = "收文號"
    .ColWidth(11) = 0
    'end 2017/05/05
    .CellAlignment = flexAlignCenterCenter
End With

'Added by Morgan 2021/8/4
With grd1(4)
    .Cols = 10
    .row = 0
    .col = 0: .Text = "V"
    .ColWidth(0) = 270
    .CellAlignment = flexAlignCenterCenter
    .col = 1: .Text = "目次"
    .ColWidth(1) = 500
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
    .col = 6: .Text = "完稿日"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = "完稿天數"
    .ColWidth(7) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 8
    If Val(frm090612.Txt1(5)) = 1 Then
        .Text = "智權人員"
    Else
        .Text = "承辦人"
    End If
    .ColWidth(8) = 800
    .col = 9: .Text = "收文號"
    .ColWidth(9) = 0
    .CellAlignment = flexAlignCenterCenter
End With
'end 2021/8/4

For i = 0 To 4 'Modified by Morgan 2021/8/4 3->4
    grd1(i).Visible = True
Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm090612_1 = Nothing
End Sub

'Added by Lydia 2017/05/05
Private Sub Grd1_Click(Index As Integer)
Dim iRow As Integer

   iRow = grd1(Index).MouseRow
   
   If iRow <> 0 And grd1(Index).Rows Then
      grdSelected Index, iRow
   End If
End Sub

'點選列反白或取消反白
Private Sub grdSelected(d_Inx As Integer, p_iRow As Integer)
   Dim stCheck As String, lColor As Long, ii As Integer
   With grd1(d_Inx)
      .row = p_iRow
      .col = 0
      If Trim(.TextMatrix(p_iRow, 1)) <> "" Then
        If Trim(.Text) = "" Then
           .Text = "V"
           lColor = &HFFC0C0
        Else
           .Text = ""
           lColor = &H80000005
        End If
        For ii = 0 To .Cols - 1
           .col = ii
           .CellBackColor = lColor
        Next
      End If
   End With
End Sub

Public Sub PubShowNextData()
Dim dInx As Integer, intRow As Integer
Dim Str01 As String

Select Case cmdState
Case 2 '案件基本資料
      Me.Enabled = False
      StrTag = ""
      For dInx = 0 To 4 'Modified by Morgan 2021/8/4 3->4
         For i = 1 To grd1(dInx).Rows - 1
            grd1(dInx).col = 0
            grd1(dInx).row = i
            If Trim(grd1(dInx).Text) = "V" Then
               grd1(dInx).col = 0
               grd1(dInx).Text = ""
               For j = 0 To grd1(dInx).Cols - 1
                    grd1(dInx).col = j
                    grd1(dInx).CellBackColor = &H80000005
               Next j

              grd1(dInx).col = 2
              Str01 = SystemNumber(grd1(dInx), 1)
              If Not IsNull(grd1(dInx).Text) Then
                  If fnSaveParentForm(Me) = False Then
                      Me.Enabled = True
                      Exit Sub
                  End If
                   Select Case Pub_RplStr(Str01)
                       Case "CFP", "FCP", "P"   '專利
                             Screen.MousePointer = vbHourglass
                             frm100101_3.Show
                             frm100101_3.Tag = Pub_RplStr(grd1(dInx).Text)
                             frm100101_3.StrMenu
                             Screen.MousePointer = vbDefault
                       Case "CFT", "FCT", "T", "TF"   '商標
                             Screen.MousePointer = vbHourglass
                             frm100101_4.Show
                             frm100101_4.Tag = Pub_RplStr(grd1(dInx).Text)
                             frm100101_4.StrMenu
                             Screen.MousePointer = vbDefault
                       Case "CFL", "FCL", "L", "LIN"          '法務
                             Screen.MousePointer = vbHourglass
                             frm100101_5.Show
                             frm100101_5.Tag = Pub_RplStr(grd1(dInx).Text)
                             frm100101_5.StrMenu
                             Screen.MousePointer = vbDefault
                       Case "LA"            '顧問
                             Screen.MousePointer = vbHourglass
                             frm100101_6.Show
                             frm100101_6.Tag = Pub_RplStr(grd1(dInx).Text)
                             frm100101_6.StrMenu
                             Screen.MousePointer = vbDefault
                       Case Else                  '服務
                            Select Case Pub_RplStr(Str01)
                                Case "TB"    '條碼
                                    Screen.MousePointer = vbHourglass
                                    frm100101_7.Show
                                    frm100101_7.Tag = Pub_RplStr(grd1(dInx).Text)
                                    frm100101_7.StrMenu
                                    Screen.MousePointer = vbDefault
                                Case "TM"
                                    Screen.MousePointer = vbHourglass
                                    frm100101_8.Show
                                    frm100101_8.Tag = Pub_RplStr(grd1(dInx).Text)
                                    frm100101_8.StrMenu
                                    Screen.MousePointer = vbDefault
                                Case "TD"
                                    Screen.MousePointer = vbHourglass
                                    frm100101_9.Show
                                    frm100101_9.Tag = Pub_RplStr(grd1(dInx).Text)
                                    frm100101_9.StrMenu
                                    Screen.MousePointer = vbDefault
                                Case "TC", "CFC"
                                    Screen.MousePointer = vbHourglass
                                    frm100101_A.Show
                                    frm100101_A.Tag = Pub_RplStr(grd1(dInx).Text)
                                    frm100101_A.StrMenu
                                    Screen.MousePointer = vbDefault
                                Case Else
                                    Screen.MousePointer = vbHourglass
                                    frm100101_B.Show
                                    frm100101_B.Tag = Pub_RplStr(grd1(dInx).Text)
                                    frm100101_B.StrMenu
                                    Screen.MousePointer = vbDefault
                             End Select
                   End Select
                    Me.Enabled = True
                    Exit Sub
              End If
           End If
        Next i
     Next dInx
     Me.Enabled = True
Case 3 '案件進度
     Me.Enabled = False
     StrTag = ""
     For dInx = 0 To 4 'Modified by Morgan 2021/8/4 3->4
        intRow = PUB_MGridGetId("收文號", grd1(dInx)) '收文號-位置
        For i = 1 To grd1(dInx).Rows - 1
           grd1(dInx).col = 0
           grd1(dInx).row = i
           If Trim(grd1(dInx).Text) = "V" Then
               grd1(dInx).col = 0
               grd1(dInx).Text = ""
               For j = 0 To grd1(dInx).Cols - 1
                  grd1(dInx).col = j
                  grd1(dInx).CellBackColor = &H80000005
               Next j
               grd1(dInx).col = intRow
               If Not IsNull(grd1(dInx).Text) Then
                  Screen.MousePointer = vbHourglass
                   If fnSaveParentForm(Me) = False Then
                       Me.Enabled = True
                       Exit Sub
                   End If
                  frm100101_C.Show
                  frm100101_C.Tag = Trim(grd1(dInx).TextMatrix(i, 2)) + "=" + Pub_RplStr(grd1(dInx).Text)
                  frm100101_C.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
           End If
        Next i
     Next dInx
     Me.Enabled = True
Case 4 '承辦進度
     Me.Enabled = False
     StrTag = ""
     For dInx = 0 To 4 'Modified by Morgan 2021/8/4 3->4
        intRow = PUB_MGridGetId("收文號", grd1(dInx)) '收文號-位置
        For i = 1 To grd1(dInx).Rows - 1
           grd1(dInx).col = 0
           grd1(dInx).row = i
           If Trim(grd1(dInx).Text) = "V" Then
               grd1(dInx).col = 0
               grd1(dInx).Text = ""
               For j = 0 To grd1(dInx).Cols - 1
                  grd1(dInx).col = j
                  grd1(dInx).CellBackColor = &H80000005
               Next j
               Str01 = SystemNumber(Trim(grd1(dInx).TextMatrix(i, 2)), 1) '系統
               grd1(dInx).col = intRow
               If Not IsNull(grd1(dInx).Text) Then
                  Screen.MousePointer = vbHourglass
                   If fnSaveParentForm(Me) = False Then
                       Me.Enabled = True
                       Exit Sub
                   End If
                    If Str01 = "P" Or Str01 = "PS" Or Str01 = "FG" Or _
                       Str01 = "FCP" Or Str01 = "CFP" Or Str01 = "CPS" Or _
                       Val(strSrvDate(1)) < Val(TMdebateStarDT) Then  '專利處工作進度
                       frm100101_F.Show
                       frm100101_F.Process Pub_RplStr(grd1(dInx).Text)
                    Else
                       frm100101_K.Show
                       frm100101_K.Process Pub_RplStr(grd1(dInx).Text)
                    End If
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
           End If
        Next i
     Next dInx
     Me.Enabled = True
Case 5 '繪圖進度
     Me.Enabled = False
     StrTag = ""
     For dInx = 0 To 4 'Modified by Morgan 2021/8/4 3->4
        intRow = PUB_MGridGetId("收文號", grd1(dInx)) '收文號-位置
        For i = 1 To grd1(dInx).Rows - 1
           grd1(dInx).col = 0
           grd1(dInx).row = i
           If Trim(grd1(dInx).Text) = "V" Then
               grd1(dInx).col = 0
               grd1(dInx).Text = ""
               For j = 0 To grd1(dInx).Cols - 1
                  grd1(dInx).col = j
                  grd1(dInx).CellBackColor = &H80000005
               Next j
               grd1(dInx).col = intRow
               If Not IsNull(grd1(dInx).Text) Then
                  Screen.MousePointer = vbHourglass
                   If fnSaveParentForm(Me) = False Then
                       Me.Enabled = True
                       Exit Sub
                   End If
                  frm100101_g.Show
                  frm100101_g.Process Pub_RplStr(grd1(dInx).Text)
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
           End If
        Next i
     Next dInx
     Me.Enabled = True
End Select
End Sub
'end 2017/05/05

'Added by Morgan 2024/3/18
Private Sub ExportWord()
   Dim iResumeCnt As Integer
   Dim stTmp As String
   
On Error GoTo ErrHnd
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   
   If g_WordAp.Visible And g_WordAp.Documents.Count > 0 Then
      If MsgBox("輸出資料是否附加在目前的文件後面？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
         g_WordAp.Selection.EndKey Unit:=wdStory
         g_WordAp.Selection.TypeParagraph
      Else
         g_WordAp.Documents.add
      End If
   Else
      g_WordAp.Documents.add
   End If
   
   With g_WordAp.Application
      .WindowState = wdWindowStateMaximize
      .Visible = True
      
      '邊框設單線
      With .Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
      End With
      '橫印
      .Selection.PageSetup.Orientation = wdOrientLandscape
      '邊界
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(1)
      
      .Selection.Font.Name = "標楷體"
      .Selection.Font.Size = 12
      
      '未完稿
      stTmp = "已齊備未完稿之新案(含201新案翻譯)： P案超過7天(含7天)" & String(2, vbTab) & "CFP案超過14天(含14天)"
      .Selection.TypeText Text:=stTmp
      .Selection.TypeParagraph
      AddNewTable
      
      '未會稿
      strSql = "select R107001 C1,replace(R107003,'-0-00','') C2,R107010 C3 from R090612,caseprogress where R107017='2' AND ID='" & strUserNum & "' and cp09(+)=R107019 and cp10 in ('101','102','201') and ( (cp01='P' and R107010>=7) or (cp01='CFP' and R107010>=14) ) order by 3 desc"
      stTmp = "已完稿未會稿之新案： P案超過4天(含4天)" & String(2, vbTab) & "CFP案超過6天(含6天)"
      .Selection.EndKey Unit:=wdStory
      .Selection.TypeText Text:=stTmp
      .Selection.TypeParagraph
      AddNewTable 1
      .Activate
   End With
   
ErrHnd:
   If Err.Number <> 0 Then
      If iResumeCnt > 3 Then
         MsgBox "錯誤 : " & Err.Description, vbCritical
      Else
         iResumeCnt = iResumeCnt + 1
         Select Case Err.Number
            Case 91:
               g_WordAp.Documents.add
               Resume Next
            Case 462:
               Set g_WordAp = New Word.Application
               Resume
            Case Else:
               MsgBox "錯誤" & Err.Number & " : " & Err.Description, vbCritical
         End Select
      End If
   End If
End Sub

'Added by Morgan 2024/3/18
Private Sub AddNewTable(Optional iType As Integer = 0)
   Dim oTable As Word.Table
   Dim iCols As Integer, iCol As Integer, iRow As Integer, ii As Integer
   
   If iType = 0 Then
      iCols = 4
   Else
      iCols = 5
   End If
   
   With g_WordAp.Application
      '新增表格
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=1, NumColumns:=iCols)
      
      'oTable.AllowAutoFit = True
      .Selection.SelectRow
      With .Selection.Borders(wdBorderTop)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderLeft)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderBottom)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderRight)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderHorizontal)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
      End With
      With .Selection.Borders(wdBorderVertical)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      
      ii = 1
      oTable.Columns(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.7), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(1).Select
      .Selection.TypeText "工程師"
      
      oTable.Columns(2).SetWidth ColumnWidth:=.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(2).Select
      .Selection.TypeText "案號"
      
      '未完稿
      If iType = 0 Then
         oTable.Columns(3).SetWidth ColumnWidth:=.CentimetersToPoints(1.8), RulerStyle:=wdAdjustProportional
         oTable.Rows(ii).Cells(3).Select
         .Selection.TypeText "天數"
         
         oTable.Rows(ii).Cells(4).Select
         .Selection.TypeText "承諾會稿日"
            
         strSql = "select R107001 C1,replace(R107003,'-0-00','') C2,R107010 C3 from R090612,caseprogress where R107017='2' AND ID='" & strUserNum & "' and cp09(+)=R107019 and cp10 in ('101','102','201') and ( (cp01='P' and R107010>=7) or (cp01='CFP' and R107010>=14) ) order by 3 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            Do While Not RsTemp.EOF
               ii = ii + 1
               oTable.Rows.add
               oTable.Rows(ii).Cells(1).Select
               .Selection.TypeText RsTemp("C1") '"工程師"
               oTable.Rows(ii).Cells(2).Select
               .Selection.TypeText RsTemp("C2") '"案號"
               oTable.Rows(ii).Cells(3).Select
               .Selection.TypeText RsTemp("C3") '"天數"
               RsTemp.MoveNext
            Loop
         End If
      
      '未會稿
      Else
         oTable.Columns(3).SetWidth ColumnWidth:=.CentimetersToPoints(1.8), RulerStyle:=wdAdjustProportional
         oTable.Rows(ii).Cells(3).Select
         .Selection.TypeText "完稿天數"
         
         oTable.Columns(4).SetWidth ColumnWidth:=.CentimetersToPoints(2.5), RulerStyle:=wdAdjustProportional
         oTable.Rows(ii).Cells(4).Select
         .Selection.TypeText "總承辦天數"
         
         oTable.Rows(ii).Cells(5).Select
         .Selection.TypeText "目前進度"
         
         strSql = "select R107001 C1,replace(R107003,'-0-00','') C2,R107013 C3,R107010 C4 from R090612,caseprogress where R107017='5' AND ID='" & strUserNum & "' and cp09(+)=R107019 and cp10 in ('101','102') and ( (cp01='P' and R107013>=4) or (cp01='CFP' and R107013>=6) ) order by 4 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            Do While Not RsTemp.EOF
               ii = ii + 1
               oTable.Rows.add
               oTable.Rows(ii).Cells(1).Select
               .Selection.TypeText RsTemp("C1") '"工程師"
               oTable.Rows(ii).Cells(2).Select
               .Selection.TypeText RsTemp("C2") '"案號"
               oTable.Rows(ii).Cells(3).Select
               .Selection.TypeText RsTemp("C3") '"完稿天數"
               oTable.Rows(ii).Cells(4).Select
               .Selection.TypeText RsTemp("C4") '"總承辦天數"
               RsTemp.MoveNext
            Loop
         End If
      End If
   End With
End Sub

