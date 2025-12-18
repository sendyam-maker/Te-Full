VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090702_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員工作量查詢-逾本所明細"
   ClientHeight    =   8880
   ClientLeft      =   -1530
   ClientTop       =   1515
   ClientWidth     =   15000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   15000
   Begin VB.CommandButton cmdok 
      Caption         =   "明細(&L)"
      Height          =   350
      Index           =   0
      Left            =   11820
      TabIndex        =   2
      Top             =   10
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   1
      Left            =   12600
      TabIndex        =   1
      Top             =   10
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Index           =   2
      Left            =   13380
      TabIndex        =   0
      Top             =   10
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   8280
      Left            =   0
      TabIndex        =   3
      Top             =   435
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   14605
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      AllowUserResizing=   1
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
Attribute VB_Name = "frm090702_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/07 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Dim i As Integer
Public StrForm1 As String


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     With grd1
        For i = 1 To .Rows - 1
            .col = 0
            .row = i
            If .Text = "V" Then
                Me.Hide
                .col = 21
                StrForm1 = .Text
                frm090702_3.Show
            Do
            DoEvents
            Loop Until Not frm090702_3.Visible
            Unload frm090702_3
            End If
        Next i
    End With
Case 1
     Me.Hide
     frm090702_1.Show
Case 2
     Me.Hide
     If frm090702.ObjForm = 2 Then
        frm090702.Show
        Unload Me
        Exit Sub
     Else
        frm090702_1.Show
        Unload Me
        Exit Sub
     End If
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
SetGrd1
StrMenu
End Sub

Sub StrMenu()
strSql = "SELECT '',R104001,R104002,R104003,R104004,R104005,R104006,R104007,R104008,R104009,R104010,R104011,R104012,R104013,R104014,R104015,R104016,R104017,R104018,R104019,R104020,r104021 FROM R090702_2 WHERE ID='" & strUserNum & "' ORDER BY R104001,R104002 "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/20
        Set grd1.Recordset = adoRecordset
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/20
    End If
End With
CheckOC
SetGrd1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090702_2 = Nothing
End Sub

Private Sub SetGrd1()
With grd1
    .Cols = 22
    .row = 0
    .col = 0:   .Text = " "
    .ColWidth(0) = 200
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "繪圖人員"
    .ColWidth(1) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "收文類別"
    .ColWidth(2) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "收文日"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "本所案號"
    .ColWidth(4) = 1550
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "案件名稱"
    .ColWidth(5) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "是否計算案件數"
    .ColWidth(6) = 1400
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "案件性質"
    .ColWidth(7) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "承辦人"
    .ColWidth(8) = 600
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "承辦期限"
    .ColWidth(9) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 10:  .Text = "點數"
    .ColWidth(10) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 11:  .Text = "草圖齊備日"
    .ColWidth(11) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 12:  .Text = "草圖完稿日"
    .ColWidth(12) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 13:  .Text = "草圖作業天數"
    .ColWidth(13) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 14:  .Text = "墨圖齊備日"
    .ColWidth(14) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 15:  .Text = "墨圖完稿日"
    .ColWidth(15) = 1000
    .CellAlignment = flexAlignCenterCenter
    .col = 16:  .Text = "墨圖作業天數"
    .ColWidth(16) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 17:  .Text = "本所期限"
    .ColWidth(17) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 18:  .Text = "發文日"
    .ColWidth(18) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 19:  .Text = "備註"
    .ColWidth(19) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 20:  .Text = "智權人員"
    .ColWidth(20) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 21:  .Text = "id"
    .ColWidth(21) = 0
    .CellAlignment = flexAlignCenterCenter
End With
End Sub

Private Sub grd1_Click()
With grd1
    .Visible = False
    .col = 0
    .row = .MouseRow
    If .MouseRow <> 0 Then
        If .Text = "V" Then
            .Text = ""
            For i = 0 To .Cols - 1
                .col = i
                .CellBackColor = QBColor(15)
            Next i
        Else
            .Text = "V"
            For i = 0 To .Cols - 1
                .col = i
                .CellBackColor = &HFFC0C0
            Next i
        End If
    End If
    .Visible = True
End With
End Sub
