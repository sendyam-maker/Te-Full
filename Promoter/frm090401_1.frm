VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090401_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "美專IDS清單"
   ClientHeight    =   5892
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9012
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5892
   ScaleWidth      =   9012
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox txtInput 
      Appearance      =   0  '平面
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1740
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      Height          =   405
      Index           =   1
      Left            =   7530
      TabIndex        =   2
      Top             =   60
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確認"
      Height          =   405
      Index           =   0
      Left            =   6060
      TabIndex        =   1
      Top             =   60
      Width           =   1425
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   1365
      Index           =   0
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   990
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   2413
      _Version        =   393216
      Cols            =   5
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "No.|Document No.　　　　　　　　|Issue/Publication Date　　　　　　　　　　　　　　|發文日　　|確認"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   1365
      Index           =   1
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2730
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   2413
      _Version        =   393216
      Cols            =   7
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "No.|Document No.　　　　　　　　|Country Code|Publication Date|English brief explanation|發文日　　|確認"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   1365
      Index           =   2
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4470
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   2413
      _Version        =   393216
      Cols            =   5
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "No.|Author, title, date, or country where published　　　　　　　　  |English brief explanation|發文日　　|確認"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame1 
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   2370
      Width           =   8775
      Begin VB.CommandButton cmdExtend 
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   30
         TabIndex        =   21
         Top             =   30
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "刪除"
         Height          =   315
         Index           =   1
         Left            =   7890
         TabIndex        =   14
         Top             =   30
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增"
         Height          =   315
         Index           =   1
         Left            =   7020
         TabIndex        =   13
         Top             =   30
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "Foreign Patent Document"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   8775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   630
      Width           =   8775
      Begin VB.CommandButton cmdExtend 
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "刪除"
         Height          =   315
         Index           =   0
         Left            =   7890
         TabIndex        =   10
         Top             =   30
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增"
         Height          =   315
         Index           =   0
         Left            =   7020
         TabIndex        =   9
         Top             =   30
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "US Patent Document"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   8775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame3"
      Height          =   345
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   4110
      Width           =   8775
      Begin VB.CommandButton cmdExtend 
         Caption         =   "+"
         Height          =   315
         Index           =   2
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "刪除"
         Height          =   315
         Index           =   2
         Left            =   7890
         TabIndex        =   18
         Top             =   30
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增"
         Height          =   315
         Index           =   2
         Left            =   7020
         TabIndex        =   17
         Top             =   30
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "Non Patent LiteratureDocument"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   8775
      End
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   1050
      TabIndex        =   24
      Top             =   330
      Width           =   4665
      VariousPropertyBits=   27
      Size            =   "8229;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1050
      TabIndex        =   23
      Top             =   30
      Width           =   1545
      VariousPropertyBits=   27
      Size            =   "2725;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   4
      Left            =   150
      TabIndex        =   7
      Top             =   60
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   6
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "frm090401_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; lbl1(index)
'Created by Morgan 2020/12/28
Option Explicit

Public m_CaseNo As String '本所案號
Public m_CP09 As String 'IDS收文號
Public m_bConfirm As Boolean '來函判發確認
Public m_bMan As Boolean '是否管理人員
Public m_bQuery As Boolean '共同查詢

Dim m_InputGrid As MSHFlexGrid, m_InputCol As Integer, m_InputRow As Integer
Dim m_lstRow(2) As Integer
Dim m_NoData As Boolean
Dim m_Pa(4) As String
Dim m_CP27 As String

Private Sub cmdAdd_Click(Index As Integer)
   Dim iRow As Integer
   With grdDataList(Index)
   If .TextMatrix(1, 0) <> "" Then
      .Rows = .Rows + 1
   'Added by Morgan 2025/4/29
   ElseIf Index = 0 Then
      MsgBox "美專公開號的數字需剛好為11位數，且不能有斜線", vbExclamation, "美專公開號輸入提醒"
   'end 2025/4/29
   End If
   
   '補資料時自動上確認
   SetValue .Rows - 1, "確認", IIf(m_bQuery, "■", "□"), grdDataList(Index)
   If m_CP09 <> "" Then
      '設定收文號
      SetValue .Rows - 1, "IL01", m_CP09, grdDataList(Index)
   End If

   If m_CP27 <> "" Then
      '設定發文日
      SetValue .Rows - 1, "發文日", ChangeWStringToTDateString(m_CP27), grdDataList(Index)
   End If
   
   For iRow = 1 To .Rows - 1
      .TextMatrix(iRow, 0) = iRow
   Next
   .row = .Rows - 1
   .col = 1
   .TopRow = .row
   ClickGrid grdDataList(Index), m_lstRow(Index)
   End With
End Sub

Private Sub cmdDel_Click(Index As Integer)
   Dim iRow As Integer
   With grdDataList(Index)
   If .row > 0 And .TextMatrix(.row, 0) <> "" Then
      If GetValue(.row, "發文日", grdDataList(Index)) = "" Or m_bMan = True Or m_bQuery = True Then
         iRow = .row
         If .Rows = 2 Then
            For intI = 0 To .Cols - 1
               .TextMatrix(iRow, intI) = ""
            Next
            .row = 0
            ClickGrid grdDataList(Index), m_lstRow(Index)
         Else
            .RemoveItem .row
            If iRow > 2 Then .row = iRow - 1
            For iRow = 1 To .Rows - 1
               .TextMatrix(iRow, 0) = iRow
            Next
            m_lstRow(Index) = 0
            SelectRow grdDataList(Index), m_lstRow(Index)
         End If
      End If
   End If
   End With
End Sub

Private Sub cmdExtend_Click(Index As Integer)
   Dim bExtend As Boolean
   If cmdExtend(Index).Caption = "+" Then
      cmdExtend(Index).Caption = "-"
      bExtend = True
   Else
      cmdExtend(Index).Caption = "+"
      bExtend = False
   End If
   Select Case Index
   Case 0
      cmdExtend(1).Caption = "+"
      cmdExtend(2).Caption = "+"
   Case 1
      cmdExtend(0).Caption = "+"
      cmdExtend(2).Caption = "+"
   Case 2
      cmdExtend(0).Caption = "+"
      cmdExtend(1).Caption = "+"
   End Select
   SetFrame Index, bExtend
End Sub

Private Sub SetFrame(Index As Integer, bExtend As Boolean)
   If bExtend = False Then
      Frame1(0).Top = 510
      grdDataList(0).Top = 870
      grdDataList(0).Height = 1365
      
      Frame1(1).Top = 2250
      grdDataList(1).Top = 2610
      grdDataList(1).Height = 1365
      
      Frame1(2).Top = 3990
      grdDataList(2).Top = 4350
      grdDataList(2).Height = 1365
   ElseIf Index = 0 Then
      Frame1(0).Top = 510
      grdDataList(0).Top = 870
      grdDataList(0).Height = 1365 + grdDataList(0).RowHeight(0) * 4
      
      Frame1(1).Top = 2250 + grdDataList(1).RowHeight(0) * 4
      grdDataList(1).Top = 2610 + grdDataList(1).RowHeight(0) * 4
      grdDataList(1).Height = 1365 - grdDataList(1).RowHeight(0) * 2
      
      Frame1(2).Top = 3990 + grdDataList(2).RowHeight(0) * 2
      grdDataList(2).Top = 4350 + grdDataList(2).RowHeight(0) * 2
      grdDataList(2).Height = 1365 - grdDataList(2).RowHeight(0) * 2
   ElseIf Index = 1 Then
      Frame1(0).Top = 510
      grdDataList(0).Top = 870
      grdDataList(0).Height = 1365 - grdDataList(0).RowHeight(0) * 2
      
      Frame1(1).Top = 2250 - grdDataList(1).RowHeight(0) * 2
      grdDataList(1).Top = 2610 - grdDataList(1).RowHeight(0) * 2
      grdDataList(1).Height = 1365 + grdDataList(1).RowHeight(0) * 4
      
      Frame1(2).Top = 3990 + grdDataList(2).RowHeight(0) * 2
      grdDataList(2).Top = 4350 + grdDataList(2).RowHeight(0) * 2
      grdDataList(2).Height = 1365 - grdDataList(2).RowHeight(0) * 2
   ElseIf Index = 2 Then
      Frame1(0).Top = 510
      grdDataList(0).Top = 870
      grdDataList(0).Height = 1365 - grdDataList(0).RowHeight(0) * 2
      
      Frame1(1).Top = 2250 - grdDataList(1).RowHeight(0) * 2
      grdDataList(1).Top = 2610 - grdDataList(1).RowHeight(0) * 2
      grdDataList(1).Height = 1365 - grdDataList(1).RowHeight(0) * 2
      
      Frame1(2).Top = 3990 - grdDataList(2).RowHeight(0) * 4
      grdDataList(2).Top = 4350 - grdDataList(2).RowHeight(0) * 4
      grdDataList(2).Height = 1365 + grdDataList(2).RowHeight(0) * 4
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   
   If Index = 0 Then
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      If CheckGrid() = True Then
         If FormSave() = True Then
            If m_bConfirm = False And m_bQuery = False And m_CP09 <> "" Then
               If MsgBox("是否產生指示信？", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                  Me.Hide
                  With frm090401
                  .Hide
                  .m_IDSCP09 = m_CP09
                  .Text1 = SystemNumber(LBL1(0).Caption, 1)
                  .Text2 = SystemNumber(LBL1(0).Caption, 2)
                  .Text3 = SystemNumber(LBL1(0).Caption, 3)
                  .Text4 = SystemNumber(LBL1(0).Caption, 4)
                  .Option1(1).Value = True '點選英文
                  .Option6.Value = True '第1次點選讀取案件資料
                  .Option6.Value = True '點選代理人
                  .Command1.Value = True
                  .Command2.Value = True
                  End With
               End If
            End If
            Unload Me
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      End If
      Me.Enabled = True
      Screen.MousePointer = vbDefault
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Activate()
   Static bActivated As Boolean
   If m_NoData = True Then
      If m_bConfirm = True Or m_CaseNo <> "" Then
         Unload Me
      End If
   End If
   
   If bActivated = False Then
      bActivated = True
      If m_CP27 = "" And m_CP09 <> "" And cmdOK(0).Visible = True Then
         ChkFailIDS
      End If
   End If
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me, True
    ReadData
    If m_CP09 <> "" Then Me.Caption = Me.Caption & "(" & m_CP09 & ")"
    
    cmdExtend(0).Visible = True
    cmdExtend(1).Visible = True
    cmdExtend(2).Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090401_1 = Nothing
End Sub

Private Function CheckGrid() As Boolean
   Dim iRow As Integer, iCol As Integer, bErr As Boolean, stExt As String
   Dim iRow2 As Integer, stExt2 As String
   
   With grdDataList(0)
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) <> "" Then
         stExt = GetValue(iRow, "IL01", grdDataList(0))
         If stExt = m_CP09 Then
            For iCol = 1 To 2
               If .TextMatrix(iRow, iCol) = "" Then
                  bErr = True
               End If
               If bErr = True Then
                  .row = iRow: .col = iCol
                  If .RowHeight(0) * (1 + .row) > .Height Then
                     .TopRow = .row
                  End If
                  MsgBox .TextMatrix(0, iCol) & " 不可空白！", vbExclamation
                  Me.Enabled = True
                  ClickGrid grdDataList(0), m_lstRow(0)
                  Exit Function
               End If
            Next
            
            '資料不可重複
            stExt = .TextMatrix(iRow, 1)
            For iRow2 = 1 To iRow - 1
               stExt2 = .TextMatrix(iRow2, 1)
               If stExt = stExt2 Then
                  stExt2 = GetValue(iRow2, "IL01", grdDataList(0))
                  If stExt2 = m_CP09 Then
                     .row = iRow: .col = 1
                     If .RowHeight(0) * (1 + .row) > .Height Then
                        .TopRow = .row
                     End If
                     MsgBox "Document No.: " & .TextMatrix(iRow, 1) & vbCrLf & "資料重複！", vbCritical, "US Patent"
                     Me.Enabled = True
                     ClickGrid grdDataList(0), m_lstRow(0)
                     Exit Function
                  Else
                     stExt2 = .TextMatrix(iRow2, 4)
                     If .TextMatrix(iRow2, 4) = "■" Then
                        .row = iRow: .col = 1
                        If .RowHeight(0) * (1 + .row) > .Height Then
                           .TopRow = .row
                        End If
                        MsgBox "Document No.: " & .TextMatrix(iRow, 1) & vbCrLf & "與" & .TextMatrix(iRow2, 3) & "發文且已確認的資料重複！", vbCritical, "US Patent"
                        Me.Enabled = True
                        ClickGrid grdDataList(0), m_lstRow(0)
                        Exit Function
                     End If
                  End If
               End If
            Next
         End If
      End If
   Next
   End With
   
   With grdDataList(1)
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) <> "" Then
         stExt = GetValue(iRow, "IL01", grdDataList(1))
         If stExt = m_CP09 Then
            For iCol = 1 To 4
               If .TextMatrix(iRow, iCol) = "" Then
                  bErr = True
               End If
               If bErr = True Then
                  .row = iRow: .col = iCol
                  If .RowHeight(0) * (1 + .row) > .Height Then
                     .TopRow = .row
                  End If
                  MsgBox .TextMatrix(0, iCol) & " 不可空白！", vbExclamation
                  Me.Enabled = True
                  ClickGrid grdDataList(1), m_lstRow(1)
                  Exit Function
               End If
            Next
            
            '資料不可重複
            stExt = ""
            For iCol = 1 To 2
               stExt = stExt & .TextMatrix(iRow, iCol) & "|"
            Next
            For iRow2 = 1 To iRow - 1
               stExt2 = ""
               For iCol = 1 To 2
                  stExt2 = stExt2 & .TextMatrix(iRow2, iCol) & "|"
               Next
               If stExt = stExt2 Then
                  stExt2 = GetValue(iRow2, "IL01", grdDataList(1))
                  If stExt2 = m_CP09 Then
                     .row = iRow: .col = 1
                     If .RowHeight(0) * (1 + .row) > .Height Then
                        .TopRow = .row
                     End If
                     MsgBox "Document No.: " & .TextMatrix(iRow, 1) & vbCrLf & "Country Code: " & .TextMatrix(iRow, 2) & vbCrLf & "資料重複！", vbCritical, "Foreign Patent"
                     Me.Enabled = True
                     ClickGrid grdDataList(1), m_lstRow(1)
                     Exit Function
                  Else
                     If .TextMatrix(iRow2, 6) = "■" Then
                        .row = iRow: .col = 1
                        If .RowHeight(0) * (1 + .row) > .Height Then
                           .TopRow = .row
                        End If
                        MsgBox "Document No.: " & .TextMatrix(iRow, 1) & vbCrLf & "Country Code: " & .TextMatrix(iRow, 2) & vbCrLf & "與" & .TextMatrix(iRow2, 3) & "發文且已確認的資料重複！", vbCritical, "Foreign Patent"
                        Me.Enabled = True
                        ClickGrid grdDataList(1), m_lstRow(1)
                        Exit Function
                     End If
                  End If
               End If
            Next
         End If
      End If
   Next
   End With
   
   With grdDataList(2)
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 0) <> "" Then
         stExt = GetValue(iRow, "IL01", grdDataList(2))
         If stExt = m_CP09 Then
            For iCol = 1 To 2
               If .TextMatrix(iRow, iCol) = "" Then
                  bErr = True
               End If
               If bErr = True Then
                  .row = iRow: .col = iCol
                  If .RowHeight(0) * (1 + .row) > .Height Then
                     .TopRow = .row
                  End If
                  MsgBox .TextMatrix(0, iCol) & " 不可空白！", vbExclamation
                  
                  If iCol = 2 Then .col = 0
                  Me.Enabled = True
                  ClickGrid grdDataList(2), m_lstRow(2)
                  Exit Function
               End If
            Next
            
            '資料不可重複
            stExt = .TextMatrix(iRow, 1)
            For iRow2 = 1 To iRow - 1
               stExt2 = .TextMatrix(iRow2, 1)
               If stExt = stExt2 Then
                  stExt2 = GetValue(iRow2, "IL01", grdDataList(2))
                  If stExt2 = m_CP09 Then
                     .row = iRow: .col = 1
                     If .RowHeight(0) * (1 + .row) > .Height Then
                        .TopRow = .row
                     End If
                     MsgBox "Document No.: " & .TextMatrix(iRow, 1) & vbCrLf & "資料重複！", vbCritical, "Non Patent Literature"
                     Me.Enabled = True
                     ClickGrid grdDataList(2), m_lstRow(2)
                     Exit Function
                  Else
                     stExt2 = .TextMatrix(iRow2, 4)
                     If .TextMatrix(iRow2, 4) = "■" Then
                        .row = iRow: .col = 1
                        If .RowHeight(0) * (1 + .row) > .Height Then
                           .TopRow = .row
                        End If
                        MsgBox "Document No.: " & .TextMatrix(iRow, 1) & vbCrLf & "與" & .TextMatrix(iRow2, 3) & "發文且已確認的資料重複！", vbCritical, "Non Patent Literature"
                        Me.Enabled = True
                        ClickGrid grdDataList(2), m_lstRow(2)
                        Exit Function
                     End If
                  End If
               End If
            Next
         End If
      End If
   Next
   End With
   
   CheckGrid = True
End Function

Private Function FormSave() As Boolean
   Dim iRow As Integer, strErr As String, strTmp As String, IL(8) As String
   Dim idx As Integer
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
   '確認
   If m_bConfirm = True Then
      For idx = 0 To 2
         Erase IL
         With grdDataList(idx)
         For iRow = 1 To .Rows - 1
            If .TextMatrix(iRow, 0) <> "" Then
               IL(1) = GetValue(iRow, "IL01", grdDataList(idx))
               IL(2) = GetValue(iRow, "IL02", grdDataList(idx))
               IL(3) = GetValue(iRow, "IL03", grdDataList(idx))
               strExc(0) = GetValue(iRow, "確認", grdDataList(idx))
               If strExc(0) = "■" Then
                  IL(8) = "sysdate"
               Else
                  IL(8) = "null"
               End If
               strSql = "update IDSList set IL08=" & IL(8) & ",IL09='" & strUserNum & "' where IL01='" & IL(1) & "' and IL02='" & IL(2) & "' and IL03=" & IL(3)
               cnnConnection.Execute strSql, intI
            End If
         Next
         End With
      Next
   
   '維護
   Else
      strSql = "delete IDSList where IL01='" & m_CP09 & "'"
      cnnConnection.Execute strSql, intI
      
      Erase IL
      With grdDataList(0)
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, 0) <> "" Then
            IL(1) = GetValue(iRow, "IL01", grdDataList(0))
            '本收文號或新增列
            If IL(1) = m_CP09 Then
               IL(1) = m_CP09
               IL(2) = "1"
               IL(3) = Val(IL(3)) + 1
               IL(4) = .TextMatrix(iRow, 1)
               IL(6) = Replace(.TextMatrix(iRow, 2), "-", "")
               strExc(0) = GetValue(iRow, "確認", grdDataList(0))
               If strExc(0) = "■" Then
                  IL(8) = "sysdate"
               Else
                  IL(8) = "null"
               End If
               
               strSql = "insert into IDSList(IL01,IL02,IL03,IL04,IL06,IL08,IL10,IL11)" & _
                  " values('" & IL(1) & "','" & IL(2) & "'," & IL(3) & ",'" & ChgSQL(IL(4)) & "'" & _
                  "," & CNULL(IL(6)) & "," & IL(8) & ",sysdate,'" & strUserNum & "')"
               cnnConnection.Execute strSql, intI
            End If
         End If
      Next
      End With
      
      Erase IL
      With grdDataList(1)
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, 0) <> "" Then
            IL(1) = GetValue(iRow, "IL01", grdDataList(1))
            '本收文號或新增列
            If IL(1) = m_CP09 Then
               IL(1) = m_CP09
               IL(2) = "2"
               IL(3) = Val(IL(3)) + 1
               IL(4) = .TextMatrix(iRow, 1)
               IL(5) = .TextMatrix(iRow, 2)
               IL(6) = Replace(.TextMatrix(iRow, 3), "-", "")
               IL(7) = Left(.TextMatrix(iRow, 4), 1)
               strExc(0) = GetValue(iRow, "確認", grdDataList(1))
               If strExc(0) = "■" Then
                  IL(8) = "sysdate"
               Else
                  IL(8) = "null"
               End If
               
               strSql = "insert into IDSList(IL01,IL02,IL03,IL04,IL05,IL06,IL07,IL08,IL10,IL11)" & _
                  " values('" & IL(1) & "','" & IL(2) & "'," & IL(3) & ",'" & ChgSQL(IL(4)) & "'" & _
                  ",'" & IL(5) & "'," & CNULL(IL(6)) & ",'" & IL(7) & "'," & IL(8) & ",sysdate,'" & strUserNum & "')"
               cnnConnection.Execute strSql, intI
            End If
         End If
      Next
      End With
      
      Erase IL
      With grdDataList(2)
      For iRow = 1 To .Rows - 1
         If .TextMatrix(iRow, 0) <> "" Then
            IL(1) = GetValue(iRow, "IL01", grdDataList(2))
            '本收文號或新增列
            If IL(1) = m_CP09 Then
               IL(1) = m_CP09
               IL(2) = "3"
               IL(3) = Val(IL(3)) + 1
               IL(4) = .TextMatrix(iRow, 1)
               IL(7) = Left(.TextMatrix(iRow, 2), 1)
               strExc(0) = GetValue(iRow, "確認", grdDataList(2))
               If strExc(0) = "■" Then
                  IL(8) = "sysdate"
               Else
                  IL(8) = "null"
               End If
               
               strSql = "insert into IDSList(IL01,IL02,IL03,IL04,IL07,IL08,IL10,IL11)" & _
                  " values('" & IL(1) & "','" & IL(2) & "'," & IL(3) & ",'" & ChgSQL(IL(4)) & "'" & _
                  ",'" & IL(7) & "'," & IL(8) & ",sysdate,'" & strUserNum & "')"
               cnnConnection.Execute strSql, intI
            End If
         End If
      Next
      End With
   End If
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   strErr = Err.Description
   cnnConnection.RollbackTrans
   MsgBox strErr, vbCritical
   
End Function

Private Sub ReadData()
   Dim iRows1 As Integer, iRows2 As Integer, iRows3 As Integer
   Dim iCol As Integer
   Dim stSQL As String, stCon As String
   Dim rsQuery As ADODB.Recordset
   Dim bCP27 As String
   
   SetGrid
   
   '案件資料
   If m_CP09 = "" Then
      stSQL = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CaseNo,pa01,pa02,pa03,pa04,pa05" & _
         " from patent where pa01='" & SystemNumber(m_CaseNo, 1) & "' and pa02='" & SystemNumber(m_CaseNo, 2) & "'" & _
         " and pa03='" & SystemNumber(m_CaseNo, 3) & "' and pa04='" & SystemNumber(m_CaseNo, 4) & "'"
   Else
      stSQL = "select cp27,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CaseNo,pa01,pa02,pa03,pa04,pa05" & _
         " from caseprogress,patent where cp09='" & m_CP09 & "' and pa01(+)=cp01 and pa02(+)=cp02" & _
         " and pa03(+)=cp03 and pa04(+)=cp04"
   End If
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      LBL1(0) = "" & rsQuery("CaseNo")
      LBL1(1) = "" & rsQuery("pa05")
      m_Pa(1) = "" & rsQuery("pa01")
      m_Pa(2) = "" & rsQuery("pa02")
      m_Pa(3) = "" & rsQuery("pa03")
      m_Pa(4) = "" & rsQuery("pa04")
      If m_CP09 <> "" Then
         m_CP27 = "" & rsQuery("CP27")
      End If
   End If
   
   If m_CP09 = "" Then
      Set rsQuery = PUB_GetIDSList(m_CaseNo, intI)
   Else
      
      If m_bConfirm = True Or m_bQuery = True Then
         stCon = " and il01='" & m_CP09 & "'"
      End If
      stSQL = "select sqldatet(b.cp27) DDate,i.* from caseprogress a,caseprogress b,IDSList i" & _
         " where a.cp09='" & m_CP09 & "' and b.cp01(+)=a.cp01 and b.cp02(+)=a.cp02" & _
         " and b.cp03(+)=a.cp03 and b.cp04(+)=a.cp04 and b.cp10(+)='214' and b.cp57 is null" & _
         " and (b.cp09=a.cp09 or b.cp27<=nvl(a.cp27,b.cp27)) and IL01(+)=b.cp09" & stCon & _
         " order by il02,DDate,il01,il03"
         
      intI = 1
      Set rsQuery = ClsLawReadRstMsg(intI, stSQL)
   End If
   '共同查詢、公文來函判發(可確認)、個人進度維護(已發文)都不可改資料
   '但110/1/11 以前發文資料可從共同查詢進度資料畫面維護(補舊資要)
   If m_CP09 = "" Or m_bConfirm = True Or (Val(m_CP27) >= 20210111 And m_bMan = False) Or (m_bQuery = True And Not Val(m_CP27) >= 20210111) Then SetReadOnly
   If intI = 1 Then
      With rsQuery
      Do While Not .EOF
         Select Case .Fields("il02")
         Case 1 'US Patent
            iRows1 = iRows1 + 1
            With grdDataList(0)
               .Rows = iRows1 + 1
               .TextMatrix(iRows1, 0) = iRows1 'No.
               .TextMatrix(iRows1, 1) = "" & rsQuery("il04") 'Document No.
               
               'Issue/Publication date
               If Not IsNull(rsQuery("il06")) Then
                  .TextMatrix(iRows1, 2) = Format(rsQuery("il06"), "@@@@-@@-@@")
               End If
               
               .TextMatrix(iRows1, 3) = "" & rsQuery("DDate") '發文日
               '已發文變灰
               If .TextMatrix(iRows1, 3) <> "" Then
                  .row = iRows1
                  For iCol = 1 To .Cols - 1
                     .col = iCol
                     .CellBackColor = &HE0E0E0
                  Next
               End If
               
               If Not IsNull(rsQuery("il08")) Then
                  .TextMatrix(iRows1, 4) = "■" '確認
               Else
                  .TextMatrix(iRows1, 4) = "□" '確認
               End If
               .TextMatrix(iRows1, 5) = rsQuery("IL01")
               .TextMatrix(iRows1, 6) = rsQuery("IL02")
               .TextMatrix(iRows1, 7) = rsQuery("IL03")
            End With
            
         Case 2 'Foreign Patent
            iRows2 = iRows2 + 1
            With grdDataList(1)
               .Rows = iRows2 + 1
               .TextMatrix(iRows2, 0) = iRows2 'No.
               .TextMatrix(iRows2, 1) = "" & rsQuery("il04") 'Document No.
               .TextMatrix(iRows2, 2) = "" & rsQuery("il05") 'Country code
               
               'Publication date
               If Not IsNull(rsQuery("il06")) Then
                  .TextMatrix(iRows2, 3) = Format(rsQuery("il06"), "@@@@-@@-@@")
               End If
               
               'English brief explanation
               If rsQuery("il07") = "Y" Then
                  .TextMatrix(iRows2, 4) = "Yes"
               ElseIf rsQuery("il07") = "N" Then
                  .TextMatrix(iRows2, 4) = "No"
               End If
               
               .TextMatrix(iRows2, 5) = "" & rsQuery("DDate") '發文日
               '已發文變灰
               If .TextMatrix(iRows2, 5) <> "" Then
                  .row = iRows2
                  For iCol = 1 To .Cols - 1
                     .col = iCol
                     .CellBackColor = &HE0E0E0
                  Next
               End If
               
               If Not IsNull(rsQuery("il08")) Then
                  .TextMatrix(iRows2, 6) = "■" '確認
               Else
                  .TextMatrix(iRows2, 6) = "□" '確認
               End If
               .TextMatrix(iRows2, 7) = rsQuery("IL01")
               .TextMatrix(iRows2, 8) = rsQuery("IL02")
               .TextMatrix(iRows2, 9) = rsQuery("IL03")
            End With
            
         Case 3 'Non Patent Literature
            iRows3 = iRows3 + 1
            With grdDataList(2)
               .Rows = iRows3 + 1
               .TextMatrix(iRows3, 0) = iRows3 'No.
               .TextMatrix(iRows3, 1) = "" & rsQuery("il04") 'Author, title, date, or country where published
               
               'English brief explanation
               If rsQuery("il07") = "Y" Then
                  .TextMatrix(iRows3, 2) = "Yes"
               ElseIf rsQuery("il07") = "N" Then
                  .TextMatrix(iRows3, 2) = "No" '
               End If
               
               .TextMatrix(iRows3, 3) = "" & rsQuery("DDate") '發文日
               '已發文變灰
               If .TextMatrix(iRows3, 3) <> "" Then
                  .row = iRows3
                  For iCol = 1 To .Cols - 1
                     .col = iCol
                     .CellBackColor = &HE0E0E0
                  Next
               End If
               If Not IsNull(rsQuery("il08")) Then
                  .TextMatrix(iRows3, 4) = "■" '確認
               Else
                  .TextMatrix(iRows3, 4) = "□" '確認
               End If
               .TextMatrix(iRows3, 5) = rsQuery("IL01")
               .TextMatrix(iRows3, 6) = rsQuery("IL02")
               .TextMatrix(iRows3, 7) = rsQuery("IL03")
            End With
         End Select
         .MoveNext
      Loop
      End With
      
      grdDataList(0).row = 0: m_lstRow(0) = grdDataList(0).row
      grdDataList(1).row = 0: m_lstRow(1) = grdDataList(1).row
      grdDataList(2).row = 0: m_lstRow(2) = grdDataList(2).row
   Else
      m_NoData = True
   End If
End Sub

Private Sub SetBox(ByRef FlexGrid As MSHFlexGrid, ByRef InputBox As TextBox, Optional bReadOnly As Boolean = False)
   Dim lngLeft As Long, lngTop As Long, iCol As Integer, ii As Integer
   With FlexGrid
      InputBox.Locked = bReadOnly
      InputBox.FontName = .CellFontName
      InputBox.FontSize = .CellFontSize
      If .CellAlignment < 3 Then
         InputBox.Alignment = 0
      ElseIf .CellAlignment > 5 Then
         InputBox.Alignment = 1
      Else
         InputBox.Alignment = 2
      End If
      InputBox.Text = .TextMatrix(.row, .col)
      InputBox.Tag = InputBox.Text
      InputBox.Width = .ColWidth(.col)
      InputBox.Height = .RowHeight(.row)
      InputBox.Tag = InputBox.Text
      InputBox.Visible = True
      If InputBox.Enabled Then InputBox.SetFocus
      TextInverse InputBox
      
      lngLeft = .Left + 20
      lngTop = .Top + .RowHeight(0) + 20
      
      lngLeft = lngLeft + .ColPos(.col)
      
      For ii = .TopRow To .row - 1
         lngTop = lngTop + .RowHeight(ii)
      Next
      InputBox.Left = lngLeft: InputBox.Top = lngTop
      m_InputRow = .row
      m_InputCol = .col
      Set m_InputGrid = FlexGrid
   End With
   
End Sub

Private Sub GrdDataList_Click(Index As Integer)
   If cmdOK(0).Visible = False Then Exit Sub
   With grdDataList(Index)
   .row = .MouseRow
   .col = .MouseCol
   End With
   ClickGrid grdDataList(Index), m_lstRow(Index)
End Sub

Private Sub grdDataList_Scroll(Index As Integer)
   If Not m_InputGrid Is Nothing Then
      If Index = m_InputGrid.Index Then
         'cmdOK(0).SetFocus
         txtInput_LostFocus
      End If
   End If
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
   If txtInput.Locked = True Then
      If UpperCase(KeyCode) = Asc("Y") Then
         txtInput = "Yes"
      ElseIf UpperCase(KeyCode) = Asc("N") Then
         txtInput = "No"
      ElseIf KeyCode = 8 Or KeyCode = 46 Then
         txtInput = ""
      Else
         Beep
      End If
   End If
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      If ChkInput() = True Then
         goNextCol m_InputGrid.Index
      End If
   ElseIf KeyAscii = vbKeyEscape Then
      txtInput = txtInput.Tag
      TextInverse txtInput
      
   '國碼控制轉大寫
   ElseIf (m_InputGrid.Index = 1 And m_InputCol = 2) Then
       KeyAscii = UpperCase(KeyAscii)
   
   End If
End Sub

Private Sub txtInput_LostFocus()
   If txtInput.Visible = True Then
      If ChkInput = True Then txtInput.Visible = False
   End If
End Sub

Private Sub SetGrid()
   Dim arrGridHeadWidth
   Dim iCol As Integer, iUbound As Integer
   
   Erase m_lstRow
   arrGridHeadWidth = Array(350, 2550, 4200, 900, 450)
   iUbound = UBound(arrGridHeadWidth)
   With grdDataList(0)
   .FormatString = "No.|Document No.|Issue/Publication Date|發文日|確認|IL01|IL02|IL03"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         If iCol = 1 Then
            .ColAlignment(iCol) = flexAlignLeftCenter
         Else
            .ColAlignment(iCol) = flexAlignCenterCenter
         End If
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
   
   arrGridHeadWidth = Array(350, 2550, 1100, 1250, 1850, 900, 450)
   iUbound = UBound(arrGridHeadWidth)
   With grdDataList(1)
   .FormatString = "No.|Document No.|Country Code|Publication Date|English brief explanation|發文日|確認|IL01|IL02|IL03"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         If iCol = 1 Then
            .ColAlignment(iCol) = flexAlignLeftCenter
         Else
            .ColAlignment(iCol) = flexAlignCenterCenter
         End If
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
   
   arrGridHeadWidth = Array(350, 4900, 1850, 900, 450)
   iUbound = UBound(arrGridHeadWidth)
   With grdDataList(2)
   .FormatString = "No.|Author, title, date, or country where published|English brief explanation|發文日|確認|IL01|IL02|IL03"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         If iCol = 1 Then
            .ColAlignment(iCol) = flexAlignLeftCenter
         Else
            .ColAlignment(iCol) = flexAlignCenterCenter
         End If
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub

Private Sub ClickGrid(ByRef FlexGrid As MSHFlexGrid, ByRef pPrevRow As Integer)
   Dim bShow As Boolean, bYesNo As Boolean, bCheck As Boolean
   
   With FlexGrid
   If .row > 0 And .TextMatrix(.row, 0) <> "" Then
      If GetValue(.row, "IL01", FlexGrid) = m_CP09 Then
         If .Index = 0 Then
            If (.col >= 1 And .col <= 2) Then
               If .TextMatrix(.row, 3) = "" Or m_bMan Or m_bQuery Then
                  bShow = True
               End If
            ElseIf .col = 4 Then
               If .TextMatrix(.row, 3) <> "" And (m_bConfirm Or m_bMan Or m_bQuery) Then
                  bCheck = True
               End If
            End If

         ElseIf .Index = 1 Then
            If (.col >= 1 And .col <= 4) Then
               If .TextMatrix(.row, 5) = "" Or m_bMan Or m_bQuery Then
                  bShow = True
                  If .col = 4 Then
                     bYesNo = True
                  End If
               End If
            ElseIf .col = 6 Then
               If .TextMatrix(.row, 5) <> "" And (m_bConfirm Or m_bMan Or m_bQuery) Then
                  bCheck = True
               End If
            End If
         ElseIf .Index = 2 Then
            If (.col >= 1 And .col <= 2) Then
               If .TextMatrix(.row, 3) = "" Or m_bMan Or m_bQuery Then
                  bShow = True
                  If .col = 2 Then
                     bYesNo = True
                  End If
               End If
            ElseIf .col = 4 Then
               If .TextMatrix(.row, 3) <> "" And (m_bConfirm Or m_bMan Or m_bQuery) Then
                  bCheck = True
               End If
            End If
         End If
         
         If bShow = True Then
            SetBox FlexGrid, txtInput, bYesNo
         
         ElseIf bCheck = True Then
            If .Text = "□" Then
               .Text = "■"
            ElseIf .Text = "■" Then
               .Text = "□"
            End If
         End If
      End If
   End If
   End With
   '勾選列變色
   SelectRow FlexGrid, pPrevRow
End Sub

Private Function SetValue(pRow As Integer, pFieldName As String, pValue As String, ByRef FlexGrid As MSHFlexGrid) As Boolean
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         .TextMatrix(pRow, iRow) = pValue
         SetValue = True
         Exit Function
      End If
   Next
   End With
End Function

Private Function GetValue(pRow As Integer, pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As String
   Dim iCol As Integer
   With FlexGrid
   For iCol = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iCol)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, iCol)
         Exit For
      End If
   Next
   End With
End Function

Private Sub SelectRow(ByRef FlexGrid As MSHFlexGrid, ByRef pPrevRow As Integer)
   Dim nRow As Integer, nCol As Integer
   
   If pPrevRow = FlexGrid.row Then Exit Sub
   
   With FlexGrid
   nRow = .row
   nCol = .col
   '本次列變色
   If .row > 0 And .TextMatrix(.row, 0) <> "" Then
      .col = 0
      .CellBackColor = .BackColorSel
      .CellForeColor = .ForeColorSel
   End If
   
   '前次列還原
   If pPrevRow > 0 Then
      .row = pPrevRow
      .col = 0
      .CellBackColor = .BackColorFixed
      .CellForeColor = .ForeColorFixed
   End If
   .row = nRow
   .col = nCol
   pPrevRow = .row
   End With
End Sub

Private Sub SetReadOnly()
   cmdAdd(0).Visible = False
   cmdAdd(1).Visible = False
   cmdAdd(2).Visible = False
   cmdDel(0).Visible = False
   cmdDel(1).Visible = False
   cmdDel(2).Visible = False
   If m_bConfirm = False Then
      cmdOK(0).Visible = False
   End If
End Sub

'跳下一格
Private Sub goNextCol(Index As Integer)
   Dim bGo As Boolean
   Dim iRows As Integer
   
   With grdDataList(Index)
   
   iRows = Round(.Height / .RowHeight(0)) - 2
   
   Select Case Index
      Case 0, 2
         If .col = 1 Then
            .col = 2
            bGo = True
         ElseIf .Rows - 1 > .row Then
            .row = .row + 1
            .col = 1
            If .row > .TopRow + iRows Then .TopRow = .TopRow + 1
            bGo = True
         End If
      Case 1
         If .col < 4 Then
            .col = .col + 1
            bGo = True
         ElseIf .Rows - 1 > .row Then
            .row = .row + 1
            .col = 1
            If .row > .TopRow + iRows Then .TopRow = .TopRow + 1
            bGo = True
         End If
   End Select
   End With
   If bGo = True Then
      ClickGrid grdDataList(Index), m_lstRow(Index)
   Else
      txtInput.Visible = False
   End If
End Sub

Private Function ChkInput() As Boolean
   Dim stExt As String
   '檢查日期格式
   If (m_InputGrid.Index = 0 And m_InputCol = 2) Or (m_InputGrid.Index = 1 And m_InputCol = 3) Then
      If txtInput <> "" Then
         stExt = Replace(txtInput, "-", "/")
         If IsNumeric(stExt) Then
            stExt = Format(stExt, "@@@@/@@/@@")
         End If
         If IsDate(stExt) = False Then
            MsgBox "日期格式輸入錯誤！" & vbCrLf & vbCrLf & "格式：yyyy-mm-dd" & vbCrLf & "範例：2007-04-12", vbCritical
            txtInput.SetFocus
            TextInverse txtInput
            Exit Function
         End If
         txtInput.Text = Format(stExt, "YYYY-MM-DD")
      End If
   
   'Added by Morgan 2024/7/31 US Patent Documents的Document No.不能輸入"/ "，若有輸入?／?，則跳提醒?Document No.不能有／,請重新輸入。"
   ElseIf m_InputGrid.Index = 0 And m_InputCol = 1 Then
      If InStr(txtInput, "/") > 0 Then
         MsgBox "Document No.不能有""/"",請重新輸入！", vbCritical, "US Patent Documents檢查"
         txtInput.SetFocus
         TextInverse txtInput
         Exit Function
      End If
   'end 2024/7/31
   End If
   m_InputGrid.TextMatrix(m_InputRow, m_InputCol) = txtInput.Text
   ChkInput = True
End Function

'檢查已發未確認之IDS
Private Sub ChkFailIDS()
   Dim stSQL As String
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select * from IDSList where IL01='" & m_CP09 & "'"
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, stSQL)
   If intI = 0 Then
      stSQL = "select * from IDSList where IL01 in (select substr(max(cp27||cp09),9)" & _
         " from caseprogress where cp01='" & m_Pa(1) & "' and cp02='" & m_Pa(2) & "'" & _
         " and cp03='" & m_Pa(3) & "' and cp04='" & m_Pa(4) & "' and cp10='214' and cp27>0 and cp57 is null) and IL08 is null"
      intI = 1
      Set rsQuery = ClsLawReadRstMsg(intI, stSQL)
      If intI = 1 Then
         If MsgBox("是否自動新增前次發文未確認之IDS資料？", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
            stSQL = "insert into IDSList(IL01,IL02,IL03,IL04,IL05,IL06,IL07,IL10,IL11)" & _
               " select '" & m_CP09 & "',IL02,IL03,IL04,IL05,IL06,IL07,sysdate,'" & strUserNum & "'" & _
               " from IDSList where IL01='" & rsQuery("IL01") & "' and il08 is null"
            cnnConnection.Execute stSQL, intI
            ReadData
         End If
      End If
   End If
End Sub
