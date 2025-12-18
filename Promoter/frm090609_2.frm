VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090609_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人工作量查詢-逾本所期限明細"
   ClientHeight    =   5715
   ClientLeft      =   270
   ClientTop       =   1530
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
      Height          =   400
      Index           =   2
      Left            =   8064
      TabIndex        =   3
      Top             =   50
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   7284
      TabIndex        =   2
      Top             =   50
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "明細(&L)"
      Height          =   400
      Index           =   0
      Left            =   6504
      TabIndex        =   1
      Top             =   50
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   5172
      Left            =   72
      TabIndex        =   0
      Top             =   504
      Width           =   9216
      _ExtentX        =   16245
      _ExtentY        =   9128
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
Attribute VB_Name = "frm090609_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/12 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Public StrForm1 As String, i As Integer
'add by nickc 2007/08/22 紀錄作用按鍵
Public cmdState As Integer
Dim j As Integer


Private Sub cmdOK_Click(Index As Integer)
cmdState = Index
PubShowNextData
Exit Sub
End Sub

'edit by nickc 2007/08/22 秀玲說改跟共同查詢相同
Public Sub PubShowNextData()
Dim strCP01 As String 'Add By Sindy 2012/5/21
Select Case cmdState
Case 0
     With grd1
        For i = 1 To .Rows - 1
            .col = 0
            .row = i
            If .Text = "V" Then
               'edit by nickc 2007/08/22 秀玲說改跟共同查詢相同
               'Me.Hide
               .col = 0
               .Text = ""
               For j = 0 To .Cols - 1
                   .col = j
                   .CellBackColor = QBColor(15)
               Next j
               .col = 22
               StrForm1 = .Text
               If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               'Modify By Sindy 2012/5/21 +if,frm100101_K
               strCP01 = GetCaseProData(Trim(Pub_RplStr(.Text)), "CP01")
               If strCP01 = "P" Or strCP01 = "PS" Or strCP01 = "FG" Or _
                  strCP01 = "FCP" Or strCP01 = "CFP" Or strCP01 = "CPS" Or _
                  Val(strSrvDate(1)) < Val(TMdebateStarDT) Then  '專利處工作進度
                  frm100101_F.Show
                  frm100101_F.Process Pub_RplStr(.Text)
               Else
                  frm100101_K.Show
                  frm100101_K.Process Pub_RplStr(.Text)
               End If
               '2012/5/21 End
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
        Next i
    End With
Case 1
     Me.Hide
     frm090609_1.Show
Case 2
     Me.Hide
     If frm090609.ObjForm = 2 Then
        frm090609.Show
        Unload Me
        Exit Sub
     Else
        frm090609_1.Show
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
'************ 90.11.29 nick
'薛說已發文的不秀
'***************************
strSql = "SELECT ' ',NVL(ST02,R104001),R104002,R104003,R104004,R104005,R104006,R104007,R104008,R104009,R104010,R104011,R104012,R104013,R104014,R104015,R104016,R104017,R104018,R104019,R104020,R104021,R104022 FROM R090609_2,STAFF WHERE ID='" & strUserNum & "' and (R104018 is null or R104018 ='') AND R104001=ST01(+) ORDER BY ST06,ST03,NVL(ST02,R104001),R104002 "
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        Set grd1.Recordset = adoRecordset
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/17
    End If
End With
CheckOC
SetGrd1
Me.grd1.row = 1: Me.grd1.col = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090609_2 = Nothing
End Sub

Private Sub SetGrd1()
With grd1
    .Cols = 23
    .row = 0
    .col = 0:   .Text = " "
    .ColWidth(0) = 200
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "承辦人"
    .ColWidth(1) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "目次"
    .ColWidth(2) = 600
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "收文類別"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "收文日"
    .ColWidth(4) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "本所案號"
    .ColWidth(5) = 1550
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "案件名稱"
    .ColWidth(6) = 1500
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "Y/N"
    .ColWidth(7) = 400
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "種類"
    .ColWidth(8) = 500
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "案件性質"
    .ColWidth(9) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 10:  .Text = "承辦期限"
    .ColWidth(10) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 11:  .Text = "本所期限"
    .ColWidth(11) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 12:  .Text = "法定期限"
    .ColWidth(12) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 13:  .Text = "齊備日"
    .ColWidth(13) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 14:  .Text = "完稿日"
    .ColWidth(14) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 15:  .Text = "會稿日"
    .ColWidth(15) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 16:  .Text = "核稿人"
    .ColWidth(16) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 17:  .Text = "會稿完成日"
    .ColWidth(17) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 18:  .Text = "發文日"
    .ColWidth(18) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 19:  .Text = "承辦天數"
    .ColWidth(19) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 20:  .Text = "備註"
    .ColWidth(20) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 21:  .Text = "智權人員"
    .ColWidth(21) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 22:  .Text = "id"
    .ColWidth(22) = 0
    .CellAlignment = flexAlignCenterCenter
End With
End Sub

Private Sub Grd1_Click()
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
