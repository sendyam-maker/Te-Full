VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090617 
   BorderStyle     =   1  '單線固定
   Caption         =   "獎金輸入"
   ClientHeight    =   5730
   ClientLeft      =   870
   ClientTop       =   3900
   ClientWidth     =   9315
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9315
   Begin VB.Frame Frame1 
      Height          =   4776
      Left            =   36
      TabIndex        =   8
      Top             =   900
      Width           =   9252
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   3708
         Left            =   36
         TabIndex        =   15
         Top             =   1056
         Width           =   9144
         _ExtentX        =   16140
         _ExtentY        =   6535
         _Version        =   393216
         Cols            =   5
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
         _Band(0).Cols   =   5
      End
      Begin VB.CommandButton cmd 
         Caption         =   "加入(&A)"
         Default         =   -1  'True
         Height          =   400
         Index           =   0
         Left            =   7488
         TabIndex        =   4
         Top             =   168
         Width           =   810
      End
      Begin VB.CommandButton cmd 
         Caption         =   "刪除(&D)"
         Height          =   400
         Index           =   1
         Left            =   8340
         TabIndex        =   5
         Top             =   168
         Width           =   810
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   3
         Left            =   3255
         TabIndex        =   3
         Top             =   645
         Width           =   975
      End
      Begin VB.TextBox txt1 
         Height          =   264
         Index           =   2
         Left            =   1095
         TabIndex        =   2
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "額外獎勵："
         Height          =   180
         Index           =   3
         Left            =   2310
         TabIndex        =   14
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "獎金："
         Height          =   180
         Index           =   1
         Left            =   450
         TabIndex        =   13
         Top             =   645
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "姓名："
         Height          =   180
         Index           =   0
         Left            =   2595
         TabIndex        =   12
         Top             =   255
         Width           =   555
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   3345
         TabIndex        =   11
         Top             =   255
         Width           =   1260
         VariousPropertyBits=   27
         Size            =   "2222;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "員工編號："
         Height          =   180
         Index           =   5
         Left            =   105
         TabIndex        =   10
         Top             =   270
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   0
         Left            =   1170
         TabIndex        =   9
         Top             =   255
         Width           =   1260
         VariousPropertyBits=   27
         Size            =   "2222;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   60
      MaxLength       =   3
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1176
      MaxLength       =   1
      TabIndex        =   1
      Top             =   600
      Width           =   405
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6375
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090617.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090617.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090617.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090617.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090617.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090617.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090617.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090617.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090617.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090617.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090617.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   1085
      ButtonWidth     =   1138
      ButtonHeight    =   1032
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "年 第"
      Height          =   180
      Index           =   2
      Left            =   696
      TabIndex        =   7
      Top             =   636
      Width           =   432
   End
   Begin VB.Label Label1 
      Caption         =   "季"
      Height          =   180
      Index           =   4
      Left            =   1620
      TabIndex        =   6
      Top             =   636
      Width           =   252
   End
End
Attribute VB_Name = "frm090617"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/17 改成Form2.0 (grd1,lbl1)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim i As Integer, j As Integer, k As Integer, s As Integer, TextOk As Boolean, SeekAction As Integer, SeekRec As Variant
Dim StrSQL6 As String, strTemp1 As Variant, SeekTemp As String, DELMenu() As String, DELTemp() As String, SeekBmk1 As Variant, SeekBmk2 As Variant, SeekBmk3 As Variant
Dim strTemp(0 To 6) As String, PLeft(0 To 7) As Integer, Page As Integer, iPrint As Integer
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'Add By Cheng 2002/05/24
Dim m_blnCancel As Boolean

Private Sub cmd_Click(Index As Integer)
If SeekAction <> 1 And SeekAction <> 0 Then
    Exit Sub
End If
With grd1
Select Case Index
Case 0
     For i = 0 To .Rows - 1
        .col = 0
        .row = i
        If .CellBackColor = &HFFC0C0 Then
            For j = 2 To 3
                .col = j
                .Text = txt1(j)
            Next j
            .col = 4
            .Text = str(Val(txt1(2)) + Val(txt1(3)))
            Exit For
        End If
     Next i
Case 1
     For i = 0 To .Rows - 1
        .col = 0
        .row = i
        If .CellBackColor = &HFFC0C0 Then
            .col = 1
            For j = 2 To 4
                .col = j
                .Text = "0"
            Next j
            Grd1_Click
            Exit For
        End If
     Next i
Case Else
End Select
End With
End Sub

Sub REFormLoad()
    SeekAction = 4
    ProcessUp
    ProcessDown
    TxtLock 3
    TxtSitu True
    ReDim DELMenu(0) As String
    ReDim DELTemp(0) As String
End Sub




Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyF2
     If SeekAction >= 4 Then
        YNEdit 0
     End If
Case vbKeyF3
     If SeekAction >= 4 Then
        YNEdit 1
     End If
Case vbKeyF5
     If SeekAction >= 4 Then
        YNEdit 2
     End If
Case vbKeyF4
     If SeekAction >= 4 Then
        YNEdit 3
     End If
Case vbKeyHome
     If SeekAction >= 4 Then
        MoveRec 0
     End If
Case vbKeyPageUp
     If SeekAction >= 4 Then
        MoveRec 1
     End If
Case vbKeyPageDown
     If SeekAction >= 4 Then
        MoveRec 2
     End If
Case vbKeyEnd
     If SeekAction >= 4 Then
        MoveRec 3
     End If
Case vbKeyF9
     If SeekAction >= 0 And SeekAction <= 3 Then
         YNEdit 4
     End If
Case vbKeyF10
     If SeekAction >= 0 And SeekAction <= 3 Then
        YNEdit 5
     End If
Case vbKeyEscape
     If SeekAction >= 4 Then
        Unload Me
     End If
Case Else
End Select
   If KeyCode <> vbKeyF2 And KeyCode <> vbKeyF3 And KeyCode <> vbKeyF4 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyEscape Then
      If SeekAction > 3 Then
         If m_bInsert Then
             TBar1.Buttons(1).Enabled = True
         Else
             TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
             TBar1.Buttons(2).Enabled = True
         Else
             TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
             TBar1.Buttons(3).Enabled = True
         Else
             TBar1.Buttons(3).Enabled = False
         End If
      End If
   End If

End Sub

Private Sub Form_Load()
       m_bInsert = IsUserHasRightOfFunction("frm090617", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm090617", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm090617", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm090617", strFind, False)
    MoveFormToCenter Me
    SeekAction = 4
    ProcessUp
    ProcessDown
    TxtLock 3
    TxtSitu True
    ReDim DELMenu(0) As String
    ReDim DELTemp(0) As String
       If m_bInsert Then
       TBar1.Buttons(1).Enabled = True
   Else
       TBar1.Buttons(1).Enabled = False
   End If
   If m_bUpdate Then
       TBar1.Buttons(2).Enabled = True
   Else
       TBar1.Buttons(2).Enabled = False
   End If
   If m_bDelete Then
       TBar1.Buttons(3).Enabled = True
   Else
       TBar1.Buttons(3).Enabled = False
   End If

End Sub

Sub ProcessUp()
'取的上半部資料
strSql = "select DISTINCT sb02-1911,sb03,(sb02-1911)||sb03 as d from staffbonus order by 1,2 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        GetDataUp
    Else
        For i = 0 To 1
            txt1(i) = ""
        Next i
    End If
End With
End Sub
        
Private Sub GetDataUp()         '取得上半部資料
If adoRecordset.RecordCount = 0 Then
    For i = 0 To 1
        txt1(i) = ""
    Next i
Else
    For i = 0 To 1
        txt1(i) = CheckStr(adoRecordset.Fields(i))
    Next i
End If
End Sub

Sub ProcessDown()
strSql = "select sb01 as 員工編號,st02 as 姓名,sb04 as 獎金,sb05 as 額外獎勵,sb04+sb05 as 小計 from staffbonus,staff where sb01=ST01(+) AND sb02=" & Val(txt1(0)) + 1911 & " and sb03=" & Val(txt1(1)) & " order by 1,2 "
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd1.Recordset = adoRecordset1
        grd1.row = 1
        TextOk = True
        GetDataDown
    Else
        lbl1(0).Caption = ""
        lbl1(1).Caption = ""
        txt1(2).Text = ""
        txt1(3).Text = ""
        grd1.Clear
        grd1.Rows = 2
        GetDataDown
    End If
End With
CheckOC2
End Sub

Private Sub GetDataDown()         '取得下半部資料
Grd1_Click
'txt1(2).SetFocus
End Sub

Private Sub TxtLock(ByVal Lt As Integer)
 Dim txt As TextBox, i As Integer
   Select Case Lt
      Case 0
         TxtLock 1
         For Each txt In frm090617.txt1
            txt.Locked = True
         Next
      Case 1
         For Each txt In frm090617.txt1
            txt.Locked = False
            txt.Enabled = True
         Next
      Case 2
         For Each txt In frm090617.txt1
            txt.Text = ""
         Next
         lbl1(0).Caption = ""
         lbl1(1).Caption = ""
         txt1(2).Text = ""
         txt1(3).Text = ""
         grd1.Clear
         grd1.Rows = 2
         SetGrd1
      Case 3
         For i = 0 To 1
            If SeekAction = 0 Or SeekAction = 1 Then
                txt1(i).Enabled = False
            Else
                txt1(i).Locked = True
            End If
         Next i
      Case 4
         For i = 2 To 3
            txt1(i).Enabled = False
         Next i
   End Select
End Sub

Private Sub TxtSitu(ByVal TF As Boolean)
 Dim i As Integer, txt As TextBox
   If TF = True Then
      TxtLock 0
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = True
         TBar1.Buttons(i + 5).Enabled = True
      Next
      TBar1.Buttons(11).Enabled = False
      TBar1.Buttons(12).Enabled = False
      TBar1.Buttons(14).Enabled = True
   Else
      TxtLock 1
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = False
         TBar1.Buttons(i + 5).Enabled = False
      Next
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
      TBar1.Buttons(14).Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090617 = Nothing
End Sub

Private Sub Grd1_Click()
With grd1
    .Visible = False
    For i = 0 To .Rows - 1
        .col = 0
        .row = i
        If .CellBackColor = &HFFC0C0 Then
            For k = 0 To .Cols - 1
                .col = k
                .CellBackColor = QBColor(15)
            Next k
            Exit For
        End If
    Next i
    .col = 0
    If TextOk = True Then
        .row = 0
        TextOk = False
    Else
        .row = .MouseRow
    End If
    If .row = 0 Then
        .row = 1
    End If
    .col = 0
    lbl1(0).Caption = .Text
    .col = 1
    lbl1(1).Caption = .Text
    .col = 2
    txt1(2).Text = .Text
    .col = 3
    txt1(3).Text = .Text
    For j = 0 To .Cols - 1
        .col = j
        If Len(.Text) <> 0 Then
            For i = 0 To .Cols - 1
                .col = i
                .CellBackColor = &HFFC0C0
            Next i
            Exit For
        End If
    Next j
    SetGrd1
    .Visible = True
End With
End Sub

Private Sub YNEdit(ByVal Strindex As Integer)

'911107 nick
On Error GoTo CheckingErr

Select Case Strindex
Case 0  'ADD
     TxtSitu False
     SeekAction = 0
     TxtLock 2
     TxtLock 4
     txt1(0).SetFocus
Case 1  'EDIT
     TxtSitu False
     SeekAction = 1
     TxtLock 3
Case 2  'DEL
     TxtSitu False
     TxtLock 4
     SeekAction = 2
     SeekRec = adoRecordset.Bookmark
     If MsgBox("是否要刪除此筆資料??", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
        YNEdit 4
     End If
     TxtSitu True
     TxtLock 0
     txt1(0).SetFocus
     txt1_GotFocus (0)
     SeekAction = 4
Case 3  'FIND
     TxtLock 2
     TxtSitu False
     TxtLock 4
     SeekAction = 3
     SeekRec = adoRecordset.Bookmark
     txt1(0).SetFocus
     Exit Sub
Case 4  'ENTER
     Select Case SeekAction
     Case 0
          grd1.row = 1
          grd1.col = 1
          If Len(txt1(0)) <> 0 And Len(txt1(1)) <> 0 And Len(txt1(2)) <> 0 And Len(grd1.Text) <> 0 Then
               'Add By Cheng 2002/05/23
               '重新檢查欄位有效性
               If TxtValidate = False Then Exit Sub
                
                '911107 nickchen
                cnnConnection.BeginTrans
                
                For i = 1 To grd1.Rows - 1
                    grd1.row = i
                    strSql = "INSERT INTO staffbonus (sb01,sb02,sb03,sb04,sb05) VALUES ('"
                    grd1.col = 0
                    strSql = strSql & Trim(grd1.Text) & "'," & Val(txt1(0)) + 1911 & "," & Val(txt1(1)) & ","
                    grd1.col = 2
                    strSql = strSql & Val(grd1.Text) & ","
                    grd1.col = 3
                    strSql = strSql & Val(grd1.Text) & ")"
                    cnnConnection.Execute strSql
                Next i
                
                '911107 nickchen
                cnnConnection.CommitTrans
                
          Else
              s = MsgBox("沒有資料可存入資料庫!!", , "USER 輸入錯誤")
          End If
     Case 1
          grd1.row = 1
          grd1.col = 1
          If Len(txt1(0)) <> 0 And Len(txt1(1)) <> 0 And Len(txt1(2)) <> 0 Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
                
                '911107 nickchen
                cnnConnection.BeginTrans
                
                For i = 1 To grd1.Rows - 1
                    grd1.row = i
                    strSql = "UPDATE staffbonus SET sb04="
                    grd1.col = 2
                    strSql = strSql & Val(grd1.Text) & ",sb05="
                    grd1.col = 3
                    strSql = strSql & Val(grd1.Text) & " "
                    grd1.col = 0
                    strSql = strSql & " WHERE sb01='" & Trim(grd1.Text) & "' AND sb02=" & Val(txt1(0)) + 1911 & " AND sb03=" & Val(txt1(1))
                    cnnConnection.Execute strSql
                Next i
                
                '911107 nickchen
                cnnConnection.CommitTrans
                
          End If
     Case 2
          SeekRec = adoRecordset.Bookmark
          
          '911107 nickchen
          cnnConnection.BeginTrans
          
          strSql = "DELETE FROM staffbonus WHERE sb02=" & Val(txt1(0)) + 1911 & " AND sb03=" & Val(txt1(1))
          cnnConnection.Execute strSql
          
          '911107 nickchen
          cnnConnection.CommitTrans
          
          TxtSitu True
          ProcessUp
          If SeekRec > adoRecordset.RecordCount Then
             If adoRecordset.EOF = True Then
             Else
                 adoRecordset.MoveFirst
             End If
          Else
             adoRecordset.Bookmark = SeekRec
          End If
          GetDataUp
          ProcessDown
          SeekAction = 4
          ReDim DELMenu(0) As String
          ReDim DELTemp(0) As String
          TxtLock 0
          txt1(0).SetFocus
          txt1_GotFocus (0)
          Exit Sub
     Case 3
          SeekRec = adoRecordset.Bookmark
          adoRecordset.Find "D='" & txt1(0) & txt1(1) & "'", 0, adSearchForward, 1
          If adoRecordset.EOF Then
              s = MsgBox("沒有符合資料!!", , "錯誤")
              adoRecordset.Bookmark = SeekRec
          End If
          TxtSitu True
          GetDataUp
          ProcessDown
          SeekAction = 4
          TxtLock 0
          txt1(0).SetFocus
          txt1_GotFocus (0)
          Exit Sub
     Case Else
     End Select
     TxtSitu True
     ProcessUp
     ProcessDown
     SeekAction = 4
     ReDim DELMenu(0) As String
     ReDim DELTemp(0) As String
     TxtLock 0
     txt1(0).SetFocus
     txt1_GotFocus (0)
Case 5  'CHANCL
     Select Case SeekAction
     Case 0
          If Len(txt1(0)) <> 0 And Len(txt1(1)) <> 0 And Len(txt1(2)) <> 0 Then
              If MsgBox("你尚未存檔, 確定離開嗎??", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
                  Exit Sub
              End If
          End If
     Case 1
          If Len(txt1(0)) <> 0 And Len(txt1(1)) <> 0 And Len(txt1(2)) <> 0 Then
              If MsgBox("你尚未存檔, 確定離開嗎??", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
                  Exit Sub
              End If
          End If
          TxtLock 1
     Case 2
          adoRecordset.Bookmark = SeekRec
          GetDataUp
          TxtSitu True
          ProcessDown
          SeekAction = 4
          ReDim DELMenu(0) As String
          ReDim DELTemp(0) As String
          TxtLock 0
          txt1(0).SetFocus
          txt1_GotFocus (0)
          Exit Sub
     Case 3
          adoRecordset.Bookmark = SeekRec
          GetDataUp
          TxtSitu True
          ProcessDown
          SeekAction = 4
          ReDim DELMenu(0) As String
          ReDim DELTemp(0) As String
          TxtLock 0
          txt1(0).SetFocus
          txt1_GotFocus (0)
          Exit Sub
     Case Else
     End Select
     TxtSitu True
     TxtLock 2
     ProcessUp
     ProcessDown
     SeekAction = 4
     ReDim DELMenu(0) As String
     ReDim DELTemp(0) As String
     TxtLock 1
     TxtLock 0
     txt1(0).SetFocus
     txt1_GotFocus (0)
Case Else
End Select
 '911107 nick transation
     Exit Sub
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
End Sub

Private Sub MoveRec(ByVal Strindex As Integer)
With adoRecordset
    If .EOF = True And .EOF = True Then Exit Sub
    Select Case Strindex
    Case 0
         .MoveFirst
    Case 1
         .MovePrevious
         If .BOF Then
            DataErrorMessage (6)
            .MoveFirst
         End If
    Case 2
         .MoveNext
         If .EOF Then
            DataErrorMessage (7)
            .MoveLast
         End If
    Case 3
         .MoveLast
    Case Else
    End Select
    GetDataUp
    ProcessDown
End With
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
         YNEdit 0
      Case 2
         If CheckRec Then
            YNEdit 1
         End If
      Case 3
         If CheckRec Then
            YNEdit 2
         End If
      Case 4
         If CheckRec Then
            YNEdit 3
         End If
      Case 6
         MoveRec 0
      Case 7
         MoveRec 1
      Case 8
         MoveRec 2
      Case 9
         MoveRec 3
      Case 11
         YNEdit 4
      Case 12
         YNEdit 5
      Case 14
         Unload Me
      End Select
         If Button.Index <> 14 And Button.Index <> 1 And Button.Index <> 2 And Button.Index <> 3 And Button.Index <> 4 Then
      If m_bInsert Then
          TBar1.Buttons(1).Enabled = True
      Else
          TBar1.Buttons(1).Enabled = False
      End If
      If m_bUpdate Then
          TBar1.Buttons(2).Enabled = True
      Else
          TBar1.Buttons(2).Enabled = False
      End If
      If m_bDelete Then
          TBar1.Buttons(3).Enabled = True
      Else
          TBar1.Buttons(3).Enabled = False
      End If
   End If

End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub SetGrd1()
With grd1
    .Visible = False
    .Cols = 5
    .row = 0
    .col = 0:   .Text = "員工編號"
    .ColWidth(0) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "姓名"
    .ColWidth(1) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "獎金"
    .ColWidth(2) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "額外獎勵"
    .ColWidth(3) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "小計"
    .ColWidth(4) = 800
    .CellAlignment = flexAlignCenterCenter
    .Visible = True
End With
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
'Add By Cheng 2002/05/24
m_blnCancel = False

If txt1(Index).Locked = True Then
    Exit Sub
End If
Select Case Index
Case 0
     If IsDate(ChangeTStringToTDateString(txt1(0) & "0101")) = False Then
        s = MsgBox("年輸入錯誤", , "USER 輸入錯誤")
        txt1(0).SetFocus
        txt1(0).SelStart = 0
        txt1(0).SelLength = Len(txt1(0))
        'Add By Cheng 2002/05/24
        m_blnCancel = True
        Exit Sub
     End If
Case 1
     If InStr(1, "1234", txt1(1)) = 0 Then
        s = MsgBox("季輸入錯誤", , "USER 輸入錯誤")
        txt1(1).SetFocus
        txt1(1).SelStart = 0
        txt1(1).SelLength = Len(txt1(1))
        'Add By Cheng 2002/05/24
        m_blnCancel = True
        Exit Sub
     End If
     If SeekAction = 0 Then
        For i = 0 To 1
            If Len(txt1(i)) = 0 Then
                s = MsgBox("年與季不可空白!!", , "USER 輸入錯誤")
                If Len(txt1(1)) = 0 Then txt1(1).SetFocus
                If Len(txt1(0)) = 0 Then txt1(0).SetFocus
                'Add By Cheng 2002/05/24
                m_blnCancel = True
                Exit Sub
            End If
        Next i
        '92.04.03 nick add left join
        'strSQL = "SELECT S1.ST01 FROM STAFF S1,STAFF S2 WHERE S1.ST04='1' AND SUBSTR(S1.ST03,1,2)=SUBSTR(S2.ST03,1,2) AND S2.ST01='" & strUserNum & "' "
        'Modified by Morgan 2018/5/23
        'strSql = "SELECT S1.ST01 FROM STAFF S1,STAFF S2 WHERE S1.ST04='1' AND SUBSTR(S2.ST03,1,2)=SUBSTR(S1.ST03,1,2)(+) AND S2.ST01='" & strUserNum & "' "
        strSql = "SELECT S1.ST01 FROM STAFF S1,STAFF S2 WHERE S1.ST04='1' AND SUBSTR(S2.ST03,1,2)=SUBSTR(S1.ST03(+),1,2) AND S2.ST01='" & strUserNum & "' "
        CheckOC2
        With adoRecordset1
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount <> 0 And .RecordCount > 0 Then
                SeekTemp = CheckStr(.Fields(0))
                strSql = "SELECT * FROM staffbonus WHERE sb01='" & SeekTemp & "' AND sb02=" & Val(txt1(0)) + 1911 & " AND sb03=" & Val(txt1(1))
                CheckOC2
                .CursorLocation = adUseClient
                .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If .RecordCount <> 0 And .RecordCount > 0 Then
                    s = MsgBox("此考核年度及考核季別與部門已存在!!", , "USER 輸入錯誤")
                    CheckOC
                    txt1(0).SetFocus
                    txt1(0).SelStart = 0
                    txt1(0).SelLength = Len(txt1(0))
                    'Add By Cheng 2002/05/24
                    m_blnCancel = True
                  Exit Sub
                End If
            End If
        End With
        CheckOC2
        '92.04.03 nick add left join
        'strSQL = "select S1.ST01,S1.st02,0,0,0 from staff S1,STAFF S2 where SUBSTR(S1.st03,1,2)=SUBSTR(S2.ST03,1,2) AND S1.ST04='1' AND S2.ST01='" & strUserNum & "' order by 1,2 "
        'Modified by Morgan 2018/5/23
        'strSql = "select S1.ST01,S1.st02,0,0,0 from staff S1,STAFF S2 where SUBSTR(S2.st03,1,2)=SUBSTR(S1.ST03,1,2)(+) AND S1.ST04='1' AND S2.ST01='" & strUserNum & "' order by 1,2 "
        strSql = "select S1.ST01,S1.st02,0,0,0 from staff S1,STAFF S2 where SUBSTR(S2.st03,1,2)=SUBSTR(S1.ST03(+),1,2) AND S1.ST04='1' AND S2.ST01='" & strUserNum & "' order by 1,2 "
        CheckOC2
        With adoRecordset1
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount <> 0 And .RecordCount > 0 Then
                Set grd1.Recordset = adoRecordset1
                grd1.row = 1
                TextOk = True
                GetDataDown
            End If
        End With
        CheckOC2
        TxtLock 1
        TxtLock 3
        txt1(2).SetFocus
     End If
Case 2, 3
     If IsNumeric(txt1(Index)) = False And Len(txt1(Index)) <> 0 Then
        s = MsgBox("輸入錯誤, 請輸入數字", , "USER 輸入錯誤")
        txt1(Index).SetFocus
        txt1(Index).SelStart = 0
        txt1(Index).SelLength = Len(txt1(Index))
        'Add By Cheng 2002/05/24
        m_blnCancel = True
        Exit Sub
     End If
Case Else
End Select
End Sub

Function CheckRec() As Boolean
If adoRecordset.RecordCount <> 0 Then CheckRec = True Else CheckRec = False
End Function

'Add By Cheng 2002/05/24
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Me.txt1
   If objTxt.Enabled = True Then
      Cancel = False
      txt1_LostFocus objTxt.Index
      If m_blnCancel = True Then
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function

