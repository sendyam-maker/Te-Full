VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090631 
   BorderStyle     =   1  '單線固定
   Caption         =   "工程師每月目標基數設定"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9285
   Begin VB.Frame Frame1 
      Height          =   4356
      Left            =   0
      TabIndex        =   2
      Top             =   1050
      Width           =   9252
      Begin VB.CommandButton Command1 
         Caption         =   "重算點數"
         Height          =   400
         Left            =   8220
         TabIndex        =   14
         Top             =   150
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton cmd 
         Caption         =   "修改(E)"
         Height          =   400
         Left            =   7380
         TabIndex        =   5
         Top             =   144
         Width           =   810
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  '靠右對齊
         Height          =   264
         Index           =   1
         Left            =   690
         MaxLength       =   7
         TabIndex        =   4
         Top             =   480
         Width           =   1100
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   3525
         Left            =   75
         TabIndex        =   3
         Top             =   795
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   6218
         _Version        =   393216
         Rows            =   3
         Cols            =   1
         FixedRows       =   2
         FixedCols       =   0
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         HighLight       =   0
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "基數點數換算:  機械 x18, 電子 x22, 生化 x27 "
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3030
         TabIndex        =   13
         Top             =   540
         Width           =   3465
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   2
         Left            =   5265
         TabIndex        =   12
         Top             =   210
         Width           =   1875
         VariousPropertyBits=   27
         Caption         =   "組別："
         Size            =   "3307;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   210
         Width           =   2205
         VariousPropertyBits=   27
         Caption         =   "員工編號："
         Size            =   "3889;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl2 
         Height          =   255
         Index           =   1
         Left            =   2670
         TabIndex        =   9
         Top             =   210
         Width           =   2265
         VariousPropertyBits=   27
         Caption         =   "姓名："
         Size            =   "3995;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "基數："
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   510
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "複製資料(&C)"
      Height          =   400
      Left            =   7110
      TabIndex        =   1
      Top             =   660
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1110
      MaxLength       =   5
      TabIndex        =   0
      Top             =   720
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8136
      Top             =   840
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
            Picture         =   "frm090631.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090631.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090631.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090631.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090631.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090631.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090631.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090631.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090631.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090631.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090631.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ex:9605  輸入完畢後按下 Tab ，將會開始計算下面資料"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   2280
      TabIndex        =   11
      Top             =   750
      Width           =   4290
   End
   Begin VB.Label Label1 
      Caption         =   "目標年月："
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   8
      Top             =   750
      Width           =   960
   End
End
Attribute VB_Name = "frm090631"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/17 改成Form2.0 (grd1,lbl2)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
'Modified by Morgan 2019/4/2 基數改可輸入小數兩位--游經理
'create by nickc 2007/04/25  copy from frm090615
Option Explicit
Dim i As Integer, j As Integer, k As Integer, s As Integer, TextOk As Boolean, SeekAction As Integer, SeekRec As Variant
Dim StrSQL6 As String, strTemp1 As Variant, SeekTemp As String, DELMenu() As String, DELTemp() As String, SeekBmk1 As Variant, SeekBmk2 As Variant, SeekBmk3 As Variant
Dim strTemp(0 To 9) As String, PLeft(0 To 9) As Integer, Page As Integer, iPrint As Integer, seekbmk, BolDbOk As Boolean
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_blnCancel As Boolean
Dim MyRs As New ADODB.Recordset
Dim MyRs1 As New ADODB.Recordset
Public TagM As String
Public SrcM As String
Public IsCopy As Boolean
Dim MyRs2 As New ADODB.Recordset
Dim m_bol108Rule As Boolean 'Added by Morgan 2019/3/14

Private Sub cmd_Click()
If SeekAction <> 1 And SeekAction <> 0 Then
    Exit Sub
End If
m_blnCancel = False
txt1_LostFocus 1
If m_blnCancel = True Then
    Exit Sub
End If
With grd1
     For i = 1 To .Rows - 1
        .col = 0
        .row = i
        If .CellBackColor = &HFFC0C0 Then
            .col = 5
            .Text = Format(txt1(1), "#####0.00")
            Cal1 i
            Exit For
        End If
     Next i
End With
End Sub
Sub Cal1(ii As Integer)
     
   With grd1
   'Added by Morgan 2019/3/14 +108考核
   If m_bol108Rule Then
      Dim stGpID As String, iRate As Integer
      Dim stPoint As String
      
      stGpID = GetGridValue("ST70", ii)
      If stGpID = "1" Then '機械
         iRate = 18
      ElseIf stGpID = "2" Then '電子
         iRate = 22
      ElseIf stGpID = "3" Then '生化
         iRate = 27
      Else '未分組
         iRate = 18
      End If
      
      'Modified by Morgan 2019/8/7 目標還是要分 P,CFP(達成情形統計要用) --王副總
      ''柄佑:目標件數點數不再區分P,CFP
      '.TextMatrix(ii, 6) = .TextMatrix(ii, 5)
      '.TextMatrix(ii, 7) = Format(Val(.TextMatrix(ii, 5)) * iRate, "0.0")
      '.TextMatrix(ii, 8) = 0
      '.TextMatrix(ii, 9) = 0
      '目標總點數
      stPoint = Format(Val(.TextMatrix(ii, 5)) * iRate, "0.0")
      If Val(.TextMatrix(ii, 3)) + Val(.TextMatrix(ii, 4)) <> 0 Then
         'CFP 浮動目標
         .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 5)) * Val(.TextMatrix(ii, 4)) / (Val(.TextMatrix(ii, 3)) + (Val(.TextMatrix(ii, 4)) * 2)), "0.00")
      Else
         'CFP 浮動目標
         .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 5)) * 1 / 4, "0.00")
      End If
      'CFP 目標點數
      .TextMatrix(ii, 9) = Format(Val(.TextMatrix(ii, 8)) * iRate * 2, "0.0")
      
      'P 浮動目標(用減的才不會有誤差)
      .TextMatrix(ii, 6) = Format(Val(.TextMatrix(ii, 5)) - Val(.TextMatrix(ii, 8)) * 2, "0.00")
      'P 目標點數(用減的才不會有誤差)
      .TextMatrix(ii, 7) = Format(stPoint - Val(.TextMatrix(ii, 9)), "0.0")
      
      'end 2019/8/7
      
   Else
   'end 2019/3/14
   
      If Val(.TextMatrix(ii, 3)) + Val(.TextMatrix(ii, 4)) <> 0 Then
         'P 浮動目標
         .TextMatrix(ii, 6) = Format(Val(.TextMatrix(ii, 5)) * Val(.TextMatrix(ii, 3)) / (Val(.TextMatrix(ii, 3)) + (Val(.TextMatrix(ii, 4)) * 2)), "######0.0")
         'P 目標點數
         .TextMatrix(ii, 7) = Format(Val(.TextMatrix(ii, 6)) * 15, "######0.0")
         'CFP 浮動目標
         .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 5)) * Val(.TextMatrix(ii, 4)) / (Val(.TextMatrix(ii, 3)) + (Val(.TextMatrix(ii, 4)) * 2)), "######0.0")
         'CFP 目標點數
         .TextMatrix(ii, 9) = Format(Val(.TextMatrix(ii, 8)) * 30, "######0.0")
         
      'Add by Morgan 2010/10/5 上月沒有完稿件數時,P和CFP各一半
      Else
         'P 浮動目標
         .TextMatrix(ii, 6) = Format(Val(.TextMatrix(ii, 5)) * 1 / 2, "######0.0")
         'P 目標點數
         .TextMatrix(ii, 7) = Format(Val(.TextMatrix(ii, 6)) * 15, "######0.0")
         'CFP 浮動目標
         .TextMatrix(ii, 8) = Format(Val(.TextMatrix(ii, 5)) * 1 / 4, "######0.0")
         'CFP 目標點數
         .TextMatrix(ii, 9) = Format(Val(.TextMatrix(ii, 8)) * 30, "######0.0")
         
      End If
   End If
   End With
End Sub

Private Sub cmdOK_Click()
     IsCopy = False
     TagM = ""
     SrcM = ""
     frm090631_1.Show vbModal
     If IsCopy = True And TagM <> "" And SrcM <> "" Then
           If SeekAction >= 4 Then
               Screen.MousePointer = vbHourglass
              '新增
               YNEdit 0
               txt1(0).Text = TagM
               txt1_LostFocus 0
               '再一一將基數資料讀出
               With grd1
                  For i = 1 To .Rows - 1
                     .row = i
                      strSql = "select * from engradix where er01=" & Mid(ChangeTStringToWString(SrcM & "01"), 1, 6) & " and er02='" & Trim(.TextMatrix(i, 1)) & "'"
                      If MyRs2.State = 1 Then MyRs2.Close
                      MyRs2.CursorLocation = adUseClient
                      MyRs2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                      If MyRs2.RecordCount <> 0 Then
                         If .TextMatrix(i, 12) = "" Then 'Add by Morgan 2010/9/14 未預設基礎值才複製
                            .TextMatrix(i, 5) = Format(CheckStr(MyRs2.Fields("er03")), "######0.00")
                            Cal1 i
                         'Add by Morgan 2011/9/1 新人基數改控制複製值高於標準值時也要更新
                         ElseIf Val(Format(.TextMatrix(i, 5))) < Val("" & MyRs2.Fields("er03")) Then
                           .TextMatrix(i, 5) = Format(CheckStr(MyRs2.Fields("er03")), "######0.00")
                           Cal1 i
                         End If
                      End If
                   Next i
              End With
             '存檔
             YNEdit 4
             '查詢
             YNEdit 3
             txt1(0).Text = TagM
             YNEdit 4
             IsCopy = False
             Screen.MousePointer = vbDefault
             '告知完成
             MsgBox "資料複製完成！", vbExclamation, "成功！"
            End If
     End If
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

Private Sub Command1_Click()
   If SeekAction <> 1 And SeekAction <> 0 Then
       Exit Sub
   End If

   Dim ii As Integer
   With grd1
   .Visible = False
   For ii = 1 To .Rows - 1
      Cal1 ii
   Next
   .Visible = True
   End With
   MsgBox "點數已重算！"
End Sub

Private Sub Form_Initialize()
Set MyRs = New ADODB.Recordset
Set MyRs1 = New ADODB.Recordset
Set MyRs2 = New ADODB.Recordset
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
Case vbKeyF9, vbKeyReturn
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
   '權限部份與目標設定一樣
   m_bInsert = IsUserHasRightOfFunction("frm090615", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm090615", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm090615", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm090615", strFind, False)
    MoveFormToCenter Me
    SeekAction = 4
    grd1.Cols = 12
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
    IsCopy = False
    
   'Added by Morgan 2019/8/7
   If Pub_StrUserSt03 = "M51" Then
      Command1.Visible = True
   Else
      Command1.Visible = False
   End If
   'end 2019/8/7
End Sub

Sub ProcessUp()
StrSQL6 = " "
'取的上半部資料
strSql = "select DISTINCT er01-191100 as D from engradix order by 1"
If MyRs.State = 1 Then MyRs.Close
With MyRs
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        'Modified by Morgan 2019/3/27 改預設在最後一筆--柄佑
        '.MoveFirst
        .MoveLast
        BolDbOk = True
        GetDataUp
    Else
        BolDbOk = False
        For i = 0 To 1
            txt1(i) = ""
        Next i
        lbl2(0).Caption = "員工編號："
        lbl2(1).Caption = "姓名："
    End If
End With
End Sub
        
Private Sub GetDataUp()         '取得上半部資料
If MyRs.RecordCount = 0 Then
    For i = 0 To 1
        txt1(i) = ""
    Next i
    lbl2(0).Caption = "員工編號："
    lbl2(1).Caption = "姓名："
Else
    txt1(0) = CheckStr(MyRs.Fields(0))
End If
End Sub

Sub ProcessDown()
Dim strLstYM As String, strCurYM As String

Screen.MousePointer = vbHourglass
grd1.MousePointer = flexHourglass
grd1.Clear
grd1.Rows = 2
SetGrd1
If Trim(txt1(0)) = "" Then grd1.MousePointer = flexDefault: Screen.MousePointer = vbDefault: Exit Sub

strCurYM = Val(txt1(0)) + 191100
strLstYM = Mid(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(ChangeTStringToWString(txt1(0) & "01")))), 1, 6)

'Modify by Morgan 2010/9/7 +不必限制上個月有案子,排除王副總及虛建的編號
'Modify by Morgan 2010/9/14 +er02,st13,Tag '計算到職未滿1年的基數用
'Modify by Morgan 2010/10/5 目標抓資料庫設定,只有修改基數時才重算
'strSql = " select decode(st06,'1','北所','2','中所','3','南所','4','高所','5','其他',''),ST01,st02,ltrim(to_char(PSCount,'999990.0')),ltrim(to_char(CFPSCount,'999990.0')),ltrim(to_char(nvl(er03,0),'999990.0')),ltrim(to_char(0,'999990.0')),ltrim(to_char(0,'999990.00')),ltrim(to_char(0,'999990.0')),ltrim(to_char(0,'999990.0')),er02,st13,'' Tag from staff,engradix , "
'Modified by Morgan 2019/3/14 +st70,GP組別
strSql = " select decode(st06,'1','北所','2','中所','3','南所','4','高所','5','其他',''),ST01,st02,ltrim(to_char(PSCount,'999990.00')),ltrim(to_char(CFPSCount,'999990.00')),ltrim(to_char(nvl(er03,0),'999990.00')),ltrim(to_char(p1.pe05,'999990.00')),ltrim(to_char(p1.pe06,'999990.0')),ltrim(to_char(p2.pe05,'999990.00')),ltrim(to_char(p2.pe06,'999990.0')),er02,st13,'' Tag,st70,CST70(st70,st03) GP from staff,engradix , "
'end 2010/10/5
strSql = strSql & " (select cp14,sum(PCount) as PSCount,sum(CFPCount) as CFPSCount from ( "
strSql = strSql & " select cp14,sum(decode(cp01,'P',cp97 * cp98 * decode(cp112,'Y',cp111,1),0)) PCount,sum(decode(cp01,'CFP',cp97 * cp98 * decode(cp112,'Y',cp111,1),0)) CFPCount "
strSql = strSql & " from caseprogress ,engineerprogress,staff  where ep02=cp09(+) and cp14=st01(+)  and st03 in ('P10','P11') "
strSql = strSql & " and ep09>=" & strLstYM & "01 and ep09<=" & strLstYM & "31 group by cp14 "
'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數
'strSql = strSql & " union select sh02 CP14,sum(Round(Decode(SH06, 'CFP', 0, Nvl(SH05,0)/4) ,2)) PCount,sum(Round(Decode(SH06, 'CFP', Nvl(SH05,0)/8, 0) ,2)) as CFPCount from supporthour where sh01>=" & strLstYM & "01 and sh01<=" & strLstYM & "31 group by sh02 "
'Modified by Morgan 2019/4/9 108考核支援時數轉換要除組別參數
'strSql = strSql & " union select sh02 CP14,sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0,1) ,2)) PCount,sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',1,0) ,2)) as CFPCount from supporthour,staff where sh01>=" & strLstYM & "01 and sh01<=" & strLstYM & "31 and st01(+)=sh02 group by sh02 "
strSql = strSql & " union select sh02 CP14,sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',0,1) / GetDivNum(st70,sh01) ,2)) PCount,sum(Round(" & Sh2EPtCode & " * decode(SH06,'CFP',1,0) / GetDivNum(st70,sh01) ,2)) as CFPCount from supporthour,staff where sh01>=" & strLstYM & "01 and sh01<=" & strLstYM & "31 and st01(+)=sh02 group by sh02 "
'end 2014/3/20
'Modify by Morgan 2010/10/5 目標抓資料庫設定,只有修改基數時才重算
'strSql = strSql & "  ) group by cp14) A  where ST01=er02(+) AND " & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6) & "=er01(+) and ST04='1' "
strSql = strSql & "  ) group by cp14) A,performance P1,performance P2 where p1.pe01(+)=st01 and p2.pe01(+)=st01 and p1.pe02(+)='P' and p2.pe02(+)='CFP' and p1.pe03(+)=" & strCurYM & " and p2.pe03(+)=" & strCurYM & " and ST01=er02(+) AND " & strCurYM & "=er01(+) and ST04='1' "
'end 2010/10/5
strSql = strSql & " AND st03 in('P10','P11') and st01=A.cp14(+) and st01>'71011' and st01<'P' order by st06,2,3 "
If MyRs1.State = 1 Then MyRs1.Close
With MyRs1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set grd1.Recordset = MyRs1
        grd1.row = 1
        TextOk = True
        GetDataDown
    Else
        grd1.Clear
        grd1.Rows = 2
        GetDataDown
    End If
End With
If MyRs1.State = 1 Then MyRs1.Close
SetGrd1
With grd1
     .Visible = False
     For i = 1 To .Rows - 1
         .row = i
         'modify by Morgan 2010/9/14
         'Cal1 i
         '新增
         If SeekAction = 0 Then
            If bolNewPromoterRule Then
               '13個月以下
               If Val(.TextMatrix(i, 11)) > Val(CompDate(1, -13, txt1(0) & "01")) Then
                  'Added by Morgan 2019/5/14 +108考核
                  If m_bol108Rule Then
                     Dim stGpID As String, iTarget As Double
                     stGpID = GetGridValue("ST70", i)
                     '目標點數
                     If stGpID = "1" Then '機械
                        iTarget = 15.3
                     ElseIf stGpID = "2" Then '電子
                        iTarget = 12.9
                     ElseIf stGpID = "3" Then '生化
                        iTarget = 10.9
                     Else '未分組
                        iTarget = 15.3
                     End If
                     
                     '13~
                     If Val(.TextMatrix(i, 11)) <= Val(CompDate(1, -12, txt1(0) & "01")) Then
                        .TextMatrix(i, 5) = iTarget
                        .TextMatrix(i, 12) = "Y"
                     '10~
                     ElseIf Val(.TextMatrix(i, 11)) <= Val(CompDate(1, -9, txt1(0) & "01")) Then
                        .TextMatrix(i, 5) = Round(iTarget * 0.9, 2) '90%
                        .TextMatrix(i, 12) = "Y"
                     '7~
                     ElseIf Val(.TextMatrix(i, 11)) <= Val(CompDate(1, -6, txt1(0) & "01")) Then
                        .TextMatrix(i, 5) = Round(iTarget * 0.75, 2) '75%
                        .TextMatrix(i, 12) = "Y"
                     '4~
                     ElseIf Val(.TextMatrix(i, 11)) <= Val(CompDate(1, -3, txt1(0) & "01")) Then
                        .TextMatrix(i, 5) = Round(iTarget * 0.6, 2) '60%
                        .TextMatrix(i, 12) = "Y"
                     End If
                     
                  'Added by Morgan 2014/3/7 4月起改基礎目標
                  ElseIf Val(txt1(0)) >= 10304 Then
                  
                     '13~
                     If Val(.TextMatrix(i, 11)) <= Val(CompDate(1, -12, txt1(0) & "01")) Then
                        .TextMatrix(i, 5) = 18
                        .TextMatrix(i, 12) = "Y"
                     '10~
                     ElseIf Val(.TextMatrix(i, 11)) <= Val(CompDate(1, -9, txt1(0) & "01")) Then
                        .TextMatrix(i, 5) = 16.2 '90%
                        .TextMatrix(i, 12) = "Y"
                     '7~
                     ElseIf Val(.TextMatrix(i, 11)) <= Val(CompDate(1, -6, txt1(0) & "01")) Then
                        .TextMatrix(i, 5) = 13.5 '75%
                        .TextMatrix(i, 12) = "Y"
                     '4~
                     ElseIf Val(.TextMatrix(i, 11)) <= Val(CompDate(1, -3, txt1(0) & "01")) Then
                        .TextMatrix(i, 5) = 10.8 '60%
                        .TextMatrix(i, 12) = "Y"
                     End If
                     
                  Else
                  'end 2014/3/7
                  
                     '13~
                     If Val(.TextMatrix(i, 11)) <= Val(CompDate(1, -12, txt1(0) & "01")) Then
                        .TextMatrix(i, 5) = 21.6
                        .TextMatrix(i, 12) = "Y"
                     '10~
                     ElseIf Val(.TextMatrix(i, 11)) <= Val(CompDate(1, -9, txt1(0) & "01")) Then
                        .TextMatrix(i, 5) = 19.4 '=21.6*90%
                        .TextMatrix(i, 12) = "Y"
                     '7~
                     ElseIf Val(.TextMatrix(i, 11)) <= Val(CompDate(1, -6, txt1(0) & "01")) Then
                        .TextMatrix(i, 5) = 16.2 '=21.6*75%
                        .TextMatrix(i, 12) = "Y"
                     '4~
                     ElseIf Val(.TextMatrix(i, 11)) <= Val(CompDate(1, -3, txt1(0) & "01")) Then
                        .TextMatrix(i, 5) = 13 '=21.6*60%
                        .TextMatrix(i, 12) = "Y"
                     End If
                     
                  End If 'Added by Morgan 2014/3/7 4月起改基礎目標
               End If
            End If
            Cal1 i
         End If
         'end 2010/9/14
     Next i
     .Visible = True
End With
grd1.MousePointer = flexDefault
Screen.MousePointer = vbDefault
End Sub

Private Sub GetDataDown()         '取得下半部資料
grd1_RowColChange
End Sub

Private Sub TxtLock(ByVal Lt As Integer)
 Dim txt As TextBox, i As Integer
   Select Case Lt
      Case 0
         TxtLock 1
         For Each txt In frm090631.txt1
            txt.Locked = True
         Next
      Case 1
         For Each txt In frm090631.txt1
            txt.Locked = False
            txt.Enabled = True
         Next
      Case 2
         For Each txt In frm090631.txt1
            txt.Text = ""
            If SeekAction = 0 Then
                txt.Locked = False
                txt.Enabled = True
            Else
                txt.Locked = True
                txt.Enabled = False
            End If
         Next
         If SeekAction = 3 Then
            txt1(0).Locked = False
            txt1(0).Enabled = True
         End If
         lbl2(0).Caption = ""
         lbl2(1).Caption = ""
         grd1.Clear
         grd1.Rows = 2
         SetGrd1
      Case 3
            If SeekAction = 0 Or SeekAction = 1 Then
                txt1(1).Enabled = True
                txt1(1).Locked = False
                txt1(0).Enabled = False
                txt1(0).Locked = True
            Else
                txt1(1).Enabled = False
                txt1(1).Locked = True
                txt1(0).Enabled = False
                txt1(0).Locked = True
            End If
      Case 4
                txt1(0).Enabled = False
                txt1(0).Locked = True
                txt1(1).Enabled = False
                txt1(1).Locked = True
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
Set frm090631 = Nothing
End Sub

Private Sub YNEdit(ByVal Strindex As Integer)


On Error GoTo CheckingErr

Select Case Strindex
Case 0  'ADD
     TxtSitu False
     SeekAction = 0
     TxtLock 2
     cmdok.Enabled = False
     Me.cmd.Default = True
     txt1(0).SetFocus
Case 1  'EDIT
     TxtSitu False
     SeekAction = 1
     cmdok.Enabled = False
     Me.cmd.Default = True
     TxtLock 3
     txt1(1).SetFocus
Case 2  'DEL
     TxtSitu False
     TxtLock 4
     SeekAction = 2
     SeekRec = MyRs.Bookmark
     If MsgBox("是否要刪除此筆資料??", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
        YNEdit 4
     End If
     TxtSitu True
     TxtLock 0
     txt1(0).SetFocus
     txt1_GotFocus (0)
     SeekAction = 4
Case 3  'FIND
     SeekAction = 3
     TxtSitu False
     TxtLock 2
     cmdok.Enabled = False
     SeekRec = MyRs.Bookmark
     txt1(0).SetFocus
     Exit Sub
Case 4  'ENTER
     Select Case SeekAction
     Case 0
          grd1.row = 1
          grd1.col = 1
          If Len(txt1(0)) <> 0 Then
            '重新檢查欄位有效性
                If TxtValidate = False Then Exit Sub
                
                'add by nickc 2007/05/08
                If IsCopy = True Then GoTo StartCopy
                
                If MsgBox("確定存檔？" & vbCrLf & "將同步更新承辦人的當月目標值，之前若有目標資料將會以本設定取代！", vbYesNo + vbExclamation, "警告！") = vbYes Then
StartCopy:
                    cnnConnection.BeginTrans
                    
                    For i = 1 To grd1.Rows - 1
                        grd1.row = i
                        strSql = "INSERT INTO engradix (er01,er02,er03,er04,er05,er06) VALUES ("
                        strSql = strSql & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6) & ","
                        grd1.col = 1
                        strSql = strSql & "'" & Trim(grd1.Text) & "',"
                        grd1.col = 5
                        strSql = strSql & grd1.Text & ","
                        strSql = strSql & "'" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI')))"
                        cnnConnection.Execute strSql
                        '檢查有無P ㄉ 目標
                        strExc(0) = "SELECT * FROM performance WHERE pe01='" & grd1.TextMatrix(i, 1) & "' and pe02='P' and pe03=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6)
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                            strSql = "update PERFORMANCE set pe05=" & grd1.TextMatrix(i, 6) & ",pe06=" & grd1.TextMatrix(i, 7) & " WHERE pe01='" & grd1.TextMatrix(i, 1) & "' and pe02='P' and pe03=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6)
                            cnnConnection.Execute strSql
                        Else
                            strSql = "INSERT INTO PERFORMANCE (PE01,PE02,PE03,PE05,PE06) VALUES ('" & Trim(grd1.TextMatrix(i, 1)) & "','P'," & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6) & "," & grd1.TextMatrix(i, 6) & "," & grd1.TextMatrix(i, 7) & ") "
                            cnnConnection.Execute strSql
                        End If
                        '檢查有無CFP ㄉ 目標
                        strExc(0) = "SELECT * FROM performance WHERE pe01='" & grd1.TextMatrix(i, 1) & "' and pe02='CFP' and pe03=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6)
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                            strSql = "update PERFORMANCE set pe05=" & grd1.TextMatrix(i, 8) & ",pe06=" & grd1.TextMatrix(i, 9) & " WHERE pe01='" & grd1.TextMatrix(i, 1) & "' and pe02='CFP' and pe03=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6)
                            cnnConnection.Execute strSql
                        Else
                            strSql = "INSERT INTO PERFORMANCE (PE01,PE02,PE03,PE05,PE06) VALUES ('" & Trim(grd1.TextMatrix(i, 1)) & "','CFP'," & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6) & "," & grd1.TextMatrix(i, 8) & "," & grd1.TextMatrix(i, 9) & ") "
                            cnnConnection.Execute strSql
                        End If
                    Next i
                    cnnConnection.CommitTrans
                    'add by nickc 2007/05/08
                    If IsCopy = False Then
                        s = MsgBox("基數、目標存檔成功!!", , "成功!!")
                    End If
                Else
                    Exit Sub
                End If
          Else
              s = MsgBox("沒有資料可存入資料庫!!", , "USER 輸入錯誤")
          End If
     Case 1
          SeekRec = MyRs.Bookmark
          grd1.row = 1
          grd1.col = 1
          If Len(txt1(0)) <> 0 Then
            If TxtValidate = False Then Exit Sub
                If MsgBox("確定存檔？" & vbCrLf & "將同步更新承辦人的當月目標值，之前若有目標資料將會以本設定取代！", vbYesNo + vbExclamation, "警告！") = vbYes Then
        
                    cnnConnection.BeginTrans
                    For i = 1 To grd1.Rows - 1
                       grd1.row = i
                       grd1.col = 1
                       strExc(0) = "SELECT * FROM engradix WHERE er01=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6) & " AND er02='" & grd1.TextMatrix(i, 1) & "' "
                       intI = 1
                       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                       If intI = 1 Then
                            strSql = "UPDATE engradix SET er03="
                            grd1.col = 5
                            strSql = strSql & grd1.Text & ",er07='" & strUserNum & "',er08=to_number(to_char(sysdate,'YYYYMMDD')),er09=to_number(to_char(sysdate,'HH24MI')) "
                            grd1.col = 1
                            strSql = strSql & " WHERE er01=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6) & " AND er02='" & grd1.TextMatrix(i, 1) & "' "
                       Else
                            strSql = "INSERT INTO engradix (er01,er02,er03,er04,er05,er06) VALUES ("
                            strSql = strSql & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6) & ","
                            grd1.col = 1
                            strSql = strSql & "'" & Trim(grd1.Text) & "',"
                            grd1.col = 5
                            strSql = strSql & grd1.Text & ","
                            strSql = strSql & "'" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI')))"
                      End If
                      cnnConnection.Execute strSql
                        '檢查有無P ㄉ 目標
                        strExc(0) = "SELECT * FROM performance WHERE pe01='" & grd1.TextMatrix(i, 1) & "' and pe02='P' and pe03=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6)
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                            strSql = "update PERFORMANCE set pe05=" & grd1.TextMatrix(i, 6) & ",pe06=" & grd1.TextMatrix(i, 7) & " WHERE pe01='" & grd1.TextMatrix(i, 1) & "' and pe02='P' and pe03=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6)
                            cnnConnection.Execute strSql
                        Else
                            strSql = "INSERT INTO PERFORMANCE (PE01,PE02,PE03,PE05,PE06) VALUES ('" & Trim(grd1.TextMatrix(i, 1)) & "','P'," & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6) & "," & Val(grd1.TextMatrix(i, 6)) & "," & Val(grd1.TextMatrix(i, 7)) & ") "
                            cnnConnection.Execute strSql
                        End If
                        '檢查有無CFP ㄉ 目標
                        strExc(0) = "SELECT * FROM performance WHERE pe01='" & grd1.TextMatrix(i, 1) & "' and pe02='CFP' and pe03=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6)
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                            strSql = "update PERFORMANCE set pe05=" & grd1.TextMatrix(i, 8) & ",pe06=" & grd1.TextMatrix(i, 9) & " WHERE pe01='" & grd1.TextMatrix(i, 1) & "' and pe02='CFP' and pe03=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6)
                            cnnConnection.Execute strSql
                        Else
                            strSql = "INSERT INTO PERFORMANCE (PE01,PE02,PE03,PE05,PE06) VALUES ('" & Trim(grd1.TextMatrix(i, 1)) & "','CFP'," & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6) & "," & Val(grd1.TextMatrix(i, 8)) & "," & Val(grd1.TextMatrix(i, 9)) & ") "
                            cnnConnection.Execute strSql
                        End If
                    Next i
                    s = MsgBox("基數、目標存檔成功!!", , "成功!!")
                    cnnConnection.CommitTrans
                Else
                    Exit Sub
                End If
          End If
          TxtSitu True
          ProcessUp
          MyRs.Bookmark = SeekRec
          GetDataUp
          ProcessDown
          cmdok.Enabled = True
          SeekAction = 4
          ReDim DELMenu(0) As String
          ReDim DELTemp(0) As String
          TxtLock 1
          TxtLock 0
          txt1(0).SetFocus
          txt1_GotFocus (0)
          Exit Sub
     Case 2
          If MsgBox("確定存檔？" & vbCrLf & "將同步更新承辦人的當月目標值，之前若有目標資料將會以本設定取代！", vbYesNo + vbExclamation, "警告！") = vbYes Then
                SeekRec = MyRs.Bookmark
                cnnConnection.BeginTrans
                '先清空目標資料
                'MODIFY BY SONIA 2014/4/11 加入 pe02 in ('P','CFP') 杜燕文有T的目標
                strSql = "update performance set pe05=0,pe06=0 where (pe01,pe02,pe03)  in (select pe01,pe02,pe03 from performance,engradix where er01=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6) & " and er01=pe03(+) and er02=pe01(+) and pe02 in ('P','CFP')) "
                cnnConnection.Execute strSql
                '刪除記錄
                strSql = "DELETE FROM engradix WHERE er01=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6)
                cnnConnection.Execute strSql
                cnnConnection.CommitTrans
                TxtSitu True
                ProcessUp
                If SeekRec > MyRs.RecordCount Then
                    If MyRs.RecordCount <> 0 Then
                        MyRs.MoveFirst
                    End If
                Else
                  MyRs.Bookmark = SeekRec
                End If
                GetDataUp
                ProcessDown
                cmdok.Enabled = True
                SeekAction = 4
                ReDim DELMenu(0) As String
                ReDim DELTemp(0) As String
                TxtLock 0
                txt1(0).SetFocus
                txt1_GotFocus (0)
                Exit Sub
            Else
                Exit Sub
            End If
     Case 3
          SeekRec = MyRs.Bookmark
          MyRs.Find "D=" & txt1(0) & " ", 0, adSearchForward, 1
          If MyRs.EOF Then
              s = MsgBox("沒有符合資料!!", , "錯誤")
              MyRs.Bookmark = SeekRec
          End If
          TxtSitu True
          GetDataUp
          ProcessDown
          cmdok.Enabled = True
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
     cmdok.Enabled = True
     SeekAction = 4
     ReDim DELMenu(0) As String
     ReDim DELTemp(0) As String
     TxtLock 1
     TxtLock 0
     txt1(0).SetFocus
     txt1_GotFocus (0)
Case 5  'CHANCL
     Select Case SeekAction
     Case 0
          If Len(txt1(0)) <> 0 Then
              If MsgBox("你尚未存檔, 確定離開嗎??", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
                  Exit Sub
              End If
          End If
     Case 1
          If Len(txt1(0)) <> 0 Then
              If MsgBox("你尚未存檔, 確定離開嗎??", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
                  Exit Sub
              End If
          End If
          TxtLock 1
     Case 2
          MyRs.Bookmark = SeekRec
          GetDataUp
          TxtSitu True
          ProcessDown
          cmdok.Enabled = True
          SeekAction = 4
          ReDim DELMenu(0) As String
          ReDim DELTemp(0) As String
          TxtLock 0
          txt1(0).SetFocus
          txt1_GotFocus (0)
          Exit Sub
     Case 3
          MyRs.Bookmark = SeekRec
          GetDataUp
          TxtSitu True
          ProcessDown
          cmdok.Enabled = True
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
     ProcessUp
     SeekAction = 4 'Add by Morgan 2010/10/5 先設定狀態以便判斷是否要預設目標數
     ProcessDown
     cmdok.Enabled = True
     'SeekAction = 4 'Remove by Morgan 2010/10/5
     ReDim DELMenu(0) As String
     ReDim DELTemp(0) As String
     TxtLock 1
     TxtLock 0
     txt1(0).SetFocus
     txt1_GotFocus (0)
Case Else
End Select
   Exit Sub
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
End Sub

Private Sub MoveRec(ByVal Strindex As Integer)
With MyRs
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

Private Sub grd1_RowColChange()
With grd1
   s = .MouseRow
    .Visible = False
    If .Cols < 10 Then .Cols = 10
    For i = 1 To .Rows - 1
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
        .row = s
    End If
    If .row = 0 Then
        .row = 1
    End If
    .col = 1
    lbl2(0).Caption = "員工編號：" & .Text
    .col = 2
    lbl2(1).Caption = "姓名：" & .Text
    lbl2(2).Caption = "組別：" & GetGridValue("GP", .row)
    .col = 5
    txt1(1).Text = Trim(.Text)
    For i = 0 To .Cols - 1
        .col = i
        .CellBackColor = &HFFC0C0
    Next i
    .Visible = True
End With

End Sub
'Added by Morgan 2019/3/14
Private Function GetGridValue(PColName As String, Optional pRowID As Integer = -1) As String
   Dim iCol As Integer
   
   With grd1
   For iCol = 0 To .Cols - 1
      If .TextMatrix(0, iCol) = PColName Then
         If pRowID = -1 Then pRowID = grd1.row
         If pRowID >= 1 Then
            GetGridValue = .TextMatrix(pRowID, iCol)
         End If
      End If
   Next
   End With
End Function

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
         YNEdit 0
      Case 2
         If BolDbOk = False Then Exit Sub
         If CheckRec Then
            YNEdit 1
         End If
      Case 3
         If BolDbOk = False Then Exit Sub
         If CheckRec Then
            YNEdit 2
         End If
      Case 4
         If BolDbOk = False Then Exit Sub
         If CheckRec Then
            YNEdit 3
         End If
      Case 6
         If BolDbOk = False Then Exit Sub
         MoveRec 0
      Case 7
         If BolDbOk = False Then Exit Sub
         MoveRec 1
      Case 8
         If BolDbOk = False Then Exit Sub
         MoveRec 2
      Case 9
         If BolDbOk = False Then Exit Sub
         MoveRec 3
      Case 11
         YNEdit 4
      Case 12
         YNEdit 5
      Case 14
         Unload Me
      End Select
      If SeekAction = 4 Then
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
            'edit by nickc 2007/05/04 不給刪，請更新為 0
'            If m_bDelete Then
'                TBar1.Buttons(3).Enabled = True
'            Else
'                TBar1.Buttons(3).Enabled = False
'            End If
        End If
   End If

End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub SetGrd1()
   Dim ii As Integer
   With grd1
   .Visible = False
   If .Cols < 10 Then .Cols = 10
      .row = 0
      .RowHeight(0) = 600
      .col = 0:   .Text = "所別"
      .ColWidth(0) = 450
      .CellAlignment = flexAlignCenterCenter
      .col = 1:   .Text = "員工編號"
      .ColWidth(1) = 900
      .CellAlignment = flexAlignCenterCenter
      .col = 2:   .Text = "姓名"
      .ColWidth(2) = 750
      .CellAlignment = flexAlignCenterCenter
      .col = 3:   .Text = "P上月" & vbCrLf & "完稿數"
      .ColWidth(3) = 1000
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(3) = flexAlignRightCenter
      .col = 4:   .Text = "CFP上月" & vbCrLf & "完稿數"
      .ColWidth(4) = 1000
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(4) = flexAlignRightCenter
      .col = 5:   .Text = "基數"
      .ColWidth(5) = 600
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(5) = flexAlignRightCenter
    
      'Added by Morgan 2019/3/14 +108考核
      If DBDATE(txt1(0) & "01") >= PUB_108RuleDate Then
         m_bol108Rule = True
         Label3.Visible = True
      Else
         m_bol108Rule = False
         Label3.Visible = False
      End If
      
      'Removed by Morgan 2019/8/7 目標還是要分 P,CFP(達成情形統計要用) --王副總
      'If m_bol108Rule Then
      '   .col = 6:   .Text = "本月" & vbCrLf & "目標基數"
      '   .ColWidth(6) = 1000
      '   .CellAlignment = flexAlignCenterCenter
      '   .ColAlignment(6) = flexAlignRightCenter
      '   .col = 7:   .Text = "本月" & vbCrLf & "目標點數"
      '   .ColWidth(7) = 1000
      '   .CellAlignment = flexAlignCenterCenter
      '   .ColAlignment(7) = flexAlignRightCenter
      '   For ii = 8 To .Cols - 1
      '     .ColWidth(ii) = 0
      '   Next
      'Else
      'end 2019/8/7
      'end 2019/3/14
        .col = 6:   .Text = "P 本月" & vbCrLf & "目標件數"
        .ColWidth(6) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignRightCenter
        .col = 7:   .Text = "P本月" & vbCrLf & "目標點數"
        .ColWidth(7) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignRightCenter
        .col = 8:   .Text = "CFP 本月" & vbCrLf & "目標件數"
        .ColWidth(8) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignRightCenter
        .col = 9:   .Text = "CFP本月" & vbCrLf & "目標點數"
        .ColWidth(9) = 1000
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignRightCenter
        For ii = 10 To .Cols - 1
          .ColWidth(ii) = 0
        Next
      'End If 'Removed by Morgan 2019/8/7 達成情形還是要分 P,CFP --王副總
      .Visible = True
   End With
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Debug.Print KeyCode
If KeyCode = 40 Then  '往下
    With grd1
         .Visible = False
         For i = 1 To .Rows - 1
            .col = 0
            .row = i
            If .CellBackColor = &HFFC0C0 Then
                If i <> .Rows - 1 Then
                    For j = 0 To .Cols - 1
                        .col = j
                        .CellBackColor = QBColor(15)
                    Next j
                    .row = i + 1
                    .col = 1
                    lbl2(0).Caption = "員工編號：" & .Text
                    .col = 2
                    lbl2(1).Caption = "姓名：" & .Text
                    .col = 5
                    txt1(1).Text = Trim(.Text)
                    For j = 0 To .Cols - 1
                        .col = j
                        .CellBackColor = &HFFC0C0
                    Next j
                    .TopRow = i + 1
                End If
                Exit For
            End If
         Next i
         txt1(Index).SelStart = 0
         txt1(Index).SelLength = Len(txt1(Index))
         .Visible = True
    End With
ElseIf KeyCode = 38 Then '往上
    With grd1
         .Visible = False
         For i = 1 To .Rows - 1
            .col = 0
            .row = i
            If .CellBackColor = &HFFC0C0 Then
                If i > 1 Then
                    For j = 0 To .Cols - 1
                        .col = j
                        .CellBackColor = QBColor(15)
                    Next j
                    .row = i - 1
                    .col = 1
                    lbl2(0).Caption = "員工編號：" & .Text
                    .col = 2
                    lbl2(1).Caption = "姓名：" & .Text
                    .col = 5
                    txt1(1).Text = Trim(.Text)
                    For j = 0 To .Cols - 1
                        .col = j
                        .CellBackColor = &HFFC0C0
                    Next j
                    .TopRow = i - 1
                End If
                Exit For
            End If
         Next i
         txt1(Index).SelStart = 0
         txt1(Index).SelLength = Len(txt1(Index))
         .Visible = True
    End With
End If
End Sub

Private Sub txt1_LostFocus(Index As Integer)
m_blnCancel = False
If txt1(Index).Locked = True Then
    Exit Sub
End If
Select Case Index
Case 0
     If Len(Trim(txt1(0))) <> 0 Then
         If IsDate(ChangeTStringToTDateString(txt1(0) & "01")) = False Then
            s = MsgBox("年月輸入錯誤", , "USER 輸入錯誤")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            m_blnCancel = True
            Exit Sub
         End If
'         If Trim(txt1(0)) < "9604" Then
'            s = MsgBox("96 年 4 月 以前尚未有基數規則！", , "USER 輸入錯誤")
'            txt1(0).SetFocus
'            txt1(0).SelStart = 0
'            txt1(0).SelLength = Len(txt1(0))
'            m_blnCancel = True
'            Exit Sub
'         End If
         'Modify by Morgan 2010/10/5 新增才產生資料
         'ProcessDown
         If SeekAction = 0 Then ProcessDown
     End If
Case 1
     If IsNumeric(txt1(Index)) = False And Len(txt1(Index)) <> 0 Then
        s = MsgBox("輸入錯誤, 請輸入數字", , "USER 輸入錯誤")
        txt1(Index).SetFocus
        txt1(Index).SelStart = 0
        txt1(Index).SelLength = Len(txt1(Index))
        m_blnCancel = True
        Exit Sub
     End If
     txt1(Index) = Format(Val(txt1(Index)), "#####0.00")
Case Else
End Select
End Sub

Function CheckRec() As Boolean
   If grd1.Rows - 1 <> 0 Then
      CheckRec = True
   Else
      CheckRec = False
   End If
End Function

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Me.txt1
   If objTxt.Enabled = True Then
      Cancel = False
      txt1_GotFocus objTxt.Index
      If m_blnCancel = True Then
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function


