VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc4170_1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "每月固定傳票資料-非分攤"
   ClientHeight    =   5115
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8760
   Begin VB.CommandButton Cmd_CopyData 
      Appearance      =   0  '平面
      Caption         =   "複製傳票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton CmdChgComp 
      Appearance      =   0  '平面
      Caption         =   "更換公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6960
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   5760
      Picture         =   "Frmacc4170_1.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   480
      Width           =   350
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   5000
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4110
      MaxLength       =   10
      TabIndex        =   1
      Top             =   480
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1308
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   612
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4170_1.frx":0102
      Height          =   3495
      Left            =   240
      TabIndex        =   12
      Top             =   1260
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "ax203"
         Caption         =   "項次"
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
         DataField       =   "ax205"
         Caption         =   "科目代號"
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
         DataField       =   "a0102"
         Caption         =   "科目名稱"
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
      BeginProperty Column03 
         DataField       =   "ax206"
         Caption         =   "借方金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "ax207"
         Caption         =   "貸方金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "ax212"
         Caption         =   "摘要"
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
      BeginProperty Column06 
         DataField       =   "ax204"
         Caption         =   "部門別"
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
      BeginProperty Column07 
         DataField       =   "ax208"
         Caption         =   "對沖代號(客)"
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
         DataField       =   "ax209"
         Caption         =   "對沖代號(業)"
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
      BeginProperty Column09 
         DataField       =   "ax214"
         Caption         =   "對沖代號(本所案號)"
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
         Size            =   344
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2399.811
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3390.236
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1950.236
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1308
      TabIndex        =   8
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   2010
      TabIndex        =   7
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "欲產生作業日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   840
      Width           =   2000
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4776
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   852
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "傳票編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc4170_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/01 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB
'Memo by Amy 2014/09/25 原ChkWorkData 函數搬至aacc_fun
'Create by Amy 2014/09/11
Option Explicit
Public adoacc020 As New ADODB.Recordset
Public adoacc021 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoacc0b0 As New ADODB.Recordset
Dim strA1R01 As String '設定字串以取得傳票編號
Dim bolNotFirst As Boolean '是否按過command1

Private Sub Cmd_CopyData_Click()
    Dim strMsg As String
    Dim bCancel As Boolean
    
    'Add by Amy 2020/04/14
    If Text1 <> MsgText(601) Then
        Call Text1_Validate(bCancel)
        If bCancel = True Then
            Text1.SetFocus
            Exit Sub
        End If
    End If
    'end 2020/04/14
    If adoadodc1.RecordCount = 0 Then
        MsgBox "無資料可複製", , MsgText(5)
        MaskEdBox2.SetFocus
        Exit Sub
    End If
    If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
        MsgBox Label1 & MsgText(52), , MsgText(5)
        MaskEdBox2.SetFocus
        Exit Sub
    End If
    Call MaskEdBox2_Validate(bCancel)
    If bCancel = True Then
        MaskEdBox2.SetFocus
        Exit Sub
    End If
    If ChkWorkData(Text1, DBDATE(MaskEdBox2), strMsg) = False Then
        MsgBox "作業日" & strMsg, , MsgText(5)
        MaskEdBox2.SetFocus
        Exit Sub
    End If
    FormSave
    FormClear
    AdodcRefresh
End Sub

Private Sub CmdChgComp_Click()
    Dim strNewComp As String
    
    Frmacc41c2.Show vbModal
    strNewComp = strCompanyNo
    If strNewComp <> "" And strNewComp <> Me.Text1 Then
        Text1 = strNewComp
        MaskEdBox1.Mask = MsgText(601)
        MaskEdBox1.Text = MsgText(601)
        MaskEdBox1.Mask = DFormat
        
        MaskEdBox2.Mask = MsgText(601)
        MaskEdBox2.Text = MsgText(601)
        MaskEdBox2.Mask = DFormat
        
        Text2 = MsgText(601)
        AdodcRefresh
      
    End If
    strCompanyNo = MsgText(601)
End Sub

Private Sub Command1_Click()
    Dim bolCancel As Boolean 'Add by Amy 2020/04/14
    
    If Text1 = "" Or Text2 = "" Then
        MsgBox MsgText(181), , MsgText(5)
        Exit Sub
    End If
    'Add by Amy 2020/04/14
    If Text1 <> MsgText(601) Then
        Call Text1_Validate(bolCancel)
        If bolCancel = True Then
            Exit Sub
        End If
    End If
    'end 2020/04/14
    bolNotFirst = True
    Text1.Enabled = False '按過後鎖住
    AdodcRefresh
    If adoadodc1.RecordCount = 0 Then
        MsgBox MsgText(9007), , MsgText(5)
        FormClear
        Text2.SetFocus
    Else
      If IsNull(adoacc020.Fields("a0205").Value) Then
        MaskEdBox1.Text = MsgText(601)
      Else
        MaskEdBox1.Text = CFDate(Trim(str(adoacc020.Fields("a0205").Value)))
      End If
        MaskEdBox1.Mask = DFormat
        MaskEdBox2.SetFocus
    End If
End Sub


Private Sub Form_Load()
    Dim intX As Integer
    Dim intY As Integer
    Dim sglWidth As Single
    Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 5700
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
    MaskEdBox1.Mask = DFormat
    MaskEdBox2.Mask = DFormat
    OpenTable
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Call PUB_GetLock("", "Frmacc4170_1") 'Add by Amy 2020/06/07
   Set Frmacc4170_1 = Nothing
End Sub

Private Sub FormSave()
    Dim strNewA0202 As String, strSave As String
On Error GoTo Checking
  
    '依公司別設定
    If Text1 = "J" Then
        strA1R01 = MsgText(819) 'JD
    ElseIf Text1 = "L" Then
        strA1R01 = MsgText(820) 'LD
    Else
        strA1R01 = MsgText(801) 'D
    End If
    
    strNewA0202 = AccAutoNo(strA1R01, 4, Val(Mid(MaskEdBox2.Text, 1, 3)), Val(Mid(MaskEdBox2.Text, 5, 2)))
    If ChkDataExists("Acc020", "A0201", "A0202", Text1, strNewA0202) Or ChkDataExists("Acc021", "AX201", "AX202", Text1, strNewA0202) Then
        MsgBox Text1 & "公司傳票 " & strNewA0202 & " 資料已存在,請洽電腦中心!!", , MsgText(5)
        Exit Sub
    End If
    
    If adoacc0b0.State = adStateOpen Then
        adoacc0b0.Close
    End If
    adoacc0b0.CursorLocation = adUseClient
    adoacc0b0.Open "Select a0b10 From Acc0b0 Where A0b10 = '01'", adoTaie, adOpenStatic, adLockReadOnly
    If adoacc0b0.RecordCount <> 0 Then
        MsgBox MsgText(197), , MsgText(5)
        adoacc0b0.Close
        Exit Sub
    End If
    adoacc0b0.Close
     
    cnnConnection.BeginTrans
   
    adoTaie.Execute "update acc0b0 set a0b10 = '01'"
   
    '新增傳票主檔
    adoacc020.AddNew
    adoacc020.Fields("A0201") = Text1
    adoacc020.Fields("A0202") = strNewA0202
    adoacc020.Fields("A0205") = Val(FCDate(MaskEdBox2.Text))
    adoacc020.Fields("A0206") = Val(strSrvDate(2))
    adoacc020.Fields("A0207") = ServerTime
    adoacc020.Fields("A0208") = strUserNum
    adoacc020.UpdateBatch
    
    '新增傳票明細
    With adoadodc1
        .MoveFirst
        Do While .EOF = False
             adoacc021.AddNew
             adoacc021.Fields("AX201") = Text1
             adoacc021.Fields("AX202") = strNewA0202
             adoacc021.Fields("AX203") = .Fields("AX203")
             adoacc021.Fields("AX204") = .Fields("AX204")
             adoacc021.Fields("AX205") = .Fields("AX205")
             adoacc021.Fields("AX206") = .Fields("AX206")
             adoacc021.Fields("AX207") = .Fields("AX207")
             adoacc021.Fields("AX208") = .Fields("AX208")
             adoacc021.Fields("AX209") = .Fields("AX209")
             adoacc021.Fields("AX211") = .Fields("AX211")
             adoacc021.Fields("AX212") = .Fields("AX212")
             adoacc021.Fields("AX213") = .Fields("AX213")
             adoacc021.Fields("AX214") = .Fields("AX214")
             adoacc021.Fields("AX215") = Null
             adoacc021.UpdateBatch
             .MoveNext
        Loop
    End With
    adoTaie.Execute "update acc0b0 set a0b10 = null"
    strSave = AccSaveAutoNo(strA1R01, Mid(strNewA0202, 7, 4), Mid(strNewA0202, 2, 3), Mid(strNewA0202, 5, 2))
    cnnConnection.CommitTrans
    
    MsgBox Text1 & "公司傳票已複製!! (新傳票號碼:" & strNewA0202 & ")"
    
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   cnnConnection.RollbackTrans
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub AdodcRefresh()
On Error GoTo Checking
    If adoacc020.State = adStateOpen Then
        adoacc020.Close
    End If
    adoacc020.CursorLocation = adUseClient
    strExc(0) = "Select * From acc020 Where a0201 = '" & Text1 & "' And a0202 = '" & Text2 & "' Order by a0201 asc, a0202 asc"
    adoacc020.Open strExc(0), adoTaie, adOpenDynamic, adLockBatchOptimistic
    
    If adoacc021.State = adStateOpen Then
        adoacc021.Close
    End If
    adoacc021.CursorLocation = adUseClient
    strExc(0) = "Select * From acc021 Where ax201 = '" & Text1 & "' And ax202 = '" & Text2 & "' Order by ax201 asc, ax202 asc, ax203 asc"
    adoacc021.Open strExc(0), adoTaie, adOpenDynamic, adLockBatchOptimistic
   
    If adoadodc1.State = adStateOpen Then
        adoadodc1.Close
    End If
   adoadodc1.CursorLocation = adUseClient
   strExc(0) = "Select * From acc021, acc010 Where ax201 = '" & Text1 & "' And ax202 = '" & Text2 & "' And ax205 = a0101 (+) Order by ax201 asc, ax202 asc, ax203 asc"
   adoadodc1.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.Requery
  
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Function ChkDataExists(strTable As String, KeyFields1 As String, KeyFields2 As String, strValue1 As String, strValue2 As String) As Boolean
    Dim adocheck As New ADODB.Recordset
    Dim strChk As String
    Dim intC As Integer
    
    strChk = "Select " & KeyFields1 & "," & KeyFields2 & " From " & strTable & _
                 " Where " & KeyFields1 & " = '" & strValue1 & "' And " & KeyFields2 & " = '" & strValue2 & "' "
    
    intC = 1
    Set adocheck = ClsLawReadRstMsg(intC, strChk)
    If adocheck.RecordCount = 0 Then
        ChkDataExists = False
    Else
        ChkDataExists = True
    End If
    Set adocheck = Nothing
End Function

Private Sub OpenTable()
On Error GoTo Checking
   
    adoacc020.CursorLocation = adUseClient
    strExc(0) = "Select * From acc020 Where RowNum <1 Order by a0201 asc, a0202 asc"
    adoacc020.Open strExc(0), adoTaie, adOpenDynamic, adLockBatchOptimistic
    
    adoacc021.CursorLocation = adUseClient
    strExc(0) = "Select * From acc021 Where RowNum <1 Order by ax201 asc, ax202 asc, ax203 asc"
    adoacc021.Open strExc(0), adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   adoadodc1.CursorLocation = adUseClient
   strExc(0) = "Select * From acc021, acc010 Where RowNum <1 And ax205 = a0101 (+) Order by ax201 asc, ax202 asc, ax203 asc"
   adoadodc1.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub FormClear()
    Text2 = ""
    MaskEdBox1.Mask = MsgText(601)
    MaskEdBox1.Text = ""
    MaskEdBox1.Mask = DFormat
    MaskEdBox2.Mask = MsgText(601)
    MaskEdBox2.Text = ""
    MaskEdBox2.Mask = DFormat
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
    If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
        Exit Sub
    End If
    If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
        Cancel = True
        MsgBox Label1 & MsgText(63), , MsgText(5)
        MaskEdBox2.SetFocus
        Exit Sub
    End If
    If ChkWorkDay(DBDATE(MaskEdBox2)) = False Then
        Cancel = True
        MsgBox Label1 & "必須為工作日!!", , MsgText(5)
        MaskEdBox2.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

'Modify by Amy 2020/04/14 原:Text1_Change ,因複製傳票鈕需檢查,故搬至此
Private Sub Text1_Validate(Cancel As Boolean)
    If Text1 = MsgText(601) Then
        Exit Sub
    End If
    'Modify by Amy 2020/04/14 改抓function
    'If Text1 <> "J" And Text1 <> "1" Then
    If InStr(GetBookKeepCmp, Text1) = 0 Then
        Text13 = ""
        Cancel = True
        'MsgBox Label3 & "只能輸入1 或 J", , MsgText(5)
        MsgBox Label3 & MsgText(63), , MsgText(5)
        Exit Sub
    End If
    'end 2020/04/14
    Text13 = A0802Query(Text1)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub
