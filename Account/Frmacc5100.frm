VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc5100 
   AutoRedraw      =   -1  'True
   Caption         =   "系統參數變更"
   ClientHeight    =   3740
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   3730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3740
   ScaleWidth      =   3730
   Begin VB.TextBox TxtAxb 
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
      Height          =   315
      Index           =   3
      Left            =   2520
      TabIndex        =   13
      Top             =   3240
      Width           =   1000
   End
   Begin VB.TextBox TxtAxb 
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
      Height          =   315
      Index           =   2
      Left            =   2520
      TabIndex        =   11
      Top             =   2880
      Width           =   1000
   End
   Begin VB.TextBox Text2 
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
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2160
      Width           =   612
   End
   Begin VB.TextBox Text1 
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
      Height          =   300
      Left            =   1800
      TabIndex        =   7
      Top             =   1680
      Width           =   612
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1572
      _ExtentX        =   2769
      _ExtentY        =   529
      _Version        =   393216
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
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   1572
      _ExtentX        =   2769
      _ExtentY        =   529
      _Version        =   393216
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   1572
      _ExtentX        =   2769
      _ExtentY        =   529
      _Version        =   393216
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
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "實績期末保留傳票日期"
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
      Left            =   200
      TabIndex        =   12
      Top             =   3240
      Width           =   2595
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "實績期末保留傳票日期"
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
      Left            =   200
      TabIndex        =   10
      Top             =   2880
      Width           =   2595
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   885
      Left            =   120
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "公司帳目別"
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
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   2535
      Left            =   240
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "作業狀態"
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
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "年結轉日期"
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
      TabIndex        =   5
      Top             =   1200
      Width           =   1212
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "月結帳日期"
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
      Top             =   720
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "過帳日期"
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
      TabIndex        =   3
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc5100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/2/9 Form2.0不用改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit
Public adoacc0b0 As New ADODB.Recordset

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 3850
   Me.Height = 4245 'Modify by Amy 2017/04/28 原:3300
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath5)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = DFormat
   'Modify by Amy 2014/02/14 顯示第一筆
   'Text2 = strAccount
   FormDisabled
   'end 2014/02/14
   OpenTable
   If adoacc0b0.RecordCount <> 0 Then
      adoacc0b0.MoveLast
      adoacc0b0.MoveFirst
      RecordShow
      FormShow
   End If
   TxtAxb(2).Enabled = False
   TxtAxb(3).Enabled = False
   Call GetMaxAxb
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Mark by Amy 2014/02/14 秀玲說固定設1
'   If Text2 = "" Or IsNumeric(Text2) = False Then
      strAccount = "1"
   'Mark by Amy 2014/02/14
'   Else
'      strAccount = Text2
'   End If
   'end 2014/02/14
   
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc5100 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         MaskEdBox2.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0b0.CursorLocation = adUseClient
   adoacc0b0.Open "select * from acc0b0", adoTaie, adOpenDynamic, adLockBatchOptimistic
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表(系統參數資料)
'
'*************************************************
Public Sub FormShow()
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0b0.Fields("a0b01").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0b0.Fields("a0b01").Value)
   End If
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc0b0.Fields("a0b02").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc0b0.Fields("a0b02").Value)
   End If
   MaskEdBox2.Mask = DFormat
   MaskEdBox3.Mask = MsgText(601)
   If IsNull(adoacc0b0.Fields("a0b03").Value) Then
      MaskEdBox3.Text = MsgText(601)
   Else
      MaskEdBox3.Text = CFDate(adoacc0b0.Fields("a0b03").Value)
   End If
   MaskEdBox3.Mask = DFormat
   If IsNull(adoacc0b0.Fields("a0b10").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc0b0.Fields("a0b10").Value
   End If
   Text2 = adoacc0b0.Fields("a0b04").Value 'Add by Amy 2014/02/14
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    If strSaveConfirm = MsgText(601) Then
        Exit Sub
    End If
    If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
       MsgBox Label1 & MsgText(52), , MsgText(5)
       Cancel = True
       MaskEdBox1.SetFocus
       Exit Sub
    End If
    If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
       MsgBox Label1 & MsgText(63), , MsgText(5)
       Cancel = True
       MaskEdBox1.SetFocus
       Exit Sub
    End If
End Sub

Private Sub MaskEdBox2_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         MaskEdBox3.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
    If strSaveConfirm = MsgText(601) Then
        Exit Sub
    End If
    If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
       MsgBox Label2 & MsgText(52), , MsgText(5)
       Cancel = True
       MaskEdBox2.SetFocus
       Exit Sub
    End If
    If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
       MsgBox Label2 & MsgText(63), , MsgText(5)
       Cancel = True
       MaskEdBox2.SetFocus
       Exit Sub
    End If
End Sub

Private Sub MaskEdBox3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         MaskEdBox1.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   Frmacc0000.StatusBar1.Panels(2).Text = adoacc0b0.Bookmark & MsgText(35) & adoacc0b0.RecordCount
End Sub

Private Sub MaskEdBox3_Validate(Cancel As Boolean)
    If strSaveConfirm = MsgText(601) Then
        Exit Sub
    End If
    If MaskEdBox3.Text = MsgText(601) Or MaskEdBox3.Text = MsgText(29) Then
       MsgBox Label3 & MsgText(52), , MsgText(5)
       Cancel = True
       MaskEdBox3.SetFocus
       Exit Sub
    End If
    If DateCheck(MaskEdBox3.Text) = MsgText(603) Then
       MsgBox Label3 & MsgText(63), , MsgText(5)
       Cancel = True
       MaskEdBox3.SetFocus
       Exit Sub
    End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If strSaveConfirm = MsgText(601) Then
        Exit Sub
    End If
    If Text1 <> MsgText(601) And Text1 <> "01" Then
        MsgBox Label4 & "只允許輸入空白或 01！", , MsgText(5)
        Cancel = True
        Text1.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

'Modify  by Amy 2014/02/14
Public Sub Frmacc5100_Save()
   On Error GoTo Checking
   
      If adoacc0b0.RecordCount = 0 Then
         adoacc0b0.AddNew
      End If
      'Add by Amy 2014/02/14
      adoacc0b0.Fields("a0b04").Value = Text2
      If Trim(Text1) <> MsgText(601) Then
         adoacc0b0.Fields("a0b10").Value = "01"
      Else
         adoacc0b0.Fields("a0b10").Value = Null
      End If
      'end 2014/02/14
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         adoacc0b0.Fields("a0b01").Value = Val(FCDate(MaskEdBox1.Text))
      Else
         adoacc0b0.Fields("a0b01").Value = Null
      End If
      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
         adoacc0b0.Fields("a0b02").Value = Val(FCDate(MaskEdBox2.Text))
      Else
         adoacc0b0.Fields("a0b02").Value = Null
      End If
      If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
         adoacc0b0.Fields("a0b03").Value = Val(FCDate(MaskEdBox3.Text))
      Else
         adoacc0b0.Fields("a0b03").Value = Null
      End If
      
      adoacc0b0.UpdateBatch
      MsgBox MsgText(17), , MsgText(5)
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   
End Sub

'Add 2014/02/14
Public Function FormCheck() As Boolean
    Dim bCancel As Boolean
    bCancel = False
    Call MaskEdBox1_Validate(bCancel)
    If bCancel = True Then
        FormCheck = False
        Exit Function
    End If
    Call MaskEdBox2_Validate(bCancel)
    If bCancel = True Then
        FormCheck = False
        Exit Function
    End If
     Call MaskEdBox3_Validate(bCancel)
    If bCancel = True Then
        FormCheck = False
        Exit Function
    End If
    Call Text1_Validate(bCancel)
    If bCancel = True Then
        FormCheck = False
        Exit Function
    End If
    FormCheck = True
End Function

Public Sub FormDisabled()
    Frmacc0000.Toolbar1.Buttons.Item(9).Enabled = False
    MaskEdBox1.Enabled = False
    MaskEdBox2.Enabled = False
    MaskEdBox3.Enabled = False
    Text1.Enabled = False
End Sub

Public Sub FormEnabled()
    Frmacc0000.Toolbar1.Buttons.Item(9).Enabled = False
    MaskEdBox1.Enabled = True
    MaskEdBox2.Enabled = True
    MaskEdBox3.Enabled = True
    Text1.Enabled = True
End Sub

Public Sub MoveFirstRecord()
    If adoacc0b0.RecordCount <> 0 Then
         adoacc0b0.MoveFirst
         FormShow
         RecordShow
      End If
End Sub

Public Sub MoveLastRecord()
    If adoacc0b0.RecordCount <> 0 Then
         adoacc0b0.MoveLast
         FormShow
         RecordShow
      End If
End Sub

Public Sub MoveNextRecord()
    If adoacc0b0.EOF = False Then
         adoacc0b0.MoveNext
         If adoacc0b0.EOF Then
            adoacc0b0.MoveLast
            MsgBox MsgText(8), , MsgText(5)
         End If
         FormShow
         RecordShow
      End If
End Sub

Public Sub MovePreviousRecord()
    If adoacc0b0.BOF = False Then
         adoacc0b0.MovePrevious
         If adoacc0b0.BOF Then
            adoacc0b0.MoveFirst
            MsgBox MsgText(7), , MsgText(5)
         End If
         FormShow
         RecordShow
      End If
End Sub
'end 2014/02/14

'Add by Amy 2017/04/28
Private Sub GetMaxAxb()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    Dim intQ As Integer
    
    strQ = "Select * From (Select '2' Sort,Max(Axb02) From Acc0b1 " & _
                           "Union Select '3' Sort,Max(Axb03) From Acc0b1) " & _
               "Order by Sort"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        Do While RsQ.EOF = False
            If Not IsNull(RsQ.Fields(1)) Then TxtAxb(Val(RsQ.Fields("Sort"))) = ChangeTStringToTDateString(RsQ.Fields(1))
            RsQ.MoveNext
        Loop
    End If
    RsQ.Close
End Sub
