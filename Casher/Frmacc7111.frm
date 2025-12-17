VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc7111 
   AutoRedraw      =   -1  'True
   Caption         =   "收款明細查詢"
   ClientHeight    =   4995
   ClientLeft      =   2385
   ClientTop       =   2385
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   8760
   Begin VB.TextBox Text11 
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
      Height          =   315
      Left            =   1410
      TabIndex        =   33
      Top             =   4500
      Width           =   7065
   End
   Begin VB.TextBox Text10 
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
      Height          =   315
      Left            =   2250
      TabIndex        =   32
      Top             =   990
      Width           =   6015
   End
   Begin VB.TextBox Text2 
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
      Height          =   315
      Left            =   4650
      MaxLength       =   15
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text9 
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
      Height          =   315
      Left            =   1650
      TabIndex        =   15
      Top             =   4140
      Width           =   1575
   End
   Begin VB.TextBox Text8 
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
      Height          =   315
      Left            =   4650
      TabIndex        =   14
      Top             =   3810
      Width           =   1575
   End
   Begin VB.TextBox Text7 
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
      Height          =   315
      Left            =   1410
      MaxLength       =   12
      TabIndex        =   12
      Top             =   3420
      Width           =   6915
   End
   Begin VB.TextBox Text6 
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
      Height          =   315
      Left            =   4650
      MaxLength       =   8
      TabIndex        =   11
      Top             =   3090
      Width           =   1575
   End
   Begin VB.TextBox Text5 
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
      Height          =   315
      Left            =   1410
      MaxLength       =   12
      TabIndex        =   10
      Top             =   3090
      Width           =   1965
   End
   Begin VB.TextBox Text3 
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
      Height          =   315
      Left            =   1410
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text16 
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
      Height          =   315
      Left            =   1410
      TabIndex        =   6
      Top             =   1980
      Width           =   1155
   End
   Begin VB.TextBox Text15 
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
      Height          =   315
      Left            =   1410
      TabIndex        =   5
      Top             =   1650
      Width           =   6825
   End
   Begin VB.TextBox Text14 
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
      Height          =   315
      Left            =   1410
      TabIndex        =   4
      Top             =   1320
      Width           =   6825
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
      Height          =   315
      Left            =   1410
      TabIndex        =   3
      Top             =   990
      Width           =   825
   End
   Begin VB.TextBox Text4 
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
      Height          =   315
      Left            =   1410
      TabIndex        =   8
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   315
      Left            =   4650
      MaxLength       =   15
      TabIndex        =   1
      Top             =   270
      Width           =   1575
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   1410
      TabIndex        =   0
      Top             =   270
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
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
      Height          =   315
      Left            =   4650
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
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
      Height          =   315
      Left            =   1410
      TabIndex        =   13
      Top             =   3810
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   34
      Top             =   4530
      Width           =   495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "留分所金額"
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
      TabIndex        =   31
      Top             =   4170
      Width           =   1635
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "扣繳金額"
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
      Left            =   3660
      TabIndex        =   30
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "扣繳日期"
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
      TabIndex        =   29
      Top             =   3810
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "付款地"
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
      TabIndex        =   28
      Top             =   3450
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "收款日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   300
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "點數"
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
      TabIndex        =   26
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "到期日"
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
      Left            =   3660
      TabIndex        =   25
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "票號"
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
      Left            =   3660
      TabIndex        =   24
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "支票"
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
      TabIndex        =   23
      Top             =   2790
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   4755
      Left            =   240
      Top             =   120
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "帳號"
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
      TabIndex        =   22
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "現金"
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
      TabIndex        =   21
      Top             =   2430
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "人工收據"
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
      Left            =   3660
      TabIndex        =   20
      Top             =   660
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "案件性質"
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
      TabIndex        =   19
      Top             =   1710
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
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
      TabIndex        =   18
      Top             =   1380
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收 款 人"
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
      TabIndex        =   17
      Top             =   1050
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "電腦收據"
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
      Left            =   3660
      TabIndex        =   16
      Top             =   300
      Width           =   1455
   End
End
Attribute VB_Name = "Frmacc7111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit
Public adoacc310 As New ADODB.Recordset

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim arrItemNo
   
    Me.Icon = LoadPicture(strIcoPath)
    strFormName = Name
    Me.Width = 8880
    Me.Height = 5355
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    Image1 = LoadPicture(strBackPicPath1)
    sglWidth = Image1.Width
    sglHeight = Image1.Height
    For intX = 0 To Int(ScaleWidth / sglWidth)
        For intY = 0 To Int(ScaleHeight / sglHeight)
            PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
        Next
    Next
    Me.MaskEdBox1.Mask = DFormat
    Me.MaskEdBox2.Mask = DFormat
    Me.MaskEdBox3.Mask = DFormat
    arrItemNo = Split(strItemNo, ",")
    Me.Text1.Text = arrItemNo(0)
    Me.Text2.Text = arrItemNo(1)
    Acc310Refresh
    If adoacc310.RecordCount <> 0 Then
        FormShow
        RecordShow
    End If
    strExitControl = MsgText(602)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strExitControl = MsgText(602) Then
      StatusClear
      strFormName = MsgText(601)
      KeyEnter vbKeyEscape
      MenuEnabled
      Set Frmacc7111 = Nothing
        tool3_enabled
        MenuDisabled
        Frmacc7110.Show
        strFormName = "Frmacc7110"
      Exit Sub
   End If
   strExitControl = MsgText(602)
   strItemNo = ""
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub MaskEdBox2_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub MaskEdBox3_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub Text1_GotFocus()
    TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub Text1_LostFocus()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer

'Modified by Morgan 2011/12/26 取消a0j03 改抓 cp10
'Modified by Morgan 2011/12/27 取消 a0j20
StrSQLa = "Select A0K20||' '||ST02, A0K04, CP10||' '||getcp10desc(cp01,cp10,a0j04) , Round(Nvl(A0J09,0)/1000,1) From ACC0K0, ACC0J0, Staff,caseprogress Where A0K01=A0J13(+) And A0K20=ST01(+) And A0K01='" & ChgSQL(Me.Text1.Text) & "' And ST06='" & pub_strUserOffice & "' and cp09(+)=a0j01 Order By cp10 "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    ii = 0
    While Not rsA.EOF
        ii = ii + 1
        If ii = 1 Then
            Me.Text13.Text = "" & rsA.Fields(0).Value
            Me.Text14.Text = "" & rsA.Fields(1).Value
            Me.Text15.Text = IIf(rsA.RecordCount > 1, "(" & ii & ")" & rsA.Fields(2).Value, "" & rsA.Fields(2).Value)
            Me.Text16.Text = "" & rsA.Fields(3).Value
        Else
            Me.Text15.Text = Me.Text15.Text & " " & IIf(rsA.RecordCount > 1, "(" & ii & ")" & rsA.Fields(2).Value, "" & rsA.Fields(2).Value)
            Me.Text16.Text = Val(Me.Text16.Text) + Val("" & rsA.Fields(3).Value)
        End If
        rsA.MoveNext
    Wend
Else
    Me.Text13.Text = ""
    Me.Text14.Text = ""
    Me.Text15.Text = ""
    Me.Text16.Text = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
If Me.Text2.Text = "" Then Me.Text2.Text = Me.Text1.Text
End Sub

Private Sub Text2_GotFocus()
    TextInverse Text2
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

On Error GoTo Checking
    'edit by nick 2004/08/20  可查他所 cancel
    strSql = "Select * From ACC310 Where A3101='" & pub_strUserOffice & "' Order By A3102, A3103, A3104 "
    'strSQL = "Select * From ACC310 Order By A3102, A3103, A3104 "
    adoacc310.CursorLocation = adUseClient
    adoacc310.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
Checking:
    If Err.Number = 0 Then
        Exit Sub
    End If
    MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
On Error GoTo ErrorHandler
    Me.MaskEdBox1.Mask = ""
    If IsNull(adoacc310.Fields("A3102").Value) Then
       Me.MaskEdBox1.Text = ""
    Else
       Me.MaskEdBox1.Text = CFDate(adoacc310.Fields("A3102").Value)
    End If
    Me.MaskEdBox1.Mask = DFormat
    Me.Text1.Text = "" & adoacc310.Fields("A3103").Value
    Me.Text2.Text = "" & adoacc310.Fields("A3104").Value
    Me.Text3.Text = Val("" & adoacc310.Fields("A3105").Value)
    Me.Text4.Text = Val("" & adoacc310.Fields("A3106").Value)
    Me.MaskEdBox2.Mask = ""
    If IsNull(adoacc310.Fields("A3107").Value) Then
       Me.MaskEdBox2.Text = ""
    Else
       Me.MaskEdBox2.Text = CFDate(adoacc310.Fields("A3107").Value)
    End If
    Me.MaskEdBox2.Mask = DFormat
    Me.Text5.Text = "" & adoacc310.Fields("A3108").Value
    Me.Text6.Text = "" & adoacc310.Fields("A3109").Value
    Me.Text7.Text = "" & adoacc310.Fields("A3110").Value
    Me.MaskEdBox3.Mask = ""
    If IsNull(adoacc310.Fields("A3111").Value) Then
       Me.MaskEdBox3.Text = ""
    Else
       Me.MaskEdBox3.Text = CFDate(adoacc310.Fields("A3111").Value)
    End If
    Me.MaskEdBox3.Mask = DFormat
    Me.Text8.Text = Val("" & adoacc310.Fields("A3112").Value)
    Me.Text9.Text = Val("" & adoacc310.Fields("A3113").Value)
    'add by nick 2004/08/19
    Me.Text13.Text = "" & adoacc310.Fields("A3121").Value
    Me.Text14.Text = "" & adoacc310.Fields("A3122").Value
    Me.Text16.Text = "" & adoacc310.Fields("A3123").Value
    Me.Text11.Text = "" & adoacc310.Fields("A3124").Value
    'Modified by Morgan 2011/12/27 取消 a0j20
    Me.Text15.Text = ReConBNOurCaseNO("" & adoacc310.Fields("A0J02").Value) & "" & adoacc310.Fields("cp10N").Value
    Text10 = Empty
    If IsEmptyText(Text13) = False Then
        Text10 = GetStaffBy7100(Text13)
    End If
    
    'Text1_LostFocus

Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
    
On Error GoTo ErrorHandler
    If adoacc310.RecordCount = 0 Then
        Exit Sub
    End If
    CountShow adoacc310.Bookmark, adoacc310.RecordCount
Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

'*************************************************
'  重新整理分所收款資料
'
'*************************************************
Public Sub Acc310Refresh()
On Error GoTo Checking
    If adoacc310.State = adStateOpen Then
        adoacc310.Close
    End If
    'edit by nick 2004/08/20  可查他所 cancel
    'strSQL = "Select * From ACC310 Where A3101='" & pub_strUserOffice & "' And A3103='" & Me.Text1.Text & "' And A3104='" & Me.Text2.Text & "' Order By A3103, A3104 "
    'Modified by Morgan 2011/12/27 取消 a0j20
    strSql = "Select acc310.*,acc0j0.A0J02,getcp10desc(cp01,cp10,a0j04) cp10N From ACC310,acc0j0,caseprogress Where A3101='" & pub_strUserOffice & "' And A3103=A0J13(+) and A3103='" & Me.Text1.Text & "' And A3104='" & Me.Text2.Text & "' and cp09(+)=a0j01 Order By A3103, A3104 "
    adoacc310.CursorLocation = adUseClient
    adoacc310.Open strSql, adoTaie, adOpenDynamic, adLockBatchOptimistic
    Do While adoacc310("A3103").Value & adoacc310("A3104").Value <= Me.Text1.Text & Me.Text2.Text
        If adoacc310("A3103").Value & adoacc310("A3104").Value = Me.Text1.Text & Me.Text2.Text Then Exit Do
        adoacc310.MoveNext
    Loop
Checking:
    If Err.Number = 0 Then
        Exit Sub
    End If
    MsgBox Err.Description, , MsgText(5)
End Sub
Private Function ReConBNOurCaseNO(strCaseNo As String) As String

If strCaseNo <> "" Then
    ReConBNOurCaseNO = Replace(Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "-" & Right(Left(strCaseNo, Len(strCaseNo) - 3), 6) & "-" & Right(Left(strCaseNo, Len(strCaseNo) - 2), 1) & "-" & Right(strCaseNo, 2), "-0-00", "")
Else
    ReConBNOurCaseNO = ""
End If

End Function

Private Sub Text2_LostFocus()
    If Me.Text1.Text = "" Then Me.Text1.Text = Me.Text2.Text
End Sub

Private Sub Text3_GotFocus()
    TextInverse Me.Text3
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub Text4_GotFocus()
    TextInverse Me.Text4
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub Text5_GotFocus()
    TextInverse Me.Text5
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub Text6_GotFocus()
    TextInverse Me.Text8
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub Text7_GotFocus()
    TextInverse Me.Text7
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub Text8_GotFocus()
    TextInverse Me.Text8
End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub Text9_GotFocus()
    TextInverse Me.Text9
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub


Function GetStaffBy7100(ByVal strStuff As String, Optional ByVal bAll As Boolean = False) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim oState As String
   GetStaffBy7100 = Empty
   
   strSql = "SELECT * FROM Staff " & _
            "WHERE ST01 = '" & strStuff & "' " & IIf(oState = "1", " and st04='1' ", "")
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("ST02")) = False Then
         GetStaffBy7100 = rsTmp.Fields("ST02")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function


