VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc3430 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "託收票據資料表"
   ClientHeight    =   3324
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5952
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3324
   ScaleWidth      =   5952
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   990
      TabIndex        =   30
      Text            =   "Combo1"
      Top             =   450
      Width           =   2375
   End
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   990
      TabIndex        =   0
      Top             =   60
      Width           =   4925
   End
   Begin VB.ComboBox Combo13 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   990
      Style           =   2  '單純下拉式
      TabIndex        =   28
      Top             =   1560
      Width           =   4925
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   3540
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   450
      Width           =   2375
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2190
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "N"
      Top             =   1200
      Width           =   615
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   6
      Top             =   3828
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   3600
      TabIndex        =   7
      Top             =   3828
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   8
      Top             =   4188
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   3600
      TabIndex        =   9
      Top             =   4188
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   10
      Top             =   4548
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo8 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   3600
      TabIndex        =   11
      Top             =   4548
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo9 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   12
      Top             =   4908
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo10 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   3600
      TabIndex        =   13
      Top             =   4908
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo11 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   960
      TabIndex        =   14
      Top             =   5268
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo12 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   3600
      TabIndex        =   15
      Top             =   5268
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   660
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   1935
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   2910
      TabIndex        =   3
      Top             =   840
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   990
      TabIndex        =   2
      Top             =   840
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "label13"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1500
      Left            =   480
      TabIndex        =   31
      Top             =   2400
      Width           =   5000
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   29
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   27
      Top             =   105
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "是否包括已兌現(Y/N)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   26
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2670
      TabIndex        =   25
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "託收日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   24
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3390
      TabIndex        =   23
      Top             =   450
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "排序方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   22
      Top             =   3552
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   720
      TabIndex        =   21
      Top             =   3828
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Image2 
      Height          =   252
      Left            =   3000
      Picture         =   "Frmacc3430.frx":0000
      Stretch         =   -1  'True
      Top             =   3828
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   720
      TabIndex        =   20
      Top             =   4188
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Image3 
      Height          =   252
      Left            =   3000
      Picture         =   "Frmacc3430.frx":0442
      Stretch         =   -1  'True
      Top             =   4188
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   720
      TabIndex        =   19
      Top             =   4548
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Image4 
      Height          =   252
      Left            =   3000
      Picture         =   "Frmacc3430.frx":0884
      Stretch         =   -1  'True
      Top             =   4548
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   720
      TabIndex        =   18
      Top             =   4908
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Image5 
      Height          =   252
      Left            =   3000
      Picture         =   "Frmacc3430.frx":0CC6
      Stretch         =   -1  'True
      Top             =   4908
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "5."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   720
      TabIndex        =   17
      Top             =   5268
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Image6 
      Height          =   252
      Left            =   3000
      Picture         =   "Frmacc3430.frx":1108
      Stretch         =   -1  'True
      Top             =   5268
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2412
      Left            =   240
      Top             =   3348
      Visible         =   0   'False
      Width           =   4692
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "銀行帳號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc3430"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/04/06 Form2.0 已修改 (Printer改以Excel印)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Dim adoacc0e0 As New ADODB.Recordset
Dim adoaccrpt303 As New ADODB.Recordset
Dim adoquery As New ADODB.Recordset
Dim strSort1, strSort2, strSort3, strSort4, strSort5 As String
'Dim dllaccrpt303 As Object 'Mark by Amy 2022/04/06 不使用
Dim strPrinter As String 'Add by Amy 2018/10/30
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/04/17
'Add by Amy 2022/04/06
Dim strField, intWidth
Dim i As Integer, intField As Integer, intR As Integer, intTitleRow As Integer

'Add by Sindy 2020/04/17
Private Sub SetCompN()
    strCmpN = "": strCmp = ""
    If Trim(CboCmp) <> MsgText(601) Then
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, False, True)
End Sub

Private Sub CboCmp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCmp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboCmp) = MsgText(601) Then Exit Sub
    
    strCmp = CboCmp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label6 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/17

Private Sub Combo10_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo11.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo11_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo12.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo12_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo1.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo4.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo4_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo5.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo5_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo6.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo6_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo7.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo7_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo8.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo8_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo9.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Combo9_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Combo10.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Command1_Click()
Dim bCancel As Boolean
   
   'Add By Sindy 2020/4/23
   If CboCmp.Text = MsgText(601) Then
      MsgBox Label6 & MsgText(52), , MsgText(5)
      Exit Sub
   End If
   Call CboCmp_Validate(bCancel)
   If bCancel = True Then
      Exit Sub
   End If
   '2020/4/23 END
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   
   Call SetCompN 'Add by Sindy 2020/04/23
   
   Screen.MousePointer = vbHourglass
   Accrpt303Delete
   ProduceData
   PUB_SetOsDefaultPrinter Combo13 'Add by Amy 2018/10/30
   If adoaccrpt303.State = adStateOpen Then
      adoaccrpt303.Close
   End If
   adoaccrpt303.CursorLocation = adUseClient
   'Modify by Amy 2022/04/06 +r30301 避免多人同時使用,改Excel印
   strExc(1) = "Select * From Accrpt303 Where r30301='" & strUserNum & "'  Order by r30301 asc, r30309 desc,  r30314 desc"
   adoaccrpt303.Open strExc(1), adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt303.RecordCount <> 0 Then
      'Modify by Amy 2022/04/06
      '20140120START Modify By eric
      'dllaccrpt303.Acc3430 ReportTitle(303) & "-" & strCmpN, Combo1, Combo2, MaskEdBox1.Text, MaskEdBox2.Text, Text3, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      'dllaccrpt303.Acc3430 ReportTitle(303), Combo1, Combo2, MaskEdBox1.Text, MaskEdBox2.Text, Text3, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      '20140120END
      ExcelSave
   End If
   'end 2022/04/06
   adoaccrpt303.Close
   PUB_SetOsDefaultPrinter strPrinter 'Add by Amy 2018/10/30
   Screen.MousePointer = vbDefault
   FormClear
   'Modify By Cheng 2002/03/29
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   'Modify by Amy 2022/04/06
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Frmacc0000.StatusBar1.Panels(1).Text = "請以A4紙印"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      'Modify By Cheng 2002/03/29
'      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   'Add by Amy 2023/11/13 +提醒文字
   Label13.Caption = "託收票據資料表會印出二份" & vbCrLf & _
               "一份交銀行" & vbCrLf & _
               "一份本所留存"
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 6075 'Modify by Amy 2023/05/23 原:5250
   Me.Height = 3746  'Modify by Amy 2023/08/18 原:2990
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Morgan 2006/7/27 預設當天--瑞婷
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = CFDate(strSrvDate(2))
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = CFDate(strSrvDate(2))
   'end 2006/7/27
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Combo4.AddItem MsgText(1)
   Combo4.AddItem MsgText(2)
   Combo6.AddItem MsgText(1)
   Combo6.AddItem MsgText(2)
   Combo8.AddItem MsgText(1)
   Combo8.AddItem MsgText(2)
   Combo10.AddItem MsgText(1)
   Combo10.AddItem MsgText(2)
   Combo12.AddItem MsgText(1)
   Combo12.AddItem MsgText(2)
   Combo4 = MsgText(1)
   Combo6 = MsgText(1)
   Combo8 = MsgText(1)
   Combo10 = MsgText(1)
   Combo12 = MsgText(1)
   ComboAdd
   'Add by Morgan 2006/7/24
   PUB_SetAccount Combo1
   PUB_SetAccount Combo2
   'end 2006/7/24
   'Modify By Cheng 2002/03/29
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   'Set dllaccrpt303 = CreateObject("AccReport.ReportSelect") 'Mark by Amy 2022/04/06 不使用
   PUB_SetPrinter Me.Name, Combo13, strPrinter 'Add by Amy 2018/10/30 +印表機
   
   'Add by Sindy 2020/04/17 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, False, False, False, , 1)
   'end 2020/04/17
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   If Me.Combo13.Text <> Me.Combo13.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo13.Name, "0", "0", Me.Combo13.Text
   End If
   'Set dllaccrpt303 = Nothing 'Mark by Amy 2022/04/06 不使用
   Set Frmacc3430 = Nothing
End Sub

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "託收銀行"
   strSort2 = "票據號碼"
   strSort3 = "收票日期"
   strSort4 = "往來對象"
   strSort5 = "收票銀行"
   Combo3.AddItem strSort1
   Combo3.AddItem strSort2
   Combo3.AddItem strSort3
   Combo3.AddItem strSort4
   Combo3.AddItem strSort5
   Combo5.AddItem strSort1
   Combo5.AddItem strSort2
   Combo5.AddItem strSort3
   Combo5.AddItem strSort4
   Combo5.AddItem strSort5
   Combo7.AddItem strSort1
   Combo7.AddItem strSort2
   Combo7.AddItem strSort3
   Combo7.AddItem strSort4
   Combo7.AddItem strSort5
   Combo9.AddItem strSort1
   Combo9.AddItem strSort2
   Combo9.AddItem strSort3
   Combo9.AddItem strSort4
   Combo9.AddItem strSort5
   Combo11.AddItem strSort1
   Combo11.AddItem strSort2
   Combo11.AddItem strSort3
   Combo11.AddItem strSort4
   Combo11.AddItem strSort5
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strOrder1, strOrder2, strOrder3, strOrder4, strOrder5 As String
Dim strSql As String
   
On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e19 asc"
         Else
            strOrder1 = " order by a0e19 desc"
         End If
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e02 asc"
         Else
            strOrder1 = " order by a0e02 desc"
         End If
      Case strSort3
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e13 asc"
         Else
            strOrder1 = " order by a0e13 desc"
         End If
      Case strSort4
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e06 asc"
         Else
            strOrder1 = " order by a0e06 desc"
         End If
      Case strSort5
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e01 asc"
         Else
            strOrder1 = " order by a0e01 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   Select Case Combo5
      Case strSort1
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e19 asc"
         Else
            strOrder2 = ", a0e19 desc"
         End If
      Case strSort2
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e02 asc"
         Else
            strOrder2 = ", a0e02 desc"
         End If
      Case strSort3
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e13 asc"
         Else
            strOrder2 = ", a0e13 desc"
         End If
      Case strSort4
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e06 asc"
         Else
            strOrder2 = ", a0e06 desc"
         End If
      Case strSort5
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e01 asc"
         Else
            strOrder2 = ", a0e01 desc"
         End If
      Case Else
         strOrder2 = MsgText(601)
   End Select
   Select Case Combo7
      Case strSort1
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e19 asc"
         Else
            strOrder3 = ", a0e19 desc"
         End If
      Case strSort2
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e02 asc"
         Else
            strOrder3 = ", a0e02 desc"
         End If
      Case strSort3
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e13 asc"
         Else
            strOrder3 = ", a0e13 desc"
         End If
      Case strSort4
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e06 asc"
         Else
            strOrder3 = ", a0e06 desc"
         End If
      Case strSort5
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e01 asc"
         Else
            strOrder3 = ", a0e01 desc"
         End If
      Case Else
         strOrder3 = MsgText(601)
   End Select
   Select Case Combo9
      Case strSort1
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e19 asc"
         Else
            strOrder4 = ", a0e19 desc"
         End If
      Case strSort2
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e02 asc"
         Else
            strOrder4 = ", a0e02 desc"
         End If
      Case strSort3
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e13 asc"
         Else
            strOrder4 = ", a0e13 desc"
         End If
      Case strSort4
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e06 asc"
         Else
            strOrder4 = ", a0e06 desc"
         End If
      Case strSort5
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e01 asc"
         Else
            strOrder4 = ", a0e01 desc"
         End If
      Case Else
         strOrder4 = MsgText(601)
   End Select
   Select Case Combo11
      Case strSort1
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e19 asc"
         Else
            strOrder5 = ", a0e19 desc"
         End If
      Case strSort2
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e02 asc"
         Else
            strOrder5 = ", a0e02 desc"
         End If
      Case strSort3
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e13 asc"
         Else
            strOrder5 = ", a0e13 desc"
         End If
      Case strSort4
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e06 asc"
         Else
            strOrder5 = ", a0e06 desc"
         End If
      Case strSort5
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e01 asc"
         Else
            strOrder5 = ", a0e01 desc"
         End If
      Case Else
         strOrder5 = MsgText(601)
   End Select
   
   '20140120START Modify By eric
   'Modify By Sindy 2020/4/23
   'strSql = " and a0e23 = '" & IIf(Text1 = "2", "J", "1") & "' "
   strSql = " and a0e23 = '" & strCmp & "' "
   '2020/4/23 END
   
   If Combo1 <> MsgText(601) Then
      'Modify by Amy 2023/05/23 原:Combo1
      strSql = strSql & " and a0e20 >= '" & Left(Combo1, InStr(Combo1, " ") - 1) & "'"
   End If
   'If Combo1 <> MsgText(601) Then
   '   strSql = " and a0e20 >= '" & Combo1 & "'"
   'End If
   '20140120END
      
   If Combo2 <> MsgText(601) Then
      'Modify by Amy 2023/05/23 原:Combo1
      strSql = strSql & " and a0e20 <= '" & Left(Combo2, InStr(Combo2, " ") - 1) & "'"
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0e14 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0e14 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   Select Case Text3
      Case MsgText(603)
         strSql = strSql & " AND (A0E21 IS NULL OR A0E21 = 0)"
   End Select
 
   adoaccrpt303.CursorLocation = adUseClient
   'Modify by Amy 2022/04/06 +r30301 避免多人同時使用
   adoaccrpt303.Open "Select * From Accrpt303 Where r30301='" & strUserNum & "' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open "select * from acc0e0 where a0e04 = '" & MsgText(18) & "' AND (A0E14 IS NOT NULL AND A0E14 <> 0)" & strSql & strOrder1 & strOrder2 & strOrder3 & strOrder4 & strOrder5, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0e0.RecordCount = 0 Then
      adoacc0e0.Close
      adoaccrpt303.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0e0.EOF = False
      adoaccrpt303.AddNew
      adoaccrpt303.Fields("r30301").Value = strUserNum
      If IsNull(adoacc0e0.Fields("a0e19").Value) Then
         adoaccrpt303.Fields("r30302").Value = Null
      Else
         adoquery.CursorLocation = adUseClient
         'Modify By Cheng 2002/03/29
'         adoquery.Open "SELECT A0G02 FROM ACC0G0 WHERE A0G01 = '" & adoacc0e0.Fields("A0E19").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         adoquery.Open "SELECT A0G02 FROM ACC0G0 WHERE A0G01 = '" & adoacc0e0.Fields("A0E01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) = False Then
               adoaccrpt303.Fields("r30302").Value = adoquery.Fields(0).Value
            End If
         End If
         adoquery.Close
      End If
      If IsNull(adoacc0e0.Fields("a0e20").Value) Then
         adoaccrpt303.Fields("r30303").Value = Null
      Else
         adoaccrpt303.Fields("r30303").Value = adoacc0e0.Fields("a0e20").Value
      End If
      adoaccrpt303.Fields("r30304").Value = adoacc0e0.Fields("a0e02").Value
      If IsNull(adoacc0e0.Fields("a0e11").Value) Then
         adoaccrpt303.Fields("r30305").Value = 0
      Else
         adoaccrpt303.Fields("r30305").Value = Val(adoacc0e0.Fields("a0e11").Value)
      End If
      If IsNull(adoacc0e0.Fields("a0e08").Value) Then
         adoaccrpt303.Fields("r30306").Value = Null
      Else
         Select Case adoacc0e0.Fields("a0e08").Value
            Case Mid(ComboItem(11), 1, 1)
               adoaccrpt303.Fields("r30306").Value = Mid(ComboItem(11), 4, 2)
            Case Mid(ComboItem(12), 1, 1)
               adoaccrpt303.Fields("r30306").Value = Mid(ComboItem(12), 4, 2)
            Case Mid(ComboItem(13), 1, 1)
               adoaccrpt303.Fields("r30306").Value = Mid(ComboItem(13), 4, 2)
         End Select
      End If
      If IsNull(adoacc0e0.Fields("a0e13").Value) Then
         adoaccrpt303.Fields("r30307").Value = Null
      Else
         adoaccrpt303.Fields("r30307").Value = adoacc0e0.Fields("a0e13").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e10").Value) Then
         adoaccrpt303.Fields("r30308").Value = Null
      Else
         adoaccrpt303.Fields("r30308").Value = adoacc0e0.Fields("a0e10").Value
      End If
      If IsNull(adoacc0e0.Fields("A0E14").Value) = False Then
         adoaccrpt303.Fields("R30309").Value = adoacc0e0.Fields("A0E14").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e06").Value) Then
         adoaccrpt303.Fields("r30310").Value = Null
      Else
         Select Case adoacc0e0.Fields("A0E05").Value
            Case "1"
               adoaccrpt303.Fields("r30310").Value = CustomerQuery(adoacc0e0.Fields("A0E06").Value, 1)
            Case "2"
               adoaccrpt303.Fields("r30310").Value = A0i02Query(adoacc0e0.Fields("A0E06").Value)
            Case "3"
               adoaccrpt303.Fields("r30310").Value = StaffQuery(adoacc0e0.Fields("A0E06").Value)
         End Select
      End If
      adoquery.CursorLocation = adUseClient
      adoquery.Open "SELECT A0G02 FROM ACC0G0 WHERE A0G01 = '" & adoacc0e0.Fields("A0E01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) = False Then
            adoaccrpt303.Fields("r30311").Value = adoquery.Fields(0).Value
         End If
      End If
      adoquery.Close
      If IsNull(adoacc0e0.Fields("a0e07").Value) Then
         adoaccrpt303.Fields("r30312").Value = Null
      Else
         adoaccrpt303.Fields("r30312").Value = adoacc0e0.Fields("a0e07").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e12").Value) Then
         adoaccrpt303.Fields("r30313").Value = Null
      Else
         adoaccrpt303.Fields("r30313").Value = adoacc0e0.Fields("a0e12").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e47").Value) Then
         adoaccrpt303.Fields("r30314").Value = Null
      Else
         adoaccrpt303.Fields("r30314").Value = adoacc0e0.Fields("a0e47").Value
      End If
      adoaccrpt303.UpdateBatch
      adoacc0e0.MoveNext
   Loop
   adoacc0e0.Close
   adoaccrpt303.Close
   'Modify by Amy 2022/04/06 +r30301 避免多人同時使用
   adoTaie.Execute "Delete From Accrpt303 Where r30301='" & strUserNum & "' And r30302 is null "
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Amy 2022/04/06 以Excel印
Private Function ExcelSave() As Boolean
    Dim Xls As New Excel.Application, Wks As New Worksheet
    Dim strWkName As String, strFileName As String, strFormat As String '工作表名稱為中文/檔案名稱/儲存格格式
    Dim strTmp As String
On Error GoTo ErrHnd

    intField = 65:  intR = 1
    strFileName = Val(Replace(MaskEdBox1.Text, "/", "")) & "-" & Val(Replace(MaskEdBox2.Text, "/", "")) & "託收票據資料表" & IIf(strCmp <> MsgText(601), strCmp & "公司", "") & ServerDate & MsgText(43)
    If Dir(strExcelPath & strFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileName
    End If
    
    Xls.SheetsInNewWorkbook = 3
    Xls.Workbooks.add
    'Xls.Visible = True
    '工作表名稱改為中文
    If strWkName = MsgText(601) Then strWkName = Left(Xls.Worksheets(1).Name, Len(Xls.Worksheets(1).Name) - 1)
    Set Wks = Xls.Worksheets(strWkName & "1")
    Call SetTitle(Wks)
    
    With Wks
        adoaccrpt303.MoveFirst
        Do While adoaccrpt303.EOF = False
            For i = LBound(strField) To UBound(strField)
                strFormat = ""
                Select Case strField(i)
                    Case "客戶名稱"
                        strTmp = PUB_StrToStr("" & adoaccrpt303.Fields("r30310"), 16)
                    Case "發票銀行"
                        strTmp = PUB_StrToStr("" & adoaccrpt303.Fields("r30302"), 12)
                    Case "發票帳號"
                        strTmp = "" & adoaccrpt303.Fields("r30312")
                        strFormat = "@"
                    Case "收票日期"
                        strTmp = "" & adoaccrpt303.Fields("r30307")
                        If strTmp <> MsgText(601) Then
                            strTmp = "" & CFDate(strTmp)
                        End If
                    Case "到期日期"
                        strTmp = "" & adoaccrpt303.Fields("r30308")
                        If strTmp <> MsgText(601) Then
                            strTmp = "" & CFDate(strTmp)
                        End If
                    Case "票據號碼"
                        strTmp = "" & adoaccrpt303.Fields("r30304")
                        strFormat = "@"
                    Case "票據金額"
                        strTmp = "" & adoaccrpt303.Fields("r30305")
                        strFormat = "#,##0"
                End Select
                If strFormat <> MsgText(601) Then
                    .Range(Chr(intField + i) & intR).NumberFormatLocal = strFormat
                    .Range(Chr(intField + i) & intR).HorizontalAlignment = xlRight
                End If
                If strField(i) = "發票帳號" Or strField(i) = "票據號碼" Then
                    .Range(Chr(intField + i) & intR).HorizontalAlignment = xlLeft
                End If
                .Range(Chr(intField + i) & intR).Font.Name = "細明體"
                .Range(Chr(intField + i) & intR).Value = strTmp
                
            Next i
            intR = intR + 1
            adoaccrpt303.MoveNext
        Loop
        .Range(Chr(intField + UBound(strField) - 1) & intR).Value = "合計:"
        .Range(Chr(intField + UBound(strField) - 1) & intR).Font.Name = "標楷體"
        .Range(Chr(intField + UBound(strField) - 1) & intR).Font.Bold = True
        .Range(Chr(intField + UBound(strField)) & intR).Value = "=Sum(" & Chr(intField + UBound(strField)) & intTitleRow + 1 & ":" & Chr(intField + UBound(strField)) & intR - 1 & ")"
        .Range(Chr(intField + UBound(strField)) & intR).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(Chr(intField + UBound(strField)) & intR).Borders(xlEdgeTop).Weight = xlThin  '細線
        .Range(Chr(intField + UBound(strField)) & intR).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range(Chr(intField + UBound(strField)) & intR).Borders(xlEdgeBottom).Weight = xlThick
        intR = intR + 1
        .Range(Chr(intField + UBound(strField)) & intR).Value = "***結束***"
        .Range(Chr(intField + UBound(strField)) & intR).Font.Size = 10
        .Range(Chr(intField + UBound(strField)) & intR).Font.Name = "標楷體"
        .Range(Chr(intField + UBound(strField)) & intR).Font.Bold = True
    End With
    
    Wks.PageSetup.PaperSize = 9 'A4
    Wks.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
    Wks.PageSetup.Orientation = xlPortrait '直印
    Wks.PageSetup.PrintTitleRows = "$1:$" & intTitleRow '標題列
    Wks.PageSetup.LeftMargin = Xls.InchesToPoints(0.5) '邊界
    Wks.PageSetup.RightMargin = Xls.InchesToPoints(0.5)
    Wks.PageSetup.TopMargin = Xls.InchesToPoints(0.5)
    Wks.PageSetup.BottomMargin = Xls.InchesToPoints(0.5)
    Wks.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
    
    '判斷版本
    If Val(Xls.Version) < 12 Then
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    'Modify by Amy 2023/11/08 原:Copies:=1 印1份
    Wks.PrintOut Copies:=2, Collate:=True
    Xls.Workbooks.Close
    Xls.Quit
    Set Wks = Nothing
    Set Xls = Nothing
    ExcelSave = True
    Exit Function
  
ErrHnd:
    If Val(Xls.Version) < 12 Then
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    Xls.Workbooks.Close
    Xls.Quit
    Set Wks = Nothing
    Set Xls = Nothing
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub SetTitle(ByRef Wks As Worksheet)
    ReDim stField(6): ReDim intWidth(6)
    strField = Array("客戶名稱", "發票銀行", "發票帳號", "收票日期", "到期日期", "票據號碼", "票據金額")
    intWidth = Array(18, 14, 12, 10, 10, 12, 9, 13)
    
    With Wks
        .Range(Chr(intField) & intR).Font.Size = 18
        .Range(Chr(intField) & intR).Font.Name = "標楷體"
        .Range(Chr(intField) & intR).Font.Bold = True
        .Range(Chr(intField) & intR).Value = "託收票據資料表"
        .Range(Chr(intField) & intR & ":" & Chr(UBound(stField) + intField) & intR).HorizontalAlignment = xlCenter
        .Range(Chr(intField) & intR & ":" & Chr(UBound(stField) + intField) & intR).MergeCells = True
        intR = intR + 1
        .Range(Chr(intField + 1) & intR).Value = "公司別："
        .Range(Chr(intField + 1) & intR).HorizontalAlignment = xlRight
        .Range(Chr(intField + 2) & intR).Value = strCmpN
        .Range(Chr(intField + 2) & intR).HorizontalAlignment = xlLeft
        intR = intR + 1
        .Range(Chr(intField + 1) & intR).Value = "銀行帳號："
        .Range(Chr(intField + 1) & intR).HorizontalAlignment = xlRight
        .Range(Chr(intField + 2) & intR).Value = Combo1 & "~" & Combo2
        .Range(Chr(intField + 2) & intR).HorizontalAlignment = xlLeft
        intR = intR + 1
        .Range(Chr(intField + 1) & intR).Value = "託收日期："
        .Range(Chr(intField + 1) & intR).HorizontalAlignment = xlRight
        .Range(Chr(intField + 2) & intR).Value = IIf(FCDate(MaskEdBox1.Text) = "", "", MaskEdBox1.Text) & "~" & IIf(FCDate(MaskEdBox2.Text) = "", "", MaskEdBox2.Text)
        .Range(Chr(intField + 2) & intR).HorizontalAlignment = xlLeft
        intR = intR + 1
        .Range(Chr(intField + 1) & intR).Value = "是否包含已兌現(Y/N)：" & Text3
        .Range(Chr(intField + 1) & intR & ":" & Chr(intField + 2) & intR).MergeCells = True
        .Range(Chr(intField + 1) & intR).HorizontalAlignment = xlLeft
        intR = intR + 1
        .Range(Chr(intField) & intR).Value = "列印人員：" & StaffQuery(strUserNum)
        .Range(Chr(intField + UBound(stField) - 1) & intR).Value = "列印日期：" & CFDate(strSrvDate(2))
        
        intR = intR + 2
        For i = LBound(strField) To UBound(strField)
            .Range(Chr(intField + i) & intR).Value = strField(i)
            .Columns(Chr(intField + i) & ":" & Chr(intField + i)).ColumnWidth = intWidth(i)
        Next i
        '設定格式
        .Range(Chr(intField) & "2:" & Chr(UBound(stField) + intField) & intR).Font.Size = 12
        .Range(Chr(intField) & "2:" & Chr(UBound(stField) + intField) & intR).Font.Name = "標楷體"
        .Range(Chr(intField) & "2:" & Chr(UBound(stField) + intField) & intR).Font.Bold = True
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(stField)) & intR).HorizontalAlignment = xlCenter
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(stField)) & intR).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(Chr(intField) & intR & ":" & Chr(intField + UBound(stField)) & intR).Borders(xlEdgeBottom).Weight = xlThin
        intTitleRow = intR
        intR = intR + 1
    End With
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
    
    For jj = 1 To UBound(strField)
        If UCase(strField(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function
'end 2022/04/06

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt303Delete()
   'Modify by Amy 2022/04/06 +r30301 避免多人同時使用
   adoTaie.Execute "Delete From Accrpt303 Where r30301='" & strUserNum & "' "
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Combo1 = ""
   Combo2 = ""
   'MaskEdBox1.Mask = ""
   'MaskEdBox1.Text = ""
   'MaskEdBox1.Mask = DFormat
   'MaskEdBox2.Mask = ""
   'MaskEdBox2.Text = ""
   'MaskEdBox2.Mask = DFormat
   Combo3 = ""
   Combo5 = ""
   Combo7 = ""
   Combo9 = ""
   Combo11 = ""
   '20140120START Modify By eric
'   Label13 = ""
'   Text1 = ""
'   Text1.SetFocus
   CboCmp.ListIndex = -1 'Add By Sindy 2020/4/23
   'Combo1.SetFocus
   '20140120END
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Combo1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Combo2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'Mark by Sindy 2020/4/23 公司別改下拉式選單
''20140120START ADD By eric
'Private Sub Text1_LostFocus()
'   If Text1.Text = "" Then
'      MsgBox "公司別不可空白 !"
'      Text1.SetFocus
'      Exit Sub
'   End If
'   If Text1.Text <> "1" And Text1.Text <> "2" Then
'      MsgBox "公司別僅能為 1 或 2 !"
'      Text1.Text = ""
'      Text1.SetFocus
'      Exit Sub
'   End If
'End Sub
''20140120END
'
''20140120START ADD By eric
'Private Sub Text1_GotFocus()
'   TextInverse Text1
'   CloseIme
'End Sub
''20140120END
'
''20140120START ADD By eric
'Private Sub Text1_Change()
'   Label13.Caption = A0802Query(IIf(Text1 = "2", "J", "1"))
'End Sub
''20140120END
