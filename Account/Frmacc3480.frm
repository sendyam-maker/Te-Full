VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc3480 
   AutoRedraw      =   -1  'True
   Caption         =   "往來對象別票據明細表"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   5160
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   30
      Width           =   3525
   End
   Begin VB.TextBox Text4 
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
      Left            =   1320
      TabIndex        =   6
      Top             =   1490
      Width           =   612
   End
   Begin VB.TextBox Text1 
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
      Left            =   1320
      TabIndex        =   1
      Top             =   410
      Width           =   612
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
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   770
      Width           =   1572
   End
   Begin VB.TextBox Text3 
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
      Left            =   3240
      TabIndex        =   3
      Top             =   770
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   1950
      Width           =   4692
   End
   Begin VB.ComboBox Combo12 
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
      Left            =   3600
      TabIndex        =   16
      Top             =   4350
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo11 
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
      Left            =   960
      TabIndex        =   15
      Top             =   4350
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo10 
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
      Left            =   3600
      TabIndex        =   14
      Top             =   3990
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo9 
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
      Left            =   960
      TabIndex        =   13
      Top             =   3990
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo8 
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
      Left            =   3600
      TabIndex        =   12
      Top             =   3610
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo7 
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
      Left            =   960
      TabIndex        =   11
      Top             =   3610
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo6 
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
      Left            =   3600
      TabIndex        =   10
      Top             =   3250
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo5 
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
      Left            =   960
      TabIndex        =   9
      Top             =   3250
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   3600
      TabIndex        =   8
      Top             =   2890
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   960
      TabIndex        =   7
      Top             =   2890
      Visible         =   0   'False
      Width           =   1812
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   1130
      Width           =   1572
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3240
      TabIndex        =   5
      Top             =   1130
      Width           =   1572
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
   Begin VB.Label Label15 
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
      TabIndex        =   32
      Top             =   100
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "(1.應收 2.應付 3.全部)"
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
      Left            =   2040
      TabIndex        =   31
      Top             =   1490
      Width           =   2652
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "票據別"
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
      TabIndex        =   30
      Top             =   1490
      Width           =   972
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   29
      Top             =   1130
      Width           =   252
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "往來日期"
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
      TabIndex        =   28
      Top             =   1130
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "往來類別"
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
      TabIndex        =   27
      Top             =   410
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "(1.客戶 2.廠商 3.員工)"
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
      Left            =   2040
      TabIndex        =   26
      Top             =   410
      Width           =   2652
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "往來對象"
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
      TabIndex        =   25
      Top             =   770
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   24
      Top             =   770
      Width           =   252
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2415
      Left            =   240
      Top             =   2430
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3480.frx":0000
      Stretch         =   -1  'True
      Top             =   4350
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "5."
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
      Left            =   720
      TabIndex        =   23
      Top             =   4350
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3480.frx":0442
      Stretch         =   -1  'True
      Top             =   3990
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "4."
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
      Left            =   720
      TabIndex        =   22
      Top             =   3990
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3480.frx":0884
      Stretch         =   -1  'True
      Top             =   3610
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "3."
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
      Left            =   720
      TabIndex        =   21
      Top             =   3610
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3480.frx":0CC6
      Stretch         =   -1  'True
      Top             =   3250
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "2."
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
      Left            =   720
      TabIndex        =   20
      Top             =   3250
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc3480.frx":1108
      Stretch         =   -1  'True
      Top             =   2890
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "1."
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
      Left            =   720
      TabIndex        =   19
      Top             =   2890
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "排序方式"
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
      Top             =   2550
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc3480"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoacc0e0 As New ADODB.Recordset
Public adoaccrpt309 As New ADODB.Recordset
Public adocust As New ADODB.Recordset
Dim strSort1 As String
Dim strSort2 As String
Dim strSort3 As String
Dim strSort4 As String
Dim strSort5 As String
Dim strCustNo As String
Dim strCustName As String
Dim dllaccrpt309 As Object
Dim r3912 As String     '20140220test
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/04/17


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
        MsgBox Label15 & MsgText(63), , MsgText(5)
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
         Text1.SetFocus
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
'20140120START Modify By eric
Private Sub Command1_Click()
Dim bCancel As Boolean
   
   'Add By Sindy 2020/4/23
   If CboCmp.Text = MsgText(601) Then
      MsgBox Label15 & MsgText(52), , MsgText(5)
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
   Accrpt309Delete
   ProduceData

   r3912 = ""
   adoaccrpt309.CursorLocation = adUseClient
   adoaccrpt309.Open "select distinct * from accrpt309 where r30912 is not null order by r30912 ", adoTaie, adOpenStatic, adLockReadOnly
   
   Do While adoaccrpt309.EOF = False
      If adoaccrpt309.RecordCount <> 0 And r3912 <> adoaccrpt309.Fields("r30912").Value Then
         r3912 = adoaccrpt309.Fields("r30912").Value
         RunReportDll
      End If
      adoaccrpt309.MoveNext
   Loop
   adoaccrpt309.Close

   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)

End Sub
'Private Sub Command1_Click()
'Dim strSql As String
'
'   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
'      Exit Sub
'   End If
'   Screen.MousePointer = vbHourglass
'   Accrpt309Delete
'   ProduceData
'   adocust.CursorLocation = adUseClient
'   strSql = ""
'   Select Case Text1
'      Case Mid(ComboItem(131), 1, 1)
'         If Text2 <> MsgText(601) Then
'            strSql = strSql & " and cu01 >= '" & Mid(Text2, 1, 8) & "'"
'         End If
'         If Text3 <> MsgText(601) Then
'            strSql = strSql & " and cu01 <= '" & Mid(Text3, 1, 8) & "'"
'         End If
'         If strSql <> "" Then
'            strSql = " where " & Mid(strSql, 5, Len(strSql) - 4)
'         End If
'         'adocust.Open "select cu01, max(cu02) from customer" & strSql & " group by cu01 order by cu01 asc", adoTaie, adOpenStatic, adLockReadOnly
'         adocust.Open "select cu01, min(cu02) from customer" & strSql & " group by cu01 order by cu01 asc", adoTaie, adOpenStatic, adLockReadOnly
'      Case Mid(ComboItem(132), 1, 1)
'         If Text2 <> MsgText(601) Then
'            strSql = strSql & " and a0i01 >= '" & Text2 & "'"
'         End If
'         If Text3 <> MsgText(601) Then
'            strSql = strSql & " and a0i01 <= '" & Text3 & "'"
'         End If
'         If strSql <> "" Then
'            strSql = " where " & Mid(strSql, 5, Len(strSql) - 4)
'         End If
'         adocust.Open "select * from acc0i0" & strSql & " order by a0i01 asc", adoTaie, adOpenStatic, adLockReadOnly
'      Case Mid(ComboItem(133), 1, 1)
'         If Text2 <> MsgText(601) Then
'            strSql = strSql & " and st01 >= '" & Text2 & "'"
'         End If
'         If Text3 <> MsgText(601) Then
'            strSql = strSql & " and st01 <= '" & Text3 & "'"
'         End If
'         If strSql <> "" Then
'            strSql = " where " & Mid(strSql, 5, Len(strSql) - 4)
'         End If
'         adocust.Open "select * from staff" & strSql & " order by st01 asc", adoTaie, adOpenStatic, adLockReadOnly
'      Case Else
'         MsgBox MsgText(161), , MsgText(5)
'         Screen.MousePointer = vbDefault
'         Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
'         Exit Sub
'   End Select
'   Do While adocust.EOF = False
'      adoaccrpt309.CursorLocation = adUseClient
'      Select Case Text1
'         Case Mid(ComboItem(131), 1, 1)
'            adoaccrpt309.Open "select * from accrpt309 where r30912 = '" & adocust.Fields("cu01").Value & adocust.Fields(1).Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'         Case Mid(ComboItem(132), 1, 1)
'            adoaccrpt309.Open "select * from accrpt309 where r30912 = '" & adocust.Fields("a0i01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'         Case Mid(ComboItem(133), 1, 1)
'            adoaccrpt309.Open "select * from accrpt309 where r30912 = '" & adocust.Fields("st01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      End Select
'      If adoaccrpt309.RecordCount <> 0 Then
'         RunReportDll
'      End If
'      adoaccrpt309.Close
'      adocust.MoveNext
'   Loop
'   adocust.Close
'   Screen.MousePointer = vbDefault
'   FormClear
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
'End Sub
'20140120END

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   Me.Height = 2900
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
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
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt309 = CreateObject("AccReport.ReportSelect")
   
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
   Set dllaccrpt309 = Nothing
   Set Frmacc3480 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text1 = Mid(ComboItem(131), 1, 1) Then
      If Len(Text2) = 6 Then
         Text2 = AfterZero(Text2)
      End If
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text1 = Mid(ComboItem(131), 1, 1) Then
      If Len(Text3) = 6 Then
         Text3 = AfterZero(Text3)
      End If
   End If
End Sub

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "銀行代號"
   strSort2 = "票據號碼"
   strSort3 = "票別"
   strSort4 = "開票日期"
   strSort5 = "到期日期"
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
Dim strOrder1 As String
Dim strOrder2 As String
Dim strOrder3 As String
Dim strOrder4 As String
Dim strOrder5 As String
Dim strSql As String
   
On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e01 asc"
         Else
            strOrder1 = " order by a0e01 desc"
         End If
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e02 asc"
         Else
            strOrder1 = " order by a0e02 desc"
         End If
      Case strSort3
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e08 asc"
         Else
            strOrder1 = " order by a0e08 desc"
         End If
      Case strSort4
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e13 asc"
         Else
            strOrder1 = " order by a0e13 desc"
         End If
      Case strSort5
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0e10 asc"
         Else
            strOrder1 = " order by a0e10 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   Select Case Combo5
      Case strSort1
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e01 asc"
         Else
            strOrder2 = ", a0e01 desc"
         End If
      Case strSort2
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e02 asc"
         Else
            strOrder2 = ", a0e02 desc"
         End If
      Case strSort3
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e08 asc"
         Else
            strOrder2 = ", a0e08 desc"
         End If
      Case strSort4
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e13 asc"
         Else
            strOrder2 = ", a0e13 desc"
         End If
      Case strSort5
         If Combo6 = MsgText(1) Then
            strOrder2 = ", a0e10 asc"
         Else
            strOrder2 = ", a0e10 desc"
         End If
      Case Else
         strOrder2 = MsgText(601)
   End Select
   Select Case Combo7
      Case strSort1
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e01 asc"
         Else
            strOrder3 = ", a0e01 desc"
         End If
      Case strSort2
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e02 asc"
         Else
            strOrder3 = ", a0e02 desc"
         End If
      Case strSort3
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e08 asc"
         Else
            strOrder3 = ", a0e08 desc"
         End If
      Case strSort4
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e13 asc"
         Else
            strOrder3 = ", a0e13 desc"
         End If
      Case strSort5
         If Combo8 = MsgText(1) Then
            strOrder3 = ", a0e10 asc"
         Else
            strOrder3 = ", a0e10 desc"
         End If
      Case Else
         strOrder3 = MsgText(601)
   End Select
   Select Case Combo9
      Case strSort1
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e01 asc"
         Else
            strOrder4 = ", a0e01 desc"
         End If
      Case strSort2
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e02 asc"
         Else
            strOrder4 = ", a0e02 desc"
         End If
      Case strSort3
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e08 asc"
         Else
            strOrder4 = ", a0e08 desc"
         End If
      Case strSort4
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e13 asc"
         Else
            strOrder4 = ", a0e13 desc"
         End If
      Case strSort5
         If Combo10 = MsgText(1) Then
            strOrder4 = ", a0e10 asc"
         Else
            strOrder4 = ", a0e10 desc"
         End If
      Case Else
         strOrder4 = MsgText(601)
   End Select
   Select Case Combo11
      Case strSort1
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e01 asc"
         Else
            strOrder5 = ", a0e01 desc"
         End If
      Case strSort2
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e02 asc"
         Else
            strOrder5 = ", a0e02 desc"
         End If
      Case strSort3
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e08 asc"
         Else
            strOrder5 = ", a0e08 desc"
         End If
      Case strSort4
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e13 asc"
         Else
            strOrder5 = ", a0e13 desc"
         End If
      Case strSort5
         If Combo12 = MsgText(1) Then
            strOrder5 = ", a0e10 asc"
         Else
            strOrder5 = ", a0e10 desc"
         End If
      Case Else
         strOrder5 = MsgText(601)
   End Select
   If Text1 <> MsgText(601) Then
      strSql = " and a0e05 = '" & Text1 & "'"
   End If
   
   '20140120START Add By eric
   'Modify By Sindy 2020/4/23
'   If Text5 <> MsgText(601) Then
'      strSql = strSql & " and a0e23 = '" & IIf(Text5 = "2", "J", "1") & "'"
'   End If
   If CboCmp.Text <> MsgText(601) Then
      strSql = strSql & " and a0e23 = '" & strCmp & "'"
   End If
   '2020/4/23 END
   '20140120END
   
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0e06 >= '" & Text2 & "'"
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0e06 <= '" & Text3 & "'"
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0e13 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0e13 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text4 <> MsgText(601) Then
      Select Case Text4
         Case "1"
            strSql = strSql & " and a0e04 = '" & MsgText(18) & "'"
         Case "2"
            strSql = strSql & " and a0e04 = '" & MsgText(19) & "'"
      End Select
   End If
   
   
   adoaccrpt309.CursorLocation = adUseClient
   adoaccrpt309.Open "select * from accrpt309", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0e0.CursorLocation = adUseClient
   adoacc0e0.Open "select * from acc0e0 where a0e25 = 0 and a0e15 = 0" & strSql & strOrder1 & strOrder2 & strOrder3 & strOrder4 & strOrder5, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0e0.RecordCount = 0 Then
      adoacc0e0.Close
      adoaccrpt309.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0e0.EOF = False
      adoaccrpt309.AddNew
      adoaccrpt309.Fields("r30901").Value = strUserNum
      Select Case adoacc0e0.Fields("a0e04").Value
         Case MsgText(18)
            adoaccrpt309.Fields("r30902").Value = adoacc0e0.Fields("a0e01").Value
            If IsNull(adoacc0e0.Fields("a0e07").Value) Then
               adoaccrpt309.Fields("r30903").Value = Null
            Else
               adoaccrpt309.Fields("r30903").Value = adoacc0e0.Fields("a0e07").Value
               adoaccrpt309.Fields("r30904").Value = A0g02Query(adoaccrpt309.Fields("r30902").Value)
            End If
            If IsNull(adoacc0e0.Fields("a0e11").Value) Then
               adoaccrpt309.Fields("r30907").Value = 0
            Else
               adoaccrpt309.Fields("r30907").Value = Val(adoacc0e0.Fields("a0e11").Value)
            End If
            If IsNull(adoacc0e0.Fields("A0E21").Value) = False Then
               adoaccrpt309.Fields("R30910").Value = adoacc0e0.Fields("A0E21").Value
            End If
         Case MsgText(19)
            adoaccrpt309.Fields("r30902").Value = adoacc0e0.Fields("a0e01").Value
            If IsNull(adoacc0e0.Fields("a0e07").Value) Then
               adoaccrpt309.Fields("r30903").Value = Null
            Else
               adoaccrpt309.Fields("r30903").Value = adoacc0e0.Fields("a0e07").Value
               adoaccrpt309.Fields("r30904").Value = A0g02Query(adoaccrpt309.Fields("r30902").Value)
            End If
            If IsNull(adoacc0e0.Fields("a0e11").Value) Then
               adoaccrpt309.Fields("r30913").Value = 0
            Else
               adoaccrpt309.Fields("r30913").Value = Val(adoacc0e0.Fields("a0e11").Value)
            End If
            If IsNull(adoacc0e0.Fields("A0E37").Value) = False Then
               adoaccrpt309.Fields("R30910").Value = adoacc0e0.Fields("A0E37").Value
            End If
         Case Else
            adoaccrpt309.Fields("r30902").Value = Null
            adoaccrpt309.Fields("r30903").Value = Null
      End Select
      adoaccrpt309.Fields("r30905").Value = adoacc0e0.Fields("a0e02").Value
      Select Case adoacc0e0.Fields("a0e08").Value
         Case Mid(ComboItem(11), 1, 1)
            adoaccrpt309.Fields("r30906").Value = Mid(ComboItem(11), 4, 2)
         Case Mid(ComboItem(12), 1, 1)
            adoaccrpt309.Fields("r30906").Value = Mid(ComboItem(12), 4, 2)
         Case Mid(ComboItem(13), 1, 1)
            adoaccrpt309.Fields("r30906").Value = Mid(ComboItem(13), 4, 2)
         Case Else
            adoaccrpt309.Fields("r30906").Value = Null
      End Select
      If IsNull(adoacc0e0.Fields("a0e13").Value) Then
         adoaccrpt309.Fields("r30908").Value = Null
      Else
         adoaccrpt309.Fields("r30908").Value = adoacc0e0.Fields("a0e13").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e10").Value) Then
         adoaccrpt309.Fields("r30909").Value = Null
      Else
         adoaccrpt309.Fields("r30909").Value = adoacc0e0.Fields("a0e10").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e05").Value) Then
         adoaccrpt309.Fields("r30911").Value = Null
      Else
         adoaccrpt309.Fields("r30911").Value = adoacc0e0.Fields("a0e05").Value
      End If
      If IsNull(adoacc0e0.Fields("a0e06").Value) Then
         adoaccrpt309.Fields("r30912").Value = Null
      Else
         adoaccrpt309.Fields("r30912").Value = adoacc0e0.Fields("a0e06").Value
      End If
      adoaccrpt309.UpdateBatch
      adoacc0e0.MoveNext
   Loop
   adoacc0e0.Close
   adoaccrpt309.Close
'   adoTaie.Execute "delete from accrpt309 where r30902 is null"
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt309Delete()
   adoTaie.Execute "delete from accrpt309"
End Sub

'*************************************************
'  執行報表之 Dll
'
'*************************************************
Private Sub RunReportDll()
Dim strSelect As String
Dim strNote As String
Dim strCustName As String
   
   '20140120START Modify By eric
   Select Case Text1
      Case Mid(ComboItem(131), 1, 1)
         strSelect = Mid(ComboItem(131), 4, 2)
         strCustName = CustomerQuery(Mid(adoaccrpt309.Fields("r30912").Value, 1, 8), 1)
      Case Mid(ComboItem(132), 1, 1)
         strSelect = Mid(ComboItem(132), 4, 2)
         strCustName = A0i02Query(Mid(adoaccrpt309.Fields("r30912").Value, 1, 5))
      Case Mid(ComboItem(133), 1, 1)
         strSelect = Mid(ComboItem(133), 4, 2)
         strCustName = StaffQuery(Mid(adoaccrpt309.Fields("r30912").Value, 1, 8))
   End Select
   'Select Case Text1
   '   Case Mid(ComboItem(131), 1, 1)
   '      strSelect = Mid(ComboItem(131), 4, 2)
   '      strCustName = CustomerQuery(adocust.Fields("cu01").Value, 1)
   '   Case Mid(ComboItem(132), 1, 1)
   '      strSelect = Mid(ComboItem(132), 4, 2)
   '      strCustName = A0i02Query(adocust.Fields("a0i01").Value)
   '   Case Mid(ComboItem(133), 1, 1)
   '      strSelect = Mid(ComboItem(133), 4, 2)
   '      strCustName = StaffQuery(adocust.Fields("st01").Value)
   'End Select
   '20140120END
   Select Case Text4
      Case "1"
         strNote = MsgText(40)
      Case "2"
         strNote = MsgText(41)
      Case "3"
         strNote = MsgText(31)
   End Select
   
   '20140120START Modify By eric
   dllaccrpt309.Acc3480 ReportTitle(309) & "-" & strCmpN, strSelect, Text2.Text, Text3.Text, strNote, MaskEdBox1.Text, MaskEdBox2.Text, r3912, strCustName, strUserNum, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   'Select Case Text1
   '   Case Mid(ComboItem(131), 1, 1)
   '      dllaccrpt309.Acc3480 ReportTitle(309), strSelect, Text2.Text, Text3.Text, strNote, MaskEdBox1.Text, MaskEdBox2.Text, adocust.Fields("cu01").Value & adocust.Fields(1).Value, strCustName, strUserNum, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   '   Case Mid(ComboItem(132), 1, 1)
   '      dllaccrpt309.Acc3480 ReportTitle(309), strSelect, Text2.Text, Text3.Text, strNote, MaskEdBox1.Text, MaskEdBox2.Text, adocust.Fields("a0i01").Value, strCustName, strUserNum, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   '   Case Mid(ComboItem(133), 1, 1)
   '      dllaccrpt309.Acc3480 ReportTitle(309), strSelect, Text2.Text, Text3.Text, strNote, MaskEdBox1.Text, MaskEdBox2.Text, adocust.Fields("st01").Value, strCustName, strUserNum, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   'End Select
   '20140120END
   
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   Text3 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text4 = ""
   Combo3 = ""
   Combo5 = ""
   Combo7 = ""
   Combo9 = ""
   Combo11 = ""
   '20140120Modify By eric
'   Label16 = ""
'   Text5 = ""
'   Text5.SetFocus
   CboCmp.ListIndex = -1 'Add By Sindy 2020/4/23
   'Text1.SetFocus
   '20140120END
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text3 <> MsgText(601) Then
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
''20140120START By eric
'Private Sub Text5_LostFocus()
'   If Text5.Text = "" Then
'      MsgBox "公司別不可空白 !"
'      Text5.SetFocus
'      Exit Sub
'   End If
'   If Text5.Text <> "1" And Text5.Text <> "2" Then
'      MsgBox "公司別僅能為 1 或 2 !"
'      Text5.Text = ""
'      Text5.SetFocus
'      Exit Sub
'   End If
'End Sub
''20140120END
'
''20140120START By eric
'Private Sub Text5_GotFocus()
'   TextInverse Text5
'   CloseIme
'End Sub
''20140120END
'
''20140120START By eric
'Private Sub Text5_Change()
'   Label16.Caption = A0802Query(IIf(Text5 = "2", "J", "1"))
'End Sub
''20140120END
