VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14e0 
   AutoRedraw      =   -1  'True
   Caption         =   "收據/請款單作廢明細表"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1440
   ScaleWidth      =   5130
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
      TabIndex        =   2
      Top             =   840
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
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
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "作廢銷帳日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Frmacc14e0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoquery As New ADODB.Recordset
Public adoaccrpt116 As New ADODB.Recordset
Dim dllaccrpt116 As Object

Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   ProcessData
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from accrpt116", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      dllaccrpt116.Acc14e0 ReportTitle(116), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   End If
   adoquery.Close
   FormClear
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
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
   Me.Height = 1850
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
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Set dllaccrpt116 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt116 = Nothing
   Set Frmacc14e0 = Nothing
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Public Sub ProcessData()
Dim strSQL1 As String
Dim strSQL2 As String
   
   adoTaie.Execute "delete from accrpt116"
   If MaskEdBox1.Text <> MsgText(29) Then
      strSQL1 = strSQL1 & " and a0k09 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSQL2 = strSQL2 & " and a0s03 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      strSQL1 = strSQL1 & " and a0k09 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSQL2 = strSQL2 & " and a0s03 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select a0902, st02, a0k01, (a0k06 + a0k07) as amount, a0k09 as cdate, 0 as ndate, a0k08 as note from acc0k0, staff, acc090 where a0k20 = st01 and st03 = a0901 and (a0k09 is not null and a0k09 <> 0)" & strSQL1 & " union " & _
                 "select a0902, st02, a0k01, (a0s05) as amount, 0 as cdate, a0s03 as ndate, a0s18 as note from acc0s0, acc0k0, staff, acc090 where a0s02 = a0k01 and a0k20 = st01 and st03 = a0901" & strSQL2 & " union " & _
                 "select a0902, st02, a0k01, (a0s06 + a0s07) as amount, 0 as cdate, a0s03 as ndate, a0s18 as note from acc0s0, acc0k0, staff, acc090 where a0s02 = a0k01 and a0k20 = st01 and st03 = a0901" & strSQL2, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount = 0 Then
      adoquery.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoquery.EOF = False
      With adoquery
         If adoquery.Fields("amount").Value <> 0 Then
            adoTaie.Execute "insert into accrpt116 values ('" & strUserNum & "', '" & .Fields("a0902").Value & "', '" & .Fields("st02").Value & "', '" & .Fields("a0k01").Value & "', " & .Fields("amount").Value & ", " & IIf(.Fields("cdate").Value = 0, "Null", .Fields("cdate").Value) & ", " & IIf(.Fields("ndate").Value = 0, "Null", .Fields("ndate").Value) & ", '" & .Fields("note").Value & "')"
         End If
      End With
      adoquery.MoveNext
   Loop
   adoquery.Close
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   MaskEdBox1.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
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

