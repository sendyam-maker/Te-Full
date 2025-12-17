VERSION 5.00
Begin VB.Form Frmacc44m0 
   AutoRedraw      =   -1  'True
   Caption         =   "費用科目分攤比率表"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2400
   ScaleWidth      =   5160
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   1800
      Width           =   3450
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
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   852
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
      Height          =   300
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   2652
   End
   Begin VB.TextBox Text5 
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
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text6 
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
      Left            =   3240
      TabIndex        =   2
      Top             =   600
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
      TabIndex        =   3
      Top             =   1200
      Width           =   4692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
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
      TabIndex        =   9
      Top             =   1830
      Width           =   975
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
      TabIndex        =   7
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      TabIndex        =   6
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label4 
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
      TabIndex        =   5
      Top             =   600
      Width           =   252
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc44m0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit
Public adoacc010 As New ADODB.Recordset
Public adoacc060 As New ADODB.Recordset
Public adoaccrpt419 As New ADODB.Recordset
Dim dllaccrpt419 As Object
Dim strPrinter As String 'Added by Lydia 2016/10/07 加印表機選項

Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   'Added by Lydia 2016/10/07 加印表機選項
   PUB_SetOsDefaultPrinter Combo1
   
   Accrpt419Delete
   ProduceData
   dllaccrpt419.Acc44m0 ReportTitle(419), Text2, Text1, Text5, Text6, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   
   'Added by Lydia 2016/10/07 加印表機選項
   PUB_SetOsDefaultPrinter strPrinter
   
   Screen.MousePointer = vbDefault
   FormClear
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
   'Modified by Lydia 2016/10/07
   'Me.Height = 2250
   Me.Height = 2800
   
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Set dllaccrpt419 = CreateObject("AccReport.ReportSelect")
   'Added by Lydia 2016/10/07 加印表機選項
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   'Added by Lydia 2016/10/07 程式結束後要還原為預設印表機
   PUB_SetOsDefaultPrinter strPrinter
   
   Set dllaccrpt419 = Nothing
   Set Frmacc44m0 = Nothing
End Sub

Private Sub Text2_Change()
   If Text2 = MsgText(601) Then
      Exit Sub
   End If
   Text1 = A0802Query(Text2)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
   CloseIme
End Sub

'2014/1/27 add by sonia
Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 = MsgText(601) Then
      MsgBox MsgText(10) & Label3, , MsgText(5)
      Cancel = True
      Text2.SetFocus
      Exit Sub
   Else
      If Text2 <> "1" And Text5 <> "J" Then
         MsgBox "只可輸入 1 或 J", vbCritical
         Cancel = True
         Text2.SetFocus
         Exit Sub
      End If
   End If
End Sub
'2014/1/27 end

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Text5.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Text6.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Text2.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strSql As String

On Error GoTo Checking
   If Text5 <> MsgText(601) Then
      strSql = " and a0101 >= '" & Text5 & "'"
   End If
   If Text6 <> MsgText(601) Then
      strSql = strSql & " and a0101 <= '" & Text6 & "'"
   End If
   If strSql <> MsgText(601) Then
      strSql = Mid(strSql, 5, Len(strSql) - 4)
   Else
      Exit Sub
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt419.CursorLocation = adUseClient
   adoaccrpt419.Open "select * from accrpt419", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc010.CursorLocation = adUseClient
   adoacc010.Open "select * from acc010 where" & strSql & " order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc010.RecordCount = 0 Then
      adoacc010.Close
      adoaccrpt419.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc010.EOF = False
      If IsNull(adoacc010.Fields("a0105").Value) = False Then
         adoacc060.CursorLocation = adUseClient
         adoacc060.Open "select * from acc060, acc070, acc090 where a0602 = a0701 (+) and a0604 = a0901 (+) and a0601 = " & Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) & " and a0602 = '" & adoacc010.Fields("a0105").Value & "' and a0603 = '" & Text2 & "' order by a0604 asc", adoTaie, adOpenStatic, adLockReadOnly
         Do While adoacc060.EOF = False
            adoaccrpt419.AddNew
            adoaccrpt419.Fields("r41901").Value = strUserNum
            adoaccrpt419.Fields("r41902").Value = adoacc010.Fields("a0101").Value
            If IsNull(adoacc010.Fields("a0102").Value) Then
               adoaccrpt419.Fields("r41903").Value = Null
            Else
               adoaccrpt419.Fields("r41903").Value = adoacc010.Fields("a0102").Value
            End If
            If IsNull(adoacc060.Fields("a0702").Value) Then
               adoaccrpt419.Fields("r41904").Value = Null
            Else
               adoaccrpt419.Fields("r41904").Value = adoacc060.Fields("a0702").Value
            End If
            If IsNull(adoacc060.Fields("a0902").Value) Then
               adoaccrpt419.Fields("r41095").Value = Null
            Else
               adoaccrpt419.Fields("r41905").Value = adoacc060.Fields("a0902").Value
            End If
            If IsNull(adoacc060.Fields("a0605").Value) Then
               adoaccrpt419.Fields("r41906").Value = 0
            Else
               adoaccrpt419.Fields("r41906").Value = Val(adoacc060.Fields("a0605").Value)
            End If
            adoaccrpt419.UpdateBatch
            adoacc060.MoveNext
         Loop
         adoacc060.Close
      End If
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt419.Close
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
Private Sub Accrpt419Delete()
   adoTaie.Execute "delete from accrpt419"
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text2 = ""
   Text1 = ""
   Text5 = ""
   Text6 = ""
   Text2.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text5 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text6 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

