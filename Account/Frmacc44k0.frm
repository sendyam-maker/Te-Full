VERSION 5.00
Begin VB.Form Frmacc44k0 
   AutoRedraw      =   -1  'True
   Caption         =   "會計科目代號對照表"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   5160
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
      Width           =   1572
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
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
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
      TabIndex        =   4
      Top             =   240
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
      TabIndex        =   3
      Top             =   240
      Width           =   252
   End
End
Attribute VB_Name = "Frmacc44k0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit
Public adoacc010 As New ADODB.Recordset
Public adoacc070 As New ADODB.Recordset
Public adoaccrpt402 As New ADODB.Recordset
Dim dllaccrpt402 As Object

Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt402Delete
   ProduceData
   dllaccrpt402.Acc44k0 ReportTitle(402), Text2, Text1, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
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
   Me.Height = 1800
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
   Set dllaccrpt402 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt402 = Nothing
   Set Frmacc44k0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strSql As String

On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt402.CursorLocation = adUseClient
   adoaccrpt402.Open "select * from accrpt402", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc010.CursorLocation = adUseClient
   If Text2 <> MsgText(601) Then
      strSql = " and a0101 >= '" & Text2 & "'"
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0101 <= '" & Text1 & "'"
   End If
   If strSql <> MsgText(601) Then
      strSql = Mid(strSql, 5, Len(strSql) - 4)
   Else
      adoaccrpt402.Close
      Exit Sub
   End If
   adoacc010.Open "select * from acc010 where" & strSql & " order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc010.RecordCount = 0 Then
      adoacc010.Close
      adoaccrpt402.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc010.EOF = False
      adoaccrpt402.AddNew
      adoaccrpt402.Fields("r40201").Value = strUserNum
      adoaccrpt402.Fields("r40202").Value = adoacc010.Fields("a0101").Value
      If IsNull(adoacc010.Fields("a0102").Value) Then
         adoaccrpt402.Fields("r40203").Value = Null
      Else
         adoaccrpt402.Fields("r40203").Value = adoacc010.Fields("a0102").Value
      End If
      If IsNull(adoacc010.Fields("a0103").Value) Then
         adoaccrpt402.Fields("r40204").Value = Null
      Else
         Select Case adoacc010.Fields("a0103").Value
            Case "1"
               adoaccrpt402.Fields("r40204").Value = Mid(ComboItem(1), 4, 1)
            Case "2"
               adoaccrpt402.Fields("r40204").Value = Mid(ComboItem(2), 4, 1)
         End Select
      End If
      If IsNull(adoacc010.Fields("a0104").Value) Then
         adoaccrpt402.Fields("r40205").Value = Null
      Else
         adoaccrpt402.Fields("r40205").Value = adoacc010.Fields("a0104").Value
      End If
      'Add by Amy 2014/02/06 +專用公司別
      If IsNull(adoacc010.Fields("a0109").Value) Then
        adoaccrpt402.Fields("r40207").Value = Null
      Else
        adoaccrpt402.Fields("r40207").Value = adoacc010.Fields("a0109").Value & "-" & Left(A0802Query(adoacc010.Fields("a0109").Value), 6)
      End If
      'end 2014/02/06
      adoacc070.CursorLocation = adUseClient
      adoacc070.Open "select * from acc070 where a0701 = '" & adoacc010.Fields("a0105").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc070.RecordCount <> 0 Then
         If IsNull(adoacc070.Fields("a0702").Value) Then
            adoaccrpt402.Fields("r40206").Value = Null
         Else
            adoaccrpt402.Fields("r40206").Value = adoacc070.Fields("a0702").Value
         End If
      Else
         adoaccrpt402.Fields("r40206").Value = Null
      End If
      adoacc070.Close
      adoaccrpt402.UpdateBatch
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt402.Close
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
Private Sub Accrpt402Delete()
   adoTaie.Execute "delete from accrpt402"
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text2 = ""
   Text1 = ""
   Text2.SetFocus
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
   FormCheck = False
End Function

