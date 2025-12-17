VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc1490 
   AutoRedraw      =   -1  'True
   Caption         =   "銷帳退費明細表"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2415
   ScaleWidth      =   5160
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
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1350
      Width           =   495
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
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   4
      Top             =   960
      Width           =   495
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
      TabIndex        =   6
      Top             =   1740
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
      Left            =   3240
      TabIndex        =   3
      Top             =   600
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
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   240
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
      TabIndex        =   1
      Top             =   240
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
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "( 1.轉應付 2.轉暫收 )"
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
      Left            =   2025
      TabIndex        =   14
      Top             =   1350
      Width           =   2730
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "退費方式"
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
      TabIndex        =   13
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "( 1.銷 2.退 3.銷+退 )"
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
      Left            =   2025
      TabIndex        =   12
      Top             =   960
      Width           =   2730
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "類別"
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
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label7 
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
      TabIndex        =   10
      Top             =   600
      Width           =   252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
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
      TabIndex        =   9
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "銷退日期"
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
      TabIndex        =   8
      Top             =   240
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
      TabIndex        =   7
      Top             =   240
      Width           =   252
   End
End
Attribute VB_Name = "Frmacc1490"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit
Public adoacc0s0 As New ADODB.Recordset
Public adocaseprogress As New ADODB.Recordset
Public adoaccrpt109 As New ADODB.Recordset
Dim dllaccrpt109 As Object

Private Sub Command1_Click()
   Dim strCon1 As String, strCon2 As String 'Add by Morgan 2004/11/16
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt109Delete
   ProduceData
   adoaccrpt109.CursorLocation = adUseClient
   adoaccrpt109.Open "select * from accrpt109", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt109.RecordCount <> 0 Then
      'Modify by Morgan 2004/11/16 加類別，退費方式輸入條件
      'dllaccrpt109.Acc1490 ReportTitle(109), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      strCon1 = MaskEdBox1.Text & " - " & MaskEdBox2.Text
      strCon2 = ""
      Select Case Text3.Text
         Case "1"
            strCon2 = "銷帳"
         Case "2"
            strCon2 = "退費"
         Case "3"
            strCon2 = "銷帳+退費"
      End Select
      strCon2 = strCon2 & ","
      Select Case Text4.Text
         Case "1"
            strCon2 = strCon2 & "轉應付"
         Case "2"
            strCon2 = strCon2 & "轉暫收"
      End Select
      dllaccrpt109.Acc1490 ReportTitle(109), strCon1, strCon2, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   End If
   adoaccrpt109.Close
   Screen.MousePointer = vbDefault
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

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
   'Modify by Amy 2023/10/11 原H2600
   Me.Height = 2880
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text1 = "X"
   Text2 = "X"
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt109 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt109 = Nothing
   Set Frmacc1490 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Len(Text1) = 6 Then
      Text1 = AfterZero(Text1)
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strSql As String

On Error GoTo Checking
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0s03 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0s03 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0k03 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0k03 <= '" & Text2 & "'"
   End If
    '若非北所員工, 只能列印該所資料
    If pub_strUserOffice <> "1" Then
        strSql = strSql & " And ''||ST06='" & pub_strUserOffice & "' "
    End If
    
   If Text3.Text <> "" Then
      strSql = strSql & " and a0s04='" & Text3.Text & "'"
   End If
   If Text4.Text <> "" Then
      strSql = strSql & " and a0s08='" & Text4.Text & "'"
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt109.CursorLocation = adUseClient
   adoaccrpt109.Open "select * from accrpt109", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0s0.CursorLocation = adUseClient
   adoacc0s0.Open "select a0s03, a0k01 as No, a0k02 as IDate, a0k03, a0k04,  a0k23, a0s05, a0s06, a0s07, a0k20, a0s08, a0s02, a0s01, '' as No2 from acc0s0, acc0k0, Staff where a0s02 = a0k01(+) And A0K20=ST01(+) " & strSql & _
                  " union select a0s03, a0t01 as No, a0t03 as IDate, a0k03, a0k04,  a0k23, a0s05, a0s06, a0s07, a0k20, a0s08, a0s02, a0s01, a0k01 as No2 from acc0s0, acc0t0, acc0m0, acc0k0, Staff where a0s02 = a0t01(+) and a0t07 = a0m01(+) and a0m02 = a0k01(+) And A0K20=ST01(+) " & strSql & " order by a0s02 asc, a0s01 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0s0.RecordCount = 0 Then
      adoacc0s0.Close
      adoaccrpt109.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0s0.EOF = False
      adoaccrpt109.AddNew
      adoaccrpt109.Fields("r10901").Value = strUserNum
      If IsNull(adoacc0s0.Fields("a0s03").Value) Then
         adoaccrpt109.Fields("r10902").Value = Null
      Else
         adoaccrpt109.Fields("r10902").Value = adoacc0s0.Fields("a0s03").Value
      End If
      adoaccrpt109.Fields("r10903").Value = adoacc0s0.Fields("No").Value
      If IsNull(adoacc0s0.Fields("IDate").Value) Then
         adoaccrpt109.Fields("r10904").Value = Null
      Else
         adoaccrpt109.Fields("r10904").Value = adoacc0s0.Fields("IDate").Value
      End If
      If IsNull(adoacc0s0.Fields("a0k03").Value) Then
         adoaccrpt109.Fields("r10905").Value = Null
      Else
         adoaccrpt109.Fields("r10905").Value = adoacc0s0.Fields("a0k03").Value
      End If
      If IsNull(adoacc0s0.Fields("a0k04").Value) Then
         adoaccrpt109.Fields("r10906").Value = Null
      Else
         adoaccrpt109.Fields("r10906").Value = adoacc0s0.Fields("a0k04").Value
      End If
      adocaseprogress.CursorLocation = adUseClient
      'Modify by Morgan 2011/8/23 改從 0j0 抓 cp
      'Modified by Morgan 2011/12/30 取消 a0j21
      adocaseprogress.Open "select CP01, CP09, cp10, na03 from acc0j0,caseprogress,nation where cp09(+) = a0j01 and na01(+)=a0j04 and a0j13 = '" & IIf(Mid(adoacc0s0.Fields("a0s02").Value, 1, 1) = MsgText(802), adoacc0s0.Fields("No").Value, adoacc0s0.Fields("No2").Value) & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocaseprogress.RecordCount <> 0 Then
         adoaccrpt109.Fields("r10907").Value = CaseNameQuery(adocaseprogress.Fields("cp09").Value, 1)
         adoaccrpt109.Fields("r10908").Value = PropertyQuery(adocaseprogress.Fields("CP01").Value, adocaseprogress.Fields("cp10").Value)
         adoaccrpt109.Fields("r10909").Value = adocaseprogress.Fields("na03").Value
      Else
         adoaccrpt109.Fields("r10907").Value = Null
         adoaccrpt109.Fields("r10908").Value = Null
         adoaccrpt109.Fields("r10909").Value = Null
      End If
      adocaseprogress.Close
      If Mid(adoacc0s0.Fields("a0s02").Value, 1, 1) = MsgText(802) Then
         If IsNull(adoacc0s0.Fields("a0s05").Value) Then
            adoaccrpt109.Fields("r10910").Value = 0
         Else
            adoaccrpt109.Fields("r10910").Value = adoacc0s0.Fields("a0s05").Value
         End If
         If IsNull(adoacc0s0.Fields("a0s06").Value) Then
            adoaccrpt109.Fields("r10911").Value = 0
         Else
            adoaccrpt109.Fields("r10911").Value = adoacc0s0.Fields("a0s06").Value
            If IsNull(adoacc0s0.Fields("a0s07").Value) = False Then
               adoaccrpt109.Fields("r10911").Value = Val(adoaccrpt109.Fields("r10911").Value) + Val(adoacc0s0.Fields("a0s07").Value)
            End If
         End If
      Else
         adoaccrpt109.Fields("r10910").Value = 0
         If IsNull(adoacc0s0.Fields("a0s05").Value) Then
            adoaccrpt109.Fields("r10911").Value = 0
         Else
            adoaccrpt109.Fields("r10911").Value = adoacc0s0.Fields("a0s05").Value
         End If
      End If
      If IsNull(adoacc0s0.Fields("a0k20").Value) Then
         adoaccrpt109.Fields("r10912").Value = Null
      Else
         adoaccrpt109.Fields("r10912").Value = StaffQuery(adoacc0s0.Fields("a0k20").Value)
      End If
      If IsNull(adoacc0s0.Fields("a0s08").Value) Then
         adoaccrpt109.Fields("r10913").Value = Null
      Else
         Select Case adoacc0s0.Fields("a0s08").Value
            Case "1"
               adoaccrpt109.Fields("r10913").Value = ComboItem(41)
            Case "2"
               adoaccrpt109.Fields("r10913").Value = ComboItem(42)
         End Select
      End If
      adoaccrpt109.UpdateBatch
      adoacc0s0.MoveNext
   Loop
   adoacc0s0.Close
   adoaccrpt109.Close
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
Private Sub Accrpt109Delete()
   adoTaie.Execute "delete from accrpt109"
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Len(Text2) = 6 Then
      Text2 = AfterZero(Text2)
   End If
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
   Text1 = "X"
   Text2 = "X"
   MaskEdBox1.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) And Text1 <> "X" Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) And Text2 <> "X" Then
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

Private Sub Text3_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text3.IMEMode = 2
   CloseIme
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
      KeyAscii = 0
   End If
End Sub
Private Sub Text4_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text3.IMEMode = 2
   CloseIme
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   End If
End Sub
