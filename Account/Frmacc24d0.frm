VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc24d0 
   AutoRedraw      =   -1  'True
   Caption         =   "代理人FC帳款明細表"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   6450
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1230
      TabIndex        =   7
      Text            =   "ALL"
      Top             =   1800
      Width           =   3510
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   1
      Left            =   3870
      TabIndex        =   6
      Top             =   1440
      Width           =   2385
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   0
      Left            =   1230
      TabIndex        =   5
      Top             =   1440
      Width           =   2385
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1710
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2190
      Width           =   612
   End
   Begin VB.CommandButton Command2 
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
      Left            =   840
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   2940
      Width           =   4692
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1230
      MaxLength       =   1
      TabIndex        =   2
      Top             =   660
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1230
      MaxLength       =   9
      TabIndex        =   0
      Top             =   270
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3090
      MaxLength       =   9
      TabIndex        =   1
      Top             =   270
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   1230
      TabIndex        =   3
      Top             =   1050
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
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
      Height          =   315
      Left            =   3060
      TabIndex        =   4
      Top             =   1050
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
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
      Caption         =   "系統類別："
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
      Index           =   0
      Left            =   270
      TabIndex        =   20
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "(Y:是)"
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
      Left            =   2370
      TabIndex        =   19
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "是否列印明細"
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
      Left            =   270
      TabIndex        =   18
      Top             =   2190
      Width           =   1455
   End
   Begin VB.Label Label8 
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
      Left            =   3660
      TabIndex        =   17
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "國籍"
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
      Left            =   270
      TabIndex        =   16
      Top             =   1440
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   2760
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2850
      TabIndex        =   15
      Top             =   1050
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "請款日期"
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
      Left            =   270
      TabIndex        =   14
      Top             =   1050
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "(1.往來帳款 2.應收帳款)"
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
      Left            =   1980
      TabIndex        =   13
      Top             =   660
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "資料性質"
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
      Left            =   270
      TabIndex        =   12
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
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
      Index           =   1
      Left            =   270
      TabIndex        =   11
      Top             =   270
      Width           =   975
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
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   270
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc24d0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/2 日期欄已修改
Option Explicit

Public adoacc1k0 As New ADODB.Recordset
Public adoacc0y0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoaccrpt213 As New ADODB.Recordset
Dim strSql As String
Dim intCounter As Integer
Dim intPage As Integer
Dim intRecord As Integer
Dim strAmount As String
Dim m_blnNoData As Boolean  '判斷是否無資料
'Add by Morgan 2006/1/9
Dim stConFA As String
Dim stConCu As String
Dim bolOneAgent As Boolean '是否指定一個代理人
Dim arrNation 'Add By Sindy 2012/8/13
Dim pLeft(0 To 10) As Integer


'Add By Sindy 2012/8/13
Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2012/8/13
Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
   Select Case Index
   Case 0, 1 '申請國家
      If Me.Combo1(Index).Text <> "" Then
         arrNation = Split(Me.Combo1(Index).Text, " ")
         StrSQLa = "Select NA01,NA03,NA59 From Nation Where Length(NA01)=3 And NA01<='9999' And substr(NA02,3,1)='0' And NA01='" & arrNation(0) & "' Order By NA01 "
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            Me.Combo1(Index).Text = "" & rsA.Fields(0).Value & " " & rsA.Fields(1).Value
         Else
            MsgBox "申請國代號輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
            Exit Sub
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If
   End Select
   If Cancel = True Then
      Me.Combo1(Index).SetFocus
      Me.Combo1(Index).SelStart = 0
      Me.Combo1(Index).SelLength = Len(Me.Combo1(Index).Text)
   End If
End Sub

Private Sub Command2_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt213Delete
   ProduceData
'   PrintData
    If m_blnNoData = False Then PrintData
   FormClear
   Screen.MousePointer = vbDefault
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
Dim intY As Integer, i As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim rsA As New ADODB.Recordset
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 6570 '5250
   Me.Height = 3840 '3450
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
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   
   'Add By Sindy 2012/8/13
   '國籍
   strSql = "Select NA01, NA03 From Nation Where Length(NA01)=3 And NA01<='9999' And substr(NA02,3,1)='0' Order By NA01 "
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   While Not rsA.EOF
       For i = 0 To 1
          Me.Combo1(i).AddItem "" & rsA.Fields(0).Value & " " & rsA.Fields(1).Value
       Next i
       rsA.MoveNext
   Wend
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   '2012/8/13 End
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc24d0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
'   If Len(Text1) = 6 Then
'      Text1 = AfterZero(Text1)
'   End If
    If Me.Text1.Text <> "" Then
        Me.Text1.Text = Left(Me.Text1.Text & "00000000", 9)
    End If
    Me.Text2.Text = Me.Text1.Text
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Len(Text2) = 6 Then
      Text2 = AfterZero(Text2)
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

'Add By Sindy 2012/8/13
Private Sub Text4_GotFocus()
   TextInverse Text4
   CloseIme
End Sub

'Add By Sindy 2012/8/13
Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2012/8/13
Private Sub Text4_Validate(Cancel As Boolean)
   If Me.Text4.Text <> "ALL" Then
      If Not CheckSysKind1("" & Me.Text4.Text) Then
         Me.Text4.SetFocus
         Cancel = True
      End If
   End If
   If Cancel Then Text4_GotFocus
End Sub
'檢查輸入的系統類別是否正確
Private Function CheckSysKind1(strSysKind As String) As Boolean
Dim arr1
Dim arr2
Dim ii As Integer
Dim jj As Integer
   
   CheckSysKind1 = False
   arr2 = Split(Me.Text4.Text, ",")
   For ii = LBound(arr2) To UBound(arr2)
      If CheckSys(arr2(ii)) = "" Then
         MsgBox "系統類別輸入錯誤!!!", vbExclamation + vbOKOnly
         Exit Function
      End If
   Next ii
   CheckSysKind1 = True
End Function

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim arr1
Dim ii As Integer, strSystemKind As String
   
On Error GoTo Checking
   
   m_blnNoData = True
   strSql = "": stConCu = "": stConFA = "" 'Add by Morgan 2006/6/6
   
   'Add by Morgan 2007/4/26
   If Text1 <> "" And Text2 = Text1 Then
      bolOneAgent = True
   Else
      bolOneAgent = False
   End If
   
   If Text1 <> MsgText(601) Then
      strSql = " and a1k03 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a1k03 <= '" & Text2 & "'"
   End If
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   'Modify By Sindy 2012/8/13
   If Combo1(0).Text <> "" Then
      arrNation = Split(Me.Combo1(0).Text, " ")
      If arrNation(0) <> MsgText(601) Then
         'Modify by Morgan 2006/1/9
         'strSQL = strSQL & " and fa10 >= '" & Text4 & "'"
         stConCu = stConCu & " and CU10 >= '" & arrNation(0) & "'"
         stConFA = stConFA & " and FA10 >= '" & arrNation(0) & "'"
      End If
   End If
   If Combo1(1).Text <> "" Then
      arrNation = Split(Me.Combo1(1).Text, " ")
      If arrNation(0) <> MsgText(601) Then
         'Modify by Morgan 2006/1/9
         'strSQL = strSQL & " and fa10 <= '" & Text5 & "z'"
         stConCu = stConCu & " and CU10 <= '" & arrNation(0) & "z'"
         stConFA = stConFA & " and FA10 <= '" & arrNation(0) & "z'"
      End If
   End If
   '2012/8/13 End
   'Add By Sindy 2012/8/13 +系統類別
   If Text4 <> "ALL" Then
      arr1 = Split(Me.Text4.Text, ",")
      For ii = LBound(arr1) To UBound(arr1)
         strSystemKind = strSystemKind & "'" & arr1(ii) & "',"
      Next ii
      strSystemKind = Left(strSystemKind, Len(strSystemKind) - 1)
      strSql = strSql & " AND A1K13 IN ( " & strSystemKind & " ) "
   End If
   '2012/8/13 End
   'Add By Cheng 2004/03/15
   '加未銷帳的條件
   strSql = strSql & " And A1K25 Is Null "
   'End
   'add by nickc 2007/07/12  加入作廢的不印
   strSql = strSql & " and a1k12 is null "
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt213.CursorLocation = adUseClient
   adoaccrpt213.Open "select * from accrpt213", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Select Case Text3
      Case Mid(ComboItem(1), 1, 1)
         Select1
      Case Mid(ComboItem(2), 1, 1)
         Select2
      Case Else
         Select1
   End Select
   adoaccrpt213.Close
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
Private Sub Accrpt213Delete()
   adoTaie.Execute "delete from accrpt213"
End Sub

'*************************************************
'  選擇往來帳款統計
'
'*************************************************
Private Sub Select1()
Dim douExchange As Double
Dim StrSQLa As String

   adoacc1k0.CursorLocation = adUseClient
    'Modify By Cheng 2004/03/15
'   adoacc1k0.Open "select * from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2014/10/2 +,A1K17
   StrSQLa = "select A1K01, A1K02, (A1K08-nvl(A1K31,0)) as A1K08, A1K09, (A1K11-nvl(A1K06,0)) as A1K11, A1K13, A1K14, A1K15, A1K16, A1K18,A1K17 from acc1k0, fagent where substr(a1k03, 1, 8) = fa01(+) and substr(a1k03, 9, 1) = fa02(+) " & strSql & stConFA
   StrSQLa = StrSQLa & " Union select A1K01, A1K02, (A1K08-nvl(A1K31,0)) as A1K08, A1K09, (A1K11-nvl(A1K06,0)) as A1K11, A1K13, A1K14, A1K15, A1K16, A1K18,A1K17 from acc1k0, Customer where substr(a1k03, 1, 8) = CU01(+) and substr(a1k03, 9, 1) = CU02(+) " & strSql & stConCu
   adoacc1k0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
    'End
   If adoacc1k0.RecordCount = 0 Then
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
    m_blnNoData = False
   Do While adoacc1k0.EOF = False
      adoacc0y0.CursorLocation = adUseClient
      '國外收款
      adoacc0y0.Open "select * from acc0z0, acc0y0 where a0z01 = a0y01 and a0z02 = '" & adoacc1k0.Fields("a1k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0y0.RecordCount = 0 Then
         'Add By Sindy 2014/10/2 無收款資料,檢查是否有a1k17值,若有則抓抵帳資料
         If "" & adoacc1k0.Fields("a1k17").Value <> "" Then
            adoacc0y0.Close
            adoacc0y0.Open "select a1k17,decode(a1h02, null, a1i03, a1h02) as DocDate,NVL(A1H03,A1I05) as Currency,(a1k08 - nvl(a1k31, 0)) as Famount,(a1k08 - nvl(a1k31, 0)) * nvl(a1g02, 0) as Namount from acc1k0,acc1g0,acc1h0,acc1i0 where a1k01='" & adoacc1k0.Fields("a1k01").Value & "' and a1k17 = a1g01(+) and a1k17 = a1h01(+) and a1k17 = a1i01(+)", adoTaie, adOpenStatic, adLockReadOnly
            If adoacc0y0.RecordCount = 0 Then
               adoaccrpt213.AddNew
               ARSave
               adoaccrpt213.UpdateBatch
            Else
               Do While adoacc0y0.EOF = False
                  adoaccrpt213.AddNew
                  ARSave
                  adoaccrpt213.Fields("r21317").Value = adoacc0y0.Fields("a1k17").Value
                  If IsNull(adoacc0y0.Fields("DocDate").Value) Then
                     adoaccrpt213.Fields("r21309").Value = Null
                  Else
                     adoaccrpt213.Fields("r21309").Value = adoacc0y0.Fields("DocDate").Value
                  End If
                  If IsNull(adoacc0y0.Fields("Currency").Value) Then
                     adoaccrpt213.Fields("r21310").Value = Null
                  Else
                     adoaccrpt213.Fields("r21310").Value = adoacc0y0.Fields("Currency").Value
                  End If
                  If IsNull(adoacc0y0.Fields("Famount").Value) Then
                     adoaccrpt213.Fields("r21311").Value = 0
                  Else
                     adoaccrpt213.Fields("r21311").Value = adoacc0y0.Fields("Famount").Value
                  End If
                  adoaccrpt213.Fields("r21312").Value = adoacc0y0.Fields("Namount").Value
                  adoaccrpt213.Fields("r21315").Value = (Val(adoaccrpt213.Fields("r21312").Value) - Val(adoaccrpt213.Fields("r21308").Value)) / 1000
                  adoaccrpt213.Fields("r21316").Value = Val(adoaccrpt213.Fields("r21306").Value) - Val(adoaccrpt213.Fields("r21312").Value)
                  adoaccrpt213.UpdateBatch
                  adoacc0y0.MoveNext
               Loop
            End If
         Else
            adoaccrpt213.AddNew
            ARSave
            adoaccrpt213.UpdateBatch
         End If
      Else
         Do While adoacc0y0.EOF = False
            adoaccrpt213.AddNew
            ARSave
            adoaccrpt213.Fields("r21317").Value = adoacc0y0.Fields("a0z01").Value
            If IsNull(adoacc0y0.Fields("a0y02").Value) Then
               adoaccrpt213.Fields("r21309").Value = Null
            Else
               adoaccrpt213.Fields("r21309").Value = adoacc0y0.Fields("a0y02").Value
            End If
            'Modify By Sindy 2012/12/7 原程式抓a0z03改抓a0y03
'            If IsNull(adoacc0y0.Fields("a0z03").Value) Then
'               adoaccrpt213.Fields("r21310").Value = Null
'            Else
'               adoaccrpt213.Fields("r21310").Value = adoacc0y0.Fields("a0z03").Value
'            End If
            If IsNull(adoacc0y0.Fields("a0y03").Value) Then
               adoaccrpt213.Fields("r21310").Value = Null
            Else
               adoaccrpt213.Fields("r21310").Value = adoacc0y0.Fields("a0y03").Value
            End If
            '2012/12/7 End
            If IsNull(adoacc0y0.Fields("a0z04").Value) Then
               adoaccrpt213.Fields("r21311").Value = 0
            Else
               adoaccrpt213.Fields("r21311").Value = adoacc0y0.Fields("a0z04").Value
            End If
            If IsNull(adoacc0y0.Fields("a0y04").Value) Then
               douExchange = 0
            Else
               douExchange = adoacc0y0.Fields("a0y04").Value
            End If
            adoaccrpt213.Fields("r21312").Value = Val(adoaccrpt213.Fields("r21311").Value) * douExchange
            'Modify by Morgan 2007/4/26 點數要扣規費(分次收款時每次都要扣)
            'adoaccrpt213.Fields("r21315").Value = Val(adoaccrpt213.Fields("r21312").Value) / 1000
            adoaccrpt213.Fields("r21315").Value = (Val(adoaccrpt213.Fields("r21312").Value) - Val(adoaccrpt213.Fields("r21308").Value)) / 1000
            'end 2007/4/26
            adoaccrpt213.Fields("r21316").Value = Val(adoaccrpt213.Fields("r21306").Value) - Val(adoaccrpt213.Fields("r21312").Value)
            adoaccrpt213.UpdateBatch
            adoacc0y0.MoveNext
         Loop
      End If
      adoacc0y0.Close
      adoacc1k0.MoveNext
   Loop
   adoacc1k0.Close
End Sub

'*************************************************
'  選擇應收帳款統計
'
'*************************************************
Private Sub Select2()
Dim douExchange As Double
Dim StrSQLa As String

   adoacc1k0.CursorLocation = adUseClient
    'Modify By Cheng 2004/03/15
'   adoacc1k0.Open "select * from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and (a1k29 is null or a1k29 = '')" & strSQL, adoTaie, adOpenDynamic, adLockBatchOptimistic
   StrSQLa = "select A1K01, A1K02, (A1K08-nvl(A1K31,0)) as A1K08, A1K09, (A1K11-nvl(A1K06,0)) as A1K11, A1K13, A1K14, A1K15, A1K16, A1K18 from acc1k0, fagent where substr(a1k03, 1, 8) = fa01(+) and substr(a1k03, 9, 1) = fa02(+) and (a1k29 is null or a1k29 = '')" & strSql & stConFA
   StrSQLa = StrSQLa & " Union select A1K01, A1K02, (A1K08-nvl(A1K31,0)) as A1K08, A1K09, (A1K11-nvl(A1K06,0)) as A1K11, A1K13, A1K14, A1K15, A1K16, A1K18 from acc1k0, Customer where substr(a1k03, 1, 8) = CU01(+) and substr(a1k03, 9, 1) = CU02(+) and (a1k29 is null or a1k29 = '')" & strSql & stConCu
   adoacc1k0.Open StrSQLa, adoTaie, adOpenDynamic, adLockBatchOptimistic
    'End
   If adoacc1k0.RecordCount = 0 Then
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
    m_blnNoData = False
   Do While adoacc1k0.EOF = False
      adoaccrpt213.AddNew
      ARSave
      adoaccrpt213.UpdateBatch
      adoacc1k0.MoveNext
   Loop
   adoacc1k0.Close
End Sub

'*************************************************
'  請款資料儲存
'
'*************************************************
Private Sub ARSave()
   adoaccrpt213.Fields("r21301").Value = strUserNum
   adoaccrpt213.Fields("r21302").Value = adoacc1k0.Fields("a1k01").Value
   If IsNull(adoacc1k0.Fields("a1k02").Value) Then
      adoaccrpt213.Fields("r21303").Value = Null
   Else
      adoaccrpt213.Fields("r21303").Value = adoacc1k0.Fields("a1k02").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k18").Value) Then
      adoaccrpt213.Fields("r21304").Value = Null
   Else
      adoaccrpt213.Fields("r21304").Value = adoacc1k0.Fields("a1k18").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k08").Value) = False Then
      adoaccrpt213.Fields("r21305").Value = Val(adoacc1k0.Fields("a1k08").Value)
   Else
      adoaccrpt213.Fields("r21305").Value = 0
   End If
   If IsNull(adoacc1k0.Fields("a1k11").Value) Then
      adoaccrpt213.Fields("r21306").Value = 0
   Else
      adoaccrpt213.Fields("r21306").Value = adoacc1k0.Fields("a1k11").Value
   End If
   
   If IsNull(adoacc1k0.Fields("a1k09").Value) Then
      adoaccrpt213.Fields("r21308").Value = 0
   Else
      adoaccrpt213.Fields("r21308").Value = adoacc1k0.Fields("a1k09").Value
   End If
   'Modify by Morgan 2007/4/26 點數要扣規費
   'adoaccrpt213.Fields("r21307").Value = Val(adoaccrpt213.Fields("r21306").Value) / 1000
   adoaccrpt213.Fields("r21307").Value = (Val(adoaccrpt213.Fields("r21306").Value) - Val(adoaccrpt213.Fields("r21308").Value)) / 1000
   'end 2007/4/26
   adoaccrpt213.Fields("r21313").Value = adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select pa77 as Yno from patent where pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                 "select tm45 as Yno from trademark where tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                 "select lc23 as Yno from lawcase where lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
                 "select sp27 as Yno from servicepractice where sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("Yno").Value) = False Then
         adoaccrpt213.Fields("r21314").Value = adoquery.Fields("Yno").Value
      End If
   End If
   adoquery.Close
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
   'Modify By Sindy 2012/8/13
   Combo1(0) = ""
   Combo1(1) = ""
   Text4 = "ALL"
   '2012/8/13 End
   Text6 = ""
   Text1.SetFocus
End Sub

'*************************************************
'  產生對帳資料
'
'*************************************************
Public Sub PrintData()
Dim strNo As String

On Error GoTo Checking
   
   intPage = 0
   If Text6 = MsgText(602) Then
      strSql = ""
      intCounter = 3
      intRecord = 1
      adoquery.CursorLocation = adUseClient
      'Modify by Morgan 2011/5/25 +fa70
      'modify by sonia 2017/12/11 FA05改為NVL(NVL(FA05,FA04),FA06)
      strSql = "select A1K03, R21301, R21302, R21303, R21304, R21305, R21306, R21307, R21308, R21309, R21310, R21311, R21312, R21313, R21314, R21315, R21316, NVL(NVL(FA05,FA04),FA06) FA05, FA18, FA19, FA20, FA21, FA22, FA32, FA33, FA34, FA35, FA36,FA70 from accrpt213, acc1k0, fagent where r21302 = a1k01 and substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 "
      strSql = strSql & " Union select A1K03, R21301, R21302, R21303, R21304, R21305, R21306, R21307, R21308, R21309, R21310, R21311, R21312, R21313, R21314, R21315, R21316, NVL(NVL(CU05,CU04),CU06) As FA05, CU24 As FA18, CU25 As FA19, CU26 As FA20, CU27 As FA21, CU28 As FA22, CU65 As FA32, CU66 As FA33, CU67 As FA34, CU68 As FA35, CU69 As FA36,cu102 as FA70 from accrpt213, acc1k0, Customer where r21302 = a1k01 and substr(a1k03, 1, 8) = CU01 and substr(a1k03, 9, 1) = CU02 "
      strSql = strSql & " Order By 2, 1 "
      adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      Do While adoquery.EOF = False
         If strNo <> adoquery.Fields("a1k03").Value Then
            If strNo <> "" Then
               Printer.NewPage
            End If
            intCounter = 3
            intRecord = 1
            intPage = intPage + 1
            PrintHead
            strNo = adoquery.Fields("a1k03").Value
         End If
         '請款編號
         Printer.CurrentX = 0
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21302").Value) Then
            Printer.Print ""
         Else
            Printer.Print adoquery.Fields("R21302").Value
         End If
         '請款日期
         Printer.CurrentX = 1500
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21303").Value) Then
            Printer.Print ""
         Else
            Printer.Print CFDate(adoquery.Fields("R21303").Value)
         End If
         '幣別
         Printer.CurrentX = 2800
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21304").Value) Then
            Printer.Print ""
         Else
            Printer.Print adoquery.Fields("R21304").Value
         End If
         '外幣款請
         If IsNull(adoquery.Fields("R21305").Value) = False Then
            strAmount = Format(Val(adoquery.Fields("R21305").Value), FDollar)
            Printer.CurrentX = 5000 - Printer.TextWidth(strAmount)
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         End If
         '台幣請款
         If IsNull(adoquery.Fields("R21306").Value) = False Then
            strAmount = Format(Val(adoquery.Fields("R21306").Value), FDollar)
            Printer.CurrentX = 6500 - Printer.TextWidth(strAmount)
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         End If
         '規費
         If IsNull(adoquery.Fields("R21308").Value) = False Then
            strAmount = Format(Val(adoquery.Fields("R21308").Value), FDollar)
            Printer.CurrentX = 8000 - Printer.TextWidth(strAmount)
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         End If
         '收款日期
         Printer.CurrentX = 8100
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21309").Value) Then
            Printer.Print ""
         Else
            Printer.Print CFDate(adoquery.Fields("R21309").Value)
         End If
         '幣別
         Printer.CurrentX = 9400
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21310").Value) Then
            Printer.Print ""
         Else
            Printer.Print adoquery.Fields("R21310").Value
         End If
         '收款金額
         If IsNull(adoquery.Fields("R21311").Value) = False Then
            strAmount = Format(Val(adoquery.Fields("R21311").Value), FDollar)
            Printer.CurrentX = 11600 - Printer.TextWidth(strAmount)
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         End If
         '台幣收款
         If IsNull(adoquery.Fields("R21312").Value) = False Then
            strAmount = Format(Val(adoquery.Fields("R21312").Value), FDollar)
            Printer.CurrentX = 13100 - Printer.TextWidth(strAmount)
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         End If
         '本所案號
         Printer.CurrentX = 13200
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21313").Value) Then
            Printer.Print ""
         Else
            Printer.Print adoquery.Fields("R21313").Value
         End If
         '彼所案號
         Printer.CurrentX = 15000
         Printer.CurrentY = 300 + intCounter * 300
         If IsNull(adoquery.Fields("R21314").Value) Then
            Printer.Print ""
         Else
            Printer.Print convForm(CheckStr(adoquery.Fields("R21314").Value), 25)
         End If
         '收款點數
         If IsNull(adoquery.Fields("R21315").Value) = False Then
            strAmount = Format(Val(adoquery.Fields("R21315").Value), FDollar)
'            Printer.CurrentX = 17400 - Printer.TextWidth(strAmount)
            Printer.CurrentX = 18400 - Printer.TextWidth(strAmount)
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
         End If
         'Modify By Sindy 2014/10/2 取消差額
'         If IsNull(adoquery.Fields("R21316").Value) = False Then
'            strAmount = Format(Val(adoquery.Fields("R21316").Value), FDollar)
'            Printer.CurrentX = 18400 - Printer.TextWidth(strAmount)
'            Printer.CurrentY = 300 + intCounter * 300
'            Printer.Print strAmount
'         End If
         '2014/10/2 END
         intCounter = intCounter + 1
         intRecord = intRecord + 1
         adoquery.MoveNext
        'Add By Cheng 2004/04/22
        If adoquery.EOF = False Then
            If strNo <> adoquery.Fields("a1k03").Value Or intRecord > 33 Then
               If strNo <> "" Then
                  Printer.NewPage
               End If
               intCounter = 3
               intRecord = 1
               intPage = intPage + 1
               PrintHead
               strNo = adoquery.Fields("a1k03").Value
            End If
        End If
        'End
      Loop
      adoquery.Close
      Printer.NewPage
   End If
   PrintSum
   Printer.EndDoc
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead()
   Printer.FontSize = 14
   Printer.CurrentX = 7000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print ReportTitle(213)
   Printer.FontSize = 12
   intCounter = intCounter + 2
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印人員: " & StaffQuery(strUserNum)
   Printer.CurrentX = 16000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印日期: " & CFDate(ACDate(ServerDate))
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "代理人編號: " & adoquery.Fields("a1k03").Value
   Printer.CurrentX = 4000
   Printer.CurrentY = 300 + intCounter * 300
   If IsNull(adoquery.Fields("fa05").Value) Then
      Printer.Print "代理人名稱(英): "
   Else
      Printer.Print "代理人名稱(英): " & adoquery.Fields("fa05").Value
   End If
   Printer.CurrentX = 16000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "頁次: " & intPage
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "代理人地址: "
   If IsNull(adoquery.Fields("fa32").Value) Then
      If IsNull(adoquery.Fields("fa18").Value) = False Then
         Printer.CurrentX = 2000
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print adoquery.Fields("fa18").Value
      End If
   Else
      Printer.CurrentX = 2000
      Printer.CurrentY = 300 + intCounter * 300
      Printer.Print adoquery.Fields("fa32").Value
   End If
   intCounter = intCounter + 1
   If IsNull(adoquery.Fields("fa32").Value) Then
      If IsNull(adoquery.Fields("fa19").Value) = False Then
         Printer.CurrentX = 2000
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print adoquery.Fields("fa19").Value
      End If
   Else
      Printer.CurrentX = 2000
      Printer.CurrentY = 300 + intCounter * 300
      Printer.Print "" & adoquery.Fields("fa33").Value
   End If
   intCounter = intCounter + 1
   If IsNull(adoquery.Fields("fa32").Value) Then
      If IsNull(adoquery.Fields("fa20").Value) = False Then
         Printer.CurrentX = 2000
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print adoquery.Fields("fa20").Value
      End If
   Else
      Printer.CurrentX = 2000
      Printer.CurrentY = 300 + intCounter * 300
      Printer.Print "" & adoquery.Fields("fa34").Value
   End If
   intCounter = intCounter + 1
   If IsNull(adoquery.Fields("fa32").Value) Then
      If IsNull(adoquery.Fields("fa21").Value) = False Then
         Printer.CurrentX = 2000
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print adoquery.Fields("fa21").Value
      End If
   Else
      Printer.CurrentX = 2000
      Printer.CurrentY = 300 + intCounter * 300
      Printer.Print "" & adoquery.Fields("fa35").Value
   End If
   intCounter = intCounter + 1
   If IsNull(adoquery.Fields("fa32").Value) Then
      If IsNull(adoquery.Fields("fa22").Value) = False Then
         Printer.CurrentX = 2000
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print adoquery.Fields("fa22").Value
      End If
      'Add by Morgan 2011/5/25
      '英文地址6
      If IsNull(adoquery.Fields("fa70").Value) = False Then
         intCounter = intCounter + 1
         Printer.CurrentX = 2000
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print adoquery.Fields("fa70").Value
      End If
   Else
      Printer.CurrentX = 2000
      Printer.CurrentY = 300 + intCounter * 300
      Printer.Print "" & adoquery.Fields("fa36").Value
   End If
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "帳款日期: " & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "請款編號"
   Printer.CurrentX = 1500
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "請款日期"
   Printer.CurrentX = 2800
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "幣別"
   Printer.CurrentX = 3600
   Printer.CurrentY = 300 + intCounter * 300
   'Modify By Sindy 2012/12/11
   'Printer.Print "美金請款"
   Printer.Print "外幣請款"
   '2012/12/11 End
   Printer.CurrentX = 5100
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "台幣請款"
   Printer.CurrentX = 6600
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "規費"
   Printer.CurrentX = 8100
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "收款日期"
   Printer.CurrentX = 9400
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "幣別"
   Printer.CurrentX = 10200
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "收款金額"
   Printer.CurrentX = 11700
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "台幣收款"
   Printer.CurrentX = 13200
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "本所案號"
   Printer.CurrentX = 15000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "彼所案號"
   Printer.CurrentX = 17500 '16500
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "點數"
   'Modify By Sindy 2014/10/2 取消差額
'   Printer.CurrentX = 17500
'   Printer.CurrentY = 300 + intCounter * 300
'   Printer.Print "差額"
   '2014/10/2 END
   Printer.Line (0, 300 + intCounter * 300 + 350)-(18500, 300 + intCounter * 300 + 350)
   intCounter = intCounter + 2
End Sub

'Add By Sindy 2012/12/11
'intType : 0.合計 1.明細
Private Sub GetPleft(intType As Integer)
   If intType = 0 Then
      pLeft(0) = 0
      pLeft(6) = 2200 '請款外幣金額合計 幣別欄
      pLeft(1) = 4000
      pLeft(2) = 4100
      pLeft(7) = 5500 '已收款之請款單外幣合計 幣別欄
      pLeft(3) = 8000
      pLeft(4) = 8100
      pLeft(5) = 11000
   End If
End Sub

'*************************************************
' 合計位置
'
'*************************************************
Private Sub PrintSum()
Dim intRow As Integer 'Add By Sindy 2012/12/12
   
   intCounter = 2
   intPage = intPage + 1
   Printer.FontSize = 14
   Printer.CurrentX = 7000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print ReportTitle(213)
   Printer.FontSize = 12
   intCounter = intCounter + 2
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印人員: " & StaffQuery(strUserNum)
   Printer.CurrentX = 16000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印日期: " & CFDate(ACDate(ServerDate))
   intCounter = intCounter + 1
   
   'Add by Morgan 2007/4/26 有輸入代理人條件時要印
   If bolOneAgent = True Then
      'Modify by Morgan 2011/5/25 +FA70
      'modify by sonia 2017/12/11 FA05改為NVL(NVL(FA05,FA04),FA06)
      strExc(0) = "select A1K03, R21301, R21302, R21303, R21304, R21305, R21306, R21307, R21308, R21309, R21310, R21311, R21312, R21313, R21314, R21315, R21316, NVL(NVL(FA05,FA04),FA06) FA05, FA18, FA19, FA20, FA21, FA22, FA32, FA33, FA34, FA35, FA36,FA70 from accrpt213, acc1k0, fagent where r21302 = a1k01 and substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 "
      strExc(0) = strExc(0) & " Union select A1K03, R21301, R21302, R21303, R21304, R21305, R21306, R21307, R21308, R21309, R21310, R21311, R21312, R21313, R21314, R21315, R21316, NVL(NVL(CU05,CU04),CU06) As FA05, CU24 As FA18, CU25 As FA19, CU26 As FA20, CU27 As FA21, CU28 As FA22, CU65 As FA32, CU66 As FA33, CU67 As FA34, CU68 As FA35, CU69 As FA36,cu102 as FA70 from accrpt213, acc1k0, Customer where r21302 = a1k01 and substr(a1k03, 1, 8) = CU01 and substr(a1k03, 9, 1) = CU02 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then

         Printer.CurrentX = 0
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print "代理人編號: " & RsTemp.Fields("a1k03").Value
         Printer.CurrentX = 4000
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print "代理人名稱(英): " & RsTemp.Fields("fa05").Value
         
         Printer.CurrentX = 16000
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print "頁次: " & intPage
         intCounter = intCounter + 1
         Printer.CurrentX = 0
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print "代理人地址: "
         If IsNull(RsTemp.Fields("fa32").Value) Then
            Printer.CurrentX = 2000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print "" & RsTemp.Fields("fa18").Value
         Else
            Printer.CurrentX = 2000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print RsTemp.Fields("fa32").Value
         End If
         intCounter = intCounter + 1
         If IsNull(RsTemp.Fields("fa32").Value) Then
            Printer.CurrentX = 2000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print "" & RsTemp.Fields("fa19").Value
         Else
            Printer.CurrentX = 2000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print "" & RsTemp.Fields("fa33").Value
         End If
         intCounter = intCounter + 1
         If IsNull(RsTemp.Fields("fa32").Value) Then
            Printer.CurrentX = 2000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print "" & RsTemp.Fields("fa20").Value
         Else
            Printer.CurrentX = 2000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print "" & RsTemp.Fields("fa34").Value
         End If
         intCounter = intCounter + 1
         If IsNull(RsTemp.Fields("fa32").Value) Then
            Printer.CurrentX = 2000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print "" & RsTemp.Fields("fa21").Value
         Else
            Printer.CurrentX = 2000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print "" & RsTemp.Fields("fa35").Value
         End If
         intCounter = intCounter + 1
         If IsNull(RsTemp.Fields("fa32").Value) Then
            Printer.CurrentX = 2000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print "" & RsTemp.Fields("fa22").Value
            
            'Add by Morgan 2011/5/25
            '英文地址6
            If Not IsNull(RsTemp.Fields("fa70").Value) Then
               intCounter = intCounter + 1
               Printer.CurrentX = 2000
               Printer.CurrentY = 300 + intCounter * 300
               Printer.Print "" & RsTemp.Fields("fa70").Value
            End If
            
         Else
            Printer.CurrentX = 2000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print "" & RsTemp.Fields("fa36").Value
         End If
      End If
   End If
   
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "帳款日期: " & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
   
   'Modify by Morgan 2007/4/26 有輸入代理人條件時不印
   If bolOneAgent = False Then
      Printer.CurrentX = 16000
      Printer.CurrentY = 300 + intCounter * 300
      Printer.Print "頁次: " & intPage
   End If
   
   Call GetPleft(0)  'Add By Sindy 2012/12/11
   
   intCounter = 5
   Printer.CurrentX = pLeft(0)
   Printer.CurrentY = 3000 + intCounter * 300
   Printer.Print "PS:"
   adoquery.CursorLocation = adUseClient
   'Modify By Sindy 2012/12/11 請款外幣合計
   adoquery.Open "select sum(r21305), sum(r21306), sum(r21307), sum(r21311), sum(r21312), sum(r21315), sum(DECODE(R21311,NULL,0,r21305)), r21304 from accrpt213 where r21301 = '" & strUserNum & "' group by r21304 order by r21304", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      adoquery.MoveFirst
      intCounter = intCounter + 2
      intRow = 0
      Do While adoquery.EOF = False
         intRow = intRow + 1
         If intRow = 1 Then
            Printer.CurrentX = pLeft(0)
            Printer.CurrentY = 3000 + intCounter * 300
            Printer.Print "請款外幣金額合計:"
         End If
         Printer.CurrentX = pLeft(6)
         Printer.CurrentY = 3000 + intCounter * 300
         Printer.Print "" & adoquery.Fields(7).Value
         If IsNull(adoquery.Fields(0).Value) = False Then
            strAmount = Format(Val(adoquery.Fields(0).Value), FDollar)
            Printer.CurrentX = pLeft(1) - Printer.TextWidth(strAmount)
            Printer.CurrentY = 3000 + intCounter * 300
            Printer.Print strAmount
         End If
         If intRow = 1 Then
            Printer.CurrentX = pLeft(2)
            Printer.CurrentY = 3000 + intCounter * 300
            Printer.Print "外幣已收款合計:"
         End If
         If IsNull(adoquery.Fields(3).Value) = False Then
            strAmount = Format(Val(adoquery.Fields(3).Value), FDollar)
            Printer.CurrentX = pLeft(3) - Printer.TextWidth(strAmount)
            Printer.CurrentY = 3000 + intCounter * 300
            Printer.Print strAmount
         End If
         If intRow = 1 Then
            Printer.CurrentX = pLeft(4)
            Printer.CurrentY = 3000 + intCounter * 300
            Printer.Print "請款點數:"
         End If
         If IsNull(adoquery.Fields(2).Value) = False Then
            strAmount = Format(Val(adoquery.Fields(2).Value), FDollar)
            Printer.CurrentX = pLeft(5) - Printer.TextWidth(strAmount)
            Printer.CurrentY = 3000 + intCounter * 300
            Printer.Print strAmount
         End If
         intCounter = intCounter + 1
         adoquery.MoveNext
      Loop
      intCounter = intCounter - 1
   End If
   adoquery.Close
   '2012/12/11 End
   
   '2009/2/20 MODIFY BY SONIA 婧瑄說加收款之請款單外幣合計
   'adoquery.Open "select sum(r21305), sum(r21306), sum(r21307), sum(r21311), sum(r21312), sum(r21315) from accrpt213 where r21301 = '" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "select sum(r21305), sum(r21306), sum(r21307), sum(r21311), sum(r21312), sum(r21315), sum(DECODE(R21311,NULL,0,r21305)) from accrpt213 where r21301 = '" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
   '2009/2/20 END
   If adoquery.RecordCount <> 0 Then
'      Printer.CurrentX = 0
'      Printer.CurrentY = 3000 + intCounter * 300
'      Printer.Print "請款美金金額合計:"
'      If IsNull(adoquery.Fields(0).Value) = False Then
'         strAmount = Format(Val(adoquery.Fields(0).Value), FDollar)
'         Printer.CurrentX = 4000 - Printer.TextWidth(strAmount)
'         Printer.CurrentY = 3000 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      Printer.CurrentX = 4100
'      Printer.CurrentY = 3000 + intCounter * 300
'      Printer.Print "外幣已收款合計:"
'      If IsNull(adoquery.Fields(3).Value) = False Then
'         strAmount = Format(Val(adoquery.Fields(3).Value), FDollar)
'         Printer.CurrentX = 8000 - Printer.TextWidth(strAmount)
'         Printer.CurrentY = 3000 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      Printer.CurrentX = 8100
'      Printer.CurrentY = 3000 + intCounter * 300
'      Printer.Print "請款點數:"
'      If IsNull(adoquery.Fields(2).Value) = False Then
'         strAmount = Format(Val(adoquery.Fields(2).Value), FDollar)
'         Printer.CurrentX = 11000 - Printer.TextWidth(strAmount)
'         Printer.CurrentY = 3000 + intCounter * 300
'         Printer.Print strAmount
'      End If
      intCounter = intCounter + 2
      Printer.CurrentX = pLeft(0)
      Printer.CurrentY = 3000 + intCounter * 300
      Printer.Print "台幣請款金額合計:"
      If IsNull(adoquery.Fields(1).Value) = False Then
         strAmount = Format(Val(adoquery.Fields(1).Value), FDollar)
         Printer.CurrentX = pLeft(1) - Printer.TextWidth(strAmount)
         Printer.CurrentY = 3000 + intCounter * 300
         Printer.Print strAmount
      End If
      Printer.CurrentX = pLeft(2)
      Printer.CurrentY = 3000 + intCounter * 300
      Printer.Print "台幣已收款合計:"
      If IsNull(adoquery.Fields(4).Value) = False Then
         strAmount = Format(Val(adoquery.Fields(4).Value), FDollar)
         Printer.CurrentX = pLeft(3) - Printer.TextWidth(strAmount)
         Printer.CurrentY = 3000 + intCounter * 300
         Printer.Print strAmount
      End If
      Printer.CurrentX = pLeft(4)
      Printer.CurrentY = 3000 + intCounter * 300
      Printer.Print "收款點數:"
      If IsNull(adoquery.Fields(5).Value) = False Then
         strAmount = Format(Val(adoquery.Fields(5).Value), FDollar)
         Printer.CurrentX = pLeft(5) - Printer.TextWidth(strAmount)
         Printer.CurrentY = 3000 + intCounter * 300
         Printer.Print strAmount
      End If
'      '2009/2/20 ADD BY SONIA 婧瑄說加收款之請款單外幣合計
'      intCounter = intCounter + 2
'      Printer.CurrentX = 3150
'      Printer.CurrentY = 3000 + intCounter * 300
'      Printer.Print "已收款之請款單外幣合計:"
'      If IsNull(adoquery.Fields(6).Value) = False Then
'         strAmount = Format(Val(adoquery.Fields(6).Value), FDollar)
'         Printer.CurrentX = 8000 - Printer.TextWidth(strAmount)
'         Printer.CurrentY = 3000 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      '2009/2/20 END
   End If
   adoquery.Close
   
   'Modify By Sindy 2012/12/11 收款之請款單外幣合計
   adoquery.Open "select r21304, sum(DECODE(R21311,NULL,0,r21305)) from accrpt213 where r21301 = '" & strUserNum & "' group by r21304 order by r21304", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      adoquery.MoveFirst
      intCounter = intCounter + 2
      intRow = 0
      Do While adoquery.EOF = False
         intRow = intRow + 1
         If intRow = 1 Then
            Printer.CurrentX = pLeft(6) '3150
            Printer.CurrentY = 3000 + intCounter * 300
            Printer.Print "已收款之請款單外幣合計:"
         End If
         'Add By Sindy 2012/12/12
         Printer.CurrentX = pLeft(7)
         Printer.CurrentY = 3000 + intCounter * 300
         Printer.Print "" & adoquery.Fields(0).Value
         '2012/12/12 End
         If IsNull(adoquery.Fields(1).Value) = False Then
            strAmount = Format(Val(adoquery.Fields(1).Value), FDollar)
            Printer.CurrentX = pLeft(3) - Printer.TextWidth(strAmount)
            Printer.CurrentY = 3000 + intCounter * 300
            Printer.Print strAmount
         End If
         intCounter = intCounter + 1
         adoquery.MoveNext
      Loop
      intCounter = intCounter - 1
   End If
   adoquery.Close
   '2012/12/11 End
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   'Modify By Sindy 2012/8/13
   If Combo1(0).Text <> MsgText(601) Then
   '2012/8/13 End
      FormCheck = True
      Exit Function
   End If
   'Modify By Sindy 2012/8/13
   If Combo1(1).Text <> MsgText(601) Then
   '2012/8/13 End
      FormCheck = True
      Exit Function
   End If
   'Add By Sindy 2012/8/13
   If Text4 <> "ALL" Then
      FormCheck = True
      Exit Function
   End If
   '2012/8/13 End
   FormCheck = False
End Function
