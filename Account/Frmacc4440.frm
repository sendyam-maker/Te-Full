VERSION 5.00
Begin VB.Form Frmacc4440 
   AutoRedraw      =   -1  'True
   Caption         =   "試算表"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   5160
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3480
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   495
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
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   612
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
      TabIndex        =   6
      Top             =   2520
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
      TabIndex        =   5
      Top             =   2520
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
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo13 
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
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   1812
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
      TabIndex        =   7
      Top             =   1080
      Width           =   4692
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
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   612
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
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   612
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "(1.台一 2.智權 空白.全部)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   240
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1335
      Left            =   240
      Top             =   1680
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc4440.frx":0000
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label13 
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
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3000
      Picture         =   "Frmacc4440.frx":0442
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label14 
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
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label15 
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
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "月份"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   600
      Width           =   612
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "年度"
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
      Width           =   612
   End
   Begin VB.Label Label1 
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
      TabIndex        =   8
      Top             =   240
      Width           =   732
   End
End
Attribute VB_Name = "Frmacc4440"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc040M As New ADODB.Recordset
Public adoacc040MM As New ADODB.Recordset
Public adoacc040Y As New ADODB.Recordset
Public adoacc021 As New ADODB.Recordset
Public adoaccrpt405 As New ADODB.Recordset
Dim strSort1, strSort2 As String
Dim dllaccrpt405 As Object

Private Sub Combo13_KeyUp(KeyCode As Integer, Shift As Integer)
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
         Text6.SetFocus
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt405Delete
   ProduceData
   If adoaccrpt405.State = adStateOpen Then
      adoaccrpt405.Close
   End If
   adoaccrpt405.CursorLocation = adUseClient
   adoaccrpt405.Open "select * from accrpt405", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt405.RecordCount <> 0 Then
      '20140123START Modify By eric 原固定為1 現增加 J-智權公司
      dllaccrpt405.Acc4440 ReportTitle(405), IIf(Text6 = "2", "J", Text6), IIf(Text6 = "", "台一　專利商標/智權", Text7), Text3, Text1, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      'dllaccrpt405.Acc4440 ReportTitle(405), Text6, Text7, Text3, Text1, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
      '20140123END
   End If
   adoaccrpt405.Close
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
   Me.Height = 2100
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Combo4.AddItem MsgText(1)
   Combo4.AddItem MsgText(2)
   Combo6.AddItem MsgText(1)
   Combo6.AddItem MsgText(2)
   Combo4 = MsgText(1)
   Combo6 = MsgText(1)
   ComboAdd
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt405 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt405 = Nothing
   Set Frmacc4440 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text6_Change()
   If Text6 = MsgText(601) Then
      Exit Sub
   End If
   
   '20140123START Add By eric
   If Text6.Text <> "1" And Text6.Text <> "2" And Text6.Text <> "" Then
      MsgBox "公司別僅可為 1 或 2 或不輸入  !"
      Text6.Text = ""
      Text6.SetFocus
      Exit Sub
   End If
  
   Select Case Text6
      Case "1"
         Text7 = A0802Query(Text6)
      Case "2"
         Text7 = A0802Query("J")
   End Select
   'Text7 = A0802Query(Text6)
   '20140123END
   
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
   '20140123START Add By eric
   CloseIme
   '20140123END
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strOrder1, strOrder2, strSql As String
Dim lngStartDate, lngEndDate As Long
Dim intCounter As Integer
Dim Text66 As String            '20140123ADD By eric 公司別可變動

   
On Error GoTo Checking
   intCounter = 0
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo13
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0405 asc"
         Else
            strOrder1 = " order by a0405 desc"
         End If
         Select Case Combo5
            Case strSort2
               If Combo6 = MsgText(1) Then
                  strOrder2 = ", a0102 asc"
               Else
                  strOrder2 = ", a0102 desc"
               End If
            Case Else
               strOrder2 = MsgText(601)
         End Select
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by a0102 asc"
         Else
            strOrder1 = " order by a0102 desc"
         End If
         Select Case Combo5
            Case strSort1
               If Combo6 = MsgText(1) Then
                  strOrder2 = ", a0405 asc"
               Else
                  strOrder2 = ", a0405 desc"
               End If
            Case Else
               strOrder2 = MsgText(601)
         End Select
      Case Else
         strOrder1 = MsgText(601)
         strOrder2 = MsgText(601)
   End Select
   adoaccrpt405.CursorLocation = adUseClient
   adoaccrpt405.Open "select * from accrpt405", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc040M.CursorLocation = adUseClient
   '20140123START Modify By eric
   If Text6 <> MsgText(601) Then
      strSql = " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
      Text66 = IIf(Text6 = "2", "J", "1")
   End If
   'If Text6 <> MsgText(601) Then
   '   strSql = " and a0403 = '" & Text6 & "'"
   'End If
   '20140123END
   
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0401 = " & Val(Text3) & ""
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0402 = " & Val(Text1) & ""
   End If
   '20140123START Modify By eric
   If Text6 <> MsgText(601) Then
      adoacc040M.Open "select a0101, a0102, a0103 from acc010 where a0101 not in ('9998', '9999') and (a0109 is null or a0109='" & Text66 & "') order by a0101", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoacc040M.Open "select a0101, a0102, a0103 from acc010 where a0101 not in ('9998', '9999') order by a0101", adoTaie, adOpenStatic, adLockReadOnly
   End If
   'adoacc040M.Open "select a0101, a0102, a0103 from acc010 where a0101 not in ('9998', '9999') order by a0101", adoTaie, adOpenStatic, adLockReadOnly
   '20140123END
   
   If adoacc040M.RecordCount = 0 Then
      adoacc040M.Close
      adoaccrpt405.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc040M.EOF = False
      adoaccrpt405.AddNew
      adoaccrpt405.Fields("r40501").Value = strUserNum
      adoaccrpt405.Fields("r40502").Value = adoacc040M.Fields(0).Value
      If IsNull(adoacc040M.Fields(1).Value) Then
         adoaccrpt405.Fields("r40503").Value = Null
      Else
         adoaccrpt405.Fields("r40503").Value = adoacc040M.Fields(1).Value
      End If
      adoacc040MM.CursorLocation = adUseClient
      adoacc040MM.Open "select sum(a0406), sum(a0407) from acc040 where a0405 = '" & adoacc040M.Fields(0).Value & "' AND A0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc040MM.RecordCount <> 0 Then
         If IsNull(adoacc040MM.Fields(0).Value) Then
            adoaccrpt405.Fields("r40504").Value = 0
         Else
            adoaccrpt405.Fields("r40504").Value = Val(adoacc040MM.Fields(0).Value)
         End If
         If IsNull(adoacc040MM.Fields(1).Value) Then
            adoaccrpt405.Fields("r40505").Value = 0
         Else
            adoaccrpt405.Fields("r40505").Value = Val(adoacc040MM.Fields(1).Value)
         End If
      Else
         adoaccrpt405.Fields("r40504").Value = 0
         adoaccrpt405.Fields("r40505").Value = 0
      End If
      adoacc040MM.Close
      adoacc040Y.CursorLocation = adUseClient
      '20140123START Modify By eric
      If Text6 <> MsgText(601) Then
         adoacc040Y.Open "select a0405, a0103, sum(a0406), sum(a0407) from acc040, acc010 where a0405 = a0101 (+) and a0401 = " & Val(Text3) & "  " & _
                         " and a0401||decode(length(a0402), 1, '0'||a0402, 2, a0402) <= " & Val(Text3) & IIf(Len(Text1) < 2, "0" & Text1, Text1) & " and a0405 = '" & adoacc040M.Fields(0).Value & "'  " & _
                         " and a0403 = '" & Text66 & "' AND A0404 = '" & MsgText(55) & "' and (a0109 is null or a0109='" & Text66 & "') group by a0405, a0103", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Else
         adoacc040Y.Open "select a0405, a0103, sum(a0406), sum(a0407) from acc040, acc010 where a0405 = a0101 (+) and a0401 = " & Val(Text3) & "  " & _
                         " and a0401||decode(length(a0402), 1, '0'||a0402, 2, a0402) <= " & Val(Text3) & IIf(Len(Text1) < 2, "0" & Text1, Text1) & " and a0405 = '" & adoacc040M.Fields(0).Value & "'  " & _
                         " AND A0404 = '" & MsgText(55) & "' group by a0405, a0103", adoTaie, adOpenDynamic, adLockBatchOptimistic
      End If
      'adoacc040Y.Open "select a0405, a0103, sum(a0406), sum(a0407) from acc040, acc010 where a0405 = a0101 (+) and a0401 = " & Val(Text3) & " and a0401||decode(length(a0402), 1, '0'||a0402, 2, a0402) <= " & Val(Text3) & IIf(Len(Text1) < 2, "0" & Text1, Text1) & " and a0405 = '" & adoacc040M.Fields(0).Value & "' and a0403 = '" & Text6 & "' AND A0404 = '" & MsgText(55) & "' group by a0405, a0103", adoTaie, adOpenDynamic, adLockBatchOptimistic
      '20140123END
      
      If adoacc040Y.RecordCount <> 0 Then
         If IsNull(adoacc040Y.Fields(2).Value) Then
            adoaccrpt405.Fields("r40506").Value = 0
         Else
            adoaccrpt405.Fields("r40506").Value = Val(adoacc040Y.Fields(2).Value)
         End If
         If IsNull(adoacc040Y.Fields(3).Value) Then
            adoaccrpt405.Fields("r40507").Value = 0
         Else
            adoaccrpt405.Fields("r40507").Value = Val(adoacc040Y.Fields(3).Value)
         End If
         If adoacc040Y.Fields(1).Value = "1" Then
            adoaccrpt405.Fields("r40508").Value = Val(adoaccrpt405.Fields("r40506").Value) - Val(adoaccrpt405.Fields("r40507").Value)
         Else
            adoaccrpt405.Fields("r40508").Value = Val(adoaccrpt405.Fields("r40507").Value) - Val(adoaccrpt405.Fields("r40506").Value)
         End If
      Else
         adoaccrpt405.Fields("r40506").Value = 0
         adoaccrpt405.Fields("r40507").Value = 0
         adoaccrpt405.Fields("r40508").Value = 0
      End If
      adoacc040Y.Close
      adoacc040Y.CursorLocation = adUseClient
      '20140123START Modify By eric
      If Text6 <> MsgText(601) Then
         adoacc040Y.Open "select a0405, a0103, sum(a0406), sum(a0407), sum(a0408) from acc040, acc010 where a0405 = a0101 (+) and a0401 = " & Val(Text3) & "  " & _
                         " and a0402 = " & Val(Text1) & " and a0405 = '" & adoacc040M.Fields(0).Value & "' and a0403 = '" & Text66 & "' AND A0404 = '" & MsgText(55) & "'  " & _
                         " and (a0109 is null or a0109='" & Text66 & "') group by a0405, a0103", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Else
         adoacc040Y.Open "select a0405, a0103, sum(a0406), sum(a0407), sum(a0408) from acc040, acc010 where a0405 = a0101 (+) and a0401 = " & Val(Text3) & " " & _
                         " and a0402 = " & Val(Text1) & " and a0405 = '" & adoacc040M.Fields(0).Value & "' AND A0404 = '" & MsgText(55) & "' group by a0405, a0103", adoTaie, adOpenDynamic, adLockBatchOptimistic
      End If
      'adoacc040Y.Open "select a0405, a0103, sum(a0406), sum(a0407), sum(a0408) from acc040, acc010 where a0405 = a0101 (+) and a0401 = " & Val(Text3) & " and a0402 = " & Val(Text1) & " and a0405 = '" & adoacc040M.Fields(0).Value & "' and a0403 = '" & Text6 & "' AND A0404 = '" & MsgText(55) & "' group by a0405, a0103", adoTaie, adOpenDynamic, adLockBatchOptimistic
      '20140123END
      If adoacc040Y.RecordCount <> 0 Then
         If IsNull(adoacc040Y.Fields(4).Value) = False Then
            adoaccrpt405.Fields("r40509").Value = Val(adoacc040Y.Fields(4).Value)
         Else
            adoaccrpt405.Fields("r40509").Value = 0
         End If
      Else
         adoaccrpt405.Fields("r40509").Value = 0
      End If
      adoacc040Y.Close
      intCounter = intCounter + 1
      adoaccrpt405.Fields("r40510").Value = intCounter
      adoaccrpt405.UpdateBatch
      adoacc040M.MoveNext
   Loop
   adoacc040M.Close
   adoaccrpt405.Close
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
Private Sub Accrpt405Delete()
   adoTaie.Execute "delete from accrpt405"
End Sub

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "科目代號"
   strSort2 = "科目名稱"
   Combo13.AddItem strSort1
   Combo13.AddItem strSort2
   Combo5.AddItem strSort1
   Combo5.AddItem strSort2
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text6 = ""
   '20140123START Modify By eric
   Text7 = "台一　專利商標/智權"
   'Text7 = ""
   '20140123END
   Text3 = ""
   Text1 = ""
   Combo13 = ""
   Combo5 = ""
   Text6.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'20140123START Add By eric
Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'20140123START By eric
Private Sub Text6_LostFocus()
   If Text6.Text <> "1" And Text6.Text <> "2" And Text6.Text <> "" Then
      MsgBox "公司別僅可為 1 / 2 或不輸入  ! "
      Text6.Text = ""
      Text6.SetFocus
      Exit Sub
   End If
   If Text6 = "" Then
      Text7 = "台一　專利商標/智權"
   End If
   
End Sub
