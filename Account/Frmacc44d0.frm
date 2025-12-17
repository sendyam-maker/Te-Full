VERSION 5.00
Begin VB.Form Frmacc44d0 
   AutoRedraw      =   -1  'True
   Caption         =   "年度部門綜合損益統計表"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   5160
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
      Height          =   300
      Left            =   1920
      TabIndex        =   10
      Top             =   720
      Width           =   2892
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
      TabIndex        =   1
      Top             =   720
      Width           =   612
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
      Left            =   4680
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   210
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
      TabIndex        =   4
      Top             =   1680
      Width           =   4692
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
      TabIndex        =   2
      Top             =   1200
      Width           =   612
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
      Height          =   300
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label4 
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
      TabIndex        =   8
      Top             =   1200
      Width           =   732
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "半年期"
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
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   732
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
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
      Top             =   720
      Width           =   732
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別              (1.台一 2.智權 空白.全部)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   4245
   End
End
Attribute VB_Name = "Frmacc44d0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit
Public adoacc010 As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoaccrpt413 As New ADODB.Recordset
Dim lngCounter As Long
Dim douTotal1(8) As Double
Dim douTotal2(8) As Double
Dim douTotal3(8) As Double
Dim douTotal4(8) As Double
Dim douTotal5(8) As Double
Dim dllaccrpt414 As Object

Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt413Delete
   ProduceData
   '2014/2/20 modify by sonia
   'dllaccrpt414.Acc44d0 ReportTitle(414), Text6, Text7, Text1, Text2, Text3, Combo3, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   dllaccrpt414.Acc44d0 ReportTitle(414), IIf(Text6 = "2", "J", Text6), IIf(Text7 = "", "台一　專利商標/智權", Text7), Text1, Text2, Text3, Combo3, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
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
   Me.Height = 2700
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Combo3.AddItem ComboItem(151)
   Combo3.AddItem ComboItem(152)
   Combo3 = ComboItem(151)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt414 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt414 = Nothing
   Set Frmacc44d0 = Nothing
End Sub

Private Sub Text1_Change()
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   Text2 = A0902Query(Text1)
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text6_Change()
   '2014/2/20 modify by sonia
   'If Text6 = MsgText(601) Then
   '   Exit Sub
   'End If
   'Text7 = A0802Query(Text6)
   Select Case Text6
      Case "1"
         Text7 = A0802Query(Text6)
      Case "2"
         Text7 = A0802Query("J")
      Case ""
         Text7 = "台一　專利商標/智權"
   End Select
   '2014/2/20 end
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim intCounter As Integer
Dim str9997 As String   'add by sonia 2016/1/28

On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   lngCounter = 0
   adoaccrpt413.CursorLocation = adUseClient
   adoaccrpt413.Open "select * from accrpt413", adoTaie, adOpenDynamic, adLockBatchOptimistic
'-------------------------------------------------
' 實際營業收入
'-------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/20 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If Text6 <> "" Then
      adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2014/2/20 end
   'end 2007/12/19
   Do While adoacc010.EOF = False
      If adoaccrpt413.RecordCount = 0 Then
         adoaccrpt413.AddNew
         adoaccrpt413.Fields("r41301").Value = strUserNum
         adoaccrpt413.UpdateBatch
      End If
      Accrpt413Save
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   PaintLine ReportSum(4)
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   adoaccrpt413.Fields("r41303").Value = ReportSum(14)
   If Mid(Combo3, 1, 1) = "1" Then
      Calculate "4", "499999", Text1, 1, 6
   Else
      Calculate "4", "499999", Text1, 7, 12
   End If
   For intCounter = 3 To 8
      If IsNull(adoaccrpt413.Fields(intCounter).Value) Then
         douTotal1(intCounter) = 0
      Else
         douTotal1(intCounter) = Val(adoaccrpt413.Fields(intCounter).Value)
      End If
   Next intCounter
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
      
   'Add By Cheng 2002/01/18
   PaintLine ReportSum(8)
   
   adoaccrpt413.UpdateBatch
'-------------------------------------------------
' 實際營業支出
'-------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/20 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If Text6 <> "" Then
      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2014/2/20 end
   'end 2007/12/19
   Do While adoacc010.EOF = False
      Accrpt413Save
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   PaintLine ReportSum(4)
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   adoaccrpt413.Fields("r41303").Value = ReportSum(15)
   If Mid(Combo3, 1, 1) = "1" Then
      Calculate "6", "699999", Text1, 1, 6
   Else
      Calculate "6", "699999", Text1, 7, 12
   End If
   For intCounter = 3 To 8
      If IsNull(adoaccrpt413.Fields(intCounter).Value) Then
         douTotal2(intCounter) = 0
      Else
         douTotal2(intCounter) = Val(adoaccrpt413.Fields(intCounter).Value)
      End If
   Next intCounter
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   
   'Add By Cheng 2002/01/18
   PaintLine ReportSum(8)
   
   adoaccrpt413.UpdateBatch
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   PaintLine ReportSum(4)
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   adoaccrpt413.Fields("r41303").Value = ReportSum(16)
   For intCounter = 3 To 8
      If douTotal1(intCounter) - douTotal2(intCounter) = 0 Then
         adoaccrpt413.Fields(intCounter).Value = Null
      Else
         adoaccrpt413.Fields(intCounter).Value = douTotal1(intCounter) - douTotal2(intCounter)
      End If
   Next intCounter
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   
   'Add By Cheng 2002/01/18
   PaintLine ReportSum(8)
   
   adoaccrpt413.UpdateBatch
'-------------------------------------------------
' 分攤管理、智權部門費用
'-------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   'modify by sonia 2016/1/28 105年起才有9997分攤法務部門費用科目
   'adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   str9997 = ""
   If Val(Text3) < 105 Then
      str9997 = " and a0101<>'9997' "
   End If
   adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and a0104 = '3' and instr(a0102,'不用')=0" & str9997 & " order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   'end 2016/1/28
   'end 2007/12/19
   Do While adoacc010.EOF = False
      Accrpt413Save
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   PaintLine ReportSum(4)
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   adoaccrpt413.Fields("r41303").Value = ReportSum(17)
   If Mid(Combo3, 1, 1) = "1" Then
      Calculate "9", "999999", Text1, 1, 6
   Else
      Calculate "9", "999999", Text1, 7, 12
   End If
   For intCounter = 3 To 8
      If IsNull(adoaccrpt413.Fields(intCounter).Value) Then
         douTotal3(intCounter) = 0
      Else
         douTotal3(intCounter) = Val(adoaccrpt413.Fields(intCounter).Value)
      End If
   Next intCounter
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   
   'Add By Cheng 2002/01/18
   PaintLine ReportSum(8)
   
   adoaccrpt413.UpdateBatch
   
'-------------------------------------------------
' 營業外收入
'-------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/20 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If Text6 <> "" Then
      adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2014/2/20 end
   'end 2007/12/19
   Do While adoacc010.EOF = False
      Accrpt413Save
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   PaintLine ReportSum(4)
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   adoaccrpt413.Fields("r41303").Value = ReportSum(5)
   If Mid(Combo3, 1, 1) = "1" Then
      Calculate "71", "719999", Text1, 1, 6
   Else
      Calculate "71", "719999", Text1, 7, 12
   End If
   For intCounter = 3 To 8
      If IsNull(adoaccrpt413.Fields(intCounter).Value) Then
         douTotal4(intCounter) = 0
      Else
         douTotal4(intCounter) = Val(adoaccrpt413.Fields(intCounter).Value)
      End If
   Next intCounter
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   
   'Add By Cheng 2002/01/18
   PaintLine ReportSum(8)
   
   adoaccrpt413.UpdateBatch
   
'-------------------------------------------------
' 營業外支出
'-------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '73' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/20 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '73' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If Text6 <> "" Then
      adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '73' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '73' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2014/2/20 end
   'end 2007/12/19
   Do While adoacc010.EOF = False
      Accrpt413Save
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   PaintLine ReportSum(4)
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   adoaccrpt413.Fields("r41303").Value = ReportSum(6)
   If Mid(Combo3, 1, 1) = "1" Then
      Calculate "72", "729999", Text1, 1, 6
   Else
      Calculate "72", "729999", Text1, 7, 12
   End If
   For intCounter = 3 To 8
      If IsNull(adoaccrpt413.Fields(intCounter).Value) Then
         douTotal5(intCounter) = 0
      Else
         douTotal5(intCounter) = Val(adoaccrpt413.Fields(intCounter).Value)
      End If
   Next intCounter
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   
   'Add By Cheng 2002/01/18
   PaintLine ReportSum(8)
   
   adoaccrpt413.UpdateBatch
   
   
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   PaintLine ReportSum(4)
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   adoaccrpt413.Fields("r41303").Value = ReportSum(21)
   For intCounter = 3 To 8
      If douTotal1(intCounter) - douTotal2(intCounter) - douTotal3(intCounter) = 0 Then
         adoaccrpt413.Fields(intCounter).Value = Null
      Else
         adoaccrpt413.Fields(intCounter).Value = douTotal1(intCounter) - douTotal2(intCounter) - douTotal3(intCounter) + douTotal4(intCounter) - douTotal5(intCounter)
      End If
   Next intCounter
   adoaccrpt413.UpdateBatch
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   PaintLine ReportSum(8)
   adoaccrpt413.UpdateBatch
   adoaccrpt413.Close
   adoTaie.Execute "delete from accrpt413 where r41302 is null"
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
Private Sub Accrpt413Delete()
   adoTaie.Execute "delete from accrpt413"
End Sub

'*************************************************
'  儲存資料表(部門損益比較表暫存檔)
'
'*************************************************
Private Sub Accrpt413Save()
Dim intCounter As Integer
      
   adoaccrpt413.AddNew
   adoaccrpt413.Fields("r41301").Value = strUserNum
   adoaccrpt413.Fields("r41302").Value = Counter
   If IsNull(adoacc010.Fields("a0102").Value) Then
      adoaccrpt413.Fields("r41303").Value = Null
   Else
      adoaccrpt413.Fields("r41303").Value = adoacc010.Fields("a0102").Value
   End If
   If Mid(Combo3, 1, 1) = "1" Then
      Calculate adoacc010.Fields("a0101").Value, adoacc010.Fields("a0101").Value, Text1, 1, 6
   Else
      Calculate adoacc010.Fields("a0101").Value, adoacc010.Fields("a0101").Value, Text1, 7, 12
   End If
   adoaccrpt413.UpdateBatch
End Sub

'*************************************************
'  計算各月份小計金額
'
'*************************************************
Private Sub Calculate(strAccNo1 As String, strAccNo2 As String, strDeptNo As String, intStartM, intEndM As Integer)
Dim douDebit, douCredit As Double
Dim intCounter, intMonth As Integer, strSql As String
      
   If Text3 <> MsgText(601) Then
      strSql = " and a0401 = " & Val(Text3) & ""
   End If
   If Text6 <> MsgText(601) Then
      '2014/2/20 modify by sonia
      'strSql = strSql & " and a0403 = '" & Text6 & "'"
      strSql = strSql & " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
      '2014/2/20 end
   End If
   If strDeptNo <> MsgText(601) Then
      'add by sonia 2016/1/28 105年起法務及投法合併,費用傳票輸在L部門
      If Val(Text3) >= 105 And strDeptNo = "L" Then
         strSql = strSql & " and a0404 in ('" & strDeptNo & "','CFL','FCL')"
      'end 2016/1/28
      'MODIFY BY SONIA 2013/11/7 102/10 CFL的416102會少計算到
      'strSql = strSql & " and a0404 = '" & strDeptNo & "'"
      ElseIf strDeptNo = "FCL" Then
         strSql = strSql & " and a0404 in ('" & strDeptNo & "','CFL')"
      Else
         strSql = strSql & " and a0404 = '" & strDeptNo & "'"
      End If
      '2013/11/7 end
   Else
      strSql = strSql & " and a0404 = '" & MsgText(55) & "'"
   End If
   If strAccNo1 <> MsgText(601) Then
      strSql = strSql & " and substr(a0405, 1, 4) >= '" & strAccNo1 & "'"
   End If
   If strAccNo2 <> MsgText(601) Then
      strSql = strSql & " and substr(a0405, 1, 4) <= '" & strAccNo2 & "'"
   End If
   intCounter = 3
   For intMonth = intStartM To intEndM
      adoacc040.CursorLocation = adUseClient
      adoacc040.Open "select sum(a0408) from acc040 where a0402 = " & intMonth & "" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc040.RecordCount <> 0 Then
         If IsNull(adoacc040.Fields(0).Value) Then
            adoaccrpt413.Fields(intCounter).Value = Null
         Else
            If adoacc040.Fields(0).Value = 0 Then
               adoaccrpt413.Fields(intCounter).Value = Null
            Else
               adoaccrpt413.Fields(intCounter).Value = adoacc040.Fields(0).Value
            End If
         End If
      Else
         adoaccrpt413.Fields(intCounter).Value = Null
      End If
      intCounter = intCounter + 1
      adoacc040.Close
   Next intMonth
End Sub

'*************************************************
'  畫線
'
'*************************************************
Private Sub PaintLine(strSign As String)
Dim intCounter As Integer

   For intCounter = 3 To 8
      adoaccrpt413.Fields(intCounter).Value = strSign
   Next intCounter
End Sub

'*************************************************
'  計數器
'
'*************************************************
Private Function Counter() As Long
   lngCounter = lngCounter + 1
   Counter = lngCounter
End Function

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text6 = ""
   Text7 = ""
   Text1 = ""
   Text2 = ""
   Text3 = ""
   'Combo3 = ""  '2014/2/20 cancel by sonia
   Text6.SetFocus
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
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

'2014/2/20 add by sonia
Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   End If
End Sub
'2014/2/20 end

