VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc44b0 
   AutoRedraw      =   -1  'True
   Caption         =   "部門綜合損益表(子科目)"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1815
   ScaleWidth      =   5160
   Begin VB.CommandButton Cmd_Excel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel"
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
      Left            =   2640
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   1200
      Width           =   2300
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
      Left            =   4800
      TabIndex        =   7
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
      TabIndex        =   3
      Top             =   1200
      Width           =   2300
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   252
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
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "年月"
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
      Top             =   720
      Width           =   612
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc44b0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc010 As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoaccrpt412 As New ADODB.Recordset
Public adoacc090 As New ADODB.Recordset
Dim lngCounter As Long
'Modify by Amy 2016/07/27 douTotalX 改名稱及型態 原Double
Dim stTotal1(13) As String
Dim stTotal2(13) As String
Dim stTotal3(13) As String
Dim stTotal4(13) As String
'end 2016/07/27
Dim dllaccrpt412 As Object
Dim strFieldN(), intWidth()  'Add by amy 2015/03/13
'Added by Lydia 2016/01/30 列印使用
Dim strTemp(0 To 11) As String
Dim PLeft(0 To 12) As Integer
Dim PTitle(0 To 11) As String
Private Const ciTitleFontSize = 14
Private Const ciFontSize = 10
Private Const ciStartX = 0
Private Const ciStartY = 500
Private Const ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim iPrint As Integer, iPage As Integer

Private Sub Cmd_Excel_Click()
    If FormCheck = False Then
        MsgBox MsgText(181), , MsgText(5)
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Accrpt412Delete
    ProduceData
    If adoaccrpt412.State = adStateOpen Then
        adoaccrpt412.Close
    End If
    adoaccrpt412.CursorLocation = adUseClient
    adoaccrpt412.Open "Select * From accrpt412 Where r41201='" & strUserNum & "' Order by r41202", adoTaie, adOpenStatic, adLockReadOnly
    If adoaccrpt412.RecordCount <> 0 Then
        ExcelSave
    End If
    If adoaccrpt412.State = adStateOpen Then
        adoaccrpt412.Close
    End If
    Screen.MousePointer = vbDefault
    FormClear
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

Private Sub Command1_Click()
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Accrpt412Delete
   ProduceData
   '2014/2/20 modify by sonia
   'dllaccrpt412.Acc44b0 ReportTitle(412), Text6, Text7, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
   'Modify by Amy 2015/04/15 財務處可能同時兩個人執行此報表,造成資料錯誤 +strUserNum
   'Modified by Lydia 2016/01/30 改成Printer
   'dllaccrpt412.Acc44b0 ReportTitle(412), IIf(Text6 = "2", "J", Text6), IIf(Text7 = "", "台一　專利商標/智權", Text7), MaskEdBox1.Text, MaskEdBox2.Text, strUserNum & "-" & StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
    adoaccrpt412.Open "select * from accrpt412 Where r41201='" & strUserNum & "' order by r41202 ", adoTaie, adOpenStatic, adLockReadOnly
    If adoaccrpt412.RecordCount <> 0 Then
        PrintData
    End If
    If adoaccrpt412.State = adStateOpen Then
        adoaccrpt412.Close
    End If
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
   Me.Height = 2200
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   MaskEdBox2.Mask = Mid(DFormat, 1, 6)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt412 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt412 = Nothing
   Set Frmacc44b0 = Nothing
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
   adoaccrpt412.CursorLocation = adUseClient
   'Modify by Amy 2015/04/15 財務處可能同時兩個人執行此報表,造成資料錯誤
   adoaccrpt412.Open "select * from accrpt412 Where r41201='" & strUserNum & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'-------------------------------------------------
' 實際營業收入
'-------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/20 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If Text6 <> "" Then
      adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 >= '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 >= '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2014/2/20 END
   'end 2007/12/19
   Do While adoacc010.EOF = False
      Accrpt412Save
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   PaintLine ReportSum(4), "4" 'Modify by Amy 2015/03/13 分隔線,+會計科目
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41203").Value = ReportSum(14)
   adoaccrpt412.Fields("r41215").Value = "4S" 'Add by Amy 2015/03/13
   For intCounter = 3 To 13
      adoaccrpt412.Fields(intCounter).Value = 0
   Next intCounter
   Calculate "4", "499999"
'   adoaccrpt412.Fields(12).Value = 0
   For intCounter = 3 To 13
      If IsNull(adoaccrpt412.Fields(intCounter).Value) Then
         stTotal1(intCounter) = "0"
      Else
         stTotal1(intCounter) = Val(adoaccrpt412.Fields(intCounter).Value)
'         adoaccrpt412.Fields(12).Value = Val(adoaccrpt412.Fields(12).Value) + Val(adoaccrpt412.Fields(intCounter).Value)
      End If
   Next intCounter
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41215").Value = "4E" 'Add by amy 2015/03/13 下方虛線
   adoaccrpt412.UpdateBatch
'-------------------------------------------------
' 實際營業支出
'-------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/20 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If Text6 <> "" Then
      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 >= '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 >= '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2014/2/20 END
   'end 2007/12/19
   Do While adoacc010.EOF = False
      Accrpt412Save
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   PaintLine ReportSum(4), "6" 'Modify by Amy 2015/03/13 分隔線,+會計科目
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41203").Value = ReportSum(15)
   adoaccrpt412.Fields("r41215").Value = "6S" 'Add by Amy 2015/03/13
   For intCounter = 3 To 13
      adoaccrpt412.Fields(intCounter).Value = 0
   Next intCounter
   Calculate "6", "699999"
'   adoaccrpt412.Fields(12).Value = 0
   For intCounter = 3 To 13
      If IsNull(adoaccrpt412.Fields(intCounter).Value) Then
         stTotal2(intCounter) = "0"
      Else
         stTotal2(intCounter) = Val(adoaccrpt412.Fields(intCounter).Value)
'         adoaccrpt412.Fields(12).Value = Val(adoaccrpt412.Fields(12).Value) + Val(adoaccrpt412.Fields(intCounter).Value)
      End If
   Next intCounter
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.UpdateBatch
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   PaintLine ReportSum(4), "6" 'Modify by Amy 2015/03/13 分隔線,+會計科目
   adoaccrpt412.Fields("r41215").Value = "6E" 'Add by Amy 2015/03/13
   adoaccrpt412.UpdateBatch
   'Modify by Amy 2016/07/27 避免報表與Excel 值不同合計可能誤差,故資料先四捨五入到小數2位
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41203").Value = ReportSum(16)
   adoaccrpt412.Fields("r41215").Value = "DS" 'Add by Amy 2015/03/13 部門損益
'   adoaccrpt412.Fields(12).Value = 0
   For intCounter = 3 To 13
      If Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)) = 0 Then
         adoaccrpt412.Fields(intCounter).Value = "0"
      Else
         adoaccrpt412.Fields(intCounter).Value = Format(Round(Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)), 2), FAmount)
'         adoaccrpt412.Fields(12).Value = Val(adoaccrpt412.Fields(12).Value) + Val(adoaccrpt412.Fields(intCounter).Value)
      End If
   Next intCounter
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41215").Value = "DE" 'Add by Amy 2015/03/13
   adoaccrpt412.UpdateBatch
'-------------------------------------------------
' 分攤管理、智權部門費用
' Memo 2016/07/27 婧瑄:分攤科目改為公式顯示(參閱 UpdAccrpt412)
'-------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/20 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   'modify by sonia 2016/1/28 105年起才有9997分攤法務部門費用科目
   'If Text6 <> "" Then
   '   adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Else
   '   adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   'End If
   str9997 = ""
   If Val(Mid(MaskEdBox1.Text, 1, 3)) < 105 Then
      str9997 = " and a0101<>'9997' "
   End If
   If Text6 <> "" Then
      adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and a0104 >= '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "')" & str9997 & " order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '999' and a0101 <= '999999' and a0104 >= '3' and instr(a0102,'不用')=0" & str9997 & " order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   'end 2016/1/28
   '2014/2/20 END
   'end 2007/12/19
   Do While adoacc010.EOF = False
      Accrpt412Save
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   PaintLine ReportSum(4), "9" 'Modify by Amy 2015/03/13 分隔線,+會計科目
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41203").Value = ReportSum(17)
   adoaccrpt412.Fields("r41215").Value = "9S" 'Add by Amy 2015/03/13
   For intCounter = 3 To 12
      adoaccrpt412.Fields(intCounter).Value = 0
   Next intCounter
   Calculate "9", "999999"
   adoaccrpt412.Fields(12).Value = 0
   adoaccrpt412.Fields(13).Value = 0
   For intCounter = 3 To 11
      If IsNull(adoaccrpt412.Fields(intCounter).Value) Then
         stTotal3(intCounter) = "0"
      Else
         stTotal3(intCounter) = Val(adoaccrpt412.Fields(intCounter).Value)
         adoaccrpt412.Fields(12).Value = Format(Round(Val(adoaccrpt412.Fields(12).Value) + Val(adoaccrpt412.Fields(intCounter).Value), 2), FAmount)
      End If
   Next intCounter
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41215").Value = "9E" 'Add by Amy 2015/03/13
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   PaintLine ReportSum(4), "9E" 'Modify by Amy 2015/03/13 分隔線,+會計科目
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41203").Value = ReportSum(18)
   adoaccrpt412.Fields("r41215").Value = "VS" 'Add by Amy 2015/03/13
'   adoaccrpt412.Fields(12).Value = 0
   For intCounter = 3 To 13
      If Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)) - Val(stTotal3(intCounter)) = 0 Then
         adoaccrpt412.Fields(intCounter).Value = 0
      Else
         adoaccrpt412.Fields(intCounter).Value = Format(Round(Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)) - Val(stTotal3(intCounter)), 2))
'         adoaccrpt412.Fields(12).Value = Val(adoaccrpt412.Fields(12).Value) + douTotal1(intCounter) - douTotal2(intCounter) - douTotal3(intCounter)
      End If
   Next intCounter
   adoaccrpt412.Fields(11).Value = 0
   adoaccrpt412.Fields(13).Value = 0   '2006/6/18 ADD BY SONIA 財務處說總所/管理之各部門營業損益印0
   'add by sonia 2016/7/13 105年起之法務部也是0
   If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
      adoaccrpt412.Fields(10).Value = 0
   End If
   'end 2016/7/13
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41215").Value = "VE" 'Add by Amy 2015/03/13
   adoaccrpt412.UpdateBatch
'-------------------------------------------------
' 營業外收入
'-------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '7' and a0101 < '8' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/20 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '7' and a0101 < '8' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'Modify by Amy 2016/07/27 拆成 營業外收入/支出拆開
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '7' and a0101 < '8' and a0104 >= '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 >= '7' and a0101 < '8' and a0104 >= '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
   If Text6 <> "" Then
      adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 >= '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 >= '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2014/2/20 END
   'end 2007/12/19
   Do While adoacc010.EOF = False
      Accrpt412Save
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   PaintLine ReportSum(4), "71" 'Modify by Amy 2015/03/13 分隔線,+會計科目
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41203").Value = ReportSum(5) 'Modify by Amy 2016/07//27 原:ReportSum(19)
   adoaccrpt412.Fields("r41215").Value = "71S" 'Add by Amy 2015/03/13
   For intCounter = 3 To 13
      adoaccrpt412.Fields(intCounter).Value = 0
   Next intCounter
   Calculate "71", "719999"
'   adoaccrpt412.Fields(12).Value = 0
   For intCounter = 3 To 13
      If IsNull(adoaccrpt412.Fields(intCounter).Value) Then
         stTotal4(intCounter) = "0"
      Else
         stTotal4(intCounter) = Val(adoaccrpt412.Fields(intCounter).Value)
'         adoaccrpt412.Fields(12).Value = Val(adoaccrpt412.Fields(12).Value) + Val(adoaccrpt412.Fields(intCounter).Value)
      End If
   Next intCounter
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41215").Value = "71E" 'Add by Amy 2015/03/13
   adoaccrpt412.UpdateBatch
   
'-------------------------------------------------
' 營業外支出
'-------------------------------------------------
 If Text6 <> "" Then
      adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '73' and a0104 >= '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '73' and a0104 >= '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   Do While adoacc010.EOF = False
      Accrpt412Save
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   PaintLine ReportSum(4), "72" '分隔線,+會計科目
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41203").Value = ReportSum(6) 'Modify by Amy 2016/07//27 原:ReportSum(19)
   adoaccrpt412.Fields("r41215").Value = "72S"
   For intCounter = 3 To 13
      adoaccrpt412.Fields(intCounter).Value = 0
   Next intCounter
   Calculate "72", "729999"
   For intCounter = 3 To 13
      If IsNull(adoaccrpt412.Fields(intCounter).Value) Then
         stTotal4(intCounter) = "0"
      Else
         stTotal4(intCounter) = Val(adoaccrpt412.Fields(intCounter).Value)
      End If
   Next intCounter
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41215").Value = "72E"
   adoaccrpt412.UpdateBatch
'end 2016/07/25
   
'-------------------------------------------------
' 計算損益
'-------------------------------------------------
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   PaintLine ReportSum(4)
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41203").Value = ReportSum(20)
   adoaccrpt412.Fields("r41215").Value = "ZZZZZZZZ" 'Add by Amy 2015/03/13
'   adoaccrpt412.Fields(12).Value = 0
   For intCounter = 3 To 13
      If Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)) - Val(stTotal3(intCounter)) + Val(stTotal4(intCounter)) = 0 Then
         adoaccrpt412.Fields(intCounter).Value = Null
      Else
         adoaccrpt412.Fields(intCounter).Value = Format(Round(Val(stTotal1(intCounter)) - Val(stTotal2(intCounter)) - Val(stTotal3(intCounter)) + Val(stTotal4(intCounter)), 2), FAmount)
'         adoaccrpt412.Fields(12).Value = Val(adoaccrpt412.Fields(12).Value) + douTotal1(intCounter) - douTotal2(intCounter) - douTotal3(intCounter) + douTotal4(intCounter)
      End If
   Next intCounter
   'end 2016/07/27
   'Add by Morgan 2006/9/6 管理部的全所損益=營業外收支
   If stTotal4(13) = 0 Then
      adoaccrpt412.Fields(13).Value = Null
   Else
      adoaccrpt412.Fields(13).Value = Val(stTotal4(13))
   End If
   'Modify by Amy 2015/03/13 智權部的全所損益=營業外收支
   'adoaccrpt412.Fields(11).Value = Null
   If stTotal4(11) = 0 Then
      adoaccrpt412.Fields(11).Value = Null
   Else
      adoaccrpt412.Fields(11).Value = Val(stTotal4(11))
   End If
   'add by sonia 2016/2/16 105年起之法務部的全所損益=營業外收支
   If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
      If stTotal4(10) = 0 Then
         adoaccrpt412.Fields(10).Value = Null
      Else
         adoaccrpt412.Fields(10).Value = Val(stTotal4(10))
      End If
   End If
   'end 2016/2/16
   
   adoaccrpt412.UpdateBatch
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   PaintLine ReportSum(8)
   adoaccrpt412.UpdateBatch
   adoaccrpt412.Close
   UpdAccrpt412 'Add by Amy 2016/07/25
   StatusClear
   Exit Sub
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   If adoacc010.State = adStateOpen Then adoacc010.Close
   If adoacc040.State = adStateOpen Then adoacc040.Close
   If adoaccrpt412.State = adStateOpen Then adoaccrpt412.Close
   If adoacc090.State = adStateOpen Then adoacc090.Close
   
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt412Delete()
  'Modify by Amy 2015/04/15 財務處可能同時兩個人執行此報表,造成資料錯誤
   adoTaie.Execute "delete from accrpt412 Where r41201='" & strUserNum & "'"
End Sub

'*************************************************
'  儲存資料表(部門損益比較表暫存檔)
'
'*************************************************
Private Sub Accrpt412Save()
Dim intCounter As Integer
      
   adoaccrpt412.AddNew
   adoaccrpt412.Fields("r41201").Value = strUserNum
   adoaccrpt412.Fields("r41202").Value = Counter
   adoaccrpt412.Fields("r41215").Value = "" & adoacc010.Fields("a0101") 'Add by Amy 2015/03/13 +會計科目代碼
  
   If IsNull(adoacc010.Fields("a0102").Value) Then
      adoaccrpt412.Fields("r41203").Value = Null
   Else
      adoaccrpt412.Fields("r41203").Value = adoacc010.Fields("a0102").Value
   End If
   For intCounter = 3 To 12
      adoaccrpt412.Fields(intCounter).Value = 0
   Next intCounter
   Calculate adoacc010.Fields("a0101").Value, adoacc010.Fields("a0101").Value
   If Mid(adoacc010.Fields("a0101").Value, 1, 1) = "9" Then
      adoaccrpt412.Fields("r41213").Value = 0
      For intCounter = 3 To 11
         If IsNull(adoaccrpt412.Fields(intCounter).Value) = False Then
            '
            adoaccrpt412.Fields("r41213").Value = Format(Round(Val(adoaccrpt412.Fields("r41213").Value) + Val(adoaccrpt412.Fields(intCounter).Value), 2), FAmount)
         End If
      Next intCounter
   End If
   adoaccrpt412.UpdateBatch
End Sub

'*************************************************
'  計算各部門小計金額
'
'*************************************************
Private Sub Calculate(strAccNo1 As String, strAccNo2 As String)
'Modify by Amy 2016/07/25 因只下 2.智權10501-06 會造成多重步驗操作…的錯誤
'Dim douDebit As Double
Dim strDebit As String
'end 2016/07/25
Dim intCounter As Integer
Dim strSql As String
Dim i
      
   intCounter = 3
   adoacc090.CursorLocation = adUseClient
   adoacc090.Open "select * from acc090 where a0904 = 'Y' order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc090.EOF = False
      adoacc040.CursorLocation = adUseClient
      strSql = MsgText(601)
      If Text6 <> MsgText(601) Then
         '2014/2/20 modify by sonia
         'strSql = " and a0403 = '" & Text6 & "'"
         strSql = " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
         '2014/2/20 END
      End If
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
         strSql = strSql & " and decode(length(a0402), 1, (a0401 * 10) || a0402, 2, a0401 || a0402)  >= " & Val(Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2)) & ""
      End If
      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
         strSql = strSql & " and decode(length(a0402), 1, (a0401 * 10) || a0402, 2, a0401 || a0402) <= " & Val(Mid(MaskEdBox2.Text, 1, 3) & Mid(MaskEdBox2.Text, 5, 2)) & ""
      End If
      If strAccNo1 <> MsgText(601) Then
         strSql = strSql & " and substr(a0405, 1, 4) >= '" & strAccNo1 & "'"
      End If
      If strAccNo2 <> MsgText(601) Then
         strSql = strSql & " and substr(a0405, 1, 4) <= '" & strAccNo2 & "'"
      End If
      '2008/1/14 modify by sonia 科目7XXX營業外收支應區分借貸方科目
      'adoacc040.Open "select sum(a0408) from acc040 where a0404 = '" & adoacc090.Fields("a0901").Value & "'" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
      If strAccNo1 >= "7" And strAccNo2 <= "799999" Then
         adoacc040.Open "select sum(decode(a0103,'1',a0408*-1,a0408)) from acc040,acc010 where a0405=a0101(+) and a0404 = '" & adoacc090.Fields("a0901").Value & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      Else
         'add by sonia 2016/1/28 105年起法務及投法合併,費用傳票輸在L部門,但此表部門名稱改法務部,放在原投法的位置
         If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
            If adoacc090.Fields("a0901").Value = "L" Then
               adoacc040.Open "select sum(a0408) from acc040 where a0404 in ('" & adoacc090.Fields("a0901").Value & "','FCL','CFL')" & strSql, adoTaie, adOpenStatic, adLockReadOnly
            Else
               adoacc040.Open "select sum(a0408) from acc040 where a0404 = '" & adoacc090.Fields("a0901").Value & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
            End If
         'end 2016/1/28
         'MODIFY BY SONIA 2013/11/7 102/10 CFL的416102會因為此段最下方的計算跑到總所/管理
         'adoacc040.Open "select sum(a0408) from acc040 where a0404 = '" & adoacc090.Fields("a0901").Value & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
         ElseIf adoacc090.Fields("a0901").Value = "FCL" Then
            adoacc040.Open "select sum(a0408) from acc040 where a0404 in ('" & adoacc090.Fields("a0901").Value & "','CFL')" & strSql, adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoacc040.Open "select sum(a0408) from acc040 where a0404 = '" & adoacc090.Fields("a0901").Value & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
         End If
         '2013/11/7 end
      End If
      '2008/1/14 end
      'Modify by Amy 2016/07/25 改 douDebit為strDebit
      If adoacc040.RecordCount <> 0 Then
         If IsNull(adoacc040.Fields(0).Value) Then
            strDebit = "0"
         Else
            strDebit = Format(adoacc040.Fields(0).Value, FAmount)
         End If
         Select Case adoacc090.Fields("a0901").Value
            Case "P"
               adoaccrpt412.Fields(3).Value = strDebit
            Case "T"
               adoaccrpt412.Fields(4).Value = strDebit
            Case "L"
               'modify by sonia 2016/1/28 105年起法務及投法合併,費用傳票輸在L部門,但此表部門名稱改法務部,放在原投法的位置
               'adoaccrpt412.Fields(5).Value = strdebit
               If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
                  adoaccrpt412.Fields(10).Value = strDebit
               Else
                  adoaccrpt412.Fields(5).Value = strDebit
               End If
               'end 2016/1/28
            Case "CFP"
               adoaccrpt412.Fields(6).Value = strDebit
            Case "CFT"
               adoaccrpt412.Fields(7).Value = strDebit
            Case "FCP"
               adoaccrpt412.Fields(8).Value = strDebit
            Case "FCT"
               adoaccrpt412.Fields(9).Value = strDebit
            Case "FCL"
               adoaccrpt412.Fields(10).Value = strDebit
            Case "SAL"
               adoaccrpt412.Fields(11).Value = strDebit
            Case "TOT"
               adoaccrpt412.Fields(12).Value = strDebit
            Case "M"
               adoaccrpt412.Fields(13).Value = strDebit
         End Select
      End If
      'end 2016/07/25
      adoacc040.Close
      adoacc090.MoveNext
   Loop
   Select Case Mid(strAccNo1, 1, 1)
      Case "6"
      Case "9"
         adoaccrpt412.Fields(13).Value = 0
      Case Else
         adoaccrpt412.Fields(13).Value = 0
         For intCounter = 3 To 11
            adoaccrpt412.Fields(13).Value = Val(Format(adoaccrpt412.Fields(13).Value, FAmount)) - Val(Format(adoaccrpt412.Fields(intCounter).Value, FAmount))
         Next intCounter
         adoaccrpt412.Fields(13).Value = Format(Val(Format(adoaccrpt412.Fields(13).Value, FAmount)) + Val(Format(adoaccrpt412.Fields(12).Value, FAmount)), FAmount)
   End Select
   adoacc090.Close
End Sub

'*************************************************
'  畫線
'
'*************************************************
'Modify by amy 2015/03/13 +strAccNo 會計科目欄位
Private Sub PaintLine(strSign As String, Optional strAccNo As String)
Dim intCounter As Integer

   For intCounter = 3 To 14
      If intCounter = 14 Then
        adoaccrpt412.Fields(intCounter).Value = strAccNo & LeftB(strSign, 2) 'TB欄位 varchar(8)
      Else
        adoaccrpt412.Fields(intCounter).Value = strSign
      End If
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
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = Mid(DFormat, 1, 6)
   Text6.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
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

'Add by Amy 2015/03/13 產生Excel
Private Sub ExcelSave()
    Dim xlsAnnuity As New Excel.Application
    Dim wksAnnuity As New Worksheet
    Dim strFileName As String, strTemp As String
    Dim ii As Integer, intField As Integer, intCounter As Integer, intTitleRow As Integer
    Dim strStartRow As String, strEndRow As String '合計起/迄始位置
    Dim strTotal(2) As String, strDSum(1) As String '加總列號(0:其他 1:智權部及總所/管理部 2:全所)/營業收入/支出加總列號
    Dim strVSum(1) '各部門加總列號(0:其他/1:全所)
    Dim strTotPos(1 To 2) As String 'Added by Lydia 2016/01/30 全所損益欄位
    'Add by Amy 2016/07/27
    Dim strOSum(1) As String '營業外收入/支出加總列號
    Dim bol105YA As Boolean '是否為105年後資料
    
    If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then bol105YA = True
   
    ReDim strFieldN(11)
    ReDim intWidth(11)
    'modify by sonia 2016/1/28 105年起法務及投法合併,費用傳票輸在L部門,但此表部門名稱改法務部,放在原投法的位置
    'strFieldN = Array("會計科目", "專利", "商標", "法務", "CFP", "CFT", "FCP", "FCT", "投法", "智權部", "總所/管理", "全所")
    'intWidth = Array(13, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10)
    If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
      strFieldN = Array("會計科目", "專利", "商標", "法務", "CFP", "CFT", "FCP", "FCT", "法務部", "智權部", "總所/管理", "全所")
      intWidth = Array(13, 10, 10, 0, 10, 10, 10, 10, 10, 10, 10, 10)
    Else
      strFieldN = Array("會計科目", "專利", "商標", "法務", "CFP", "CFT", "FCP", "FCT", "投法", "智權部", "總所/管理", "全所")
      intWidth = Array(13, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10)
    End If
   'end 2016/1/28 end
                                
On Error GoTo ErrHnd
    
    intField = 65:  intCounter = 1
    strFileName = Val(Replace(MaskEdBox1.Text, "/", "")) & "-" & Val(Replace(MaskEdBox2.Text, "/", "")) & "部門綜合損益表-子科目" & ServerDate & MsgText(43)
    If Dir(strExcelPath & strFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
             MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileName
    End If
    
    xlsAnnuity.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
    xlsAnnuity.Workbooks.add
    Set wksAnnuity = xlsAnnuity.Worksheets(1)
    
    With wksAnnuity
        '***表頭設定***
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField) & intCounter).Value = "部門綜合損益表 (子科目)"
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField) & intCounter).HorizontalAlignment = xlCenter
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField) & intCounter).VerticalAlignment = xlCenter
        intCounter = intCounter + 1
        
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField - 1) & intCounter).Value = "公司別："
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField) & intCounter).Value = Trim(Text6) & IIf(Text7 = "", "台一　專利商標/智權", Text7)
        intCounter = intCounter + 1
        
        .Range(Chr(intField) & intCounter).Value = "列印日期：" & CFDate(ACDate(ServerDate))
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField - 1) & intCounter).Value = "年　月："
        .Range(Chr(Fix(UBound(strFieldN) / 2) + intField) & intCounter).Value = MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
        intCounter = intCounter + 1
        
        .Range(Chr(intField) & intCounter).Value = "列印人員：" & strUserName
        intCounter = intCounter + 1
        
        For ii = 0 To UBound(strFieldN)
            .Columns(Chr(intField + ii) & ":" & Chr(intField + ii)).ColumnWidth = intWidth(ii)
            .Range(Chr(intField + ii) & intCounter).Value = strFieldN(ii)
            .Range(Chr(intField + ii) & intCounter).HorizontalAlignment = xlCenter
        Next ii
        'Add by Amy 2016/07/27 +框線
        Call SetExcelLine(0, wksAnnuity, Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter)
       
        intTitleRow = intCounter: intCounter = intCounter + 1: strStartRow = intCounter
        '列印資料
        Do While adoaccrpt412.EOF = False
            If "" & adoaccrpt412.Fields(2) <> "" Then
                .Range(Chr(intField) & intCounter).Value = adoaccrpt412.Fields(2) '會計科目欄位
            End If
            
            If "" & adoaccrpt412.Fields("r41215") = "ZZZZZZZZ" Then
                    For ii = 1 To UBound(strFieldN)
                        Select Case ii
                            'modify by sonia 2016/2/16 +法務部
                            Case GetValue("智權部"), GetValue("總所/管理"), GetValue("法務部")
                                'Modify by Amy 2016/07/25 改為營業外收入-營業外支出
                                'strTemp = "," & Chr(intField + ii) & strTotal(1)
                                strTemp = Chr(intField + ii) & strOSum(0) & "-" & Chr(intField + ii) & strOSum(1)
                            Case GetValue("全所")
                                'Modify by Amy 2016/07/27 排除不計入全所的列(分攤合計列)
                                'strTemp = Replace(strTotal(2), ",", "," & Chr(intField + ii))
                                strTemp = Replace(Replace(strTotal(0), strTotal(2), ""), ",", "," & Chr(intField + ii))
                                'Added by Lydia 2016/01/30
                                strTotPos(1) = Chr(intField + ii): strTotPos(2) = intCounter
                            Case Else
                                strTemp = Replace(strTotal(0), ",", "," & Chr(intField + ii))
                        End Select
                        'Modify by Amy 2016/07/27改為營業外收入-營業外支出
                        If GetValue("智權部") <> ii And GetValue("總所/管理") <> ii And GetValue("法務部") <> ii Then
                           strTemp = "sum(" & Right(Replace(strTemp, Chr(intField + ii) & "-", "-" & Chr(intField + ii)), Len(strTemp) - 1) & ")"
                        End If
                        .Range(Chr(intField + ii) & intCounter).Formula = "=" & strTemp
                        'end 2016/07/27
                    Next ii
                   'Add by Amy 2016/07/27 +框線
                   Call SetExcelLine(1, wksAnnuity, Chr(intField) & intCounter - 1 & ":" & Chr(UBound(strFieldN) + intField) & intCounter)
            ElseIf InStr("" & adoaccrpt412.Fields("r41215"), "S") > 0 Then
                   'Add by Amy 2016/07/27 +框線
                   If "" & adoaccrpt412.Fields("r41215") = "DS" Or "" & adoaccrpt412.Fields("r41215") = "VS" Then
                        Call SetExcelLine(0, wksAnnuity, Chr(intField) & intCounter - 1 & ":" & Chr(UBound(strFieldN) + intField) & intCounter - 1)
                   Else
                        If strStartRow = intTitleRow + 1 Then
                            Call SetExcelLine(2, wksAnnuity, Chr(intField) & strStartRow & ":" & Chr(UBound(strFieldN) + intField) & strEndRow)
                        Else
                            Call SetExcelLine(2, wksAnnuity, Chr(intField) & strStartRow - 1 & ":" & Chr(UBound(strFieldN) + intField) & strEndRow)
                        End If
                   End If
                   
                    '*** 合計
                    For ii = 1 To UBound(strFieldN)
                        .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "#,##0.00 ;[紅色]-#,##0.00"
                        '部門損益
                        If "" & adoaccrpt412.Fields("r41215") = "DS" Then
                            .Range(Chr(intField + ii) & intCounter).Formula = "=" & Chr(intField + ii) & strDSum(0) & "-" & Chr(intField + ii) & strDSum(1)
                        '各部門營業損益
                        ElseIf "" & adoaccrpt412.Fields("r41215") = "VS" Then
                            'modify by sonia 2016/7/13 法務部也是0
                            'If GetValue("智權部") = ii Or GetValue("總所/管理") = ii Then
                            If GetValue("智權部") = ii Or GetValue("總所/管理") = ii Or GetValue("法務部") = ii Then
                                .Range(Chr(intField + ii) & intCounter).Value = 0
                            ElseIf GetValue("全所") = ii Then
                                .Range(Chr(intField + ii) & intCounter).Formula = "=" & Chr(intField + ii) & strVSum(1)
                            Else
                                strTemp = Replace(strVSum(0), ",", "," & Chr(intField + ii))
                                .Range(Chr(intField + ii) & intCounter).Formula = "=sum(" & Replace(strTemp, Chr(intField + ii) & "-", "-" & Chr(intField + ii)) & ")"
                            End If
                        Else
                            .Range(Chr(intField + ii) & intCounter).Formula = "=sum(" & Chr(intField + ii) & strStartRow & ":" & Chr(intField + ii) & strEndRow & ")"
                        End If
                    Next ii
                    'Add by Amy 2016/07/27 +框線
                    Call SetExcelLine(0, wksAnnuity, Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter)

                    'strDSum():部門損益 / strVSum():各部門營業損益 / strTotal():全所損益 計算欄位
                    strTemp = Left("" & adoaccrpt412.Fields("r41215"), 1)
                    Select Case strTemp
                        Case "4" '營業收入
                            strDSum(0) = intCounter
                        Case "6" '營業支出
                             strDSum(1) = intCounter
                        Case "7" '營業外收支
                            'Add by Amy 2016/07/27 收入支出拆開
'                            strTotal(1) = intCounter
'                            strTotal(2) = strTotal(2) & "," & intCounter
                            If Left(adoaccrpt412.Fields("r41215"), 2) = "71" Then
                               strOSum(0) = intCounter
                               strTotal(0) = strTotal(0) & "," & intCounter
                               strTotal(1) = strTotal(1) & "," & intCounter
                            Else
                               strOSum(1) = intCounter
                               strTotal(0) = strTotal(0) & ",-" & intCounter
                               strTotal(1) = strTotal(1) & ",-" & intCounter
                            End If
                        Case "9" '分攤費用
                            strVSum(0) = strVSum(0) & ",-" & intCounter
                            strTotal(0) = strTotal(0) & ",-" & intCounter
                            strTotal(2) = ",-" & intCounter 'Add by Amy 2016/07/27
                        Case "D" '部門損益
                            strVSum(0) = strVSum(0) & "," & intCounter
                            strVSum(1) = intCounter
                            strTotal(0) = strTotal(0) & "," & intCounter
                         'Mark by Amy 2016/07/27 營業外收支拆開
'                        Case "V" '各部門營業損益
'                            strTotal(2) = strTotal(2) & "," & intCounter
                        Case Else
                    End Select
            'Add by Amy 2016/07/25 +判斷為分攤科目改為公式顯示-婧瑄
            ElseIf "" & adoaccrpt412.Fields("r41215") = "9997" Or "" & adoaccrpt412.Fields("r41215") = "9998" Or "" & adoaccrpt412.Fields("r41215") = "9999" Then
                For ii = UBound(strFieldN) To 1 Step -1
                    .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "#,##0.00 ;[紅色]-#,##0.00"
                    If GetValue("全所") = ii Then
                        Select Case "" & adoaccrpt412.Fields("r41215")
                            Case "9997"
                                strTemp = "=" & Chr(intField + GetValue("法務部")) & strDSum(1)
                            Case "9998"
                                strTemp = "=" & Chr(intField + GetValue("總所/管理")) & strDSum(1)
                            Case Else
                                strTemp = "=" & Chr(intField + GetValue("智權部")) & strDSum(1)
                        End Select
                    ElseIf GetValue("專利") = ii Then
                        '與其他欄位一樣用算的再加總全所會造成小數位與XX部門費用合計不符
                        strTemp = "=Round(" & Chr(intField + GetValue("全所")) & intCounter & "-Sum(" & Chr(intField + GetValue("商標")) & intCounter & ":" & Chr(intField + GetValue("總所/管理")) & intCounter & "),2)"
                    ElseIf "" & adoaccrpt412.Fields("r41215") = "9997" Then
                        strTemp = "=Round(" & Chr(intField + GetValue("法務部")) & strDSum(1) & "*(" & Chr(intField + ii) & strDSum(0) & "/" & Chr(intField + GetValue("全所")) & strDSum(0) & "),2)"
                    ElseIf "" & adoaccrpt412.Fields("r41215") = "9998" Then
                        strTemp = "=Round(" & Chr(intField + GetValue("總所/管理")) & strDSum(1) & "*(" & Chr(intField + ii) & strDSum(0) & "/" & Chr(intField + GetValue("全所")) & strDSum(0) & "),2)"
                    Else
                        If GetValue("FCP") = ii Or GetValue("FCT") = ii Or GetValue("法務部") = ii Or GetValue("智權部") = ii Or GetValue("總所/管理") = ii Then
                            strTemp = "0"
                        Else
                            strTemp = "=Round(" & Chr(intField + GetValue("智權部")) & strDSum(1) & "*(" & Chr(intField + ii) & strDSum(0) & "/(" & Chr(intField + GetValue("全所")) & strDSum(0) & "-" & _
                                                Chr(intField + GetValue("FCP")) & strDSum(0) & "-" & Chr(intField + GetValue("FCT")) & strDSum(0) & IIf(bol105YA = True, "", "-" & Chr(intField + GetValue("投法")) & strDSum(0)) & ")),2)"
                        End If
                    End If
                    If InStr(strTemp, "=") > 0 Then
                        .Range(Chr(intField + ii) & intCounter).Formula = strTemp
                    Else
                        .Range(Chr(intField + ii) & intCounter).Value = Val(strTemp)
                    End If
                Next ii
            Else
                   'Add by Amy 2016/07/25
                    If "" & adoaccrpt412.Fields("r41215") = "" Or InStr("" & adoaccrpt412.Fields("r41215"), "－") > 0 Then
                         intCounter = intCounter - 1
                    End If
               
                    '*** 資料
                    For ii = 1 To UBound(strFieldN)
                        If InStr("" & adoaccrpt412.Fields("r41215"), "－") > 0 Then strEndRow = intCounter 'Modify by Amy 2016/07/25 原:- 1- 1 '更新合計結束位置
                        If InStr("" & adoaccrpt412.Fields("r41215"), "E") > 0 Then strStartRow = intCounter + 1 '更新合計起始位置
                        If InStr("" & adoaccrpt412.Fields("r41215"), "－") > 0 Or InStr("" & adoaccrpt412.Fields("r41215"), "E") > 0 Or _
                            "" & adoaccrpt412.Fields("r41215") = "" Or "" & adoaccrpt412.Fields("r41215") = "＝" Then
                             'Mark by Amy 2016/07/25 －/＝不印
                             '.Range(Chr(intField + ii) & intCounter).Value = adoaccrpt412.Fields(ii + 2)
                        Else
                            .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "#,##0.00 ;[紅色]-#,##0.00" '數字資料設小數2位
                            '資料
                            Select Case ii
                                Case GetValue("總所/管理")
                                    .Range(Chr(intField + ii) & intCounter).Value = Val(adoaccrpt412.Fields(ii + 3))
                                Case GetValue("全所")
                                    .Range(Chr(intField + ii) & intCounter).Value = Val(adoaccrpt412.Fields(ii + 1))
                                Case Else
                                    .Range(Chr(intField + ii) & intCounter).Value = Val(adoaccrpt412.Fields(ii + 2))
                            End Select
                        End If
                    Next ii
            End If
            adoaccrpt412.MoveNext
            intCounter = intCounter + 1
        Loop
        
       'Added by Lydia 2016/01/30 +利潤率(部門損益/全所損益)
       .Range(Chr(intField) & intCounter).Value = "利潤率"
        For ii = 1 To UBound(strFieldN)
            If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
               strExc(0) = "法務部,智權部,總所/管理,全所"
            Else
               strExc(0) = "智權部,總所/管理,全所"
            End If
            If InStr(strExc(0), strFieldN(ii)) = 0 Then
               .Range(Chr(intField + ii) & intCounter).Formula = "=" & Chr(intField + ii) & strTotPos(2) & "/$" & strTotPos(1) & "$" & strTotPos(2)
               .Range(Chr(intField + ii) & intCounter).NumberFormatLocal = "#,##0.00%"
            End If
        Next ii
        intCounter = intCounter + 2
        'end 2016/01/30
        'Add by Amy 2016/07/27 +備註
        .Range(Chr(intField) & intCounter).Value = "備註："
        intCounter = intCounter + 1
        .Range(Chr(intField) & intCounter).Value = "1.分攤法務部門費用: 法務部費用總合＊各該部門當月實際收入／全所實際收入"
        intCounter = intCounter + 1
        .Range(Chr(intField) & intCounter).Value = "2.分攤管理部門費用: 管理部費用總合＊該部門當月實際收入／全所實際收入"
        intCounter = intCounter + 1
        .Range(Chr(intField) & intCounter).Value = "3.分攤智權部門費用: 智權部費用總合＊該部門當月實際收入／（全所實際收入－ＦＣＰ收入－ＦＣＴ收入）"
    End With
    'Add by Amy 2016/07/27 Excel字型大小設定
    With wksAnnuity.Range(Chr(intField) & "1:" & Chr(UBound(strFieldN) + intField) & intCounter)
        .Font.Name = "新細明體"
        .Font.Size = 10
    End With
    With wksAnnuity
        .PageSetup.PaperSize = 9 'Add by Amy 2016/07/27 設A4
        .PageSetup.PrintTitleRows = "$1:$" & intTitleRow
        .PageSetup.Orientation = xlLandscape '橫印
        .PageSetup.TopMargin = xlsAnnuity.InchesToPoints(0.78) '上
        .PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.78) '下
        .PageSetup.LeftMargin = xlsAnnuity.InchesToPoints(0.78) '左邊界
        .PageSetup.RightMargin = xlsAnnuity.InchesToPoints(0.5) '右邊界
    End With
    'Modify by Amy2016/05/06 判斷若版本2007以上改變存格式
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    Set wksAnnuity = Nothing
    Set xlsAnnuity = Nothing
    MsgBox "檔案已產生~"
    Exit Sub
   
ErrHnd:
    If adoaccrpt412.State = adStateOpen Then adoaccrpt412.Close
    'Modify by Amy2016/05/06 判斷若版本2007以上改變存格式
    If Val(xlsAnnuity.Version) < 12 Then
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        xlsAnnuity.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    'end 2016/05/06
    xlsAnnuity.Workbooks.Close
    xlsAnnuity.Quit
    Set wksAnnuity = Nothing
    Set xlsAnnuity = Nothing
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strFieldN)
       If UCase(strFieldN(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function
'Added by Lydia 2016/01/30 從AccReport改成Printer
Private Sub PrintData()
Dim strRBase As String '全所損益
Dim strRD(0 To 10) As String
Dim ii As Integer
    Printer.EndDoc
    Printer.Orientation = 1 '1.直印 2.橫印
    Printer.PaperSize = PUB_GetPaperSize(15) '美國標準
       
    lngPageHeight = Printer.ScaleHeight
    lngPageWidth = Printer.ScaleWidth
    lngLineHeight = 300
           
    iPage = 0
    GetPleft
    Erase strRD
    
    PrintHeader '列印表頭
    With adoaccrpt412
        Do While Not .EOF
        '列印明細
           iPage = iPage + 1
           strTemp(0) = "" & .Fields("R41203") '會計科目
           strTemp(1) = "" & .Fields("R41204") '專利
           strTemp(2) = "" & .Fields("R41205") '商標
           strTemp(3) = "" & .Fields("R41206") '法務->105年以後併入"法務部"
           strTemp(4) = "" & .Fields("R41207") 'CFP
           strTemp(5) = "" & .Fields("R41208") 'CFT
           strTemp(6) = "" & .Fields("R41209") 'FCP
           strTemp(7) = "" & .Fields("R41210") 'FCT
           strTemp(8) = "" & .Fields("R41211") '投法->105年以後"法務部"
           strTemp(9) = "" & .Fields("R41212") '智權部
           strTemp(10) = "" & .Fields("R41214") '總所/管理
           strTemp(11) = "" & .Fields("R41213") '全所
           If .Fields("R41215") = "ZZZZZZZZ" Then
              strRBase = strTemp(11)
              strRD(0) = "利潤率"
              If Val(strRBase) <> 0 Then
                 For ii = 1 To 10
                    strRD(ii) = Format(Val(strTemp(ii)) / Val(strRBase), "##0.00%")
                 Next
              End If
           End If
           For ii = 0 To UBound(strFieldN)
              If intWidth(ii) > 0 Then
                 If strTemp(0) = "" And InStr(strTemp(1), "－") = 0 And InStr(strTemp(1), "＝") = 0 And Val(strTemp(11)) = 0 Then
                    '空一行
                    Exit For
                 Else
                    '靠左
                    If ii = 0 Or InStr(strTemp(ii), "－") > 0 Or InStr(strTemp(ii), "＝") > 0 Then
                        Printer.CurrentX = PLeft(ii) + 50
                        Printer.CurrentY = iPrint
                        If ii < UBound(strFieldN) Then
                           Printer.Print strTemp(ii)
                        Else
                            If InStr(strTemp(ii), "－") > 0 Then
                               Printer.Print String(7, "－")
                            ElseIf InStr(strTemp(ii), "＝") > 0 Then
                               Printer.Print String(7, "＝")
                            End If
                        End If
                    '靠右
                    Else
                        Printer.CurrentX = PLeft(ii + 1) - Printer.TextWidth(Format(Val(strTemp(ii)), "###,##0.00")) - ciColGap
                        Printer.CurrentY = iPrint
                        Printer.Print Format(Val(strTemp(ii)), "###,##0.00")
                    End If
                 End If
              End If
           Next
           PrintNewLine
           
           .MoveNext
        Loop
    End With
    
    '利潤率
    For ii = 0 To 10
       If intWidth(ii) > 0 Then
          '靠左
          If ii = 0 Then
              Printer.CurrentX = PLeft(ii)
              Printer.CurrentY = iPrint
              Printer.Print strRD(ii)
          '靠右
          Else
              If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
                 strExc(0) = "法務部,智權部,總所/管理,全所"
              Else
                 strExc(0) = "智權部,總所/管理,全所"
              End If
              If InStr(strExc(0), strFieldN(ii)) = 0 Then
                 Printer.CurrentX = PLeft(ii + 1) - Printer.TextWidth(strRD(ii)) - ciColGap
                 Printer.CurrentY = iPrint
                 Printer.Print strRD(ii)
              End If
          End If
       End If
    Next
'Add by Amy 2016/07/27最後備註
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "備註："
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "1.分攤法務部門費用: 法務部費用總合＊各該部門當月實際收入／全所實際收入"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "2.分攤管理部門費用: 管理部費用總合＊該部門當月實際收入／全所實際收入"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "3.分攤智權部門費用: 智權部費用總合＊該部門當月實際收入／（全所實際收入－ＦＣＰ收入－ＦＣＴ收入）"
'end 2016/07/27
Printer.EndDoc
ShowPrintOk

End Sub


Private Sub GetPleft() '明細表邊界
Dim inX As Integer

'105年起法務及投法合併,費用傳票輸在L部門,但此表部門名稱改法務部,放在原投法的位置
If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
   strFieldN = Array("會計科目", "專利", "商標", "法務", "CFP", "CFT", "FCP", "FCT", "法務部", "智權部", "總所/管理", "全所")
   intWidth = Array(16, 9, 9, 0, 9, 9, 9, 9, 9, 9, 9, 10)
Else
   strFieldN = Array("會計科目", "專利", "商標", "法務", "CFP", "CFT", "FCP", "FCT", "投法", "智權部", "總所/管理", "全所")
   intWidth = Array(16, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 10)
End If
   
Printer.Font.Name = "新細明體"
Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False

Erase PLeft
Erase PTitle
  
   PLeft(0) = ciStartX
   For inX = 1 To UBound(strFieldN)
       If intWidth(inX - 1) = 0 Then
           PLeft(inX) = PLeft(inX - 1)
       Else
           PLeft(inX) = PLeft(inX - 1) + Printer.TextWidth(String(intWidth(inX - 1), "A")) + ciColGap
       End If
   Next
   PLeft(12) = PLeft(11) + Printer.TextWidth(String(10, "A")) + ciColGap
End Sub
Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader

   End If
End Sub

Private Sub PrintHeader()
Dim strPTmp As String
Dim pa1 As Integer
Dim ii As Integer
iPrint = ciStartY
Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.Font.Underline = False

'報表抬頭
strPTmp = ReportTitle(412)
Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
Printer.CurrentY = iPrint
Printer.Print strPTmp

PrintNewLine
PrintNewLine

Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False

strPTmp = "公司別：" & IIf(Text6 = "2", "J", Text6) & " " & IIf(Text7 = "", "台一　專利商標/智權", Text7)
pa1 = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
Printer.CurrentX = pa1
Printer.CurrentY = iPrint
Printer.Print strPTmp

Printer.CurrentX = 15500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & CFDate(strSrvDate(2))

PrintNewLine

Printer.CurrentX = pa1
Printer.CurrentY = iPrint
Printer.Print "年　月：" & MaskEdBox1.Text & " ～ " & MaskEdBox2.Text

Printer.CurrentX = 15500
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page

PrintNewLine

Printer.CurrentX = ciStartX
Printer.CurrentY = iPrint
Printer.Print "列印人員：" & strUserName

PrintNewLine

For ii = 0 To UBound(strFieldN)
   '顯示／欄位
   If intWidth(ii) > 0 And strFieldN(ii) <> "" Then
       strPTmp = strFieldN(ii)
       Printer.CurrentX = PLeft(ii) + ((PLeft(ii + 1) - PLeft(ii) - Printer.TextWidth(strPTmp)) / 2) - ciColGap
       Printer.CurrentY = iPrint
       Printer.Print strPTmp
   End If
Next

PrintNewLine

PrintLine

End Sub

Private Sub PrintLine()
   Printer.Line (PLeft(0) - 50, iPrint)-(PLeft(12) + 50, iPrint)
   iPrint = iPrint + 150
End Sub
'end 2016/01/30

'Added by Lydia 2016/02/17
Private Sub MaskEdBox1_LostFocus()
   If MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
      If InStr(MaskEdBox1.Text, "_") > 0 Then
         MsgBox "請輸入完整年月!", , MsgText(5)
         MaskEdBox1.SetFocus
      Else
         strExc(1) = Replace(MaskEdBox1.Text, "/", "") & "01"
         If CheckIsTaiwanDate(strExc(1)) = False Then
            MaskEdBox1.SetFocus
         End If
      End If
   End If
End Sub
'Added by Lydia 2016/02/17
Private Sub MaskEdBox2_LostFocus()
   If MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
      If InStr(MaskEdBox2.Text, "_") > 0 Then
         MsgBox "請輸入完整年月!", , MsgText(5)
         MaskEdBox2.SetFocus
      Else
         strExc(1) = Replace(MaskEdBox2.Text, "/", "") & "01"
         If CheckIsTaiwanDate(strExc(1)) = False Then
            MaskEdBox2.SetFocus
         End If
      End If
   End If
End Sub

'Add by Amy 2016/07/25 +更新分攤費用值(改成公式計算)-婧瑄
Private Sub UpdAccrpt412()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strUpd As String
    Dim i As Integer, intQ As Integer
    Dim strVal(10) As String, strE As String '更新值(10:合計)/相對費用值
    Dim bol105YA As Boolean '是否為105年後資料
    
    If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then bol105YA = True
    
    strQ = "Select * From accrpt412 Where r41201='" & strUserNum & "' And R41215 in ('9997','9998','9999') Order by R41215"
    If adoaccrpt412.State <> adStateClosed Then adoaccrpt412.Close
    adoaccrpt412.CursorLocation = adUseClient
    adoaccrpt412.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    '抓取相關資料
    strQ = "Select * From " & _
            "(Select R41204 as P,R41205 as T,R41206 as L,R41207 as CFP,R41208 as CFT,R41209 as FCP,R41210 as FCT,R41211 as Law,R41212 as S,R41214 as M,R41213 as Total " & _
             "From accrpt412 Where r41201='" & strUserNum & "' And R41215='4S')," & _
            "(Select R41211 as LE,R41214 as ME,R41212 as SE From accrpt412 Where r41201='" & strUserNum & "' And R41215='6S' )"
    If RsQ.State <> adStateClosed Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If adoaccrpt412.RecordCount > 0 And RsQ.RecordCount > 0 Then
      With adoaccrpt412
        .MoveFirst
        Do While Not .EOF
            Select Case "" & .Fields("R41215")
                Case "9997"
                    strE = Val("" & RsQ.Fields("LE"))
                Case "9998"
                    strE = Val("" & RsQ.Fields("ME"))
                Case "9999"
                    strE = Val("" & RsQ.Fields("SE"))
            End Select
            For i = 1 To 9
               If "" & .Fields("R41215") = "9999" Then
                    If bol105YA = False Then
                        If i + 4 >= 9 And i + 4 <= 14 And i + 4 <> 11 Then
                            strVal(i) = "0"
                        Else
                            '105年以前需剔除「投法」
                            strVal(i) = Format(Round(strE * (Val("" & RsQ.Fields(i)) / (Val(RsQ.Fields("Total")) - Val(RsQ.Fields("FCP")) - Val(RsQ.Fields("FCT")) - Val(RsQ.Fields("Law")))), 2), FAmount)
                        End If
                    Else
                        If i + 4 >= 9 And i + 4 <= 14 Then
                            strVal(i) = "0"
                        Else
                            strVal(i) = Format(Round(strE * (Val("" & RsQ.Fields(i)) / (Val(RsQ.Fields("Total")) - Val(RsQ.Fields("FCP")) - Val(RsQ.Fields("FCT")))), 2), FAmount)
                        End If
                    End If
               Else
                    strVal(i) = Format(Round(strE * (Val("" & RsQ.Fields(i)) / Val(RsQ.Fields("Total"))), 2), FAmount)
               End If
                strUpd = strUpd & ",R412" & IIf(i + 4 > 9, "", "0") & IIf(i = 9, i + 5, i + 4) & "='" & strVal(i) & "'"
                strVal(0) = Val(strVal(0)) + Val(strVal(i))
            Next i
            strVal(0) = Val(strE) - Val(strVal(0))
            '更新
            If strUpd <> MsgText(601) Then
                strUpd = "Update Accrpt412 Set R41213='" & strE & "',R41204='" & strVal(0) & "'" & strUpd & " Where R41201='" & strUserNum & "' And R41215='" & .Fields("R41215") & "'"
                cnnConnection.Execute strUpd
                strUpd = ""
            End If
            strVal(0) = ""
             .MoveNext
        Loop
       End With
    End If
    adoaccrpt412.Close
    RsQ.Close
    
     '更新 分攤費用(9S)
    For i = 0 To 10
        strVal(i) = ""
    Next i
    '抓取相關資料
    strQ = "Select R41204 as P,R41205 as T,R41206 as L,R41207 as CFP,R41208 as CFT,R41209 as FCP," & _
               "R41210 as FCT,R41211 as Law,R41212 as S,R41214 as M,R41213 as Total " & _
             "From accrpt412 Where r41201='" & strUserNum & "' And R41215 in ('9997','9998','9999')"
    If RsQ.State <> adStateClosed Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
      strUpd = ""
      With RsQ
        .MoveFirst
        Do While Not .EOF
            For i = 0 To 9
                strVal(i) = Val(strVal(i)) + Val("" & .Fields(i))
            Next i
            strVal(10) = Val(strVal(10)) + Val("" & .Fields("Total"))
            .MoveNext
        Loop
        If strVal(10) <> MsgText(601) Then
            For i = 0 To 9
                strUpd = strUpd & ",R412" & IIf(i + 4 > 9, "", "0") & IIf(i = 9, i + 5, i + 4) & "='" & strVal(i) & "'"
            Next i
            If strUpd <> MsgText(601) Then
                strUpd = "Update Accrpt412 Set R41213='" & strVal(10) & "' " & strUpd & " Where r41201='" & strUserNum & "' And R41215='9S' "
                cnnConnection.Execute strUpd
            End If
        End If
      End With
    End If
    RsQ.Close
    '更新 各部門營業損益(VS)
     strQ = "Select R41204 as P,R41205 as T,R41206 as L,R41207 as CFP,R41208 as CFT,R41209 as FCP," & _
               "R41210 as FCT,R41211 as Law,R41212 as S,R41214 as M,R41213 as Total From accrpt412 Where r41201='" & strUserNum & "' And R41215 ='DS' "
    If RsQ.State <> adStateClosed Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
      strUpd = ""
      With RsQ
            .MoveFirst
            Do While Not .EOF
                For i = 0 To 9
                    If bol105YA = False Then
                        If i >= 8 And i <= 9 Then
                            strVal(i) = "0"
                        Else
                            strVal(i) = Val("" & .Fields(i)) - Val(strVal(i))
                        End If
                    Else
                        If i >= 7 And i <= 9 Then
                            strVal(i) = "0"
                        Else
                            strVal(i) = Val("" & .Fields(i)) - Val(strVal(i))
                        End If
                    End If
                Next i
                strVal(10) = Val("" & .Fields("Total"))
                .MoveNext
            Loop
            For i = 0 To 9
                strUpd = strUpd & ",R412" & IIf(i + 4 > 9, "", "0") & IIf(i = 9, i + 5, i + 4) & "='" & strVal(i) & "'"
            Next i
            strUpd = "Update Accrpt412 Set R41213='" & strVal(10) & "' " & strUpd & " Where r41201='" & strUserNum & "' And R41215='VS' "
            cnnConnection.Execute strUpd
      End With
    End If
    '更新 全所損益(ZZZZZZZZ)
     strQ = "Select R41204 as P,R41205 as T,R41206 as L,R41207 as CFP,R41208 as CFT,R41209 as FCP," & _
               "R41210 as FCT,R41211 as Law,R41212 as S,R41214 as M,R41213 as Total,R41215 From accrpt412 " & _
               "Where r41201='" & strUserNum & "' And R41215 in ('71S','72S') Order by R41215"
    If RsQ.State <> adStateClosed Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
      strUpd = ""
      With RsQ
            .MoveFirst
            Do While Not .EOF
                For i = 0 To 10
                    If .Fields("R41215") = "71S" Then
                        '+ 營業外收入
                        strVal(i) = Val(strVal(i)) + Val("" & .Fields(i))
                    Else
                        '- 營業外支出
                        strVal(i) = Val(strVal(i)) - Val("" & .Fields(i))
                    End If
                Next i
                .MoveNext
            Loop
            For i = 0 To 9
                strUpd = strUpd & ",R412" & IIf(i + 4 > 9, "", "0") & IIf(i = 9, i + 5, i + 4) & "='" & strVal(i) & "'"
            Next i
            strUpd = "Update Accrpt412 Set R41213='" & strVal(10) & "' " & strUpd & " Where r41201='" & strUserNum & "' And R41215='ZZZZZZZZ' "
            cnnConnection.Execute strUpd
      End With
    End If
End Sub

'增加框線設定-婉莘
Private Sub SetExcelLine(intChoose As Integer, ByRef m_Xls As Worksheet, strField As String)

    With m_Xls.Range(strField)
        Select Case intChoose
            Case 0 '抬頭/合計
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlHairline
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
            Case 1 '最後合計
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThick
                .Borders(xlEdgeBottom).LineStyle = xlDouble
                .Borders(xlEdgeBottom).Weight = xlThick
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlHairline
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).Weight = xlThick
            Case 2 '資料內容
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlHairline
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).Weight = xlHairline
        End Select
    End With
End Sub
'end 2016/07/27
