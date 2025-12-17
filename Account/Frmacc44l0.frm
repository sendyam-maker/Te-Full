VERSION 5.00
Begin VB.Form Frmacc44l0 
   AutoRedraw      =   -1  'True
   Caption         =   "預算資料表"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2760
   ScaleWidth      =   5325
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
      Top             =   120
      Width           =   3500
   End
   Begin VB.TextBox Text7 
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
      Left            =   2430
      MaxLength       =   1
      TabIndex        =   18
      Top             =   3270
      Width           =   435
   End
   Begin VB.TextBox Text9 
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
      Top             =   1260
      Width           =   612
   End
   Begin VB.TextBox Text8 
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
      TabIndex        =   3
      Top             =   900
      Width           =   852
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
      Height          =   300
      Left            =   2670
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3660
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel"
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
      Left            =   165
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   1770
      Width           =   4692
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
      Top             =   540
      Width           =   1572
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
      Top             =   540
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
      Height          =   300
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3660
      Width           =   852
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   540
      TabIndex        =   19
      Top             =   180
      Width           =   675
   End
   Begin VB.Label Label11 
      Caption         =   "PS：1. 若資料來源選實績而該月尚未結算則會改抓預算資料            2. 產生之Excel檔案子科目數字由人工自行填寫"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   90
      TabIndex        =   17
      Top             =   2205
      Width           =   5190
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1:實績 2:預算)"
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
      Left            =   2070
      TabIndex        =   16
      Top             =   1290
      Width           =   1530
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "資料來源"
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
      Left            =   315
      TabIndex        =   15
      Top             =   1290
      Width           =   915
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "預算年度"
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
      Left            =   315
      TabIndex        =   14
      Top             =   930
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "是否產生Excel檔案"
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
      Left            =   405
      TabIndex        =   13
      Top             =   3300
      Width           =   1905
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "(Y/N)"
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
      Left            =   2970
      TabIndex        =   12
      Top             =   3300
      Width           =   540
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
      Left            =   2430
      TabIndex        =   11
      Top             =   3660
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   135
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
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   540
      Width           =   255
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
      Height          =   255
      Left            =   315
      TabIndex        =   9
      Top             =   570
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   510
      TabIndex        =   8
      Top             =   3660
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc44l0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
'Modify by Morgan 2008/12/1
'畫面部門別條件鎖住TOT且不顯示(報表會用到所以只預設暫不拿掉)
Option Explicit

Public adoacc080 As New ADODB.Recordset
Public adoacc090 As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoaccrpt418 As New ADODB.Recordset
Dim lngCounter As Long
Dim dllaccrpt418 As Object
Dim strSQL1 As String, strSQL2 As String
Dim strSql As String
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/4/27


'Add by Sindy 2020/4/27
Private Sub SetCompN()
    strCmpN = "": strCmp = ""
    If Trim(CboCmp) <> MsgText(601) Then
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, True, True)
End Sub

Private Sub CboCmp_GotFocus()
    TextInverse CboCmp
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
    If InStr(GetBookKeepCmp & 組合作帳公司 & ",", strCmp) = 0 Then
        MsgBox Label1 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/4/27

Private Sub Command1_Click()
   Dim stCon010 As String, stCon040 As String, stAccDate As String
   
   If FormCheck = False Then
      Exit Sub
   End If
   
   Call SetCompN 'Add by Sindy 2020/4/27
   
   Screen.MousePointer = vbHourglass
   'Add by Morgan 2008/11/26
   '存成Excel
'   If Text7 = "Y" Then
      stCon010 = ""
      stCon040 = ""
      '公司別
      'Modify By Sindy 2020/4/27
'      If Text2 <> "" Then
'         '2014/3/3 modify by sonia
'         'stCon040 = stCon040 & " and a0403(+) = '" & Text2 & "'"
'         stCon040 = stCon040 & " and a0403(+) = '" & IIf(Text2 = "2", "J", "1") & "'"
'         '2014/3/3 end
      If strCmp <> "" Then
         If InStr(strCmp, "+") > 0 Then
            stCon040 = stCon040 & " and a0403(+) In ('" & Replace(strCmp, "+", "','") & "')"
         Else
            stCon040 = stCon040 & " and a0403(+) = '" & strCmp & "'"
         End If
      End If
      
      '會計科目起
      If Text5 <> "" Then
         stCon010 = stCon010 & " and a0101 >= '" & Text5 & "'"
      End If
      '會計科目迄
      If Text6 <> "" Then
         stCon010 = stCon010 & " and a0101 <= '" & Text6 & "'"
      End If
      '年度
      If Text8 <> "" Then
         stCon040 = stCon040 & " and a0401(+) = '" & Text8 & "'"
      End If

      '實績
      If Text9 = "1" Then
         '若未結算時抓預算
         '2014/3/3 modify by sonia 依公司別抓ACC0B0,合併則抓1公司,另下合併公司時要sum()
         'strExc(0) = "select a0101,a0402,a0102" & _
            ",decode(sign(a0401||a0402-trunc(a0b01/100)),1,a0409,a0408) a0409" & _
            " From acc010, acc040, acc0b0" & _
            " where a0104 in ('3','4') and instr(a0102,'/不用')=0 " & stCon010 & _
            " and a0405(+)=a0101 and a0404(+)='TOT'" & stCon040 & " order by a0101,a0402"
         
         'Modify By Sindy 2020/4/27
'         strExc(0) = "select a0101,a0402,a0102" & _
'            ",sum(decode(sign(a0401||a0402-trunc(a0b01/100)),1,a0409,a0408)) a0409" & _
'            " From acc010, acc040, acc0b0" & _
'            " where a0104 in ('3','4') and instr(a0102,'/不用')=0 " & stCon010 & _
'            " and a0405(+)=a0101 and a0b04='" & IIf(Text2 = "2", "J", "1") & "' and a0404(+)='TOT'" & stCon040 & " group by a0101,a0402,a0102 order by a0101,a0402"
         'Modify By Sindy 2020/10/26 Mark,下列SQL直接下 a0b04(+)=a0403
'         If InStr(strCmp, "+") > 0 Then
'            strExc(10) = " and a0b04 In ('" & Replace(strCmp, "+", "','") & "')"
'         Else
'            ' and a0b04='" & IIf(Text2 = "2", "J", "1") & "'
'            strExc(10) = " and a0b04='" & IIf(strCmp = "", "1", strCmp) & "'"
'         End If
         strExc(0) = "select a0101,a0402,a0102" & _
            ",sum(decode(sign(a0401||a0402-trunc(a0b01/100)),1,a0409,a0408)) a0409" & _
            " From acc010, acc040, acc0b0" & _
            " where a0104 in ('3','4') and instr(a0102,'/不用')=0 " & stCon010 & _
            " and a0405(+)=a0101 and a0b04(+)=a0403 and a0404(+)='TOT'" & stCon040 & " group by a0101,a0402,a0102 order by a0101,a0402"
      '預算
      Else
         '若未結算時抓預算
         '2014/3/3 modify by sonia 下合併公司時要sum()
         'strExc(0) = "select a0101,a0402,a0102,a0409" & _
            " From acc010, acc040" & _
            " where a0104 in ('3','4') and instr(a0102,'/不用')=0 " & stCon010 & _
            " and a0405(+)=a0101 and a0404(+)='TOT'" & stCon040 & " order by a0101,a0402"
         strExc(0) = "select a0101,a0402,a0102,sum(a0409)" & _
            " From acc010, acc040" & _
            " where a0104 in ('3','4') and instr(a0102,'/不用')=0 " & stCon010 & _
            " and a0405(+)=a0101 and a0404(+)='TOT'" & stCon040 & " group by a0101,a0402,a0102 order by a0101,a0402"
      End If
      
      intI = 1
      Set adoaccrpt418 = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Export2Excel adoaccrpt418
         adoaccrpt418.Close
      End If
'   'end 2008/11/26
'   Else
'      strSql = ""
'      If Text3 <> "" Then
'         strSql = strSql & " and r41803 >= '" & Text3 & "'"
'      End If
'      If Text4 <> "" Then
'         strSql = strSql & " and r41803 <= '" & Text4 & "'"
'      End If
'      If strSql <> "" Then
'         strSql = " where " & Mid(strSql, 5, Len(strSql) - 4)
'      End If
'      Accrpt418Delete
'      ProduceData
'      adoacc080.CursorLocation = adUseClient
'      '2014/3/3 modify by sonia
'      'adoacc080.Open "select distinct r41802 from accrpt418 where r41802 = '" & Text2 & "' order by r41802 asc", adoTaie, adOpenStatic, adLockReadOnly
'      If Text2 <> "" Then
'         adoacc080.Open "select distinct r41802 from accrpt418 where r41802 = '" & IIf(Text2 = "2", "J", Text2) & "' order by r41802 asc", adoTaie, adOpenStatic, adLockReadOnly
'      Else
'         adoacc080.Open "select distinct r41802 from accrpt418 order by r41802 asc", adoTaie, adOpenStatic, adLockReadOnly
'      End If
'      '2014/3/3 end
'      Do While adoacc080.EOF = False
'         adoacc090.CursorLocation = adUseClient
'         adoacc090.Open "select distinct r41803 from accrpt418" & strSql & " order by r41803 asc", adoTaie, adOpenStatic, adLockReadOnly
'         Do While adoacc090.EOF = False
'            RunReportDll
'            adoacc090.MoveNext
'         Loop
'         adoacc090.Close
'         adoacc080.MoveNext
'      Loop
'      adoacc080.Close
'   End If
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
   Me.Width = 5445
   Me.Height = 3165 '3630
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   'Add by Sindy 2020/4/27 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/4/27
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt418 = CreateObject("AccReport.ReportSelect")
   FormClear
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt418 = Nothing
   Set Frmacc44l0 = Nothing
End Sub

'Modify by Sindy 2020/4/27 公司別改下拉
'Private Sub Text2_Change()
'   '2014/3/3 modify by sonia
'   'If Text2 = MsgText(601) Then
'   '   Exit Sub
'   'End If
'   'Text1 = A0802Query(Text2)
'   Select Case Text2
'      Case "1"
'         Text1 = A0802Query(Text2)
'      Case "2"
'         Text1 = A0802Query("J")
'      Case ""
'         Text1 = "台一　專利商標/智權"
'   End Select
'   '2014/3/3 end
'End Sub
'
'Private Sub Text2_GotFocus()
'   TextInverse Text2
'End Sub
'
''2014/3/3 add by sonia
'Private Sub Text2_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub
''2014/3/3 end

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

''*************************************************
''  產生報表資料
''
''*************************************************
'Private Sub ProduceData()
'Dim intCounter As Integer
'
'On Error GoTo Checking
'   strSQL1 = ""
'   strSQL2 = ""
'   If Text3 <> MsgText(601) Then
'      strSQL1 = " and a0901 >= '" & Text3 & "'"
'   End If
'   If Text4 <> MsgText(601) Then
'      strSQL1 = strSQL1 & " and a0901 <= '" & Text4 & "'"
'   End If
'   If strSQL1 <> MsgText(601) Then
'      strSQL1 = " where " & Mid(strSQL1, 5, Len(strSQL1) - 4)
'   End If
'   If Text5 <> MsgText(601) Then
'      strSQL2 = " and a0405 >= '" & Text5 & "'"
'   End If
'   If Text6 <> MsgText(601) Then
'      strSQL2 = strSQL2 & " and a0405 <= '" & Text6 & "'"
'   End If
'
'   '2014/3/3 add by sonia 加公司別條件
'   'Modify By Sindy 2020/4/27
''   If Text2 <> "" Then
''      strSQL2 = strSQL2 & " and a0403(+) = '" & IIf(Text2 = "2", "J", "1") & "'"
''   End If
'   If strCmp <> "" Then
'      If InStr(strCmp, "+") > 0 Then
'         strSQL2 = strSQL2 & " and a0403(+) In ('" & Replace(strCmp, "+", "','") & "')"
'      Else
'         strSQL2 = strSQL2 & " and a0403(+) = '" & strCmp & "'"
'      End If
'   End If
'   '2020/4/27 END
'   '2014/3/3 end
'
'   intCounter = 0
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   '2014/3/3 cancel by sonia 因為可合併
'   'adoacc080.CursorLocation = adUseClient
'   'adoacc080.Open "select * from acc080 where a0801 = '" & Text2 & "' order by a0801 asc", adoTaie, adOpenStatic, adLockReadOnly
'   'If adoacc080.RecordCount = 0 Then
'   '   adoacc080.Close
'   '   MsgBox MsgText(28), , MsgText(5)
'   '   Exit Sub
'   'End If
'   'Do While adoacc080.EOF = False
'      adoacc090.CursorLocation = adUseClient
'      adoacc090.Open "select * from acc090" & strSQL1 & " order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
'      Do While adoacc090.EOF = False
'         adoacc040.CursorLocation = adUseClient
'         'Modify by Morgan 2008/11/27 年度改抓畫面欄位
'         'adoacc040.Open "select a0409, a0405, a0402, a0102 from acc040, acc010 where a0405 = a0101 and a0401 = " & Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) & " and a0403 = '" & adoacc080.Fields("a0801").Value & "' and a0404 = '" & adoacc090.Fields("a0901").Value & "' and a0409 <> 0" & strSQL2 & " order by a0402 asc, a0405 asc", adoTaie, adOpenStatic, adLockReadOnly
'         'Modified by Morgan 2013/1/23 改先依科目排序--辜
'         'adoacc040.Open "select a0409, a0405, a0402, a0102 from acc040, acc010 where a0405 = a0101 and a0401 = " & Val(Text8) & " and a0403 = '" & adoacc080.Fields("a0801").Value & "' and a0404 = '" & adoacc090.Fields("a0901").Value & "' and a0409 <> 0" & strSQL2 & " order by a0402 asc, a0405 asc", adoTaie, adOpenStatic, adLockReadOnly
'         '2014/3/3 modify by sonia a0403條件併入strSQL2,下合併公司時要sum()
'         'adoacc040.Open "select a0409, a0405, a0402, a0102 from acc040, acc010 where a0405 = a0101 and a0401 = " & Val(Text8) & " and a0403 = '" & adoacc080.Fields("a0801").Value & "' and a0404 = '" & adoacc090.Fields("a0901").Value & "' and a0409 <> 0" & strSQL2 & " order by a0405 asc,a0402 asc", adoTaie, adOpenStatic, adLockReadOnly
'         adoacc040.Open "select sum(a0409), a0405, a0402, a0102 from acc040, acc010 where a0405 = a0101 and a0401 = " & Val(Text8) & " and a0404 = '" & adoacc090.Fields("a0901").Value & "' and a0409 <> 0" & strSQL2 & " group by a0405, a0402, a0102 order by a0405 asc,a0402 asc", adoTaie, adOpenStatic, adLockReadOnly
'         Do While adoacc040.EOF = False
'            adoaccrpt418.CursorLocation = adUseClient
'            '2014/3/3 modify by sonia
'            'adoaccrpt418.Open "select * from accrpt418 where r41801 = '" & strUserNum & "' and r41802 = '" & adoacc080.Fields("a0801").Value & "' and r41803 = '" & adoacc090.Fields("a0901").Value & "' and r41805 = '" & adoacc040.Fields("a0405").Value & adoacc040.Fields("a0102").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'            'Modify By Sindy 2020/4/27
'            'If Text2 <> "" Then
'            '   adoaccrpt418.Open "select * from accrpt418 where r41801 = '" & strUserNum & "' and r41802 = '" & IIf(Text2 = "2", "J", Text2) & "' and r41803 = '" & adoacc090.Fields("a0901").Value & "' and r41805 = '" & adoacc040.Fields("a0405").Value & adoacc040.Fields("a0102").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'            If strCmp <> "" Then
'               If InStr(strCmp, "+") > 0 Then
'                  adoaccrpt418.Open "select * from accrpt418 where r41801 = '" & strUserNum & "' and r41802 In ('" & Replace(strCmp, "+", "','") & "') and r41803 = '" & adoacc090.Fields("a0901").Value & "' and r41805 = '" & adoacc040.Fields("a0405").Value & adoacc040.Fields("a0102").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'               Else
'                  adoaccrpt418.Open "select * from accrpt418 where r41801 = '" & strUserNum & "' and r41802 = '" & strCmp & "' and r41803 = '" & adoacc090.Fields("a0901").Value & "' and r41805 = '" & adoacc040.Fields("a0405").Value & adoacc040.Fields("a0102").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'               End If
'            '2020/4/27 END
'            Else
'               adoaccrpt418.Open "select * from accrpt418 where r41801 = '" & strUserNum & "' and r41803 = '" & adoacc090.Fields("a0901").Value & "' and r41805 = '" & adoacc040.Fields("a0405").Value & adoacc040.Fields("a0102").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'            End If
'            '2014/3/3 end
'            If adoaccrpt418.RecordCount = 0 Then
'               adoaccrpt418.AddNew
'               adoaccrpt418.Fields("r41801").Value = strUserNum
'               '2014/3/3 modify by sonia
'               'adoaccrpt418.Fields("r41802").Value = adoacc080.Fields("a0801").Value
'               adoaccrpt418.Fields("r41802").Value = strCmp 'IIf(Text2 = "2", "J", Text2) Modify By Sindy 2020/4/27
'               adoaccrpt418.Fields("r41803").Value = adoacc090.Fields("a0901").Value
'               adoaccrpt418.Fields("r41804").Value = Counter
'               adoaccrpt418.Fields("r41805").Value = adoacc040.Fields("a0405").Value & adoacc040.Fields("a0102").Value
'               adoaccrpt418.Fields("r41818").Value = "0"
'            End If
'            PutMonth
'            adoaccrpt418.UpdateBatch
'            adoaccrpt418.Close
'            adoacc040.MoveNext
'         Loop
'         If adoacc040.RecordCount <> 0 Then
'            adoaccrpt418.CursorLocation = adUseClient
'            adoaccrpt418.Open "select * from accrpt418", adoTaie, adOpenDynamic, adLockBatchOptimistic
'            adoaccrpt418.AddNew
'            adoaccrpt418.Fields("r41801").Value = strUserNum
'            '2014/3/3 modify by sonia
'            'adoaccrpt418.Fields("r41802").Value = adoacc080.Fields("a0801").Value
'            adoaccrpt418.Fields("r41802").Value = strCmp 'IIf(Text2 = "2", "J", Text2) Modify By Sindy 2020/4/27
'            adoaccrpt418.Fields("r41803").Value = adoacc090.Fields("a0901").Value
'            adoaccrpt418.Fields("r41804").Value = Counter
'            For intCounter = 1 To 13
'               adoaccrpt418.Fields(intCounter + 4) = ReportSum(4)
'            Next intCounter
'            adoaccrpt418.UpdateBatch
'            adoaccrpt418.Close
'            For intCounter = 1 To 12
'               If adoacc040.State = adStateOpen Then
'                  adoacc040.Close
'               End If
'               adoacc040.CursorLocation = adUseClient
'               '2014/3/3 modify by sonia a0403條件併入strSQL2
'               'adoacc040.Open "select sum(a0409), a0402 from acc040 where a0401 = " & Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) & " and a0402 = " & intCounter & " and a0403 = '" & adoacc080.Fields("a0801").Value & "' and a0404 = '" & adoacc090.Fields("a0901").Value & "'" & strSQL2 & " group by a0402", adoTaie, adOpenStatic, adLockReadOnly
'               adoacc040.Open "select sum(a0409), a0402 from acc040 where a0401 = " & Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) & " and a0402 = " & intCounter & " and a0404 = '" & adoacc090.Fields("a0901").Value & "'" & strSQL2 & " group by a0402", adoTaie, adOpenStatic, adLockReadOnly
'               If adoacc040.RecordCount <> 0 Then
'                  adoaccrpt418.CursorLocation = adUseClient
'                  '2014/3/3 modify by sonia
'                  'adoaccrpt418.Open "select * from accrpt418 where r41801 = '" & strUserNum & "' and r41802 = '" & adoacc080.Fields("a0801").Value & "' and r41803 = '" & adoacc090.Fields("a0901").Value & "' and r41805 = '" & ReportSum(24) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'                  If Text2 <> "" Then
'                     adoaccrpt418.Open "select * from accrpt418 where r41801 = '" & strUserNum & "' and r41802 = '" & IIf(Text2 = "2", "J", Text2) & "' and r41803 = '" & adoacc090.Fields("a0901").Value & "' and r41805 = '" & ReportSum(24) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'                  Else
'                     adoaccrpt418.Open "select * from accrpt418 where r41801 = '" & strUserNum & "' and r41803 = '" & adoacc090.Fields("a0901").Value & "' and r41805 = '" & ReportSum(24) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'                  End If
'                  '2014/3/3 end
'                  If adoaccrpt418.RecordCount = 0 Then
'                     adoaccrpt418.AddNew
'                     adoaccrpt418.Fields("r41818").Value = "0"
'                  End If
'                  adoaccrpt418.Fields("r41801").Value = strUserNum
'                  '2014/3/3 modify by sonia
'                  'adoaccrpt418.Fields("r41802").Value = adoacc080.Fields("a0801").Value
'                  adoaccrpt418.Fields("r41802").Value = IIf(Text2 = "2", "J", Text2)
'                  adoaccrpt418.Fields("r41803").Value = adoacc090.Fields("a0901").Value
'                  adoaccrpt418.Fields("r41804").Value = Counter
'                  adoaccrpt418.Fields("r41805").Value = ReportSum(24)
'                  PutMonth
'                  adoaccrpt418.UpdateBatch
'                  adoaccrpt418.Close
'               End If
'               adoacc040.Close
'            Next intCounter
'         End If
'         If adoacc040.State = adStateOpen Then
'            adoacc040.Close
'         End If
'         adoacc090.MoveNext
'         If adoacc090.EOF Then
'            adoacc090.MoveLast
'            adoaccrpt418.CursorLocation = adUseClient
'            adoaccrpt418.Open "select * from accrpt418", adoTaie, adOpenDynamic, adLockBatchOptimistic
'            adoaccrpt418.AddNew
'            adoaccrpt418.Fields("r41801").Value = strUserNum
'            '2014/3/3 modify by sonia
'            'adoaccrpt418.Fields("r41802").Value = adoacc080.Fields("a0801").Value
'            adoaccrpt418.Fields("r41802").Value = IIf(Text2 = "2", "J", Text2)
'            adoaccrpt418.Fields("r41803").Value = adoacc090.Fields("a0901").Value
'            adoaccrpt418.Fields("r41804").Value = Counter
'            For intCounter = 1 To 13
'               adoaccrpt418.Fields(intCounter + 4) = ReportSum(4)
'            Next intCounter
'            adoaccrpt418.UpdateBatch
'            adoaccrpt418.Close
'            For intCounter = 1 To 12
'               adoacc040.CursorLocation = adUseClient
'               '2014/3/3 modify by sonia a0403條件併入strSQL2
'               'adoacc040.Open "select sum(a0409), a0402 from acc040 where a0401 = " & Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) & " and a0402 = " & intCounter & " and a0403 = '" & adoacc080.Fields("a0801").Value & "'" & strSQL2 & " group by a0402", adoTaie, adOpenStatic, adLockReadOnly
'               adoacc040.Open "select sum(a0409), a0402 from acc040 where a0401 = " & Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) & " and a0402 = " & intCounter & strSQL2 & " group by a0402", adoTaie, adOpenStatic, adLockReadOnly
'               If adoacc040.RecordCount <> 0 Then
'                  adoaccrpt418.CursorLocation = adUseClient
'                  '2014/3/3 modify by sonia
'                  'adoaccrpt418.Open "select * from accrpt418 where r41801 = '" & strUserNum & "' and r41802 = '" & adoacc080.Fields("a0801").Value & "' and r41805 = '" & ReportSum(25) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'                  If Text2 <> "" Then
'                     adoaccrpt418.Open "select * from accrpt418 where r41801 = '" & strUserNum & "' and r41802 = '" & IIf(Text2 = "2", "J", Text2) & "' and r41805 = '" & ReportSum(25) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'                  Else
'                     adoaccrpt418.Open "select * from accrpt418 where r41801 = '" & strUserNum & "' and r41805 = '" & ReportSum(25) & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'                  End If
'                  '2014/3/3 end
'                  If adoaccrpt418.RecordCount = 0 Then
'                     adoaccrpt418.AddNew
'                     adoaccrpt418.Fields("r41818").Value = "0"
'                  End If
'                  adoaccrpt418.Fields("r41801").Value = strUserNum
'                  '2014/3/3 modify by sonia
'                  'adoaccrpt418.Fields("r41802").Value = adoacc080.Fields("a0801").Value
'                  adoaccrpt418.Fields("r41802").Value = IIf(Text2 = "2", "J", Text2)
'                  adoaccrpt418.Fields("r41803").Value = adoacc090.Fields("a0901").Value
'                  adoaccrpt418.Fields("r41804").Value = Counter
'                  adoaccrpt418.Fields("r41805").Value = ReportSum(25)
'                  PutMonth
'                  adoaccrpt418.UpdateBatch
'                  adoaccrpt418.Close
'               End If
'               adoacc040.Close
'            Next intCounter
'            adoaccrpt418.CursorLocation = adUseClient
'            adoaccrpt418.Open "select * from accrpt418", adoTaie, adOpenDynamic, adLockBatchOptimistic
'            adoaccrpt418.AddNew
'            adoaccrpt418.Fields("r41801").Value = strUserNum
'            '2014/3/3 modify by sonia
'            'adoaccrpt418.Fields("r41802").Value = adoacc080.Fields("a0801").Value
'            adoaccrpt418.Fields("r41802").Value = IIf(Text2 = "2", "J", Text2)
'            adoaccrpt418.Fields("r41803").Value = adoacc090.Fields("a0901").Value
'            adoaccrpt418.Fields("r41804").Value = Counter
'            For intCounter = 1 To 13
'               adoaccrpt418.Fields(intCounter + 4) = ReportSum(8)
'            Next intCounter
'            adoaccrpt418.UpdateBatch
'            adoaccrpt418.Close
'            adoacc090.MoveNext
'         End If
'      Loop
'      adoacc090.Close
'   '   adoacc080.MoveNext
'   'Loop
'   'adoacc080.Close
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub

''*************************************************
''  刪除報表資料
''
''*************************************************
'Private Sub Accrpt418Delete()
'   adoTaie.Execute "delete from accrpt418"
'End Sub

'*************************************************
'  畫線
'
'*************************************************
Private Sub PaintLine(strSign As String)
Dim intCounter As Integer

   For intCounter = 5 To 16
      adoaccrpt418.Fields(intCounter).Value = strSign
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
'  依月份存放至適當的位置中
'
'*************************************************
Private Sub PutMonth()
   Select Case adoacc040.Fields("a0402").Value
      Case 1
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt418.Fields("r41806").Value = Null
         Else
            adoaccrpt418.Fields("r41806").Value = adoacc040.Fields(0).Value
            adoaccrpt418.Fields("r41818").Value = Val(adoaccrpt418.Fields("r41818").Value) + Val(adoacc040.Fields(0).Value)
         End If
      Case 2
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt418.Fields("r41807").Value = Null
         Else
            adoaccrpt418.Fields("r41807").Value = adoacc040.Fields(0).Value
            adoaccrpt418.Fields("r41818").Value = Val(adoaccrpt418.Fields("r41818").Value) + Val(adoacc040.Fields(0).Value)
         End If
      Case 3
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt418.Fields("r41808").Value = Null
         Else
            adoaccrpt418.Fields("r41808").Value = adoacc040.Fields(0).Value
            adoaccrpt418.Fields("r41818").Value = Val(adoaccrpt418.Fields("r41818").Value) + Val(adoacc040.Fields(0).Value)
         End If
      Case 4
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt418.Fields("r41809").Value = Null
         Else
            adoaccrpt418.Fields("r41809").Value = adoacc040.Fields(0).Value
            adoaccrpt418.Fields("r41818").Value = Val(adoaccrpt418.Fields("r41818").Value) + Val(adoacc040.Fields(0).Value)
         End If
      Case 5
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt418.Fields("r41810").Value = Null
         Else
            adoaccrpt418.Fields("r41810").Value = adoacc040.Fields(0).Value
            adoaccrpt418.Fields("r41818").Value = Val(adoaccrpt418.Fields("r41818").Value) + Val(adoacc040.Fields(0).Value)
         End If
      Case 6
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt418.Fields("r41811").Value = Null
         Else
            adoaccrpt418.Fields("r41811").Value = adoacc040.Fields(0).Value
            adoaccrpt418.Fields("r41818").Value = Val(adoaccrpt418.Fields("r41818").Value) + Val(adoacc040.Fields(0).Value)
         End If
      Case 7
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt418.Fields("r41812").Value = Null
         Else
            adoaccrpt418.Fields("r41812").Value = adoacc040.Fields(0).Value
            adoaccrpt418.Fields("r41818").Value = Val(adoaccrpt418.Fields("r41818").Value) + Val(adoacc040.Fields(0).Value)
         End If
      Case 8
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt418.Fields("r41813").Value = Null
         Else
            adoaccrpt418.Fields("r41813").Value = adoacc040.Fields(0).Value
            adoaccrpt418.Fields("r41818").Value = Val(adoaccrpt418.Fields("r41818").Value) + Val(adoacc040.Fields(0).Value)
         End If
      Case 9
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt418.Fields("r41814").Value = Null
         Else
            adoaccrpt418.Fields("r41814").Value = adoacc040.Fields(0).Value
            adoaccrpt418.Fields("r41818").Value = Val(adoaccrpt418.Fields("r41818").Value) + Val(adoacc040.Fields(0).Value)
         End If
      Case 10
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt418.Fields("r41815").Value = Null
         Else
            adoaccrpt418.Fields("r41815").Value = adoacc040.Fields(0).Value
            adoaccrpt418.Fields("r41818").Value = Val(adoaccrpt418.Fields("r41818").Value) + Val(adoacc040.Fields(0).Value)
         End If
      Case 11
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt418.Fields("r41816").Value = Null
         Else
            adoaccrpt418.Fields("r41816").Value = adoacc040.Fields(0).Value
            adoaccrpt418.Fields("r41818").Value = Val(adoaccrpt418.Fields("r41818").Value) + Val(adoacc040.Fields(0).Value)
         End If
      Case 12
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt418.Fields("r41817").Value = Null
         Else
            adoaccrpt418.Fields("r41817").Value = adoacc040.Fields(0).Value
            adoaccrpt418.Fields("r41818").Value = Val(adoaccrpt418.Fields("r41818").Value) + Val(adoacc040.Fields(0).Value)
         End If
   End Select
End Sub

''*************************************************
''  執行報表之 Dll
''
''*************************************************
'Private Sub RunReportDll()
'   '2014/3/3 modify by sonia
'   'dllaccrpt418.Acc44l0 ReportTitle(418), Text2, Text1, Text3, Text4, Text5, Text6, adoacc080.Fields("r41802").Value, adoacc090.Fields("r41803").Value, strUserNum, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'   dllaccrpt418.Acc44l0 ReportTitle(418), IIf(Text2 = "2", "J", Text2), IIf(Text1 = "", "台一　專利商標/智權", Text1), Text3, Text4, Text5, Text6, IIf(Text2 = "2", "J", Text2), adoacc090.Fields("r41803").Value, strUserNum, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
'   Text2 = ""
'   Text1 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
'   Text7 = "Y"
   Text9 = "2"
   
   'Add by Mogan 2008/12/1 加年度,部門預設為TOT
   Text8 = strSrvDate(2) \ 10000
   Text3 = "TOT"
   Text4 = "TOT"
   'end 2008/12/1
   
   If Me.Visible = True Then
'      Text2.SetFocus
      'Add By Sindy 2020/4/27
      CboCmp.ListIndex = -1
      CboCmp.SetFocus
      '2020/4/27 END
   End If
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
'Modify by Morgan 2008/12/1
Public Function FormCheck() As Boolean
Dim bolCancel As Boolean
   
   '2014/3/3 cancel by sonia
   'If Text2 = "" Then
   '   MsgBox "公司別不可空白!!", vbExclamation
   '   Text2.SetFocus
   '   Exit Function
   'End If
   '2014/3/3 end
   'Add by Sindy 2020/4/27 +公司別判斷
   If Trim(CboCmp) <> MsgText(601) Then
      Call CboCmp_Validate(bolCancel)
      If bolCancel = True Then
          Exit Function
      End If
   End If
   'end 2020/4/27
   
   If Text8 = "" Then
      MsgBox "預算年度不可空白!!", vbExclamation
      Text8.SetFocus
      Exit Function
   End If
   
   If Text9 = "" Then
      MsgBox "資料來源不可空白!!", vbExclamation
      Text9.SetFocus
      Exit Function
   '2014/3/3 add by sonia
   Else
      'Modify By Sindy 2020/4/27 Mark
'      If Text7 <> "Y" And Text9 = "1" Then
'         MsgBox "選擇產生Excel檔案才可選實績資料!!", vbExclamation
'         Text9.SetFocus
'         Exit Function
'      End If
   '2014/3/3 end
   End If
   
   FormCheck = True
End Function

''Add by Morgan 2008/11/26
'Private Sub Text7_GotFocus()
'   TextInverse Text7
'   CloseIme
'End Sub
'
'Private Sub Text7_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
'      KeyAscii = 0
'      Beep
'   End If
'End Sub

'Add by Morgan 2008/11/26
'*************************************************
'  轉成Excel檔案
'
'*************************************************
Private Sub Export2Excel(adoRst As ADODB.Recordset)
   
   Dim xlsSalesPoint As New Excel.Application
   Dim wksaccrpt418 As New Worksheet
   Dim xlsSelect As Selection
   Dim strFileName As String
   Dim iRow As Integer, iRowCount As Integer
   Dim stCode1 As String, stCode2 As String
   Dim stRptName As String, stCellID As String, stCellFormat As String
   Dim bolShowZezo As Boolean, stCode1Name As String
   Dim Rc As String '欄位座標
   
   stCellFormat = "#,##0.00 ;[紅色]-#,##0.00 "
   
   If Text9 = "1" Then
      stRptName = Val(Text8) & "年度實績預算資料表"
   Else
      stRptName = Val(Text8) & "年度預算資料表"
   End If
   
   strFileName = strExcelPath & stRptName & Format(Now, "yyyymmddhhmmss") & MsgText(43)
   If Dir(strFileName) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strFileName
   End If
   
   xlsSalesPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsSalesPoint.Workbooks.add
   Set wksaccrpt418 = xlsSalesPoint.Worksheets(1)
   With wksaccrpt418
      iRow = 1
      .Range("a" & iRow).Value = stRptName
      Rc = Chr(Asc("a") + 15) & iRow '總共16個欄位+1個隱藏欄位(判斷合計公式用)
      With .Range("a" & iRow & ":" & Rc)
         .Font.Size = 18
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
'         .VerticalAlignment = xlBottom
'         .WrapText = False
'         .Orientation = 0
'         .AddIndent = False
'         .ShrinkToFit = False
         .MergeCells = True
      End With
      
      iRow = iRow + 2
      .Range("f" & iRow).Value = "公司別："
      With .Range("f" & iRow & ":g" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlRight
         .MergeCells = True
      End With

      '2014/3/3 modify by sonia
      '.Range("h" & iRow).Value = Text2 & "  " & Text1
      'Modify By Sindy 2020/4/27
      '.Range("h" & iRow).Value = IIf(Text2 = "2", "J", Text2) & "  " & IIf(Text1 = "", "台一　專利商標/智權", Text1)
      .Range("h" & iRow).Value = strCmp & "  " & strCmpN
      '2020/4/27 END
      With .Range("h" & iRow & ":l" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
         .MergeCells = True
      End With
      
      iRow = iRow + 1
      .Range("f" & iRow).Value = "會計科目："
      With .Range("f" & iRow & ":g" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlRight
         .MergeCells = True
      End With
      .Range("h" & iRow).Value = Text5 & " ~ " & Text6
      With .Range("h" & iRow & ":j" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
         .MergeCells = True
      End With
      
      iRow = iRow + 2
      .Range("a" & iRow).Value = "會計科目"
      
      'Add by Morgan 2009/1/10
      .Range("b" & iRow).Value = "當年度預算總額"
      .Range("c" & iRow).Value = "合計"
      
      For intI = 1 To 12
         .Range(Chr(Asc("c") + intI) & iRow).Value = PUB_ChgNumber2Chinese(str(intI)) & "月"
      Next
      Rc = Chr(Asc("c") + 13) & iRow
      .Range(Rc).Value = "差額"
      
      With .Range("a" & iRow & ":" & Rc)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
      End With

      stCode1 = ""
      stCode2 = ""
      Set RsTemp = adoRst.Clone
      adoRst.MoveFirst
      Do While Not adoRst.EOF
         bolShowZezo = True
         '主科目不同
         If stCode1 <> Left("" & adoRst.Fields(0), 4) Then
            If stCode1 <> "" Then
               iRow = iRow + 1
               '有子科目的印合計
               If stCode2 <> stCode1 Then
                  .Range("a" & iRow).Value = stCode1Name & "－" & "合計"
                  Rc = Chr(Asc("c") + 14) & iRow
                  .Range(Rc).Value = "　"
                  '設定列合計公式
                  For intI = 1 To 15
                     stCellID = Chr(Asc("a") + intI) & iRow
                     .Range(stCellID).NumberFormatLocal = stCellFormat
                     .Range(stCellID).FormulaR1C1 = "=SUM(R[-" & iRowCount & "]C:R[-1]C)"
                  Next
                  
                  iRow = iRow + 1
               End If
            End If
            
            stCode1 = Left("" & adoRst.Fields(0), 4)
            stCode1Name = "" & adoRst.Fields(2)
            iRowCount = 0 '合計列數歸零
         End If
                  
         '科目不同時跳列
         If stCode2 <> "" & adoRst.Fields(0) Then
            stCode2 = "" & adoRst.Fields(0)
            '主科目下若有子科目時[零]不顯示
            If stCode1 = stCode2 Then
               RsTemp.MoveFirst
               RsTemp.Find "a0101 > '" & stCode1 & "'"
               If Not RsTemp.EOF Then
                  If stCode1 = Left("" & RsTemp.Fields(0), 4) Then
                     bolShowZezo = False
                  End If
               End If
            End If
            
            If bolShowZezo = True Then
               iRow = iRow + 1
               iRowCount = iRowCount + 1
               '會計科目
               .Range("a" & iRow).Value = Replace("" & adoRst.Fields(2), stCode1Name & "－", "")
               For intI = 1 To 12
                  stCellID = Chr(Asc("c") + intI) & iRow
                  .Range(stCellID).NumberFormatLocal = stCellFormat
                  .Range(stCellID).Value = 0
               Next
               '設定列合計公式
               stCellID = "c" & iRow
               .Range(stCellID).NumberFormatLocal = stCellFormat
               .Range(stCellID).FormulaR1C1 = "=SUM(RC[1]:RC[12])"
               '設定差額公式
               stCellID = Chr(Asc("a") + 15) & iRow
               .Range(stCellID).NumberFormatLocal = stCellFormat
               .Range(stCellID).FormulaR1C1 = "=(RC2-RC3)"
            End If
         End If
         
         '預算月份
         intI = Val("" & adoRst.Fields(1))
         If intI > 0 Then
            stCellID = Chr(Asc("c") + intI) & iRow
            .Range(stCellID).NumberFormatLocal = stCellFormat
            .Range(stCellID).Value = "" & adoRst.Fields(3)
         End If
         adoRst.MoveNext
      Loop
      
      iRow = iRow + 1
      
      '有子科目的印合計
      If stCode2 <> stCode1 Then
         .Range("a" & iRow).Value = stCode1Name & "－" & "合計"
         For intI = 1 To 15
            stCellID = Chr(Asc("a") + intI) & iRow
            .Range(stCellID).NumberFormatLocal = stCellFormat
            .Range(stCellID).FormulaR1C1 = "=SUM(R[-" & iRowCount & "]C:R[-1]C)"
         Next
         iRow = iRow + 1
      End If
      
      iRow = iRow + 1
      .Range("a" & iRow).Value = "總計"
      For intI = 1 To 15
         stCellID = Chr(Asc("a") + intI) & iRow
         .Range(stCellID).NumberFormatLocal = stCellFormat
         .Range(stCellID).FormulaR1C1 = "=SUM(R7C:R[-1]C)-SUMIF(R7C[" & 16 - intI & "]:R[-1]C[" & 16 - intI & "],""　"",R7C:R[-1]C)"
      Next
      Rc = Chr(Asc("a") + 15)
      .Columns("a:" & Rc).EntireColumn.AutoFit
      
      'Removed by Morgan 2013/12/11 預設印表機可能不是點陣指定紙張會有錯誤,取消預設列印格式--婧瑄
      '.PageSetup.PaperSize = PUB_GetPaperSize(7) '設定紙張　15x11
      '.PageSetup.Orientation = xlPortrait '直印
      '.PageSetup.PrintTitleRows = "$1:$6" '表頭保留6列
      ''Modified by Morgan 2011/12/14 修更列印範圍
      '.PageSetup.PrintArea = "$A$1:$P$" & iRow '設定列印範圍
      '.PageSetup.LeftMargin = xlsSalesPoint.InchesToPoints(0.2) '左邊界
      '.PageSetup.RightMargin = xlsSalesPoint.InchesToPoints(0.2) '右邊界
      '.PageSetup.Zoom = 100 '縮放比例
      'end 2013/12/11
      
   End With
   'Modify by Amy 2016/06/23 +判斷版本
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   'end 2016/06/23
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   StatusClear
   
   MsgBox "Excel檔案已產生！（檔案位置：" & strFileName & "）"
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
   CloseIme
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
   CloseIme
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
      Beep
   End If
End Sub

