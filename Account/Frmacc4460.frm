VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc4460 
   AutoRedraw      =   -1  'True
   Caption         =   "綜合損益表"
   ClientHeight    =   1710
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1710
   ScaleWidth      =   5160
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   300
      Width           =   3350
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
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   1200
      Width           =   4692
   End
   Begin VB.CheckBox Check1 
      Caption         =   "是否產生Excel檔案"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1710
      Width           =   3450
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   855
      _ExtentX        =   1517
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
      Top             =   840
      Width           =   855
      _ExtentX        =   1517
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
      Left            =   300
      TabIndex        =   8
      Top             =   1740
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1680
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
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   255
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
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   5
      Top             =   300
      Width           =   675
   End
End
Attribute VB_Name = "Frmacc4460"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoacc021 As New ADODB.Recordset
Public adoaccrpt407 As New ADODB.Recordset
Public adoacc010 As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim lngStartDate, lngEndDate As Long
Dim intCounter As Integer
'Dim dllaccrpt407 As Object 'Mark by Amy 2018/12/27 不使用
Dim strPrinter As String
'Add by Amy 2018/12/28
Dim strF, intWidth()
Dim i As Integer, intField As Integer, intRow As Integer, intTitleRow As Integer
Dim strCmp As String, strCmpN As String 'Add by Amy 2020/04/23

'Add by Amy 2020/04/23
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
'end 2020/04/23

Private Sub Command1_Click()
'Modify by Amy 2020/04/23
Dim bolShowMsg As Boolean

On Error GoTo Checking
   If FormCheck(bolShowMsg) = False Then
      If bolShowMsg = False Then MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   strCmp = "": strCmpN = ""
   If Trim(CboCmp) <> MsgText(601) Then
       strCmp = CboCmp
       If InStr(strCmp, "　") > 0 Then
             strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
       End If
   End If
   'end 2020/04/23
   Screen.MousePointer = vbHourglass
   Accrpt407Delete
   'Modify by Amy 2018/12/28 ProduceData設為function,已有Excel 列印不使用
   If ProduceData = True Then
       'Add By Sindy 2013/5/27
       If Check1.Value = 1 Then
          Call IsExcelSave
    '   Else
    '   '2013/5/27 End
    '      PUB_SetOsDefaultPrinter Combo1
    '      '2014/1/23 modify by sonia
    '      'dllaccrpt407.Acc4460 ReportTitle(407), Text6, Text7, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
    '      dllaccrpt407.Acc4460 ReportTitle(407), IIf(Text6 = "2", "J", Text6), IIf(Text7 = "", "台一　專利商標/智權", Text7), MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
    '      PUB_SetOsDefaultPrinter strPrinter
        End If
   End If
   'end 2018/12/28
   Screen.MousePointer = vbDefault
   FormClear
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101) 'Mark by Amy 2018/12/28
   Exit Sub
   
Checking:
   PUB_SetOsDefaultPrinter strPrinter
   Screen.MousePointer = vbDefault
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   'Mark by Amy 2020/04/23
'   If KeyCode <> vbKeyEscape Then
'      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
'   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5280 '5250
   Me.Height = 2270 'Modify by Amy 2018/12/28 原:3270
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2020/04/23 公司別下拉
   CboCmp.Clear
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/04/23
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   MaskEdBox2.Mask = Mid(DFormat, 1, 6)
   Check1.Value = 1 'Add by Amy 2016/09/29 預設勾Excel-瑞婷
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101) 'Mark by Amy 2020/04/23
   'Set dllaccrpt407 = CreateObject("AccReport.ReportSelect") 'Mark by Amy 2018/12/28 不使用
   
   PUB_SetPrinter Me.Name, Combo1, strPrinter
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   'Set dllaccrpt407 = Nothing 'Mark by Amy 2018/12/28 不使用
   Set Frmacc4460 = Nothing
End Sub

'Mark by Amy 2020/04/23 公司別改下拉
'Private Sub Text6_Change()
'   '2014/1/23 modify by sonia
'   'If Text6 = MsgText(601) Then
'   '   Exit Sub
'   'End If
'   'Text7 = A0802Query(Text6)
'   Select Case Text6
'      Case "1"
'         Text7 = A0802Query(Text6)
'      Case "2"
'         Text7 = A0802Query("J")
'      Case ""
'         Text7 = "台一　專利商標/智權"
'   End Select
'   '2014/1/23 end
'End Sub
'
'Private Sub Text6_GotFocus()
'   TextInverse Text6
'End Sub
'
''2014/1/23 add by sonia
'Private Sub Text6_KeyPress(KeyAscii As Integer)
'   'Modify by Amy 2018/12/28 畫面公司別加3.4選項
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> Asc("4") Then
'      KeyAscii = 0
'   End If
'End Sub
''2014/1/23 end
'end 2020/04/23

'*************************************************
'  產生報表資料
'
'*************************************************
'Add by Amy 2018/12/28 畫面公司別加3.4選項,加總改寫法
Private Function ProduceData() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim intQ As Integer
    Dim strQ As String, strWhere As String, strField As String

    ProduceData = False: intCounter = 0
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
    '刪除暫存檔
    strQ = "select * from accrpt407 Where R40701='" & strUserNum & "'"
    If adoaccrpt407.State = adStateOpen Then adoaccrpt407.Close
    adoaccrpt407.CursorLocation = adUseClient
    adoaccrpt407.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
   
    strField = ",R40705"
    'Modify by Amy 2020/04/23 公司別改下拉 原:Text6
'    If Text6 = "1" Then
'        strWhere = "and (a0109 is null or a0109='1') "
'    End If
'    If Text6 = "2" Then
'        strWhere = "and (a0109 is null or a0109='J') "
'    End If
'    If Text6 = "3" Then strField = ""
    If strCmp <> MsgText(601) Then
        If InStr(strCmp, "+") > 0 Then
            strWhere = "and (a0109 is null or a0109 In ('" & Replace(strCmp, "+", "','") & "') )"
            strField = ""
        Else
            strWhere = "and (a0109 is null or a0109='" & strCmp & "')"
        End If
   End If
   'end 2020/04/23
    
'Memo 取消instr(a0102,'不用')=0,改為沒有數字才不出現,ex:106/11之1132備抵呆帳－應收票據/不用 有數字
'------------------------------------------------
' 營業收入明細
'------------------------------------------------
    strQ = "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' " & strWhere & " order by a0101 asc"
    If adoacc010.State = adStateOpen Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        Call Accrpt407Save1("" & adoacc010.Fields("a0101"), "" & adoacc010.Fields("a0102"))
        adoacc010.MoveNext
    Loop
    adoacc010.Close
    
'------------------------------------------------
' 營業支出明細
'------------------------------------------------
    strQ = "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' " & strWhere & " order by a0101 asc"
    If adoacc010.State = adStateOpen Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        Call Accrpt407Save1("" & adoacc010.Fields("a0101"), "" & adoacc010.Fields("a0102"))
        adoacc010.MoveNext
    Loop
    adoacc010.Close
    
'------------------------------------------------
' 非營業收入
'------------------------------------------------
    strQ = "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' " & strWhere & " order by a0101 asc"
    If adoacc010.State = adStateOpen Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        Call Accrpt407Save1("" & adoacc010.Fields("a0101"), "" & adoacc010.Fields("a0102"))
        adoacc010.MoveNext
    Loop
   adoacc010.Close
   
'------------------------------------------------
' 非營業支出明細
'------------------------------------------------
    strQ = "select * from acc010 where a0101 >= '72' and a0101 < '8' and a0104 = '3' " & strWhere & " order by a0101 asc"
    If adoacc010.State = adStateOpen Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        Call Accrpt407Save1("" & adoacc010.Fields("a0101"), "" & adoacc010.Fields("a0102"))
        adoacc010.MoveNext
    Loop
    adoacc010.Close
    
'------------------------------------------------
' 營業收入小計
'------------------------------------------------
    strQ = "Insert Into Accrpt407(R40701,R40702,R40703,R40704" & strField & ") " & _
              "Select '" & strUserNum & "','4T','營業收入',Sum(R40704)" & strField & " From Accrpt407 Where R40701='" & strUserNum & "' And R40702>= '4' and R40702 < '5' " & IIf(strField <> "", " Group by R40705", "") & " Having Sum(R40704)<>0"
    cnnConnection.Execute strQ
'------------------------------------------------
' 營業支出小計
'------------------------------------------------
    strQ = "Insert Into Accrpt407(R40701,R40702,R40703,R40704" & strField & ") " & _
              "Select '" & strUserNum & "','6T','營業費用',Sum(R40704)" & strField & " From Accrpt407 Where R40701='" & strUserNum & "' And R40702>= '6' and R40702 < '7' " & IIf(strField <> "", " Group by R40705", "") & " Having Sum(R40704)<>0"
    cnnConnection.Execute strQ
'------------------------------------------------
' 營業損益
'------------------------------------------------
    strQ = "Insert Into Accrpt407(R40701,R40702,R40703,R40704" & strField & ") " & _
              "Select '" & strUserNum & "','6ZT','營業損益',Nvl(T1,0)-Nvl(T2,0) " & IIf(strField <> "", ",C1", "") & " From" & _
              "(Select '6ZT' as AccNo,'4T',Sum(R40704) as T1" & IIf(strField <> "", strField, ",''") & " as C1 From Accrpt407 Where R40701='" & strUserNum & "' And R40702= '4T' " & IIf(strField <> "", " Group by R40705", "") & ")" & _
             ",(Select '6ZT' as AccNo1,'6T',Sum(R40704) as T2" & IIf(strField <> "", strField, ",''") & "  as C2 From Accrpt407 Where R40701='" & strUserNum & "' And R40702= '6T' " & IIf(strField <> "", " Group by R40705", "") & ")" & _
            " Where AccNo=AccNo1(+) And C1=C2(+)"
    cnnConnection.Execute strQ
'------------------------------------------------
' 非營業收入小計
'------------------------------------------------
    strQ = "Insert Into Accrpt407(R40701,R40702,R40703,R40704" & strField & ") " & _
              "Select '" & strUserNum & "','71T','營業外收入',Sum(R40704)" & strField & " From Accrpt407 Where R40701='" & strUserNum & "' And R40702>= '71' and R40702 < '72' " & IIf(strField <> "", " Group by R40705", "") & " Having Sum(R40704)<>0"
    cnnConnection.Execute strQ
'------------------------------------------------
' 非營業支出小計
'------------------------------------------------
    strQ = "Insert Into Accrpt407(R40701,R40702,R40703,R40704" & strField & ") " & _
              "Select '" & strUserNum & "','72T','營業外支出',Sum(R40704)" & strField & " From Accrpt407 Where R40701='" & strUserNum & "' And R40702>= '72' and R40702 < '8' " & IIf(strField <> "", " Group by R40705", "") & " Having Sum(R40704)<>0"
    cnnConnection.Execute strQ
'------------------------------------------------
' 稅前損益
'------------------------------------------------
    strQ = "Insert Into Accrpt407(R40701,R40702,R40703,R40704" & strField & ") " & _
              "Select '" & strUserNum & "','ZZT','稅前淨損益',Nvl(T1,0)-Nvl(T2,0) " & IIf(strField <> "", ",C1", "") & " From" & _
              "(Select 'ZZT' as AccNo,'6ZT',Sum(R40704) as T1" & IIf(strField <> "", strField, ",''") & " as C1 From Accrpt407 Where R40701='" & strUserNum & "' And (R40702= '6ZT'  or R40702= '71T') " & IIf(strField <> "", " Group by R40705", "") & ")" & _
             ",(Select 'ZZT' as AccNo1,'72T',Sum(R40704) as T2" & IIf(strField <> "", strField, ",''") & " as C2 From Accrpt407 Where R40701='" & strUserNum & "' And R40702= '72T' " & IIf(strField <> "", " Group by R40705", "") & ")" & _
            " Where AccNo=AccNo1(+) And C1=C2(+)"
    cnnConnection.Execute strQ
'------------------------------------------------
' 符號
'------------------------------------------------
    strExc(1) = ""
    'Modify by Amy 2020/04/23 公司別改抓變數
'    If Text6 = "1" Then strExc(1) = ",'1'"
'    If Text6 = "2" Then strExc(1) = ",'J'"
'    If Text6 = "4" Then strField = ""
    If strCmp <> MsgText(601) Then
        If InStr(strCmp, "+") = 0 Then
            strExc(1) = ",'" & strCmp & "'"
        End If
    '原項目4.顯示 1/J/Sum(1+J)
    Else
        strField = ""
    End If
    strExc(0) = "Select Distinct SubStr(R40702,1,1) From Accrpt407 Where R40701='" & strUserNum & "' And R40702 in ('4T','6T') And R40704<>'0' " & _
           "Union Select Distinct SubStr(R40702,1,2) From Accrpt407 Where R40701='" & strUserNum & "' And R40702 in ('71T','72T') And R40704<>'0' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strExc(0))
    If intQ = 1 Then
        RsQ.MoveFirst
        Do While Not RsQ.EOF
            strQ = "Insert Into Accrpt407(R40701,R40702" & strField & ") Values('" & strUserNum & "','" & RsQ.Fields(0) & "S'" & strExc(1) & ") "
            cnnConnection.Execute strQ
            strQ = "Insert Into Accrpt407(R40701,R40702" & strField & ") Values('" & strUserNum & "','" & RsQ.Fields(0) & "V'" & strExc(1) & ") "
            cnnConnection.Execute strQ
            RsQ.MoveNext
        Loop
    End If
    RsQ.Close
    strQ = "Insert Into Accrpt407(R40701,R40702" & strField & ") " & _
               "Select '" & strUserNum & "','6ZV'" & strExc(1) & " From Dual " & _
   "Union Select '" & strUserNum & "','ZZV'" & strExc(1) & " From Dual "
    cnnConnection.Execute strQ

    StatusClear
    ProduceData = True
End Function

Private Sub IsExcelSave()
    Dim xlsSalesPoint As New Excel.Application
    Dim wksAccrpt407 As New Worksheet
    Dim stSQL As String, strFileName As String, stCellFormat As String, stRptName As String
    Dim stAccOld As String, stTmp As String, stCol As String, stIncome As String, stCost As String
    Dim intStartR As Integer, intEndR As Integer
    Dim bolPer As Boolean '顯示比例
    Dim ii As Integer, arrTmp 'Add by Amy 2020/04/23
    
On Error GoTo ErrHnd
    
    'Modify by Amy 2020/04/23 公司別改抓變數 原:Text6=4(顯示1/J/1+J,原1+J改用公式)
    '公司別=4先抓會計科目串各公司值
    If strCmp = MsgText(601) Then
'        ReDim strF(4): ReDim intWidth(4)
'        strF = Array("會計科目", "科目名稱", A0802Query("1"), A0802Query("J"), "台一　專利商標/智權")
'        intWidth = Array(11, 18, 22.5, 22.5, 22.5)
        strCmpN = GetAccReportCmpN(strCmp, , True)
        strFileName = Replace(strCmpN, "/", " ")
        strF = Split("會計科目,科目名稱," & Replace(strCmpN, "/", ",") & "," & strCmpN, ",")
    Else
        'ReDim strF(2): ReDim intWidth(2)
        If InStr(strCmp, "+") > 0 Then
            arrTmp = Split(strCmp, "+")
            For ii = LBound(arrTmp) To UBound(arrTmp)
                strCmpN = strCmpN & "/ " & A0802Query("" & arrTmp(ii), True)
            Next ii
            strCmpN = Mid(strCmpN, 3)
            strFileName = Replace(strCmpN, "/", "")
        End If
        If strCmpN = MsgText(601) Then
            strCmpN = A0802Query(strCmp)
            strFileName = A0802Query(strCmp, True)
        End If
        strF = Split("會計科目,科目名稱,金額 ", ",")
    End If
    
    ReDim intWidth(UBound(strF))
    For ii = LBound(strF) To UBound(strF)
        Select Case strF(ii)
            Case "會計科目"
                intWidth(ii) = 10
            Case "科目名稱"
                intWidth(ii) = 18
            Case Else
                If strCmp = MsgText(601) Then
                    If ii = UBound(strF) Then
                        intWidth(ii) = 18
                    Else
                        intWidth(ii) = 13
                    End If
                Else
                    intWidth(ii) = 22.5
                End If
        End Select
    Next ii
    
    '畫面公司別有值(會寫公司編號) 或 選1+J (不會寫公司編號)
    stSQL = "Select R40702,R40703,R40704 From Accrpt407 Where R40701='" & strUserNum & "' Order by r40702 asc"
    '抓所有公司別(會寫公司編號)
    If strCmp = MsgText(601) Then
        '"(Select R40702,Sum(R40704) as ATol From Accrpt407 Where R40701='" & strUserNum & "' Group by R40702) TJ " ->改公式
        stSQL = "Select a.R40702,a.R40703,Nvl(TTol,0) TTol,Nvl(JTol,0) JTol,Nvl(LTol,0) LTol,'' From" & _
                        "(Select Distinct R40702,R40703 From Accrpt407 Where R40701='" & strUserNum & "' ) a," & _
                        "(Select R40702,R40704 as TTol From Accrpt407 Where R40701='" & strUserNum & "' and R40705='1' ) T," & _
                        "(Select R40702,R40704 as JTol From Accrpt407 Where R40701='" & strUserNum & "' and R40705='J') J," & _
                        "(Select R40702,R40704 as LTol From Accrpt407 Where R40701='" & strUserNum & "' and R40705='L') L " & _
                      "Where a.R40702=T.R40702(+) And a.R40702=J.R40702(+) And a.R40702=L.R40702(+) " & _
                      "Order by R40702 asc"
    End If
    'end 2020/04/23
    
    If adoaccrpt407.State = adStateOpen Then adoaccrpt407.Close
    adoaccrpt407.CursorLocation = adUseClient
    adoaccrpt407.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
    If adoaccrpt407.RecordCount <= 0 Then
       adoaccrpt407.Close
       Exit Sub
    End If
   
    stCellFormat = "#,##0.00 ;[紅色]-#,##0.00 "
    stRptName = ReportTitle(407)
    intField = 65: intRow = 1
    
    'Mark by Amy 2020/04/23 公司別改抓變數搬至上面
'    strFileName = "台一"
'    If Text6 = "1" Then strFileName = "專利商標"
'    If Text6 = "2" Then strFileName = "智權"
    'end 2020/04/23
    
    strFileName = strExcelPath & Trim(Replace(stRptName, "*", "")) & " " & strFileName & Format(Now, "yyyymmddhhmmss") & MsgText(43)
    If Dir(strFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strFileName
    End If
   
    xlsSalesPoint.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
    xlsSalesPoint.Workbooks.add
    Set wksAccrpt407 = xlsSalesPoint.Worksheets(1)
    Call SetField(wksAccrpt407)
    intStartR = intRow
    adoaccrpt407.MoveFirst
    Do While Not adoaccrpt407.EOF
        bolPer = False
        '符號
        If Right("" & adoaccrpt407.Fields(0), 1) = "S" Or Right("" & adoaccrpt407.Fields(0), 1) = "V" Then
            'Modify by Amy 2020/04/23 公司別改抓變數 原:Text6 = "4"
            If strCmp = MsgText(601) And Right("" & adoaccrpt407.Fields(0), 1) = "V" And "" & adoaccrpt407.Fields(0) <> "6ZV" And Left("" & adoaccrpt407.Fields(0), 1) <> "7" Then
                bolPer = True
                Select Case Left("" & adoaccrpt407.Fields(0), 1)
                    Case "4"
                        stTmp = "收入占比"
                    Case "6"
                        stTmp = "費用占比"
                    Case "Z"
                        stTmp = "損益占比"
                End Select
                wksAccrpt407.Range(Chr(intField + GetValue("科目名稱")) & intRow).Value = stTmp
            End If
            For i = GetValue("科目名稱") + 1 To UBound(strF)
                 stTmp = "－－－－－－"
                 If Right("" & adoaccrpt407.Fields(0), 1) = "V" Then
                    stTmp = "＝＝＝＝＝＝"
                    If bolPer = True Then
                        stTmp = "=" & Chr(intField + i) & intRow - 1 & "/" & Chr(intField + UBound(strF)) & intRow - 1
                        If i = UBound(strF) Then stTmp = ""
                    End If
                 End If
                 wksAccrpt407.Range(Chr(intField + i) & intRow).Value = stTmp
                 wksAccrpt407.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlRight
                 If bolPer = True And stTmp <> MsgText(601) Then
                    wksAccrpt407.Range(Chr(intField + i) & intRow).NumberFormatLocal = "0.00%"
                    wksAccrpt407.Range(Chr(intField + i) & intRow).Interior.ColorIndex = 40 '設置儲存格填充色(膚)
                 End If
                 If Right("" & adoaccrpt407.Fields(0), 1) = "S" Then intEndR = intRow - 1
            Next i
        '合計
        ElseIf InStr("" & adoaccrpt407.Fields(0), "T") > 0 Then
            For i = GetValue("科目名稱") To UBound(strF)
                If i = GetValue("科目名稱") Then
                    stTmp = "" & adoaccrpt407.Fields(i)
                    stCol = Chr(intField + i)
                    If "" & adoaccrpt407.Fields(0) = "4T" Then
                        stIncome = Chr(intField + i) & intRow
                    ElseIf "" & adoaccrpt407.Fields(0) = "6T" Then
                        stCost = Chr(intField + i) & intRow
                    ElseIf "" & adoaccrpt407.Fields(0) = "71T" Then
                        stIncome = stIncome & "+" & Chr(intField + i) & intRow
                    ElseIf "" & adoaccrpt407.Fields(0) = "72T" Then
                        stCost = Chr(intField + i) & intRow
                    End If
                ElseIf "" & adoaccrpt407.Fields(0) = "6ZT" Or "" & adoaccrpt407.Fields(0) = "ZZT" Then
                    stTmp = "=" & IIf(stIncome <> MsgText(601), Replace(stIncome, Chr(intField + GetValue("科目名稱")), Chr(intField + i)), "") & _
                                            IIf(stCost <> MsgText(601), "-" & Replace(stCost, Chr(intField + GetValue("科目名稱")), Chr(intField + i)), "")
                    If i = UBound(strF) Then
                        stIncome = Chr(intField + GetValue("科目名稱")) & intRow
                        stCost = ""
                    End If
                Else
                    stTmp = Replace(stCol, stCol, Chr(intField + i))
                    stTmp = "=Sum(" & stTmp & intStartR & ":" & stTmp & intEndR & ")"
                End If
                wksAccrpt407.Range(Chr(intField + i) & intRow).Value = stTmp
                If i <= GetValue("科目名稱") Then
                    wksAccrpt407.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlLeft
                Else
                    wksAccrpt407.Range(Chr(intField + i) & intRow).NumberFormatLocal = stCellFormat
                End If
            Next i
        '資料
        Else
            If stAccOld <> MsgText(601) And (Left(stAccOld, 1) <> Left("" & adoaccrpt407.Fields(0), 1) _
              Or (Left(stAccOld, 2) = "71" And Left(stAccOld, 2) <> Left("" & adoaccrpt407.Fields(0), 2))) Then
                intStartR = intRow
            End If
            For i = LBound(strF) To UBound(strF)
                stTmp = "" & adoaccrpt407.Fields(i) 'Modify by Amy 2020/04/23
                'Add by Amy 2020/04/23 公司別空白最後一欄顯示加總公式
                If i = UBound(strF) And strCmp = MsgText(601) Then
                    stTmp = "=Sum(" & Chr(intField + GetValue("科目名稱") + 1) & intRow & ":" & Chr(intField + UBound(strF) - 1) & intRow & ")"
                ElseIf i <= GetValue("科目名稱") Then
                    wksAccrpt407.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlLeft
                Else
                    wksAccrpt407.Range(Chr(intField + i) & intRow).NumberFormatLocal = stCellFormat
                End If
                wksAccrpt407.Range(Chr(intField + i) & intRow).Value = stTmp 'Modify by Amy 2020/04/23
            Next i
        End If
        intRow = intRow + 1
        stAccOld = "" & adoaccrpt407.Fields(0)
        adoaccrpt407.MoveNext
    Loop
    adoaccrpt407.Close
    'Add by Amy 2020/04/23 顯示所有公司字型改小(因欄位多)
    If strCmp = MsgText(601) Then
        wksAccrpt407.Range(Chr(intField) & intTitleRow + 1 & ":" & Chr(intField + UBound(strF)) & intRow).Font.Size = 11
    End If
    Call SetField(wksAccrpt407, True)
    wksAccrpt407.PageSetup.PaperSize = 9 '設定紙張 A4
    wksAccrpt407.PageSetup.Orientation = xlPortrait '直印
    wksAccrpt407.PageSetup.PrintTitleRows = "$1:$" & intTitleRow '表頭保留
    wksAccrpt407.PageSetup.LeftMargin = xlsSalesPoint.InchesToPoints(0.5)
    wksAccrpt407.PageSetup.RightMargin = xlsSalesPoint.InchesToPoints(0.5)
    wksAccrpt407.PageSetup.TopMargin = xlsSalesPoint.InchesToPoints(0.3)
    wksAccrpt407.PageSetup.BottomMargin = xlsSalesPoint.InchesToPoints(0.3)
    wksAccrpt407.PageSetup.HeaderMargin = xlsSalesPoint.InchesToPoints(0.5)
    wksAccrpt407.PageSetup.FooterMargin = xlsSalesPoint.InchesToPoints(0.5)
    wksAccrpt407.PageSetup.Zoom = 100 '縮放比例
    
    '判斷若版本2007以上改變存格式
    If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
    Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
    End If
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    Set xlsSalesPoint = Nothing
    Set wksAccrpt407 = Nothing
    MsgBox "Excel檔案產生完成！（檔案位置：" & strFileName & "）"
    Exit Sub

ErrHnd:
    If Not xlsSalesPoint Is Nothing Then
        If Val(xlsSalesPoint.Version) < 12 Then
              xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
        Else
              xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
        End If
        xlsSalesPoint.Workbooks.Close
        xlsSalesPoint.Quit
        Set xlsSalesPoint = Nothing
        Set wksAccrpt407 = Nothing
    End If
   MsgBox Err.Description
End Sub

Private Sub SetField(ByRef Wks As Worksheet, Optional ByVal bolIsLast As Boolean = False)
    Dim stTemp As String
    
    If bolIsLast = False Then
        Wks.Range(Chr(intField) & intRow).Value = ReportTitle(407)
        Wks.Range(Chr(intField) & intRow).Font.Size = 14
        Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strF)) & intRow).HorizontalAlignment = xlCenter
        Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strF)) & intRow).MergeCells = True
        intRow = intRow + 1
        
        For i = LBound(strF) + 1 To UBound(strF)
            stTemp = strF(i)
            If Chr(intField + i) = "B" Then
                stTemp = "公司別："
            'Add by Amy 2020/04/23 公司名稱改抓變數
            ElseIf Chr(intField + i) = "C" And strCmp <> MsgText(601) Then
                stTemp = strCmpN
            End If
            Wks.Range(Chr(intField + i) & intRow).Font.Size = 12
            Wks.Range(Chr(intField + i) & intRow).Value = stTemp
            If Chr(intField + i) = "B" Then
                Wks.Range(Chr(intField + i) & intRow).Font.Bold = True
                Wks.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlRight
            ElseIf strCmp = MsgText(601) Then
                Wks.Range(Chr(intField + i) & intRow).Font.Size = 11
                Wks.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlCenter
            Else
                Wks.Range(Chr(intField + i) & intRow).HorizontalAlignment = xlLeft
            End If
        Next i
        intRow = intRow + 1
        
        'Modify by Amy 2020/04/23 顯示所有公司字型改小(因欄位多)
        If strCmp = MsgText(601) Then
            Wks.Range(Chr(intField + 1) & intRow).Font.Size = 11
        Else
            Wks.Range(Chr(intField + 1) & intRow).Font.Size = 12
        End If
        Wks.Range(Chr(intField + 1) & intRow).Font.Bold = True
        Wks.Range(Chr(intField + 1) & intRow).Value = "年月份："
        Wks.Range(Chr(intField + 1) & intRow).HorizontalAlignment = xlRight
        Wks.Range(Chr(intField + 2) & intRow).Value = MaskEdBox1.Text & "~" & MaskEdBox2.Text
        Wks.Range(Chr(intField + 2) & intRow & ":" & Chr(intField + UBound(strF)) & intRow).MergeCells = True
        Wks.Range(Chr(intField + 2) & intRow).HorizontalAlignment = xlLeft
        intRow = intRow + 1
        
        For i = LBound(strF) To UBound(strF)
            'Modify by Amy 2020/04/23 顯示所有公司字型改小(因欄位多)
            If strCmp = MsgText(601) Then
                Wks.Range(Chr(intField + i) & intRow).Font.Size = 11
            Else
                Wks.Range(Chr(intField + i) & intRow).Font.Size = 12
            End If
            Wks.Range(Chr(intField + i) & intRow).Font.Bold = True
            Wks.Range(Chr(intField + i) & intRow).Value = strF(i)
            Wks.Range(Chr(intField + i) & intRow).ColumnWidth = intWidth(i)
        Next i
        intTitleRow = intRow: intRow = intRow + 1
    Else
        For i = GetValue("科目名稱") + 1 To UBound(strF)
            Wks.Range(Chr(intField + i) & intTitleRow).Value = "金額"
            Wks.Range(Chr(intField + i) & intTitleRow).HorizontalAlignment = xlCenter
        Next i
    End If
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = LBound(strF) To UBound(strF)
       If UCase(strF(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function
'end 2018/12/28

'Mark by Amy 2018/12/28 部分改寫法
Private Sub ProduceData_Old()
'Dim douManaIn, douManaOut, douExManaIn, douExManaOut As Double
'
'On Error GoTo Checking
'   intCounter = 0
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   adoaccrpt407.CursorLocation = adUseClient
'   adoaccrpt407.Open "select * from accrpt407", adoTaie, adOpenDynamic, adLockBatchOptimistic
''------------------------------------------------
'' 營業收入明細
''------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/1/23 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 >= '4' and a0101 < '5' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   '2014/1/23 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      adoaccrpt407.AddNew
'      Accrpt407Save
'      adoaccrpt407.UpdateBatch
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
''------------------------------------------------
'' 營業收入小計
''------------------------------------------------
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(r40704) from accrpt407 where r40701 = '" & strUserNum & "' and r40702 >= '4' and r40702 < '5'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      adoaccrpt407.AddNew
'      adoaccrpt407.Fields("r40701").Value = strUserNum
'      adoaccrpt407.Fields("r40704").Value = ReportSum(4)
'      adoaccrpt407.Fields("r40705").Value = Counter
'      adoaccrpt407.UpdateBatch
'      adoaccrpt407.AddNew
'      adoaccrpt407.Fields("r40701").Value = strUserNum
'      adoaccrpt407.Fields("r40703").Value = ReportSum(1)
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt407.Fields("r40704").Value = 0
'         douManaIn = 0
'      Else
'         adoaccrpt407.Fields("r40704").Value = Val(adoaccsum.Fields(0).Value)
'         douManaIn = Val(adoaccsum.Fields(0).Value)
'      End If
'      adoaccrpt407.Fields("r40705").Value = Counter
'      adoaccrpt407.UpdateBatch
'
'      'Add By Cheng 2002/01/18
'      adoaccrpt407.AddNew
'      adoaccrpt407.Fields("r40701").Value = strUserNum
'      adoaccrpt407.Fields("r40704").Value = ReportSum(8)
'      adoaccrpt407.Fields("r40705").Value = Counter
'      adoaccrpt407.UpdateBatch
'
'   End If
'   adoaccsum.Close
''------------------------------------------------
'' 營業支出明細
''------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/1/23 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 >= '6' and a0101 < '7' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   '2014/1/23 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      adoaccrpt407.AddNew
'      Accrpt407Save
'      adoaccrpt407.UpdateBatch
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
''------------------------------------------------
'' 營業支出小計
''------------------------------------------------
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(r40704) from accrpt407 where r40701 = '" & strUserNum & "' and r40702 >= '6' and r40702 < '7'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      adoaccrpt407.AddNew
'      adoaccrpt407.Fields("r40701").Value = strUserNum
'      adoaccrpt407.Fields("r40704").Value = ReportSum(4)
'      adoaccrpt407.Fields("r40705").Value = Counter
'      adoaccrpt407.UpdateBatch
'      adoaccrpt407.AddNew
'      adoaccrpt407.Fields("r40701").Value = strUserNum
'      adoaccrpt407.Fields("r40703").Value = ReportSum(2)
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt407.Fields("r40704").Value = 0
'         douManaOut = 0
'      Else
'         adoaccrpt407.Fields("r40704").Value = Val(adoaccsum.Fields(0).Value)
'         douManaOut = Val(adoaccsum.Fields(0).Value)
'      End If
'      adoaccrpt407.Fields("r40705").Value = Counter
'      adoaccrpt407.UpdateBatch
'
'      'Add By Cheng 2002/01/18
'      adoaccrpt407.AddNew
'      adoaccrpt407.Fields("r40701").Value = strUserNum
'      adoaccrpt407.Fields("r40704").Value = ReportSum(8)
'      adoaccrpt407.Fields("r40705").Value = Counter
'      adoaccrpt407.UpdateBatch
'
'   End If
'   adoaccsum.Close
''------------------------------------------------
'' 營業損益
''------------------------------------------------
'   adoaccrpt407.AddNew
'   adoaccrpt407.Fields("r40701").Value = strUserNum
'   adoaccrpt407.Fields("r40703").Value = ReportSum(3)
'   adoaccrpt407.Fields("r40705").Value = Counter
'   adoaccrpt407.Fields("r40704").Value = douManaIn - douManaOut
'   adoaccrpt407.UpdateBatch
'
'   'Add By Cheng 2002/01/1
'   adoaccrpt407.AddNew
'   adoaccrpt407.Fields("r40701").Value = strUserNum
'   adoaccrpt407.Fields("r40704").Value = ReportSum(8)
'   adoaccrpt407.Fields("r40705").Value = Counter
'   adoaccrpt407.UpdateBatch
'
''------------------------------------------------
'' 非營業收入
''------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/1/23 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 >= '71' and a0101 < '72' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   '2014/1/23 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      adoaccrpt407.AddNew
'      Accrpt407Save
'      adoaccrpt407.UpdateBatch
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
''------------------------------------------------
'' 非營業收入小計
''------------------------------------------------
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(r40704) from accrpt407 where r40701 = '" & strUserNum & "' and r40702 >= '71' and r40702 < '72'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      adoaccrpt407.AddNew
'      adoaccrpt407.Fields("r40701").Value = strUserNum
'      adoaccrpt407.Fields("r40704").Value = ReportSum(4)
'      adoaccrpt407.Fields("r40705").Value = Counter
'      adoaccrpt407.UpdateBatch
'      adoaccrpt407.AddNew
'      adoaccrpt407.Fields("r40701").Value = strUserNum
'      adoaccrpt407.Fields("r40703").Value = ReportSum(5)
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt407.Fields("r40704").Value = 0
'         douExManaIn = 0
'      Else
'         adoaccrpt407.Fields("r40704").Value = Val(adoaccsum.Fields(0).Value)
'         douExManaIn = Val(adoaccsum.Fields(0).Value)
'      End If
'      adoaccrpt407.Fields("r40705").Value = Counter
'      adoaccrpt407.UpdateBatch
'
'      'Add By Cheng 2002/01/1
'      adoaccrpt407.AddNew
'      adoaccrpt407.Fields("r40701").Value = strUserNum
'      adoaccrpt407.Fields("r40704").Value = ReportSum(8)
'      adoaccrpt407.Fields("r40705").Value = Counter
'      adoaccrpt407.UpdateBatch
'
'   End If
'   adoaccsum.Close
''------------------------------------------------
'' 非營業支出明細
''------------------------------------------------
'   adoacc010.CursorLocation = adUseClient
'   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
'   'adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '8' and a0104 = '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   '2014/1/23 modify by sonia 加入a0109條件
'   'adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '8' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '8' and a0104 = '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 >= '72' and a0101 < '8' and a0104 = '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   '2014/1/23 end
'   'end 2007/12/19
'   Do While adoacc010.EOF = False
'      adoaccrpt407.AddNew
'      Accrpt407Save
'      adoaccrpt407.UpdateBatch
'      adoacc010.MoveNext
'   Loop
'   adoacc010.Close
''------------------------------------------------
'' 非營業支出小計
''------------------------------------------------
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select sum(r40704) from accrpt407 where r40701 = '" & strUserNum & "' and r40702 >= '72' and r40702 < '8'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      adoaccrpt407.AddNew
'      adoaccrpt407.Fields("r40701").Value = strUserNum
'      adoaccrpt407.Fields("r40704").Value = ReportSum(4)
'      adoaccrpt407.Fields("r40705").Value = Counter
'      adoaccrpt407.UpdateBatch
'      adoaccrpt407.AddNew
'      adoaccrpt407.Fields("r40701").Value = strUserNum
'      adoaccrpt407.Fields("r40703").Value = ReportSum(6)
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         adoaccrpt407.Fields("r40704").Value = 0
'         douExManaOut = 0
'      Else
'         adoaccrpt407.Fields("r40704").Value = Val(adoaccsum.Fields(0).Value)
'         douExManaOut = Val(adoaccsum.Fields(0).Value)
'      End If
'      adoaccrpt407.Fields("r40705").Value = Counter
'      adoaccrpt407.UpdateBatch
'
'      'Add By Cheng 2002/01/1
'      adoaccrpt407.AddNew
'      adoaccrpt407.Fields("r40701").Value = strUserNum
'      adoaccrpt407.Fields("r40704").Value = ReportSum(8)
'      adoaccrpt407.Fields("r40705").Value = Counter
'      adoaccrpt407.UpdateBatch
'
'   End If
'   adoaccsum.Close
''------------------------------------------------
'' 稅前損益
''------------------------------------------------
'   adoaccrpt407.AddNew
'   adoaccrpt407.Fields("r40701").Value = strUserNum
'   adoaccrpt407.Fields("r40703").Value = ReportSum(7)
'   adoaccrpt407.Fields("r40704").Value = douManaIn - douManaOut + douExManaIn - douExManaOut
'   adoaccrpt407.Fields("r40705").Value = Counter
'   adoaccrpt407.UpdateBatch
'   adoaccrpt407.AddNew
'   adoaccrpt407.Fields("r40701").Value = strUserNum
'   adoaccrpt407.Fields("r40704").Value = ReportSum(8)
'   adoaccrpt407.Fields("r40705").Value = Counter
'   adoaccrpt407.UpdateBatch
'   adoaccrpt407.Close
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt407Delete()
   adoTaie.Execute "delete from accrpt407 Where R40701='" & strUserNum & "'" 'Moidfy by Amy 2018/12/28 避免同時多人操作資料錯誤
End Sub

'*************************************************
'  將會計科目餘額儲存至損益表資料暫存檔中
'
'*************************************************
'Add by Amy 2018/12/28 畫面公司別加3.4選項
Private Sub Accrpt407Save1(ByVal strA0101 As String, ByVal strA0102 As String)
    Dim strSql As String, strField As String
    
    strField = ",a0403"
    'Modify by Amy 2020/04/23 公司別改抓變數
'    If Text6 = "1" Then strSql = " And a0403='1'"
'    If Text6 = "2" Then strSql = " And a0403='J'"
'    If Text6 = "3" Then strField = ""
    If strCmp <> MsgText(601) Then
        If InStr(strCmp, "+") > 0 Then
            strSql = " And a0403 In ('" & Replace(strCmp, "+", "','") & "')"
            strField = ""
        Else
            strSql = " And a0403='" & strCmp & "'"
        End If
    End If
    'end 2020/04/23
    If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
        strSql = strSql & " and a0401 || decode(length(a0402),1, '0'||a0402, a0402) >= " & Val(Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2)) & ""
    End If
    If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
        strSql = strSql & " and a0401 || decode(length(a0402),1, '0'||a0402, a0402) <= " & Val(Mid(MaskEdBox2.Text, 1, 3) & Mid(MaskEdBox2.Text, 5, 2)) & ""
    End If
    
    '公司別可能不止一家,判斷餘額沒有數字不出現
    strSql = "select sum(a0408)" & strField & " from acc040 where substr(a0405, 1, 4) = '" & adoacc010.Fields("a0101").Value & "' and a0404 = '" & MsgText(55) & "'" & strSql & IIf(strField <> "", " Group by a0403", "") & " Having sum(a0408)<>0"
    adoacc040.CursorLocation = adUseClient
    adoacc040.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
    If adoacc040.RecordCount > 0 Then
        adoacc040.MoveFirst
        Do While Not adoacc040.EOF
            adoaccrpt407.AddNew
            adoaccrpt407.Fields("r40701").Value = strUserNum
            adoaccrpt407.Fields("r40702").Value = strA0101 '會計科目
            adoaccrpt407.Fields("r40703").Value = strA0102 '科目名稱
            adoaccrpt407.Fields("r40704").Value = "" & adoacc040.Fields(0).Value '餘額
            If strField <> MsgText(601) Then
                adoaccrpt407.Fields("r40705").Value = "" & adoacc040.Fields("a0403") '公司別
            End If
            adoaccrpt407.UpdateBatch
            adoacc040.MoveNext
        Loop
    End If
    
    adoacc040.Close
End Sub

'Mark by Amy 2018/12/28 畫面公司別加3.4選項,可能不止一個公司別
Private Sub Accrpt407Save1_Old()
'Dim strSql As String
'
'   adoacc040.CursorLocation = adUseClient
'   If Text6 <> MsgText(601) Then
'      '2014/1/23 modify by sonia
'      'strSql = " and a0403 = '" & Text6 & "'"
'      strSql = " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
'      '2014/1/23 end
'   End If
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = strSql & " and a0401 || decode(length(a0402),1, '0'||a0402, a0402) >= " & Val(Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a0401 || decode(length(a0402),1, '0'||a0402, a0402) <= " & Val(Mid(MaskEdBox2.Text, 1, 3) & Mid(MaskEdBox2.Text, 5, 2)) & ""
'   End If
'   adoacc040.Open "select sum(a0408) from acc040 where substr(a0405, 1, 4) = '" & adoacc010.Fields("a0101").Value & "' and a0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc040.RecordCount <> 0 Then
'      If IsNull(adoacc040.Fields(0).Value) Then
'          adoaccrpt407.Fields("r40704").Value = 0
'      Else
'         adoaccrpt407.Fields("r40704").Value = Val(adoacc040.Fields(0).Value)
'      End If
'   Else
'     adoaccrpt407.Fields("r40704").Value = 0
'   End If
'   adoaccrpt407.Fields("r40705").Value = Counter
'   adoacc040.Close
End Sub

'*************************************************
' 設定初始及結束年月
'
'*************************************************
Private Sub Accrpt407Save()
   'Mark by Amy 2018/12/28 不使用
'   adoaccrpt407.Fields("r40701").Value = strUserNum
'   adoaccrpt407.Fields("r40702").Value = adoacc010.Fields("a0101").Value
'   If IsNull(adoacc010.Fields("a0102").Value) Then
'      adoaccrpt407.Fields("r40703").Value = Null
'   Else
'      adoaccrpt407.Fields("r40703").Value = adoacc010.Fields("a0102").Value
'   End If
'   Accrpt407Save1_Old
End Sub

'*************************************************
' 序號計算
'
'*************************************************
Private Function Counter() As Integer
   intCounter = intCounter + 1
   Counter = intCounter
End Function

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   'Modify by Amy 2020/04/23 公司別改下拉
'   Text6 = ""
'   Text7 = "台一　專利商標/智權"
   CboCmp = ""
   'end 2020/04/23
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = Mid(DFormat, 1, 6)
   CboCmp.SetFocus 'Modify by Amy 2020/04/23 原:Text6
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
'Modify by Amy 2020/04/23 +bolShowMsg
Public Function FormCheck(bolShowMsg As Boolean) As Boolean
   Dim bCancel As Boolean 'Add by Amy 2020/04/23
   
   'Add by Amy 2018/12/28 公司別必填
   'Modify by Amy 2020/04/23 公司別改下拉後,可能有空白=抓全部公司 原:Text6
   If Trim(CboCmp) <> MsgText(601) Then
      Call CboCmp_Validate(bCancel)
      If bCancel = True Then
            FormCheck = False
            bolShowMsg = True
            CboCmp.SetFocus
            Exit Function
      End If
   End If
   'end 2020/04/04/23
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

'Add By Sindy 2013/5/27
'產生Excel檔案
Public Sub IsExcelSave_Old()
'Mark by Amy 2018/12/28 畫面公司別加3.4選項,格式修改
'Dim xlsSalesPoint As New Excel.Application
'Dim wksaccrpt407 As New Worksheet
'Dim strFileName As String
'Dim iRow As Integer
'Dim stCellFormat As String
'Dim stRptName As String
'Dim Rc As String '欄位座標
'Dim MaxCol As String '最右的欄位代碼
'
'On Error GoTo ErrHnd
'
'   '讀取綜合損益表資料
'   If adoaccrpt407.State = adStateOpen Then
'      adoaccrpt407.Close
'   End If
'   adoaccrpt407.CursorLocation = adUseClient
'   adoaccrpt407.Open "select * from accrpt407 order by r40705 asc", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccrpt407.RecordCount <= 0 Then
'      adoaccrpt407.Close
'      Exit Sub
'   End If
'
'   MaxCol = Chr(Asc("a") + 5)
'
'   stCellFormat = "#,##0.00 ;[紅色]-#,##0.00 "
'
'   stRptName = ReportTitle(407)
'
'   strFileName = strExcelPath & Trim(Replace(stRptName, "*", "")) & Format(Now, "yyyymmddhhmmss") & MsgText(43)
'   If Dir(strFileName) = MsgText(601) Then
'      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'         MkDir strExcelPath
'      End If
'   Else
'      Kill strFileName
'   End If
'
'   xlsSalesPoint.Workbooks.add
'   Set wksaccrpt407 = xlsSalesPoint.Worksheets(1)
'   With wksaccrpt407
'      iRow = 1
'      .Range("a" & iRow).Value = stRptName
'      Rc = MaxCol & iRow
'      With .Range("a" & iRow & ":" & Rc)
'         .Font.Size = 18
'         .Font.Bold = True
'         .HorizontalAlignment = xlCenter
'         .MergeCells = True
'      End With
'
'      iRow = iRow + 2
'      .Range("c" & iRow).Value = "公司別："
'      With .Range("c" & iRow)
'         .Font.Size = 12
'         .Font.Bold = True
'         .HorizontalAlignment = xlRight
'      End With
'
'      'Modify By Sindy 2014/9/3
'      '.Range("d" & iRow).Value = Text6 & "  " & Text7
'      .Range("d" & iRow).Value = IIf(Text6 = "2", "J", Text6) & "  " & IIf(Text7 = "", "台一　專利商標/智權", Text7)
'      '2014/9/3 END
'      .Columns("d").ColumnWidth = 24
'      With .Range("d" & iRow)
'         .Font.Size = 12
'         .Font.Bold = False
'         .HorizontalAlignment = xlLeft
'         .MergeCells = True
'      End With
'
'      iRow = iRow + 1
'      .Range("c" & iRow).Value = "年月份："
'      With .Range("c" & iRow)
'         .Font.Size = 12
'         .Font.Bold = True
'         .HorizontalAlignment = xlRight
'      End With
'
'      .Range("d" & iRow).Value = MaskEdBox1 & " ~ " & MaskEdBox2
'      With .Range("d" & iRow)
'         .Font.Size = 12
'         .Font.Bold = False
'         .HorizontalAlignment = xlLeft
'         .MergeCells = True
'      End With
'
'      iRow = iRow + 1
'
'      .Range("a" & iRow).Value = "列印人員："
'      With .Range("a" & iRow)
'         .Font.Size = 12
'         .Font.Bold = True
'         .HorizontalAlignment = xlRight
'      End With
'
'      .Range("b" & iRow).Value = StaffQuery(strUserNum)
'      With .Range("b" & iRow)
'         .Font.Size = 12
'         .Font.Bold = False
'         .HorizontalAlignment = xlLeft
'      End With
'
'      .Range("e" & iRow).Value = "列印日期："
'      With .Range("e" & iRow)
'         .Font.Size = 12
'         .Font.Bold = True
'         .HorizontalAlignment = xlRight
'         .MergeCells = True
'      End With
'
'      .Range("f" & iRow).Value = CFDate(ACDate(ServerDate))
'      With .Range("f" & iRow)
'         .Font.Size = 12
'         .Font.Bold = False
'         .HorizontalAlignment = xlLeft
'      End With
'
'      iRow = iRow + 2
'
'      .Range("a" & iRow).Value = "會計科目"
'      .Columns("a").ColumnWidth = 12
'      With .Range("a" & iRow)
'         .Font.Size = 12
'         .Font.Bold = True
'         .HorizontalAlignment = xlLeft
'      End With
'
'      .Range("c" & iRow).Value = "科目名稱"
'      .Columns("c").ColumnWidth = 14
'      With .Range("c" & iRow)
'         .Font.Size = 12
'         .Font.Bold = True
'         .HorizontalAlignment = xlLeft
'      End With
'
'      .Columns("e").ColumnWidth = 5
'      .Range("f" & iRow).Value = "金額"
'      .Columns("f").ColumnWidth = 14
'      With .Range("f" & iRow)
'         .Font.Size = 12
'         .Font.Bold = True
'         .HorizontalAlignment = xlCenter
'      End With
'
'      adoaccrpt407.MoveFirst
'      Do While Not adoaccrpt407.EOF
'         iRow = iRow + 1
'
'         With .Range("A" & iRow)
'            .Font.Size = 12
'            .HorizontalAlignment = xlLeft
'         End With
'         With .Range("F" & iRow)
'            .Font.Size = 12
'            .HorizontalAlignment = xlRight
'            .NumberFormatLocal = stCellFormat
'         End With
'
'         .Range("a" & iRow).Value = "" & adoaccrpt407.Fields("r40702")
'         .Range("c" & iRow).Value = "" & adoaccrpt407.Fields("r40703")
'         .Range("f" & iRow).Value = "" & adoaccrpt407.Fields("r40704")
'         adoaccrpt407.MoveNext
'      Loop
'      adoaccrpt407.Close
'      iRow = iRow + 1
'      .Range("c" & iRow).Value = "*** 結束 ***"
'
'      'Modify by Amy 2015/05/21 原使用函數以為是抓A4紙張
'      .PageSetup.PaperSize = 9 '設定紙張 A4
'      .PageSetup.Orientation = xlPortrait '直印 xlLandscape.橫印
'      .PageSetup.PrintTitleRows = "$1:$6" '表頭保留7列
'      .PageSetup.PrintArea = "$A$1:$F$" & iRow '設定列印範圍
''      .PageSetup.LeftMargin = xlsSalesPoint.InchesToPoints(0.5) '左邊界
''      .PageSetup.RightMargin = xlsSalesPoint.InchesToPoints(0.5) '右邊界
'
'
'      .PageSetup.LeftMargin = xlsSalesPoint.InchesToPoints(0.5)
'      .PageSetup.RightMargin = xlsSalesPoint.InchesToPoints(0.5)
'      .PageSetup.TopMargin = xlsSalesPoint.InchesToPoints(0.3)
'      .PageSetup.BottomMargin = xlsSalesPoint.InchesToPoints(0.3)
'      .PageSetup.HeaderMargin = xlsSalesPoint.InchesToPoints(0.5)
'      .PageSetup.FooterMargin = xlsSalesPoint.InchesToPoints(0.5)
'
'      .PageSetup.Zoom = 100 '縮放比例
'   End With
'   'Modify by Amy2016/05/06 判斷若版本2007以上改變存格式
'   If Val(xlsSalesPoint.Version) < 12 Then
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
'  Else
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
'   End If
'   'end 2016/05/06
'   xlsSalesPoint.Workbooks.Close
'   xlsSalesPoint.Quit
'   Set xlsSalesPoint = Nothing
'   Set wksaccrpt407 = Nothing
'   MsgBox "Excel檔案產生完成！（檔案位置：" & strFileName & "）"
'   Exit Sub
'
'ErrHnd:
'   If Not xlsSalesPoint Is Nothing Then
'      xlsSalesPoint.Quit
'      Set xlsSalesPoint = Nothing
'      Set wksaccrpt407 = Nothing
'   End If
'   MsgBox Err.Description
End Sub

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

