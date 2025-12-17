VERSION 5.00
Begin VB.Form Frmacc4480 
   AutoRedraw      =   -1  'True
   Caption         =   "資產負債表"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2295
   ScaleWidth      =   5160
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
      Top             =   300
      Width           =   3350
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
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1995
   End
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
      Left            =   1380
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   2310
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
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
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
      TabIndex        =   5
      Top             =   1770
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
      Height          =   315
      Left            =   4050
      MaxLength       =   2
      TabIndex        =   2
      Top             =   840
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
      Height          =   315
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   1
      Top             =   840
      Width           =   612
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
      TabIndex        =   11
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "是否含子目(Y:包含)"
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
      TabIndex        =   10
      Top             =   1320
      Width           =   2200
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "截止月份"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   840
      Width           =   1095
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
      Height          =   255
      Left            =   360
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   300
      Width           =   675
   End
End
Attribute VB_Name = "Frmacc4480"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

'Modify by Amy 2020/05/22 連線原Public
Dim adoacc010 As New ADODB.Recordset
Dim adoacc021 As New ADODB.Recordset
Dim adoacc040 As New ADODB.Recordset
Dim adoaccrpt409 As New ADODB.Recordset
Dim adoacc0b0 As New ADODB.Recordset
Dim adoaccsum As New ADODB.Recordset
'end 2020/05/22
Dim intAutoNo As Integer
'Dim dllaccrpt409 As Object 'Mark by Amy 2020/05/29 不使用
Dim strSql As String
Dim strPrinter As String
Dim stra0403 As String  '2014/1/23 add by sonia
Dim strCmp As String, strCmpN As String 'Add by Amy 2020/04/17
'Add by Amy 2020/05/29
Dim stRptName As String, iRow As Integer, intField As Integer, intTitleRow As Integer, intCount(3) As Integer
Dim intSum(2 To 3) As String, stCellFormat As String '負債、業主權益 合計位置/格式
Dim strFieldN(), intWidth()

'Add by Amy 2020/04/17
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
'end 200/04/16

Private Sub Command1_Click()
   Dim bolShowMsg As Boolean 'Add by Amy 2020/04/17
   Dim strQ As String 'Add by Amy 2020/05/29
'on error GoTo Checking
   'Modify by Amy 2020/04/17 +bolShowMsg,公司改下拉
   strCmp = "": strCmpN = ""
   If Trim(CboCmp) <> MsgText(601) Then
      strCmp = CboCmp
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
   End If
   
   If FormCheck(bolShowMsg) = False Then
      If bolShowMsg = False Then MsgBox MsgText(181), , MsgText(5)
    'end 2020/04/17
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   'Modify by Amy 2020/04/17
   bolShowMsg = False
   strCmpN = GetAccReportCmpN(CboCmp, , True)
   Accrpt409Delete
   'Modfiy by Amy 2020/05/29 改寫法
'   Call ProduceData(bolShowMsg)
'   If bolShowMsg = True Then Screen.MousePointer = vbDefault: Exit Sub
'   '無資料產生跳離開,否則程式會錯
'    If CheckData = False Then Screen.MousePointer = vbDefault: Exit Sub
   intCount(0) = 0: intCount(1) = 0: intCount(2) = 0: intCount(3) = 0
   If ProduceData = False Then Exit Sub
   'end 2020/05/29
   'end 2020/04/17
   
   'Modify by Amy 2020/05/29 拿掉列印
   'Add By Sindy 2013/5/27
'   If Check1.Value = 1 Then
   strQ = "select * from accrpt409 where r40901='" & strUserNum & "' "
   If adoaccrpt409.State = adStateOpen Then adoaccrpt409.Close
   adoaccrpt409.CursorLocation = adUseClient
   adoaccrpt409.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt409.RecordCount <= 0 Then
      MsgBox "無資料產生！"
      Screen.MousePointer = vbDefault
      Exit Sub
   Else
      Call IsExcelSave
   End If
'   Else
'   '2013/5/27 End
'      PUB_SetOsDefaultPrinter Combo1
'      If adoaccrpt409.State = adStateOpen Then
'         adoaccrpt409.Close
'      End If
'      adoaccrpt409.CursorLocation = adUseClient
'      'Modify by Amy 2015/04/09 財務處可能同時兩個人執行此報表,造成資料錯誤
'      adoaccrpt409.Open "select * From accrpt409 where r40901='" & strUserNum & "' ", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccrpt409.RecordCount <> 0 Then
'         '2014/1/23 modify by sonia
'         'dllaccrpt409.Acc4480 ReportTitle(409), Text6, Text7, Text3, Text1, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'         'Modify by Amy 2020/04/17 公司別/名稱改抓變數
'         'dllaccrpt409.Acc4480 ReportTitle(409), IIf(Text6 = "2", "J", Text6), IIf(Text7 = "", "台一　專利商標/智權", Text7), Text3, Text1, strUserNum & "-" & StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'         dllaccrpt409.Acc4480 ReportTitle(409), strCmp, strCmpN, Text3, Text1, strUserNum & "-" & StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'      End If
'      'end 2015/04/09
'      adoaccrpt409.Close
'      Set adoaccrpt409 = Nothing
'      PUB_SetOsDefaultPrinter strPrinter
'   End If
   'end 2020/05/29
   Screen.MousePointer = vbDefault
   FormClear
   'Modify by Amy 2020/05/29 原:MsgText(102) & " / " & MsgText(139)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(139)
   Exit Sub
   
Checking:
   adoaccrpt409.Close
   Set adoaccrpt409 = Nothing
   'PUB_SetOsDefaultPrinter strPrinter 'Mark by Amy 2020/05/29
   Screen.MousePointer = vbDefault
   MsgBox Err.Description
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      'Modify by Amy 2020/05/29 原:MsgText(102) & " / " & MsgText(139)
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(139)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5280 '5250
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
   'Add by Amy 2020/04/17 公司別下拉
   CboCmp.Clear
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/04/17
   'Modify by Amy 2020/05/29 原:MsgText(102) & " / " & MsgText(139),並拿掉列印
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(139)
   'Set dllaccrpt409 = CreateObject("AccReport.ReportSelect")
   'Add By Cheng 2003/02/14
   '預設不含子目
   'modify by sonia 2016/2/15 改預設含子目
   Me.Text2.Text = "Y"
   Check1.Value = 1 'Add by Amy 2016/09/29 預勾產生Excel-瑞婷
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

   'Set dllaccrpt409 = Nothing'Mark by Amy 2020/05/29
   Set Frmacc4480 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add by Amy 2020/05/29
   If KeyAscii <> Asc("Y") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

'Add by Amy 2020/05/29 改寫法,若負債業主權益列數>資料列數,資料會不完整
Private Function ProduceData() As Boolean
    Dim strQ As String, strA As String
    Dim strA1 As String, strA2 As String, strA0109 As String

On Error GoTo ErrHnd
    ProduceData = False
    strQ = "Select * From accrpt409 Where r40901='" & strUserNum & "' "
    If adoaccrpt409.State = adStateOpen Then adoaccrpt409.Close
    adoaccrpt409.CursorLocation = adUseClient
    adoaccrpt409.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
    
    '公司別
    If strCmp <> MsgText(601) Then
        If InStr(strCmp, "+") > 0 Then
            strA1 = strA1 & " And a0403 In ('" & Replace(strCmp, "+", "','") & "')"
            strA2 = strA2 & " And (a0109 is null Or A0109 In ('" & Replace(strCmp, "+", "','") & "') " & ")"
        Else
            strA1 = strA1 & " And a0403 = '" & strCmp & "' "
            strA2 = strA2 & " And (a0109 is null Or A0109 ='" & strCmp & "') "
        End If
    End If
    '年度
    If Text3 <> MsgText(601) Then
        strA1 = strA1 & " And a0401 = " & Val(Text3) & ""
    End If
    '月份
    If Text1 <> MsgText(601) Then
        strA1 = strA1 & " And a0402 = " & Val(Text1) & ""
    End If
    '不含子科目
    If Trim(Text2) <> MsgText(602) Then
        strA2 = strA2 & " And A0104 <'4' "
    End If
    
    '抓取會計科目餘額資料
    strQ = "Insert into accrpt409_check " & _
                "Select '" & strUserNum & "', SUBSTR(A0405, 1, 2) r409c2 From acc040 " & _
                "Where  Length(a0405) > '2' " & strA1 & _
                "Group by SUBSTR(A0405, 1, 2) Having Sum(Nvl(A0408,0)) <> 0 "
    adoTaie.Execute strQ
    
    '------------------------------------------------
    '  資產
    '------------------------------------------------
    '抓取「資產」會計科目資料及餘額檔,寫入暫存檔
    strQ = "Select * From acc010 Where  a0101 < '2' " & strA2 & _
            " And SUBSTR(A0101, 1, 2) in (Select r409c2 From accrpt409_check Where R409C1='" & strUserNum & "' ) " & _
            "Order by a0101 asc"
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        strA = "Select * From acc040, acc010 Where a0405 = a0101" & strA1 & _
                    " And a0405 = '" & adoacc010.Fields("a0101").Value & "' "
        If adoacc040.State <> adStateClosed Then adoacc040.Close
        adoacc040.Open strA, adoTaie, adOpenStatic, adLockReadOnly
        If adoacc040.RecordCount <> 0 Then
            If Not adoacc040.EOF Then
                adoacc040.MoveFirst
                Call Accrpt409Save2(strA1)
            End If
        End If
        adoacc010.MoveNext
    Loop
    intCount(1) = intCount(0): intCount(0) = 0
    
    '------------------------------------------------
    ' 負債
    '------------------------------------------------
    '抓取「負債」會計科目資料及餘額檔,寫入暫存檔
    strQ = "Select * From acc010 Where a0101 >= '2' And a0101 < '3' " & strA2 & _
            " And SUBSTR(A0101, 1, 2) in (Select r409c2 From accrpt409_check Where R409C1='" & strUserNum & "' ) " & _
            "Order by a0101 asc"
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        strA = "Select * From acc040, acc010 Where a0405 = a0101" & strA1 & _
                    " And a0405 = '" & adoacc010.Fields("a0101").Value & "' "
        If adoacc040.State <> adStateClosed Then adoacc040.Close
        adoacc040.Open strA, adoTaie, adOpenStatic, adLockReadOnly
        If adoacc040.RecordCount <> 0 Then
            If Not adoacc040.EOF Then
                adoacc040.MoveFirst
                Call Accrpt409Save2(strA1)
            End If
        End If
        adoacc010.MoveNext
    Loop
    intCount(2) = intCount(0): intCount(0) = 0
    
    If intCount(1) = 0 And intCount(2) = 0 Then ProduceData = True: Exit Function
    
    ' 負債小計
    If intCount(2) > 0 Then
        strQ = "Insert Into Accrpt409 (r40901,r40902,r40903) " & _
                    "Select '" & strUserNum & "','2z','" & ReportSum(10001) & "' From Dual "
        adoTaie.Execute strQ
    End If
    
    '------------------------------------------------
    ' 股東權益
    '------------------------------------------------
     '抓取「股東權益」會計科目資料及餘額檔,寫入暫存檔
     strQ = "Select * From acc010 Where a0101 >= '3' And a0101 < '4' " & strA2 & _
            "Order by a0101 asc"
    If adoacc010.State <> adStateClosed Then adoacc010.Close
    adoacc010.CursorLocation = adUseClient
    adoacc010.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    Do While adoacc010.EOF = False
        If adoacc010.Fields(0).Value = "3222" Then
            '4字頭+6字頭+71字頭-72字頭
            strA = "Select Nvl(sum(decode(substr(a0405, 1, 1), '4', a0408, '6', a0408 * -1, DECODE(SUBSTR(A0405,1,2),'71',a0408,A0408*-1))),0) " & _
                    "From acc040, acc010 Where a0405 = a0101 (+) And a0405 >= '4' And a0405 < '8' And a0404 = '" & MsgText(55) & "' " & strA1
        Else
            strA = "Select * From acc040, acc010 Where a0405 = a0101 And a0405 = '" & adoacc010.Fields("a0101") & "' " & strA1
        End If
        If adoacc040.State <> adStateClosed Then adoacc040.Close
        adoacc040.Open strA, adoTaie, adOpenStatic, adLockReadOnly
        If adoacc040.RecordCount <> 0 Then
            If Not adoacc040.EOF Then
                adoacc040.MoveFirst
                If adoacc010.Fields(0).Value = "3222" Then
                    Call Accrpt409Save4
                Else
                    Call Accrpt409Save2(strA1)
                End If
            End If
        End If
        adoacc010.MoveNext
    Loop
    intCount(3) = intCount(0): intCount(0) = 0
    
    If intCount(3) > 0 Then
    ' 股東權益小計
        strQ = "Insert Into Accrpt409 (r40901,r40902,r40903) " & _
                    "Select '" & strUserNum & "','3z','" & ReportSum(12001) & "' From Dual "
        adoTaie.Execute strQ
    End If
    ProduceData = True
    Exit Function
    
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Mark by Amy 2020/04/17 公司別改下拉
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

''2014/1/23 add by sonia
'Private Sub Text6_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub
''2014/1/23 end
'end 2020/04/17

'*************************************************
'  產生報表資料
'
'*************************************************
'Modify by Amy 2020/04/17 +bolShowMsg
Private Sub ProduceData_Old(bolShowMsg As Boolean)
'on error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   If adoaccrpt409.State = adStateOpen Then
      adoaccrpt409.Close
   End If
   adoaccrpt409.CursorLocation = adUseClient
   'Modify by Amy 2015/04/09 財務處可能同時兩個人執行此報表,造成資料錯誤
   adoaccrpt409.Open "select * from accrpt409 Where r40901='" & strUserNum & "' ", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoaccrpt409.AddNew
   adoaccrpt409.Fields("r40901").Value = strUserNum
   adoaccrpt409.UpdateBatch
   Call Accrpt409Save1(bolShowMsg) 'Modify by Amy 2020/04/17
   adoaccrpt409.Close
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
Private Sub Accrpt409Delete()
   'Modify by Amy 2015/04/09 財務處可能同時兩個執行此報表,造成資料錯誤
   adoTaie.Execute "delete from accrpt409 Where r40901='" & strUserNum & "' "
   'Modified by Lydia 2014/12/17
   adoTaie.Execute "delete from accrpt409_check Where R409C1='" & strUserNum & "' "
   'end 2015/04/09
End Sub
'*************************************************
'  計算會計科目餘額並儲存於資產負債表暫存檔中(1090529不使用)
'
'*************************************************
'Modify by Amy 2020/04/17 +bolShowMsg
Private Sub Accrpt409Save1(bolShowMsg As Boolean)
Dim douTotal2 As Double
Dim douTotal3 As Double
Dim strTotal3 As String
'Modified by Lydia 2014/12/17
Dim strA1 As String
Dim strA2 As String, strQ As String 'Add by Amy 2020/04/17

   intAutoNo = 0
'------------------------------------------------
'  資產明細
'------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   '92.12.31 MODIFY BY SONIA
   'adoacc010.Open "select * from acc010 where a0101 < '2' AND A0101<>'1134' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '93.10.5 MODIFY BY SONIA
   'adoacc010.Open "select * from acc010 where a0101 < '2' AND A0101<>'1134' AND A0101<>'1305' AND A0101<>'1307' AND A0101<>'1308' AND A0101<>'1309' AND A0101<>'1310' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 < '2' AND A0101<>'1134' AND A0101<>'1305' AND A0101<>'1307' AND A0101<>'1308' AND A0101<>'1309' AND A0101<>'1310' AND A0101<>'110211' AND A0101<>'110212' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/1/23 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 < '2' AND A0101<>'1134' AND A0101<>'1305' AND A0101<>'1307' AND A0101<>'1308' AND A0101<>'1309' AND A0101<>'1310' AND A0101<>'110211' AND A0101<>'110212' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
  'Modified by Lydia 2014/12/17 判斷分類會計科目合計金額為零的會計科目不顯示
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 < '2' AND A0101<>'1134' AND A0101<>'1305' AND A0101<>'1307' AND A0101<>'1308' AND A0101<>'1309' AND A0101<>'1310' AND A0101<>'110211' AND A0101<>'110212' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc010.Open "select * from acc010 where a0101 < '2' AND A0101<>'1134' AND A0101<>'1305' AND A0101<>'1307' AND A0101<>'1308' AND A0101<>'1309' AND A0101<>'1310' AND A0101<>'110211' AND A0101<>'110212' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
    'Modified by Lydia 2022/05/25  debug: 會計科目不等於0才要出現
    'strA1 = "insert into accrpt409_check select '" & strUserNum & "', SUBSTR(A0405, 1, 2) r409c2 from acc040 where "
    strA1 = "insert into accrpt409_check select '" & strUserNum & "', SUBSTR(A0405, 1, 2) r409c2 from (select * acc040 where "
    'Modify by Amy 2020/04/17 公司別改下拉 原:Text6/IIf(Text6 = "2", "J", Text6),並加組合公司
    If strCmp <> "" Then
        If InStr(strCmp, "+") > 0 Then
            strA1 = strA1 & " a0403 In ('" & Replace(strCmp, "+", "','") & "') and"
            strA2 = " A0109 In ('" & Replace(strCmp, "+", "','") & "') "
        Else
            strA1 = strA1 & " a0403 = '" & strCmp & "' and"
            strA2 = " A0109 ='" & strCmp & "' "
        End If
    End If
    'end 2020/04/17
    'Modified by Lydia 2022/05/25  debug: 會計科目不等於0才要出現
    'strA1 = strA1 & " A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " and length(a0405) > 2 group by SUBSTR(A0405, 1, 2) having SUM(NVL(A0408,0)) <> 0 "
    strA1 = strA1 & " A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " and length(a0405) > 2 and nvl(a0408,0) <> 0) group by SUBSTR(A0405, 1, 2) "
    adoTaie.Execute strA1
    'Modify by Amy 2015/04/09  財務處可能同時兩個人執行此報表,造成資料錯誤+員編條件
   'modify by sonia 2018/8/22 取消AND A0101<>'1134' AND A0101<>'1305' AND A0101<>'1307' AND A0101<>'1308' AND A0101<>'1309' AND A0101<>'1310' AND A0101<>'110211' AND A0101<>'110212'條件,一律用instr(a0102,'不用')=0
   'modify by sonia 2018/12/28 取消instr(a0102,'不用')=0條件,1公司106/11之1132備抵呆帳－應收票據/不用 有數字但不出現而不平
   'Modify by Amy 2020/04/17 公司別改下拉 原:Text6
   If strCmp <> "" Then
      strA1 = "select * from acc010 where a0101 < '2' and (a0109 is null or " & strA2 & ") " & _
      "and SUBSTR(A0101, 1, 2) in (select r409c2 from accrpt409_check Where R409C1='" & strUserNum & "' ) order by a0101 asc"
   Else
      strA1 = "select * from acc010 where a0101 < '2' " & _
      "and SUBSTR(A0101, 1, 2) in (select r409c2 from accrpt409_check Where R409C1='" & strUserNum & "' ) order by a0101 asc"
   End If
   'end 2015/04/09
   adoacc010.Open strA1, adoTaie, adOpenStatic, adLockReadOnly
   'end 2014/12/17
   '2014/1/23 end
   'end 2007/12/19
   '93.10.5 END
   '92.12.31 END
   Do While adoacc010.EOF = False
      If Text2 <> MsgText(602) Then
         If adoacc010.Fields("a0104").Value = "4" Then
            GoTo Check1
         End If
      End If
      adoacc040.CursorLocation = adUseClient
      strSql = MsgText(601)
      stra0403 = MsgText(601) '2014/1/23 add by sonia
      'Modify by Amy 2020/04/17 公司別改下拉 原:Text6
      If strCmp <> MsgText(601) Then
         If InStr(strCmp, "+") > 0 Then
            strSql = " and a0403 In ('" & Replace(strCmp, "+", "','") & "')"
            stra0403 = " a0403 In ('" & Replace(strCmp, "+", "','") & "') and"
         Else
            '2014/1/23 modify by sonia
            'strSql = " and a0403 = '" & Text6 & "'"
            strSql = " and a0403 = '" & strCmp & "'"
            stra0403 = " a0403 = '" & strCmp & "' and"
            '2014/1/23 end
        End If
      End If
      If Text3 <> MsgText(601) Then
         strSql = strSql & " and a0401 = " & Val(Text3) & ""
      End If
      If Text1 <> MsgText(601) Then
         strSql = strSql & " and a0402 = " & Val(Text1) & ""
      End If
      strSql = strSql & " and a0405 = '" & adoacc010.Fields("a0101").Value & "'"
      adoacc040.Open "select * from acc040, acc010 where acc040.a0405 = acc010.a0101" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc040.RecordCount <> 0 Then
         Accrpt409Save2
      Else
         adoaccrpt409.AddNew
         adoaccrpt409.Fields("r40901").Value = strUserNum
         adoaccrpt409.Fields("r40902").Value = adoacc010.Fields("a0101").Value
         If IsNull(adoacc010.Fields("a0102").Value) Then
            adoaccrpt409.Fields("r40903").Value = Null
         Else
            adoaccrpt409.Fields("r40903").Value = adoacc010.Fields("a0102").Value
         End If
         Select Case adoacc010.Fields("A0104").Value
            Case "3"
               adoaccsum.CursorLocation = adUseClient
               '2014/1/23 modify by sonia
               'adoaccsum.Open "select SUM(A0408) from acc040 where A0403 = '" & Text6 & "' AND A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND SUBSTR(A0405, 1, 4) = '" & adoacc010.Fields("A0101").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
               adoaccsum.Open "select SUM(A0408) from acc040 where " & stra0403 & " A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND SUBSTR(A0405, 1, 4) = '" & adoacc010.Fields("A0101").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoaccsum.RecordCount <> 0 Then
                  If IsNull(adoaccsum.Fields(0).Value) Then
                     adoaccrpt409.Fields("r40904").Value = 0
                  Else
                     adoaccrpt409.Fields("r40904").Value = Format(adoaccsum.Fields(0).Value, FAmount)
                  End If
               Else
                  adoaccrpt409.Fields("R40904").Value = 0
               End If
               adoaccsum.Close
            Case "4"
               adoaccrpt409.Fields("R40909").Value = 0
            Case Else
               adoaccrpt409.Fields("r40904").Value = 0
         End Select
        'Modified by Lydia 2014/12/17 判斷分類會計科目合計金額為零的會計科目不顯示
        If adoaccrpt409.Fields("r40904").Value = 0 Or adoaccrpt409.Fields("r40909").Value = 0 Then
           adoaccrpt409.CancelUpdate
        Else
           AutoNoSave
           adoaccrpt409.UpdateBatch
        End If
      End If
      adoacc040.Close
Check1:
      adoacc010.MoveNext
   Loop
   adoacc010.Close
'------------------------------------------------
' 資產總額
'------------------------------------------------
   adoaccrpt409.AddNew
   adoaccrpt409.Fields("r40901").Value = strUserNum
   adoaccrpt409.Fields("r40904").Value = ReportSum(4)
   AutoNoSave
   adoaccrpt409.UpdateBatch
   adoaccrpt409.AddNew
   adoaccrpt409.Fields("r40901").Value = strUserNum
   'Modify By Cheng 2002/01/18
'   adoaccrpt409.Fields("r40903").Value = ReportSum(9)
   adoaccrpt409.Fields("r40903").Value = ReportSum(9001)
   adoaccsum.CursorLocation = adUseClient
   'Modfiy by Amy 2015/04/09 財務處可能同時兩個人執行此報表,造成資料錯誤
   adoaccsum.Open "select sum(r40904) from accrpt409 where r40901='" & strUserNum & "' And r40902 < '2'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         adoaccrpt409.Fields("r40904").Value = 0
      Else
         adoaccrpt409.Fields("r40904").Value = Format(adoaccsum.Fields(0).Value, FAmount)
      End If
   Else
      adoaccrpt409.Fields("r40904").Value = 0
   End If
   adoaccsum.Close
   AutoNoSave
   adoaccrpt409.UpdateBatch
   adoaccrpt409.AddNew
   adoaccrpt409.Fields("r40901").Value = strUserNum
   adoaccrpt409.Fields("r40904").Value = ReportSum(8)
   AutoNoSave
   adoaccrpt409.UpdateBatch
   '2014/1/24 add by sonia
   'Modify by Amy 2020/04/17 公司別改下拉 原:Text6 = "2"
   If strCmp = "J" Then
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.UpdateBatch
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.UpdateBatch
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.UpdateBatch
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.UpdateBatch
      'Add by Amy 2016/11/15
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.UpdateBatch
      adoaccrpt409.AddNew
      adoaccrpt409.Fields("r40901").Value = strUserNum
      AutoNoSave
      adoaccrpt409.UpdateBatch
   End If
   '2014/1/24 end
   adoaccrpt409.Close
'------------------------------------------------
' 負債明細
'------------------------------------------------
   'Modify by Amy 2015/04/09 財務處可能同時兩個人執行此報表,造成資料錯誤
   adoTaie.Execute "delete from accrpt409 where r40901='" & strUserNum & "' And (r40908 is null or r40908 = 0)"
   adoaccrpt409.CursorLocation = adUseClient
   adoaccrpt409.Open "select * from accrpt409 where r40901='" & strUserNum & "' order by r40901 asc, r40908 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   
   adoacc010.CursorLocation = adUseClient
   'Modified by Lydia 2014/12/17 判斷分類會計科目合計金額為零的會計科目不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '2' and a0101 < '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoacc010.Open "select * from acc010 where a0101 >= '2' and a0101 < '3' and SUBSTR(A0101, 1, 2) in (select r409c2 from accrpt409_check Where R409C1='" & strUserNum & "' ) order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   'end 2015/04/09
   Do While adoacc010.EOF = False
      If Text2 <> MsgText(602) Then
         If adoacc010.Fields("a0104").Value = "4" Then
            GoTo Check2
         End If
      End If
      adoacc040.CursorLocation = adUseClient
      strSql = MsgText(601)
      'Modify by Amy 2020/04/17 公司別改下拉 原:Text6
      If strCmp <> MsgText(601) Then
         If InStr(strCmp, "+") > 0 Then
            strSql = " and a0403 In ('" & Replace(strCmp, "+", "','") & "')"
         Else
            '2014/1/23 modify by sonia
            'strSql = " and a0403 = '" & Text6 & "'"
            strSql = " and a0403 = '" & strCmp & "'"
            '2014/1/23 end
         End If
      End If
      If Text3 <> MsgText(601) Then
         strSql = strSql & " and a0401 = " & Val(Text3) & ""
      End If
      If Text1 <> MsgText(601) Then
         strSql = strSql & " and a0402 = " & Val(Text1) & ""
      End If
      strSql = strSql & " and a0405 = '" & adoacc010.Fields("a0101").Value & "'"
      adoacc040.Open "select * from acc040, acc010 where acc040.a0405 = acc010.a0101" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc040.RecordCount <> 0 Then
         Accrpt409Save3
      Else
         If adoaccrpt409.EOF Then
            adoaccrpt409.AddNew
         End If
         adoaccrpt409.Fields("r40901").Value = strUserNum
         adoaccrpt409.Fields("r40905").Value = adoacc010.Fields("a0101").Value
         If IsNull(adoacc010.Fields("a0102").Value) Then
            adoaccrpt409.Fields("r40906").Value = Null
         Else
            adoaccrpt409.Fields("r40906").Value = adoacc010.Fields("a0102").Value
         End If
         Select Case adoacc010.Fields("A0104").Value
            Case "3"
               adoaccsum.CursorLocation = adUseClient
               '2014/1/23 modify by sonia
               'adoaccsum.Open "select SUM(A0408) from acc040 where A0403 = '" & Text6 & "' AND A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND SUBSTR(A0405, 1, 4) = '" & adoacc010.Fields("A0101").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
               adoaccsum.Open "select SUM(A0408) from acc040 where " & stra0403 & " A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND SUBSTR(A0405, 1, 4) = '" & adoacc010.Fields("A0101").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
               If adoaccsum.RecordCount <> 0 Then
                  If IsNull(adoaccsum.Fields(0).Value) Then
                     adoaccrpt409.Fields("r40907").Value = 0
                  Else
                     adoaccrpt409.Fields("r40907").Value = Format(adoaccsum.Fields(0).Value, FAmount)
                  End If
               Else
                  adoaccrpt409.Fields("R40907").Value = 0
               End If
               adoaccsum.Close
            Case "4"
               adoaccrpt409.Fields("R40910").Value = 0
            Case Else
               adoaccrpt409.Fields("r40907").Value = 0
         End Select
        'Modified by Lydia 2014/12/17 判斷分類會計科目合計金額為零的會計科目不顯示
        If adoaccrpt409.Fields("r40907").Value = 0 Or adoaccrpt409.Fields("r40910").Value = 0 Then
           adoaccrpt409.CancelUpdate
        Else
           adoaccrpt409.UpdateBatch
           adoaccrpt409.MoveNext
        End If
      End If
      adoacc040.Close
Check2:
      adoacc010.MoveNext
   Loop
   adoacc010.Close
'------------------------------------------------
' 負債小計
'------------------------------------------------
   'Modified by Lydia 2014/12/17 排除無資料
   If adoaccrpt409.RecordCount = 0 Then
      bolShowMsg = True
      MsgBox MsgText(9211)
      Exit Sub
   End If
   
   adoaccrpt409.MoveNext
   If adoaccrpt409.EOF Then
      adoaccrpt409.AddNew
   End If
   adoaccrpt409.Fields("r40901").Value = strUserNum
   adoaccrpt409.Fields("r40907").Value = ReportSum(4)
   adoaccrpt409.UpdateBatch
   adoaccrpt409.MoveNext
   If adoaccrpt409.EOF Then
      adoaccrpt409.AddNew
   End If
   adoaccrpt409.Fields("r40901").Value = strUserNum
   'Modify By Cheng 2002/01/18
'   adoaccrpt409.Fields("r40906").Value = ReportSum(10)
   adoaccrpt409.Fields("r40906").Value = ReportSum(10001)
   adoaccsum.CursorLocation = adUseClient
   'Modify Amy 2015/04/09 財務處可能同時兩個人執行此報表,造成資料錯誤
   'Modify by Amy 2020/05/22 欄位有文字會錯
   adoaccsum.Open "select  sum(Nvl(r40907,0))  from accrpt409 where r40901='" & strUserNum & "' And r40905 >= '2' and r40905 < '3' And instr(r40907,'－'）=0 And instr(r40907,'＝'）=0 ", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         adoaccrpt409.Fields("r40907").Value = 0
         douTotal2 = 0
      Else
         adoaccrpt409.Fields("r40907").Value = Format(adoaccsum.Fields(0).Value, FAmount)
         douTotal2 = Val(Format(adoaccsum.Fields(0).Value, FAmount))
      End If
   Else
      adoaccrpt409.Fields("r40907").Value = 0
      douTotal2 = 0
   End If
   adoaccsum.Close
   adoaccrpt409.UpdateBatch
'------------------------------------------------
' 股東權益明細
'------------------------------------------------
   adoaccrpt409.MoveNext
   'Modify by Amy 2020/04/17 沒資料會  Error
   If adoaccrpt409.EOF = True Then
      bolShowMsg = True
      MsgBox MsgText(9211)
      Exit Sub
   End If
   adoaccrpt409.MoveNext
   adoacc010.CursorLocation = adUseClient
   adoacc010.Open "select * from acc010 where a0101 >= '3' and a0101 < '4' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc010.EOF = False
      If Text2 <> MsgText(602) Then
         If adoacc010.Fields("a0104").Value = "4" Then
            GoTo Check3
         End If
      End If
      strSql = MsgText(601)
      'Modify by Amy 2020/04/17 公司別改下拉 原:Text6,加作帳公司
      If strCmp <> MsgText(601) Then
        If InStr(strCmp, "+") > 0 Then
            strSql = " and a0403 In ('" & Replace(strCmp, "+", "','") & "')"
        Else
            '2014/1/23 modify by sonia
            'strSql = " and a0403 = '" & Text6 & "'"
            strSql = " and a0403 = '" & strCmp & "'"
            '2014/1/23 end
         End If
      End If
      If Text3 <> MsgText(601) Then
         strSql = strSql & " and a0401 = " & Val(Text3) & ""
      End If
      If Text1 <> MsgText(601) Then
         strSql = strSql & " and a0402 = " & Val(Text1) & ""
      End If
      If adoacc010.Fields(0).Value <> "3222" Then
         strSql = strSql & " and a0405 = '" & adoacc010.Fields("a0101").Value & "'"
      End If
      If adoacc010.Fields(0).Value = "3222" Then
         adoacc040.CursorLocation = adUseClient
         '2007/5/25 MODIFY BY SONIA 分71科目,72科目
         'adoacc040.Open "select sum(decode(substr(a0405, 1, 1), '4', a0408, '6', a0408 * -1, '7', a0408)) from acc040, acc010 where a0405 = a0101 (+) and a0405 >= '4' and a0405 < '8' and a0404 = '" & MsgText(55) & "'" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
         adoacc040.Open "select nvl(sum(decode(substr(a0405, 1, 1), '4', a0408, '6', a0408 * -1, DECODE(SUBSTR(A0405,1,2),'71',a0408,A0408*-1))),0) from acc040, acc010 where a0405 = a0101 (+) and a0405 >= '4' and a0405 < '8' and a0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
         'adoacc040.Open "select r40704 from accrpt407 where r40703 = '" & ReportSum(7) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc040.RecordCount <> 0 Then
            If adoaccrpt409.EOF Then
               adoaccrpt409.AddNew
            End If
            adoaccrpt409.Fields("r40901").Value = strUserNum
            adoaccrpt409.Fields("r40905").Value = adoacc010.Fields("a0101").Value
            If IsNull(adoacc010.Fields("a0102").Value) Then
               adoaccrpt409.Fields("r40906").Value = Null
            Else
               adoaccrpt409.Fields("r40906").Value = adoacc010.Fields("a0102").Value
            End If
            adoaccsum.CursorLocation = adUseClient
            '2014/1/23 modify by sonia 加公司別"1"
            'adoaccsum.Open "select * from acc0b0", adoTaie, adOpenStatic, adLockReadOnly
            'modify by sonia 2016/1/21
            'adoaccsum.Open "select * from acc0b0 where a0b04='1'", adoTaie, adOpenStatic, adLockReadOnly
            'Moidfy by Amy 2020/04/17 公司別改下拉 原:Text6,加作帳公司
            strQ = ""
            If strCmp <> MsgText(601) Then
               If InStr(strCmp, "+") > 0 Then
                    strQ = "And a0b04 In ('" & Replace(strCmp, "+", "','") & "')"
               Else
                    strQ = "And a0b04='" & strCmp & "'"
               End If
               strQ = "Select * from acc0b0 where 1=1 " & strQ
               adoaccsum.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
            Else
               adoaccsum.Open "select * from acc0b0 where a0b04='1'", adoTaie, adOpenStatic, adLockReadOnly
            End If
            'end 2016/1/21
            If adoaccsum.RecordCount <> 0 Then
               If Val(Text1) = 12 Then
                  If IsNull(adoaccsum.Fields("a0b03").Value) Then
                     adoaccrpt409.Fields("r40907").Value = Format(adoacc040.Fields(0).Value, FAmount)
                  Else
                     If Len(adoaccsum.Fields("a0b03").Value) = 6 Then
                        If Mid(adoaccsum.Fields("a0b03").Value, 1, 2) = Val(Text3) Then
                           adoaccrpt409.Fields("r40907").Value = 0
                        Else
                           adoaccrpt409.Fields("r40907").Value = Format(adoacc040.Fields(0).Value, FAmount)
                        End If
                     Else
                        If Mid(adoaccsum.Fields("a0b03").Value, 1, 3) >= Val(Text3) Then
                           adoaccrpt409.Fields("r40907").Value = 0
                        Else
                           adoaccrpt409.Fields("r40907").Value = Format(adoacc040.Fields(0).Value, FAmount)
                        End If
                     End If
                  End If
               Else
                  If IsNull(adoacc040.Fields(0).Value) Then
                     adoaccrpt409.Fields("r40907").Value = 0
                  Else
                     adoaccrpt409.Fields("r40907").Value = Format(adoacc040.Fields(0).Value, FAmount)
                  End If
               End If
            Else
               adoaccrpt409.Fields("r40907").Value = Format(adoacc040.Fields(0).Value, FAmount)
            End If
            adoaccsum.Close
         End If
         adoaccrpt409.UpdateBatch     '2009/4/30 add by sonia因外帳9公司97/1印不出來才發現
         adoaccrpt409.MoveNext        '2009/4/30 add by sonia
         adoacc040.Close
      Else
         adoacc040.CursorLocation = adUseClient
         adoacc040.Open "select * from acc040, acc010 where acc040.a0405 = acc010.a0101" & strSql, adoTaie, adOpenStatic, adLockReadOnly
         If adoacc040.RecordCount <> 0 Then
            Accrpt409Save3
         Else
            If adoaccrpt409.EOF Then
               adoaccrpt409.AddNew
            End If
            adoaccrpt409.Fields("r40901").Value = strUserNum
            adoaccrpt409.Fields("r40905").Value = adoacc010.Fields("a0101").Value
            If IsNull(adoacc010.Fields("a0102").Value) Then
               adoaccrpt409.Fields("r40906").Value = Null
            Else
               adoaccrpt409.Fields("r40906").Value = adoacc010.Fields("a0102").Value
            End If
            adoaccrpt409.Fields("r40907").Value = 0
            adoaccrpt409.UpdateBatch
            adoaccrpt409.MoveNext
         End If
         adoacc040.Close
      End If
Check3:
      adoacc010.MoveNext
   Loop
   adoacc010.Close
'------------------------------------------------
' 股東權益小計
'------------------------------------------------
   adoaccrpt409.MoveNext
   If adoaccrpt409.EOF Then
      adoaccrpt409.AddNew
   End If
   adoaccrpt409.Fields("r40901").Value = strUserNum
   adoaccrpt409.Fields("r40907").Value = ReportSum(4)
   adoaccrpt409.UpdateBatch
   adoaccrpt409.MoveNext
   If adoaccrpt409.EOF Then
      adoaccrpt409.AddNew
   End If
   adoaccrpt409.Fields("r40901").Value = strUserNum
   'Modify By Cheng 2002/01/18
'   adoaccrpt409.Fields("r40906").Value = ReportSum(12)
   adoaccrpt409.Fields("r40906").Value = ReportSum(12001)
   adoaccsum.CursorLocation = adUseClient
   'Modify by Amy 2015/04/09 財務處可能同時兩個人執行此報表,造成資料錯誤
   adoaccsum.Open "select sum(r40907) from accrpt409 where r40901='" & strUserNum & "' And r40905 >= '3'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         adoaccrpt409.Fields("r40907").Value = 0
         douTotal3 = 0
         strTotal3 = "0"
      Else
         adoaccrpt409.Fields("r40907").Value = Format(adoaccsum.Fields(0).Value, FAmount)
         douTotal3 = Val(Format(adoaccsum.Fields(0).Value, FAmount))
         strTotal3 = Format(adoaccsum.Fields(0).Value, FAmount)
      End If
   Else
      adoaccrpt409.Fields("r40907").Value = 0
      douTotal3 = 0
      strTotal3 = "0"
   End If
   adoaccsum.Close
   adoaccrpt409.UpdateBatch
'------------------------------------------------
' 負債總額
'------------------------------------------------
  adoaccrpt409.MoveLast
  adoaccrpt409.MovePrevious
  adoaccrpt409.MovePrevious
  If adoaccrpt409.EOF Then
     adoaccrpt409.AddNew
  End If
  adoaccrpt409.Fields("r40901").Value = strUserNum
  adoaccrpt409.Fields("r40907").Value = ReportSum(4)
  adoaccrpt409.UpdateBatch
  adoaccrpt409.MoveNext
  If adoaccrpt409.EOF Then
     adoaccrpt409.AddNew
  End If
  adoaccrpt409.Fields("r40901").Value = strUserNum
   'Modify By Cheng 2002/01/18
'  adoaccrpt409.Fields("r40906").Value = ReportSum(13)
  adoaccrpt409.Fields("r40906").Value = ReportSum(13001)
  adoaccrpt409.Fields("r40907").Value = douTotal2 + douTotal3
  adoaccrpt409.UpdateBatch
  adoaccrpt409.MoveNext
  If adoaccrpt409.EOF Then
     adoaccrpt409.AddNew
  End If
  adoaccrpt409.Fields("r40901").Value = strUserNum
  adoaccrpt409.Fields("r40907").Value = ReportSum(8)
  adoaccrpt409.UpdateBatch
End Sub

'Modify by Amy 2020/05/29 +strA1 畫面條件
Private Sub Accrpt409Save2(Optional ByVal strA1 As String = "")
   Dim stSQL As String 'Add by Amy 2020/05/29
   
   adoaccrpt409.AddNew
   adoaccrpt409.Fields("r40901").Value = strUserNum
   adoaccrpt409.Fields("r40902").Value = adoacc040.Fields("a0405").Value
   If IsNull(adoacc040.Fields("a0102").Value) Then
      adoaccrpt409.Fields("r40903").Value = Null
   Else
      adoaccrpt409.Fields("r40903").Value = adoacc040.Fields("a0102").Value
   End If
   Select Case adoacc010.Fields("A0104").Value
      Case "3"
         adoaccsum.CursorLocation = adUseClient
         '2014/1/23 modify by sonia
         'adoaccsum.Open "select SUM(A0408) from acc040 where A0403 = '" & Text6 & "' AND A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND SUBSTR(A0405, 1, 4) = '" & adoacc010.Fields("A0101").Value & "' and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
         'Modify by Amy 2020/05/29 傳入Where 條件
         stSQL = "select SUM(A0408) from acc040 where SUBSTR(A0405, 1, 4) = '" & adoacc010.Fields("A0101").Value & "' and a0404 = '" & MsgText(55) & "' " & strA1
         adoaccsum.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
         'end 2020/05/29
         If adoaccsum.RecordCount <> 0 Then
            If IsNull(adoaccsum.Fields(0).Value) Then
               adoaccrpt409.Fields("r40904").Value = 0
            Else
               adoaccrpt409.Fields("r40904").Value = Format(adoaccsum.Fields(0).Value, FAmount)
            End If
         Else
            adoaccrpt409.Fields("R40904").Value = 0
         End If
         adoaccsum.Close
      Case "4"
         adoaccsum.CursorLocation = adUseClient
         '2014/1/23 modify by sonia
         'adoaccsum.Open "select SUM(A0408) from acc040 where A0403 = '" & Text6 & "' AND A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND A0405 = '" & adoacc010.Fields("A0101").Value & "' and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
         'Modify by Amy 2020/05/29 傳入Where 條件
         stSQL = "select SUM(A0408) from acc040 where A0405 = '" & adoacc010.Fields("A0101").Value & "' and a0404 = '" & MsgText(55) & "' " & strA1
         adoaccsum.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
         If IsNull(adoaccsum.Fields(0).Value) Then
            adoaccrpt409.Fields("r40909").Value = 0
         Else
            adoaccrpt409.Fields("r40909").Value = Format(adoaccsum.Fields(0).Value, FAmount)
         End If
         adoaccsum.Close
      Case Else
         adoaccsum.CursorLocation = adUseClient
         '2014/1/23 modify by sonia
         'adoaccsum.Open "select SUM(A0408) from acc040 where A0403 = '" & Text6 & "' AND A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND A0405 = '" & adoacc010.Fields("A0101").Value & "' and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
         'Modify by Amy 2020/05/29 傳入Where 條件
         stSQL = "select SUM(A0408) from acc040 where A0405 = '" & adoacc010.Fields("A0101").Value & "' and a0404 = '" & MsgText(55) & "' " & strA1
         adoaccsum.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
         If IsNull(adoaccsum.Fields(0).Value) Then
            adoaccrpt409.Fields("r40904").Value = Null
         Else
            If adoaccsum.Fields(0).Value = 0 Then
               adoaccrpt409.Fields("r40904").Value = Null
            Else
               adoaccrpt409.Fields("r40904").Value = Format(adoaccsum.Fields(0).Value, FAmount)
            End If
         End If
         adoaccsum.Close
   End Select
   
'Modified by Lydia 2014/12/17 令金額為零的會計科目不顯示
If adoaccrpt409.Fields("r40904").Value = 0 Or adoaccrpt409.Fields("r40909").Value = 0 Then
   adoaccrpt409.CancelUpdate
Else
   'AutoNoSave 'Mark by Amy 不使用
   adoaccrpt409.UpdateBatch
   intCount(0) = intCount(0) + 1
End If
End Sub
'*************************************************
'  儲存於資產負債表右方之資料(1090529不使用)
'
'*************************************************
Private Sub Accrpt409Save3()
   If adoaccrpt409.EOF Then
      adoaccrpt409.AddNew
   End If
   adoaccrpt409.Fields("r40901").Value = strUserNum
   adoaccrpt409.Fields("r40905").Value = adoacc040.Fields("a0405").Value
   If IsNull(adoacc040.Fields("a0102").Value) Then
      adoaccrpt409.Fields("r40906").Value = Null
   Else
      adoaccrpt409.Fields("r40906").Value = adoacc040.Fields("a0102").Value
   End If
   Select Case adoacc010.Fields("A0104").Value
      Case "3"
         adoaccsum.CursorLocation = adUseClient
         '2014/1/23 modify by sonia
         'adoaccsum.Open "select SUM(A0408) from acc040 where A0403 = '" & Text6 & "' AND A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND SUBSTR(A0405, 1, 4) = '" & adoacc010.Fields("A0101").Value & "' and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
         adoaccsum.Open "select SUM(A0408) from acc040 where " & stra0403 & " A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND SUBSTR(A0405, 1, 4) = '" & adoacc010.Fields("A0101").Value & "' and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoaccsum.RecordCount <> 0 Then
            If IsNull(adoaccsum.Fields(0).Value) Then
               adoaccrpt409.Fields("r40907").Value = 0
            Else
               adoaccrpt409.Fields("r40907").Value = Format(adoaccsum.Fields(0).Value, FAmount)
            End If
         Else
            adoaccrpt409.Fields("R40907").Value = 0
         End If
         adoaccsum.Close
      Case "4"
         adoaccsum.CursorLocation = adUseClient
         '2014/1/23 modify by sonia
         'adoaccsum.Open "select SUM(A0408) from acc040 where A0403 = '" & Text6 & "' AND A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND A0405 = '" & adoacc010.Fields("A0101").Value & "' and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
         adoaccsum.Open "select SUM(A0408) from acc040 where " & stra0403 & " A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND A0405 = '" & adoacc010.Fields("A0101").Value & "' and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If IsNull(adoaccsum.Fields(0).Value) Then
            adoaccrpt409.Fields("r40910").Value = 0
         Else
            adoaccrpt409.Fields("r40910").Value = Format(adoaccsum.Fields(0).Value, FAmount)
         End If
         adoaccsum.Close
      Case Else
         adoaccsum.CursorLocation = adUseClient
         '2014/1/23 modify by sonia
         'adoaccsum.Open "select SUM(A0408) from acc040 where A0403 = '" & Text6 & "' AND A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND A0405 = '" & adoacc010.Fields("A0101").Value & "' and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
         adoaccsum.Open "select SUM(A0408) from acc040 where " & stra0403 & " A0401 = " & Val(Text3) & " AND A0402 = " & Val(Text1) & " AND A0405 = '" & adoacc010.Fields("A0101").Value & "' and a0404 = '" & MsgText(55) & "'", adoTaie, adOpenStatic, adLockReadOnly
         If IsNull(adoaccsum.Fields(0).Value) Then
            adoaccrpt409.Fields("r40907").Value = Null
         Else
            If adoaccsum.Fields(0).Value = 0 Then
               adoaccrpt409.Fields("r40907").Value = Null
            Else
               adoaccrpt409.Fields("r40907").Value = Format(adoaccsum.Fields(0).Value, FAmount)
            End If
         End If
         adoaccsum.Close
   End Select
   
'Modified by Lydia 2014/12/17 令金額為零的會計科目不顯示
If adoaccrpt409.Fields("r40907").Value = 0 Or adoaccrpt409.Fields("r40910").Value = 0 Then
   adoaccrpt409.CancelUpdate
Else
   adoaccrpt409.UpdateBatch
   adoaccrpt409.MoveNext
End If
End Sub

'Add by Amy 2020/05/29 儲存會科3222「股東權益」資料於資產負債表右方
Private Sub Accrpt409Save4()
    Dim stSQL As String
    
    adoaccrpt409.AddNew
    adoaccrpt409.Fields("r40901").Value = strUserNum
    adoaccrpt409.Fields("r40902").Value = adoacc010.Fields("a0101").Value
    If IsNull(adoacc010.Fields("a0102").Value) Then
        adoaccrpt409.Fields("r40903").Value = Null
    Else
        adoaccrpt409.Fields("r40903").Value = adoacc010.Fields("a0102").Value
    End If
    If strCmp <> MsgText(601) Then
        If InStr(strCmp, "+") > 0 Then
            stSQL = "And a0b04 In ('" & Replace(strCmp, "+", "','") & "')"
        Else
            stSQL = "And a0b04='" & strCmp & "' "
        End If
    Else
        stSQL = "And a0b04='1' "
    End If
    stSQL = "Select * from acc0b0 where 1=1 " & stSQL
    adoaccsum.CursorLocation = adUseClient
    adoaccsum.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
    If adoaccsum.RecordCount <> 0 Then
        '畫面輸12月
        If Val(Text1) = 12 Then
            '年度結轉日為空(未結轉)
            If IsNull(adoaccsum.Fields("a0b03").Value) Then
                adoaccrpt409.Fields("r40904").Value = Format(adoacc040.Fields(0).Value, FAmount)
            '年度已結轉
            Else
                '年度結轉日長度為6
                If Len(adoaccsum.Fields("a0b03").Value) = 6 Then
                    If Mid(adoaccsum.Fields("a0b03").Value, 1, 2) = Val(Text3) Then
                        adoaccrpt409.Fields("r40904").Value = 0
                    Else
                        adoaccrpt409.Fields("r40904").Value = Format(adoacc040.Fields(0).Value, FAmount)
                    End If
                '年度結轉日長度「不為」6
                Else
                    If Mid(adoaccsum.Fields("a0b03").Value, 1, 3) >= Val(Text3) Then
                        adoaccrpt409.Fields("r40904").Value = 0
                    Else
                        adoaccrpt409.Fields("r40904").Value = Format(adoacc040.Fields(0).Value, FAmount)
                    End If
                End If
            End If
        '不是12月
        Else
            If IsNull(adoacc040.Fields(0).Value) Then
                adoaccrpt409.Fields("r40904").Value = 0
            Else
                adoaccrpt409.Fields("r40904").Value = Format(adoacc040.Fields(0).Value, FAmount)
            End If
        End If
    'adoaccsum.RecordCount=0
    Else
        adoaccrpt409.Fields("r40904").Value = Format(adoacc040.Fields(0).Value, FAmount)
    End If
    adoaccsum.Close
    adoaccrpt409.UpdateBatch
    adoaccrpt409.MoveNext
End Sub

'改寫暫存檔資料寫入方式,故Excel 也需調整
Private Function IsExcelSave() As Boolean
    Dim xlsSalesPoint As New Excel.Application
    Dim WksAccrpt409 As New Worksheet
    Dim strQ As String, strFileName As String
    Dim bolExcel As Boolean
    Dim iRow_R As Integer  '資產總列數
    
    IsExcelSave = False
    
    ReDim strFieldN(9)
    ReDim intWidth(9)

    strFieldN = Array("會計科目", "科目名稱", "金額", "借方金額")
    intWidth = Array(13, 20, 13, 13)
    
    stRptName = ReportTitle(409)
    
    strFileName = strExcelPath & Trim(Replace(stRptName, "*", "")) & Format(Now, "yyyymmddhhmmss") & MsgText(43)
    If Dir(strFileName) = MsgText(601) Then
       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
          MkDir strExcelPath
       End If
    Else
       Kill strFileName
    End If
   
    xlsSalesPoint.SheetsInNewWorkbook = 3 '預設工作表數量
    xlsSalesPoint.Workbooks.add
    Set WksAccrpt409 = xlsSalesPoint.Worksheets(1)
    'xlsSalesPoint.Visible = True
    
    stCellFormat = "#,##0.00 ;[紅色]-#,##0.00"
    intField = 65:  iRow = 1: intSum(2) = 1: intSum(3) = 1
    
    Call SetTitle(WksAccrpt409, False)
    intTitleRow = iRow
    iRow = iRow + 1
    
    '讀取「資產」資料
    strQ = "Select r40902,r40903,r40909,r40904,r40905,r40906,r40910,r40907 From Accrpt409 Where r40901='" & strUserNum & "' And r40902<'2' order by r40902"
    If adoaccrpt409.State = adStateOpen Then adoaccrpt409.Close
    adoaccrpt409.CursorLocation = adUseClient
    adoaccrpt409.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoaccrpt409.RecordCount <= 0 Then
        WksAccrpt409.Range(Chr(intField + GetValue("科目名稱")) & iRow).Value = "無資產資料"
        iRow = iRow + 1
    Else
        Call PutData(WksAccrpt409, True)
    End If
    iRow = iRow + 1
    iRow_R = iRow
    
    iRow = intTitleRow + 1
    '讀取「負債/業主權益」資料
    strQ = "Select r40902,r40903,r40909,r40904,r40905,r40906,r40910,r40907 From Accrpt409 Where r40901='" & strUserNum & "' And r40902>='2' order by r40902"
    If adoaccrpt409.State = adStateOpen Then adoaccrpt409.Close
    adoaccrpt409.CursorLocation = adUseClient
    adoaccrpt409.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoaccrpt409.RecordCount <= 0 Then
        WksAccrpt409.Range(Chr(intField + GetValue("科目名稱")) & iRow).Value = "無負債/業主權益資料"
        iRow = iRow + 1
    Else
        Call PutData(WksAccrpt409, False)
    End If
    iRow = iRow + 1
    
    If iRow_R > iRow Then iRow = iRow_R
    
    With WksAccrpt409
        '資產
        .Range(Chr(intField + GetValue("科目名稱")) & iRow).Value = ReportSum(9001)
        .Range(Chr(intField + GetValue("科目名稱")) & iRow).HorizontalAlignment = xlLeft
        .Range(Chr(intField + GetValue("借方金額")) & iRow).Value = "=Sum(" & Chr(intField + GetValue("借方金額")) & intTitleRow + 1 & ":" & _
                                                                                                                                               Chr(intField + GetValue("借方金額")) & iRow - 1 & ")"
        .Range(Chr(intField + GetValue("借方金額")) & iRow).HorizontalAlignment = xlRight
        '負債/業主權益
        .Range(Chr(intField + (GetValue("科目名稱") + UBound(strFieldN) + 1)) & iRow).Value = ReportSum(13001)
        .Range(Chr(intField + (GetValue("科目名稱") + UBound(strFieldN) + 1)) & iRow).HorizontalAlignment = xlLeft
        .Range(Chr(intField + (GetValue("借方金額") + UBound(strFieldN) + 1)) & iRow).Value = "=" & Chr(intField + (GetValue("借方金額") + UBound(strFieldN) + 1)) & intSum(2) & "+" & _
                                                                                                                                                                                         Chr(intField + (GetValue("借方金額") + UBound(strFieldN) + 1)) & intSum(3)
        .Range(Chr(intField + (GetValue("借方金額") + UBound(strFieldN) + 1)) & iRow).HorizontalAlignment = xlRight
   
        .Range(Chr(intField) & intTitleRow & ":" & Chr(intField + (2 * UBound(strFieldN) + 1)) & iRow).Font.Size = 11
        '設定格線
        Call SetExcelLine(2, WksAccrpt409, Chr(intField) & intTitleRow + 1 & ":" & Chr(intField + (2 * UBound(strFieldN) + 1)) & iRow)
        Call SetExcelLine(1, WksAccrpt409, Chr(intField) & intTitleRow & ":" & Chr(intField) & iRow)
        Call SetExcelLine(1, WksAccrpt409, Chr(intField + GetValue("會計科目") + UBound(strFieldN) + 1) & intTitleRow & ":" & Chr(intField + GetValue("會計科目") + UBound(strFieldN) + 1) & iRow)
        '小計格線
        Call SetExcelLine(3, WksAccrpt409, Chr(intField + GetValue("借方金額") + UBound(strFieldN) + 1) & intSum(2))
        Call SetExcelLine(3, WksAccrpt409, Chr(intField + GetValue("借方金額") + UBound(strFieldN) + 1) & intSum(3))
        '總額格線
        Call SetExcelLine(4, WksAccrpt409, Chr(intField + GetValue("借方金額")) & iRow)
        Call SetExcelLine(4, WksAccrpt409, Chr(intField + GetValue("借方金額") + UBound(strFieldN) + 1) & iRow)
        
        Call SetTitle(WksAccrpt409, True) 'Add by Amy 2020/06/09
        .PageSetup.PaperSize = 9 '設定紙張 A4
        .PageSetup.Orientation = xlPortrait '直印 '橫印 xlLandscape
        .PageSetup.PrintTitleRows = "$1:$" & intTitleRow '表頭保留7列
        .PageSetup.PrintArea = "$A$1:$" & Chr(intField + (2 * UBound(strFieldN) + 1)) & "$" & iRow '設定列印範圍
        
        .PageSetup.TopMargin = xlsSalesPoint.InchesToPoints(0.3)
        .PageSetup.BottomMargin = xlsSalesPoint.InchesToPoints(0.7)
        .PageSetup.LeftMargin = xlsSalesPoint.InchesToPoints(0.39) '左邊界
        .PageSetup.RightMargin = xlsSalesPoint.InchesToPoints(0) '右邊界
        .PageSetup.HeaderMargin = xlsSalesPoint.InchesToPoints(0.5)
        .PageSetup.FooterMargin = xlsSalesPoint.InchesToPoints(0.5)
        
        .PageSetup.Zoom = 80 '縮放比例 100
    End With
   '版本
    If Val(xlsSalesPoint.Version) < 12 Then
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
    Else
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
    End If
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   Set xlsSalesPoint = Nothing
   Set WksAccrpt409 = Nothing
   MsgBox "Excel檔案產生完成！（檔案位置：" & strFileName & "）"
   
End Function

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = 1 To UBound(strFieldN)
       If UCase(strFieldN(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Private Sub SetTitle(ByRef Wks As Worksheet, ByVal IsLast As Boolean)
    Dim ii As Integer, strTp As String, intCol As Integer
    
    With Wks
        If IsLast = False Then
            '***表頭設定***
            .Range(Chr(intField) & iRow).Value = stRptName
            .Range(Chr(intField) & iRow).Font.Bold = True
            .Range(Chr(intField) & iRow & ":" & Chr(UBound(strFieldN) * 2 + 67) & iRow).HorizontalAlignment = xlCenter
            .Range(Chr(intField) & iRow & ":" & Chr(UBound(strFieldN) * 2 + 67) & iRow).MergeCells = True
            iRow = iRow + 1
            
            If strCmp = MsgText(601) Then
                strTp = strCmpN
            Else
                strTp = strCmp & " " & strCmpN
            End If
            .Range(Chr(intField + GetValue("金額")) & iRow).Value = "公司別："
            .Range(Chr(intField + GetValue("金額")) & iRow).Font.Bold = True
            .Range(Chr(intField + GetValue("金額")) & iRow).HorizontalAlignment = xlRight
            
            .Range(Chr(intField + GetValue("借方金額")) & iRow).Value = strTp
            .Range(Chr(intField + GetValue("借方金額")) & iRow).Font.Bold = True
            .Range(Chr(intField + GetValue("借方金額")) & iRow).HorizontalAlignment = xlLeft
            iRow = iRow + 1
            
            .Range(Chr(intField + GetValue("金額")) & iRow).Value = "年度："
            .Range(Chr(intField + GetValue("金額")) & iRow).Font.Bold = True
            .Range(Chr(intField + GetValue("金額")) & iRow).HorizontalAlignment = xlRight
            
            .Range(Chr(intField + GetValue("借方金額")) & iRow).Value = Text3
            .Range(Chr(intField + GetValue("借方金額")) & iRow).Font.Bold = True
            .Range(Chr(intField + GetValue("借方金額")) & iRow).HorizontalAlignment = xlLeft
            iRow = iRow + 1
            
            .Range(Chr(intField) & iRow).Value = "列印人員："
            .Range(Chr(intField) & iRow).Font.Bold = True
            .Range(Chr(intField) & iRow).HorizontalAlignment = xlRight
            
            .Range(Chr(intField + GetValue("科目名稱")) & iRow).Value = strUserName
            .Range(Chr(intField + GetValue("科目名稱")) & iRow).Font.Bold = True
            .Range(Chr(intField + GetValue("科目名稱")) & iRow).HorizontalAlignment = xlLeft
            
            .Range(Chr(intField + GetValue("金額")) & iRow).Value = "截止月份："
            .Range(Chr(intField + GetValue("金額")) & iRow).Font.Bold = True
            .Range(Chr(intField + GetValue("金額")) & iRow).HorizontalAlignment = xlRight
            
            .Range(Chr(intField + GetValue("借方金額")) & iRow).Value = Text1
            .Range(Chr(intField + GetValue("借方金額")) & iRow).Font.Bold = True
            .Range(Chr(intField + GetValue("借方金額")) & iRow).HorizontalAlignment = xlLeft
            
            .Range(Chr(intField + (2 * GetValue("借方金額"))) & iRow).Value = "列印日期："
            .Range(Chr(intField + (2 * GetValue("借方金額"))) & iRow).Font.Bold = True
            .Range(Chr(intField + (2 * GetValue("借方金額"))) & iRow).HorizontalAlignment = xlRight
            
            .Range(Chr(intField + (2 * GetValue("借方金額") + 1)) & iRow).Value = CFDate(ACDate(ServerDate))
            .Range(Chr(intField + (2 * GetValue("借方金額") + 1)) & iRow).Font.Bold = True
            .Range(Chr(intField + (2 * GetValue("借方金額") + 1)) & iRow).HorizontalAlignment = xlLeft
            iRow = iRow + 1
            
            For ii = 0 To UBound(strFieldN)
                intCol = intField + ii
                .Columns(Chr(intCol) & ":" & Chr(intCol)).ColumnWidth = intWidth(ii)
                .Range(Chr(intCol) & iRow).Value = strFieldN(ii)
                .Range(Chr(intCol) & iRow).HorizontalAlignment = xlCenter
                
                intCol = intField + ii + UBound(strFieldN) + 1
                .Columns(Chr(intCol) & ":" & Chr(intCol)).ColumnWidth = intWidth(ii)
                .Range(Chr(intCol) & iRow).Value = strFieldN(ii)
                .Range(Chr(intCol) & iRow).HorizontalAlignment = xlCenter
            Next ii
            '框線
            Call SetExcelLine(0, Wks, Chr(intField) & iRow & ":" & Chr(intField + (2 * UBound(strFieldN) + 1)) & iRow)
        Else
            For ii = 0 To UBound(strFieldN)
                If ii = GetValue("金額") Then
                    .Range(Chr(intField + ii) & intTitleRow).Value = ""
                    .Range(Chr(intField + ii + UBound(strFieldN) + 1) & intTitleRow).Value = ""
                'Add by Amy 2020/06/09
                ElseIf ii = GetValue("借方金額") Then
                    .Range(Chr(intField + ii + UBound(strFieldN) + 1) & intTitleRow).Value = "貸方金額"
                End If
            Next ii
        End If
    End With
End Sub

Private Sub PutData(ByRef Wks As Worksheet, ByVal IsAssets As Boolean)
    Dim ii As Integer, intCol As Integer
    Dim bolRight As Boolean '靠右
    Dim strTemp As String
  
    adoaccrpt409.MoveFirst
    Do While adoaccrpt409.EOF = False
        For ii = LBound(strFieldN) To UBound(strFieldN)
            intCol = ii: bolRight = False
            If IsAssets = False Then intCol = intCol + UBound(strFieldN) + 1
            strTemp = "" & adoaccrpt409.Fields(ii)
        
            '合計
            If InStr(adoaccrpt409.Fields("r40902"), "z") > 0 And IsAssets = False Then
                If ii = GetValue("借方金額") Then
                    '負債
                    If Left(adoaccrpt409.Fields("r40902"), 1) = "2" Then
                        intSum(2) = iRow
                        Wks.Range(Chr(intField + intCol) & iRow).Value = "=Sum(" & Chr(intField + intCol) & intTitleRow + 1 & ":" & Chr(intField + intCol) & iRow - 1 & ")"
                    '業主權益
                    Else
                        intSum(3) = iRow
                        Wks.Range(Chr(intField + intCol) & iRow).Value = "=Sum(" & Chr(intField + intCol) & intSum(2) + 1 & ":" & Chr(intField + intCol) & iRow - 1 & ")"
                    End If
                ElseIf ii <> GetValue("會計科目") Then
                    Wks.Range(Chr(intField + intCol) & iRow).Value = strTemp
                End If
            '資料
            Else
                Wks.Range(Chr(intField + intCol) & iRow).Value = strTemp
            End If
            If ii = GetValue("金額") Or ii = GetValue("借方金額") Then
                Wks.Range(Chr(intField + intCol) & iRow).HorizontalAlignment = xlRight
                Wks.Range(Chr(intField + intCol) & iRow).NumberFormatLocal = stCellFormat
            Else
                Wks.Range(Chr(intField + intCol) & iRow).HorizontalAlignment = xlLeft
            End If
        Next ii
        iRow = iRow + 1
        adoaccrpt409.MoveNext
    Loop
End Sub
'end 2020/05/29


'*************************************************
' 自動編號存入 r40908 欄位
'
'*************************************************
Private Sub AutoNoSave()
   intAutoNo = intAutoNo + 1
   adoaccrpt409.Fields("r40908").Value = intAutoNo
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   'Modify by Amy 2020/04/17 公司改下拉
'   Text6 = ""
'   Text7 = "台一　專利商標/智權"
   CboCmp = ""
   'end 2020/04/17
   Text3 = ""
   Text1 = ""
   'modify by sonia 2016/2/15 改預設含子目
   Text2 = "Y"
   CboCmp.SetFocus 'Modify by Amy 2020/04/17 原:Text6
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
'Modify by Amy 2020/04/17 +bolShowMsg
Public Function FormCheck(bolShowMsg As Boolean) As Boolean
   'Add by Amy 2020/04/17
   Dim bCancel As Boolean
   
   If Trim(CboCmp) <> MsgText(601) Then
      Call CboCmp_Validate(bCancel)
      If bCancel = True Then
          bolShowMsg = True
          Exit Function
      End If
   End If
   'end 2020/04/17
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

'Add By Sindy 2013/5/27
'產生Excel檔案
Public Sub IsExcelSave_Old()
Dim xlsSalesPoint As New Excel.Application
Dim WksAccrpt409 As New Worksheet
Dim strFileName As String
Dim iRow As Integer, iRow_R As Integer, iRow_E 'Modify by Amy 2016/08/10 拆成左右列/最後一筆 總額列號
Dim stCellFormat As String
Dim stRptName As String
Dim Rc As String '欄位座標
Dim MaxCol As String '最右的欄位代碼
'Add by Amy 2016/08/10
Dim intTitleRow As Integer, intSum(1) As Integer '抬頭列數/0:負債小計/1:股東權益小計 列號
Dim strTotal1(1) As String, strTotal2(1) As String '0:左邊總額名稱 1:數值/ 0:右邊總額名稱 1:數值

'on error GoTo ErrHnd
   
   '讀取資料
   If adoaccrpt409.State = adStateOpen Then
      adoaccrpt409.Close
   End If
   adoaccrpt409.CursorLocation = adUseClient
   'Modify by Amy 2015/04/09 財務處可能同時兩個人執行此報表,造成資料錯誤
   adoaccrpt409.Open "select * from accrpt409 where r40901='" & strUserNum & "' order by r40908 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt409.RecordCount <= 0 Then
      adoaccrpt409.Close
      Exit Sub
   End If
   
   MaxCol = Chr(Asc("a") + 7)
   
   stCellFormat = "#,##0.00 ;[紅色]-#,##0.00 "
   
   stRptName = ReportTitle(409)
   
   strFileName = strExcelPath & Trim(Replace(stRptName, "*", "")) & Format(Now, "yyyymmddhhmmss") & MsgText(43)
   If Dir(strFileName) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strFileName
   End If
   
   xlsSalesPoint.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
   xlsSalesPoint.Workbooks.add
   Set WksAccrpt409 = xlsSalesPoint.Worksheets(1)
   With WksAccrpt409
      iRow = 1
      .Range("a" & iRow).Value = stRptName
      Rc = MaxCol & iRow
      With .Range("a" & iRow & ":" & Rc)
         .Font.Size = 18
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
         .MergeCells = True
      End With
      
      iRow = iRow + 2
      .Range("c" & iRow).Value = "公司別："
      With .Range("c" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlRight
      End With
      
      'Modify By Sindy 2014/9/3
      '.Range("d" & iRow).Value = Text6 & "  " & Text7
      'Modify by Amy 2020/04/17 公司別改抓變數
      '.Range("d" & iRow).Value = IIf(Text6 = "2", "J", Text6) & "  " & IIf(Text7 = "", "台一　專利商標/智權", Text7)
      .Range("d" & iRow).Value = strCmp & "  " & strCmpN
      '2014/9/3 END
      .Columns("d").ColumnWidth = 24
      With .Range("d" & iRow)
         .Font.Size = 12
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
         .MergeCells = True
      End With
      
      iRow = iRow + 1
      .Range("c" & iRow).Value = "年度："
      With .Range("c" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlRight
      End With

      .Range("d" & iRow).Value = Text3
      With .Range("d" & iRow)
         .Font.Size = 12
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
         .MergeCells = True
      End With
      
      iRow = iRow + 1
      .Range("a" & iRow).Value = "列印人員：" & StaffQuery(strUserNum)
      With .Range("a" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
      End With
      'Mark by Amy 2016/08/10 列印時字無法完全顯示,故合併
'      .Range("b" & iRow).Value = StaffQuery(strUserNum)
'      With .Range("b" & iRow)
'         .Font.Size = 12
'         .Font.Bold = False
'         .HorizontalAlignment = xlLeft
'      End With
      
      .Range("c" & iRow).Value = "截止月份："
      With .Range("c" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlRight
      End With

      .Range("d" & iRow).Value = Text1
      With .Range("d" & iRow)
         .Font.Size = 12
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
         .MergeCells = True
      End With
      
      .Range("g" & iRow).Value = "列印日期："
      With .Range("g" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlRight
         .MergeCells = True
      End With
      
      .Range("h" & iRow).Value = CFDate(ACDate(ServerDate))
      With .Range("h" & iRow)
         .Font.Size = 12
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
      End With
      
      iRow = iRow + 2
      
      .Range("a" & iRow).Value = "會計科目"
      .Columns("a").ColumnWidth = 8
      With .Range("a" & iRow)
         .Font.Size = 11
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
      End With
      
      .Range("b" & iRow).Value = "科目名稱"
      .Columns("b").ColumnWidth = 20
      With .Range("b" & iRow)
         .Font.Size = 11
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
      End With
      
      .Range("c" & iRow).Value = ""
      .Columns("c").ColumnWidth = 13
      With .Range("c" & iRow)
         .Font.Size = 11
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
      End With
      
      .Range("d" & iRow).Value = "借方金額"
      .Columns("d").ColumnWidth = 13.5
      With .Range("d" & iRow)
         .Font.Size = 11
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
      End With
      
      .Range("e" & iRow).Value = "會計科目"
      .Columns("e").ColumnWidth = 8
      With .Range("e" & iRow)
         .Font.Size = 11
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
      End With
      
      .Range("f" & iRow).Value = "科目名稱"
      .Columns("f").ColumnWidth = 20
      With .Range("f" & iRow)
         .Font.Size = 11
         .Font.Bold = True
         .HorizontalAlignment = xlLeft
      End With
      
      .Range("g" & iRow).Value = ""
      .Columns("g").ColumnWidth = 13
      With .Range("g" & iRow)
         .Font.Size = 11
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
      End With
      
      .Range("h" & iRow).Value = "貸方金額"
      .Columns("h").ColumnWidth = 13.5
      With .Range("h" & iRow)
         .Font.Size = 11
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
      End With
      
      'Modify by Amy 2016/08/10 +框線
      intTitleRow = iRow: iRow_R = iRow: iRow_E = 0
      Call SetExcelLine(0, WksAccrpt409, "a" & intTitleRow & ":" & MaxCol & intTitleRow)
      
      iRow = iRow + 1: iRow_R = iRow_R + 1
      adoaccrpt409.MoveFirst
      Do While Not adoaccrpt409.EOF
         
         With .Range("A" & iRow)
            .Font.Size = 11
            .HorizontalAlignment = xlLeft
         End With
         With .Range("B" & iRow)
            .Font.Size = 11
            .HorizontalAlignment = xlLeft
         End With
         With .Range("C" & iRow)
            .Font.Size = 11
            .HorizontalAlignment = xlRight
            .NumberFormatLocal = stCellFormat
         End With
         With .Range("D" & iRow)
            .Font.Size = 11
            .HorizontalAlignment = xlRight
            .NumberFormatLocal = stCellFormat
         End With
         'Mark by Amy 2018/12/28 往下搬,避免右邊比較多資料沒設到
'         With .Range("E" & iRow)
'            .Font.Size = 12
'            .HorizontalAlignment = xlLeft
'         End With
'         With .Range("F" & iRow)
'            .Font.Size = 11
'            .HorizontalAlignment = xlLeft
'         End With
'         With .Range("G" & iRow)
'            .Font.Size = 11
'            .HorizontalAlignment = xlRight
'            .NumberFormatLocal = stCellFormat
'         End With
'         With .Range("H" & iRow)
'            .Font.Size = 11
'            .HorizontalAlignment = xlRight
'            .NumberFormatLocal = stCellFormat
'         End With
         
         '左邊資料-資產
         If InStr("" & adoaccrpt409.Fields("r40903"), "資產總額") > 0 Then
            strTotal1(0) = "" & adoaccrpt409.Fields("r40903")
            strTotal1(1) = "" & adoaccrpt409.Fields("r40904")
            If iRow_E = 0 Or iRow > iRow_R Then
                iRow_E = iRow + 1
            Else
                iRow_E = iRow_R '右邊資料比較多
            End If
         ElseIf InStr("" & adoaccrpt409.Fields("r40904"), "－") = 0 And InStr("" & adoaccrpt409.Fields("r40904"), "＝") = 0 _
           And Not (Trim("" & adoaccrpt409.Fields("r40902")) = MsgText(601) And Trim("" & adoaccrpt409.Fields("r40903")) = MsgText(601) _
           And Trim("" & adoaccrpt409.Fields("r40904")) = MsgText(601) And Trim("" & adoaccrpt409.Fields("r40909")) = MsgText(601)) Then
            .Range("a" & iRow).Value = "" & adoaccrpt409.Fields("r40902")
            .Range("b" & iRow).Value = "" & adoaccrpt409.Fields("r40903")
            .Range("c" & iRow).Value = "" & adoaccrpt409.Fields("r40909")
            .Range("d" & iRow).Value = "" & adoaccrpt409.Fields("r40904")
            iRow = iRow + 1
         End If
         
         '右邊資料-負債/股東權益
         If InStr("" & adoaccrpt409.Fields("r40906"), "負債與股東權益總額") > 0 Then
            strTotal2(0) = "" & adoaccrpt409.Fields("r40906")
            strTotal2(1) = "" & adoaccrpt409.Fields("r40907")
            If iRow_E = 0 Then
                iRow_E = iRow_R + 1
            ElseIf iRow > iRow_R Then
                iRow_E = iRow
            Else
                iRow_E = iRow_R + 1 '右邊資料比較多
            End If
         ElseIf InStr("" & adoaccrpt409.Fields("r40907"), "－") = 0 And InStr("" & adoaccrpt409.Fields("r40907"), "＝") = 0 _
           And Not (Trim("" & adoaccrpt409.Fields("r40905")) = MsgText(601) And Trim("" & adoaccrpt409.Fields("r40906")) = MsgText(601) _
           And Trim("" & adoaccrpt409.Fields("r40907")) = MsgText(601) And Trim("" & adoaccrpt409.Fields("r40910")) = MsgText(601)) Then
            .Range("e" & iRow_R).Value = "" & adoaccrpt409.Fields("r40905")
            'Modify by Amy 2018/12/28 格式從上搬下來,避免右邊比較多資料沒設到
            .Range("e" & iRow_R).Font.Size = 12: .Range("e" & iRow_R).HorizontalAlignment = xlLeft
            .Range("f" & iRow_R).Value = "" & adoaccrpt409.Fields("r40906")
            .Range("f" & iRow_R).Font.Size = 11: .Range("f" & iRow_R).HorizontalAlignment = xlLeft
            .Range("g" & iRow_R).Value = "" & adoaccrpt409.Fields("r40910")
            .Range("g" & iRow_R).Font.Size = 11: .Range("g" & iRow_R).HorizontalAlignment = xlRight: .Range("g" & iRow_R).NumberFormatLocal = stCellFormat
            .Range("h" & iRow_R).Value = "" & adoaccrpt409.Fields("r40907")
            .Range("h" & iRow_R).Font.Size = 11: .Range("h" & iRow_R).HorizontalAlignment = xlRight: .Range("h" & iRow_R).NumberFormatLocal = stCellFormat
            'end 2018/12/28
            
            If InStr("" & adoaccrpt409.Fields("r40906"), "負債小計") > 0 Then intSum(0) = iRow_R
            If InStr("" & adoaccrpt409.Fields("r40906"), "股東權益小計") > 0 Then intSum(1) = iRow_R
            iRow_R = iRow_R + 1
         End If
        
         adoaccrpt409.MoveNext
      Loop
      'end 2016/08/10
      adoaccrpt409.Close
      'Mark by Amy 2016/08/10
'      iRow = iRow + 1
'      .Range("e" & iRow).Value = "*** 結束 ***"
      
      'Add by Amy 2016/08/10
      '總額
      .Range("b" & iRow_E).Value = strTotal1(0)
      .Range("d" & iRow_E).Value = strTotal1(1)
      'Add by Amy 2018/12/28 設格式
      .Range("d" & iRow_E).Font.Size = 11: .Range("d" & iRow_E).HorizontalAlignment = xlRight: .Range("d" & iRow_E).NumberFormatLocal = stCellFormat
      .Range("f" & iRow_E).Value = strTotal2(0)
      .Range("h" & iRow_E).Value = strTotal2(1)
      'Add by Amy 2018/12/28 設格式
      .Range("h" & iRow_E).Font.Size = 11: .Range("h" & iRow_E).HorizontalAlignment = xlRight: .Range("h" & iRow_E).NumberFormatLocal = stCellFormat
      '設定格線
      Call SetExcelLine(2, WksAccrpt409, "a" & intTitleRow + 1 & ":" & "h" & iRow_E + 1)
      Call SetExcelLine(1, WksAccrpt409, "a" & intTitleRow & ":" & "a" & iRow_E)
      Call SetExcelLine(1, WksAccrpt409, "e" & intTitleRow & ":" & "e" & iRow_E)
      '小計格線
      Call SetExcelLine(3, WksAccrpt409, "h" & intSum(0))
      Call SetExcelLine(3, WksAccrpt409, "h" & intSum(1))
      '總額格線
      Call SetExcelLine(4, WksAccrpt409, "d" & iRow_E)
      Call SetExcelLine(4, WksAccrpt409, "h" & iRow_E)
      
      .Range("a2:h" & IIf(iRow > iRow_R, iRow, iRow_R)).RowHeight = 21 'Add by Amy 2016/09/29 調高度-瑞婷
      
      'Modify by Amy 2015/05/21 原使用函數以為是抓A4紙張
      .PageSetup.PaperSize = 9 '設定紙張 A4
      .PageSetup.Orientation = xlPortrait 'Modfiy by Amy 2016/08/10 直印-婉莘 原:橫印 xlLandscape '
      .PageSetup.PrintTitleRows = "$1:$" & intTitleRow 'Modify by Amy 2016/08/10 改變數 原:表頭保留7列
      .PageSetup.PrintArea = "$A$1:$H$" & IIf(iRow > iRow_R, iRow, iRow_R) '設定列印範圍
      
      .PageSetup.TopMargin = xlsSalesPoint.InchesToPoints(0.3)
      'Modify by Amy 2016/09/29 改下、左、右邊界
      .PageSetup.BottomMargin = xlsSalesPoint.InchesToPoints(0.7)
      .PageSetup.LeftMargin = xlsSalesPoint.InchesToPoints(0.39) 'Modfiy by Amy 2016/08/10 左邊界左設2-婉莘 原:0.5
      .PageSetup.RightMargin = xlsSalesPoint.InchesToPoints(0) '右邊界
      'end 2016/09/29
      .PageSetup.HeaderMargin = xlsSalesPoint.InchesToPoints(0.5)
      .PageSetup.FooterMargin = xlsSalesPoint.InchesToPoints(0.5)
      
      .PageSetup.Zoom = 80 '縮放比例 100
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
   Set xlsSalesPoint = Nothing
   Set WksAccrpt409 = Nothing
   MsgBox "Excel檔案產生完成！（檔案位置：" & strFileName & "）"
   Exit Sub

ErrHnd:
   If Not xlsSalesPoint Is Nothing Then
        'Modify by Amy 2016/06/23 +判斷版本
        If Val(xlsSalesPoint.Version) < 12 Then
             xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
        Else
             xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
        End If
        xlsSalesPoint.Workbooks.Close
        'end 2016/06/23
      xlsSalesPoint.Quit
      Set xlsSalesPoint = Nothing
      Set WksAccrpt409 = Nothing
   End If
   MsgBox Err.Description
End Sub

'Add by Amy 2016/08/10 增加框線設定-婉莘
Private Sub SetExcelLine(intChoose As Integer, ByRef m_Xls As Worksheet, strField As String)
    With m_Xls.Range(strField)
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        Select Case intChoose
            Case 0 '抬頭
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlHairline
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin '實線
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlHairline
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
            Case 1 '會計科目
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThin
            Case 2 '資料內容
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlHairline
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlHairline
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlHairline
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).Weight = xlHairline
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlHairline
            Case 3 '小計
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
            Case 4 '總額
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlDouble
                .Borders(xlEdgeBottom).Weight = xlThick '雙線
        End Select
    End With
End Sub

'Mark by Amy 2020/05/29 改寫法不使用
''Add by Amy 2020/04/17
'Private Function CheckData() As Boolean
'    Dim RsQ As New ADODB.Recordset
'    Dim strQ As String
'
'    CheckData = False
'    '新增資產明細後新增分線符號,但其實無資料
'    strQ = "select * from accrpt409 where r40901='" & strUserNum & "' And r40902 is not null  order by r40908 asc"
'    RsQ.CursorLocation = adUseClient
'    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
'    If RsQ.RecordCount <= 0 Then
'        MsgBox "無資料產生!!"
'        Exit Function
'    End If
'    CheckData = True
'End Function


