VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmacc4320 
   AutoRedraw      =   -1  'True
   Caption         =   "過帳及分攤作業"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2490
   ScaleWidth      =   5160
   Begin VB.ComboBox CboComp 
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
      Text            =   "CboComp"
      Top             =   270
      Width           =   3500
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1890
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "過帳及分攤(&E)"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   1320
      Width           =   4452
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
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
      TabIndex        =   2
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
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
      Left            =   360
      TabIndex        =   6
      Top             =   270
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc4320"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 (無需修改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc010 As New ADODB.Recordset
Public adoacc021 As New ADODB.Recordset
Public adoacc021w As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoacc060 As New ADODB.Recordset
Public adoacc0b0 As New ADODB.Recordset
Public adoacc090 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset

'Add by Amy 2020/04/15
Private Sub CboComp_GotFocus()
    TextInverse CboComp
End Sub

Private Sub CboComp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboComp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboComp) = MsgText(601) Then
        MsgBox "請輸入公司別!!", , MsgText(5)
        Cancel = True
        CboComp.SetFocus
        Exit Sub
    End If
    strCmp = CboComp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label3 & MsgText(63), , MsgText(5)
        Cancel = True
        CboComp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboComp)) = 1 Then
        CboComp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/15

Private Sub Command1_Click()
   Dim strA0b05 As String
   'Modify by Amy 2017/09/22
   Dim bolAxbHasDt As Boolean
   Dim strAxb(1 To 17) As String, strMsg As String, bolHasActualP As Boolean 'Modify by Amy 2023/04/17
   
   strExc(0) = GetA0b01(strA0b05)
      
   'ADD BY SONIA  2014/7/7
   If Mid(MaskEdBox1, 1, 3) & "/" & Mid(MaskEdBox1, 5, 2) <> Mid(MaskEdBox2, 1, 3) & "/" & Mid(MaskEdBox2, 5, 2) Then
      MsgBox "不可跨月份過帳 !", , MsgText(5)
      Exit Sub
   End If
   'END 2014/7/7
   'Add by Amy 2016/02/15 +業績輸入需先關閉才可過帳
   If Val(strA0b05) < Val(Mid(FCDate(MaskEdBox1), 1, 5)) Then
      MsgBox "請先執行「智權點數實績與結餘輸入」關閉 !", , MsgText(5)
      Exit Sub
   End If
   'Add by Amy 2017/04/28 +1公司檢查axb02/03
   'Modify by Amy 2017/09/22 加判斷SalesPoint 轉撥總經理欄修改日>傳票修改日不可過帳(實績、結餘分開判斷)
   'Modify by Amy 2020/04/15 公司別改下拉 原:Text3
   If Left(Trim(CboComp), 1) = "1" Then
        bolAxbHasDt = bolAcc0b1(0, Mid(FCDate(MaskEdBox1), 1, 5), strAxb())
        If bolAxbHasDt = False Then
           MsgBox "實績、結餘期末傳票尚未產生,不可過帳 !", , MsgText(5)
           Exit Sub
        ElseIf strAxb(4) = MsgText(601) Then
           MsgBox "實績期末傳票尚未處理,不可過帳 !", , MsgText(5)
           Exit Sub
        ElseIf strAxb(9) = MsgText(601) Then
           MsgBox "結餘期末傳票尚未處理,不可過帳 !", , MsgText(5)
           Exit Sub
        End If
        '判斷SalesPoint 總經理欄有修改,對映傳票是否有更正
        If bolAutoVoucherNS(1, Mid(FCDate(MaskEdBox1), 1, 5), strAxb(4), strAxb(5)) = True Then
           MsgBox "實績輸入有修改,但傳票尚未更正,不可過帳 !", , MsgText(5)
           Exit Sub
        End If
        'Add by Amy 2023/04/17
        '實績期未保留傳票已產生,當月實績傳票有修改,需彈訊息,不可過帳
        strMsg = ""
        bolHasActualP = HasActualP(2, Round(Val(FCDate(MaskEdBox1.Text)) / 100, 0)) '實績保留
        If bolHasActualP = True Then
            MsgBox "每月點數開放後「當月實績」有修改，" & vbCrLf & _
                          "請至「智權期末實績保留傳票產生」更正傳票", , MsgText(5)
            Exit Sub
        End If
        '判斷是否有ACS收入傳票,是否已產生實績期未保留傳票
        strMsg = ""
        If ChkACSIncomeAndEndAmt(Me.Name, Mid(FCDate(MaskEdBox1), 1, 5), strAxb(17), strMsg) = False Then
            MsgBox strMsg, , MsgText(5)
            Exit Sub
        End If
        'end 2023/04/17
        If bolAutoVoucherNS(3, Mid(FCDate(MaskEdBox1), 1, 5), strAxb(6)) = True Then
           MsgBox "實績轉撥輸入有修改,但傳票尚未更正,不可過帳 !", , MsgText(5)
           Exit Sub
        End If
        If bolAutoVoucherNS(2, Mid(FCDate(MaskEdBox1), 1, 5), strAxb(9), strAxb(10)) = True Then
           MsgBox "結餘輸入有修改,但傳票尚未更正,不可過帳 !", , MsgText(5)
           Exit Sub
        End If
        If bolAutoVoucherNS(4, Mid(FCDate(MaskEdBox1), 1, 5), strAxb(11)) = True Then
           MsgBox "結餘轉撥輸入有修改,但傳票尚未更正,不可過帳 !", , MsgText(5)
           Exit Sub
        End If
        'Add by Amy 2022/11/07 每月點數開放後有修改「結餘」,尚未更正,不可過帳
        'Modify by Amy 2023/04/17 原: strAxb16(0) 不使用
        If strAxb(16) = "Y" Then '結餘保留
            MsgBox "每月點數開放後「結餘傳票資料」有修改" & vbCrLf & _
                          "請確認是否刪除「智權期末結餘保留資料」" & vbCrLf & _
                           "若需刪除請至「智權期末結餘保留資料刪除」作業刪除" & vbCrLf & _
                           "刪除後,請至「智權期未結餘保留」更正傳票", , MsgText(5)
            Exit Sub
        End If
        'end 2022/11/07
   End If
   'end 2017/09/22
   'end 2017/04/28
   Screen.MousePointer = vbHourglass
   CalculateTable
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5280
   Me.Height = 2895
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath3)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2020/04/15 公司別改下拉
   CboComp.Clear
   Call Pub_SetCboCmp(CboComp, False, False, False, , 0)
   'end 2020/04/15
   'MaskEdBox1.Text = Mid(ACDate(ServerDate), 1, 3) & "/" & Mid(MsgText(16), 2, 2) & "/" & Mid(MsgText(16), 2, 2)
   MaskEdBox1.Text = TransDate(CompDate(1, -1, (Left(strSrvDate(1), 6) & "01")), 1)   '預設上月1日
   MaskEdBox1.Text = Mid(MaskEdBox1.Text, 1, 3) & "/" & Mid(MaskEdBox1.Text, 4, 2) & "/" & Mid(MaskEdBox1.Text, 6, 2)
   MaskEdBox1.Mask = DFormat
   'MaskEdBox2.Text = Mid(ACDate(ServerDate), 1, 3) & "/" & Mid(ACDate(ServerDate), 4, 2) & "/" & Mid(ACDate(ServerDate), 6, 2)
   MaskEdBox2.Text = TransDate(CompDate(2, -1, (Left(strSrvDate(1), 6) & "01")), 1) '預設上月最後一日
   MaskEdBox2.Text = Mid(MaskEdBox2.Text, 1, 3) & "/" & Mid(MaskEdBox2.Text, 4, 2) & "/" & Mid(MaskEdBox2.Text, 6, 2)
   MaskEdBox2.Mask = DFormat
   OpenTable
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Call PUB_GetLock("", "Frmacc4320")  'add by sonia 2014/8/8
   Set Frmacc4320 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   '2014/1/22 cancel by sonia 移到CalculateTable再依公司別讀檔
   'adoacc0b0.CursorLocation = adUseClient
   'adoacc0b0.Open "select * from acc0b0", adoTaie, adOpenDynamic, adLockBatchOptimistic
   '2014/1/22 end
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  傳票過帳及分攤計算
'
'*************************************************
Private Sub CalculateTable()
Dim intYear As Integer, intStartMonth As Integer, intEndMonth As Integer
Dim douAmount As Double, douDAmount As Double, douCAmount As Double, douTotalAmount As Double
Dim strQ As String, strCmp As String 'Add by Amy 2020/04/15

On Error GoTo Checking
   ProgressBar1.Value = 0
   StatusView MsgText(26)
   'Modify by Amy 2020/04/15 公司別改下拉 原:IIf(Text3 = "2", "J", Text3)
   strCmp = CboComp
   If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
   End If
   strQ = "Select * from acc0b0 where a0b04 = '" & strCmp & "'"
   '2014/1/22 add by sonia
   If adoacc0b0.State = adStateOpen Then adoacc0b0.Close
   adoacc0b0.CursorLocation = adUseClient
   adoacc0b0.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
   'end 2020/04/14
   '2014/1/22 end
   If adoacc0b0.RecordCount <> 0 Then
      If IsNull(adoacc0b0.Fields("a0b01").Value) = False Then
         If Mid(CFDate(adoacc0b0.Fields("a0b01").Value), 1, 6) < Mid(MaskEdBox1.Text, 1, 6) Then
             ' add by sonia 2014/7/7 上月未月結, 本月不可過帳
             'modify by sonia 2016/3/9
             'If Val(adoacc0b0.Fields("a0b02").Value) <> Val(adoacc0b0.Fields("a0b01").Value) Then
             If Left(Val(adoacc0b0.Fields("a0b02").Value) + 19110000, 6) <> Left(Val(adoacc0b0.Fields("a0b01").Value) + 19110000, 6) Then
                MsgBox "上月月結算尚未作業, 本月不可執行過帳作業 ! ", , MsgText(5)
                adoacc0b0.Close
                Exit Sub
             End If
            'end 2014/7/7
            Call AccountSave(strCmp) 'Modify by Amy 2020/04/15 +公司別參數
         End If
      Else
         Call AccountSave(strCmp)  'Modify by Amy 2020/04/15 +公司別參數
      End If
   Else
      Call AccountSave(strCmp)  'Modify by Amy 2020/04/15 +公司別參數
   End If
   adoacc0b0.Close
'   AccountSave
   ProgressBar1.Value = 0
   If adoacc021.State = adStateOpen Then adoacc021.Close
   adoacc021.CursorLocation = adUseClient
   '2014/1/22 modify by sonia 加入公司別條件
   'adoacc021.Open "select * from acc021, acc020 where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & " and ax210 is null order by ax201 asc, ax202 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2020/04/15 公司別改下拉 原:IIf(Text3 = "2", "J", Text3)
   strQ = "select * from acc021, acc020 where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and a0201 = '" & strCmp & "' and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & " and ax210 is null order by ax201 asc, ax202 asc"
   adoacc021.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   'end 2020/04/15
   If adoacc021.RecordCount = 0 Then
      StatusClear
      MsgBox MsgText(15), , MsgText(21)
      adoacc021.Close
      Exit Sub
   Else
      ProgressBar1.max = adoacc021.RecordCount
   End If
   Do While adoacc021.EOF = False
   '部門過帳
      If adoacc040.State = adStateOpen Then adoacc040.Close
      adoacc040.CursorLocation = adUseClient
      adoacc040.Open "select * from acc040 where a0401 = " & Val(Mid(CFDate(adoacc021.Fields("a0205").Value), 1, 3)) & " " & _
                     "and a0402 = " & Val(Mid(CFDate(adoacc021.Fields("a0205").Value), 5, 2)) & " and a0403 = '" & adoacc021.Fields("ax201").Value & "' and " & _
                     "a0404 = '" & IIf(IsNull(adoacc021.Fields("ax204").Value), MsgText(55), adoacc021.Fields("ax204").Value) & "' and a0405 = '" & adoacc021.Fields("ax205").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoacc040.RecordCount = 0 Then
         adoacc040.AddNew
         Acc040Save
         If IsNull(adoacc021.Fields("AX204").Value) = False And adoacc021.Fields("ax204").Value <> MsgText(55) Then
            adoacc040.Fields("a0404").Value = adoacc021.Fields("ax204").Value
         Else
            adoacc040.Fields("A0404").Value = MsgText(55)
         End If
      Else
         If IsNull(adoacc021.Fields("AX204").Value) = False And adoacc021.Fields("ax204").Value <> MsgText(55) Then
            adoacc040.Fields("a0404").Value = adoacc021.Fields("ax204").Value
         Else
            adoacc040.Fields("A0404").Value = MsgText(55)
         End If
      End If
      adoacc040.Fields("a0406").Value = Val(adoacc040.Fields("a0406").Value) + Val(adoacc021.Fields("ax206").Value)
      adoacc040.Fields("a0407").Value = Val(adoacc040.Fields("a0407").Value) + Val(adoacc021.Fields("ax207").Value)
      adoacc040.UpdateBatch
      adoacc040.Close
      
   '總所過帳
      If adoacc021.Fields("ax204").Value <> MsgText(55) And IsNull(adoacc021.Fields("ax204").Value) = False Then
         If adoacc040.State = adStateOpen Then adoacc040.Close
         adoacc040.CursorLocation = adUseClient
         adoacc040.Open "select * from acc040 where a0401 = " & Val(Mid(CFDate(adoacc021.Fields("a0205").Value), 1, 3)) & " " & _
                        "and a0402 = " & Val(Mid(CFDate(adoacc021.Fields("a0205").Value), 5, 2)) & " and a0403 = '" & adoacc021.Fields("ax201").Value & "' and " & _
                        "a0404 = '" & MsgText(55) & "' and a0405 = '" & adoacc021.Fields("ax205").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
         If adoacc040.RecordCount = 0 Then
            adoacc040.AddNew
            Acc040Save
            adoacc040.Fields("A0404").Value = "TOT"
         End If
         adoacc040.Fields("a0406").Value = Val(adoacc040.Fields("a0406").Value) + Val(adoacc021.Fields("ax206").Value)
         adoacc040.Fields("a0407").Value = Val(adoacc040.Fields("a0407").Value) + Val(adoacc021.Fields("ax207").Value)
         adoacc040.UpdateBatch
         adoacc040.Close
      End If
      
   '第一次分攤
   '部門過帳
      If adoacc021.Fields("ax204").Value = MsgText(55) Or IsNull(adoacc021.Fields("ax204").Value) Then
         If adoacc060.State = adStateOpen Then adoacc060.Close
         adoacc060.CursorLocation = adUseClient
         adoacc060.Open "select * from acc060, acc010 where a0602 = a0105 and a0601 = " & Val(Mid(CFDate(adoacc021.Fields("a0205").Value), 1, 3)) & " and a0603 = '" & adoacc021.Fields("ax201").Value & "' and a0101 = '" & adoacc021.Fields("ax205").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         Do While adoacc060.EOF = False
            If adoacc040.State = adStateOpen Then adoacc040.Close
            adoacc040.CursorLocation = adUseClient
            adoacc040.Open "select * from acc040 where a0401 = " & Val(Mid(CFDate(adoacc021.Fields("a0205").Value), 1, 3)) & " " & _
                           "and a0402 = " & Val(Mid(CFDate(adoacc021.Fields("a0205").Value), 5, 2)) & " and a0403 = '" & adoacc021.Fields("ax201").Value & "' and " & _
                           "a0404 = '" & adoacc060.Fields("a0604").Value & "' and a0405 = '" & adoacc021.Fields("ax205").Value & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
            If adoacc040.RecordCount = 0 Then
               adoacc040.AddNew
               Acc040Save
               adoacc040.Fields("a0404").Value = adoacc060.Fields("a0604").Value
            End If
            adoacc040.Fields("a0406").Value = Val(adoacc040.Fields("a0406").Value) + Val(Format(Val(adoacc021.Fields("ax206").Value) * Val(adoacc060.Fields("a0605").Value) / 100, FAmount))
            adoacc040.Fields("a0407").Value = Val(adoacc040.Fields("a0407").Value) + Val(Format(Val(adoacc021.Fields("ax207").Value) * Val(adoacc060.Fields("a0605").Value) / 100, FAmount))
            adoacc040.UpdateBatch
            adoacc040.Close
            adoacc060.MoveNext
         Loop
         adoacc060.Close
      End If
         
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoTaie.Execute "update acc021 set ax210 = " & Val(ACDate(ServerDate())) & " where ax201 = '" & adoacc021.Fields("ax201").Value & "' and ax202 = '" & adoacc021.Fields("ax202").Value & "' and ax203 = '" & adoacc021.Fields("ax203").Value & "'"
      adoacc021.MoveNext
   Loop
   adoacc021.Close
   
' 計算資產負債之累計
   If adoacc040.State = adStateOpen Then adoacc040.Close
   adoacc040.CursorLocation = adUseClient
   'Modify by Amy 2020/04/15 公司別改下拉 原:IIf(Text3 = "2", "J", Text3)
   strQ = "select * from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0402 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & " and a0403 = '" & strCmp & "' and substr(a0405, 1, 1) in ('1', '2', '3') order by a0401 asc, a0402 asc"
   adoacc040.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
   'end 2020/04/15
   Do While adoacc040.EOF = False
      If adoacc040.Fields("a0405").Value = "3222" Then
         adoacc040.Fields("a0408").Value = Val(adoacc040.Fields("a0407").Value) - Val(adoacc040.Fields("a0406").Value)
      Else
         If Mid(adoacc040.Fields("a0405").Value, 1, 1) = "1" Then
            'Modify by Morgan 2005/10/17會有小數點後三位被捨去問題(Ex.81772.8699999996->81772.86)
            'adoacc040.Fields("a0408").Value = Val(adoacc040.Fields("a0406").Value) - Val(adoacc040.Fields("a0407").Value) + GetLastMonthBalance(Val(adoacc040.Fields("a0401").Value), Val(adoacc040.Fields("a0402").Value), adoacc040.Fields("a0403").Value, adoacc040.Fields("a0404").Value, adoacc040.Fields("a0405").Value)
            adoacc040.Fields("a0408").Value = Format(Val(adoacc040.Fields("a0406").Value) - Val(adoacc040.Fields("a0407").Value) + GetLastMonthBalance(Val(adoacc040.Fields("a0401").Value), Val(adoacc040.Fields("a0402").Value), adoacc040.Fields("a0403").Value, adoacc040.Fields("a0404").Value, adoacc040.Fields("a0405").Value), "0.00")
         Else
            'Modify by Morgan 2005/10/17會有小數點後三位被捨去問題
            'adoacc040.Fields("a0408").Value = Val(adoacc040.Fields("a0407").Value) - Val(adoacc040.Fields("a0406").Value) + GetLastMonthBalance(Val(adoacc040.Fields("a0401").Value), Val(adoacc040.Fields("a0402").Value), adoacc040.Fields("a0403").Value, adoacc040.Fields("a0404").Value, adoacc040.Fields("a0405").Value)
            adoacc040.Fields("a0408").Value = Format(Val(adoacc040.Fields("a0407").Value) - Val(adoacc040.Fields("a0406").Value) + GetLastMonthBalance(Val(adoacc040.Fields("a0401").Value), Val(adoacc040.Fields("a0402").Value), adoacc040.Fields("a0403").Value, adoacc040.Fields("a0404").Value, adoacc040.Fields("a0405").Value), "0.00")
         End If
      End If
      adoacc040.UpdateBatch
      adoacc040.MoveNext
   Loop
   adoacc040.Close
   
' 計算收入費用餘額
  If adoacc040.State = adStateOpen Then adoacc040.Close
   adoacc040.CursorLocation = adUseClient
   'Modify by Amy 2020/04/15 公司別改下拉 原:IIf(Text3 = "2", "J", Text3)
   strQ = "select * from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0402 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & " and a0403 = '" & strCmp & "' and substr(a0405, 1, 1) not in ('1', '2', '3') order by a0401 asc, a0402 asc"
   adoacc040.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
   'end 2020/04/15
   Do While adoacc040.EOF = False
      If Mid(adoacc040.Fields("a0405").Value, 1, 1) <> "4" And Mid(adoacc040.Fields("a0405").Value, 1, 2) <> "71" Then
         'Modify by Morgan 2006/4/4會有小數點後三位被捨去問題
         'adoacc040.Fields("a0408").Value = Val(adoacc040.Fields("a0406").Value) - Val(adoacc040.Fields("a0407").Value)
         adoacc040.Fields("a0408").Value = Format(Val(adoacc040.Fields("a0406").Value) - Val(adoacc040.Fields("a0407").Value), "0.00")
      Else
         'Modify by Morgan 2006/4/4會有小數點後三位被捨去問題
         'adoacc040.Fields("a0408").Value = Val(adoacc040.Fields("a0407").Value) - Val(adoacc040.Fields("a0406").Value)
         adoacc040.Fields("a0408").Value = Format(Val(adoacc040.Fields("a0407").Value) - Val(adoacc040.Fields("a0406").Value), "0.00")
      End If
      adoacc040.UpdateBatch
      adoacc040.MoveNext
   Loop
   adoacc040.Close
   
' 第二次分攤
   'add by sonia 2016/1/28 105年起加9997分攤法務部門費用
   'Modify by Amy 2020/04/15 公司別改抓變數
   If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
      adoTaie.Execute "delete from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0402 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & " and a0403 = '" & strCmp & "' and a0405 in ('9997', '9998', '9999')"
   Else
   'end 2016/1/28
      adoTaie.Execute "delete from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0402 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & " and a0403 = '" & strCmp & "' and a0405 in ('9998', '9999')"
   End If
   'end 2020/04/15
   
   'Mark by Amy 2020/08/12 不使用
'   'add by sonia 2016/1/28
'' 計算法務部門分攤之費用
'   If Val(Mid(MaskEdBox1.Text, 1, 3)) >= 105 Then
'      If adoacc060.State = adStateOpen Then adoacc060.Close
'      adoacc060.CursorLocation = adUseClient
'      'Modify by Amy 2020/04/15 公司別改下拉 原:IIf(Text3 = "2", "J", Text3)
'      strQ = "select a0401, a0402, a0403, sum(a0406) as Debit, sum(a0407) as Credit, sum(a0408) as Total from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0402 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & " and a0403 = '" & strCmp & "' and a0404 = 'L' and substr(a0405, 1, 1) = '6' group by a0401, a0402, a0403"
'      adoacc060.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
'      'end 2020/04/15
'      Do While adoacc060.EOF = False
'         douTotalAmount = 0
'         If adoacc090.State = adStateOpen Then adoacc090.Close
'         adoacc090.CursorLocation = adUseClient
'         adoacc090.Open "select a0901 from acc090 where a0904 = '" & MsgText(602) & "' and a0901 not in ('M', 'SAL', 'TOT', 'L', 'CFL', 'FCL') order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
'         Do While adoacc090.EOF = False
'            adoacc040.CursorLocation = adUseClient
'            adoacc040.Open "select * from acc040 where a0401 = " & Val(adoacc060.Fields("a0401").Value) & " " & _
'                              "and a0402 = " & Val(adoacc060.Fields("a0402").Value) & " and a0403 = '" & adoacc060.Fields("a0403").Value & "' and " & _
'                                  "a0404 = '" & adoacc090.Fields("a0901").Value & "' and a0405 = '9997'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'            If adoacc040.RecordCount = 0 Then
'               adoacc040.AddNew
'               adoacc040.Fields("a0401").Value = Val(adoacc060.Fields("a0401").Value)
'               adoacc040.Fields("a0402").Value = Val(adoacc060.Fields("a0402").Value)
'               adoacc040.Fields("a0403").Value = adoacc060.Fields("a0403").Value
'               adoacc040.Fields("a0404").Value = adoacc090.Fields("a0901").Value
'               adoacc040.Fields("a0405").Value = "9997"
'               adoacc040.Fields("a0406").Value = 0
'               adoacc040.Fields("a0407").Value = 0
'               adoacc040.Fields("a0408").Value = 0
'               adoacc040.Fields("a0409").Value = 0
'            End If
'            douAmount = GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "P") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "T") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "CFP") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "CFT") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "FCP") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "FCT")
'            If douAmount <> 0 Then
'               adoacc040.Fields("a0406").Value = Val(adoacc040.Fields("a0406").Value) + Val(Format(Val(adoacc060.Fields("Debit").Value) * Val(GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, adoacc090.Fields("a0901").Value)) / douAmount, FAmount))
'               adoacc040.Fields("a0407").Value = Val(adoacc040.Fields("a0407").Value) + Val(Format(Val(adoacc060.Fields("Credit").Value) * Val(GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, adoacc090.Fields("a0901").Value)) / douAmount, FAmount))
'               adoacc040.Fields("a0408").Value = Val(adoacc040.Fields("a0408").Value) + Val(Format(Val(adoacc060.Fields("Total").Value) * Val(GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, adoacc090.Fields("a0901").Value)) / douAmount, FAmount))
'            End If
'            douTotalAmount = douTotalAmount + Val(adoacc040.Fields("a0408").Value)
'            adoacc040.UpdateBatch
'            adoacc040.Close
'            adoacc090.MoveNext
'         Loop
'         adoacc090.Close
'         adoTaie.Execute "update acc040 set a0408 = " & Val(adoacc060.Fields("Total").Value) - douTotalAmount & " + a0408 where a0401 = " & Val(adoacc060.Fields("a0401").Value) & " and a0402 = " & Val(adoacc060.Fields("a0402").Value) & " and a0403 = '" & adoacc060.Fields("a0403").Value & "' and a0404 = 'P' and a0405 = '9997'"
'         adoacc060.MoveNext
'      Loop
'      adoacc060.Close
'   End If
'   'end 2016/1/28
   'end 2020/08/12

' 計算管理部門分攤之費用
   If adoacc060.State = adStateOpen Then adoacc060.Close
   adoacc060.CursorLocation = adUseClient
   'Modify by Amy 2020/04/15 公司別改下拉 原:IIf(Text3 = "2", "J", Text3)
   strQ = "select a0401, a0402, a0403, sum(a0406) as Debit, sum(a0407) as Credit, sum(a0408) as Total from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0402 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & " and a0403 = '" & strCmp & "' and a0404 = 'M' and substr(a0405, 1, 1) = '6' group by a0401, a0402, a0403"
   adoacc060.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   'end 2020/04/15
   Do While adoacc060.EOF = False
      douTotalAmount = 0
      If adoacc090.State = adStateOpen Then adoacc090.Close
      adoacc090.CursorLocation = adUseClient
      'modify by sonia 2016/1/28 105年起不含L故加入'L',且CFL,FCL部門已無傳票
      adoacc090.Open "select a0901 from acc090 where a0904 = '" & MsgText(602) & "' and a0901 not in ('M', 'SAL', 'TOT', 'L', 'CFL', 'FCL') order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
      Do While adoacc090.EOF = False
         If adoacc040.State = adStateOpen Then adoacc040.Close
         adoacc040.CursorLocation = adUseClient
         adoacc040.Open "select * from acc040 where a0401 = " & Val(adoacc060.Fields("a0401").Value) & " " & _
                           "and a0402 = " & Val(adoacc060.Fields("a0402").Value) & " and a0403 = '" & adoacc060.Fields("a0403").Value & "' and " & _
                                "a0404 = '" & adoacc090.Fields("a0901").Value & "' and a0405 = '9998'", adoTaie, adOpenDynamic, adLockBatchOptimistic
         If adoacc040.RecordCount = 0 Then
            adoacc040.AddNew
            adoacc040.Fields("a0401").Value = Val(adoacc060.Fields("a0401").Value)
            adoacc040.Fields("a0402").Value = Val(adoacc060.Fields("a0402").Value)
            adoacc040.Fields("a0403").Value = adoacc060.Fields("a0403").Value
            adoacc040.Fields("a0404").Value = adoacc090.Fields("a0901").Value
            adoacc040.Fields("a0405").Value = "9998"
            adoacc040.Fields("a0406").Value = 0
            adoacc040.Fields("a0407").Value = 0
            adoacc040.Fields("a0408").Value = 0
            adoacc040.Fields("a0409").Value = 0
         End If
         'modify by sonia 2016/1/28 105年起不含L故取消L,CFL,FCL,且CFL,FCL部門已無傳票
         douAmount = GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "P") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "T") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "CFP") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "CFT") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "FCP") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "FCT")
         'douAmount = douAmount + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "FCL") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "CFL")
         If douAmount <> 0 Then
            adoacc040.Fields("a0406").Value = Val(adoacc040.Fields("a0406").Value) + Val(Format(Val(adoacc060.Fields("Debit").Value) * Val(GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, adoacc090.Fields("a0901").Value)) / douAmount, FAmount))
            adoacc040.Fields("a0407").Value = Val(adoacc040.Fields("a0407").Value) + Val(Format(Val(adoacc060.Fields("Credit").Value) * Val(GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, adoacc090.Fields("a0901").Value)) / douAmount, FAmount))
            adoacc040.Fields("a0408").Value = Val(adoacc040.Fields("a0408").Value) + Val(Format(Val(adoacc060.Fields("Total").Value) * Val(GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, adoacc090.Fields("a0901").Value)) / douAmount, FAmount))
         End If
         douTotalAmount = douTotalAmount + Val(adoacc040.Fields("a0408").Value)
         adoacc040.UpdateBatch
         adoacc040.Close
         adoacc090.MoveNext
      Loop
      adoacc090.Close
      adoTaie.Execute "update acc040 set a0408 = " & Val(adoacc060.Fields("Total").Value) - douTotalAmount & " + a0408 where a0401 = " & Val(adoacc060.Fields("a0401").Value) & " and a0402 = " & Val(adoacc060.Fields("a0402").Value) & " and a0403 = '" & adoacc060.Fields("a0403").Value & "' and a0404 = 'P' and a0405 = '9998'"
      adoacc060.MoveNext
   Loop
   adoacc060.Close
      
' 計算智權分攤之費用
  If adoacc060.State = adStateOpen Then adoacc060.Close
   adoacc060.CursorLocation = adUseClient
   'Modify by Amy 2020/04/15 公司別改下拉 原:IIf(Text3 = "2", "J", Text3)
   strQ = "select a0401, a0402, a0403, sum(a0406) as Debit, sum(a0407) as Credit, sum(a0408) as Total from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0402 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & " and a0403 = '" & strCmp & "' and a0404 = 'SAL' and substr(a0405, 1, 1) = '6' group by a0401, a0402, a0403"
   'end 2020/04/15
   adoacc060.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc060.EOF = False
      douTotalAmount = 0
      If adoacc090.State = adStateOpen Then adoacc090.Close
      adoacc090.CursorLocation = adUseClient
      'modify by sonia 2016/1/28 105年起不含L故加入'L',且CFL,FCL部門已無傳票
      adoacc090.Open "select a0901 from acc090 where a0904 = '" & MsgText(602) & "' and a0901 not in ('M', 'SAL', 'TOT', 'FCP', 'FCT', 'CFL', 'FCL', 'L') order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
      Do While adoacc090.EOF = False
         adoacc040.CursorLocation = adUseClient
         adoacc040.Open "select * from acc040 where a0401 = " & Val(adoacc060.Fields("a0401").Value) & " " & _
                           "and a0402 = " & Val(adoacc060.Fields("a0402").Value) & " and a0403 = '" & adoacc060.Fields("a0403").Value & "' and " & _
                               "a0404 = '" & adoacc090.Fields("a0901").Value & "' and a0405 = '9999'", adoTaie, adOpenDynamic, adLockBatchOptimistic
         If adoacc040.RecordCount = 0 Then
            adoacc040.AddNew
            adoacc040.Fields("a0401").Value = Val(adoacc060.Fields("a0401").Value)
            adoacc040.Fields("a0402").Value = Val(adoacc060.Fields("a0402").Value)
            adoacc040.Fields("a0403").Value = adoacc060.Fields("a0403").Value
            adoacc040.Fields("a0404").Value = adoacc090.Fields("a0901").Value
            adoacc040.Fields("a0405").Value = "9999"
            adoacc040.Fields("a0406").Value = 0
            adoacc040.Fields("a0407").Value = 0
            adoacc040.Fields("a0408").Value = 0
            adoacc040.Fields("a0409").Value = 0
         End If
         'modify by sonia 2016/1/28 105年起不含L故取消L
         douAmount = GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "P") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "T") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "CFP") + GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, "CFT")
         If douAmount <> 0 Then
            adoacc040.Fields("a0406").Value = Val(adoacc040.Fields("a0406").Value) + Val(Format(Val(adoacc060.Fields("Debit").Value) * GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, adoacc090.Fields("a0901").Value) / douAmount, FAmount))
            adoacc040.Fields("a0407").Value = Val(adoacc040.Fields("a0407").Value) + Val(Format(Val(adoacc060.Fields("Credit").Value) * GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, adoacc090.Fields("a0901").Value) / douAmount, FAmount))
            adoacc040.Fields("a0408").Value = Val(adoacc040.Fields("a0408").Value) + Val(Format(Val(adoacc060.Fields("Total").Value) * GetMonthIncome(adoacc060.Fields("a0401").Value, adoacc060.Fields("a0402").Value, adoacc060.Fields("a0403").Value, adoacc090.Fields("a0901").Value) / douAmount, FAmount))
         End If
         douTotalAmount = douTotalAmount + Val(adoacc040.Fields("a0408").Value)
         adoacc040.UpdateBatch
         adoacc040.Close
         adoacc090.MoveNext
      Loop
      adoacc090.Close
      adoTaie.Execute "update acc040 set a0408 = " & Val(adoacc060.Fields("Total").Value) - douTotalAmount & " + a0408 where a0401 = " & Val(adoacc060.Fields("a0401").Value) & " and a0402 = " & Val(adoacc060.Fields("a0402").Value) & " and a0403 = '" & adoacc060.Fields("a0403").Value & "' and a0404 = 'P' and a0405 = '9999'"
      adoacc060.MoveNext
   Loop
   adoacc060.Close
   
' 計算管理部門分攤之費用
'   adoacc040.CursorLocation = adUseClient
'   adoacc040.Open "select a0401, a0402, a0403, sum(a0408) from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0402 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & " and a0403 = '1' and a0404 = 'M' and substr(a0405, 1, 1) = '6' group by a0401, a0402, a0403", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoacc040.EOF = False
'      adoacc090.CursorLocation = adUseClient
'      adoacc090.Open "select * from acc090 where a0904 = 'Y' and a0901 not in ('TOT', 'M', 'SAL') order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
'      Do While adoacc090.EOF = False
'         adoTaie.Execute "delete from acc040 where a0401 = " & Val(adoacc040.Fields("a0401").Value) & " and a0402 = " & Val(adoacc040.Fields("a0402").Value) & " and a0403 = '" & adoacc040.Fields("a0403").Value & "' and a0404 = '" & adoacc090.Fields("a0901").Value & "' and a0405 = '9998'"
'         douAmount = Val(Format(Val(adoacc040.Fields(3).Value) * DeptPercentM(Val(adoacc040.Fields("a0401").Value), Val(adoacc040.Fields("a0402").Value), adoacc040.Fields("a0403").Value, adoacc090.Fields("a0901").Value) / 100, FAmount))
'         adoTaie.Execute "insert into acc040 values (" & Val(adoacc040.Fields("a0401").Value) & ", " & Val(adoacc040.Fields("a0402").Value) & ", '" & adoacc040.Fields("a0403").Value & "', '" & adoacc090.Fields("a0901").Value & "', '9998', 0, 0, " & douAmount & ", 0, '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", null, null, null)"
'         adoacc090.MoveNext
'      Loop
'      adoacc090.Close
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select sum(a0408) from acc040 where a0401 = " & Val(adoacc040.Fields("a0401").Value) & " and a0402 = " & Val(adoacc040.Fields("a0402").Value) & " and a0403 = '" & adoacc040.Fields("a0403").Value & "' and a0405 = '9998' and a0404 not in ('TOT', 'M', 'SAL', 'P')", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         adoTaie.Execute "update acc040 set a0408 = " & Val(Format(Val(adoacc040.Fields(3).Value) - Val(IIf(IsNull(adoquery.Fields(0).Value), 0, adoquery.Fields(0).Value)), FAmount)) & " where a0401 = " & Val(adoacc040.Fields("a0401").Value) & " and a0402 = " & Val(adoacc040.Fields("a0402").Value) & " and a0403 = '" & adoacc040.Fields("a0403").Value & "' and a0404 = 'P' and a0405 = '9998'"
'      End If
'      adoquery.Close
'      adoTaie.Execute "delete from acc040 where a0401 = " & Val(adoacc040.Fields("a0401").Value) & " and a0402 = " & Val(adoacc040.Fields("a0402").Value) & " and a0403 = '" & adoacc040.Fields("a0403").Value & "' and a0404 = 'M' and a0405 = '9998'"
'      adoTaie.Execute "insert into acc040 values (" & Val(adoacc040.Fields("a0401").Value) & ", " & Val(adoacc040.Fields("a0402").Value) & ", '" & adoacc040.Fields("a0403").Value & "', 'M', '9998', 0, 0, " & IIf(IsNull(adoacc040.Fields(3).Value), 0, Val(adoacc040.Fields(3).Value) * -1) & ", 0, '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", null, null, null)"
'      adoacc040.MoveNext
'   Loop
'   adoacc040.Close
   
' 計算智權部門分攤之費用
'   adoacc040.CursorLocation = adUseClient
'   adoacc040.Open "select a0401, a0402, a0403, sum(a0408) from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 >= " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0402 <= " & Val(Mid(MaskEdBox2.Text, 5, 2)) & " and a0403 = '1' and a0404 = 'SAL' and substr(a0405, 1, 1) = '6' group by a0401, a0402, a0403", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoacc040.EOF = False
'      adoacc090.CursorLocation = adUseClient
'      adoacc090.Open "select * from acc090 where a0904 = 'Y' and a0901 not in ('TOT', 'M', 'SAL', 'FCL', 'FCP', 'FCT', 'FL') order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
'      Do While adoacc090.EOF = False
'         adoTaie.Execute "delete from acc040 where a0401 = " & Val(adoacc040.Fields("a0401").Value) & " and a0402 = " & Val(adoacc040.Fields("a0402").Value) & " and a0403 = '" & adoacc040.Fields("a0403").Value & "' and a0404 = '" & adoacc090.Fields("a0901").Value & "' and a0405 = '9999'"
'         douAmount = Val(Format(Val(adoacc040.Fields(3).Value) * DeptPercentS(Val(adoacc040.Fields("a0401").Value), Val(adoacc040.Fields("a0402").Value), adoacc040.Fields("a0403").Value, adoacc090.Fields("a0901").Value) / 100, FAmount))
'         adoTaie.Execute "insert into acc040 values (" & Val(adoacc040.Fields("a0401").Value) & ", " & Val(adoacc040.Fields("a0402").Value) & ", '" & adoacc040.Fields("a0403").Value & "', '" & adoacc090.Fields("a0901").Value & "', '9999', 0, 0, " & douAmount & ", 0, '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", null, null, null)"
'         adoacc090.MoveNext
'      Loop
'      adoacc090.Close
'      adoquery.CursorLocation = adUseClient
'      adoquery.Open "select sum(a0408) from acc040 where a0401 = " & Val(adoacc040.Fields("a0401").Value) & " and a0402 = " & Val(adoacc040.Fields("a0402").Value) & " and a0403 = '" & adoacc040.Fields("a0403").Value & "' and a0405 = '9999' and a0404 not in ('TOT', 'M', 'SAL', 'P', 'FCL', 'FCP', 'FCT', 'FL')", adoTaie, adOpenStatic, adLockReadOnly
'      If adoquery.RecordCount <> 0 Then
'         adoTaie.Execute "update acc040 set a0408 = " & Val(Format(Val(adoacc040.Fields(3).Value) - Val(IIf(IsNull(adoquery.Fields(0).Value), 0, adoquery.Fields(0).Value)), FAmount)) & " where a0401 = " & Val(adoacc040.Fields("a0401").Value) & " and a0402 = " & Val(adoacc040.Fields("a0402").Value) & " and a0403 = '" & adoacc040.Fields("a0403").Value & "' and a0404 = 'P' and a0405 = '9999'"
'      End If
'      adoquery.Close
'      adoTaie.Execute "delete from acc040 where a0401 = " & Val(adoacc040.Fields("a0401").Value) & " and a0402 = " & Val(adoacc040.Fields("a0402").Value) & " and a0403 = '" & adoacc040.Fields("a0403").Value & "' and a0404 = 'SAL' and a0405 = '9999'"
'      adoTaie.Execute "insert into acc040 values (" & Val(adoacc040.Fields("a0401").Value) & ", " & Val(adoacc040.Fields("a0402").Value) & ", '" & adoacc040.Fields("a0403").Value & "', 'SAL', '9999', 0, 0, " & IIf(IsNull(adoacc040.Fields(3).Value), 0, Val(adoacc040.Fields(3).Value) * -1) & ", 0, '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", null, null, null)"
'      adoacc040.MoveNext
'   Loop
'   adoacc040.Close
   '2014/1/22 modify by sonia 加公司別
   'adoTaie.Execute "update acc0b0 set a0b01 = " & Val(FCDate(MaskEdBox2.Text)) & ""
   'Modify by Amy 2020/0415 公司別改下拉 原:IIf(Text3 = "2", "J", Text3)
   adoTaie.Execute "update acc0b0 set a0b01 = " & Val(FCDate(MaskEdBox2.Text)) & " where a0b04 = '" & strCmp & "'"
   StatusClear
   MsgBox MsgText(15), , MsgText(21)
   Exit Sub
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
   Exit Sub
End Sub

'*************************************************
'  儲存資料表(會計科目餘額資料)
'
'*************************************************
Private Sub Acc040Save()
   adoacc040.Fields("a0401").Value = Val(Mid(CFDate(adoacc021.Fields("a0205").Value), 1, 3))
   adoacc040.Fields("a0402").Value = Val(Mid(CFDate(adoacc021.Fields("a0205").Value), 5, 2))
   adoacc040.Fields("a0403").Value = adoacc021.Fields("ax201").Value
   adoacc040.Fields("a0405").Value = adoacc021.Fields("ax205").Value
   adoacc040.Fields("a0406").Value = 0
   adoacc040.Fields("a0407").Value = 0
   adoacc040.Fields("a0408").Value = 0
   adoacc040.Fields("a0409").Value = 0
End Sub

'*************************************************
'  儲存資料表 (會計科目資料)
'
'*************************************************
'Modify by Amy 2020/04/15 +公司別參數
Private Sub AccountSave(strCmp As String)

On Error GoTo Checking
   adoacc010.CursorLocation = adUseClient
   '2014/1/22 modify by sonia 加a0109條件
   'adoacc010.Open "select * from acc010 where substr(a0101, 1, 1) in ('1', '2', '3') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2020/04/15 公司別IIf(Text3 = "2", "J", Text3) 改抓參數
   adoacc010.Open "select * from acc010 where substr(a0101, 1, 1) in ('1', '2', '3') and (a0109 is null or a0109='" & strCmp & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   ProgressBar1.max = adoacc010.RecordCount
   Do While adoacc010.EOF = False
      adoacc090.CursorLocation = adUseClient
      adoacc090.Open "select * from acc090 where a0901 = '" & MsgText(55) & "' order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
      Do While adoacc090.EOF = False
         adoacc040.CursorLocation = adUseClient
         'Modify by Amy 2020/04/15 公司別IIf(Text3 = "2", "J", Text3) 改抓參數
         adoacc040.Open "select a0401 from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 = " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0403 = '" & strCmp & "' and a0404 = '" & adoacc090.Fields("a0901").Value & "' and a0405 = '" & adoacc010.Fields("a0101").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoacc040.RecordCount = 0 Then
            'Modify by Amy 2020/04/15 公司別IIf(Text3 = "2", "J", Text3) 改抓參數
            adoTaie.Execute "insert into acc040 (a0401, a0402, a0403, a0404, a0405,a0406, a0407, a0408, a0409, a0411, a0412, a0413) values " & _
                            "(" & Val(Mid(MaskEdBox1.Text, 1, 3)) & ", " & Val(Mid(MaskEdBox1.Text, 5, 2)) & ", '" & strCmp & "', '" & adoacc090.Fields("a0901").Value & "', '" & adoacc010.Fields("a0101").Value & "', 0, 0, 0, 0, " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strUserNum & "')"
         End If
         adoacc040.Close
         adoacc090.MoveNext
      Loop
      adoacc090.Close
      ProgressBar1.Value = ProgressBar1.Value + 1
      adoacc010.MoveNext
   Loop
   adoacc010.Close

Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)

End Sub

'Mark by Amy 2020/04/15 公司別改下拉
''2014/1/22 add by sonia
'Private Sub Text3_GotFocus()
'   TextInverse Text3
'End Sub
'
'Private Sub Text3_Validate(Cancel As Boolean)
'
'   If Text3 <> MsgText(601) Then
'      If Text3 <> "1" And Text3 <> "2" Then
'         MsgBox "公司別只可輸入 1 或 2", , MsgText(5)
'         Cancel = True
'         Text3.SetFocus
'         Exit Sub
'      End If
'   Else
'      MsgBox "請輸入公司別!!", , MsgText(5)
'      Cancel = True
'      Text3.SetFocus
'      Exit Sub
'   End If
'End Sub
'2014/1/22 end
