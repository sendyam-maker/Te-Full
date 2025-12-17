VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc4330 
   AutoRedraw      =   -1  'True
   Caption         =   "月結算作業"
   ClientHeight    =   2060
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2060
   ScaleWidth      =   3630
   Begin VB.ComboBox CboComp 
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
      Left            =   1230
      TabIndex        =   0
      Text            =   "CboComp"
      Top             =   240
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "執行(&E)"
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
      Left            =   570
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1995
      _ExtentX        =   3528
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
      Left            =   330
      TabIndex        =   4
      Top             =   240
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "結算日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   330
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc4330"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 (無需修改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc0b0 As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoacc010 As New ADODB.Recordset
Public adoacc090 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset

'Add by Amy 2020/04/15
Private Sub CboComp_GotFocus()
    TextInverse cboComp
End Sub

Private Sub CboComp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboComp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(cboComp) = MsgText(601) Then
        MsgBox "請輸入公司別!!", , MsgText(5)
        Cancel = True
        cboComp.SetFocus
        Exit Sub
    End If
    strCmp = cboComp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label3 & MsgText(63), , MsgText(5)
        Cancel = True
        cboComp.SetFocus
        Exit Sub
    ElseIf Len(Trim(cboComp)) = 1 Then
        cboComp = Trim(strCmp) & "　" & A0802Query(strCmp, True)
    End If
End Sub
'end 2020/04/15

Private Sub Command1_Click()
   Me.MousePointer = vbHourglass
   CalculateTable
   Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
'2012/6/15 add by sonia
Dim strYear As String
Dim strMonth As String
Dim strDay As String
'2012/6/15 end
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 3750
   Me.Height = 2460
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
   cboComp.Clear
   Call Pub_SetCboCmp(cboComp, False, True, False, , 0)
   'end 2020/04/15
   'Modify by Morgan 2007/3/1
   'MaskEdBox1.Text = CFDate(ACDate(ServerDate))
   '2012/6/15 modify by sonia 改預設前一月份之月底
   'MaskEdBox1.Text = CFDate(strSrvDate(2))
   MaskEdBox1.Text = CFDate(strSrvDate(2) - 100)
   strYear = Val(Mid(MaskEdBox1.Text, 1, 3))
   strMonth = Val(Mid(MaskEdBox1.Text, 5, 2))
   If strMonth = "0" Then
      strYear = Val(Mid(MaskEdBox1.Text, 1, 3)) - 1
      strMonth = 12
   End If
   If Len(strMonth) < 2 Then
      strMonth = "0" & strMonth
   End If
   Select Case Val(strMonth)
      Case 1, 3, 5, 7, 8, 10, 12
         strDay = "31"
      Case 4, 6, 9, 11
         strDay = "30"
      Case Else
         strDay = "29"
         If CheckIsTaiwanDate(Val(strYear & strMonth & strDay), False) = False Then
            strDay = "28"
         End If
   End Select
   MaskEdBox1.Text = CFDate(strYear & strMonth & strDay)
   '2012/6/15 end
   'end 2007/3/1
   MaskEdBox1.Mask = DFormat
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Call PUB_GetLock("", "Frmacc4320")  'add by sonia 2014/8/8
   Set Frmacc4330 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  月結算計算
'
'*************************************************
Private Sub CalculateTable()
Dim douAmount As Double
Dim strAccNo As String
Dim strSave As String
Dim strYear As String
Dim strMonth As String
Dim strDay As String
Dim strWdate As String   '2013/3/5 add by sonia
Dim strQ As String, strCmp As String 'Add by Amy 2020/04/15

On Error GoTo Checking
  'Modify by Amy 2020/04/15 公司別改下拉 原:IIf(Text3 = "2", "J", Text3)
   strCmp = cboComp
   If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
   End If
   'end 2020/04/15
   strYear = Val(Mid(MaskEdBox1.Text, 1, 3))
   strMonth = Val(Mid(MaskEdBox1.Text, 5, 2))
   If strMonth = "12" Then
      MsgBox MsgText(141), , MsgText(5)
      Exit Sub
   End If
   strMonth = Val(strMonth) + 1
   If Len(strMonth) < 2 Then
      strMonth = "0" & strMonth
   End If
'   Select Case Val(strMonth)
'      Case 1, 3, 5, 7, 8, 10, 12
'         strDay = "31"
'      Case 4, 6, 9, 11
'         strDay = "30"
'      Case Else
'         If (Val(strYear) Mod 4) = 0 Then
'            strDay = "29"
'         Else
'            strDay = "28"
'         End If
'   End Select

   'Modify by Morgan 2007/3/1
   'strDay = Mid(CFDate(ACDate(ServerDate)), 8, 2)
   'Modify by Morgan 2008/11/5 若跨月才做時抓畫面日期次月的最後一天
   'strDay = Right(strSrvDate(1), 2)
   'modify by sonia 2016/3/9
   'If Mid(strSrvDate(1), 5, 2) = strMonth Then
   If Mid(strSrvDate(1), 1, 6) = Val(strYear + strMonth) + 191100 Then
      strDay = Right(strSrvDate(1), 2)
   Else
      strDay = Right(CompDate(2, -1, strYear & Format(Val(strMonth) + 1, "00") & "01"), 2)
   End If
   'end 2008/11/5
   'end 2007/3/1
   '2013/3/5 add by sonia
   strWdate = TransDate(PUB_GetWorkDay1(Val(strYear & Format(strMonth, "00") & Format(strDay, "00")), True), 1)
   strYear = Val(Mid(strWdate, 1, 3))
   strMonth = Val(Mid(strWdate, 4, 2))
   strDay = Right(strWdate, 2)
   '2013/3/5 end
   
   adoacc0b0.CursorLocation = adUseClient
   '2014/1/23 modify by sonia 加公司別
   'adoacc0b0.Open "select * from acc0b0", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2020/04/15 公司別改變數
   strQ = "select * from acc0b0 where a0b04 = '" & strCmp & "'"
   adoacc0b0.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   'end 2020/04/15
   If adoacc0b0.RecordCount <> 0 Then
      If Val(FCDate(MaskEdBox1.Text)) <= IIf(IsNull(adoacc0b0.Fields("a0b02").Value), 0, adoacc0b0.Fields("a0b02").Value) Then
         MsgBox MsgText(77), , MsgText(5)
         adoacc0b0.Close
         Exit Sub
      End If
      'add by sonia 2014/7/7 月結日=過帳日表示尚未過帳, 不可月結
      'modify by sonia 2016/3/9
      'If Val(adoacc0b0.Fields("a0B02").Value) = Val(adoacc0b0.Fields("a0b01").Value) Then
      If Left(Val(adoacc0b0.Fields("a0b02").Value) + 19110000, 6) = Left(Val(adoacc0b0.Fields("a0b01").Value) + 19110000, 6) Then
         MsgBox "尚未過帳, 不可執行月結算作業 ! ", , MsgText(5)
         adoacc0b0.Close
         Exit Sub
      End If
      'end 2014/7/7
   End If
   adoacc0b0.Close
   'add by sonia 2025/1/16 避免過帳後新增傳票未補過帳直接做月結
   strQ = "select distinct a0202 from acc020,acc021 where a0201 = '" & strCmp & "' and a0205>= " & Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2) & "01 and a0205 <= " & Mid(MaskEdBox1.Text, 1, 3) & Mid(MaskEdBox1.Text, 5, 2) & "31 and a0201=ax201(+) and a0202=ax202(+) and ax210 is null"
   adoacc0b0.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0b0.RecordCount <> 0 Then
      MsgBox "尚有傳票未過帳, 不可執行月結算作業 ! ", , MsgText(5)
      adoacc0b0.Close
      Exit Sub
   End If
   adoacc0b0.Close
   'end 2025/1/16
   
   adoacc040.CursorLocation = adUseClient
   '2014/1/23 modify by sonia 加公司別
   'adoacc040.Open "select sum(a0408) from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 = " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0403 = '1' and a0404 = '" & MsgText(55) & "' and (substr(a0405, 1, 1) = '4' or substr(a0405, 1, 2) = '71')", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2020/04/15 公司別改變數
   strQ = "select sum(a0408) from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 = " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0403 = '" & strCmp & "' and a0404 = '" & MsgText(55) & "' and (substr(a0405, 1, 1) = '4' or substr(a0405, 1, 2) = '71')"
   adoacc040.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc040.RecordCount <> 0 Then
      If IsNull(adoacc040.Fields(0).Value) Then
         douAmount = 0
      Else
         douAmount = adoacc040.Fields(0).Value
      End If
   Else
      douAmount = 0
   End If
   adoacc040.Close
   adoacc040.CursorLocation = adUseClient
   '2014/1/23 modify by sonia 加公司別
   'adoacc040.Open "select sum(a0408) from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 = " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0403 = '1' and a0404 = '" & MsgText(55) & "' and (substr(a0405, 1, 1) = '6' or substr(a0405, 1, 2) = '72')", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2020/05/04 公司別改變數
   strQ = "select sum(a0408) from acc040 where a0401 = " & Val(Mid(MaskEdBox1.Text, 1, 3)) & " and a0402 = " & Val(Mid(MaskEdBox1.Text, 5, 2)) & " and a0403 = '" & strCmp & "' and a0404 = '" & MsgText(55) & "' and (substr(a0405, 1, 1) = '6' or substr(a0405, 1, 2) = '72')"
   adoacc040.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc040.RecordCount <> 0 Then
      If IsNull(adoacc040.Fields(0).Value) = False Then
         douAmount = douAmount - adoacc040.Fields(0).Value
      End If
   End If
   adoacc040.Close
   If douAmount <> 0 Then
      '2014/1/23 MODIFY BY SONIA 判斷公司別
      'strAccNo = AccAutoNo(MsgText(801), 4, Val(strYear), Val(strMonth))
      'strSave = AccSaveAutoNo(MsgText(801), Mid(strAccNo, 7, 4), Val(strYear), Val(strMonth))
      'Modify by Amy 2020/04/15 公司別改下拉 原:Text3,並加L 公司
      If strCmp = "J" Then  'J公司用MsgText(819)
         strAccNo = AccAutoNo(MsgText(819), 4, Val(strYear), Val(strMonth))
         strSave = AccSaveAutoNo(MsgText(819), Mid(strAccNo, 7, 4), Val(strYear), Val(strMonth))
      ElseIf strCmp = "L" Then
         strAccNo = AccAutoNo(MsgText(820), 4, Val(strYear), Val(strMonth))
         strSave = AccSaveAutoNo(MsgText(820), Mid(strAccNo, 7, 4), Val(strYear), Val(strMonth))
      Else
         strAccNo = AccAutoNo(MsgText(801), 4, Val(strYear), Val(strMonth))
         strSave = AccSaveAutoNo(MsgText(801), Mid(strAccNo, 7, 4), Val(strYear), Val(strMonth))
      End If
      'end 2020/04/18
      '2014/1/23 END
      'Modify by Morgan 2007/3/1
      'adoTaie.Execute "insert into acc020 values ('1', '" & strAccNo & "', " & Val(strYear & strMonth & strDay) & ", '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", null, null, null)"
      '2014/1/23 modify by sonia 改公司別
      'adoTaie.Execute "insert into acc020 values ('1', '" & strAccNo & "', " & Val(strYear & Format(strMonth, "00") & Format(strDay, "00")) & ", '" & strUserNum & "', " & Val(strSrvDate(2)) & ", to_char(sysdate, 'HH24MISS'), null, null, null)"
      'Modify by Amy 2020/04/15 公司別改變數
      adoTaie.Execute "insert into acc020 values ('" & strCmp & "', '" & strAccNo & "', " & Val(strYear & Format(strMonth, "00") & Format(strDay, "00")) & ", '" & strUserNum & "', " & Val(strSrvDate(2)) & ", to_char(sysdate, 'HH24MISS'), null, null, null)"
      'end 2007/3/1
      If douAmount > 0 Then
         adoTaie.Execute "insert into acc021 values ('" & strCmp & "', '" & strAccNo & "', '001', 'TOT', '3222', " & douAmount & ", 0, NULL, NULL, NULL, NULL, NULL, '" & MsgText(168) & "', NULL, NULL)"
         adoTaie.Execute "insert into acc021 values ('" & strCmp & "', '" & strAccNo & "', '002', 'TOT', '3221', 0, " & douAmount & ", NULL, NULL, NULL, NULL, NULL, '" & MsgText(168) & "', NULL, NULL)"
      Else
         adoTaie.Execute "insert into acc021 values ('" & strCmp & "', '" & strAccNo & "', '001', 'TOT', '3221', " & douAmount * (-1) & ", 0, NULL, NULL, NULL, NULL, NULL, '" & MsgText(168) & "', NULL, NULL)"
         adoTaie.Execute "insert into acc021 values ('" & strCmp & "', '" & strAccNo & "', '002', 'TOT', '3222', 0, " & douAmount * (-1) & ", NULL, NULL, NULL, NULL, NULL, '" & MsgText(168) & "', NULL, NULL)"
      End If
      'end 2020/04/15
   End If
   
   '2014/1/23 modify by sonia 加公司別
   'adoTaie.Execute "update acc0b0 set a0b02 = " & Val(FCDate(MaskEdBox1.Text)) & ""
   'Modify by Amy 2020/04/15 公司別改變數
   adoTaie.Execute "update acc0b0 set a0b02 = " & Val(FCDate(MaskEdBox1.Text)) & " where a0b04 = '" & strCmp & "'"
   MsgBox MsgText(76), , MsgText(21)
Checking:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'2012/5/14 ADD BY SONIA
Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If CheckIsTaiwanDate(ChangeTDateStringToTString(MaskEdBox1)) = False Then
      MaskEdBox1.SetFocus
      Cancel = True
      Exit Sub
   End If
End Sub
'2012/514 END

'Mark by Amy 2020/04/15 公司別改下拉
''2014/1/23 add by sonia
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
'2014/1/23 end

