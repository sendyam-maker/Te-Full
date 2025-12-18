VERSION 5.00
Begin VB.Form frm090615_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人、繪圖人員目標資料維護(複製資料)"
   ClientHeight    =   1965
   ClientLeft      =   1230
   ClientTop       =   2475
   ClientWidth     =   4425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4425
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2376
      TabIndex        =   5
      Top             =   50
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3156
      TabIndex        =   6
      Top             =   50
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2304
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1644
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1215
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1644
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1215
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1224
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1215
      MaxLength       =   3
      TabIndex        =   1
      Top             =   864
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1215
      MaxLength       =   3
      TabIndex        =   0
      Top             =   444
      Width           =   525
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   2160
      TabIndex        =   11
      Top             =   900
      Width           =   1740
   End
   Begin VB.Line Line1 
      X1              =   1668
      X2              =   2223
      Y1              =   1752
      Y2              =   1752
   End
   Begin VB.Label Label1 
      Caption         =   "複製年月："
      Height          =   180
      Index           =   3
      Left            =   132
      TabIndex        =   10
      Top             =   1668
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "被複製年月："
      Height          =   180
      Index           =   2
      Left            =   132
      TabIndex        =   9
      Top             =   1272
      Width           =   1116
   End
   Begin VB.Label Label1 
      Caption         =   "部門別："
      Height          =   180
      Index           =   1
      Left            =   132
      TabIndex        =   8
      Top             =   888
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   132
      TabIndex        =   7
      Top             =   492
      Width           =   948
   End
End
Attribute VB_Name = "frm090615_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/17 改成Form2.0 (無)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim TestOk As Boolean, i As Integer, j As Integer, MonthMenu() As Long, strTemp As String, MonthTmp() As Long
Dim s As Integer, strTemp1 As Variant, strTemp2 As Variant
'Add By Cheng 2002/05/24
Dim m_blnCancel As Boolean

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         If Len(txt1(1)) = 0 Then
             s = MsgBox("部門別不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             Exit Sub
         Else
             If Len(txt1(2)) = 0 Then
                 s = MsgBox("被複製年月不可空白!!", , "USER 輸入錯誤")
                 txt1(2).SetFocus
                 Exit Sub
             Else
                 If Len(txt1(3)) = 0 Or Len(txt1(4)) = 0 Then
                     s = MsgBox("複製年月不可空白!!", , "USER 輸入錯誤")
                     'If Len(txt1(4)) = 0 Then txt1(4).SetFocus
                     If Len(txt1(3)) = 0 Then txt1(3).SetFocus
                     Exit Sub
                 Else
                     'Add By Cheng 2002/05/24
                     '重新檢查欄位有效性
                     If TxtValidate = False Then Exit Sub
                 
                     Screen.MousePointer = vbHourglass
                     Me.Enabled = False
                     Process
                     Me.Enabled = True
                     Screen.MousePointer = vbDefault
                 End If
             End If
         End If
     End If
Case 1
     Me.Hide
     frm090615.REFormLoad
     frm090615.Show
     Unload Me
Case Else
End Select
End Sub

Sub Process()
strSql = "SELECT PE01,PE02,PE03,PE04,PE05,PE06,PE07,PE08,PE09,PE10,PE11,PE12,PE13,PE14,PE15,PE16,PE17,PE18,PE19,PE20,PE21,PE22,PE23,PE24,PE25,PE26,PE27,PE28,PE29 FROM PERFORMANCE,STAFF WHERE ST01=PE01(+) AND ST04='1' AND ST03='" & txt1(1) & "' AND PE02='" & txt1(0) & "' AND PE03>=" & Val(txt1(3)) + 191100 & " AND PE03<=" & Val(txt1(4)) + 191100
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        TestOk = True
    Else
        TestOk = False
    End If
End With
CheckOC2
If TestOk = True Then
    s = MsgBox("資料庫中已經有被複製年月區間了!!", , "無法新增")
    Exit Sub
End If
TestOk = False
MonthMenu(0) = Val(txt1(3))
If txt1(3) <> txt1(4) Then
        strTemp = txt1(3)
    Do While TestOk = False
        '900629  不能用 FORMAT 因為有些人的 OS 不能用     NICK 改
        'strTemp = Format(DateAdd("M", 1, ChangeWStringToWDateString(ChangeTStringToWString(strTemp & "01"))), "EEMM")
        strTemp = Trim(str(Val(Mid(ChangeWDateStringToWString(DateAdd("M", 1, ChangeWStringToWDateString(ChangeTStringToWString(strTemp & "01")))), 1, 6)) - 191100))
        ReDim MonthTmp(UBound(MonthMenu) + 1) As Long
        For i = 0 To UBound(MonthMenu)
            MonthTmp(i) = MonthMenu(i)
        
        Next i
        ReDim MonthMenu(UBound(MonthMenu) + 1) As Long
        For i = 0 To UBound(MonthTmp)
            MonthMenu(i) = MonthTmp(i)
        Next i
        MonthMenu(UBound(MonthMenu)) = Val(strTemp)
        If strTemp = txt1(4) Then
            TestOk = True
        End If
    Loop
End If
strSql = ""
For i = 0 To UBound(MonthMenu)
    'Modify by Morgan 2011/4/1
    '100/4以後改1件7張(原來1件5張)
    If Val(txt1(2)) < 10004 And Val(MonthMenu(i)) >= 10004 Then
      strSql = strSql + " SELECT PE01,PE02," & MonthMenu(i) + 191100 & ",PE04,PE05,PE06,PE07,PE08,PE09,round(PE09*7) PE10,PE11,PE12,PE13,PE14,PE15,PE16,PE17,PE18,PE19,PE20,PE21,PE22,PE23,PE24,PE25,PE26,PE27,PE28,PE29 FROM PERFORMANCE,STAFF WHERE ST01=PE01(+) AND ST04='1' AND ST03='" & txt1(1) & "' AND PE02='" & txt1(0) & "' AND PE03=" & Val(txt1(2)) + 191100
    Else
      strSql = strSql + " SELECT PE01,PE02," & MonthMenu(i) + 191100 & ",PE04,PE05,PE06,PE07,PE08,PE09,PE10,PE11,PE12,PE13,PE14,PE15,PE16,PE17,PE18,PE19,PE20,PE21,PE22,PE23,PE24,PE25,PE26,PE27,PE28,PE29 FROM PERFORMANCE,STAFF WHERE ST01=PE01(+) AND ST04='1' AND ST03='" & txt1(1) & "' AND PE02='" & txt1(0) & "' AND PE03=" & Val(txt1(2)) + 191100
    End If
    If i <> UBound(MonthMenu) Then
        strSql = strSql + " UNION "
    End If
Next i
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        strSql = "INSERT INTO PERFORMANCE (PE01,PE02,PE03,PE04,PE05,PE06,PE07,PE08,PE09,PE10,PE11,PE12,PE13,PE14,PE15,PE16,PE17,PE18,PE19,PE20,PE21,PE22,PE23,PE24,PE25,PE26,PE27,PE28,PE29) " & strSql
        cnnConnection.Execute strSql
    Else
        s = MsgBox("無符合被複製年月的資料!!", , "錯誤")
        Me.Enabled = True
        txt1(2).SetFocus
        txt1_GotFocus (2)
        CheckOC2
        Exit Sub
    End If
End With
CheckOC
s = MsgBox("複製OK！！", , "OK！！")
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
ReDim MonthMenu(0) As Long
ReDim MonthTmp(0) As Long
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090615_1 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
'Add By Cheng 2002/05/24
m_blnCancel = False

Select Case Index
Case 0 '系統類別
      'Add By Cheng 2002/01/07
      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
     strTemp1 = Split(UCase(Systemkind_g), ",")
     strTemp2 = Split(UCase(txt1(0)), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp2(i) = strTemp1(j) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            'Add By Cheng 2002/05/24
            m_blnCancel = True
            Exit Sub
        End If
     Next i
Case 1
     If Len(Trim(txt1(1))) <> 0 Then
        strSql = "SELECT NVL(A0902,A0903) FROM ACC090 WHERE A0901='" & txt1(1) & "' "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            lbl1.Caption = CheckStr(adoRecordset1.Fields(0))
        Else
            s = MsgBox("部門別輸入錯誤找不到!!", , "USER 輸入錯誤")
            lbl1.Caption = ""
            txt1(1).SetFocus
            txt1(1).SelStart = 0
            txt1(1).SelLength = Len(txt1(1))
            CheckOC2
            'Add By Cheng 2002/05/24
            m_blnCancel = True
            Exit Sub
        End If
     End If
Case Else
End Select
End Sub

'Add By Cheng 2002/05/24
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In txt1
   If objTxt.Enabled = True Then
      Cancel = False
      txt1_LostFocus objTxt.Index
      If m_blnCancel = True Then
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function

