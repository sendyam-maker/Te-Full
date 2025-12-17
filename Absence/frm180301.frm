VERSION 5.00
Begin VB.Form frm180301 
   BorderStyle     =   1  '單線固定
   Caption         =   "出缺勤查詢"
   ClientHeight    =   4370
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   6290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4370
   ScaleWidth      =   6290
   Begin VB.ComboBox cboDept 
      Height          =   260
      Index           =   1
      Left            =   3780
      TabIndex        =   5
      Text            =   "cboDept"
      Top             =   1320
      Width           =   1965
   End
   Begin VB.ComboBox cboDept 
      Height          =   260
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Text            =   "cboDept"
      Top             =   1320
      Width           =   1965
   End
   Begin VB.TextBox txtST06 
      Height          =   300
      Index           =   1
      Left            =   2310
      MaxLength       =   1
      TabIndex        =   11
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtST06 
      Height          =   300
      Index           =   0
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   10
      Top             =   2640
      Width           =   495
   End
   Begin VB.ComboBox CboB1008 
      Height          =   300
      ItemData        =   "frm180301.frx":0000
      Left            =   1680
      List            =   "frm180301.frx":0002
      TabIndex        =   9
      Top             =   2310
      Width           =   1665
   End
   Begin VB.OptionButton Option1 
      Caption         =   "明細"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   690
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      Caption         =   "僅個人出缺勤統計"
      Height          =   255
      Index           =   1
      Left            =   2700
      TabIndex        =   1
      Top             =   690
      Width           =   1815
   End
   Begin VB.ComboBox CboB1002 
      Height          =   300
      ItemData        =   "frm180301.frx":0004
      Left            =   1680
      List            =   "frm180301.frx":0006
      TabIndex        =   8
      Top             =   1980
      Width           =   1665
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Index           =   1
      Left            =   2760
      MaxLength       =   7
      TabIndex        =   3
      Top             =   990
      Width           =   945
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Index           =   0
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   2
      Top             =   990
      Width           =   945
   End
   Begin VB.TextBox txtDept 
      Height          =   300
      Index           =   0
      Left            =   3510
      MaxLength       =   3
      TabIndex        =   12
      Top             =   3150
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtDept 
      Height          =   300
      Index           =   1
      Left            =   4140
      MaxLength       =   3
      TabIndex        =   13
      Top             =   3150
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtB1003 
      Height          =   300
      Index           =   0
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1650
      Width           =   765
   End
   Begin VB.TextBox txtB1003 
      Height          =   300
      Index           =   1
      Left            =   2580
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1650
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   360
      Index           =   0
      Left            =   4290
      TabIndex        =   14
      Top             =   150
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Index           =   1
      Left            =   5190
      TabIndex        =   15
      Top             =   150
      Width           =   800
   End
   Begin VB.Label Label8 
      Caption         =   "（依勞基法規定：每月加班不得超過46小時）"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   3540
      Width           =   4815
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "(1.北所 2.中所 3.南所 4.高所 5.其他)"
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   2850
      TabIndex        =   24
      Top             =   2700
      Width           =   2775
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "所　　別："
      Height          =   180
      Left            =   750
      TabIndex        =   23
      Top             =   2700
      Width           =   900
   End
   Begin VB.Line Line3 
      X1              =   2100
      X2              =   2580
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Label Label5 
      Caption         =   "備註：您有權限查詢的部門別為"
      ForeColor       =   &H00000080&
      Height          =   420
      Left            =   60
      TabIndex        =   22
      Top             =   3900
      Width           =   6180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "假　　別："
      Height          =   180
      Left            =   750
      TabIndex        =   21
      Top             =   2370
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "表單類別："
      Height          =   180
      Left            =   750
      TabIndex        =   20
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查詢內容："
      Height          =   180
      Index           =   1
      Left            =   750
      TabIndex        =   19
      Top             =   750
      Width           =   900
   End
   Begin VB.Line Line4 
      X1              =   2520
      X2              =   3060
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日　　期："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   750
      TabIndex        =   18
      Top             =   1050
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   3630
      X2              =   4110
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "部  門  別："
      Height          =   180
      Left            =   750
      TabIndex        =   17
      Top             =   1380
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "表單當事人："
      Height          =   180
      Index           =   0
      Left            =   570
      TabIndex        =   16
      Top             =   1680
      Width           =   1080
   End
   Begin VB.Line Line1 
      X1              =   2430
      X2              =   2670
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "frm180301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2023/12/19 修改抓新部門程式
'Memo By Sindy 2021/12/20 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Sindy 2011/8/5
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Public cmdState As Integer '紀錄作用按鍵
Public m_IsAbsBossST03 As String
Public m_strEmp As String 'Add By Sindy 2021/12/21 所屬簽核的人員


Public Sub PubShowNextData()
Dim Cancel As Boolean

   Select Case cmdState
      Case 0 '查詢
         cmdState = -1
         If txtDate(0) = "" And txtDate(1) = "" _
            And CboB1002 = "" And CboB1008 = "" _
            And cboDept(0).Text = "" And cboDept(1).Text = "" _
            And txtB1003(0) = "" And txtB1003(1) = "" _
            And txtST06(0) = "" And txtST06(1) = "" Then
            MsgBox "請輸入查詢條件！"
            Exit Sub
         End If
         
         If txtDate(0) <> "" And txtDate(1) = "" Then txtDate(1) = txtDate(0)
         If txtDate(1) <> "" And txtDate(0) = "" Then txtDate(0) = txtDate(1)
         'If txtDept(0) <> "" And txtDept(1) = "" Then txtDept(1) = txtDept(0)
         If cboDept(0) <> "" And cboDept(1) = "" Then cboDept(1) = cboDept(0)
         'If txtDept(1) <> "" And txtDept(0) = "" Then txtDept(0) = txtDept(1)
         If cboDept(1) <> "" And cboDept(0) = "" Then cboDept(0) = cboDept(1)
         If txtB1003(0) <> "" And txtB1003(1) = "" Then txtB1003(1) = txtB1003(0)
         If txtB1003(1) <> "" And txtB1003(0) = "" Then txtB1003(0) = txtB1003(1)
         If txtST06(0) <> "" And txtST06(1) = "" Then txtST06(1) = txtST06(0)
         If txtST06(1) <> "" And txtST06(0) = "" Then txtST06(0) = txtST06(1)
         
         'Add By Sindy 2024/11/5
         For i = 0 To 1
            If txtB1003(i) <> "" Then
               Call txtB1003_Validate(i, Cancel)
               If Cancel = True Then
                  Exit Sub
               End If
            End If
         Next i
         For i = 0 To 1
            If txtDept(i) <> "" Then
               Call txtDept_Validate(i, Cancel)
               If Cancel = True Then
                  Exit Sub
               End If
            End If
         Next i
         For i = 0 To 1
            If cboDept(i) <> "" Then
               Call CboDept_Validate(i, Cancel)
               If Cancel = True Then
                  Exit Sub
               End If
            End If
         Next i
         For i = 0 To 1
            If txtDate(i) <> "" Then
               Call txtDate_Validate(i, Cancel)
               If Cancel = True Then
                  Exit Sub
               End If
            End If
         Next i
         If CboB1002 <> "" Then
            Call CboB1002_Validate(Cancel)
            If Cancel = True Then
               Exit Sub
            End If
         End If
         If CboB1008 <> "" Then
            Call CboB1008_Validate(Cancel)
            If Cancel = True Then
               Exit Sub
            End If
         End If
         For i = 0 To 1
            If txtST06(i) <> "" Then
               Call txtST06_Validate(i, Cancel)
               If Cancel = True Then
                  Exit Sub
               End If
            End If
         Next i
         '2024/11/5 END
         
         Me.Enabled = False
         Screen.MousePointer = vbHourglass
         Me.Hide
         If Option1(0).Value = True Then '明細
            frm180301_01.Show
            'Add By Sindy 2020/5/28
            If Me.cmdOK(1).Tag = "SysQuery" Then
               frm180301_01.cmdOK(1).Enabled = False
            End If
            '2020/5/28 END
            frm180301_01.QueryData
         ElseIf Option1(1).Value = True Then '統計
            frm180301_02.Show
            frm180301_02.QueryData
         End If
         Screen.MousePointer = vbDefault
         Me.Enabled = True
      Case 1 '結束
         Unload Me
      Case Else
   End Select
End Sub

'Modify By Sindy 2020/5/28 Private => Public
Public Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

'Private Sub SetComboData()
''宣告變數
'Dim Rs As New ADODB.Recordset
'Dim ii As Integer
'
'   Me.CboDept(0).Clear
'   Me.CboDept(1).Clear
'   Rs.CursorLocation = adUseClient
'   '2014/2/11 modify by sonia 除電腦中心及人事處外,其他人只能看到有在職員工的部門(王副總提需求江總同意)
'   'rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' Order By A0901", _
'            cnnConnection, adOpenStatic, adLockReadOnly
'   'Modify By Sindy 2023/12/19
'   If strSrvDate(1) >= 新部門啟用日 Then
'      Call SetST93Combo(CboDept(0))
'      Call SetST93Combo(CboDept(1))
'   Else
'      If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M21" Then
'         Rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' Order By A0901", _
'                  cnnConnection, adOpenStatic, adLockReadOnly
'      Else
'         Rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' and a0901<>'P29' and a0901 in (select distinct st03 from staff where st04='1' and st01>'6' and substr(st01,1,1)<'G' and substr(st01,4,1)<>'9') Order By A0901", _
'                  cnnConnection, adOpenStatic, adLockReadOnly
'      End If
'      '2014/2/11 end
'      Me.CboDept(0).AddItem ""
'      Me.CboDept(1).AddItem ""
'      While Not Rs.EOF
'         Me.CboDept(0).AddItem Left(Rs.Fields(0).Value & Space(5), 5) & Rs.Fields(1).Value
'         Me.CboDept(1).AddItem Left(Rs.Fields(0).Value & Space(5), 5) & Rs.Fields(1).Value
'         Rs.MoveNext
'      Wend
'      If Rs.State <> adStateClosed Then Rs.Close
'      Set Rs = Nothing
'   End If
'
'   '預設值
'   'Modify By Sindy 2023/12/19
'   If strSrvDate(1) >= 新部門啟用日 Then
'      txtDept(0) = Pub_StrUserSt93
'      txtDept(1) = Pub_StrUserSt93
'   Else
'      txtDept(0) = Pub_StrUserSt03
'      txtDept(1) = Pub_StrUserSt03
'   End If
'   'Add By Sindy 2021/12/21
'   If Pub_GetSpecMan("專利處出缺勤可查詢權限") <> "" And InStr(Pub_GetSpecMan("專利處出缺勤可查詢權限"), strUserNum) > 0 Then
'      'Modify By Sindy 2023/12/19
'      If strSrvDate(1) >= 新部門啟用日 Then
''         txtDept(0) = "P00"
''         txtDept(1) = "P41" 'Modify By Sindy 2024/1/30 "P99"
'         'Modify By Sindy 2025/3/5 用SQL抓出起迄部門別
'         strSql = "select A0921,A0922 From ACC090New where A0921 in(" & _
'                  "select A0921 From staff,ACC090New " & _
'                  "where substr(st01,1,1) in (" & ST01CodeNum1 & ") " & _
'                  "and st04='1' and substr(st93,1,1)='P'" & _
'                  "and substr(st01,4,1)<>'9' " & _
'                  "and st01 not in('60000','96029','96030') " & _
'                  "and ST93=A0921(+) and A0921 is not null group by A0921) " & _
'                  "order by A0921 asc "
'         Rs.CursorLocation = adUseClient
'         Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If Rs.RecordCount > 0 Then
'            Rs.MoveFirst
'            txtDept(0) = Rs.Fields("A0921")
'            Rs.MoveLast
'            txtDept(1) = Rs.Fields("A0921")
'         End If
'         '2025/3/5 END
'      Else
'      '2023/12/19 END
'         txtDept(0) = "P10"
'         txtDept(1) = "P14"
'      End If
'   End If
'
'   For ii = 1 To CboDept(0).ListCount - 1
'      If Left(CboDept(0).List(ii), Len(txtDept(0))) = txtDept(0) Then
'         CboDept(0).ListIndex = ii
'         Exit For
'      End If
'   Next ii
'   For ii = 1 To CboDept(1).ListCount - 1
'      If Left(CboDept(1).List(ii), Len(txtDept(1))) = txtDept(1) Then
'         CboDept(1).ListIndex = ii
'         Exit For
'      End If
'   Next ii
'   '2021/12/21 END
'End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   cmdState = -1
   
   Me.cmdOK(1).Tag = "" 'Add By Sindy 2020/5/28
   '預設值
   SetB1002Combo CboB1002
   CboB1002.AddItem "04 外出" 'Add By Sindy 2013/6/26
   SetB1008Combo CboB1008
   
   'Modify By Sindy 2025/3/19
   'SetComboData 'Add By Sindy 2021/12/20
   Call PUB_SetQFormCol_ABS(m_IsAbsBossST03, m_strEmp, Me.Name, txtDate(0), txtDate(1), txtB1003(0), txtB1003(1), _
      cboDept(0), cboDept(1), txtDept(0), txtDept(1), txtST06(0), txtST06(1), Me.Label5)
   '2025/3/19 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180301 = Nothing
End Sub

Private Sub CboB1002_GotFocus()
   InverseTextBox CboB1002
End Sub

Private Sub CboB1002_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboB1002_LostFocus()
   If CboB1002.Text > "" Then
      For i = 0 To CboB1002.ListCount - 1
         If Left(CboB1002.List(i), 2) = CboB1002.Text Then CboB1002.Text = CboB1002.List(i): Exit For
      Next i
   End If
End Sub

Private Sub CboB1002_Validate(Cancel As Boolean)
Dim bolComp As Boolean
   
   If CboB1002 <> "" Then
      bolComp = False
      For i = 0 To CboB1002.ListCount
         If Left(CboB1002, 2) = Left(CboB1002.List(i), 2) Then
            bolComp = True
            Exit For
         End If
      Next i
      If bolComp = False Then
         MsgBox "表單類別有誤!!!", vbExclamation + vbOKOnly
         Call CboB1002_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub CboB1008_GotFocus()
   InverseTextBox CboB1008
End Sub

Private Sub CboB1008_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboB1008_LostFocus()
   If CboB1008.Text > "" Then
      For i = 0 To CboB1008.ListCount - 1
         If Left(CboB1008.List(i), 2) = CboB1008.Text Then CboB1008.Text = CboB1008.List(i): Exit For
      Next i
   End If
End Sub

Private Sub CboB1008_Validate(Cancel As Boolean)
Dim MyRs As New ADODB.Recordset
Dim MyArr As Variant
   
   If CboB1008.Text <> "" Then
      MyArr = Split(CboB1008, " ")
      Set MyRs = New ADODB.Recordset
      If MyRs.State = 1 Then MyRs.Close
      ' 排除不須要的代碼 : 01.忘打卡 02.遲到 03.曠職 04.出差 16.加班 17.扣年終產假 18.扣年終流產假
      strSql = "select ac02||' '||ac03 from allcode where ac01='04' and ac02='" & MyArr(0) & "' and ac02 not in ('01','02','03','04','16','17','18') order by ac02"
      MyRs.CursorLocation = adUseClient
      MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If MyRs.RecordCount <> 0 Then
         CboB1008.Text = "" & MyRs.Fields(0).Value
      Else
         MsgBox "假別代號輸入錯誤!!!", vbExclamation + vbOKOnly
         Call CboB1008_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtB1003_GotFocus(Index As Integer)
   InverseTextBox txtB1003(Index)
End Sub

Private Sub txtB1003_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'2013/7/26 ADD BY SONIA
Private Sub txtB1003_LostFocus(Index As Integer)
   Select Case Index
      Case 0
         txtB1003(1) = txtB1003(0)
   End Select
End Sub
'2013/7/26 END

Private Sub txtB1003_Validate(Index As Integer, Cancel As Boolean)
   If txtB1003(Index).Text <> "" Then
      If ChkStaffID(txtB1003(Index)) Then
         Call txtB1003_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
   End If
   If Index = 0 Then
      If txtB1003(Index) <> "" And txtB1003(Index + 1) = "" Then
         txtB1003(Index + 1) = txtB1003(Index)
      End If
      If txtB1003(Index) > txtB1003(Index + 1) Then
         txtB1003(Index + 1) = txtB1003(Index)
      End If
   ElseIf Index = 1 Then
      If txtB1003(Index) <> "" And txtB1003(Index - 1) = "" Then
         txtB1003(Index - 1) = txtB1003(Index)
      End If
      If txtB1003(Index - 1) <> "" And txtB1003(Index) <> "" Then
         If RunNick(txtB1003(Index - 1), txtB1003(Index)) Then
            Call txtB1003_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtDept_GotFocus(Index As Integer)
   InverseTextBox txtDept(Index)
End Sub

Private Sub txtDept_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDept_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
      If txtDept(Index) <> "" And txtDept(Index + 1) = "" Then
         txtDept(Index + 1) = txtDept(Index)
      End If
      If txtDept(Index) > txtDept(Index + 1) Then
         txtDept(Index + 1) = txtDept(Index)
      End If
   ElseIf Index = 1 Then
      If txtDept(Index) <> "" And txtDept(Index - 1) = "" Then
         txtDept(Index - 1) = txtDept(Index)
      End If
      If txtDept(Index - 1) <> "" And txtDept(Index) <> "" Then
         If RunNick(txtDept(Index - 1), txtDept(Index)) Then
            Call txtDept_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

'Add By Sindy 2021/12/21
Private Sub CboDept_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub CboDept_Validate(Index As Integer, Cancel As Boolean)
Dim strDept0 As String, strDept1 As String

   If Index = 0 Then
      If cboDept(Index) <> "" And cboDept(Index + 1) = "" Then
         cboDept(Index + 1) = cboDept(Index)
      End If
      If cboDept(Index) > cboDept(Index + 1) Then
         cboDept(Index + 1) = cboDept(Index)
      End If
   ElseIf Index = 1 Then
      If cboDept(Index) <> "" And cboDept(Index - 1) = "" Then
         cboDept(Index - 1) = cboDept(Index)
      End If
      strDept0 = Left(Trim(cboDept(Index - 1)), 3)
      strDept1 = Left(Trim(cboDept(Index)), 3)
      If strDept0 <> "" And strDept1 <> "" Then
         If RunNick(strDept0, strDept1) Then
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub
'2021/12/21 END

Private Sub txtDate_GotFocus(Index As Integer)
   InverseTextBox txtDate(Index)
End Sub

Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
   If txtDate(Index).Text <> "" Then
      If ChkDate(txtDate(Index)) = False Then
         Call txtDate_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
   End If
   If Index = 0 Then
      If txtDate(Index) <> "" And txtDate(Index + 1) = "" Then
         txtDate(Index + 1) = txtDate(Index)
      End If
      If Val(txtDate(Index)) > Val(txtDate(Index + 1)) Then
         txtDate(Index + 1) = txtDate(Index)
      End If
   ElseIf Index = 1 Then
      If txtDate(Index) <> "" And txtDate(Index - 1) = "" Then
         txtDate(Index - 1) = txtDate(Index)
      End If
      If txtDate(Index - 1) <> "" And txtDate(Index) <> "" Then
         If RunNick2(txtDate(Index - 1), txtDate(Index)) Then
            Call txtDate_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtST06_GotFocus(Index As Integer)
   InverseTextBox txtST06(Index)
End Sub

Private Sub txtST06_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtST06_Validate(Index As Integer, Cancel As Boolean)
   If txtST06(Index) <> "" Then
      If CheckLengthIsOK(txtST06(Index), txtST06(Index).MaxLength) = False Then
          Call txtST06_GotFocus(Index)
          Cancel = True
          Exit Sub
      End If
      If Trim(txtST06(Index)) <> "" Then
         If txtST06(Index) <> "1" And txtST06(Index) <> "2" And txtST06(Index) <> "3" And _
            txtST06(Index) <> "4" And txtST06(Index) <> "5" Then
            MsgBox "所別代碼有誤!!!", vbExclamation + vbOKOnly
            Call txtST06_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
   If Index = 0 Then
      If txtST06(Index) <> "" And txtST06(Index + 1) = "" Then
         txtST06(Index + 1) = txtST06(Index)
      End If
      If txtST06(Index) > txtST06(Index + 1) Then
         txtST06(Index + 1) = txtST06(Index)
      End If
   ElseIf Index = 1 Then
      If txtST06(Index) <> "" And txtST06(Index - 1) = "" Then
         txtST06(Index - 1) = txtST06(Index)
      End If
      If txtST06(Index - 1) <> "" And txtST06(Index) <> "" Then
         If RunNick(txtST06(Index - 1), txtST06(Index)) Then
            Call txtST06_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
CloseIme
End Sub
