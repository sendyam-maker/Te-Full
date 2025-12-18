VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880024 
   BorderStyle     =   1  '單線固定
   Caption         =   "各部門人員"
   ClientHeight    =   4700
   ClientLeft      =   4670
   ClientTop       =   900
   ClientWidth     =   3490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4700
   ScaleWidth      =   3490
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   1290
      TabIndex        =   2
      Top             =   45
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   2250
      TabIndex        =   3
      Top             =   45
      Width           =   1200
   End
   Begin VB.Label Label4 
      Caption         =   "１. 快速點二下，即直接帶入前畫面"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   4230
      Width           =   3225
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "可輸入員編或姓名後，按Tab直接帶入"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   480
      TabIndex        =   6
      Top             =   570
      Width           =   2970
   End
   Begin MSForms.ComboBox cboDepName 
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Top             =   780
      Width           =   2340
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4128;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstEmp 
      Height          =   3090
      Left            =   240
      TabIndex        =   1
      Top             =   1110
      Width           =   3080
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "5433;5450"
      MatchEntry      =   0
      MultiSelect     =   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "２. 點選多筆：按住 Ctrl 點選資料列"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4470
      Width           =   3225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "部　門："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   210
      TabIndex        =   4
      Top             =   870
      Width           =   720
   End
End
Attribute VB_Name = "frm880024"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Sindy 2022/7/19
Option Explicit

Dim m_PrevForm As Form


Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Public Sub cboDepName_Click()
Dim Rs As New ADODB.Recordset
Dim strDept As String
Dim strST16Con As String, strST16Order As String
   
   If cboDepName.Text <> "" Then
      strDept = Left(cboDepName, 3)
   Else
      'Modify By Sindy 2023/12/29
      If strSrvDate(1) >= 新部門啟用日 Then
         strDept = Pub_StrUserSt93
      Else
      '2023/12/29 END
         strDept = Pub_StrUserSt03
      End If
   End If
   Me.lstEmp.Clear
   'Modify By Sindy 2023/12/29
   If strSrvDate(1) >= 新部門啟用日 Then
      strExc(0) = "Select ST01,ST02,st16 From STAFF" & _
                  " Where ST93='" & strDept & "' and ST04='1' and substr(st01,1,1)<'F' and substr(st01,4,1)<>'9'" & _
                  strST16Order
   Else
      strST16Con = "''"
      strST16Order = ""
      If strDept = "F21" Then '外專工程師
         strST16Con = "decode(st16,'1','(電子電機組)','2','(化學組)','3','(日文組)','4','(機械設計組)',st16)"
         strST16Order = " Order By ST16 asc,ST01 asc"
      ElseIf strDept = "F23" Then '外專承辦
         strST16Con = "decode(st16,'1','(英文組)','2','(日文組)',st16)"
         strST16Order = " Order By ST16 asc,ST01 asc"
      ElseIf Left(strDept, 2) = "F1" Then '外商
         strST16Con = "decode(st16,'2','(英文組)','4','(日文組)','6','(CF案)',st16)"
         strST16Order = " Order By ST16 asc,ST01 asc"
      Else
         strST16Order = " Order By ST01 asc"
      End If
      strExc(0) = "Select ST01,ST02||' '||" & strST16Con & ",st16 From STAFF" & _
                  " Where ST03='" & strDept & "' and ST04='1' and substr(st01,1,1)<'F' and substr(st01,4,1)<>'9'" & _
                  strST16Order
   End If
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Rs.MoveFirst
      While Not Rs.EOF
         Me.lstEmp.AddItem Left(Rs.Fields(0).Value & Space(6), 6) & Rs.Fields(1).Value
         Rs.MoveNext
      Wend
   End If
   
   If cboDepName.Text = "" And InStr(Me.Caption, "轉寄收受者") > 0 Then
      If Pub_StrUserSt03 = "F23" Then
         Me.lstEmp.AddItem "外商群組 國外部轉信外商群組 *" '洪琬姿 ,葉易雲,沈佳穎,陳蒲璇
         Me.lstEmp.AddItem "新知群組 國外部轉信新知群組 *" '閻?泰,EXTERNAL_NEWS@taie.com.tw,顏裕洋,鄒宜珊  ? 妳們轉寄給這群組, David也是一員, 但你會因若人員上了不處理 且 主管核准了,信件就沖銷了哦~~ 以上也是如此…
         Me.lstEmp.AddItem "開拓群組 國外部轉信開拓群組 *" '閻?泰,楊雯芳,陳增廣,鄒宜珊
         Me.lstEmp.AddItem "外法群組 國外部轉信外法群組;國外部轉信外專承辦日文組長 *" 'Add By Sindy 2023/3/31
         Me.lstEmp.AddItem "代理人通知 國外部轉信外商群組;國外部轉信外專群組;patent;99033;A4024 *" 'Add By Sindy 2023/3/31
         Me.lstEmp.AddItem "10F傳真機 25011666@taie.com.tw *" 'Add By Sindy 2023/3/31
         Me.lstEmp.AddItem "國內信件 國內信件管理人員 *" 'Add By Sindy 2023/3/31
         Me.lstEmp.AddItem "Patent Patent@taie.com.tw *"
         Me.lstEmp.AddItem "TM TM@taie.com.tw *"
         Me.lstEmp.AddItem "account account@taie.com.tw *" 'Add By Sindy 2023/3/31
      'Add By Sindy 2023/5/10
      ElseIf Left(Pub_StrUserSt03, 2) = "F1" Then
         Me.lstEmp.AddItem "IPDept IPDept@taie.com.tw *"
         Me.lstEmp.AddItem "Patent Patent@taie.com.tw *"
         Me.lstEmp.AddItem "TM TM@taie.com.tw *"
         Me.lstEmp.AddItem "account account@taie.com.tw *"
         '2023/5/10 END
         Me.lstEmp.AddItem "lawoffice@taie.com.tw" 'Add By Sindy 2024/4/8
      End If
      Me.lstEmp.AddItem "" 'Add By Sindy 2023/3/31
   End If
   
   If Me.lstEmp.ListCount > 0 Then Me.lstEmp.ListIndex = 0 'Add By Sindy 2023/6/7
   Set Rs = Nothing
End Sub

Private Sub cboDepName_GotFocus()
   cboDepName.SelStart = 0
   cboDepName.SelLength = Len(cboDepName.Text)
End Sub
Private Sub cboDepName_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub cboDepName_LostFocus()
Dim strText As String
Dim bolHad As Boolean '有抓到資料
   
   bolHad = False
   cboDepName.Text = Trim(cboDepName.Text)
   If cboDepName.Text <> "" Then
      '依員工姓名抓取員工編號
      strText = GetPrjSalesNM_2(cboDepName.Text)
      If strText <> "" Then
         cboDepName.Text = strText & " " & cboDepName.Text
         bolHad = True
      Else
         '依員工編號抓取員工姓名
         strText = GetPrjSalesNM(Left(cboDepName.Text, 5))
         If strText <> "" Then
            cboDepName.Text = Left(cboDepName.Text, 5) & " " & strText
            bolHad = True
         End If
      End If
      'Add By Sindy 2024/4/8 @taie.com.tw
      If InStr(UCase(cboDepName.Text), UCase("@taie.com.tw")) > 0 Then
         bolHad = True
      End If
      '2024/4/8 END
      If bolHad = True Then
         m_PrevForm.m_LstEmp = cboDepName.Text
         Unload Me
      End If
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim ArrStr As Variant
Dim i As Integer

If Index = 0 Then '確定
   m_PrevForm.m_LstEmp = ""
   For i = 0 To lstEmp.ListCount - 1
      If lstEmp.Selected(i) = True Then
         ArrStr = Split(lstEmp.List(i), " ")
         'Add By Sindy 2023/3/31
         If UBound(ArrStr) >= 0 Then
         '2023/3/31 END
            If m_PrevForm.m_LstEmp <> "" Then m_PrevForm.m_LstEmp = m_PrevForm.m_LstEmp & ";"
            If InStr(lstEmp.List(i), "*") > 0 Then
               m_PrevForm.m_LstEmp = m_PrevForm.m_LstEmp & ArrStr(0) & " " & ArrStr(1)
            Else
               m_PrevForm.m_LstEmp = m_PrevForm.m_LstEmp & ArrStr(0)
            End If
         End If
      End If
   Next i
End If
Unload Me
End Sub

Private Sub Form_Activate()
   'Modify By Sindy 2023/6/7
   'cboDepName.SetFocus
   lstEmp.SetFocus
   '2023/6/7 END
End Sub

Private Sub Form_Load()
   Call SetComboData
   'Call cboDepName_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   Set frm880024 = Nothing
End Sub

Private Sub SetComboData()
Dim Rs As New ADODB.Recordset
   
   'Modify By Sindy 2023/12/29
   If strSrvDate(1) >= 新部門啟用日 Then
      Call SetST93Combo(cboDepName)
   Else
   '2023/12/29 END
      Me.cboDepName.Clear
      Rs.CursorLocation = adUseClient
      '除電腦中心及人事處外,其他人只能看到有在職員工的部門(王副總提需求江總同意)
      If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M21" Then
         Rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' Order By A0901", _
                  cnnConnection, adOpenStatic, adLockReadOnly
      Else
         Rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' and a0901<>'P29' and a0901 in (select distinct st03 from staff where st04='1' and st01>'6' and substr(st01,1,1)<'G' and substr(st01,4,1)<>'9') Order By A0901", _
                  cnnConnection, adOpenStatic, adLockReadOnly
      End If
      Me.cboDepName.AddItem ""
      While Not Rs.EOF
         Me.cboDepName.AddItem Left(Rs.Fields(0).Value & Space(5), 5) & Rs.Fields(1).Value
         Rs.MoveNext
      Wend
      If Rs.State <> adStateClosed Then Rs.Close
      Set Rs = Nothing
   End If
End Sub

Private Sub lstEmp_Click()
'   strExc(0) = Format(Mid(lstEmp.Text, 26, 9))
'   If Val(strExc(0)) > 0 Then
'      txtMoney = strExc(0)
'      strExc(0) = Mid(lstEmp.Text, 36)
'      If strExc(0) <> "" Then
'         intI = InStr(strExc(0), " ")
'         If intI > 0 Then
'            lblAgent = Trim(Mid(strExc(0), intI))
'            strExc(0) = Left(strExc(0), intI - 1)
'         End If
'         cboDepName = strExc(0)
'      End If
'   Else
'     txtMoney = ""
'     cboDepName = "" 'Added by Morgan 2020/8/7
'     lblAgent = "" 'Added by Morgan 2020/8/7
'   End If
'   If txtMoney.Visible Then
'      txtMoney.SetFocus
'      txtMoney_GotFocus
'   End If
End Sub

Private Sub lstEmp_DblClick(Cancel As MSForms.ReturnBoolean)
   If lstEmp.Selected(lstEmp.ListIndex) = True Then
      Call cmdOK_Click(0)
   End If
End Sub
