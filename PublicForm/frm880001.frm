VERSION 5.00
Begin VB.Form frm880001 
   BorderStyle     =   1  '單線固定
   Caption         =   "指定國家"
   ClientHeight    =   5610
   ClientLeft      =   450
   ClientTop       =   990
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5760
   Begin VB.CommandButton Command1 
      Caption         =   "延伸國"
      Height          =   375
      Index           =   1
      Left            =   1020
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "成員國"
      Height          =   375
      Index           =   0
      Left            =   150
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "清除(&C)"
      Height          =   400
      Index           =   2
      Left            =   4752
      TabIndex        =   4
      Top             =   1320
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "刪除(&D)"
      Height          =   400
      Index           =   1
      Left            =   3780
      TabIndex        =   3
      Top             =   1320
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "新增(&A)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Index           =   0
      Left            =   2820
      TabIndex        =   2
      Top             =   1320
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   3576
      TabIndex        =   5
      Top             =   10
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   4512
      TabIndex        =   6
      Top             =   10
      Width           =   1200
   End
   Begin VB.TextBox txtNation 
      Height          =   264
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   0
      Top             =   600
      Width           =   732
   End
   Begin VB.ListBox lstCountry 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2040
      Width           =   5532
   End
   Begin VB.Label Label2 
      Caption         =   "國家名稱(英)　　　　　國家名稱(中)　　　　　國家代號"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   180
      TabIndex        =   11
      Top             =   1800
      Width           =   5472
   End
   Begin VB.Label lblNationName 
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   5472
   End
   Begin VB.Label lblFund 
      Caption         =   "國家代號："
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   972
   End
End
Attribute VB_Name = "frm880001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit
'StrCountry存放指定國家
Public strCountry As String
'Add by Morgan 2005/9/12 strPA10存放申請日
Public strPA10 As String
Private Sub cmdMove_Click(Index As Integer)
Dim i As Integer, intlastIndex As Integer
Dim stCon As String, stCon1 As String

If Index = 0 Then
   If lblNationName <> "" Then
      'Add by Morgan 2009/10/16 EPC的指定國才要限定(馬德里不用)
      If intPWhere = 國外_CF Then
         stCon = "1"
      Else
         '2012/3/21 MODIFY BY SONIA 馬德里改為mc01='2',TF-000500
         'stCon = ""
         stCon = "2"
      End If
      '2012/3/29 ADD BY SONIA
      If strPA10 = "" Then
         stCon1 = strSrvDate(1)
      Else
         stCon1 = strPA10
      End If
      '2012/3/29 END
      
      '2012/3/28 modify by sonia 加mc02條件
      'strExc(0) = "select 1 from membercountry where mc03='" & txtNation & "' and rownum<2" & stCon
      strExc(0) = "select 1 from membercountry where mc03='" & txtNation & "' and mc01='" & stCon & "'" & _
                  " and mc02=(select max(b.mc02) from membercountry b where b.mc01='" & stCon & "' and b.mc02<=" & stCon1 & " )"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         cmdMove(1).Enabled = True
         cmdMove(2).Enabled = True
         For i = 0 To lstCountry.ListCount - 1
            'Modify by Morgan 2004/7/26
            '國家代號移到最右邊
            'If txtNation = Left(lstCountry.List(i), 3) Then
            If txtNation = Right(lstCountry.List(i), 3) Then
               Exit For
            End If
         Next
         If i = lstCountry.ListCount Then
            'Modify by Morgan 2004/7/26
            '顯示內容修改
            'lstCountry.AddItem txtNation + vbTab + "     " + lblNationName
            lstCountry.AddItem PUB_GetCountryData(txtNation)
            txtNation = ""
            If lstCountry.ListCount = 1 Then lstCountry.ListIndex = 0
         Else
            ShowMsg lblNationName & MsgText(9199)
            txtNation_GotFocus
         End If
      Else
         MsgBox lblNationName & "不是指定國家!!"
         txtNation_GotFocus
      End If
   End If
ElseIf Index = 1 Then
   If lstCountry.ListIndex = -1 Then
      ShowMsg MsgText(8006)
   Else
      intlastIndex = lstCountry.ListIndex
      lstCountry.RemoveItem lstCountry.ListIndex
      If lstCountry.ListCount = 0 Then
         cmdMove(1).Enabled = False
         cmdMove(2).Enabled = False
      Else
         If intlastIndex = lstCountry.ListCount Then
            lstCountry.ListIndex = lstCountry.ListCount - 1
         Else
            lstCountry.ListIndex = intlastIndex
         End If
      End If
   End If
Else
   lstCountry.Clear
   cmdMove(1).Enabled = False
   cmdMove(2).Enabled = False
End If
txtNation.SetFocus
End Sub
'將ListBox中之輸入以逗號分隔合併為字串
Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer

If Index = 0 Then
   strCountry = ""
   If lstCountry.ListCount > 1 Then
      For i = 0 To lstCountry.ListCount - 2
         'Modify by Morgan 2004/7/26
         '國家代號移到最右邊
         'strCountry = strCountry + Left(lstCountry.List(i), 3) + ","
         strCountry = strCountry + Right(lstCountry.List(i), 3) + ","
      Next
      'Modify by Morgan 2004/7/26
      '國家代號移到最右邊
      'strCountry = strCountry + Left(lstCountry.List(i), 3)
      strCountry = strCountry + Right(lstCountry.List(i), 3)
      
'      Dim varTmp As Variant
'      varTmp = Split(strCountry, ",")
'      Sort varTmp, UBound(varTmp) - 1
'      strCountry = Join(varTmp, ",")
   ElseIf lstCountry.ListCount = 1 Then
      'Modify by Morgan 2004/7/26
      '國家代號移到最右邊
      'strCountry = Left(lstCountry.List(0), 3)
      strCountry = Right(lstCountry.List(0), 3)
   End If
End If
Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
 Dim varTmp As Variant, i As Integer, strTemp As String
   If Index = 0 Then
      'Modify by Morgan 2004/12/20 更新
      'strExc(0) = "201,203,204,205,206,207,208,209,211,212,213,214,216,217,220,227,228,230,231,232,235,318"
      'Modify by Morgan 2005/4/18 改抓MemberCountry
      'strExc(0) = "206,209,226,205,230,223,231,216,242,211,217,203,201,212,219,220,204,227,208,318,207,222,213,228,214,232,234,235"
      'Modify by Morgan 2005/9/12
      'strExc(0) = PUB_GetMemberCountry("1")
      strExc(0) = PUB_GetMemberCountry("1", strPA10)
      
   Else
      'Modify by Morgan 2004/10/13 加 242
      'Modify by Morgan 2004/12/20 更新
      'strExc(0) = "240,241,242,248,249"
      'Modify by Morgan 2005/4/18 改抓MemberCountry
      'strExc(0) = "248,251,241,240,249"
      'Modify by Morgan 2005/9/12
      'strExc(0) = PUB_GetMemberCountry("2")
      strExc(0) = PUB_GetMemberCountry("2", strPA10)
   End If
   'Modify by Morgan 2004/7/26
   '若清單無資料時以整批方式加入以提高效率
   If lstCountry.ListCount = 0 Then
      strTemp = PUB_GetCountryData(strExc(0))
      varTmp = Split(strTemp, ",")
      For i = 0 To UBound(varTmp)
         lstCountry.AddItem varTmp(i)
      Next
      cmdMove(1).Enabled = True
      cmdMove(2).Enabled = True
      If lstCountry.ListCount > 0 Then lstCountry.ListIndex = 0
   Else
      varTmp = Split(strExc(0), ",")
      For i = 0 To UBound(varTmp)
         txtNation.Text = Format(varTmp(i))
         cmdMove_Click 0
      Next
      txtNation.Text = ""
   End If
End Sub

'分析字串並存入ListBox
Private Sub Form_Load()
Dim i As Integer, varCountryTemp As Variant, strTemp As String

'Modify by Morgan 2004/7/26
'依照國家英文排序
'Dim objPublicData As Object
'Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
'varCountryTemp = Split(strCountry, ",")
'For i = 0 To UBound(varCountryTemp)
'       If objPublicData.GetNation(CStr(varCountryTemp(i)), strTemp) Then
'          lstCountry.AddItem varCountryTemp(i) + vbTab + "     " + strTemp
'       End If
'Next
strTemp = PUB_GetCountryData(strCountry)
varCountryTemp = Split(strTemp, ",")
For i = 0 To UBound(varCountryTemp)
   lstCountry.AddItem varCountryTemp(i)
Next
'2004/7/26 end

If lstCountry.ListCount > 0 Then
   lstCountry.ListIndex = 0
Else
   cmdMove(1).Enabled = False
   cmdMove(2).Enabled = False
End If
'edit by nickc 2007/02/02 不用 dll 了
'Set objPublicData = Nothing

If intPWhere <> 國外_CF Then
   Command1(0).Visible = False
   Command1(1).Visible = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   'Set frm880001 = Nothing
End Sub

Private Sub txtNation_Change()
Dim strTemp As String
'edit by nickc 2007/02/06 不用 dll 了
'Dim objPublicData As Object
'edit by nickc 2007/02/06 不用 dll 了
'Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
cmdMove(0).Enabled = False
If Len(txtNation) = 3 Then
   'edit by nickc 2007/02/06 不用 dll 了
   'If objPublicData.GetNation(txtNation, strTemp) Then
   If ClsPDGetNation(txtNation, strTemp) Then
      lblNationName = strTemp
      cmdMove(0).Enabled = True
   Else
      txtNation_GotFocus
   End If
Else
   lblNationName = ""
End If
'edit by nickc 2007/02/06 不用 dll 了
'Set objPublicData = Nothing
End Sub
Private Sub txtNation_GotFocus()
txtNation.SelStart = 0
txtNation.SelLength = Len(txtNation)
End Sub
