VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880009 
   BorderStyle     =   1  '單線固定
   Caption         =   "指定國註冊費"
   ClientHeight    =   5040
   ClientLeft      =   4665
   ClientTop       =   900
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8520
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "新增(&A)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2805
      TabIndex        =   2
      Top             =   240
      Width           =   912
   End
   Begin VB.TextBox txtMoney 
      Height          =   264
      Left            =   1080
      TabIndex        =   0
      Top             =   270
      Width           =   1572
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
      Height          =   2940
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1260
      Width           =   8295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   4800
      TabIndex        =   4
      Top             =   70
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   5736
      TabIndex        =   5
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.Label lblAgent 
      Height          =   300
      Left            =   2700
      TabIndex        =   9
      Top             =   600
      Width           =   5715
      VariousPropertyBits=   27
      Size            =   "10081;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   630
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "費用金額："
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   270
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "國家代號　國家名稱　　　　　　　　　　　費用金額　代理人"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   960
      Width           =   8235
   End
End
Attribute VB_Name = "frm880009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/16 改成Form2.0 ;lblAgent ；lstCountry因為換成Form2.0在點選時不是很準確，所以維持Form1.0元件
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit
'StrCountry存放指定國家  strMoneyCountry存放繳費國家 strMoney存放費用
Public strCountry As String, strMoneyCountry As String, strMoney As String
Public strFagentNo As String 'Added by Morgan 2020/8/7 代理人

Private Sub cmdEnter_Click()
   'Added by Morgan 2020/8/7
   Dim bCancel As Boolean
   If lstCountry.ListIndex = -1 Then
      MsgBox "請先點選國家！", vbExclamation
      Exit Sub
   End If
   'end 2020/8/7
   
   strExc(2) = Format(txtMoney, "#,###")
   If Len(strExc(2)) < 9 Then
      strExc(1) = Left(lstCountry.List(lstCountry.ListIndex), 25) & String(9 - Len(strExc(2)), " ") & strExc(2)
      'Added by Morgan 2020/8/7
      If Combo1 = "" Then
         MsgBox "代理人不可空白！", vbExclamation
         Combo1.SetFocus
         Exit Sub
      Else
         If lblAgent = "" Then
            Combo1_Validate bCancel
            If bCancel = True Then
               MsgBox "代理人輸入錯誤！", vbExclamation
               Combo1.SetFocus
               Exit Sub
            End If
         End If
         strExc(1) = strExc(1) & "　" & Combo1 & " " & lblAgent
      End If
      'end 2020/8/7
   Else
      'Modified by Morgan 2020/8/9 應該不可能超過9位數
      'strExc(1) = Left(lstCountry.List(lstCountry.ListIndex), 25) & strExc(2)
      MsgBox "金額太大！", vbCritical
      Exit Sub
      'end 2020/8/9
   End If
   
   
   lstCountry.List(lstCountry.ListIndex) = strExc(1)
   If lstCountry.ListIndex < lstCountry.ListCount - 1 Then
      lstCountry.ListIndex = lstCountry.ListIndex + 1
   End If
   txtMoney.SetFocus
   txtMoney_GotFocus
End Sub
Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer

If Index = 0 Then
   strMoneyCountry = ""
   strMoney = ""
   strFagentNo = "" 'Added by Morgan 2020/8/7
   For i = 0 To lstCountry.ListCount - 1
          If Val(Mid(lstCountry.List(i), 26)) > 0 Then
             strMoneyCountry = strMoneyCountry + Left(lstCountry.List(i), 3) + ","
             'Modified by Morgan 2020/8/7
             'strMoney = strMoney + Format(Mid(lstCountry.List(i), 26)) + ","
             strMoney = strMoney + Format(Mid(lstCountry.List(i), 26, 9)) + ","
             strExc(0) = Format(Mid(lstCountry.List(i), 36))
             intI = InStr(strExc(0), " ")
             If intI > 0 Then
               strExc(0) = Left(strExc(0), intI - 1)
             End If
             strFagentNo = strFagentNo + strExc(0) + ","
             'end 2020/8/7
          Else
             strMoneyCountry = strMoneyCountry + ","
             strMoney = strMoney + ","
             strFagentNo = strFagentNo + ","
          End If
   Next
   If Right(strMoneyCountry, 1) = "," Then
      strMoneyCountry = Mid(strMoneyCountry, 1, Len(strMoneyCountry) - 1)
      strMoney = Mid(strMoney, 1, Len(strMoney) - 1)
      strFagentNo = Mid(strFagentNo, 1, Len(strFagentNo) - 1) 'Added by Morgan 2020/8/7
   End If
End If
Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   Dim strCusTemp As String, strTemp As String
   
   lblAgent = ""
   If Combo1.Text <> "" Then
      strCusTemp = Combo1
      If GetFAgentName(strCusTemp, strTemp) = True Then
         lblAgent = strTemp
         For intI = 0 To Combo1.ListCount - 1
            If Combo1.List(intI) = strCusTemp Then
               Combo1.ListIndex = intI
               Exit For
            End If
         Next
         '代理人第一次輸入加入清單並檢查
         If intI = Combo1.ListCount Then
            Combo1 = strCusTemp
            Combo1.AddItem strCusTemp, 0
         
            intI = InStr(strCusTemp, "-")
            If intI > 0 Then
               strCusTemp = Left(strCusTemp, intI - 1)
            End If
            If PUB_CheckStatus(strCusTemp) = False Then
               Cancel = True
               lblAgent = ""
            Else
               strExc(0) = "select FA29 from Fagent where " & ChgFagent(strCusTemp) & " and FA29 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "" & RsTemp(0), vbExclamation, "代理人備註"
               End If
            End If
         End If
      Else
         Cancel = True
      End If
   End If
   
End Sub
'Added by Morgan 2020/8/7
Private Function GetFAgentName(ByRef pNo As String, ByRef pName As String) As Boolean
   '聯絡人
   If InStr(pNo, "-") > 0 Then
      GetFAgentName = ClsPDGetContact(pNo, pName)
   '代理人
   Else
      GetFAgentName = ClsPDGetAgent(pNo, pName)
   End If
End Function

'分析字串並存入ListBox
Private Sub Form_Load()

   Dim i As Integer, strTemp As String, j As Integer
   Dim varCountryTemp, varMoneyCountryTemp, varMoneyTemp, varFAgentNoTemp
   Dim strCusTemp As String
   
   varCountryTemp = Split(strCountry, ",")
   '若已輸過資料
   If strMoneyCountry <> "" Then
      varMoneyCountryTemp = Split(strMoneyCountry, ",")
      varMoneyTemp = Split(strMoney, ",")
      varFAgentNoTemp = Split(strFagentNo, ",")
   End If
   
   lstCountry.Clear
   For i = UBound(varCountryTemp) To 0 Step -1
      If ClsPDGetNation(CStr(varCountryTemp(i)), strTemp) Then
         strExc(1) = varCountryTemp(i) & String(7, " ") & Trim(strTemp)
         strExc(1) = strExc(1) & String(25 - Len(strExc(1)), "　")
         If strMoneyCountry <> "" Then
            If varMoneyCountryTemp(i) = varCountryTemp(i) Then
               strExc(2) = Format(varMoneyTemp(i), "#,###")
               If Len(strExc(2)) < 9 Then
                  strExc(1) = strExc(1) & String(9 - Len(strExc(2)), " ") & strExc(2)
                  'Added by Morgan 2020/8/7
                  If varFAgentNoTemp(i) <> "" Then
                     strCusTemp = varFAgentNoTemp(i)
                     If GetFAgentName(strCusTemp, strTemp) = True Then
                        strExc(1) = strExc(1) & "　" & strCusTemp & " " & strTemp
                     End If
                  End If
                  'end 2020/8/7
               Else
                  strExc(1) = strExc(1) & strExc(2)
               End If
            End If
         'Remove by Morgan 2009/3/12 不必預設,否則無法區分是否金額已修正--禧佩
         'Else
         '   strExc(0) = "select yf06,yf07 from patentyearfee where yf01='" & varCountryTemp(i) & "' and yf02='1' and yf03='Y00000000' and yf04='224' and yf05='1'"
         '   intI = 1
         '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         '   If intI = 1 Then
         '      strExc(2) = Format(Val("" & RsTemp(0)) + Val("" & RsTemp(1)), "#,###")
         '      If Len(strExc(2)) < 9 Then
         '         strExc(1) = strExc(1) & String(9 - Len(strExc(2)), " ") & strExc(2)
         '      Else
         '         strExc(1) = strExc(1) & strExc(2)
         '      End If
         '   End If
         End If
         lstCountry.AddItem strExc(1), 0
       End If
   Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   'Set frm880009 = Nothing
End Sub

Private Sub lstCountry_Click()
   'Modified by Morgan 2020/8/7
   'strExc(0) = Format(Mid(lstCountry.Text, 26))
   strExc(0) = Format(Mid(lstCountry.Text, 26, 9))
   'end 2020/8/7
   If Val(strExc(0)) > 0 Then
      txtMoney = strExc(0)
      'Added by Morgan 2020/8/7
      strExc(0) = Mid(lstCountry.Text, 36)
      If strExc(0) <> "" Then
         intI = InStr(strExc(0), " ")
         If intI > 0 Then
            lblAgent = Trim(Mid(strExc(0), intI))
            strExc(0) = Left(strExc(0), intI - 1)
         End If
         Combo1 = strExc(0)
         
      End If
      'end 2020/8/7
   Else
     txtMoney = ""
     Combo1 = "" 'Added by Morgan 2020/8/7
     lblAgent = "" 'Added by Morgan 2020/8/7
   End If
   If txtMoney.Visible Then
      txtMoney.SetFocus
      txtMoney_GotFocus
   End If
End Sub
Private Sub txtMoney_GotFocus()
   txtMoney.SelStart = 0
   txtMoney.SelLength = Len(txtMoney)
End Sub

Private Sub txtMoney_KeyDown(KeyCode As Integer, Shift As Integer)
   '下
   If KeyCode = 40 Then
      If lstCountry.ListIndex < lstCountry.ListCount - 1 Then
         lstCountry.ListIndex = lstCountry.ListIndex + 1
      End If
   ElseIf KeyCode = 38 Then
      If lstCountry.ListIndex > 0 Then
         lstCountry.ListIndex = lstCountry.ListIndex - 1
      End If
   End If
End Sub

