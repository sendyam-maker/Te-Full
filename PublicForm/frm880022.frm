VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880022 
   BorderStyle     =   1  '單線固定
   Caption         =   "寄發信函-往來記錄"
   ClientHeight    =   4245
   ClientLeft      =   450
   ClientTop       =   990
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5295
   Begin VB.TextBox txtCU20 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      TabIndex        =   12
      Top             =   1080
      Width           =   3930
   End
   Begin VB.TextBox txtNO 
      BackColor       =   &H8000000F&
      Height          =   300
      Left            =   1230
      MaxLength       =   9
      TabIndex        =   11
      Top             =   420
      Width           =   1275
   End
   Begin VB.ComboBox cboSort 
      Height          =   300
      ItemData        =   "frm880022.frx":0000
      Left            =   1230
      List            =   "frm880022.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   10
      Top             =   1410
      Width           =   3930
   End
   Begin VB.CheckBox Check1 
      Caption         =   "代表號"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1086
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Height          =   400
      Index           =   0
      Left            =   2820
      TabIndex        =   2
      Top             =   90
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面"
      Height          =   400
      Index           =   1
      Left            =   3870
      TabIndex        =   3
      Top             =   90
      Width           =   1200
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1230
      TabIndex        =   13
      Top             =   750
      Width           =   3930
      VariousPropertyBits=   679495707
      BackColor       =   -2147483648
      DisplayStyle    =   3
      Size            =   "6932;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtMeeting 
      Height          =   300
      Left            =   1230
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1740
      Width           =   3930
      VariousPropertyBits=   671105051
      Size            =   "6932;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstData 
      Height          =   1830
      Left            =   60
      TabIndex        =   1
      Top             =   2370
      Width           =   5205
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "9181;3228"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "名稱："
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   783
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "編號："
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "往來類別："
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   1464
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "會議名稱："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1770
      Width           =   930
   End
   Begin VB.Label Label3 
      Caption         =   "聯絡人：(可複選)"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   2100
      Width           =   2160
   End
End
Attribute VB_Name = "frm880022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/07 改成Form2.0 ; Combo1、lstData、txtMeeting
'Add By Sindy 2019/10/1
Option Explicit

Public m_strNo As String
Public m_PrevF As Form
Public m_PCC02 As String
Dim m_strOftPath As String
Dim bolActiveFirst As Boolean

Private Sub cboSort_Click()
   Dim iPos As Integer
   
   iPos = InStr(cboSort.Text, Chr(1))
   If iPos > 0 Then
      cboSort.Text = Left(cboSort.Text, iPos - 1)
   End If
End Sub

Private Sub cboSort_GotFocus()
   If cboSort.Locked = False Then
      CloseIme
      SendMessage cboSort.hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim ii As Integer, varTmp As Variant
'Dim Cancel As Boolean
Dim objOutLook As Object
Dim objMail As Object

   If Index = 0 Then '確定
    
On Error GoTo ErrHnd
      
      If cboSort = "" Then
         MsgBox "往來記錄不可空白！", vbExclamation
         cboSort.SetFocus
         Exit Sub
      ElseIf Left(cboSort, 3) = "B11" Then
         If Trim(txtMeeting) = "" Then
            MsgBox "會議名稱不可空白！", vbExclamation
            txtMeeting.SetFocus
            Exit Sub
         End If
      End If
      
      '呼叫新郵件：
      If Dir(m_strOftPath) = "" Then
         MsgBox "無郵件範本檔！" & vbCrLf & _
                "(" & m_strOftPath & ")", vbExclamation
         Exit Sub
      End If
      
      'Call m_PrevF_Unload
      Set objOutLook = CreateObject("Outlook.Application")
      '收件者.To
      '副本.cc
      '密件副本.BCC
      '主旨.Subject
      If Check1.Value = 1 Then 'Check1.Enabled = True And
         Set objMail = objOutLook.CreateItemFromTemplate(m_strOftPath)
         objMail.To = txtCU20.Text
         objMail.Subject = " [Our Ref:" & txtNo & "." & Left(cboSort, 3) & IIf(Trim(txtMeeting) <> "", "." & Trim(txtMeeting), "") & "]"
         objMail.Display
      End If
      '有聯絡人
      For ii = 0 To lstData.ListCount - 1
         If lstData.Selected(ii) = True Then
            Set objMail = objOutLook.CreateItemFromTemplate(m_strOftPath)
            varTmp = Split(lstData.List(ii), "-")
            If UBound(varTmp) > 0 Then
               objMail.To = Trim(varTmp(UBound(varTmp)))
               objMail.Subject = " [Our Ref:" & txtNo & "." & Left(cboSort, 3) & IIf(Trim(txtMeeting) <> "", "." & Trim(txtMeeting), "") & IIf(lstData.List(ii) <> "", "." & Left(lstData.List(ii), 2), "") & "]"
               objMail.Display
            End If
         End If
      Next ii
      
      Set objMail = Nothing
      Set objOutLook = Nothing
   
'   Else '回前畫面
'      Call m_PrevF_Unload
   End If
   Unload Me
   Call m_PrevF_Unload
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description
End Sub

Private Sub Form_Load()
Dim str_DownloadOftPath As String
   
   MoveFormToCenter Me
   
   '取得郵件範本
   str_DownloadOftPath = "$$TOT-000M31-0-02.oft"
   Call PUB_GetSampleFile(str_DownloadOftPath, Replace(Left(str_DownloadOftPath, Len(str_DownloadOftPath) - 4), "$$", ""))
   m_strOftPath = App.path & "\" & str_DownloadOftPath
End Sub

Public Function QueryData() As Boolean
Dim intRow As Integer
   
   If m_strNo = "" Then MsgBox "請傳入客戶/代理人/潛在客戶編號！", vbExclamation: Exit Function
   
   '編號
   txtNo = Left(Trim(m_strNo) & "00000000", 9)
   
   '檢查國內外權限
   'If Len(txtNO) <> 9 Then txtNO = txtNO & "0"
   If CheckSR12(txtNo) = False Then
      Screen.MousePointer = vbDefault
      'tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Me.m_PrevF.Show
      Unload Me
      Exit Function
   End If
   
   '名稱
   Select Case Left(m_strNo, 1)
      Case "X"
         strSql = "SELECT CU04,RTRIM(cu05||' '||cu88||' '||cu89||' '||cu90),CU06,CU20 FROM customer Where CU01='" & Left(txtNo, 8) & "' AND CU02='" & Right(txtNo, 1) & "'"
      Case "Y" '代理人
         strSql = "SELECT fa04,RTRIM(fa05||' '||fa63||' '||fa64||' '||fa65),fa06,fa16 FROM fagent where fa01='" & Left(txtNo, 8) & "' AND fa02='" & Right(txtNo, 1) & "'"
      Case "R"
         strSql = "SELECT PCU08,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06),PCU07,PCU18 FROM POTCUSTOMER WHERE PCU01='" & Left(txtNo, 8) & "' AND PCU02='" & Right(txtNo, 1) & "'"
   End Select
   '聯絡人
   Combo1.Clear
   txtCU20 = "" ': Check1.Enabled = False
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Trim(RsTemp.Fields(0)) <> "" Then
         Combo1.AddItem "中:" & RsTemp.Fields(0)
      End If
      If Trim(RsTemp.Fields(1)) <> "" Then
         Combo1.AddItem "英:" & RsTemp.Fields(1)
      End If
      If Trim(RsTemp.Fields(2)) <> "" Then
         Combo1.AddItem "日:" & RsTemp.Fields(2)
      End If
      If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
      '代表號
      If Trim(RsTemp.Fields(3)) <> "" Then
         txtCU20 = RsTemp.Fields(3)
         'Check1.Enabled = True
      End If
   End If
   
   '往來類別
   cboSort.Clear
   strExc(0) = "select ac02,ac03 from allcode where ac01='11' order by ac01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         cboSort.AddItem RsTemp.Fields("ac02") & " " & RsTemp.Fields("ac03")
         RsTemp.MoveNext
      Loop
   End If
   
   '聯絡人
   lstData.Clear
   strExc(0) = "select pcc02,pcc05,pcc03,pcc04,pcc08 from potcustcont where pcc01='" & Left(txtNo, 8) & "' order by pcc02 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   intRow = 0
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         strExc(10) = ""
         If Trim(" " & RsTemp.Fields(1)) <> "" Then
            strExc(10) = strExc(10) & RsTemp.Fields(1)
         End If
         If Trim(" " & RsTemp.Fields(2)) <> "" Then
            strExc(10) = strExc(10) & RsTemp.Fields(2)
         End If
         If Trim(" " & RsTemp.Fields(3)) <> "" Then
            strExc(10) = strExc(10) & RsTemp.Fields(3)
         End If
         lstData.AddItem RsTemp.Fields(0) & " " & strExc(10) & " - " & RsTemp.Fields(4)
         If m_PCC02 <> "" Then
            If m_PCC02 = RsTemp.Fields(0) Then
               lstData.Selected(intRow) = True
            End If
         End If
         intRow = intRow + 1
         RsTemp.MoveNext
      Loop
   Else
      If Trim(txtCU20.Text) <> "" Then
         Check1.Value = 1
      End If
   End If
   
   QueryData = True
End Function

Private Sub m_PrevF_Unload()
   m_PrevF.Show
   If m_PrevF.Name = "frm100114_1" Then
      m_PrevF.cmdState = 10
      m_PrevF.PubShowNextData
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm880022 = Nothing
End Sub
