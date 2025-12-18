VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_h 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文-電話聯絡單"
   ClientHeight    =   3800
   ClientLeft      =   270
   ClientTop       =   960
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3800
   ScaleWidth      =   8760
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1125
      MaxLength       =   3
      TabIndex        =   0
      Top             =   615
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1605
      MaxLength       =   6
      TabIndex        =   1
      Top             =   615
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2445
      MaxLength       =   1
      TabIndex        =   2
      Top             =   615
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   2685
      MaxLength       =   2
      TabIndex        =   3
      Top             =   615
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Left            =   3075
      TabIndex        =   4
      Top             =   570
      Width           =   800
   End
   Begin VB.TextBox txtCP113 
      Height          =   300
      Left            =   3285
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2610
      Width           =   600
   End
   Begin VB.TextBox txtCP14 
      Height          =   300
      Left            =   4995
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   7
      Top             =   2610
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7455
      TabIndex        =   10
      Top             =   70
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   6435
      TabIndex        =   9
      Top             =   70
      Width           =   975
   End
   Begin VB.TextBox txtCP27 
      Height          =   300
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2610
      Width           =   1095
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   1125
      TabIndex        =   34
      Top             =   915
      Width           =   6780
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11959;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   615
      Left            =   1125
      TabIndex        =   8
      Top             =   3000
      Width           =   7410
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13070;1085"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCP14T 
      Height          =   255
      Left            =   6150
      TabIndex        =   33
      Top             =   2633
      Width           =   1710
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "3016;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   3
      Left            =   4560
      TabIndex        =   32
      Top             =   1890
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   4
      Left            =   1140
      TabIndex        =   31
      Top             =   2160
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   30
      Top             =   1350
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   5
      Left            =   4560
      TabIndex        =   29
      Top             =   2160
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   9
      Left            =   3645
      TabIndex        =   28
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "本所期限:"
      Height          =   180
      Index           =   8
      Left            =   180
      TabIndex        =   27
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Index           =   7
      Left            =   3645
      TabIndex        =   26
      Top             =   1890
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Index           =   3
      Left            =   3660
      TabIndex        =   25
      Top             =   1350
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   24
      Top             =   960
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   23
      Top             =   1350
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   22
      Top             =   660
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   6
      Left            =   180
      TabIndex        =   21
      Top             =   1890
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   2
      Left            =   1140
      TabIndex        =   20
      Top             =   1890
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   1
      Left            =   4560
      TabIndex        =   19
      Top             =   1350
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   18
      Top             =   1620
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文日:"
      Height          =   180
      Index           =   5
      Left            =   3660
      TabIndex        =   17
      Top             =   1620
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   16
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   1
      Left            =   4560
      TabIndex        =   15
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   12
      Left            =   2430
      TabIndex        =   14
      Top             =   2655
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Index           =   11
      Left            =   4275
      TabIndex        =   13
      Top             =   2655
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   8520
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   8520
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Index           =   13
      Left            =   180
      TabIndex        =   12
      Top             =   3060
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Index           =   10
      Left            =   180
      TabIndex        =   11
      Top             =   2655
      Width           =   585
   End
End
Attribute VB_Name = "frm060104_h"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/24 改成Form2.0 ; txtCP64、Label3(6)=>lblCP14T、Combo1
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Created by Morgan 2012/5/22
Option Explicit

Dim pa() As String
Dim m_CP10 As String, m_CP09 As String, m_ST16 As String

Private Sub Command1_Click()
Dim Rs As New ADODB.Recordset

   If Text1(3) = "" Then Text1(3) = "0"
   If Text1(4) = "" Then Text1(4) = "00"
   
   strExc(0) = "select cp09,st16,cp14 from caseprogress,staff where cp01='" & Text1(1) & "' and cp02='" & Text1(2) & "' and cp03='" & Text1(3) & "' and cp04='" & Text1(4) & "' and cp10='945' and cp27||cp57 is null and st01(+)=cp14"
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modify By Sindy 2024/4/17 排除員編第4號是9的人員(支援人員) + And PUB_NeedChkFCPST16("" & Rs("cp14")) = True
      If Pub_StrUserSt03 <> "M51" And m_ST16 <> Rs("st16") And PUB_NeedChkFCPST16("" & Rs("cp14")) = True Then
         MsgBox "工程師組別不同不可發文！"
      Else
         Label3(0) = Rs(0)
         ReadPatent
      End If
   Else
      MsgBox "無符合資料！"
   End If
   
   Set Rs = Nothing
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   FormClear
   ReDim pa(TF_PA)
   m_ST16 = PUB_GetStaffST16(strUserNum)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060104_h = Nothing
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         If TxtValidate = False Then Exit Sub
         If FormSave = False Then
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
            Exit Sub
         Else
            FormClear True
            Text1(1).SetFocus
         End If
      Case 1
         Unload Me
   End Select
   
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   txtCP27_Validate bCancel
   If bCancel = True Then Exit Function
'   txtCP14_Validate bCancel
'   If bCancel = True Then Exit Function
   txtCP113_Validate bCancel
   If bCancel = True Then Exit Function
   
    'Added by Lydia 2021/09/24 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If
   
   TxtValidate = True

End Function

Private Function FormSave() As Boolean
   
On Error GoTo CheckingErr
   cnnConnection.BeginTrans
   strSql = "Update Caseprogress set cp27=" & DBDATE(txtCP27) & ",cp14='" & txtCP14 & "',cp113=" & IIf(txtCP113 = "", "NULL", txtCP113) & ",cp64='" & ChgSQL(txtCP64) & "' where cp09='" & m_CP09 & "'"
   cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   
   FormSave = True
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function


Private Sub ReadPatent()
   pa(1) = Text1(1)
   pa(2) = Text1(2)
   pa(3) = Text1(3)
   pa(4) = Text1(4)
   Select Case pa(1)
      Case "FCP"
         If ClsPDReadPatentDatabase(pa(), 國外_FC) Then
            Label2(0) = pa(11)
            Label2(1) = pa(10)
            AddCboName Combo1, pa(5), pa(6), pa(7)
         End If
      Case "FG"
         If PUB_ReadServicePracticeDatabase(pa(), 國外_FC) Then
            Label2(0) = pa(11)
            Label2(1) = pa(10)
            AddCboName Combo1, pa(5), pa(6), pa(7)
         End If
   End Select
   strExc(0) = "select cp05,cp10,cpm03,cp08,cp06,cp07,cp14,cp113,cp64,st02" & _
      " from caseprogress,casepropertymap,staff where cp09='" & Label3(0) & "'" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp14"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
   With RsTemp
      Label3(1) = .Fields("cp05") - 19110000
      Label3(2) = .Fields("cp10") & " " & .Fields("cpm03")
      m_CP10 = .Fields("cp10")
      Label3(3) = "" & .Fields("cp08")
      If Not IsNull(.Fields("cp06")) Then
         Label3(4) = .Fields("cp06") - 19110000
      End If
      If Not IsNull(.Fields("cp07")) Then
         Label3(5) = .Fields("cp07") - 19110000
      End If
      txtCP14 = "" & .Fields("cp14")
      'modify by sonia 2015/9/21
      'Label3(6) = "" & .Fields("st02")
      txtCP14_Validate False
      'end 2015/9/21
      txtCP113 = "" & .Fields("cp113")
      txtCP64 = "" & .Fields("cp64")
   End With
   m_CP09 = Label3(0)
   FormEnable True
   End If
End Sub

Private Sub Text1_Change(Index As Integer)
   If m_CP09 <> "" Then
      m_CP09 = ""
      FormClear
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
   CloseIme
End Sub

Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   Cancel = Not PUB_CheckCP113(txtCP113, pa(1), m_CP10, txtCP14)
End Sub

Private Sub txtCP14_Change()
   'Modified by Lydia 2021/09/24 Label3(6)=>lblCP14T
   lblCP14T = ""
End Sub

Private Sub txtCP14_GotFocus()
   TextInverse txtCP14
End Sub

Private Sub txtCP14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCP14_Validate(Cancel As Boolean)
   If txtCP14 = "" Then
      MsgBox "承辦人不可空白 !", vbCritical
      Cancel = True
   Else
      'ADD BY SONIA 2015/9/21 承辦人為外專程序時,改為操作人員
      txtCP14 = GetFCPUser(txtCP14)
      'END 2015/9/21
      'Modified by Lydia 2021/09/24 Label3(6)=>lblCP14T
      lblCP14T = GetStaffName(txtCP14, True)
   End If
End Sub

Private Sub txtCP27_GotFocus()
   TextInverse txtCP27
End Sub

Private Sub txtCP27_Validate(Cancel As Boolean)
   If txtCP27 = "" Then
      MsgBox "發文日不可空白 !", vbCritical
      Cancel = True
   ElseIf Not ChkDate(txtCP27) Then
      txtCP27_GotFocus
      Cancel = True
   End If
End Sub

Private Sub FormClear(Optional pbolAll As Boolean)
   Dim oLabel As LABEL
   
   Combo1.Clear
   For Each oLabel In Label2
      oLabel.Caption = ""
   Next
   For Each oLabel In Label3
      oLabel.Caption = ""
   Next
   txtCP27 = strSrvDate(2)
   txtCP113.Text = ""
   txtCP14 = ""
   txtCP64 = ""
   
   If pbolAll Then
      Text1(1) = ""
      Text1(2) = ""
      Text1(3) = ""
      Text1(4) = ""
      m_CP09 = ""
   End If
   FormEnable False
End Sub

Private Sub FormEnable(pValue As Boolean)
   txtCP27.Enabled = pValue
   txtCP113.Enabled = pValue
   cmdOK(0).Enabled = pValue
End Sub


