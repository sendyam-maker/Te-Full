VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_a_1 
   Caption         =   "工作進度資料維護－管制期限提醒－EMail智權人員"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   5490
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox txtEP28 
      Height          =   300
      Left            =   3870
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1725
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "取消(&C)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4320
      TabIndex        =   3
      Top             =   60
      Width           =   850
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&Y)"
      Height          =   400
      Index           =   1
      Left            =   3465
      TabIndex        =   2
      Top             =   60
      Width           =   850
   End
   Begin MSForms.TextBox txtContent 
      Height          =   1635
      Left            =   150
      TabIndex        =   1
      Top             =   2880
      Width           =   5085
      VariousPropertyBits=   -1466939365
      ScrollBars      =   2
      Size            =   "8969;2884"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseProperty 
      Height          =   255
      Left            =   1035
      TabIndex        =   27
      Top             =   969
      Width           =   1875
      VariousPropertyBits=   27
      Caption         =   "lblCaseProperty"
      Size            =   "3307;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   255
      Left            =   1035
      TabIndex        =   26
      Top             =   720
      Width           =   4275
      VariousPropertyBits=   27
      Caption         =   "lblCaseName"
      Size            =   "7541;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   255
      Index           =   10
      Left            =   2925
      TabIndex        =   25
      Top             =   1222
      Width           =   900
   End
   Begin VB.Label lblCP06 
      AutoSize        =   -1  'True
      Caption         =   "lblCP06"
      Height          =   255
      Left            =   3870
      TabIndex        =   24
      Top             =   1222
      Width           =   570
   End
   Begin VB.Label lblRecNo 
      AutoSize        =   -1  'True
      Caption         =   "lblRecNo"
      Height          =   255
      Left            =   1035
      TabIndex        =   23
      Top             =   210
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "收文號："
      Height          =   255
      Index           =   9
      Left            =   270
      TabIndex        =   22
      Top             =   210
      Width           =   720
   End
   Begin VB.Label lblSalesDate 
      AutoSize        =   -1  'True
      Caption         =   "lblSalesDate"
      Height          =   255
      Left            =   3870
      TabIndex        =   21
      Top             =   1475
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員管制日："
      Height          =   255
      Left            =   2385
      TabIndex        =   20
      Top             =   1475
      Width           =   1440
   End
   Begin VB.Label lblEP09 
      AutoSize        =   -1  'True
      Caption         =   "lblEP09"
      Height          =   255
      Left            =   1035
      TabIndex        =   19
      Top             =   1725
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   2565
      Left            =   60
      Top             =   2070
      Width           =   5325
   End
   Begin VB.Label lblCP48 
      AutoSize        =   -1  'True
      Caption         =   "lblCP48"
      Height          =   255
      Left            =   1035
      TabIndex        =   18
      Top             =   1475
      Width           =   570
   End
   Begin VB.Label lblEP06 
      AutoSize        =   -1  'True
      Caption         =   "lblEP06"
      Height          =   255
      Left            =   1035
      TabIndex        =   17
      Top             =   1222
      Width           =   555
   End
   Begin VB.Label lblCaseNo 
      AutoSize        =   -1  'True
      Caption         =   "lblCaseNo"
      Height          =   255
      Left            =   1035
      TabIndex        =   16
      Top             =   463
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "完稿日："
      Height          =   255
      Index           =   8
      Left            =   270
      TabIndex        =   15
      Top             =   1725
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "承辦期限："
      Height          =   255
      Index           =   7
      Left            =   90
      TabIndex        =   14
      Top             =   1475
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "齊備日："
      Height          =   255
      Index           =   6
      Left            =   270
      TabIndex        =   13
      Top             =   1222
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   255
      Index           =   5
      Left            =   90
      TabIndex        =   12
      Top             =   969
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   4
      Left            =   90
      TabIndex        =   11
      Top             =   716
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Index           =   3
      Left            =   90
      TabIndex        =   10
      Top             =   463
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "預定會稿日："
      Height          =   255
      Left            =   2745
      TabIndex        =   9
      Top             =   1725
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "內容："
      Height          =   255
      Index           =   2
      Left            =   450
      TabIndex        =   8
      Top             =   2550
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "主旨："
      Height          =   255
      Index           =   1
      Left            =   450
      TabIndex        =   7
      Top             =   2340
      Width           =   540
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "lblSubject"
      Height          =   255
      Left            =   1035
      TabIndex        =   6
      Top             =   2340
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "收件人："
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   5
      Top             =   2130
      Width           =   720
   End
   Begin VB.Label lblReveiver 
      Caption         =   "lblReveiver"
      Height          =   255
      Left            =   1035
      TabIndex        =   4
      Top             =   2130
      Width           =   1065
   End
End
Attribute VB_Name = "frm090201_a_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; lblCaseName、lblCaseProperty、txtContent
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
Dim stContent1 As String, stContent2 As String
Public stToNo As String
Public stCP10 As String

Private Function FormSave() As Boolean
   cnnConnection.BeginTrans
On Error GoTo ErrHnd
   
   If txtEP28 <> txtEP28.Tag Then
      strSql = "update engineerprogress set ep28=" & DBDATE(txtEP28) & ",EP30=nvl(EP30,0)+1 where ep02='" & lblRecNo & "'"
      cnnConnection.Execute strSql, intI
   End If
   strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
      " values ('" & strUserNum & "','" & stToNo & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & ChgSQL(lblSubject) & "','" & ChgSQL(txtContent) & "')"
   cnnConnection.Execute strSql, intI

   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

Private Function ChkEP28(pEP28 As Long, pEP06 As Long, pCP01 As String, pCP10 As String, pCP06 As Long, pCP48 As Long) As Boolean
   Dim strMsg As String, strDate As String, iDays As Integer
   
   If pEP28 < Val(strSrvDate(1)) Then
      MsgBox "預定會稿日不可早於系統日！"
      Exit Function
   
   ElseIf pEP28 < pEP06 Then
      MsgBox "預定會稿日不可早於齊備日！"
      Exit Function
   Else
      If pCP06 > 0 Then
         strDate = ""
         strMsg = ""
         If pCP01 = "P" Or InStr(",103,105,901,902,1002,1006,1201,1205,1206,1209,", "," & pCP10 & ",") > 0 Then
            strDate = CompWorkDay(1, CompDate(2, -1, pCP06), 1)
            strMsg = "前1個工作天"
         Else
            strDate = CompWorkDay(2, CompDate(2, -1, pCP06), 1)
            strMsg = "前2個工作天"
         End If
         
         If pEP28 > Val(strDate) And pEP28 > pEP06 Then
            MsgBox "預定會稿日不可晚於〔本所期限" & strMsg & "〕！", vbExclamation, "操作錯誤！"
            Exit Function
         End If
      End If
      
      If pCP48 > 0 Then
         strDate = ""
         strMsg = ""
         If pCP01 = "P" Or InStr(",103,105,901,902,1002,1006,1201,1205,1206,1209,", "," & pCP10 & ",") > 0 Then
            'Modified by Morgan 2012/7/19
            'iDays = 5
            iDays = 6
         Else
            'Modified by Morgan 2012/7/19
            'iDays = 10
            iDays = 12
         End If
         
         strDate = CompWorkDay(iDays, CompDate(2, 1, pCP48))
         strMsg = "+" & iDays & "個工作天"
         
         If strDate <> "" Then
            If pEP28 > Val(strDate) Then
               MsgBox "預定會稿日不可晚於〔承辦期限" & strMsg & "〕！", vbExclamation, "操作錯誤！"
               Exit Function
            End If
         End If
      End If
   End If
   
   ChkEP28 = True
End Function

Private Function TxtValidate() As Boolean
   If txtEP28 = "" Then
      MsgBox "預定會稿日不可空白!!"
      Exit Function
      
   ElseIf Not ChkDate(txtEP28) Then
      Exit Function
      
   ElseIf txtEP28 = txtEP28.Tag Then
      If DBDATE(txtEP28) <= DBDATE(lblSalesDate) Then
         MsgBox "預定會稿日沒有晚於智權人員管制日，不需要發EMail!!"
         Exit Function
      End If
      
   ElseIf ChkEP28(Val(DBDATE(txtEP28)), Val(DBDATE(lblEP06)), Left(lblCaseNo, InStr(lblCaseNo, "-") - 1), stCP10, Val(DBDATE(lblCP06)), Val(DBDATE(lblCP48))) = False Then
      Exit Function
   
   End If
   TxtValidate = True
End Function

Private Sub cmdOK_Click(Index As Integer)
   If Index = 1 Then
      If Not TxtValidate Then
         txtEP28.SetFocus
         txtEP28_GotFocus
         Exit Sub
      End If
      'Added by Lydia 2021/12/21 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         txtContent.SetFocus
         txtContent_GotFocus
         Exit Sub
      End If
      'end 2021/12/21
      
      If FormSave Then
         frm090201_a.bolCancel = False
      End If
   Else
      frm090201_a.bolCancel = True
   End If
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090201_a_1 = Nothing
End Sub

Public Sub SetContent(p_Date1 As String, Optional p_Date2 As String)
   'Modified by Morgan 2012/7/19 修改內容--柄佑
   If stContent1 = "" Then
      stContent1 = "敬閱者:" & vbCrLf & vbCrLf & _
         "　　因為工作安排因素，本案無法如期於 " & p_Date2 & " 會稿，依據估算本件預計可以在 "
   End If
   
   If stContent2 = "" Then
      stContent2 = " 會稿，請諒悉！" & vbCrLf & vbCrLf & _
        "　　若您與客戶有特別約定，煩請告知，以便於重新安排會稿時程。"
   End If
   txtContent = stContent1 & p_Date1 & stContent2
End Sub

Private Sub txtEP28_Change()
   SetContent Format(txtEP28, "@@@/@@/@@")
End Sub

Private Sub txtEP28_GotFocus()
   TextInverse txtEP28
   CloseIme
End Sub

Private Sub txtEP28_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2021/12/21
Private Sub txtContent_GotFocus()
   TextInverse txtContent
End Sub
