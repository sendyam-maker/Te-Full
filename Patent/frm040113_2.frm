VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040113_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "公文來函判發作業-退回程序"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4920
   StartUpPosition =   3  '系統預設值
   Begin VB.OptionButton Option1 
      Caption         =   "內部收文 -->"
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   3
      Top             =   1830
      Width           =   1275
   End
   Begin VB.OptionButton Option1 
      Caption         =   "修改定稿"
      Height          =   255
      Index           =   1
      Left            =   135
      TabIndex        =   2
      Top             =   1440
      Width           =   1365
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "取消(&C)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   4005
      TabIndex        =   1
      Top             =   3690
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   345
      Left            =   3150
      TabIndex        =   0
      Top             =   3690
      Width           =   800
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   2
      Left            =   1050
      TabIndex        =   14
      Top             =   765
      Width           =   3500
      VariousPropertyBits=   27
      Caption         =   "lblFM2(2)"
      Size            =   "6174;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   0
      Left            =   1050
      TabIndex        =   13
      Top             =   120
      Width           =   3500
      VariousPropertyBits=   27
      Caption         =   "lblFM2(0)"
      Size            =   "6174;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   1
      Left            =   1050
      TabIndex        =   12
      Top             =   435
      Width           =   3500
      VariousPropertyBits=   27
      Caption         =   "lblFM2(1)"
      Size            =   "6174;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   3
      Left            =   1050
      TabIndex        =   11
      Top             =   1080
      Width           =   3500
      VariousPropertyBits=   27
      Caption         =   "lblFM2(3)"
      Size            =   "6174;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   1275
      Left            =   1050
      TabIndex        =   4
      Top             =   2340
      Width           =   3750
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "6615;2249"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收文日："
      Height          =   180
      Index           =   3
      Left            =   135
      TabIndex        =   10
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   9
      Top             =   765
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   8
      Top             =   435
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   7
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "內部收文其他(910)本所期限7天,C類來函發文日改上111111"
      ForeColor       =   &H000040C0&
      Height          =   405
      Left            =   1485
      TabIndex        =   6
      Top             =   1830
      Width           =   3255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "退回意見："
      Height          =   180
      Left            =   135
      TabIndex        =   5
      Top             =   2370
      Width           =   900
   End
End
Attribute VB_Name = "frm040113_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/07 改成Form2.0 ; Label3(index)=>lblFM2(index)、Text1
'Created by Morgan 2019/1/9
Option Explicit
Public intRow As Integer
Public strCP09 As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, strCP83 As String, strPA09 As String

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click()
   If Option1(1).Value = False And Option1(2).Value = False Then
      MsgBox "請點選退回選項！", vbExclamation
      Exit Sub
   ElseIf Text1 = "" Then
      MsgBox "請輸入退回意見！", vbExclamation
      Exit Sub
   End If
      
   If FormSave = True Then
      Unload Me
   End If
End Sub

Private Sub Form_Activate()
   'Added by Morgan 2019/1/19 台灣案台可以內部收文--玲玲
   If Not (strCP01 = "P" And strPA09 = "000") Then
      Option1(2).Visible = False
      Label1.Visible = False
      Option1(1).Value = True
   End If
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
ReadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040113_2 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   Text1.SetFocus
End Sub

Private Function FormSave() As Boolean
   Dim stSQL As String
   Dim strSub As String, strContent As String
   Dim strBCP09 As String, strBCP12 As String, strBCP13 As String, strBCP06 As String
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
   '修改定稿
   If Option1(1).Value = True Then
      '更新退回意見,日期
      strSql = "update letterprogress set lp36=sysdate,lp37='" & ChgSQL(Text1) & "' where lp01='" & strCP09 & "'"
      cnnConnection.Execute strSql, intI
      '刪除客戶函
      PUB_DelFtpFile2 strCP09, " and instr(upper(cpp02),'.CUS.PDF')>0"
      strSql = "delete casepaperpdf where cpp01='" & strCP09 & "' and instr(upper(cpp02),'.CUS.PDF')>0"
      cnnConnection.Execute strSql, intI
      
   '內部收文(專利用)
   Else
      '更新判發人,日期,不通知
      strSql = "update letterprogress set lp04='" & strUserNum & "',lp05=" & strSrvDate(1) & ",lp10='N' where lp01='" & strCP09 & "'"
      cnnConnection.Execute strSql, intI
      
      '刪除客戶函
      PUB_DelFtpFile2 strCP09, " and instr(upper(cpp02),'.CUS.PDF')>0"
      strSql = "delete casepaperpdf where cpp01='" & strCP09 & "' and instr(upper(cpp02),'.CUS.PDF')>0"
      cnnConnection.Execute strSql, intI
      
      'C類來函發文日上111111
      strSql = "update caseprogress set cp27=19221111 where cp09='" & strCP09 & "'"
      cnnConnection.Execute strSql, intI
      
      '內部收文其他
      strBCP09 = AutoNo("B", 6)
      strBCP13 = PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04)
      strBCP12 = GetSalesArea(strBCP13)
      strBCP06 = PUB_GetWorkDay1(CompDate(2, 7, strSrvDate(1)), True)
      
      strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp08,cp09,cp10,cp12,cp13,cp14,cp16,cp17,cp18,cp20,cp26,cp32,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp43,cp64)" & _
         " select cp01,cp02,cp03,cp04," & strSrvDate(1) & "," & strBCP06 & ",cp08,'" & strBCP09 & "','910','" & strBCP12 & "','" & strBCP13 & "'" & _
         ",cp65,0,0,0,'N','N','N',cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp09,'" & ChgSQL(Text1) & "'" & _
         " from caseprogress where cp09='" & strCP09 & "'"
      cnnConnection.Execute strSql, intI
   End If
      
   '發EMail給程序
   strSub = "客戶函判發退回:" & lblFM2(0) & "(" & strCP09 & ")"
   strContent = "本所案號：" & lblFM2(0)
   strContent = strContent & vbCrLf & "案件名稱：" & lblFM2(1)
   strContent = strContent & vbCrLf & "案件性質：" & lblFM2(2)
   strContent = strContent & vbCrLf & "收文日：" & lblFM2(3)
   strContent = strContent & vbCrLf & "退回選項：" & IIf(Option1(1).Value = True, "修改定稿", "內部收文")
   strContent = strContent & vbCrLf & "　　意見：" & Text1
   
   strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
      " values( '" & strUserNum & "','" & strCP83 & "',to_char(sysdate,'yyyymmdd')" & _
      ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & ChgSQL(strContent) & "')"
   cnnConnection.Execute strSql, intI
      
   frm040113.UpdateGrid1 intRow
   cnnConnection.CommitTrans
   FormSave = True
   PUB_SendMailCache
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub ReadData()
   'Modify by Amy 2019/12/04 +Trademark,ServicePractice
   strExc(0) = "select cp01,cp02,cp03,cp04,cp83,Decode(pa01,null,Decode(tm01,null,sp09,tm10),pa09) as pa09 " & _
      "from caseprogress,patent,Trademark,ServicePractice " & _
      "where cp09='" & strCP09 & "'" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " And tm01(+)=cp01 And tm02(+)=cp02 And tm03(+)=cp03 And tm04(+)=cp04 " & _
      " And sp01(+)=cp01 And sp02(+)=cp02 And sp03(+)=cp03 And sp04(+)=cp04 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      strCP01 = .Fields("cp01")
      strCP02 = .Fields("cp02")
      strCP03 = .Fields("cp03")
      strCP04 = .Fields("cp04")
      strCP83 = .Fields("cp83")
      strPA09 = .Fields("pa09")
      End With
   End If
End Sub
