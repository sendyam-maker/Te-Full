VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880021 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外ID號數"
   ClientHeight    =   4605
   ClientLeft      =   450
   ClientTop       =   990
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4635
   Begin VB.CommandButton cmdMove 
      Caption         =   "清除(&C)"
      Height          =   400
      Index           =   2
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "刪除(&D)"
      Height          =   400
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   912
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "新增(&A)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox txtID 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   960
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1427
      Width           =   2295
   End
   Begin VB.ListBox lstData 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      ItemData        =   "frm880021.frx":0000
      Left            =   360
      List            =   "frm880021.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Width           =   3855
   End
   Begin MSForms.Label lblCust 
      Height          =   300
      Left            =   2040
      TabIndex        =   15
      Top             =   720
      Width           =   2340
      VariousPropertyBits=   27
      Size            =   "4128;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblTitle 
      Caption         =   "申請人1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblNation 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   1005
      Width           =   1995
   End
   Begin VB.Label lbl2 
      Caption         =   "申請國家："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1005
      Width           =   1455
   End
   Begin VB.Label lbl1 
      Caption         =   "客戶編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label Label4 
      Caption         =   "ID號碼"
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
      Left            =   1800
      TabIndex        =   10
      Top             =   2520
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   2400
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "ID號碼："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "客戶編號"
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
      Left            =   360
      TabIndex        =   7
      Top             =   2520
      Width           =   795
   End
End
Attribute VB_Name = "frm880021"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/16 改成Form2.0 ; lblCust
'Add by Lydia 2015/02/02 輸入國外ID號數(多筆)
Option Explicit
Public m_PrevF As Form
Public strXNo As String, strXNation As String
Dim strSPSign As String '分隔符號
Dim bolRec As Boolean '判斷先前是否有資料
Dim bolUpdate As Boolean '判斷是否有變動
Private Sub cmdOK_Click(Index As Integer)
   Dim i As Integer, strTmp As String
   Dim varTemp As Variant
   
   If Index = 0 Then '確定
      
      '先前是否有資料的情況
      If bolRec = True And lstData.ListCount = 0 Then
         If MsgBox("是否刪除全部ID號碼?", vbYesNo) = vbNo Then
            Exit Sub
         End If
      Else
         If lstData.ListCount = 0 And txtID.Text = "" Then
            Unload Me
            Exit Sub
         End If
      End If
On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
      
      strTmp = "delete from ApplicantForeignID where AFID01='" & strXNo & "' and AFID02='" & strXNation & "' "
      cnnConnection.Execute strTmp
      
      For i = 0 To lstData.ListCount - 1
         varTemp = Split(lstData.List(i), strSPSign)
         strTmp = "insert into ApplicantForeignID values ('" & strXNo & "','" & strXNation & "','" & Trim(varTemp(UBound(varTemp))) & "') "
         cnnConnection.Execute strTmp
      Next i

   cnnConnection.CommitTrans
   
   Else
      Unload Me
      Exit Sub

   End If
   
   Unload Me
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description
End Sub

'分析字串並存入ListBox
Private Sub Form_Load()
Dim i As Integer, varCountryTemp As Variant, strTemp As String
MoveFormToCenter Me

strSPSign = " " & vbVerticalTab
bolRec = False: bolUpdate = False
strXNo = strXNo + String(8 - Len(strXNo), "0")
lbl1.Caption = "客戶編號：" & strXNo
lbl2.Caption = "申請國家：" & strXNation
lstData.Clear

strTemp = "select * from ApplicantForeignID where afid01='" & strXNo & "' and afid02='" & strXNation & "' "

intI = 1
Set RsTemp = ClsLawReadRstMsg(intI, strTemp)
If intI = 1 Then
    bolRec = True
    RsTemp.MoveFirst
    For i = 0 To RsTemp.RecordCount - 1
       lstData.AddItem RsTemp.Fields("AFID01") & strSPSign & RsTemp.Fields("AFID02") & strSPSign & RsTemp.Fields("AFID03")
       RsTemp.MoveNext
    Next i
End If
If lstData.ListCount > 0 Then
   lstData.ListIndex = 0
End If

End Sub

Private Sub lstData_DblClick()
   If lstData.ListCount > 0 Then
      txtID = lstData.Text
   End If
End Sub

Private Sub cmdMove_Click(Index As Integer)
   Dim i As Integer, intlastIndex As Integer
   Dim varTemp As Variant
   
If Index = 0 Then '新增
    
   If txtID.Text = "" Then Exit Sub

   For i = 0 To lstData.ListCount - 1
      '改為用全型空白
       varTemp = Split(lstData.List(i), strSPSign)
       '檢查是否有相同ID
       If txtID = varTemp(UBound(varTemp)) Then
          MsgBox "ID號碼重複!!", vbExclamation
          Exit For
       End If
   Next
   '沒有重複時新增
   If i = lstData.ListCount Then
      lstData.AddItem strXNo & strSPSign & strXNation & strSPSign & Trim(txtID.Text)
      If lstData.ListCount = 1 Then lstData.ListIndex = 0
      txtID.Text = ""
      bolUpdate = True
   End If
   
ElseIf Index = 1 Then '刪除
   If lstData.ListIndex = -1 Then
      ShowMsg MsgText(8006)
   Else
      intlastIndex = lstData.ListIndex
      lstData.RemoveItem lstData.ListIndex
      If lstData.ListCount <> 0 Then
         If intlastIndex = lstData.ListCount Then
            lstData.ListIndex = lstData.ListCount - 1
         Else
            lstData.ListIndex = intlastIndex
         End If
         bolUpdate = True
      End If
   End If
Else
   txtID.Text = ""

End If
txtID.SetFocus
End Sub

Private Sub txtid_GotFocus()
   TextInverse txtID
   CloseIme
End Sub


