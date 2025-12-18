VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_12 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶發明人資料查詢"
   ClientHeight    =   4500
   ClientLeft      =   200
   ClientTop       =   1040
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   9310
   Begin VB.CommandButton CmdOk 
      Caption         =   "結束"
      Height          =   400
      Index           =   1
      Left            =   8325
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "回前畫面"
      Height          =   400
      Index           =   0
      Left            =   7140
      TabIndex        =   0
      Top             =   70
      Width           =   1185
   End
   Begin MSForms.TextBox txt1 
      Height          =   330
      Index           =   6
      Left            =   1530
      TabIndex        =   8
      Top             =   4620
      Visible         =   0   'False
      Width           =   7695
      VariousPropertyBits=   -1466941409
      ScrollBars      =   2
      Size            =   "13582;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   360
      Index           =   5
      Left            =   1512
      TabIndex        =   7
      Top             =   3497
      Width           =   7700
      VariousPropertyBits=   -1467989985
      ScrollBars      =   2
      Size            =   "13582;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   360
      Index           =   4
      Left            =   1512
      TabIndex        =   6
      Top             =   3094
      Width           =   7700
      VariousPropertyBits=   -1467989985
      ScrollBars      =   2
      Size            =   "13582;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   360
      Index           =   3
      Left            =   1512
      TabIndex        =   5
      Top             =   2691
      Width           =   7700
      VariousPropertyBits=   -1467989985
      ScrollBars      =   2
      Size            =   "13582;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   360
      Index           =   2
      Left            =   1536
      TabIndex        =   4
      Top             =   2288
      Width           =   7700
      VariousPropertyBits=   -1467989985
      ScrollBars      =   2
      Size            =   "13582;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   360
      Index           =   1
      Left            =   1512
      TabIndex        =   3
      Top             =   1885
      Width           =   7700
      VariousPropertyBits=   -1467989985
      ScrollBars      =   2
      Size            =   "13582;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   360
      Index           =   0
      Left            =   1512
      TabIndex        =   2
      Top             =   1482
      Width           =   7700
      VariousPropertyBits=   -1467989985
      ScrollBars      =   2
      Size            =   "13582;635"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   255
      Left            =   1530
      TabIndex        =   25
      Top             =   4200
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "法定代表人職稱："
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   24
      Top             =   4200
      Width           =   1440
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   23
      Top             =   1184
      Width           =   7700
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13582;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      Height          =   255
      Left            =   1185
      TabIndex        =   22
      Top             =   1184
      Width           =   225
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   1536
      TabIndex        =   21
      Top             =   3900
      Width           =   7704
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13589;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   20
      Top             =   886
      Width           =   7700
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13582;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   19
      Top             =   588
      Width           =   7704
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "13589;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國籍："
      Height          =   255
      Index           =   9
      Left            =   990
      TabIndex        =   18
      Top             =   3900
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發明人地址(英)："
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   3094
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發明人地址(中)："
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   2691
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發明人名稱(日)："
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   2288
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發明人名稱(英)："
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   1885
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發明人名稱(中)："
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1482
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發明人代號："
      Height          =   255
      Index           =   3
      Left            =   405
      TabIndex        =   12
      Top             =   886
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人編號："
      Height          =   255
      Index           =   2
      Left            =   405
      TabIndex        =   11
      Top             =   588
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "代表人："
      Height          =   180
      Index           =   1
      Left            =   720
      TabIndex        =   10
      Top             =   4710
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發明人地址(日)："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   3497
      Width           =   1380
   End
End
Attribute VB_Name = "frm100101_12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; lbl1(index)、txt1(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
Option Explicit

'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer


'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 1
     fnCloseAllFrm100
Case Else
End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
''92.04.16 nick 以下無效
'Select Case Index
'Case 0
'     Me.Hide
'Case 1
'     bolToEndByNick = True
'     Unload Me
'     Exit Sub
'End Select
End Sub

Sub StrMenu()
Dim strSql  As String, s As Integer
Dim Str01 As String, Str02 As String   ', str03 As String, str04 As String
Dim strArr(11) As String, i As Integer, StrOk(2) As String, StrOkTxt(6) As String

Str01 = ""
Str02 = ""
pub_QL05 = ";客戶發明人編號：" & Me.Tag & "(基本資料)" 'Add By Sindy 2025/8/13
Me.Tag = Mid(Me.Tag, 1, 8) & Mid(Me.Tag, 10, 2)  '2008/9/2 ADD BY SONIA 取消畫面發明人編號之'-'
Str01 = Mid(Trim(Me.Tag), 1, 8)
Str02 = Mid(Trim(Me.Tag), 9, 2)
LBL1(1).Caption = Str02
'顯示申請人編號及名稱
strSql = "select NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) FROM CUSTOMER WHERE CU01='" & Str01 & "' AND CU02='0'"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If Not IsNull(adoRecordset.Fields(0)) Then
         LBL1(0).Caption = Str01 + "  " + adoRecordset.Fields(0)
    Else
         LBL1(0).Caption = Str01
    End If
End If
CheckOC
'查詢發明人資料
strSql = "SELECT * FROM INVENTOR WHERE IN01='" & Mid(Me.Tag, 1, 8) & "' AND IN02='" & Mid(Me.Tag, 9, 2) & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
     If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/8/13
     If IsNull(adoRecordset.Fields(2)) Then
        LBL1(3).Caption = ""
     Else
        LBL1(3).Caption = adoRecordset.Fields(2)
     End If
     If IsNull(adoRecordset.Fields(3)) Then
          txt1(0).Text = ""
     Else
          txt1(0).Text = adoRecordset.Fields(3)
     End If
     If IsNull(adoRecordset.Fields(4)) Then
          txt1(1).Text = ""
     Else
          txt1(1).Text = adoRecordset.Fields(4)
     End If
     If IsNull(adoRecordset.Fields(5)) Then
          txt1(2).Text = ""
     Else
          txt1(2).Text = adoRecordset.Fields(5)
     End If
     If IsNull(adoRecordset.Fields(6)) Then
          txt1(3).Text = ""
     Else
          txt1(3).Text = adoRecordset.Fields(6)
     End If
     If IsNull(adoRecordset.Fields(7)) Then
          txt1(4).Text = ""
     Else
          txt1(4).Text = adoRecordset.Fields(7)
     End If
     If IsNull(adoRecordset.Fields(8)) Then
          txt1(5).Text = ""
     Else
          txt1(5).Text = adoRecordset.Fields(8)
     End If
     If IsNull(adoRecordset.Fields(10)) Then
          LBL1(2).Caption = ""
     Else
          LBL1(2).Caption = adoRecordset.Fields(10)
     End If
     If IsNull(adoRecordset.Fields(9)) Then
          txt1(6).Text = ""
     Else
          txt1(6).Text = adoRecordset.Fields(9)
     End If
     If IsNull(adoRecordset.Fields("in12")) Then
          Label3.Caption = ""
     Else
          Label3.Caption = adoRecordset.Fields("in12")
     End If
Else
     If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
     ShowNoData
     Screen.MousePointer = vbDefault
     '920416 nick
     'Me.Hide
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
     Exit Sub
End If
CheckOC
strSql = "SELECT NA03 FROM NATION WHERE NA01='" & LBL1(2).Caption & "'"
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If Not IsNull(adoRecordset.Fields(0)) Then
        LBL1(2).Caption = LBL1(2).Caption + "  " + adoRecordset.Fields(0)
    End If
End If
CheckOC

End Sub

Private Sub Form_Load()
bolToEndByNick = False
   MoveFormToCenter Me
'92.04.16 nick
cmdState = -1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100101_12 = Nothing
End Sub

'Added by Lydia 2016/10/29 修正Win7 輸入法問題
Private Sub txt1_GotFocus(Index As Integer)
   OpenIme
End Sub
