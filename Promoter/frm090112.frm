VERSION 5.00
Begin VB.Form frm090112 
   BorderStyle     =   1  '單線固定
   Caption         =   "組群統計"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   2670
      TabIndex        =   2
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2370
      MaxLength       =   7
      TabIndex        =   1
      Top             =   960
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   6
      Text            =   "990614"
      Top             =   960
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   0
      Top             =   510
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "備註：查名資料從990614開始統計"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   1500
      Width           =   2700
   End
   Begin VB.Line Line1 
      X1              =   1950
      X2              =   2760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "查名時間："
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   990
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "組群："
      Height          =   180
      Left            =   480
      TabIndex        =   4
      Top             =   540
      Width           =   540
   End
End
Attribute VB_Name = "frm090112"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/21 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
'add by nick 2004/11/10
Option Explicit
Dim strSql As String

Private Sub cmdOK_Click(Index As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim Cancel As Boolean
Dim strSql As String
Select Case Index
Case 0
        If Trim(txt1(0).Text) = "" Then
           strMsg = "組群不能空白, 請重新輸入"
           strTit = "資料檢核"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           txt1(0).SetFocus
           Exit Sub
        End If
        If Trim(txt1(2).Text) = "" Then
           strMsg = "查名時間不能空白, 請重新輸入"
           strTit = "資料檢核"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           txt1(2).SetFocus
           Exit Sub
        End If
        Dim objTxt As Object
        For Each objTxt In txt1
           If objTxt.Enabled = True Then
              Cancel = False
              txt1_Validate objTxt.Index, Cancel
              If Cancel = True Then
                 Exit Sub
              End If
           End If
        Next
        'Modified by Lydia 2015/06/16 改tmqclass
        'strSql = "select * from tmqctl where tmqc02='" & txt1(0).Text & "' "
         strSql = "select tmqc02 from tmqclass where tmqc01='" & txt1(0).Text & "' "
        CheckOC3
        With AdoRecordSet3
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount = 0 Then
                strMsg = "組群不存在, 請重新輸入"
                strTit = "資料檢核"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                txt1(0).SetFocus
                Exit Sub
            End If
        End With
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/13 清除查詢印表記錄檔欄位
        CheckOC3
        frm090112_1.strTM09 = txt1(0).Text
        frm090112_1.StrDateStart = ChangeTStringToWString(txt1(1).Text)
        frm090112_1.StrDateEnd = ChangeTStringToWString(txt1(2).Text)
        Set frm090112_1.UpForm = Me
        frm090112_1.InToCombo
        'frm090112_1.queryData
        frm090112_1.Show
        Me.Hide
Case 1
        Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me '將畫面移至中央
   txt1(2).Text = ChangeWStringToTString(ServerDate)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090112 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
TextInverse txt1(Index)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
If Trim(txt1(Index)) = "" Then Exit Sub
Dim strTit As String
Dim strMsg As String
Dim nResponse
Select Case Index
Case 0

Case 2
        If CheckIsTaiwanDate(txt1(2), False) = False Then
           strMsg = "日期不正確, 請重新輸入"
           strTit = "資料檢核"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           Cancel = True
           Exit Sub
        End If
Case Else
End Select
End Sub
