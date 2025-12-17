VERSION 5.00
Begin VB.Form frm010021 
   BorderStyle     =   1  '單線固定
   Caption         =   "每月收文統計明細查詢"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3690
      TabIndex        =   2
      Top             =   180
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   1
      Top             =   180
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Left            =   1230
      MaxLength       =   5
      TabIndex        =   0
      Top             =   300
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "退件件數："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   5
      Left            =   210
      TabIndex        =   15
      Top             =   2130
      Width           =   1275
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   5
      Left            =   2580
      TabIndex        =   14
      Top             =   2130
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "智權人員收文件數："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   4
      Left            =   210
      TabIndex        =   13
      Top             =   780
      Width           =   2295
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   4
      Left            =   2580
      TabIndex        =   12
      Top             =   780
      Width           =   1755
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   3
      Left            =   2580
      TabIndex        =   11
      Top             =   1860
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "其他信件件數："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   3
      Left            =   210
      TabIndex        =   10
      Top             =   1860
      Width           =   1785
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   2
      Left            =   2580
      TabIndex        =   9
      Top             =   1590
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "客戶信件件數："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Index           =   2
      Left            =   210
      TabIndex        =   8
      Top             =   1590
      Width           =   1785
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Index           =   1
      Left            =   2580
      TabIndex        =   7
      Top             =   1320
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "國外信件件數："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Index           =   1
      Left            =   210
      TabIndex        =   6
      Top             =   1320
      Width           =   1785
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   2580
      TabIndex        =   5
      Top             =   1050
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "政府機關來函件數："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Top             =   1050
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "統計年月："
      Height          =   180
      Left            =   270
      TabIndex        =   3
      Top             =   330
      Width           =   900
   End
End
Attribute VB_Name = "frm010021"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/23 Form2.0已修改(無需修改)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/23 日期欄已修改
Option Explicit


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
        StrMenu txt1
Case 1
        Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1 = Trim(Val(Mid(Trim(ServerDate), 1, 6)) - 191100)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm010021 = Nothing
End Sub

Private Sub txt1_GotFocus()
InverseTextBox txt1
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub txt1_Validate(Cancel As Boolean)
If txt1 <> "" Then
    If CheckIsTaiwanDate(txt1 & "01", False) = False Then
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入民國年月不含/！", vbInformation, "輸入年月錯誤"
End Sub

Sub StrMenu(oDate As String)
Dim rsTmp As New ADODB.Recordset
Set rsTmp = New ADODB.Recordset
lbl1(0).Caption = "0"
lbl1(1).Caption = "0"
lbl1(2).Caption = "0"
lbl1(3).Caption = "0"
lbl1(4).Caption = "0"
Screen.MousePointer = vbHourglass
'add by nickc 2007/12/04 加入智權人員收文
'Modify By Sindy 2009/06/03 增加退件件數
strSql = "select sum(a1) b1,sum(a2) b2,sum(a3) b3,sum(a4) b4,sum(a5) b5,sum(a6) b6 from ("
strSql = strSql & "select nvl(count(*),0) A1,0 A2,0 A3,0 A4,0 A5,0 A6 from mailrec where mr02>=" & ChangeTStringToWString(oDate & "01") & " and mr02<=" & ChangeTStringToWString(oDate & "31") & "  "
strSql = strSql & "union select 0,nvl(count(*),0),0,0,0,0 from letterinput where li01>=" & ChangeTStringToWString(oDate & "01") & "  and li01<=" & ChangeTStringToWString(oDate & "31") & " and li08='3' "
strSql = strSql & "union select 0,0,nvl(count(*),0),0,0,0 from letterinput where li01>=" & ChangeTStringToWString(oDate & "01") & "  and li01<=" & ChangeTStringToWString(oDate & "31") & " and li08='4' "
strSql = strSql & "union select 0,0,0,nvl(count(*),0),0,0 from letterinput where li01>=" & ChangeTStringToWString(oDate & "01") & "  and li01<=" & ChangeTStringToWString(oDate & "31") & " and li08 in ('1','2') "
'Modified by Lydia 2022/09/27 排除TT-999999案號 and cp01||cp02<>'TT999999'
strSql = strSql & "union select 0,0,0,0,nvl(count(cp09),0),0 from caseprogress where cp66>=" & ChangeTStringToWString(oDate & "01") & " and cp66<=" & ChangeTStringToWString(oDate & "31") & " and substr(cp09,1,1)='A' and cp01||cp02<>'TT999999' "
'modify by sonia 2016/5/23 1.刪除記錄檔會有重覆的資料故加distinct(T-203082之申請有二筆) 2.剔除已刪除但又救回來的進度T-203085之申請故加入dd14 not in條件
'strSql = strSql & "union select 0,0,0,0,nvl(count(dd14),0),0 from datadeleterecord where dd25>=" & ChangeTStringToWString(oDate & "01") & " and dd25<=" & ChangeTStringToWString(oDate & "31") & " and substr(dd14,1,1)='A' and dd18 is not null  "
strSql = strSql & "union select 0,0,0,0,nvl(count(distinct dd14),0),0 from datadeleterecord where dd25>=" & ChangeTStringToWString(oDate & "01") & " and dd25<=" & ChangeTStringToWString(oDate & "31") & " and substr(dd14,1,1)='A' and dd18 is not null  "
'Modified by Lydia 2022/09/27 排除TT-999999案號 and cp01||cp02<>'TT999999'
strSql = strSql & "         and dd14 not in (select cp09 from caseprogress where cp66>=" & ChangeTStringToWString(oDate & "01") & " and cp66<=" & ChangeTStringToWString(oDate & "31") & " and substr(cp09,1,1)='A' and cp01||cp02<>'TT999999' ) "
'end 2016/5/23
strSql = strSql & "union select 0,0,0,0,0,nvl(count(*),0) from letterinput where li01>=" & ChangeTStringToWString(oDate & "01") & "  and li01<=" & ChangeTStringToWString(oDate & "31") & " and li08 in ('5') "

strSql = strSql & ") AA "
If rsTmp.State = 1 Then rsTmp.Close
rsTmp.CursorLocation = adUseClient
rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If rsTmp.RecordCount <> 0 Then
    lbl1(0) = CheckStr(rsTmp.Fields("b1"))
    lbl1(1) = CheckStr(rsTmp.Fields("b2"))
    lbl1(2) = CheckStr(rsTmp.Fields("b3"))
    lbl1(3) = CheckStr(rsTmp.Fields("b4"))
    lbl1(4) = CheckStr(rsTmp.Fields("b5"))
    'Add By Sindy 2009/06/03 增加退件件數
    lbl1(5) = CheckStr(rsTmp.Fields("b6"))
End If
Screen.MousePointer = vbDefault
End Sub
