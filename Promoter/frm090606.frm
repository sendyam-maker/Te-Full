VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm090606 
   BorderStyle     =   1  '單線固定
   Caption         =   "每月目次重編作業"
   ClientHeight    =   1380
   ClientLeft      =   4860
   ClientTop       =   3300
   ClientWidth     =   3600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   3600
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   228
      Left            =   60
      TabIndex        =   3
      Top             =   1044
      Width           =   3528
      _ExtentX        =   6218
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   2364
      TabIndex        =   1
      Top             =   24
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1584
      TabIndex        =   0
      Top             =   24
      Width           =   756
   End
   Begin VB.Label lbl1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   108
      TabIndex        =   2
      Top             =   504
      Width           =   2952
   End
End
Attribute VB_Name = "frm090606"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/17 改成Form2.0 (無)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim i As Integer, k As Integer, strTemp3 As String, s As Integer

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Screen.MousePointer = vbHourglass
     Me.Enabled = False
     Process
     Me.Enabled = True
     Screen.MousePointer = vbDefault
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()

'911107 nick
On Error GoTo CheckingErr

'StrSQL6 + " and CP14='" & strUserNum & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Mid(GetTodayDate, 1, 6) & "01 AND CP27<=" & Mid(GetTodayDate, 1, 6) & "31 and cp57 is null ) or (CP57>=" & Mid(GetTodayDate, 1, 6) & "01 AND CP57<=" & Mid(GetTodayDate, 1, 6) & "31 and cp27 is null))) and cp05>=19980101"
'Modify By Cheng 2004/01/02
'strSQL = "SELECT EP01,EP02,ep05,EP10,CP05 FROM ENGINEERPROGRESS,CASEPROGRESS WHERE cp09=ep02(+) AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Mid(GetTodayDate, 1, 6) & "01 AND CP27<=" & Mid(GetTodayDate, 1, 6) & "31 and cp57 is null ) or (CP57>=" & Mid(GetTodayDate, 1, 6) & "01 AND CP57<=" & Mid(GetTodayDate, 1, 6) & "31 and cp27 is null))) AND EP05 IS NOT NULL and cp05>=19980101 "
'edit by nickc 2005/05/13
'strSQL = "SELECT EP01,EP02,ep05,EP10,CP05 FROM ENGINEERPROGRESS,CASEPROGRESS WHERE cp09=ep02(+) AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP27<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp57 is null ) or (CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP57<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp27 is null))) AND EP05 IS NOT NULL and cp05>=19980101 "
strSql = "SELECT EP01,EP02,ep05,EP10,CP05 FROM ENGINEERPROGRESS,CASEPROGRESS WHERE cp09=ep02(+) AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP27<=" & Mid(strSrvDate(1), 1, 6) & "31 ) or (CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP57<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp27 is null))) AND EP05 IS NOT NULL and cp05>=19980101 "

'End
'strSQL = strSQL & " UNION all  SELECT EP01,EP02,EP05,EP10 FROM ENGINEERPROGRESS,CASEPROGRESS WHERE CP09=ep02(+) AND cp57 is null and CP27>=" & Mid(GetTodayDate, 1, 6) & "01  AND EP05 IS NOT NULL and cp05>=19980101 "
'strSQL = strSQL & " UNION all  SELECT EP01,EP02,EP05,EP10 FROM ENGINEERPROGRESS,CASEPROGRESS WHERE CP09=ep02(+) AND CP57>=" & Mid(GetTodayDate, 1, 6) & "01 and CP27 is null AND EP05 IS NOT NULL and cp05>=19980101
strSql = strSql & " ORDER BY EP05,ep01,cp05 "
'SELECT EP01,EP02,ep05,EP10 FROM ENGINEERPROGRESS,CASEPROGRESS WHERE CP09=EP02(+) AND ((CP57 IS NULL AND CP27 IS NULL) OR (CP27>=20010701 OR CP57>=20010701)) AND EP10 IS NULL AND EP05 IS NOT NULL and cp05>=19980101
'ORDER BY EP05,EP01
CheckOC
i = 0
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PB1.Min = 0
        PB1.max = .RecordCount
        PB1.Value = 0
        strTemp3 = CheckStr(.Fields(2))
        i = 0
        Do While .EOF = False
            PB1.Value = PB1.Value + 1
            If strTemp3 <> CheckStr(.Fields(2)) Then
                i = 1
                strTemp3 = CheckStr(.Fields(2))
            Else
                i = i + 1
            End If
             cnnConnection.Execute "UPDATE ENGINEERPROGRESS SET EP01=" & i & " WHERE EP02='" & CheckStr(.Fields(1)) & "' "
            .MoveNext
            DoEvents
        Loop
    End If
End With
CheckOC
s = MsgBox("已重編完畢!!", , "OK")
PB1.Value = 0
'911107 nick
     Exit Sub
CheckingErr:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
lbl1.Caption = Mid(ChangeTStringToTDateString(ChangeWStringToTString(GetTodayDate)), 1, InStr(1, ChangeTStringToTDateString(ChangeWStringToTString(GetTodayDate)), "/") - 1) & " 年 " & Format(ChangeWStringToWDateString(GetTodayDate), "mm") & " 月 "
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090606 = Nothing
End Sub
