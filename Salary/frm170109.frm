VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm170109 
   BorderStyle     =   1  '單線固定
   Caption         =   "福利金轉檔"
   ClientHeight    =   3840
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5628
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5628
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   7.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1608
      ItemData        =   "frm170109.frx":0000
      Left            =   192
      List            =   "frm170109.frx":0002
      TabIndex        =   7
      Top             =   2064
      Width           =   5184
   End
   Begin VB.TextBox txtYEAR 
      Height          =   270
      Left            =   1392
      MaxLength       =   3
      TabIndex        =   0
      Top             =   576
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "轉檔(&T)"
      Height          =   405
      Left            =   3096
      TabIndex        =   1
      Top             =   100
      Width           =   1065
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   405
      Left            =   4296
      TabIndex        =   2
      Top             =   100
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   216
      Left            =   192
      TabIndex        =   6
      Top             =   1776
      Width           =   4692
      _ExtentX        =   8276
      _ExtentY        =   381
      _Version        =   393216
      Appearance      =   1
      Max             =   20
   End
   Begin VB.Label Label3 
      Caption         =   "最遲須於12月薪資計算前轉檔，否則會無法及時扣繳併入三節獎金之補充保費!!!!"
      ForeColor       =   &H000000FF&
      Height          =   516
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   24
      Width           =   2328
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "　　所得或獎金年度固定為福利金年度的次年"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   216
      TabIndex        =   9
      Top             =   1416
      Width           =   3600
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "0/1"
      Height          =   180
      Left            =   5016
      TabIndex        =   8
      Top             =   1800
      Width           =   216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "　　智權公司轉入每月獎金(併入三節獎金)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   216
      TabIndex        =   5
      Top             =   1176
      Width           =   3360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PS：智慧所及法律所轉入其他各類所得資料(格式92 給付項目8A)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   216
      TabIndex        =   4
      Top             =   936
      Width           =   5052
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "福利金年度：                   (ex:112)"
      Height          =   180
      Index           =   1
      Left            =   228
      TabIndex        =   3
      Top             =   624
      Width           =   2568
   End
End
Attribute VB_Name = "frm170109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2023/11/30
Option Explicit

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click()
   If TxtValidate = True Then
      If MsgBox("是否已至查詢系統確認資料無誤？", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
         Process
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtYEAR = strSrvDate(2) \ 10000 - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170109 = Nothing
End Sub

Private Sub txtYEAR_GotFocus()
   TextInverse txtYEAR
End Sub

Private Sub txtYEAR_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Function TxtValidate() As Boolean
   If txtYEAR = "" Then
      MsgBox "請輸入年度！", vbExclamation
      txtYEAR.SetFocus
      Exit Function
   ElseIf Val(txtYEAR) < 100 Or Val(txtYEAR) > Val(strSrvDate(2) \ 10000) Then
      MsgBox "年度輸入錯誤！", vbCritical
      txtYEAR.SetFocus
      Exit Function
      
   ElseIf Val(txtYEAR + 1911) >= Left(strSrvDate(1), 4) Then
      MsgBox "年度必須小於系統年！", vbCritical
      txtYEAR.SetFocus
      Exit Function
      
   End If
   TxtValidate = True
End Function

Private Sub Process()
   Dim stMB01 As String
   Dim stDate As String, stYear As String, stMon As String '給付日期,年,月份
   Dim stHDate As String '補充保費代扣日期
   Dim stJDate As String '每月獎金日期
   Dim bolInTran As Boolean
   Dim iErr As Integer
   
On Error GoTo ErrHnd
   
   iErr = 0
   List1.Clear
   ProgressBar1.Value = 0
   
   stMB01 = (Val(txtYEAR) + 1911)
   stDate = strSrvDate(1)
   'If stMB01 = "2023" Then stDate = "20240202" '轉112年度用
   'Modified by Morgan 2025/1/23 改固定在福利金年度的隔年
   'stYear = Left(stDate, 4)
   stYear = Val(stMB01) + 1
   'end 2025/1/23
   stMon = Mid(stDate, 5, 2)
   '1~4月,4月薪資扣
   If Val(stMon) <= 4 Then
      stHDate = stYear & "0430"
   '5~8月,8月薪資扣
   ElseIf Val(stMon) <= 8 Then
      stHDate = stYear & "0831"
   '9~12月,12月薪資扣
   Else
      stHDate = stYear & "1231"
   End If
      
   strExc(0) = "select a.*,st02 from (select sd19,mb04,sum(mb05) mb05,(mb01-1911)||'年'||max(decode(mb02,'01',ac03||'獎金',mb06)) mb06,mb02" & _
      " from miscbonus,salarydata,allcode where mb01=" & stMB01 & " and mb11=0 and sd01(+)=mb04 and sd19='J' and ac01(+)='17' and ac02(+)=mb02" & _
      " group by sd19,mb04,mb01,mb02" & _
      " union select sd19,mb04,sum(mb05) mb05,'' mb06,'' mb02 from miscbonus,salarydata where mb01=" & stMB01 & " and mb11=0 and sd01(+)=mb04 and sd19<>'J' group by sd19,mb04) a,staff where st01(+)=mb04 order by sd19,mb04,mb02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      ProgressBar1.max = .RecordCount
      lblProgress = ProgressBar1.Value & "/" & ProgressBar1.max
      bolInTran = False
      Do While Not .EOF
         cnnConnection.BeginTrans
         bolInTran = True
         
         List1.AddItem .Fields("mb04") & .Fields("st02") & ">>", 0
         strSql = "update miscbonus set mb11=" & strSrvDate(1) & " where mb01=" & stMB01 & " and mb04='" & .Fields("mb04") & "' and mb11=0"
         'Added by Morgan 2025/1/23
         If "" & .Fields("mb02") <> "" Then
            strSql = strSql & " and mb02='" & .Fields("mb02") & "'"
         End If
         cnnConnection.Execute strSql, intI
         If intI > 0 Then
            If .Fields("sd19") = "J" Then
               List1.List(0) = List1.List(0) & "轉獎金.."
               stJDate = GetMBDate(stDate, .Fields("mb04"))
               strSql = "INSERT INTO MonthBonus (MB01,MB02,MB03,MB04,MB11,MB13,MB14)" & _
                  " Values (" & stJDate & ",'" & .Fields("mb04") & "'," & .Fields("mb05") & ",null,'" & .Fields("sd19") & "'," & stHDate & ",'" & .Fields("mb06") & "')"
               
            Else
               List1.List(0) = List1.List(0) & "轉所得.."
               strSql = "insert into OtherIncomeData(oid01,oid02,oid03,oid04,oid05,oid06,oid07,oid08,oid09)" & _
                  " values(" & stYear & ",'" & .Fields("mb04") & "','" & .Fields("sd19") & "','92'," & stMon & "," & stMon & ",'8A'," & .Fields("mb05") & ",0)"
            End If
            cnnConnection.Execute strSql, intI
            cnnConnection.CommitTrans
            List1.List(0) = List1.List(0) & "成功!"
         Else
            cnnConnection.RollbackTrans
            List1.List(0) = List1.List(0) & "失敗!找不到資料"
            Exit Sub
         End If
         bolInTran = False
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = ProgressBar1.Value & "/" & ProgressBar1.max
         DoEvents
         .MoveNext
      Loop
      End With
   Else
      MsgBox "無資料可轉檔！", vbExclamation
   End If
   Exit Sub
   
ErrHnd:
   If bolInTran Then cnnConnection.RollbackTrans
   'MsgBox Err.Description, vbCritical
   List1.List(0) = List1.List(0) & "失敗!" & Err.Description
End Sub

'Added by Morgan 2025/1/23
'取得每月獎金日期
Private Function GetMBDate(pMB01, pMB02) As String
   Dim stSQL As String, intQ As Integer, ii As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stDDate As String
   Dim bolOK As Boolean
   
   stDDate = pMB01
   For ii = 1 To 100
      stSQL = "select * from MonthBonus where mb01=" & stDDate & " and mb02='" & pMB02 & "'"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 0 Then
         GetMBDate = stDDate
         Exit For
      Else
         stDDate = CompDate(2, 1, stDDate)
         stDDate = PUB_GetWorkDay1(stDDate, False)
      End If
   Next
   Set rsQuery = Nothing
End Function
