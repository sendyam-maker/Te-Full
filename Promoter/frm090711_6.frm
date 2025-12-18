VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090711_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖超時內部收文"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4575
   StartUpPosition =   3  '系統預設值
   Begin VB.CheckBox Check1 
      Caption         =   "發 EMail 通知相關人員"
      Height          =   225
      Left            =   225
      TabIndex        =   4
      Top             =   4020
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回(&B)"
      Height          =   345
      Left            =   3375
      TabIndex        =   6
      Top             =   90
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "存檔(&S)"
      Height          =   345
      Left            =   2205
      TabIndex        =   5
      Top             =   90
      Width           =   1095
   End
   Begin VB.TextBox txtCP104 
      Alignment       =   1  '靠右對齊
      Enabled         =   0   'False
      Height          =   270
      Left            =   1755
      TabIndex        =   3
      Top             =   3660
      Width           =   465
   End
   Begin VB.TextBox txtCP101 
      Alignment       =   1  '靠右對齊
      Enabled         =   0   'False
      Height          =   270
      Left            =   1755
      TabIndex        =   2
      Top             =   3030
      Width           =   465
   End
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   1260
      MaxLength       =   4
      TabIndex        =   1
      Top             =   2370
      Width           =   465
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   1125
      Left            =   1215
      TabIndex        =   0
      Top             =   1140
      Width           =   3165
      VariousPropertyBits=   -1467989989
      MaxLength       =   1000
      ScrollBars      =   2
      Size            =   "5583;1984"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1215
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   840
      Width           =   3165
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "5583;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      Caption         =   "1"
      Height          =   180
      Index           =   33
      Left            =   1440
      TabIndex        =   18
      Top             =   2820
      Width           =   780
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "草圖加乘註記："
      Height          =   180
      Left            =   225
      TabIndex        =   17
      Top             =   3075
      Width           =   1260
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "草圖計件值："
      Height          =   180
      Left            =   225
      TabIndex        =   16
      Top             =   2820
      Width           =   1080
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      Caption         =   "1"
      Height          =   180
      Index           =   34
      Left            =   1470
      TabIndex        =   15
      Top             =   3435
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "墨圖加乘註記："
      Height          =   180
      Left            =   225
      TabIndex        =   14
      Top             =   3705
      Width           =   1260
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "墨圖計件值："
      Height          =   180
      Left            =   225
      TabIndex        =   13
      Top             =   3420
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "工作時數：               小時"
      Height          =   180
      Left            =   270
      TabIndex        =   12
      Top             =   2430
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(進度備註)"
      Height          =   180
      Left            =   270
      TabIndex        =   11
      Top             =   1110
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "超時原因："
      Height          =   180
      Left            =   270
      TabIndex        =   9
      Top             =   900
      Width           =   900
   End
   Begin VB.Label lblCP09 
      AutoSize        =   -1  'True
      Caption         =   "B"
      Height          =   180
      Left            =   1260
      TabIndex        =   8
      Top             =   540
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號："
      Height          =   180
      Left            =   270
      TabIndex        =   7
      Top             =   540
      Width           =   720
   End
End
Attribute VB_Name = "frm090711_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (Combo1,txtCP64)
'Memo By Morgan 2012/12/10 智權人員欄已修改
Option Explicit

Public p_CP43 As String
Dim m_CP09 As String

Private Function TxtValidate() As Boolean
   If Combo1 = "" Then
      MsgBox "請輸入超時原因！"
      Combo1.SetFocus
      Exit Function
   End If
   
   If txtCP113 = "" Then
      MsgBox "請輸入工作時數！"
      txtCP113.SetFocus
      Exit Function
   End If
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
   Dim stSQL As String, stBNo As String
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   stBNo = m_CP09
   If stBNo = "" Then
      stBNo = AutoNo("B", 6) 'B類總收文號
      stSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13" & _
         ",cp14,cp20,cp26,cp29,cp27,cp43,cp107) " & _
         " select cp01,cp02,cp03,cp04," & strSrvDate(1) & ",'" & stBNo & "','943',cp12,cp13" & _
         ",cp29,'N','N',cp29,cp27,'" & p_CP43 & "','Y' from caseprogress where cp09='" & p_CP43 & "'"
      cnnConnection.Execute stSQL, intI
      
      'Modify by Morgan 2011/5/3 改都上系統日(update ep20 是為了要觸發 Trigger 計算計件值)
      stSQL = "update engineerprogress a set ep06=" & strSrvDate(1) & ",ep14=" & strSrvDate(1) & ",ep15=" & strSrvDate(1) & _
         ",ep17=" & strSrvDate(1) & ",ep18=" & strSrvDate(1) & ",ep20=null where ep02='" & stBNo & "'"
      cnnConnection.Execute stSQL, intI
      
'Removed by Morgan 2012/9/4 取消 EMail 通知--瓊玉
'      'Add by Morgan 2011/5/18 加有勾選才要發 EMail
'      'Modify by Morgan 2011/5/25 區主管和王副總改固定要通知
'      If Check1.Value = 1 Then
'         'Modify by Morgan 2011/5/2 +王副總71011也要寄
'         'Modify by Morgan 2011/5/18 文字調整--游經理
'         'Modified by Morgan 2012/1/18 智權部收信人全部為杜副總--張瓊玉(2012/1/20確認)
'         'stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'            " select '" & strUserNum & "',cp13,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'            ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'案('||cp09||')" & _
'            "的繪圖作業，提醒智權同仁如下訊息：','因案件內容" & Combo1 & "，造成作業成本增加，特提醒有此狀況，請斟酌" & _
'            "向客戶反應收取適當服務費。" & vbCrLf & vbCrLf & "（註：以下為專利訊息，繪圖系統以Ｂ類收文調整繪圖人員計件值）'" & _
'            ",'71011;'||OMAN from caseprogress,staff,SetSpecMan where cp09='" & p_CP43 & "'" & _
'            " and st01(+)=cp29 and OCODE(+)=decode(st06,'1','A2','2','A3','3','A4','4','A5')"
'         stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'            " select '" & strUserNum & "','68006',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'            ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'案('||cp09||')" & _
'            "的繪圖作業，提醒智權同仁如下訊息：','因案件內容" & Combo1 & "，造成作業成本增加，特提醒有此狀況。" & vbCrLf & vbCrLf & _
'            "（註：以下為專利訊息，繪圖系統以Ｂ類收文調整繪圖人員計件值）'" & _
'            ",'71011;'||OMAN from caseprogress,staff,SetSpecMan where cp09='" & p_CP43 & "'" & _
'            " and st01(+)=cp29 and OCODE(+)=decode(st06,'1','A2','2','A3','3','A4','4','A5')"
'      Else
'         stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'            " select '" & strUserNum & "',OMAN,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'            ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'案('||cp09||')" & _
'            "的繪圖作業，因案件內容" & Combo1 & "，繪圖系統以Ｂ類收文調整繪圖人員計件值。','如旨','71011'" & _
'            " from caseprogress,staff,SetSpecMan where cp09='" & p_CP43 & "'" & _
'            " and st01(+)=cp29 and OCODE(+)=decode(st06,'1','A2','2','A3','3','A4','4','A5')"
'      End If
'      cnnConnection.Execute stSQL, intI
      
   End If
   '新增會預設加乘註記,所以一定要用更新的方式才不會被覆蓋
   stSQL = "update caseprogress set cp64='" & ChgSQL(Combo1 & IIf(Combo1 = "", "", ":") & txtCP64) & "',cp101=" & Val(txtCP101) & ",cp104=" & Val(txtCP104) & _
      ",cp113=" & Val(txtCP113) & " where cp09='" & stBNo & "'"
   cnnConnection.Execute stSQL, intI
         
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub Combo1_Click()
   
   If txtCP64.Enabled And txtCP64.Visible Then txtCP64.SetFocus
End Sub

Private Sub Command1_Click()
   If TxtValidate = True Then
      If FormSave = True Then
         If m_CP09 = "" Then
            MsgBox "若要顯示繪圖超時內部收文，請重新查詢本月資料！"
         End If
         Unload Me
      End If
   End If
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'Modify by Morgan 2011/5/18 文字調整--游經理
   'Modified by Morgan 2012/9/21 瓊玉
   'Combo1.AddItem "較複雜"
   'Combo1.AddItem "有大幅修改"
   'Combo1.AddItem "屬重辦案件"
   Combo1.AddItem "複雜"
   Combo1.AddItem "重辦"
   Combo1.AddItem "修改"
   Combo1.AddItem "描圖"
   Combo1.AddItem "其他"
   Combo1.AddItem "成組設計" 'Added by Morgan 2014/1/6 瓊玉
   Combo1.ListIndex = -1
   'end 2012/9/21
   
   ReadData
End Sub

Private Function ReadData()
   Dim stSQL As String
   Dim intR As Integer
   Dim adoRst As ADODB.Recordset
   Dim ii As Integer, strDesc As String
   
   stSQL = "select * from caseprogress where cp43='" & p_CP43 & "' and cp10='943' order by cp09 desc"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      m_CP09 = adoRst("cp09")
      lblCP09 = m_CP09
      If Not IsNull(adoRst("cp64")) Then
         For ii = 0 To Combo1.ListCount - 1
            If InStr(adoRst("cp64"), Combo1.List(ii) & ":") = 1 Then
               Combo1.ListIndex = ii
               txtCP64 = Mid(adoRst("cp64"), Len(Combo1.List(ii) & ":") + 1)
               Exit For
            End If
         Next
         If Combo1.ListIndex = -1 Then
            txtCP64.Text = adoRst("cp64")
         End If
      End If
      txtCP113 = "" & adoRst("cp113")
      Me.Caption = Me.Caption & "(修改)"
      Check1.Enabled = False
   Else
      txtCP64 = ""
      Me.Caption = Me.Caption & "(新增)"
      Check1.Enabled = True
   End If
   If adoRst.State <> adStateClosed Then adoRst.Close
   Set adoRst = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm090711_6 = Nothing
End Sub

Private Sub txtCP113_Change()
   Dim dblPlus As Double
   If txtCP113 <> "" Then
      'Modified by Morgan 2014/4/23 改 6.5 小時
      'dblPlus = Round(Val(txtCP113) / 8#, 1)
      dblPlus = Round(Val(txtCP113) / 6.5, 1)
      txtCP101 = dblPlus
      txtCP104 = dblPlus
   End If
End Sub

Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

Private Sub txtCP113_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc(".") And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCP64_GotFocus()
   OpenIme
End Sub
