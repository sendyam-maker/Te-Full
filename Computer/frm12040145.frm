VERSION 5.00
Begin VB.Form frm12040145 
   BorderStyle     =   1  '單線固定
   Caption         =   "本所期限工作天推算作業"
   ClientHeight    =   3660
   ClientLeft      =   3480
   ClientTop       =   3192
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6120
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2724
      TabIndex        =   1
      Top             =   1116
      Width           =   2500
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   4332
      TabIndex        =   5
      Top             =   96
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3540
      TabIndex        =   4
      Top             =   96
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2892
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1536
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1620
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1536
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   2700
      TabIndex        =   0
      Text            =   "P,PS,CFP,CPS,FCT,CFT,CFC,S,T,TB,TC,TD,TF,TM,TR,TS,TT"
      Top             =   672
      Width           =   2500
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "　　語法在  TaieNew\電腦中心日常工作"
      ForeColor       =   &H0000C000&
      Height          =   180
      Left            =   108
      TabIndex        =   14
      Top             =   3072
      Width           =   3096
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "　　FCP及FG案件更新下一年度假日本所期限為前一工作日後檢查語法.txt"
      ForeColor       =   &H0000C000&
      Height          =   180
      Left            =   108
      TabIndex        =   13
      Top             =   3276
      Width           =   5784
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "PS：執行完檢查FCP及FG案件，詢問是否要通知代理人"
      ForeColor       =   &H0000C000&
      Height          =   180
      Left            =   108
      TabIndex        =   12
      Top             =   2844
      Width           =   4308
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "　　工作日調整為假日之承辦期限更新處理.doc"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   2448
      Width           =   3708
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "　　語法在  TaieNew\電腦中心日常工作"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   2244
      Width           =   3096
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "PS：若為平時調整為假日者，要檢查承辦期限的資料"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   4152
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "限台灣案之系統類別:"
      Height          =   180
      Left            =   804
      TabIndex        =   8
      Top             =   1116
      Width           =   1668
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   780
      TabIndex        =   7
      Top             =   1572
      Width           =   768
   End
   Begin VB.Line Line2 
      X1              =   2532
      X2              =   2772
      Y1              =   1656
      Y2              =   1656
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "不限申請國家系統類別:"
      Height          =   180
      Left            =   780
      TabIndex        =   6
      Top             =   672
      Width           =   1848
   End
End
Attribute VB_Name = "frm12040145"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/10 Form2.0不用改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Dim s As Integer, i As Integer, j As Integer, iLine As Integer, k As Integer
Dim strSql As String, StrTest As String
Dim strTemp As Variant, strTemp1 As Variant, StrTempP As Variant, StrTempP2 As Variant
Dim Page As Integer, iPrint As Integer, St As String, TmpArea As String
Dim PLeft1(0 To 7) As Integer, Pleft2(0 To 10) As Integer, PLeft3(0 To 8) As Integer
Dim strSQL2 As String, strSQL1 As String, StrSQL3 As String, StrSQL6 As String
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
Dim STRSTRING As String
Dim m_blnNoData1 As Boolean
Dim m_blnNoData2 As Boolean

Private Sub cmdok_Click(Index As Integer)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   Select Case Index
      Case 0 '確定
          blnClkSure = False
          If Len(txt1(0)) = 0 Then
              s = MsgBox("系統類別不可空白", , "USER 輸入錯誤")
              txt1(0).SetFocus
              Exit Sub
          End If
          If PUB_CheckKeyInDate(Me.txt1(3)) = -1 Then
              Me.txt1(3).SetFocus
              txt1_GotFocus 3
              Exit Sub
          End If
          If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
              Me.txt1(4).SetFocus
              txt1_GotFocus 4
              Exit Sub
          End If
          If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
              If Val(Me.txt1(3).Text) > Val(Me.txt1(4).Text) Then
                  'edit by nick 2004/09/27
                  'MsgBox "法定期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  MsgBox "本所期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
              End If
          End If
          If Len(Trim(txt1(4))) = 0 Then
              'edit by nick 2004/09/27
              's = MsgBox("法定期限不可空白", , "USER 輸入錯誤")
              s = MsgBox("本所期限不可空白", , "USER 輸入錯誤")
              txt1(3).SetFocus
              txt1_GotFocus (3)
              Exit Sub
          End If
          Screen.MousePointer = vbHourglass
          Me.Enabled = False
          'StrMenu
          StrMenu2
          If m_blnNoData1 = True And m_blnNoData2 = True Then
              ShowNoData
          Else
              MsgBox "本所期限重整完畢!!!", vbExclamation + vbOKOnly
          End If
          
          'add by sonia 2016/8/26 若為某日臨時調整為假日時才跑
          If ChangeTStringToWString(Me.txt1(3).Text) = ChangeTStringToWString(Me.txt1(4).Text) Then
             StrSQLa = "Select * From Staff_CALENDAR Where SC01=" & ChangeTStringToWString(Me.txt1(3).Text) & " "
             rsA.CursorLocation = adUseClient
             rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
             If rsA.RecordCount > 0 Then
                If IsNull(rsA.Fields(0)) = False Then
                   MsgBox "國外部行事曆有管制此日期的資料, 請通知國外部自行決定是否調整 !!!", vbExclamation + vbOKOnly
                End If
             End If
             rsA.Close
          End If
          'end 2016/8/26
          
          Me.Enabled = True
          Screen.MousePointer = vbDefault
      Case 1 '結束
          Unload Me
      Case Else
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'add by sonia 2016/8/25 原只有P,CFP加入其他系統別, 但 2014/11/20 FCP,FG改回舊規則 2019/7/15FCP,FG又改
   '但僅P,PS,CFP,CPS是不管申請國家都改,其他系統類別僅限台灣案才改本所期限,故分二個欄位
   'modify by sonia 2020/7/13 內商及外商CF系統也都改為本所期限要工作日
   'txt1(0) = "P,PS,CFP,CPS,FCP,FG,FCT"
   'txt1(1) = "S,T,TB,TC,TD,TM,TR,TS,TT"
   '2024/2/16 +ACS，只剩法顧案件未預設
   txt1(0) = "ACS,P,PS,CFP,CPS,FCP,FG,FCT,CFT,CFC,S,T,TB,TC,TD,TF,TM,TR,TS,TT"
   txt1(1) = ""
   'end 2020/7/13
End Sub

'add by nick 2004/10/06
Sub StrMenu2()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strWD01 As String

On Error GoTo ErrorHandler
      
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      
      'Added by Lydia 2025/11/12
      Dim strUpdNP08 As String, strUpdNP23 As String
      Dim strUpdCP06 As String, strUpdCP48 As String
      strUpdNP08 = ", np15=np15||decode(np15,NULL,NULL,';')||sqldatet(to_char(SYSDATE,'yyyymmdd'))||'電腦中心更新本所期限，原為'||sqldatet(np08)"
      strUpdCP06 = ", cp64=cp64||decode(cp64,NULL,NULL,';')||sqldatet(to_char(SYSDATE,'yyyymmdd'))||'電腦中心更新本所期限，原為'||sqldatet(cp06)"
      strUpdNP23 = ", np15=np15||decode(np15,NULL,NULL,';')||sqldatet(to_char(SYSDATE,'yyyymmdd'))||'電腦中心更新約定期限，原為'||sqldatet(np23)"
      strUpdCP48 = ", cp64=cp64||decode(cp64,NULL,NULL,';')||sqldatet(to_char(SYSDATE,'yyyymmdd'))||'電腦中心更新承辦期限，原為'||sqldatet(cp48)"
      'end 2025/11/12
      cnnConnection.BeginTrans
      
      '不限申請國家系統類別 txt1(0)
      StrSQLa = ""
      If Len(txt1(0)) <> 0 Then
          StrSQLa = StrSQLa & " AND CP01 IN (" & GetAddStr(txt1(0)) & ") "
      End If
      If Me.txt1(3).Text <> "" Then
          StrSQLa = StrSQLa & " AND CP06>=" & ChangeTStringToWString(Me.txt1(3).Text) & " "
      End If
      If Me.txt1(4).Text <> "" Then
          StrSQLa = StrSQLa & " AND CP06<=" & ChangeTStringToWString(Me.txt1(4).Text) & " "
      End If
      
      'MODIFY BY SONIA 2016/4/26
      'StrSQLa = "update caseprogress set cp06=(select max(wd01) from workday where wd01<=cp06) where CP27 IS NULL AND CP57 IS NULL " & StrSQLa
      'modify by sonia 2016/8/25 cp27,cp57改用新欄位cp158,cp159
      'StrSQLa = "update caseprogress set cp06=(select max(wd01) from workday where wd01<=cp06) where NVL(CP27,0)=0 AND NVL(CP57,0)=0 " & StrSQLa
      'Modified by Lydia 2025/11/12 增加備註strUpdCP06
      StrSQLa = "update caseprogress set cp06=(select max(wd01) from workday where wd01<=cp06) " & strUpdCP06 & " where CP158=0 AND CP159=0 " & StrSQLa
      cnnConnection.Execute StrSQLa, intI
      
      'add by sonia 2025/3/14 更新承辦期限
      StrSQLa = ""
      If Len(txt1(0)) <> 0 Then
          StrSQLa = StrSQLa & " AND CP01 IN (" & GetAddStr(txt1(0)) & ") "
      End If
      If Me.txt1(3).Text <> "" Then
          StrSQLa = StrSQLa & " AND CP48>=" & ChangeTStringToWString(Me.txt1(3).Text) & " "
      End If
      If Me.txt1(4).Text <> "" Then
          StrSQLa = StrSQLa & " AND CP48<=" & ChangeTStringToWString(Me.txt1(4).Text) & " "
      End If
      
      'Modified by Lydia 2025/11/12 增加備註strUpdCP48
      StrSQLa = "update caseprogress set cp48=(select max(wd01) from workday where wd01<=cp48) " & strUpdCP48 & " where CP158=0 AND CP159=0 " & StrSQLa
      cnnConnection.Execute StrSQLa, intI
      'end 2025/3/14

      StrSQLa = ""
      If Len(txt1(0)) <> 0 Then
          StrSQLa = StrSQLa & " AND Np02 IN (" & GetAddStr(txt1(0)) & ") "
      End If
      If Me.txt1(3).Text <> "" Then
          StrSQLa = StrSQLa & " AND NP08>=" & ChangeTStringToWString(Me.txt1(3).Text) & " "
      End If
      If Me.txt1(4).Text <> "" Then
          StrSQLa = StrSQLa & " AND NP08<=" & ChangeTStringToWString(Me.txt1(4).Text) & " "
      End If
      'Modified by Lydia 2025/11/12 增加備註strUpdNP08
      StrSQLa = "update nextprogress set np08=(select max(wd01) from workday where wd01<=np08) " & strUpdNP08 & " where NP06 IS NULL " & StrSQLa
      'Add by Morgan 2008/12/5 +控制來函期限通知的 Trigger 不被觸發
      StrSQLa = "begin user_data.user_notrigger:=1; " & StrSQLa & "; user_data.user_notrigger:=0; end;"
      'end 2008/12/5
      cnnConnection.Execute StrSQLa, intI
      
      'Add by Morgan 2011/1/11 更新約定期限
      StrSQLa = ""
      If Len(txt1(0)) <> 0 Then
          StrSQLa = StrSQLa & " AND Np02 IN (" & GetAddStr(txt1(0)) & ") "
      End If
      If Me.txt1(3).Text <> "" Then
          StrSQLa = StrSQLa & " AND NP23>=" & ChangeTStringToWString(Me.txt1(3).Text) & " "
      End If
      If Me.txt1(4).Text <> "" Then
          StrSQLa = StrSQLa & " AND NP23<=" & ChangeTStringToWString(Me.txt1(4).Text) & " "
      End If
      'Modified by Lydia 2025/11/12 增加備註strUpdNP23
      StrSQLa = "update nextprogress set np23=(select max(wd01) from workday where wd01<=np23) " & strUpdNP23 & " where NP06 IS NULL " & StrSQLa
      'Add by Morgan 2008/12/5 +控制來函期限通知的 Trigger 不被觸發
      StrSQLa = "begin user_data.user_notrigger:=1; " & StrSQLa & "; user_data.user_notrigger:=0; end;"
      'end 2008/12/5
      cnnConnection.Execute StrSQLa, intI
      
      'add by sonia 2016/8/25 限台灣案之系統類別 txt1(1)
      StrSQLa = ""
      If Len(txt1(1)) <> 0 Then
          StrSQLa = StrSQLa & " AND CP01 IN (" & GetAddStr(txt1(1)) & ") "
      End If
      If Me.txt1(3).Text <> "" Then
          StrSQLa = StrSQLa & " AND CP06>=" & ChangeTStringToWString(Me.txt1(3).Text) & " "
      End If
      If Me.txt1(4).Text <> "" Then
          StrSQLa = StrSQLa & " AND CP06<=" & ChangeTStringToWString(Me.txt1(4).Text) & " "
      End If
      'Modified by Lydia 2025/11/12 增加備註strUpdCP06
      StrSQLa = "update caseprogress set cp06=(select max(wd01) from workday where wd01<=cp06) " & strUpdCP06 & " where CP158=0 AND CP159=0 " & StrSQLa & _
                "   and cp09 in (      select cp09 from caseprogress,trademark where CP158=0 AND CP159=0 " & StrSQLa & " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm10='000' " & _
                "                union select cp09 from caseprogress,patent where CP158=0 AND CP159=0 " & StrSQLa & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09='000' " & _
                "                union select cp09 from caseprogress,lawcase where CP158=0 AND CP159=0 " & StrSQLa & " and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and lc15='000' " & _
                "                union select cp09 from caseprogress,servicepractice where CP158=0 AND CP159=0 " & StrSQLa & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and sp09='000' " & _
                "                union select cp09 from caseprogress,hirecase where CP158=0 AND CP159=0 " & StrSQLa & " and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and hc01 is not null" & ")"
      cnnConnection.Execute StrSQLa, intI
      
      'add by sonia 2025/3/14 更新承辦期限
      StrSQLa = ""
      If Len(txt1(1)) <> 0 Then
          StrSQLa = StrSQLa & " AND CP01 IN (" & GetAddStr(txt1(1)) & ") "
      End If
      If Me.txt1(3).Text <> "" Then
          StrSQLa = StrSQLa & " AND CP48>=" & ChangeTStringToWString(Me.txt1(3).Text) & " "
      End If
      If Me.txt1(4).Text <> "" Then
          StrSQLa = StrSQLa & " AND CP48<=" & ChangeTStringToWString(Me.txt1(4).Text) & " "
      End If
      'Modified by Lydia 2025/11/12 增加備註strUpdCP48
      StrSQLa = "update caseprogress set cp48=(select max(wd01) from workday where wd01<=cp48) " & strUpdCP48 & " where CP158=0 AND CP159=0 " & StrSQLa & _
                "   and cp09 in (      select cp09 from caseprogress,trademark where CP158=0 AND CP159=0 " & StrSQLa & " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and tm10='000' " & _
                "                union select cp09 from caseprogress,patent where CP158=0 AND CP159=0 " & StrSQLa & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09='000' " & _
                "                union select cp09 from caseprogress,lawcase where CP158=0 AND CP159=0 " & StrSQLa & " and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and lc15='000' " & _
                "                union select cp09 from caseprogress,servicepractice where CP158=0 AND CP159=0 " & StrSQLa & " and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and sp09='000' " & _
                "                union select cp09 from caseprogress,hirecase where CP158=0 AND CP159=0 " & StrSQLa & " and cp01=hc01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and hc01 is not null" & ")"
      cnnConnection.Execute StrSQLa, intI
      'end 2025/3/14
      
      StrSQLa = ""
      If Len(txt1(1)) <> 0 Then
          StrSQLa = StrSQLa & " AND Np02 IN (" & GetAddStr(txt1(1)) & ") "
      End If
      If Me.txt1(3).Text <> "" Then
          StrSQLa = StrSQLa & " AND NP08>=" & ChangeTStringToWString(Me.txt1(3).Text) & " "
      End If
      If Me.txt1(4).Text <> "" Then
          StrSQLa = StrSQLa & " AND NP08<=" & ChangeTStringToWString(Me.txt1(4).Text) & " "
      End If
      'Modified by Lydia 2025/11/12 增加備註strUpdNP08
      StrSQLa = "update nextprogress set np08=(select max(wd01) from workday where wd01<=np08) " & strUpdNP08 & " where NP06 IS NULL " & StrSQLa & _
                "   and (np01,np07,np22) in (      select np01,np07,np22 from nextprogress,trademark where NP06 IS NULL " & StrSQLa & " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and tm10='000' " & _
                "                            union select np01,np07,np22 from nextprogress,patent where NP06 IS NULL " & StrSQLa & " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and pa09='000' " & _
                "                            union select np01,np07,np22 from nextprogress,lawcase where NP06 IS NULL " & StrSQLa & " and np02=lc01(+) and np03=lc02(+) and np04=lc03(+) and np05=lc04(+) and lc15='000' " & _
                "                            union select np01,np07,np22 from nextprogress,servicepractice where NP06 IS NULL " & StrSQLa & " and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) and sp09='000' " & _
                "                            union select np01,np07,np22 from nextprogress,hirecase where NP06 IS NULL " & StrSQLa & " and np02=hc01(+) and np03=hc02(+) and np04=hc03(+) and np05=hc04(+) and hc01 is not null" & ")"
      '控制來函期限通知的 Trigger 不被觸發
      StrSQLa = "begin user_data.user_notrigger:=1; " & StrSQLa & "; user_data.user_notrigger:=0; end;"
      cnnConnection.Execute StrSQLa, intI
      
      '更新約定期限
      StrSQLa = ""
      If Len(txt1(1)) <> 0 Then
          StrSQLa = StrSQLa & " AND Np02 IN (" & GetAddStr(txt1(1)) & ") "
      End If
      If Me.txt1(3).Text <> "" Then
          StrSQLa = StrSQLa & " AND NP23>=" & ChangeTStringToWString(Me.txt1(3).Text) & " "
      End If
      If Me.txt1(4).Text <> "" Then
          StrSQLa = StrSQLa & " AND NP23<=" & ChangeTStringToWString(Me.txt1(4).Text) & " "
      End If
      'Modified by Lydia 2025/11/12 增加備註strUpdNP23
      StrSQLa = "update nextprogress set np23=(select max(wd01) from workday where wd01<=np23) " & strUpdNP23 & " where NP06 IS NULL " & StrSQLa & _
                "   and (np01,np07,np22) in (      select np01,np07,np22 from nextprogress,trademark where NP06 IS NULL " & StrSQLa & " and np02=tm01(+) and np03=tm02(+) and np04=tm03(+) and np05=tm04(+) and tm10='000' " & _
                "                            union select np01,np07,np22 from nextprogress,patent where NP06 IS NULL " & StrSQLa & " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and pa09='000' " & _
                "                            union select np01,np07,np22 from nextprogress,lawcase where NP06 IS NULL " & StrSQLa & " and np02=lc01(+) and np03=lc02(+) and np04=lc03(+) and np05=lc04(+) and lc15='000' " & _
                "                            union select np01,np07,np22 from nextprogress,servicepractice where NP06 IS NULL " & StrSQLa & " and np02=sp01(+) and np03=sp02(+) and np04=sp03(+) and np05=sp04(+) and sp09='000' " & _
                "                            union select np01,np07,np22 from nextprogress,hirecase where NP06 IS NULL " & StrSQLa & " and np02=hc01(+) and np03=hc02(+) and np04=hc03(+) and np05=hc04(+) and hc01 is not null" & ")"
      '控制來函期限通知的 Trigger 不被觸發
      StrSQLa = "begin user_data.user_notrigger:=1; " & StrSQLa & "; user_data.user_notrigger:=0; end;"
      cnnConnection.Execute StrSQLa, intI
      
      'Add By Sindy 2021/2/26 檢查工作所在地資料中,非工作日資料刪除
      StrSQLa = "delete FROM STAFF_WORKPLACE WHERE sp01 IN(" & _
                " SELECT sp01 FROM STAFF_WORKPLACE,workday" & _
                " WHERE SP01=wd01(+) AND wd01 IS NULL" & _
                " and sp01>=" & ChangeTStringToWString(Me.txt1(3).Text) & " and sp01<=" & ChangeTStringToWString(Me.txt1(4).Text) & _
                " GROUP BY sp01" & _
                ")"
      cnnConnection.Execute StrSQLa, intI
      '檢查是否有補班資料漏掉了
      StrSQLa = "SELECT wd01 FROM workday,STAFF_WORKPLACE" & _
                " WHERE wd01>=" & ChangeTStringToWString(Me.txt1(3).Text) & " AND wd01<=" & ChangeTStringToWString(Me.txt1(4).Text) & _
                " AND wd06='Y' and wd01=sp01(+) and wd01 is not null"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      While Not rsA.EOF
         strWD01 = CompDate(2, -1, rsA.Fields("wd01"))
         StrSQLa = "INSERT INTO staff_workplace(sp01,sp02,sp03,sp04)" & _
                   " SELECT " & rsA.Fields("wd01") & ",sp02,sp03,'QPGMR'" & _
                   " FROM STAFF_WORKPLACE WHERE sp01=" & strWD01
         cnnConnection.Execute StrSQLa, intI
         rsA.MoveNext
      Wend
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      '2021/2/26 END
      
      cnnConnection.CommitTrans
      'end 2016/8/25
            
      Me.Enabled = True
      Screen.MousePointer = vbDefault
      Exit Sub

ErrorHandler:
    Screen.MousePointer = vbDefault
    cnnConnection.RollbackTrans
    If Err.Number <> 0 Then MsgBox "(" & Err.Number & ")" & Err.Description

End Sub

'edit by nick 2004/10/06  以下作廢
'Sub StrMenu()
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'Dim strDate(0 To 3) As String
'
'   Screen.MousePointer = vbHourglass
'   Me.Enabled = False
'   m_blnNoData1 = True: m_blnNoData2 = True
'   '更新進度檔的本所期限
'   Erase strDate
'   StrSQLa = ""
'   If Len(txt1(0)) <> 0 Then
'       StrSQLa = StrSQLa & " AND CP01 IN (" & GetAddStr(txt1(0)) & ") "
'   End If
'   If Me.txt1(3).Text <> "" Then
'       'edit by nick 2004/09/27
'       'strSQLA = strSQLA & " AND CP07>=" & ChangeTStringToWString(Me.txt1(3).Text) & " "
'       StrSQLa = StrSQLa & " AND CP06>=" & ChangeTStringToWString(Me.txt1(3).Text) & " "
'   End If
'   If Me.txt1(4).Text <> "" Then
'       'edit by nick 2004/09/27
'       'strSQLA = strSQLA & " AND CP07<=" & ChangeTStringToWString(Me.txt1(4).Text) & " "
'       StrSQLa = StrSQLa & " AND CP06<=" & ChangeTStringToWString(Me.txt1(4).Text) & " "
'   End If
'   '93.9.24 MODIFY BY SONIA
'   'strSQLA = "Select * From CaseProgress, Patent Where CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 " & strSQLA
'   StrSQLa = "Select * From CaseProgress, Patent Where CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) AND CP27 IS NULL AND CP57 IS NULL " & StrSQLa
'   '93.9.24 END
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'   While Not rsA.EOF
'       If m_blnNoData1 = True Then m_blnNoData1 = False
'       strDate(1) = "" & rsA("CP01").Value
'       strDate(2) = "" & rsA("PA09").Value
'       strDate(3) = "" & rsA("CP07").Value
'       GetCtrlDT strDate()
'       StrSQLa = "Update CaseProgress Set CP06=" & PUB_GetWorkDay1(strDate(0), True) & " Where CP09='" & rsA("CP09").Value & "' "
'       cnnConnection.Execute StrSQLa
'       rsA.MoveNext
'   Wend
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   '更新下一程序檔的本所期限
'   Erase strDate
'   StrSQLa = ""
'   If Len(txt1(0)) <> 0 Then
'       StrSQLa = StrSQLa & " AND NP02 IN (" & GetAddStr(txt1(0)) & ") "
'   End If
'   If Me.txt1(3).Text <> "" Then
'       StrSQLa = StrSQLa & " AND NP09>=" & ChangeTStringToWString(Me.txt1(3).Text) & " "
'   End If
'   If Me.txt1(4).Text <> "" Then
'       StrSQLa = StrSQLa & " AND NP09<=" & ChangeTStringToWString(Me.txt1(4).Text) & " "
'   End If
'   '93.9.24 MODIFY BY SONIA
'   'strSQLA = "Select * From NextProgress, Patent Where NP02=PA01 And NP03=PA02 And NP04=PA03 And NP05=PA04 " & strSQLA
'   StrSQLa = "Select * From NextProgress, Patent Where NP02=PA01(+) And NP03=PA02(+) And NP04=PA03(+) And NP05=PA04(+) AND NP06 IS NULL " & StrSQLa
'   '93.9.24 END
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'   While Not rsA.EOF
'       If m_blnNoData2 = True Then m_blnNoData2 = False
'       strDate(1) = "" & rsA("NP02").Value
'       strDate(2) = "" & rsA("PA09").Value
'       strDate(3) = "" & rsA("NP09").Value
'       GetCtrlDT strDate()
'       StrSQLa = "Update NextProgress Set NP08=" & PUB_GetWorkDay1(strDate(0), True) & " Where NP01='" & rsA("NP01").Value & "' And NP07=" & rsA("NP07").Value & " And NP22=" & rsA("NP22").Value
'       cnnConnection.Execute StrSQLa
'       rsA.MoveNext
'   Wend
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   Me.Enabled = True
'   Screen.MousePointer = vbDefault
'End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040145 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
      Case 0, 1
        If Len(txt1(Index)) <> 0 Then
           STRSTRING = ""
           StrTempP = Split(Replace(txt1(Index), ",,", ""), ",")
           StrTempP2 = Split(Replace(GetSystemKindByNick, ",,", ""), ",")
           For i = 0 To UBound(StrTempP)
               s = 0
               For j = 0 To UBound(StrTempP2)
                   If StrTempP(i) = StrTempP2(j) Then
                       s = 1
                   End If
               Next j
               If s = 0 Then
                   STRSTRING = STRSTRING + StrTempP(i) + " "
               End If
           Next i
           If Len(STRSTRING) <> 0 Then
               s = MsgBox(STRSTRING + " 不是 " + strUserNum + " 的權限!!", , "警告!!!")
               txt1(Index).SetFocus
               Exit Sub
           End If
         End If
      Case 2, 4
         If blnClkSure = False Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               txt1(Index - 1).SetFocus
            End If
         Else
            blnClkSure = False
         End If
   End Select

End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 3, 4  '法定期限起, 迄
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Cancel = True
         End If
      End Select
   If Cancel Then TextInverse txt1(Index)
End Sub
