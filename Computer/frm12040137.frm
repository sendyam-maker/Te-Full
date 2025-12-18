VERSION 5.00
Begin VB.Form frm12040137 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作天維護"
   ClientHeight    =   7644
   ClientLeft      =   -108
   ClientTop       =   372
   ClientWidth     =   9228
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7644
   ScaleWidth      =   9228
   Begin VB.CommandButton Command1 
      Caption         =   "結束(&X)"
      Height          =   300
      Index           =   1
      Left            =   8316
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "儲存(&O)"
      Default         =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   7512
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      Left            =   1056
      Style           =   2  '單純下拉式
      TabIndex        =   0
      Top             =   48
      Width           =   1215
   End
   Begin Computer.UsrCalendar Cal1 
      Height          =   2385
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   390
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   4043
   End
   Begin Computer.UsrCalendar Cal1 
      Height          =   2385
      Index           =   1
      Left            =   2370
      TabIndex        =   4
      Top             =   390
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   4043
   End
   Begin Computer.UsrCalendar Cal1 
      Height          =   2385
      Index           =   2
      Left            =   4650
      TabIndex        =   5
      Top             =   390
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   4043
   End
   Begin Computer.UsrCalendar Cal1 
      Height          =   2385
      Index           =   3
      Left            =   6930
      TabIndex        =   6
      Top             =   390
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   4043
   End
   Begin Computer.UsrCalendar Cal1 
      Height          =   2355
      Index           =   4
      Left            =   90
      TabIndex        =   7
      Top             =   2790
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   4043
   End
   Begin Computer.UsrCalendar Cal1 
      Height          =   2355
      Index           =   5
      Left            =   2370
      TabIndex        =   8
      Top             =   2790
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   4043
   End
   Begin Computer.UsrCalendar Cal1 
      Height          =   2355
      Index           =   6
      Left            =   4650
      TabIndex        =   9
      Top             =   2790
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   4043
   End
   Begin Computer.UsrCalendar Cal1 
      Height          =   2355
      Index           =   7
      Left            =   6930
      TabIndex        =   10
      Top             =   2790
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   4043
   End
   Begin Computer.UsrCalendar Cal1 
      Height          =   2355
      Index           =   8
      Left            =   90
      TabIndex        =   11
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   4043
   End
   Begin Computer.UsrCalendar Cal1 
      Height          =   2355
      Index           =   9
      Left            =   2370
      TabIndex        =   12
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   4043
   End
   Begin Computer.UsrCalendar Cal1 
      Height          =   2355
      Index           =   10
      Left            =   4650
      TabIndex        =   13
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   4043
   End
   Begin Computer.UsrCalendar Cal1 
      Height          =   2355
      Index           =   11
      Left            =   6930
      TabIndex        =   14
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   4043
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "選擇年度"
      Height          =   180
      Left            =   216
      TabIndex        =   15
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frm12040137"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/10 Form2.0不用改
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

' 90.07.16 modify by Ken (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Private Sub Combo1_Click()
   Screen.MousePointer = vbHourglass
   ShowCal Val(Combo1.Text)
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click(Index As Integer)
 Dim i As Integer, j As Integer, k As Integer, kk As Integer, varTmp As Variant, strTmp As String, iTmp As Integer
 Dim strTxt(1 To 380) As String
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         iTmp = 1
         For i = 1 To 12
            'Modified by Morgan 2014/11/26 控制有設定/取消的日期才要新增或刪除,否則颱風假或補班的設定會不見
            'strTxt(iTmp) = "DELETE FROM WORKDAY WHERE SUBSTR(WD01,1,6)=" & Combo1.Text & Format(i, "00")
            'strTmp = Cal1(i - 1).SaveDate
            'iTmp = iTmp + 1
            'If strTmp <> "" Then
            '   varTmp = Split(strTmp, ",")
            '   For j = 0 To UBound(varTmp)
            '      strTxt(iTmp) = "INSERT INTO WORKDAY (WD01) VALUES (" & Combo1.Text & Format(i, "00") & Format(varTmp(j), "00") & ")"
            '      iTmp = iTmp + 1
            '   Next
            'End If
            
            strTmp = Cal1(i - 1).SaveDate
            If strTmp <> "" Then
               strTxt(iTmp) = "DELETE FROM WORKDAY WHERE SUBSTR(WD01,1,6)=" & Combo1.Text & Format(i, "00") & " and substr(wd01,-2) not in (" & strTmp & ")" 'Added by Morgan 2015/10/29
               iTmp = iTmp + 1
            
               varTmp = Split(strTmp, ",")
               For j = 0 To UBound(varTmp)
                  '非工作天
                  'Removed by Morgan 2015/10/29 '改上面批次刪
                  'If j = 0 Then
                  '   kk = 1
                  'Else
                  '   kk = Val(varTmp(j - 1)) + 1
                  'End If
                  
                  'For k = kk To varTmp(j) - 1
                  '   cnnConnection.Execute "update workday set wd01=wd01 where wd01=" & Combo1.Text & Format(i, "00") & Format(k, "00"), intI
                  '   '有資料才刪
                  '   If intI = 1 Then
                  '      strTxt(iTmp) = "DELETE FROM WORKDAY WHERE WD01=" & Combo1.Text & Format(i, "00") & Format(k, "00")
                  '      iTmp = iTmp + 1
                  '   End If
                  'Next
                  'end 2015/10/29
                  
                  cnnConnection.Execute "update workday set wd01=wd01 where wd01=" & Combo1.Text & Format(i, "00") & Format(varTmp(j), "00"), intI
                  '沒資料才新增
                  If intI = 0 Then
                     '工作天
                     strTxt(iTmp) = "INSERT INTO WORKDAY (WD01) VALUES (" & Combo1.Text & Format(i, "00") & Format(varTmp(j), "00") & ")"
                     iTmp = iTmp + 1
                  End If
               Next
               
            '整月不上班(不太可能發生)
            Else
               strTxt(iTmp) = "DELETE FROM WORKDAY WHERE SUBSTR(WD01,1,6)=" & Combo1.Text & Format(i, "00")
               iTmp = iTmp + 1
            End If
            'end 2014/11/26
         Next
         'edit by nickc 2007/02/09 不用 dll 了
         'If Not objLawDll.ExecSQL(iTmp - 1, strTxt) Then
         If Not ClsLawExecSQL(iTmp - 1, strTxt) Then
            Screen.MousePointer = vbDefault
            MsgBox Combo1.Text & " 年 " & i & " 月存檔失敗，請洽系統管理員 !", vbCritical
            Exit Sub
         Else
            PUB_WriteHoliday 'Added by Morgan 2020/9/8 更新門禁機假日表
            
            'Add By Sindy 2021/2/26 檢查星期六資料,更新為補班wd06=Y
            strSql = "update workday set wd06='Y'" & _
                     " WHERE wd01>=" & Combo1.Text & "0101 AND wd01<=" & Combo1.Text & "1231" & _
                     " and TO_CHAR(TO_DATE(wd01,'yyyymmdd'), 'D')=7"
            cnnConnection.Execute strSql, intI
            '2021/2/26 END
            
            MsgBox "存檔成功 ！", vbInformation
'            MsgBox "存檔成功 ！" & vbCrLf & vbCrLf & _
'                   "請確認【補班日】是否有加註為（補班wd06=Y）！", vbInformation
         End If
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Private Sub ShowCal(ByVal iYear As Integer)
 Dim i As Integer, j As Integer, strTmp As String
   Screen.MousePointer = vbHourglass
   For i = 1 To 12
      strExc(0) = "SELECT WD01 FROM WORKDAY WHERE SUBSTR(WD01,1,6)=" & Format(iYear) & Format(i, "00")
      intI = 1
      'edit by nickc 2007/02/09 不用 dll 了
      'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      strTmp = ""
      If intI = 1 Then
         With RsTemp
            Do While Not RsTemp.EOF
               strTmp = strTmp & Val(Right(.Fields(0), 2)) & ","
               .MoveNext
            Loop
         End With
      End If
      If Right(strTmp, 1) = "," Then strTmp = Left(strTmp, Len(strTmp) - 1)
      Cal1(i - 1).InitCalendar iYear, i, strTmp
   Next
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
 Dim i As Integer
 
   ' 90.07.16 modify by Ken (取得使用者執行各項功能的權限)
   m_bUpdate = IsUserHasRightOfFunction("frm12040137", strEdit, False)
   ' Ken 90.07.16 -- End
   
   MoveFormToCenter Me
   Combo1.Clear
   For i = 1974 To 2040
      Combo1.AddItem i
   Next
   ShowCal Year(Date)
   Combo1.Text = Year(Date)
   
   ' Ken 90.07.16 -- start
   If m_bUpdate Then
       Command1(0).Enabled = True
   Else
       Command1(0).Enabled = False
   End If
   ' Ken 90.07.16 -- End
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040137 = Nothing
End Sub
