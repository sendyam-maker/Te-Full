VERSION 5.00
Begin VB.Form frm073009 
   BorderStyle     =   1  '單線固定
   Caption         =   "顧問地址條"
   ClientHeight    =   2355
   ClientLeft      =   1530
   ClientTop       =   1920
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4770
   Begin VB.CheckBox Check1 
      Caption         =   "含寄顧問電子報客戶"
      Height          =   255
      Left            =   384
      TabIndex        =   3
      Top             =   1740
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1344
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1245
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2700
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1245
      Width           =   975
   End
   Begin VB.ComboBox cboDep 
      Height          =   300
      ItemData        =   "frm073009.frx":0000
      Left            =   1344
      List            =   "frm073009.frx":0013
      Style           =   2  '單純下拉式
      TabIndex        =   0
      Top             =   708
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3708
      TabIndex        =   5
      Top             =   70
      Width           =   760
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2880
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "統計顧問客戶數時要勾選此項"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   648
      TabIndex        =   8
      Top             =   2052
      Width           =   2340
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2640
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所         別："
      Height          =   180
      Index           =   1
      Left            =   384
      TabIndex        =   7
      Top             =   756
      Width           =   948
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "到期日期 ："
      Height          =   180
      Index           =   0
      Left            =   384
      TabIndex        =   6
      Top             =   1290
      Width           =   945
   End
End
Attribute VB_Name = "frm073009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/11 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

'Add By Cheng 2002/09/09
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
    'Add By Cheng 2003/04/17
    Screen.MousePointer = vbHourglass
   'Add By Cheng 2002/09/09
   blnClkSure = False
   'Add By Cheng 2002/03/22
   If PUB_CheckKeyInDate(Me.Text1(0)) = -1 Then
      Me.Text1(0).SetFocus
      Text1_GotFocus 0
        'Add By Cheng 2003/04/17
        Screen.MousePointer = vbDefault
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text1(1)) = -1 Then
      Me.Text1(1).SetFocus
      Text1_GotFocus 1
        'Add By Cheng 2003/04/17
        Screen.MousePointer = vbDefault
      Exit Sub
   End If
   'Add By Cheng 2002/09/09
   If Me.Text1(0).Text <> "" And Me.Text1(1).Text <> "" Then
      If Val(Me.Text1(0).Text) > Val(Me.Text1(1).Text) Then
         MsgBox "到期日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         blnClkSure = True
         Me.Text1(0).SetFocus
         Text1_GotFocus 0
            'Add By Cheng 2003/04/17
            Screen.MousePointer = vbDefault
         Exit Sub
      End If
   End If
   
'edit by nick 2004/10/29
'   strExc(0) = "SELECT '',DECODE(CP13,ST01,ST02),hc05," & _
'      "DECODE(hc05,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
'      " MIN(SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2)),cp01||'-'||cp02||'-'||cp03||'-'||cp04, CP12 " & _
'      "FROM CASEPROGRESS,HIRECASE,CUSTOMER,STAFF WHERE CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04" & _
'      strGetcdnSQL & " AND (SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+)) AND " & _
'      "CP13=ST01(+) " & _
'      " GROUP BY DECODE(CP13,ST01,ST02),hc05,DECODE(hc05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))),cp01||'-'||cp02||'-'||cp03||'-'||cp04,CP12, Cp13 " & _
'      " ORDER BY CP12, CP13, HC05"
   'Modify By Sindy 2011/2/11
'   strExc(0) = "SELECT '',DECODE(CP13,ST01,ST02),hc05," & _
'      "DECODE(hc05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06)))," & _
'      " Max(SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2)),cp01||'-'||cp02||'-'||cp03||'-'||cp04,CP12,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))) " & _
'      "FROM CASEPROGRESS,HIRECASE,CUSTOMER C1,CUSTOMER C2,STAFF WHERE HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) " & _
'      strGetcdnSQL & " AND (SUBSTR(HC05,1,8)=C1.CU01(+) AND SUBSTR(HC05,9,1)=C1.CU02(+)) AND (SUBSTR(HC24,1,8)=C2.CU01(+) AND SUBSTR(HC24,9,1)=C2.CU02(+)) AND " & _
'      "CP13=ST01(+) and cp27 is null " & _
'      " GROUP BY DECODE(CP13,ST01,ST02),hc05,DECODE(hc05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))),cp01||'-'||cp02||'-'||cp03||'-'||cp04,CP12,Cp13,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))) " & _
'      " ORDER BY CP12, CP13, HC05"
   'Modify By Sindy 2011/3/17
   'Modify by Amy 2022/07/27 排除X03072-台一國際智慧財產事務所 (因1110425 將不寄顧問電子報由Y改為null,11107月顧問地址條會印出X03072010,故排除X03072關系企業-秀玲)
   strExc(0) = "SELECT '',DECODE(CP13,ST01,ST02),decode(HC05,'X65299000','●'||HC24,HC05)," & _
      "NVL(C1.CU04,NVL(C1.CU05,C1.CU06))," & _
      " Max(SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2)),cp01||'-'||cp02||'-'||cp03||'-'||cp04,CP12 " & _
      "FROM CASEPROGRESS,HIRECASE,CUSTOMER C1,STAFF WHERE HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) " & _
      strGetcdnSQL & " AND (SUBSTR(decode(HC05,'X65299000',HC24,HC05),1,8)=C1.CU01(+) AND SUBSTR(decode(HC05,'X65299000',HC24,HC05),9,1)=C1.CU02(+)) " & _
      IIf(Check1.Value = 0, " AND (C1.CU153='N' or C1.CU153 is null) ", "") & _
      " AND CP13=ST01(+) and cp27 is null And Substr(c1.cu01,1,6)<>'X03072' " & _
      " GROUP BY DECODE(CP13,ST01,ST02),hc05,NVL(C1.CU04,NVL(C1.CU05,C1.CU06)),cp01||'-'||cp02||'-'||cp03||'-'||cp04,CP12,Cp13,HC24 " & _
      " ORDER BY CP12, CP13, HC05"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      frm073010.Show
      Me.Hide
   End If
    'Add By Cheng 2003/04/17
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   cboDep.ListIndex = 0
End Sub

Private Function strGetcdnSQL() As String
 Dim St As String
   '2007/11/1 modify by sonia 所別改判斷ST06,不抓CP12
   Select Case cboDep.ListIndex
      Case 0
         'St = " AND SUBSTR(CP12,1,2) NOT IN ('S2','S3','S4')"
         St = " AND ST06='1'"
      Case 1
         'St = " AND SUBSTR(CP12,1,2)='S2'"
         St = " AND ST06='2'"
      Case 2
         'St = " AND SUBSTR(CP12,1,2)='S3'"
         St = " AND ST06='3'"
      Case 3
         'St = " AND SUBSTR(CP12,1,2)='S4'"
         St = " AND ST06='4'"
      Case 4
         St = ""
   End Select
   strExc(1) = St
   If Text1(0) = "" And Text1(1) <> "" Then
      strExc(1) = strExc(1) & " AND CP54<=" & Text1(1)
   ElseIf Text1(0) <> "" And Text1(1) <> "" Then
      strExc(1) = strExc(1) & " AND (CP54 BETWEEN " & ChangeTStringToWString(Text1(0)) & " AND " & ChangeTStringToWString(Text1(1)) & ")"
   'Add By Cheng 2002/03/22
   ElseIf Text1(0) <> "" And Text1(1) = "" Then
      strExc(1) = strExc(1) & " AND (CP54 BETWEEN " & ChangeTStringToWString(Text1(0)) & " AND " & ServerDate & ")"
   End If
   strExc(1) = strExc(1) & " AND CP10='0' AND CP57 IS NULL"
   strGetcdnSQL = strExc(1)
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm073009 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
      Case 1
         'Add/Modify By Cheng 2002/09/09
         If blnClkSure = False Then
            If RunNick(Text1(Index - 1), Text1(Index)) Then
               Text1(Index - 1).SetFocus
            End If
         Else
            blnClkSure = False
         End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Text1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0, 1
         If CheckIsTaiwanDate(Text1(Index)) = False Then Cancel = True
      End Select
      If Cancel Then TextInverse Text1(Index)
End Sub
