VERSION 5.00
Begin VB.Form frm072003 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦/協辦人員案件查詢"
   ClientHeight    =   2160
   ClientLeft      =   1710
   ClientTop       =   795
   ClientWidth     =   4815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4815
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   0
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   1
      Top             =   660
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   1
      Left            =   3060
      MaxLength       =   6
      TabIndex        =   2
      Top             =   660
      Width           =   1332
   End
   Begin VB.OptionButton Option2 
      Caption         =   "協辦人員："
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   996
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "承  辦  人："
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   660
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   3
      Left            =   3060
      MaxLength       =   6
      TabIndex        =   5
      Top             =   996
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   4
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1428
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   5
      Left            =   3060
      MaxLength       =   7
      TabIndex        =   7
      Top             =   1428
      Width           =   1332
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3828
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   70
      Width           =   760
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3000
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   2
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   4
      Top             =   996
      Width           =   1332
   End
   Begin VB.Line Line3 
      X1              =   2820
      X2              =   3060
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收文日期："
      Height          =   180
      Left            =   396
      TabIndex        =   10
      Top             =   1428
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2820
      X2              =   3060
      Y1              =   1548
      Y2              =   1548
   End
   Begin VB.Line Line1 
      X1              =   2820
      X2              =   3060
      Y1              =   1116
      Y2              =   1116
   End
End
Attribute VB_Name = "frm072003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/09/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim blnCom1 As Boolean, blnCom2 As Boolean, blnCom3 As Boolean, blnCom4 As Boolean
Dim strMerger1 As String, strMerger2 As String


Private Sub cmdBack_Click()
   Unload Me
End Sub


Private Sub cmdSure_Click()
 Dim strSQL1 As String, strSQL2 As String
 Dim strCP01 As String
 Screen.MousePointer = 11
   If Option1.Value = True Then
      If ChkRange(Text(0), Text(1), "承辦人代號") = False Then
         Screen.MousePointer = 0
         Exit Sub
      End If
   End If
   If Option2.Value = True Then
      'Modified by Lydia 2015/10/05
      'If ChkRange(Text(2), Text(3), "法務人員代號") = False Then
      If ChkRange(Text(2), Text(3), "協辦人員代號") = False Then
         Screen.MousePointer = 0
         Exit Sub
      End If
   End If
   'Add By Cheng 2002/03/22
   If PUB_CheckKeyInDate(Me.Text(4)) = -1 Then
      Me.Text(4).SetFocus
      Text_GotFocus 4
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text(5)) = -1 Then
      Me.Text(5).SetFocus
      Text_GotFocus 5
      Exit Sub
   End If
   
   If CechkDate(Text(4), Text(5), "收文日期") = False Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   If Me.Tag = 0 Then
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      strCP01 = "CP01 IN ('L','LA','FCL','LIN','ACS')"
   ElseIf Me.Tag = 1 Then
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      strCP01 = "CP01 IN ('CFL','FCL','LIN')"
   End If
   MergerSQL
   
   'Modify By Cheng 2002/04/26
   '若已閉卷, 在本所案號後加"*"號
   'Modified by Lydia 2015/10/05
   strSQL1 = "select  cp01||'-'||cp02||decode(cp03,'0','','-'||cp03)||decode(cp04,'00','','-'||cp04)||DECODE(LC08,'Y','＊','')  本所案號," & _
      "substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 收文日," + _
       "decode(lc11,cu01||cu02,nvl(cu04,nvl(cu05,cu06))) 當事人," & _
       "decode(cp10,cpm02,cpm03,cpm04) 案件性質,decode(cp14,S1.ST01,S1.ST02) 承辦人," & _
       "decode(cp29,S2.ST01,S2.ST02) 協辦人員," + _
       "decode(cp27,null,'',substr(cp27,1,4)-1911||'/'||substr(cp27,5,2)||'/'||substr(cp27,7,2)) 發文日,cp09 " & _
       "from lawcase,caseprogress,STAFF S1,STAFF S2,casepropertymap,customer where " & strCP01 & _
       " AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 and cp01=cpm01(+) and " + _
       "CP14=S1.ST01(+) and CP29=S2.ST01(+) AND (substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+))  and " & _
       "cp10=cpm02(+)  and CP09<'C' AND " + strMerger1 + "" + strMerger2
   'Modified by Lydia 2015/10/05
   strSQL2 = "select cp01||'-'||cp02||decode(cp03,'0','','-'||cp03)||decode(cp04,'00','','-'||cp04)||DECODE(HC09,'Y','＊','')  本所案號,substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 收文日," + _
      " decode(hc05,cu01||cu02,nvl(cu04,nvl(cu05,cu06))) 當事人,decode(cp10,cpm02,cpm03,cpm04) 案件性質,decode(cp14,S1.ST01,S1.ST02) 承辦人,decode(cp29,S2.ST01,S2.ST02) 協辦人員," + _
      " decode(cp27,null,'',substr(cp27,1,4)-1911||'/'||substr(cp27,5,2)||'/'||substr(cp27,7,2)) 發文日,cp09 from hirecase,caseprogress,STAFF S1,STAFF S2,casepropertymap,customer" + _
      " where " & strCP01 & " AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 and cp01=cpm01(+) and CP14=S1.ST01(+) and CP29=S2.ST01(+) and substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+)  and cp10=cpm02(+) and CP09<'C' AND " + strMerger1 + strMerger2
   strExc(1) = strSQL1 + " Union " + strSQL2
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(1))
   If intI <> 1 Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   frm072004.Show
   frm072003.Hide
   Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
' Text(0).SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Text(2).Enabled = False
   Text(3).Enabled = False
   Text(0).Enabled = True
   Text(1).Enabled = True
End Sub

Private Sub ComEable()
'If (blnCom1 And blnCom2 And blnCom3 And blnCom4) Then cmdSure.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm072003 = Nothing
End Sub

Private Sub Option1_Click()
Option2.Value = False
Text(2).Enabled = False
Text(3).Enabled = False
Text(0).Enabled = True
Text(1).Enabled = True


End Sub
Private Sub Option2_Click()
Option1.Value = False
Text(0).Enabled = False
Text(1).Enabled = False
Text(2).Enabled = True
Text(3).Enabled = True

End Sub

Private Sub Text_GotFocus(Index As Integer)
Select Case Index
Case Index
 TextInverse Text(Index)
 End Select
End Sub

'Add By Sindy 2010/11/26
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case 4, 5
         KeyAscii = Pub_NumAscii(KeyAscii)
   End Select
End Sub

Private Sub Text_LostFocus(Index As Integer)
   Select Case Index
      Case 1, 3, 5
      If RunNick(Text(Index - 1), Text(Index)) Then
         Text(Index - 1).SetFocus
      End If
   End Select
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp As String
Select Case Index
Case 1, 3
'    If Text(Index) <> "" Then
'       If Not objPublicData.GetStaff(Text(Index), strTemp) Then Cancel = True
'    End If
 Case 0, 2
'   If Text(Index) <> "" Then
'      If Not objPublicData.GetStaff(Text(Index), strTemp) Then Cancel = True
'   End If
 
Case 4, 5
    If Text(Index) <> "" Then
        If CheckIsTaiwanDate(Text(Index)) Then
           'If Val(GetTaiwanTodayDate) - Val(Text(Index)) < 0 Then
           '    MsgBox "輸入日期大於系統日", vbCritical
           '    Cancel = True
           ' Else
           '    Cancel = False
           ' End If
        Else
           Cancel = True
        End If
    End If
End Select
If Cancel Then TextInverse Text(Index)
End Sub
Private Sub MergerSQL()
If Option1.Value Then
    If Text(0) <> "" Then
        strMerger1 = " (cp14  between '" + Text(0) + "' and '" + Text(1) + "') and "
     Else
         strMerger1 = " cp14 <='" + Text(1) + "' and "
     End If
 ElseIf Option2.Value Then
    If Text(2) <> "" Then
         strMerger1 = " (cp29  between '" + Text(2) + "' and '" + Text(3) + "') and "
     Else
        strMerger1 = " cp29 <= '" + Text(3) + "' and "
     End If
  End If
If Text(4) <> "" Then
        strMerger2 = " (cp05 between '" + ChangeTStringToWString(Text(4)) + "' and  '" + ChangeTStringToWString(Text(5)) + "')"
 Else
        strMerger2 = " cp05 < = '" + ChangeTStringToWString(Text(5)) + "'"
 End If

End Sub
Private Function AllTextBeforeSaveCheck() As Boolean
If Option1.Value Then
   If Text(1) = "" Then
      DataErrorMessage 5, "承辦人"
      Text(1).SetFocus
      Exit Function
   End If
ElseIf Option2.Value Then
   If Text(3) = "" Then
      'Modified by Lydia 2015/10/05
      'DataErrorMessage 5, "法務人員"
      DataErrorMessage 5, "協辦人員"
      Text(3).SetFocus
      Exit Function
   End If
End If
If Text(5) = "" Then
      DataErrorMessage 5, "收文日期"
      Text(5).SetFocus
      Exit Function

End If
AllTextBeforeSaveCheck = False
End Function

Private Function CechkDate(txt1 As TextBox, txt2 As TextBox, ByVal St As String) As Boolean
 On Error Resume Next
   CechkDate = True
   If txt2 = "" Or (txt1 = "" And txt2 = "") Then
      txt1.SetFocus
      MsgBox St & "不得為空值 !", vbCritical
      CechkDate = False
   ElseIf txt1 <> "" And txt2 <> "" Then
      If Val(txt1.Text) > Val(txt2.Text) Then
         MsgBox St & "範圍不正確 !", vbCritical
         txt1.SetFocus
         CechkDate = False
      End If
   End If
End Function
