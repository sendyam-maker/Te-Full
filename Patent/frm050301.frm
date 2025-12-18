VERSION 5.00
Begin VB.Form frm050301 
   BorderStyle     =   1  '單線固定
   Caption         =   "詢問進度函"
   ClientHeight    =   2856
   ClientLeft      =   2748
   ClientTop       =   3072
   ClientWidth     =   4584
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2856
   ScaleWidth      =   4584
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   2724
      MaxLength       =   2
      TabIndex        =   3
      Top             =   555
      Width           =   345
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   2484
      MaxLength       =   1
      TabIndex        =   2
      Top             =   555
      Width           =   225
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1776
      MaxLength       =   6
      TabIndex        =   1
      Top             =   555
      Width           =   750
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3690
      TabIndex        =   13
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2895
      TabIndex        =   12
      Top             =   90
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   1284
      MaxLength       =   1
      TabIndex        =   11
      Top             =   2070
      Width           =   275
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   2604
      MaxLength       =   7
      TabIndex        =   5
      Top             =   870
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1284
      MaxLength       =   7
      TabIndex        =   4
      Top             =   870
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1284
      MaxLength       =   3
      TabIndex        =   0
      Top             =   555
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2604
      MaxLength       =   9
      TabIndex        =   10
      Top             =   1770
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1284
      MaxLength       =   9
      TabIndex        =   9
      Top             =   1770
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1296
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1470
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2604
      MaxLength       =   9
      TabIndex        =   7
      Top             =   1170
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1284
      MaxLength       =   9
      TabIndex        =   6
      Top             =   1170
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   84
      TabIndex        =   19
      Top             =   600
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "發文日："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   84
      TabIndex        =   20
      Top             =   930
      Width           =   1035
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "7.答辯提申(通用) 8.年費提申"
      Height          =   180
      Left            =   1656
      TabIndex        =   23
      Top             =   2556
      Width           =   2232
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   180
      Left            =   2220
      TabIndex        =   22
      Top             =   1530
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "4.催領證 5.催證書 6.催讓渡"
      Height          =   180
      Left            =   1656
      TabIndex        =   21
      Top             =   2316
      Width           =   2112
   End
   Begin VB.Line Line3 
      X1              =   2250
      X2              =   2490
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Line Line2 
      X1              =   2250
      X2              =   2490
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Line Line1 
      X1              =   2250
      X2              =   2490
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "1.收達 2.申請日號 3.申請案結果 "
      Height          =   180
      Left            =   1656
      TabIndex        =   18
      Top             =   2076
      Width           =   2520
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "詢問函格式："
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   180
      TabIndex        =   17
      Top             =   2070
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   330
      TabIndex        =   16
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   330
      TabIndex        =   15
      Top             =   1470
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   330
      TabIndex        =   14
      Top             =   1170
      Width           =   720
   End
End
Attribute VB_Name = "frm050301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim strReceiveNo As String

Private Sub cmdOK_Click(Index As Integer)
Dim TmpCls As ClsSysName, i As Integer, strTmp As String
'Add By Cheng 2003/04/15
Dim strSQL1 As String, strSQL2 As String
Dim rsA As New ADODB.Recordset
'Add By Cheng 2003/06/18
Dim strCaseNo As String '本所案號
Dim stLetter As String 'Add by Morgan 2004/10/6
'Added by Morgan 2018/8/7
Dim bolSave As Boolean
Dim stCP09 As String, stCP10 As String
Dim strCon As String 'Added by Morgan 2023/11/14


   If Index = 1 Then Unload Me: Exit Sub
    Screen.MousePointer = vbHourglass
   If Txt1(11) = "" Then
      MsgBox "詢問函格式不得空白 !", vbCritical
      Txt1(11).SetFocus
        Screen.MousePointer = vbDefault
      Exit Sub
   End If
   If Txt1(0) <> "" And Txt1(1) <> "" Then
      If Left(Txt1(0), 6) <> Left(Txt1(1), 6) Then
         MsgBox "申請人前六碼必須相同 !", vbCritical
         Txt1(0).SetFocus
        Screen.MousePointer = vbDefault
         Exit Sub
      Else
         If Not ChkRange(Txt1(0), Txt1(1), "申請人") Then
            Txt1(0).SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      End If
   ElseIf (Txt1(0) = "" And Txt1(1) <> "") Or (Txt1(0) <> "" And Txt1(1) = "") Then
      MsgBox "申請人兩者皆須有值 !", vbCritical
      Txt1(0).SetFocus
        Screen.MousePointer = vbDefault
      Exit Sub
   End If
   If Txt1(3) <> "" And Txt1(4) <> "" Then
      If Left(Txt1(3), 6) <> Left(Txt1(4), 6) Then
         MsgBox "代理人前六碼必須相同 !", vbCritical
         Txt1(3).SetFocus
        Screen.MousePointer = vbDefault
         Exit Sub
      Else
         If Not ChkRange(Txt1(3), Txt1(4), "代理人") Then
            Txt1(3).SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      End If
   ElseIf (Txt1(3) = "" And Txt1(4) <> "") Or (Txt1(3) <> "" And Txt1(4) = "") Then
      MsgBox "代理人兩者皆須有值 !", vbCritical
      Txt1(3).SetFocus
        Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/3 清除查詢印表記錄檔欄位
   strExc(0) = ""
   'Add By Cheng 2003/04/15
   strSQL1 = "": strSQL2 = ""

   Select Case Me.Txt1(11).Text
'Modified by Morgan 2023/11/13 調整選項--玫音
'   Case "1" '申請日
'      strSQL1 = strSQL1 & " And PA10 IS NULL "
'      strSQL2 = strSQL2 & " And SP10 IS NULL "
'      pub_QL05 = pub_QL05 & ";" & Label6 & "1.申請日" 'Add By Sindy 2010/12/3
'   Case "2" '申請案號
'      strSQL1 = strSQL1 & " And PA11 IS NULL "
'      strSQL2 = strSQL2 & " And SP11 IS NULL "
'      pub_QL05 = pub_QL05 & ";" & Label6 & "2.申請案號" 'Add By Sindy 2010/12/3
'   Case "3" '申請案結果
'      strSQL1 = strSQL1 & " And (PA20 IS NULL  or pa20<>'1')"
'      strSQL2 = strSQL2
'      pub_QL05 = pub_QL05 & ";" & Label6 & "3.申請案結果" 'Add By Sindy 2010/12/3
'   Case "4" '發證日號
'      strSQL1 = strSQL1 & " And (PA21 IS NULL And PA22 IS NULL) "
'      strSQL2 = strSQL2 & " And (SP12 IS NULL And SP14 IS NULL) "
'      pub_QL05 = pub_QL05 & ";" & Label6 & "4.發證日號" 'Add By Sindy 2010/12/3
'   Case "5" '證書
'      strSQL1 = strSQL1 & " And (PA24 IS NULL And PA25 IS NULL) "
'      strSQL2 = strSQL2 & " And (SP20 IS NULL And SP21 IS NULL) "
'      pub_QL05 = pub_QL05 & ";" & Label6 & "5.證書" 'Add By Sindy 2010/12/3
'   Case "6"
'      pub_QL05 = pub_QL05 & ";" & Label6 & "6.是否收達" 'Add By Sindy 2010/12/3
'   Case "7"
'      pub_QL05 = pub_QL05 & ";" & Label6 & "7.催讓渡" 'Add By Sindy 2010/12/3
'   Case "8"
'      pub_QL05 = pub_QL05 & ";" & Label6 & "8.答辯提申" 'Added By Morgan 2012/4/18
   
   Case "1" '收達
      strCon = " and cp46 is null"
      pub_QL05 = pub_QL05 & ";" & Label6 & "1.收達"
   Case "2" '申請日號
      strCon = " and cp10 in (" & NewCasePtyList & ") and cp47 is null"
      strSQL1 = strSQL1 & " And (PA10 IS NULL or PA11 IS NULL) "
      strSQL2 = strSQL2 & " And (SP10 IS NULL or SP11 IS NULL) "
      pub_QL05 = pub_QL05 & ";" & Label6 & "2.申請日號"
   Case "3" '申請案結果
      'Modified by Morgan 2023/11/16 不限定性質(其他性質可修改使用)--禧佩
      'strCon = " and cp10 in (" & NewCasePtyList & ") and cp24 is null"
      strCon = " and cp24 is null"
      'end 2023/11/16
      strSQL1 = strSQL1 & " And (PA20 IS NULL  or pa20<>'1')"
      strSQL2 = strSQL2
      pub_QL05 = pub_QL05 & ";" & Label6 & "3.申請案結果"
   Case "4" '催領證
      strCon = " and cp10='601' and cp47 is null"
      pub_QL05 = pub_QL05 & ";" & Label6 & "4.催領證"
   Case "5" '催證書
      'strCon = " and cp10='601'" 'Removed by Morgan 2023/12/1 自動發證沒有領證--慧汶 Ex:CFP-33537
      strSQL1 = strSQL1 & " And (PA24 IS NULL And PA25 IS NULL) "
      strSQL2 = strSQL2 & " And (SP20 IS NULL And SP21 IS NULL) "
      pub_QL05 = pub_QL05 & ";" & Label6 & "5.催證書"
   Case "6" '催讓渡
      strCon = " and cp10='701'"
      pub_QL05 = pub_QL05 & ";" & Label6 & "6.催讓渡"
   Case "7" '答辯提申
      'Modified by Morgan 2023/11/15 不限定性質(他性質可修改使用)--玫音
      'strCon = " and cp10='107' and cp47 is null"
      strCon = " and cp47 is null"
      'end 2023/11/15
      pub_QL05 = pub_QL05 & ";" & Label6 & "7.答辯提申(通用)"
   Case "8" '年費提申
      'Modified by Morgan 2024/2/1 +606,607
      strCon = " and cp10 in ('605','606','607') and cp47 is null"
      pub_QL05 = pub_QL05 & ";" & Label6 & "8.年費提申"
'end 2023/11/13
   End Select
   
   '選擇本所案號
   If Option1(0).Value = True Then
      If Txt1(5) = "" Or Txt1(6) = "" Then
         MsgBox "本所案號錯誤，請重新輸入 !", vbCritical
         Txt1(5).SetFocus
         Screen.MousePointer = vbDefault
         Exit Sub
      Else
         If Txt1(7) = "" Then Txt1(7) = "0"
         If Txt1(8) = "" Then Txt1(8) = "00"
         pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Txt1(5) & "-" & Txt1(6) & "-" & Txt1(7) & "-" & Txt1(8) 'Add By Sindy 2010/12/3
         '判斷系統類別
         Select Case Txt1(5)
            Case "P", "CFP", "FCP"
               strExc(0) = "CP01='" & Txt1(5) & "' AND CP02='" & Txt1(6) & "' AND CP03='" & Txt1(7) & "' AND CP04='" & Txt1(8) & "' AND "
               If Txt1(0) <> "" And Txt1(1) <> "" Then
                  pub_QL05 = pub_QL05 & ";" & Label1 & Txt1(0) & "-" & Txt1(1) 'Add By Sindy 2010/12/3
                  strExc(2) = ChangeCustomerL(Txt1(0))
                  strExc(3) = ChangeCustomerL(Txt1(1))
                  strExc(0) = strExc(0) & "(PA26 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "' OR " & _
                     "PA27 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "' OR " & _
                     "PA28 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "' OR " & _
                     "PA29 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "' OR " & _
                     "PA30 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "') AND "
               End If
               If Txt1(2) <> "" Then
                  pub_QL05 = pub_QL05 & ";" & Label2 & Txt1(2) & Label5 'Add By Sindy 2010/12/3
                  strExc(0) = strExc(0) & "PA09='" & ChangeCustomerL(Txt1(2)) & "' AND "
               End If
               'Modified by Morgan 2023/11/14 +CP10
               strExc(0) = "SELECT CP09, PA09, PA08, CP01, CP02, CP03, CP04, CP05,CP10 FROM PATENT,CASEPROGRESS WHERE " & strExc(0) & "cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+)"
               '2014/3/5 modify by sonia 改A,B類原只抓A類 CFP-023721
               strExc(0) = strExc(0) & strSQL1 & " And CP09 < 'C' "
               'Modify by Morgan 2008/5/8 改抓最後發文的代理人
               'strExc(0) = strExc(0) & " Order By CP01, CP02, CP03, CP04, CP05, CP09 "
               'Modified by Morgan 2019/11/25 +判斷有代理人(有程序沒有CP44 Ex:CFP-031230繪圖超時)
               strExc(0) = strExc(0) & " and cp27>0 and cp44 is not null " & strCon & " Order By CP01, CP02, CP03, CP04, CP27 desc, CP09 desc"
               'end 2008/5/8
               
            Case Else
               strExc(0) = "CP01='" & Txt1(5) & "' AND CP02='" & Txt1(6) & "' AND CP03='" & Txt1(7) & "' AND CP04='" & Txt1(8) & "' AND "
               If Txt1(0) <> "" And Txt1(1) <> "" Then
                  pub_QL05 = pub_QL05 & ";" & Label1 & Txt1(0) & "-" & Txt1(1) 'Add By Sindy 2010/12/3
                  strExc(2) = ChangeCustomerL(Txt1(0))
                  strExc(3) = ChangeCustomerL(Txt1(1))
                  strExc(0) = strExc(0) & "(SPA08 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "' OR " & _
                     "SP58 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "' OR " & _
                     "SP59 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "') AND "
               End If
               If Txt1(2) <> "" Then
                  pub_QL05 = pub_QL05 & ";" & Label2 & Txt1(2) & Label5 'Add By Sindy 2010/12/3
                  strExc(0) = strExc(0) & "SP09='" & ChangeCustomerL(Txt1(2)) & "' AND "
               End If
               'Modified by Morgan 2023/11/14 +CP10
               strExc(0) = "SELECT CP09, SP09, '0', CP01, CP02, CP03, CP04, CP05,CP10 FROM SERVICEPRACTICE,CASEPROGRESS WHERE " & strExc(0) & "cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+)"
               '2014/3/5 modify by sonia 改A,B類原只抓A類 CFP-023721
               strExc(0) = strExc(0) & strSQL2 & " And CP09 < 'C' "
               'Modify by Morgan 2008/5/8 改抓最後發文的代理人
               'strExc(0) = strExc(0) & " Order By CP01, CP02, CP03, CP04, CP05, CP09 "
               'Modified by Morgan 2019/11/25 +判斷有代理人(有程序沒有CP44 Ex:CFP-031230繪圖超時)
               strExc(0) = strExc(0) & " and cp27>0 and cp44 is not null " & strCon & " Order By CP01, CP02, CP03, CP04, CP27 desc, CP09 desc"
               'end 2008/5/8
         End Select
      End If
   '選擇發文日區間
   Else
      'Add By Cheng 2002/03/18
      If PUB_CheckKeyInDate(Me.Txt1(9)) = -1 Then
         Me.Txt1(9).SetFocus
         txt1_GotFocus 9
        Screen.MousePointer = vbDefault
         Exit Sub
      End If
      If PUB_CheckKeyInDate(Me.Txt1(10)) = -1 Then
         Me.Txt1(10).SetFocus
         txt1_GotFocus 10
        Screen.MousePointer = vbDefault
         Exit Sub
      End If
      If (Txt1(9) = "" And Txt1(10) <> "") Or (Txt1(9) <> "" And Txt1(10) = "") Or (Txt1(9) = "" And Txt1(10) = "") Then
         MsgBox "發文日必須有值，請重新輸入 !", vbCritical
         Txt1(9).SetFocus
        Screen.MousePointer = vbDefault
         Exit Sub
      End If
      If ChkRange(Txt1(9), Txt1(10), "發文日") Then
         pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Txt1(9) & "-" & Txt1(10) 'Add By Sindy 2010/12/3
         strExc(5) = ""
         '取系統別
         strTmp = ""
         For Each TmpCls In ColSysName
            If InStr(1, strTmp, TmpCls.SysId) <= 0 Then
               strTmp = strTmp & "'" & TmpCls.SysId & "',"
            End If
         Next
         If Right(strTmp, 1) = "," Then strTmp = Left(strTmp, Len(strTmp) - 1)
         If strTmp <> "" Then
            strExc(4) = "PA01 IN (" & strTmp & ") AND "
            strExc(5) = "SP01 IN (" & strTmp & ") AND "
         End If
         If Txt1(0) <> "" And Txt1(1) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label1 & Txt1(0) & "-" & Txt1(1) 'Add By Sindy 2010/12/3
            strExc(2) = ChangeCustomerL(Txt1(0))
            strExc(3) = ChangeCustomerL(Txt1(1))
            strExc(4) = strExc(4) & "(PA26 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "' OR " & _
               "PA27 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "' OR " & _
               "PA28 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "' OR " & _
               "PA29 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "' OR " & _
               "PA30 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "') AND "
            strExc(5) = strExc(5) & "(SP08 BETWEEN '" & strExc(0) & "' AND '" & strExc(1) & "' OR " & _
               "SP58 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "' OR " & _
               "SP59 BETWEEN '" & strExc(2) & "' AND '" & strExc(3) & "') AND "
         End If
         If Txt1(2) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label2 & Txt1(2) & Label5 'Add By Sindy 2010/12/3
            strExc(4) = strExc(4) & "PA09='" & Txt1(2) & "' AND "
            strExc(5) = strExc(5) & "SP09='" & Txt1(2) & "' AND "
         End If
         '2014/3/5 modify by sonia 改A,B類原只抓A類 CFP-023721
         'Modified by Morgan 2023/11/14
         'strExc(0) = "SELECT CP09, PA09, PA08, CP01, CP02, CP03, CP04, CP05,CP27 FROM PATENT,CASEPROGRESS WHERE CP27 BETWEEN " & TransDate(Txt1(9), 2) & " AND " & TransDate(Txt1(10), 2) & " AND " & _
            strExc(4) & "cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) " & strSQL1 & " And CP09<'C' AND CP31='Y' UNION " & _
            "SELECT CP09, SP09, '0', CP01, CP02, CP03, CP04, CP05,CP27 FROM SERVICEPRACTICE,CASEPROGRESS WHERE CP27 BETWEEN " & TransDate(Txt1(9), 2) & " AND " & TransDate(Txt1(10), 2) & " AND " & _
            strExc(5) & "cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) " & strSQL2 & " And CP09<'C' AND CP31='Y' "
         strExc(0) = "SELECT CP09, PA09, PA08, CP01, CP02, CP03, CP04, CP05,CP27,CP10 FROM PATENT,CASEPROGRESS WHERE CP27 BETWEEN " & TransDate(Txt1(9), 2) & " AND " & TransDate(Txt1(10), 2) & " AND " & _
            strExc(4) & "cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) " & strSQL1 & " And CP09<'C' and cp27>0 and cp44 is not null " & strCon & " UNION " & _
            "SELECT CP09, SP09, '0', CP01, CP02, CP03, CP04, CP05,CP27,CP10 FROM SERVICEPRACTICE,CASEPROGRESS WHERE CP27 BETWEEN " & TransDate(Txt1(9), 2) & " AND " & TransDate(Txt1(10), 2) & " AND " & _
            strExc(5) & "cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) " & strSQL2 & " And CP09<'C' and cp27>0 and cp44 is not null "
         'end 2023/11/14
         'Modify by Morgan 2008/5/8改抓最後發文的代理人
         'strExc(0) = strExc(0) & " Order By CP01, CP02, CP03, CP04, CP05, CP09 "
         strExc(0) = strExc(0) & " Order By CP01, CP02, CP03, CP04, CP27 desc, CP09 desc"
         'end 2008/5/8
      Else
         Exit Sub
      End If
   End If
   strTmp = ""
   intI = 0
   Set rsA = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strCaseNo = ""
      InsertQueryLog (rsA.RecordCount) 'Add By Sindy 2010/12/3
      Do While Not rsA.EOF
        If strCaseNo <> "" & rsA("CP01").Value & rsA("CP02").Value & rsA("CP03").Value & rsA("CP04").Value Then
            strCaseNo = "" & rsA("CP01").Value & rsA("CP02").Value & rsA("CP03").Value & rsA("CP04").Value
            '判斷詢問函
             Select Case Txt1(11)
'Modified by Morgan 2023/11/13 調整選項--玫音
'                'Add By Cheng 2003/04/15
'                '加申請日條件
'                Case "1" '申請日
'                   If rsA.Fields(1) = 美國國家代號 Then
'                      strTmp = "01"
'                   Else
'                      strTmp = "02"
'                   End If
'                Case "2" '申請案號
'                   If rsA.Fields(1) = 美國國家代號 Then
'                      strTmp = "01"
'                   Else
'                      strTmp = "02"
'                   End If
'                Case "3" '申請案結果
'                   strTmp = "03"
'                Case "4" '發證日號
'                   '93.12.30 MODIFY BY SONIA
'                   'If rsA.Fields(1) = 美國國家代號 Then strTmp = "04"
'                   If rsA.Fields(1) = 美國國家代號 Then
'                     strTmp = "04"
'                   Else
'                     strTmp = "08"
'                   End If
'                   '93.12.30 END
'                Case "5" '證書
'                   If rsA.Fields(1) = 美國國家代號 Then
'                      strTmp = "05"
'                   Else
'                      strTmp = "06"
'                   End If
'               'Add by Morgan 2006/5/19
'               Case "6" '是否收達
'                  strTmp = "07"
'               'Add by Morgan 2010/1/12
'               Case "7" '催讓渡
'                  strTmp = "10"
'               'Added by Morgan 2012/4/18
'               Case "8" '答辯提申
'                  strTmp = "11"

               Case "1" '收達
                  strTmp = "07"
               Case "2" '申請日號
                  strTmp = "02"
               Case "3" '申請案結果
                  strTmp = "03"
               Case "4" '催領證
                  strTmp = "04"
               Case "5" '催證書
                  strTmp = "05"
               Case "6" '催讓渡
                  strTmp = "10"
               Case "7" '答辯提申
                  strTmp = "11"
               Case "8" '年費提申
                  strTmp = "04"
'end 2023/11/13
             End Select
             
            If strTmp <> "" Then
               strReceiveNo = rsA.Fields(0).Value
               'StartLetter "9", strTmp
                
               'Added by Morgan 2018/8/20 CFP電子化
               stCP09 = "": stCP10 = ""
               If rsA("CP01") = "CFP" And strSrvDate(1) >= CFP指示信電子化啟用日 Then
                  If FormSave(strReceiveNo, stCP09, stCP10, rsA.Fields("cp10")) = False Then Exit Sub
               End If
               'end 2018/8/20
               
               'Modify by Morgan 2004/10/6
               '加印傳真封面
               'NowPrint strReceiveNo, "09", strTmp, True, strUserNum, 0
               'Modify by Morgan 2004/10/13 傳真封面改用90(不預設Acknowledge)
               'NowPrint strReceiveNo, "01", "99", False, strUserNum, , , True, stLetter
               'Removed by Morgan 2018/10/22 取消傳真封面--慧汶
               'NowPrint strReceiveNo, "01", "90", False, strUserNum, , , True, stLetter, , , , , , , , , stCP09
               'If stCP09 <> "" Then Sleep 1000 '等1秒以確保letterdemand不會發生dupe錯誤 Added by Morgan 2018/8/20 CFP電子化
               'end 2018/10/22
               'Added by Morgan 2023/11/14 催審要抓該進度代理人故總收文號傳該催審程序,催收達/提申則維持被催的總收文號(因是對原代且指示信要抓發文日)
               If stCP10 = "411" Then
                  NowPrint stCP09, "09", strTmp, True, strUserNum, 0, stLetter, , , , , , , , , , , stCP09
               Else
               'end 2023/11/14
               
                  NowPrint strReceiveNo, "09", strTmp, True, strUserNum, 0, stLetter, , , , , , , , , , , stCP09
               End If
                  
               'Added by Morgan 2018/8/20 CFP電子化
               If stCP09 <> "" Then
                  frm1105_1.m_RecNo = stCP09
                  frm1105_1.m_PdfName = PUB_CaseNo2FileName(rsA("CP01"), rsA("CP02"), rsA("CP03"), rsA("CP04")) & "." & stCP10 & ".DATA.PDF"
                  frm1105_1.Show
               End If
               'end 2018/8/20
            End If
        End If
         rsA.MoveNext
      Loop
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/3
   End If
   Screen.MousePointer = vbDefault
End Sub

'Modified by Morgan 2023/11/10
Private Function FormSave(pCP09 As String, ByRef pNewCP09 As String, ByRef pNewCP10 As String, ByVal pCP10 As String) As Boolean
   Dim bolInTran As Boolean
   Dim stPA01 As String, stPA02 As String, stPA03 As String, stPA04 As String, stPA09 As String, stPA11 As String
   Dim stCP45 As String, stCP13 As String, stCP12 As String, stCP43 As String, stCP64 As String
   Dim stAF01 As String, stAF06 As String, stRefCP10 As String
   Dim strCon As String, strSubject As String
   'Added by Lydia 2019/02/13 CF案最新代理人
   Dim strCP44 As String, strCP45 As String
   
On Error GoTo ErrHandle
   
   Select Case Txt1(11)
'Modified by Morgan 2023/11/13 調整選項--玫音
'      Case "1"
'         pNewCP10 = "953" '催提申
'         stCP64 = "催申請日"
'         strCon = " and cp10 in (" & NewCasePtyList & ") and cp47 is null"
'      Case "2"
'         pNewCP10 = "953" '催提申
'         stCP64 = "催申請案號"
'         strCon = " and cp10 in (" & NewCasePtyList & ") and cp47 is null"
'      Case "3"
'         pNewCP10 = "411" '催審
'         stCP64 = "催申請案結果"
'         strCon = " and cp10 in (" & NewCasePtyList & ") and cp24 is null"
'      Case "4"
'         pNewCP10 = "411" '催審
'         stCP64 = "催發證日號"
'         strCon = " and cp10='601'"
'      Case "5"
'         pNewCP10 = "411" '催審
'         stCP64 = "催證書"
'         strCon = " and cp10='601'"
'      Case "6"
'        pNewCP10 = "952" '催收達
'        stCP64 = "催收達"
'        strCon = " and cp46 is null"
'      Case 7
'        pNewCP10 = "411" '催審
'        stCP64 = "催讓渡"
'        strCon = " and cp10='701'"
'      Case "8"
'        pNewCP10 = "953" '催提申
'        stCP64 = "催答辯提申"
'        strCon = " and cp10='107'"

      Case "1" '收達
         pNewCP10 = "952" '催收達
         stCP64 = "催收達"
         strCon = " and cp46 is null"
      Case "2" '申請日號
         pNewCP10 = "953" '催提申
         stCP64 = "催申請日號"
         strCon = " and cp10 in (" & NewCasePtyList & ") and cp47 is null"
      Case "3" '申請案結果
         pNewCP10 = "411" '催審
         stCP64 = "催申請案結果"
         strCon = " and cp10 in (" & NewCasePtyList & ") and cp24 is null"
      Case "4" '催領證
         pNewCP10 = "953" '催提申
         stCP64 = "催領證"
         strCon = " and cp10='601' and cp47 is null"
      Case "5" '催證書
         pNewCP10 = "411" '催審
         stCP64 = "催證書"
         strCon = " and cp10='601'"
      Case "6" '催讓渡
         pNewCP10 = "411" '催審
         stCP64 = "催讓渡"
         strCon = " and cp10='701'"
      Case "7" '答辯提申
         pNewCP10 = "953" '催提申
         stCP64 = "催答辯提申"
         strCon = " and cp10='107' and cp47 is null"
      Case "8" '年費提申
         pNewCP10 = "953" '催提申
         stCP64 = "催年費提申"
         'Modified by Morgan 2024/2/1 +606,607
         strCon = " and cp10 in ('605','606','607') and cp47 is null"
   End Select
      
   '讀取本所案號,申請國家,彼所案號
   '檢查是否已收文未發(只跑定稿)
   '檢查是否有指示信紀錄
   strExc(0) = "select a.cp45,pa01,pa02,pa03,pa04,pa09,pa11,b.cp09 Ncp09,c.cp09 Rcp09,c.cp10 Rcp10,af01" & _
      " from caseprogress a,patent,caseprogress b,caseprogress c,appform" & _
      " where a.cp09='" & pCP09 & "'" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04" & _
      " and b.cp01(+)=pa01 and b.cp02(+)=pa02 and b.cp03(+)=pa03 and b.cp04(+)=pa04" & _
      " and b.cp10(+)='" & pNewCP10 & "' and b.cp158(+)=0 and b.cp159(+)=0" & _
      " and c.cp09(+)=b.cp43 and af01(+)=b.cp09"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      stPA01 = RsTemp("pa01")
      stPA02 = RsTemp("pa02")
      stPA03 = RsTemp("pa03")
      stPA04 = RsTemp("pa04")
      stPA09 = RsTemp("pa09")
      stPA11 = "" & RsTemp("pa11")
      stCP45 = "" & RsTemp("cp45")
      
      pNewCP09 = "" & RsTemp("Ncp09")
      stCP43 = "" & RsTemp("Rcp09")
      stRefCP10 = "" & RsTemp("Rcp10")
      stAF01 = "" & RsTemp("af01")
   Else
      Exit Function
   End If
   
   '非自動收文
   If pNewCP09 = "" Then
      'Added by Morgan 2023/11/14
      stCP43 = pCP09
      stRefCP10 = pCP10
      'Modified by Morgan 2023/12/1 +催證書
      If Txt1(11) = "3" Or Txt1(11) = "5" Then '催申請案結果優先抓申請程序(因考慮其他性質催審也會用,前面抓的收文號沒有限制案件性質)
      'end 2023/11/14
         strExc(0) = "select cp09,cp10 from caseprogress where cp01='" & stPA01 & "' and cp02='" & stPA02 & "'" & _
            " and cp03='" & stPA03 & "' and cp04='" & stPA04 & "' and cp27>0" & strCon & _
            " order by cp27 desc,cp09 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            stCP43 = RsTemp("cp09")
            stRefCP10 = RsTemp("cp10")
         End If
      End If
      'end 2023/11/16
   End If
   
   cnnConnection.BeginTrans
   bolInTran = True
   
   If pNewCP09 = "" Then
      pNewCP09 = AutoNo("B", 6)
      stCP13 = PUB_GetAKindSalesNo(stPA01, stPA02, stPA03, stPA04)
      stCP12 = GetSalesArea(stCP13)
      'Modified by Lydia 2019/02/13 改抓最新進度檔最新一道程序的代理人(ex.CFP-026875)
      'strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
         "CP12,CP13,CP14,CP20,CP26,CP28,CP32,CP43,CP44,CP45,CP64)" & _
         " select cp01,cp02,cp03,cp04," & strSrvDate(1) & ",'" & pNewCP09 & "','" & pNewCP10 & "','" & stCP12 & "'" & _
         ",'" & stCP13 & "','" & strUserNum & "','N','N','" & pNewCP09 & "','N','" & stCP43 & "',cp44,cp45,'" & stCP64 & "'" & _
         " from caseprogress where cp09='" & pCP09 & "'"
      'Added by Morgan 2023/11/14
      '催審才要抓最後代理人, 催收達/提申應該對原代
      If pNewCP10 = "952" Or pNewCP10 = "953" Then
         strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
         "CP12,CP13,CP14,CP20,CP26,CP28,CP32,CP43,CP44,CP45,CP64)" & _
         " select cp01,cp02,cp03,cp04," & strSrvDate(1) & ",'" & pNewCP09 & "','" & pNewCP10 & "','" & stCP12 & "'" & _
         ",'" & stCP13 & "','" & strUserNum & "','N','N','" & pNewCP09 & "','N','" & stCP43 & "',cp44,cp45,'" & stCP64 & "'" & _
         " from caseprogress where cp09='" & pCP09 & "'"
         cnnConnection.Execute strSql, intI
      Else
      'end 2023/11/14
      
         Call PUB_GetCP44(stPA01, stPA02, stPA03, stPA04, strCP44, strExc(0), strCP45)
         strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
            "CP12,CP13,CP14,CP20,CP26,CP28,CP32,CP43,CP44,CP45,CP64)" & _
            " select cp01,cp02,cp03,cp04," & strSrvDate(1) & ",'" & pNewCP09 & "','" & pNewCP10 & "','" & stCP12 & "'" & _
            ",'" & stCP13 & "','" & strUserNum & "','N','N','" & pNewCP09 & "','N','" & stCP43 & "', '" & strCP44 & "', '" & strCP45 & "','" & stCP64 & "'" & _
            " from caseprogress where cp09='" & pCP09 & "'"
         'end 2019/02/13
         cnnConnection.Execute strSql, intI
         
      End If
   End If
   
   If stAF01 = "" Then
      stAF06 = PUB_GetLetterJudgeNew("2", "CFP", pNewCP10, stPA09, stRefCP10)
      strSubject = PUB_GetSubject(stPA01, stPA02, stPA03, stPA04, pNewCP10, stPA11, stCP45)
      PUB_AddAppForm pNewCP09, True, stAF06, strSubject  '自行判發,不轉檔(一定要先看過)
   End If
   
   cnnConnection.CommitTrans
   bolInTran = False
   
   FormSave = True
   Exit Function
   
ErrHandle:
   If bolInTran Then cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   Option1_Click 0
   Label5.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm050301 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
 Dim i As Integer
On Error Resume Next
   If Index = 0 Then
      For i = 5 To 8
         Txt1(i).Enabled = True
      Next
      Txt1(9).Enabled = False
      Txt1(10).Enabled = False
      Txt1(5).SetFocus
   Else
      For i = 5 To 8
         Txt1(i).Enabled = False
      Next
      Txt1(9).Enabled = True
      Txt1(10).Enabled = True
      Txt1(9).SetFocus
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse Txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 3, 4, 5, 7
         KeyAscii = UpperCase(KeyAscii)
      Case 11 '詢問函格式
         'Modify by Morgan 2006/5/19 加6
         'If (KeyAscii < 49 Or KeyAscii > 53) And KeyAscii <> 8 Then
         'Modify by Morgan 2006/5/19 加7
         'Modify by Morgan 2012/4/18 加8
         If (KeyAscii < 49 Or KeyAscii > 56) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Select Case Index
      Case 1
         If Txt1(0).Text > Txt1(1).Text Then
            MsgBox "申請人範圍錯誤 !", vbCritical
            Txt1(0).SetFocus
            TextInverse Txt1(0)
         End If
      Case 4
         If Txt1(3).Text > Txt1(4).Text Then
            MsgBox "代理人範圍錯誤 !", vbCritical
            Txt1(3).SetFocus
            TextInverse Txt1(3)
         End If
      Case 10
         If IsEmpty(Txt1(10)) = False Then
            If CheckIsTaiwanDate(Txt1(10), False) = False Then
               strMsg = "請輸入正確的發文日 !"
               strTit = "資料檢核"
               nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
               Txt1(10).SetFocus
               TextInverse Txt1(10)
            Else
               If Not ChkRange(Txt1(9), Txt1(10), "發文日") Then
               
               End If
            End If
         Else
            strMsg = "發文日必須輸入"
            strTit = "檢核輸入"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Txt1(9).SetFocus
         End If
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If Txt1(Index).Text = "" Then Exit Sub
   If Index = 10 Then Exit Sub
   Select Case Index
      Case 2
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(txt1(Index).Text, strExc(0)) Then
         If ClsPDGetNation(Txt1(Index).Text, strExc(0)) Then
            Label5.Caption = strExc(0)
         Else
            Label5.Caption = ""
            Cancel = True
         End If
      Case 5
         Cancel = Not ChkSysName(Txt1(Index))
      Case 9
         Cancel = Not ChkDate(Txt1(Index).Text)
   End Select
   If Cancel = True Then TextInverse Txt1(Index)
End Sub
