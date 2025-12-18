VERSION 5.00
Begin VB.Form frm100129_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外部業務消長分析表"
   ClientHeight    =   5820
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5500
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   31
      Left            =   1485
      MaxLength       =   1
      TabIndex        =   22
      Top             =   5400
      Width           =   405
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   22
      Left            =   1485
      MaxLength       =   3
      TabIndex        =   21
      Top             =   5040
      Width           =   405
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   21
      Left            =   1575
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2730
      Width           =   405
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   20
      Left            =   1485
      MaxLength       =   1
      TabIndex        =   20
      Top             =   4710
      Width           =   405
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   19
      Left            =   1485
      MaxLength       =   1
      TabIndex        =   19
      Top             =   4380
      Width           =   405
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   18
      Left            =   2475
      MaxLength       =   9
      TabIndex        =   18
      Top             =   4050
      Width           =   1035
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   17
      Left            =   1170
      MaxLength       =   9
      TabIndex        =   17
      Top             =   4050
      Width           =   1035
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   16
      Left            =   2475
      MaxLength       =   9
      TabIndex        =   16
      Top             =   3720
      Width           =   1035
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   15
      Left            =   1170
      MaxLength       =   9
      TabIndex        =   15
      Top             =   3720
      Width           =   1035
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   14
      Left            =   1575
      MaxLength       =   1
      TabIndex        =   14
      Top             =   3375
      Width           =   405
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   1170
      TabIndex        =   13
      Top             =   3060
      Width           =   1980
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   11
      Top             =   2400
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   1575
      MaxLength       =   4
      TabIndex        =   10
      Top             =   2400
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   2340
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2070
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1395
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2070
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   2115
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1740
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1170
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1740
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2115
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1410
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1170
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1410
      Width           =   630
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1170
      TabIndex        =   3
      Text            =   "ALL"
      Top             =   1080
      Width           =   2985
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1170
      MaxLength       =   5
      TabIndex        =   1
      Top             =   750
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2385
      MaxLength       =   5
      TabIndex        =   2
      Top             =   750
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   0
      Top             =   420
      Width           =   492
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   4485
      TabIndex        =   24
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   3585
      TabIndex        =   23
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "FCP工程師組別："
      Height          =   180
      Left            =   0
      TabIndex        =   48
      Top             =   5460
      Width           =   1380
   End
   Begin VB.Label lblName 
      Height          =   180
      Left            =   1980
      TabIndex        =   47
      Top             =   5460
      Width           =   1440
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "(  Ex.: 101-103,125,301-309 )"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   3240
      TabIndex        =   46
      Top             =   3120
      Width           =   2145
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   " (ALL：全部)"
      Height          =   180
      Left            =   4230
      TabIndex        =   45
      Top             =   1122
      Width           =   1035
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "增減件數："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   44
      Top             =   5085
      Width           =   900
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "( A.律師事務所 B.公司直接委辦 C.其他 )"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2070
      TabIndex        =   43
      Top             =   2775
      Width           =   3135
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人性質："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   42
      Top             =   2775
      Width           =   1290
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "( 1.洲 2.國家 3.FC代理人 4.申請人 )"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1935
      TabIndex        =   41
      Top             =   4770
      Width           =   2715
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "( 1.FC代理人性質 2.系統別 )"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1935
      TabIndex        =   40
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "縱向統計方式："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   39
      Top             =   4755
      Width           =   1260
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "橫向統計方式："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   38
      Top             =   4425
      Width           =   1260
   End
   Begin VB.Line Line8 
      X1              =   2250
      X2              =   2370
      Y1              =   4170
      Y2              =   4170
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   37
      Top             =   4095
      Width           =   930
   End
   Begin VB.Line Line7 
      X1              =   2250
      X2              =   2370
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   36
      Top             =   3765
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "專利/商標種類："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   35
      Top             =   3420
      Width           =   1305
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   34
      Top             =   3105
      Width           =   900
   End
   Begin VB.Line Line4 
      X1              =   2295
      X2              =   2415
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人國籍："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   33
      Top             =   2445
      Width           =   1290
   End
   Begin VB.Line Line3 
      X1              =   2115
      X2              =   2235
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請人國籍："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   32
      Top             =   2115
      Width           =   1080
   End
   Begin VB.Line Line2 
      X1              =   1890
      X2              =   2010
      Y1              =   1865
      Y2              =   1865
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   31
      Top             =   1782
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   1890
      X2              =   2010
      Y1              =   1535
      Y2              =   1535
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "業務區別："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   30
      Top             =   1452
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   180
      TabIndex        =   29
      Top             =   1122
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   28
      Top             =   792
      Width           =   540
   End
   Begin VB.Line Line5 
      X1              =   2145
      X2              =   2265
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "輸入民國年月"
      Height          =   180
      Left            =   3420
      TabIndex        =   27
      Top             =   792
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查詢別："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   26
      Top             =   462
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "（1. 收文   2. 發文）"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1755
      TabIndex        =   25
      Top             =   465
      Width           =   1575
   End
End
Attribute VB_Name = "frm100129_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/ Form2.0不用改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Create by Morgan 2010/9/2
Option Explicit

Public cmdState As Integer
Dim stSys(3) As String

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   If Left(Pub_StrUserSt03, 2) = "F1" Then
      txt1(4) = "F10"
      txt1(5) = "F19"
      'Modified by Morgan 2014/2/11 改組合條件
      txt1(12) = "101"
      'txt1(13) = "101"
      'end 2014/2/11
      
   'ElseIf Left(Pub_StrUserSt03, 2) = "F2" Then
   Else
   
      txt1(4) = "F20"
      txt1(5) = "F29"
      'Modified by Morgan 2014/2/11 改組合條件
      'txt1(12) = "101"
      'txt1(13) = "103"
      txt1(12) = "101-103,125,301-309"
      'end 2014/2/11
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100129_1 = Nothing
End Sub

Sub PubShowNextData()
   Select Case cmdState
      Case 0
         cmdState = -1
         Screen.MousePointer = vbHourglass
         DoEvents
         If ConstrainCheck = True Then
            ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/16 清除查詢印表記錄檔欄位
            doQuery
         End If
         Screen.MousePointer = vbDefault
      Case 1
         fnCloseAllFrm100
      Case Else
   End Select
End Sub

Private Function ConstrainCheck() As Boolean
   If txt1(0) = "" Then
      MsgBox "查詢別不可空白!!!"
      txt1(0).SetFocus
      Exit Function
   End If
   If txt1(1) = "" Or txt1(2) = "" Then
      MsgBox "日期條件不完整!!!"
      If txt1(1) = "" Then txt1(1).SetFocus: Exit Function
      If txt1(2) = "" Then txt1(2).SetFocus: Exit Function
   ElseIf Not ChkDate(txt1(1) & "01") Then
      txt1(1).SetFocus: Exit Function
   ElseIf Not ChkDate(txt1(2) & "01") Then
      txt1(2).SetFocus: Exit Function
   End If
   If txt1(3) = "" Then
      MsgBox "系統類別不可空白!!!"
      txt1(3).SetFocus: Exit Function
   End If
   'Modified by Morgan 2014/2/11 改組合條件
   'If txt1(12) = "" Or txt1(13) = "" Then
   '   MsgBox "案件性質條件不完整!!!"
   If txt1(12) = "" Then
      MsgBox "案件性質不可空白!!!"
   'end 2014/2/11
      If txt1(12) = "" Then txt1(12).SetFocus: Exit Function
   '   If txt1(13) = "" Then txt1(13).SetFocus: Exit Function
   End If
   
   If (txt1(15) <> "" And txt1(16) = "") Or (txt1(15) = "" And txt1(16) <> "") Then
      MsgBox "申請人條件不完整!!!"
      If txt1(15) = "" Then txt1(15).SetFocus: Exit Function
      If txt1(16) = "" Then txt1(16).SetFocus: Exit Function
   End If
   If (txt1(17) <> "" And txt1(18) = "") Or (txt1(17) = "" And txt1(18) <> "") Then
      MsgBox "FC申請人條件不完整!!!"
      If txt1(17) = "" Then txt1(17).SetFocus: Exit Function
      If txt1(18) = "" Then txt1(18).SetFocus: Exit Function
   End If
   
   If txt1(19) = "" Then
      MsgBox "請輸入橫向統計方式!!!"
      txt1(19).SetFocus: Exit Function
   ElseIf txt1(19) = "2" Then
      If txt1(3) = "ALL" Or SysCount(txt1(3)) > 2 Then
         MsgBox "橫向統計方式選 [ 2.系統別 ] 時請輸入 3 個以下系統別!!!" & vbCrLf & "(Ex.FCP,P,CFP)"
         txt1(3).SetFocus: Exit Function
      End If
   End If
   If txt1(20) = "" Then
      MsgBox "請輸入縱向統計方式!!!"
      txt1(20).SetFocus: Exit Function
   End If
   ConstrainCheck = True
End Function

Private Function SysCount(p_Sys As String) As Integer
   Dim arr1
   arr1 = Split(p_Sys, ",")
   SysCount = UBound(arr1)
End Function

Private Function ChkTbl(p_Sys As String, p_tbid As String) As Boolean
   If p_Sys = "ALL" Then
      ChkTbl = True: Exit Function
   End If
   strExc(0) = "select * from systemkind where sk02='" & p_tbid & "' and instr('," & p_Sys & ",',','||sk01||',')>0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ChkTbl = True
   End If
End Function
'Modified by Morgan 2015/1/19 1.三年度任一有案件者都要列出 2.排序當年度增減後再加上一年度增減 3.互惠加判斷系統別
Private Sub doQuery()
   Dim stCon As String, stConCP12 As String, stDate1 As String, stDate2 As String, ii As Integer, jj As Integer
   Dim stTable As String, stTableX As String, stTableY As String
   Dim arTable(3) As String
   Dim strSystemKind As String
   Dim stConPA As String, stConTM As String, stConLC As String, stConSP As String
   Dim stConCu As String, stConFa As String
   Dim arConDate(3) As String
   Dim stSelect As String, stSelectX As String, stSelectV As String
   Dim stWhere As String, stGroupByV As String, stOrderBy As String
   Dim arr1() As String, iPos As Integer, stConCP10 As String
   Dim fc04 As String, fc05 As String, stConFC06 As String
   
   '互惠年度
   fc04 = strSrvDate(2) \ 10000
   '互惠期間
   If Val(Right(strSrvDate(2), 4)) < 700 Then
      fc05 = "1" '上半年
   Else
      fc05 = "2" '下半年
   End If
   
   stTable = ""
   Erase arTable
   
   stDate1 = DBDATE(txt1(1) & "01")
   'Modified by Morgan 2013/3/5 固定用最大日期 31 否則 2月若為閏年會少統計
   'stDate2 = CompDate(2, -1, CompDate(1, 1, txt1(2) & "01"))
   stDate2 = (txt1(2) + 191100) & "31"
   
   stCon = ""
   Erase stSys
   jj = 1
   If txt1(3) <> "ALL" Then
      arr1 = Split(txt1(3), ",")
      strSystemKind = ""
      For ii = LBound(arr1) To UBound(arr1)
         strSystemKind = strSystemKind & "'" & arr1(ii) & "',"
         If arr1(ii) <> "" And jj <= 3 Then
            stSys(jj) = arr1(ii)
            jj = jj + 1
         End If
      Next ii
      strSystemKind = Left(strSystemKind, Len(strSystemKind) - 1)
      stCon = stCon & " AND CP01||'' IN ( " & strSystemKind & " ) "
      
      'Added by Morgan 2015/1/19
      If txt1(19) = "2" Then
         stConFC06 = " and fc06 IN ( " & strSystemKind & " ) "
      End If
      'end 2015/1/19
   End If
   If Len(txt1(3)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label4 & txt1(3)  'Add By Sindy 2010/11/16
   End If
   
   Erase arConDate
   If stDate1 <> "" Then
      If txt1(0) = "1" Then
         arConDate(1) = arConDate(1) & " and cp05>=" & stDate1
         arConDate(2) = arConDate(2) & " and cp05>=" & (Val(stDate1) - 10000)
         arConDate(3) = arConDate(3) & " and cp05>=" & (Val(stDate1) - 20000)
      Else
         arConDate(1) = arConDate(1) & " and cp27>=" & stDate1
         arConDate(2) = arConDate(2) & " and cp27>=" & (Val(stDate1) - 10000)
         arConDate(3) = arConDate(3) & " and cp27>=" & (Val(stDate1) - 20000)
      End If
   End If
   If stDate2 <> "" Then
      If txt1(0) = "1" Then
         arConDate(1) = arConDate(1) & " and cp05<=" & stDate2
         arConDate(2) = arConDate(2) & " and cp05<=" & (Val(stDate2) - 10000)
         arConDate(3) = arConDate(3) & " and cp05<=" & (Val(stDate2) - 20000)
      Else
         arConDate(1) = arConDate(1) & " and cp27<=" & stDate2
         arConDate(2) = arConDate(2) & " and cp27<=" & (Val(stDate2) - 10000)
         arConDate(3) = arConDate(3) & " and cp27<=" & (Val(stDate2) - 20000)
      End If
   End If
   If stDate1 <> "" Or stDate2 <> "" Then
      If txt1(0) = "1" Then
         pub_QL05 = pub_QL05 & ";收文" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/11/16
      Else
         pub_QL05 = pub_QL05 & ";發文" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/11/16
      End If
   End If
   
   '業務區
   stConCP12 = ""
   If txt1(4) <> "" Then
      stConCP12 = stConCP12 & " and cp12>='" & txt1(4) & "'"
   End If
   If txt1(5) <> "" Then
      stConCP12 = stConCP12 & " and cp12<='" & txt1(5) & "'"
   End If
   If txt1(4) <> "" Or txt1(5) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label6 & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/11/16
   End If
   
   '申請國家
   If txt1(6) <> "" Then
      stConPA = stConPA & " and pa09>='" & txt1(6) & "'"
      stConTM = stConTM & " and tm10>='" & txt1(6) & "'"
      stConSP = stConSP & " and sp09>='" & txt1(6) & "'"
      stConLC = stConLC & " and lc15>='" & txt1(6) & "'"
   End If
   If txt1(7) <> "" Then
      stConPA = stConPA & " and pa09<='" & txt1(7) & "'"
      stConTM = stConTM & " and tm10<='" & txt1(7) & "'"
      stConSP = stConSP & " and sp09<='" & txt1(7) & "'"
      stConLC = stConLC & " and lc15<='" & txt1(7) & "'"
   End If
   If txt1(6) <> "" Or txt1(7) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label3 & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/11/16
   End If
   
   If txt1(31) <> "" Then stConPA = stConPA & " and pa150='" & txt1(31) & "'"   'ADD BY SONIA 2014/7/30
   
   '申請人國籍
   If txt1(8) <> "" Then
      stConCu = stConCu & " and cu10>='" & txt1(8) & "'"
   End If
   If txt1(9) <> "" Then
      'Modified by Morgan 2021/3/18 國籍第4碼改有英文
      'stConCu = stConCu & " and cu10<='" & txt1(9) & "9'"
      stConCu = stConCu & " and cu10<='" & txt1(9) & "Z'"
   End If
   If txt1(8) <> "" Or txt1(9) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label9 & txt1(8) & "-" & txt1(9) 'Add By Sindy 2010/11/16
   End If
   
   'FC代理人國籍
   If txt1(10) <> "" Then
      stConFa = stConFa & " and fa10>='" & txt1(10) & "'"
   End If
   If txt1(11) <> "" Then
      'Modified by Morgan 2021/3/18 國籍第4碼改有英文
      'stConFa = stConFa & " and fa10<='" & txt1(11) & "9'"
      stConFa = stConFa & " and fa10<='" & txt1(11) & "Z'"
   End If
   If txt1(10) <> "" Or txt1(11) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label10 & txt1(10) & "-" & txt1(11) 'Add By Sindy 2010/11/16
   End If
   
   'FC代理人性質
   If txt1(21) <> "" Then
      stConFa = stConFa & " and fa76='" & txt1(21) & "'"
      pub_QL05 = pub_QL05 & ";" & Label7 & txt1(21) & Label20 'Add By Sindy 2010/11/16
   End If
   
   '案件性質
'Modified by Morgan 2014/2/11 改組合條件
'   If txt1(12) <> "" Then
'      stCon = stCon & " and cp10>='" & txt1(12) & "'"
'   End If
'   If txt1(13) <> "" Then
'      stCon = stCon & " and cp10<='" & txt1(13) & "'"
'   End If
'   If txt1(12) <> "" Or txt1(13) <> "" Then
'      pub_QL05 = pub_QL05 & ";" & Label11 & txt1(12) & "-" & txt1(13) 'Add By Sindy 2010/11/16
'   End If
   If txt1(12) <> "" Then
      arr1 = Split(txt1(12), ",")
      stConCP10 = ""
      For ii = LBound(arr1) To UBound(arr1)
         If arr1(ii) <> "" Then
            iPos = InStr(arr1(ii), "-")
            If stConCP10 <> "" Then stConCP10 = stConCP10 & " or "
            If iPos > 0 Then
               strExc(0) = Trim(Left(arr1(ii), iPos - 1))
               strExc(1) = Trim(Mid(arr1(ii), iPos + 1))
               stConCP10 = stConCP10 & " (cp10>='" & strExc(0) & "' and cp10<='" & strExc(1) & "')"
            Else
               stConCP10 = stConCP10 & " cp10='" & Trim(arr1(ii)) & "'"
            End If
         End If
      Next
      If stConCP10 <> "" Then
         stCon = stCon & " and (" & stConCP10 & ")"
      End If
   End If
'end 2014/2/11
   
   '專利/商標種類
   If txt1(14) <> "" Then
      stConPA = stConPA & " and pa08='" & txt1(14) & "'"
      stConTM = stConTM & " and tm08='" & txt1(14) & "'"
      pub_QL05 = pub_QL05 & ";" & Label12 & txt1(14) 'Add By Sindy 2010/11/16
   End If
   
   '申請人
   If txt1(15) <> "" And txt1(16) <> "" Then
      stConPA = stConPA & " and ((pa26>='" & txt1(15) & "' and pa26<='" & txt1(16) & "') or (pa27>='" & txt1(15) & "' and pa27<='" & txt1(16) & "') or (pa28>='" & txt1(15) & "' and pa28<='" & txt1(16) & "') or (pa29>='" & txt1(15) & "' and pa29<='" & txt1(16) & "') or (pa30>='" & txt1(15) & "' and pa30<='" & txt1(16) & "'))"
      stConTM = stConTM & " and ((tm23>='" & txt1(15) & "' and tm23<='" & txt1(16) & "') or (tm78>='" & txt1(15) & "' and tm78<='" & txt1(16) & "') or (tm79>='" & txt1(15) & "' and tm79<='" & txt1(16) & "') or (tm80>='" & txt1(15) & "' and tm80<='" & txt1(16) & "') or (tm81>='" & txt1(15) & "' and tm81<='" & txt1(16) & "'))"
      stConSP = stConSP & " and ((sp08>='" & txt1(15) & "' and sp08<='" & txt1(16) & "') or (sp58>='" & txt1(15) & "' and sp58<='" & txt1(16) & "') or (sp59>='" & txt1(15) & "' and sp59<='" & txt1(16) & "') or (sp65>='" & txt1(15) & "' and sp65<='" & txt1(16) & "') or (sp66>='" & txt1(15) & "' and sp66<='" & txt1(16) & "'))"
      stConLC = stConLC & " and (lc11>='" & txt1(15) & "' and lc11<='" & txt1(16) & "')"
      pub_QL05 = pub_QL05 & ";" & Label13 & txt1(15) & "-" & txt1(16) 'Add By Sindy 2010/11/16
   End If
   
   'FC代理人
   If txt1(17) <> "" And txt1(18) <> "" Then
      stConPA = stConPA & " and pa75>='" & txt1(17) & "' and pa75<='" & txt1(18) & "'"
      stConTM = stConTM & " and tm44>='" & txt1(17) & "' and tm44<='" & txt1(18) & "'"
      stConSP = stConSP & " and sp26>='" & txt1(17) & "' and sp26<='" & txt1(18) & "'"
      stConLC = stConLC & " and lc22>='" & txt1(17) & "' and lc22<='" & txt1(18) & "'"
      pub_QL05 = pub_QL05 & ";" & Label14 & txt1(17) & "-" & txt1(18) 'Add By Sindy 2010/11/16
   End If
   
   Select Case txt1(19)
      Case "1" 'FC代理人性質
         pub_QL05 = pub_QL05 & ";" & Label15 & "1.FC代理人性質" 'Add By Sindy 2010/11/16
         Select Case txt1(20)
            Case "1" '洲
               pub_QL05 = pub_QL05 & ";" & Label17 & "1.洲" 'Add By Sindy 2010/11/16
               stSelect = " c1 c01" & _
                  ",nvl(c2_3,0) c02" & _
                  ",nvl(c3_3,0) c03,decode(sign(c2_3),1,round(nvl(c3_3,0)/c2_3*100)||'%') c04" & _
                  ",nvl(c4_3,0) c05,decode(sign(c2_3),1,round(nvl(c4_3,0)/c2_3*100)||'%') c06" & _
                  ",nvl(c2_2,0) c07" & _
                  ",nvl(c3_2,0) c08,decode(sign(c2_2),1,round(nvl(c3_2,0)/c2_2*100)||'%') c09" & _
                  ",nvl(c4_2,0) c10,decode(sign(c2_2),1,round(nvl(c4_2,0)/c2_2*100)||'%') c11" & _
                  ",nvl(c2_2,0)-nvl(c2_3,0) c12" & _
                  ",decode(sign(c2_3),1,round((nvl(c2_2,0)-c2_3)/c2_3*100)||'%') c13" & _
                  ",nvl(c2,0) c14" & _
                  ",nvl(c3,0) c15,decode(sign(c2),1,round(nvl(c3,0)/c2*100)||'%') c16" & _
                  ",nvl(c4,0) c17,decode(sign(c2),1,round(nvl(c4,0)/c2*100)||'%') c18" & _
                  ",nvl(c2,0)-nvl(c2_2,0) c19" & _
                  ",decode(sign(c2_2),1,round((nvl(c2,0)-c2_2)/c2_2*100)||'%') c20"
                  
               stSelectX = " decode(substr(nvl(fa10,cu10),1,1),'0','亞洲','1','美洲','歐非洲') x1" & _
                  ",decode(fa76,'A',1,0) x2,decode(fa76,'A',0,1) x3"
               stSelectV = " x1 c1,sum(decode(x0,1,1)) c2,sum(decode(x0,1,x2)) c3,sum(decode(x0,1,x3)) c4,sum(decode(x0,2,1)) c2_2,sum(decode(x0,3,1)) c2_3,sum(decode(x0,2,x2)) c3_2,sum(decode(x0,3,x2)) c3_3,sum(decode(x0,2,x3)) c4_2,sum(decode(x0,3,x3)) c4_3"
               stGroupByV = " x1"
               stOrderBy = " 1 asc"
               
            Case "2" '國家
               pub_QL05 = pub_QL05 & ";" & Label17 & "2.國家" 'Add By Sindy 2010/11/16
               stSelect = " na03 c01" & _
                  ",nvl(c2_3,0) c02" & _
                  ",nvl(c3_3,0) c03,decode(sign(c2_3),1,round(nvl(c3_3,0)/c2_3*100)||'%') c04" & _
                  ",nvl(c4_3,0) c05,decode(sign(c2_3),1,round(nvl(c4_3,0)/c2_3*100)||'%') c06" & _
                  ",nvl(c2_2,0) c07" & _
                  ",nvl(c3_2,0) c08,decode(sign(c2_2),1,round(nvl(c3_2,0)/c2_2*100)||'%') c09" & _
                  ",nvl(c4_2,0) c10,decode(sign(c2_2),1,round(nvl(c4_2,0)/c2_2*100)||'%') c11" & _
                  ",nvl(c2_2,0)-nvl(c2_3,0) c12" & _
                  ",decode(sign(c2_3),1,round((nvl(c2_2,0)-c2_3)/c2_3*100)||'%') c13" & _
                  ",nvl(c2,0) c14" & _
                  ",nvl(c3,0) c15,decode(sign(c2),1,round(nvl(c3,0)/c2*100)||'%') c16" & _
                  ",nvl(c4,0) c17,decode(sign(c2),1,round(nvl(c4,0)/c2*100)||'%') c18" & _
                  ",nvl(c2,0)-nvl(c2_2,0) c19" & _
                  ",decode(sign(c2_2),1,round((nvl(c2,0)-c2_2)/c2_2*100)||'%') c20"
                  
               stSelectX = " substr(nvl(fa10,cu10),1,3) x1,decode(fa76,'A',1,0) x2,decode(fa76,'A',0,1) x3"
               stSelectV = " x1 c1,sum(decode(x0,1,1)) c2,sum(decode(x0,1,x2)) c3,sum(decode(x0,1,x3)) c4,sum(decode(x0,2,1)) c2_2,sum(decode(x0,3,1)) c2_3,sum(decode(x0,2,x2)) c3_2,sum(decode(x0,3,x2)) c3_3,sum(decode(x0,2,x3)) c4_2,sum(decode(x0,3,x3)) c4_3"
               If Val(txt1(22)) > 0 Then
                  stWhere = " and abs(nvl(c2,0)-nvl(c2_2,0))>=" & Val(txt1(22))
                  pub_QL05 = pub_QL05 & ";" & Label22 & txt1(22) 'Add By Sindy 2010/11/16
               End If
               stGroupByV = " x1"
               stOrderBy = " sign(c14) desc,c19 asc,c12 asc,c01 asc"
               
            Case "3" 'FC代理人
               pub_QL05 = pub_QL05 & ";" & Label17 & "3.FC代理人" 'Add By Sindy 2010/11/16
               stSelect = " c0||' '||na03 c0,decode(fc01,null,null,'＊')||c3 c01,c1_1 c01_1" & _
                  ",nvl(c2_3,0) c02" & _
                  ",nvl(c2_2,0) c03" & _
                  ",nvl(c2_2,0)-nvl(c2_3,0) c04" & _
                  ",nvl(c2,0) c05" & _
                  ",nvl(c2,0)-nvl(c2_2,0) c06"
                  
               stSelectX = " nvl(fa01||fa02,cu01||cu02) x1" & _
                  ",decode(fa01,null,nvl(cu05,nvl(cu04,cu06)),nvl(fa05,nvl(fa04,fa06))) x2" & _
                  ",substr(nvl(fa10,cu10),1,3) x3,fa76"
               stSelectV = " x3 c0,x1 c1,sum(decode(x0,1,1)) c2,max(x2) c3,max(fa76) c1_1,sum(decode(x0,2,1)) c2_2,sum(decode(x0,3,1)) c2_3"
               If Val(txt1(22)) > 0 Then
                  stWhere = " and abs(nvl(c2,0)-nvl(c2_2,0))>=" & Val(txt1(22))
                  pub_QL05 = pub_QL05 & ";" & Label22 & txt1(22) 'Add By Sindy 2010/11/16
               End If
               stGroupByV = " x3,x1"
               stOrderBy = " c0 asc,sign(c05) desc,c06 asc,c04 asc,c01_1 asc"
               
            Case "4" '申請人
               pub_QL05 = pub_QL05 & ";" & Label17 & "4.申請人" 'Add By Sindy 2010/11/16
               stSelect = " c0||' '||na03 c0,c3 c01" & _
                  ",nvl(c2_3,0) c02" & _
                  ",nvl(c2_2,0) c03" & _
                  ",nvl(c2_2,0)-nvl(c2_3,0) c04" & _
                  ",nvl(c2,0) c05" & _
                  ",nvl(c2,0)-nvl(c2_2,0) c06"
               stSelectX = " cu01||cu02 x1,nvl(cu05,nvl(cu04,cu06)) x2,substr(cu10,1,3) x3"
               stSelectV = " x3 c0,x1 c1,sum(decode(x0,1,1)) c2,max(x2) c3,sum(decode(x0,2,1)) c2_2,sum(decode(x0,3,1)) c2_3"
               If Val(txt1(22)) > 0 Then
                  stWhere = " and abs(nvl(c2,0)-nvl(c2_2,0))>=" & Val(txt1(22))
                  pub_QL05 = pub_QL05 & ";" & Label22 & txt1(22) 'Add By Sindy 2010/11/16
               End If
               stGroupByV = " x3,x1"
               stOrderBy = " c0 asc,sign(c05) desc,c06 asc,c04 asc,c01 asc"
               
         End Select
         
      Case "2" '系統別
         pub_QL05 = pub_QL05 & ";" & Label15 & "2.系統別" 'Add By Sindy 2010/11/16
         Select Case txt1(20)
            Case "1" '洲
               pub_QL05 = pub_QL05 & ";" & Label17 & "1.洲" 'Add By Sindy 2010/11/16
               stSelect = " c1 c01"
               stSelectX = " decode(substr(nvl(fa10,cu10),1,1),'0','亞洲','1','美洲','歐非洲') x1,cp01 x4"
               stSelectV = " x1 c1,sum(decode(x4,'" & stSys(1) & "',decode(x0,1,1))) c2,sum(decode(x4,'" & stSys(1) & "',decode(x0,2,1))) c2_2,sum(decode(x4,'" & stSys(1) & "',decode(x0,3,1))) c2_3"
               stGroupByV = " x1"
               stOrderBy = " 1 asc"
               
            Case "2" '國家
               pub_QL05 = pub_QL05 & ";" & Label17 & "2.國家" 'Add By Sindy 2010/11/16
               stSelect = " na03 c01"
               stSelectX = " substr(nvl(fa10,cu10),1,3) x1,cp01 x4"
               stSelectV = " x1 c1,sum(decode(x4,'" & stSys(1) & "',decode(x0,1,1))) c2,sum(decode(x4,'" & stSys(1) & "',decode(x0,2,1))) c2_2,sum(decode(x4,'" & stSys(1) & "',decode(x0,3,1))) c2_3"
               stGroupByV = " x1"
               stOrderBy = " c06 asc,c01 asc"
               
            Case "3" 'FC代理人
               pub_QL05 = pub_QL05 & ";" & Label17 & "3.FC代理人" 'Add By Sindy 2010/11/16
               stSelect = " c0||' '||na03 c0,decode(fc01,null,null,'＊')||c3 c01,c1_1 c01_1"
               stSelectX = " substr(nvl(fa10,cu10),1,3) x3,nvl(fa01||fa02,cu01||cu02) x1,decode(fa01,null,nvl(cu05,nvl(cu04,cu06)),nvl(fa05,nvl(fa04,fa06))) x2,cp01 x4,fa76"
               stSelectV = " x3 c0,x1 c1,sum(decode(x4,'" & stSys(1) & "',decode(x0,1,1))) c2,max(x2) c3,max(fa76) c1_1,sum(decode(x4,'" & stSys(1) & "',decode(x0,2,1))) c2_2,sum(decode(x4,'" & stSys(1) & "',decode(x0,3,1))) c2_3"
               stGroupByV = " x3,x1"
               stOrderBy = " c0 asc,sign(c05) desc,c06 asc,c04 asc,c01 asc"
               
            Case "4" '申請人
               pub_QL05 = pub_QL05 & ";" & Label17 & "4.申請人" 'Add By Sindy 2010/11/16
               stSelect = " c0||' '||na03 c0,c3 c01"
               stSelectX = " substr(cu10,1,3) x3,cu01||cu02 x1,nvl(cu05,nvl(cu04,cu06)) x2,cp01 x4"
               stSelectV = " x3 c0,x1 c1,sum(decode(x4,'" & stSys(1) & "',decode(x0,1,1))) c2,max(x2) c3,sum(decode(x4,'" & stSys(1) & "',decode(x0,2,1))) c2_2,sum(decode(x4,'" & stSys(1) & "',decode(x0,3,1))) c2_3"
               stGroupByV = " x3,x1"
               stOrderBy = " c0 asc,sign(c05) desc,c06 asc,c04 asc,c01 asc"
         End Select
         stSelect = stSelect & _
            ",nvl(c2_3,0) c02" & _
            ",nvl(c2_2,0) c03" & _
            ",nvl(c2_2,0)-nvl(c2_3,0) c04" & _
            ",nvl(c2,0) c05" & _
            ",nvl(c2,0)-nvl(c2_2,0) c06"
            
         If stSys(2) <> "" Then
            stSelect = stSelect & _
            ",nvl(c4_3,0) c07" & _
            ",nvl(c4_2,0) c08" & _
            ",nvl(c4_2,0)-nvl(c4_3,0) c09" & _
            ",nvl(c4,0) c10" & _
            ",nvl(c4,0)-nvl(c4_2,0) c11"
         End If
         
         If stSys(3) <> "" Then
            stSelect = stSelect & _
            ",nvl(c5_3,0) c12" & _
            ",nvl(c5_2,0) c13" & _
            ",nvl(c5_2,0)-nvl(c5_3,0) c14" & _
            ",nvl(c5,0) c15" & _
            ",nvl(c5,0)-nvl(c5_2,0) c16"
         End If
         
         If stSys(2) <> "" Then stSelectV = stSelectV & ",sum(decode(x4,'" & stSys(2) & "',decode(x0,1,1))) c4,sum(decode(x4,'" & stSys(2) & "',decode(x0,2,1))) c4_2,sum(decode(x4,'" & stSys(2) & "',decode(x0,3,1))) c4_3"
         If stSys(3) <> "" Then stSelectV = stSelectV & ",sum(decode(x4,'" & stSys(3) & "',decode(x0,1,1))) c5,sum(decode(x4,'" & stSys(2) & "',decode(x0,2,1))) c5_2,sum(decode(x4,'" & stSys(2) & "',decode(x0,3,1))) c5_3"
         If Val(txt1(22)) > 0 Then
            stWhere = " and abs(nvl(c2,0)-nvl(c2_2,0))>=" & Val(txt1(22))
            pub_QL05 = pub_QL05 & ";" & Label22 & txt1(22) 'Add By Sindy 2010/11/16
         End If
   End Select
   
   'PA
   If ChkTbl(txt1(3), 1) = True Then
      If arTable(1) <> "" Then
         arTable(1) = arTable(1) & " union all "
         arTable(2) = arTable(2) & " union all "
         arTable(3) = arTable(3) & " union all "
      End If
      
      'Modify by Morgan 2010/10/8 關係企業合併--David
      'Modified by Morgan 2016/11/29 申請人更名要併到最新的 --David
      stTable = " from caseprogress, Patent, fagent, customer" & _
         " where cp09<'B' and cp57 is null " & stCon & stConCP12 & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
         " and fa01(+)=substr(pa75,1,6)||'00' and fa02(+)='0'" & stConFa & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)='0'" & stConCu
            
      For intI = 1 To 3
         arTable(intI) = arTable(intI) & " select " & intI & " x0," & stSelectX & stTable & arConDate(intI)
         If txt1(19) = "2" Then
            arTable(intI) = arTable(intI) & " union all select " & intI & " x0," & stSelectX & " From caseprogress" & _
               ",(select distinct pa75 F01" & stTable & arConDate(intI) & ") FF" & _
               ",Patent,fagent,customer where cp09<'B' and cp57 is null and cp01='CFP' " & stCon & _
               " and F01(+)=cp44 and F01 is not null" & _
               " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & _
               " and fa01(+)=substr(F01,1,6)||'00' and fa02(+)='0'" & _
               " and cu01(+)=substr(pa26,1,8) and cu02(+)='0'" & arConDate(intI)
         End If
      Next
   End If
   'TM
   If ChkTbl(txt1(3), 2) = True Then
      If arTable(1) <> "" Then
         arTable(1) = arTable(1) & " union all "
         arTable(2) = arTable(2) & " union all "
         arTable(3) = arTable(3) & " union all "
      End If
      
      stTable = " From caseprogress, Trademark, fagent, customer" & _
         " where cp09<'B' and cp57 is null" & stCon & stConCP12 & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & stConTM & _
         " and fa01(+)=substr(tm44,1,8) and fa02(+)='0'" & stConFa & _
         " and cu01(+)=substr(tm23,1,8) and cu02(+)='0'" & stConCu
      For intI = 1 To 3
         arTable(intI) = arTable(intI) & " select " & intI & " x0," & stSelectX & stTable & arConDate(intI)
      Next
   End If
   'SP
   If ChkTbl(txt1(3), 4) = True Then
      If arTable(1) <> "" Then
         arTable(1) = arTable(1) & " union all "
         arTable(2) = arTable(2) & " union all "
         arTable(3) = arTable(3) & " union all "
      End If
      
      stTable = " From caseprogress, servicepractice, fagent, customer" & _
         " where cp09<'B' and cp57 is null" & stCon & stConCP12 & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
         " and fa01(+)=substr(sp26,1,8) and fa02(+)='0'" & stConFa & _
         " and cu01(+)=substr(sp08,1,8) and cu02(+)='0'" & stConCu
      
      For intI = 1 To 3
         arTable(intI) = arTable(intI) & " select " & intI & " x0," & stSelectX & stTable & arConDate(intI)
      Next
   End If
   'LC
   If ChkTbl(txt1(3), 3) = True Then
      If arTable(1) <> "" Then
         arTable(1) = arTable(1) & " union all "
         arTable(2) = arTable(2) & " union all "
         arTable(3) = arTable(3) & " union all "
      End If
      
      stTable = " From caseprogress, lawcase, fagent, customer" & _
         " where cp09<'B' and cp57 is null" & stCon & stConCP12 & _
         " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04 and lc01 is not null" & stConLC & _
         " and fa01(+)=substr(lc22,1,8) and fa02(+)='0'" & stConFa & _
         " and cu01(+)=substr(lc11,1,8) and cu02(+)='0'" & stConCu
      
      For intI = 1 To 3
         arTable(intI) = arTable(intI) & " select " & intI & " x0," & stSelectX & stTable & arConDate(intI)
      Next
   End If

   Select Case txt1(20)
      Case "1"
         strExc(0) = "select " & stSelect & _
            " from (select " & stSelectV & " From (" & arTable(1) & " union all " & arTable(2) & " union all " & arTable(3) & ") x group by " & stGroupByV & ") v1" & _
            " where 1=1" & stWhere & _
            " order by " & stOrderBy
      Case "2"
         strExc(0) = "select " & stSelect & _
            " from (select " & stSelectV & " From (" & arTable(1) & " union all " & arTable(2) & " union all " & arTable(3) & ") x group by " & stGroupByV & ") v1" & _
            ",nation where na01(+)=c1" & stWhere & _
            " order by " & stOrderBy
      Case "3"
         'Modified by Morgan 2013/11/28 互惠代理人關係企業合併--David
         strExc(0) = "select " & stSelect & _
            " from (select " & stSelectV & " From (" & arTable(1) & " union all " & arTable(2) & " union all " & arTable(3) & ") x group by " & stGroupByV & ") v1" & _
            ",nation,(select substr(fc01,1,6) fc01 from FAGENTCONFIG where fc02='0' and fc04=" & fc04 & _
            " and fc05=" & fc05 & stConFC06 & " group by substr(fc01,1,6)) TFC" & _
            " where na01(+)=c0" & stWhere & _
            " and fc01(+)=substr(c1,1,6) order by " & stOrderBy
         'end 2013/11/28
      Case "4"
         strExc(0) = "select " & stSelect & _
            " from (select " & stSelectV & " From (" & arTable(1) & " union all " & arTable(2) & " union all " & arTable(3) & ") x group by " & stGroupByV & ") v1" & _
            ",nation where na01(+)=c0" & stWhere & _
            " order by " & stOrderBy
   End Select
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If fnSaveParentForm(Me) Then
         InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/11/16
         With frm100129_2
         If txt1(0) = "1" Then
            .Caption = .Caption & "(收文)"
         Else
            .Caption = .Caption & "(發文)"
         End If
         If txt1(19) = "2" Then
            .Caption = txt1(3) & .Caption
         End If
         If txt1(20) = "1" Then
            .Caption = .Caption & "-洲別"
         ElseIf txt1(20) = "2" Then
            .Caption = .Caption & "-國家別"
         ElseIf txt1(20) = "3" Then
            .Caption = .Caption & "-FC代理人別"
         ElseIf txt1(20) = "4" Then
            .Caption = .Caption & "-申請人別"
         End If
         
         .m_RptType = txt1(19) & txt1(20)
         .SetGrid RsTemp, GetFormatString
         
         If txt1(20) <> "3" Then .lblMemo = ""
         If IsUserHasRightOfFunction(Me.Name, strPrint, False) Then
            .cmdOK(2).Enabled = True
         Else
            .cmdOK(2).Enabled = False
         End If
         
         End With
         Me.Hide
      End If
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/11/16
      ShowNoData
   End If
End Sub

Private Sub txt1_Change(Index As Integer)
   Select Case Index
      Case 31
           lblName = PUB_GetFCPGrpName(txt1(31))
   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 Then
      Select Case Index
         Case 0
            If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
               KeyAscii = 0
               Beep
            End If
         
         Case 19
            If KeyAscii < Asc("1") Or KeyAscii > Asc("2") Then
               KeyAscii = 0
               Beep
            End If
         
         Case 21
            If KeyAscii < Asc("A") Or KeyAscii > Asc("C") Then
               KeyAscii = 0
               Beep
            End If
            
         Case 20
            If KeyAscii < Asc("1") Or KeyAscii > Asc("4") Then
               KeyAscii = 0
               Beep
            End If
            
         Case 22
            If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
               KeyAscii = 0
               Beep
            End If
            
      End Select
   End If
End Sub

Private Function GetFormatString() As String
   Dim stFromatString As String, stTmp As String
   
   If txt1(20) = "1" Then
      stFromatString = "洲　別"
   ElseIf txt1(20) = "2" Then
      stFromatString = "國家別"
   ElseIf txt1(20) = "3" Then
      stFromatString = "國家別|FC代理人|性質"
   ElseIf txt1(20) = "4" Then
      stFromatString = "國家別|申請人"
   End If
      
   
   If txt1(19) & txt1(20) = "11" Or txt1(19) & txt1(20) = "12" Then
      stTmp = txt1(1) \ 100 - 2 & "." & Val(Right(txt1(1), 2))
      If Val(txt1(2)) <> Val(txt1(1)) Then
         stTmp = stTmp & " - " & Val(Right(txt1(2), 2))
      End If
      stFromatString = stFromatString & "|" & stTmp & "|事務所Ａ|比率|廠商Ｂ|比率"
      
      stTmp = txt1(1) \ 100 - 1 & "." & Val(Right(txt1(1), 2))
      If Val(txt1(2)) <> Val(txt1(1)) Then
         stTmp = stTmp & " - " & Val(Right(txt1(2), 2))
      End If
      stFromatString = stFromatString & "|" & stTmp & "|事務所Ａ|比率|廠商Ｂ|比率|增減|比率"
      
      stTmp = txt1(1) \ 100 & "." & Val(Right(txt1(1), 2))
      If Val(txt1(2)) <> Val(txt1(1)) Then
         stTmp = stTmp & " - " & Val(Right(txt1(2), 2))
      End If
      stFromatString = stFromatString & "|" & stTmp & "|事務所Ａ|比率|廠商Ｂ|比率|增減|比率"
   
   Else
      If txt1(19) = "1" Then
         stTmp = txt1(1) \ 100 - 2 & "." & Val(Right(txt1(1), 2))
         If Val(txt1(2)) <> Val(txt1(1)) Then
            stTmp = stTmp & " - " & Val(Right(txt1(2), 2))
         End If
         stFromatString = stFromatString & "|" & stTmp
         
         stTmp = txt1(1) \ 100 - 1 & "." & Val(Right(txt1(1), 2))
         If Val(txt1(2)) <> Val(txt1(1)) Then
            stTmp = stTmp & " - " & Val(Right(txt1(2), 2))
         End If
         stFromatString = stFromatString & "|" & stTmp & "|增減"
         
         stTmp = txt1(1) \ 100 & "." & Val(Right(txt1(1), 2))
         If Val(txt1(2)) <> Val(txt1(1)) Then
            stTmp = stTmp & " - " & Val(Right(txt1(2), 2))
         End If
         stFromatString = stFromatString & "|" & stTmp & "|增減"
         
      ElseIf txt1(19) = "2" Then
         For intI = 1 To 3
            If stSys(intI) = "" Then Exit For
            stTmp = txt1(1) \ 100 - 2 & "." & Val(Right(txt1(1), 2))
            If Val(txt1(2)) <> Val(txt1(1)) Then
               stTmp = stTmp & " - " & Val(Right(txt1(2), 2))
            End If
            stFromatString = stFromatString & "|" & stSys(intI) & stTmp
            
            stTmp = txt1(1) \ 100 - 1 & "." & Val(Right(txt1(1), 2))
            If Val(txt1(2)) <> Val(txt1(1)) Then
               stTmp = stTmp & " - " & Val(Right(txt1(2), 2))
            End If
            stFromatString = stFromatString & "|" & stSys(intI) & stTmp & "|增減"
            
            stTmp = txt1(1) \ 100 & "." & Val(Right(txt1(1), 2))
            If Val(txt1(2)) <> Val(txt1(1)) Then
               stTmp = stTmp & " - " & Val(Right(txt1(2), 2))
            End If
            stFromatString = stFromatString & "|" & stSys(intI) & stTmp & "|增減"
            
         Next
      End If
   End If
   
   GetFormatString = stFromatString
End Function
