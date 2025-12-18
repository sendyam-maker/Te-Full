VERSION 5.00
Begin VB.Form frm060404 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文與承辦期限比較統計表"
   ClientHeight    =   4455
   ClientLeft      =   1290
   ClientTop       =   2955
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5940
   Begin VB.OptionButton Option2 
      Caption         =   "已收文未發文已逾承辦期限"
      Height          =   180
      Left            =   270
      TabIndex        =   20
      Top             =   1260
      Width           =   3255
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   1
      Left            =   4095
      MaxLength       =   3
      TabIndex        =   13
      Text            =   "30"
      Top             =   1965
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   5040
      ScaleHeight     =   435
      ScaleWidth      =   630
      TabIndex        =   12
      Top             =   2580
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   14
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2310
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   720
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1620
      Width           =   240
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1050
      TabIndex        =   0
      Top             =   504
      Width           =   3270
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1455
      MaxLength       =   7
      TabIndex        =   1
      Top             =   945
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2550
      MaxLength       =   7
      TabIndex        =   2
      Top             =   945
      Width           =   990
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1950
      Width           =   240
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4050
      TabIndex        =   6
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   4845
      TabIndex        =   7
      Top             =   60
      Width           =   800
   End
   Begin VB.OptionButton Option1 
      Caption         =   "發文日期："
      Height          =   180
      Left            =   270
      TabIndex        =   19
      Top             =   990
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.Line Line1 
      X1              =   2385
      X2              =   2550
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Shape Shape1 
      Height          =   675
      Left            =   135
      Top             =   870
      Width           =   3660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "4. 案件更改承辦人者，遲延記錄將會統計於新的承辦人。"
      Height          =   180
      Index           =   6
      Left            =   180
      TabIndex        =   18
      Top             =   3960
      Width           =   4500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "3. 本統計表已排除收文後尚不能發文的主動修正。"
      Height          =   180
      Index           =   5
      Left            =   180
      TabIndex        =   17
      Top             =   3750
      Width           =   3960
   End
   Begin VB.Label Label1 
      Caption         =   "2. 統計方式為承辦期限或核稿期限減去發文日所得的工作天數作為數值，負值即代表遲延天數。另後附括號以表示遲延的總件數。"
      Height          =   390
      Index           =   4
      Left            =   180
      TabIndex        =   16
      Top             =   3330
      Width           =   5670
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1. 本統計表僅統計發文日超過承辦期限或核稿期限的案件。"
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   15
      Top             =   3120
      Width           =   4680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   14
      Top             =   2850
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "輸出方式：  　   ( 1.螢幕 2.印表機 )"
      Height          =   180
      Index           =   9
      Left            =   150
      TabIndex        =   11
      Top             =   2340
      Width           =   2730
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "組別：          ( 1.電子電機 2.化學 3.日文 4.機械設計 5.其他 )"
      Height          =   180
      Index           =   7
      Left            =   150
      TabIndex        =   10
      Top             =   1665
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   9
      Top             =   555
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "報表別：             ( 1.工程師統計 2.各組統計 3.超過           個工作天明細  )"
      Height          =   180
      Index           =   8
      Left            =   150
      TabIndex        =   8
      Top             =   2010
      Width           =   5610
   End
End
Attribute VB_Name = "frm060404"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Create by Morgan 2012/4/2
Option Explicit

Dim m_bPrint As Boolean '列印權限
Dim m_bQuery As Boolean '跨組查詢權限

Dim PLeft(0 To 20) As Integer, iPrint As Long, m_iMaxNum As Integer, m_iColMax As Integer
Dim m_bPrinter As Boolean, m_iPages As Integer, m_Device
Dim adoReport As ADODB.Recordset
Dim m_stGrp As String, arrNameList() As String, m_iRound As Integer, m_stTotal As String, m_stCount As String
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If Option1.Value Then
            If Len(Txt1(2)) = 0 Then
               MsgBox "日期區間不可空白!!", , "USER 輸入錯誤"
               Txt1(2).SetFocus
               txt1_GotFocus (2)
               Exit Sub
            ElseIf Len(Txt1(3)) = 0 Then
               MsgBox "日期區間不可空白!!", , "USER 輸入錯誤"
               Txt1(3).SetFocus
               txt1_GotFocus (3)
               Exit Sub
            ElseIf PUB_CheckKeyInDate(Me.Txt1(2)) = -1 Then
               Me.Txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            ElseIf PUB_CheckKeyInDate(Me.Txt1(3)) = -1 Then
               Me.Txt1(3).SetFocus
               txt1_GotFocus 3
               Exit Sub
            ElseIf Val(Me.Txt1(2).Text) > Val(Me.Txt1(3).Text) Then
               MsgBox "日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.Txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
         End If
            
         If Len(Txt1(0)) = 0 Then
            MsgBox "系統類別不可空白!!", , "USER 輸入錯誤"
            Txt1(0).SetFocus
            txt1_GotFocus (0)
         ElseIf Me.Txt1(12).Text = "" Then
            MsgBox "請輸入報表別!!!", vbExclamation + vbOKOnly
            Me.Txt1(12).SetFocus
            txt1_GotFocus 12
            
         ElseIf Me.Txt1(14).Text = "" Then
            MsgBox "請選擇輸出方式!!!", vbExclamation + vbOKOnly
            Me.Txt1(14).SetFocus
            txt1_GotFocus 14
            
         ElseIf Me.Txt1(12).Text = "3" And Txt1(1) = "" Then
            MsgBox "選擇明細報表時請輸入天數!!!", vbExclamation + vbOKOnly
            Me.Txt1(1).SetFocus
            txt1_GotFocus 1
            
         Else
            Screen.MousePointer = vbHourglass
            Me.Enabled = False
            Process
            Me.Enabled = True
            Screen.MousePointer = vbDefault
         End If
      Case 1
         Unload Me
      Case Else
   End Select
End Sub

Private Sub Form_Load()
   m_bPrint = IsUserHasRightOfFunction(Me.Name, strPrint, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   '跨組查詢
   If Not m_bQuery Then
      Txt1(11) = PUB_GetStaffST16(strUserNum)
      If Txt1(11) = "" Then Txt1(11) = "5"
      Txt1(11).Enabled = False
   End If
   '列印
   If Not m_bPrint Then
      Txt1(14) = "1"
      Txt1(14).Enabled = False
   End If
   MoveFormToCenter Me
   Txt1(0) = "FCP,FG,P,CFP"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060404 = Nothing
End Sub

Private Sub Option1_Click()
   Txt1(2).Enabled = True
   Txt1(2).SetFocus
   Txt1(3).Enabled = True
   Label1(3) = Replace(Label1(3), "未發文且今日", "發文日")
   Label1(4) = Replace(Label1(4), "今日", "發文日")
End Sub

Private Sub Option2_Click()
   Txt1(2).Enabled = False
   Txt1(3).Enabled = False
   Label1(3) = Replace(Label1(3), "發文日", "未發文且今日")
   Label1(4) = Replace(Label1(4), "發文日", "今日")
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse Txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 Then
      Select Case Index
         Case 1
            If Not IsNumeric(Chr(KeyAscii)) Then
               KeyAscii = 0
               Beep
            End If
         Case 2, 3 '發文日
            If Not IsNumeric(Chr(KeyAscii)) Then
               KeyAscii = 0
               Beep
            End If
         Case 11 '組別
            If Chr(KeyAscii) < "1" Or Chr(KeyAscii) > "5" Then
               KeyAscii = 0
               Beep
            End If
         Case 12 '報表別
            If Chr(KeyAscii) < "1" Or Chr(KeyAscii) > "3" Then
               KeyAscii = 0
               Beep
            End If
         Case 14 '輸出方式
            If Chr(KeyAscii) < "1" Or Chr(KeyAscii) > "2" Then
               KeyAscii = 0
               Beep
            End If
      End Select
   End If
End Sub

'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷，列出排除案件
Private Function ProcExceptList(ByVal pSQL As String) As String
Dim intJ As Integer, strGrp As String, strTmp1 As String
Dim rsR1 As New ADODB.Recordset

    ProcExceptList = ""
    If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
        intJ = 1
        Set rsR1 = ClsLawReadRstMsg(intJ, pSQL)
        If intJ = 1 Then
            With rsR1
                 .MoveFirst
                 Do While Not .EOF
                     If strGrp <> "" & .Fields("CASENO") Then
                        If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & .Fields("CASENO"), "" & .Fields("cust01") & "," & .Fields("cust02") & "," & .Fields("cust03") & "," & .Fields("cust04") & "," & .Fields("cust05"), "" & .Fields("fcno")) = False Then
                            intCufaCnt = intCufaCnt + 1
                            strTmp1 = strTmp1 & "," & .Fields("CASENO")
                        End If
                     End If
                     strGrp = "" & .Fields("CASENO")
                     .MoveNext
                 Loop
            End With
        End If
        Set rsR1 = Nothing
        
        If strTmp1 <> "" Then
            ProcExceptList = " AND CASENO NOT IN (" & GetAddStr(strTmp1) & ") "
        End If
    End If
End Function

Private Sub Process()
   Dim stCon As String, stSys As String
   Dim stPty As String
   Dim ii As Integer, iFrom As Integer, iTo As Integer
   Dim stVTB As String, stConEP As String, stConCP As String, stConPA As String, stConSP As String
   Dim strExcept As String 'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷，列出排除案件
   
   ClearQueryLog (Me.Name)
   
   stSys = "'" & Join(Split(Txt1(0), ","), "','") & "'"
   stCon = " and cp01 in (" & stSys & ")"
   'Added by Lydia 2019/11/01 利益衝突案件
   m_AllSys = IIf(Txt1(0) <> "ALL", Txt1(0), GetAllSysKind(, Txt1(0)))
   intCufaCnt = 0
   'end 2019/11/01
   
   pub_QL05 = pub_QL05 & ";" & Label1(0) & Txt1(0)
   
   'Modified by Morgan 2012/5/4 +已收文未發文已逾承辦期限報表
   If Option1.Value Then
      If Txt1(2) <> "" Then
         'Modified by Lydia 2016/12/21 排除D類收文
         'stCon = stCon & " and cp27>=" & DBDATE(Txt1(2))
         stCon = stCon & " and cp158>=" & DBDATE(Txt1(2)) & " and substr(cp09,1,1) <> 'D' "
      End If
      If Txt1(3) <> "" Then
         'Modified by Lydia 2016/12/21 排除D類收文
         'stCon = stCon & " and cp27<=" & DBDATE(txt1(3))
         stCon = stCon & " and cp158<=" & DBDATE(Txt1(3)) & " and substr(cp09,1,1) <> 'D' "
      End If
      pub_QL05 = pub_QL05 & Option1 & Txt1(2) & "-" & Txt1(3)
   Else
      'Modified by Lydia 2016/12/21 排除D類收文
      'stCon = stCon & " and cp27||cp57 is null and cp05>" & (strSrvDate(1) - 30000)
      stCon = stCon & " and cp158=0 and cp159=0 and cp05>" & (strSrvDate(1) - 30000) & " and substr(cp09,1,1) <> 'D' "
      stConCP = " and cp48+0<" & strSrvDate(1)
      stConEP = " and ep08<" & strSrvDate(1)
      stConPA = " and pa57||pa108 is null"
      stConSP = " and sp15||sp61 is null"
      pub_QL05 = pub_QL05 & Option2
   End If
   
   If Txt1(11) <> "" Then
      stCon = stCon & " and st16='" & Txt1(11) & "'"
      pub_QL05 = pub_QL05 & ";" & Left(Label1(7), 3) & Txt1(11) & Trim(Mid(Label1(7), 4))
   End If
   
   'Modified by Morgan 2012/4/18 只要負值的資料
   
   '排除主動修正(或補充說明)其承辦期限=本所期限且>收文日(或中說發文日)+10個工作天者
   '新案翻譯天數差=承辦期限-完稿日
   '核稿天數差=核稿期限-發文日
   '工程師
   If Txt1(12) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1(8), 4) & "1.工程師統計"
      'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
      'Modified by Lydia 2019/11/01 增加欄位:本所案號CaseNo,申請人1~5(cust01~cust05),FC代理人
      stVTB = " select st16,cp14,st02,decode(nvl(pa09,'000'),'000',cpm03,cpm04) C1,workdaydiff(nvl(cp27," & strSrvDate(1) & "),cp48) C2" & _
         ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO" & _
         " From caseprogress a, staff,patent,casepropertymap" & _
         " where cp48+0>0 " & stCon & stConCP & " and cp10<>'201' and cp14 is not null and st01(+)=cp14 and st03 in ('F21','F81') AND CP14<>'F4102' " & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
         " and (not(cp10 in ('203','206') and cp48=nvl(cp06,0)) or cp48<=workdayadd(10,cp05)" & _
         " or exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04" & _
         " and b.cp10 in ('201','209','210') and b.cp27>=a.cp05 and b.cp27<=nvl(a.cp27," & strSrvDate(1) & ") and a.cp48<=workdayadd(10,b.cp27)))" & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
         " union all select st16,cp14,st02,cpm03 C1,workdaydiff(nvl(ep09," & strSrvDate(1) & "),cp48) C2" & _
         ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO" & _
         " From caseprogress a,patent,engineerprogress, staff,casepropertymap" & _
         " where cp10='201' and cp14 is not null and cp48+0>0 and ep02(+)=cp09 " & stCon & stConCP & " and st01(+)=cp14 and st03 in ('F21','F81') AND CP14<>'F4102'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
         " union all select st16,ep04,st02,'核稿' C1,workdaydiff(nvl(cp27," & strSrvDate(1) & "),ep08) C2" & _
         ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO" & _
         " From caseprogress a,patent,engineerprogress, staff" & _
         " where cp10='201' and ep02(+)=cp09 and ep08>0 and ep04 is not null" & stCon & stConEP & " and st01(+)=ep04 and st03 in ('F21','F81') AND EP04<>'F4102' " & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA
         
      'Modified by Lydia 2019/11/01 增加欄位:本所案號CaseNo,申請人1~5(cust01~cust05),FC代理人
      stVTB = stVTB & " union all" & _
         " select st16,cp14,st02,decode(nvl(sp09,'000'),'000',cpm03,cpm04) C1,workdaydiff(nvl(cp27," & strSrvDate(1) & "),cp48) C2" & _
         " ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,SP08 AS CUST01,SP58 AS CUST02,SP59 AS CUST03,SP65 AS CUST04,SP66 AS CUST05,SP26 AS FCNO" & _
         " From caseprogress a, staff,servicepractice,casepropertymap" & _
         " where cp48>0 " & stCon & stConCP & " and cp10<>'201' and cp14 is not null and st01(+)=cp14 and st03 in ('F21','F81') " & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
         " union all select st16,cp14,st02,cpm03 C1,workdaydiff(nvl(ep09," & strSrvDate(1) & "),cp48) C2" & _
         " ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,SP08 AS CUST01,SP58 AS CUST02,SP59 AS CUST03,SP65 AS CUST04,SP66 AS CUST05,SP26 AS FCNO" & _
         " From caseprogress a,servicepractice,engineerprogress, staff,casepropertymap" & _
         " where cp10='201' and cp14 is not null and cp48+0>0 and ep02(+)=cp09 " & stCon & stConCP & " and st01(+)=cp14 and st03 in ('F21','F81') " & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
         " union all select st16,ep04,st02,'核稿' C1,workdaydiff(nvl(cp27," & strSrvDate(1) & "),ep08) C2" & _
         " ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,SP08 AS CUST01,SP58 AS CUST02,SP59 AS CUST03,SP65 AS CUST04,SP66 AS CUST05,SP26 AS FCNO" & _
         " From caseprogress a,servicepractice,engineerprogress, staff" & _
         " where cp10='201' and ep02(+)=cp09 and ep08>0 and ep04 is not null " & stCon & stConEP & " and st01(+)=ep04 and st03 in ('F21','F81') " & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP
      
      strExcept = ProcExceptList("select X.* from (" & stVTB & ") X where c2<0 order by CASENO ") 'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷，列出排除案件
      
      'Modifeid by Lydia 2019/11/01 + 排除案件strExcept
      strExc(0) = "select st16 R01,cp14 R02,max(st02) R03,C1 R04,sum(C2) R05,sum(1) R06" & _
         " from (" & stVTB & ") where C2<0 " & strExcept & " group by st16,C1,cp14"
   '各組
   ElseIf Txt1(12) = "2" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1(8), 4) & "2.各組統計"
      'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
      'Modified by Lydia 2019/11/01 增加欄位:本所案號CaseNo,申請人1~5(cust01~cust05),FC代理人
      stVTB = "select st16,decode(nvl(pa09,'000'),'000',cpm03,cpm04) C1,workdaydiff(nvl(cp27," & strSrvDate(1) & "),cp48) C2" & _
         ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO" & _
         " From caseprogress a, staff,patent,casepropertymap" & _
         " where cp48+0>0 " & stCon & stConCP & " and cp10<>'201' and cp14 is not null and st01(+)=cp14 and st03 in ('F21','F81')" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
         " and (not(cp10 in ('203','206')  and cp48=nvl(cp06,0)) or cp48<=workdayadd(10,cp05)" & _
         " or exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04" & _
         " and b.cp10 in ('201','209','210') and b.cp27>=a.cp05 and b.cp27<=nvl(a.cp27," & strSrvDate(1) & ") and a.cp48<=workdayadd(10,b.cp27)))" & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
         " union all select st16,cpm03 C1,workdaydiff(nvl(ep09," & strSrvDate(1) & "),cp48) C2" & _
         ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO" & _
         " From caseprogress a,patent,engineerprogress, staff,casepropertymap" & _
         " where cp10='201' and cp14 is not null and cp48+0>0 and ep02(+)=cp09 " & stCon & stConCP & " and st01(+)=cp14 and st03 in ('F21','F81') AND CP14<>'F4102' " & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
         " union all select st16,'核稿' C1,workdaydiff(nvl(cp27," & strSrvDate(1) & "),ep08) C2" & _
         ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO" & _
         " From caseprogress a,patent,engineerprogress, staff" & _
         " where cp10='201' and ep02(+)=cp09 and ep08>0 and ep04 is not null " & stCon & stConEP & " and st01(+)=ep04 and st03 in ('F21','F81') AND EP04<>'F4102' " & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA
      
      'Modified by Lydia 2019/11/01 增加欄位:本所案號CaseNo,申請人1~5(cust01~cust05),FC代理人
      stVTB = stVTB & " union all " & _
         "select st16,decode(nvl(sp09,'000'),'000',cpm03,cpm04) C1,workdaydiff(nvl(cp27," & strSrvDate(1) & "),cp48) C2" & _
         " ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,SP08 AS CUST01,SP58 AS CUST02,SP59 AS CUST03,SP65 AS CUST04,SP66 AS CUST05,SP26 AS FCNO" & _
         " From caseprogress a, staff,servicepractice,casepropertymap" & _
         " where cp48>0 " & stCon & stConCP & " and cp10<>'201' and cp14 is not null and st01(+)=cp14 and st03 in ('F21','F81')" & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
         " union all select st16,cpm03 C1,workdaydiff(nvl(ep09," & strSrvDate(1) & "),cp48) C2" & _
         " ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,SP08 AS CUST01,SP58 AS CUST02,SP59 AS CUST03,SP65 AS CUST04,SP66 AS CUST05,SP26 AS FCNO" & _
         " From caseprogress a,servicepractice,engineerprogress, staff,casepropertymap" & _
         " where cp10='201' and cp14 is not null and cp48+0>0 and ep02(+)=cp09 " & stCon & stConCP & " and st01(+)=cp14 and st03 in ('F21','F81')" & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
         " union all select st16,'核稿' C1,workdaydiff(nvl(cp27," & strSrvDate(1) & "),ep08) C2" & _
         " ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,SP08 AS CUST01,SP58 AS CUST02,SP59 AS CUST03,SP65 AS CUST04,SP66 AS CUST05,SP26 AS FCNO" & _
         " From caseprogress a,servicepractice,engineerprogress, staff" & _
         " where cp10='201' and ep02(+)=cp09 and ep08>0 and ep04 is not null " & stCon & stConEP & " and st01(+)=ep04 and st03 in ('F21','F81')" & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP
      
      strExcept = ProcExceptList("select X.* from (" & stVTB & ") X where c2<0 order by CASENO ")  'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷，列出排除案件
      'Modifeid by Lydia 2019/11/01 + 排除案件strExcept
      strExc(0) = "select 1 R01,st16 R02,cst16(st16) R03,C1 R04,sum(C2) R05,sum(1) R06" & _
         " from (" & stVTB & ") where C2<0 " & strExcept & " group by C1,st16"
   '明細
   Else
      pub_QL05 = pub_QL05 & ";" & Left(Label1(8), 4) & "3.超過 " & Txt1(1) & " 個工作天明細"
      'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
      'Modified by Lydia 2019/11/01 增加欄位:本所案號CaseNo,申請人1~5(cust01~cust05),FC代理人
      stVTB = "select st16 R01,cp14 R02,st02 R03,decode(nvl(pa09,'000'),'000',cpm03,cpm04) R04,workdaydiff(nvl(cp27," & strSrvDate(1) & "),cp48) R05" & _
         ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) R06, sqldatet(cp27) R07, sqldatet(cp48) R08" & _
         ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO" & _
         " From caseprogress a, staff,patent,casepropertymap" & _
         " where cp48+0>0 " & stCon & stConCP & " and cp10<>'201' and cp14 is not null and st01(+)=cp14 and st03 in ('F21','F81') AND CP14<>'F4102'" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
         " and (not(cp10 in ('203','206')  and cp48=nvl(cp06,0)) or cp48<=workdayadd(10,cp05)" & _
         " or exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04" & _
         " and b.cp10 in ('201','209','210') and b.cp27>=a.cp05 and b.cp27<=nvl(a.cp27," & strSrvDate(1) & ") and a.cp48<=workdayadd(10,b.cp27)))" & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
         " union all select st16 R01,cp14 R02,st02 R03,cpm03 R04,workdaydiff(nvl(ep09," & strSrvDate(1) & "),cp48) R05,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) R06, sqldatet(ep09) R07, sqldatet(cp48) R08" & _
         ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO" & _
         " From caseprogress a,patent,engineerprogress, staff,casepropertymap" & _
         " where cp10='201' and cp14 is not null and cp48+0>0 and ep02(+)=cp09 " & stCon & stConCP & " and st01(+)=cp14 and st03 in ('F21','F81') AND CP14<>'F4102' " & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA & _
         " union all select st16 R01,ep04 R02,st02 R03,'核稿' R04,workdaydiff(nvl(cp27," & strSrvDate(1) & "),ep08) R05" & _
         ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) R06, sqldatet(cp27) R07, sqldatet(ep08) R08" & _
         ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,PA26 AS CUST01,PA27 AS CUST02,PA28 AS CUST03,PA29 AS CUST04,PA30 AS CUST05,PA75 AS FCNO" & _
         " From caseprogress a,patent,engineerprogress, staff" & _
         " where cp10='201' and ep02(+)=cp09 and ep08>0 and ep04 is not null " & stCon & stConEP & " and st01(+)=ep04 and st03 in ('F21','F81') AND EP04<>'F4102' " & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & stConPA
      
      'Modified by Lydia 2019/11/01 增加欄位:本所案號CaseNo,申請人1~5(cust01~cust05),FC代理人
      stVTB = stVTB & " union all select st16 R01,cp14 R02,st02 R03,decode(nvl(sp09,'000'),'000',cpm03,cpm04) R04,workdaydiff(nvl(cp27," & strSrvDate(1) & "),cp48) R05" & _
         ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) R06, sqldatet(cp27) R07, sqldatet(cp48) R08" & _
         " ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,SP08 AS CUST01,SP58 AS CUST02,SP59 AS CUST03,SP65 AS CUST04,SP66 AS CUST05,SP26 AS FCNO" & _
         " From caseprogress a, staff,servicepractice,casepropertymap" & _
         " where cp48+0>0 " & stCon & stConCP & " and cp10<>'201' and cp14 is not null and st01(+)=cp14 and st03 in ('F21','F81')" & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
         " union all select st16 R01,cp14 R02,st02 R03,cpm03 R04,workdaydiff(nvl(ep09," & strSrvDate(1) & "),cp48) R05,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) R06, sqldatet(ep09) R07, sqldatet(cp48) R08" & _
         " ,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,SP08 AS CUST01,SP58 AS CUST02,SP59 AS CUST03,SP65 AS CUST04,SP66 AS CUST05,SP26 AS FCNO" & _
         " From caseprogress a,servicepractice,engineerprogress, staff,casepropertymap" & _
         " where cp10='201' and cp14 is not null and cp48+0>0 and ep02(+)=cp09 " & stCon & stConCP & " and st01(+)=cp14 and st03 in ('F21','F81')" & _
         " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP & _
         " union all select st16 R01,ep04 R02,st02 R03,'核稿' R04,workdaydiff(nvl(cp27," & strSrvDate(1) & "),ep08) R05" & _
         ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) R06, sqldatet(cp27) R07, sqldatet(ep08) R08" & _
         ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS CASENO,SP08 AS CUST01,SP58 AS CUST02,SP59 AS CUST03,SP65 AS CUST04,SP66 AS CUST05,SP26 AS FCNO" & _
         " From caseprogress a,servicepractice,engineerprogress, staff" & _
         " where cp10='201' and ep02(+)=cp09 and ep08>0 and ep04 is not null " & stCon & stConEP & " and st01(+)=ep04 and st03 in ('F21','F81')" & _
         " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & stConSP
         
      strExcept = ProcExceptList("select X.* from (" & stVTB & ") X where R05<0 and abs(R05)>" & Val(Txt1(1))) 'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷，列出排除案件
      'Modifeid by Lydia 2019/11/01 + 排除案件strExcept
      strExc(0) = "select * from (" & stVTB & ") where R05<0 and abs(R05)>" & Val(Txt1(1)) & strExcept & " order by R01,R02,R04,R06,R07"
   End If

   intI = 1
   Set adoReport = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Added by Lydia 2019/11/01
      If intCufaCnt > 0 Then
           MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
      End If
      'end 2019/11/01
      
      If Txt1(14) = "1" Then
         m_bPrinter = False
         Set m_Device = Picture1
         m_Device.AutoRedraw = True
         If Txt1(12) = "3" Then
            m_Device.Width = 11904
            m_Device.Height = 16836
         Else
            m_Device.Width = 16836
            m_Device.Height = 11904
         End If
         DelPic
      Else
         m_bPrinter = True
         Set m_Device = Printer
         If Txt1(12) = "3" Then
            Printer.Orientation = 1
         Else
            Printer.Orientation = 2
         End If
      End If
      
      If Txt1(12) = "1" Then
         GetPleft1
      ElseIf Txt1(12) = "2" Then
         GetPleft2
      Else
         GetPleft3
      End If
      
      m_iPages = 0
      m_stTotal = ""
      m_stCount = ""
      With adoReport
      m_stGrp = .Fields("R01")
      SetNameList m_stGrp
      m_iRound = 1
      PrintTitle
      
      If Txt1(12) = "3" Then
         Do While Not .EOF
            If m_stGrp <> .Fields("R01") Then
               PrintCount
               m_stGrp = .Fields("R01")
               m_stCount = ""
               PrintTitle
            End If
            NewLine , True
            m_Device.CurrentX = PLeft(1)
            m_Device.CurrentY = iPrint
            m_Device.Print "" & .Fields("R03") '工程師
            m_Device.CurrentX = PLeft(2)
            m_Device.CurrentY = iPrint
            m_Device.Print "" & .Fields("R06") '本所案號
            m_Device.CurrentX = PLeft(3)
            m_Device.CurrentY = iPrint
            m_Device.Print "" & .Fields("R07") '發文日
            m_Device.CurrentX = PLeft(4)
            m_Device.CurrentY = iPrint
            m_Device.Print "" & .Fields("R08") '承辦期限
            strExc(1) = "" & .Fields("R05") '天數差
            m_Device.CurrentX = PLeft(6) - 240 - m_Device.TextWidth(strExc(1))
            m_Device.CurrentY = iPrint
            m_Device.Print strExc(1)
            m_Device.CurrentX = PLeft(6)
            m_Device.CurrentY = iPrint
            m_Device.Print "" & .Fields("R04") '案件性質
            m_stCount = Val(m_stCount) + 1
            .MoveNext
         Loop
         PrintCount
         
      Else
         Do While Not .EOF
            '組別不同跳頁 'Memo by Lydia 2019/09/02 誤刪,補回原程式
            If m_stGrp <> .Fields("R01") Then
               PrintSubTot
               '欄位數超過
               If UBound(arrNameList, 2) > m_iRound * m_iMaxNum Then
                  m_iRound = m_iRound + 1
                  PrintTitle
                  .MoveFirst
                  .Find "R01='" & m_stGrp & "'"
               Else
                  PrintTotal
                  m_stGrp = .Fields("R01")
                  SetNameList m_stGrp
                  m_iRound = 1
                  m_stTotal = ""
                  m_stCount = ""
                  PrintTitle
               End If
               stPty = ""
            End If
            'end 2019/09/02
            
            '列印-案件性質
            If stPty <> "" & .Fields("R04") Then
               NewLine , True
               stPty = "" & .Fields("R04")
               m_Device.CurrentX = PLeft(1)
               m_Device.CurrentY = iPrint
               m_Device.Print StrToStr(stPty, 12)
            End If
            
            '比對員工編號符合,才列出
            iFrom = (m_iRound - 1) * m_iMaxNum + 1
            If UBound(arrNameList, 2) > iFrom + m_iMaxNum - 1 Then
               iTo = iFrom + m_iMaxNum - 1
            Else
               iTo = UBound(arrNameList, 2)
            End If
            For ii = iFrom To iTo
               If arrNameList(1, ii) = .Fields("R02") Then
                  If Not IsNull(.Fields("R05")) Then
                     strExc(1) = .Fields("R05") & Format("(" & .Fields("R06") & ")", "@@@@")
                     m_Device.CurrentX = PLeft(IIf(ii Mod m_iMaxNum = 0, m_iMaxNum, ii Mod m_iMaxNum) + 2) - 240 - m_Device.TextWidth(strExc(1))
                     m_Device.CurrentY = iPrint
                     m_Device.Print strExc(1)
                     arrNameList(3, ii) = Val(arrNameList(3, ii)) + Val("" & .Fields("R05"))
                     arrNameList(4, ii) = Val(arrNameList(4, ii)) + Val("" & .Fields("R06"))
                  End If
                  Exit For
               End If
            Next
            .MoveNext
            If .EOF Then 'Memo by Lydia 2019/08/19 尚未列印所有資料,因為欄位數超過1頁
               PrintSubTot
               '欄位數超過
               If UBound(arrNameList, 2) > m_iRound * m_iMaxNum Then
                  m_iRound = m_iRound + 1
                  PrintTitle
                  .MoveFirst
                  .Find "R01='" & m_stGrp & "'"
               End If
            End If
         Loop
         PrintTotal
      End If
      
      PrintMemo
      
      If m_bPrinter = True Then
         Printer.EndDoc
         ShowPrintOk
      ElseIf m_iPages > 0 Then
         SetPic m_iPages
         frm060404_1.m_ImageW = m_Device.Width
         frm060404_1.m_ImageH = m_Device.Height
         frm060404_1.m_iPages = m_iPages
         frm060404_1.Show
      End If
      
      InsertQueryLog (.RecordCount)
      End With
   Else
      InsertQueryLog (0)
      MsgBox "無符合資料!!"
   End If
End Sub

Private Sub PrintMemo()
   Dim ii As Integer
   
   NewLine
   For ii = 2 To 6
      NewLine
      m_Device.CurrentX = PLeft(1)
      m_Device.CurrentY = iPrint
      If Txt1(12) = "3" Then
         strExc(0) = StrToStr(Label1(ii), 45)
         m_Device.Print strExc(0)
         If Mid(Label1(ii), Len(strExc(0)) + 1) <> "" Then
            NewLine
            m_Device.CurrentX = PLeft(1)
            m_Device.CurrentY = iPrint
            m_Device.Print Mid(Label1(ii), Len(strExc(0)) + 1)
         End If
      Else
         m_Device.Print Label1(ii)
      End If
   Next
End Sub

Private Sub NewLine(Optional iHeight As Integer = 400, Optional bDrawLine As Boolean)
   iPrint = iPrint + iHeight
   If iPrint > m_Device.ScaleHeight - 800 Then
      If bDrawLine Then
         iPrint = iPrint - iHeight
         PrintLine
      End If
      PrintTitle
      iPrint = iPrint + 300
   End If
End Sub

Private Sub SetNameList(pGrp As String)
   Dim ii As Integer, stLstR02 As String
   Set RsTemp = adoReport.Clone
   RsTemp.Sort = "R01,R02"
   With RsTemp
   .MoveFirst
   Erase arrNameList
   ii = 0
   stLstR02 = ""
   Do While Not .EOF
      If .Fields("R01") = pGrp And stLstR02 <> .Fields("R02") Then
         ii = ii + 1
         ReDim Preserve arrNameList(4, ii) As String
         arrNameList(1, ii) = "" & .Fields("R02")
         arrNameList(2, ii) = "" & .Fields("R03")
         arrNameList(3, ii) = 0
         arrNameList(4, ii) = 0
         stLstR02 = "" & .Fields("R02")
      End If
      .MoveNext
   Loop
   End With
End Sub

Private Sub DelPic()
   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_*.tmp"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
End Sub


Private Sub SetPic(idx As Integer)

   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_" & idx & ".tmp"
   
'   Clipboard.Clear
'   Clipboard.SetData Picture1.Image
'   Set m_Pictures(m_iPages - 1) = Clipboard.GetData
'   Set m_Pictures(idx) = Picture1.Image

   SavePicture Picture1.Image, strPicFileName
   '要用覆蓋的否則會錯誤--VB Bug
   'Picture1.Cls
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
   
End Sub

Private Sub PrintLine(Optional iType As Integer = 0)
   iPrint = iPrint + 300
   m_Device.CurrentX = PLeft(1)
   m_Device.CurrentY = iPrint
   If iType = 1 Then
      m_Device.Line (PLeft(1), iPrint + 150)-(m_Device.ScaleWidth - 200, iPrint + 150)
   Else
      m_Device.Print String(Round((m_Device.ScaleWidth - PLeft(1) - 200) / m_Device.TextWidth("-")), "-")
   End If
End Sub


Private Sub PrintTitle()
   Dim stCon As String
   Dim stTmp As String
   
   m_iPages = m_iPages + 1
      
   If m_iPages > 1 Then
      If m_bPrinter = False Then
         SetPic m_iPages - 1
      ElseIf m_iPages > 1 Then
         Printer.NewPage
      End If
   End If
   
   If Option1.Value Then
      If Val(Txt1(12)) = 3 Then
         stTmp = "工程師發文與承辦期限比較明細表"
      Else
         stCon = ""
         If Val(Txt1(12)) = 2 Then
            stCon = "各組"
         End If
         stTmp = stCon & "工程師發文與承辦期限比較統計表"
      End If
   Else
      If Val(Txt1(12)) = 3 Then
         stTmp = "工程師已收文未發文逾承辦期限比較明細表"
      Else
         stCon = ""
         If Val(Txt1(12)) = 2 Then
            stCon = "各組"
         End If
         stTmp = stCon & "工程師已收文未發文逾承辦期限比較統計表"
      End If
   End If
   
   iPrint = 500
   m_Device.FontName = "細明體"
   m_Device.Font.Size = 22
   m_Device.Font.Bold = True
   m_Device.Font.Underline = True
   m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(stTmp)) / 2
   m_Device.CurrentY = iPrint
   m_Device.Print stTmp
      
   
   m_Device.Font.Bold = False
   m_Device.Font.Underline = False
   m_Device.Font.Size = 12
      
   iPrint = iPrint + 500
   If Option1.Value Then
      stTmp = "發文日：" & Format(ChangeTStringToTDateString(Txt1(2)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(Txt1(3))
      m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(stTmp)) / 2
      m_Device.CurrentY = iPrint
      m_Device.Print stTmp
   End If
   
   iPrint = iPrint + 300
   m_Device.CurrentX = 500
   m_Device.CurrentY = iPrint
   m_Device.Print "列印人：" & strUserName
   
   m_Device.CurrentX = m_Device.Width - 3400
   m_Device.CurrentY = iPrint
   m_Device.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   iPrint = iPrint + 300
   
   If Val(Txt1(12)) <> 2 Then
      m_Device.CurrentX = 500
      m_Device.CurrentY = iPrint
      m_Device.Print "工程師組別：" & PUB_GetFCPGrpName(m_stGrp, True)
   End If
            
   m_Device.CurrentX = m_Device.Width - 3400
   m_Device.CurrentY = iPrint
   m_Device.Print "頁    次：" & str(m_iPages)
   PrintLine 1
   
   If Txt1(12) = "3" Then
      PrintTitle2
   Else
      PrintTitle1
   End If
End Sub

'列印欄位名稱
Sub PrintTitle1()
   Dim ii As Integer, jj As Integer
   Dim iY1 As Long, iY2 As Long
   
   
   iPrint = iPrint + 300
   
   iY1 = iPrint
   iY2 = iPrint
   
   
   For ii = 1 To m_iMaxNum
      jj = m_iMaxNum * (m_iRound - 1) + ii
      If UBound(arrNameList, 2) >= jj Then
         strExc(0) = StrToStr(arrNameList(2, jj), Val(m_iColMax))
         If strExc(0) <> arrNameList(2, jj) Then
            iY2 = iY1 + 300
            Exit For
         End If
      End If
   Next
   
   m_Device.CurrentX = PLeft(1)
   m_Device.CurrentY = iY2
   m_Device.Print "案件性質"
   
   For ii = 1 To m_iMaxNum
      jj = m_iMaxNum * (m_iRound - 1) + ii
      If UBound(arrNameList, 2) >= jj Then
         m_Device.CurrentX = PLeft(ii + 1)
         
         strExc(0) = StrToStr(arrNameList(2, jj), Val(m_iColMax))
         If strExc(0) = arrNameList(2, jj) Then
            m_Device.CurrentY = iY2
            m_Device.Print arrNameList(2, jj)
         Else
            m_Device.CurrentY = iY1
            m_Device.Print strExc(0)
            m_Device.CurrentX = PLeft(ii + 1)
            m_Device.CurrentY = iY2
            m_Device.Print StrToStr(Mid(arrNameList(2, jj), Len(strExc(0)) + 1), Val(m_iColMax))
         End If
      Else
         Exit For
      End If
   Next
   iPrint = iY2 - 100
   PrintLine 1
End Sub


'列印欄位名稱(明細)
Sub PrintTitle2()
   
   iPrint = iPrint + 300
   
   m_Device.CurrentX = PLeft(3)
   m_Device.CurrentY = iPrint
   m_Device.Print "完搞日/"
   m_Device.CurrentX = PLeft(4)
   m_Device.CurrentY = iPrint
   m_Device.Print "核稿期限/"
   
   iPrint = iPrint + 300
   
   m_Device.CurrentX = PLeft(1)
   m_Device.CurrentY = iPrint
   m_Device.Print "工程師"
   m_Device.CurrentX = PLeft(2)
   m_Device.CurrentY = iPrint
   m_Device.Print "本所案號"
   m_Device.CurrentX = PLeft(3)
   m_Device.CurrentY = iPrint
   m_Device.Print "發文日"
   m_Device.CurrentX = PLeft(4)
   m_Device.CurrentY = iPrint
   m_Device.Print "承辦期限"
   m_Device.CurrentX = PLeft(5)
   m_Device.CurrentY = iPrint
   m_Device.Print "天數差"
   m_Device.CurrentX = PLeft(6)
   m_Device.CurrentY = iPrint
   m_Device.Print "案件性質"
   
   iPrint = iPrint - 100
   PrintLine 1
End Sub

Sub GetPleft1()
   Dim ii As Integer
      
   Erase PLeft
   
   m_iMaxNum = 10 '每頁工程師數
   m_iColMax = 3 '工程師名稱每列字數
   
   PLeft(0) = 500
   PLeft(1) = 500
   PLeft(2) = PLeft(1) + 3120
   For ii = 0 To m_iMaxNum
      PLeft(3 + ii) = PLeft(2 + ii) + 1200
   Next
End Sub

Sub GetPleft2()
   Dim ii As Integer
      
   Erase PLeft
   
   m_iMaxNum = 6 '每頁組別數
   m_iColMax = 7 '組別名稱每列字數
   
   PLeft(0) = 500
   PLeft(1) = 500
   PLeft(2) = PLeft(1) + 3120
   For ii = 0 To m_iMaxNum
      PLeft(3 + ii) = PLeft(2 + ii) + 1920
   Next
End Sub


Sub GetPleft3()
   
   Erase PLeft
      
   PLeft(0) = 500
   PLeft(1) = 500 '工程師
   PLeft(2) = PLeft(1) + 1680 '本所案號
   PLeft(3) = PLeft(2) + 2160 '發文日
   PLeft(4) = PLeft(3) + 1200 '承辦期限
   PLeft(5) = PLeft(4) + 1200 '天數差
   PLeft(6) = PLeft(5) + 960 '案件性質
   PLeft(7) = PLeft(6) + 3120
End Sub

Private Sub PrintSubTot()
   Dim ii As Integer, iFrom As Integer, iTo As Integer
   
   PrintLine 1
   NewLine
   
   m_Device.CurrentX = PLeft(2) - 480 - m_Device.TextWidth("合計")
   m_Device.CurrentY = iPrint
   m_Device.Print "合計"
   
   iFrom = (m_iRound - 1) * m_iMaxNum + 1
   If UBound(arrNameList, 2) > iFrom + m_iMaxNum - 1 Then
      iTo = iFrom + m_iMaxNum - 1
   Else
      iTo = UBound(arrNameList, 2)
   End If
   For ii = iFrom To iTo
      strExc(1) = arrNameList(3, ii) & Format("(" & arrNameList(4, ii) & ")", "@@@@")
      m_Device.CurrentX = PLeft(IIf(ii Mod m_iMaxNum = 0, m_iMaxNum, ii Mod m_iMaxNum) + 2) - 240 - m_Device.TextWidth(strExc(1))
      m_Device.CurrentY = iPrint
      m_Device.Print strExc(1)
      m_stTotal = Val(m_stTotal) + Val(arrNameList(3, ii))
      m_stCount = Val(m_stCount) + Val(arrNameList(4, ii))
   Next
End Sub

Private Sub PrintTotal()
   NewLine
   m_Device.CurrentX = PLeft(2) - 480 - m_Device.TextWidth("總計")
   m_Device.CurrentY = iPrint
   m_Device.Print "總計"
   
   strExc(1) = m_stTotal & Format("(" & m_stCount & ")", "@@@@")
   m_Device.CurrentX = PLeft(3) - 240 - m_Device.TextWidth(strExc(1))
   m_Device.CurrentY = iPrint
   m_Device.Print strExc(1)
End Sub


Private Sub PrintCount()
   NewLine
   m_Device.CurrentX = PLeft(1)
   m_Device.CurrentY = iPrint
   m_Device.Print "共 " & m_stCount & " 筆資料"
End Sub
