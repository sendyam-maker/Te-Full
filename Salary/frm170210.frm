VERSION 5.00
Begin VB.Form frm170210 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工勞健保保費及勞退自提明細表"
   ClientHeight    =   2544
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2544
   ScaleWidth      =   4740
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   1530
      MaxLength       =   5
      TabIndex        =   0
      Top             =   770
      Width           =   675
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1530
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "1"
      Top             =   1110
      Width           =   200
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   2
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   7
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2580
      TabIndex        =   3
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3660
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "類　　別：      (1：勞保 2：健保 3：勞退自提)"
      Height          =   180
      Left            =   600
      TabIndex        =   8
      Top             =   1155
      Width           =   3630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "薪資月份："
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   810
      Width           =   900
   End
End
Attribute VB_Name = "frm170210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/8/30 改成Form2.0 (以圖片方式列印Unicode文字)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2009/2/2 add by sonia
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblAmt1 As Double, dblAmt2 As Double, dblAmt3 As Double        '小計
Dim dblTotAmt1 As Double, dblTotAmt2 As Double, dblTotAmt3 As Double  '合計
Dim dblCnt As Double
Dim dblTotCnt As Double

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If txt1(0) = "" Then
            MsgBox "薪資月份不可以空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
         End If
         If txt1(0) <> "" Then
            If Len(txt1(0)) <= 3 Then
               MsgBox "薪資月份輸入錯誤！", vbInformation, "操作錯誤！"
               txt1(0).SetFocus
               Exit Sub
            End If
            If ChkDate(txt1(0) & "01") = False Then
               txt1(0).SetFocus
               Exit Sub
            End If
         End If
         If txt1(1) = "" Then
            MsgBox "類別不可以空白！", vbInformation, "操作錯誤！"
            txt1(1).SetFocus
            Exit Sub
         End If
         
         Screen.MousePointer = vbHourglass
         StrMenu
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Sub StrMenu()
Dim strYM As String
Dim WDs As String '當月工作天數   '2010/5/27 ADD BY SONIA
 
   strYM = Val(txt1(0)) + 191100
   '2010/5/27 ADD BY SONIA
   WDs = Right(CompDate(2, -1, CompDate(1, 1, Format(100 * strYM + 1))), 2)
   
   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   
   m_str = ""
   m_StrSQL = " AND SM02= " & strYM
   
   Select Case txt1(1)
      Case "1"  '勞保
         If Val(strYM) < 200900 Then
            '2009/11/26 MODIFY BY SONIA 投保薪資改抓每月薪資檔
            'm_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
                  "NVL(SM14,0) 勞保費,Round(SI02 * DECODE(ST24,'F',5.5,6.5) / 100 * IR04 / 100, 0) 投保單位保險費,A0802 公司名稱 " & _
                  "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
                  "WHERE SM01=ST01(+) and nvl(sm14,0)>0 AND SM01=SD01(+) AND SM37=A0801(+) " & _
                  "AND 'L'=SI01(+) AND SI03<=NVL(SD12,(NVL(SD20,0)+NVL(SD21,0)+NVL(SD23,0))) AND SI04>=NVL(SD12,(NVL(SD20,0)+NVL(SD21,0)+NVL(SD23,0))) " & m_StrSQL
            '2010/5/27 MODIFY BY SONIA 離職人員的投保單位費用也要依工作天數計算
            'm_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
                  "NVL(SM14,0) 勞保費,Round(SI02 * DECODE(ST24,'F',5.5,6.5) / 100 * IR04 / 100, 0) 投保單位保險費,A0802 公司名稱 " & _
                  "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
                  "WHERE SM01=ST01(+) and nvl(sm14,0)>0 AND SM01=SD01(+) AND SM37=A0801(+) " & _
                  "AND 'L'=SI01(+) AND SI03<=NVL(SM40,0) AND SI04>=NVL(SM40,0)) " & m_StrSQL
                  
            'Modified by Morgan 2012/11/28 改判斷勞保投保薪資>0 (因94026勞保是全額補助)
            m_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
                  "NVL(SM14,0) 勞保費,DECODE(SM27," & WDs & ",Round(SI02 * DECODE(ST24,'F',5.5,6.5) / 100 * IR04 / 100, 0),Round(SI02 * DECODE(ST24,'F',5.5,6.5) / 100 * IR04 / 100 * SM27 / " & WDs & ", 0)) 投保單位保險費,A0802 公司名稱 " & _
                  "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
                  "WHERE SM01=ST01(+) and SM40>0 AND SM01=SD01(+) AND SM37=A0801(+) " & _
                  "AND 'L'=SI01(+) AND SI03<=NVL(SM40,0) AND SI04>=NVL(SM40,0)) " & m_StrSQL
         
         ElseIf Val(strYM) < 201300 Then
            '2009/11/26 MODIFY BY SONIA 投保薪資改抓每月薪資檔
            'm_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
                  "NVL(SM14,0) 勞保費,Round(SI02 * DECODE(ST24,'F',IR02,IR01) / 100 * IR04 / 100, 0) 投保單位保險費,A0802 公司名稱 " & _
                  "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
                  "WHERE SM01=ST01(+) and nvl(sm14,0)>0 AND SM01=SD01(+) AND SM37=A0801(+) " & _
                  "AND 'L'=SI01(+) AND SI03<=NVL(SD12,(NVL(SD20,0)+NVL(SD21,0)+NVL(SD23,0))) AND SI04>=NVL(SD12,(NVL(SD20,0)+NVL(SD21,0)+NVL(SD23,0))) " & m_StrSQL
            '2010/5/27 MODIFY BY SONIA 離職人員的投保單位費用也要依工作天數計算
            'm_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
                  "NVL(SM14,0) 勞保費,Round(SI02 * DECODE(ST24,'F',IR02,IR01) / 100 * IR04 / 100, 0) 投保單位保險費,A0802 公司名稱 " & _
                  "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
                  "WHERE SM01=ST01(+) and nvl(sm14,0)>0 AND SM01=SD01(+) AND SM37=A0801(+) " & _
                  "AND 'L'=SI01(+) AND SI03<=NVL(SM40,0) AND SI04>=NVL(SM40,0) " & m_StrSQL
                  
            'Modify by Morgan 2010/10/28 勞保費費率要個別四捨五入計算,且僅雇主沒有就業保險
            'm_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
                  "NVL(SM14,0) 勞保費,DECODE(SM27," & WDs & ",Round(SI02 * DECODE(ST24,'F',IR02,IR01) / 100 * IR04 / 100, 0),Round(SI02 * DECODE(ST24,'F',IR02,IR01) / 100 * IR04 / 100 * SM27 / " & WDs & ", 0)) 投保單位保險費,A0802 公司名稱 " & _
                  "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
                  "WHERE SM01=ST01(+) and nvl(sm14,0)>0 AND SM01=SD01(+) AND SM37=A0801(+) " & _
                  "AND 'L'=SI01(+) AND SI03<=NVL(SM40,0) AND SI04>=NVL(SM40,0) " & m_StrSQL
                  
            'Modified by Morgan 2012/11/28 改判斷勞保投保薪資>0 (因94026勞保是全額補助)
            m_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
                  "NVL(SM14,0) 勞保費,Round(DECODE(SM27," & WDs & ",1,SM27 / " & WDs & ")*(ROUND(SI02*IR01/100*IR04/100)+DECODE(ST20,'11',0,ROUND(SI02*IR02/100*IR04/100))+ROUND(SI02*IR17/100))) 投保單位保險費,A0802 公司名稱 " & _
                  "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
                  "WHERE SM01=ST01(+) and SM40>0 AND SM01=SD01(+) AND SM37=A0801(+) " & _
                  "AND 'L'=SI01(+) AND SI03<=NVL(SM40,0) AND SI04>=NVL(SM40,0) " & m_StrSQL
            
            'Add by Morgan 2009/6/26 98年6月起第2家也要繳勞保費
            'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
            'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
            m_str = m_str & " UNION ALL SELECT SM37 公司別,0 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號" & _
                  ",NVL(SM14,0) 勞保費,0 投保單位保險費,A0802 公司名稱 " & _
                  " FROM SALARYMONTH,STAFF,ACC080 " & _
                  " WHERE SUBSTR(SM01,3,1)='A' AND SM14>0 AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=ST01(+)" & _
                  " AND SM37=A0801(+)" & m_StrSQL
                  
         'Added by Morgan 2013/3/29 是否含失業給付改判斷是否以合夥人投保
         ElseIf Val(strYM) < 202205 Then
            'Modified by Morgan 2016/1/15 +其他給付扣款資料
            'm_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
                  "NVL(SM14,0) 勞保費,Round(DECODE(SM27," & WDs & ",1,SM27 / " & WDs & ")*(ROUND(SI02*IR01/100*IR04/100)+DECODE(SD11,'Y',0,ROUND(SI02*IR02/100*IR04/100))+ROUND(SI02*IR17/100))) 投保單位保險費,A0802 公司名稱 " & _
                  "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
                  "WHERE SM01=ST01(+) and SM40>0 AND SM01=SD01(+) AND SM37=A0801(+) " & _
                  "AND 'L'=SI01(+) AND SI03<=NVL(SM40,0) AND SI04>=NVL(SM40,0) " & m_StrSQL
            'Modified by Morgan 2017/12/27 台一投資、智權公司職災保險費率不同改抓IR18
            'Modified by Morgan 2020/11/26 改智慧所職災保險費率抓IR17其他抓IR18
            'Modified by Morgan 2022/8/30 +判斷有提撥的不管是否已設定為合夥人身分投保(未生效月份 Ex:蔣律師)--辜
            m_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號 " & _
                  ",NVL(SM14,0)+NVL(OD05,0) 勞保費,Round(DECODE(SM27," & WDs & ",1,SM27 / " & WDs & ")*(ROUND(SI02*IR01/100*IR04/100)+DECODE(decode(nvl(sm30,0),0,sd11),'Y',0,ROUND(SI02*IR02/100*IR04/100))+ROUND(SI02*DECODE(SM37,'2',IR17,IR18)/100))) 投保單位保險費,A0802 公司名稱 " & _
                  " FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
                  ",(select od03,od14,sum(decode(od04,'31',od05,'32',-1*od05)) od05 from SALARYMONTH,othersalarydata where od03(+)=sm01 and od14(+)=sm02 and od04 in ('31','32')" & m_StrSQL & " group by od03,od14) X" & _
                  " WHERE SM01=ST01(+) and SM40>0 AND SM01=SD01(+) AND SM37=A0801(+) and od03(+)=sm01 and od14(+)=sm02" & _
                  " AND 'L'=SI01(+) AND SI03<=NVL(SM40,0) AND SI04>=NVL(SM40,0) " & m_StrSQL
            'end 2016/1/15
            'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
            m_str = m_str & " UNION ALL SELECT SM37 公司別,0 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號" & _
                  ",NVL(SM14,0) 勞保費,0 投保單位保險費,A0802 公司名稱 " & _
                  " FROM SALARYMONTH,STAFF,ACC080 " & _
                  " WHERE SUBSTR(SM01,3,1)='A' AND SM14>0 AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=ST01(+)" & _
                  " AND SM37=A0801(+)" & m_StrSQL
                  
         'end 2013/3/29
         'Added by Morgan 2022/11/30 111/5起職災級數調整 --辜
         Else
            m_str = "SELECT SM37 公司別,NVL(S1.SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號 " & _
                  ",NVL(SM14,0)+NVL(OD05,0) 勞保費,Round(DECODE(SM27," & WDs & ",1,SM27 / " & WDs & ")*(ROUND(S1.SI02*IR01/100*IR04/100)+DECODE(decode(nvl(sm30,0),0,sd11),'Y',0,ROUND(S1.SI02*IR02/100*IR04/100))+ROUND(S2.SI02*DECODE(SM37,'2',IR17,IR18)/100))) 投保單位保險費,A0802 公司名稱 " & _
                  ",S2.SI02 職災級數 FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance S1,SalaryInsurance S2,InsuranceRate,ACC080 " & _
                  ",(select od03,od14,sum(decode(od04,'31',od05,'32',-1*od05)) od05 from SALARYMONTH,othersalarydata where od03(+)=sm01 and od14(+)=sm02 and od04 in ('31','32')" & m_StrSQL & " group by od03,od14) X" & _
                  " WHERE SM01=ST01(+) and SM40>0 AND SM01=SD01(+) AND SM37=A0801(+) and od03(+)=sm01 and od14(+)=sm02" & _
                  " AND 'L'=S1.SI01(+) AND S1.SI03<=NVL(SM40,0) AND S1.SI04>=NVL(SM40,0) " & m_StrSQL & _
                  " AND 'D'=S2.SI01(+) AND S2.SI03<=NVL(SM40,0) AND S2.SI04>=NVL(SM40,0) "
            'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
            m_str = m_str & " UNION ALL SELECT SM37 公司別,0 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號" & _
                  ",NVL(SM14,0) 勞保費,0 投保單位保險費,A0802 公司名稱 " & _
                  ",0 職災級數 FROM SALARYMONTH,STAFF,ACC080 " & _
                  " WHERE SUBSTR(SM01,3,1)='A' AND SM14>0 AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=ST01(+)" & _
                  " AND SM37=A0801(+)" & m_StrSQL
         'end 2022/11/30
         End If
      Case "2"  '健保(96006因個人減免為0但公司仍要提撥,故改以SM14>0判斷)
         '2009/2/25 MODIFY BY SONIA 投保單位保險費應再負擔1.7平均眷口數
         '2009/11/26 MODIFY BY SONIA 投保薪資改抓每月薪資檔
         'm_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
               "NVL(SM15,0) 健保費,Round(SI02 * IR06 / 100 * IR08 / 100 * IR16, 0) 投保單位保險費,A0802 公司名稱 " & _
               "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
               "WHERE SM01=ST01(+) and nvl(sm14,0)>0 AND SM01=SD01(+) AND SM37=A0801(+) " & _
               "AND 'H'=SI01(+) AND SI03<=NVL(SD13,(NVL(SD20,0)+NVL(SD21,0)+NVL(SD23,0))) AND SI04>=NVL(SD13,(NVL(SD20,0)+NVL(SD21,0)+NVL(SD23,0))) " & m_StrSQL
         'Modified by Morgan 2012/11/28 改判斷健保投保薪資>0 (因94026勞保也是全額補助)
         
         'Modified by Morgan 2013/3/28 有月薪資但未投保者(Ex.非當月到職但當月離職員工)或低收入戶健保全額補助者,投保單位也不必負擔健保費
         'm_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
               "NVL(SM15,0) 健保費,Round(SI02 * IR06 / 100 * IR08 / 100 * IR16, 0) 投保單位保險費,A0802 公司名稱 " & _
               "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
               "WHERE SM01=ST01(+) and SM41>0 and SM01=SD01(+) AND SM37=A0801(+) " & _
               "AND 'H'=SI01(+) AND SI03<=NVL(SM41,0) AND SI04>=NVL(SM41,0) " & m_StrSQL
               
         '判斷補助類別為12(低收入戶)或健保費為0且無健保明細者(因判斷當月離職還需考慮該是否為當月到職且全額補助者...)
         'Modified by Morgan 2013/5/28 以合夥人身分投保 100% 個人負擔
         'Modified by Morgan 2016/1/15 +其他給付扣款資料
         'm_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
               "NVL(SM15,0) 健保費,decode(sd11,'Y',0,decode(sign(sm15),1,1,decode(hm06,'12',0,null,0,1))) * Round(SI02 * IR06 / 100 * IR08 / 100 * IR16, 0) 投保單位保險費,A0802 公司名稱 " & _
               "FROM SALARYMONTH,Himonth,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
               "WHERE SM01=ST01(+) and SM41>0 and hm03(+)=sm02 and hm01(+)=sm01 and hm02(+)=0 and SM01=SD01(+) AND SM37=A0801(+) " & _
               "AND 'H'=SI01(+) AND SI03<=NVL(SM41,0) AND SI04>=NVL(SM41,0) " & m_StrSQL
         'Modified by Morgan 2022/8/30 +判斷有提撥的不管是否已設定為合夥人身分投保(未生效月份 Ex:蔣律師)--辜
         m_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
               " NVL(SM15,0)+NVL(OD05,0) 健保費,decode(decode(nvl(sm30,0),0,sd11),'Y',0,decode(sign(sm15),1,1,decode(hm06,'12',0,null,0,1))) * Round(SI02 * IR06 / 100 * IR08 / 100 * IR16, 0) 投保單位保險費,A0802 公司名稱 " & _
               " FROM SALARYMONTH,Himonth,STAFF,SALARYDATA,SalaryInsurance,InsuranceRate,ACC080 " & _
               ",(select od03,od14,sum(decode(od04,'35',od05,'36',-1*od05)) od05 from SALARYMONTH,othersalarydata where od03(+)=sm01 and od14(+)=sm02 and od04 in ('35','36')" & m_StrSQL & " group by od03,od14) X" & _
               " WHERE SM01=ST01(+) and SM41>0 and hm03(+)=sm02 and hm01(+)=sm01 and hm02(+)=0 and SM01=SD01(+) AND SM37=A0801(+) and od03(+)=sm01 and od14(+)=sm02" & _
               " AND 'H'=SI01(+) AND SI03<=NVL(SM41,0) AND SI04>=NVL(SM41,0) " & m_StrSQL
         'end 2016/1/15
         
      Case "3"  '勞退自提
         '2009/11/26 MODIFY BY SONIA 投保薪資改抓每月薪資檔
         'm_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
               "NVL(SM16,0) 勞退自提,NVL(SM30,0) 公司提撥,A0802 公司名稱 " & _
               "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,ACC080 " & _
               "WHERE SUBSTR(SM01,3,1)<>'A' AND SM01=ST01(+) and nvl(sm16,0)+nvl(sm30,0)>0 AND SM01=SD01(+) AND SM37=A0801(+) " & _
               "AND 'R'=SI01(+) AND SI03<=NVL(SD27,(NVL(SD20,0)+NVL(SD21,0)+NVL(SD23,0))) AND SI04>=NVL(SD27,(NVL(SD20,0)+NVL(SD21,0)+NVL(SD23,0))) " & m_StrSQL
         'Modified by Morgan 2016/1/15 +其他給付扣款資料
         'm_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
               "NVL(SM16,0) 勞退自提,NVL(SM30,0) 公司提撥,A0802 公司名稱 " & _
               "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,ACC080 " & _
               "WHERE SUBSTR(SM01,3,1)<>'A' AND SM01=ST01(+) and nvl(sm16,0)+nvl(sm30,0)>0 AND SM01=SD01(+) AND SM37=A0801(+) " & _
               "AND 'R'=SI01(+) AND SI03<=NVL(SM38,0) AND SI04>=NVL(SM38,0) " & m_StrSQL
         m_str = "SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
               " NVL(SM16,0)+NVL(OD05,0) 勞退自提,NVL(SM30,0) 公司提撥,A0802 公司名稱 " & _
               " FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,ACC080 " & _
               ",(select od03,od14,sum(decode(od04,'33',od05,'34',-1*od05)) od05 from SALARYMONTH,othersalarydata where od03(+)=sm01 and od14(+)=sm02 and od04 in ('33','34')" & m_StrSQL & " group by od03,od14) X" & _
               " WHERE SUBSTR(SM01,3,1)<>'A' AND SM01=ST01(+) and nvl(sm16,0)+nvl(sm30,0)>0 AND SM01=SD01(+) AND SM37=A0801(+) and od03(+)=sm01 and od14(+)=sm02" & _
               " AND 'R'=SI01(+) AND SI03<=NVL(SM38,0) AND SI04>=NVL(SM38,0) " & m_StrSQL
         'end 2016/1/15
         
         '2009/11/26 END
         'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
         'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
         m_str = m_str & " UNION SELECT SM37 公司別,NVL(SI02,0) 月投保金額,SM01 員工編號,ST02 姓名,ST26 身分證字號, " & _
               "NVL(SM16,0) 勞退自提,NVL(SM30,0) 公司提撥,A0802 公司名稱 " & _
               "FROM SALARYMONTH,STAFF,SALARYDATA,SalaryInsurance,ACC080 " & _
               "WHERE SUBSTR(SM01,3,1)='A' AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=ST01(+) and nvl(sm16,0)+nvl(sm30,0)>0 AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=SD01(+) AND SM37=A0801(+) " & _
               "AND 'R'=SI01(+) AND SI03<=NVL(SM38,0) AND SI04>=NVL(SM38,0) " & m_StrSQL
   End Select
   
   'Added by Morgan 2022/12/1
   If txt1(1) = "1" And Val(strYM) >= 202205 Then
      m_str = m_str & " order by 公司別,月投保金額,職災級數,員工編號"
   Else
   'end 2022/12/1
      m_str = m_str & " order by 公司別,月投保金額,員工編號"
   End If
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         iLine = 1
         strType = "" '切頁條件
         dblAmt1 = 0: dblAmt2 = 0: dblAmt3 = 0: dblCnt = 0
         dblTotAmt1 = 0: dblTotAmt2 = 0: dblTotAmt3 = 0: dblTotCnt = 0
         
         Do While Not m_rs.EOF
             
            For m_i = 1 To 10
               strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields(0))  '公司別
            strTemp(2) = CheckStr(m_rs.Fields(1))  '月投保金額
            strTemp(3) = CheckStr(m_rs.Fields(2))  '員工編號
            strTemp(4) = CheckStr(m_rs.Fields(3))  '姓名
            'Modified by Morgan 2025/4/14 取消(個資)--婉莘
            'strTemp(5) = CheckStr(m_rs.Fields(4))  '身分證字號
            strTemp(5) = ""
            'end 2025/4/14
            strTemp(6) = CheckStr(m_rs.Fields(5))  '個人費用
            strTemp(7) = CheckStr(m_rs.Fields(6))  '公司費用
            strTemp(8) = CheckStr(m_rs.Fields(7))  '公司名稱
            
            If iLine > 50 Or iLine = 1 Or strType <> strTemp(1) Then
                     
               If (strType <> "" And strType <> strTemp(1)) Then
                  PrintEnd '小計
               End If
               
               If strType <> "" Then Printer.NewPage
               iLine = 1
               PrintTitle '列印表頭
            End If
            
            PrintDetail '列印表中
            
            strType = strTemp(1) '依公司別跳頁
            
            dblAmt1 = dblAmt1 + strTemp(6)        '小計
            dblAmt2 = dblAmt2 + strTemp(7)        '小計
            dblAmt3 = dblAmt3 + strTemp(2)        '小計 Added by Morgan 2012/11/27
            dblTotAmt1 = dblTotAmt1 + strTemp(6)  '合計
            dblTotAmt2 = dblTotAmt2 + strTemp(7)  '合計
            dblTotAmt3 = dblTotAmt3 + strTemp(2)  '合計 Added by Morgan 2012/11/27
            dblCnt = dblCnt + 1
            dblTotCnt = dblTotCnt + 1
            m_rs.MoveNext
         Loop
          
         '列印表尾
         PrintEnd '小計
         
         iLine = iLine + 1
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 300
         Printer.Print String(140, "-")
         
         iLine = iLine + 1
         Printer.CurrentX = PLeft(2) - Printer.TextWidth(dblTotCnt & "人")
         Printer.CurrentY = iLine * 300
         Printer.Print dblTotCnt & "人"
         Printer.CurrentX = PLeft(3)
         Printer.CurrentY = iLine * 300
         Printer.Print "合　計："
         Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblTotAmt1, "#,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt1, "#,###,##0")
         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblTotAmt2, "#,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt2, "#,###,##0")
         'Added by Morgan 2012/11/27
         Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(dblTotAmt3, "#,###,##0")) - 300
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt3, "#,###,##0")
         'end 2012/11/27
      End With
   Else
      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintTitle()
   GetPleft
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Select Case txt1(1)
      Case "1"  '勞保
         Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("勞　保　保　費　明　細　表") / 2)
         Printer.CurrentY = iLine * 300
         Printer.Print "勞　保　保　費　明　細　表"
      Case "2"  '健保
         Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("健　保　保　費　明　細　表") / 2)
         Printer.CurrentY = iLine * 300
         Printer.Print "健　保　保　費　明　細　表"
      Case "3"  '勞退自提
         Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("勞　退　自　提　明　細　表") / 2)
         Printer.CurrentY = iLine * 300
         Printer.Print "勞　退　自　提　明　細　表"
   End Select
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   'Modified by Morgan 2022/8/30
   'Printer.Print "列印人：" & strUserName
   PUB_PrintUnicodeText "列印人：" & strUserName, Printer.CurrentX, Printer.CurrentY, 0
   'end 2022/8/30
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("000 年 00 月") / 2)
   Printer.CurrentY = iLine * 300
   If Len(txt1(0)) = 5 Then
      Printer.Print Left(Trim(txt1(0)), 3) & "  年  " & Right(Trim(txt1(0)), 2) & "  月"
   Else
      Printer.Print Left(Trim(txt1(0)), 2) & "  年  " & Right(Trim(txt1(0)), 2) & "  月"
   End If
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "公司別：" & strTemp(1) & "　" & strTemp(8)
   
   iLine = iLine + 2
   Printer.CurrentX = PLeft(1) - Printer.TextWidth("月投保金額")
   Printer.CurrentY = iLine * 300
   Printer.Print "月投保金額"
   Printer.CurrentX = PLeft(2) - Printer.TextWidth("員工編號")
   Printer.CurrentY = iLine * 300
   Printer.Print "員工編號"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("姓　名")
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   'Modified by Morgan 2025/4/14 取消(個資)--婉莘
   'Printer.CurrentX = PLeft(4) - Printer.TextWidth("身份證字號")
   'Printer.CurrentY = iLine * 300
   'Printer.Print "身份證字號"
   'end 2025/4/14
   Select Case txt1(1)
      Case "1"  '勞保
         Printer.CurrentX = PLeft(5) - Printer.TextWidth("勞保費")
         Printer.CurrentY = iLine * 300
         Printer.Print "勞保費"
      Case "2"  '健保
         Printer.CurrentX = PLeft(5) - Printer.TextWidth("健保費")
         Printer.CurrentY = iLine * 300
         Printer.Print "健保費"
      Case "3"  '勞退自提
         Printer.CurrentX = PLeft(5) - Printer.TextWidth("勞退自提")
         Printer.CurrentY = iLine * 300
         Printer.Print "勞退自提"
   End Select
   Printer.CurrentX = PLeft(6) - Printer.TextWidth("投保單位費用")
   Printer.CurrentY = iLine * 300
   Printer.Print "投保單位費用"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
End Sub

Sub PrintEnd()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
   Printer.CurrentX = PLeft(2) - Printer.TextWidth(dblCnt & "人")
   Printer.CurrentY = iLine * 300
   Printer.Print dblCnt & "人"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print "小　計："
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmt1, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt1, "##,###,###")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblAmt2, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt2, "##,###,###")
   
   'Added by Morgan 2012/11/27
   Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(dblAmt3, "#,###,##0")) - 300
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt3, "#,###,##0")
   'end 2012/11/27
         
   dblAmt1 = 0
   dblAmt2 = 0
   dblAmt3 = 0 'Added by Morgan 2012/11/27
   dblCnt = 0
End Sub

Sub GetPleft()
   PLeft(1) = 2000
   PLeft(2) = 3500
   PLeft(3) = 5000
   PLeft(4) = 7000
   PLeft(5) = 8500
   PLeft(6) = 10500
End Sub

Sub PrintDetail()
   Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(strTemp(2), "#,###,##0")) - 300
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(2), "#,###,##0")
   Printer.CurrentX = PLeft(2) - 700
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(3) - 700
   Printer.CurrentY = iLine * 300
   'Modified by Morgan 2022/8/30
   'Printer.Print strTemp(4)
   PUB_PrintUnicodeText strTemp(4), Printer.CurrentX, Printer.CurrentY, 0
   'end 2022/8/30
   Printer.CurrentX = PLeft(4) - 1200
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(5)
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(6), "#,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "#,###,##0")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(7), "#,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(7), "#,###,##0")
   
   iLine = iLine + 1
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSql Then
         SeekPrint = i
      End If
   Next i
   
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170210 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 1
         If KeyAscii < Asc("1") Or KeyAscii > Asc("3") Then
            KeyAscii = 0
            Beep
         End If
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "01") = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
