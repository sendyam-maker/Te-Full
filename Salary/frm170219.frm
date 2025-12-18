VERSION 5.00
Begin VB.Form frm170219 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工應稅薪資檢核表"
   ClientHeight    =   2952
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5064
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2952
   ScaleWidth      =   5064
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2200
      MaxLength       =   1
      TabIndex        =   1
      Top             =   600
      Width           =   300
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1710
      MaxLength       =   1
      TabIndex        =   0
      Top             =   600
      Width           =   300
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   20
      TabIndex        =   9
      Top             =   2160
      Width           =   5000
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   10
         Top             =   180
         Width           =   4200
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   11
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1710
      MaxLength       =   3
      TabIndex        =   2
      Top             =   930
      Width           =   435
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   5
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1620
      Width           =   765
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   6
      Left            =   2610
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1620
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   1710
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1290
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   2200
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1290
      Width           =   375
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2700
      TabIndex        =   7
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3780
      TabIndex        =   8
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "不確定何時用"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   1080
   End
   Begin VB.Line Line2 
      X1              =   1830
      X2              =   2250
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "公  司  別："
      Height          =   180
      Index           =   0
      Left            =   780
      TabIndex        =   15
      Top             =   630
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "薪資月份："
      Height          =   180
      Index           =   2
      Left            =   780
      TabIndex        =   14
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "薪資年度："
      Height          =   180
      Left            =   780
      TabIndex        =   13
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   3
      Left            =   780
      TabIndex        =   12
      Top             =   1650
      Width           =   900
   End
   Begin VB.Line Line3 
      X1              =   2220
      X2              =   2880
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2460
      Y1              =   1410
      Y2              =   1410
   End
End
Attribute VB_Name = "frm170219"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2009/1/16 add by sonia
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL1 As String     '每月薪資資料
Dim m_StrSQL2 As String     '年終獎金資料
Dim m_StrSQL3 As String     '每月獎金資料
Dim m_i As Integer
Dim PLeft(1 To 20) As Integer
Dim strTemp(1 To 20) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblAmtT As Double, dblAmtS As Double, dblAmtO As Double, dblAmtT1 As Double, dblAmtT2 As Double, dblAmtA As Double, dblAmtL As Double, dblAmt01 As Double, dblAmt02 As Double
Dim dblAmtS1 As Double, dblAmtS2 As Double
Dim dblCntT As Double, dblCntT1 As Double, dblCntT2 As Double
Dim dblTAmtT As Double, dblTAmtS As Double, dblTAmtO As Double, dblTAmtT1 As Double, dblTAmtT2 As Double, dblTAmtA As Double, dblTAmtL As Double, dblTAmt01 As Double, dblTAmt02 As Double
Dim dblTAmtS1 As Double, dblTAmtS2 As Double
Dim dblTCntT As Double, dblTCntT1 As Double, dblTCntT2 As Double
Dim dblOldAmt As Double

Private Sub cmdok_Click(Index As Integer)
Dim strYM As String

   Select Case Index
      Case 0
         If txt1(2) = "" Then
            MsgBox "薪資年度不可空白！", vbInformation, "操作錯誤！"
            txt1(2).SetFocus
            Exit Sub
         End If
'         If txt1(3) = "" And txt1(4) = "" Then
'            MsgBox "薪資月份不可空白！", vbInformation, "操作錯誤！"
'            txt1(3).SetFocus
'            Exit Sub
'         End If
         If RunNick(txt1(3), txt1(4)) Then
            txt1(3).SetFocus
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

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 2 '1.直印 2.橫印
   
   m_StrSQL1 = "": m_StrSQL2 = "": m_StrSQL3 = ""
   If txt1(0) <> "" Then '公司別起
      m_StrSQL1 = m_StrSQL1 & " and sm37 >='" & Trim(txt1(0)) & "' "
      m_StrSQL2 = m_StrSQL2 & " and yb24 >='" & Trim(txt1(0)) & "' "
      'Modified by Morgan 2013/5/7 改語法
      'm_StrSQL3 = m_StrSQL3 & " and COMP >='" & Trim(txt1(0)) & "' "
      '2014/2/19 modify by sonia 每月獎金之公司別改抓MB11,原抓SD19
      m_StrSQL3 = m_StrSQL3 & " and MB11 >='" & Trim(txt1(0)) & "' "
   End If
   If txt1(1) <> "" Then '公司別迄
      m_StrSQL1 = m_StrSQL1 & " and SM37 <='" & Trim(txt1(1)) & "' "
      m_StrSQL2 = m_StrSQL2 & " and yb24 <='" & Trim(txt1(1)) & "' "
      'Modified by Morgan 2013/5/7 改語法
      'm_StrSQL3 = m_StrSQL3 & " and COMP <='" & Trim(txt1(1)) & "' "
      '2014/2/19 modify by sonia 每月獎金之公司別改抓MB11,原抓SD19
      m_StrSQL3 = m_StrSQL3 & " and MB11 <='" & Trim(txt1(1)) & "' "
   End If
   If txt1(2) <> "" Then '薪資年度, 年終獎金資料要抓前一年
      m_StrSQL1 = m_StrSQL1 & " and substr(SM02,1,4)='" & Val(txt1(2)) + 1911 & "' "
      m_StrSQL2 = m_StrSQL2 & " and yb01=" & Val(txt1(2)) + 1911 - 1
      m_StrSQL3 = m_StrSQL3 & " and substr(MB01,1,4)='" & Val(txt1(2)) + 1911 & "' "
   End If
   If txt1(3) <> "" Then '薪資月份起
      m_StrSQL1 = m_StrSQL1 & " and substr(SM02,5,2) >='" & Trim(txt1(3)) & "' "
      'Modified by Morgan 2013/3/8 獎金年月已改為給付日期,14月抓給付日期1~4月,15月抓給付日期5~8月,16月抓給付日期9~12月
      'm_StrSQL3 = m_StrSQL3 & " and substr(MB01,5,2) >='" & Trim(txt1(3)) & "' "
      m_StrSQL3 = m_StrSQL3 & " and 14+TRUNC((SUBSTR(MB01,5,2)-1)/4)>=" & Val(txt1(3))
   End If
   If txt1(4) <> "" Then '薪資月份迄
      m_StrSQL1 = m_StrSQL1 & " and substr(SM02,5,2) <='" & Trim(txt1(4)) & "' "
      'Modified by Morgan 2013/3/8 獎金年月已改為給付日期,14月抓給付日期1~4月,15月抓給付日期5~8月,16月抓給付日期9~12月
      'm_StrSQL3 = m_StrSQL3 & " and substr(MB01,5,2) <='" & Trim(txt1(4)) & "' "
      m_StrSQL3 = m_StrSQL3 & " and 14+TRUNC((SUBSTR(MB01,5,2)-1)/4) <=" & Val(txt1(4))
   End If
   If txt1(5) <> "" Then '員工編號起
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      m_StrSQL1 = m_StrSQL1 & " and substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) >='" & Trim(txt1(5)) & "' "
      m_StrSQL2 = m_StrSQL2 & " and substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4) >='" & Trim(txt1(5)) & "' "
      m_StrSQL3 = m_StrSQL3 & " and substr(MB02,1,2)||replace(substr(MB02,3,1),'A','0')||substr(MB02,4) >='" & Trim(txt1(5)) & "' "
   End If
   If txt1(6) <> "" Then '員工編號迄
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      m_StrSQL1 = m_StrSQL1 & " and substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) <='" & Trim(txt1(6)) & "' "
      m_StrSQL2 = m_StrSQL2 & " and substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4) <='" & Trim(txt1(6)) & "' "
      m_StrSQL3 = m_StrSQL3 & " and substr(MB02,1,2)||replace(substr(MB02,3,1),'A','0')||substr(MB02,4) <='" & Trim(txt1(6)) & "' "
   End If
   '每月薪資資料
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan2013/4/23 od01 改放年月日
   'modify by sonia 2015/12/24 扣繳表之加班費欄應扣除超時加班費SM28
   'Modify By Sindy 2020/6/25 + 證照津貼
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   m_str = "SELECT SM37 公司別, SM02-191100 年月,SM03 部門,SM01 員工編號,ST02 姓名, " & _
           "TO_CHAR(NVL(SM04,0)+NVL(SM05,0)+NVL(SM45,0)+NVL(SM28,0)-NVL(SM21,0),'9G999G999G999') 應稅薪資," & _
           "TO_CHAR(NVL(SM24,0),'9G999G999G999') 所得稅,TO_CHAR(NVL(SM25,0),'9G999G999G999') 年終基準月薪,TO_CHAR(SM27,'99') 工作天," & _
           "TO_CHAR(NVL(SM14,0),'9G999G999G999') 勞保費,TO_CHAR(NVL(SM15,0),'9G999G999G999') 健保費,TO_CHAR(NVL(OD03,0),'9G999G999G999') 其他給付," & _
           "TO_CHAR(NVL(SM12,0)-NVL(SM28,0),'9G999G999G999') 加班費,TO_CHAR(NVL(SM07,0),'9G999G999G999') 午餐津貼,substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4) NO " & _
           "From SALARYMONTH, STAFF, OtherPayData " & _
           "WHERE substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=ST01(+) AND SM02=substr(OD01(+),1,6) AND SM01=OD02(+) " & m_StrSQL1
   '年終獎金資料(前一年),未下月份條件或月份跨過13月者才抓年終資料
   If (txt1(3) = "" And txt1(4) = "") Or (Val(txt1(3)) <= 13 And Val(txt1(4)) >= 13) Then
   '2009/5/15 MODIFY BY SONIA 改UNION為UNION ALL否則89047之9714與9715相同6,000加總應為12,000否則會只有6,000
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'modify by sonia 2018/1/11 +YB26
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      m_str = m_str & " UNION ALL " & _
              "SELECT YB24 公司別, (YB01-1911)*100+13 年月,YB03 部門,YB02 員工編號,ST02 姓名, " & _
              "TO_CHAR((NVL(YB05,0)+NVL(YB06,0)+NVL(YB26,0)-NVL(YB15,0)),'9G999G999G999') 年終獎金," & _
              "TO_CHAR(NVL(YB17,0),'9G999G999G999') 所得稅,'','','','','','','',substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4) NO " & _
              "FROM YEARBONUS,STAFF WHERE substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4)=ST01(+) " & m_StrSQL2
   End If
   '每月獎金資料
   '2009/5/15 MODIFY BY SONIA 改UNION為UNION ALL否則89047之9714與9715相同6,000加總應為12,000否則會只有6,000
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2013/3/8 獎金年月已改為給付日期,14月抓給付日期1~4月,15月抓給付日期5~8月,16月抓給付日期9~12月
   'm_str = m_str & " UNION ALL " & _
           "SELECT comp 公司別, mb01-191100 年月,st03 部門,MB02 員工編號,ST02 姓名," & _
           "TO_CHAR(NVL(total,0),'9G999G999G999') 獎金總額," & _
           "TO_CHAR(NVL(tax,0),'9G999G999G999') 扣繳稅額,'','','','','','','',substr(MB02,1,1)||replace(substr(MB02,2),'A','0') NO from STAFF,(" & _
           "SELECT sd19 comp, MB01,MB02,NVL(MB03,0) total, NVL(MB04,0) tax " & _
           "FROM MonthBonus,salarydata WHERE substr(mb02,3,1)='0' AND MB02=Sd01(+) UNION " & _
           "SELECT sd28 comp, MB01,MB02,NVL(MB03,0) total, NVL(MB04,0) tax " & _
           "FROM MonthBonus,salarydata WHERE substr(mb02,3,1)='A' AND substr(MB02,1,1)||replace(substr(MB02,2),'A','0')=Sd01(+)) " & _
           "WHERE substr(MB02,1,1)||replace(substr(MB02,2),'A','0')=ST01(+) " & m_StrSQL3
   'Modified by Morgan 2013/5/7 第14,15,16月獎金要合併
   'm_str = m_str & " UNION ALL " & _
           "SELECT comp 公司別, (SUBSTR(MB01,1,4)*100-191100)+14+TRUNC((SUBSTR(MB01,5,2)-1)/4) 年月,st03 部門,MB02 員工編號,ST02 姓名," & _
           "TO_CHAR(NVL(total,0),'9G999G999G999') 獎金總額," & _
           "TO_CHAR(NVL(tax,0),'9G999G999G999') 扣繳稅額,'','','','','','','',substr(MB02,1,1)||replace(substr(MB02,2),'A','0') NO from STAFF,(" & _
           "SELECT sd19 comp, MB01,MB02,NVL(MB03,0) total, NVL(MB04,0) tax " & _
           "FROM MonthBonus,salarydata WHERE substr(mb02,3,1)='0' AND MB02=Sd01(+) UNION " & _
           "SELECT sd28 comp, MB01,MB02,NVL(MB03,0) total, NVL(MB04,0) tax " & _
           "FROM MonthBonus,salarydata WHERE substr(mb02,3,1)='A' AND substr(MB02,1,1)||replace(substr(MB02,2),'A','0')=Sd01(+)) " & _
           "WHERE substr(MB02,1,1)||replace(substr(MB02,2),'A','0')=ST01(+) " & m_StrSQL3
   '2014/2/19 modify by sonia 每月獎金之公司別改抓MB11,原抓SD19,sd28
   'Modified by Morgan 2023/12/25 +新部門st93
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   m_str = m_str & " UNION ALL " & _
           "SELECT comp 公司別,MB01  年月,decode(sign(mb01-" & (Left(新部門啟用日, 6) - 191100) & "),-1,st03,st93) 部門,MB02 員工編號,ST02 姓名," & _
           "TO_CHAR(NVL(sum(total),0),'9G999G999G999') 獎金總額," & _
           "TO_CHAR(NVL(sum(tax),0),'9G999G999G999') 扣繳稅額,'','','','','','','',substr(MB02,1,2)||replace(substr(MB02,3,1),'A','0')||substr(MB02,4) NO " & _
           "from STAFF,(" & _
           "SELECT MB11 comp, (SUBSTR(MB01,1,4)*100-191100)+14+TRUNC((SUBSTR(MB01,5,2)-1)/4) MB01,MB02,NVL(MB03,0) total, NVL(MB04,0) tax " & _
           "FROM MonthBonus,salarydata WHERE substr(mb02,3,1)='0' AND MB02=Sd01(+) " & m_StrSQL3 & _
           "UNION SELECT MB11 comp, (SUBSTR(MB01,1,4)*100-191100)+14+TRUNC((SUBSTR(MB01,5,2)-1)/4) MB01,MB02,NVL(MB03,0) total, NVL(MB04,0) tax " & _
           "FROM MonthBonus,salarydata WHERE substr(mb02,3,1)='A' AND substr(MB02,1,2)||replace(substr(MB02,3,1),'A','0')||substr(MB02,4)=Sd01(+)" & m_StrSQL3 & _
           ") WHERE substr(MB02,1,2)||replace(substr(MB02,3,1),'A','0')||substr(MB02,4)=ST01(+) group by comp,MB01,decode(sign(mb01-" & (Left(新部門啟用日, 6) - 191100) & "),-1,st03,st93),MB02,ST02"
           
   m_str = m_str & " ORDER BY 部門,NO,姓名,年月"
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       With m_rs
           m_rs.MoveFirst
           
           '預設值
           iLine = 1
           strType = "" '切頁條件
'           dblAmtT = 0
'           dblAmtS = 0
'           dblAmtO = 0
'           dblAmtT1 = 0
'           dblAmtT2 = 0
'           dblAmtS1 = 0
'           dblAmtS2 = 0
'           dblAmtA = 0
'           dblAmtL = 0
'           dblAmt01 = 0
'           dblAmt02 = 0
'           dblCntT = 0
'           dblCntT1 = 0
'           dblCntT2 = 0
'           dblOldAmt = 0
'
'           dblTAmtT = 0
'           dblTAmtS = 0
'           dblTAmtO = 0
'           dblTAmtT1 = 0
'           dblTAmtT2 = 0
'           dblTAmtS1 = 0
'           dblTAmtS2 = 0
'           dblTCntT = 0
'           dblTCntT1 = 0
'           dblTCntT2 = 0
'
           Do While Not m_rs.EOF
               
               For m_i = 1 To 12
                   strTemp(m_i) = ""
               Next m_i
               
               strTemp(1) = CheckStr(m_rs.Fields(3))   '員工編號
               strTemp(2) = CheckStr(m_rs.Fields(4))   '姓名
               strTemp(3) = CheckStr(m_rs.Fields(1))   '薪資年月
               strTemp(4) = CheckStr(m_rs.Fields(5))   '應稅薪資
               strTemp(5) = CheckStr(m_rs.Fields(6))   '所得稅
               strTemp(6) = CheckStr(m_rs.Fields(7))   '年終基準月薪
               strTemp(7) = CheckStr(m_rs.Fields(8))   '工作天
               strTemp(8) = CheckStr(m_rs.Fields(9))   '勞保費
               strTemp(9) = CheckStr(m_rs.Fields(10))  '健保費
               strTemp(10) = CheckStr(m_rs.Fields(11)) '其他給付
               strTemp(11) = CheckStr(m_rs.Fields(12)) '加班費
               strTemp(12) = CheckStr(m_rs.Fields(13)) '午餐津貼
               
               If iLine > 36 Or iLine = 1 Then
'                     (strType <> strTemp(7)) Then
'
'                   If (strType <> "" And strType <> strTemp(7)) Then
'                      PrintEnd '小計
'                   End If
'
'                   'If .AbsolutePosition <> .RecordCount Then
                       If iLine <> 1 Then Printer.NewPage
                       iLine = 1
                       PrintTitle '列印表頭
'                   'End If
               End If
               
               PrintDetail '列印表中
               
               strType = strTemp(1) '暫不跳頁
               
'               '小計 ********************************
'               dblAmtT = dblAmtT + strTemp(12)
'               dblAmtS = dblAmtS + strTemp(13)
'               dblAmtO = dblAmtO + strTemp(14)
'               dblCntT = dblCntT + 1
'               If strTemp(13) > 0 Then '所得稅>0
'                  dblAmtT1 = dblAmtT1 + strTemp(12) '應稅
'                  dblAmtS1 = dblAmtS1 + strTemp(13)
'                  dblCntT1 = dblCntT1 + 1
'               Else
'                  dblAmtT2 = dblAmtT2 + strTemp(12) '未稅
'                  dblAmtS2 = dblAmtS2 + strTemp(13)
'                  dblCntT2 = dblCntT2 + 1
'               End If
'               dblAmtA = dblAmtA + strTemp(15)
'               dblAmtL = dblAmtL + strTemp(16)
'               dblAmt01 = dblAmt01 + strTemp(17)
'               dblAmt02 = dblAmt02 + strTemp(18)
               
'               '合計 ********************************
'               dblTAmtT = dblTAmtT + strTemp(12)
'               dblTAmtS = dblTAmtS + strTemp(13)
'               dblTAmtO = dblTAmtO + strTemp(14)
'               dblTCntT = dblTCntT + 1
'               If strTemp(13) > 0 Then '所得稅>0
'                  dblTAmtT1 = dblTAmtT1 + strTemp(12) '應稅
'                  dblTAmtS1 = dblTAmtS1 + strTemp(13)
'                  dblTCntT1 = dblTCntT1 + 1
'               Else
'                  dblTAmtT2 = dblTAmtT2 + strTemp(12) '未稅
'                  dblTAmtS2 = dblTAmtS2 + strTemp(13)
'                  dblTCntT2 = dblTCntT2 + 1
'               End If
               
               m_rs.MoveNext
           Loop
           
'            '列印表尾
'            PrintEnd '小計
            
       End With
   Else
       MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
       Exit Sub
   End If
   
   Printer.EndDoc
   ShowPrintOk

End Sub

'Sub PrintEnd()
'   Printer.CurrentX = 500
'   Printer.CurrentY = iLine * 300
'   Printer.Print String(203, "-")
'
'   iLine = iLine + 1
'   Printer.CurrentX = 2000 - Printer.TextWidth(dblCntT & "人")
'   Printer.CurrentY = iLine * 300
'   Printer.Print dblCntT & "人"
'   Printer.CurrentX = PLeft(2)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "小　計："
'   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtT, "##,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmtT, "##,###")
'   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblAmtS, "##,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmtS, "##,###")
'   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmtO, "##,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmtO, "##,###")
'   iLine = iLine + 1
'   Printer.CurrentX = 2000 - Printer.TextWidth(dblCntT1 & "人")
'   Printer.CurrentY = iLine * 300
'   Printer.Print dblCntT1 & "人"
'   Printer.CurrentX = PLeft(2)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "應稅小計："
'   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtT1, "##,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmtT1, "##,###")
'   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblAmtS1, "##,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmtS1, "##,###")
'   iLine = iLine + 1
'   Printer.CurrentX = 2000 - Printer.TextWidth(dblCntT2 & "人")
'   Printer.CurrentY = iLine * 300
'   Printer.Print dblCntT2 & "人"
'   Printer.CurrentX = PLeft(2)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "未稅小計："
'   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtT2, "##,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmtT2, "##,###")
'   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblAmtS2, "##,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmtS2, "##,###")
'   iLine = iLine + 1
'   Printer.CurrentX = PLeft(2)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "加班費："
'   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtA, "##,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmtA, "##,###")
'   iLine = iLine + 1
'   Printer.CurrentX = PLeft(2)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "午餐津貼："
'   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtL, "##,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmtL, "##,###")
'   iLine = iLine + 1
'   Printer.CurrentX = PLeft(2)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "舊制應提退休金："
'   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(Round(Val(dblOldAmt) * 4 / 100, 0), "##,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(Round(Val(dblOldAmt) * 4 / 100, 0), "##,###")
'   iLine = iLine + 1
'   Printer.CurrentX = PLeft(2)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "新制應提退休金："
'   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmt01, "##,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmt01, "##,###")
'   iLine = iLine + 1
'   Printer.CurrentX = PLeft(2)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "新制員工自提退休金："
'   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmt02, "##,###"))
'   Printer.CurrentY = iLine * 300
'   Printer.Print Format(dblAmt02, "##,###")
'
'   dblAmtT = 0
'   dblAmtS = 0
'   dblAmtO = 0
'   dblAmtT1 = 0
'   dblAmtT2 = 0
'   dblAmtS1 = 0
'   dblAmtS2 = 0
'   dblAmtA = 0
'   dblAmtL = 0
'   dblAmt01 = 0
'   dblAmt02 = 0
'   dblCntT = 0
'   dblCntT1 = 0
'   dblCntT2 = 0
'   dblOldAmt = 0
'End Sub

Sub PrintTitle()

   GetPleft
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("員工應稅薪資檢核表") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "員工應稅薪資檢核表"
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
'   iLine = iLine + 2
'   Printer.CurrentX = PLeft(1)
'   Printer.CurrentY = iLine * 300
'   Printer.Print "公司別：" & strTemp(7) & "　" & strTemp(5)

   iLine = iLine + 2
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("年月")
   Printer.CurrentY = iLine * 300
   Printer.Print "薪資"
   Printer.CurrentX = PLeft(6) - Printer.TextWidth("基準月薪")
   Printer.CurrentY = iLine * 300
   Printer.Print "年終獎金"
   Printer.CurrentX = PLeft(7) - Printer.TextWidth("天數")
   Printer.CurrentY = iLine * 300
   Printer.Print "工作"
   
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "編號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("年月")
   Printer.CurrentY = iLine * 300
   Printer.Print "年月"
   Printer.CurrentX = PLeft(4) - Printer.TextWidth("應稅薪資")
   Printer.CurrentY = iLine * 300
   Printer.Print "應稅薪資"
   Printer.CurrentX = PLeft(5) - Printer.TextWidth("所得稅")
   Printer.CurrentY = iLine * 300
   Printer.Print "所得稅"
   Printer.CurrentX = PLeft(6) - Printer.TextWidth("基準月薪")
   Printer.CurrentY = iLine * 300
   Printer.Print "基準月薪"
   Printer.CurrentX = PLeft(7) - Printer.TextWidth("天數")
   Printer.CurrentY = iLine * 300
   Printer.Print "天數"
   Printer.CurrentX = PLeft(8) - Printer.TextWidth("勞保費")
   Printer.CurrentY = iLine * 300
   Printer.Print "勞保費"
   Printer.CurrentX = PLeft(9) - Printer.TextWidth("健保費")
   Printer.CurrentY = iLine * 300
   Printer.Print "健保費"
   Printer.CurrentX = PLeft(10) - Printer.TextWidth("其他給付")
   Printer.CurrentY = iLine * 300
   Printer.Print "其他給付"
   Printer.CurrentX = PLeft(11) - Printer.TextWidth("加班費")
   Printer.CurrentY = iLine * 300
   Printer.Print "加班費"
   Printer.CurrentX = PLeft(12) - Printer.TextWidth("午餐津貼")
   Printer.CurrentY = iLine * 300
   Printer.Print "午餐津貼"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(203, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 500
   PLeft(2) = 1400
   PLeft(3) = 3000
   PLeft(4) = 4500
   PLeft(5) = 6000
   PLeft(6) = 7500
   PLeft(7) = 8300
   PLeft(8) = 9980
   PLeft(9) = 11480
   PLeft(10) = 12980
   PLeft(11) = 14480
   PLeft(12) = 15980
End Sub

Sub PrintDetail()
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   'Modified by Morgan 2023/12/25
   'Printer.Print strTemp(2)
   PUB_PrintUnicodeText strTemp(2), Printer.CurrentX, Printer.CurrentY, 0
   'end 2023/12/25
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(3), "##/##"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(3), "##/##")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(strTemp(4), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(4), "##,###,###")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(5), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(5), "##,###,###")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(6), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "##,###,###")
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(strTemp(7), "##"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(7), "##")
   Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(strTemp(8), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(8), "##,###,###")
   Printer.CurrentX = PLeft(9) - Printer.TextWidth(Format(strTemp(9), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(9), "##,###,###")
   Printer.CurrentX = PLeft(10) - Printer.TextWidth(Format(strTemp(10), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(10), "##,###,###")
   Printer.CurrentX = PLeft(11) - Printer.TextWidth(Format(strTemp(11), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(11), "##,###,###")
   Printer.CurrentX = PLeft(12) - Printer.TextWidth(Format(strTemp(12), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(12), "##,###,###")
   
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
   Set frm170219 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 2, 3, 4
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 0, 1, 5, 6
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1
         If txt1(Index) <> "" Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
            End If
         End If
      Case 2
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "0101") = False Then
               Cancel = True
            End If
         End If
      Case 4
         If txt1(Index) <> "" Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
            End If
         End If
      Case 5, 6
         ' 判斷員工代號須為 6~9 或 F 開頭
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Cancel = True
            End If
         End If
         If Index = 5 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 6 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
            End If
         End If
      Case Else
   End Select
   
   If Cancel = True Then TextInverse txt1(Index)
      
End Sub
