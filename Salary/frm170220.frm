VERSION 5.00
Begin VB.Form frm170220 
   BorderStyle     =   1  '單線固定
   Caption         =   "薪資所得稅明細表"
   ClientHeight    =   3516
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4692
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3516
   ScaleWidth      =   4692
   Begin VB.CheckBox Check1 
      Caption         =   "含年終獎金"
      Height          =   408
      Left            =   3204
      TabIndex        =   17
      Top             =   1404
      Width           =   1344
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   4
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1470
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2070
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1470
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   6
      Left            =   2970
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   5
      Left            =   2070
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1800
      Width           =   765
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2070
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1110
      Width           =   435
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3660
      TabIndex        =   8
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   7
      Top             =   30
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   0
      TabIndex        =   10
      Top             =   2340
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   11
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   12
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   2070
      MaxLength       =   1
      TabIndex        =   0
      Top             =   780
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   1
      Left            =   2460
      MaxLength       =   1
      TabIndex        =   1
      Top             =   780
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "註：A公司舊制退休金自106/7起提列15%"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   3180
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2820
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line Line3 
      X1              =   2580
      X2              =   3240
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   3
      Left            =   1140
      TabIndex        =   15
      Top             =   1830
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "薪資年度："
      Height          =   180
      Left            =   1140
      TabIndex        =   14
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "薪資月份："
      Height          =   180
      Index           =   2
      Left            =   1140
      TabIndex        =   13
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "公  司  別："
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   9
      Top             =   810
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2190
      X2              =   2610
      Y1              =   900
      Y2              =   900
   End
End
Attribute VB_Name = "frm170220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by SINDY 2009/01/07
'2009/1/17 MODIFY BY SONIA 跨月加總且加入年終及每月獎金
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL1 As String    '每月薪資資料
Dim m_StrSQL2 As String    '年終獎金資料 2009/1/17 add by sonia
Dim m_StrSQL3 As String    '每月獎金資料 2009/1/17 add by sonia
Dim m_StrSQL4 As String    '同仁其他給付資料 Added by Morgan 2015/10/13
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
Dim strCompname As String    '2011/7/27 ADD BY SONIA

Private Sub cmdok_Click(Index As Integer)
Dim strYM As String

   Select Case Index
      Case 0
         If txt1(2) = "" Then
            MsgBox "薪資年度不可空白！", vbInformation, "操作錯誤！"
            txt1(2).SetFocus
            Exit Sub
         End If
         If txt1(3) = "" And txt1(4) = "" Then
            MsgBox "薪資月份不可空白！", vbInformation, "操作錯誤！"
            txt1(3).SetFocus
            Exit Sub
         End If
         
         Screen.MousePointer = vbHourglass
         StrMenu1
         Screen.MousePointer = vbDefault
      
      Case 1
         Unload Me
   End Select
End Sub

Sub StrMenu1()

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   
   '2009/1/17 modify by sonia 加入年終獎金(前一年13月)及每月獎金(當年14,15,16月)
   'm_StrSQL = ""
   'If txt1(0) <> "" Then '公司別起
   '   If m_StrSQL <> "" Then m_StrSQL = m_StrSQL & " and "
   '   m_StrSQL = m_StrSQL & " SM37 >='" & Trim(txt1(0)) & "' "
   'End If
   'If txt1(1) <> "" Then '公司別迄
   '   If m_StrSQL <> "" Then m_StrSQL = m_StrSQL & " and "
   '   m_StrSQL = m_StrSQL & " SM37 <='" & Trim(txt1(1)) & "' "
   'End If
   'If txt1(2) <> "" Then '薪資年度
   '   strYM = Left(ChangeTStringToWString(Trim(txt1(2)) & "0101"), 4)
   '   If m_StrSQL <> "" Then m_StrSQL = m_StrSQL & " and "
   '   m_StrSQL = m_StrSQL & " substr(SM02,1,4)='" & strYM & "' "
   'End If
   'If txt1(3) <> "" Then '薪資月份起
   '   If m_StrSQL <> "" Then m_StrSQL = m_StrSQL & " and "
   '   m_StrSQL = m_StrSQL & " substr(SM02,5,2) >='" & Trim(txt1(3)) & "' "
   'End If
   'If txt1(4) <> "" Then '薪資月份迄
   '   If m_StrSQL <> "" Then m_StrSQL = m_StrSQL & " and "
   '   m_StrSQL = m_StrSQL & " substr(SM02,5,2) <='" & Trim(txt1(4)) & "' "
   'End If
   'If txt1(5) <> "" Then '員工編號起
   '   If m_StrSQL <> "" Then m_StrSQL = m_StrSQL & " and "
   '   m_StrSQL = m_StrSQL & " replace(SM01,'A','0') >='" & Trim(txt1(5)) & "' "
   'End If
   'If txt1(6) <> "" Then '員工編號迄
   '   If m_StrSQL <> "" Then m_StrSQL = m_StrSQL & " and "
   '   m_StrSQL = m_StrSQL & " replace(SM01,'A','0') <='" & Trim(txt1(6)) & "' "
   'End If
   
   '2009/1/15 modify by sonia 應稅薪資再扣除勞退自提
   'm_str = "SELECT nvl(SD16,''),nvl(ST13,''),ST02,ST03,a0802,T.* " & _
            "FROM acc080,Staff,SalaryData, " & _
            "(SELECT SM01,SM37,SUM(nvl(SM04,0)) as T04,SUM(nvl(SM05,0)) as T05,SUM(nvl(SM21,0)) as T21,SUM(nvl(SM28,0)) as T28, " & _
            "SUM((nvl(SM04,0)+nvl(SM05,0)-nvl(SM21,0)+nvl(SM28,0)-nvl(SM16,0))) as TAmt, " & _
            "SUM(nvl(SM24,0)) as T24,SUM(nvl(OD03,0)) as TOD3, " & _
            "SUM(nvl(SM12,0)) as T12,SUM(nvl(SM07,0)) as T07, " & _
            "SUM(nvl(SM30,0)) as T30,SUM(nvl(SM16,0)) as T16 " & _
            "From SalaryMonth, OtherPayData " & _
            "WHERE " & m_StrSQL & _
            "AND SM02=OD01(+) " & _
            "AND SM01=OD02(+) " & _
            "group by SM01,SM37) T " & _
            "WHERE T.SM37=a0801(+) " & _
            "AND replace(T.SM01,'A','0')=ST01(+) " & _
            "AND replace(T.SM01,'A','0')=SD01(+) " & _
            "order by T.SM37,ST03,T.SM01 "
   
   m_StrSQL1 = "": m_StrSQL2 = "": m_StrSQL3 = "": m_StrSQL4 = ""
   If txt1(0) <> "" Then '公司別起
      m_StrSQL1 = m_StrSQL1 & " and sm37 >='" & Trim(txt1(0)) & "' "
      m_StrSQL2 = m_StrSQL2 & " and yb24 >='" & Trim(txt1(0)) & "' "
      'Modified by Morgan 2013/5/7 改語法
      'm_StrSQL3 = m_StrSQL3 & " and COMP >='" & Trim(txt1(0)) & "' "
      '2014/2/19 modify by sonia 每月獎金之公司別改抓MB11,原抓sd19
      m_StrSQL3 = m_StrSQL3 & " and MB11 >='" & Trim(txt1(0)) & "' "
   End If
   If txt1(1) <> "" Then '公司別迄
      m_StrSQL1 = m_StrSQL1 & " and SM37 <='" & Trim(txt1(1)) & "' "
      m_StrSQL2 = m_StrSQL2 & " and yb24 <='" & Trim(txt1(1)) & "' "
      'Modified by Morgan 2013/5/7 改語法
      'm_StrSQL3 = m_StrSQL3 & " and COMP <='" & Trim(txt1(1)) & "' "
      '2014/2/19 modify by sonia 每月獎金之公司別改抓MB11,原抓sd19
      m_StrSQL3 = m_StrSQL3 & " and MB11 <='" & Trim(txt1(1)) & "' "
   End If
   If txt1(2) <> "" Then '薪資年度, 年終獎金資料要抓前一年
      m_StrSQL1 = m_StrSQL1 & " and substr(SM02,1,4)='" & Val(txt1(2)) + 1911 & "' "
      m_StrSQL2 = m_StrSQL2 & " and yb01=" & Val(txt1(2)) + 1911 - 1
      m_StrSQL3 = m_StrSQL3 & " and substr(MB01,1,4)='" & Val(txt1(2)) + 1911 & "' "
      m_StrSQL4 = m_StrSQL4 & " and substr(OD01,1,4)>='" & Val(txt1(2)) + 1911 & "' " 'Added by Morgan 2015/10/13
   End If
   If txt1(3) <> "" Then '薪資月份起
      m_StrSQL1 = m_StrSQL1 & " and substr(SM02,5,2) >=" & Trim(txt1(3)) & " "
      'Modified by Morgan 2013/3/8 獎金年月已改為給付日期,14月抓給付日期1~4月,15月抓給付日期5~8月,16月抓給付日期9~12月
      'm_StrSQL3 = m_StrSQL3 & " and substr(MB01,5,2) >='" & Trim(txt1(3)) & "' "
      m_StrSQL3 = m_StrSQL3 & " and 14+TRUNC((SUBSTR(MB01,5,2)-1)/4) >=" & Val(txt1(3))
      m_StrSQL4 = m_StrSQL4 & " and substr(OD01,5,2)>='" & Trim(txt1(3)) & "' " 'Added by Morgan 2015/10/13
   End If
   If txt1(4) <> "" Then '薪資月份迄
      m_StrSQL1 = m_StrSQL1 & " and substr(SM02,5,2) <=" & Trim(txt1(4)) & " "
      'Modified by Morgan 2013/3/8 獎金年月已改為給付日期,14月抓給付日期1~4月,15月抓給付日期5~8月,16月抓給付日期9~12月
      'm_StrSQL3 = m_StrSQL3 & " and substr(MB01,5,2) <='" & Trim(txt1(4)) & "' "
      m_StrSQL3 = m_StrSQL3 & " and 14+TRUNC((SUBSTR(MB01,5,2)-1)/4) <=" & Val(txt1(4))
      m_StrSQL4 = m_StrSQL4 & " and substr(OD01,5,2)<='" & Trim(txt1(4)) & "' " 'Added by Morgan 2015/10/13
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
   '2009/1/20 MODIFY BY SONIA 婧瑄又說不扣勞退自提
   '2010/8/20 MODIFY BY SONIA 婧瑄加印SD16舊制符號◎
   'Modified by Morgan 2023/12/26 st03-->sm03
   m_str = "SELECT SUM(舊制公司提撥),ST02,SM03,A0802,員工編號,公司別,SUM(應稅薪資),SUM(所得稅),SUM(其他給付),SUM(加班費)," & _
           "SUM(午餐津貼),SUM(勞退公司提撥),SUM(勞退自提),DECODE(SD16,'','◎','') 舊制 FROM ( "
   '每月薪資資料(非F編號,19850101<=到職日<20050701,不管新舊制,公司舊制提撥都要提)
   '2009/3/20婧瑄說改19850101為19860101,每月薪資資料(非F編號,19860101<=到職日<20050701,不管新舊制,公司舊制提撥都要提)
   'Modify By Sindy 2022/4/21                        修改，上述條件剔除法律所同仁
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2015/10/14 同仁其他給付資料一個月可能多筆
   'm_str = m_str & "SELECT SUM(nvl(SM04,0)+nvl(SM05,0)-nvl(SM21,0)+nvl(SM28,0)) 舊制公司提撥,ST02,SM03,a0802,SM01 員工編號,SM37 公司別," & _
           "SUM((nvl(SM04,0)+nvl(SM05,0)-nvl(SM21,0)+nvl(SM28,0))) as 應稅薪資,SUM(nvl(SM24,0)) as 所得稅," & _
           "SUM(nvl(OD03,0)) as 其他給付, SUM(nvl(SM12,0)) as 加班費,SUM(nvl(SM07,0)) as 午餐津貼," & _
           "SUM(nvl(SM30,0)) as 勞退公司提撥, SUM(nvl(SM16,0)) as 勞退自提,SD16 " & _
           "From acc080,Staff,SalaryData,SALARYMONTH, OtherPayData " & _
           "WHERE SM02=substr(od01(+),1,6) AND SM01=OD02(+) " & m_StrSQL1 & "AND SM37=a0801(+) AND substr(sm01,1,1)||replace(substr(sm01,2),'A','0')=ST01(+) AND NVL(ST13,0)>=19860101 AND NVL(ST13,0)<20050701 AND SM01<'F' " & _
           "AND substr(sm01,1,1)||replace(substr(sm01,2),'A','0')=SD01(+) group by SD16,ST02,SM03,a0802,SM01,SM37,SD16"
   'modify by sonia 2015/12/24 扣繳表之加班費欄應扣除超時加班費SM28
   'Modify By Sindy 2020/6/25 + 證照津貼
   'Modified by Morgan 2023/12/26 st03-->sm03
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   'Modified by Morgan 2025/5/21 114/6薪資(7月初計算)開始暫停提撥舊制公司提撥(台一投資A公司不動) +*decode(SM37,'A',1,decode(sign(sm02-202505),1,0,1))
   m_str = m_str & "SELECT SUM((nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)-nvl(SM21,0)+nvl(SM28,0))*decode(SM37,'A',1,decode(sign(sm02-202505),1,0,1))) 舊制公司提撥,ST02,SM03,a0802,SM01 員工編號,SM37 公司別," & _
           "SUM((nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)-nvl(SM21,0)+nvl(SM28,0))) as 應稅薪資,SUM(nvl(SM24,0)) as 所得稅," & _
           "SUM(nvl(OD03,0)) as 其他給付, SUM(nvl(SM12,0)-NVL(SM28,0)) as 加班費,SUM(nvl(SM07,0)) as 午餐津貼," & _
           "SUM(nvl(SM30,0)) as 勞退公司提撥, SUM(nvl(SM16,0)) as 勞退自提,SD16 " & _
           "From acc080,Staff,SalaryData,SALARYMONTH,(select substr(od01,1,6) X1, od02, sum(od03) od03 from OtherPayData where 1=1 " & m_StrSQL4 & " group by substr(od01,1,6),od02) X " & _
           "WHERE SM02=X1(+) AND SM01=OD02(+) " & m_StrSQL1 & "AND SM37=a0801(+) AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=ST01(+) AND NVL(ST13,0)>=19860101 AND NVL(ST13,0)<20050701 AND SM01<'F' AND SD19<>'L' " & _
           "AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=SD01(+) group by SD16,ST02,SM03,a0802,SM01,SM37,SD16"
   'end 2015/10/14
   '每月薪資資料(F編號,到職日<19850101,到職日>=20050701,公司舊制提撥都不提)
   '2009/3/20婧瑄說改19850101為19860101,每月薪資資料(F編號,到職日<19860101,到職日>=20050701,公司舊制提撥都不提)
   'Modify By Sindy 2022/4/21                        修改，上述條件剔除法律所同仁
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2015/10/14 同仁其他給付資料一個月可能多筆
   'm_str = m_str & " UNION ALL SELECT 0 舊制公司提撥,ST02,SM03,a0802,SM01 員工編號,SM37 公司別," & _
           "SUM((nvl(SM04,0)+nvl(SM05,0)-nvl(SM21,0)+nvl(SM28,0))) as 應稅薪資,SUM(nvl(SM24,0)) as 所得稅," & _
           "SUM(nvl(OD03,0)) as 其他給付, SUM(nvl(SM12,0)) as 加班費,SUM(nvl(SM07,0)) as 午餐津貼," & _
           "SUM(nvl(SM30,0)) as 勞退公司提撥, SUM(nvl(SM16,0)) as 勞退自提,SD16 " & _
           "From acc080,Staff,SalaryData,SALARYMONTH, OtherPayData " & _
           "WHERE SM02=substr(od01(+),1,6) AND SM01=OD02(+) " & m_StrSQL1 & "AND SM37=a0801(+) AND substr(sm01,1,1)||replace(substr(sm01,2),'A','0')=ST01(+) AND (NVL(ST13,0)<19860101 OR NVL(ST13,0)>=20050701 OR SM01>'F') " & _
           "AND substr(sm01,1,1)||replace(substr(sm01,2),'A','0')=SD01(+) group by nvl(SD16,''),nvl(ST13,''),ST02,SM03,a0802,SM01,SM37,SD16 "
   'modify by sonia 2015/12/24 扣繳表之加班費欄應扣除超時加班費SM28
   'Modify By Sindy 2020/6/25 + 證照津貼
   'Modified by Morgan 2023/12/26 st03-->sm03
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   m_str = m_str & " UNION ALL SELECT 0 舊制公司提撥,ST02,SM03,a0802,SM01 員工編號,SM37 公司別," & _
           "SUM((nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)-nvl(SM21,0)+nvl(SM28,0))) as 應稅薪資,SUM(nvl(SM24,0)) as 所得稅," & _
           "SUM(nvl(OD03,0)) as 其他給付, SUM(nvl(SM12,0)-NVL(SM28,0)) as 加班費,SUM(nvl(SM07,0)) as 午餐津貼," & _
           "SUM(nvl(SM30,0)) as 勞退公司提撥, SUM(nvl(SM16,0)) as 勞退自提,SD16 " & _
           "From acc080,Staff,SalaryData,SALARYMONTH,(select substr(od01,1,6) X1, od02, sum(od03) od03 from OtherPayData where 1=1 " & m_StrSQL4 & " group by substr(od01,1,6),od02) X " & _
           "WHERE SM02=X1(+) AND SM01=OD02(+) " & m_StrSQL1 & "AND SM37=a0801(+) AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=ST01(+) AND (NVL(ST13,0)<19860101 OR NVL(ST13,0)>=20050701 OR SM01>'F' OR SD19='L') " & _
           "AND substr(sm01,1,2)||replace(substr(sm01,3,1),'A','0')||substr(sm01,4)=SD01(+) group by nvl(SD16,''),nvl(ST13,''),ST02,SM03,a0802,SM01,SM37,SD16 "
   'end 2015/10/14
   '年終獎金資料(前一年),未下月份條件或月份跨過13月者才抓年終資料
   'Modified by Morgan 2018/3/5 +年終獎金選項--辜
   If Check1.Value = vbChecked And ((txt1(3) = "" And txt1(4) = "") Or (Val(txt1(3)) <= 13 And Val(txt1(4)) >= 13)) Then
      '2009/5/15 MODIFY BY SONIA 改UNION為UNION ALL否則89047之9714與9715相同6,000加總應為12,000否則會只有6,000
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'modify by sonia 2018/1/11 +YB26
      'Modified by Morgan 2023/12/26 st03-->YB03 SM03
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      m_str = m_str & " UNION ALL " & _
           "select 0 舊制公司提撥,ST02,YB03 SM03,a0802,YB02 員工編號,YB24 公司別," & _
           "NVL(YB05,0)+NVL(YB06,0)+NVL(YB26,0)-NVL(YB15,0) 應稅薪資,NVL(YB17,0) 所得稅,0 其他給付,0 加班費,0 午餐津貼,0 勞退公司提撥,0 勞退自提,SD16 " & _
           "FROM acc080,Staff,SalaryData,YEARBONUS WHERE YB24=a0801(+) " & _
           "AND substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4)=ST01(+) AND substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4)=SD01(+) " & m_StrSQL2
   End If
   
   '每月獎金資料
   '2009/5/15 MODIFY BY SONIA 改UNION為UNION ALL否則89047之9714與9715相同6,000加總應為12,000否則會只有6,000
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'Modified by Morgan 2013/5/7 第14,15,16月獎金要合併(可以不必合併,因為最後全部要加總)
   'm_str = m_str & " UNION ALL select 0 舊制公司提撥,ST02,ST03,a0802,MB02 員工編號,comp 公司別," & _
           "NVL(total,0) 應稅薪資,NVL(tax,0) 所得稅,0 其他給付,0 加班費,0 午餐津貼,0 勞退公司提撥,0 勞退自提,SD16 from STAFF,ACC080,(" & _
           "SELECT sd19 comp, MB01,MB02,NVL(MB03,0) total, NVL(MB04,0) tax,SD16 FROM MonthBonus,salarydata " & _
           "WHERE substr(mb02,3,1)='0' AND MB02=Sd01(+) UNION " & _
           "SELECT sd28 comp, MB01,MB02,NVL(MB03,0) total, NVL(MB04,0) tax,SD16 FROM MonthBonus,salarydata " & _
           "WHERE substr(mb02,3,1)='A' AND substr(MB02,1,1)||replace(substr(MB02,2),'A','0')=Sd01(+) ) " & _
           "WHERE COMP=A0801(+) AND substr(MB02,1,1)||replace(substr(MB02,2),'A','0')=ST01(+) " & m_StrSQL3
   ''2014/2/19 modify by sonia 每月獎金之公司別改抓MB11,原抓sd19,sd28
   'Modified by Morgan 2023/12/26 +新部門判斷
   'm_str = m_str & " UNION ALL select 0 舊制公司提撥,ST02,ST03,a0802,MB02 員工編號,comp 公司別," & _
           "NVL(sum(total),0) 應稅薪資,NVL(sum(tax),0) 所得稅,0 其他給付,0 加班費,0 午餐津貼,0 勞退公司提撥,0 勞退自提,SD16" & _
           " from STAFF,ACC080,(" & _
           "SELECT MB11 comp,TO_CHAR(14+TRUNC((SUBSTR(MB01,5,2)-1)/4),'99') MB01,MB02,NVL(MB03,0) total, NVL(MB04,0) tax,SD16 FROM MonthBonus,salarydata " & _
           "WHERE substr(mb02,3,1)='0' AND MB02=Sd01(+) " & m_StrSQL3 & _
           " UNION ALL SELECT MB11 comp,TO_CHAR(14+TRUNC((SUBSTR(MB01,5,2)-1)/4),'99') MB01,MB02,NVL(MB03,0) total, NVL(MB04,0) tax,SD16 FROM MonthBonus,salarydata " & _
           "WHERE substr(mb02,3,1)='A' AND substr(MB02,1,1)||replace(substr(MB02,2),'A','0')=Sd01(+) " & m_StrSQL3 & _
           " ) WHERE COMP=A0801(+) AND substr(MB02,1,1)||replace(substr(MB02,2),'A','0')=ST01(+) " & _
           " group by st02,st03,a0802,mb02,comp,sd16"
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   m_str = m_str & " UNION ALL select 0 舊制公司提撥,ST02,sm03,a0802,MB02 員工編號,comp 公司別," & _
           "NVL(sum(total),0) 應稅薪資,NVL(sum(tax),0) 所得稅,0 其他給付,0 加班費,0 午餐津貼,0 勞退公司提撥,0 勞退自提,SD16" & _
           " from STAFF,ACC080,(" & _
           "SELECT MB11 comp,decode(sign(mb01-" & (Left(新部門啟用日, 6) - 191100) & "),-1,st03,st93) SM03,MB02,NVL(MB03,0) total, NVL(MB04,0) tax,SD16 FROM MonthBonus,salarydata,staff " & _
           "WHERE substr(mb02,3,1)='0' AND MB02=Sd01(+) and st01(+)=sd01 " & m_StrSQL3 & _
           " UNION ALL SELECT MB11 comp,decode(sign(mb01-" & (Left(新部門啟用日, 6) - 191100) & "),-1,st03,st93) SM03,MB02,NVL(MB03,0) total, NVL(MB04,0) tax,SD16 FROM MonthBonus,salarydata,staff " & _
           "WHERE substr(mb02,3,1)='A' AND substr(MB02,1,2)||replace(substr(MB02,3,1),'A','0')||substr(MB02,4)=Sd01(+) and st01(+)=sd01" & m_StrSQL3 & _
           " ) WHERE COMP=A0801(+) AND substr(MB02,1,2)||replace(substr(MB02,3,1),'A','0')||substr(MB02,4)=ST01(+) " & _
           " group by st02,sm03,a0802,mb02,comp,sd16"
   'Modified by Morgan 2023/12/26 st03-->sm03
   m_str = m_str & ") GROUP BY ST02,SM03,a0802,員工編號,公司別,SD16 ORDER BY 公司別,SM03,員工編號"
   '2009/1/17 END
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         '預設值
         iLine = 1
         strType = "" '切頁條件
         dblAmtT = 0
         dblAmtS = 0
         dblAmtO = 0
         dblAmtT1 = 0
         dblAmtT2 = 0
         dblAmtS1 = 0
         dblAmtS2 = 0
         dblAmtA = 0
         dblAmtL = 0
         dblAmt01 = 0
         dblAmt02 = 0
         dblCntT = 0
         dblCntT1 = 0
         dblCntT2 = 0
         dblOldAmt = 0
         
         dblTAmtT = 0
         dblTAmtS = 0
         dblTAmtO = 0
         dblTAmtT1 = 0
         dblTAmtT2 = 0
         dblTAmtS1 = 0
         dblTAmtS2 = 0
         dblTCntT = 0
         dblTCntT1 = 0
         dblTCntT2 = 0
         
         Do While Not m_rs.EOF
             
            For m_i = 1 To 20
                strTemp(m_i) = ""
            Next m_i
               
   '01.SUM((nvl(SM04,0)+nvl(SM05,0)-nvl(SM21,0)+nvl(SM28,0))) as 舊制公司提撥,
   '02.ST02,
   '03.ST03,
   '04.a0802,
   '05.SM01,
   '06.SM37,
   '07.SUM((nvl(SM04,0)+nvl(SM05,0)-nvl(SM21,0)+nvl(SM28,0))) as TAmt,
   '08.SUM(nvl(SM24,0)) as T24,
   '09.SUM(nvl(OD03,0)) as TOD3,
   '10.SUM(nvl(SM12,0)) as T12,
   '11.SUM(nvl(SM07,0)) as T07,
   '12.SUM(nvl(SM30,0)) as T30,
   '13.SUM(nvl(SM16,0)) as T16
            strTemp(1) = CheckStr(m_rs.Fields(0)) '舊制公司提撥
            strTemp(2) = CheckStr(m_rs.Fields(1)) '姓名
            strTemp(3) = CheckStr(m_rs.Fields(2)) '部門代號
            strTemp(4) = CheckStr(m_rs.Fields(3)) '公司名稱
            strTemp(5) = CheckStr(m_rs.Fields(4)) '編號
            strTemp(6) = CheckStr(m_rs.Fields(5)) '公司別
            strTemp(7) = CheckStr(m_rs.Fields(6)) '應稅薪資
            strTemp(8) = CheckStr(m_rs.Fields(7)) '所得稅
            strTemp(9) = CheckStr(m_rs.Fields(8)) '其他給付
            strTemp(10) = CheckStr(m_rs.Fields(9)) '加班費
            strTemp(11) = CheckStr(m_rs.Fields(10)) '午餐津貼
            strTemp(12) = CheckStr(m_rs.Fields(11)) '新制公司應提退休金
            strTemp(13) = CheckStr(m_rs.Fields(12)) '新制員工自提退休金
            strTemp(14) = CheckStr(m_rs.Fields(13)) '舊制員工    2010/8/20 ADD BY SONIA
               
            If iLine > 50 Or iLine = 1 Or _
               (strType <> strTemp(6)) Then
               
               If (strType <> "" And strType <> strTemp(6)) Then
                  PrintEnd '小計
               End If
               
               'If .AbsolutePosition <> .RecordCount Then
                   If strType <> "" Then Printer.NewPage
                   iLine = 1
                   strType = strTemp(6)     '2011/7/27 ADD BY SONIA
                   strCompname = strTemp(4) '2011/7/27 ADD BY SONIA
                   PrintTitle '列印表頭
               'End If
            End If
               
          PrintDetail '列印表中
          
          strType = strTemp(6) '依公司別跳頁
          strCompname = strTemp(4) '2011/7/27 ADD BY SONIA
          
          '小計 ********************************
          dblAmtT = dblAmtT + strTemp(7)
          dblAmtS = dblAmtS + strTemp(8)
          dblAmtO = dblAmtO + strTemp(9)
          dblCntT = dblCntT + 1
          If strTemp(8) > 0 Then '所得稅>0
             dblAmtT1 = dblAmtT1 + strTemp(7) '應稅
             dblAmtS1 = dblAmtS1 + strTemp(8)
             dblCntT1 = dblCntT1 + 1
          Else
             dblAmtT2 = dblAmtT2 + strTemp(7) '未稅
             dblAmtS2 = dblAmtS2 + strTemp(8)
             dblCntT2 = dblCntT2 + 1
          End If
          dblAmtA = dblAmtA + strTemp(10)
          dblAmtL = dblAmtL + strTemp(11)
          dblAmt01 = dblAmt01 + strTemp(12)
          dblAmt02 = dblAmt02 + strTemp(13)
          '舊制為94.7.1以前入所且適用勞退新制SD16=NULL的應稅薪資*4/100
          '2009/1/19 MODIFY BY SONIA
          'If strTemp(2) < "20050701" And strTemp(1) = "" Then
          If strTemp(1) > 0 Then
             dblOldAmt = dblOldAmt + strTemp(1)
          End If
          
          '合計 ********************************
          dblTAmtT = dblTAmtT + strTemp(7)
          dblTAmtS = dblTAmtS + strTemp(8)
          dblTAmtO = dblTAmtO + strTemp(9)
          dblTCntT = dblTCntT + 1
          If strTemp(8) > 0 Then '所得稅>0
             dblTAmtT1 = dblTAmtT1 + strTemp(7) '應稅
             dblTAmtS1 = dblTAmtS1 + strTemp(8)
             dblTCntT1 = dblTCntT1 + 1
          Else
             dblTAmtT2 = dblTAmtT2 + strTemp(7) '未稅
             dblTAmtS2 = dblTAmtS2 + strTemp(8)
             dblTCntT2 = dblTCntT2 + 1
          End If
          
          m_rs.MoveNext
      Loop
         
          '列印表尾
          PrintEnd '小計
          
          iLine = iLine + 1
            Printer.CurrentX = 500
            Printer.CurrentY = iLine * 300
            Printer.Print String(140, "-")
            
            iLine = iLine + 1
            Printer.CurrentX = 2000 - Printer.TextWidth(dblTCntT & "人")
            Printer.CurrentY = iLine * 300
            Printer.Print dblTCntT & "人"
            Printer.CurrentX = PLeft(2)
            Printer.CurrentY = iLine * 300
            Printer.Print "合　計："
            Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblTAmtT, "##,###,###"))
            Printer.CurrentY = iLine * 300
            Printer.Print Format(dblTAmtT, "##,###,###")
            Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblTAmtS, "##,###,###"))
            Printer.CurrentY = iLine * 300
            Printer.Print Format(dblTAmtS, "##,###,###")
            Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblTAmtO, "##,###,###"))
            Printer.CurrentY = iLine * 300
            Printer.Print Format(dblTAmtO, "##,###,###")
            iLine = iLine + 1
            Printer.CurrentX = 2000 - Printer.TextWidth(dblTCntT1 & "人")
            Printer.CurrentY = iLine * 300
            Printer.Print dblTCntT1 & "人"
            Printer.CurrentX = PLeft(2)
            Printer.CurrentY = iLine * 300
            Printer.Print "應稅合計："
            Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblTAmtT1, "##,###,###"))
            Printer.CurrentY = iLine * 300
            Printer.Print Format(dblTAmtT1, "##,###,###")
            Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblTAmtS1, "##,###,###"))
            Printer.CurrentY = iLine * 300
            Printer.Print Format(dblTAmtS1, "##,###,###")
            iLine = iLine + 1
            Printer.CurrentX = 2000 - Printer.TextWidth(dblTCntT2 & "人")
            Printer.CurrentY = iLine * 300
            Printer.Print dblTCntT2 & "人"
            Printer.CurrentX = PLeft(2)
            Printer.CurrentY = iLine * 300
            Printer.Print "未稅合計："
            Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblTAmtT2, "##,###,###"))
            Printer.CurrentY = iLine * 300
            Printer.Print Format(dblTAmtT2, "##,###,###")
            Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblTAmtS2, "##,###,###"))
            Printer.CurrentY = iLine * 300
            Printer.Print Format(dblTAmtS2, "##,###,###")
       End With
   Else
       MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
       Exit Sub
   End If
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub PrintEnd()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   '2009/4/9 add by sonia
   If iLine + 9 > 50 Then
      If strType <> "" Then Printer.NewPage
      iLine = 1
      PrintTitle '列印表頭
   End If
   '2009/4/9 end
   
   If iLine <> 9 Then iLine = iLine + 1
   Printer.CurrentX = 2000 - Printer.TextWidth(dblCntT & "人")
   Printer.CurrentY = iLine * 300
   Printer.Print dblCntT & "人"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "小　計："
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtT, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtT, "##,###,###")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblAmtS, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS, "##,###,###")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmtO, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtO, "##,###,###")
   iLine = iLine + 1
   Printer.CurrentX = 2000 - Printer.TextWidth(dblCntT1 & "人")
   Printer.CurrentY = iLine * 300
   Printer.Print dblCntT1 & "人"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "應稅小計："
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtT1, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtT1, "##,###,###")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblAmtS1, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS1, "##,###,###")
   iLine = iLine + 1
   Printer.CurrentX = 2000 - Printer.TextWidth(dblCntT2 & "人")
   Printer.CurrentY = iLine * 300
   Printer.Print dblCntT2 & "人"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "未稅小計："
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtT2, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtT2, "##,###,###")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblAmtS2, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS2, "##,###,###")
   iLine = iLine + 1
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "加班費："
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtA, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtA, "##,###,###")
   iLine = iLine + 1
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "午餐津貼："
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmtL, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtL, "##,###,###")
   'Modify By Sindy 2022/5/20 智權公司J公司、法律所L公司：請取消舊制應提退休金：(有 v 的應稅薪資總合 * 0.07)的那一行。
   If strType <> "L" And strType <> "J" Then
   '2022/5/20 END
      iLine = iLine + 1
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = iLine * 300
      Printer.Print "舊制應提退休金："
      'add by sonia 2017/8/3 2017/7起台一投資A公司改15%
      'modify by sonia 2018/3/15 107年會跑0.07
      'If strType = "A" And txt1(2) >= 106 And txt1(3) >= 7 Then
      If strType = "A" And ((txt1(2) = 106 And txt1(3) >= 7) Or txt1(2) >= 107) Then
         Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(Round(Val(dblOldAmt) * 15 / 100, 0), "##,###,###"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(Round(Val(dblOldAmt) * 15 / 100, 0), "##,###,###")
         Printer.CurrentX = PLeft(5) - Printer.TextWidth("(有 v 的應稅薪資總合 * 0.15)")
         Printer.CurrentY = iLine * 300
         Printer.Print "(有 v 的應稅薪資總合 * 0.15)"
      Else
         'Modify By Sindy 2022/5/20 智慧所1公司：退休金提撥率由7%調降為2%; 台一投資A公司：仍維持15%
         'Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(Round(Val(dblOldAmt) * 7 / 100, 0), "##,###,###"))
         Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(Round(Val(dblOldAmt) * 2 / 100, 0), "##,###,###"))
         Printer.CurrentY = iLine * 300
         'Printer.Print Format(Round(Val(dblOldAmt) * 7 / 100, 0), "##,###,###")
         Printer.Print Format(Round(Val(dblOldAmt) * 2 / 100, 0), "##,###,###")
         '2009/1/20 ADD BY SONIA
         'Printer.CurrentX = PLeft(5) - Printer.TextWidth("(有 v 的應稅薪資總合 * 0.07)")
         Printer.CurrentX = PLeft(5) - Printer.TextWidth("(有 v 的應稅薪資總合 * 0.02)")
         Printer.CurrentY = iLine * 300
         'Printer.Print "(有 v 的應稅薪資總合 * 0.07)"
         Printer.Print "(有 v 的應稅薪資總合 * 0.02)"
         '2009/1/20 END
      End If
   End If
   iLine = iLine + 1
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "新制應提退休金："
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmt01, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt01, "##,###,###")
   iLine = iLine + 1
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "新制員工自提退休金："
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmt02, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt02, "##,###,###")
   
   dblAmtT = 0
   dblAmtS = 0
   dblAmtO = 0
   dblAmtT1 = 0
   dblAmtT2 = 0
   dblAmtS1 = 0
   dblAmtS2 = 0
   dblAmtA = 0
   dblAmtL = 0
   dblAmt01 = 0
   dblAmt02 = 0
   dblCntT = 0
   dblCntT1 = 0
   dblCntT2 = 0
   dblOldAmt = 0
End Sub

Sub PrintTitle()

   GetPleft
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   '2009/1/17 modify by sonia
   'Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("薪資所得稅明細表") / 2)
   Printer.CurrentX = 4750
   Printer.CurrentY = iLine * 300
   Printer.Print "薪資所得稅明細表"
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "列印人：" & strUserName
   '2009/1/17 add by sonia
   Printer.CurrentX = 4500
   Printer.CurrentY = iLine * 300
   Printer.Print "薪資年月：" & txt1(2) & "/" & txt1(3) & "--" & txt1(2) & "/" & txt1(4)
   '2009/1/17 end
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 2
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   '2011/7/27 MODIFY BY SONIA 100/6的商標公司合計因印在單獨一頁但會印成專利
   'Printer.Print "公司別：" & strTemp(6) & "　" & strTemp(4)
   Printer.Print "公司別：" & strType & "　" & strCompname
   '2011/7/27 END
   '2010/8/20 add by sonia
   Printer.CurrentX = PLeft(3) + 1000
   Printer.CurrentY = iLine * 300
   Printer.Print "v為公司舊制退休金，◎為勞退舊制"
   '2010/8/20 end
   
   iLine = iLine + 2
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "編　號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("應稅薪資")
   Printer.CurrentY = iLine * 300
   Printer.Print "應稅薪資"
   Printer.CurrentX = PLeft(4) - Printer.TextWidth("所得稅")
   Printer.CurrentY = iLine * 300
   Printer.Print "所得稅"
   Printer.CurrentX = PLeft(5) - Printer.TextWidth("其他給付")
   Printer.CurrentY = iLine * 300
   Printer.Print "其他給付"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 1000
   PLeft(2) = 2500
   PLeft(3) = 6500
   PLeft(4) = 8500
   PLeft(5) = 10500
End Sub

Sub PrintDetail()
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   '2009/1/20 MODIFY BY SONIA 有提公司舊制退休金者加註v
   'Printer.Print strTemp(5)
   If strTemp(1) > 0 Then
      Printer.Print strTemp(5) & " v"
   Else
      Printer.Print strTemp(5)
   End If
   '2009/1/20 END
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2) & strTemp(14)
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(7), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(7), "##,###,###")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(strTemp(8), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(8), "##,###,###")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(9), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(9), "##,###,###")
   
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
   Set frm170220 = Nothing
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
      Case 0, 1
         If txt1(Index) = "" Then Exit Sub
         If Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 2
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "0101") = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
      Case 3, 4
         '2009/1/17 cancel by sonia 加年終及每月獎金故取消檢查
         'If txt1(Index) <> "" Then
         '   txt1(Index).Text = Right("00" & Trim(txt1(Index).Text), 2)
         '   If ChkDate("99" & txt1(Index) & "01") = False Then
         '       Call txt1_GotFocus(Index)
         '       Cancel = True
         '       Exit Sub
         '   End If
         'End If
         If Index = 4 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 5, 6
         ' 判斷員工代號須為 6~9 或 F 開頭
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 5 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 6 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
