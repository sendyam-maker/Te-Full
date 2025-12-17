Attribute VB_Name = "basPerson"
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Add By Sindy 2011/8/3 出缺勤簽核系統
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

'Add By Sindy 2017/8/21 因人事雙職代
Public Const intDutyItem = 20
Public PubABS001_1(1 To intDutyItem) As String '職代組數
Public PubABS001_A, PubABS001_B '雙職代的A,B區
'2017/8/21 END
Public Const 人事處 = "M21"
'Add By Sindy 2021/5/17 (PUB_ChkByPassWork)
Public Const 分流上班起始日期 = "20210517"
Public Const 分流上班截止日期 = "20211130"
Public intByPassArea As Integer
Public strByPassStarTime(1 To 8) As String
Public strByPassEndTime(1 To 8) As String
Public m_bolByPassWork As Boolean '特殊分流上班
'2021/5/17 END

'表單類別
Public Const 表單類別_請假 = "01"
Public Const 表單類別_加班 = "02"
Public Const 表單類別_出差 = "03"
Public Const 表單類別_出缺勤統計 = "04"
Public Const 表單類別_個人資料明細 = "05"

'表單狀態
Public Const 會簽職代 = "01"
Public Const 主管審核中 = "02"
Public Const 退回 = "03"
Public Const 送人事處簽收 = "04"
Public Const 已核准 = "05"
Public Const 註銷 = "06"
Public Const 重送 = "07"
Public Const 主管代填 = "08"

Public Const 重送更改通知 = "71"
Public Const 退回通知主管 = "72"

Public Const ST01CodeNum1 As String = "'6','7','8','9','A','B','C','D','E'"
Public Const B1002CName As String = "decode(B1002,'01','請假','02','加班','03','出差')"
Public Const B1014CName As String = "decode(B1014,'1','長程','2','短程','3','大陸','4','國外')"
Public Const B1018CName As String = "decode(B1018,'01','會簽職代','02','主管審核中','03','退回','04','送人事處簽收','05','已核准','06','註銷','07','重送','08','主管代填')"
Public Const B1206CName As String = "decode(B1206,'01','會簽職代','02','主管審核中','03','退回','04','送人事處簽收','05','已核准','06','註銷','07','重送','08','主管代填')"
Public Const B1102CName As String = "decode(B1102,'1','職代','2','審核主管')"
Public Const B1107CName As String = "decode(B1107,'1','同意','2','退回')"
'Add By Sindy 2023/12/27 人事部門名稱
Public Const A0925For1Code As String = "B,F,J,L,M,P,S,T,W,Y,R" '第1碼
Public Const A0925CName As String = "decode(substr(st93,1,1),'B','業務拓展部','F','專利國外部','J','專利日本部','L','法律所','M','管理部','P','專利國內部','S','智權部','T','商標部','W','顧問服務組','Y','創新業務部','R','台一投資',substr(st93,1,1))"
'2023/12/27 END

'Add By Sindy 2020/4/13
'Modify By Sindy 2020/5/8 + ,'09','108號8F(805室)'
Public Const SP03WorkPlace As String = "decode(SP03,'01','居家辦公','02','大都會'" & _
   ",'03','108號4F','04','108號5F'" & _
   ",'05','北所','06','中所'" & _
   ",'07','南所','08','高所','09','108號8F(805室)'" & _
   ",SP03)"

Public pub_CallNextABSForm As Boolean '是否呼叫下一個表單(出缺勤使用)
Public bolfrm180203ExitForm As Boolean
Public Const 接收打卡時間公告 = "尖峰時段刷卡後半小時即可查詢（固定 ９：１５ 發送異常通知）"
Public Const 中午休息時間改1200 = 20170502

'假別List,統計時數及人數:
'01   忘打卡
'02   遲到
'03   曠職
'04   出差
'05   事假
'06   病假
'07   公假
'08   特別假
'09   婚假
'10   產假
'11   流產假
'12   喪假
'13   公傷假
'14   補休
'15   其他
'16   加班
'17   扣年終產假
'18   扣年終流產假
'19   陪產假
'20   生理假
'21   產檢假
'22   家庭照顧假
'23   健檢假
'24   防疫照顧假
'25   天災不給薪
Public dblHour(25) As Double, dblCnt(25) As Double 'Add By Sindy 2014/12/31
Public g_strB1002(10) As String, g_strB1008(10) As String, g_strDay(10) As String 'Modify By Sindy 2023/9/15


'Modify by Sindy 2009/01/23
'*************************************************
' 計算年資
' Input : oST01                  ==> 員工編號
' Input : strCntEndDate      ==> 計算年資的截止日期 (西元日期 ex.20081231)
' Return : 年資
'*************************************************
Public Function CalYear(oST01 As String, strCntEndDate As String) As Double
Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_str2 As String
Dim m_i As Integer
Dim int_i As Integer
Dim strStarDate As String
Dim strEndDate As String
Dim LngReturnDay As Long
Dim intReturnYear As Integer
Dim m_month As Integer
Dim m_Year1_Std As String     '員工入所當年的第一個工作天
Dim strBackTaieDate As String 'Add By Sindy 2019/6/25

LngReturnDay = 0
intReturnYear = 0

m_str = "SELECT ST13 " & _
             "FROM staff " & _
             "WHERE st01='" & oST01 & "' and ST13 is not null "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        
        Do While Not m_rs.EOF
            
            '為固定任職期間寫死的條件判斷
            If oST01 = "63001" Or oST01 = "63002" Or _
               oST01 = "64001" Or oST01 = "65001" Or _
               oST01 = "68001" Or oST01 = "72010" Then
               If oST01 = "63001" Then
                  strStarDate = "19740222": strEndDate = "19760930"
               End If
               If oST01 = "63002" Then
                  strStarDate = "19740222": strEndDate = "19760930"
               End If
               If oST01 = "64001" Then
                  strStarDate = "19750301": strEndDate = "19760930"
               End If
               If oST01 = "65001" Then
                  strStarDate = "19760501": strEndDate = "19760930"
               End If
               If oST01 = "68001" Then
                  strStarDate = "19760401": strEndDate = "19760930"
               End If
               If oST01 = "72010" Then
                  strStarDate = "19830420": strEndDate = "19880630"
               End If
               Call PUB_NianZiDaysYear(strStarDate, strEndDate, LngReturnDay, intReturnYear)
            End If
            
            strStarDate = CheckStr(m_rs.Fields(0)) '到職日
            
            '抓該入所年的第一個工作天
            m_Year1_Std = GetYearStdDay(Mid(strStarDate, 1, 4))
            '假如 該入所年的第一個工作天與到職日同一天, 算一年的年資
            If strStarDate = m_Year1_Std And _
               (Left(strStarDate, 4) < Left(strCntEndDate, 4)) Then
               intReturnYear = 1
               strStarDate = CStr(Val(Left(strStarDate, 4) + 1)) & "0101"
            End If
            
            '任職時間
            '02.復職 03.離職 04.留職停薪 08.退休 09.撤職 10.資遣
            m_str2 = "select sc02,sc03 " & _
                           "from staff_change " & _
                           "where sc03 in ('02','03','04','08','09','10')  " & _
                           "and sc01='" & oST01 & "' " & _
                           "order by sc02 "
            If m_rs2.State = 1 Then m_rs2.Close
            m_rs2.CursorLocation = adUseClient
            m_rs2.Open m_str2, cnnConnection, adOpenStatic, adLockReadOnly
            If Not m_rs2.EOF And Not m_rs2.BOF Then
                m_rs2.MoveFirst
                int_i = 0
                Do While Not m_rs2.EOF
                    int_i = int_i + 1
                    
                    '有到職日並且異動資料只有一筆
                    If strStarDate <> "" And m_rs2.RecordCount = 1 Then
                        If CheckStr(m_rs2.Fields(1)) = "04" Then
                           strEndDate = ChangeWDateStringToWString(DateAdd("d", -1, ChangeTStringToWDateString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(oST01, ChangeWStringToTDateString(CheckStr(m_rs2.Fields(0))))))))
                        Else
                           strEndDate = strCntEndDate
                        End If
                        If strStarDate <= strEndDate Then
                           Call PUB_NianZiDaysYear(strStarDate, strEndDate, LngReturnDay, intReturnYear)
                        End If
                        
                    '有到職日並且非異動資料的最後一筆
                    ElseIf strStarDate <> "" And m_rs2.RecordCount > 1 And m_rs2.AbsolutePosition <> m_rs2.RecordCount Then
                        strEndDate = ChangeWDateStringToWString(DateAdd("d", -1, ChangeTStringToWDateString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(oST01, ChangeWStringToTDateString(CheckStr(m_rs2.Fields(0))))))))
                        If int_i Mod 2 <> 0 Then
                           If strStarDate <= strEndDate Then
                              Call PUB_NianZiDaysYear(strStarDate, strEndDate, LngReturnDay, intReturnYear)
                           End If
                        End If
                        strStarDate = ChangeTStringToWString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(oST01, ChangeWStringToTDateString(CheckStr(m_rs2.Fields(0))))))
                        
                    Else
                        strEndDate = ChangeWDateStringToWString(DateAdd("d", -1, ChangeTStringToWDateString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(oST01, ChangeWStringToTDateString(CheckStr(m_rs2.Fields(0))))))))
                        If int_i Mod 2 <> 0 Then
                           If strStarDate <= strEndDate Then
                              Call PUB_NianZiDaysYear(strStarDate, strEndDate, LngReturnDay, intReturnYear)
                           End If
                        End If
                        strStarDate = ChangeTStringToWString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(oST01, ChangeWStringToTDateString(CheckStr(m_rs2.Fields(0))))))
                        
                        If CheckStr(m_rs2.Fields(1)) = "04" Then
                           strEndDate = ChangeWDateStringToWString(DateAdd("d", -1, ChangeTStringToWDateString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(oST01, ChangeWStringToTDateString(CheckStr(m_rs2.Fields(0))))))))
                        Else
                           strEndDate = strCntEndDate
                        End If
                        If strStarDate <= strEndDate Then
                           Call PUB_NianZiDaysYear(strStarDate, strEndDate, LngReturnDay, intReturnYear)
                        End If
                    End If
                    m_rs2.MoveNext
                Loop
            '沒異動資料
            Else
               '為固定任職期間寫死的條件判斷
               If oST01 = "72010" Then
                  strStarDate = "19880701"
               End If
               strEndDate = strCntEndDate
               If strStarDate <= strEndDate Then
                  Call PUB_NianZiDaysYear(strStarDate, strEndDate, LngReturnDay, intReturnYear)
               End If
            End If
            m_rs.MoveNext
        Loop
    End With
End If

If intReturnYear = 0 And LngReturnDay = 0 Then
   CalYear = 0
Else
   '滿年
   CalYear = intReturnYear + (Round(LngReturnDay / 365, 2) * 100 \ 100)
   If CalYear < 1 Then
      'Add By Sindy 2010/01/18 檢查到職日是否為當月第一天
      Dim dblTempDay As Double, dblTempDay2 As Double
      dblTempDay = PUB_GetWorkDayAfterSysDate(Val(strStarDate), -1) + 19110000
      dblTempDay2 = ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(strStarDate)))
      If Left(Trim(dblTempDay), 6) = Left(Trim(dblTempDay2), 6) Then
         strStarDate = Left(Trim(strStarDate), 6) & "01" '到職日
      End If
      '2010/01/18 End
      '未滿一年算月份
      m_month = DateDiff("m", DateAdd("d", -1, ChangeWStringToWDateString(strStarDate)), ChangeWStringToWDateString(strEndDate))
      CalYear = Format(m_month / 12, "0.00")
   End If
End If
End Function

''add by nickc 2007/12/28 計算年資(新進,留職停薪,離職,復職)
'Public Function CalYear_M(oST01 As String, strEndDate As String, Optional oIsMonth As Boolean = False) As Double
'Dim m_month As Integer   '記錄年資月份
'Dim m_StrSQL As String
'Dim m_rs As New ADODB.Recordset
'Dim m_ST13_Day As String        '員工入所日
'Dim m_Year1_Std As String     '員工入所當年的第一個工作天
'Dim m_Cal_day As String
''add by nickc 2008/01/30
'Dim m_IsMonthStd As Boolean
'Dim strCntDate As String '計算年資截止日
'
'   '先抓員工的入所日
'   m_StrSQL = "select * from staff where st01='" & oST01 & "' "
'   Set m_rs = New ADODB.Recordset
'   If m_rs.State = 1 Then m_rs.Close
'   m_rs.CursorLocation = adUseClient
'   m_rs.Open m_StrSQL, cnnConnection, adOpenStatic, adLockReadOnly
'   If m_rs.RecordCount <> 0 Then
'      m_ST13_Day = CheckStr(m_rs.Fields("ST13"))
'   Else
'      m_ST13_Day = ""
'      CalYear_M = 0
'      Exit Function
'   End If
'
'   '抓該入所年的第一個工作天
'   m_Year1_Std = GetYearStdDay(Mid(m_ST13_Day, 1, 4))
'
'   '假如  入所日與到職日同一天，算一年的年資，若沒有，算月份
'   If m_ST13_Day = m_Year1_Std Then
'      m_month = 12
'   Else
'      m_month = DateDiff("m", DateAdd("d", -1, ChangeWStringToWDateString(m_ST13_Day)), Mid(m_ST13_Day, 1, 4) & "/12/31")
'   End If
'
'   '已經計算到入所下一年
'   m_Cal_day = GetYearStdDay(Trim(Val(Mid(m_ST13_Day, 1, 4)) + 1))
'
'   If strEndDate = "" Then
'      strCntDate = strSrvDate(1) '計算年資截止日：系統日期
'   Else
'      strCntDate = strEndDate
'   End If
'
'   If m_Cal_day <= strCntDate Then
'      '檢查有無異動紀錄
'      m_StrSQL = "select * from staff_change where sc01='" & oST01 & "' and sc03 in ('02','03','04','08','09','10') order by sc02 "
'      Set m_rs = New ADODB.Recordset
'      If m_rs.State = 1 Then m_rs.Close
'      m_rs.CursorLocation = adUseClient
'      m_rs.Open m_StrSQL, cnnConnection, adOpenStatic, adLockReadOnly
'      If m_rs.RecordCount <> 0 Then
'         m_rs.MoveFirst
'         Do While Not m_rs.EOF
'            '離職
'            If m_rs.Fields("sc03") = "03" Or m_rs.Fields("sc03") = "04" Or _
'               m_rs.Fields("sc03") = "08" Or m_rs.Fields("sc03") = "09" Or _
'               m_rs.Fields("sc03") = "10" Then
'               If m_Cal_day <> "" Then
'                  m_month = m_month + DateDiff("m", DateAdd("d", -1, ChangeWStringToWDateString(m_Cal_day)), ChangeWStringToWDateString(CheckStr(m_rs.Fields("sc02"))))
'                  m_Cal_day = ""
'               End If
'            End If
'            '復職
'            If m_rs.Fields("sc03") = "02" Then
'               m_Cal_day = CheckStr(m_rs.Fields("sc02"))
'            End If
'            m_rs.MoveNext
'         Loop
'      End If
'      'add by nickc 2008/01/30 檢查是否為該月第一天
'      m_IsMonthStd = False
'
'      'edit by nickc 2008/05/22  加判斷，若是入所日等於當年第一日時，不檢查
'      If m_ST13_Day = GetMonthStdDay(Mid(m_ST13_Day, 1, 6)) And Mid(m_ST13_Day, 7, 2) <> "01" And m_ST13_Day <> m_Year1_Std Then
'         m_IsMonthStd = True
'      End If
'
'      If strEndDate = "" Then
'         '計算剩餘的到目前時間的前一年年底
'         'edit by nickc 2008/05/22  加入控制，不可以大於系統前年底
'         If m_Cal_day <> "" And m_Cal_day <= Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(strSrvDate(1))), "yyyy") & "1231" Then
'            'edit by nickc 2008/05/22  修正，應該是要算到系統前一年年底
'            m_month = m_month + DateDiff("m", ChangeWStringToWDateString(m_Cal_day), ChangeWStringToWDateString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(strSrvDate(1))), "yyyy") & "1231")) + 1
'         End If
'      Else
'         '計算剩餘的到strEndDate
'         If m_Cal_day <> "" And m_Cal_day <= strEndDate Then
'            m_month = m_month + DateDiff("m", ChangeWStringToWDateString(m_Cal_day), ChangeWStringToWDateString(strEndDate)) + 1
'         End If
'      End If
'
'      'add by nickc 2008/01/30 若是等於該月第一天，要 + 1 個月
'      If m_IsMonthStd Then
'         m_month = m_month + 1
'      End If
'   Else
'      m_month = 0
'   End If
'
'   '寫死-補大頭們年資
'   If oST01 = "63001" Then
'      m_month = m_month + 31
'   ElseIf oST01 = "63002" Then
'      m_month = m_month + 31
'   ElseIf oST01 = "64001" Then
'      m_month = m_month + 19
'   ElseIf oST01 = "65001" Then
'      m_month = m_month + 5
'   ElseIf oST01 = "68001" Then
'      m_month = m_month + 6
'   End If
'
'   If oIsMonth Then
'      CalYear_M = Int(m_month)
'   Else
'      CalYear_M = Format(m_month / 12, "#0.00")
'   End If
'End Function

'add by nickc 2007/12/03 人事使用(共用) 計算請假天數和時數
'日期時間計算差距  小時
' Optional bwk4hour As Boolean ==> Add By Sindy 2010/7/14 是否1天工作4小時
' Optional bwk5hour As Boolean ==> Modify By Sindy 2011/3/8 是否1天工作5小時
'Modify By Sindy 2012/7/9 將 bwk5hour 變數改為 PUB_bWkSpec 變數
'Add by Sindy 2013/
Function CalDateTime(StrST01 As String, oDT1 As String, oDT2 As String, Optional PUB_bWkSpec As Boolean, Optional strSTime As String, Optional strETime As String, Optional bolChkWorkDay As Boolean = True) As String
Dim otd1, ott1, otd2, ott2, tmpott1, tmpott2, tmpott
Dim tmpvar As Variant
Dim CalD, CalH, CalM, CalTmp
Dim i As Integer, dblDay As Double 'Add By Sindy 2011/9/8
Dim bolRest1Day As Boolean 'Add By Sindy 2011/9/16 是否請整天
Dim strStarWorkTime As String, strEndWorkTime As String
Dim j As Integer
   
   Call Pub_GetSpecWorkHour(StrST01, Left(oDT1, 8), strStarWorkTime, strEndWorkTime)
   
   '取出1
   If InStr(1, oDT1, "/") <> 0 Then
      tmpvar = Split(oDT1, "/")
      If UBound(tmpvar) > 1 Then
         otd1 = tmpvar(0) & "/" & tmpvar(1) & "/" & Mid(tmpvar(2), 1, 2)
         If InStr(1, oDT1, ":") <> 0 Then
             ott1 = Replace(Right(oDT1, 5), ":", "")
         Else
             ott1 = Right(oDT1, 4)
         End If
      Else
         CalDateTime = ""
         Exit Function
      End If
   ElseIf InStr(1, oDT1, ":") <> 0 Then
      otd1 = Mid(oDT1, 1, Len(oDT1) - 5)
      ott1 = Replace(Right(oDT1, 5), ":", "")
   Else
      otd1 = Mid(oDT1, 1, Len(oDT1) - 4)
      ott1 = Right(oDT1, 4)
   End If
   otd1 = DBDATE(otd1)
   '取出2
   If InStr(1, oDT2, "/") <> 0 Then
      tmpvar = Split(oDT2, "/")
      If UBound(tmpvar) > 1 Then
         otd2 = tmpvar(0) & "/" & tmpvar(1) & "/" & Mid(tmpvar(2), 1, 2)
         If InStr(1, oDT2, ":") <> 0 Then
             ott2 = Replace(Right(oDT2, 5), ":", "")
         Else
             ott2 = Right(oDT2, 4)
         End If
      Else
         CalDateTime = ""
         Exit Function
      End If
   ElseIf InStr(1, oDT2, ":") <> 0 Then
      otd2 = Mid(oDT2, 1, Len(oDT2) - 5)
      ott2 = Replace(Right(oDT2, 5), ":", "")
   Else
      otd2 = Mid(oDT2, 1, Len(oDT2) - 4)
      ott2 = Right(oDT2, 4)
   End If
   otd2 = DBDATE(otd2)
   '檢查資料
   If otd2 < otd1 Then '迄比起還小
      CalDateTime = ""
      Exit Function
   ElseIf otd1 = otd2 And ott2 < ott1 Then  '同一天，但是時間迄比時間起還小
      CalDateTime = ""
      Exit Function
   End If
   '開始比較
   'CalD = Val(GetWorkDay(otd2, otd1)) - 1 '扣除工作天
   '天數 ***********
   'CalD = Val(otd2) - Val(otd1)
   'CalD = DateValue(ChangeWStringToWDateString(Val(otd2))) - DateValue(ChangeWStringToWDateString(Val(otd1)))
   
   'Add By Sindy 2011/9/8
   If otd1 <> otd2 Then '請多天
      '計算工作天
      CalD = 0: strDate = otd1
      For dblDay = otd1 To otd2
         dblDay = strDate
         If bolChkWorkDay = True Then '要檢查是否工作天
            If ChkWorkDay(DBDATE(strDate)) Then
               CalD = CalD + 1
               'Add By Sindy 2013/7/23 廖宗岳特殊上班時間(休週五算2天假)
               If PUB_bWkSpec = True And StrST01 = "73029" And Weekday(Format(DBDATE(strDate), "####-##-##")) = 6 Then
                  '週五在請假區間內的直接算休2天
                  If DBDATE(strDate) <> DBDATE(otd1) And DBDATE(strDate) <> DBDATE(otd2) Then
                     CalD = CalD + 1
                  End If
               End If
               '2013/7/23 END
            End If
         Else 'Add By Sindy 2011/10/14 劉經理：國外及大陸出差天數計算應含休假日,國內出差則以實際工作時數計算
            CalD = CalD + 1
         End If
         strDate = DBDATE(ChangeWStringToTString(DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(CStr(dblDay)))))))
         If strDate > otd2 Then dblDay = strDate
      Next dblDay
      '第一天和最後一天分別計算時數
      If CalD > 0 Then
         'Modify By Sindy 2011/11/9
         If bolChkWorkDay = True Then '要檢查是否工作天
            If ChkWorkDay(DBDATE(otd1)) = False And ChkWorkDay(DBDATE(otd2)) = False Then
               dblDay = 0
            Else
               If ChkWorkDay(DBDATE(otd1)) = False Or ChkWorkDay(DBDATE(otd2)) = False Then
                  CalD = CalD - 1
               Else
                  CalD = CalD - 2
               End If
               dblDay = 2
            End If
         Else
         '2011/11/9 End
            CalD = CalD - 2
            dblDay = 2
         End If
      Else
         dblDay = 2
      End If
   Else
      dblDay = 1 '請一天
   End If
   '天數 END ***********
   For i = 1 To dblDay 'dblDay只有等於1或2
      'Add By Sindy 2011/11/9
      If bolChkWorkDay = True Then '要檢查是否工作天
         If i = 1 And ChkWorkDay(DBDATE(otd1)) = False Then GoTo ReadNext
         If i = 2 And ChkWorkDay(DBDATE(otd2)) = False Then GoTo ReadNext
      End If
      
      If dblDay = 1 Then '請1天(內)
         tmpott1 = ott1
         tmpott2 = ott2
      Else
         '第1天
         If i = 1 Then
            Call PUB_ChkByPassWork(PUB_GetST06(StrST01), Left(oDT1, 8)) 'Add By Sindy 2021/8/13
            tmpott1 = ott1
            'Modify By Sindy 2012/7/9
            'If bwk5hour = True Then
            If PUB_bWkSpec = True And Not (StrST01 = "73029" And Weekday(Format(DBDATE(otd1), "####-##-##")) = 6) Then
'               'Modify by Sindy 2013/7/23
'               If strST01 = "73029" Then
'                  tmpott2 = "1730"
'               Else
'               '2013/7/23 END
'                  If PUB_intWkHour = 4 Then
'                     tmpott2 = "1200"
'                  Else
'               '2012/7/9 End
'                     tmpott2 = "1730" 'PUB_intWkHour = 5
'                  End If
'               End If
               tmpott2 = strEndWorkTime 'Add By Sindy 2015/9/17
            Else
'               tmpott2 = "1700" '預設值
'               '中途下班
'               If strSTime <> "" Then
'                  tmpott = strSTime
'               Else
'                  tmpott = ott1
'               End If
'               If CStr(tmpott) <= "0800" Then
'                  tmpott2 = "1700"
'               ElseIf CStr(tmpott) > "0800" And CStr(tmpott) <= "0830" Then
'                  tmpott2 = "1730"
'               ElseIf CStr(tmpott) > "0830" And CStr(tmpott) <= "0900" Then
'                  tmpott2 = "1800"
'               End If
               'Add By Sindy 2021/8/13
               tmpott2 = Format(strByPassEndTime(1), "HHMM") '預設值
               '中途下班
               If strSTime <> "" Then
                  tmpott = strSTime
               Else
                  tmpott = ott1
               End If
               For j = 1 To intByPassArea
                  If CStr(tmpott) = CStr(Format(strByPassStarTime(j), "HHMM")) Then
                     tmpott2 = Format(strByPassEndTime(j), "HHMM")
                     Exit For
                  End If
               Next j
               '2021/8/13 END
            End If
         '最後1天
         Else
            Call PUB_ChkByPassWork(PUB_GetST06(StrST01), Left(oDT2, 8)) 'Add By Sindy 2021/8/13
            tmpott2 = ott2
            'Modify By Sindy 2012/7/9
            'If bwk5hour = True Then
            If PUB_bWkSpec = True And Not (StrST01 = "73029" And Weekday(Format(DBDATE(otd2), "####-##-##")) = 6) Then
'               'Modify by Sindy 2013/7/23
'               If strST01 = "73029" Then
'                  tmpott1 = "1330"
'               Else
'               '2013/7/23 END
'                  If PUB_intWkHour = 4 Then
'                     tmpott1 = "0800"
'                  Else
'               '2012/7/9 End
'                     'Modify By Sindy 2012/11/15
'                     'tmpott1 = "1230" 'PUB_intWkHour = 5
'                     tmpott1 = "1130" 'PUB_intWkHour = 5
'                     '2012/11/15 End
'                  End If
'               End If
               tmpott1 = strStarWorkTime 'Add By Sindy 2015/9/17
            Else
'               tmpott1 = "0800" '預設值
'               '中途上班
'               If strETime <> "" Then
'                  tmpott = strETime
'               Else
'                  tmpott = ott2
'               End If
'               If CStr(tmpott) >= "1700" And CStr(tmpott) <= "1729" Then
'                  tmpott1 = "0800"
'               ElseIf CStr(tmpott) >= "1730" And CStr(tmpott) <= "1759" Then
'                  tmpott1 = "0830"
'               ElseIf CStr(tmpott) >= "1800" Then
'                  tmpott1 = "0900"
'               End If
               'Add By Sindy 2021/8/13
               tmpott1 = Format(strByPassStarTime(1), "HHMM") '預設值
               '中途上班
               If strETime <> "" Then
                  tmpott = strETime
               Else
                  tmpott = ott2
               End If
               For j = 1 To intByPassArea
                  If CStr(tmpott) = CStr(Format(strByPassEndTime(j), "HHMM")) Then
                     tmpott1 = Format(strByPassStarTime(j), "HHMM")
                     Exit For
                  End If
               Next j
               '2021/8/13 END
            End If
         End If
      End If
      
      'Add By Sindy 2022/5/26 增加判斷是否為整日
      If dblDay > 1 Then
         If i = 1 And Val(tmpott1) <= Val(IIf(strStarWorkTime = "", 900, Val(strStarWorkTime))) Then
            CalD = CalD + 1
            GoTo ReadNext
         ElseIf i = 2 And Val(tmpott2) >= Val(IIf(strEndWorkTime = "", 1700, Val(strEndWorkTime))) Then
            CalD = CalD + 1
            GoTo ReadNext
         End If
      End If
      '2022/5/26 END
      
      CalTmp = ((Val(Mid(tmpott2, 1, 2)) * 60) + Val(Mid(tmpott2, 3))) - ((Val(Mid(tmpott1, 1, 2)) * 60) + Val(Mid(tmpott1, 3)))
      'Add By Sindy 98/03/13 起始時間<=12時並且迄止時間>=13時30分者，減1小時
      'If CStr(tmpott1) <= "1200" And CStr(tmpott2) >= "1330" Then
      '若輸入09:10-18:00時,只工作了10分鐘,應算請整日
      If CalTmp <= 540 And CalTmp > 510 Then
         CalTmp = 540
      End If
      'Modify By Sindy 2011/9/8
      'Modify By Sindy 2012/7/9
      'If bwk5hour = True Or CalTmp >= 540 Then '一天工作五小時或請整日的人
      If PUB_bWkSpec = True Or CalTmp >= 540 Then '一天工作四或五小時或請整日的人
      '2012/7/9 End
         CalM = CalM + CalTmp
         'Modify By Sindy 2012/7/9
         'If bwk5hour = False Then
         'Modify by Sindy 2013/7/23
         'If PUB_bWkSpec = False Then
         '廖宗岳特殊上班時間(週五上班整日)
         If PUB_bWkSpec = True And StrST01 = "73029" Then
            If i = 1 Then strDate = DBDATE(otd1)
            If i = 2 Then strDate = DBDATE(otd2)
            If Weekday(Format(strDate, "####-##-##")) = 6 Then
               CalM = CalM - 60
            End If
         '2013/7/23 END
         ElseIf PUB_bWkSpec = False Then
         '2012/7/9 End
            CalM = CalM - 60
         End If
      Else '非整日
         '拆上午及下午計算(因1200-1210為工作時間,難判斷何時要減60分鐘何時要減80分鐘,因此上午下午分開計算)
         'Modify By Sindy 2017/4/17
         If strSrvDate(1) >= 中午休息時間改1200 Then
            If CStr(tmpott1) < "1200" Then '上午
               If CStr(tmpott2) >= "1200" Then
                  tmpott = "1200"
               Else
                  tmpott = tmpott2
               End If
               CalM = CalM + ((Val(Mid(tmpott, 1, 2)) * 60) + Val(Mid(tmpott, 3))) - ((Val(Mid(tmpott1, 1, 2)) * 60) + Val(Mid(tmpott1, 3)))
            End If
         Else
         '2017/4/17 END
            If CStr(tmpott1) < "1210" Then '上午
               If CStr(tmpott2) >= "1210" Then
                  tmpott = "1210"
               Else
                  tmpott = tmpott2
               End If
               CalM = CalM + ((Val(Mid(tmpott, 1, 2)) * 60) + Val(Mid(tmpott, 3))) - ((Val(Mid(tmpott1, 1, 2)) * 60) + Val(Mid(tmpott1, 3)))
            End If
         End If
         If CStr(tmpott2) > "1330" Then '下午
            If CStr(tmpott1) <= "1330" Then
               tmpott = "1330"
            Else
               tmpott = tmpott1
            End If
            CalM = CalM + ((Val(Mid(tmpott2, 1, 2)) * 60) + Val(Mid(tmpott2, 3))) - ((Val(Mid(tmpott, 1, 2)) * 60) + Val(Mid(tmpott, 3)))
         End If
      End If
      
ReadNext: 'Add By Sindy 2011/11/9
   Next i
'   CalM = ((Val(Mid(ott2, 1, 2)) * 60) + Val(Mid(ott2, 3))) - ((Val(Mid(ott1, 1, 2)) * 60) + Val(Mid(ott1, 3)))
'   If CalM < 0 Then '不夠減
'      CalD = CalD - 1
'      CalM = (((Val(Mid(ott2, 1, 2)) * 60) + Val(Mid(ott2, 3))) - (8 * 60)) + ((17 * 60) - ((Val(Mid(ott1, 1, 2)) * 60) + Val(Mid(ott1, 3))))
'   End If
   '2011/9/8 End
   
   CalH = CalM \ 60
   ' 2008/12/19 Modify BY SINDY
   'CalM = ((CalM - (CalH * 60)) \ 30) * 0.5
   CalM = Round((CalM - (CalH * 60)) / 60, 2)
   ' 2008/12/19 END
   
   'Add By Sindy 2011/9/8 未滿半小時以半小時計
   If Val(CalM) <= 0.5 And CalM > 0 Then CalM = 0.5
   If Val(CalM) > 0.5 And CalM > 0 Then CalM = 1
   
   'Modify By Sindy 2010/7/14
   'Modify By Sindy 2011/3/8 改5小時
   'Modify By Sindy 2012/7/9
   'If bwk5hour = True Then
   CalDateTime = (CalD * PUB_intWkHour) + CalH + CalM
End Function
'Removed by Morgan 2025/7/29 114/7/28起廢止婚喪互助辦法
''Add by Morgan 2008/12/23
''取得婚喪戶助金額
'Public Function PUB_GetHelpFee(strUserNo As String, Optional strHelpFee As String) As Boolean
'Dim stSQL As String, intR As Integer, stPosCode As String
'Dim adoRst As ADODB.Recordset
'
'   strHelpFee = ""
'   stSQL = "select st21 from staff where st01='" & strUserNo & "'"
'   intR = 1
'   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
'   If intR = 1 Then
'      stPosCode = "" & adoRst(0)
'      If stPosCode <> "" Then
'         '71~
'         If Val(stPosCode) > 70 Then
'            strHelpFee = 30
'         '61-70
'         ElseIf Val(stPosCode) > 60 Then
'            strHelpFee = 50
'         '26-60
'         ElseIf Val(stPosCode) > 25 Then
'            strHelpFee = 70
'         '1-25
'         Else
'            strHelpFee = 100
'         End If
'
'      'Added by Morgan 2025/1/13 未設定職位時預設最低金額(因0或Null都是為不扣),否則後續修改時都不會再更新 Ex:B3032
'      Else
'         strHelpFee = 30
'      'end 2025/1/13
'      End If
'      PUB_GetHelpFee = True
'   End If
'   Set adoRst = Nothing
'End Function

'Add by SINDY 2008/12/24
'*************************************************
' 取得各假別時數
' Input : strStaffSQL     ==> 員工檔Where條件 ex. and ST01='97037'
' Input : strSDate           ==> 起始日期
' Input : strEDate           ==> 終止日期
' Output : strHelpHour() ==> 時數
' Output : strHelpCnt()   ==> 人數
'假別List:
'01   忘打卡
'02   遲到
'03   曠職
'04   出差
'05   事假
'06   病假
'07   公假
'08   特別假
'09   婚假
'10   產假
'11   流產假
'12   喪假
'13   公傷假
'14   補休
'15   其他
'16   加班
'17   扣年終產假
'18   扣年終流產假
'19   陪產假
'20   生理假
'21   產檢假
'22   家庭照顧假
'23   健檢假
'24   防疫照顧假
'25   天災不給薪 Add By Sindy 2025/11/13
'*************************************************
Public Function PUB_GetAbsenceHour(strStaffSQL As String, strSDate As String, strEDate As String, strHelpHour() As Double, strHelpCnt() As Double) As Boolean
Dim stSQL As String, intR As Integer, dblHour As Double, dblHour2 As Double
Dim adoRst As New ADODB.Recordset
Dim i As Integer, strWhere As String, dblCnt As Double, j As Long
Dim strStarDate As String, strEndDate As String
Dim intNotWork As Integer 'Add By Sindy 2014/3/5 記錄有遲到曠職
   
   'strHelpHour(0) 此陣列值不使用
   'strHelpCnt(0)   此陣列值不使用
   For i = 1 To 25 '24 '23 '22 '20
      strHelpHour(i) = 0
      strHelpCnt(i) = 0
   Next i
   
   '********************************
   '出缺勤檔 - 忘打卡,曠職
   'Modify By Sindy 2010/7/14 增加ST01
   'Modify By Sindy 2010/10/7 因為遲到須要做換算所以提出來, 另外在下段程式處理
   'Modified by Morgan 2025/8/15 修改曠職時間改以「分」計算
   If strSrvDate(1) >= 曠職以分計啟用日 Then
      'Modify By Sindy 2025/11/14 嘉渝提加總要顯示分鐘,不要轉換為小時
      'stSQL = "select sum(nvl(SA03,0)),sum(nvl(SA04,0)),sum(nvl(SA05,0)),sum(nvl(SA06,0)/60),ST01 "
      stSQL = "select sum(nvl(SA03,0)),sum(nvl(SA04,0)),sum(nvl(SA05,0)),sum(nvl(SA06,0)),ST01 " & _
            "From staff, staff_Assist " & _
            "Where ST01 = sa01 " & _
            "and (sa02 between '" & strSDate & "' and '" & strEDate & "') " & strStaffSQL & " group by ST01 "
   Else
      stSQL = "select sum(nvl(SA03,0)),sum(nvl(SA04,0)),sum(nvl(SA05,0)),sum(nvl(SA06,0)),ST01 " & _
            "From staff, staff_Assist " & _
            "Where ST01 = sa01 " & _
            "and (sa02 between '" & strSDate & "' and '" & strEDate & "') " & strStaffSQL & " group by ST01 "
   End If
   If adoRst.State = 1 Then adoRst.Close
   adoRst.CursorLocation = adUseClient
   adoRst.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF And Not adoRst.BOF Then
      If adoRst.RecordCount <> 0 Then
         adoRst.MoveFirst
         Call Pub_GetSpecWorkHour(Trim("" & adoRst.Fields("ST01")), strSDate) 'Add By Sindy 2012/7/9 上班時數為特殊者
         While Not adoRst.EOF
            '01.忘打卡
            If Not IsNull(adoRst(0)) Then strHelpHour(1) = strHelpHour(1) + adoRst(0)
'            '02.遲到
'            If Not IsNull(adoRst(1)) Then strHelpHour(2) = strHelpHour(2) + adoRst(1)
            '03.曠職
            If IsNull(adoRst(2)) Then dblHour = 0 Else dblHour = adoRst(2)
            If IsNull(adoRst(3)) Then dblHour2 = 0 Else dblHour2 = adoRst(3)
            If dblHour > 0 Or dblHour2 > 0 Then
               'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
               'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
               'Modify By Sindy 2012/7/9 上班時數為特殊者
''               If Trim("" & adoRst.Fields("ST01")) = "99029" Then
''                  dblHour = (dblHour * 5) + dblHour2
'               dblHour = (dblHour * PUB_intWkHour) + dblHour2
'               strHelpHour(3) = strHelpHour(3) + dblHour
               'Modify By Sindy 2025/11/14 加總要顯示分鐘,不要轉換為小時
               dblHour = (dblHour * PUB_intWkHour) * 60 '(天數*1日工時)*60分鐘
               strHelpHour(3) = strHelpHour(3) + dblHour + dblHour2 '分鐘
            End If
            adoRst.MoveNext
         Wend
      End If
   End If
   
   '出缺勤檔 - 忘打卡,曠職人數
   'Modify By Sindy 2010/10/7 因為遲到須要做換算所以提出來, 另外在下段程式處理
   For i = 1 To 3 Step 2
         If i = 1 Then
            strWhere = " and sa03>0 " '忘打卡
'         ElseIf i = 2 Then
'            strWhere = " and sa04>0 " '遲到
         ElseIf i = 3 Then
            strWhere = " and (sa05>0 or sa06>0) " '曠職
         End If
         stSQL = "select nvl(sum(count(distinct sa01)),0) " & _
                  "From staff, staff_Assist " & _
                  "Where ST01 = sa01 " & _
                  "and (sa02 between '" & strSDate & "' and '" & strEDate & "') " & strWhere & strStaffSQL & _
                  " group by sa01 "
         If adoRst.State = 1 Then adoRst.Close
         adoRst.CursorLocation = adUseClient
         adoRst.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
         If Not adoRst.EOF And Not adoRst.BOF Then
            If IsNull(adoRst(0)) Then dblCnt = 0 Else dblCnt = adoRst(0)
            If dblCnt > 0 Then strHelpCnt(i) = dblCnt
         End If
   Next i
      
   'Modify By Sindy 2010/10/7
   For j = Val(Left(strSDate, 6)) To Val(Left(strEDate, 6))
      If Left(strSDate, 6) = Left(strEDate, 6) Then
         strStarDate = strSDate
         strEndDate = strEDate
      Else
         If j = Val(Left(strSDate, 6)) Then
            strStarDate = strSDate
         Else
            strStarDate = CStr(j) & "01"
         End If
         If j = Val(Left(strEDate, 6)) Then
            strEndDate = strEDate
         Else
            strEndDate = CStr(j) & "31"
         End If
      End If
      '抓出該月總遲到次數
      'Modified by Morgan 2025/8/15 修改曠職時間改以「分」計算
      If strSrvDate(1) >= 曠職以分計啟用日 Then
         'Modify By Sindy 2025/11/14 嘉渝提加總要顯示分鐘,不要轉換為小時
         'stSQL = "SELECT sa01,nvl(sum(nvl(sa04,0)),0),nvl(sum(nvl(sa05,0)),0),nvl(sum(nvl(sa06,0)/60),0)
         stSQL = "SELECT sa01,nvl(sum(nvl(sa04,0)),0),nvl(sum(nvl(sa05,0)),0),nvl(sum(nvl(sa06,0)),0) " & _
                          "From staff, staff_Assist " & _
                        "Where ST01 = sa01 " & _
                             "and sa02 between " & strStarDate & " and " & strEndDate & " " & strStaffSQL & _
                             "group by sa01 "
      Else
         stSQL = "SELECT sa01,nvl(sum(nvl(sa04,0)),0),nvl(sum(nvl(sa05,0)),0),nvl(sum(nvl(sa06,0)),0) " & _
                          "From staff, staff_Assist " & _
                        "Where ST01 = sa01 " & _
                             "and sa02 between " & strStarDate & " and " & strEndDate & " " & strStaffSQL & _
                             "group by sa01 "
      End If
      If adoRst.State = 1 Then adoRst.Close
      adoRst.CursorLocation = adUseClient
      adoRst.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If Not adoRst.EOF And Not adoRst.BOF Then
         adoRst.MoveFirst
         While Not adoRst.EOF
            If IsNull(adoRst(1)) Then dblCnt = 0 Else dblCnt = adoRst(1)
            '出缺勤檔 - 遲到,遲到曠職時數 (若同月3次遲到就算1.5曠職, 第4次遲到算0.5曠職, 以此推算)
            'Modify By Sindy 2012/12/18 依江總指示,自101年1月起月遲到三次以上視為曠職之規定修正為:
            '                           前二次仍以遲到計算,第三次以後才以曠職計算(每次仍為30分)
            intNotWork = 0 'Add By Sindy 2014/3/5
            If dblCnt > 0 Then
               '2012年01月之後的新規定
'               If Val(Left(strStarDate, 6)) >= 201201 Then
                  '02.遲到
                  If dblCnt <= 2 Then
                     strHelpHour(2) = strHelpHour(2) + dblCnt
                  Else
                     strHelpHour(2) = strHelpHour(2) + 2 '前2次算遲到,之後全算曠職
                  End If
                  '03.遲到曠職
                  If dblCnt > 2 Then
                     'Modify By Sindy 2025/11/14 加總要顯示分鐘,不要轉換為小時
                     'strHelpHour(3) = strHelpHour(3) + ((dblCnt - 2) * 0.5)
                     strHelpHour(3) = strHelpHour(3) + ((dblCnt - 2) * 30) '改用分鐘記錄
                     intNotWork = 1 'Add By Sindy 2014/3/5
                  End If
               '2012/12/18 End
'               Else '2012年01月之前的算法
'                  If dblCnt <= 2 Then
'                     '02.遲到
'                     strHelpHour(2) = strHelpHour(2) + dblCnt
'                  Else
'                     '03.遲到曠職
'                     strHelpHour(3) = strHelpHour(3) + (dblCnt * 0.5)
'                     intNotWork = 1 'Add By Sindy 2014/3/5
'                  End If
'               End If
               '出缺勤檔 - 遲到,遲到曠職人數
               'Modify By Sindy 2012/12/18
'               If dblCnt <= 2 Then
'                  strHelpCnt(2) = strHelpCnt(2) + 1 '遲到人數+1
'               Else
'                  If adoRst(2) = 0 And adoRst(3) = 0 Then
'                     strHelpCnt(3) = strHelpCnt(3) + 1 '曠職人數+1
'                  End If
'               End If
               'If Val(strHelpHour(2)) > 0 Then
               If dblCnt > 0 Then
                  strHelpCnt(2) = strHelpCnt(2) + 1 '遲到人數+1
               End If
               'If adoRst(2) = 0 And adoRst(3) = 0 Then
               If intNotWork > 0 Then
                  strHelpCnt(3) = strHelpCnt(3) + 1 '曠職人數+1
               End If
               '2012/12/18 End
            End If
            adoRst.MoveNext
         Wend
      End If
   Next j
   '2010/10/7 End
   
   '********************************
   '04.出差檔 & 請假檔(假別:05~23) & 16.加班檔 - 時數
   'Modify By Sindy 2010/7/14 增加ST01,item
   'Modify By Sindy 2012/1/4  增加 19.陪產假
   'Modify By Sindy 2014/12/5 +    20.生理假,21.產檢假,22.家庭照顧假
   'Modify By Sindy 2014/12/31 +   23.健檢假
   'Modify By Sindy 2020/2/4 +     24.防疫照顧假
   'Modify By Sindy 2025/11/13 +   25.天災不給薪
   stSQL = "select sum(nvl(SB06,0)),sum(nvl(SB07,0)),ST01,4 as item From staff, staff_busi_trip Where ST01 = sb01 and (sb02 >= '" & strSDate & "' and sb04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,5 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='05' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,6 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='06' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,7 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='07' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,8 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='08' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,9 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='09' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,10 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='10' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,11 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='11' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,12 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='12' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,13 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='13' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,14 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='14' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,15 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='15' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select 0,sum(nvl(So05,0))+sum(nvl(So06,0)),ST01,16 as item  From staff, Staff_Overtime Where ST01 = so01 and (so02 >= '" & strSDate & "' and so02 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,17 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='17' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,18 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='18' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,19 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='19' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,20 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='20' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,21 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='21' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,22 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='22' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,23 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='23' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,24 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='24' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 Union All " & _
            "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),ST01,25 as item  From staff, staff_Absence Where ST01 = sa01 and SA06='25' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by ST01 "
   If adoRst.State = 1 Then adoRst.Close
   adoRst.CursorLocation = adUseClient
   adoRst.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF And Not adoRst.BOF Then
      If adoRst.RecordCount <> 0 Then
         adoRst.MoveFirst
         'i = 1
         Call Pub_GetSpecWorkHour(Trim("" & adoRst.Fields("ST01")), strSDate) 'Add By Sindy 2012/7/9 上班時數為特殊者
         While Not adoRst.EOF
            If IsNull(adoRst(0)) Then dblHour = 0 Else dblHour = adoRst(0)
            If IsNull(adoRst(1)) Then dblHour2 = 0 Else dblHour2 = adoRst(1)
            If dblHour > 0 Or dblHour2 > 0 Then
               'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
               'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
               'Modify By Sindy 2012/7/9 上班時數為特殊者
'               If Trim("" & adoRst.Fields("ST01")) = "99029" Then
'                  dblHour = (dblHour * 5) + dblHour2
               dblHour = (dblHour * PUB_intWkHour) + dblHour2
               'strHelpHour(i + 3) = dblHour
               strHelpHour(adoRst.Fields("item")) = strHelpHour(adoRst.Fields("item")) + dblHour
            End If
            'i = i + 1
            adoRst.MoveNext
         Wend
      End If
   End If
   
   '04.出差檔 & 請假檔(假別:05~23) & 16.加班檔 - 人數
   'Modify By Sindy 2020/2/4 +     24.防疫照顧假
   'Modify By Sindy 2025/11/13 +   25.天災不給薪
   stSQL = "select sum(count(distinct sb01)) From staff, staff_busi_trip Where ST01 = sb01 and (sb02 >= '" & strSDate & "' and sb04 <= '" & strEDate & "') " & strStaffSQL & " group by sb01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='05' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='06' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='07' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='08' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='09' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06 in ('10','17') and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06 in ('11','18') and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='12' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='13' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='14' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='15' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct so01)) From staff, Staff_Overtime Where ST01 = so01 and (so02 >= '" & strSDate & "' and so02 <= '" & strEDate & "') " & strStaffSQL & " group by so01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='17' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='18' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='19' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='20' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='21' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='22' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='23' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='24' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 Union All " & _
            "select sum(count(distinct sa01)) From staff, staff_Absence Where ST01 = sa01 and SA06='25' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') " & strStaffSQL & " group by sa01 "
   If adoRst.State = 1 Then adoRst.Close
   adoRst.CursorLocation = adUseClient
   adoRst.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF And Not adoRst.BOF Then
      If adoRst.RecordCount <> 0 Then
         adoRst.MoveFirst
         i = 1
         While Not adoRst.EOF
            If IsNull(adoRst(0)) Then dblCnt = 0 Else dblCnt = adoRst(0)
            If dblCnt > 0 Then strHelpCnt(i + 3) = dblCnt
            i = i + 1
            adoRst.MoveNext
         Wend
      End If
   End If
   
   PUB_GetAbsenceHour = True
   Set adoRst = Nothing
End Function

'公司名稱查詢
'Modified by Morgan 2020/3/23 +pLang: 1=中 2=英 3=日 4=簡稱(中)
'Modify By Sindy 2020/3/26 + Optional ByRef strAddr As String : 抓地址
Public Function CompNameQuery(InputNo As String, Optional pLang As String = "1", _
   Optional ByRef strAddr As String) As String
Dim adoacc080 As New ADODB.Recordset
   
   adoacc080.CursorLocation = adUseClient
   adoacc080.Open "select * from acc080 where a0801 = '" & InputNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc080.RecordCount <> 0 Then
      If IsNull(adoacc080.Fields("a0802").Value) Then
         CompNameQuery = MsgText(601)
      Else
         'Added by Morgan 2020/3/23
         If pLang = "2" Then '英
            CompNameQuery = "" & adoacc080.Fields("a0803").Value
         ElseIf pLang = "3" Then '日
            CompNameQuery = "" & adoacc080.Fields("a0828").Value
         ElseIf pLang = "4" Then '簡稱(中)
            CompNameQuery = "" & adoacc080.Fields("a0820").Value
         Else
         'end 2020/3/23
            CompNameQuery = "" & adoacc080.Fields("a0802").Value
         End If 'Added by Morgan 2020/3/23
         strAddr = "" & adoacc080.Fields("a0804").Value '抓地址
      End If
   Else
      CompNameQuery = MsgText(601)
   End If
   adoacc080.Close
End Function

'2008/12/26 add by sonia
'取得其他所得人名稱
'modify by sonia 2016/1/20 +stroi02
Public Function ClsPDGetOtherIncomer(ByVal strNo As String, ByRef strName As String, Optional stroi02 As String = "") As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset

On Error GoTo ErrHand
   
   ClsPDGetOtherIncomer = False
   strName = ""
   'modify by sonia 2016/1/20 +oi02
   strSql = "select oi04,oi02 from OtherIncomer where oi01=" + CNULL(strNo)
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      If Not IsNull(rsRecordset.Fields(0)) Then strName = rsRecordset.Fields(0)
      If Not IsNull(rsRecordset.Fields(1)) Then stroi02 = rsRecordset.Fields(1)   'add by sonia 2016/1/20 +oi02
      ClsPDGetOtherIncomer = True
   End If
   rsRecordset.Close
   Exit Function

ErrHand:
   MsgBox Err.Description
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得員工的職稱
' Input : strStuff ==> 員工的代碼
' Output : 傳回員工的職稱
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modify By Sindy 2025/5/29 + , Optional intType As Integer = 0: 0=職稱中文 1=職稱代碼
Public Function GetStaffST20(ByVal strStuff As String, Optional intType As Integer = 0) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetStaffST20 = Empty
   
   strSql = "SELECT ac03,st20 FROM Staff,allcode " & _
            "WHERE ST01 = '" & strStuff & "' and (ac02(+)=ST20 and ac01='01') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields(0)) = False Then
         GetStaffST20 = rsTmp.Fields(0)
      End If
      'Add By Sindy 2025/5/29 1=職稱代碼
      If intType = 1 Then
         If IsNull(rsTmp.Fields(1)) = False Then
            GetStaffST20 = rsTmp.Fields(1)
         End If
      End If
      '2025/5/29 END
   End If
   
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'2008/12/31 若為薪資系統的檢查,當第三碼為'A'時改為用'0'取代來抓資料
'取得員工之公司名稱(含離職)
Public Function ClsPDGetStaffComp(ByVal StrStaff As String, Optional strStaffCompanyNo As String, Optional bolIsSalary As Boolean) As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset
Dim strNo As String
 
On Error GoTo ErrHand

   strNo = StrStaff
   If Mid(StrStaff, 3, 1) = "A" Then
      strNo = Left(StrStaff, 2) & "0" & Mid(StrStaff, 4)
   End If
   
   strStaffCompanyNo = ""
   strSql = "select sd19,sd28 from salarydata where sd01=" + CNULL(strNo)
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      strStaffCompanyNo = "" & rsRecordset.Fields(0)    '正常抓第一家公司別
      If bolIsSalary And Mid(StrStaff, 3, 1) = "A" Then '薪資系統員工編號第三碼'A'則抓第二家
         strStaffCompanyNo = "" & rsRecordset.Fields(1)
      End If
      ClsPDGetStaffComp = True
   Else
      ShowMsg "公司別錯誤 !"
   End If
   rsRecordset.Close
   Exit Function

ErrHand:
   MsgBox Err.Description
End Function

' 取得年終獎金基準月數 0,1,2,9公司都抓1公司的基準月數
Public Function GetYearBonusMonth(ByVal strYear As String, ByVal StrStaff As String) As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   GetYearBonusMonth = Empty
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'modify by sonia 2019/1/17 改為未輸基準月數的公司都抓1公司的基準月數
   'strSql = "SELECT c1.ybm03,c2.ybm03 FROM SalaryData,YearBonusMonth c1,YearBonusMonth c2 WHERE sd01= '" & Left(StrStaff, 1) & Replace(Mid(StrStaff, 2), "A", "0") & "' " & _
            "and c1.ybm01(+) = '" & Val(strYear) + 1911 & "' and c1.ybm02(+) = decode(sd19,'0','1','2','1','9','1',sd19) " & _
            "and c2.ybm01(+) = '" & Val(strYear) + 1911 & "' and c2.ybm02(+) = decode(sd28,'0','1','2','1','9','1',sd28) "
   'modify by sonia 2023/1/13 改為2021年起未輸基準月數的公司都抓2公司的基準月數
   'strSql = "SELECT nvl(c1.ybm03,c3.ybm03),nvl(c2.ybm03,c4.ybm03) FROM SalaryData,YearBonusMonth c1,YearBonusMonth c2,YearBonusMonth c3,YearBonusMonth c4 " & _
            "WHERE sd01= '" & Left(StrStaff, 1) & Replace(Mid(StrStaff, 2), "A", "0") & "' " & _
            "and c1.ybm01(+) = '" & Val(strYear) + 1911 & "' and c1.ybm02(+) = sd19 and c2.ybm01(+) = '" & Val(strYear) + 1911 & "' and c2.ybm02(+) = sd28 " & _
            "and c3.ybm01(+) = '" & Val(strYear) + 1911 & "' and c3.ybm02(+) = '2'  and c4.ybm01(+) = '" & Val(strYear) + 1911 & "' and c4.ybm02(+) = '2' "
   strSql = "SELECT nvl(c1.ybm03,c3.ybm03),nvl(c2.ybm03,c4.ybm03) FROM SalaryData,YearBonusMonth c1,YearBonusMonth c2,YearBonusMonth c3,YearBonusMonth c4 " & _
            "WHERE sd01= '" & Left(StrStaff, 1) & Replace(Mid(StrStaff, 2), "A", "0") & "' " & _
            "and c1.ybm01(+) = '" & Val(strYear) + 1911 & "' and c1.ybm02(+) = sd19 and c2.ybm01(+) = '" & Val(strYear) + 1911 & "' and c2.ybm02(+) = sd28 " & _
            "and c3.ybm01(+) = '" & Val(strYear) + 1911 & "' and c3.ybm02(+) = DECODE(SIGN(" & Val(strYear) + 1911 & "-2020),0,'1',-1,'1','2') " & _
            "and c4.ybm01(+) = '" & Val(strYear) + 1911 & "' and c4.ybm02(+) = DECODE(SIGN(" & Val(strYear) + 1911 & "-2020),0,'1',-1,'1','2') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      '依員工代號決定第一家或第二家
      If Mid(StrStaff, 3, 1) = "A" Then
         If IsNull(rsTmp.Fields(1)) = False Then
            GetYearBonusMonth = rsTmp.Fields(1)
         End If
      Else
         If IsNull(rsTmp.Fields(0)) = False Then
            GetYearBonusMonth = rsTmp.Fields(0)
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 取得考績及核發獎金基數
Public Function GetYearMerit(ByVal strYear As String, ByVal StrStaff As String, ByRef strym03 As String, ByRef strYearMerit As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   GetYearMerit = False
   strym03 = "": strYearMerit = "%"
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'strSQL = "SELECT ym02 FROM YearMerit WHERE ym01 = '" & Val(strYear) + 1911 & "' and ym03= '" & Replace(StrStaff, "A", "0") & "' "
   strSql = "SELECT ym02 FROM YearMerit WHERE ym01 = '" & Val(strYear) + 1911 & "' and ym03= '" & Left(StrStaff, 1) & Replace(Mid(StrStaff, 2), "A", "0") & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then
         Select Case rsTmp.Fields(0)
            Case "1"
               strym03 = "優等"
               strYearMerit = "110" & strYearMerit
            '2010/2/2 ADD BY SONIA
            Case "2"
               strym03 = "甲等"
               strYearMerit = "100" & strYearMerit
            '2010/2/2 END
            Case "3"
               strym03 = "乙等"
               strYearMerit = "85" & strYearMerit
            Case "4"
               strym03 = "丙等"
               strYearMerit = "60" & strYearMerit
            'ADD BY SONIA 2016/1/5
            Case "*"
               strym03 = "不參加考核"
               strYearMerit = "100" & strYearMerit
            'END 2016/1/5
         End Select
      End If
   Else
      strym03 = "無考績資料"   'modify by sonia 2016/1/5 讀不到之甲等改為無考績
      strYearMerit = "100" & strYearMerit
   End If
   GetYearMerit = True
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Add by SINDY 2009/01/17
'*************************************************
' 考勤與考績
' Input : strEmpID    ==> 員工編號
' Input : strSDate      ==> 起始日期
' Input : strEDate      ==> 終止日期
' Input : strType        ==> 考勤類別 : 0.全部 1.遲到 2.曠職 3.全勤
' Return 分數
'*************************************************
Public Function GetAssistAbsenceGrade(strEmpID As String, strSDate As String, strEDate As String, strType As String) As Double
Dim stSQL As String
Dim adoRst As New ADODB.Recordset
Dim dblHour As Double, dblSACnt As Double
Dim dblGrade0 As Double, dblGrade1 As Double, dblGrade2 As Double, dblGrade3 As Double
Dim j As Double, dblCnt As Double
Dim strStarDate As String, strEndDate As String
   
   '1.遲到            -0.2
   '2.曠職 8小時 -3
   '    未滿8小時 -1.5
   '3.全勤            +3   無請過(05.事假、06.病假、SA04.遲到、SA05+SA06曠職)者,均視為全勤
   
   dblHour = 0
   dblSACnt = 0
   dblGrade0 = 0
   dblGrade1 = 0
   dblGrade2 = 0
   dblGrade3 = 0
   
   GetAssistAbsenceGrade = 0
   '1.遲到 2.曠職
'   stSQL = "SELECT sum(nvl(SA04,0)) as T04,sum(nvl(SA05,0)) as T05,sum(nvl(SA06,0)) as T06 " & _
'                 "From Staff_Assist " & _
'                 "WHERE SA01='" & strEmpID & "' " & _
'                 "AND SA02 >='" & strSDate & "' AND SA02 <='" & strEDate & "' " & _
'                 "GROUP BY SA01 " & _
'                 "ORDER BY SA01 "
'   If adoRst.State = 1 Then adoRst.Close
'   adoRst.CursorLocation = adUseClient
'   adoRst.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
'   If adoRst.RecordCount <> 0 And adoRst.RecordCount > 0 Then
'      dblGrade1 = CheckStr(adoRst.Fields(0)) * -0.2
'      dblHour = (CheckStr(adoRst.Fields(1)) * 8) + CheckStr(adoRst.Fields(2))
'      dblGrade2 = ((dblHour \ 8) * -3)
'      If ((dblHour * 10) Mod (8 * 10)) / 10 <> 0 Then
'         dblGrade2 = dblGrade2 + (-1.5)
'      End If
'   End If
   'Modify By Sindy 2010/10/7
   For j = Val(Left(strSDate, 6)) To Val(Left(strEDate, 6))
      Call Pub_GetSpecWorkHour(strEmpID, strStarDate) 'Modify By Sindy 2013/12/16
      If Left(strSDate, 6) = Left(strEDate, 6) Then
         strStarDate = strSDate
         strEndDate = strEDate
      Else
         If j = Val(Left(strSDate, 6)) Then
            strStarDate = strSDate
         Else
            strStarDate = CStr(j) & "01"
         End If
         If j = Val(Left(strEDate, 6)) Then
            strEndDate = strEDate
         Else
            strEndDate = CStr(j) & "31"
         End If
      End If
      '抓出該月總遲到次數
      stSQL = "SELECT sa01,nvl(sum(nvl(sa04,0)),0),nvl(sum(nvl(sa05,0)),0),nvl(sum(nvl(sa06,0)),0) " & _
                "From staff_Assist " & _
               "Where SA01='" & strEmpID & "' " & _
                 "and sa02 between " & strStarDate & " and " & strEndDate & _
              " group by sa01 "
      If adoRst.State = 1 Then adoRst.Close
      adoRst.CursorLocation = adUseClient
      adoRst.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If adoRst.RecordCount <> 0 And adoRst.RecordCount > 0 Then
         If IsNull(adoRst(1)) Then dblCnt = 0 Else dblCnt = adoRst(1)
         '出缺勤檔 - 遲到
         'Modify By Sindy 2012/12/18 依江總指示,自101年1月起月遲到三次以上視為曠職之規定修正為:
         '                           前二次仍以遲到計算,第三次以後才以曠職計算(每次仍為30分)
         If dblCnt > 0 Then
            '2012年01月之後的新規定
            If Val(Left(strStarDate, 6)) >= 201201 Then
               '02.遲到
               If dblCnt <= 2 Then
                  dblGrade1 = dblGrade1 + dblCnt
               Else
                  dblGrade1 = dblGrade1 + 2 '前2次算遲到,之後全算曠職
               End If
               '03.遲到曠職
               If dblCnt > 2 Then
                  dblHour = dblHour + ((dblCnt - 2) * 0.5)
               End If
            '2012/12/18 End
            Else '2012年01月之前的算法
               If dblCnt <= 2 Then
                  '02.遲到
                  dblGrade1 = dblGrade1 + dblCnt
               Else
                  '03.遲到曠職
                  dblHour = dblHour + (dblCnt * 0.5)
               End If
            End If
         End If
         '曠職
         '99029伊恩一天只上4個小時
         'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
         'Modify By Sindy 2012/7/9 上班時數為特殊者
         If Val(adoRst.Fields(2)) <> 0 Or Val(adoRst.Fields(3)) <> 0 Then
            'Call Pub_GetSpecWorkHour(strEmpID, strStarDate)
   '         If strEmpID = "99029" Then
   '            dblHour = dblHour + (CheckStr(adoRst.Fields(2)) * 5) + CheckStr(adoRst.Fields(3))
            dblHour = dblHour + (Val(adoRst.Fields(2)) * PUB_intWkHour) + Val(adoRst.Fields(3))
         End If
      End If
   Next j
   dblGrade1 = dblGrade1 * -0.2
   '99029伊恩一天只上4個小時
   'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
   'Modify By Sindy 2012/7/9 上班時數為特殊者
'   If strEmpID = "99029" Then
'      dblGrade2 = ((dblHour \ 5) * -3)
'      If ((dblHour * 10) Mod (5 * 10)) / 10 <> 0 Then
'         dblGrade2 = dblGrade2 + (-1.5)
'      End If
   dblGrade2 = ((dblHour \ PUB_intWkHour) * -3)
   If ((dblHour * 10) Mod (PUB_intWkHour * 10)) / 10 <> 0 Then
      dblGrade2 = dblGrade2 + (-1.5)
   End If
      
   If strType = "1" Then
      GetAssistAbsenceGrade = dblGrade1
      Exit Function
   ElseIf strType = "2" Then
      GetAssistAbsenceGrade = dblGrade2
      Exit Function
   End If
   
   '3.全勤
   'Modify By Sindy 2012/1/5 要整年度都在公司者才算全勤,排除本年度新進同事及復職者
   stSQL = "SELECT * From Staff,Staff_Change " & _
            "WHERE ST01='" & strEmpID & "' " & _
              "AND ST01=SC01 " & _
              "AND (ST13>=" & strSDate & " OR (SC02 between " & strSDate & " AND " & strEDate & " AND SC03='02')) "
   If adoRst.State = 1 Then adoRst.Close
   adoRst.CursorLocation = adUseClient
   adoRst.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRst.RecordCount = 0 Then
   '2012/1/5 End
      stSQL = "SELECT Count(*) " & _
                "From Staff_Absence " & _
               "WHERE SA01='" & strEmpID & "' " & _
                 "AND SA06 in ('05','06') " & _
                 "AND SA02 >='" & strSDate & "' AND SA04 <='" & strEDate & "' " & _
              "GROUP BY SA01 " & _
              "ORDER BY SA01 "
      If adoRst.State = 1 Then adoRst.Close
      adoRst.CursorLocation = adUseClient
      adoRst.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If adoRst.RecordCount <> 0 And adoRst.RecordCount > 0 Then
         dblSACnt = CheckStr(adoRst.Fields(0))
      End If
      If dblSACnt = 0 And dblGrade1 = 0 And dblGrade2 = 0 Then
         dblGrade3 = 3
      End If
   End If
   
   If strType = "3" Then
      GetAssistAbsenceGrade = dblGrade3
      Exit Function
   ElseIf strType = "0" Then
      GetAssistAbsenceGrade = dblGrade1 + dblGrade2 + dblGrade3
      Exit Function
   End If
End Function

'Add by SINDY 2009/01/17
'*************************************************
' 獎懲與考績
' Input : strEmpID    ==> 員工編號
' Input : strSDate      ==> 起始日期
' Input : strEDate      ==> 終止日期
' Return 分數
'*************************************************
Public Function GetRewardGrade(strEmpID As String, strSDate As String, strEDate As String) As Double
Dim stSQL As String
Dim adoRst As New ADODB.Recordset
   
   '01.大功 9
   '02.小功 3
   '03.嘉獎 1
   '04.大過 -9
   '05.小過 -3
   '06.申誡 -1
   GetRewardGrade = 0
   '2009/11/30 MODIFY BY SONIA 加SR11次數
   stSQL = "SELECT sum(decode(SR03,'01','+9','02','+3','03','+1','04','-9','05','-3','06','-1')*NVL(SR11,1)) " & _
                 "From staff_reward " & _
                 "WHERE SR01='" & strEmpID & "' " & _
                 "AND SR02 >='" & strSDate & "' AND SR02 <='" & strEDate & "' " & _
                 "GROUP BY SR01 " & _
                 "ORDER BY SR01 "
   If adoRst.State = 1 Then adoRst.Close
   adoRst.CursorLocation = adUseClient
   adoRst.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRst.RecordCount <> 0 And adoRst.RecordCount > 0 Then
      GetRewardGrade = CheckStr(adoRst.Fields(0))
   End If
End Function

'Add by Sindy 2009/01/21
'*************************************************
' 轉換年資格式
' Input : dblNianZi    ==> 年資
' Return 年資(格式化)
'*************************************************
Public Function PUB_ChangeNianZi(dblNianZi As Double) As String
Dim dblYear As Double, dblMonth As Double
   
   If dblNianZi >= 1 Then
       dblYear = (dblNianZi * 100) \ 100
       PUB_ChangeNianZi = dblYear & " 年 "
       'dblMonth = ((dblNianZi * 100) Mod 100) * 0.01
       'PUB_ChangeNianZi = PUB_ChangeNianZi & Trim(Round(dblMonth * 12)) & " 個月"
   Else
       PUB_ChangeNianZi = Trim(Round(dblNianZi * 12)) & " 個月"
   End If
End Function

'Add by Sindy 2009/01/23
'*************************************************
' 任職時間的例外狀況
' Input : strST01    ==> 員工編號
' Input : strDate     ==> 異動日期 (民國日期 ex.86/01/31)
' Return 異動日期
'*************************************************
Function PUB_ScDateWriteDeal(StrST01 As String, strDate As String) As String
   If StrST01 = "79068" And strDate = "86/01/31" Then
      PUB_ScDateWriteDeal = "86/02/01"
   ElseIf StrST01 = "79068" And strDate = "86/12/31" Then
      PUB_ScDateWriteDeal = "87/01/01"
   ElseIf StrST01 = "84052" And strDate = "89/07/31" Then
      PUB_ScDateWriteDeal = "89/08/01"
   Else
      PUB_ScDateWriteDeal = strDate
   End If
End Function

'Add by Sindy 2009/01/23
'*************************************************
' 計算年資Sub
' 依據起迄日期取得天數或(及)年數
' Input : strSDate      ==> 起始日期 (西元日期 ex.20081231)
' Input : strEDate      ==> 截止日期 (西元日期 ex.20081231)
' Output : LngDays   ==> 天數
' Output : intYear      ==> 年數
'*************************************************
Function PUB_NianZiDaysYear(strSDate As String, strEDate As String, LngDays As Long, intYear As Integer)
'   LngDays = 0
'   intYear = 0
   
'   If Right(strSDate, 4) = "0229" Then
'      strSDate = Left(strSDate, 4) & "0228"
'   End If
'   If Right(strEDate, 4) = "0229" Then
'      strEDate = Left(strEDate, 4) & "0228"
'   End If
   
   'StarDate和EndDate同年度
   If Left(strSDate, 4) = Left(strEDate, 4) Then
      'LngDays = LngDays + DateDiff("d", ChangeTStringToTDateString(ChangeWStringToTString(strSDate)), ChangeTStringToTDateString(ChangeWStringToTString(strEDate))) + 1
      LngDays = LngDays + DateDiff("d", ChangeWStringToWDateString(strSDate), ChangeWStringToWDateString(strEDate)) + 1
'      '月份有含2月時
'      If Right(strSDate, 4) <= "0228" And Right(strEDate, 4) >= "0228" Then
'         '檢查是否有遇到2/29, 若有, 則天數+1
'         If PUB_GetMonthDays(Left(strSDate, 4), 2) = 29 Then
'            LngDays = LngDays + 1
'         End If
'      End If
      
   Else
      '算StarDate天數
      'LngDays = LngDays + DateDiff("d", ChangeTStringToTDateString(ChangeWStringToTString(strSDate)), ChangeTStringToTDateString(ChangeWStringToTString(Left(strSDate, 4) & "1231"))) + 1
      LngDays = LngDays + DateDiff("d", ChangeWStringToWDateString(strSDate), ChangeWStringToWDateString(Left(strSDate, 4) & "1231")) + 1
'      '月份有含2月時
'      If Right(strSDate, 4) <= "0228" Then
'         '檢查是否有遇到2/29, 若有, 則天數+1
'         If PUB_GetMonthDays(Left(strSDate, 4), 2) = 29 Then
'            LngDays = LngDays + 1
'         End If
'      End If
      
      '若EndDate為12/31時, 計算年數
      If Right(strEDate, 4) = "1231" Then
         'intYear = intYear + DateDiff("yyyy", ChangeTStringToTDateString(ChangeWStringToTString(CStr(Val(Left(strSDate, 4)) + 1) & "0101")), ChangeTStringToTDateString(ChangeWStringToTString(strEDate))) + 1
         intYear = intYear + DateDiff("yyyy", ChangeWStringToWDateString(CStr(Val(Left(strSDate, 4)) + 1) & "0101"), ChangeWStringToWDateString(strEDate)) + 1
      Else
         '算StarDate和EndDate中間年數
         'intYear = intYear + DateDiff("yyyy", ChangeTStringToTDateString(ChangeWStringToTString(CStr(Val(Left(strSDate, 4)) + 1) & "0101")), ChangeTStringToTDateString(ChangeWStringToTString(CStr(Val(Left(strEDate, 4)) - 1) & "1231"))) + 1
         intYear = intYear + DateDiff("yyyy", ChangeWStringToWDateString(CStr(Val(Left(strSDate, 4)) + 1) & "0101"), ChangeWStringToWDateString(CStr(Val(Left(strEDate, 4)) - 1) & "1231")) + 1
         
         '算EndDate天數
         'LngDays = LngDays + DateDiff("d", ChangeTStringToTDateString(ChangeWStringToTString(Left(strEDate, 4) & "0101")), ChangeTStringToTDateString(ChangeWStringToTString(strEDate))) + 1
         LngDays = LngDays + DateDiff("d", ChangeWStringToWDateString(Left(strEDate, 4) & "0101"), ChangeWStringToWDateString(strEDate)) + 1
'         '月份有含2月時
'         If Right(strEDate, 4) >= "0228" Then
'            '檢查是否有遇到2/29, 若有, 則天數+1
'            If PUB_GetMonthDays(Left(strEDate, 4), 2) = 29 Then
'               LngDays = LngDays + 1
'            End If
'         End If
      End If
   End If
End Function

'Add by Morgan 2010/4/14
'依照投保金額之級距取得勞健保保費計算金額
Public Function PUB_GetInsureBase(pInsureSalary As Long, pKind As String, Optional pRate As Double) As Long
   Dim stSQL As String, intR As Integer, adoRst As ADODB.Recordset
   stSQL = "select si02,si11 from SalaryInsurance" & _
      " where si01='" & pKind & "' and si03<=" & pInsureSalary & " and si04>=" & pInsureSalary
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , True)
   If intR = 1 Then
      PUB_GetInsureBase = Val("" & adoRst.Fields(0))
      pRate = Val("" & adoRst.Fields(1))
   End If
   Set adoRst = Nothing
End Function

'Add by Morgan 2010/4/15
'計算健保費
'健保費 = 健保等級 * 健保費率 * 健保個人負擔比例
Public Function PUB_GetHIFee(lngBase As Long, dblRate As Double, intShareRate As Integer, Optional dblFreeRate As Double) As String
   Dim lngNewFee As Long
   Dim lngOldFee As Long
   Dim intFree As Integer
   
   lngNewFee = Round(lngBase * dblRate / 100 * intShareRate / 100)
   If dblFreeRate = 0 Then
      PUB_GetHIFee = lngNewFee
   Else
      lngOldFee = Round(lngBase * 4.55 / 100 * intShareRate / 100)
      If dblFreeRate = 100 Then
         PUB_GetHIFee = lngOldFee
      Else
         intFree = Round((lngNewFee - lngOldFee) * dblFreeRate / 100)
         PUB_GetHIFee = lngNewFee - intFree
      End If
   End If
End Function

'************************************* Add By Sindy 2011/8/4  *************************************
'************************************* 新增出缺勤相關共用函數 *************************************
'*************************************
'檢查 員工是否離職
'   strST01     員工編號
'*************************************
'Modify By Sindy 2011/10/17 +strDate
Public Function ChkStaffST04(StrST01 As String, Optional bolMsg As Boolean = True, Optional strDate As String = "") As Boolean
   Dim s As Integer
   Dim rsTmp As New ADODB.Recordset
   
   If Trim(StrST01) = "" Then Exit Function 'Add By Sindy 2024/8/27
   CheckOC3
   
   '在職
   strSql = "select st01 from staff where st01='" & StrST01 & "' and st04='1'"
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         ChkStaffST04 = False '在職
      Else
         'Modify By Sindy 2011/10/17 傳入日期與離職日做比較,若大於等於離職日就不可作業
         If Val(strDate) > 0 Then
            '離職
            strSql = "select st01 from staff where st01='" & StrST01 & "' and st04<>'1' and ST51<=" & DBDATE(strDate)
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               ChkStaffST04 = True
               If bolMsg = True Then s = MsgBox("此人員已離職！！", , "人員錯誤！！")
            Else
               ChkStaffST04 = False
            End If
         Else
            ChkStaffST04 = True
            If bolMsg = True Then s = MsgBox("此人員不存在或已離職！！", , "人員錯誤！！")
         End If
      End If
   End With
   CheckOC3
End Function

'*************************************
'檢查 員工不可為”不寄信”
'   strST01     員工編號
'*************************************
Public Function ChkStaffST14(StrST01 As String, Optional bolMsg As Boolean = True) As Boolean
   Dim s As Integer
   
   CheckOC3
   strSql = "select st01 from staff where st01='" & StrST01 & "' and st14='99997' "
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         ChkStaffST14 = True
         If bolMsg = True Then s = MsgBox("此人員無公司MAIL，無法收信！", , "人員錯誤！！")
      Else
         ChkStaffST14 = False
      End If
   End With
   CheckOC3
End Function

'檢查是否為審核主管的身份
Public Function ChkIsAbsBoss(StrST01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
   
   ChkIsAbsBoss = False
   strSql = "SELECT count(*) FROM ABS001 WHERE B0108='" & StrST01 & "' or " & _
                                       "B0109='" & StrST01 & "' or " & _
                                       "B0110='" & StrST01 & "' or " & _
                                       "B0111='" & StrST01 & "' "
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Not IsNull(rsTmp.Fields(0)) Then
         If rsTmp.Fields(0) > 0 Then ChkIsAbsBoss = True
      End If
   End If
End Function

'Add By Sindy 2023/12/19
'取得此審核主管的部門別權限 (改抓 ST93)
'Add By Sindy 2021/12/21 + ByRef strEmp As String : 所屬簽核的人員
Public Function GetIsAbsBossST93(StrST01 As String, ByRef strEmp As String) As String
Dim rsTmp As New ADODB.Recordset
   
   '*************************************************
   'Modify By Sindy 2023/12/19 部門調整改抓 ST93
   '*************************************************
   GetIsAbsBossST93 = "": strEmp = ""
   '抓部門
   'Modify By Sindy 2015/11/2 + AND st04='1'
   strSql = "select distinct ST93 from (SELECT distinct ST93 " & _
            "From ABS001,Staff,Acc090NEW " & _
            "WHERE (B0108='" & StrST01 & "' or B0109='" & StrST01 & "' or B0110='" & StrST01 & "' or B0111='" & StrST01 & "') " & _
            "AND B0101=ST01(+) AND ST93=A0921(+) AND st04='1'"
   'Modify By Sindy 2022/5/3 開放中所林柄佑協理可查詢中所之所有同仁的出缺勤資料，但法律所中所人員除外，除非職代表有設定。
   If StrST01 = "82026" Then '林柄佑
      strSql = strSql & " union select distinct ST93 " & _
                        "from staff where st04='1' and st06='2' and substr(ST93,1,1)<>'L' " & _
                        "and st01>'6' and st01<'F' " & _
                        "AND substr(st01,4,1)<>'9' "
   End If
   'Add By Sindy 2022/7/19
   If InStr(Pub_GetSpecMan("專利處出缺勤可查詢權限"), strUserNum) > 0 Then
      strSql = strSql & " union select distinct ST93 " & _
                        "from staff where st04='1' and substr(st93,1,1)='P' " & _
                        "and st01>'6' and st01<'F' " & _
                        "AND substr(st01,4,1)<>'9' "
   End If
   '2022/7/19 END
   strSql = strSql & ")"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With rsTmp
         .MoveFirst
         Do While Not .EOF
            If Not IsNull(rsTmp.Fields(0)) Then
               'Add By Sindy 2025/11/6
               If InStr(GetIsAbsBossST93, rsTmp.Fields(0)) = 0 Then
               '2025/11/6 END
                  GetIsAbsBossST93 = GetIsAbsBossST93 & "'" & rsTmp.Fields(0) & "',"
               End If
            End If
            .MoveNext
         Loop
      End With
      If GetIsAbsBossST93 <> "" Then
         GetIsAbsBossST93 = Left(GetIsAbsBossST93, Len(GetIsAbsBossST93) - 1)
         'Modify By Sindy 2022/6/23 & ",'" & Pub_StrUserST93 & "'" :  +含自己部門
         'Add By Sindy 2025/11/6
         If InStr(GetIsAbsBossST93, Pub_StrUserSt93) = 0 Then
         '2025/11/6 END
            GetIsAbsBossST93 = GetIsAbsBossST93 & ",'" & Pub_StrUserSt93 & "'"
         End If
      End If
   End If
   'Add By Sindy 2021/12/21
   '抓所屬簽核的人員
   If GetIsAbsBossST93 <> "" Then
      'Modify By Sindy 2022/4/22 + 新人未設定職代表時,部門主管及假日加班簽核主管可以查看
      'Modify By Sindy 2023/1/10 + and substr(ST01,4,1)<>'9'
      'Modify By Sindy 2024/3/1 僅為抓權限可以不用排除: AND (st14<>'99997' or st14 is null)
      strSql = "select distinct st01 from (SELECT distinct ST01 " & _
               "From ABS001,Staff,Acc090NEW " & _
               "WHERE (B0108='" & StrST01 & "' or B0109='" & StrST01 & "' or B0110='" & StrST01 & "' or B0111='" & StrST01 & "') " & _
               "AND B0101=ST01(+) AND ST93=A0921(+) AND st04='1' " & _
               "union SELECT distinct ST01 " & _
               "From ABS001,Staff,Acc090NEW " & _
               "WHERE B0101(+)=ST01 AND ST93=A0921(+) AND st04='1' AND B0101 is null " & _
               "AND ST93 in(" & GetIsAbsBossST93 & ") and (A0924='" & strUserNum & "' or instr(A0928,'" & strUserNum & "')>0) " & _
               "AND st01<'F' and st01>='6' and substr(st01,4,1)<>'9' and st01 not in('60000','96029','96030') " & _
               "AND substr(ST93,1,1)<>'R' " & _
               "and substr(ST01,4,1)<>'9' "
      'Modify By Sindy 2022/5/3 開放中所林柄佑協理可查詢中所之所有同仁的出缺勤資料，但法律所中所人員除外，除非職代表有設定。
      If StrST01 = "82026" Then '林柄佑
         strSql = strSql & " union select distinct ST01 " & _
                           "from staff where st04='1' and st06='2' and substr(ST93,1,1)<>'L' " & _
                           "and st01>'6' and st01<'F' " & _
                           "AND substr(st01,4,1)<>'9' "
      End If
      'Add By Sindy 2022/7/19
      If InStr(Pub_GetSpecMan("專利處出缺勤可查詢權限"), strUserNum) > 0 Then
         strSql = strSql & " union select distinct ST01 " & _
                           "from staff where st04='1' and substr(st93,1,1)='P' " & _
                           "and st01>'6' and st01<'F' " & _
                           "AND substr(st01,4,1)<>'9' "
      End If
      '2022/7/19 END
      strSql = strSql & ")"
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         With rsTmp
            .MoveFirst
            Do While Not .EOF
               If Not IsNull(rsTmp.Fields(0)) Then
                  strEmp = strEmp & "'" & rsTmp.Fields(0) & "',"
               End If
               .MoveNext
            Loop
         End With
         If strEmp <> "" Then
            'Modify By Sindy 2021/12/22 & ",'" & strUserNum & "'" : +含自己
            strEmp = Left(strEmp, Len(strEmp) - 1) & ",'" & strUserNum & "'"
         End If
      End If
   End If
   '2021/12/21 END
End Function

'取得此審核主管的部門別權限
'Add By Sindy 2021/12/21 + ByRef strEmp As String : 所屬簽核的人員
Public Function GetIsAbsBossST03(StrST01 As String, ByRef strEmp As String) As String
Dim rsTmp As New ADODB.Recordset
   
   GetIsAbsBossST03 = "": strEmp = ""
   '抓部門
   'Modify By Sindy 2015/11/2 + AND st04='1'
   strSql = "select distinct st03 from (SELECT distinct ST03 " & _
            "From ABS001,Staff,Acc090 " & _
            "WHERE (B0108='" & StrST01 & "' or B0109='" & StrST01 & "' or B0110='" & StrST01 & "' or B0111='" & StrST01 & "') " & _
            "AND B0101=ST01(+) AND ST03=A0901(+) AND st04='1'"
   'Modify By Sindy 2022/5/3 開放中所林柄佑協理可查詢中所之所有同仁的出缺勤資料，但法律所中所人員除外，除非職代表有設定。
   If StrST01 = "82026" Then '林柄佑
      strSql = strSql & " union select distinct ST03 " & _
                        "from staff where st04='1' and st06='2' and substr(st03,1,1)<>'L' " & _
                        "and st01>'6' and st01<'F' " & _
                        "AND substr(st01,4,1)<>'9' "
   End If
   'Add By Sindy 2022/7/19
   If InStr(Pub_GetSpecMan("專利處出缺勤可查詢權限"), strUserNum) > 0 Then
      strSql = strSql & " union select distinct ST03 " & _
                        "from staff where st04='1' and st03>='P10' and st03<='P14' " & _
                        "and st01>'6' and st01<'F' " & _
                        "AND substr(st01,4,1)<>'9' "
   End If
   '2022/7/19 END
   strSql = strSql & ")"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With rsTmp
         .MoveFirst
         Do While Not .EOF
            If Not IsNull(rsTmp.Fields(0)) Then
               GetIsAbsBossST03 = GetIsAbsBossST03 & "'" & rsTmp.Fields(0) & "',"
            End If
            .MoveNext
         Loop
      End With
      If GetIsAbsBossST03 <> "" Then
         'Modify By Sindy 2022/6/23 & ",'" & Pub_StrUserSt03 & "'" :  +含自己部門
         GetIsAbsBossST03 = Left(GetIsAbsBossST03, Len(GetIsAbsBossST03) - 1) & ",'" & Pub_StrUserSt03 & "'"
      End If
   End If
   'Add By Sindy 2021/12/21
   '抓所屬簽核的人員
   If GetIsAbsBossST03 <> "" Then
      'Modify By Sindy 2022/4/22 + 新人未設定職代表時,部門主管及假日加班簽核主管可以查看
      'Modify By Sindy 2023/1/10 + and substr(ST01,4,1)<>'9'
      'Modify By Sindy 2024/3/1 僅為抓權限可以不用排除: AND (st14<>'99997' or st14 is null)
      strSql = "select distinct st01 from (SELECT distinct ST01 " & _
               "From ABS001,Staff,Acc090 " & _
               "WHERE (B0108='" & StrST01 & "' or B0109='" & StrST01 & "' or B0110='" & StrST01 & "' or B0111='" & StrST01 & "') " & _
               "AND B0101=ST01(+) AND ST03=A0901(+) AND st04='1' " & _
               "union SELECT distinct ST01 " & _
               "From ABS001,Staff,Acc090 " & _
               "WHERE B0101(+)=ST01 AND ST03=A0901(+) AND st04='1' AND B0101 is null " & _
               "AND st03 in(" & GetIsAbsBossST03 & ") and (A0908='" & strUserNum & "' or instr(A0915,'" & strUserNum & "')>0) " & _
               "AND st01<'F' and st01>='6' and substr(st01,4,1)<>'9' and st01 not in('60000','96029','96030') " & _
               "AND substr(st03,1,1)<>'R' " & _
               "and substr(ST01,4,1)<>'9' "
      'Modify By Sindy 2022/5/3 開放中所林柄佑協理可查詢中所之所有同仁的出缺勤資料，但法律所中所人員除外，除非職代表有設定。
      If StrST01 = "82026" Then '林柄佑
         strSql = strSql & " union select distinct ST01 " & _
                           "from staff where st04='1' and st06='2' and substr(st03,1,1)<>'L' " & _
                           "and st01>'6' and st01<'F' " & _
                           "AND substr(st01,4,1)<>'9' "
      End If
      'Add By Sindy 2022/7/19
      If InStr(Pub_GetSpecMan("專利處出缺勤可查詢權限"), strUserNum) > 0 Then
         strSql = strSql & " union select distinct ST01 " & _
                           "from staff where st04='1' and st03>='P10' and st03<='P14' " & _
                           "and st01>'6' and st01<'F' " & _
                           "AND substr(st01,4,1)<>'9' "
      End If
      '2022/7/19 END
      strSql = strSql & ")"
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         With rsTmp
            .MoveFirst
            Do While Not .EOF
               If Not IsNull(rsTmp.Fields(0)) Then
                  strEmp = strEmp & "'" & rsTmp.Fields(0) & "',"
               End If
               .MoveNext
            Loop
         End With
         If strEmp <> "" Then
            'Modify By Sindy 2021/12/22 & ",'" & strUserNum & "'" : +含自己
            strEmp = Left(strEmp, Len(strEmp) - 1) & ",'" & strUserNum & "'"
         End If
      End If
   End If
   '2021/12/21 END
End Function

Public Function SetCboStaffName(strEmpID As String) As String
   SetCboStaffName = Left(Left(strEmpID, 5) & Space(5), 7) & GetPrjSalesNM(Left(strEmpID, 5))
End Function

'設定人事大部門的員工代號下拉式選單
'Modify By Sindy 2023/12/20 +, strA0925 As String
Public Sub SetB1003Combo(ByRef cboTemp As Object, strA0911 As String, strA0925 As String)
Dim Rs As New ADODB.Recordset
   
   cboTemp.Clear
   cboTemp.AddItem ""
   Rs.CursorLocation = adUseClient
   
   'Modify By Sindy 2023/12/19
   If strA0925 <> "" Then
      strSql = "select ST01,ST02 " & _
               "From staff " & _
               "where substr(st01,1,1) in (" & ST01CodeNum1 & ") " & _
               "and st04='1' " & _
               "and substr(st01,4,1)<>'9' " & _
               "and st01 not in('60000','96029','96030') " & _
               "and ST93 in(Select A0921 From ACC090NEW Where A0925='" & strA0925 & "') " & _
               "order by st01 "
   Else
   '2023/12/19 END
      strSql = "select ST01,ST02 " & _
               "From staff " & _
               "where substr(st01,1,1) in (" & ST01CodeNum1 & ") " & _
               "and st04='1' " & _
               "and substr(st01,4,1)<>'9' " & _
               "and st01 not in('60000','96029','96030') " & _
               "and ST03 in(Select A0901 From ACC090 Where A0911='" & strA0911 & "') " & _
               "order by st01 "
   End If
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   While Not Rs.EOF
      cboTemp.AddItem Left(Rs.Fields(0).Value & Space(7), 7) & Rs.Fields(1).Value
      Rs.MoveNext
   Wend
   If Rs.State <> adStateClosed Then Rs.Close
   Set Rs = Nothing
   If cboTemp.ListCount > 0 Then cboTemp.ListIndex = 0
End Sub

'Add By Sindy 2021/8/11
'設定上下班時段的下拉式選單
'intType: 1.上班 2.下班
'Modify By Sindy 2025/10/29 Optional ByVal bolShowAllTime As Boolean = False: 顯示上下班起迄時間
Public Sub SetB102829Combo(ByRef cboTemp As Object, ByVal intType As Integer, _
   Optional ByVal strDate As String = "", Optional ByVal StrST01 As String = "", _
   Optional ByVal bolShowAllTime As Boolean = False)
   
Dim i As Integer

   If Val(strDate) = 0 Then strDate = strSrvDate(1)
   If StrST01 = "" Then StrST01 = strUserNum
   
   Call PUB_ChkByPassWork(PUB_GetST06(StrST01), strDate)
   cboTemp.Clear
   If intType = 1 Then
      For i = 1 To intByPassArea
         cboTemp.AddItem strByPassStarTime(i) & IIf(bolShowAllTime = True, "~" & strByPassEndTime(i), "")
      Next i
   Else
      For i = 1 To intByPassArea
         cboTemp.AddItem strByPassEndTime(i)
      Next i
   End If
End Sub

'設定表單類別的下拉式選單
Public Sub SetB1002Combo(ByRef cboTemp As Object)
   cboTemp.Clear
   cboTemp.AddItem ""
   cboTemp.AddItem "01 請假"
   cboTemp.AddItem "02 加班"
   cboTemp.AddItem "03 出差"
   If cboTemp.ListCount > 0 Then cboTemp.ListIndex = 0
End Sub

'設定假別的下拉式選單
Public Sub SetB1008Combo(ByRef cboTemp As Object)
Dim Rs As New ADODB.Recordset
   
   cboTemp.Clear
   cboTemp.AddItem ""
   Rs.CursorLocation = adUseClient
   strSql = "select * from allcode where ac01='04' and ac02 not in('01','02','03','04','16','17','18') " & _
            "order by ac02 asc "
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   While Not Rs.EOF
      'Modify By Sindy 2020/2/4
      'cboTemp.AddItem Left(rs.Fields("ac02").Value & Space(7), 7) & rs.Fields("ac03").Value
      cboTemp.AddItem Left(Rs.Fields("ac02").Value & Space(4), 4) & Rs.Fields("ac03").Value
      '2020/2/4 END
      Rs.MoveNext
   Wend
   If Rs.State <> adStateClosed Then Rs.Close
   Set Rs = Nothing
   If cboTemp.ListCount > 0 Then cboTemp.ListIndex = 0
End Sub

'讀取表單流程備註檔
Public Sub SetABS012TextBox(ByRef txtTempBox As Object, StrST01 As String)
Dim rsTmp As New ADODB.Recordset
   
   txtTempBox.Text = ""
   'Modify By Sindy 2014/2/24 sqltime(B1205) ==> sqltime(substr('000000'||B1205,-6)) ; ex.09:05: ==> 00:09:05
   'Modify By Sindy 2023/12/29
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "SELECT sqldateT(B1204),sqltime(substr('000000'||B1205,-6)),decode(B1206,'05','',nvl(A0922,ST02))," & B1206CName & ",B1207 FROM ABS012,Staff,Acc090NEW WHERE B1201='" & StrST01 & "' and B1203=ST01(+) and B1203=A0921(+) order by B1202 asc "
   Else
   '2023/12/29 END
      strSql = "SELECT sqldateT(B1204),sqltime(substr('000000'||B1205,-6)),decode(B1206,'05','',nvl(A0902,ST02))," & B1206CName & ",B1207 FROM ABS012,Staff,Acc090 WHERE B1201='" & StrST01 & "' and B1203=ST01(+) and B1203=A0901(+) order by B1202 asc "
   End If
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With rsTmp
         .MoveFirst
         Do While Not .EOF
            If txtTempBox.Text <> "" Then
               txtTempBox.Text = txtTempBox.Text & vbCrLf
               txtTempBox.Text = txtTempBox.Text & "-----------------------------------------------------" & vbCrLf 'Add By Sindy 2022/10/28
            End If
            txtTempBox.Text = txtTempBox.Text & .Fields(0) & " " & _
                                                .Fields(1) & " " & _
                                                IIf(IsNull(.Fields(2)), "", .Fields(2) & " ") & _
                                                IIf(IsNull(.Fields(3)), "", .Fields(3) & " ") & _
                                                IIf(Not IsNull(.Fields(4)) And .Fields(4) > "", IIf(IsNull(.Fields(2)) And IsNull(.Fields(3)), "", "："), "") & _
                                                .Fields(4)
            .MoveNext
         Loop
      End With
   End If
End Sub

'取得表單類別欄位值
Public Function GetB1002Value(strTemp As String) As String
   Select Case strTemp
      Case "01"
         GetB1002Value = "01 請假"
      Case "02"
         GetB1002Value = "02 加班"
      Case "03"
         GetB1002Value = "03 出差"
   End Select
End Function

'顯示目前特別假的休假狀況
'Modify By Sindy 2012/4/27 +intType
'intType=0：@今年特別假：" & dST40 & "天，已休" & dblDay & "天
'intType=1：已休天數
'intType=2：未休天數
'Modify By Sindy 2014/1/3 +strYear
Public Function GetCurrSpecRestDay(StrST01 As String, Optional intType As Integer = 0, Optional ByVal strYear As String = "") As String
Dim dST40 As Double
Dim strSDate As String, strEDate As String
Dim dblHour As Double, dblDay As Double, dblTmpDay As Double
Dim rsTmp As New ADODB.Recordset
   
   GetCurrSpecRestDay = ""
   
   If strYear <> "" Then
      If Len(strYear) = 3 Then
         strYear = CStr(Val(strYear) + 1911)
      End If
      strSDate = strYear & "0101"
      strEDate = strYear & "1231"
   Else
      strSDate = Left(strSrvDate(1), 4) & "0101"
      strEDate = Left(strSrvDate(1), 4) & "1231"
   End If
   
   Call Pub_GetSpecWorkHour(StrST01, strSrvDate(1)) 'Add By Sindy 2017/11/3
   '取得目前已累計的特別假時數
   strSql = "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),sa01 From staff_Absence Where sa01='" & StrST01 & "' and SA06='08' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') group by sa01 "
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      '特別假只能請整日,因此直接Sum天數引用
      If IsNull(rsTmp.Fields(0)) Then dblDay = 0 Else dblDay = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) Then dblHour = 0 Else dblHour = rsTmp.Fields(1)
      'Modify By Sindy 2017/11/3 107年1月1日起開始特別假可以請半天(4小時)
      If Fix(dblHour / PUB_intWkHour) > 0 Then
         dblDay = dblDay + Fix(dblHour / PUB_intWkHour)
         dblHour = dblHour Mod PUB_intWkHour
      End If
      '2017/11/3 END
   End If
   
   '取得可休特別假
   'Modify By Sindy 2014/1/3
   'strSql = "select st40 from staff where st01='" & strST01 & "'"
   If strYear <> "" Then
      strSql = "select yv04 from YearVacation where yv01=" & strYear & " and yv02='" & StrST01 & "'"
   Else
      strSql = "select yv04 from YearVacation where yv01=" & Left(strSrvDate(1), 4) & " and yv02='" & StrST01 & "'"
   End If
   '2014/1/3 END
   intI = 1
   dST40 = 0
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Not IsNull(rsTmp.Fields(0).Value) Then
         dST40 = rsTmp.Fields(0).Value
      End If
   'Add By Sindy 2014/12/3
   Else
      If intType = 0 Then
         'Add By Sindy 2014/12/8
         If strYear = "" Then
            GetCurrSpecRestDay = "@未設定特別假天數"
         Else
         '2014/12/8 END
            GetCurrSpecRestDay = "@" & strYear - 1911 & "年未設定特別假天數"
         End If
         Exit Function
      End If
   '2014/12/3 END
   End If
   If intType = 0 Then
      'Modify By Sindy 2014/12/3
      'GetCurrSpecRestDay = "@今年特別假：" & dST40 & "天，已休" & dblDay & "天"
      If strYear <> "" And strYear <> Left(strSrvDate(1), 4) Then
         'Modify By Sindy 2017/11/3
         GetCurrSpecRestDay = "@" & strYear - 1911 & "年特別假：" & dST40 & "天，已休" & dblDay & "天" & IIf(dblHour > 0, dblHour & "小時", "")
      Else
         'Modify By Sindy 2017/11/3
         GetCurrSpecRestDay = "@今年特別假：" & dST40 & "天，已休" & dblDay & "天" & IIf(dblHour > 0, dblHour & "小時", "")
      End If
      '2014/12/3 END
   '已休天數
   ElseIf intType = 1 Then
      'Modify By Sindy 2017/11/3
      'GetCurrSpecRestDay = dblDay
      GetCurrSpecRestDay = dblDay & "天" & IIf(dblHour > 0, dblHour & "小時", "")
   '未休天數
   ElseIf intType = 2 Then
      'Modify By Sindy 2017/11/3
      'GetCurrSpecRestDay = Val(dST40) - Val(dblDay)
      dblTmpDay = ((Val(dST40) * PUB_intWkHour) - (Val(dblDay) * PUB_intWkHour + dblHour)) / PUB_intWkHour
      GetCurrSpecRestDay = IIf(Fix(dblTmpDay) > 0, Fix(dblTmpDay) & "天", "") & IIf(dblTmpDay - Fix(dblTmpDay) > 0, (dblTmpDay - Fix(dblTmpDay)) * PUB_intWkHour & "小時", "")
   End If
End Function

'Add By Sindy 2024/12/10
'顯示目前補休假的休假狀況
'strST01：員工編號
'intType=0：@可補休：剩餘" & dblDay & "天
'        2：未休天數
'strCountStartDate：到期起始日期
'strFirstSRR01：取得要計算的發生日期(最早)
'dblTotSRR03：計算區間中的可補休時數
'dblTmpHour：未休時數
'strCountEndDate：到期截止日期
'dblRestHour：已休時數
'strCanRestSRR01：取得尚有可補休(最早)發生日期
Public Function GetCurrFor14RestDay(ByVal StrST01 As String, Optional ByVal intType As Integer = 0, _
   Optional ByVal strCountStartDate As String = "", _
   Optional ByRef strFirstSRR01 As String, Optional ByRef dblTotSRR03 As Double = 0, _
   Optional ByRef dblTmpHour As Double = 0, Optional ByVal strCountEndDate As String = "", _
   Optional ByRef dblRestHour As Double = 0, Optional ByRef strCanRestSRR01 As String = "") As String
   
Dim strSDate As String, strEDate As String
Dim dblHour As Double, dblDay As Double, dblTmpDay As Double
Dim rsTmp As New ADODB.Recordset
Dim intDo As Integer
Dim dblCountHour As Double
   
   GetCurrFor14RestDay = ""
   strFirstSRR01 = "": dblTotSRR03 = 0: dblTmpDay = 0: dblTmpHour = 0
   strCanRestSRR01 = ""
   
   '到期起始日期
   If strCountStartDate <> "" Then
      strSDate = DBDATE(strCountStartDate)
   Else
      strSDate = strSrvDate(1)
   End If
   '到期截止日期
   If strCountEndDate <> "" Then '有傳入此值,目前為檢查是否過期用
      strEDate = DBDATE(strCountEndDate)
   Else
      strEDate = DBDATE(DateAdd("m", 12, Format(strSDate, "####/##/##")))
      If Left(strEDate, 4) < Left(strSrvDate(1), 4) Then
         GetCurrFor14RestDay = "": Exit Function
      ElseIf Left(strEDate, 4) = Left(strSrvDate(1), 4) Then
         strEDate = Left(strEDate, 4) & "1231"
      End If
   End If
   Call Pub_GetSpecWorkHour(StrST01, strSDate)
   
   '取得要計算的可補休發生日期:
   '依可補休資料
   strSql = "select * from Staff_RepayRest" & _
            " where SRR05>=" & strSDate & " and SRR05<=" & strEDate & " and SRR02='" & StrST01 & "'" & _
            " order by SRR01 asc,SRR05 asc"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      rsTmp.MoveFirst
      strFirstSRR01 = rsTmp.Fields("SRR01").Value '要計算的發生日期
      '抓出相關的補休區間
      intDo = 1
      Do While intDo > 0
'         strSql = "select * From staff_Absence Where sa01='" & strST01 & "'" & _
'                  " and SA06='14' and (sa02 >= '" & strSRR01 & "' and sa04 <= '" & strEDate & "')" & _
'                  " order by sa02 asc,sa04 asc"
'         intDo = 1
'         Set rsTmp = ClsLawReadRstMsg(intDo, strSql)
'         If intDo = 1 Then
'            rsTmp.MoveFirst
            '再取得有效的可補休發生日期
            strSql = "select * from Staff_RepayRest" & _
                     " where " & strFirstSRR01 & " between SRR01 and SRR05" & _
                     " and SRR02='" & StrST01 & "'" & _
                     " order by SRR01 asc,SRR05 asc"
            intDo = 1
            Set rsTmp = ClsLawReadRstMsg(intDo, strSql)
            If intDo = 1 Then
               rsTmp.MoveFirst
               If rsTmp.Fields("SRR01").Value < strFirstSRR01 Then
                  strFirstSRR01 = rsTmp.Fields("SRR01").Value '要計算的發生日期
               Else
                  intDo = 0
               End If
            Else
               intDo = 0
            End If
'         End If
      Loop
   End If
   
   If strFirstSRR01 <> "" Then
      '依可補休發生日期起,抓有請補休的時數
      strSql = "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),sa01 From staff_Absence Where sa01='" & StrST01 & "'" & _
               " and SA06='14' and (sa02 >= '" & strFirstSRR01 & "' and sa04 <= '" & strEDate & "') group by sa01 "
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If IsNull(rsTmp.Fields(0)) Then dblDay = 0 Else dblDay = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) Then dblHour = 0 Else dblHour = rsTmp.Fields(1)
         If Fix(dblHour / PUB_intWkHour) > 0 Then
            'Modify By Sindy 2025/5/26
            'dblDay = dblDay + Fix(dblHour / PUB_intWkHour)
            'dblHour = dblHour Mod PUB_intWkHour
            strExc(10) = Fix(dblHour / PUB_intWkHour)
            dblDay = dblDay + Val(strExc(10))
            dblHour = dblHour - (Val(strExc(10)) * PUB_intWkHour)
            '2025/5/26 END
         End If
         '已休時數
         dblRestHour = Val(dblDay) * PUB_intWkHour + dblHour
      End If
      
      '取得可補休時數
      strSql = "select * from Staff_RepayRest" & _
               " where SRR01>=" & strFirstSRR01 & " and SRR05<=" & strEDate & " and SRR02='" & StrST01 & "'" & _
               " order by SRR05 asc,SRR01 asc"
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         rsTmp.MoveFirst
         strCanRestSRR01 = rsTmp.Fields("SRR01") '取得尚有可補休(最早)發生日期
         Do While Not rsTmp.EOF
            If Not IsNull(rsTmp.Fields("SRR03").Value) Then
               dblTotSRR03 = dblTotSRR03 + rsTmp.Fields("SRR03").Value
            End If
            '過期時數
            If Not IsNull(rsTmp.Fields("SRR12").Value) Then
               dblTotSRR03 = dblTotSRR03 - rsTmp.Fields("SRR12").Value
            End If
            rsTmp.MoveNext
         Loop
         '未休時數
         dblTmpHour = dblTotSRR03 - dblRestHour
         If dblTmpHour > 0 Then
            'Add By Sindy 2025/2/12 取得尚有可補休(最早)發生日期
            If dblRestHour > 0 Then
               'If dblTmpHour > 0 Then
                  dblCountHour = dblRestHour
                  rsTmp.MoveFirst
                  Do While Not rsTmp.EOF
                     dblCountHour = dblCountHour - (rsTmp.Fields("SRR03").Value - Val("" & rsTmp.Fields("SRR12").Value))
                     strCanRestSRR01 = rsTmp.Fields("SRR01")
                     If dblCountHour = 0 Then
                        rsTmp.MoveNext '再抓下一筆
                        strCanRestSRR01 = rsTmp.Fields("SRR01")
                        Exit Do
                     ElseIf dblCountHour < 0 Then
                        Exit Do
                     End If
                     rsTmp.MoveNext
                  Loop
               'End If
            End If
            '2025/2/12 END
         Else
            strCanRestSRR01 = ""
         End If
      End If
      
      '未休天數
      dblTmpDay = dblTmpHour / PUB_intWkHour
      If dblTmpDay > 0 Then
         GetCurrFor14RestDay = "剩餘 " & IIf(Fix(dblTmpDay) > 0, Fix(dblTmpDay) & "天", "") & IIf(dblTmpDay - Fix(dblTmpDay) > 0, (dblTmpDay - Fix(dblTmpDay)) * PUB_intWkHour & "小時", "")
      Else
         GetCurrFor14RestDay = "剩餘 0 天"
      End If
      If intType = 0 Then
         GetCurrFor14RestDay = "@可補休：" & GetCurrFor14RestDay
      End If
   Else
      If intType = 0 Then
         GetCurrFor14RestDay = "@無可補休時數"
      Else
         GetCurrFor14RestDay = ""
      End If
   End If
End Function

'出缺勤電子簽核-表單編號(電腦自動給號)
Public Function AutoNo_ABS(InputItem As String, InputLength As Integer) As String
Dim adoaccnum As New ADODB.Recordset
Dim strItem As String, strYes As String

''911106 NICK '911106 nick 避免相同連線作做2次 transation
'Dim BolTransOk As Boolean
'BolTransOk = True
'On Error GoTo TransErr

'   adoTaie.BeginTrans
   adoTaie.Execute "update autonumber set au03 = au03 where au01 = '" & InputItem & "'"
   If Trim(InputItem) = "ABS" Then
      strItem = ""
   End If
   adoaccnum.CursorLocation = adUseClient
   adoaccnum.Open "select * from autonumber where au01 = '" & InputItem & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccnum.RecordCount = 0 Then
      AutoNo_ABS = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo("0", InputLength)
   Else
      If adoaccnum.Fields("au02").Value <> Val(Mid(strSrvDate(1), 1, 4)) Then
         AutoNo_ABS = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo("0", InputLength)
      Else
         AutoNo_ABS = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo(str(adoaccnum.Fields("au03").Value), InputLength)
      End If
   End If
   strYes = SaveAutoNo(InputItem, Mid(AutoNo_ABS, 4, InputLength))
   adoaccnum.Close
   
'   If BolTransOk Then
'      adoTaie.CommitTrans
'   End If
''911106 nick 避免相同連線作做2次 transation
'   Exit Function
'
'TransErr:
'   If Err.Number = -2147168237 Then
'      BolTransOk = False
'      Resume Next
'   End If
End Function

'取得Insert表單流程備註Sql
Public Function GetInsertABS012Sql(strB1201 As String, strB1203 As String, strB1204 As String, strB1205 As String, strB1206 As String, strB1207 As String) As String
Dim intSeqno As Integer
Dim rsTmp As New ADODB.Recordset
   
   strSql = "select * from ABS012 where B1201='" & strB1201 & "' "
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strSql = "select max(B1202) from ABS012 where B1201='" & strB1201 & "' "
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Not IsNull(rsTmp.Fields(0).Value) Then intSeqno = rsTmp.Fields(0).Value
      End If
   End If
   GetInsertABS012Sql = "insert into ABS012 (B1201,B1202,B1203,B1204,B1205,B1206,B1207) " & _
                        "values(" & CNULL(strB1201) & "," & (intSeqno + 1) & "," & _
                        CNULL(strB1203) & "," & strB1204 & "," & strB1205 & "," & _
                        CNULL(strB1206) & "," & CNULL(strB1207) & ")"
End Function

'人事取得職務代理人
'Modify By Amy 2017/02/07 + bolOnlyOne 是否只取職代(1)
'Modify By Sindy 2017/8/22 + , Optional ByRef DutyAorB As String = ""
'    DutyAorB : 傳入 職代人員ID; 回傳 雙職代的A區或B區人員資料
Public Sub GetABS001_1(StrST01 As String, ByRef strABS001_1 As String, ByRef strABS001_2 As String, ByRef strABS001_3 As String, _
   Optional ByVal bolOnlyOne As Boolean = False, Optional ByRef DutyAorB As String = "")
Dim rsTmp As New ADODB.Recordset
   
   strABS001_1 = ""
   strABS001_2 = ""
   strABS001_3 = ""
   strSql = "select * from abs001 where b0101=" & CNULL(StrST01)
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Not IsNull(rsTmp.Fields("B0102")) Or Not IsNull(rsTmp.Fields("B0103")) Then
         If Not IsNull(rsTmp.Fields("B0102")) Then
            If ChkStaffST04(rsTmp.Fields("B0102"), False) = False Then strABS001_1 = strABS001_1 & rsTmp.Fields("B0102") & ","
         End If
         'Modify by Amy 2017/02/07
         If bolOnlyOne = False Then
            If Not IsNull(rsTmp.Fields("B0103")) Then
               If ChkStaffST04(rsTmp.Fields("B0103"), False) = False Then strABS001_1 = strABS001_1 & rsTmp.Fields("B0103") & ","
            End If
         End If
         If strABS001_1 <> "" Then strABS001_1 = Left(strABS001_1, Len(strABS001_1) - 1)
      End If
      If Not IsNull(rsTmp.Fields("B0104")) Or Not IsNull(rsTmp.Fields("B0105")) Then
         If Not IsNull(rsTmp.Fields("B0104")) Then
            If ChkStaffST04(rsTmp.Fields("B0104"), False) = False Then strABS001_2 = strABS001_2 & rsTmp.Fields("B0104") & ","
         End If
         'Modify by Amy 2017/02/07
         If bolOnlyOne = False Then
            If Not IsNull(rsTmp.Fields("B0105")) Then
               If ChkStaffST04(rsTmp.Fields("B0105"), False) = False Then strABS001_2 = strABS001_2 & rsTmp.Fields("B0105") & ","
            End If
         End If
         If strABS001_2 <> "" Then strABS001_2 = Left(strABS001_2, Len(strABS001_2) - 1)
      End If
      If Not IsNull(rsTmp.Fields("B0106")) Or Not IsNull(rsTmp.Fields("B0107")) Then
         If Not IsNull(rsTmp.Fields("B0106")) Then
            If ChkStaffST04(rsTmp.Fields("B0106"), False) = False Then strABS001_3 = strABS001_3 & rsTmp.Fields("B0106") & ","
         End If
         'Modify by Amy 2017/02/07
         If bolOnlyOne = False Then
            If Not IsNull(rsTmp.Fields("B0107")) Then
               If ChkStaffST04(rsTmp.Fields("B0107"), False) = False Then strABS001_3 = strABS001_3 & rsTmp.Fields("B0107") & ","
            End If
         End If
         If strABS001_3 <> "" Then strABS001_3 = Left(strABS001_3, Len(strABS001_3) - 1)
      End If
      'Add By Sindy 2017/8/21 設定職代組數
      Call SetPubABS001(strABS001_1, strABS001_2, strABS001_3, DutyAorB)
   End If
End Sub

'案件取得職務代理人
'Modify By Sindy 2017/8/22 + , Optional ByRef DutyAorB As String = ""
'    DutyAorB : 傳入 職代人員ID; 回傳 雙職代的A區或B區人員資料
Public Sub GetABS001_3(StrST01 As String, ByRef strABS001_1 As String, ByRef strABS001_2 As String, ByRef strABS001_3 As String, _
   strKind As String, Optional ByRef DutyAorB As String = "")
Dim rsTmp As New ADODB.Recordset
   
   strABS001_1 = ""
   strABS001_2 = ""
   strABS001_3 = ""
   strSql = "select * from abs001 where b0101=" & CNULL(StrST01)
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Not IsNull(rsTmp.Fields("B0117")) Or Not IsNull(rsTmp.Fields("B0119")) Then
         If Not IsNull(rsTmp.Fields("B0117")) Then
            If "" & rsTmp.Fields("B0116") = "" Or "" & rsTmp.Fields("B0116") = strKind Or strKind = "" Then
               If ChkStaffST04(rsTmp.Fields("B0117"), False) = False Then strABS001_1 = strABS001_1 & rsTmp.Fields("B0117") & ","
            End If
         End If
         If Not IsNull(rsTmp.Fields("B0119")) Then
            If "" & rsTmp.Fields("B0118") = "" Or "" & rsTmp.Fields("B0118") = strKind Or strKind = "" Then
               If ChkStaffST04(rsTmp.Fields("B0119"), False) = False Then strABS001_1 = strABS001_1 & rsTmp.Fields("B0119") & ","
            End If
         End If
         If strABS001_1 <> "" Then strABS001_1 = Left(strABS001_1, Len(strABS001_1) - 1)
      End If
      If Not IsNull(rsTmp.Fields("B0121")) Or Not IsNull(rsTmp.Fields("B0123")) Then
         If Not IsNull(rsTmp.Fields("B0121")) Then
            If "" & rsTmp.Fields("B0120") = "" Or "" & rsTmp.Fields("B0120") = strKind Or strKind = "" Then
               If ChkStaffST04(rsTmp.Fields("B0121"), False) = False Then strABS001_2 = strABS001_2 & rsTmp.Fields("B0121") & ","
            End If
         End If
         If Not IsNull(rsTmp.Fields("B0123")) Then
            If "" & rsTmp.Fields("B0122") = "" Or "" & rsTmp.Fields("B0122") = strKind Or strKind = "" Then
               If ChkStaffST04(rsTmp.Fields("B0123"), False) = False Then strABS001_2 = strABS001_2 & rsTmp.Fields("B0123") & ","
            End If
         End If
         If strABS001_2 <> "" Then strABS001_2 = Left(strABS001_2, Len(strABS001_2) - 1)
      End If
      'Add By Sindy 2017/8/21 設定職代組數
      Call SetPubABS001(strABS001_1, strABS001_2, strABS001_3, DutyAorB)
   End If
End Sub

'Add By Sindy 2017/8/21 設定職代組數
Private Sub SetPubABS001(strABS001_1 As String, strABS001_2 As String, strABS001_3 As String, _
   Optional ByRef DutyAorB As String = "")
Dim intItem As Integer, strData As String, varTemp As Variant, varTemp2 As Variant
Dim i As Integer, j As Integer, k As Integer, h As Integer
Dim strEmp As String
   
   '清除職代組數
   intItem = 0
   For k = 1 To intDutyItem
      PubABS001_1(k) = ""
   Next k
   PubABS001_A = "": PubABS001_B = "" '雙職代的A,B區
   strData = ""
   '設定職代組數
   For i = 1 To 3
      strEmp = ""
      If i = 1 Then strEmp = strABS001_1
      If i = 2 Then strEmp = strABS001_2
      If i = 3 Then strEmp = strABS001_3
      If strEmp <> "" Then
         varTemp = Split(strEmp, ",")
         PubABS001_A = PubABS001_A & varTemp(0) & ","
         For j = 0 To UBound(varTemp)
            If varTemp(j) <> "" Then
               If InStr(strData, varTemp(j)) = 0 Then
                  strData = strData & varTemp(j) & ","
               End If
            End If
         Next j
         For k = 1 To 3
            strEmp = ""
            If k = 1 And InStr(strABS001_1, ",") > 0 Then strEmp = strABS001_1
            If k = 2 And InStr(strABS001_2, ",") > 0 Then strEmp = strABS001_2
            If k = 3 And InStr(strABS001_3, ",") > 0 Then strEmp = strABS001_3
            If strEmp <> "" Then
               varTemp2 = Split(strEmp, ",")
               If i = k Then
                  PubABS001_B = PubABS001_B & varTemp2(1) & ","
               End If
               intItem = intItem + 1
               PubABS001_1(intItem) = PubABS001_1(intItem) & varTemp(0) & ","
               PubABS001_1(intItem) = PubABS001_1(intItem) & varTemp2(1) & ","
            Else
               If i = k Then
                  intItem = intItem + 1
                  PubABS001_1(intItem) = PubABS001_1(intItem) & varTemp(0) & ","
               End If
            End If
         Next k
      End If
   Next i
   If InStr(strABS001_1, ",") > 0 Or InStr(strABS001_2, ",") > 0 Or InStr(strABS001_3, ",") > 0 Then
      If strData <> "" Then
         strData = Left(strData, Len(strData) - 1)
         varTemp = Split(strData, ",")
         For k = 0 To UBound(varTemp)
            intItem = intItem + 1
            PubABS001_1(intItem) = PubABS001_1(intItem) & varTemp(k) & ","
         Next k
      End If
   End If
   For k = 1 To intItem
      If PubABS001_1(k) <> "" Then
         PubABS001_1(k) = Left(PubABS001_1(k), Len(PubABS001_1(k)) - 1)
      End If
   Next k
   If PubABS001_A <> "" Then PubABS001_A = Left(PubABS001_A, Len(PubABS001_A) - 1)
   If PubABS001_B <> "" Then PubABS001_B = Left(PubABS001_B, Len(PubABS001_B) - 1)
   '回傳 雙職代的A區或B區人員資料
   If DutyAorB <> "" Then
      If InStr(PubABS001_A, DutyAorB) > 0 And InStr(PubABS001_B, DutyAorB) = 0 Then
         DutyAorB = PubABS001_A
      ElseIf InStr(PubABS001_A, DutyAorB) = 0 And InStr(PubABS001_B, DutyAorB) > 0 Then
         DutyAorB = PubABS001_B
      ElseIf InStr(PubABS001_A, DutyAorB) >= InStr(PubABS001_B, DutyAorB) Then
         DutyAorB = PubABS001_A
      Else
         DutyAorB = PubABS001_B
      End If
   End If
End Sub

'Add By Sindy 2014/5/6
'案件系統取得職務代理人
'strB0124:請假、出差核准後通知人員編號 'Add by Amy 2020/01/11
Public Sub GetABS001_CaseSys(StrST01 As String, ByRef strABS001_1 As String, ByRef strABS001_2 As String, ByRef strABS001_3 As String, Optional ByRef strB0124 As String)
Dim rsTmp As New ADODB.Recordset
   
   strABS001_1 = ""
   strABS001_2 = ""
   strABS001_3 = ""
   strB0124 = "" 'Add by Amy 2020/01/11 請假、出差核准後通知人員編號(多筆;區隔)
   strSql = "select * from abs001 where b0101=" & CNULL(StrST01)
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      '案件職代優先
      If Not IsNull(rsTmp.Fields("B0117")) Then
         If Not IsNull(rsTmp.Fields("B0117")) Or Not IsNull(rsTmp.Fields("B0119")) Then
            If Not IsNull(rsTmp.Fields("B0117")) Then
               If ChkStaffST04(rsTmp.Fields("B0117"), False) = False Then strABS001_1 = strABS001_1 & rsTmp.Fields("B0117") & ","
            End If
            If Not IsNull(rsTmp.Fields("B0119")) Then
               If ChkStaffST04(rsTmp.Fields("B0119"), False) = False Then strABS001_1 = strABS001_1 & rsTmp.Fields("B0119") & ","
            End If
            If strABS001_1 <> "" Then strABS001_1 = Left(strABS001_1, Len(strABS001_1) - 1)
         End If
         If Not IsNull(rsTmp.Fields("B0121")) Or Not IsNull(rsTmp.Fields("B0123")) Then
            If Not IsNull(rsTmp.Fields("B0121")) Then
               If ChkStaffST04(rsTmp.Fields("B0121"), False) = False Then strABS001_2 = strABS001_2 & rsTmp.Fields("B0121") & ","
            End If
            If Not IsNull(rsTmp.Fields("B0123")) Then
               If ChkStaffST04(rsTmp.Fields("B0123"), False) = False Then strABS001_2 = strABS001_2 & rsTmp.Fields("B0123") & ","
            End If
            If strABS001_2 <> "" Then strABS001_2 = Left(strABS001_2, Len(strABS001_2) - 1)
         End If
      Else
         '無案件職代,才讀取人事職代
         If Not IsNull(rsTmp.Fields("B0102")) Or Not IsNull(rsTmp.Fields("B0103")) Then
            If Not IsNull(rsTmp.Fields("B0102")) Then
               If ChkStaffST04(rsTmp.Fields("B0102"), False) = False Then strABS001_1 = strABS001_1 & rsTmp.Fields("B0102") & ","
            End If
            If Not IsNull(rsTmp.Fields("B0103")) Then
               If ChkStaffST04(rsTmp.Fields("B0103"), False) = False Then strABS001_1 = strABS001_1 & rsTmp.Fields("B0103") & ","
            End If
            If strABS001_1 <> "" Then strABS001_1 = Left(strABS001_1, Len(strABS001_1) - 1)
         End If
         If Not IsNull(rsTmp.Fields("B0104")) Or Not IsNull(rsTmp.Fields("B0105")) Then
            If Not IsNull(rsTmp.Fields("B0104")) Then
               If ChkStaffST04(rsTmp.Fields("B0104"), False) = False Then strABS001_2 = strABS001_2 & rsTmp.Fields("B0104") & ","
            End If
            If Not IsNull(rsTmp.Fields("B0105")) Then
               If ChkStaffST04(rsTmp.Fields("B0105"), False) = False Then strABS001_2 = strABS001_2 & rsTmp.Fields("B0105") & ","
            End If
            If strABS001_2 <> "" Then strABS001_2 = Left(strABS001_2, Len(strABS001_2) - 1)
         End If
         If Not IsNull(rsTmp.Fields("B0106")) Or Not IsNull(rsTmp.Fields("B0107")) Then
            If Not IsNull(rsTmp.Fields("B0106")) Then
               If ChkStaffST04(rsTmp.Fields("B0106"), False) = False Then strABS001_3 = strABS001_3 & rsTmp.Fields("B0106") & ","
            End If
            If Not IsNull(rsTmp.Fields("B0107")) Then
               If ChkStaffST04(rsTmp.Fields("B0107"), False) = False Then strABS001_3 = strABS001_3 & rsTmp.Fields("B0107") & ","
            End If
            If strABS001_3 <> "" Then strABS001_3 = Left(strABS001_3, Len(strABS001_3) - 1)
         End If
      End If
      strB0124 = "" & rsTmp.Fields("B0124") 'Add by Amy 2021/01/11 +請假、出差核准後通知人員編號
   End If
End Sub

'取得審核主管
'增加判斷天數...以上(不含)
'Modify By Sindy 2014/2/17 dblDay As Double ==> Optional dblDay As Double = 0
Public Function GetABS001_2(StrST01 As String, Optional dblDay As Double = 0) As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "select * from abs001 where b0101=" & CNULL(StrST01)
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Not IsNull(rsTmp.Fields("B0108")) Then
         If ChkStaffST04(rsTmp.Fields("B0108"), False) = False Then
            'Modify By Sindy 2022/11/30 排除設定99天的主管
            If Not IsNull(rsTmp.Fields("B0112")) And Val("" & rsTmp.Fields("B0112")) = 99 Then
            Else
            '2022/11/30 END
               If Not IsNull(rsTmp.Fields("B0112")) And dblDay <> 0 Then
                  If dblDay > Val(rsTmp.Fields("B0112")) Then GetABS001_2 = GetABS001_2 & rsTmp.Fields("B0108") & ","
               Else
                  GetABS001_2 = GetABS001_2 & rsTmp.Fields("B0108") & ","
               End If
            End If
         End If
      End If
      If Not IsNull(rsTmp.Fields("B0109")) Then
         If ChkStaffST04(rsTmp.Fields("B0109"), False) = False Then
            'Modify By Sindy 2022/11/30 排除設定99天的主管
            If Not IsNull(rsTmp.Fields("B0113")) And Val("" & rsTmp.Fields("B0113")) = 99 Then
            Else
            '2022/11/30 END
               If Not IsNull(rsTmp.Fields("B0113")) And dblDay <> 0 Then
                  If dblDay > Val(rsTmp.Fields("B0113")) Then GetABS001_2 = GetABS001_2 & rsTmp.Fields("B0109") & ","
               Else
                  GetABS001_2 = GetABS001_2 & rsTmp.Fields("B0109") & ","
               End If
            End If
         End If
      End If
      If Not IsNull(rsTmp.Fields("B0110")) Then
         If ChkStaffST04(rsTmp.Fields("B0110"), False) = False Then
            'Modify By Sindy 2022/11/30 排除設定99天的主管
            If Not IsNull(rsTmp.Fields("B0114")) And Val("" & rsTmp.Fields("B0114")) = 99 Then
            Else
            '2022/11/30 END
               If Not IsNull(rsTmp.Fields("B0114")) And dblDay <> 0 Then
                  If dblDay > Val(rsTmp.Fields("B0114")) Then GetABS001_2 = GetABS001_2 & rsTmp.Fields("B0110") & ","
               Else
                  GetABS001_2 = GetABS001_2 & rsTmp.Fields("B0110") & ","
               End If
            End If
         End If
      End If
      If Not IsNull(rsTmp.Fields("B0111")) Then
         If ChkStaffST04(rsTmp.Fields("B0111"), False) = False Then
            'Modify By Sindy 2022/11/30 排除設定99天的主管
            If Not IsNull(rsTmp.Fields("B0115")) And Val("" & rsTmp.Fields("B0115")) = 99 Then
            Else
            '2022/11/30 END
               If Not IsNull(rsTmp.Fields("B0115")) And dblDay <> 0 Then
                  If dblDay > Val(rsTmp.Fields("B0115")) Then GetABS001_2 = GetABS001_2 & rsTmp.Fields("B0111") & ","
               Else
                  GetABS001_2 = GetABS001_2 & rsTmp.Fields("B0111") & ","
               End If
            End If
         End If
      End If
      If GetABS001_2 <> "" Then GetABS001_2 = Left(GetABS001_2, Len(GetABS001_2) - 1)
   End If
End Function

'檢查人員當時是否休假
'注意傳入的時間須要是時間格式 hh:mm
'Modify By Sindy 2012/9/13  +strRestKind 休假狀況:1=請假 3=出差
'Modify By Sindy 2024/7/29                        增加辨識 4=颱風假
'Modify By Sindy 2012/11/15 +bolIsRest1Day 是否請整日
'Modify By Sindy 2013/6/25 +bolNotABS010 不含簽核中的資料
'Modify By Sindy 2013/11/8 +strStarTime,strEndTime,strWOnTime,strWOffTime 改為回傳值
Public Function CheckIsPersonRest(ByVal StrST01 As String, ByVal strDate As String, ByVal strTime As String, _
                     Optional ByRef strRestKind As String = "", Optional ByRef bolIsRest1Day As Boolean = False, _
                     Optional ByVal bolNotABS010 As Boolean = False, _
                     Optional ByRef strStarTime As String = "", Optional ByRef strEndTime As String = "", _
                     Optional ByRef strWOnTime As String = "", Optional ByRef strWOffTime As String = "") As Boolean
Dim strCompDate As String
Dim strCompTime As String
Dim strStarDate As String
'Dim strStarTime As String
Dim strEndDate As String
'Dim strEndTime As String
'Dim strWOnTime As String
'Dim strWOffTime As String
Dim intDay As Integer
Dim dblHour As Double
Dim adoRst As ADODB.Recordset
Dim adoRst2 As ADODB.Recordset 'Add By Sindy 2022/10/25
Dim strStarWorkTime As String, strEndWorkTime As String 'Add By Sindy 2013/9/18
Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2013/10/30
Dim strMinPr02 As String, strMaxPr02 As String 'Add By Sindy 2013/11/8
Dim j As Integer, bolChkOk As Boolean 'Add By Sindy 2021/8/13
Dim bolChkRst2 As Boolean, dblTotHour As Double
Dim strST06 As String 'Add By Sindy 2023/10/11
   
   bolChkRst2 = False
   CheckIsPersonRest = False
   strRestKind = ""
   bolIsRest1Day = False
   strStarTime = ""
   strEndTime = ""
   strWOnTime = ""
   strWOffTime = ""
   
   'Add By Sindy 2012/4/18
   If StrST01 = "" Then Exit Function
   If strDate = "" Then Exit Function
   If strTime = "" Then Exit Function
   '2012/4/18 End
   
   '注意比較的日期格式為 YYYYMMDDHHMM
   strCompTime = IIf(strTime = "24:00", "2400", Format(strTime, "hhmm"))
   strCompDate = DBDATE(strDate) & strCompTime
   
   'Add By Sindy 2023/10/11 增加檢查當日是否為颱風假
   strST06 = PUB_GetST06(StrST01)
   strSql = ""
   If strST06 = "1" Then
      strSql = " and wd02='Y'"
   ElseIf strST06 = "2" Then
      strSql = " and wd03='Y'"
   ElseIf strST06 = "3" Then
      strSql = " and wd04='Y'"
   ElseIf strST06 = "4" Then
      strSql = " and wd05='Y'"
   End If
   If strSql <> "" Then
      strSql = "select * from workday where wd01=" & DBDATE(strDate) & strSql
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'Modify By Sindy 2024/7/29
         'strRestKind = "1"
         strRestKind = "4" '颱風假
         '2024/7/29 END
         bolIsRest1Day = True
         CheckIsPersonRest = True
         Set adoRst = Nothing
         Exit Function
      End If
   End If
   '2023/10/11 END
   
   'Modify By Sindy 2012/2/3 增加判斷若請整日者(B1028 is null),日期區期重覆到即視為請假
   'Modify By Sindy 2012/4/11 人事資料增加起迄上下班時段,請假時,無起迄上下班時段則視為請整日
'   strSql = "SELECT SA01 FROM staff_Absence WHERE SA01='" & strST01 & "' and (" & strCompDate & " between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4) or (SA16 is null and " & DBDATE(strDate) & " between SA02 and SA04)) " & _
'      "union SELECT SB01 FROM staff_busi_trip WHERE SB01='" & strST01 & "' and " & strCompDate & " between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4) " & _
'      "union SELECT B1003 FROM ABS010 WHERE B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "') and B1018 not in('" & 註銷 & "','" & 已核准 & "') and B1003='" & strST01 & "' and (" & strCompDate & " between B1004||substr(decode(B1005,0,0800,B1005)+10000,2,4) and B1006||substr(decode(B1007,0,1800,B1007)+10000,2,4) or (B1002='01' and B1028 is null and " & DBDATE(strDate) & " between B1004 and B1006)) "
'   intI = 1
'   Set adoRst = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      CheckIsPersonRest = True
'      Exit Function
'   Else
      'Add By Sindy 2012/4/11 改寫檢查語法
      '檢查是否有該日請假資料
      'Modify By Sindy 2012/12/18 +退回的表單視為無休假
      strSql = "SELECT SA01,SA02,SA03,SA04,SA05,SA16,SA17,SA07,SA08,'01' as sort FROM staff_Absence WHERE SA01='" & StrST01 & "' and (" & DBDATE(strDate) & " between SA02 and SA04)" & _
               " Union " & _
               "SELECT SB01,SB02,SB03,SB04,SB05,SB17,SB18,SB06,SB07,'03' as sort FROM staff_busi_trip WHERE SB01='" & StrST01 & "' and (" & DBDATE(strDate) & " between SB02 and SB04)"
      'Modify By Sindy 2013/10/18 +增加專利處的外出登記
      strSql = strSql & " Union " & _
                        "SELECT OG03,OG02,0+replace(OG19,':',''),OG02,0+replace(OG20,':',''),0,0,0,OG05,'03' as sort FROM outgoing WHERE OG03='" & StrST01 & "' and OG02=" & DBDATE(strDate)
      '2013/10/18 END
      'Modify By Sindy 2013/6/25
      If bolNotABS010 = False Then
         'Modify By Sindy 2025/1/9 排除人事已先行建檔資料
         '+ and not exists(select * from staff_Absence where sa09=b1001) and not exists(select * from staff_busi_trip where sb10=b1001)
         strSql = strSql & " Union " & _
                           "SELECT B1003,B1004,B1005,B1006,B1007,B1028,B1029,B1009,B1010,decode(b1002,'01','11','03','13') as sort" & _
                           " FROM ABS010 WHERE B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "')" & _
                           " and B1018 not in('" & 退回 & "','" & 註銷 & "','" & 已核准 & "')" & _
                           " and B1003='" & StrST01 & "' and (" & DBDATE(strDate) & " between B1004 and B1006)" & _
                           " and not exists(select * from staff_Absence where sa09=b1001)" & _
                           " and not exists(select * from staff_busi_trip where sb10=b1001)"
      End If
      strSql = strSql & " order by sort asc"
      '2013/6/25 End
      intI = 1
      Set adoRst2 = ClsLawReadRstMsg(intI, strSql) 'Add By Sindy 2022/10/25 另外，判斷一天多張假單使用
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         adoRst.MoveFirst
         'Do While CheckIsPersonRest = False
         Dim strB1028 As String
         Dim strB1029 As String
         Do While CheckIsPersonRest = False And Not adoRst.EOF
            strStarDate = adoRst.Fields(1)
            strStarTime = Format(adoRst.Fields(2), "0000")
            strEndDate = adoRst.Fields(3)
            strEndTime = Format(adoRst.Fields(4), "0000")
            intDay = Val("" & adoRst.Fields(7))
            dblHour = Val("" & adoRst.Fields(8))
            strB1028 = "" & adoRst.Fields(5)
            strB1029 = "" & adoRst.Fields(6)
            Call Pub_GetSpecWorkHour(StrST01, strDate, strStarWorkTime, strEndWorkTime) 'Add By Sindy 2013/9/18
            'Modify By Sindy 2013/11/8 若有刷卡資料以刷卡資料為主,辨識其上下班時段
            strSql = "select scd01,pr01,nvl(min(pr02),0) as min_pr02,nvl(max(pr02),0) as max_pr02" & _
                     " from pollrecord,staffcarddata" & _
                     " where pr01=" & DBDATE(strDate) & _
                     " and pr03=scd02(+) and scd01='" & StrST01 & "'" & _
                     " group by scd01,pr01"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            strMinPr02 = ""
            strMaxPr02 = ""
            If intI = 1 Then
               strMinPr02 = rsTmp.Fields(2)
               strMaxPr02 = rsTmp.Fields(3)
            End If
            rsTmp.Close
            If PUB_bWkSpec = False And strMinPr02 <> "" Then
               strWOnTime = Format("" & adoRst.Fields(5), "0000")
               strWOffTime = Format("" & adoRst.Fields(6), "0000")
               If strWOnTime = "0000" Then
                  'Modify By Sindy 2021/5/17
                  m_bolByPassWork = PUB_ChkByPassWork(strST06, strDate, strMinPr02, strMaxPr02, strWOnTime, strWOffTime)
                  strWOnTime = Format(strWOnTime, "HHMM")
                  strWOffTime = Format(strWOffTime, "HHMM")
                  '2021/5/17 END
'                  If Val(strMinPr02) <= 90000 Then
'                     If Val(strMinPr02) < 80000 Then
'                        strWOnTime = "0800"
'                        strWOffTime = "1700"
'                     ElseIf Val(strMinPr02) < 83000 Then
'                        strWOnTime = "0830"
'                        strWOffTime = "1730"
'                     Else
'                        strWOnTime = "0900"
'                        strWOffTime = "1800"
'                     End If
'                  ElseIf Val(strMaxPr02) >= 170000 Then
'                     If Val(strMaxPr02) >= 180000 Then
'                        strWOnTime = "0900"
'                        strWOffTime = "1800"
'                     ElseIf Val(strMaxPr02) >= 173000 Then
'                        strWOnTime = "0830"
'                        strWOffTime = "1730"
'                     Else
'                        strWOnTime = "0800"
'                        strWOffTime = "1700"
'                     End If
'                  End If
               End If
            ElseIf PUB_bWkSpec = True Then '特殊人員
               strWOnTime = strStarWorkTime
               strWOffTime = strEndWorkTime
            Else
               strWOnTime = Format("" & adoRst.Fields(5), "0000")
               If strWOnTime = "" Then strWOnTime = "0000" 'Add By Sindy 2015/12/18
               strWOffTime = Format("" & adoRst.Fields(6), "0000")
               If strWOffTime = "" Then strWOffTime = "0000" 'Add By Sindy 2015/12/18
            End If
            '2013/11/8 END
            
            'Add By Sindy 2012/9/13 休假狀況
            If adoRst.Fields(9) = "01" Or adoRst.Fields(9) = "11" Then
               strRestKind = "1" '請假
            Else 'adoRst.Fields(9) = "03" Or adoRst.Fields(9) = "13" Then
               strRestKind = "3" '出差
            End If
            '2012/9/13 End
            
            'Added by Lydia 2020/08/07 限查名人員才判斷異地上班; 因為有異地上班人員又出差到客戶方
            strSql = "select tmqm01 from tmqmember where tmqm01=" & CNULL(StrST01)
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
            'end 2020/08/07
                'Added by Lydia 2020/04/14 檢查是否異地上班
                strSql = "select sp01,sp02,sp03 from staff_workplace where sp01='" & strDate & "' and sp02=" & CNULL(StrST01)
                intI = 1
                Set rsTmp = ClsLawReadRstMsg(intI, strSql)
                'Modify By Sindy 2020/4/14 增加判斷必須是出差
                If intI = 1 And strRestKind = "3" Then
                    '異地上班=正常上班
                    CheckIsPersonRest = False
                    Set rsTmp = Nothing
                    Set adoRst = Nothing
                    Set adoRst2 = Nothing 'Add By Sindy 2022/10/25
                    Exit Function
                End If
                'end 2020/04/14
            End If 'Added by Lydia 2020/08/07
            
            Call PUB_ChkByPassWork(PUB_GetST06(StrST01), strStarDate) 'Add By Sindy 2021/8/13
            
            'Modify By Sindy 2015/11/19 因人員有可能請11/16 8:30 到11/18 12:10,一樣是2天但以11/18而言是非整日
            'If Not IsNull(adoRst.Fields(8)) Then
            If Not IsNull(adoRst.Fields(8)) And strStarDate = strEndDate Then
            '2015/11/19 END
               '若時數為0時,代表請整日
               'Modify By Sindy 2013/11/8 因出差時數算法和請假不同,雖算整日,但有可能是上班中途才出差去
               If dblHour = 0 Then
                  If strRestKind = "1" Then
                     bolIsRest1Day = True 'Add By Sindy 2012/11/15 請整日
                     CheckIsPersonRest = True
                     Set adoRst = Nothing
                     Set adoRst2 = Nothing 'Add By Sindy 2022/10/25
                     Exit Function
                  'Add By Sindy 2017/12/4
                  ElseIf strRestKind = "3" Then '出差
                     'Modify By Sindy 2021/8/13
                     'If intDay = 1 And strStarTime <= Format("900", "0000") And strEndTime >= "1700" Then
                     If intDay = 1 And _
                        strStarTime <= Format(Format(strByPassStarTime(intByPassArea), "hhmm"), "0000") And _
                        strEndTime >= Format(Format(strByPassEndTime(1), "hhmm"), "0000") Then
                     '2021/8/13 END
                        bolIsRest1Day = True 'Add By Sindy 2012/11/15 請整日
                        CheckIsPersonRest = True
                        Set adoRst = Nothing
                        Set adoRst2 = Nothing 'Add By Sindy 2022/10/25
                        Exit Function
                     End If
                  End If
                  '2017/12/4 END
               End If
            End If
            '若請多天時...
            If strStarDate <> strEndDate Then
               If strStarDate <> DBDATE(strDate) And strEndDate <> DBDATE(strDate) Then
                  '請假區間的中間日期一定是請整日
                  bolIsRest1Day = True 'Add By Sindy 2012/11/15 請整日
                  CheckIsPersonRest = True
                  Set adoRst = Nothing
                  Set adoRst2 = Nothing 'Add By Sindy 2022/10/25
                  Exit Function
               Else
                  '計算出該日請假的起迄時間:
                  '為請假的第一天時,必須計算出正確的下班時段
                  If strStarDate = DBDATE(strDate) Then
                     'Modify By Sindy 2013/9/18
                     If PUB_bWkSpec = True Then
                        strEndTime = strEndWorkTime
                     Else
                     '2013/9/18 END
                        Call PUB_ChkByPassWork(PUB_GetST06(StrST01), strStarDate) 'Add By Sindy 2021/8/13
                        
                        'Add By Sindy 2016/9/23 若多天,第一天的上班時間若有在正常的上班時間內,則以第一天的記錄去推算下班時間
                        'If strStarTime <= "0900" Then strWOnTime = "0000"
                        If strStarTime <= Format(Format(strByPassStarTime(intByPassArea), "hhmm"), "0000") Then strWOnTime = "0000"
                        
                        If strWOnTime = "0000" Then
'                           'Modify By Sindy 2013/11/8
'                           If strStarTime <= "0900" Then
'                              If strStarTime < "0800" Then
'                                 strEndTime = "1700"
'                              ElseIf strStarTime < "0830" Then
'                                 strEndTime = "1730"
'                              Else
'                                 strEndTime = "1800"
'                              End If
'                              strWOnTime = strStarTime
'                           Else
'                              strEndTime = "1800"
'                              strWOnTime = "0900"
'                           End If
'                           '2013/11/8 END
                           'Add By Sindy 2021/8/13
                           If strStarTime <= Format(Format(strByPassStarTime(intByPassArea), "hhmm"), "0000") Then '"0900"
                              For j = 1 To intByPassArea
                                 If strStarTime <= Format(Format(strByPassStarTime(j), "hhmm"), "0000") Then '"0800"
                                    strEndTime = Format(Format(strByPassEndTime(j), "hhmm"), "0000")
                                    Exit For
                                 End If
                                 If j = intByPassArea Then
                                    strEndTime = Format(Format(strByPassEndTime(intByPassArea), "hhmm"), "0000")
                                 End If
                              Next j
                              strWOnTime = strStarTime
                           Else
                              strEndTime = Format(Format(strByPassEndTime(intByPassArea), "hhmm"), "0000") '"1800"
                              strWOnTime = Format(Format(strByPassStarTime(intByPassArea), "hhmm"), "0000") '"0900"
                           End If
                           '2021/8/13 END
                           
                        Else
'                           If CStr(strWOnTime) <= "0800" Then
'                              strEndTime = "1700"
'                           ElseIf CStr(strWOnTime) > "0800" And CStr(strWOnTime) <= "0830" Then
'                              strEndTime = "1730"
'                           ElseIf CStr(strWOnTime) > "0830" And CStr(strWOnTime) <= "0900" Then
'                              strEndTime = "1800"
'                           End If
                           'Add By Sindy 2021/8/13
                           For j = 1 To intByPassArea
                              If CStr(strWOnTime) <= Format(Format(strByPassStarTime(j), "hhmm"), "0000") Then '"0800"
                                 strEndTime = Format(Format(strByPassEndTime(j), "hhmm"), "0000") '"1700"
                                 Exit For
                              End If
                           Next j
                           '2021/8/13 END
                        End If
                        strWOffTime = strEndTime
                     End If
                  
                  '為請假的最後一天時,必須計算出正確的上班時段
                  ElseIf strEndDate = DBDATE(strDate) Then
                     'Modify By Sindy 2013/9/18
                     If PUB_bWkSpec = True Then
                        strStarTime = strStarWorkTime
                     Else
                     '2013/9/18 END
                        Call PUB_ChkByPassWork(PUB_GetST06(StrST01), strEndDate) 'Add By Sindy 2021/8/13
                        
                        'Add By Sindy 2016/9/23 若多天,最後一天的下班時間若已過正常下班時間,則以最後一天的記錄去推算上班時間
                        'If strEndTime >= "1700" Then strWOffTime = "0000"
                        If strEndTime >= Format(Format(strByPassEndTime(1), "hhmm"), "0000") Then strWOffTime = "0000"
                        
                        If strWOffTime = "0000" Then
                           'Modify By Sindy 2013/11/8
'                           If strEndTime >= "1700" Then
'                              If strEndTime >= "1800" Then
'                                 strStarTime = "0900"
'                              ElseIf strEndTime >= "1730" Then
'                                 strStarTime = "0830"
'                              Else
'                                 strStarTime = "0800"
'                              End If
'                              strWOffTime = strEndTime
'                           Else
'                              strStarTime = "0800"
'                              strWOffTime = "1700"
'                           End If
                           '2013/11/8 END
                           'Add By Sindy 2021/8/13
                           If strEndTime >= Format(Format(strByPassEndTime(1), "hhmm"), "0000") Then '"1700"
                              For j = intByPassArea To 1 Step -1
                                 If strEndTime >= Format(Format(strByPassEndTime(j), "hhmm"), "0000") Then '"1800"
                                    strStarTime = Format(Format(strByPassStarTime(j), "hhmm"), "0000") '"0900"
                                    Exit For
                                 End If
                                 If j = 1 Then
                                    strStarTime = Format(Format(strByPassStarTime(1), "hhmm"), "0000")
                                 End If
                              Next j
                              strWOffTime = strEndTime
                           Else
                              strStarTime = Format(Format(strByPassStarTime(1), "hhmm"), "0000") '"0800"
                              strWOffTime = Format(Format(strByPassEndTime(1), "hhmm"), "0000") '"1700"
                           End If
                           '2021/8/13 END
                           
                        Else
'                           If CStr(strWOffTime) >= "1700" And CStr(strWOffTime) <= "1729" Then
'                              strStarTime = "0800"
'                           ElseIf CStr(strWOffTime) >= "1730" And CStr(strWOffTime) <= "1759" Then
'                              strStarTime = "0830"
'                           ElseIf CStr(strWOffTime) >= "1800" Then
'                              strStarTime = "0900"
'                           End If
                           'Add By Sindy 2021/8/13
                           For j = 1 To intByPassArea
                              If CStr(strWOffTime) >= Format(Format(strByPassEndTime(j), "hhmm"), "0000") Then
                                 strStarTime = Format(Format(strByPassStarTime(j), "hhmm"), "0000")
                                 Exit For
                              End If
                           Next j
                           '2021/8/13 END
                        End If
                        strWOnTime = strStarTime
                     End If
                  End If
               End If
            
            Else
               'Add By Sindy 2022/10/25 檢查是否有同一天多張假單
               If bolChkRst2 = False Then
                  bolChkRst2 = True: dblTotHour = 0
                  adoRst2.MoveFirst
                  Do While Not adoRst2.EOF
                     'Modify By Sindy 2025/4/25
                     'If strDate = "" & adoRst.Fields(1) Then
                     If strDate = "" & adoRst2.Fields(1) Then
                        'intTotDay = intTotDay + Val("" & adoRst.Fields(7))
                        'dblTotHour = dblTotHour + Val("" & adoRst.Fields(8))
                        dblTotHour = dblTotHour + Val("" & adoRst2.Fields(8))
                        '2025/4/25 END
                     End If
'                     strStarTime = Format(adoRst.Fields(2), "0000")
'                     strEndDate = adoRst.Fields(3)
'                     strEndTime = Format(adoRst.Fields(4), "0000")
'                     intDay = Val("" & adoRst.Fields(7))
'                     dblHour = Val("" & adoRst.Fields(8))
'                     strB1028 = "" & adoRst.Fields(5)
'                     strB1029 = "" & adoRst.Fields(6)
                     adoRst2.MoveNext
                  Loop
                  If dblTotHour >= PUB_intWkHour - 0.5 Then
                     bolIsRest1Day = True '請整日
                     CheckIsPersonRest = True
                     Set adoRst = Nothing
                     Set adoRst2 = Nothing
                     Exit Function
                  End If
               End If
            End If
            
            'Modify By Sindy 2017/2/20 若比對的日期為未來日期,則時間就用設定時間比對
            If DBDATE(strDate) > strSrvDate(1) Then
               If strWOnTime = "0000" Then
                  strWOnTime = strCompTime
               End If
            Else
            '2017/2/20 END
               If strWOnTime = "0000" Then
                  'If Val(strStarTime) <= 900 Then
                  If Val(strStarTime) <= Val(Format(strByPassStarTime(intByPassArea), "hhmm")) Then
                     strWOnTime = strStarTime
                  Else
                     'strWOnTime = "0800"
                     strWOnTime = Format(Format(strByPassStarTime(1), "hhmm"), "0000") '
                  End If
               End If
            End If
            
            'Add By Sindy 2012/11/15 檢查是否請整日
            'Modify By Sindy 2016/7/14 +(strStarTime <= strB1028 And strEndTime >= "1800") ==> (strStarTime <= "0800" And strEndTime >= "1800")
'            If (strStarTime = "0800" And strEndTime = "1700") Or _
'               (strStarTime = "0830" And strEndTime = "1730") Or _
'               (strStarTime = "0900" And strEndTime = "1800") Or _
'               (strStarTime <= IIf(strB1028 <> "", Format(strB1028, "0000"), "0800") And _
'                strEndTime >= IIf(strB1029 <> "", Format(strB1029, "0000"), "1800")) Or _
'               (strStarTime = strWOnTime And strEndTime = strWOffTime) Then
            'Add By Sindy 2021/8/13
            bolChkOk = False
            For j = 1 To intByPassArea
               If strStarTime = Format(Format(strByPassStarTime(j), "hhmm"), "0000") And _
                  strEndTime = Format(Format(strByPassEndTime(j), "hhmm"), "0000") Then
                  bolChkOk = True
                  Exit For
               End If
            Next j
            If bolChkOk = True Or _
               (strStarTime <= IIf(strB1028 <> "", Format(strB1028, "0000"), Format(Format(strByPassStarTime(1), "hhmm"), "0000")) And _
                strEndTime >= IIf(strB1029 <> "", Format(strB1029, "0000"), Format(Format(strByPassEndTime(intByPassArea), "hhmm"), "0000"))) Or _
               (strStarTime = strWOnTime And strEndTime = strWOffTime) Then
            '2021/8/13 END
               bolIsRest1Day = True
               'Add By Sindy 2016/8/29
               CheckIsPersonRest = True
               Set adoRst = Nothing
               Set adoRst2 = Nothing 'Add By Sindy 2022/10/25
               Exit Function
               '2016/8/29 END
            End If
            '2012/11/15 End
            
            '非整日的檢查
            'Modify By Sindy 2018/3/26 摩根接到雅娟電話說107/3/26早上8:30請假沒有轉職代
'            If Val(strCompTime) < 800 Then '比對的時間小於上班日
'               '請假起始時間若和上班時間一樣代表一早就請假了
'               If Val(strStarTime) = Val(strWOnTime) Then
'                  CheckIsPersonRest = True
'                  Set adoRst = Nothing
'                  Exit Function
'               End If
'            Else
'               '比較時間有在請假時間的區間內代表請假
'               If Val(strCompTime) >= Val(strStarTime) And Val(strCompTime) <= Val(strEndTime) Then
               'Modify By Sindy 2018/12/12 + Or (Val(strCompTime) >= 1700 And Val(strEndTime) >= 1700)
               '  請假時間為16:00~17:00,檢查假單在17:00之後
'               If (Val(strCompTime) >= Val(strStarTime) And Val(strCompTime) <= Val(strEndTime)) Or _
'                  (Val(strCompTime) <= 900 And Val(strStarTime) <= 900) Or _
'                  (Val(strCompTime) >= 1700 And Val(strEndTime) >= 1700) Then
            '2018/3/26 END
               'Add By Sindy 2021/8/13
               If (Val(strCompTime) >= Val(strStarTime) And Val(strCompTime) <= Val(strEndTime)) Or _
                  (Val(strCompTime) <= Val(Format(strByPassStarTime(intByPassArea), "hhmm")) And _
                   Val(strStarTime) <= Val(Format(strByPassStarTime(intByPassArea), "hhmm"))) Or _
                  (Val(strCompTime) >= Val(Format(strByPassEndTime(1), "hhmm")) And _
                   Val(strEndTime) >= Val(Format(strByPassEndTime(1), "hhmm"))) Then
               '2021/8/13 END
                  CheckIsPersonRest = True
                  Set adoRst = Nothing
                  Set adoRst2 = Nothing 'Add By Sindy 2022/10/25
                  Exit Function
               End If
'            End If
            adoRst.MoveNext
         Loop
      End If
      '2012/4/11 End
'   End If
   Set adoRst = Nothing
   Set adoRst2 = Nothing 'Add By Sindy 2022/10/25
End Function

'Add By Sindy 2016/4/20
'檢查人員是否請假區間”完全相同”
'注意傳入的時間須要是時間格式 hh:mm
'(含外出記錄做檢查)
Public Function CheckIsPersonRestSectorSame(StrST01 As String, strDateS As String, strTimeS As String, strDateE As String, strTimeE As String, strB1001 As String) As Boolean
Dim strCompDateS As String, strCompDateE As String, strCon As String
Dim adoRst As ADODB.Recordset
   
   CheckIsPersonRestSectorSame = False
   
   If StrST01 = "" Then Exit Function
   If strDateS = "" Then Exit Function
   If strTimeS = "" Then Exit Function
   If strDateE = "" Then Exit Function
   If strTimeE = "" Then Exit Function
   
   '注意比較的日期格式為 YYYYMMDDHHMM
   strCompDateS = DBDATE(strDateS) & IIf(strTimeS = "24:00", "2400", Format(strTimeS, "hhmm"))
   strCompDateE = DBDATE(strDateE) & IIf(strTimeE = "24:00", "2400", Format(strTimeE, "hhmm"))
   
   If strB1001 <> "" Then strCon = " and B1001<>" & CNULL(strB1001) & " "
   
   '檢查起始日期:
   '增加判斷若請整日者(B1028 is null),日期區期重覆到即視為請假
   '退回的表單視為無休假
   '含外出記錄做檢查
   strSql = "SELECT SA01 FROM staff_Absence WHERE SA01='" & StrST01 & "' and (" & strCompDateS & " between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4)) and (" & strCompDateE & " between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4)) " & _
      "union SELECT SB01 FROM staff_busi_trip WHERE SB01='" & StrST01 & "' and (" & strCompDateS & " between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4)) and (" & strCompDateE & " between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4)) " & _
      "union SELECT B1003 FROM ABS010 WHERE B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "') and B1018 not in('" & 退回 & "','" & 註銷 & "','" & 已核准 & "') and B1003='" & StrST01 & "' and (((" & strCompDateS & " between B1004||substr(decode(B1005,0,0800,B1005)+10000,2,4) and B1006||substr(decode(B1007,0,1800,B1007)+10000,2,4)) and (" & strCompDateE & " between B1004||substr(decode(B1005,0,0800,B1005)+10000,2,4) and B1006||substr(decode(B1007,0,1800,B1007)+10000,2,4))) or (B1028 is null and " & DBDATE(strDateS) & " between b1004 and b1006 and " & DBDATE(strDateE) & " between b1004 and b1006)) " & strCon & _
      "union SELECT og01 FROM outgoing WHERE og03='" & StrST01 & "' and (" & strCompDateS & " between og02||replace(og19,':','') and og02||replace(og20,':','')) and (" & strCompDateE & " between og02||replace(og19,':','') and og02||replace(og20,':','')) "
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      CheckIsPersonRestSectorSame = True: Exit Function
   End If
   Set adoRst = Nothing
End Function
'2016/4/20 END

'檢查人員區間內是否休假
'注意傳入的時間須要是時間格式 hh:mm
'Modify By Sindy 2015/1/16 +含外出記錄做檢查
Public Function CheckIsPersonRestSector(StrST01 As String, strDateS As String, strTimeS As String, strDateE As String, strTimeE As String, strB1001 As String) As Boolean
Dim strCompDateS As String, strCompDateE As String, strCon As String
Dim adoRst As ADODB.Recordset
   
   CheckIsPersonRestSector = False
   
   'Add By Sindy 2012/4/18
   If StrST01 = "" Then Exit Function
   If strDateS = "" Then Exit Function
   If strTimeS = "" Then Exit Function
   If strDateE = "" Then Exit Function
   If strTimeE = "" Then Exit Function
   '2012/4/18 End
   
'   'Add By Sindy 2011/10/25 事後補請假時,不控管請假區期是否相同
'   If strSrvDate(1) > Val(DBDATE(strDateS)) And strSrvDate(1) > Val(DBDATE(strDateE)) Then Exit Function
   
   '注意比較的日期格式為 YYYYMMDDHHMM
   strCompDateS = DBDATE(strDateS) & IIf(strTimeS = "24:00", "2400", Format(strTimeS, "hhmm"))
   strCompDateE = DBDATE(strDateE) & IIf(strTimeE = "24:00", "2400", Format(strTimeE, "hhmm"))
   
   If strB1001 <> "" Then strCon = " and B1001<>" & CNULL(strB1001) & " "
   
   '檢查起始日期
   'Modify By Sindy 2012/2/3 增加判斷若請整日者(B1028 is null),日期區期重覆到即視為請假 ==> (B1002='01' and B1028 is null ==取消(B1002='01' and)
   'Modify By Sindy 2012/12/18 +退回的表單視為無休假
   'Modify By Sindy 2015/1/16 +含外出記錄做檢查
   strSql = "SELECT SA01 FROM staff_Absence WHERE SA01='" & StrST01 & "' and " & strCompDateS & " between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4) " & _
      "union SELECT SB01 FROM staff_busi_trip WHERE SB01='" & StrST01 & "' and " & strCompDateS & " between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4) " & _
      "union SELECT B1003 FROM ABS010 WHERE B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "') and B1018 not in('" & 退回 & "','" & 註銷 & "','" & 已核准 & "') and B1003='" & StrST01 & "' and (" & strCompDateS & " between B1004||substr(decode(B1005,0,0800,B1005)+10000,2,4) and B1006||substr(decode(B1007,0,1800,B1007)+10000,2,4) or (B1028 is null and " & DBDATE(strDateS) & " between B1004 and B1006)) " & strCon & _
      "union SELECT og01 FROM outgoing WHERE og03='" & StrST01 & "' and " & strCompDateS & " between og02||replace(og19,':','') and og02||replace(og20,':','') "
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      CheckIsPersonRestSector = True: Exit Function
   End If
   If strCompDateS <> strCompDateE Then
      '檢查迄止日期
      'Modify By Sindy 2015/1/16 +含外出記錄做檢查
      strSql = "SELECT SA01 FROM staff_Absence WHERE SA01='" & StrST01 & "' and " & strCompDateE & " between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4) " & _
         "union SELECT SB01 FROM staff_busi_trip WHERE SB01='" & StrST01 & "' and " & strCompDateE & " between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4) " & _
         "union SELECT B1003 FROM ABS010 WHERE B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "') and B1018 not in('" & 退回 & "','" & 註銷 & "','" & 已核准 & "') and B1003='" & StrST01 & "' and (" & strCompDateE & " between B1004||substr(decode(B1005,0,0800,B1005)+10000,2,4) and B1006||substr(decode(B1007,0,1800,B1007)+10000,2,4) or (B1028 is null and " & DBDATE(strDateE) & " between B1004 and B1006)) " & strCon & _
         "union SELECT og01 FROM outgoing WHERE og03='" & StrST01 & "' and " & strCompDateE & " between og02||replace(og19,':','') and og02||replace(og20,':','') "
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         CheckIsPersonRestSector = True: Exit Function
      End If
      '檢查起迄區間
      'Modify By Sindy 2015/1/16 +含外出記錄做檢查
      strSql = "SELECT SA01 FROM staff_Absence WHERE SA01='" & StrST01 & "' and " & strCompDateS & "<= SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04>0 and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4)<=" & strCompDateE & " " & _
         "union SELECT SB01 FROM staff_busi_trip WHERE SB01='" & StrST01 & "' and " & strCompDateS & "<= SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04>0 and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4)<=" & strCompDateE & " " & _
         "union SELECT B1003 FROM ABS010 WHERE B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "') and B1018 not in('" & 退回 & "','" & 註銷 & "','" & 已核准 & "') and B1003='" & StrST01 & "' and ((" & strCompDateS & "<= B1004||substr(decode(B1005,0,0800,B1005)+10000,2,4) and B1006>0 and B1006||substr(decode(B1007,0,1800,B1007)+10000,2,4)<=" & strCompDateE & ") or (B1028 is null and (" & DBDATE(strDateS) & "<= B1004 and B1006>0 and B1006<=" & DBDATE(strDateE) & "))) " & strCon & _
         "union SELECT og01 FROM outgoing WHERE og03='" & StrST01 & "' and " & strCompDateS & "<= og02||replace(og19,':','') and og02||replace(og20,':','')<=" & strCompDateE & " "
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         CheckIsPersonRestSector = True: Exit Function
      End If
   End If
   Set adoRst = Nothing
End Function

'Add By Sindy 2018/11/15
'假單核准或主管代填請假時,檢查在此請假區間中是否有幫他人做職代
'若有,發Mail通知其他職代
Public Function CheckIsPersonRestSectorMail(strB1001 As String) As Boolean
Dim strCompDateS As String, strCompDateE As String
Dim strCompDateS_2 As String, strCompDateE_2 As String
Dim strContent As String
Dim adoRst As ADODB.Recordset
Dim adoTmp As ADODB.Recordset
Dim strB1003 As String, strB1004 As String, strB1005 As String, strB1006 As String, strB1007 As String
Dim QueryID As String
Dim strContent_Main As String 'Add By Sindy 2025/5/20
   
   CheckIsPersonRestSectorMail = False
   
   strSql = "SELECT * FROM ABS010 WHERE B1001='" & strB1001 & "' and B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "')"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strB1003 = adoRst.Fields("b1003")
      strB1004 = adoRst.Fields("b1004")
      strB1005 = adoRst.Fields("b1005")
      strB1006 = adoRst.Fields("b1006")
      strB1007 = adoRst.Fields("b1007")
      
      If adoRst.Fields("b1006") < strSrvDate(1) Then
         Set adoRst = Nothing
         Exit Function
      Else
         strContent_Main = "但其職務代理人(" & GetPrjSalesNM(adoRst.Fields("b1003")) & ")" & "" & "於[表單編號:" & strB1001 & "] " & _
            Format(ChangeWStringToTString(strB1004), "###/##/##") & " " & _
            Format(Right("00" & strB1005, 4), "##:##") & _
            " ～ " & _
            Format(ChangeWStringToTString(strB1006), "###/##/##") & " " & _
            Format(Right("00" & strB1007, 4), "##:##") & IIf(adoRst.Fields("b1002") = "01", " 同時請假", " 同時出差")
         
         '注意比較的日期格式為 YYYYMMDDHHMM
         If adoRst.Fields("b1004") >= strSrvDate(1) Then
            strCompDateS = adoRst.Fields("b1004") & Format(adoRst.Fields("b1005"), "0000")
         Else
            strCompDateS = strSrvDate(1) & "0800"
         End If
         strCompDateE = adoRst.Fields("b1006") & Format(adoRst.Fields("b1007"), "0000")
      End If
      
      'Modify By Sindy 2022/10/28 EX:11106604簽核表單修改,導至下面mail出下列訊息(但10/27雅雯是沒休假的)
      'Subject: ◎職務代理人通知[11106881] [寄件者：北所(十樓)#330]
      '內容:
      '王雅雯同仁於[表單編號:11106604] 111/10/27 8:00 ～ 111/10/31 17:00 請假
      '但其職務代理人(劉湘芸)於[表單編號:11106881] 111/10/27 8:00 ～ 111/10/27 12:00 同時請假
      '通知由您擔任王雅雯同仁的職務代理人。
      
      '檢查是否有幫他人做職代
      'Modify By Sindy 2022/10/28 修改SQL,已簽核完的資料改抓人事資料
      strSql = "select B1001,B1002,B1003,B1004,B1005,B1006,B1007,st02 from abs010,abs011,staff where B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "') and B1019 is null" & _
               " and (b1004>=" & strSrvDate(1) & " or b1006>=" & strSrvDate(1) & ")" & _
               " and b1001=b1101 and b1102='1'" & _
               " and b1104='" & strB1003 & "'" & _
               " and B1018='" & 已核准 & "' and b1003=st01(+) and st04='1'" & _
               " and ((b1004||substr(decode(b1005,0,0800,b1005)+10000,2,4) between '" & strCompDateS & "' and '" & strCompDateE & "'" & _
                  " or b1006||substr(decode(b1007,0,1800,b1007)+10000,2,4) between '" & strCompDateS & "' and '" & strCompDateE & "')" & _
                 " or ('" & strCompDateS & "' between b1004||substr(decode(b1005,0,0800,b1005)+10000,2,4) and b1006||substr(decode(b1007,0,1800,b1007)+10000,2,4)" & _
                  " or '" & strCompDateE & "' between b1004||substr(decode(b1005,0,0800,b1005)+10000,2,4) and b1006||substr(decode(b1007,0,1800,b1007)+10000,2,4))" & _
               ")"
      strSql = strSql & " union " & _
               "select SA09 as B1001,'01' as B1002,SA01 as B1003,SA02 as B1004,SA03 as B1005,SA04 as B1006,SA05 as B1007,st02 from Staff_Absence,abs011,staff where SA09 is not null" & _
               " and (SA02>=" & strSrvDate(1) & " or SA04>=" & strSrvDate(1) & ")" & _
               " and SA09=b1101 and b1102='1'" & _
               " and b1104='" & strB1003 & "'" & _
               " and SA01=st01(+) and st04='1'" & _
               " and ((SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) between '" & strCompDateS & "' and '" & strCompDateE & "'" & _
                  " or SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4) between '" & strCompDateS & "' and '" & strCompDateE & "')" & _
                 " or ('" & strCompDateS & "' between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4)" & _
                  " or '" & strCompDateE & "' between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4))" & _
               ")"
      strSql = strSql & " union " & _
               "select SB10 as B1001,'01' as B1002,SB01 as B1003,SB02 as B1004,SB03 as B1005,SB04 as B1006,SB05 as B1007,st02 from Staff_Busi_Trip,abs011,staff where SB10 is not null" & _
               " and (SB02>=" & strSrvDate(1) & " or SB04>=" & strSrvDate(1) & ")" & _
               " and SB10=b1101 and b1102='1'" & _
               " and b1104='" & strB1003 & "'" & _
               " and SB01=st01(+) and st04='1'" & _
               " and ((SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) between '" & strCompDateS & "' and '" & strCompDateE & "'" & _
                  " or SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4) between '" & strCompDateS & "' and '" & strCompDateE & "')" & _
                 " or ('" & strCompDateS & "' between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4)" & _
                  " or '" & strCompDateE & "' between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4))" & _
               ")"
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         adoRst.MoveFirst
         Do While Not adoRst.EOF
            '注意比較的日期格式為 YYYYMMDDHHMM
            If adoRst.Fields("b1004") >= strSrvDate(1) Then
               strCompDateS_2 = adoRst.Fields("b1004") & Format(adoRst.Fields("b1005"), "0000")
            Else
               strCompDateS_2 = strSrvDate(1) & "0800"
            End If
            strCompDateE_2 = adoRst.Fields("b1006") & Format(adoRst.Fields("b1007"), "0000")
            
            '檢查此假單是否還有其他人做職代,若有,則不用再抓職代通知
            strSql = "SELECT B1104 FROM ABS011,staff" & _
                     " WHERE B1101='" & adoRst.Fields("b1001") & "' and b1102='1'" & _
                     " and b1104=st01(+) and st04='1' and B1104<>'" & strB1003 & "'"
            intI = 1
            Set adoTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Then
               '抓取人事職代
               QueryID = ""
               strSql = "SELECT B0102 FROM ABS001,staff WHERE B0101='" & adoRst.Fields("b1003") & "' and b0102=st01(+) and st04='1' and B0102<>'" & strB1003 & "'" & _
                        "union SELECT B0103 FROM ABS001,staff WHERE B0101='" & adoRst.Fields("b1003") & "' and b0103=st01(+) and st04='1' and B0103<>'" & strB1003 & "'" & _
                        "union SELECT B0104 FROM ABS001,staff WHERE B0101='" & adoRst.Fields("b1003") & "' and b0104=st01(+) and st04='1' and B0104<>'" & strB1003 & "'" & _
                        "union SELECT B0105 FROM ABS001,staff WHERE B0101='" & adoRst.Fields("b1003") & "' and b0105=st01(+) and st04='1' and B0105<>'" & strB1003 & "'" & _
                        "union SELECT B0106 FROM ABS001,staff WHERE B0101='" & adoRst.Fields("b1003") & "' and b0106=st01(+) and st04='1' and B0106<>'" & strB1003 & "'" & _
                        "union SELECT B0107 FROM ABS001,staff WHERE B0101='" & adoRst.Fields("b1003") & "' and b0107=st01(+) and st04='1' and B0107<>'" & strB1003 & "'"
               intI = 1
               Set adoTmp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  adoTmp.MoveFirst
                  Do While Not adoTmp.EOF
                     QueryID = QueryID & "," & adoTmp.Fields(0) '記錄新的職代人員
                     adoTmp.MoveNext
                  Loop
               End If
               If QueryID <> "" Then
                  '檢查其他職代是否也同時請假或出差,若無,才發Mail通知人員做職代
                  'Modify By Sindy 2022/10/28 修改SQL,已簽核完的資料改抓人事資料
                  strSql = "select b1003,st02 from abs010,staff where B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "') and B1019 is null" & _
                           " and b1003 in('" & Replace(Mid(QueryID, 2), ",", "','") & "')" & _
                           " and B1018='" & 已核准 & "' and b1003=st01(+) and st04='1'" & _
                           " and ((b1004||substr(decode(b1005,0,0800,b1005)+10000,2,4) between '" & strCompDateS & "' and '" & strCompDateE & "'" & _
                              " or b1006||substr(decode(b1007,0,1800,b1007)+10000,2,4) between '" & strCompDateS & "' and '" & strCompDateE & "')" & _
                             " or ('" & strCompDateS & "' between b1004||substr(decode(b1005,0,0800,b1005)+10000,2,4) and b1006||substr(decode(b1007,0,1800,b1007)+10000,2,4)" & _
                              " or '" & strCompDateE & "' between b1004||substr(decode(b1005,0,0800,b1005)+10000,2,4) and b1006||substr(decode(b1007,0,1800,b1007)+10000,2,4))" & _
                             " or (b1004||substr(decode(b1005,0,0800,b1005)+10000,2,4) between '" & strCompDateS_2 & "' and '" & strCompDateE_2 & "'" & _
                              " or b1006||substr(decode(b1007,0,1800,b1007)+10000,2,4) between '" & strCompDateS_2 & "' and '" & strCompDateE_2 & "')" & _
                             " or ('" & strCompDateS_2 & "' between b1004||substr(decode(b1005,0,0800,b1005)+10000,2,4) and b1006||substr(decode(b1007,0,1800,b1007)+10000,2,4)" & _
                              " or '" & strCompDateE_2 & "' between b1004||substr(decode(b1005,0,0800,b1005)+10000,2,4) and b1006||substr(decode(b1007,0,1800,b1007)+10000,2,4))" & _
                           ")"
                  strSql = strSql & " union " & _
                           "select SA01 as b1003,st02 from Staff_Absence,staff where SA09 is not null" & _
                           " and SA01 in('" & Replace(Mid(QueryID, 2), ",", "','") & "')" & _
                           " and SA01=st01(+) and st04='1'" & _
                           " and ((SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) between '" & strCompDateS & "' and '" & strCompDateE & "'" & _
                              " or SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4) between '" & strCompDateS & "' and '" & strCompDateE & "')" & _
                             " or ('" & strCompDateS & "' between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4)" & _
                              " or '" & strCompDateE & "' between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4))" & _
                             " or (SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) between '" & strCompDateS_2 & "' and '" & strCompDateE_2 & "'" & _
                              " or SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4) between '" & strCompDateS_2 & "' and '" & strCompDateE_2 & "')" & _
                             " or ('" & strCompDateS_2 & "' between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4)" & _
                              " or '" & strCompDateE_2 & "' between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4))" & _
                           ")"
                  strSql = strSql & " union " & _
                           "select SB01 as b1003,st02 from Staff_Busi_Trip,staff where SB10 is not null" & _
                           " and SB01 in('" & Replace(Mid(QueryID, 2), ",", "','") & "')" & _
                           " and SB01=st01(+) and st04='1'" & _
                           " and ((SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) between '" & strCompDateS & "' and '" & strCompDateE & "'" & _
                              " or SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4) between '" & strCompDateS & "' and '" & strCompDateE & "')" & _
                             " or ('" & strCompDateS & "' between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4)" & _
                              " or '" & strCompDateE & "' between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4))" & _
                             " or (SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) between '" & strCompDateS_2 & "' and '" & strCompDateE_2 & "'" & _
                              " or SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4) between '" & strCompDateS_2 & "' and '" & strCompDateE_2 & "')" & _
                             " or ('" & strCompDateS_2 & "' between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4)" & _
                              " or '" & strCompDateE_2 & "' between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4))" & _
                           ")"
                  intI = 1
                  Set adoTmp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     adoTmp.MoveFirst
                     Do While Not adoTmp.EOF
                        QueryID = Replace(QueryID, "," & adoTmp.Fields("b1003"), "") '清除
                        adoTmp.MoveNext
                     Loop
                  End If
                  If QueryID <> "" Then
                     strContent = adoRst.Fields("st02") & "同仁於[表單編號:" & adoRst.Fields("b1001") & "] " & _
                        Format(ChangeWStringToTString(adoRst.Fields("b1004")), "###/##/##") & " " & _
                        Format(Right("00" & adoRst.Fields("b1005"), 4), "##:##") & _
                        " ～ " & _
                        Format(ChangeWStringToTString(adoRst.Fields("b1006")), "###/##/##") & " " & _
                        Format(Right("00" & adoRst.Fields("b1007"), 4), "##:##") & IIf(adoRst.Fields("b1002") = "01", " 請假", " 出差") & vbCrLf & _
                        strContent_Main & vbCrLf & _
                        "通知由您擔任" & adoRst.Fields("st02") & "同仁的職務代理人。"
                     PUB_SendMail strUserNum, Replace(Mid(QueryID, 2), ",", ";"), "", "職務代理人通知[" & strB1001 & "]", strContent, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , False, , , , "QPGMR", "系統管理員", , True, False
                     CheckIsPersonRestSectorMail = True
                  End If
               End If
            End If
            adoRst.MoveNext
         Loop
      End If
   End If
   
   Set adoRst = Nothing
   Set adoTmp = Nothing
End Function

'檢查請假,加班,出差的資料是否已存在
'注意傳入的時間須要是時間格式 hh:mm
'Modify By Sindy 2015/1/16 +含外出記錄做檢查
'Modify By Sindy 2017/10/25 +Optional strAbsenceType As String = "":表單種類(01.請假 02.加班 03.出差)
Public Function CheckIsAbsenceExist(StrST01 As String, strDateS As String, _
   strTimeS As String, strDateE As String, strTimeE As String, strB1001 As String, _
   Optional strAbsenceType As String = "") As Boolean
Dim strCompDateS As String, strCompDateE As String
Dim strCon As String, strConSA As String, strConSO As String, strConSB As String
Dim adoRst As ADODB.Recordset
Dim strConOG As String 'Add By Sindy 2015/1/16
Dim strSql As String
   
   CheckIsAbsenceExist = False
   
   'Add By Sindy 2012/4/18
   If StrST01 = "" Then Exit Function
   If strDateS = "" Then Exit Function
   If strTimeS = "" Then Exit Function
   If strDateE = "" Then Exit Function
   If strTimeE = "" Then Exit Function
   '2012/4/18 End
   
   '注意比較的日期格式為 YYYYMMDDHHMM
   strCompDateS = DBDATE(strDateS) & IIf(strTimeS = "24:00", "2400", Format(strTimeS, "hhmm"))
   strCompDateE = DBDATE(strDateE) & IIf(strTimeE = "24:00", "2400", Format(strTimeE, "hhmm"))
   'Add By Sindy 2012/7/4 假單可以輸入重覆的時間,如第一張假單填8:00到(10:00),第二張假單可以填(10:00)到11:00
   strCompDateS = Val(strCompDateS) + 1
   '2012/7/4 End
   'Add By Sindy 2013/11/18 假單可以輸入重覆的時間,如第一張假單填(12:10)到17:30,第二張假單可以填11:20到(12:10)
   strCompDateE = Val(strCompDateE) - 1
   '2013/11/18 End
   
   '加班或全部
   If strAbsenceType = "02" Or strAbsenceType = "" Then
      If strB1001 <> "" Then strConSO = " and So13<>" & CNULL(strB1001) & " "
      If strB1001 <> "" Then strCon = " and B1001<>" & CNULL(strB1001) & " "
   End If
   '非加班
   If strAbsenceType <> "02" Then
      If strB1001 <> "" Then strCon = " and B1001<>" & CNULL(strB1001) & " "
      If strB1001 <> "" Then strConSA = " and SA09<>" & CNULL(strB1001) & " "
      If strB1001 <> "" Then strConSB = " and SB10<>" & CNULL(strB1001) & " "
      If strB1001 <> "" Then strConOG = " and og01<>" & CNULL(strB1001) & " " 'Add By Sindy 2015/1/16
   End If
   
   '檢查起始日期
   'Modify By Sindy 2012/2/3 增加判斷若請整日者(B1028 is null),日期區期重覆到即視為請假
   'Modify By Sindy 2015/1/16 +含外出記錄做檢查
   '加班或全部
   strSql = ""
   If strAbsenceType = "02" Or strAbsenceType = "" Then
      'Modify By Sindy 2018/1/5 取消 or (B1002 in('" & 表單類別_加班 & "') and B1028 is null and " & DBDATE(strDateS) & " between B1004 and B1004))
      strSql = "SELECT So01 FROM Staff_Overtime WHERE So01='" & StrST01 & "' and " & strCompDateS & " between So02||substr(decode(So03,0,0800,So03)+10000,2,4) and So02||substr(decode(So04,0,1800,So04)+10000,2,4) " & strConSO & _
         "union SELECT B1001 FROM ABS010 WHERE B1018 not in('" & 註銷 & "','" & 已核准 & "') and B1003='" & StrST01 & "' and " & strCompDateS & " between B1004||substr(decode(B1005,0,0800,B1005)+10000,2,4) and B1004||substr(decode(B1007,0,1800,B1007)+10000,2,4) " & strCon
   End If
   '非加班
   If strAbsenceType <> "02" Then
      If strSql <> "" Then strSql = strSql & " union "
      'Modify By Sindy 2018/1/5 取消 or (B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "') and B1028 is null and " & DBDATE(strDateS) & " between B1004 and B1006)
      strSql = strSql & _
               "SELECT SA01 FROM staff_Absence WHERE SA01='" & StrST01 & "' and " & strCompDateS & " between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4) " & strConSA & _
         "union SELECT SB01 FROM staff_busi_trip WHERE SB01='" & StrST01 & "' and " & strCompDateS & " between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4) " & strConSB & _
         "union SELECT B1001 FROM ABS010 WHERE B1018 not in('" & 註銷 & "','" & 已核准 & "') and B1003='" & StrST01 & "' and (" & strCompDateS & " between B1004||substr(decode(B1005,0,0800,B1005)+10000,2,4) and B1006||substr(decode(B1007,0,1800,B1007)+10000,2,4)) " & strCon & _
         "union SELECT og01 FROM outgoing WHERE og03='" & StrST01 & "' and " & strCompDateS & " between og02||replace(og19,':','') and og02||replace(og20,':','') " & strConOG
   End If
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      CheckIsAbsenceExist = True: Exit Function
   End If
   If strCompDateS <> strCompDateE Then
      '檢查迄止日期
      'Modify By Sindy 2015/1/16 +含外出記錄做檢查
      '加班或全部
      strSql = ""
      If strAbsenceType = "02" Or strAbsenceType = "" Then
         'Modify By Sindy 2018/1/5 取消 or (B1002 in('" & 表單類別_加班 & "') and B1028 is null and " & DBDATE(strDateE) & " between B1004 and B1004))
         strSql = "SELECT So01 FROM Staff_Overtime WHERE So01='" & StrST01 & "' and " & strCompDateE & " between So02||substr(decode(So03,0,0800,So03)+10000,2,4) and So02||substr(decode(So04,0,1800,So04)+10000,2,4) " & strConSO & _
            "union SELECT B1001 FROM ABS010 WHERE B1018 not in('" & 註銷 & "','" & 已核准 & "') and B1003='" & StrST01 & "' and " & strCompDateE & " between B1004||substr(decode(B1005,0,0800,B1005)+10000,2,4) and B1004||substr(decode(B1007,0,1800,B1007)+10000,2,4) " & strCon
      End If
      '非加班
      If strAbsenceType <> "02" Then
         If strSql <> "" Then strSql = strSql & " union "
         'Modify By Sindy 2018/1/5 取消 or (B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "') and B1028 is null and " & DBDATE(strDateE) & " between B1004 and B1006)
         strSql = strSql & _
                  "SELECT SA01 FROM staff_Absence WHERE SA01='" & StrST01 & "' and " & strCompDateE & " between SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4) " & strConSA & _
            "union SELECT SB01 FROM staff_busi_trip WHERE SB01='" & StrST01 & "' and " & strCompDateE & " between SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4) " & strConSB & _
            "union SELECT B1001 FROM ABS010 WHERE B1018 not in('" & 註銷 & "','" & 已核准 & "') and B1003='" & StrST01 & "' and (" & strCompDateE & " between B1004||substr(decode(B1005,0,0800,B1005)+10000,2,4) and B1006||substr(decode(B1007,0,1800,B1007)+10000,2,4)) " & strCon & _
            "union SELECT og01 FROM outgoing WHERE og03='" & StrST01 & "' and " & strCompDateE & " between og02||replace(og19,':','') and og02||replace(og20,':','') " & strConOG
      End If
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         CheckIsAbsenceExist = True: Exit Function
      End If
      '檢查起迄區間
      'Modify By Sindy 2015/1/16 +含外出記錄做檢查
      '加班或全部
      strSql = ""
      If strAbsenceType = "02" Or strAbsenceType = "" Then
         'Modify By Sindy 2018/1/5 取消  or (B1002 in('" & 表單類別_加班 & "') and B1028 is null and (" & DBDATE(strDateS) & "<= B1004 and B1004<=" & DBDATE(strDateE) & "))
         strSql = "SELECT So01 FROM Staff_Overtime WHERE So01='" & StrST01 & "' and " & strCompDateS & "<= So02||substr(decode(So03,0,0800,So03)+10000,2,4) and So02||substr(decode(So04,0,1800,So04)+10000,2,4)<=" & strCompDateE & " " & strConSO & _
            "union SELECT B1001 FROM ABS010 WHERE B1018 not in('" & 註銷 & "','" & 已核准 & "') and B1003='" & StrST01 & "' and " & strCompDateS & "<= B1004||substr(decode(B1005,0,0800,B1005)+10000,2,4) and B1004||substr(decode(B1007,0,1800,B1007)+10000,2,4)<=" & strCompDateE & strCon
      End If
      '非加班
      If strAbsenceType <> "02" Then
         If strSql <> "" Then strSql = strSql & " union "
         'Modify By Sindy 2018/1/5 取消 or (B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "') and B1028 is null and (" & DBDATE(strDateS) & "<= B1004 and B1006>0 and B1006<=" & DBDATE(strDateE) & "))
         strSql = strSql & _
                  "SELECT SA01 FROM staff_Absence WHERE SA01='" & StrST01 & "' and " & strCompDateS & "<= SA02||substr(decode(SA03,0,0800,SA03)+10000,2,4) and SA04>0 and SA04||substr(decode(SA05,0,1800,SA05)+10000,2,4)<=" & strCompDateE & " " & strConSA & _
            "union SELECT SB01 FROM staff_busi_trip WHERE SB01='" & StrST01 & "' and " & strCompDateS & "<= SB02||substr(decode(SB03,0,0800,SB03)+10000,2,4) and SB04>0 and SB04||substr(decode(SB05,0,1800,SB05)+10000,2,4)<=" & strCompDateE & " " & strConSB & _
            "union SELECT B1001 FROM ABS010 WHERE B1018 not in('" & 註銷 & "','" & 已核准 & "') and B1003='" & StrST01 & "' and ((" & strCompDateS & "<= B1004||substr(decode(B1005,0,0800,B1005)+10000,2,4) and B1006>0 and B1006||substr(decode(B1007,0,1800,B1007)+10000,2,4)<=" & strCompDateE & ")) " & strCon & _
            "union SELECT og01 FROM outgoing WHERE og03='" & StrST01 & "' and " & strCompDateS & "<= og02||replace(og19,':','') and og02||replace(og20,':','')<=" & strCompDateE & " " & strConOG
      End If
      intI = 1
      Set adoRst = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         CheckIsAbsenceExist = True: Exit Function
      End If
   End If
   Set adoRst = Nothing
End Function

'傳回已同意的審核主管
Public Function GetBossB1107_2_1(strKEY01 As String) As String
Dim adoRst As ADODB.Recordset
   
   GetBossB1107_2_1 = ""
   strSql = "SELECT B1104 FROM ABS011 " & _
            "WHERE B1101='" & strKEY01 & "' and B1102='2' and B1107='1' order by B1103 asc "
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With adoRst
         .MoveFirst
         Do While Not .EOF
            GetBossB1107_2_1 = GetBossB1107_2_1 & adoRst.Fields(0) & ";"
            .MoveNext
         Loop
         If GetBossB1107_2_1 <> "" Then GetBossB1107_2_1 = Left(GetBossB1107_2_1, Len(GetBossB1107_2_1) - 1)
      End With
   End If
   Set adoRst = Nothing
End Function

'Add By Sindy 2012/3/20
'傳回審核主管未簽核的尚有幾位
Public Function GetBossNotSignsCnt(strKEY01 As String) As Integer
Dim adoRst As ADODB.Recordset
   
   GetBossNotSignsCnt = 0
   strSql = "SELECT count(*) FROM ABS011 " & _
            "WHERE B1101='" & strKEY01 & "' and B1102='2' and B1107 is null "
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      GetBossNotSignsCnt = adoRst.Fields(0)
   End If
   Set adoRst = Nothing
End Function

'傳回已同意的職代
Public Function GetBossB1107_1_1(strKEY01 As String) As String
Dim rsTmp As New ADODB.Recordset
   
   GetBossB1107_1_1 = ""
   strSql = "SELECT B1104 FROM ABS011 " & _
            "WHERE B1101='" & strKEY01 & "' and B1102='1' and B1107='1' order by B1103 asc "
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With rsTmp
         .MoveFirst
         Do While Not .EOF
            GetBossB1107_1_1 = GetBossB1107_1_1 & rsTmp.Fields(0) & ";"
            .MoveNext
         Loop
         If GetBossB1107_1_1 <> "" Then GetBossB1107_1_1 = Left(GetBossB1107_1_1, Len(GetBossB1107_1_1) - 1)
      End With
   End If
End Function

'傳回已簽核的職代和審核主管
Public Function GetBossB1107_All(strKEY01 As String) As String
Dim rsTmp As New ADODB.Recordset
   
   GetBossB1107_All = ""
   strSql = "SELECT distinct B1104 FROM ABS011 " & _
            "WHERE B1101='" & strKEY01 & "' and B1107 is not null order by B1104 asc "
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With rsTmp
         .MoveFirst
         Do While Not .EOF
            GetBossB1107_All = GetBossB1107_All & rsTmp.Fields(0) & ";"
            .MoveNext
         Loop
         If GetBossB1107_All <> "" Then GetBossB1107_All = Left(GetBossB1107_All, Len(GetBossB1107_All) - 1)
      End With
   End If
End Function

'出缺勤系統
'過濾人員是否有離職，休假狀況；傳回下一處理人員
'注意：AutoBatchDay會使用到該函數
'Modify By Sindy 2012/9/27 +bolAutoBatch
Public Function GetNextProPerson(strKEY01 As String, strB1003 As String, ByRef strB1104 As String, strB1016 As String, Optional bolAutoBatch As Boolean = False) As Boolean
Dim strTemp As Variant, i As Integer, j As Integer
Dim intB1103 As Integer
Dim strB1018 As String
Dim rsTmp As New ADODB.Recordset
Dim Rs As New ADODB.Recordset
Dim bolRunLoopFindNextP As Boolean
Dim m_ABS001_1 As String, m_ABS001_2 As String, m_ABS001_3 As String
Dim strData As String, strOldB1104 As String, strB1017 As String
'Add By Sindy 2012/3/20
Dim strB1002 As String, strB1004 As String, strB1005 As String, strB1006 As String
Dim strB1007 As String, strB1008 As String, strB1011 As String
'2012/3/20 End
Dim m_bolIsRest1Day As Boolean 'Add By Sindy 2012/11/15
Dim strB1109 As String 'Add By Sindy 2015/11/18
Dim strB1108 As String 'Add By Sindy 2017/6/9
   
On Error GoTo ErrHand
   
   GetNextProPerson = True
   
   '讀取表單主檔資料
   strSql = "SELECT * FROM ABS010 WHERE B1001='" & strKEY01 & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strB1002 = "" & RsTemp.Fields("B1002")
      strB1004 = "" & RsTemp.Fields("B1004")
      strB1005 = "" & RsTemp.Fields("B1005")
      strB1006 = "" & RsTemp.Fields("B1006")
      strB1007 = "" & RsTemp.Fields("B1007")
      strB1008 = "" & RsTemp.Fields("B1008")
      strB1011 = "" & RsTemp.Fields("B1011")
      strB1017 = "" & RsTemp.Fields("B1017")
'      'AutoBatch時,才須要檢查下列條件
'      If bolAutoBatch = True Then
'         '若該筆表單的下一處理人員為自己或人事處,則不須往下執行
'         If RsTemp.Fields("B1003") = "" & RsTemp.Fields("B1017") Or _
'            "" & RsTemp.Fields("B1017") = "M21" Then
'            Exit Function
'         End If
'      End If
   End If
   
   bolRunLoopFindNextP = True '預設一進此函數都會Run一次do迴圈
   Do While bolRunLoopFindNextP = True
      bolRunLoopFindNextP = False
      
      'Modify By Sindy 2015/11/18 + B1109
      strSql = "SELECT B1102,B1103,B1104,B1109,B1108 FROM ABS011 " & _
               "WHERE B1101='" & strKEY01 & "' and B1107 is null order by B1102,B1103 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         
         intB1103 = rsTmp.Fields("B1103")
         strB1104 = rsTmp.Fields("B1104")
         strB1108 = "" & rsTmp.Fields("B1108") '(代) 'Add By Sindy 2017/6/9
         strB1109 = "" & rsTmp.Fields("B1109") 'Add By Sindy 2015/11/18
         strOldB1104 = rsTmp.Fields("B1104")
         If rsTmp.Fields("B1102") = "1" Then '職代
            '在一開始成立表單時,會過濾職代是否有和請假當事人重覆到請假區間,及是否為離職人員
            strB1018 = 會簽職代
            
            '下一處理人員若為離職或休假
            'Modify By Sindy 2012/11/15 +m_bolIsRest1Day
            If ChkStaffST04(strB1104, False) = True Or _
               CheckIsPersonRest(strB1104, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , m_bolIsRest1Day) = True Then
               
               'Add By Sindy 2012/11/15
               'If bolAutoBatch = True And m_bolIsRest1Day = False Then '若為批次執行,人員非休整日,不轉職代
               'If m_bolIsRest1Day = False And strB1004 <> strSrvDate(1) Then '簽核人員不是休整日並且不是當日假單,不轉職代
               'Modify By Sindy 2016/4/21 批核主管未全天請假或出差時，系統仍會留給原主管批核，不會轉交主管職代批核
               If m_bolIsRest1Day = False Then
                  strB1104 = strOldB1104 'Add By Sindy 2013/2/19
                  GoTo ReadChkEnd
               End If
               '2012/11/15 End
               
               '請假當事人的職代名單
               strData = strB1104
               Call GetABS001_1(strB1003, m_ABS001_1, m_ABS001_2, m_ABS001_3, , strData)
               'Add By Sindy 2017/8/22
               '回傳 雙職代的A區或B區人員資料
               If InStr(m_ABS001_1, ",") > 0 Then '雙職代
                  If strData <> "" Then
                     strTemp = Split(strData, ",")
                     For i = 0 To UBound(strTemp)
                        '不可為離職不可為休假
                        If ChkStaffST04(CStr(strTemp(i)), False) = False And _
                           CheckIsPersonRest(CStr(strTemp(i)), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False Then
                           '檢查是否已在職代簽核人員名單中
                           If Rs.State <> 0 Then Rs.Close
                           strSql = "SELECT count(*) FROM ABS011 " & _
                                    "WHERE B1101='" & strKEY01 & "' and B1102='1' and B1104='" & CStr(strTemp(i)) & "' "
                           Rs.CursorLocation = adUseClient
                           Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                           If Rs.Fields(0) = 0 Then '尚不是職代時
                              '表單當事人的請假資料
                              If Rs.State <> 0 Then Rs.Close
                              strSql = "SELECT * FROM ABS010 WHERE B1001='" & strKEY01 & "' "
                              Rs.CursorLocation = adUseClient
                              Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                              If Rs.RecordCount > 0 Then
                                 '檢查取得的職代和表單當事人是否有相同的請假區間,若有,則找下一職代
                                 If CheckIsPersonRestSector(CStr(strTemp(i)), Rs.Fields("B1004"), Right("00" & Format(CStr(Rs.Fields("B1005")), "##:##"), 5), Rs.Fields("B1006"), Right("00" & Format(CStr(Rs.Fields("B1007")), "##:##"), 5), strKEY01) = False Then
                                    '取代已離職或休假人員
                                    strSql = "update ABS011 " & _
                                             "set B1104='" & CStr(strTemp(i)) & "' " & _
                                             ",B1108='(代" & CStr(i + 1) & ")' " & _
                                             "WHERE B1101='" & strKEY01 & "' and B1102='1' and B1104='" & strOldB1104 & "' and B1107 is null"
                                    cnnConnection.Execute strSql
                                    strB1104 = CStr(strTemp(i)) '已找到下一處理人員
                                    Exit For
                                 End If
                              End If
                           End If
                        End If
                     Next i
                  End If
               Else
               '2017/8/22 END
                  For j = 1 To 3 '有3組職代
                     strData = ""
                     If j = 1 And m_ABS001_1 <> "" Then strData = m_ABS001_1
                     If j = 2 And m_ABS001_2 <> "" Then strData = m_ABS001_2
                     If j = 3 And m_ABS001_3 <> "" Then strData = m_ABS001_3
                     If strData <> "" Then
                        strTemp = Split(strData, ",")
                        For i = 0 To UBound(strTemp)
                           '不可為離職不可為休假
                           If ChkStaffST04(CStr(strTemp(i)), False) = False And _
                              CheckIsPersonRest(CStr(strTemp(i)), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False Then
                              '檢查是否已在職代簽核人員名單中
                              If Rs.State <> 0 Then Rs.Close
                              strSql = "SELECT count(*) FROM ABS011 " & _
                                       "WHERE B1101='" & strKEY01 & "' and B1102='1' and B1104='" & CStr(strTemp(i)) & "' "
                              Rs.CursorLocation = adUseClient
                              Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                              If Rs.Fields(0) = 0 Then '尚不是職代時
                                 '表單當事人的請假資料
                                 If Rs.State <> 0 Then Rs.Close
                                 strSql = "SELECT * FROM ABS010 WHERE B1001='" & strKEY01 & "' "
                                 Rs.CursorLocation = adUseClient
                                 Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                                 If Rs.RecordCount > 0 Then
                                    '檢查取得的職代和表單當事人是否有相同的請假區間,若有,則找下一職代
                                    If CheckIsPersonRestSector(CStr(strTemp(i)), Rs.Fields("B1004"), Right("00" & Format(CStr(Rs.Fields("B1005")), "##:##"), 5), Rs.Fields("B1006"), Right("00" & Format(CStr(Rs.Fields("B1007")), "##:##"), 5), strKEY01) = False Then
                                       '取代已離職或休假人員
                                       strSql = "update ABS011 " & _
                                                "set B1104='" & CStr(strTemp(i)) & "' " & _
                                                ",B1108='(代" & CStr(j) & ")' " & _
                                                "WHERE B1101='" & strKEY01 & "' and B1102='1' and B1104='" & strOldB1104 & "' and B1107 is null"
                                       cnnConnection.Execute strSql
                                       strB1104 = CStr(strTemp(i))
                                       Exit For
                                    End If
                                 End If
                              End If
                           End If
                        Next i
                        If strB1104 <> "" And strB1104 <> strOldB1104 Then Exit For '已找到下一處理人員
                     End If
                  Next j
               End If
               If strB1104 = "" Then strB1104 = strOldB1104 '維持原簽核人員
               
               '下一處理人員若為離職者且又找不到職代時
               If ChkStaffST04(strB1104, False) = True And _
                  strB1104 = strOldB1104 Then
                  '刪除該離職人員待簽核資料
                  strSql = "delete from ABS011 " & _
                           "WHERE B1101='" & strKEY01 & "' and B1102='1' and B1104='" & strB1104 & "' and B1107 is null"
                  cnnConnection.Execute strSql
                  strB1104 = "" 'Add By Sindy 2012/3/20
                  '記錄刪除簽核
                  strSql = GetInsertABS012Sql(strKEY01, "", strSrvDate(1), Right("000000" & ServerTime, 6), "", "[系統自動刪除]" & GetPrjSalesNM(strB1104) & "已離職且找不到職代")
                  cnnConnection.Execute strSql
                  '改送下一處理人員
                  bolRunLoopFindNextP = True
               End If
               If Rs.State <> 0 Then Rs.Close
            End If
            
         Else '2.審核主管
            strB1018 = 主管審核中
            
            'Add By Sindy 2017/6/9
            '檢查目前的簽核主管是否為代理的,若是,先檢查原簽核主管是否還休假中,若未休假了則轉回該主管簽核
            If strB1108 = "(代)" And strB1104 <> strB1109 And strB1109 <> "" Then
               '原簽核主管是否還休假中,若未休假了則轉回該主管簽核
               'Modify By Sindy 2017/12/4
               'If CheckPerCurrRestReturnPer(strB1109, strB1003, m_bolIsRest1Day) = False Then
               If CheckIsPersonRest(strB1109, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , m_bolIsRest1Day) = False Then
               '2017/12/4 END
                  strB1104 = strB1109
                  strSql = "update ABS011 set " & _
                           "B1104='" & strB1109 & "',B1108=null " & _
                           "where B1101='" & strKEY01 & "' and B1102='2' and B1103=" & intB1103
                  cnnConnection.Execute strSql
                  GoTo ReadChkEnd
               Else
                  '主管未全天請假或出差時，系統仍會留給原主管批核
                  If m_bolIsRest1Day = False Then
                     strB1104 = strB1109
                     strSql = "update ABS011 set " & _
                              "B1104='" & strB1109 & "',B1108=null " & _
                              "where B1101='" & strKEY01 & "' and B1102='2' and B1103=" & intB1103
                     cnnConnection.Execute strSql
                     GoTo ReadChkEnd
                  End If
               End If
            End If
            '2017/6/9 END
            
            'Add By Sindy 2013/4/10 若待簽核人員是王副總時,請假/出差單要剔除請假日14天以後的假單暫不需進行是否需轉職代的檢查
            'Modify By Sindy 2023/4/24 秀玲說5/11自然取消，不用修改程式
            If strB1104 = "71011" And (strB1002 = "01" Or strB1002 = "03") And Val(strB1004) > Val(CompDate(2, 14, strSrvDate(1))) Then
               GoTo ReadChkEnd
            End If
            '2013/4/10 End
            
            '檢查審核主管是否有休假
            'Modify By Sindy 2012/11/15 +m_bolIsRest1Day
            'Modify By Sindy 2015/11/18 + strB1109
            If CheckPerCurrRestReturnPer(strB1104, strB1003, m_bolIsRest1Day, strB1109) = True Then '為休假
               'Add By Sindy 2012/11/15
               'If bolAutoBatch = True And m_bolIsRest1Day = False Then '若為批次執行,人員非休整日,不轉職代
               'If m_bolIsRest1Day = False And strB1004 <> strSrvDate(1) Then '簽核人員不是休整日並且不是當日假單,不轉職代
               'Modify By Sindy 2016/4/21 審核主管未全天請假或出差時，系統仍會留給原主管批核，不會轉交主管職代批核
               'Modify By Sindy 2022/4/1 增加判斷當天的假單, 若假單的開始時間主管也在休假中, 得轉職代或下一次主管
               If m_bolIsRest1Day = False And _
                  Not (strB1004 = strSrvDate(1) And _
                       CheckIsPersonRest(strOldB1104, strB1004, Left(Right("0000" & strB1005, 4), 2) & ":" & Mid(Right("0000" & strB1005, 4), 3, 2)) = True) Then
                  strB1104 = strOldB1104 'Add By Sindy 2013/3/26 流程停住,維持原簽核人員
                  GoTo ReadChkEnd
               End If
               '2012/11/15 End
               
               'Add By Sindy 2012/3/20
               '71011.王副總提專利處(王副總假單,及審核主管為王副總者除外)審核主管若休假,不走職代,送下一審核主管
               'Modify By Sindy 2015/7/23 智權部69005.簡協理,69010.蘇特助請假時,若不是同部門,他們的職代不可以代理簽核,送呈林總
               'Add By Sindy 2023/4/24 增加檢查if的條件 +73022
               'Modify By Sindy 2023/12/5 And Left(Trim(PUB_GetST03(strB1104)), 1) = "P" => 取消; 總經理休假變成職代都非P單位才導至流程停止
               'Modified by Morgan 2025/2/20 71011->82026
               If (Left(Trim(PUB_GetST03(strB1003)), 2) = "P1" _
                   And ((strB1003 <> "82026" And strB1104 <> "82026" And strOldB1104 <> "82026") _
                         Or (Val(strSrvDate(1)) >= 20230501 And strB1003 <> "73022" And strB1104 <> "73022" And strOldB1104 <> "73022") _
                       ) _
                  ) _
                  Or _
                  (Left(Trim(PUB_GetST03(strB1003)), 1) = "S" _
                   And (strOldB1104 = "69005" Or strOldB1104 = "69010") _
                   And PUB_GetST03(strB1104) <> PUB_GetST03(strB1003) _
                  ) Then
                  
                  '[專利處簽核流程]
                  '若未簽核的審核主管只有1位時,固定設王副總簽核
                  If GetBossNotSignsCnt(strKEY01) = 1 Then
                     '智權部=>林總
                     If Left(Trim(PUB_GetST03(strB1003)), 1) = "S" Then
                        strB1104 = "94007"
                     Else
                        'Modify By Sindy 2023/12/5
                        'If strB1104 <> "94007" Then
                        If strOldB1104 <> "94007" Then
                        '2023/12/5 END
                           'Modify By Sindy 2023/4/24 5/1 改為游
                           If Val(strSrvDate(1)) >= 20230501 Then
                              'Modified by Morgan 2025/2/21
                              'strB1104 = "73022" '專利處:游經理
                              pub_PMan = Pub_GetSpecMan("專利處特定編號")
                              strB1104 = Left(pub_PMan, 5)
                              'end 2025/2/21
                           Else
                           '2023/4/24 END
                              strB1104 = "71011" '專利處:王副總
                           End If
                        End If
                     End If
                     'Add By Sindy 2017/12/4 要先檢查此簽核主管是否也休假
                     'Call CheckPerCurrRestReturnPer(strB1104, strB1003, m_bolIsRest1Day)
                     '2017/12/4 END
                     'Modify By Sindy 2023/5/18
                     '游經理請假時，應由游經理簽核的假單轉由李柏翰簽核。
                     '一、郭雅娟直接轉林總簽核。 => 2024/6/5 有問過游協理,不用轉林總,停置等他回來再簽
                     '二、P12及P14人員轉由郭雅娟簽核。
                     '若本人與李柏翰同時請假時，則假單轉郭雅娟簽核。
                     '若本人與郭雅娟同時請假時，則假單轉李柏翰簽核。
                     If CheckPerCurrRestReturnPer(strB1104, strB1003, m_bolIsRest1Day) = True Then
                        If Left(Trim(PUB_GetST03(strB1003)), 2) = "P1" And strB1104 <> "94007" Then
                           If Left(Trim(PUB_GetST03(strB1003)), 3) = "P12" Or Left(Trim(PUB_GetST03(strB1003)), 3) = "P14" Then
'                              'Add By Sindy 2024/6/5 雅娟的假單(11303308) 有問過游協理,不用轉林總,停置等他回來再簽
'                              If strB1003 <> "79075" Then
'                              '2024/6/5 END
                              strB1104 = "79075" '郭雅娟
                              'Modify By Sindy 2025/3/13 + Or strB1003 = "79075" 增加判斷是雅娟的假單
                              If CheckPerCurrRestReturnPer(strB1104, strB1003, m_bolIsRest1Day) = True Or strB1003 = "79075" Then
                                 strB1104 = "99050" '李柏翰
                                 'Modify By Sindy 2024/10/3
                                 If CheckPerCurrRestReturnPer(strB1104, strB1003, m_bolIsRest1Day) = True Then
                                    strB1104 = strOldB1104 '維持原簽核人員
                                 End If
                                 '2024/10/3 END
                              End If
'                              End If
                           Else
                              strB1104 = "99050" '李柏翰
                              'Modify By Sindy 2025/3/13 + Or strB1003 = "99050" 增加判斷是柏翰的假單
                              If CheckPerCurrRestReturnPer(strB1104, strB1003, m_bolIsRest1Day) = True Or strB1003 = "99050" Then
                                 strB1104 = "79075" '郭雅娟
                                 'Modify By Sindy 2024/10/3
                                 If CheckPerCurrRestReturnPer(strB1104, strB1003, m_bolIsRest1Day) = True Then
                                    strB1104 = strOldB1104 '維持原簽核人員
                                 End If
                                 '2024/10/3 END
                              End If
                           End If
                        End If
                     End If
                     '2023/5/18 END
                     strSql = "update ABS011 set " & _
                              "B1104='" & strB1104 & "' " & _
                              ",B1108='(代)' " & _
                              "where B1101='" & strKEY01 & "' and B1102='2' and B1103=" & intB1103
                     cnnConnection.Execute strSql
                  Else
                     'Modify By Sindy 2013/5/7 不刪除改上簽核日期為19221111,EX.10202140
'                     '刪除休假主管待簽核資料
'                     strSql = "delete from ABS011 " & _
'                              "WHERE B1101='" & strKEY01 & "' and B1102='2' and B1104='" & strOldB1104 & "' and B1107 is null"
'                     cnnConnection.Execute strSql
                     strSql = "update ABS011 set " & _
                              "B1105=19221111,B1106=0,B1107='1' " & _
                              "WHERE B1101='" & strKEY01 & "' and B1102='2' and B1104='" & strOldB1104 & "' and B1107 is null"
                     cnnConnection.Execute strSql
                     strB1104 = ""
                     '記錄刪除簽核
                     strSql = GetInsertABS012Sql(strKEY01, "", strSrvDate(1), Right("000000" & ServerTime, 6), "", "[系統自動更新]" & GetPrjSalesNM(strOldB1104) & "休假不轉職代，由下一級審核主管簽核")
                     cnnConnection.Execute strSql
                     
                     'Add By Sindy 2024/10/3
                     '改送下一處理人員
                     bolRunLoopFindNextP = True
                     '2024/10/3 END
                  End If
                  'Modify By Sindy 2024/10/3 mark
'                  '改送下一處理人員
'                  bolRunLoopFindNextP = True
                  '2024/10/3 END
                  '發信通知請假主管
                  'Modify By Sindy 2012/9/27 AutoBatch時,不可執行PUB_SendMail
                  If bolAutoBatch = False Then
                  '2012/9/27 End
                     PUB_SendMail strUserNum, strOldB1104, "", GetPrjSalesNM(strB1003) & "表單轉由下一級審核主管簽核通知！", _
                     vbCrLf & "　　因您不在事務所，故此表單已轉由下一級審核主管簽核！" & vbCrLf & vbCrLf & _
                     "表單人員：" & strB1003 & " " & GetPrjSalesNM(strB1003) & vbCrLf & _
                     "表單類別：" & IIf(strB1002 = "01", "請假", IIf(strB1002 = "02", "加班", IIf(strB1002 = "03", "出差", ""))) & vbCrLf & _
                     IIf(strB1002 = "01", "假　　別：" & GetAllCode04(strB1008) & vbCrLf, "") & _
                     "表單日期：" & ChangeWStringToTDateString(strB1004) & " " & Format(Right("00" & strB1005, 4), "##:##") & " ~ " & IIf(strB1006 = "", ChangeWStringToTDateString(strB1004), ChangeWStringToTDateString(strB1006)) & " " & Format(Right("00" & strB1007, 4), "##:##") & vbCrLf & _
                     "事　　由：" & strB1011, , , , , , , , , , True
                  End If
               Else
               '2012/3/20 End
                  'Add By Sindy 2012/4/24 王副總該簽的假單，若副總請假則轉職代73022.游登銘，若2人都同時休假流程則停住。
                  'Modify By Sindy 2023/4/24 秀玲說等職代表更新
                  'If strOldB1104 = "71011" And strB1104 <> "73022" Then
                  'Modify By Sindy 2023/5/18
                  'If (strOldB1104 = "71011" Or (Val(strSrvDate(1)) >= 20230501 And strOldB1104 = "73022")) And strB1104 <> "73022" Then
                  If Left(Trim(PUB_GetST03(strB1003)), 2) = "P1" Then
                  '2023/5/18 END
                     strB1104 = strOldB1104 '流程停住,維持原簽核人員
                  Else
                  '2012/4/24 End
                     '[職代簽核流程]
                     If strB1104 <> "" Then
                        '更新簽核人員
                        strTemp = Split(strB1104, ",")
                        '若傳回1人以上者,須異動簽核檔的"順序"欄位值
                        If UBound(strTemp) > 0 Then
                           strSql = "update ABS011 set " & _
                                    "B1103=B1103+" & UBound(strTemp) & _
                                    " where B1101='" & strKEY01 & "' and B1102='2' and B1103>" & intB1103
                           cnnConnection.Execute strSql
                        End If
                        For i = 0 To UBound(strTemp)
                           intB1103 = intB1103 + i
                           '更新傳回的第1人
                           If i = 0 Then
                              strB1104 = strTemp(i)
                              strSql = "update ABS011 set " & _
                                       "B1104='" & strTemp(i) & "' " & _
                                       ",B1108='(代)' " & _
                                       "where B1101='" & strKEY01 & "' and B1102='2' and B1103=" & intB1103
                              cnnConnection.Execute strSql
                           Else
                              '新增傳回1人以上
                              'Modify By Sindy 2015/11/13 + B1109
                              strSql = "insert into ABS011 (B1101,B1102,B1103,B1104,B1108,B1109) values(" & CNULL(strKEY01) & ",'2'," & intB1103 & ",'" & strTemp(i) & "','(代)','" & strOldB1104 & "')"
                              cnnConnection.Execute strSql
                           End If
                        Next i
                     Else
                        If strB1104 = "" Then strB1104 = strOldB1104 '維持原簽核人員
                        '下一處理人員若為離職者且又找不到職代時
                        If ChkStaffST04(strB1104, False) = True Then
                           '刪除該離職人員待簽核資料
                           strSql = "delete from ABS011 " & _
                                    "WHERE B1101='" & strKEY01 & "' and B1102='2' and B1104='" & strB1104 & "' and B1107 is null"
                           cnnConnection.Execute strSql
                           strB1104 = "" 'Add By Sindy 2012/3/20
                           '記錄刪除簽核
                           strSql = GetInsertABS012Sql(strKEY01, "", strSrvDate(1), Right("000000" & ServerTime, 6), "", "[系統自動刪除]" & GetPrjSalesNM(strB1104) & "已離職且找不到職代")
                           cnnConnection.Execute strSql
                           '改送下一處理人員
                           bolRunLoopFindNextP = True
                        End If
                     End If
                  End If
               End If
            End If
            
         End If
ReadChkEnd:
      Else
         strB1104 = 人事處
         strB1018 = 送人事處簽收
      End If
      If rsTmp.State <> 0 Then rsTmp.Close
      
      '檢查下一處理人員是否已簽核過,若是,則直接Update其相同的簽核日期資料
      If strB1104 <> "" Then 'Add By Sindy 2012/3/20 +if
         If Rs.State <> 0 Then Rs.Close
         strSql = "SELECT * FROM ABS011 " & _
                  "WHERE B1101='" & strKEY01 & "' and B1104='" & strB1104 & "' and B1107 is not null order by B1102,B1103 asc "
         Rs.CursorLocation = adUseClient
         Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If Rs.RecordCount > 0 Then
            'Add By Sindy 2013/11/11
            If Val(Rs.Fields("b1105")) <> 19221111 Then
            '2013/11/11 END
               strSql = "update ABS011 set " & _
                        "B1105=" & Rs.Fields("B1105") & _
                        ",B1106=" & Rs.Fields("B1106") & _
                        ",B1107='" & Rs.Fields("B1107") & "' " & _
                        "WHERE B1101='" & strKEY01 & "' and B1104='" & strB1104 & "' and B1107 is null"
               cnnConnection.Execute strSql
               '改送下一處理人員
               bolRunLoopFindNextP = True
            End If
         End If
      End If
   Loop
   
   '有異動下一處理人員時才須更新資料
   If strB1104 <> "" And strB1104 <> strB1017 Then
      strSql = "update ABS010 set " & _
               "B1016='" & strB1016 & "'" & _
               ",B1017='" & strB1104 & "'" & _
               ",B1018='" & strB1018 & "'" & _
               " where B1001='" & strKEY01 & "' "
      cnnConnection.Execute strSql
   End If
   
'   'Modify By Sindy 2011/10/17
'   If bolAutoBatch = True Then
'      If strB1104 = 人事處 Then strB1104 = Pub_GetSpecMan("人事處出缺勤電子簽核")
'   End If
   
   Set rsTmp = Nothing
   Set Rs = Nothing
   Exit Function
   
ErrHand:
   GetNextProPerson = False
End Function

'********************************************************
'出缺勤系統
'檢查人員目前是否休假或離職, 若是, 則回傳職務代理人
'********************************************************
'Modify By Sindy 2012/11/15 +bolIsRest1Day 是否請整日
'Modify By Sindy 2015/11/18 + strB1109 原簽核主管
Public Function CheckPerCurrRestReturnPer(ByRef strB1104 As String, ByVal strB1003 As String, _
                                          Optional ByRef bolIsRest1Day As Boolean = False, _
                                          Optional ByVal strB1109 As String = "") As Boolean
Dim i As Integer, j As Integer, strTemp As String
Dim rsTmp As New ADODB.Recordset
Dim strOrgB1104 As String 'Add By Sindy 2013/10/17
   
   CheckPerCurrRestReturnPer = False
   
   strOrgB1104 = strB1104 'Add By Sindy 2013/10/17 記錄原簽核人員
   
   'Modify By Sindy 2012/11/15 +bolIsRest1Day 是否請整日
   If CheckIsPersonRest(strB1104, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolIsRest1Day) = True Or _
      ChkStaffST04(strB1104, False) = True Then
      CheckPerCurrRestReturnPer = True 'Add By Sindy 2022/3/23
      '以主管簽核特殊身份職代優先(1.審核主管+當事人 2.審核主管+部門別 3.審核主管+審核主管)，若無，再依人事職代
      'Modify By Sindy 2012/2/14 增加3.審核主管+審核主管
'      strSql = "SELECT '1',B0202,B0203,B0204,B0205,B0206,B0207 FROM ABS002 WHERE B0201='" & strB1104 & "' and B0208='" & strB1003 & "' " & _
'               "Union " & _
'               "SELECT '2',B0202,B0203,B0204,B0205,B0206,B0207 FROM ABS002 WHERE B0201='" & strB1104 & "' and B0208='" & PUB_GetST03(strB1003) & "' " & _
'               "Union " & _
'               "SELECT '3',B0202,B0203,B0204,B0205,B0206,B0207 FROM ABS002 WHERE B0201='" & strB1104 & "' and B0208='" & strB1104 & "' " & _
'               "Union " & _
'               "SELECT '4',B0102,B0103,B0104,B0105,B0106,B0107 FROM ABS001 WHERE B0101='" & strB1104 & "' " & _
'               "order by 1 asc"
      'Modify By Sindy 2012/3/20 王副總不要另設審核主管職代
      'Modify By Sindy 2015/11/18 strB1104 ==> IIf(strB1109 <> "", strB1109, strB1104)
      'Modify By Sindy 2023/5/4 + and B0209='1':1.人事
      'Modify By Sindy 2025/2/24 PUB_GetST03(strB1003) => 改用 PUB_GetST93(strB1003)
      strSql = "SELECT '1',B0202,B0203,B0204,B0205,B0206,B0207 FROM ABS002 WHERE B0201='" & IIf(strB1109 <> "", strB1109, strB1104) & "' and B0208='" & strB1003 & "' and B0209='1' " & _
               "Union " & _
               "SELECT '2',B0202,B0203,B0204,B0205,B0206,B0207 FROM ABS002 WHERE B0201='" & IIf(strB1109 <> "", strB1109, strB1104) & "' and B0208='" & PUB_GetST93(strB1003) & "' and B0209='1' " & _
               "Union " & _
               "SELECT '4',B0102,B0103,B0104,B0105,B0106,B0107 FROM ABS001 WHERE B0101='" & IIf(strB1109 <> "", strB1109, strB1104) & "' " & _
               "order by 1 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         strB1104 = ""
         'Modify By Sindy 2025/1/9
         Do While Not rsTmp.EOF
         '2025/1/9 END
            For i = 1 To 3
               For j = 1 To 2
                  strTemp = ""
                  If i = 1 And j = 1 Then If Not IsNull(rsTmp.Fields(1)) Then strTemp = rsTmp.Fields(1)
                  If i = 1 And j = 2 Then If Not IsNull(rsTmp.Fields(2)) Then strTemp = rsTmp.Fields(2)
                  If i = 2 And j = 1 Then If Not IsNull(rsTmp.Fields(3)) Then strTemp = rsTmp.Fields(3)
                  If i = 2 And j = 2 Then If Not IsNull(rsTmp.Fields(4)) Then strTemp = rsTmp.Fields(4)
                  If i = 3 And j = 1 Then If Not IsNull(rsTmp.Fields(5)) Then strTemp = rsTmp.Fields(5)
                  If i = 3 And j = 2 Then If Not IsNull(rsTmp.Fields(6)) Then strTemp = rsTmp.Fields(6)
                  If strTemp <> "" And strTemp <> strB1003 Then '不能為請假當事人
                     If CheckIsPersonRest(strTemp, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False And _
                        ChkStaffST04(strTemp, False) = False Then
                        strB1104 = strB1104 & strTemp & ","
                     End If
                  End If
               Next j
               If strB1104 <> "" Then
                  strB1104 = Left(strB1104, Len(strB1104) - 1)
                  CheckPerCurrRestReturnPer = True
                  Exit Do
               End If
            Next i
         'Modify By Sindy 2025/1/9
            rsTmp.MoveNext
         Loop
         '2025/1/9 END
      End If
      'Add By Sindy 2013/10/17 若簽核人員休假但職代也同時均休假,則維持原簽核人員
      If strB1104 = "" Then
         strB1104 = strOrgB1104
      End If
      '2013/10/17 END
      rsTmp.Close
      Set rsTmp = Nothing
   End If
End Function

'E-Mail簽核通知內容
'Add By Sindy 2019/5/24
'm_EditMode=1.新增資料
'           2.修改資料
'           3.刪除資料
Public Function GetEMailContent(ByVal strB1001 As String, ByRef strSubject As String, _
   Optional ByVal strSendKind As String = "", Optional ByVal strSendText As String = "", _
   Optional ByVal bolGetOnlyText As Boolean = False, _
   Optional ByVal strFomType As String, _
   Optional ByVal strKey1 As String, Optional ByVal StrKey2 As String, Optional ByVal strKey3 As String, _
   Optional ByVal m_EditMode As Integer = 0) As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strTemp As String, strText As String
   
   GetEMailContent = "": strTemp = "": strText = ""
   strSql = "SELECT * FROM ABS010,allcode WHERE B1001 = '" & strB1001 & "' and ac01(+)='04' and B1008=ac02(+) "
   'Modify By Sindy 2019/5/23
   If strB1001 = "" And strFomType <> "" Then
      If strFomType = 表單類別_請假 Then
         strSql = "SELECT '01' as B1002,sa01 as B1003,sa02 as B1004,sa03 as B1005,sa04 as B1006,sa05 as B1007,sa07 as B1009,sa08 as B1010,0 as B1012,0 as B1013,' ' as B1018,ac03" & _
                  " FROM staff_absence,allcode WHERE sa01='" & strKey1 & "'" & _
                  " and sa02=" & StrKey2 & " and sa03=" & strKey3 & _
                  " and ac01(+)='04' and sa06=ac02(+)"
      ElseIf strFomType = 表單類別_加班 Then
         strSql = "SELECT '02' as B1002,so01 as B1003,so02 as B1004,so03 as B1005,0 as B1006,so04 as B1007,0 as B1009,0 as B1010,so05 as B1012,so06 as B1013,' ' as B1018,' ' as ac03" & _
                  " FROM staff_overtime WHERE so01='" & strKey1 & "'" & _
                  " and so02=" & StrKey2 & " and so03=" & strKey3
      Else '出差
         strSql = "SELECT '03' as B1002,sb01 as B1003,sb02 as B1004,sb03 as B1005,sb04 as B1006,sb05 as B1007,sb06 as B1009,sb07 as B1010,0 as B1012,0 as B1013,' ' as B1018,' ' as ac03" & _
                  " FROM staff_busi_trip WHERE sb01='" & strKey1 & "'" & _
                  " and sb02=" & StrKey2 & " and sb03=" & strKey3
      End If
   End If
   '2019/5/23 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'Modify By Sindy 2016/11/16 半型空格在Outlook2007會被截掉,字全連在一起;改用全型空格
      '迄止日期
      If Not IsNull(rsTmp.Fields("B1006")) Then strTemp = Format(ChangeWStringToTString(rsTmp.Fields("B1006")), "###/##/##") & "　"
      '日,時
      If rsTmp.Fields("B1002") = "02" Then
         strText = "時數：" & IIf(IsNull(rsTmp.Fields("B1012")), rsTmp.Fields("B1013"), rsTmp.Fields("B1012"))
      Else
         strText = rsTmp.Fields("B1009") & "日" & rsTmp.Fields("B1010") & "時"
      End If
      GetEMailContent = GetPrjSalesNM(rsTmp.Fields("B1003")) & "(" & _
                        IIf(IsNull(rsTmp.Fields("ac03")), Right(GetB1002Value(rsTmp.Fields("B1002")), 2), rsTmp.Fields("ac03")) & "　" & _
                        Format(ChangeWStringToTString(rsTmp.Fields("B1004")), "###/##/##") & "　" & _
                        Format(Right("00" & rsTmp.Fields("B1005"), 4), "##:##") & _
                        "　～　" & strTemp & _
                        Format(Right("00" & rsTmp.Fields("B1007"), 4), "##:##") & _
                        "　" & strText & ")"
      
      'Add By Sindy 2011/10/7 因為配合劉經理要收到國外大陸出差的假單通知,只取表單內容
      If bolGetOnlyText = True Then
         rsTmp.Close
         Set rsTmp = Nothing
         Exit Function
      End If
      
      If Trim(strSendKind) = "" Then strSendKind = rsTmp.Fields("B1018")
      If m_EditMode = 3 And strSendKind <> 註銷 Then
         strSendKind = "刪除"
      ElseIf strB1001 = "" And Trim(strSendKind) = "" Then
         If strFomType = 表單類別_請假 Then
            strSendKind = "請假"
         ElseIf strFomType = 表單類別_加班 Then
            strSendKind = "加班"
         ElseIf strFomType = 表單類別_出差 Then
            strSendKind = "出差"
         End If
      End If
      Select Case strSendKind
         Case 會簽職代, 主管審核中, 送人事處簽收:
            GetEMailContent = GetEMailContent & "電子簽核通知"
         Case 退回:
            GetEMailContent = GetEMailContent & "電子簽核退回通知"
         Case 已核准:
            GetEMailContent = GetEMailContent & "已核准"
         Case 註銷:
            GetEMailContent = GetEMailContent & "已註銷"
         Case 重送:
            GetEMailContent = GetEMailContent & "電子簽核重送通知"
         Case 重送更改通知, 退回通知主管:
            GetEMailContent = GetEMailContent & "電子簽核" & strSendText
         Case Else
            GetEMailContent = GetEMailContent & IIf(strB1001 <> "", "電子簽核", "") & strSendKind & "通知"
      End Select
      'strSubject = GetEMailContent
      strSubject = Replace(GetEMailContent, "　", " ")
      strSubject = Replace(strSubject, "～", "~")
      
      If strB1001 <> "" Then 'Add By Sindy 2019/5/23
         'Add By Sindy 2013/3/5 若流程備註裡有除了已核准的訊息外,才顯示流程備註內容
         If strSendKind = 已核准 Then
            strSql = "select count(*) from abs012 where b1201='" & strB1001 & "' and B1207 is not null"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If rsTmp.Fields(0) > 0 Then
                  'Modify By Sindy 2023/12/29
                  If strSrvDate(1) >= 新部門啟用日 Then
                     strSql = "SELECT sqldateT(B1204),sqltime(B1205),decode(B1206,'05','',nvl(A0922,ST02))," & B1206CName & ",B1207 FROM ABS012,Staff,Acc090NEW WHERE B1201='" & strB1001 & "' and B1207 is not null and B1203=ST01(+) and B1203=A0921(+) order by B1202 asc"
                  Else
                  '2023/12/29 END
                     strSql = "SELECT sqldateT(B1204),sqltime(B1205),decode(B1206,'05','',nvl(A0902,ST02))," & B1206CName & ",B1207 FROM ABS012,Staff,Acc090 WHERE B1201='" & strB1001 & "' and B1207 is not null and B1203=ST01(+) and B1203=A0901(+) order by B1202 asc"
                  End If
                  intI = 1
                  Set rsTmp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     With rsTmp
                        .MoveFirst
                        GetEMailContent = GetEMailContent & vbCrLf & vbCrLf & "【意見或注意事項】" & vbCrLf
                        Do While Not .EOF
                           GetEMailContent = GetEMailContent & .Fields(0) & " " & _
                           Format(.Fields(1), "HH:MM:SS") & " " & _
                           IIf(IsNull(.Fields(2)), "", .Fields(2) & " ") & _
                           IIf(IsNull(.Fields(3)), "", .Fields(3) & " ") & _
                           IIf(Not IsNull(.Fields(4)) And .Fields(4) > "", IIf(IsNull(.Fields(2)) And IsNull(.Fields(3)), "", "："), "") & _
                           .Fields(4) & vbCrLf
                           .MoveNext
                        Loop
                     End With
                  End If
               End If
            End If
         End If
         '2013/3/5 End
         If strSendKind = 退回 Then
            GetEMailContent = GetEMailContent & vbCrLf & vbCrLf & "請至案件管理系統的一般作業項目中，進行表單退回處理。"
         ElseIf strSendKind = 退回通知主管 Then
            GetEMailContent = GetEMailContent & vbCrLf & vbCrLf & "請至案件管理系統的一般作業項目中，進行表單退回處理。(副本收件人純屬被通知，不須做任何處理)"
         ElseIf strSendKind <> 已核准 And strSendKind <> 註銷 And strSendKind <> 重送更改通知 Then
            GetEMailContent = GetEMailContent & vbCrLf & vbCrLf & "請至案件管理系統的一般作業項目中，進行表單簽核處理。"
         End If
      End If
'      MsgBox "Mail主旨 = " & strSubject & vbCrLf & "Mail本文 = " & GetEMailContent
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得員工的人事大部門別
' Input : strStuff ==> 員工的代碼
' Output : 傳回員工所屬的人事大部門別
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modify By Sindy 2023/12/19 +, ByRef strA0925 As String
Public Function GetStaffA0911(ByVal strStuff As String, ByRef strA0925 As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetStaffA0911 = Empty
   
   'Modify By Sindy 2023/12/19 加抓A0925
   strSql = "SELECT A0911,A0925 FROM Staff,Acc090,Acc090NEW " & _
            "WHERE ST01 = '" & strStuff & "' and ST03=A0901(+) and ST93=A0921(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("A0911")) = False Then
         GetStaffA0911 = rsTmp.Fields("A0911")
      End If
      If IsNull(rsTmp.Fields("A0925")) = False Then
         strA0925 = rsTmp.Fields("A0925")
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Function

'表單狀態對照表
Public Sub GetB1018CodeOrCName(ByRef strCode As String, ByRef strCName As String)
   If strCode <> "" Then
      Select Case strCode
         Case "01"
            strCName = "會簽職代"
         Case "02"
            strCName = "主管審核中"
         Case "03"
            strCName = "退回"
         Case "04"
            strCName = "送人事處簽收"
         Case "05"
            strCName = "已核准"
         Case "06"
            strCName = "註銷"
         Case "07"
            strCName = "重送"
         Case "08"
            strCName = "主管代填"
         Case Else
            strCName = strCode
      End Select
   ElseIf strCName <> "" Then
      Select Case strCName
         Case "會簽職代"
            strCode = "01"
         Case "主管審核中"
            strCode = "02"
         Case "退回"
            strCode = "03"
         Case "送人事處簽收"
            strCode = "04"
         Case "已核准"
            strCode = "05"
         Case "註銷"
            strCode = "06"
         Case "重送"
            strCode = "07"
         Case "主管代填"
            strCode = "08"
         Case Else
            strCode = strCName
      End Select
   End If
End Sub

'傳回表單確認(每月出缺勤統計確認)的下一處理人員
Public Function GetNextB1303(strKEY01 As String, strKEY02 As String) As String
Dim rsTmp As New ADODB.Recordset

   GetNextB1303 = ""
   'Modify By Sindy 2015/7/7 +and nvl(B0112,0)<99 ...
   strSql = "SELECT 1,B0108 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and B0108 is not null and B1304 is null and nvl(B0112,0)<99 " & _
            "Union " & _
            "SELECT 2,B0109 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and B0109 is not null and B1305 is null and nvl(B0113,0)<99 " & _
            "Union " & _
            "SELECT 3,B0110 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and B0110 is not null and B1306 is null and nvl(B0114,0)<99 " & _
            "Union " & _
            "SELECT 4,B0111 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and B0111 is not null and B1307 is null and nvl(B0115,0)<99 " & _
            "order by 1 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GetNextB1303 = rsTmp.Fields(1)
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'傳回表單確認(每月出缺勤統計確認)的目前處理人員其審核主管為何
Public Function GetCurrB1303Seqno(strKEY01 As String, strKEY02 As String, strKEY03 As String) As Integer
Dim rsTmp As New ADODB.Recordset

   GetCurrB1303Seqno = 0

   strSql = "SELECT 1,B0108 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and '" & strKEY03 & "'=B0108(+) " & _
            "Union " & _
            "SELECT 2,B0109 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and '" & strKEY03 & "'=B0109(+) " & _
            "Union " & _
            "SELECT 3,B0110 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and '" & strKEY03 & "'=B0110(+) " & _
            "Union " & _
            "SELECT 4,B0111 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and '" & strKEY03 & "'=B0111(+) " & _
            "order by 1 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      With rsTmp
         .MoveFirst
         Do While Not .EOF
            If Not IsNull(rsTmp.Fields(1)) Then
               GetCurrB1303Seqno = rsTmp.Fields(0): Exit Do
            End If
            .MoveNext
         Loop
      End With
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'取得職代檔裡是第幾組的職代
Public Function GetPersonSeqno(strKEY01 As String, strKEY02 As String) As String
Dim rsTmp As New ADODB.Recordset
   
   GetPersonSeqno = ""

   strSql = "select '1' from ABS001 where B0101='" & strKEY01 & "' and (B0102='" & strKEY02 & "' or B0103='" & strKEY02 & "') " & _
            "Union " & _
            "select '2' from ABS001 where B0101='" & strKEY01 & "' and (B0104='" & strKEY02 & "' or B0105='" & strKEY02 & "') " & _
            "Union " & _
            "select '3' from ABS001 where B0101='" & strKEY01 & "' and (B0106='" & strKEY02 & "' or B0107='" & strKEY02 & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GetPersonSeqno = rsTmp.Fields(0)
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'檢查人事系統裡是否已有此表單編號
Public Function ChkPerSysB1001Exist(strKey1 As String, StrKey2 As String, Optional bolMsg As Boolean = True) As Boolean
Dim rsTmp As New ADODB.Recordset
   
   ChkPerSysB1001Exist = False
   
   strSql = "select sa09 from staff_absence where sa09='" & strKey1 & "' and sa01='" & StrKey2 & "' " & _
            "union select so13 from staff_overtime where so13='" & strKey1 & "' and so01='" & StrKey2 & "' " & _
            "union select sb10 from staff_busi_trip where sb10='" & strKey1 & "' and sb01='" & StrKey2 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ChkPerSysB1001Exist = True
      If bolMsg = True Then MsgBox "此表單人事處已先行作業，不可再修改請假內容！"
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'檢查出缺勤系統裡是否已有此表單編號
Public Function ChkAbsSysB1001Exist(strKey1 As String, StrKey2 As String, strKey3 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
   
   ChkAbsSysB1001Exist = False
   
   strSql = "select B1001,B1002 from ABS010 where B1001='" & strKey1 & "' and B1003='" & strKey3 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If rsTmp.Fields("B1002") <> StrKey2 Then
         MsgBox "表單類別不符！"
      Else
         ChkAbsSysB1001Exist = True
      End If
   Else
      MsgBox "表單編號不存在！"
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'取得假別資料
Public Function GetAllCode04(strKey1 As String) As String
Dim rsTmp As New ADODB.Recordset
   
   GetAllCode04 = ""
   
   strSql = "select ac02||' '||ac03 from allcode where ac01='04' and ac02='" & strKey1 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GetAllCode04 = rsTmp.Fields(0)
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'********************************************************
'案件系統:
'檢查人員是否有休假,取得工作職代
'********************************************************
'Modify By Sindy 2012/9/13 +strRestKind休假狀況:1.請假 3.出差
'Modify by Sindy 2016/10/12 + bolInSpecialDuty:是否含特殊職代
'Modify by Sindy 2020/7/20 + , Optional strDutyAgentKind As String = "" : 指定職代種類; 1.人事職代 2.案件職代 A.指定抓全部職代(2021/11/12)
'Modify by Sindy 2024/3/27 + , Optional ByVal stCaseNo As String = "" : 傳入本所案號
'Modify By Sindy 2025/2/24 此函數 bolInSpecialDuty 已取消使用
Public Function GetCaseDutyAgent(ByVal stReceiver As String, ByVal stReceiveNo As String, _
         Optional ByVal bolMsg As Boolean = True, Optional ByRef strRestKind As String, _
         Optional bolInSpecialDuty As Boolean = False, Optional strDutyAgentKind As String = "", _
         Optional ByVal stCaseNo As String = "") As String
Dim intMaxPerson As Integer, i As Integer, j As Integer, k As Integer
Dim StrST01 As Variant, strTemp As Variant
Dim strData As String, strGetPer As String, strText As String
Dim strTime As String
Dim m_ABS001_1 As String
Dim m_ABS001_2 As String
Dim m_ABS001_3 As String
Dim strKind As String
Dim rsTmp As New ADODB.Recordset
Dim bolSendMailAll As Boolean 'Add By Sindy 2015/10/19
Dim strCompTime As String, strCompDate As String
Dim ii As Integer 'Add By Sindy 2016/10/12
Dim jj As Integer 'Add By Sindy 2025/5/26
Dim tmpDutyAgent As String 'Add By Sindy 2018/1/3
'Add By Sindy 2024/7/29 職代的休假狀況
Dim strRestKind_CC As String, bolHadPerWork As Boolean, bolAllRestTo4 As Boolean
Dim bolRest As Boolean
'2024/7/29 END
   
   'Add By Sindy 2024/8/1
   '預設值:
   bolHadPerWork = False '判斷是否有收受者上班
   bolAllRestTo4 = True '收受者和職代全部因颱風天停止上班
   '2024/8/1 END
   
   If InStr(1, stReceiver, ",") > 0 Then
      StrST01 = Split(stReceiver & ",", ",")
      intMaxPerson = UBound(StrST01) - 1
   Else
      StrST01 = Split(stReceiver & ";", ";")
      intMaxPerson = UBound(StrST01) - 1
   End If
   
   '目前系統日期時間
   strDate = strSrvDate(1)
   strTime = Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)
   'Add By Sindy 2012/4/12 若過下班時間,檢查日期改為下一個工作天,時間預設為07:30
   'Modify By Sindy 2017/7/10 增加檢查非工作日,檢查日期改為下一個工作天,時間預設為07:30
   If Val(Format(strTime, "hhmm")) > 1800 Or ChkWorkDay(strDate) = False Then
      strDate = Format(DateAdd("d", 1, ChangeWStringToWDateString(strDate)), "YYYYMMDD") '+1天 Add By Sindy 2017/7/26
      Do While ChkWorkDay(strDate) = False
         strDate = CompWorkDay(1, strDate, 0)
      Loop
      strTime = "07:30"
   End If
   '2017/7/10 END
   '2012/4/12 End
   
   For i = 0 To intMaxPerson
      '先檢查收件人是否當日休假
      If Trim(CStr(StrST01(i))) = "" Then Exit For
      'Modify By Sindy 2012/9/13 +strRestKind休假狀況
      If CheckIsPersonRest(CStr(StrST01(i)), strDate, strTime, strRestKind) = True Then
         strRestKind_CC = "" 'Add By Sindy 2024/8/1
         tmpDutyAgent = "" 'Add By Sindy 2018/1/3
         
         'Add By Sindy 2015/10/19
         If InStr(strText, GetPrjSalesNM(CStr(StrST01(i)))) = 0 Then
         '2015/10/19 END
            strText = strText & GetPrjSalesNM(CStr(StrST01(i))) '& "、" 'Modify By Sindy 2012/10/3
         Else
            GoTo ReadNext 'Add By Sindy 2024/9/9 收件者重覆了
         End If
         
         'Add By Sindy 2016/10/12
         'Modify By Sindy 2025/2/24 應該都要 含特殊職代 做檢查
'         If bolInSpecialDuty = True Then '含特殊職代
         '2025/2/24 END
            'Modify By Sindy 2022/12/21 + and B0208='" & strUserNum & "' : 簽核對象 + 部門判斷
            'Modify By Sindy 2023/5/3 增加 3.審核主管+審核主管
            'Modify By Sindy 2023/5/4 + and B0209='2'
            'Modify By Sindy 2025/2/24 PUB_GetStaffST15(strUserNum, "1") => 改用 PUB_GetST93(strUserNum)
            strSql = "select B0202,B0203,B0204,B0205,B0206,B0207,1 SoType from abs002 where b0201='" & StrST01(i) & "' and B0208='" & strUserNum & "' and B0209='2'" & _
                     " union select B0202,B0203,B0204,B0205,B0206,B0207,2 SoType from abs002 where b0201='" & StrST01(i) & "' and B0208='" & PUB_GetST93(strUserNum) & "' and B0209='2'" & _
                     " union select B0202,B0203,B0204,B0205,B0206,B0207,3 SoType from abs002 where b0201='" & StrST01(i) & "' and B0208='" & StrST01(i) & "' and B0209='2'" & _
                     " order by SoType asc"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               'Modify By Sindy 2025/5/26
               k = 0
               'For ii = 0 To 5
               For ii = 1 To 3 '僅三組職代
                  For jj = 1 To 2 '雙職代
                     If "" & rsTmp.Fields(k) <> "" Then
                        If ChkStaffST04(rsTmp.Fields(k), False) = False Then '人員在職時,才須再往下檢查
                           'Modify By Sindy 2024/7/29 +strRestKind_CC: 職代的休假狀況
                           'If CheckIsPersonRest(rsTmp.Fields(k), strDate, strTime, strRestKind_CC) = False Then
                           bolRest = CheckIsPersonRest(rsTmp.Fields(k), strDate, strTime, strRestKind_CC)
                           '判斷職代為颱風假時,不再往下檢查休假問題
                           If bolRest = False Or strRestKind_CC = "4" Then
                           '2024/7/29 END
                              tmpDutyAgent = tmpDutyAgent & rsTmp.Fields(k) & ";"
                           End If
                        End If
                     End If
                     k = k + 1
                  Next jj
                  If tmpDutyAgent <> "" Then
                     Exit For
                  End If
               '2025/5/26 END
               Next ii
            End If
            rsTmp.Close
'         End If
         
         'Add By Sindy 2023/4/28
         '清除職代組數
         For ii = 1 To intDutyItem
            PubABS001_1(ii) = ""
         Next ii
         PubABS001_A = "": PubABS001_B = "" '雙職代的A,B區
         '2023/4/28 END
         If tmpDutyAgent = "" Then
         '2016/10/12 END
            '有休假時,才需要讀取案件職代或一般人事職代作寄件副本
            strData = "": strGetPer = ""
            m_ABS001_1 = "": m_ABS001_2 = "": m_ABS001_3 = "": strKind = ""
            
            If stReceiveNo <> "" Then
               '先取得案件申請國家
               strSql = "select PA09 from caseprogress,patent where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp09='" & stReceiveNo & "' " & _
                  "Union select TM10 from caseprogress,trademark where cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and cp09='" & stReceiveNo & "' " & _
                  "Union select SP09 from caseprogress,servicepractice where cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and cp09='" & stReceiveNo & "' " & _
                  "Union select LC15 from caseprogress,lawcase where cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and cp09='" & stReceiveNo & "' "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount > 0 Then
                  If Trim("" & rsTmp.Fields(0)) <= "010" Then
                     strKind = "1" '1.案件申請國家為台灣
                  Else
                     strKind = "2" '2.案件申請國家為”非”台灣
                  End If
               End If
               rsTmp.Close
            'Modify by Sindy 2024/3/27 + , Optional ByVal stCaseNo As String = "" : 傳入本所案號
            ElseIf stCaseNo <> "" Then
               strKind = GetPrjNation1(stCaseNo)
               If strKind <> "" Then
                  If strKind <= "010" Then
                     strKind = "1" '1.案件申請國家為台灣
                  Else
                     strKind = "2" '2.案件申請國家為”非”台灣
                  End If
               End If
            '2024/3/27 END
            End If
            
            'Get職代檔
            'Modify By Sindy 2015/10/19 +B0126.案件職代是否全部發mail通知
            'strSql = "select B0117 from abs001 where b0101='" & strST01(i) & "'"
            strSql = "select * from abs001 where b0101='" & StrST01(i) & "'"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               'Add By Sindy 2015/10/19
               If "" & rsTmp.Fields("B0126") = "Y" Then
                  bolSendMailAll = True
               Else
                  bolSendMailAll = False
               End If
               '2015/10/19 END
               'Added by Lydia 2021/11/12 A.指定抓全部職代; 用在frm090127,frm090128內商查名之權限管控
               If strDutyAgentKind = "A" Then
                   bolSendMailAll = True
               End If
               'end 2021/11/12
               
               'Modify by Sindy 2020/7/20 + 指定職代種類: 1.人事職代 2.案件職代
               If strDutyAgentKind = "1" Then
                  '(2)一般人事職代
                  Call GetABS001_1(CStr(StrST01(i)), m_ABS001_1, m_ABS001_2, m_ABS001_3)
               ElseIf strDutyAgentKind = "2" Then
                  '(1)案件職代
                  Call GetABS001_3(CStr(StrST01(i)), m_ABS001_1, m_ABS001_2, m_ABS001_3, strKind)
               Else
               '2020/7/20 END
                  If Not IsNull(rsTmp.Fields("B0117")) And "" & rsTmp.Fields("B0117") > "" Then
                     '(1)案件職代
                     Call GetABS001_3(CStr(StrST01(i)), m_ABS001_1, m_ABS001_2, m_ABS001_3, strKind)
                  Else
                     '(2)一般人事職代
                     Call GetABS001_1(CStr(StrST01(i)), m_ABS001_1, m_ABS001_2, m_ABS001_3)
                  End If
               End If
            End If
            rsTmp.Close
            
            'Modify By Sindy 2017/8/24
            'If InStr(m_ABS001_1, ",") > 0 And bolSendMailAll = False Then '雙職代 且 沒有要全發
            If bolSendMailAll = False Then '沒有要全發
               For j = 1 To intDutyItem 'n組職代
                  strGetPer = ""
                  If PubABS001_1(j) <> "" Then
                     '檢查取得的職代是否當日休假,若有,則找下一職代
                     strTemp = Split(PubABS001_1(j), ",")
                     For k = 0 To UBound(strTemp)
                        '不可當日休假
                        If Not IsNull(strTemp(k)) Then
                           'Modify By Sindy 2024/7/29 +strRestKind_CC: 職代的休假狀況
                           'If CheckIsPersonRest(CStr(strTemp(k)), strDate, strTime, strRestKind_CC) = False Then
                           bolRest = CheckIsPersonRest(CStr(strTemp(k)), strDate, strTime, strRestKind_CC)
                           '判斷職代為颱風假時,不再往下檢查休假問題
                           If bolRest = False Or strRestKind_CC = "4" Then
                           '2024/7/29 END
                              If k = UBound(strTemp) Then '最後一位時才一起寫入職代資料
                                 For ii = 0 To UBound(strTemp)
                                    If strGetPer = "" Then
                                       strGetPer = strTemp(ii) & ";"
                                    ElseIf InStr(strGetPer, strTemp(ii)) = 0 Then
                                       strGetPer = strGetPer & strTemp(ii) & ";"
                                    End If
                                 Next ii
                              End If
                           Else
                              Exit For
                           End If
                        Else
                           Exit For
                        End If
                     Next k
                     If strGetPer <> "" Then
                        If InStr(tmpDutyAgent, strGetPer) = 0 Then
                           tmpDutyAgent = tmpDutyAgent & strGetPer
                        End If
                        'Add By Sindy 2015/10/19
                        If bolSendMailAll = False Then
                        '2015/10/19 END
                           Exit For '若有取得職代,則離開迴圈
                        End If
                     End If
                  Else
                     Exit For '離開迴圈
                  End If
               Next j
            Else
            '2017/8/24 END
               For j = 1 To 3 '有3組職代
                  strData = "": strGetPer = "" 'Add By Sindy 2015/10/19
                  If j = 1 And m_ABS001_1 <> "" Then strData = m_ABS001_1
                  If j = 2 And m_ABS001_2 <> "" Then strData = m_ABS001_2
                  If j = 3 And m_ABS001_3 <> "" Then strData = m_ABS001_3
                  If strData <> "" Then
                     '檢查取得的職代是否當日休假,若有,則找下一職代
                     strTemp = Split(strData, ",")
                     For k = 0 To UBound(strTemp)
                        '不可當日休假
                        If Not IsNull(strTemp(k)) Then
                           'Modify By Sindy 2024/7/29 +strRestKind_CC: 職代的休假狀況
                           'If CheckIsPersonRest(CStr(strTemp(k)), strDate, strTime, strRestKind_CC) = False Then
                           bolRest = CheckIsPersonRest(CStr(strTemp(k)), strDate, strTime, strRestKind_CC)
                           '判斷職代為颱風假時,不再往下檢查休假問題
                           If bolRest = False Or strRestKind_CC = "4" Then
                           '2024/7/29 END
                              'Modify By Sindy 2015/8/14 敏莉:FCP-52049反應她沒有收到mail,因她是職代
                              'Add By Sindy 2015/6/15 職代若為操作者,不需再收到E-Mail
                              'If strTemp(k) <> strUserNum Then
                              'Sindy 2015/8/14 END
                                 If strGetPer = "" Then
                                    strGetPer = strTemp(k) & ";"
                                 ElseIf InStr(strGetPer, strTemp(k)) = 0 Then
                              '2015/6/15 END
                                    strGetPer = strGetPer & strTemp(k) & ";"
                                 End If
                              'End If
                           End If
                        End If
                     Next k
                     If strGetPer <> "" Then
                        If InStr(tmpDutyAgent, strGetPer) = 0 Then
                           tmpDutyAgent = tmpDutyAgent & strGetPer
                        End If
                        'Add By Sindy 2015/10/19
                        If bolSendMailAll = False Then
                        '2015/10/19 END
                           Exit For '若有取得職代,則離開迴圈
                        End If
                     End If
                  End If
               Next j
            End If
            
            'Add By Sindy 2016/5/25 增加假單裡的職代
            If tmpDutyAgent = "" Then
               '注意比較的日期格式為 YYYYMMDDHHMM
               strCompTime = IIf(strTime = "24:00", "2400", Format(strTime, "hhmm"))
               strCompDate = strSrvDate(1) & strCompTime
               'modify by sonia 2017/6/3 A0029之6/3假單職代A1010於當日離職故要剔除
               strSql = "SELECT B1104 FROM ABS011,STAFF" & _
                        " where B1101 in(" & _
                        " SELECT SA09 From staff_Absence" & _
                        " WHERE " & strCompDate & " between SA02||substr('0000'||SA03,-4) and SA04||substr('0000'||SA05,-4) and SA01='" & Trim(CStr(StrST01(i))) & "'" & _
                        " Union" & _
                        " SELECT SB10 From staff_busi_trip" & _
                        " WHERE " & strCompDate & " between SB02||substr('0000'||SB03,-4) and SB04||substr('0000'||SB05,-4) and SB01='" & Trim(CStr(StrST01(i))) & "'" & _
                        ") and B1102='1' and b1104=st01(+) and st04='1' order by B1103 asc"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  With RsTemp
                     .MoveFirst
                     Do While Not .EOF
                        If Not IsNull(RsTemp.Fields(0)) Then
                           '不可當日休假
                           'Modify By Sindy 2024/7/29 +strRestKind_CC: 職代的休假狀況
                           'If CheckIsPersonRest(CStr(RsTemp.Fields(0)), strDate, strTime, strRestKind_CC) = False Then
                           bolRest = CheckIsPersonRest(CStr(RsTemp.Fields(0)), strDate, strTime, strRestKind_CC)
                           '判斷職代為颱風假時,不再往下檢查休假問題
                           If bolRest = False Or strRestKind_CC = "4" Then
                           '2024/7/29 END
                              If bolSendMailAll = False Then
                                 tmpDutyAgent = tmpDutyAgent & Trim(RsTemp.Fields(0)) & ";"
                                 Exit Do '若有取得職代,則離開迴圈
                              Else
                                 If InStr(tmpDutyAgent, Trim(RsTemp.Fields(0))) = 0 Then
                                    tmpDutyAgent = tmpDutyAgent & Trim(RsTemp.Fields(0)) & ";"
                                 End If
                              End If
                           End If
                        End If
                        .MoveNext
                     Loop
                  End With
               End If
            End If
            '2016/5/25 END
         End If
         
         'Add By Sindy 2021/6/3 加居家職代
         strSql = "select * from abs001 where b0101='" & StrST01(i) & "'"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If "" & rsTmp.Fields("B0127") = "Y" Then '居家無法連線
               If Len(Trim("" & rsTmp.Fields("B0128"))) > 0 Then
                   tmpDutyAgent = tmpDutyAgent & Trim(rsTmp.Fields("B0128")) & ";"
               End If
               If Len(Trim("" & rsTmp.Fields("B0129"))) > 0 Then
                   tmpDutyAgent = tmpDutyAgent & Trim(rsTmp.Fields("B0129")) & ";"
               End If
            End If
         End If
         rsTmp.Close 'Add By Sindy 2021/6/7
         '2021/6/3 END
         
         'Add By Sindy 2017/6/7 收件人休假,無讀取到副本收件人,改發部門主管
         If tmpDutyAgent = "" Then
            'Add By Sindy 2022/10/24 改職代也同時請假時，改發設定表的審核主管(抓沒有設定天數的審核主管)，若審核主管也都請假才抓部門主管A0908
            '但抓到設定表的審核主管時，若該審核主管的部門前二碼與請假人的部門前二碼不同時，則直接抓A0908。
            strSql = "select st01,st02,st15,st04,1 as sort from abs001,staff where B0101='" & StrST01(i) & "' and B0108 is not null and nvl(B0112,0)=0 and B0108=st01(+) and st04='1'" & _
                     " union all select st01,st02,st15,st04,2 as sort from abs001,staff where B0101='" & StrST01(i) & "' and B0109 is not null and nvl(B0113,0)=0 and B0109=st01(+) and st04='1'" & _
                     " union all select st01,st02,st15,st04,3 as sort from abs001,staff where B0101='" & StrST01(i) & "' and B0110 is not null and nvl(B0114,0)=0 and B0110=st01(+) and st04='1'" & _
                     " union all select st01,st02,st15,st04,4 as sort from abs001,staff where B0101='" & StrST01(i) & "' and B0111 is not null and nvl(B0115,0)=0 and B0111=st01(+) and st04='1'" & _
                     " order by sort asc"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               Do While Not rsTmp.EOF
                  If "" & rsTmp.Fields("st01") <> "" And "" & rsTmp.Fields("st15") <> "" Then
                     If Left(rsTmp.Fields("st15"), 2) <> Left(PUB_GetStaffST15(CStr(StrST01(i)), "1"), 2) Then
                        Exit Do
                     End If
                     '當日未休假
                     'Modify By Sindy 2024/7/29 +strRestKind_CC: 職代的休假狀況
                     'If CheckIsPersonRest(rsTmp.Fields("st01"), strDate, strTime, strRestKind_CC) = False Then
                     bolRest = CheckIsPersonRest(rsTmp.Fields("st01"), strDate, strTime, strRestKind_CC)
                     '判斷職代為颱風假時,不再往下檢查休假問題
                     If bolRest = False Or strRestKind_CC = "4" Then
                     '2024/7/29 END
                        tmpDutyAgent = rsTmp.Fields("st01") & ";"
                        Exit Do
                     End If
                  End If
                  rsTmp.MoveNext
               Loop
            End If
            rsTmp.Close
            If tmpDutyAgent = "" Then
            '2022/10/24 END
               'Modify By Sindy 2024/6/13
               'strSql = "select a0908 from staff,acc090 where st01='" & strST01(i) & "' and st15=a0901(+)"
               strSql = "select a0924 from staff,acc090new where st01='" & StrST01(i) & "' and st93=a0921(+)"
               '2024/6/13 END
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  If IsNull(RsTemp.Fields(0)) = False Then
                     tmpDutyAgent = RsTemp.Fields(0) & ";"
                  End If
               End If
            End If
         End If
         '2017/6/7 END
         
         If strRestKind <> "4" And strRestKind <> "" Then bolAllRestTo4 = False 'Add By Sindy 2024/8/1
         If strRestKind_CC <> "4" And strRestKind_CC <> "" Then bolAllRestTo4 = False 'Add By Sindy 2024/8/1
         
         'Add By Sindy 2012/10/3
         If strRestKind = "3" Then
            strText = strText & "出差" & "、"
         ElseIf strRestKind = "4" Then
            strText = strText & "颱風天停止上班" & "、"
         Else
            strText = strText & "請假" & "、"
         End If
         '2012/10/3 End
      Else
         bolHadPerWork = True '有收受者上班 Add By Sindy 2024/8/1
      End If
      'Add By Sindy 2018/1/3
      If tmpDutyAgent <> "" And InStr(GetCaseDutyAgent, tmpDutyAgent) = 0 Then
         GetCaseDutyAgent = GetCaseDutyAgent & tmpDutyAgent
      End If
      '2018/1/3 END
ReadNext: 'Add By Sindy 2024/9/9
   Next i
   If strText <> "" Then strText = Left(strText, Len(strText) - 1)
   If GetCaseDutyAgent <> "" Then GetCaseDutyAgent = Left(GetCaseDutyAgent, Len(GetCaseDutyAgent) - 1)
   If bolMsg = True And GetCaseDutyAgent <> "" And UCase(strUserNum) <> "QPGMR" Then '晚上批次發信則不必彈訊息
      'Modify By Sindy 2012/9/13 王副總提能否就實際狀況顯示,若出差則顯示出差
      'MsgBox strText & "請假，副本發案件職代！"
'      If strRestKind = "1" Then '請假
'         MsgBox strText & "請假，副本發案件職代！"
'      Else '出差
'         MsgBox strText & "出差，副本發案件職代！"
'      End If
      '2012/9/13 End
      'Modify By Sindy 2012/10/3
      MsgBox strText & "，副本發案件職代！"
      'strRestKind = strText 'Modify By Sindy 2013/10/3 Mark
   End If
   
   'Add By Sindy 2024/8/1
   If bolHadPerWork = False And bolAllRestTo4 = True Then '沒人上班,全部因颱風天停止上班
      strRestKind = "全體人員因颱風天停止上班"
   Else
   '2024/8/1 END
      strRestKind = strText 'Modify By Sindy 2013/10/3 Move
   End If
   
   Set rsTmp = Nothing
End Function

'Modify By Sindy 2012/1/3
Public Function PUB_AutoM21Receive(strUserNum As String, strUpdDate As String, strUpdTime As String _
, strB1001 As String, strB1002 As String, strB1003 As String _
, strB1004 As String, strB1005 As String, strB1006 As String, strB1007 As String _
, strB1008 As String, strB1009 As String, strB1010 As String _
, strB1030 As String, strB101213 As String _
, strB1014 As String, strB1015 As String _
, strB1028 As String, strB1029 As String _
, ByRef strSubject As String, ByRef strContent As String _
, Optional bolMsg As Boolean = True, Optional bolAutoBatch As Boolean = False) As Boolean

Dim strB1018 As String
Dim dblDay As Double, dbl_cYM As Double
Dim strMsg As String
Dim i As Integer ', strTo As String
Dim strUpdB1004 As String, strUpdB1005 As String, strUpdB1006 As String, strUpdB1007 As String
Dim strUpdB1009 As String, strUpdB1010 As String
Dim strUpdB1028 As String, strUpdB1029 As String
Dim m_Day As Integer, m_Hour As Double
'Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2013/7/12
'Dim varTemp As Variant 'Add By Sindy 2015/7/15
Dim strChkYm As String 'Add By Sindy 2015/12/25
Dim j As Integer 'Add By Sindy 2021/8/13
   
On Error GoTo ErrHand
   
   PUB_AutoM21Receive = True
   
   strB1018 = 已核准
   
   '流程備註檔
   'strSql = GetInsertABS012Sql(Trim(txtB1001), 人事處, strUpdDate, strUpdTime, m_B1018, "")
   strSql = GetInsertABS012Sql(Trim(strB1001), strUserNum, strUpdDate, strUpdTime, strB1018, "")
   cnnConnection.Execute strSql
   
   '更新出缺勤電子簽核主檔
   strSql = "update ABS010 set " & _
            "B1018=" & CNULL(strB1018) & _
            ",B1019=" & CNULL(strUserNum) & _
            ",B1020=" & CNULL(strUpdDate) & _
            ",B1021=" & CNULL(strUpdTime) & _
            " where B1001=" & CNULL(strB1001)
   cnnConnection.Execute strSql
   
   '先取得E-Mail主旨,本文內容
   strContent = GetEMailContent(strB1001, strSubject)
   
   '檢查人事系統裡是否已有表單編號,若資料已存在時,只純做表單簽收不更新人事資料
'   strMsg = "" '預設值
   If ChkPerSysB1001Exist(strB1001, strB1003, False) = False Then
      '有跨月份時
      If Not IsNull(strB1006) And Trim(strB1006) > "" And _
         (Left(DBDATE(strB1004), 6) <> Left(DBDATE(strB1006), 6)) Then
         dbl_cYM = Left(strB1004, 6) '目前年月
         '目前年月小於等於迄日年月時,就要執行下列迴圈程式
         Do While dbl_cYM <= Left(strB1006, 6)
            '系統取得起日
            If dbl_cYM = Left(strB1004, 6) Then
               strUpdB1004 = strB1004
               strUpdB1005 = strB1005
            Else
               For i = 1 To 31
                  dblDay = Val(dbl_cYM) * 100 + i
                  '正確的日期格式
                  If CheckIsTaiwanDate(ChangeWStringToTString(CStr(dblDay)), False) = True Then
                     'Add By Sindy 2012/9/21
                     If strB1002 = 表單類別_出差 Then
                        strUpdB1004 = CStr(dblDay)
                        '預設值
                        strUpdB1005 = "0000"
                        Exit For
                     Else
                     '2012/9/21 End
                        '工作天
                        'Modify By Sindy 2012/10/22 產假和流產假可輸入工作天
                        'If ChkWorkDay(dblDay) = True Then
                        If ChkWorkDay(dblDay) = True Or (strB1002 = 表單類別_請假 And (strB1008 = "10" Or strB1008 = "11")) Then
                        '2012/10/22 End
                           strUpdB1004 = CStr(dblDay)
                           '預設值
                           Call PUB_ChkByPassWork(PUB_GetST06(strB1003), strUpdB1004, , strUpdB1007, strUpdB1005) 'Add By Sindy 2021/8/13
                           'strUpdB1005 = Format(strByPassStarTime(1), "hhmm") '"800"
                           If InStr(strUpdB1005, ":") > 0 Then strUpdB1005 = Format(strUpdB1005, "hhmm")
                           Exit For
                        End If
                     End If
                  End If
               Next i
            End If
            '系統取得迄日
            If dbl_cYM = Left(strB1006, 6) Then
               strUpdB1006 = strB1006
               strUpdB1007 = strB1007
            Else
               For i = 31 To 1 Step -1
                  dblDay = Val(dbl_cYM) * 100 + i
                  '正確的日期格式
                  If CheckIsTaiwanDate(ChangeWStringToTString(CStr(dblDay)), False) = True Then
                     'Add By Sindy 2012/9/21
                     If strB1002 = 表單類別_出差 Then
                        strUpdB1006 = CStr(dblDay)
                        '預設值
                        strUpdB1007 = "2400"
                        Exit For
                     Else
                     '2012/9/21 End
                        '工作天
                        'Modify By Sindy 2012/10/22 產假和流產假可輸入工作天
                        'If ChkWorkDay(dblDay) = True Then
                        If ChkWorkDay(dblDay) = True Or (strB1002 = 表單類別_請假 And (strB1008 = "10" Or strB1008 = "11")) Then
                        '2012/10/22 End
                           strUpdB1006 = CStr(dblDay)
                           '預設值
                           Call PUB_ChkByPassWork(PUB_GetST06(strB1003), strUpdB1006, strUpdB1005, , , strUpdB1007) 'Add By Sindy 2021/8/13
                           'strUpdB1007 = Format(strByPassEndTime(1), "hhmm") '"1700"
                           If InStr(strUpdB1007, ":") > 0 Then strUpdB1007 = Format(strUpdB1007, "hhmm")
                           Exit For
                        End If
                     End If
                  End If
               Next i
            End If
            
            Call PUB_ChkByPassWork(PUB_GetST06(strB1003), strUpdB1004) 'Add By Sindy 2021/8/13
            
            '若有輸入起日上班時段及迄日下班時段時
'            cboSTime.Visible = False
'            cboETime.Visible = False
            strUpdB1028 = "": strUpdB1029 = ""
            If strB1028 <> "" And strB1029 <> "" Then
               '系統自動取得的起日和迄日為同一天
               If DBDATE(strUpdB1004) = DBDATE(strUpdB1006) Then
                  '若系統抓到的請假日為請假的第一天  ,必須計算出正確的下班時段
                  If DBDATE(strUpdB1004) = strB1004 Then
                     'Modify By Sindy 2012/9/25
                     strUpdB1028 = strB1028
'                     If CStr(strB1028) <= "0800" Then
'                        strUpdB1029 = "1700"
'                     ElseIf CStr(strB1028) > "0800" And CStr(strB1028) <= "0830" Then
'                        strUpdB1029 = "1730"
'                     ElseIf CStr(strB1028) > "0830" And CStr(strB1028) <= "0900" Then
'                        strUpdB1029 = "1800"
'                     End If
                     'Modify By Sindy 2021/8/13
                     For j = 1 To intByPassArea
                        If CStr(strB1028) = Format(strByPassStarTime(j), "hhmm") Then
                           strUpdB1029 = Format(strByPassEndTime(j), "hhmm")
                           Exit For
                        End If
                     Next j
                     '2021/8/13 END
                     
                     If strUpdB1005 <= strUpdB1029 Then
                        strUpdB1007 = strUpdB1029
'                        If CStr(strB1028) <= "0800" Then
'                           strUpdB1007 = "1700"
'                        ElseIf CStr(strB1028) > "0800" And CStr(strB1028) <= "0830" Then
'                           strUpdB1007 = "1730"
'                        ElseIf CStr(strB1028) > "0830" And CStr(strB1028) <= "0900" Then
'                           strUpdB1007 = "1800"
'                        End If
'                        'Add By Sindy 2012/4/11
'                        strUpdB1028 = strB1028
'                        strUpdB1029 = strUpdB1007
'                        '2012/4/11 End
                     End If
                     
                  '若系統抓到的請假日為請假的最後一天,必須計算出正確的上班時段
                  ElseIf DBDATE(strUpdB1004) = strB1006 Then
'                     If CStr(strB1029) >= "1700" And CStr(strB1029) <= "1729" Then
'                        strUpdB1005 = "800"
'                     ElseIf CStr(strB1029) >= "1730" And CStr(strB1029) <= "1759" Then
'                        strUpdB1005 = "830"
'                     ElseIf CStr(strB1029) >= "1800" Then
'                        strUpdB1005 = "900"
'                     End If
                     'Add By Sindy 2021/8/13
                     For j = 1 To intByPassArea
                        If CStr(strB1029) = Format(strByPassEndTime(j), "hhmm") Then
                           strUpdB1005 = Format(strByPassStarTime(j), "hhmm")
                           Exit For
                        End If
                     Next j
                     '2021/8/13 END
                     
                     'Add By Sindy 2012/4/11
                     strUpdB1028 = strUpdB1005
                     strUpdB1029 = strB1029
                     '2012/4/11 End
                  End If
               Else
                  If DBDATE(strUpdB1004) = strB1004 Then
'                     cboSTime.Visible = True
'                     cboETime.Visible = True
                     '起日上班時段
'                     For i = 0 To cboSTime.ListCount - 1
'                        If cboSTime.List(i) = Format(strB1028, "##:##") Then
'                           cboSTime.ListIndex = i
'                           Exit For
'                        End If
'                     Next i
                     strUpdB1028 = strB1028
'                     cboETime.ListIndex = 0 '預設值
                     'strUpdB1029 = Format(strByPassEndTime(1), "hhmm") '"1700" 'Modify By Sindy 2024/12/17 mark
                     strUpdB1029 = strB1029 'Modify By Sindy 2024/12/17
                     
                  ElseIf DBDATE(strUpdB1006) = strB1006 Then
'                     cboSTime.Visible = True
'                     cboETime.Visible = True
'                     cboSTime.ListIndex = 0 '預設值
                      
                     'strUpdB1028 = Format(strByPassStarTime(1), "hhmm") '"800" 'Modify By Sindy 2024/12/17 mark
                     '迄日下班時段
'                     For i = 0 To cboETime.ListCount - 1
'                        If cboETime.List(i) = Format(strB1029, "##:##") Then
'                           cboETime.ListIndex = i
'                           Exit For
'                        End If
'                     Next i
                     strUpdB1028 = strB1028 'Modify By Sindy 2024/12/17
                     strUpdB1029 = strB1029
                  End If
               End If
            End If
            '計算日,時
            'Add By Sindy 2012/4/12
            If strB1002 = 表單類別_出差 Then
               Call PUB_CountHour_Busi_Trip(strUpdB1004, strUpdB1005, strUpdB1006, strUpdB1007, m_Day, m_Hour)
               strUpdB1009 = m_Day
               strUpdB1010 = m_Hour
            Else
            '2012/4/12 End
               'Modify by Sindy 2012/10/12
               'Call PUB_CountDayHour(strB1003, strUpdB1004, strUpdB1005, strUpdB1006, strUpdB1007, strUpdB1028, strUpdB1029, strUpdB1009, strUpdB1010, bolMsg)
               If strUpdB1007 = "" Then strUpdB1007 = strUpdB1029 'Add By Sindy 2024/12/17
               Call PUB_CountDayHour(strB1003, strUpdB1004, strUpdB1005, strUpdB1006, strUpdB1007, strUpdB1028, strUpdB1029, strUpdB1009, strUpdB1010, strB1008, bolMsg)
            End If
            If Val(strUpdB1009) = 0 And Val(strUpdB1010) = 0 Then
               If bolMsg = True Then MsgBox "系統自動執行人事處簽收時，計算日/時有誤！", vbExclamation
               PUB_AutoM21Receive = False
               Exit Function
            End If
            '存檔
            strChkYm = strChkYm & IIf(strChkYm <> "", ",", "") & Left(strUpdB1004, 6) 'Add By Sindy 2015/12/25
            If PUB_SavePMainFile(strB1001, strB1002, strB1003, strUpdB1004, strUpdB1005, strUpdB1006, strUpdB1007, strB1008, strUpdB1009, strUpdB1010, strB1030, strB101213, strB1014, strB1015, strUpdB1028, strUpdB1029, bolMsg) = False Then
               PUB_AutoM21Receive = False
               Exit Function
            End If
'            '串多筆假單訊息
'            If strMsg = "" Then
'               strMsg = "請假單跨月份，系統自動拆多筆，內容如下：" & vbCrLf
'            Else
'               strMsg = strMsg & vbCrLf
'            End If
'            strNowData = "" '預設值
''            If cboSTime.Visible = True And cboETime.Visible = True And _
''               cboSTime.Text <> "" And cboETime.Text <> "" Then
''               strNowData = strNowData & "非整日," & cboSTime.Text & "," & cboETime.Text
''            End If
'            strNowData = strNowData & "," & ChangeWStringToTDateString(DBDATE(strUpdB1004)) & "," & Format(txtB1005_1 & Format("00" & txtB1005_2, "00"), "##:##")
'            strNowData = strNowData & "," & ChangeWStringToTDateString(DBDATE(txtB1006)) & "," & Format(txtB1007_1 & Format("00" & txtB1007_2, "00"), "##:##")
'            strNowData = strNowData & "," & txtB1009 & "日," & txtB1010 & "時"
'            If Left(strNowData, 1) = "," Then strNowData = Right(strNowData, Len(strNowData) - 1)
'            strMsg = strMsg & strNowData
            '下一個年月
            If Val(Right(dbl_cYM, 2)) = 12 Then
               dbl_cYM = (Val(Left(dbl_cYM, 4)) + 1) * 100 + 1
            Else
               dbl_cYM = Val(dbl_cYM) + 1
            End If
         Loop
'         '顯示拆單訊息
'         MsgBox strMsg, vbExclamation
      Else
         '存檔
         strUpdB1004 = strB1004
         strUpdB1005 = strB1005
         strUpdB1006 = strB1006
         strUpdB1007 = strB1007
         strUpdB1009 = strB1009
         strUpdB1010 = strB1010
         'Add By Sindy 2012/4/11 人事資料增加非整日起迄上下班時段
         strUpdB1028 = strB1028
         strUpdB1029 = strB1029
         '2012/4/11 End
         strChkYm = Left(strUpdB1004, 6) 'Add By Sindy 2015/12/25
         If PUB_SavePMainFile(strB1001, strB1002, strB1003, strUpdB1004, strUpdB1005, strUpdB1006, strUpdB1007, strB1008, strUpdB1009, strUpdB1010, strB1030, strB101213, strB1014, strB1015, strUpdB1028, strUpdB1029, bolMsg) = False Then
            PUB_AutoM21Receive = False
            Exit Function
         End If
      End If
   End If
   
   cnnConnection.CommitTrans
   
   'Modify By Sindy 2019/5/23 表單簽核完成,後續資料檢查及SendMail
   Call PUB_AutoM21Receive_SendMail(strB1001, strB1002, strB1003, strUpdB1004, strUpdB1005, _
      strUpdB1006, strB1014, strChkYm, bolAutoBatch, , , , strB101213)
   
'   'Modify By Sindy 2013/9/11 系統自動確認打卡異常資料
'   Call PUB_UpdateB14Data(strB1003)
'   'Modify By Sindy 2018/11/15 假單核准時,檢查在此請假區間中是否有幫他人做職代,若有,發Mail通知其他職代
'   If strB1002 = 表單類別_請假 Or strB1002 = 表單類別_出差 Then
'      Call CheckIsPersonRestSectorMail(strB1001)
'   'Add by Sindy 2019/4/30 加班要檢查是否超過46小時
'   Else
'      strMsg = PUB_PerFormRemindMsg(表單類別_加班, "1", strB1003, DBDATE(strUpdB1004) - 19110000, 0, False)
'      If strMsg <> "" Then
'         PUB_SendMail strUserNum, "68010", "", "【加班累計超過40小時】" & strSubject, strContent & vbCrLf & vbCrLf & strMsg, , , , , , , , , , True
'      End If
'   End If
'
''   'Add By Sindy 2013/7/12
''   '當假單核准時,檢查是否有該時間區間的打卡異常資料尚未核銷,若有,則一併核銷異常資料
''   If strB1002 = 表單類別_請假 Or strB1002 = 表單類別_出差 Then
''      '已請假的區間去檢查是否有異常資料
''      strSql = "select b1401,b1402,b1403,b1404,V4.min_pr02,decode(b1404,null,V4.min_pr02,b1404)" & _
''               " from abs014," & _
''               "(select scd01,pr01,nvl(min(pr02),0) as min_pr02,nvl(max(pr02),0) as max_pr02 from pollrecord,staffcarddata where pr03=scd02(+) and pr01>=" & strB1004 & " and pr01<=" & strB1006 & " and scd01='" & strB1003 & "' group by scd01,pr01) V4" & _
''               " where b1401='" & strB1003 & "'" & _
''               " and b1401=V4.scd01 and b1402=V4.pr01" & _
''               " and b1402||substr('0000'||substr(decode(b1404,null,V4.min_pr02,b1404),1,length(decode(b1404,null,V4.min_pr02,b1404))-2),-4) between " & strB1004 & strB1005 & " and " & strB1006 & strB1007 & _
''               " and b1411 is null"
''      intI = 1
''      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
''      If intI = 1 Then
''         If rsTmp.RecordCount > 0 Then
''            rsTmp.MoveFirst
''            Do While Not rsTmp.EOF
''               'Modify By Sindy 2013/8/22
'''               strSql = "update ABS014 set B1405='" & IIf(strB1002 = 表單類別_出差, "6", "1") & "'" & _
'''                                         ",B1411='A'" & _
'''                                         ",B1412=" & strUpdDate & _
'''                                         ",B1413=" & strUpdTime & _
'''                        " where b1401='" & strB1003 & "'" & _
'''                        " and b1402||substr('0000'||substr(b1404,1,length(b1404)-2),-4) between " & strB1004 & strB1005 & " and " & strB1006 & strB1007 & _
'''                        " and b1411 is null"
''               strSql = "update ABS014 set B1405='" & IIf(strB1002 = 表單類別_出差, "6", "1") & "'" & _
''                                         ",B1411='A'" & _
''                                         ",B1412=" & strUpdDate & _
''                                         ",B1413=" & strUpdTime & _
''                        " where b1401='" & rsTmp.Fields("b1401") & "'" & _
''                          " and b1402=" & rsTmp.Fields("b1402") & _
''                          " and b1403='" & rsTmp.Fields("b1403") & "'" & _
''                          " and b1411 is null"
''               '2013/8/22 END
''               cnnConnection.Execute strSql
''               rsTmp.MoveNext
''            Loop
''         End If
''      End If
''      '檢查請假當事者是否有那一天上/下班打卡都出異常並且該天有請整日的假單，則將該日異常資料核銷
''      strSql = "UPDATE abs014" & _
''               " set B1405='" & IIf(strB1002 = 表單類別_出差, "6", "1") & "',B1411='A',B1412=" & strUpdDate & ",B1413=" & strUpdTime & _
''               " where b1401||b1402 in(" & _
''               " select b14.b1401||b14.b1402 from staff_absence,(" & _
''               " select b1401,b1402,count(*) as ErrCnt from abs014 where b1401='" & strB1003 & "' and b1411 is null group by b1401,b1402) b14" & _
''               " Where sa01 = b14.b1401" & _
''               " and b14.b1402 between sa02 and sa04" & _
''               " and sa08=0" & _
''               " Union" & _
''               " select b14.b1401||b14.b1402 from staff_busi_trip,(" & _
''               " select b1401,b1402,count(*) as ErrCnt from abs014 where b1401='" & strB1003 & "' and b1411 is null group by b1401,b1402) b14" & _
''               " Where sb01 = b14.b1401" & _
''               " and b14.b1402 between sb02 and sb04" & _
''               " and sb07=0" & _
''               ")"
''      cnnConnection.Execute strSql
''      '中午休息時間下班者(打卡時間在121001~132959)
''      If strB1005 = "1330" Then
''         strSql = "update ABS014 set B1405='" & IIf(strB1002 = 表單類別_出差, "6", "1") & "'" & _
''                                   ",B1411='A'" & _
''                                   ",B1412=" & strUpdDate & _
''                                   ",B1413=" & strUpdTime & _
''                  " where b1401='" & strB1003 & "'" & _
''                  " and b1402=" & strB1004 & _
''                  " and b1404 between 121001 and 132959" & _
''                  " and b1411 is null"
''         cnnConnection.Execute strSql
''      End If
''      '中午休息時間上班者(打卡時間在121001~132959)
''      If strB1007 = "1210" Then
''         strSql = "update ABS014 set B1405='" & IIf(strB1002 = 表單類別_出差, "6", "1") & "'" & _
''                                   ",B1411='A'" & _
''                                   ",B1412=" & strUpdDate & _
''                                   ",B1413=" & strUpdTime & _
''                  " where b1401='" & strB1003 & "'" & _
''                  " and b1402=" & strB1006 & _
''                  " and b1404 between 121001 and 132959" & _
''                  " and b1411 is null"
''         cnnConnection.Execute strSql
''      End If
''   End If
''   '2013/7/12 END
'
'   'If bolAutoBatch = False Then
'      '發E-Mail通知當事人
'      'Modify By Sindy 2015/7/29
'      If bolAutoBatch = True Then
'         PUB_SendMail strUserNum, strB1003, "", strSubject, strContent, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , False, , , , "QPGMR", "系統管理員", , True, False
'      Else
'      '2015/7/29 END
'         PUB_SendMail strUserNum, strB1003, "", strSubject, strContent, , , , , , , , , , True
'      End If
'
'      'Add By Sindy 2015/12/25 加班單或請假單主管核准時,若當月的薪資已計算過,發E-MAIL通知財務處
'      If strB1002 = 表單類別_請假 Or strB1002 = 表單類別_加班 Then
'         varTemp = Split(strChkYm, ",")
'         For i = 0 To UBound(varTemp)
'            strSql = "select sm02,count(*) from SalaryMonth where sm02=" & varTemp(i) & " group by sm02"
'            intI = 1
'            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               If bolAutoBatch = True Then
'                  PUB_SendMail strUserNum, "71005", "", strB1003 & GetPrjSalesNM(strB1003) & "補輸" & IIf(strB1002 = 表單類別_請假, "請假", "加班") & "資料，請重新做每月(" & varTemp(0) & ")薪資計算！", strContent, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , False, , , , "QPGMR", "系統管理員", , True, False
'               Else
'                  PUB_SendMail strUserNum, "71005", "", strB1003 & GetPrjSalesNM(strB1003) & "補輸" & IIf(strB1002 = 表單類別_請假, "請假", "加班") & "資料，請重新做每月(" & varTemp(0) & ")薪資計算！", strContent, , , , , , , , , , True
'               End If
'            End If
'         Next i
'      End If
'      '2015/12/25 END
'
''      'Add By Sindy 2013/12/20 增列國外部投資法律處(F31),在請假或出差核准通知請假人時,同時也MAIL通知桂所長76012
''      If GetStaffDepartment(strB1003) = "F31" And (strB1002 = 表單類別_請假 Or strB1002 = 表單類別_出差) Then
''         PUB_SendMail strUserNum, "76012", "", strSubject, strContent, , , , , , , , , , True
''      'Modify By Sindy 2015/5/21 因 F31 人員將改編制 L02, 但有些人的假單會經過桂所長簽核, 有些人不必,
''      '                          因此若為 L02部門的人, 只要假單或出差單未經桂所長簽核, 在核准時都要發MAIL通知桂所長
''      ElseIf GetStaffDepartment(strB1003) = "L02" And (strB1002 = 表單類別_請假 Or strB1002 = 表單類別_出差) And _
''             InStr(GetBossB1107_2_1(strB1001), "76012") = 0 Then
''         PUB_SendMail strUserNum, "76012", "", strSubject, strContent, , , , , , , , , , True
''      'Add By Sindy 2015/7/2 日文顧問,請假/出差假單核准時,一併通知王文安經理
''      ElseIf GetStaffDepartment(strB1003) = "F71" And (strB1002 = 表單類別_請假 Or strB1002 = 表單類別_出差) Then
''         PUB_SendMail strUserNum, "88003", "", strSubject, strContent, , , , , , , , , , True
''      '2015/7/2 END
''      End If
''      '2013/12/20 END
''      'Add By Sindy 2011/11/11 74018.杜燕文請假及出差核准後,知會68006.杜副總
''      If strB1003 = "74018" And _
''         (strB1002 = 表單類別_請假 Or strB1002 = 表單類別_出差) Then
''         PUB_SendMail strUserNum, "68006", "", strSubject, strContent, , , , , , , , , , True
''      End If
'      'Modify By Sindy 2015/7/15 檢查是否有需要核准後再通知其他人員
'      strSql = "select b0124,b0125 from abs001 where b0101='" & strB1003 & "' and b0124 is not null"
'      intI = 1
'      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         strTo = rsTmp.Fields("b0124")
'         If "" & rsTmp.Fields("b0125") = "N" Then '不含簽核主管
'            varTemp = Split(strTo, ";")
'            strTo = ""
'            For i = 0 To UBound(varTemp)
'               If InStr(GetBossB1107_2_1(strB1001), varTemp(i)) = 0 Then
'                  strTo = strTo & ";" & varTemp(i)
'               End If
'            Next i
'            If Left(strTo, 1) = ";" Then strTo = Mid(strTo, 2)
'         End If
'         If strTo <> "" Then
'            'Add By Sindy 2019/4/25 劉經理:董事長指示89037蘇月星請假、出差及加班單在所長核准後應會知董事長
'            If (strB1002 = 表單類別_請假 Or strB1002 = 表單類別_出差) Or strB1003 = "89037" Then
'            '2019/4/25 END
'               'Modify By Sindy 2015/7/29
'               If bolAutoBatch = True Then
'                  PUB_SendMail strUserNum, strTo, "", strSubject, strContent, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , False, , , , "QPGMR", "系統管理員", , True, False, , , , , , , , True
'               Else
'               '2015/7/29 END
'                  PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True, , , , , , , , , True
'               End If
'            End If
'         End If
'      End If
'      '2015/7/15 END
'
'      'Add By Sindy 2012/2/10 專利處P10-P14人員當日臨時請假核准後,必須E-Mail通知71011王副總
'      'Modify By Sindy 2012/3/2 請假起迄日小於等於系統日請假核准者,均要通知王副總
'      If strB1003 <> "71011" And InStr(GetBossB1107_2_1(strB1001), "71011") = 0 And _
'         (strB1002 = 表單類別_請假 Or strB1002 = 表單類別_出差) And _
'         (GetStaffDepartment(strB1003) >= "P10" And GetStaffDepartment(strB1003) <= "P14") And _
'         (strUpdB1004 <= strSrvDate(1) And strUpdB1006 <= strSrvDate(1)) Then
'         'Modify By Sindy 2015/7/29
'         If bolAutoBatch = True Then
'            PUB_SendMail strUserNum, "71011", "", strSubject, strContent, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , False, , , , "QPGMR", "系統管理員", , True, False
'         Else
'         '2015/7/29 END
'            PUB_SendMail strUserNum, "71011", "", strSubject, strContent, , , , , , , , , , True
'         End If
'      End If
'      '2012/2/10 End
'      '國外大陸出差的假單在人事處簽收時，一併要發Mail通知劉經理（增加系統特殊設定：國外大陸出差通知人員）
'      If strB1002 = 表單類別_出差 And (strB1014 = "3" Or strB1014 = "4") Then
'         '取得國外大陸出差通知人員
'         strTo = Pub_GetSpecMan("國外大陸出差通知人員")
'         'Add By Sindy 2014/9/24 桂所長及閻副所長國外出差時(含大陸出差),在核准後知會人事時,一併通知林總經理
'         If strB1003 = "76012" Or strB1003 = "81040" Then
'            If InStr(GetBossB1107_2_1(strB1001), "94007") = 0 Then
'               strTo = strTo & ";94007"
'            End If
'         End If
'         '2014/9/24 END
'         If strTo <> "" Then
'            strContent = GetEMailContent(strB1001, strSubject, , , True)
'            strSubject = strContent & IIf(strB1014 = "3", " 大陸", " 國外") & "出差"
'            'Modify By Sindy 2015/7/29
'            If bolAutoBatch = True Then
'               PUB_SendMail strUserNum, strTo, "", strSubject, strContent, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , False, , , , "QPGMR", "系統管理員", , True, False
'            Else
'            '2015/7/29 END
'               PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True
'            End If
'         End If
'      End If
'   'End If
'   Set rsTmp = Nothing
   
   Exit Function
   
ErrHand:
   PUB_AutoM21Receive = False
   cnnConnection.RollbackTrans
   If bolMsg = True Then MsgBox "系統自動簽收表單進人事系統失敗！" & vbCrLf & Err.Description
End Function

'Add By Sindy 2019/5/23 表單簽核完成,後續資料檢查及SendMail
'm_EditMode=1.新增資料
'           2.修改資料
'           3.刪除資料
Public Sub PUB_AutoM21Receive_SendMail(m_strB1001 As String, m_strB1002 As String, m_strB1003 As String, _
   strB1004 As String, strB1005 As String, strB1006 As String, strB1014 As String, _
   strChkYm As String, Optional bolAutoBatch As Boolean = False, _
   Optional m_strSubject As String = "", Optional m_strContent As String = "", _
   Optional m_EditMode As Integer = 0, Optional m_strB101213 As String)
   
Dim rsTmp As New ADODB.Recordset
Dim strMsg As String
Dim varTemp As Variant
Dim strTo As String, i As Integer
Dim strSubject As String, strContent As String
Dim strTemp As String
   
   'Modify By Sindy 2013/9/11 系統自動確認打卡異常資料
   If m_EditMode <> 2 And m_EditMode <> 3 Then
      Call PUB_UpdateB14Data(m_strB1003)
   End If
   
   If m_strSubject = "" Then
      '先取得E-Mail主旨,本文內容
      strContent = GetEMailContent(m_strB1001, strSubject, , , , m_strB1002, m_strB1003, strB1004, Val(strB1005), m_EditMode)
      If m_strContent <> "" Then strContent = strContent & vbCrLf & vbCrLf & m_strContent  'Add By Sindy 2024/10/18
   Else
      strSubject = m_strSubject
      strContent = m_strContent
   End If
   
   'Modify By Sindy 2018/11/15 假單核准時,檢查在此請假區間中是否有幫他人做職代,若有,發Mail通知其他職代
   If m_strB1002 = 表單類別_請假 Or m_strB1002 = 表單類別_出差 Then
      If m_EditMode <> 2 And m_EditMode <> 3 Then
         Call CheckIsPersonRestSectorMail(m_strB1001)
      End If
   'Add by Sindy 2019/4/30 加班要檢查是否超過46小時
   Else
      strMsg = PUB_PerFormRemindMsg(表單類別_加班, "1", m_strB1003, DBDATE(strB1004) - 19110000, m_strB101213, False)
      If strMsg <> "" Then
         'Modify By Sindy 2024/3/4 改使用 Pub_GetSpecMan("國外大陸出差通知人員")
         PUB_SendMail "QPGMR", Pub_GetSpecMan("國外大陸出差通知人員"), "", "【加班累計超過40小時】" & strSubject, strContent & vbCrLf & vbCrLf & strMsg, , , , , , , , , , True
      End If
   End If
   
   'If bolAutoBatch = False Then
      'Add By Sindy 2015/12/25 加班單或請假單主管核准時,若當月的薪資已計算過,發E-MAIL通知財務處
      If m_strB1002 = 表單類別_請假 Or m_strB1002 = 表單類別_加班 Then
         varTemp = Split(strChkYm, ",")
         For i = 0 To UBound(varTemp)
            strSql = "select sm02,count(*) from SalaryMonth where sm02=" & varTemp(i) & " group by sm02"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If bolAutoBatch = True Then
                  'Modify By Sindy 2024/3/4 改使用 Pub_GetSpecMan("試用期滿追蹤薪資人員")
                  PUB_SendMail strUserNum, Pub_GetSpecMan("試用期滿追蹤薪資人員"), "", m_strB1003 & GetPrjSalesNM(m_strB1003) & IIf(m_EditMode = 3, "刪除", IIf(m_EditMode = 2, "修改", "補輸")) & IIf(m_strB1002 = 表單類別_請假, "請假", "加班") & "資料，請重新做每月(" & varTemp(0) & ")薪資計算！", strContent, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , False, , , , "QPGMR", "系統管理員", , True, False
               Else
                  'Modify By Sindy 2024/3/4 改使用 Pub_GetSpecMan("試用期滿追蹤薪資人員")
                  PUB_SendMail strUserNum, Pub_GetSpecMan("試用期滿追蹤薪資人員"), "", m_strB1003 & GetPrjSalesNM(m_strB1003) & IIf(m_EditMode = 3, "刪除", IIf(m_EditMode = 2, "修改", "補輸")) & IIf(m_strB1002 = 表單類別_請假, "請假", "加班") & "資料，請重新做每月(" & varTemp(0) & ")薪資計算！", strContent, , , , , , , , , , True
               End If
            End If
         Next i
      End If
      '2015/12/25 END
      
      '發E-Mail通知當事人
      'Modify By Sindy 2015/7/29
      'Modify By Sindy 2019/5/23 剔除員工為”不寄信”者
      If ChkStaffST14(m_strB1003, False) = False Then
      '2019/5/23 END
         If m_EditMode <> 2 And m_EditMode <> 3 Then
            If bolAutoBatch = True Then
               PUB_SendMail strUserNum, m_strB1003, "", strSubject, strContent, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , False, , , , "QPGMR", "系統管理員", , True, False
            Else
            '2015/7/29 END
               PUB_SendMail strUserNum, m_strB1003, "", strSubject, strContent, , , , , , , , , , True
            End If
         End If
      End If
      
      '******************************************
      ' 加發通知各收受者
      '******************************************
'      'Add By Sindy 2013/12/20 增列國外部投資法律處(F31),在請假或出差核准通知請假人時,同時也MAIL通知桂所長76012
'      If GetStaffDepartment(m_strB1003) = "F31" And (m_strB1002 = 表單類別_請假 Or m_strB1002 = 表單類別_出差) Then
'         PUB_SendMail strUserNum, "76012", "", strSubject, strContent, , , , , , , , , , True
'      'Modify By Sindy 2015/5/21 因 F31 人員將改編制 L02, 但有些人的假單會經過桂所長簽核, 有些人不必,
'      '                          因此若為 L02部門的人, 只要假單或出差單未經桂所長簽核, 在核准時都要發MAIL通知桂所長
'      ElseIf GetStaffDepartment(m_strB1003) = "L02" And (m_strB1002 = 表單類別_請假 Or m_strB1002 = 表單類別_出差) And _
'             InStr(GetBossB1107_2_1(m_strB1001), "76012") = 0 Then
'         PUB_SendMail strUserNum, "76012", "", strSubject, strContent, , , , , , , , , , True
'      'Add By Sindy 2015/7/2 日文顧問,請假/出差假單核准時,一併通知王文安經理
'      ElseIf GetStaffDepartment(m_strB1003) = "F71" And (m_strB1002 = 表單類別_請假 Or m_strB1002 = 表單類別_出差) Then
'         PUB_SendMail strUserNum, "88003", "", strSubject, strContent, , , , , , , , , , True
'      '2015/7/2 END
'      End If
'      '2013/12/20 END
'      'Add By Sindy 2011/11/11 74018.杜燕文請假及出差核准後,知會68006.杜副總
'      If m_strB1003 = "74018" And _
'         (m_strB1002 = 表單類別_請假 Or m_strB1002 = 表單類別_出差) Then
'         PUB_SendMail strUserNum, "68006", "", strSubject, strContent, , , , , , , , , , True
'      End If
      'Modify By Sindy 2015/7/15 檢查是否有需要核准後再通知其他人員
      strTo = ""
      'Add By Sindy 2019/4/25 劉經理:董事長指示89037蘇月星請假、出差及加班單在所長核准後應會知董事長
      If (m_strB1002 = 表單類別_請假 Or m_strB1002 = 表單類別_出差) Or m_strB1003 = "89037" Then
         strSql = "select b0124,b0125 from abs001 where b0101='" & m_strB1003 & "' and b0124 is not null"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strTo = rsTmp.Fields("b0124")
            If "" & rsTmp.Fields("b0125") = "N" Then  '不含簽核主管
               varTemp = Split(strTo, ";")
               strTo = ""
               For i = 0 To UBound(varTemp)
                  If InStr(GetBossB1107_2_1(m_strB1001), varTemp(i)) = 0 Then
                     strTo = strTo & ";" & varTemp(i)
                  End If
               Next i
            End If
         End If
      End If
      '2015/7/15 END
      'Add By Sindy 2012/2/10 專利處P10-P14人員當日臨時請假核准後,必須E-Mail通知71011王副總
      'Modify By Sindy 2012/3/2 請假起迄日小於等於系統日請假核准者,均要通知王副總
      If m_strB1003 <> "71011" And _
         (m_strB1002 = 表單類別_請假 Or m_strB1002 = 表單類別_出差) And _
         (GetStaffDepartment(m_strB1003) >= "P10" And GetStaffDepartment(m_strB1003) <= "P14") And _
         (strB1004 <= strSrvDate(1) And strB1006 <= strSrvDate(1)) Then
         
         'Modify By Sindy 2023/4/24 5/1 加入游、李，5/11取消王副總
         If Val(strSrvDate(1)) >= 20230501 Then
            If Val(strSrvDate(1)) < 20230511 Then
               If InStr(GetBossB1107_2_1(m_strB1001), "71011") = 0 Then
                  If InStr(strTo, "71011") = 0 Then
                     strTo = strTo & ";" & "71011"
                  End If
               End If
            End If
            
            'Modified by Morgan 2025/2/21
            'If m_strB1003 <> "73022" Then
            '   If InStr(GetBossB1107_2_1(m_strB1001), "73022") = 0 Then
            '      If InStr(strTo, "73022") = 0 Then
            '         strTo = strTo & ";" & "73022"
            '      End If
            '   End If
            'End If
            pub_PMan = Pub_GetSpecMan("專利處特定編號")
            If m_strB1003 <> Right(pub_PMan, 5) Then
               If InStr(GetBossB1107_2_1(m_strB1001), Right(pub_PMan, 5)) = 0 Then
                  If InStr(strTo, Right(pub_PMan, 5)) = 0 Then
                     strTo = strTo & ";" & Right(pub_PMan, 5)
                  End If
               End If
            End If
            If m_strB1003 <> Left(pub_PMan, 5) Then
               If InStr(GetBossB1107_2_1(m_strB1001), Left(pub_PMan, 5)) = 0 Then
                  If InStr(strTo, Left(pub_PMan, 5)) = 0 Then
                     strTo = strTo & ";" & Left(pub_PMan, 5)
                  End If
               End If
            End If
            'end 2025/2/21
            
            If m_strB1003 <> "99050" Then
               If InStr(GetBossB1107_2_1(m_strB1001), "99050") = 0 Then
                  If InStr(strTo, "99050") = 0 Then
                     strTo = strTo & ";" & "99050"
                  End If
               End If
            End If
         Else
         '2023/4/24 END
            If InStr(GetBossB1107_2_1(m_strB1001), "71011") = 0 Then
               If InStr(strTo, "71011") = 0 Then
                  strTo = strTo & ";" & "71011"
               End If
            End If
         End If
'         'Modify By Sindy 2015/7/29
'         If bolAutoBatch = True Then
'            PUB_SendMail strUserNum, "71011", "", strSubject, strContent, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , False, , , , "QPGMR", "系統管理員", , True, False
'         Else
'         '2015/7/29 END
'            PUB_SendMail strUserNum, "71011", "", strSubject, strContent, , , , , , , , , , True
'         End If
      End If
      '2012/2/10 End
      'Add By Sindy 2019/5/24 人事處修改/刪除假單時,也要一併通知相關假單人員
      If m_EditMode = 2 Or m_EditMode = 3 Then
         '寄E-Mail通知當事人有異動內容
         'Modify By Sindy 2012/7/17 發E-Mail通知當事人之外，已簽核的職代及審核主管亦也要通知
         strTemp = GetBossB1107_All(m_strB1001)
         If strTemp <> "" Then
            If strTo = "" Then
               strTo = strTemp
            Else
               varTemp = Split(strTemp, ";")
               For i = 0 To UBound(varTemp)
                  If InStr(strTo, varTemp(i)) = 0 Then
                     strTo = strTo & ";" & varTemp(i)
                  End If
               Next i
            End If
         End If
         'Add By Sindy 2012/7/17 專利處P10-P14,必須另外E-Mail通知71011王副總
         'Modify By Sindy 2023/4/24 5/1 加入游、李，5/11取消王副總
         If Val(strSrvDate(1)) >= 20230501 And _
            (GetStaffDepartment(m_strB1003) >= "P10" And GetStaffDepartment(m_strB1003) <= "P14") Then
            If Val(strSrvDate(1)) < 20230511 Then
               If InStr(strTo, "71011") = 0 Then
                  strTo = strTo + ";71011"
               End If
            End If
            
            'Modified by Moran 2025/2/21
            'If InStr(strTo, "73022") = 0 Then
            '   strTo = strTo + ";73022"
            'End If
            pub_PMan = Pub_GetSpecMan("專利處特定編號")
            If InStr(strTo, Right(pub_PMan, 5)) = 0 Then
               strTo = strTo + ";" & Right(pub_PMan, 5)
            End If
            If InStr(strTo, Left(pub_PMan, 5)) = 0 Then
               strTo = strTo + ";" & Left(pub_PMan, 5)
            End If
            'end 2025/2/21
            
            If InStr(strTo, "99050") = 0 Then
               strTo = strTo + ";99050"
            End If
         Else
         '2023/4/24 END
            If (GetStaffDepartment(m_strB1003) >= "P10" And GetStaffDepartment(m_strB1003) <= "P14") And _
               InStr(strTo, "71011") = 0 Then
               strTo = strTo + ";71011"
            End If
         End If
      End If
      If strTo <> "" And Left(strTo, 1) = ";" Then strTo = Mid(strTo, 2)
      'Add By Sindy 2019/5/24 人事處修改/刪除假單時,也要一併通知相關假單人員
      If m_EditMode = 2 Or m_EditMode = 3 Then
         If ChkStaffST14(m_strB1003, False) = False Then
            PUB_SendMail strUserNum, m_strB1003, "", strSubject, strContent, , , , , , strTo, , , , True, , , , , , , , , IIf((m_strB1002 = 表單類別_請假 Or m_strB1002 = 表單類別_出差) Or m_strB1003 = "89037", True, False)
         ElseIf strTo <> "" Then
            PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True, , , , , , , , , IIf((m_strB1002 = 表單類別_請假 Or m_strB1002 = 表單類別_出差) Or m_strB1003 = "89037", True, False)
         End If
      '核准/新增假單時
      Else
         If strTo <> "" Then
            'Modify By Sindy 2015/7/29
            If bolAutoBatch = True Then
               PUB_SendMail strUserNum, strTo, "", strSubject, strContent, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , False, , , , "QPGMR", "系統管理員", , True, False, , , , , , , , True
            Else
            '2015/7/29 END
               PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True, , , , , , , , , True
            End If
         End If
      End If
      
      '國外大陸出差的假單在人事處簽收時，一併要發Mail通知劉經理（增加系統特殊設定：國外大陸出差通知人員）
      If m_strB1002 = 表單類別_出差 And (strB1014 = "3" Or strB1014 = "4") Then
         '取得國外大陸出差通知人員
         strTo = Pub_GetSpecMan("國外大陸出差通知人員")
         'Add By Sindy 2014/9/24 桂所長及閻副所長國外出差時(含大陸出差),在核准後知會人事時,一併通知林總經理
         If m_strB1003 = "76012" Or m_strB1003 = "81040" Then
            If InStr(GetBossB1107_2_1(m_strB1001), "94007") = 0 Then
               strTo = strTo & ";94007"
            End If
         End If
         '2014/9/24 END
         If strTo <> "" Then
            If m_strSubject <> "" Then
               strSubject = m_strSubject & IIf(strB1014 = "3", " (大陸", " (國外") & "出差)"
               strContent = m_strContent
            Else
               strContent = GetEMailContent(m_strB1001, strSubject, , , True, m_strB1002, m_strB1003, strB1004, Val(strB1005), m_EditMode)
               strSubject = strContent & IIf(strB1014 = "3", " (大陸", " (國外") & "出差)"
            End If
            'Modify By Sindy 2015/7/29
            If bolAutoBatch = True Then
               PUB_SendMail strUserNum, strTo, "", strSubject, strContent, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", , False, , , , "QPGMR", "系統管理員", , True, False
            Else
            '2015/7/29 END
               PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True
            End If
         End If
      End If
   'End If
   Set rsTmp = Nothing
End Sub

'更新人事系統請假主檔
'Add By Sindy 2012/4/11 人事資料增加非整日起迄上下班時段 +strUpdB1028 As String, strUpdB1029 As String
Public Function PUB_SavePMainFile(strB1001 As String, strB1002 As String, strB1003 As String _
, strUpdB1004 As String, strUpdB1005 As String, strUpdB1006 As String, strUpdB1007 As String _
, strB1008 As String, strUpdB1009 As String, strUpdB1010 As String _
, strB1030 As String, strB101213 As String, strB1014 As String, strB1015 As String _
, strUpdB1028 As String, strUpdB1029 As String _
, Optional bolMsg As Boolean = True) As Boolean

Dim strErrTxt As String
Dim rsTmp As New ADODB.Recordset
Dim strB1012 As String, strB1013 As String 'Add by Sindy 2016/12/26
   
On Error GoTo ErrHand
   
   PUB_SavePMainFile = True
   
   '新增人事系統該筆表單資料
   If strB1002 = 表單類別_請假 Then
      '檢查是否有資料重覆
      strSql = "SELECT * FROM Staff_Absence WHERE SA01='" & strB1003 & _
                  "' and SA02=" & strUpdB1004 & " and SA03=" & strUpdB1005
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strErrTxt = "人事員工請假資料重覆！"
         GoTo ErrHand
      End If
      
      'Modify By Sindy 2012/4/11 人事資料增加非整日起迄上下班時段
      strSql = "insert into Staff_Absence(SA01,SA02,SA03,SA04,SA05,SA06,SA07,SA08,SA09,SA16,SA17) " & _
               "values(" & CNULL(strB1003) & "," & strUpdB1004 & "," & strUpdB1005 & _
               "," & strUpdB1006 & "," & strUpdB1007 & "," & CNULL(strB1008) & _
               "," & CNULL(strUpdB1009) & "," & CNULL(strUpdB1010) & "," & CNULL(strB1001) & _
               "," & CNULL(strUpdB1028) & "," & CNULL(strUpdB1029) & ")"
   ElseIf strB1002 = 表單類別_加班 Then
      '檢查是否有資料重覆
      strSql = "SELECT * FROM Staff_Overtime WHERE So01='" & strB1003 & _
                  "' and So02=" & strUpdB1004 & " and So03=" & strUpdB1005
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strErrTxt = "人事員工加班資料重覆！"
         GoTo ErrHand
      End If
      'Add by Sindy 2016/12/26
      If ChkWorkDay(strUpdB1004, strB1003, True) = False Then '假日
         strB1012 = ""
         strB1013 = strB101213
      Else '平日
         strB1012 = strB101213
         strB1013 = ""
      End If
      '2016/12/26 END
      'Modify By Sindy 2016/12/26 +,SO15
      strSql = "insert into Staff_Overtime(SO01,SO02,SO03,SO04,SO05,SO06,SO13,SO15) " & _
               "values(" & CNULL(strB1003) & "," & strUpdB1004 & "," & strUpdB1005 & _
               "," & strUpdB1007 & "," & CNULL(strB1012) & "," & CNULL(strB1013) & _
               "," & CNULL(strB1001) & "," & CNULL(strB1030) & ")"
   ElseIf strB1002 = 表單類別_出差 Then
      '檢查是否有資料重覆
      strSql = "SELECT * FROM Staff_Busi_Trip WHERE SB01='" & strB1003 & _
                  "' and SB02=" & strUpdB1004 & " and SB03=" & strUpdB1005
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strErrTxt = "人事員工出差資料重覆！"
         GoTo ErrHand
      End If
      
      'Modify By Sindy 2012/4/11 人事資料增加非整日起迄上下班時段
      strSql = "insert into Staff_Busi_Trip(SB01,SB02,SB03,SB04,SB05,SB06,SB07,SB08,SB09,SB10,SB17,SB18) " & _
               "values(" & CNULL(strB1003) & "," & strUpdB1004 & "," & strUpdB1005 & _
               "," & strUpdB1006 & "," & strUpdB1007 & "," & CNULL(strUpdB1009) & "," & CNULL(strUpdB1010) & _
               "," & CNULL(strB1014) & "," & CNULL(strB1015) & "," & CNULL(strB1001) & _
               "," & CNULL(strUpdB1028) & "," & CNULL(strUpdB1029) & ")"
   End If
   cnnConnection.Execute strSql
   Exit Function
   
ErrHand:
   PUB_SavePMainFile = False
   cnnConnection.RollbackTrans
   If bolMsg = True Then MsgBox "新增失敗！" & vbCrLf & Err.Description & strErrTxt
End Function

'計算時數-請假單
'Modify by Sindy 2012/10/12 +CboB1008
Public Function PUB_CountDayHour(strB1003 As String, strB1004 As String, strB1005 As String, strB1006 As String, strB1007 As String, strSTime As String, strETime As String, ByRef strUpdB1009 As String, ByRef strUpdB1010 As String, CboB1008 As String, Optional bolMsg As Boolean = True, Optional bolResetCount As Boolean = True)
Dim tmpCalH As String
Dim temp As Variant ', bwk5hour As Boolean
   
   If bolResetCount = False Then
      '若日數或時數已有值,則不重新計算
      If Val(strUpdB1009) > 0 Or Val(strUpdB1010) > 0 Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2010/7/14 99029伊恩一天只上4個小時
   'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
   'Modify By Sindy 2012/7/9 上班時數為特殊者
'   bwk5hour = False
'   If strB1003 = "99029" Then bwk5hour = True
   Call Pub_GetSpecWorkHour(strB1003, strB1004)
   '2012/7/9 End
   
   If Trim(strB1004) <> "" And Trim(strB1005) <> "" And Trim(strB1006) <> "" And Trim(strB1007) <> "" Then
       If CheckIsTaiwanDate(ChangeWStringToTString(strB1004), False) = True And CheckIsTaiwanDate(ChangeWStringToTString(strB1006), False) = True Then
           'Modify By Sindy 2010/7/14 增加傳入bwk4hour
           'Modify By Sindy 2011/3/8 增加傳入bwk5hour
'              strSTime = "": strETime = ""
'              If cboSTime.Visible = True Then strSTime = Format(cboSTime.text, "hhmm")
'              If cboETime.Visible = True Then strETime = Format(cboETime.text, "hhmm")
           
           'Add By Sindy 2011/10/14 劉經理：國外及大陸出差天數計算應含休假日,國內出差則以實際工作時數計算
           'If Frame03.Visible = True And (txtB1014 = "3" Or txtB1014 = "4") Then
           'If strB1014 = "3" Or strB1014 = "4" Then
           'Modify By Sindy 2012/3/14
'           If strB1014 <> "" Then '出差
'               tmpCalH = CalDateTime(strB1004 & Format(strB1005, "0000"), strB1006 & Format(strB1007, "0000"), bwk5hour, strSTime, strETime, False)
'           Else
               'Modify By Sindy 2012/10/12 10.產假及11.流產假均有含假日天數
               If Left(Trim(CboB1008), 2) = "10" Or Left(Trim(CboB1008), 2) = "11" Then
                  tmpCalH = CalDateTime(strB1003, strB1004 & Format(strB1005, "0000"), strB1006 & Format(strB1007, "0000"), PUB_bWkSpec, strSTime, strETime, False)
               Else
               '2012/10/12 End
                  tmpCalH = CalDateTime(strB1003, strB1004 & Format(strB1005, "0000"), strB1006 & Format(strB1007, "0000"), PUB_bWkSpec, strSTime, strETime)
               End If
'           End If
           
'              'Add By Sindy 98/03/13 起始時間<=12時並且迄止時間>=13時30分者，減1小時
'              dblSTime = Val(txtB1005_1 & txtB1005_2)
'              dblETime = Val(txtB1007_1 & txtB1007_2)
'              If dblSTime <= 1200 And dblETime >= 1330 Then
'                  tmpCalH = tmpCalH - 1
'              End If
           
           If tmpCalH > "" Then
               'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
               'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
               'Modify By Sindy 2012/7/9 上班時數為特殊者
'               If strB1003 = "99029" Then
'                   If tmpCalH < 5 Then
'                       strUpdB1009 = 0
'                   Else
'                       temp = Split(CStr(Val(tmpCalH) / 5), ".")
'                       strUpdB1009 = temp(0)
'                   End If
'                   strUpdB1010 = Val(tmpCalH) - (Val(strUpdB1009) * 5)
               If Val(tmpCalH) < Val(PUB_intWkHour) Then
                   strUpdB1009 = 0
               Else
                   temp = Split(CStr(Val(tmpCalH) / PUB_intWkHour), ".")
                   strUpdB1009 = temp(0)
               End If
               strUpdB1010 = Val(tmpCalH) - (Val(strUpdB1009) * PUB_intWkHour)
           Else
               strUpdB1009 = ""
               strUpdB1010 = ""
               If bolMsg = True Then MsgBox "日期時間設定錯誤！", vbInformation, "輸入錯誤！"
               Exit Function
           End If
       Else
           strUpdB1009 = ""
           strUpdB1010 = ""
       End If
   Else
       strUpdB1009 = ""
       strUpdB1010 = ""
   End If
End Function

'Add By Sindy 2012/2/13
'檢查表單職代是否符合人事職代裡的設定資料
Public Function ChkIsDutyAgent(strB1001 As String, strB1003 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
   
   ChkIsDutyAgent = True
   
   strSql = "SELECT * FROM abs010 WHERE b1001='" & strB1001 & "' "
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      '加班單排除檢查
      If rsTmp.Fields("b1002") = 表單類別_加班 Then Exit Function
   End If
   
   strSql = "select * from abs001 where b0101='" & strB1003 & "' " & _
               "and (b0102 in (select b1104 from abs011 where b1101='" & strB1001 & "' AND b1102='1') " & _
                     "or b0103 in (select b1104 from abs011 where b1101='" & strB1001 & "' AND b1102='1') " & _
                     "or b0104 in (select b1104 from abs011 where b1101='" & strB1001 & "' AND b1102='1') " & _
                     "or b0105 in (select b1104 from abs011 where b1101='" & strB1001 & "' AND b1102='1') " & _
                     "or b0106 in (select b1104 from abs011 where b1101='" & strB1001 & "' AND b1102='1') " & _
                     "or b0107 in (select b1104 from abs011 where b1101='" & strB1001 & "' AND b1102='1') " & _
                    ") "
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI <> 1 Or rsTmp.RecordCount = 0 Then ChkIsDutyAgent = False
End Function

'Add By Sindy 2012/3/14 計算出差單時數
Public Function PUB_CountHour_Busi_Trip(strB1004 As String, strB1005 As String, strB1006 As String, strB1007 As String, ByRef m_Day As Integer, ByRef m_Hour As Double)
Dim dblDay As Double, dblHour As Double
   
   strB1004 = DBDATE(strB1004)
   If InStr(strB1005, ":") > 0 Then
      If strB1005 = "24:00" Then
         strB1005 = "2400"
      Else
         strB1005 = Format(strB1005, "hhmm")
      End If
   Else
      strB1005 = Format(strB1005, "0000")
   End If
   strB1006 = DBDATE(strB1006)
   If InStr(strB1007, ":") > 0 Then
      If strB1007 = "24:00" Then
         strB1007 = "2400"
      Else
         strB1007 = Format(strB1007, "hhmm")
      End If
   Else
      strB1007 = Format(strB1007, "0000")
   End If
   
   If Val(strB1004) > 0 Then
      m_Day = 0: m_Hour = 0: strDate = strB1004
      If Val(strB1004) <> Val(strB1006) Then
         For dblDay = Val(strB1004) To Val(strB1006)
            dblDay = strDate
            If dblDay <> Val(strB1004) And dblDay <> Val(strB1006) Then '第一天和最後一天要另外計算
               m_Day = m_Day + 1
            End If
            strDate = DBDATE(ChangeWStringToTString(DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(CStr(dblDay)))))))
            If strDate > strB1006 Then dblDay = strDate
         Next dblDay
         '第一天
         '以分鐘為單位, 取至小數第一位, 四捨五入
         dblHour = Round(((Val(24) * 60) - ((Val(Left(strB1005, 2)) * 60) + Val(Mid(strB1005, 3, 2)))) / 60, 1)
         If dblHour >= 8 Then
            m_Day = m_Day + 1
         Else
            m_Hour = m_Hour + dblHour
         End If
         '最後一天
         '以分鐘為單位, 取至小數第一位, 四捨五入
         dblHour = Round(((Val(Left(strB1007, 2)) * 60) + Val(Mid(strB1007, 3, 2))) / 60, 1)
         If dblHour >= 8 Then
            m_Day = m_Day + 1
         Else
            m_Hour = m_Hour + dblHour
         End If
      Else
         '以分鐘為單位, 取至小數第一位, 四捨五入
         dblHour = Round((((Val(Left(strB1007, 2)) * 60) + Val(Mid(strB1007, 3, 2))) - ((Val(Left(strB1005, 2)) * 60) + Val(Mid(strB1005, 3, 2)))) / 60, 1)
         If dblHour >= 8 Then
            m_Day = m_Day + 1
         Else
            m_Hour = m_Hour + dblHour
         End If
      End If
   End If
End Function

'Add By Sindy 2016/12/27 計算加班時數
Public Function PUB_CountHour_Overtime(txtB1004 As String, txtB1003 As String, txtB1007_1 As String, _
   txtB1007_2 As String, txtB1005_1 As String, txtB1005_2 As String, ByRef txtB101213 As String) As Double
   
   PUB_CountHour_Overtime = 0
   '以半小時為單位
   'PUB_CountHour_Overtime = ((((Val(txtB1007_1) * 60) + Val(txtB1007_2)) - ((Val(txtB1005_1) * 60) + Val(txtB1005_2))) \ 30) * 0.5
   '以分鐘為單位, 取至小數第一位, 四捨五入
   PUB_CountHour_Overtime = Round((((Val(txtB1007_1) * 60) + Val(txtB1007_2)) - ((Val(txtB1005_1) * 60) + Val(txtB1005_2))) / 60, 1)
   If PUB_CountHour_Overtime < 0 Then PUB_CountHour_Overtime = 0
   
   '假日時數
   'Modify By Sindy 2012/8/15 增加檢查颱風假
   'If ChkWorkDay(ChangeTStringToWString(txtB1004)) = False Then
   'Modify By Sindy 2016/12/26
   If DBDATE(txtB1004) >= 20161223 Then '(2016/12/23開始實施)
'      If ChkWorkDay(DBDATE(txtB1004), txtB1003, True) = False Then '假日加班
'         '非週六的加班計算
'         If Weekday(Format(DBDATE(txtB1004), "####-##-##")) <> 7 Then
'            If Val(PUB_CountHour_Overtime) <= 8 Then
'               txtB101213 = 8
'            Else
'               txtB101213 = PUB_CountHour_Overtime
'            End If
'         '週六的加班計算
'         Else
'            If Val(PUB_CountHour_Overtime) <= 4 Then
'               txtB101213 = 4
'            ElseIf Val(PUB_CountHour_Overtime) <= 8 Then
'               txtB101213 = 8
'            ElseIf Val(PUB_CountHour_Overtime) <= 12 Then
'               txtB101213 = 12
'            Else
'               txtB101213 = PUB_CountHour_Overtime
'            End If
'         End If
'      Else '平日加班(工作日加班沒有換算的問題)
'         txtB101213 = PUB_CountHour_Overtime
'      End If
      'Modify By Sindy 2017/1/3 換算加班時數
      txtB101213 = PUB_Overtime_TransDay(txtB1004, txtB1003, PUB_CountHour_Overtime)
   '原加班多少時數,就是多少時數
   Else
      txtB101213 = PUB_CountHour_Overtime
   End If
End Function

'Add By Sindy 2017/1/3 換算加班時數
Public Function PUB_Overtime_TransDay(txtB1004 As String, txtB1003 As String, dbl_Day As Double) As String
   If ChkWorkDay(DBDATE(txtB1004), txtB1003, True) = False Then '假日加班
      '非週六的加班計算
      If Weekday(Format(DBDATE(txtB1004), "####-##-##")) <> 7 Then
         If Val(dbl_Day) <= 8 Then
            PUB_Overtime_TransDay = 8
         Else
            PUB_Overtime_TransDay = dbl_Day
         End If
      '週六的加班計算
      Else
         'Modify By Sindy 2018/2/8
         If Val(DBDATE(txtB1004)) >= 20180301 Then
            PUB_Overtime_TransDay = dbl_Day '實際加班時數
         Else
         '2018/2/8 END
            If Val(dbl_Day) <= 4 Then
               PUB_Overtime_TransDay = 4
            ElseIf Val(dbl_Day) <= 8 Then
               PUB_Overtime_TransDay = 8
            ElseIf Val(dbl_Day) <= 12 Then
               PUB_Overtime_TransDay = 12
            Else
               PUB_Overtime_TransDay = dbl_Day
            End If
         End If
      End If
   Else '平日加班(工作日加班沒有換算的問題)
      PUB_Overtime_TransDay = dbl_Day
   End If
End Function

'Add By Sindy 2012/7/9 人事系統:上班時數為特殊者
'******************************************************************************************
'注意:
'  frmHTAauto: ChkWorkTime 打卡異常檢查 (若有特殊的打卡時段,要檢查此函數是否需要一併修正)
'******************************************************************************************
'strDate:適用時間
Public Sub Pub_GetSpecWorkHour(ByVal strUser As String, ByVal strDate As String, _
                               Optional ByRef strStarWorkTime As String, _
                               Optional ByRef strEndWorkTime As String)
   '預設值
   strStarWorkTime = ""
   strEndWorkTime = ""
   PUB_bWkSpec = False
   PUB_intWkHour = 8     '2013/1/22 modify by sonia 改預設8小時,原為預設0
   PUB_bSpecY = False '上班時數特殊者,在過渡期的那一個年度不列印年度累計
   
   '*********************************************************************
   '2013/1/22 ADD BY SONIA 注意此段若有修改要通知秀玲改年終獎金計算程式
   '*********************************************************************
   'Add By Sindy 2013/10/30 使用於打卡系統
   'Modify By Sindy 2023/7/25 B2024.劉美英接朱小姐的工作,上班時段一樣
   If strUser = "96006" Or strUser = "B2024" Then '96006.朱苡甄上班時間為13:30-20:30
      PUB_bWkSpec = True
      PUB_intWkHour = 7 '8 'Modify By Sindy 2015/7/27 劉經理指示改為7小時
      strStarWorkTime = "1330"
      strEndWorkTime = "2030"
   '2013/10/30 END
   'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
   'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
   ElseIf strUser = "99029" Then
      'Add By Sindy 2012/10/9 99029伊恩一天只上6個小時,20121001開始適用, 11:30-17:30
      If Val(DBDATE(strDate)) >= 20121001 Then
         PUB_bWkSpec = True
         PUB_intWkHour = 6
         strStarWorkTime = "1130"
         strEndWorkTime = "1730"
      Else
      '2012/10/9 End
         PUB_bWkSpec = True
         PUB_intWkHour = 5
         strStarWorkTime = "1230"
         strEndWorkTime = "1730"
      End If
      'Add By Sindy 2012/10/9
      If Left(DBDATE(strDate), 4) = Left(20121001, 4) Then
         PUB_bSpecY = True '適用的過渡期
      End If
      '2012/10/9 End
   'Modify By Sindy 2012/7/9 84043尤春彬一天只上4個小時, 8:00-12:00
   ElseIf strUser = "84043" Then
      '20120701 開始適用
      'Modify By Sindy 2016/4/12 尤春彬自2016/3/1起改全職
      'If Val(DBDATE(strDate)) >= 20120701 Then
      If Val(DBDATE(strDate)) >= 20120701 And Val(DBDATE(strDate)) <= 20160229 Then
      '2016/4/12 END
         PUB_bWkSpec = True
         PUB_intWkHour = 4
         strStarWorkTime = "0800"
         strEndWorkTime = "1200"
      End If
      If Left(DBDATE(strDate), 4) = Left(20120701, 4) Then
         PUB_bSpecY = True '適用的過渡期
      'Add By Sindy 2016/4/12 尤春彬自2016/3/1起改全職
      ElseIf Left(DBDATE(strDate), 4) = Left(20160229, 4) Then
         PUB_bSpecY = True '適用的過渡期
      '2016/4/12 END
      End If
   'Modify By Sindy 2013/7/23 73029廖宗岳一天只上4個小時 13:30-17:30,週五全天(休週五算請2天)
   ElseIf strUser = "73029" Then
      '20130801 開始適用
      If Val(DBDATE(strDate)) >= 20130801 Then
         PUB_bWkSpec = True
         If Weekday(Format(DBDATE(strDate), "####-##-##")) <> 6 Then '非星期五
            PUB_intWkHour = 4
            strStarWorkTime = "1330"
            strEndWorkTime = "1730"
         End If
      End If
      If Left(DBDATE(strDate), 4) = Left(20130801, 4) Then
         PUB_bSpecY = True '適用的過渡期
      End If
   ElseIf strUser = "A4016" Then 'A4016.德文顧問.黎康翰上班時間為9:00-16:20
      'Modify By Sindy 2018/4/9
      '20180501 截止特殊設定,改為正常上下班
      If Val(DBDATE(strDate)) < 20180501 Then
      '2018/4/9 END
         PUB_bWkSpec = True
         PUB_intWkHour = 6
         strStarWorkTime = "0900"
         strEndWorkTime = "1620"
      End If
      If Left(DBDATE(strDate), 4) = Left(20180501, 4) Then
         PUB_bSpecY = True '適用的過渡期
      End If
   'Add By Sindy 2018/4/27
   ElseIf strUser = "A7007" Then 'A7007.伊恩上班時間為 13:30-17:30，每日4小時
      '20180501 開始適用
      If Val(DBDATE(strDate)) >= 20180501 Then
         PUB_bWkSpec = True
         PUB_intWkHour = 4
         strStarWorkTime = "1330"
         strEndWorkTime = "1730"
      End If
      If Left(DBDATE(strDate), 4) = Left(20180501, 4) Then
         PUB_bSpecY = True '適用的過渡期
      End If
   'Add By Sindy 2020/6/4 + 鄭皓云(A9004)13:30-17:30，每日4小時 (同Iain)
   ElseIf strUser = "A9004" Then
      PUB_bWkSpec = True
      PUB_intWkHour = 4
      strStarWorkTime = "1330"
      strEndWorkTime = "1730"
   'Add By Sindy 2018/8/31
   ElseIf strUser = "A7023" Then 'A7023.宗家澔上班時間為 8:00-12:00，每日4小時
      'Modify By Sindy 2018/10/15 宗家澔自10/15起上全日班
      '20180901 開始適用
'      If Val(DBDATE(strDate)) >= 20180901 Then
'         PUB_bWkSpec = True
'         PUB_intWkHour = 4
'         strStarWorkTime = "0800"
'         strEndWorkTime = "1200"
'      End If
      If Val(DBDATE(strDate)) < 20181015 Then
      '2018/10/15 END
         PUB_bWkSpec = True
         PUB_intWkHour = 4
         strStarWorkTime = "0800"
         strEndWorkTime = "1200"
      End If
      If Left(DBDATE(strDate), 4) = Left(20180901, 4) Then
         PUB_bSpecY = True '適用的過渡期
      End If
   End If
End Sub

'Add By Sindy 2012/7/12 移為共用函數
'檢查特別假
'Modify By Sindy 2017/11/3 int08Day as Integer ==> dbl08Day as Double
'Add By Sindy 2017/11/3 + , Optional dbl08Hour As Double
'Modify by Sindy 2018/12/3 + , Optional strB1001 As String = ""
Public Function ChkSA06_08(txtB1009 As String, txtB1010 As String, txtB1003 As String, _
   txtB1004 As String, txtB1005_1 As String, txtB1005_2 As String, txtB1006 As String, _
   txtB1007_1 As String, txtB1007_2 As String, CboB1008 As String, dbl08Day As Integer, _
   Optional dbl08Hour As Double, Optional strB1001 As String = "") As Boolean
Dim dST40 As Double, dblRestHour As Double
Dim strSDate As String, strEDate As String
'Dim bwk5hour As Boolean
'Dim strSTime As String, strETime As String
Dim strB1009 As String, temp As Variant
Dim strB1010 As String, dblTmpHour As Double
Dim i As Integer, intCnt As Integer
   
   ChkSA06_08 = True
   
   If (txtB1009 = "" Or txtB1009 = "0") And (txtB1010 = "" Or txtB1010 = "0") Then Exit Function
   
   If txtB1003 <> "" And Trim(txtB1004) <> "" And (Trim(txtB1005_1) <> "" And Trim(txtB1005_1) <> "00") And Trim(txtB1005_2) <> "" And Trim(txtB1006) <> "" And (Trim(txtB1007_1) <> "" And Trim(txtB1007_1) <> "00") And Trim(txtB1007_2) <> "" Then
      If Left(Trim(CboB1008), 2) = "08" Then '08.特別假
         If Left(DBDATE(txtB1004), 4) = Left(DBDATE(txtB1006), 4) Then
            '同年
            intCnt = 2 '跑1次迴圈
         Else
            '跨年
            intCnt = 1 '跑2次迴圈
         End If
         For i = intCnt To 2
            dblRestHour = 0
            '取得目前已累計的特別假時數
            If i = 1 Then
               strSDate = Left(DBDATE(txtB1004), 4) & "0101"
               strEDate = Left(DBDATE(txtB1004), 4) & "1231"
            Else
               strSDate = Left(DBDATE(txtB1006), 4) & "0101"
               strEDate = Left(DBDATE(txtB1006), 4) & "1231"
            End If
            
            strSql = "select sum(nvl(SA07,0)) t1,sum(nvl(SA08,0)) t2,sa01 From staff_Absence Where sa01='" & txtB1003 & "' and SA06='08' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') group by sa01 "
            'Add By Sindy 2018/12/3 加未核准假單
            'Modify By Sindy 2018/12/20 + 過濾已核准簽核資料
            strSql = strSql & " union" & _
                     " select sum(nvl(b1009,0)) t1,sum(nvl(b1010,0)) t2,b1003 from abs010 where b1002='" & 表單類別_請假 & "' and b1003='" & txtB1003 & "' and b1008='08'" & _
                     " and (b1004 >= '" & strSDate & "' and b1006 <= '" & strEDate & "')" & _
                     " and b1001 not in(select sa09 from staff_absence where sa01='" & txtB1003 & "' and sa06='08' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') and sa09 is not null)" & _
                     IIf(strB1001 <> "", " and b1001<>'" & strB1001 & "'", "") & _
                     " and B1018 not in('" & 註銷 & "','" & 已核准 & "')" & _
                     " group by b1003"
            strSql = "select sum(t1),sum(t2),sa01 from(" & strSql & ") group by sa01"
            '2018/12/3 END
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               '特別假只能請整日,因此直接Sum天數引用
               'Modify By Sindy 2017/11/3 107年1月1日起,開始特別假可以請半天(4小時)
               If Not IsNull(RsTemp.Fields(0)) Then
                  dblRestHour = RsTemp.Fields(0) * PUB_intWkHour
               End If
               If Not IsNull(RsTemp.Fields(1)) Then
                  dblRestHour = dblRestHour + RsTemp.Fields(1)
               End If
            End If
            
            '取得可休特別假
            strSql = "select st40 from staff where st01='" & txtB1003 & "'"
            intI = 1
            dST40 = 0
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If Not IsNull(RsTemp.Fields(0).Value) Then
                  'dST40 = RsTemp.Fields(0).Value
                  dST40 = RsTemp.Fields(0).Value * PUB_intWkHour
               End If
            End If
            'Modify By Sindy 2017/11/2 改新制還適用嗎?先Mark起來
'            'Add By Sindy 2011/11/14 提前請明年的特休且可休日未滿7日者,以7日計
'            If i = 1 Then
'               If Left(DBDATE(txtB1004), 4) > Left(strSrvDate(1), 4) And dST40 < 7 Then
'                  dST40 = 7
'               End If
'            Else
'               If Left(DBDATE(txtB1006), 4) > Left(strSrvDate(1), 4) And dST40 < 7 Then
'                  dST40 = 7
'               End If
'            End If
            
            dblTmpHour = 0 'Add By Sindy 2018/12/3
            If Left(DBDATE(txtB1004), 4) <> Left(DBDATE(txtB1006), 4) Then
               '99029伊恩一天只上5個小時
'               bwk5hour = False
'               If txtB1003 = "99029" Then bwk5hour = True
'               strSTime = "": strETime = ""
'               If cboSTime <> "" Then strSTime = Format(cboSTime, "hhmm")
'               If cboETime <> "" Then strETime = Format(cboETime, "hhmm")
               If i = 1 Then
                  'Call Pub_GetSpecWorkHour(txtB1003, txtB1004) 'Add By Sindy 2012/7/9 上班時數為特殊者
                  '傳回小時
                  dblTmpHour = CDbl(CalDateTime(txtB1003, txtB1004 & Format(txtB1005_1, "00") & Format(txtB1005_2, "00"), Left(txtB1004, Len(txtB1004) - 4) & "1231" & Format(txtB1007_1, "00") & Format(txtB1007_2, "00"), PUB_bWkSpec, "", ""))
               Else
                  'Call Pub_GetSpecWorkHour(txtB1003, Left(txtB1006, Len(txtB1006) - 4) & "0101") 'Add By Sindy 2012/7/9 上班時數為特殊者
                  '傳回小時
                  dblTmpHour = CDbl(CalDateTime(txtB1003, Left(txtB1006, Len(txtB1006) - 4) & "0101" & Format(txtB1005_1, "00") & Format(txtB1005_2, "00"), txtB1006 & Format(txtB1007_1, "00") & Format(txtB1007_2, "00"), PUB_bWkSpec, "", ""))
               End If
               'If strB1009 > "" Then
               If dblTmpHour > 0 Then
                  'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
                  'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
                  'Modify By Sindy 2012/7/9 上班時數為特殊者
'                  If txtB1003 = "99029" Then
'                      If strB1009 < 5 Then
'                          strB1009 = 0
'                      Else
'                          temp = Split(CStr(Val(strB1009) / 5), ".")
'                          strB1009 = temp(0)
'                      End If
                  'Modify By Sindy 2017/11/3 107年1月1日起,特別假可以請半天(4小時)
'                  If Val(strB1009) < Val(PUB_intWkHour) Then
'                      strB1009 = 0
'                  Else
'                      temp = Split(CStr(Val(strB1009) / PUB_intWkHour), ".")
'                      strB1009 = temp(0)
'                  End If
                  strB1009 = Fix(dblTmpHour / PUB_intWkHour) '天
                  strB1010 = dblTmpHour Mod PUB_intWkHour '小時
               End If
            Else
               strB1009 = txtB1009
               strB1010 = txtB1010 'Add By Sindy 2017/11/3
               'Add By Sindy 2018/12/3
               '傳回小時
               If strB1009 <> "" Then dblTmpHour = strB1009 * PUB_intWkHour
               If strB1010 <> "" Then dblTmpHour = dblTmpHour + strB1010
               '2018/12/3 END
            End If
            
            If dST40 = 0 Then
               MsgBox "無特別假!!!", vbExclamation + vbOKOnly
               ChkSA06_08 = False
               Exit Function
            End If
            
            '檢查特別假是否有超額
            If txtB1009 <> "" Or txtB1010 <> "" Then
            'If Val(strB1009) <> 0 Then
               'Add By Sindy 2011/9/21 資料修改時,(當年已累計之特別假時數)會含目前該筆時數在裡面,須先扣除再比對
               If dbl08Day > 0 Then dblRestHour = dblRestHour - (dbl08Day * PUB_intWkHour)
               '2011/9/21 End
               If dbl08Hour > 0 Then dblRestHour = dblRestHour - dbl08Hour 'Add By Sindy 2017/11/3
               If (dblRestHour + Val(dblTmpHour)) > dST40 Then
                  MsgBox "特別假天數已超額!!!", vbExclamation + vbOKOnly
                  ChkSA06_08 = False
                  Exit Function
               End If
            End If
         Next i
      End If
   End If
End Function

'Add By Sindy 2014/12/31
'檢查健檢假 : 一年只能請一天(不可超過8小時)且最多只能分2次請假
Public Function ChkSA06_23(txtB1009 As String, txtB1010 As String, txtB1003 As String, txtB1004 As String, txtB1005_1 As String, txtB1005_2 As String, txtB1006 As String, txtB1007_1 As String, txtB1007_2 As String, CboB1008 As String, int23Day As Integer, dbl23Hour As Double) As Boolean
Dim dblHour As Double, dblTempHour As Double
Dim strSDate As String, strEDate As String
Dim i As Integer, intCnt As Integer
   
   ChkSA06_23 = True
   
   If (txtB1009 = "" Or txtB1009 = "0") And (txtB1010 = "" Or txtB1010 = "0") Then Exit Function
   
   If txtB1003 <> "" And Trim(txtB1004) <> "" And (Trim(txtB1005_1) <> "" And Trim(txtB1005_1) <> "00") And Trim(txtB1005_2) <> "" And Trim(txtB1006) <> "" And (Trim(txtB1007_1) <> "" And Trim(txtB1007_1) <> "00") And Trim(txtB1007_2) <> "" Then
      If Left(Trim(CboB1008), 2) = "23" Then '23.健檢假
         If Left(DBDATE(txtB1004), 4) = Left(DBDATE(txtB1006), 4) Then
            '同年
            intCnt = 2 '跑1次迴圈
         Else
            '跨年
            intCnt = 1 '跑2次迴圈
         End If
         For i = intCnt To 2
            dblHour = 0: dblTempHour = 0
            '取得目前已累計的時數
            If i = 1 Then
               strSDate = Left(DBDATE(txtB1004), 4) & "0101"
               strEDate = Left(DBDATE(txtB1004), 4) & "1231"
            Else
               strSDate = Left(DBDATE(txtB1006), 4) & "0101"
               strEDate = Left(DBDATE(txtB1006), 4) & "1231"
            End If
            
            '上班時數是否為特殊者
            If i = 1 Then
               Call Pub_GetSpecWorkHour(txtB1003, txtB1004)
            Else
               Call Pub_GetSpecWorkHour(txtB1003, Left(txtB1006, Len(txtB1006) - 4) & "0101")
            End If
            
            '檢查已請次數
            strSql = "select count(*) From staff_Absence Where sa01='" & txtB1003 & "' and SA06='23' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If RsTemp.Fields(0) >= 2 Then
                  MsgBox "健檢假只能分2次申請!!!", vbExclamation + vbOKOnly
                  ChkSA06_23 = False
                  Exit Function
               End If
            End If
            
            '檢查已請幾小時
            strSql = "select sum(nvl(SA07,0)),sum(nvl(SA08,0)),sa01 From staff_Absence Where sa01='" & txtB1003 & "' and SA06='23' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') group by sa01 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If Not IsNull(RsTemp.Fields(0)) Then
                  dblHour = Val(RsTemp.Fields(0)) * PUB_intWkHour
               End If
               If Not IsNull(RsTemp.Fields(1)) Then
                  dblHour = dblHour + Val(RsTemp.Fields(1))
               End If
            End If
            
            '計算這次填的假單時數
            If Left(DBDATE(txtB1004), 4) <> Left(DBDATE(txtB1006), 4) Then
               If i = 1 Then
                  dblTempHour = CalDateTime(txtB1003, txtB1004 & Format(txtB1005_1, "00") & Format(txtB1005_2, "00"), Left(txtB1004, Len(txtB1004) - 4) & "1231" & Format(txtB1007_1, "00") & Format(txtB1007_2, "00"), PUB_bWkSpec, "", "")
               Else
                  dblTempHour = CalDateTime(txtB1003, Left(txtB1006, Len(txtB1006) - 4) & "0101" & Format(txtB1005_1, "00") & Format(txtB1005_2, "00"), txtB1006 & Format(txtB1007_1, "00") & Format(txtB1007_2, "00"), PUB_bWkSpec, "", "")
               End If
            Else
               If txtB1009 <> "" Then
                  dblTempHour = Val(txtB1009) * PUB_intWkHour
               End If
               If txtB1010 <> "" Then
                  dblTempHour = dblTempHour + Val(txtB1010)
               End If
            End If
            
            '檢查是否有超額請假
            If txtB1009 <> "" Or txtB1010 <> "" Then
               '資料修改時,(當年已累計之時數)會含目前該筆時數在裡面,須先扣除再比對
               If int23Day > 0 Then dblHour = dblHour - (int23Day * PUB_intWkHour)
               If dbl23Hour > 0 Then dblHour = dblHour - dbl23Hour
               If (dblHour + dblTempHour) > PUB_intWkHour Then
                  MsgBox "健檢假時數已超額!!!", vbExclamation + vbOKOnly
                  ChkSA06_23 = False
                  Exit Function
               End If
            End If
         Next i
      End If
   End If
End Function

'Add By Sindy 2024/12/10
'檢查可補休
Public Function ChkSA06_14(txtB1009 As String, txtB1010 As String, txtB1003 As String, _
   txtB1004 As String, txtB1005_1 As String, txtB1005_2 As String, txtB1006 As String, _
   txtB1007_1 As String, txtB1007_2 As String, CboB1008 As String, dbl08Day As Integer, _
   Optional dbl08Hour As Double, Optional strB1001 As String = "") As Boolean
Dim dblRestHour As Double
Dim strSDate As String, strEDate As String
Dim strB1009 As String, temp As Variant
Dim strB1010 As String, dblTmpHour As Double
Dim i As Integer, intCnt As Integer
Dim strSRR01 As String, dblTotSRR03 As Double
   
   ChkSA06_14 = True
   
   'R投資單位不控管補休
   If Left(PUB_GetST03(txtB1003), 1) = "R" Then Exit Function
   
   If (txtB1009 = "" Or txtB1009 = "0") And (txtB1010 = "" Or txtB1010 = "0") Then Exit Function
   
   If txtB1003 <> "" And Trim(txtB1004) <> "" And (Trim(txtB1005_1) <> "" And Trim(txtB1005_1) <> "00") And Trim(txtB1005_2) <> "" And Trim(txtB1006) <> "" And (Trim(txtB1007_1) <> "" And Trim(txtB1007_1) <> "00") And Trim(txtB1007_2) <> "" Then
      If Left(Trim(CboB1008), 2) = "14" Then '14.補休
         If Left(DBDATE(txtB1004), 4) = Left(DBDATE(txtB1006), 4) Then
            '同年
            intCnt = 2 '跑1次迴圈
         Else
            '跨年
            intCnt = 1 '跑2次迴圈
         End If
         For i = intCnt To 2
            dblRestHour = 0
            If i = 1 Then
               strSDate = DBDATE(txtB1004)
            Else
               If Left(DBDATE(txtB1004), 4) = Left(DBDATE(txtB1006), 4) Then
                  strSDate = DBDATE(txtB1004)
               Else
                  strSDate = Left(DBDATE(txtB1006), 4) & "0101"
               End If
            End If
            '取得有效的可補休發生日期和剩餘時數
            Call GetCurrFor14RestDay(txtB1003, 2, strSDate, strSRR01, dblTotSRR03)
            If strSRR01 <> "" Then
               strSDate = DBDATE(strSRR01)
            End If
            strEDate = Left(DBDATE(strSDate), 4) & "1231"
            
            If dblTotSRR03 = 0 Then
               MsgBox "無可補休假!!!" & vbCrLf & "（請先通知人事處輸入可補休資料）", vbExclamation + vbOKOnly
               ChkSA06_14 = False
               Exit Function
            End If
            
            strSql = "select sum(nvl(SA07,0)) t1,sum(nvl(SA08,0)) t2,sa01 From staff_Absence Where sa01='" & txtB1003 & "'" & _
                     " and SA06='14' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') group by sa01 "
            '加未核准假單
            '過濾已核准簽核資料
            strSql = strSql & " union" & _
                     " select sum(nvl(b1009,0)) t1,sum(nvl(b1010,0)) t2,b1003 from abs010 where b1002='" & 表單類別_請假 & "' and b1003='" & txtB1003 & "' and b1008='14'" & _
                     " and (b1004 >= '" & strSDate & "' and b1006 <= '" & strEDate & "')" & _
                     " and b1001 not in(select sa09 from staff_absence where sa01='" & txtB1003 & "'" & _
                     " and sa06='14' and (sa02 >= '" & strSDate & "' and sa04 <= '" & strEDate & "') and sa09 is not null)" & _
                     IIf(strB1001 <> "", " and b1001<>'" & strB1001 & "'", "") & _
                     " and B1018 not in('" & 註銷 & "','" & 已核准 & "')" & _
                     " group by b1003"
            strSql = "select sum(t1),sum(t2),sa01 from(" & strSql & ") group by sa01"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If Not IsNull(RsTemp.Fields(0)) Then
                  dblRestHour = RsTemp.Fields(0) * PUB_intWkHour
               End If
               If Not IsNull(RsTemp.Fields(1)) Then
                  dblRestHour = dblRestHour + RsTemp.Fields(1)
               End If
            End If
            
            dblTmpHour = 0
            If Left(DBDATE(txtB1004), 4) <> Left(DBDATE(txtB1006), 4) Then
               If i = 1 Then
                  '傳回小時
                  dblTmpHour = CDbl(CalDateTime(txtB1003, txtB1004 & Format(txtB1005_1, "00") & Format(txtB1005_2, "00"), Left(txtB1004, Len(txtB1004) - 4) & "1231" & Format(txtB1007_1, "00") & Format(txtB1007_2, "00"), PUB_bWkSpec, "", ""))
               Else
                  '傳回小時
                  dblTmpHour = CDbl(CalDateTime(txtB1003, Left(txtB1006, Len(txtB1006) - 4) & "0101" & Format(txtB1005_1, "00") & Format(txtB1005_2, "00"), txtB1006 & Format(txtB1007_1, "00") & Format(txtB1007_2, "00"), PUB_bWkSpec, "", ""))
               End If
               If dblTmpHour > 0 Then
                  strB1009 = Fix(dblTmpHour / PUB_intWkHour) '天
                  strB1010 = dblTmpHour Mod PUB_intWkHour '小時
               End If
            Else
               strB1009 = txtB1009
               strB1010 = txtB1010
               '傳回小時
               If strB1009 <> "" Then dblTmpHour = strB1009 * PUB_intWkHour
               If strB1010 <> "" Then dblTmpHour = dblTmpHour + strB1010
            End If
            
            '檢查可補休假是否有超額
            If txtB1009 <> "" Or txtB1010 <> "" Then
               '資料修改時,會含目前該筆時數在裡面,須先扣除再比對
               If dbl08Day > 0 Then dblRestHour = dblRestHour - (dbl08Day * PUB_intWkHour)
               If dbl08Hour > 0 Then dblRestHour = dblRestHour - dbl08Hour
               If (dblRestHour + Val(dblTmpHour)) > dblTotSRR03 Then
                  MsgBox "可補休時數已超額!!!", vbExclamation + vbOKOnly
                  ChkSA06_14 = False
                  Exit Function
               End If
            End If
         Next i
      End If
   End If
End Function

'Added by Morgan 2016/2/23
'讀取健保費率
Public Function PUB_GetNhiRate(pDate As String) As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select NHR02 from NHI2NDRATE where NHR01 <= " & Val(pDate) & " order by NHR01 desc"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      PUB_GetNhiRate = Val("" & rsQuery.Fields(0))
   End If
   Set rsQuery = Nothing
End Function

'2013/1/17 ADD BY SONIA 計算補充保費
Public Sub PUB_NHI2nd(ByVal strNHI01 As String, ByVal strNHI02 As String, ByVal strNHI03 As String, ByVal strNHI04 As String, ByVal strNHI07 As String, Optional ByRef strNHI05 As String, Optional ByRef strNHI06 As String, Optional ByRef strNHI08 As String, Optional ByRef strNHI10 As String, Optional ByRef strPayCompany As String, Optional ByRef strNHI13 As String)
'傳入參數
'strNHI01 所得人代號
'strNHI02 給付日期
'strNHI03 格式代號
'strNHI04 資料來源      1年終獎金, 2獎金明細, 3同仁其他給付, 4翻譯費, 5複委託, 0 其他來源
'strNHI07 給付金額
'傳出參數
'strNHI05 當月投保金額
'strNHI06 補充保費
'strNHI08 補充保險費費基

Dim intQ As Integer, stSQL As String
Dim stHiComp As String '健保投保公司別
Dim rsQuery As ADODB.Recordset
'Modified by Morgan 2013/2/25
'需考慮多個編號 Ex.林信昌有68007,68091,68092,F5162,F5644,F5645多個編號
'Dim strNo As String      '內翻之所內編號
Dim strNoList As String   '其他相同ID的編號
'end 2013/2/25
Dim strNHIRate As Double  '費率
Dim lngLimit As Long      '兼職所得扣繳下限(薪資所得)
Dim lngLimitOther As Long '非薪資所得扣繳下限(利息,股利,租金,執行業務收入) add by sonia 2016/1/7
Dim strST04 As String     '2015/2/2 ADD BY SONIA 退休人員隔年發放年終時,財務自行輸入年終資料,補充保費列入其他所得人(即NHI05=0)
     
   '兼職所得 超過1千萬部分及未達一定金額者,免予扣取
   
   'add by sonia 2024/2/26 113/1/1 起基本薪資改 27470 --辜
   If DBDATE(strNHI02) >= "20240101" Then
      lngLimit = 27470
      
   'Added by Morgan 2019/1/7 108/1/1 起基本薪資改 23100 --辜
   ElseIf DBDATE(strNHI02) >= "20190101" Then
      lngLimit = 23100
      
   'Added by Morgan 2018/6/25 107/1/1 起基本薪資改 22000
   ElseIf DBDATE(strNHI02) >= "20180101" Then
      lngLimit = 22000
      
   'Added by Morgan 2015/6/30 104/7/1 起基本薪資改 20008
   ElseIf DBDATE(strNHI02) >= "20150701" Then
      lngLimit = 20008
      
   'Added by Morgan2014/7/31 103/9/1 起基本薪資改 19273
   ElseIf DBDATE(strNHI02) >= "20140901" Then
      lngLimit = 19273
   
   Else
      lngLimit = 5000
   End If
   
   'add by sonia 2016/1/7 '非薪資所得扣繳下限 105/1/1起下限改20000
   If DBDATE(strNHI02) >= "20160101" Then
      lngLimitOther = 20000
   Else
      lngLimitOther = 5000
   End If
   'end 2016/1/7
   
   strNHI05 = "": strNHI06 = ""
   strNHI08 = Val(strNHI07)
   strNHI13 = ""
   
   '讀取當時費率
   'Modified by Morgan 2016/2/23 改共用
   'stSQL = "select NHR02 from NHI2NDRATE where NHR01 <= " & Val(strNHI02) & " order by NHR01 desc"
   'intQ = 1
   'Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   'If intQ = 1 Then
   '   strNHIRate = Val("" & rsQuery.Fields(0))
   'End If
   strNHIRate = Val(PUB_GetNhiRate(strNHI02))
   'end 2016/2/23
      
   '判斷格式代號
   If strNHI03 <> "50" Then  '非薪資所得超過下限以上才扣,上限1千萬
   
      'Added by Morgan 2024/6/21
      '股利的部分, 除了超過2萬元才需要扣,雇主另有計算公式:"單次給付金額超過已列入投保金額計算部分達2萬元"
      '目前只有何金柱68099，因無月薪資所以固定抓薪資基本檔的投保金額*12個月計算--婉莘
      If strNHI03 = "54" Then
         stSQL = "select sd47 from salarydata where sd01='" & strNHI01 & "' and sd11='Y' and sd19='" & strPayCompany & "' and sd47>0" & _
            " union select sd47 from OtherIncomer,staff,salarydata where oi01='" & strNHI01 & "' and st26(+)=oi02 and st04='1' and sd01(+)=st01 and sd11='Y' and sd19='" & strPayCompany & "' and sd47>0"
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            strNHI07 = strNHI07 - (rsQuery(0) * 12)
         End If
      End If
      'end 2024/6/21
      
      'modify by sonia 2016/1/7
      'If Val(strNHI07) < 5000 Then
      If Val(strNHI07) < lngLimitOther Then
         strNHI08 = 0
         GoTo EXITSUB
      ElseIf Val(strNHI07) >= 10000000 Then
         strNHI08 = 10000000
      End If

      '執行業務所得判斷是否要扣補充保費
      If strNHI03 = "9A" Then
         stSQL = "select OI13,OI02 from OtherIncomer where OI01='" & strNHI01 & "'"
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            'Modified by Morgan 2017/1/23 '不是個人也不扣
            'If "" & rsQuery.Fields(0).Value = "N" Then  '不扣補充保費
            'Modified by Morgan 2017/6/26
            'If "" & rsQuery.Fields(0).Value = "N" Or Len("" & rsQuery(0)) <> 10 Then
            If "" & rsQuery.Fields("OI13").Value = "N" Or Len("" & rsQuery("OI02")) <> 10 Then
            'end 2017/1/23
               strNHI08 = 0
               GoTo EXITSUB
            End If
         End If
      End If
      
      'Added by Morgan 2013/1/22
      '租金個人才要扣
      'Modified by Morgan 2013/12/30 +股利 54
      If strNHI03 = "51" Or strNHI03 = "54" Then
         stSQL = "select OI02 from OtherIncomer where OI01='" & strNHI01 & "'"
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            If Len("" & rsQuery(0)) <> 10 Then
               strNHI08 = 0
               GoTo EXITSUB
            End If
         End If
      End If
      'end 2013/1/22

   
   Else   '薪資所得獎金/兼職所得
      
      '翻譯費
      If strNHI04 = "4" Then
         'Removed by Morgan 2013/2/18 移到下面,非F編號也要考慮F編號的獎金收入
         'If Left(strNHI01, 1) = "F" Then
         '   '內翻編號同時抓所內編號資料
         '   'Modified by Morgan 2013/1/29 改抓身分證號相同者考慮有可能是離職回籠抓排序大者
         '   'stSQL = "select SIM01 from Staff_IdMap where SIM02='" & strNHI01 & "'"
         '   stSQL = "select st01 from staff a where substr(st01,1,1)<>'F' and st26=(select b.st26 from staff b where b.st01='" & strNHI01 & "') order by st01 desc"
         '   intQ = 1
         '   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         '   If intQ = 1 Then strNo = "" & rsQuery.Fields(0).Value
         'End If
         
         '以員工檔部門判斷是薪資所得(內翻)或兼職所得(外翻,再以薪資基本檔判斷是否扣補充保費)
         'Modified by Morgan 2014/12/22 考慮留職停薪翻譯費仍用轉帳但健保會轉出所以改抓收文部門判斷是否為兼職所得 Ex.F5519 103/10,103/11
         'stSQL = "select ST03,SD49 from STAFF,SALARYDATA where ST01='" & strNHI01 & "' and sd01(+)=St01"
         stSQL = "select ST15,SD49 from STAFF,SALARYDATA where ST01='" & strNHI01 & "' and sd01(+)=St01"
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            If "" & rsQuery.Fields(0).Value = "F51" Then  '外翻為兼職所得,超過下限以上才扣,上限1千萬
               If "" & rsQuery.Fields(1).Value = "N" Then  '不扣補充保費
                  strNHI08 = 0
                  GoTo EXITSUB
               End If
               If Val(strNHI07) < lngLimit Then
                  strNHI08 = 0
                  GoTo EXITSUB
               Else
                  If Val(strNHI07) >= 10000000 Then
                     strNHI08 = 10000000
                  End If
                  GoTo CompStep
               End If
            Else
               '內翻併入獎金計算
            End If
         End If
         
      'Added by Morgan 2013/1/23
      '若有傳給付公司別要再判斷是否為兼職薪資所得
      ElseIf strPayCompany <> "" Then
         '本所同仁
         stSQL = "select sd47,sd19,sd49 from SALARYDATA where SD01='" & strNHI01 & "'"
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            'PUB_GetStaffNoByNumDate(strnhi01,strnhi02)
            '本所投保
            If rsQuery("sd47") > 0 Then
               If strNHI04 <> "1" Then 'Added by Morgan 2025/2/14 要排除年終獎金,因即使退休也要用獎金規則計算(四倍薪資) Ex:68010--婉莘
                  '給付公司非投保公司(兼職薪資)
                  'Modified by Morgan 2024/1/3 給付時已非在職員工也算兼職所得 Ex:B0002(1121211)
                  'If rsQuery("sd19") <> strPayCompany Then
                  If (rsQuery("sd19") <> strPayCompany Or PUB_GetStaffNoByNumDate(strNHI01, strNHI02) = "") Then
                     '要扣
                     If Val(strNHI07) < lngLimit Then
                        strNHI08 = 0
                     Else
                        If Val(strNHI07) >= 10000000 Then
                           strNHI08 = 10000000
                        End If
                        GoTo CompStep
                     End If
                     GoTo EXITSUB
                  Else
                     '獎金
                  End If
               End If
            '非本所投保
            Else
               '免扣
               If rsQuery("sd49") = "N" Then
                  strNHI08 = 0
               '要扣
               Else
                  If Val(strNHI07) < lngLimit Then
                     strNHI08 = 0
                  Else
                     If Val(strNHI07) >= 10000000 Then
                        strNHI08 = 10000000
                     End If
                     GoTo CompStep
                  End If
               End If
               GoTo EXITSUB
            End If
            
         '其他所得人
         Else
            stSQL = "select OI14 from OtherIncomer where OI01='" & strNHI01 & "'"
            intQ = 1
            Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
            If intQ = 1 Then
               '免扣
               If rsQuery("oi14") = "N" Then
                  strNHI08 = 0
               '要扣
               Else
                  If Val(strNHI07) < lngLimit Then
                     strNHI08 = 0
                  Else
                     If Val(strNHI07) >= 10000000 Then
                        strNHI08 = 10000000
                     End If
                     GoTo CompStep
                  End If
               End If
               
               GoTo EXITSUB
            End If
         End If
      'end 2013/1/22
      End If

      'Modified by Morgan 2013/2/25
      '所有相同id的獎金都要合併計算
      stSQL = "select st01 from staff a where st26=(select b.st26 from staff b where b.st01='" & strNHI01 & "') and st01<>'" & strNHI01 & "'"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         strNoList = "'" & rsQuery.Fields(0).Value & "'"
         rsQuery.MoveNext
         Do While Not rsQuery.EOF
            strNoList = strNoList & ",'" & rsQuery.Fields(0).Value & "'"
            rsQuery.MoveNext
         Loop
      End If
      'end 2013/2/25
      
      '抓當月投保金額,若為年終獎金則直接抓薪資基本檔,其他則先抓每月薪資檔若尚未產生則抓薪資基本檔
      If strNHI04 <> "1" Then
         'Modified by Morgan 2013/2/25
         '只會有1筆有投保薪資且所有編號都要考慮
         'stSQL = "select sum(sm42) from (" & _
                " select nvl(sm42,0) sm42 from salarymonth where sm01='" & strNHI01 & "' and sm02=" & Left(strNHI02, 6) & _
                " union select nvl(sm42,0) sm42 from salarymonth where sm01='" & strNo & "' and sm02=" & Left(strNHI02, 6) & ")"
         stSQL = "select nvl(sm42,0) sm42 from salarymonth where sm01='" & strNHI01 & "' and sm02=" & Left(strNHI02, 6) & " and sm42>0"
         If strNoList <> "" Then
            stSQL = stSQL & " union select nvl(sm42,0) sm42 from salarymonth where sm01 in (" & strNoList & ") and sm02=" & Left(strNHI02, 6) & " and sm42>0"
         End If
         'end 2013/2/25
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            If Val("" & rsQuery.Fields(0)) > 0 Then strNHI05 = Val("" & rsQuery.Fields(0))
         End If
      End If
      
      If Val(strNHI05) = 0 Then
         'Modified by Morgan 2013/2/25
         '考慮可能有復職情形抓最大號(多個編號者只會有1個有投保薪資)
         'stSQL = "select sum(sd47) from (" & _
                " select nvl(sd47,0) sd47 from salarydata where sd01='" & strNHI01 & "'" & _
                " union select nvl(sd47,0) sd47 from salarydata where sd01='" & strNo & "')"
         stSQL = " select nvl(sd47,0) sd47,sd01 from salarydata where sd01='" & strNHI01 & "' and sd47>0"
         If strNoList <> "" Then
            stSQL = stSQL & " union select nvl(sd47,0) sd47,sd01 from salarydata where sd01 in (" & strNoList & ")  and sd47>0"
         End If
         stSQL = stSQL & " order by sd01 desc"
         'end 2013/2/25
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            If Val("" & rsQuery.Fields(0)) > 0 Then
               strNHI05 = Val("" & rsQuery.Fields(0))
            End If
         End If
      End If

      '2015/2/2 ADD BY SONIA 退休人員隔年發放年終時,財務自行輸入年終資料,補充保費列入其他所得人(即NHI05=0)
      'Removed by Morgan 2025/2/14 即使退休也要用獎金規則計算(四倍薪資) Ex:68010--婉莘
      'If strNHI04 = "1" Then
      '   strST04 = "1"
      '   stSQL = "select st01,st04 from staff where st01='" & strNHI01 & "'"
      '   intQ = 1
      '   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      '   If intQ = 1 Then
      '      If "" & rsQuery.Fields(1).Value <> "1" Then strST04 = "" & rsQuery.Fields(1).Value
      '   End If
      '   If strST04 <> "1" Then strNHI05 = 0
      'End If
      'end 2025/2/14
      '2015/2/2 END
      
      If strPayCompany = "" Then Err.Raise vbObjectError + 513, "公司別參數 strPayCompany 不可空白，累計獎金計算失敗！" 'Added by Morgan 2014/7/25
   
      '計算累計獎金(但要剔除當筆)
      'Modified by Morgan 2013/1/23 +NHI10條件
      'Modified by Morgan 2014/5/1 +NHI11條件(轉公司不可累計)
      'Modified by Morgan 2014/7/25 改寫法,因用 union 當金額相同時會只算一次
      'stSQL = "SELECT SUM(NHI07) FROM (" & _
      '         "      SELECT NVL(NHI07,0) NHI07 FROM NHI2ND WHERE NHI01='" & strNHI01 & "' AND NHI03='50' AND SUBSTR(NHI02,1,4)=" & Val(Left(strNHI02, 4)) & " AND NHI02*1000000+NHI10<" & (Val(strNHI02) * 1000000 + Val(strNHI10)) & " and nhi11='" & strPayCompany & "'"
      ''所內編號也要合併計算
      ''Modified by Morgan 2013/2/25
      ''會有多個編號 Ex.林信昌
      ''If strNo <> "" Then
      ''   stSQL = stSQL & " UNION SELECT NVL(NHI07,0) NHI07 FROM NHI2ND WHERE NHI01='" & strNo & "' AND NHI03='50' AND SUBSTR(NHI02,1,4)=" & Val(Left(strNHI02, 4)) & " AND NHI02*1000000+NHI10<" & (Val(strNHI02) * 1000000 + Val(strNHI10))
      ''End If
      'If strNoList <> "" Then
      '   stSQL = stSQL & " UNION SELECT NVL(NHI07,0) NHI07 FROM NHI2ND WHERE NHI01 in (" & strNoList & ") AND NHI03='50' AND SUBSTR(NHI02,1,4)=" & Val(Left(strNHI02, 4)) & " AND NHI02*1000000+NHI10<" & (Val(strNHI02) * 1000000 + Val(strNHI10)) & " and nhi11='" & strPayCompany & "'"
      'End If
      ''end 2013/2/25
      'Modified by Morgan 2015/1/6 +判斷有投保薪資者,因留職停薪又復職時同年會有兼職及薪資所得
      stSQL = "SELECT SUM(NHI07) FROM NHI2ND WHERE NHI01 in ('" & strNHI01 & "'" & IIf(strNoList = "", "", "," & strNoList) & ") AND NHI03='50' AND SUBSTR(NHI02,1,4)=" & Val(Left(strNHI02, 4)) & " AND NHI02*1000000+NHI10<" & (Val(strNHI02) * 1000000 + Val(strNHI10)) & " and nhi11='" & strPayCompany & "' and nhi05>0"
      'end 2014/7/25
      
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         strNHI13 = Val("" & rsQuery.Fields(0)) + Val(strNHI07) 'Added by Morgan 2013/3/11
         If Val("" & rsQuery.Fields(0)) + Val(strNHI07) <= Val(strNHI05) * 4 Then          '未超過投保金額4倍不扣
            strNHI08 = 0
            GoTo EXITSUB
         Else
            strNHI08 = Val("" & rsQuery.Fields(0)) + Val(strNHI07) - Val(strNHI05) * 4     '累計超過4倍投保金額之獎金
            If Val(strNHI08) >= Val(strNHI07) Then    '若>=當次發放獎金時以當次發放獎金為補充保險費費基,否則以累計超過4倍投保金額之獎金為補充保險費費基
               strNHI08 = Val(strNHI07)
            End If
         End If
      Else
         strNHI13 = Val(strNHI07)  'Added by Morgan 2013/3/11
      End If

   End If

CompStep:
   '計算補充保費
   If Val(strNHI08) > 0 Then
      strNHI06 = Round(Val(strNHI08) * Val(strNHIRate) / 100)
   End If

EXITSUB:
   Set rsQuery = Nothing
   
End Sub

Public Sub PUB_InsertNHI2nd(ByRef strNHI() As String, Optional pIsPerson As Boolean = True)
   
   '檢查是否該所得人的最後一筆 ??? => 新增前呼叫 PUB_ChkNHi2nd memo by Morgan 2013/2/26
   'Modified by Morgan 2013/2/26 +nhi11
   
   '刪除當筆資料
   strSql = "DELETE NHI2ND WHERE NHI01='" & strNHI(1) & "' AND NHI02=" & Val(strNHI(2)) & " AND NHI03='" & strNHI(3) & "' AND NHI04='" & strNHI(4) & "' AND NHI10='" & strNHI(10) & "' AND NHI11='" & strNHI(11) & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   
   If pIsPerson Then 'Added by Morgan 2017/1/23 個人才要新增
      '新增補充保費明細及維護記錄檔
      'Modified by Morgan 2013/4/24 +NHI14
      strSql = "INSERT INTO NHI2ND (NHI01,NHI02,NHI03,NHI04,NHI05,NHI06,NHI07,NHI08,NHI10,NHI11,NHI13,NHI14) " & _
               "Values (" & CNULL(strNHI(1)) & "," & Val(strNHI(2)) & "," & CNULL(strNHI(3)) & "," & CNULL(strNHI(4)) & "," & Val(strNHI(5)) & "," & Val(strNHI(6)) & _
               "," & Val(strNHI(7)) & "," & Val(strNHI(8)) & "," & Val(strNHI(10)) & ",'" & strNHI(11) & "'," & Val(strNHI(13)) & "," & IIf(Val(strNHI(14)) > 0, Val(strNHI(14)), Val(strNHI(2))) & ") "
      
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
   
End Sub
'2013/1/17 end

'Added by Morgan 2013/1/24
'檢查是否有未申報之補充保費
'傳入: pbolAddBonus:是否含獎金(預設不含,因為主要是檢查外翻及所外所得人)
'回傳: pNotPaidFee:未申報補充保費
Public Function PUB_ChkNotPaidNhiFee(pNHI01 As String, Optional pNHI03 As String, Optional pNotPaidFee As String, Optional pbolNoBonus As Boolean = True) As Boolean
   Dim stSQL As String, intQ As Integer, stCon As String
   Dim rsQuery As ADODB.Recordset
   
   stCon = ""
   If pNHI03 <> "" Then stCon = stCon & " and nhi03='" & pNHI03 & "'"
   If pbolNoBonus = True Then stCon = stCon & " and nvl(nhi05,0)=0"
   
   stSQL = "select count(*) C1,nvl(sum(nhi06),0) C2 from nhi2nd where nhi01='" & pNHI01 & "' and nhi09 is null" & stCon
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      If rsQuery(0) > 0 Then
         pNotPaidFee = rsQuery(1)
         PUB_ChkNotPaidNhiFee = True
      End If
   End If
   Set rsQuery = Nothing
End Function

'Added by Morgan 2013/2/5
'檢查月薪資是否已計算
Public Function PUB_ExistsSalaryMonth(pDate As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   PUB_ExistsSalaryMonth = True
   
   stSQL = "select * from salarymonth where sm02=" & DBDATE(pDate) \ 100 & " and rownum<2"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 0 Then
      PUB_ExistsSalaryMonth = False
   End If
   Set rsQuery = Nothing
End Function

'Added by Morgan 2013/2/6
'檢查補充保費資料(若為員工獎金則不可有晚於該筆的資料)
'
Public Function PUB_ChkNHi2nd(pPayeeNo As String, pPayDate As String, pPayTime As String, Optional pbolAddRec As Boolean = False, Optional pbolInTrans As Boolean = True) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stStaffName As String
   'Dim stStaffNo As String 'Removed by Morgan 2014/7/25 考慮有復職員工且編號不同及一人有多個外譯編號故改寫法

   If Left(pPayeeNo, 1) = "F" Then
      'Modified by Morgan 2014/7/25 考慮有復職員工且編號不同及一人有多個外譯編號故改寫法
      'Modified by Morgan 2014/12/22 考慮留職停薪翻譯費仍用轉帳但健保會轉出所以改抓收文部門判斷是否為兼職所得 Ex.F5519 103/10,103/11
      'stSQL = "select a.st03,b.st01,a.st02 from staff a,staff b where a.st01='" & pPayeeNo & "' and b.st26(+)=a.st26 and b.st01(+)<'F' order by 2 desc"
      stSQL = "select a.st15,b.st01,a.st02 from staff a,staff b where a.st01='" & pPayeeNo & "' and b.st26(+)=a.st26 and b.st01(+)<'F' order by 2 desc"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         '外翻不必控制
         If rsQuery(0) = "F51" Then
            PUB_ChkNHi2nd = True
            GoTo EscPoint
         'Removed by Morgan 2014/7/25 考慮有復職員工且編號不同故改寫法
         'Else
         '   stStaffNo = "" & rsQuery(1)
         'end 2014/7/25
         End If
      End If
   
   'Added by Morgan 2013/2/18 非F編號也要考慮F編號的獎金收入
   'Removed by Morgan 2014/7/25 考慮有復職員工且編號不同及一人有多個外譯編號故改寫法
   'Else
   '   stSQL = "select a.st03,b.st01,a.st02 from staff a,staff b where a.st01='" & pPayeeNo & "' and b.st26(+)=a.st26 and b.st01>'F' order by 2 desc"
   '   intQ = 1
   '   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   '   If intQ = 1 Then
   '      stStaffNo = rsQuery(1)
   '   End If
   'end 2014/7/25
   End If
   
   'Modified by Morgan 2014/1/7 只需同年度的資料
   'Modified by Morgan 2014/7/25 考慮有復職員工且編號不同及一人有多個外譯編號故改寫法
   'stSQL = "select nhi01 from NHI2nd where nhi01='" & pPayeeNo & "' and nhi02<" & DBDATE((Val(pPayDate) \ 10000 + 1) & "0101") & " and nhi03='50' and nhi05>0 and 1000000*nhi02+nhi10" & IIf(pbolAddRec, ">=", ">") & (1000000 * Val(DBDATE(pPayDate)) + Val(pPayTime))
   'If stStaffNo <> "" Then
   '   stSQL = stSQL & "union select nhi01 from NHI2nd where nhi01='" & stStaffNo & "' and nhi02<" & DBDATE((Val(pPayDate) \ 10000 + 1) & "0101") & " and nhi03='50' and nhi05>0 and 1000000*nhi02+nhi10" & IIf(pbolAddRec, ">=", ">") & (1000000 * Val(DBDATE(pPayDate)) + Val(pPayTime))
   'End If
   'Modified by Morgan 2016/1/20 只需判斷同一年的資料
   'stSQL = "select nhi01,nhi02,nhi10 from staff s1,staff s2,NHI2nd where s1.st01='" & pPayeeNo & "' and s2.st26(+)=s1.st26 and nhi01(+)=s2.st01 and nhi03='50' and nhi05>0 and nhi02||lpad(nhi10,6,'0')>=" & Val(DBDATE(pPayDate)) & "||lpad(" & Val(pPayTime) & ",6,'0')" & IIf(pbolAddRec, "", " and not (nhi01='" & pPayeeNo & "' and nhi02=" & Val(DBDATE(pPayDate)) & " and nhi10=" & Val(pPayTime) & ")")
   stSQL = "select nhi01,nhi02,nhi10 from staff s1,staff s2,NHI2nd where s1.st01='" & pPayeeNo & "' and s2.st26(+)=s1.st26 and nhi01(+)=s2.st01 and nhi03='50' and nhi05>0 and substr(nhi02,1,4)='" & Left(DBDATE(pPayDate), 4) & "' and nhi02||lpad(nhi10,6,'0')>=" & Val(DBDATE(pPayDate)) & "||lpad(" & Val(pPayTime) & ",6,'0')" & IIf(pbolAddRec, "", " and not (nhi01='" & pPayeeNo & "' and nhi02=" & Val(DBDATE(pPayDate)) & " and nhi10=" & Val(pPayTime) & ")")
   'end 2014/7/25
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      '若在 Transaction 內要取消鎖定
      If pbolInTrans = True Then
         cnnConnection.RollbackTrans
      End If
      MsgBox pPayeeNo & "( 或相同ID的員工編號 ) 已存在給付時間晚(等)於本筆的獎金資料" & IIf(pbolInTrans, "，作業取消！", "！"), vbExclamation, "補充保費檢查"
   ElseIf intQ = 0 Then
      PUB_ChkNHi2nd = True
   End If
EscPoint:
   Set rsQuery = Nothing
End Function

'Added by Morgan 2013/2/26
'讀取薪資公司別(先抓月薪資再抓基本薪資)
Public Function GetSalaryCompany(pNo As String, pDate As String) As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   'Modified by Morgan 2014/7/25 補排序 order by srt
   stSQL = "select 1 srt,sm37 from salarymonth where sm01='" & pNo & "' and sm02=" & Left(DBDATE(pDate), 6) & _
      " union all select 2 srt,sd19 from salarydata where sd01='" & pNo & "' order by srt"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      GetSalaryCompany = "" & rsQuery(1)
   End If

   Set rsQuery = Nothing
End Function

'Added by Morgan 2013/3/7
Public Function ChkHi2ndIsPaid(ByVal pYM As String, Optional ByRef pPayDate As String, Optional pComp As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stCon As String
   
   If pComp <> "" Then
      stCon = " and NHI11='" & pComp & "'"
   End If
   
   stSQL = "select nhi12 from nhi2nd where nhi02>=" & pYM & "01 and nhi02<=" & pYM & "31 and nhi12>0" & stCon & " and rownuM<2"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      pPayDate = rsQuery(0)
      ChkHi2ndIsPaid = True
   End If
   Set rsQuery = Nothing
End Function

'Add by Sindy 2013/7/3 檢查工作時數是否符合規定
Public Function ChkTaieWorkingHour(dblMinTime As Double, dblMaxTime As Double, strWorkStar As String, strWorkEnd As String) As Boolean
Dim ii As Integer
Dim dblChkStarTime As Double

   ChkTaieWorkingHour = False
   If Val(dblMinTime) = 0 Or Val(dblMaxTime) = 0 Then Exit Function
   'Modify By Sindy 2015/5/5 特殊上班時段的人員,只要檢查是否有在符合的上下班時段裡就好,不用考慮彈性時段的問題
   'If strWorkStar = "09:00" Then
   If PUB_bWkSpec = False Then
   '2015/5/5 END
      'If (Val(dblMinTime) < 900 And Val(dblMinTime) <> 0) Then
         'Modify By Sindy 2021/5/17
         For ii = 1 To intByPassArea
            If Right(strByPassStarTime(ii), 2) = "00" Then
               dblChkStarTime = Val(Left(strByPassStarTime(ii), 2) & "59") - 100
            Else
               dblChkStarTime = Val(Format(strByPassStarTime(ii), "HHMM")) - 1
            End If
            If ((Val(dblMaxTime) >= Val(Format(strByPassEndTime(ii), "HHMM")) And _
                 Val(dblMaxTime) <= Val(Format(strByPassEndTime(ii), "HHMM")) + 29) And _
                 Val(dblMinTime) <= dblChkStarTime) = True Then
               ChkTaieWorkingHour = True
            End If
         Next ii
         If (Val(dblMaxTime) >= Val(Format(strByPassEndTime(intByPassArea), "HHMM")) And _
             Val(dblMinTime) <= dblChkStarTime) = True Then
            ChkTaieWorkingHour = True
         End If
         '2021/5/17 END
'         If ((Val(dblMaxTime) >= 1700 And Val(dblMaxTime) <= 1729) And Val(dblMinTime) <= 759) = True Or _
'            ((Val(dblMaxTime) >= 1730 And Val(dblMaxTime) <= 1759) And Val(dblMinTime) <= 829) = True Or _
'             (Val(dblMaxTime) >= 1800 And Val(dblMinTime) <= 859) = True Then
'            ChkTaieWorkingHour = True
'         End If
      'End If
   Else
      If (Val(dblMaxTime) >= Val(Format(strWorkEnd, "hhmm")) And _
          Val(dblMinTime) <= Val(Format(strWorkStar, "hhmm"))) = True Then
         ChkTaieWorkingHour = True
      End If
   End If
End Function

'Add by Sindy 2013/7/3 打卡異常E-Mail通知
Public Sub StaffCardErrSendMail(strB1401 As String, strB1402 As String, strB1403 As String, strB1404 As String, chkMinTime As Double, chkMaxTime As Double, min_pr02 As String, max_pr02 As String, Optional strSendPerson As String = "")
Dim strSubject As String, strContent As String
Dim strST59 As String, strTo As String
   
   strB1404 = Replace(strB1404, ":", "")
   If strB1403 = "A" Then
      strSubject = ChangeWStringToTDateString(DBDATE(strB1402)) & " 上班打卡異常通知"
   Else
      strSubject = ChangeWStringToTDateString(DBDATE(strB1402)) & " 下班打卡異常通知"
   End If
   strContent = "打卡日期：" & strSubject & vbCrLf
   If strB1403 = "A" Then
      strContent = strContent & "打卡時間：" & IIf(strB1404 = "", "未打卡", Format(strB1404, "00:00:00")) & vbCrLf
   Else
      strContent = strContent & "最早打卡時間：" & IIf(min_pr02 = "", "上班未打卡", Format(min_pr02, "00:00:00")) & vbCrLf
      strContent = strContent & "最後打卡時間：" & IIf(strB1404 = "", "下班未打卡", Format(max_pr02, "00:00:00")) & vbCrLf
      If strB1403 = "P" And chkMaxTime >= 1700 Then
         strContent = strContent & "（上下班時段不符規定：上班" & Format(chkMinTime, "00:00") & "下班" & Format(chkMaxTime, "00:00") & "）" & vbCrLf
      End If
   End If
   strContent = strContent & "處　　理：請至案件管理系統（一般作業->出缺勤作業->表單->打卡異常個人處理）中，進行處理。" & vbCrLf
   strContent = strContent & "員工代號：" & strB1401 & vbCrLf
   
   strST59 = PUB_GetST59(strB1401)
   If Not IsNull(strST59) And strST59 <> "" Then
      strTo = strST59
   Else
      strTo = strB1401
   End If
   'Add By Sindy 2013/9/18 檢查收件人是否為不寄信者
   If Len(strTo) = 5 Then
      If ChkStaffST14(strTo, False) = True Then
         Exit Sub
      End If
   End If
   '2013/9/18 END
   If strSendPerson = "" Then
      strContent = strContent & vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***"
      If Val(DBDATE(strB1402)) >= 20130801 Then
         PUB_SendMail "administrator", strTo, "", strSubject, strContent, , , , , , , "administrator", "系統管理員", , True
      End If
   Else
      If Val(DBDATE(strB1402)) >= 20130801 Then
         PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True
      End If
   End If
End Sub

'Add By Sindy 2013/7/16
'取得員工代收郵件人員2
Public Function PUB_GetST59(ByVal p_User As String) As String
   strSql = "select st59 from staff where st01='" & p_User & "'"
   
On Error GoTo ErrHnd

   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         PUB_GetST59 = "" & .Fields(0)
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Add By Sindy 2013/7/24 查詢請假、出差、加班資料
Public Function PUB_QueryData_ABS(strB1401 As String, strB1402 As String, ByRef rsTmp As ADODB.Recordset) As Boolean
'Dim rstmp As New ADODB.Recordset
Dim strSql As String, strCon As String
Dim strConABS As String, strConSA As String, strConSO As String, strConSB As String
'Modify By Sindy 2023/12/29
Dim strA09TaNm1 As String
Dim strA09Pkey1 As String
Dim strA09TaNm2 As String
Dim strA09Pkey2 As String
'2023/12/29 END
   
   PUB_QueryData_ABS = False
   
   If strB1401 = "" Or strB1402 = "" Then Exit Function
   
'   grd2.Clear
'   SetGrd2
   strCon = "": strConABS = "": strConSA = "": strConSO = "": strConSB = "":
   '日期
   If strB1402 <> "" Then
      strConABS = strConABS & " and (" & DBDATE(strB1402) & " between B1004 and B1006 or " & _
                              DBDATE(strB1402) & " between B1004 and B1006 or " & _
                              "B1004 between " & DBDATE(strB1402) & " and " & DBDATE(strB1402) & " or " & _
                              "B1006 between " & DBDATE(strB1402) & " and " & DBDATE(strB1402) & ") "
      strConSA = strConSA & " and (" & DBDATE(strB1402) & " between SA02 and SA04 or " & _
                              DBDATE(strB1402) & " between SA02 and SA04 or " & _
                              "SA02 between " & DBDATE(strB1402) & " and " & DBDATE(strB1402) & " or " & _
                              "SA04 between " & DBDATE(strB1402) & " and " & DBDATE(strB1402) & ") "
      strConSO = strConSO & " and (So02 between " & DBDATE(strB1402) & " and " & DBDATE(strB1402) & ") "
      strConSB = strConSB & " and (" & DBDATE(strB1402) & " between SB02 and SB04 or " & _
                              DBDATE(strB1402) & " between SB02 and SB04 or " & _
                              "SB02 between " & DBDATE(strB1402) & " and " & DBDATE(strB1402) & " or " & _
                              "SB04 between " & DBDATE(strB1402) & " and " & DBDATE(strB1402) & ") "
   End If
   '員工代號
   If strB1401 <> "" Then
      strCon = strCon & " and s1.ST01=" & CNULL(strB1401)
   End If
   Screen.MousePointer = vbHourglass
   'Modify By Sindy 2023/12/29
   If strSrvDate(1) >= 新部門啟用日 Then
      strA09TaNm1 = "ACC090NEW a1"
      strA09Pkey1 = "s1.ST93=a1.A0921"
      strA09TaNm2 = "ACC090NEW a2"
      strA09Pkey2 = "a2.A0921"
   Else
      strA09TaNm1 = "ACC090 a1"
      strA09Pkey1 = "s1.ST03=a1.A0901"
      strA09TaNm2 = "ACC090 a2"
      strA09Pkey2 = "a2.A0901"
   End If
   '出缺勤電子簽核主檔(人事處未簽收或註銷的表單),員工請假資料,員工加班資料,員工出差資料
   strSql = "Select 'V' as V,s1.ST01 員工代號,B1001 表單編號,'1' TableID,B1004,B1005 " & _
            "From ABS010,Staff s1," & strA09TaNm1 & ",allcode,Staff s2," & strA09TaNm2 & _
            " Where B1003=s1.ST01(+) and " & strA09Pkey1 & "(+) and B1017=" & strA09Pkey2 & "(+) and ac01(+)='04' and B1008=ac02(+) and B1017=s2.ST01(+) and (B1019 is null or B1018='" & 註銷 & "') " & strCon & strConABS
   strSql = strSql & " union " & _
            "Select 'V' as V,s1.ST01 員工代號,SA09 表單編號,'2' TableID,SA02,SA03 " & _
            "From Staff_Absence,Staff s1," & strA09TaNm1 & ",allcode " & _
            "Where SA01=s1.ST01(+) and " & strA09Pkey1 & "(+) and ac01(+)='04' and SA06=ac02(+) and (SA09 is null or (SA09 in(select B1001 from abs010 where B1002='01' and B1003=SA01 and B1019 is not null))) " & strCon & strConSA
   strSql = strSql & " union " & _
            "Select 'V' as V,s1.ST01 員工代號,SO13 表單編號,'3' TableID,So02,So03 " & _
            "From Staff_Overtime,Staff s1," & strA09TaNm1 & " " & _
            "Where SO01=s1.ST01(+) and " & strA09Pkey1 & "(+) and (SO13 is null or (SO13 in(select B1001 from abs010 where B1002='02' and B1003=SO01 and B1019 is not null))) " & strCon & strConSO
   strSql = strSql & " union " & _
            "Select 'V' as V,s1.ST01 員工代號,SB10 表單編號,'4' TableID,SB02,SB03 " & _
            "From Staff_Busi_Trip,Staff s1," & strA09TaNm1 & " " & _
            "Where SB01=s1.ST01(+) and " & strA09Pkey1 & "(+) and (SB10 is null or (SB10 in(select B1001 from abs010 where B1002='03' and B1003=SB01 and B1019 is not null))) " & strCon & strConSB
   strSql = strSql & " order by TableID asc "
'   Set rstmp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      'Set grd2.Recordset = rsTmp
'      PUB_QueryData_ABS = True
'   Else
'      ShowNoData
'      rstmp.Close
'      Set rstmp = Nothing
'      Screen.MousePointer = vbDefault
'      Exit Function
'   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Screen.MousePointer = vbDefault
   If rsTmp.RecordCount > 0 Then
'      Set grd2.Recordset = rsTmp
      PUB_QueryData_ABS = True
   Else
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
   
'   rstmp.Close
'   Set rstmp = Nothing
'   Screen.MousePointer = vbDefault
'   Call PubShowNextData
End Function

'Modify By Sindy 2021/5/17
'檢查是否分流上班,上班時段
Public Function PUB_ChkByPassWork(ByVal strST06 As String, ByVal strDay As String, _
Optional ByVal strMinPr02 As String, Optional ByVal strMaxPr02 As String, _
Optional ByRef strStarWorkTime As String, Optional ByRef strEndWorkTime As String) As Boolean

Dim ii As Integer

   PUB_ChkByPassWork = False
   strDay = DBDATE(strDay)
   
   If Len(strMinPr02) <= 4 And strMinPr02 <> "" Then strMinPr02 = strMinPr02 & "00"
   If Len(strMaxPr02) <= 4 And strMaxPr02 <> "" Then strMaxPr02 = strMaxPr02 & "00"
   intByPassArea = Val(Pub_GetSpecMan("彈性上下班時段")) 'Add By Sindy 2021/11/23
   
   'strST06 = "1"'北所才適用
   'Modify By Sindy 2021/5/19 全所均分流 mark:And strST06 = "1"
   'Modify By Sindy 2021/11/23 + Or intByPassArea = 6
   If (Val(strDay) >= 分流上班起始日期 And Val(strDay) <= 分流上班截止日期) Or _
      intByPassArea = 6 Then
      PUB_ChkByPassWork = True
      
      intByPassArea = 6
      strByPassStarTime(1) = "07:30"
      strByPassEndTime(1) = "16:30"
      strByPassStarTime(2) = "08:00"
      strByPassEndTime(2) = "17:00"
      strByPassStarTime(3) = "08:30"
      strByPassEndTime(3) = "17:30"
      strByPassStarTime(4) = "09:00"
      strByPassEndTime(4) = "18:00"
      strByPassStarTime(5) = "09:30"
      strByPassEndTime(5) = "18:30"
      strByPassStarTime(6) = "10:00"
      strByPassEndTime(6) = "19:00"
      
      If Val(strMinPr02) <= 100000 And strMinPr02 <> "" Then
         If Val(strMinPr02) <= 73000 Then
            strStarWorkTime = "07:30"
            strEndWorkTime = "16:30"
         ElseIf Val(strMinPr02) <= 80000 Then
            strStarWorkTime = "08:00"
            strEndWorkTime = "17:00"
         ElseIf Val(strMinPr02) <= 83000 Then
            strStarWorkTime = "08:30"
            strEndWorkTime = "17:30"
         ElseIf Val(strMinPr02) <= 90000 Then
            strStarWorkTime = "09:00"
            strEndWorkTime = "18:00"
         ElseIf Val(strMinPr02) <= 93000 Then
            strStarWorkTime = "09:30"
            strEndWorkTime = "18:30"
         Else
            strStarWorkTime = "10:00"
            strEndWorkTime = "19:00"
         End If
      ElseIf Val(strMaxPr02) >= 163000 Then
         If Val(strMaxPr02) >= 190000 Then
            strStarWorkTime = "10:00"
            strEndWorkTime = "19:00"
         ElseIf Val(strMaxPr02) >= 183000 Then
            strStarWorkTime = "09:30"
            strEndWorkTime = "18:30"
         ElseIf Val(strMaxPr02) >= 180000 Then
            strStarWorkTime = "09:00"
            strEndWorkTime = "18:00"
         ElseIf Val(strMaxPr02) >= 173000 Then
            strStarWorkTime = "08:30"
            strEndWorkTime = "17:30"
         ElseIf Val(strMaxPr02) >= 170000 Then
            strStarWorkTime = "08:00"
            strEndWorkTime = "17:00"
         Else
            strStarWorkTime = "07:30"
            strEndWorkTime = "16:30"
         End If
      End If
   
   Else
      '預設值
      intByPassArea = 3
      strByPassStarTime(1) = "08:00"
      strByPassEndTime(1) = "17:00"
      strByPassStarTime(2) = "08:30"
      strByPassEndTime(2) = "17:30"
      strByPassStarTime(3) = "09:00"
      strByPassEndTime(3) = "18:00"
      
      If Val(strMinPr02) <= 90000 And strMinPr02 <> "" Then
         If Val(strMinPr02) <= 80000 Then
            strStarWorkTime = "08:00"
            strEndWorkTime = "17:00"
         ElseIf Val(strMinPr02) <= 83000 Then
            strStarWorkTime = "08:30"
            strEndWorkTime = "17:30"
         Else
            strStarWorkTime = "09:00"
            strEndWorkTime = "18:00"
         End If
      ElseIf Val(strMaxPr02) >= 170000 Then
         If Val(strMaxPr02) >= 180000 Then
            strStarWorkTime = "09:00"
            strEndWorkTime = "18:00"
         ElseIf Val(strMaxPr02) >= 173000 Then
            strStarWorkTime = "08:30"
            strEndWorkTime = "17:30"
         Else
            strStarWorkTime = "08:00"
            strEndWorkTime = "17:00"
         End If
      End If
   End If
   
   '清空其他變數值
   For ii = intByPassArea + 1 To 8
      strByPassStarTime(ii) = ""
      strByPassEndTime(ii) = ""
   Next ii
End Function

'Add By Sindy 2013/9/11 系統自動確認打卡異常資料
Public Function PUB_UpdateB14Data(StrST01 As String, Optional strB1402Date As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim bolRest1Day As Boolean, strRestKind As String
Dim strUpdDate As String, strUpdTime As String
Dim strPollData As String
Dim strStarWorkTime As String, strEndWorkTime As String
Dim strMinPr02 As String, strMaxPr02 As String
Dim strWOnTime As String, strWOffTime As String 'Add By Sindy 2016/8/30
Dim bolChk As Boolean, ii As Integer 'Add By Sindy 2021/5/17
   
   PUB_UpdateB14Data = False 'Add By Sindy 2016/8/9
   strUpdDate = Format(Now, "YYYYMMDD")
   strUpdTime = Format(time, "HHMMSS")
'   '請假
'   strSql = "update ABS014 set B1405='1'" & _
'                             ",B1411='A'" & _
'                             ",B1412=" & strUpdDate & _
'                             ",B1413=" & strUpdTime & _
'            " where b1401='" & strST01 & "' and b1404 between 121001 and 132959" & _
'              " and b1401||b1402 in(select sa01||sa02 from staff_absence Where sa01=b1401 and sa02=b1402 and sa03=1330" & _
'                            " Union select sa01||sa04 from staff_absence Where sa01=b1401 and sa04=b1402 and sa05=1210)" & _
'            " and b1411 is null"
'   cnnConnection.Execute strSql
'   '出差
'   strSql = "update ABS014 set B1405='6'" & _
'                             ",B1411='A'" & _
'                             ",B1412=" & strUpdDate & _
'                             ",B1413=" & strUpdTime & _
'            " where b1401='" & strST01 & "' and b1404 between 121001 and 132959" & _
'              " and b1401||b1402 in(select sb01||sb02 from staff_busi_trip Where sb01=b1401 and sb02=b1402 and sb03=1330" & _
'                             "Union select sb01||sb04 from staff_busi_trip Where sb01=b1401 and sb04=b1402 and sb05=1210)" & _
'            " and b1411 is null"
'   cnnConnection.Execute strSql
   
   '逐筆檢查是否有系統需要直接確認的
   strSql = "select * from ABS014 where b1411 is null and b1401='" & StrST01 & "'" & IIf(strB1402Date = "", "", " and b1402=" & strB1402Date & "")
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         '取得當天的打卡資料
         strExc(0) = "select scd01,pr01,nvl(min(pr02),0) as min_pr02,nvl(max(pr02),0) as max_pr02 from pollrecord,staffcarddata where pr03=scd02(+) and pr01=" & rsTmp.Fields("b1402") & " and scd01='" & rsTmp.Fields("b1401") & "' group by scd01,pr01"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         strMinPr02 = ""
         strMaxPr02 = ""
         If intI = 1 Then
            strMinPr02 = RsTemp.Fields("min_pr02")
            strMaxPr02 = RsTemp.Fields("max_pr02")
         End If
         '抓出應該工作的上下班時段
         Call Pub_GetSpecWorkHour(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), strStarWorkTime, strEndWorkTime)
         strStarWorkTime = Format(strStarWorkTime, "00:00")
         strEndWorkTime = Format(strEndWorkTime, "00:00")
         If strStarWorkTime = "" And PUB_bWkSpec = False And strMinPr02 <> "" Then
            'Modify By Sindy 2021/5/17
            m_bolByPassWork = PUB_ChkByPassWork(PUB_GetST06(rsTmp.Fields("b1401")), strB1402Date, strMinPr02, strMaxPr02, strStarWorkTime, strEndWorkTime)
            '2021/5/17 END
'            If Val(strMinPr02) <= 90000 Then
'               If Val(strMinPr02) < 80000 Then
'                  strStarWorkTime = "08:00"
'                  strEndWorkTime = "17:00"
'               ElseIf Val(strMinPr02) < 83000 Then
'                  strStarWorkTime = "08:30"
'                  strEndWorkTime = "17:30"
'               Else
'                  strStarWorkTime = "09:00"
'                  strEndWorkTime = "18:00"
'               End If
'            ElseIf Val(strMaxPr02) >= 170000 Then
'               If Val(strMaxPr02) >= 180000 Then
'                  strStarWorkTime = "09:00"
'                  strEndWorkTime = "18:00"
'               ElseIf Val(strMaxPr02) >= 173000 Then
'                  strStarWorkTime = "08:30"
'                  strEndWorkTime = "17:30"
'               Else
'                  strStarWorkTime = "08:00"
'                  strEndWorkTime = "17:00"
'               End If
'            End If
         End If
         
'         'Test使用
'         If rsTmp.Fields("b1401") = "96027" Then
'            MsgBox rsTmp.Fields("b1401")
'         End If
'         'Test END
         
         'Modify By Sindy 2013/11/8 增加非整日時請假時間判斷
         '有打卡時間
         If Val("" & rsTmp.Fields("b1404")) > 0 Then
            If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), Left(Format(rsTmp.Fields("b1404"), "00:00:00"), 5), strRestKind, bolRest1Day, True) = True Then
               'Add By Sindy 2015/8/10 ex.91010曾維揚104/8/7下班異常誤判為請假
               'Modify By Sindy 2016/8/30 + , , , strWOnTime, strWOffTime
               '上班異常
               If rsTmp.Fields("b1403") = "A" Then
                  If strStarWorkTime <> "" Then
                     If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), strStarWorkTime, strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = False Then GoTo RunEnd
                  Else
                     'Modify By Sindy 2021/5/17
                     bolChk = False
                     For ii = 1 To intByPassArea
                        bolChk = CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), strByPassStarTime(ii), strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime)
                        If bolChk = True Then Exit For
                     Next ii
                     If bolChk = False Then
                        GoTo RunEnd
                     End If
                     '2021/5/17 END
'                     If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "08:00", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = False And _
'                        CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "08:30", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = False And _
'                        CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "09:00", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = False Then
'                        GoTo RunEnd
'                     End If
                  End If
               '下班異常
               Else
                  If strEndWorkTime <> "" Then
                     If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), strEndWorkTime, strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = False Then GoTo RunEnd
                  Else
                     'Modify By Sindy 2021/5/17
                     bolChk = False
                     For ii = 1 To intByPassArea
                        bolChk = CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), strByPassEndTime(ii), strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime)
                        If bolChk = True Then Exit For
                     Next ii
                     If bolChk = False Then
                        GoTo RunEnd
                     End If
                     '2021/5/17 END
'                     If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "17:00", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = False And _
'                        CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "17:30", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = False And _
'                        CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "18:00", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = False Then
'                        GoTo RunEnd
'                     End If
                  End If
               End If
               '2015/8/10 END
               'If Val(Replace(strStarWorkTime, ":", "")) = Val(strWOnTime) Then 'Add By Sindy 2016/8/30 +if
               'Modify By Sindy 2018/8/16 + Val(strStarWorkTime) <> 0 And Val(strWOnTime) <> 0
               'Modify By Sindy 2019/3/20
               If PUB_ChkWorkDtTiUpdABS("A", rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), rsTmp.Fields("b1403"), _
                  strRestKind, strWOnTime, strWOffTime, strStarWorkTime, strEndWorkTime, _
                  strUpdDate, strUpdTime) = True Then GoTo RunEnd
               '2019/3/20 END
            End If
            'Add By Sindy 2013/9/30
            'Modify By Sindy 2017/4/18
'            If strSrvDate(1) >= 中午休息時間改1200 Then
               '中午打上班卡
               If "" & rsTmp.Fields("b1403") = "A" And _
                  Val(Val("" & rsTmp.Fields("b1404"))) > 120000 And Val(Val("" & rsTmp.Fields("b1404"))) < 133000 Then
                  If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "12:00", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = True Then
                     'Modify By Sindy 2019/3/20
                     If PUB_ChkWorkDtTiUpdABS("A", rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), rsTmp.Fields("b1403"), _
                        strRestKind, strWOnTime, strWOffTime, strStarWorkTime, strEndWorkTime, _
                        strUpdDate, strUpdTime) = True Then GoTo RunEnd
                     '2019/3/20 END
                  End If
               End If
               '中午打下班卡
               If "" & rsTmp.Fields("b1403") = "P" And _
                  Val(Val("" & rsTmp.Fields("b1404"))) > 120000 And Val(Val("" & rsTmp.Fields("b1404"))) < 133000 Then
                  If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "13:30", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = True Then
                     'Modify By Sindy 2019/3/20
                     If PUB_ChkWorkDtTiUpdABS("P", rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), rsTmp.Fields("b1403"), _
                        strRestKind, strWOnTime, strWOffTime, strStarWorkTime, strEndWorkTime, _
                        strUpdDate, strUpdTime) = True Then GoTo RunEnd
                     '2019/3/20 END
                  End If
               End If
               'Add By Sindy 2018/8/16
               '當打卡時間有落在假單時間中，系統自動核銷
               If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), Left(Format(Val("" & rsTmp.Fields("b1404")), "00:00:00"), 5), strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = True Then
                  'If Val(Replace(strEndWorkTime, ":", "")) = Val(strWOffTime) Then 'Add By Sindy 2016/8/30 +if
                  'If Val(Replace(strEndWorkTime, ":", "")) >= Val(strWOffTime) Then 'Add By Sindy 2016/9/12 +if
                  'If Val(Replace(strStarWorkTime, ":", "")) <= Val(strWOnTime) Then 'Add By Sindy 2016/11/16 +if 寫反判斷,用上班時間檢查
                  'Modify By Sindy 2019/3/20
                  If PUB_ChkWorkDtTiUpdABS("A", rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), rsTmp.Fields("b1403"), _
                     strRestKind, strWOnTime, strWOffTime, strStarWorkTime, strEndWorkTime, _
                     strUpdDate, strUpdTime) = True Then GoTo RunEnd
                  '2019/3/20 END
               End If
'            Else
'            '2017/4/18 END
'               '中午打上班卡
'               If "" & rsTmp.Fields("b1403") = "A" And _
'                  Val(Val("" & rsTmp.Fields("b1404"))) > 121000 And Val(Val("" & rsTmp.Fields("b1404"))) < 133000 Then
'                  'Modify By Sindy 2016/8/30 + , , , strWOnTime, strWOffTime
'                  If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "12:10", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = True Then
'                     'If Val(Replace(strStarWorkTime, ":", "")) = Val(strWOnTime) Then 'Add By Sindy 2016/8/30 +if
'                     'Modify By Sindy 2019/3/20
'                     If PUB_ChkWorkDtTiUpdABS("A", rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), rsTmp.Fields("b1403"), _
'                        strRestKind, strWOnTime, strWOffTime, strStarWorkTime, strEndWorkTime, _
'                        strUpdDate, strUpdTime) = True Then GoTo RunEnd
'                     '2019/3/20 END
'                  End If
'               End If
'               '中午打下班卡
'               If "" & rsTmp.Fields("b1403") = "P" And _
'                  Val(Val("" & rsTmp.Fields("b1404"))) > 121000 And Val(Val("" & rsTmp.Fields("b1404"))) < 133000 Then
'                  'Modify By Sindy 2016/8/30 + , , , strWOnTime, strWOffTime
'                  If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "13:30", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = True Then
'                     'If Val(Replace(strEndWorkTime, ":", "")) = Val(strWOffTime) Then 'Add By Sindy 2016/8/30 +if
'                     'Modify By Sindy 2019/3/20
'                     If PUB_ChkWorkDtTiUpdABS("P", rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), rsTmp.Fields("b1403"), _
'                        strRestKind, strWOnTime, strWOffTime, strStarWorkTime, strEndWorkTime, _
'                        strUpdDate, strUpdTime) = True Then GoTo RunEnd
'                     '2019/3/20 END
'                  End If
'               End If
'            End If
            '2013/9/30 END
         
         '未打卡--異常
         Else
            '檢查是否整日請假的（以下午二點做檢查,bolRest1Day = True則為整日休）
            If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "14:00", strRestKind, bolRest1Day, True) = True Then
               If bolRest1Day = True Then
                  strSql = "update ABS014 set B1405='" & IIf(strRestKind = "3", "6", "1") & "'" & _
                                            ",B1411='A'" & _
                                            ",B1412=" & strUpdDate & _
                                            ",B1413=" & strUpdTime & _
                           " where b1401='" & rsTmp.Fields("b1401") & "'" & _
                             " and b1402=" & rsTmp.Fields("b1402") & _
                             " and b1403='" & rsTmp.Fields("b1403") & "'"
                  cnnConnection.Execute strSql
                  PUB_UpdateB14Data = True 'Add By Sindy 2016/8/9
                  GoTo RunEnd
               End If
            End If
            '************************************************************************************
            'Modify By Sindy 2013/9/30 劉經理:因是忘打卡所以系統不可自動核銷
            '(ex.96027林佳芳/1020917 上班異常被系統自動確認掉,但其實是忘打卡)
            '************************************************************************************
            'Modify by Sindy 2013/10/18 To:劉柏翰(人事經理)/Cc:林靖蓉(人事處) 提恢復
            '增加控管填寫假單時,起迄時間不可以輸入中午休息時段,起始時間13:30及迄止時間12:10除外
            '************************************************************************************
            '取得當天的打卡資料
            strExc(0) = "select scd01,pr01,nvl(min(pr02),0) as min_pr02,nvl(max(pr02),0) as max_pr02 from pollrecord,staffcarddata where pr03=scd02(+) and pr01=" & rsTmp.Fields("b1402") & " and scd01='" & rsTmp.Fields("b1401") & "' group by scd01,pr01"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            strMinPr02 = ""
'            strMaxPr02 = ""
            If intI = 1 Then
'               strMinPr02 = RsTemp.Fields("min_pr02")
'               strMaxPr02 = RsTemp.Fields("max_pr02")
               strPollData = ""
               '上班異常
               'Modify By Sindy 2016/8/30 + , , , strWOnTime, strWOffTime
               If rsTmp.Fields("b1403") = "A" Then
                   '上午時段
                  If Val("" & RsTemp.Fields("min_pr02")) > 0 Then
                     'Modify By Sindy 2018/8/16 靖蓉:因是忘打卡所以系統不可自動核銷
                     '                          (ex.85036陳金妙/1070816 上班異常被系統自動確認掉,但其實是忘打卡)
                     If Val("" & RsTemp.Fields("min_pr02")) < 120000 And _
                        Val("" & RsTemp.Fields("min_pr02")) <> Val("" & RsTemp.Fields("max_pr02")) Then
                     '2018/8/16 END
                        strPollData = RsTemp.Fields("min_pr02")
                     End If
                  End If
                  '中午打上班卡
                  'Modify By Sindy 2017/4/18
'                  If strSrvDate(1) >= 中午休息時間改1200 Then
                     If Val(strPollData) > 120000 And Val(strPollData) < 133000 Then
                        If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "12:00", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = True Then
                           'Modify By Sindy 2019/3/20
                           If PUB_ChkWorkDtTiUpdABS("A", rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), rsTmp.Fields("b1403"), _
                              strRestKind, strWOnTime, strWOffTime, strStarWorkTime, strEndWorkTime, _
                              strUpdDate, strUpdTime) = True Then GoTo RunEnd
                           '2019/3/20 END
                        End If
                     End If
'                  Else
'                  '2017/4/18 END
'                     If Val(strPollData) > 121000 And Val(strPollData) < 133000 Then
'                        If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "12:10", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = True Then
'                           'If Val(Replace(strStarWorkTime, ":", "")) = Val(strWOnTime) Then 'Add By Sindy 2016/8/30 +if
'                           'Modify By Sindy 2019/3/20
'                           If PUB_ChkWorkDtTiUpdABS("A", rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), rsTmp.Fields("b1403"), _
'                              strRestKind, strWOnTime, strWOffTime, strStarWorkTime, strEndWorkTime, _
'                              strUpdDate, strUpdTime) = True Then GoTo RunEnd
'                           '2019/3/20 END
'                        End If
'                     End If
'                  End If
               '下班異常
               Else
                  If Val("" & RsTemp.Fields("max_pr02")) > 0 Then
                     'Modify By Sindy 2018/8/16 靖蓉:因是忘打卡所以系統不可自動核銷
                     '                          (ex.85036陳金妙/1070816 上班異常被系統自動確認掉,但其實是忘打卡)
                     If Val("" & RsTemp.Fields("max_pr02")) >= 120000 And _
                        Val("" & RsTemp.Fields("min_pr02")) <> Val("" & RsTemp.Fields("max_pr02")) Then
                     '2018/8/16 END
                        strPollData = RsTemp.Fields("max_pr02")
                     End If
                  End If
                  'Modify By Sindy 2017/4/18
'                  If strSrvDate(1) >= 中午休息時間改1200 Then
                     '中午打下班卡
                     If Val(strPollData) > 120000 And Val(strPollData) < 133000 Then
                        If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "13:30", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = True Then
                           'Modify By Sindy 2019/3/20
                           If PUB_ChkWorkDtTiUpdABS("P", rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), rsTmp.Fields("b1403"), _
                              strRestKind, strWOnTime, strWOffTime, strStarWorkTime, strEndWorkTime, _
                              strUpdDate, strUpdTime) = True Then GoTo RunEnd
                           '2019/3/20 END
                        End If
                     End If
'                  Else
'                  '2017/4/18 END
'                     '中午打下班卡
'                     If Val(strPollData) > 121000 And Val(strPollData) < 133000 Then
'                        If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), "13:30", strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = True Then
'                           'If Val(Replace(strEndWorkTime, ":", "")) = Val(strWOffTime) Then 'Add By Sindy 2016/8/30 +if
'                           'Modify By Sindy 2019/3/20
'                           If PUB_ChkWorkDtTiUpdABS("P", rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), rsTmp.Fields("b1403"), _
'                              strRestKind, strWOnTime, strWOffTime, strStarWorkTime, strEndWorkTime, _
'                              strUpdDate, strUpdTime) = True Then GoTo RunEnd
'                           '2019/3/20 END
'                        End If
'                     End If
'                  End If
               End If
               '當打卡時間有落在假單時間中，系統自動核銷
               If Val(strPollData) > 0 Then
                  If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), Left(Format(strPollData, "00:00:00"), 5), strRestKind, bolRest1Day, True, , , strWOnTime, strWOffTime) = True Then
                     'If Val(Replace(strEndWorkTime, ":", "")) = Val(strWOffTime) Then 'Add By Sindy 2016/8/30 +if
                     'If Val(Replace(strEndWorkTime, ":", "")) >= Val(strWOffTime) Then 'Add By Sindy 2016/9/12 +if
                     'If Val(Replace(strStarWorkTime, ":", "")) <= Val(strWOnTime) Then 'Add By Sindy 2016/11/16 +if 寫反判斷,用上班時間檢查
                     'Modify By Sindy 2019/3/20
                     If PUB_ChkWorkDtTiUpdABS("A", rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), rsTmp.Fields("b1403"), _
                        strRestKind, strWOnTime, strWOffTime, strStarWorkTime, strEndWorkTime, _
                        strUpdDate, strUpdTime) = True Then GoTo RunEnd
                     '2019/3/20 END
                  End If
               End If
            End If
         End If
         
         'Add By Sindy 2013/11/8
         '在同仁*****該上下班的時段*****裡, 有出差單, 則不管有無刷卡,有異常時均系統自動確認
'            Call Pub_GetSpecWorkHour(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), strStarWorkTime, strEndWorkTime)
'            strStarWorkTime = Format(strStarWorkTime, "00:00")
'            strEndWorkTime = Format(strEndWorkTime, "00:00")
''            If strStarWorkTime = "00:00" Or strStarWorkTime = "" Then
''               strStarWorkTime = "09:00"
''               strEndWorkTime = "17:00"
''            End If
'            If strStarWorkTime = "" And PUB_bWkSpec = False And strMinPr02 <> "" Then
'               If Val(strMinPr02) <= 90000 Then
'                  If Val(strMinPr02) < 80000 Then
'                     strStarWorkTime = "08:00"
'                     strEndWorkTime = "17:00"
'                  ElseIf Val(strMinPr02) < 83000 Then
'                     strStarWorkTime = "08:30"
'                     strEndWorkTime = "17:30"
'                  Else
'                     strStarWorkTime = "09:00"
'                     strEndWorkTime = "18:00"
'                  End If
'               ElseIf Val(strMaxPr02) >= 170000 Then
'                  If Val(strMaxPr02) >= 180000 Then
'                     strStarWorkTime = "09:00"
'                     strEndWorkTime = "18:00"
'                  ElseIf Val(strMaxPr02) >= 173000 Then
'                     strStarWorkTime = "08:30"
'                     strEndWorkTime = "17:30"
'                  Else
'                     strStarWorkTime = "08:00"
'                     strEndWorkTime = "17:00"
'                  End If
'               End If
'            End If
         
         '上班異常
         If rsTmp.Fields("b1403") = "A" Then
            If strStarWorkTime <> "" Then
               If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), strStarWorkTime, strRestKind, bolRest1Day, True) = True Then
                  If strRestKind = "3" Then
                     strSql = "update ABS014 set B1405='" & IIf(strRestKind = "3", "6", "1") & "'" & _
                                               ",B1411='A'" & _
                                               ",B1412=" & strUpdDate & _
                                               ",B1413=" & strUpdTime & _
                              " where b1401='" & rsTmp.Fields("b1401") & "'" & _
                                " and b1402=" & rsTmp.Fields("b1402") & _
                                " and b1403='" & rsTmp.Fields("b1403") & "'"
                     cnnConnection.Execute strSql
                     PUB_UpdateB14Data = True 'Add By Sindy 2016/8/9
                     GoTo RunEnd
                  End If
               End If
            End If
         '下班異常
         Else
            If strEndWorkTime <> "" Then
               If CheckIsPersonRest(rsTmp.Fields("b1401"), rsTmp.Fields("b1402"), strEndWorkTime, strRestKind, bolRest1Day, True) = True Then
                  If strRestKind = "3" Then
                     strSql = "update ABS014 set B1405='" & IIf(strRestKind = "3", "6", "1") & "'" & _
                                               ",B1411='A'" & _
                                               ",B1412=" & strUpdDate & _
                                               ",B1413=" & strUpdTime & _
                              " where b1401='" & rsTmp.Fields("b1401") & "'" & _
                                " and b1402=" & rsTmp.Fields("b1402") & _
                                " and b1403='" & rsTmp.Fields("b1403") & "'"
                     cnnConnection.Execute strSql
                     PUB_UpdateB14Data = True 'Add By Sindy 2016/8/9
                     GoTo RunEnd
                  End If
               End If
            End If
         End If
         '2013/11/8 END
         
RunEnd:
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Add By Sindy 2019/3/21 用請假的日期時間,檢查打卡時間,更新打卡異常資料
Public Function PUB_ChkWorkDtTiUpdABS(strType As String, strB1401 As String, strB1402 As String, strB1403 As String, _
   strRestKind As String, strWOnTime As String, strWOffTime As String, strStarWorkTime As String, strEndWorkTime As String, _
   strUpdDate As String, strUpdTime As String) As Boolean
   
   PUB_ChkWorkDtTiUpdABS = False
   If strType = "A" Then
      'If Val(Replace(strStarWorkTime, ":", "")) = Val(strWOnTime) Then 'Add By Sindy 2016/8/30 +if
      'If Val(Replace(strStarWorkTime, ":", "")) <= Val(strWOnTime) Then 'Add By Sindy 2016/9/12 +if
      'Modify By Sindy 2019/3/20
      If Val(strWOnTime) > 0 And Val(Replace(strStarWorkTime, ":", "")) > 0 Then
         If Val(Replace(strWOnTime, ":", "")) <= Val(Replace(strStarWorkTime, ":", "")) Then
      '2019/3/20 END
            strSql = "update ABS014 set B1405='" & IIf(strRestKind = "3", "6", "1") & "'" & _
                                      ",B1411='A'" & _
                                      ",B1412=" & strUpdDate & _
                                      ",B1413=" & strUpdTime & _
                     " where b1401='" & strB1401 & "'" & _
                       " and b1402=" & strB1402 & _
                       " and b1403='" & strB1403 & "'"
            cnnConnection.Execute strSql
            PUB_ChkWorkDtTiUpdABS = True
            Exit Function
         End If
      End If
   Else
      'If Val(Replace(strEndWorkTime, ":", "")) = Val(strWOffTime) Then 'Add By Sindy 2016/8/30 +if
      'If Val(Replace(strEndWorkTime, ":", "")) >= Val(strWOffTime) Then 'Add By Sindy 2016/9/12 +if
      'Modify By Sindy 2019/3/20 ex:99025-20190308-P
      If Val(strWOffTime) > 0 And Val(Replace(strEndWorkTime, ":", "")) > 0 Then
         If Val(Replace(strWOffTime, ":", "")) >= Val(Replace(strEndWorkTime, ":", "")) Then
      '2019/3/20 END
            strSql = "update ABS014 set B1405='" & IIf(strRestKind = "3", "6", "1") & "'" & _
                                      ",B1411='A'" & _
                                      ",B1412=" & strUpdDate & _
                                      ",B1413=" & strUpdTime & _
                     " where b1401='" & strB1401 & "'" & _
                       " and b1402=" & strB1402 & _
                       " and b1403='" & strB1403 & "'"
            cnnConnection.Execute strSql
            PUB_ChkWorkDtTiUpdABS = True
            Exit Function
         End If
      End If
   End If
End Function

'Add By Sindy 2013/9/12 檢查當時是否需要為他人職代
'Modified by Morgan 2015/5/26 +pNoList:被代理人清單
'Modify by Sindy 2016/10/12 + bolInSpecialDuty:是否含特殊職代
'Modify By Sindy 2025/2/24 此函數 bolInSpecialDuty 已取消使用
'Modify By Sindy 2025/10/27 + bolChkCaseDuty:是否要管案件職代有無設定: True-要 False-不用
Public Sub Pub_SetForOthersEmpCombo(ByVal StrST01 As String, Optional objCbo As Object, _
                  Optional bolClear As Boolean = True, Optional pNoList As String, _
                  Optional bolInSpecialDuty As Boolean = False, Optional bolChkCaseDuty As Boolean = True)
Dim strText As String, strEmp As String
Dim ii As Integer
Dim strCompTime As String, strCompDate As String
Dim bolRest As Boolean, bolIsRest1Day As Boolean 'Add By Sindy 2016/7/19
Dim varArr As Variant
Dim strChkST01 As String
   
   If Not objCbo Is Nothing Then
      If bolClear = True Then objCbo.Clear
   End If
   
   'Add By Sindy 2024/3/29
   strChkST01 = StrST01
   If Mid(strChkST01, 4, 1) = "9" Then '人員是否為虛建員編
      '抓回正式員編
      strSql = "SELECT ST01,ST02 FROM Staff" & _
               " where st01 in(SELECT ST14 FROM Staff WHERE ST01='" & strChkST01 & "' AND ST04='1' AND ST14 is not null)" & _
               " and substr(st01,1,1)<'F' and substr(st01,1,1)>='6'" & _
               " AND ST04='1'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If "" & RsTemp.Fields("st01") <> "" Then
            strChkST01 = "" & RsTemp.Fields("st01")
         End If
      End If
   End If
   '2024/3/29 END
   
'   '先抓案件職代
'   strSql = "      SELECT B0101,1 FROM ABS001,Staff WHERE B0117='" & strChkST01 & "' AND B0117=ST01(+) AND ST04='1' " & _
'            "Union SELECT B0101,2 FROM ABS001,Staff WHERE B0119='" & strChkST01 & "' AND B0119=ST01(+) AND ST04='1' " & _
'            "Union SELECT B0101,3 FROM ABS001,Staff WHERE B0121='" & strChkST01 & "' AND B0121=ST01(+) AND ST04='1' " & _
'            "Union SELECT B0101,4 FROM ABS001,Staff WHERE B0123='" & strChkST01 & "' AND B0123=ST01(+) AND ST04='1' " & _
'            "order by 2 asc"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 0 Then
'      '再抓人事職代
'      strSql = "      SELECT B0101,1 FROM ABS001,Staff WHERE B0102='" & strChkST01 & "' AND B0102=ST01(+) AND ST04='1' " & _
'               "Union SELECT B0101,2 FROM ABS001,Staff WHERE B0103='" & strChkST01 & "' AND B0103=ST01(+) AND ST04='1' " & _
'               "Union SELECT B0101,3 FROM ABS001,Staff WHERE B0104='" & strChkST01 & "' AND B0104=ST01(+) AND ST04='1' " & _
'               "Union SELECT B0101,4 FROM ABS001,Staff WHERE B0105='" & strChkST01 & "' AND B0105=ST01(+) AND ST04='1' " & _
'               "Union SELECT B0101,5 FROM ABS001,Staff WHERE B0106='" & strChkST01 & "' AND B0106=ST01(+) AND ST04='1' " & _
'               "Union SELECT B0101,6 FROM ABS001,Staff WHERE B0107='" & strChkST01 & "' AND B0107=ST01(+) AND ST04='1' " & _
'               "order by 2 asc"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   End If
   'Modify By Sindy 2014/6/23 讀取案件職代及未設定案件職代的人事職代 Ex.例如游經理請假,文雄為職代
   'Modify By Sindy 2021/5/24 +,B0127
   'Modify By Sindy 2025/10/27 + bolChkCaseDuty:是否要管案件職代有無設定: True-要 False-不用
   If bolChkCaseDuty = True Then
      strExc(10) = "AND B0117 is null AND B0119 is null AND B0121 is null AND B0123 is null "
   Else
      strExc(10) = ""
   End If
   '2025/10/27 END
   strSql = "      SELECT B0101,1,B0127 FROM ABS001,Staff WHERE B0117='" & strChkST01 & "' AND B0101=ST01(+) AND ST04='1' " & _
            "Union SELECT B0101,2,B0127 FROM ABS001,Staff WHERE B0119='" & strChkST01 & "' AND B0101=ST01(+) AND ST04='1' " & _
            "Union SELECT B0101,3,B0127 FROM ABS001,Staff WHERE B0121='" & strChkST01 & "' AND B0101=ST01(+) AND ST04='1' " & _
            "Union SELECT B0101,4,B0127 FROM ABS001,Staff WHERE B0123='" & strChkST01 & "' AND B0101=ST01(+) AND ST04='1' " & _
            "Union SELECT B0101,5,B0127 FROM ABS001,Staff WHERE B0102='" & strChkST01 & "' AND B0101=ST01(+) AND ST04='1' " & strExc(10) & _
            "Union SELECT B0101,6,B0127 FROM ABS001,Staff WHERE B0103='" & strChkST01 & "' AND B0101=ST01(+) AND ST04='1' " & strExc(10) & _
            "Union SELECT B0101,7,B0127 FROM ABS001,Staff WHERE B0104='" & strChkST01 & "' AND B0101=ST01(+) AND ST04='1' " & strExc(10) & _
            "Union SELECT B0101,8,B0127 FROM ABS001,Staff WHERE B0105='" & strChkST01 & "' AND B0101=ST01(+) AND ST04='1' " & strExc(10) & _
            "Union SELECT B0101,9,B0127 FROM ABS001,Staff WHERE B0106='" & strChkST01 & "' AND B0101=ST01(+) AND ST04='1' " & strExc(10) & _
            "Union SELECT B0101,10,B0127 FROM ABS001,Staff WHERE B0107='" & strChkST01 & "' AND B0101=ST01(+) AND ST04='1' " & strExc(10)
   'Add By Sindy 2016/10/12 + 含特殊職代
   'Modify By Sindy 2025/2/24 應該都要 含特殊職代 做檢查
'   If bolInSpecialDuty = True Then
   '2025/2/24 END
      'Modify By Sindy 2023/5/4 + and B0209='2' :案件簽核職代
      strSql = strSql & _
            "Union SELECT B0201,11,'' FROM ABS002,Staff WHERE B0202='" & strChkST01 & "' AND B0201=ST01(+) AND ST04='1' and B0209='2' " & _
            "Union SELECT B0201,12,'' FROM ABS002,Staff WHERE B0203='" & strChkST01 & "' AND B0201=ST01(+) AND ST04='1' and B0209='2' " & _
            "Union SELECT B0201,13,'' FROM ABS002,Staff WHERE B0204='" & strChkST01 & "' AND B0201=ST01(+) AND ST04='1' and B0209='2' " & _
            "Union SELECT B0201,14,'' FROM ABS002,Staff WHERE B0205='" & strChkST01 & "' AND B0201=ST01(+) AND ST04='1' and B0209='2' " & _
            "Union SELECT B0201,15,'' FROM ABS002,Staff WHERE B0206='" & strChkST01 & "' AND B0201=ST01(+) AND ST04='1' and B0209='2' " & _
            "Union SELECT B0201,16,'' FROM ABS002,Staff WHERE B0207='" & strChkST01 & "' AND B0201=ST01(+) AND ST04='1' and B0209='2' "
'   End If
   '2016/10/12 END
   strSql = strSql & " order by 2 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            If Not IsNull(RsTemp.Fields(0)) Then
               '檢查當事人,是否當時為休假狀況
               'Modify By Sindy 2016/7/19
               bolRest = CheckIsPersonRest(RsTemp.Fields(0), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolIsRest1Day)
               '有休假或整日休假
               'Modify By Sindy 2021/5/24 + Or RsTemp.Fields(2) = "Y"
               If bolRest = True Or bolIsRest1Day = True Or RsTemp.Fields(2) = "Y" Then
               '2016/7/19 END
                  If Not objCbo Is Nothing Then
                     strText = Trim(RsTemp.Fields(0)) & " " & GetPrjSalesNM(Trim(RsTemp.Fields(0)))
                     'Modify By Sindy 2015/12/21
                     For ii = 0 To objCbo.ListCount - 1
                        If InStr(objCbo.List(ii), Trim(RsTemp.Fields(0))) = 1 Then
                           Exit For
                        End If
                     Next ii
                     If ii = objCbo.ListCount Then
                     '2015/12/21 END
                        objCbo.AddItem strText
                     End If
                  End If
                  If InStr(pNoList, Trim(RsTemp.Fields(0))) = 0 Then
                     pNoList = pNoList & Trim(RsTemp.Fields(0)) & ";" 'Added by Morgan 2015/5/26
                     'Add By Sindy 2021/7/12 檢查是否有需代理的MCTF
                     strText = GetMCTF0XAllCode(Trim(RsTemp.Fields(0)))
                     If strText <> "" Then
                        varArr = Split(strText, "','")
                        For ii = 0 To UBound(varArr)
                           If InStr(pNoList, Trim(varArr(ii))) = 0 Then
                              If Not objCbo Is Nothing Then
                                 objCbo.AddItem Trim(varArr(ii)) & " " & GetPrjSalesNM(Trim(varArr(ii)))
                              End If
                              pNoList = pNoList & Trim(varArr(ii)) & ";"
                           End If
                        Next ii
                     End If
                     '2021/7/12 END
                  End If
               End If
            End If
            .MoveNext
         Loop
      End With
   End If
   
   'Add By Sindy 2016/3/30
   '增加假單裡的職代
   '注意比較的日期格式為 YYYYMMDDHHMM
'   strCompTime = Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)
'   strCompTime = IIf(strCompTime = "24:00", "2400", Format(strCompTime, "hhmm"))
'   strCompDate = strSrvDate(1) & strCompTime
'   strSql = "SELECT B1003 FROM ABS011,ABS010" & _
'            " where B1101 in(" & _
'            " SELECT SA09 From staff_Absence" & _
'            " WHERE " & strCompDate & " between SA02||substr('0000'||SA03,-4) and SA04||substr('0000'||SA05,-4)" & _
'            " Union" & _
'            " SELECT SB10 From staff_busi_trip" & _
'            " WHERE " & strCompDate & " between SB02||substr('0000'||SB03,-4) and SB04||substr('0000'||SB05,-4)" & _
'            ") and B1102='1' and B1104='" & strChkST01 & "' and B1101=B1001"
   'Modify By Sindy 2024/6/20 加簽核中當日表單
   strSql = "SELECT B1003 FROM ABS011,ABS010" & _
            " where B1101 in(" & _
            " SELECT SA09 From staff_Absence" & _
            " WHERE " & strSrvDate(1) & " between SA02 and SA04" & _
            " Union" & _
            " SELECT SB10 From staff_busi_trip" & _
            " WHERE " & strSrvDate(1) & " between SB02 and SB04" & _
            " Union" & _
            " SELECT B1001 FROM ABS010 WHERE B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "') and B1018 not in('" & 退回 & "','" & 註銷 & "','" & 已核准 & "') and " & strSrvDate(1) & " between B1004 and B1006" & _
            ") and B1102='1' and B1104='" & strChkST01 & "' and B1101=B1001 and B1003 is not null" & _
            " group by B1003"
   'Add By Sindy 2024/12/25 當日的 主管代填 抓人事職代
   strSql = strSql & " union SELECT B1003 FROM ABS010,ABS001 WHERE B1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "') and B1018 in('" & 主管代填 & "')" & _
            " and " & strSrvDate(1) & " between B1004 and B1006" & _
            " and (B0102='" & strChkST01 & "' or B0103='" & strChkST01 & "' or B0104='" & strChkST01 & "' or B0105='" & strChkST01 & "' or B0106='" & strChkST01 & "' or B0107='" & strChkST01 & "')" & _
            " and B1003=B0101"
   '2024/12/25 END
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            '檢查當事人,是否當時為休假狀況
            'Modify By Sindy 2016/7/19
            bolRest = CheckIsPersonRest(RsTemp.Fields(0), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolIsRest1Day)
            '有休假或整日休假
            If bolRest = True Or bolIsRest1Day = True Then
            '2016/7/19 END
               If Not objCbo Is Nothing Then
                  strText = Trim(RsTemp.Fields(0)) & " " & GetPrjSalesNM(Trim(RsTemp.Fields(0)))
                  For ii = 0 To objCbo.ListCount - 1
                     If InStr(objCbo.List(ii), Trim(RsTemp.Fields(0))) = 1 Then
                        Exit For
                     End If
                  Next
                  If ii = objCbo.ListCount Then
                     objCbo.AddItem strText
                  End If
               End If
               If InStr(pNoList, Trim(RsTemp.Fields(0))) = 0 Then
                  pNoList = pNoList & Trim(RsTemp.Fields(0)) & ";"
                  'Add By Sindy 2021/7/12 檢查是否有需代理的MCTF
                  strText = GetMCTF0XAllCode(Trim(RsTemp.Fields(0)))
                  If strText <> "" Then
                     varArr = Split(strText, "','")
                     For ii = 0 To UBound(varArr)
                        If InStr(pNoList, Trim(varArr(ii))) = 0 Then
                           If Not objCbo Is Nothing Then
                              objCbo.AddItem Trim(varArr(ii)) & " " & GetPrjSalesNM(Trim(varArr(ii)))
                           End If
                           pNoList = pNoList & Trim(varArr(ii)) & ";"
                        End If
                     Next ii
                  End If
                  '2021/7/12 END
               End If
            End If
            .MoveNext
         Loop
      End With
   End If
   '2016/3/30 END
   
   'Add By Sindy 2021/5/24 居家無法連線者 => 居家職代(1) / 居家職代(2)
   strSql = "      SELECT B0101,1 FROM ABS001,Staff WHERE B0127='Y' AND B0128='" & strChkST01 & "' AND B0101=ST01(+) AND ST04='1' " & _
            "Union SELECT B0101,2 FROM ABS001,Staff WHERE B0127='Y' AND B0129='" & strChkST01 & "' AND B0101=ST01(+) AND ST04='1' "
   strSql = strSql & " order by 2 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            If Not IsNull(RsTemp.Fields(0)) Then
               If Not objCbo Is Nothing Then
                  strText = Trim(RsTemp.Fields(0)) & " " & GetPrjSalesNM(Trim(RsTemp.Fields(0)))
                  For ii = 0 To objCbo.ListCount - 1
                     If InStr(objCbo.List(ii), Trim(RsTemp.Fields(0))) = 1 Then
                        Exit For
                     End If
                  Next
                  If ii = objCbo.ListCount Then
                     objCbo.AddItem strText
                  End If
               End If
               If InStr(pNoList, Trim(RsTemp.Fields(0))) = 0 Then
                  pNoList = pNoList & Trim(RsTemp.Fields(0)) & ";"
                  'Add By Sindy 2021/7/12 檢查是否有需代理的MCTF
                  strText = GetMCTF0XAllCode(Trim(RsTemp.Fields(0)))
                  If strText <> "" Then
                     varArr = Split(strText, "','")
                     For ii = 0 To UBound(varArr)
                        If InStr(pNoList, Trim(varArr(ii))) = 0 Then
                           If Not objCbo Is Nothing Then
                              objCbo.AddItem Trim(varArr(ii)) & " " & GetPrjSalesNM(Trim(varArr(ii)))
                           End If
                           pNoList = pNoList & Trim(varArr(ii)) & ";"
                        End If
                     Next ii
                  End If
                  '2021/7/12 END
               End If
            End If
            .MoveNext
         Loop
      End With
   End If
   '2021/5/24 END
   
   'Add By Sindy 2023/5/15 增加檢查信件未沖銷但已離職的人員
   strSql = "select ir04,st14,count(*) from(" & _
            "select ir04,ir01,ir03,ir22 from tminput,inputrecord,staff" & _
            " Where nvl(ti16, 0) = 0 And ti01 = iR01 And ti03 = iR03 And iR04 = st01" & _
            " and ir08=0 and st04='2'" & _
            " Union All" & _
            " select ir04,ir01,ir03,ir22 from ipdeptinput,inputrecord,staff" & _
            " Where nvl(ii16, 0) = 0 And ii01 = iR01 And ii03 = iR03 And iR04 = st01" & _
            " and ir08=0 and st04='2'" & _
            " Union All" & _
            " select ir04,ir01,ir03,ir22 from patentinput,inputrecord,staff" & _
            " Where nvl(pi16, 0) = 0 And pi01 = iR01 And pi03 = iR03 And iR04 = st01" & _
            " and ir08=0 and st04='2'" & _
            ") a,staff where a.ir04=st01 and instr(st14,'" & strChkST01 & "')>0 group by ir04,st14"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            If Not IsNull(RsTemp.Fields("ir04")) Then
               strText = Trim(RsTemp.Fields("ir04")) & " " & GetPrjSalesNM(Trim(RsTemp.Fields("ir04")))
               If UCase(TypeName(objCbo)) <> "NOTHING" Then 'Added by Lydia 2023/05/16
                  For ii = 0 To objCbo.ListCount - 1
                     If InStr(objCbo.List(ii), Trim(strText)) = 1 Then
                        Exit For
                     End If
                  Next ii
                  If ii = objCbo.ListCount Then
                     objCbo.AddItem strText
                  End If
               End If 'Added by Lydia 2023/05/16
               If InStr(pNoList, Trim(RsTemp.Fields("ir04"))) = 0 Then
                  pNoList = pNoList & Trim(RsTemp.Fields("ir04")) & ";"
               End If
            End If
            .MoveNext
         Loop
      End With
   End If
   '2023/5/15 END
   
   'Add By Sindy 2023/4/10 檢查休假人員是否有虛建員編(依內部郵件收件員工編號抓資料) ex:A3014蕭茹曣 有另一帳號 A3099蕭茹曣CFP
   If pNoList <> "" Then
      varArr = Split(pNoList, ";")
      For ii = 0 To UBound(varArr)
         If Trim(varArr(ii)) <> "" Then
            'Modify By Sindy 2024/3/29 呂達陽B1016,有外專的身分B1098; 增加部門判斷: and st03='" & Pub_StrUserSt03 & "'
            strSql = "SELECT ST01,ST02 FROM Staff WHERE ST14='" & Trim(varArr(ii)) & "' AND ST04='1'" & _
                     " and length(ST14)=5 and ST14<>'99997'" & _
                     " and substr(st01,1,1)<'F' and substr(st01,1,1)>='6' and st03='" & Pub_StrUserSt03 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If InStr(pNoList, RsTemp.Fields("ST01")) = 0 Then
                  If Not objCbo Is Nothing Then
                     objCbo.AddItem Trim(RsTemp.Fields("ST01")) & " " & GetPrjSalesNM(Trim(RsTemp.Fields("ST01")))
                  End If
                  pNoList = pNoList & Trim(RsTemp.Fields("ST01")) & ";"
               End If
            End If
         End If
      Next ii
   End If
   '2023/4/10 END
   
   'Add By Sindy 2021/11/8 林律師(98003)若請假時，商爭案件的判發人，針對T類案的核判，設定江協理(98020)為林律師的職代。
   If strChkST01 = "98020" Then
      strEmp = "98003"
      '檢查當事人,是否當時為休假狀況
      bolRest = CheckIsPersonRest(strEmp, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolIsRest1Day)
      '有休假或整日休假
      If bolRest = True Or bolIsRest1Day = True Then
         If Not objCbo Is Nothing Then
            strText = strEmp & " " & GetPrjSalesNM(strEmp)
            For ii = 0 To objCbo.ListCount - 1
               If InStr(objCbo.List(ii), strEmp) = 1 Then
                  Exit For
               End If
            Next
            If ii = objCbo.ListCount Then
               objCbo.AddItem strText
            End If
         End If
      End If
   End If
   '2021/11/8 END
   
   'Add By Sindy 2024/4/3 虛建員編只能代替虛建員編操作
   If Mid(StrST01, 4, 1) = "9" Then
      If Not objCbo Is Nothing Then
         For ii = objCbo.ListCount - 1 To 1 Step -1
            If Mid(Trim(objCbo.List(ii)), 4, 1) <> "9" Then
               objCbo.RemoveItem ii
            End If
         Next ii
      End If
      If pNoList <> "" Then
         varArr = Split(pNoList, ";")
         For ii = 0 To UBound(varArr)
            If Mid(Trim(varArr(ii)), 4, 1) <> "9" And Trim(varArr(ii)) <> "" Then
               pNoList = Replace(pNoList, ";" & Trim(varArr(ii)), "")
               pNoList = Replace(pNoList, Trim(varArr(ii)) & ";", "")
               pNoList = Replace(pNoList, Trim(varArr(ii)), "")
            End If
         Next ii
      End If
   End If
   '2024/4/3 END
End Sub

'Add By Sindy 2014/2/12
'檢查是否有主管代填連續假單
'Modify By Sindy 2014/3/7 + , Optional bolShowMsg As Boolean = True
'Modify By Sindy 2020/5/27 + , Optional strKind As String = ""
Public Function PUB_ChkSerialRest(ByVal strB1001 As String, ByVal strB1002 As String, ByVal strB1003 As String, _
                                  Optional bolShowMsg As Boolean = True, Optional strKind As String = "") As Boolean
                                  
Dim strCon As String, strText As String, strPrvB1001 As String, strPrvB1002 As String, bolPrvIs1Day As Boolean
Dim m_Day As Integer, m_Hour As Double
Dim strB1009 As String, strB1010 As String
Dim strSql_A As String, strSql_B As String
Dim strConSql As String, strConSqla As String, strConSqlb As String
Dim rsTmp As New ADODB.Recordset
   
On Error GoTo ErrHnd
   
   PUB_ChkSerialRest = False
   'Modify By Sindy 2014/10/30 排除人事處已核准的主管代填表單
   'Modify By Sindy 2020/5/27 檢查在新增假單時,是否有連續假單
   If strB1001 <> "" Then
      strConSql = "and (b1018='" & 主管代填 & "' or b1001='" & strB1001 & "')"
      If strKind <> "" Then
         If Left(strB1002, 2) = 表單類別_請假 Then
            strConSqla = " and (b1008 is null or b1008='" & strKind & "')"
         Else
            strConSqlb = " and (b1014 is null or b1014='" & strKind & "')"
         End If
      End If
   Else
      strConSql = "and b1018='" & 主管代填 & "'"
   End If
   '2020/5/27 END
   strSql_A = "select abs010.* from abs010,staff_Absence where b1003='" & strB1003 & "'" & strConSql & strConSqla & _
              " and b1002='" & 表單類別_請假 & "' and b1001=SA09(+) and sa09 is null"
   strSql_B = "select abs010.* from abs010,staff_busi_trip where b1003='" & strB1003 & "'" & strConSql & strConSqlb & _
              " and b1002='" & 表單類別_出差 & "' and b1001=SB10(+) and SB10 is null"
    If strB1002 = "" Then
      strSql = strSql_A & " union " & strSql_B
   Else
      If Left(strB1002, 2) = 表單類別_請假 Then
         strSql = strSql_A
      Else
         strSql = strSql_B
      End If
   End If
'   If strB1002 = "" Then
'      strCon = " and b1002 in('" & 表單類別_請假 & "','" & 表單類別_出差 & "')"
'   Else
'      strCon = " and b1002='" & Left(strB1002, 2) & "'"
'   End If
'   strSql = "select * from abs010 where b1003='" & strB1003 & "' and (b1018='" & 主管代填 & "' or b1018 is null)" & strCon & _
'            " order by b1002 asc,b1004 asc"
   strSql = "select * from (" & strSql & ") order by b1002 asc,b1004 asc"
   '2014/10/30 END
   If rsTmp.State = 1 Then rsTmp.Close
   With rsTmp
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         .MoveFirst
         strDate = ""
         Do While Not .EOF
            If strDate <> "" Then
               If strPrvB1002 = .Fields("B1002") Then '同性質表單類別才需要檢查
                  If Val(strDate) = Val(.Fields("B1004")) Then
                     '檢查是否整日
                     strB1009 = "" & .Fields("B1009")
                     strB1010 = "" & .Fields("B1010")
                     If strB1009 = "" Then
                        If .Fields("B1002") = 表單類別_出差 Then
                           Call PUB_CountHour_Busi_Trip(.Fields("B1004"), Format(.Fields("B1005"), "0000"), .Fields("B1006"), Format(.Fields("B1007"), "0000"), m_Day, m_Hour)
                           If m_Day > 0 Then
                              strB1009 = m_Day
                           Else
                              strB1009 = "0"
                           End If
                           If m_Hour > 0 Then
                              strB1010 = m_Hour
                           Else
                              strB1010 = "0"
                           End If
                        ElseIf .Fields("B1002") = 表單類別_請假 Then
                           Call PUB_CountDayHour(.Fields("B1003"), .Fields("B1004"), Format(.Fields("B1005"), "0000"), .Fields("B1006"), Format(.Fields("B1007"), "0000"), "", "", strB1009, strB1010, "" & .Fields("B1008"), False)
                        End If
                     End If
                     '此表單為整日或(前表單為整日,此表單非整日亦也算連續假單)
                     If Val(strB1010) = 0 Or _
                        (bolPrvIs1Day = True And strB1010 > 0) Then
                        PUB_ChkSerialRest = True
                        If InStr(strText, strPrvB1001) = 0 Then
                           strText = strText & "," & strPrvB1001 '前一筆
                        End If
                        strText = strText & "," & .Fields("B1001")
                     End If
                     '記錄前一筆是整日或非整日
                     If Val(strB1010) = 0 Then
                        bolPrvIs1Day = True
                     Else
                        bolPrvIs1Day = False
                     End If
                  End If
               End If
            End If
            strPrvB1001 = .Fields("B1001")
            strPrvB1002 = .Fields("B1002")
            '第一筆:
            If strDate = "" Then
               '檢查是否整日
               strB1009 = "" & .Fields("B1009")
               strB1010 = "" & .Fields("B1010")
               If strB1009 = "" Then
                  If .Fields("B1002") = 表單類別_出差 Then
                     Call PUB_CountHour_Busi_Trip(.Fields("B1004"), Format(.Fields("B1005"), "0000"), .Fields("B1006"), Format(.Fields("B1007"), "0000"), m_Day, m_Hour)
                     If m_Day > 0 Then
                        strB1009 = m_Day
                     Else
                        strB1009 = "0"
                     End If
                     If m_Hour > 0 Then
                        strB1010 = m_Hour
                     Else
                        strB1010 = "0"
                     End If
                  ElseIf .Fields("B1002") = 表單類別_請假 Then
                     Call PUB_CountDayHour(.Fields("B1003"), .Fields("B1004"), Format(.Fields("B1005"), "0000"), .Fields("B1006"), Format(.Fields("B1007"), "0000"), "", "", strB1009, strB1010, "" & .Fields("B1008"), False)
                  End If
               End If
               '記錄前一筆是整日或非整日
               If Val(strB1010) = 0 Then
                  bolPrvIs1Day = True
               Else
                  bolPrvIs1Day = False
               End If
            End If
            '計算下一天
            strDate = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(.Fields("B1006")))))
            If .Fields("B1002") = 表單類別_請假 Then
               Do While ChkWorkDay(strDate) = False
                  strDate = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(strDate)))
               Loop
            End If
            
            .MoveNext
         Loop
      End If
   End With
   If PUB_ChkSerialRest = True Then
      strText = Mid(strText, 2)
      If strB1001 <> "" And InStr(strText, strB1001) = 0 Then '以防正在處理的假單,不在連續假單裡
         PUB_ChkSerialRest = False
      Else
         If bolShowMsg = True Then
            MsgBox "您有連續假單, 請刪除代填假單（" & strText & "）重新合併填寫成一張處理!!", vbExclamation
         End If
      End If
   End If
   Set rsTmp = Nothing
   Exit Function
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Add By Sindy 2020/5/28
'檢查是否有連續假單(含已核准,未核准)
'彈訊息給主管知道 及 查看全部天數決定其當筆假單要簽核到那一層級的主管
Public Function PUB_ChkSerialRest_ToSir(ByVal strB1001 As String, _
   Optional bolShowMsg As Boolean = True, Optional ByRef dblDay As Double = 0, _
   Optional ByVal strB1002 As String, Optional ByVal strB1003 As String, _
   Optional ByVal strB1004 As String, Optional ByVal strB1006 As String, _
   Optional ByVal dblB1009 As Double, Optional ByVal dblB1010 As Double) As Boolean
   
Dim strID As String
Dim strConSql As String
Dim strQuyDateS As String '查詢起始日期
Dim strQuyDateE As String '查詢迄止日期
Dim oForm As Form
Dim ii As Integer, intCnt As Integer, bolFind As Boolean, strB1008 As String 'Add By Sindy 2023/9/15
   
On Error GoTo ErrHnd
   
   PUB_ChkSerialRest_ToSir = False
   
   'Add By Sindy 2023/9/15
   intCnt = 0
   For ii = 1 To 10
      g_strB1002(ii) = ""
      g_strB1008(ii) = ""
      g_strDay(ii) = ""
   Next ii
   '2023/9/15 END
   
   If strB1001 <> "" Then
      'Modify By Sindy 2020/10/26 呼叫此函數時,就要傳入這些欄位值
'      strSql = "select * from abs010 where b1001='" & strB1001 & "'"
'      CheckOC3
'      With AdoRecordSet3
'         .CursorLocation = adUseClient
'         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If .RecordCount > 0 Then
'            strB1002 = .Fields("B1002")
'            strB1003 = .Fields("B1003")
'            strB1004 = "" & .Fields("B1004") '起始日期
'            strB1006 = "" & .Fields("B1006") '迄止日期
'            dblB1009 = Val("" & .Fields("B1009")) '日
'            dblB1010 = Val("" & .Fields("B1010")) '時
'         End If
'      End With
'      CheckOC3
   Else
      '要傳入欲檢查的假單資料,若無,則無法檢查離開
      If strB1002 = "" Or strB1003 = "" Or strB1004 = "" Or strB1006 = "" _
         Or (Val(dblB1009) = 0 And Val(dblB1010) = 0) Then
         Exit Function
      End If
   End If
   
   If strB1002 <> 表單類別_請假 And strB1002 <> 表單類別_出差 Then Exit Function
   Call Pub_GetSpecWorkHour(strB1003, strB1004) '特殊人員的工作時數
   
   '以起始日期往前面找尋表單:
   '計算前一個工作天
   strDate = strB1004 '起始日期
   strID = ""
Count_Previous1:
   strDate = DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strDate)))
'   If strB1002 = 表單類別_請假 Then
      Do While ChkWorkDay(strDate) = False
         strDate = DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strDate)))
      Loop
      '讀取符合請假區間的資料
      'Modify By Sindy 2020/10/26 + IIf(strB1001 <> "", " and b1001<>'" & strB1001 & "'", "")
      'Modify By Sindy 2020/10/26 排除註銷 and b1018<>'06'
      'Modify By Sindy 2023/9/15 +,B1002,B1008,B1014
      strSql = "select B1001,B1004,B1006,B1009,B1010,sa09,sa02,sa04,sa07,sa08,B1002,B1008,B1014" & _
               " from abs010,staff_Absence where b1003='" & strB1003 & "'" & _
               " and b1002='" & strB1002 & "' and b1001=SA09(+) and b1018<>'06'" & _
               " and " & strDate & " between b1004 and b1006" & IIf(strB1001 <> "", " and b1001<>'" & strB1001 & "'", "") & _
               " union " & _
               "select B1001,B1004,B1006,B1009,B1010,SB10 sa09,SB02 sa02,SB04 sa04,SB06 sa07,SB07 sa08,B1002,B1008,B1014" & _
               " from abs010,staff_busi_trip where b1003='" & strB1003 & "'" & _
               " and b1002='" & strB1002 & "' and b1001=SB10(+) and b1018<>'06'" & _
               " and " & strDate & " between b1004 and b1006" & IIf(strB1001 <> "", " and b1001<>'" & strB1001 & "'", "") & _
               " order by B1004 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If InStr(strID, RsTemp.Fields("B1001")) = 0 Then
            If "" & RsTemp.Fields("sa09") <> "" Then '已收錄至人事系統
               If Not (Val(strDate) >= Val(RsTemp.Fields("sa02")) And _
                      Val(strDate) <= Val(RsTemp.Fields("sa04"))) Then
                  GoTo GetChkNext
               Else
                  dblB1009 = dblB1009 + Val(RsTemp.Fields("sa07")) '日
                  dblB1010 = dblB1010 + Val(RsTemp.Fields("sa08")) '時
                  PUB_ChkSerialRest_ToSir = True
                  'Add By Sindy 2023/9/15 記錄起來,後續程式使用
                  bolFind = False
                  If RsTemp.Fields("B1002") = 表單類別_請假 Then
                     strB1008 = "" & RsTemp.Fields("B1008")
                  Else
                     strB1008 = "" & RsTemp.Fields("B1014")
                  End If
                  For ii = 1 To intCnt
                     If g_strB1002(ii) = RsTemp.Fields("B1002") _
                        And g_strB1008(ii) = strB1008 Then
                        g_strDay(ii) = g_strDay(ii) + Val(RsTemp.Fields("sa07")) + (Val(RsTemp.Fields("sa08")) / PUB_intWkHour)
                        bolFind = True
                        Exit For
                     End If
                  Next
                  If bolFind = False Then
                     intCnt = intCnt + 1
                     g_strB1002(intCnt) = RsTemp.Fields("B1002")
                     g_strB1008(intCnt) = strB1008
                     g_strDay(intCnt) = Val(RsTemp.Fields("sa07")) + (Val(RsTemp.Fields("sa08")) / PUB_intWkHour)
                  End If
                  '2023/9/15 END
               End If
            Else
               dblB1009 = dblB1009 + Val("" & RsTemp.Fields("B1009")) '日
               dblB1010 = dblB1010 + Val("" & RsTemp.Fields("B1010")) '時
               PUB_ChkSerialRest_ToSir = True
               'Add By Sindy 2023/9/15 記錄起來,後續程式使用
               bolFind = False
               If RsTemp.Fields("B1002") = 表單類別_請假 Then
                  strB1008 = "" & RsTemp.Fields("B1008")
               Else
                  strB1008 = "" & RsTemp.Fields("B1014")
               End If
               For ii = 1 To intCnt
                  If g_strB1002(ii) = "" & RsTemp.Fields("B1002") _
                     And g_strB1008(ii) = strB1008 Then
                     g_strDay(ii) = g_strDay(ii) + Val("" & RsTemp.Fields("B1009")) + (Val("" & RsTemp.Fields("B1010")) / PUB_intWkHour)
                     bolFind = True
                     Exit For
                  End If
               Next
               If bolFind = False Then
                  intCnt = intCnt + 1
                  g_strB1002(intCnt) = "" & RsTemp.Fields("B1002")
                  g_strB1008(intCnt) = strB1008
                  g_strDay(intCnt) = Val("" & RsTemp.Fields("B1009")) + (Val("" & RsTemp.Fields("B1010")) / PUB_intWkHour)
               End If
               '2023/9/15 END
            End If
         End If
         If InStr(strID, RsTemp.Fields("B1001")) = 0 Then
            strQuyDateS = RsTemp.Fields("B1004")
            If strID = "" Then
               strID = RsTemp.Fields("B1001")
            Else
               strID = strID & "," & RsTemp.Fields("B1001")
            End If
         End If
         GoTo Count_Previous1
      End If
      
'   '表單類別_出差
'   Else
'      Do While ChkWorkDay(strDate) = False
'         strDate = DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strDate)))
'      Loop
'      '讀取符合請假區間的資料
'      strSql = "select * from abs010,staff_busi_trip where b1003='" & strB1003 & "'" & _
'        " and b1002='" & strB1002 & "' and b1001=SB10(+)" & _
'        " and " & strDate & " between b1004 and b1006"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         If InStr(strID, RsTemp.Fields("B1001")) = 0 Then
'            If "" & RsTemp.Fields("SB10") <> "" Then '已收錄至人事系統
'               If Not (Val(strDate) >= Val(RsTemp.Fields("sb02")) And _
'                      Val(strDate) <= Val(RsTemp.Fields("sb04"))) Then
'                  GoTo GetChkNext
'               Else
'                  dblB1009 = dblB1009 + Val(RsTemp.Fields("sb06")) '日
'                  dblB1010 = dblB1010 + Val(RsTemp.Fields("sb07")) '時
'                  PUB_ChkSerialRest_ToSir = True
'               End If
'            Else
'               dblB1009 = dblB1009 + Val(RsTemp.Fields("B1009")) '日
'               dblB1010 = dblB1010 + Val(RsTemp.Fields("B1010")) '時
'               PUB_ChkSerialRest_ToSir = True
'            End If
'         End If
'         If InStr(strID, RsTemp.Fields("B1001")) = 0 Then
'            strQuyDateS = RsTemp.Fields("B1004")
'            If strID = "" Then
'               strID = RsTemp.Fields("B1001")
'            Else
'               strID = strID & "," & RsTemp.Fields("B1001")
'            End If
'         End If
'         GoTo Count_Previous1
'      End If
'   End If
   
GetChkNext:
   '以迄止日期往後面找尋表單:
   '計算後一個工作天
   strDate = strB1006 '迄止日期
   strID = ""
Count_Next1:
   strDate = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(strDate)))
'   If strB1002 = 表單類別_請假 Then
      Do While ChkWorkDay(strDate) = False
         strDate = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(strDate)))
      Loop
      '讀取符合請假區間的資料
      'Modify By Sindy 2020/10/26 + IIf(strB1001 <> "", " and b1001<>'" & strB1001 & "'", "")
      'Modify By Sindy 2020/10/26 排除註銷 and b1018<>'06'
      'Modify By Sindy 2023/9/15 +,B1002,B1008,B1014
      strSql = "select B1001,B1004,B1006,B1009,B1010,sa09,sa02,sa04,sa07,sa08,B1002,B1008,B1014" & _
               " from abs010,staff_Absence where b1003='" & strB1003 & "'" & _
               " and b1002='" & strB1002 & "' and b1001=SA09(+) and b1018<>'06'" & _
               " and " & strDate & " between b1004 and b1006" & IIf(strB1001 <> "", " and b1001<>'" & strB1001 & "'", "") & _
               " union " & _
               "select B1001,B1004,B1006,B1009,B1010,SB10 sa09,SB02 sa02,SB04 sa04,SB06 sa07,SB07 sa08,B1002,B1008,B1014" & _
               " from abs010,staff_busi_trip where b1003='" & strB1003 & "'" & _
               " and b1002='" & strB1002 & "' and b1001=SB10(+) and b1018<>'06'" & _
               " and " & strDate & " between b1004 and b1006" & IIf(strB1001 <> "", " and b1001<>'" & strB1001 & "'", "") & _
               " order by B1004 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If InStr(strID, RsTemp.Fields("B1001")) = 0 Then
            If "" & RsTemp.Fields("sa09") <> "" Then '已收錄至人事系統
               If Not (Val(strDate) >= Val(RsTemp.Fields("sa02")) And _
                      Val(strDate) <= Val(RsTemp.Fields("sa04"))) Then
                  GoTo ExitEnd
               Else
                  dblB1009 = dblB1009 + Val(RsTemp.Fields("sa07")) '日
                  dblB1010 = dblB1010 + Val(RsTemp.Fields("sa08")) '時
                  PUB_ChkSerialRest_ToSir = True
                  'Add By Sindy 2023/9/15 記錄起來,後續程式使用
                  bolFind = False
                  If RsTemp.Fields("B1002") = 表單類別_請假 Then
                     strB1008 = "" & RsTemp.Fields("B1008")
                  Else
                     strB1008 = "" & RsTemp.Fields("B1014")
                  End If
                  For ii = 1 To intCnt
                     If g_strB1002(ii) = RsTemp.Fields("B1002") _
                        And g_strB1008(ii) = strB1008 Then
                        g_strDay(ii) = g_strDay(ii) + Val(RsTemp.Fields("sa07")) + (Val(RsTemp.Fields("sa08")) / PUB_intWkHour)
                        bolFind = True
                        Exit For
                     End If
                  Next
                  If bolFind = False Then
                     intCnt = intCnt + 1
                     g_strB1002(intCnt) = RsTemp.Fields("B1002")
                     g_strB1008(intCnt) = strB1008
                     g_strDay(intCnt) = Val(RsTemp.Fields("sa07")) + (Val(RsTemp.Fields("sa08")) / PUB_intWkHour)
                  End If
                  '2023/9/15 END
               End If
            Else
               dblB1009 = dblB1009 + Val("" & RsTemp.Fields("B1009")) '日
               dblB1010 = dblB1010 + Val("" & RsTemp.Fields("B1010")) '時
               PUB_ChkSerialRest_ToSir = True
               'Add By Sindy 2023/9/15 記錄起來,後續程式使用
               bolFind = False
               If RsTemp.Fields("B1002") = 表單類別_請假 Then
                  strB1008 = "" & RsTemp.Fields("B1008")
               Else
                  strB1008 = "" & RsTemp.Fields("B1014")
               End If
               For ii = 1 To intCnt
                  If g_strB1002(ii) = RsTemp.Fields("B1002") _
                     And g_strB1008(ii) = strB1008 Then
                     g_strDay(ii) = g_strDay(ii) + Val("" & RsTemp.Fields("B1009")) + (Val("" & RsTemp.Fields("B1010")) / PUB_intWkHour)
                     bolFind = True
                     Exit For
                  End If
               Next
               If bolFind = False Then
                  intCnt = intCnt + 1
                  g_strB1002(intCnt) = RsTemp.Fields("B1002")
                  g_strB1008(intCnt) = strB1008
                  g_strDay(intCnt) = Val("" & RsTemp.Fields("B1009")) + (Val("" & RsTemp.Fields("B1010")) / PUB_intWkHour)
               End If
               '2023/9/15 END
            End If
         End If
         If InStr(strID, RsTemp.Fields("B1001")) = 0 Then
            strQuyDateE = RsTemp.Fields("B1006")
            If strID = "" Then
               strID = RsTemp.Fields("B1001")
            Else
               strID = strID & "," & RsTemp.Fields("B1001")
            End If
         End If
         GoTo Count_Next1
      End If
      
'   '表單類別_出差
'   Else
'      Do While ChkWorkDay(strDate) = False
'         strDate = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(strDate)))
'      Loop
'      '讀取符合請假區間的資料
'      strSql = "select * from abs010,staff_busi_trip where b1003='" & strB1003 & "'" & _
'        " and b1002='" & strB1002 & "' and b1001=SB10(+)" & _
'        " and " & strDate & " between b1004 and b1006"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         If InStr(strID, RsTemp.Fields("B1001")) = 0 Then
'            If "" & RsTemp.Fields("SB10") <> "" Then '已收錄至人事系統
'               If Not (Val(strDate) >= Val(RsTemp.Fields("sb02")) And _
'                      Val(strDate) <= Val(RsTemp.Fields("sb04"))) Then
'                  GoTo ExitEnd
'               Else
'                  dblB1009 = dblB1009 + Val(RsTemp.Fields("sb06")) '日
'                  dblB1010 = dblB1010 + Val(RsTemp.Fields("sb07")) '時
'                  PUB_ChkSerialRest_ToSir = True
'               End If
'            Else
'               dblB1009 = dblB1009 + Val(RsTemp.Fields("B1009")) '日
'               dblB1010 = dblB1010 + Val(RsTemp.Fields("B1010")) '時
'               PUB_ChkSerialRest_ToSir = True
'            End If
'         End If
'         If InStr(strID, RsTemp.Fields("B1001")) = 0 Then
'            strQuyDateE = RsTemp.Fields("B1006")
'            If strID = "" Then
'               strID = RsTemp.Fields("B1001")
'            Else
'               strID = strID & "," & RsTemp.Fields("B1001")
'            End If
'         End If
'         GoTo Count_Next1
'      End If
'   End If
ExitEnd:
   
   If PUB_ChkSerialRest_ToSir = True Then
      dblDay = dblB1009 + (dblB1010 / PUB_intWkHour)
      If dblDay < 3 Then
         PUB_ChkSerialRest_ToSir = False
      Else
         If bolShowMsg = True Then
            If Val(strQuyDateS) > Val(strB1004) Or Val(strQuyDateS) = 0 Then strQuyDateS = strB1004
            If Val(strQuyDateE) < Val(strB1006) Or Val(strQuyDateE) = 0 Then strQuyDateE = strB1006
            If MsgBox("本假單有連續假單存在，合計超過3日(含)以上，請到「出缺勤查詢」查看。" & vbCrLf & vbCrLf & _
               "現在是否要進入「出缺勤查詢」？", vbYesNo, "詢問") = vbYes Then
               Set oForm = Forms(0).GetForm("frm180301")
               With oForm
                  .Hide
                  .Option1(0).Value = True '明細
                  .txtDate(0) = Val(strQuyDateS) - 19110000 '日期
                  .txtDate(1) = Val(strQuyDateE) - 19110000
                  'Modify By Sindy 2023/12/29
                  'Modify By Sindy 2025/5/20
'                  If strSrvDate(1) >= 新部門啟用日 Then
                     strExc(10) = PUB_GetST93(strB1003)
                     For ii = 0 To .cboDept(0).ListCount - 1
                        If InStr(.cboDept(0).List(ii), strExc(10)) > 0 Then
                           .cboDept(0).ListIndex = ii
                           .cboDept(1).ListIndex = ii
                           Exit For
                        End If
                     Next ii
'                     .txtDept(0) = PUB_GetST93(strB1003) '部門
'                     .txtDept(1) = PUB_GetST93(strB1003)
'                  Else
'                  '2023/12/29 END
'                     .txtDept(0) = PUB_GetST03(strB1003) '部門
'                     .txtDept(1) = PUB_GetST03(strB1003)
'                  End If
                  '2025/5/20 END
                  .txtB1003(0) = strB1003 '請假人員
                  .txtB1003(1) = strB1003
                  .CboB1002 = strB1002 '表單類別
                  .txtST06(0) = PUB_GetST06(strB1003) '所別
                  .txtST06(1) = PUB_GetST06(strB1003)
                  .cmdOK(1).Tag = "SysQuery"
                  .m_IsAbsBossST03 = "" 'Add By Sindy 2025/5/20 代理簽核主管(江協理)沒有當事人(林文雄)的簽核權限
                  Call .cmdok_Click(0) '查詢
               End With
            End If
         End If
      End If
   End If
   
   Set oForm = Nothing
   Exit Function
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Added by Morgan 2015/1/28 檢查是否超過65歲
Public Function PUB_ChkOver65(pNo As String) As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select st23 from staff where st01='" & pNo & "' and st23<=" & CompDate(0, -65, strSrvDate(1))
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      PUB_ChkOver65 = True
   End If
   Set rsQuery = Nothing
End Function

'Add By Sindy 2015/12/25 出缺勤表單的提醒訊息
'strFormType : 01.請假單 02.加班單 03.出差單
'strToSeePeople : 0.當事人 1.簽核人員
Public Function PUB_PerFormRemindMsg(strFormType As String, strToSeePeople As String, StrST01 As String, _
   Optional txtB1004 As String, Optional txtB101213 As String, Optional bolShowMsg As Boolean = True) As String
   
Dim dblHour As Double
   
   If strFormType = 表單類別_加班 Then
      'Add By Sindy 2015/12/25 增加檢查同仁加班合計是否有超過40小時
      'strSql = "select sum(nvl(So05,0))+sum(nvl(So06,0)) from staff_overtime "
      strSql = "select So02,nvl(So05,0) So05,nvl(So06,0) So06 from staff_overtime " & _
                "where So01='" & StrST01 & "' " & _
                  "and So02 between " & CStr(Val(Left(txtB1004, Len(txtB1004) - 2)) + 191100) & "01 and " & CStr(Val(Left(txtB1004, Len(txtB1004) - 2)) + 191100) & "31"
      intI = 1
      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'Modify By Sindy 2016/12/26
         '逐天累計
         'dblHour = Val("" & adoRecordset.Fields(0))
         'If Val(IIf(txtB101213 = "", 0, txtB101213)) > 0 Then dblHour = dblHour + Val(txtB101213)
         adoRecordset.MoveFirst
         Do While Not adoRecordset.EOF
            If ChkWorkDay(adoRecordset.Fields("So02"), StrST01, True) = False Then '假日加班
               If Weekday(Format(adoRecordset.Fields("So02"), "####-##-##")) <> 7 And _
                  adoRecordset.Fields("So02") >= 20161223 Then '非週六  (2016/12/23開始實施)
                  '假日且非週六,要扣8個小時
                  dblHour = dblHour + (Val("" & adoRecordset.Fields("So06")) - 8)
               Else
                  dblHour = dblHour + Val("" & adoRecordset.Fields("So06"))
               End If
            Else '平日加班
               dblHour = dblHour + Val("" & adoRecordset.Fields("So05"))
            End If
            adoRecordset.MoveNext
         Loop
         
         'Modify By Sindy 2021/7/23
         If Val(txtB101213) > 0 And txtB1004 <> "" Then
            '檢查上列的統計是否已包含此次加班的時數
            strSql = "select So02 from staff_overtime " & _
                      "where So01='" & StrST01 & "' " & _
                        "and (So02 between " & CStr(Val(Left(txtB1004, Len(txtB1004) - 2)) + 191100) & "01 and " & CStr(Val(Left(txtB1004, Len(txtB1004) - 2)) + 191100) & "31)" & _
                        " and So02=" & DBDATE(txtB1004)
            intI = 1
            Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Then
            '2021/7/23 END
               '假日加班且非週六,要扣8個小時
               If ChkWorkDay(DBDATE(txtB1004), StrST01, True) = False And _
                  Weekday(Format(DBDATE(txtB1004), "####-##-##")) <> 7 And _
                  DBDATE(txtB1004) >= 20161223 Then '非週六  (2016/12/23開始實施)
                  dblHour = dblHour + (Val(txtB101213) - 8)
               Else
                  dblHour = dblHour + Val(txtB101213)
               End If
            End If
         End If
         '2016/12/26 END
         
         If Val(dblHour) > 40 Then
            If strToSeePeople = "0" Then '同仁
               PUB_PerFormRemindMsg = "您加班已超過40小時，請注意勿違反勞基法每月加班不得超過46小時之規定，謝謝！"
            Else '簽核主管
               PUB_PerFormRemindMsg = "此同仁目前加班時數已超過40小時，請注意勿違反勞基法每月加班不得超過46小時之規定！"
            End If
            If bolShowMsg = True Then
               MsgBox PUB_PerFormRemindMsg, vbExclamation
            End If
         End If
      End If
      '2015/12/25 END
   End If
End Function

'Added by Morgan 2016/1/22
'以所得人代碼/員工號取得身分證號
Public Function PUB_GetIdByNum(pNum As String) As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select oi02 from otherincomer where oi01='" & pNum & "' and oi02 is not null" & _
      " union select st26 from staff where st01='" & pNum & "' and st26 is not null"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      PUB_GetIdByNum = rsQuery(0)
   End If
   Set rsQuery = Nothing
End Function

'Added by Morgan 2016/1/22
'以所得人代碼/員工號及日期取得在職員工號
Public Function PUB_GetStaffNoByNumDate(pNum As String, pDate As String) As String
   Dim strNo As String
   
   strNo = PUB_GetIdByNum(pNum)
   If strNo <> "" Then
      PUB_GetStaffNoByNumDate = PUB_GetStaffNoByIdDate(strNo, pDate)
   End If
End Function

'Added by Morgan 2016/1/22
'以身分證號及日期檢查取得職員工號
'pID:身分證號, pDate:日期, 回傳:在職員工號
Public Function PUB_GetStaffNoByIdDate(pID As String, pDate As String) As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset, rsQuery2 As ADODB.Recordset
   
   '檢查ID是否有對應的員工
   stSQL = "select st01,st13 from staff  where st26='" & pID & "' and st01>'6' and st01<'F' and st13<=" & DBDATE(pDate)
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      Do While Not .EOF
         If .Fields("st13") <= DBDATE(pDate) Then
            '檢查最後異動
            stSQL = "select st13,sc02,sc03 from staff,staff_change" & _
               " where st01='" & .Fields("st01") & "' and sc01(+)=st01 and sc03 in ('02','03','04','08','09','10')" & _
               " and sc02<=" & DBDATE(pDate) & " order by sc02 desc"
            intQ = 1
            Set rsQuery2 = ClsLawReadRstMsg(intQ, stSQL)
            If intQ = 1 Then
               '復職->在職
               If rsQuery2("sc03") = "02" Then
                  PUB_GetStaffNoByIdDate = .Fields("st01")
                  Exit Do
               End If
            '無異動->在職
            Else
               PUB_GetStaffNoByIdDate = .Fields("st01")
               Exit Do
            End If
         End If
         .MoveNext
      Loop
      End With
   End If
   
   Set rsQuery = Nothing
   Set rsQuery2 = Nothing
End Function

'Add By Sindy 2017/1/6
'Modify By Sindy 2012/9/27 + Optional ByVal bolAutoBatch As Boolean = False
Public Function CountRestDay(strST13 As String, strST40 As String, strYV01 As String, _
   Optional ByRef strFormulation As String, Optional ByRef m_Type As String, _
   Optional ByVal bolAutoBatch As Boolean = False) As String
Dim ii As Integer
Dim strMonthWorkDate As String
Dim intMonth As Integer
Dim strFormulationA As String, strFormulationB As String
Dim strStarDate As String, strEndDate As String
Dim intYearDay1 As Integer
Dim intCountDay As Integer, m_DaysNew As Double
Dim varArr As Variant
   
   If Len(strYV01) < 4 Then
      strYV01 = Val(strYV01) + 1911
   End If
   
   '當年滿1年者
   If DBDATE(DateAdd("m", 12, ChangeWStringToWDateString(strST13))) <= strSrvDate(1) Then
      'Modified by Morgan 2024/1/11 修正日期轉字串會因顯示格式而不同問題
      'If Val(Left(DateAdd("m", 12, ChangeWStringToWDateString(strST13)), 4)) = Val(strYV01) Then
      If Val(Left(DBDATE(DateAdd("m", 12, ChangeWStringToWDateString(strST13))), 4)) = Val(strYV01) Then
      'end 2024/1/11
         '*************************************************************************
         'A公式:
         '*************************************************************************
         '檢查到職日是否為當月工作的第1天,若是該月則足月計算,否則不列入該月統計
         For ii = 1 To 31
            strMonthWorkDate = Left(strST13, 6) & Format(ii, "00")
            If IsDate(ChangeWStringToWDateString(strMonthWorkDate)) = True _
               And ChkWorkDay(strMonthWorkDate) = True Then
               Exit For
            End If
         Next ii
         If strST13 <= strMonthWorkDate Then
            intMonth = 12 - Val(Mid(strST13, 5, 2)) + 1
         Else
            intMonth = 12 - Val(Mid(strST13, 5, 2))
         End If
         CountRestDay = (0.5 * intMonth) + 1
         strFormulationA = "計算公式 = ((0.5 * " & intMonth & ") + 1)"
         '*************************************************************************
         'B公式:
         '*************************************************************************
         strStarDate = CStr(strYV01) & "/01/01"
         strEndDate = CStr(strYV01) & "/12/31"
         intYearDay1 = DateDiff("d", strStarDate, strEndDate) + 1 '今年總天數
         'intYearDay2 = DateDiff("d", CStr(Val(strYV01) + 1) & "/01/01", CStr(Val(strYV01) + 1) & "/12/31") + 1 '明年總天數
         '到職日~12/31
         intCountDay = DateDiff("d", CStr(strYV01) & Mid(ChangeWStringToWDateString(strST13), 5), strEndDate) + 1
         'Modify By Sindy 2019/12/30 檢查若年總天數為366則天數減1
         If intYearDay1 = 366 Then
            intCountDay = intCountDay - 1
         End If
         '2019/12/30 END
         'Modify By Sindy 2019/12/30 劉經理說固定365計算
         m_DaysNew = Round(7 * (intCountDay / 365), 2)
         'Modify By Sindy 2019/12/30 劉經理說固定365計算
         strFormulationB = "計算公式 = (7 * (" & intCountDay & " / " & 365 & "))"
         If InStr(m_DaysNew, ".") > 0 Then
            varArr = Split(m_DaysNew, ".")
            '滿半日不滿一日者,以一日加計
            If Val("0." & CStr(varArr(1))) > 0.5 Then
               m_DaysNew = Int(m_DaysNew) + 1
            '不滿0.5日者,以0.5日加計
            ElseIf Val("0." & CStr(varArr(1))) < 0.5 Then
               m_DaysNew = Int(m_DaysNew) + 0.5
            End If
         End If
         '*************************************************************************
         '比較A式和B式,擇天數多者
         '*************************************************************************
         If Val(m_DaysNew) > Val(CountRestDay) Then
            m_Type = "B" '週年制
            CountRestDay = m_DaysNew
            strFormulation = strFormulationB
         Else
            m_Type = "A" '曆年制
            strFormulation = strFormulationA
         End If
         '*************************************************************************
         '到職日滿半年落在計算當年者,特別假要累計
         '*************************************************************************
         'Add By Sindy 2020/10/14 (傳國峻)滿一年時,可能遇到前年底半年特休往後留存,所以直接累計
         If bolAutoBatch = True Then
            If Val(strST40) > 0 Then
               CountRestDay = CountRestDay + Val(strST40)
               strFormulation = strFormulation & " + " & CStr(Val(strST40))
            End If
         Else
         '2020/10/14 END
            'Modified by Morgan 2024/1/11 修正日期轉字串會因顯示格式而不同問題
            'If Val(Left(DateAdd("m", 6, ChangeWStringToWDateString(strST13)), 4)) = Val(strYV01) Then
            If Val(Left(DBDATE(DateAdd("m", 6, ChangeWStringToWDateString(strST13))), 4)) = Val(strYV01) Then
            'end 2024/1/11
               If Val(strST40) > 0 Then
                  CountRestDay = CountRestDay + Val(strST40)
                  strFormulation = strFormulation & " + " & CStr(Val(strST40))
               End If
            End If
         End If
      End If
   '當年滿半年者
   ElseIf DBDATE(DateAdd("m", 6, ChangeWStringToWDateString(strST13))) <= strSrvDate(1) Then
      'Modified by Morgan 2024/1/11 修正日期轉字串會因顯示格式而不同問題
      'If Val(Left(DateAdd("m", 6, ChangeWStringToWDateString(strST13)), 4)) = Val(strYV01) Then
      If Val(Left(DBDATE(DateAdd("m", 6, ChangeWStringToWDateString(strST13))), 4)) = Val(strYV01) Then
      'end 2024/1/11
         CountRestDay = 3
      End If
   End If
End Function

'工作時數
'Added by Morgan 2013/8/8
'Modified by Morgan 2017/7/10 配合薪資明細查詢改共用函數並增加 pYM 參數
Public Function GetDaiyHour(pStaffNo As String, Optional pYM As String = "999999") As Integer
'Modified by Morgan 2017/9/27
'   If pStaffNo = "99029" Then
'      'Modified by Morgan 2017/7/10
'      'GetDaiyHour = cDailyHr99029
'      If pYM >= "201110" Then
'         GetDaiyHour = 6
'      ElseIf pYM >= "201103" Then
'         GetDaiyHour = 5
'      Else
'         GetDaiyHour = 4
'      End If
'      'end 2017/7/10
'
'   'Removed by Morgan 2017/7/11 105年薪資查詢上線後 84043,73029 皆無加班紀錄
'   ''101/7月起尤春彬 84043 工作時數改 4 小時
'   ''Removed by Morgan 2016/4/14 尤春彬84043 105/3/1 改回全職
'   ''ElseIf pStaffNo = "84043" Then
'   ''   GetDaiyHour = cDailyHr84043
'   ''end 2016/4/14
'   'ElseIf pStaffNo = "73029" Then
'   '   GetDaiyHour = cDailyHr73029
'   'end 2017/7/11
'
'   Else
'      GetDaiyHour = 8
'   End If
   Pub_GetSpecWorkHour pStaffNo, pYM & "01"
   GetDaiyHour = PUB_intWkHour
'end 2017/9/27
End Function

'計算時薪
'Added by Morgan 2008/12/25
'Modified by Morgan 2017/7/10 配合薪資明細查詢改共用函數並增加 pYM 參數
'Modified by Morgan 2019/7/1 pYM-->pDate 時薪改以日期判斷(原有調薪時以平均薪資計算)--婧瑄
Public Function GetHourPay(pUserNo As String, Optional pDate As String, Optional pDailyHours As Integer = 8) As Double
   Dim stSQL As String, intR As Integer, adoRst As New ADODB.Recordset
   Dim dblSalary As Double
   
'Modified by Morgan 2019/7/1
'108/5月薪資的加班費及缺勤扣款計算修改為以當日的薪資逐日計算(非1號調薪的計算規則調整)
'上線前為維持相容性先保留舊程式
If Len(pDate) = 6 Then

   '判斷該月份有異動(非1號)
   'stSQL = "select SL02 from salarylog where SL01='" & pUserNo & "' and SL02<>" & pDate & "01 and SUBSTR(SL02,1,6)=" & pDate & " order by 1 asc"
   
   stSQL = "select SL02,nvl(sc02,st13) d1,st51 d2 from STAFF" & _
      ",(select SL01,SL02 from salarylog where SL01='" & pUserNo & "'" & _
      " and SUBSTR(SL02,1,6)=" & pDate & " and SL02<>" & pDate & "01) a" & _
      ",(select sc01,sc02 from Staff_Change where sc01='" & pUserNo & "'" & _
      " and substr(sc02,1,6)=" & pDate & " and sc03='02') b" & _
      " where ST01='" & pUserNo & "' and SL01(+)=ST01 and sc01(+)=ST01" & _
      " order by SL02 asc"
      
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL, , True)
   If intR = 1 Then
      '沒異動或當月到(復)職,抓基本薪資資料
      'modify by sonia 2013/11/22 總經理指示缺勤及加班費之時薪計算要含職務津貼故加入SD21,SD30
      If IsNull(adoRst("SL02")) Or Left(adoRst("d1"), 6) = pDate Then
         'Modified by Morgan 2017/7/11
         'stSQL = "select nvl(SD20,0)+nvl(SD21,0)+nvl(SD23,0)+nvl(SD29,0)+nvl(SD30,0)+nvl(SD32,0)" & _
            " from salarydata a where SD01='" & pUserNo & "'"
         'Modified by Morgan 2017/7/11 配合薪資明細查詢改都抓異動
         'Modify By Sindy 2020/6/25 + 證照津貼 +nvl(b.SL39,0)
         stSQL = "select nvl(b.SL11,0)+nvl(b.SL12,0)+nvl(b.SL39,0)+nvl(b.SL14,0)+nvl(b.SL19,0)+nvl(b.SL20,0)+nvl(b.SL22,0)" & _
            " from salarylog b where SL01='" & pUserNo & "' and SL02=(select max(c.SL02) from salarylog c" & _
            " where c.SL01=b.SL01 and c.SL02<=" & pDate & "31)"
         
      '當月有異動
      Else
         'modify by sonia 2013/11/22 總經理指示缺勤及加班費之時薪計算要含職務津貼故加入SD21,SD30,SL12,SL20
         'Modified by Morgan 2017/7/11 配合薪資明細查詢改都抓異動
         'stSQL = " select 0.5*(nvl(a.SD20,0)+nvl(a.SD21,0)+nvl(a.SD23,0)+nvl(a.SD29,0)+nvl(a.SD30,0)+nvl(a.SD32,0)" & _
            "+nvl(b.SL11,0)+nvl(b.SL12,0)+nvl(b.SL14,0)+nvl(b.SL19,0)+nvl(b.SL20,0)+nvl(b.SL22,0))" & _
            " from salarydata a,salarylog b where a.SD01='" & pUserNo & "'" & _
            " and b.SL01(+)=a.SD01" & _
            " and b.SL02=(select max(c.SL02) from salarylog c" & _
            " where c.SL01=a.SD01 and c.SL02<" & pDate * 100 & ")"
         'Modify By Sindy 2020/6/25 + 證照津貼
         stSQL = " select 0.5*(nvl(a.SL11,0)+nvl(a.SL12,0)+nvl(a.SL39,0)+nvl(a.SL14,0)+nvl(a.SL19,0)+nvl(a.SL20,0)+nvl(a.SL22,0)" & _
            "+nvl(b.SL11,0)+nvl(b.SL12,0)+nvl(b.SL39,0)+nvl(b.SL14,0)+nvl(b.SL19,0)+nvl(b.SL20,0)+nvl(b.SL22,0))" & _
            " from salarylog a,salarylog b where a.SL01='" & pUserNo & "'" & _
            " and a.SL02=(select max(c.SL02) from salarylog c" & _
            " where c.SL01=a.SL01 and c.SL02<=" & pDate & "31)" & _
            " and b.SL01(+)=a.SL01" & _
            " and b.SL02=(select max(c.SL02) from salarylog c" & _
            " where c.SL01=a.SL01 and c.SL02<" & pDate & "00)"
      End If
   End If
   
Else
   'Modify By Sindy 2020/6/25 + 證照津貼
   stSQL = "select nvl(b.SL11,0)+nvl(b.SL12,0)+nvl(b.SL39,0)+nvl(b.SL14,0)+nvl(b.SL19,0)+nvl(b.SL20,0)+nvl(b.SL22,0)" & _
      " from salarylog b where SL01='" & pUserNo & "' and SL02=(select max(c.SL02) from salarylog c" & _
      " where c.SL01=b.SL01 and c.SL02<=" & pDate & ")"
End If
'end 2019/7/1
         
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      dblSalary = Val("" & adoRst.Fields(0))
   End If
   'Modify by Morgan 2010/7/14
   'GetHourPay = dblSalary / 30 / 8
   'Modified by Morgan 2017/7/10
   'iDailyHours = GetDaiyHour(pUserNo)
   'GetHourPay = dblSalary / 30 / iDailyHours
   GetHourPay = dblSalary / 30 / pDailyHours
   'end 2017/7/10
   'end 2010/7/14
   Set adoRst = Nothing
End Function

'計算病假時數
'Add by Morgan 2009/3/9
'Modified by Morgan 2014/12/12 +p_IsGirlSick: True=生理假, False=病假
'Modified by Morgan 2017/7/10 配合薪資明細查詢改共用函數
Public Function GetSickHour(p_StaffNo As String, p_FromDate As Double, p_ToDate As Double, Optional p_IsGirlSick As Boolean = False) As Double
   Dim stSQL As String, adoRst As ADODB.Recordset, intR As Integer
   
   'GetDaiyHour(stLstNo)
   'Modified by Morgan 2020/10/7 修正特殊工時(非8小時)問題
   'stSQL = "select nvl(sum(" & GetDaiyHourSQL("sa01") & "*nvl(sa07,0)+nvl(sa08,0)),0) y2" & _
      " from Staff_Absence where sa01='" & p_StaffNo & "'" & _
      " and sa02>=" & p_FromDate & " and sa02<=" & p_ToDate & " and sa06='" & IIf(p_IsGirlSick = True, "20", "06") & "'"
   stSQL = "select nvl(sum(" & GetDaiyHour(p_StaffNo) & "*nvl(sa07,0)+nvl(sa08,0)),0) y2" & _
      " from Staff_Absence where sa01='" & p_StaffNo & "'" & _
      " and sa02>=" & p_FromDate & " and sa02<=" & p_ToDate & " and sa06='" & IIf(p_IsGirlSick = True, "20", "06") & "'"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      GetSickHour = adoRst.Fields(0)
   End If
   Set adoRst = Nothing
End Function

'工作時數語法
'Added by Morgan 2013/8/8
'Modified by Morgan 2017/7/10 配合薪資明細查詢改共用函數
'Removed by Morgan 2020/10/7 取消,相關程式改寫用函數判斷
'Public Function GetDaiyHourSQL(PColName As String) As String
'   'Modified by Morgan 2016/4/14 尤春彬84043 105/3/1 改回全職
'   'Modified by Morgan 2017/7/10 廖宗岳73029 已離職
'   'GetDaiyHourSQL = "decode(" & PColName & ",'99029'," & cDailyHr99029 & ",'73029'," & cDailyHr73029 & ",8)"
'   GetDaiyHourSQL = "decode(" & PColName & ",'99029'," & GetDaiyHour("99029") & ",8)"
'   'end 2017/7/10
'End Function

'Add By Sindy 2017/10/24
'取得職務代理人
'Modify By Sindy 2018/4/17 + Optional bolChkRest As Boolean = True
Public Function PUB_GetWorkDeputyEmp(strUser As String, Optional bolChkRest As Boolean = True) As String
Dim m_ABS001_1 As String
Dim m_ABS001_2 As String
Dim m_ABS001_3 As String
Dim strData As String, strTemp As Variant
Dim j As Integer, i As Integer
   
   PUB_GetWorkDeputyEmp = ""
   '(1)案件職代
   Call GetABS001_3(strUser, m_ABS001_1, m_ABS001_2, m_ABS001_3, "")
   If m_ABS001_1 = "" Then
      '(2)一般人事職代
      Call GetABS001_1(strUser, m_ABS001_1, m_ABS001_2, m_ABS001_3)
   End If
   For j = 1 To 3 '有3組職代
      strData = ""
      If j = 1 And m_ABS001_1 <> "" Then strData = m_ABS001_1
      If j = 2 And m_ABS001_2 <> "" Then strData = m_ABS001_2
      If j = 3 And m_ABS001_3 <> "" Then strData = m_ABS001_3
      If strData <> "" Then
         '檢查取得的職代是否請假
         strTemp = Split(strData, ",")
         For i = 0 To UBound(strTemp)
            'Modify By Sindy 2018/4/17
'            If CheckIsPersonRest(CStr(strTemp(i)), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False Then
'               PUB_GetWorkDeputyEmp = strTemp(i)
'               Exit Function
'            End If
            If bolChkRest = True Then
               If CheckIsPersonRest(CStr(strTemp(i)), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = False Then
                  PUB_GetWorkDeputyEmp = strTemp(i)
                  Exit Function
               End If
            Else
               PUB_GetWorkDeputyEmp = strTemp(i)
               Exit Function
            End If
            '2018/4/17 END
         Next i
      End If
   Next j
End Function

'Add by Amy 2017/12/27 取得傳入之人員是誰的職代
'intChoose:1-全部/2-案件/3-人事
'Modified by Lydia 2017/12/29 +傳入人員和職代皆在職(dblChk), 是否取得假單的職代 (bAddRest)
'Modified by Lydia 2018/03/07 +是否取得審核主管bAddMan
Public Function GetSOAgent(ByVal intChoose As Integer, ByVal StrST01 As String, Optional ByVal dblChk As Boolean = False, Optional ByVal bAddRest As Boolean = False, Optional ByVal bAddMan As Boolean = False) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strB0101 As String
    Dim intQ As Integer
    
    GetSOAgent = ""
    If intChoose = 1 Or intChoose = 2 Then
        '案件職代
        'Modified by Lydia 2017/12/29 判斷在職
'        strQ = "       Select B0101 From ABS001,Staff Where B0117='" & strST01 & "' And B0117=ST01(+) And ST04='1' " & _
'                "Union Select B0101 From ABS001,Staff Where B0119='" & strST01 & "' And B0119=ST01(+) And ST04='1' " & _
'                "Union Select B0101 From ABS001,Staff Where B0121='" & strST01 & "' And B0121=ST01(+) And ST04='1' " & _
'                "Union Select B0101 From ABS001,Staff Where B0123='" & strST01 & "' And B0123=ST01(+) And ST04='1' "
        strQ = "       Select B0101 From ABS001,Staff S1, Staff S2 Where B0117='" & StrST01 & "' And B0117=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "") & _
                "Union Select B0101 From ABS001,Staff S1, Staff S2 Where B0119='" & StrST01 & "' And B0119=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "") & _
                "Union Select B0101 From ABS001,Staff S1, Staff S2 Where B0121='" & StrST01 & "' And B0121=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "") & _
                "Union Select B0101 From ABS001,Staff S1, Staff S2 Where B0123='" & StrST01 & "' And B0123=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "")
    End If
    If intChoose = 1 Or intChoose = 3 Then
        '人事職代
        'Modified by Lydia 2017/12/29 判斷在職
        'strQ = IIf(intChoose = 1, strQ & " Union", "") & _
                "          Select B0101 From ABS001,Staff Where B0102='" & strST01 & "' And B0102=ST01(+) And ST04='1' " & _
                "Union Select B0101 From ABS001,Staff Where B0103='" & strST01 & "' And B0103=ST01(+) And ST04='1' " & _
                "Union Select B0101 From ABS001,Staff Where B0104='" & strST01 & "' And B0104=ST01(+) And ST04='1' " & _
                "Union Select B0101 From ABS001,Staff Where B0105='" & strST01 & "' And B0105=ST01(+) And ST04='1' " & _
                "Union Select B0101 From ABS001,Staff Where B0106='" & strST01 & "' And B0106=ST01(+) And ST04='1' " & _
                "Union Select B0101 From ABS001,Staff Where B0107='" & strST01 & "' And B0107=ST01(+) And ST04='1' "
        strQ = IIf(intChoose = 1, strQ & " Union", "") & _
                "          Select B0101 From ABS001,Staff S1, Staff S2 Where B0102='" & StrST01 & "' And B0102=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "") & _
                "Union Select B0101 From ABS001,Staff S1, Staff S2 Where B0103='" & StrST01 & "' And B0103=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "") & _
                "Union Select B0101 From ABS001,Staff S1, Staff S2 Where B0104='" & StrST01 & "' And B0104=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "") & _
                "Union Select B0101 From ABS001,Staff S1, Staff S2 Where B0105='" & StrST01 & "' And B0105=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "") & _
                "Union Select B0101 From ABS001,Staff S1, Staff S2 Where B0106='" & StrST01 & "' And B0106=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "") & _
                "Union Select B0101 From ABS001,Staff S1, Staff S2 Where B0107='" & StrST01 & "' And B0107=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "")
    End If
    'Added by Lydia 2018/03/07 審核主管
    If bAddMan = True Then
        strQ = strQ & " Union Select B0101 From ABS001,Staff S1, Staff S2 Where B0108='" & StrST01 & "' And B0102=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "") & _
                "Union Select B0101 From ABS001,Staff S1, Staff S2 Where B0109='" & StrST01 & "' And B0103=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "") & _
                "Union Select B0101 From ABS001,Staff S1, Staff S2 Where B0110='" & StrST01 & "' And B0104=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "") & _
                "Union Select B0101 From ABS001,Staff S1, Staff S2 Where B0111='" & StrST01 & "' And B0105=S1.ST01(+) And S1.ST04='1'  And B0101=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "")
    End If
    'end 2018/03/07
    
    'Added by Lydia 2017/12/29 取得假單的職代
    If bAddRest = True Then
         strQ = strQ & "Union SELECT B1003 FROM ABS011,ABS010,STAFF S1,STAFF S2 " & _
            " where B1101 in( SELECT SA09 From staff_Absence WHERE " & strSrvDate(1) & " between SA02 and SA04" & _
            " Union SELECT SB10 From staff_busi_trip WHERE " & strSrvDate(1) & " between SB02 and SB04" & _
            ") and B1102='1' and B1104='" & StrST01 & "' AND B1104=S1.ST01(+) AND S1.ST04='1' AND B1003=S2.ST01(+) " & IIf(dblChk = True, " AND S2.ST04='1' ", "") & _
            " and B1101=B1001 and B1003 is not null" & _
            " group by B1003"
    End If
    'end 2017/12/29
    
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        Do While RsQ.EOF = False
            'Modified by Lydia 2017/12/29  修正並排除重複編號
            'strB0101 = "," & RsQ.Fields("B0101")
            If "" & RsQ.Fields("B0101") <> "" Then
                If strB0101 = "" Then
                    strB0101 = "," & RsQ.Fields("B0101")
                ElseIf InStr(strB0101, "" & RsQ.Fields("B0101")) = 0 Then
                    strB0101 = strB0101 & "," & RsQ.Fields("B0101")
                End If
            End If
            'end 2017/12/29
            RsQ.MoveNext
        Loop
        GetSOAgent = Mid(strB0101, 2)
    End If
    RsQ.Close
End Function

'Add By Sindy 2019/6/18
'計算年資和特別假天數
'm_Type = A.曆年制
'         B.週年制
Public Function PUB_GetSeniorityYearVacation(StrST01 As String, strYV01 As String, _
   Optional ByRef m_Year As String = "0", Optional ByRef m_Type As String = "", _
   Optional ByRef m_Formulation As String = "") As Double

Dim m_rs As New ADODB.Recordset
Dim strST13 As String
Dim strStarDate As String, strEndDate As String
Dim strStarDate_sub As String, strEndDate_sub As String
Dim m_Days As Double, m_Days2 As Double
Dim intYearDay1 As Integer, intYearDay2 As Integer
Dim intCountDay1 As Integer, intCountDay2 As Integer
Dim m_DaysNew As Double
Dim varArr As Variant
Dim LongWorkDay As Long, m_Year1_Std As String
Dim strBackDt As String
Dim m_YearDay As Long '年度總天數
Dim strTempDate As String, strST40 As String
Dim strFormulationA As String
Dim strNote As String
Dim strBackTaieDate As String
Dim strCountDate As String
Dim strTemp As String
Dim intRow As Integer
Dim strDateNote As String
   
   PUB_GetSeniorityYearVacation = 0
   m_Formulation = "": strFormulationA = "": strTemp = ""
   strDateNote = ""
   
   strSql = "select * from staff " & _
            "where ST01='" & StrST01 & "' "
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If m_rs.RecordCount <> 0 Then
      strST13 = "" & m_rs.Fields("st13")
      strST40 = "" & m_rs.Fields("st40")
      If strST13 = "" Then
         Exit Function
      End If
   Else
      Exit Function
   End If
   
   strBackTaieDate = Pub_BackTaieToDate(StrST01, strYV01, strNote)
   If Val(strBackTaieDate) > 0 Then
      strCountDate = strBackTaieDate '用留職停薪後起算日計算
   Else
      strCountDate = strST13 '用到職日計算
   End If
   
   If Len(strYV01) < 4 Then
      strYV01 = Val(strYV01) + 1911
   End If
   strEndDate = CStr(Val(CStr(Val(strYV01) - 1) & "1231"))
   
   m_Days = 0: m_Days2 = 0
   '計算年資
   m_Year = Trim(CalYear(StrST01, strEndDate))
   If Val(m_Year) <= 0 Then
      m_Year = "0"
   Else
      'A.曆年制
      If Val(m_Year) >= 10 Then '滿 10 年
         m_Days = 16 + (Int(m_Year) - 10) '滿10年者以16天起算
         If m_Days >= 30 Then
            m_Days = 30
            m_Days2 = 30
         Else
            m_Days2 = m_Days + 1
         End If
      ElseIf Val(m_Year) >= 5 Then    '滿 5 年
         m_Days = 15
         If m_Year = 9 Then
            m_Days2 = 16
         Else
            m_Days2 = m_Days
         End If
      ElseIf Val(m_Year) >= 3 Then ' 滿 3 年
         m_Days = 14
         If m_Year = 4 Then
            m_Days2 = 15
         Else
            m_Days2 = m_Days
         End If
      ElseIf Val(m_Year) >= 2 Then ' 滿 2 年
         m_Days = 10
         m_Days2 = 14
      ElseIf Val(m_Year) >= 1 Then  '滿1年
         m_Days = 7
         m_Days2 = 10
'      'Modify by Sindy 2019/6/19
'      Else '未滿1年
'         m_Days = CountRestDay(strCountDate, strST40, strYV01, strFormulationA)
'         '當年滿1年者
'         If Val(Left(DateAdd("m", 12, ChangeWStringToWDateString(strCountDate)), 4)) = Val(strYV01) Then
'            m_Days2 = 7
'         End If
''      'Modify by Sindy 2017/1/3 未滿一年都是每天跑批次計算特別假
''      ElseIf Val(m_Year) = 0.5 Then '滿6月
''         m_Days = 3
''         m_Days2 = 0
      End If
      If m_Days < 0 Then m_Days = 0
      
'                   '到職日為前一年者
'                   If CheckStr(m_rs.Fields("st13")) > Val((Val(textYV01_1) + 1911 - 1) & "0101") Then
'                        '2009/12/3 add by sonia 每年12/1到職者特別假給0.5天
'                        If CheckStr(m_rs.Fields("st13")) = Val((Val(textYV01_1) + 1911 - 1) & "1201") Then
'                           m_Days = 0.5
'                        '每年12/1以後到職者無特別假
'                        ElseIf CheckStr(m_rs.Fields("st13")) > Val((Val(textYV01_1) + 1911 - 1) & "1201") Then
'                           m_Days = 0
''                        Else
''                        '2009/12/3 end
''2009/12/15 CANCEL BY SONIA 上面已處理不必再做
''                           '取得計算年度之前年工作總時數
''                           m_StrSQL2 = "select sum(nvl(sm27,0))+31 " & _
''                                       "from salarymonth " & _
''                                       "where sm01='" & CheckStr(m_rs.Fields("st01")) & "' " & _
''                                       "and sm02>='" & CStr(Val(textYV01_1) + 1911 - 1) & "01" & "' " & _
''                                       "and sm02<='" & CStr(Val(textYV01_1) + 1911 - 1) & "12" & "' "
''                           If m_rs2.State = 1 Then m_rs2.Close
''                           m_rs2.CursorLocation = adUseClient
''                           m_rs2.Open m_StrSQL2, cnnConnection, adOpenStatic, adLockReadOnly
''                           If m_rs2.RecordCount <> 0 Then
''                               m_rs2.MoveFirst
''                               If m_YearDay <> CLng(CheckStr(m_rs2.Fields(0))) Then
''                                    '因前年未工作滿一年, 所以特別假要按照工作總時數計算
''                                    'm_Days = Round(m_Days * CLng(CheckStr(m_rs2.Fields(0))) / m_YearDay, 1)
''                                    m_Days = Round(7 * CLng(CheckStr(m_rs2.Fields(0))) / m_YearDay, 0)
''                               End If
''                           End If
''2009/12/15 end
'                        End If
'                   End If
      
      'Add By Sindy 2017/1/3 滿一年以上,當年復職者,依工作天數比例給假
      If Val(m_Year) >= 1 Then
         '取得計算年度之前年總天數
         If PUB_GetMonthDays((Val(strYV01) - 1), 2) = 28 Then
            m_YearDay = 365
         Else
            m_YearDay = 366
         End If
         '為當年復職者
         'Modify By Sindy 2025/1/2 修改 取消and substr(sc02,1,4)='" & CStr(Val(strYV01) - 1) & "' "
         ' A3021隆杉夢維 20250101才復職
         strStarDate = CStr(Val(strYV01) - 1) & "0101" '前年的第一天
         strEndDate = CStr(GetYearStdDay(CStr(strYV01))) '新年的第一個工作天
         strSql = "select sc02 " & _
                  "from staff_change " & _
                  "where sc03='02' and sc02 between '" & strStarDate & "' and '" & strEndDate & "' " & _
                  "and sc01='" & StrST01 & "' "
         '2025/1/2 END
         If m_rs.State = 1 Then m_rs.Close
         m_rs.CursorLocation = adUseClient
         m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         LongWorkDay = 0
         If m_rs.RecordCount > 0 Then
            m_Year1_Std = GetYearStdDay(CStr(Val(strYV01) - 1)) '抓前年的第一個工作天
            
            '假如 該年的第一個工作天與復職日同一天, 算做滿整年不用算比例給假
            If m_rs.Fields("sc02") <> m_Year1_Std Then
               strBackDt = m_rs.Fields("sc02")
               If Mid(m_rs.Fields("sc02"), 5, 2) = "12" Then  '復職日為12月份時
                  Call PUB_NianZiDaysYear(m_rs.Fields("sc02"), CStr(Val(strYV01) - 1) & "1231", LongWorkDay, 0) '工作天數
               Else
                  'Modify By Sindy 2025/1/2 人員新年的第一個工作天才復職,不算第12月
                  LongWorkDay = 0
                  If Not (m_rs.Fields("sc02") >= CStr(strYV01) & "0101" And m_rs.Fields("sc02") <= strEndDate) Then
                  '2025/1/2 END
                     '檢查12月份薪水是否已產生
                     strSql = "select sum(nvl(sm27,0)) from SalaryMonth where sm01='" & StrST01 & "' " & _
                              "and sm02='" & CStr(Val(strYV01) - 1) & "12' "
                     If m_rs.State = 1 Then m_rs.Close
                     m_rs.CursorLocation = adUseClient
                     m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If m_rs.RecordCount > 0 Then
                        If Val("" & m_rs.Fields(0)) > 0 Then
                           LongWorkDay = m_rs.Fields(0)
                        Else
                           LongWorkDay = 31
                        End If
                     Else
                        LongWorkDay = 31
                     End If
                  End If
                  
                  strSql = "select sum(nvl(sm27,0)) from SalaryMonth where sm01='" & StrST01 & "' " & _
                           "and sm02>='" & CStr(Val(strYV01) - 1) & "01' " & _
                           "and sm02<='" & CStr(Val(strYV01) - 1) & "11' "
                  If m_rs.State = 1 Then m_rs.Close
                  m_rs.CursorLocation = adUseClient
                  m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If m_rs.RecordCount > 0 Then
                     If m_rs.Fields(0) > 0 Then
                        LongWorkDay = LongWorkDay + m_rs.Fields(0)
                     End If
                  End If
               End If
'               m_Days = Round(m_Days * (LongWorkDay / m_YearDay), 1)
            End If
         End If
      End If
      '2017/1/3 END
      
      'B公式：週年制
      strStarDate = CStr(strYV01) & "/01/01"
      strEndDate = CStr(strYV01) & "/12/31"
      intYearDay1 = DateDiff("d", strStarDate, strEndDate) + 1 '今年總天數
      intYearDay2 = DateDiff("d", CStr(Val(strYV01) + 1) & "/01/01", CStr(Val(strYV01) + 1) & "/12/31") + 1 '明年總天數
'      '有復職日
'      If strBackDt <> "" Then
'         '取得留職停薪/復職日
'         strSql = "select sc02,sc03 " & _
'                  "from staff_change " & _
'                  "where sc03 in('04','02') and substr(sc02,1,4)='" & CStr(Val(strYV01) - 1) & "' " & _
'                  "and sc01='" & strST01 & "' " & _
'                  " order by sc02"
'         If m_rs.State = 1 Then m_rs.Close
'         m_rs.CursorLocation = adUseClient
'         m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If m_rs.RecordCount > 0 Then
'            m_rs.MoveFirst
'            intRow = 0
'            Do While Not m_rs.EOF
'               intRow = intRow + 1
'               strStarDate_sub = ""
'               strEndDate_sub = ""
'               '離職日
'               If m_rs.Fields("sc03") = "04" Then
'                  strEndDate_sub = DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(m_rs.Fields(0))))
'                  If intRow = 1 Then
'                     strStarDate_sub = CStr(Val(strYV01) - 1) & "0101"
'                  Else
'                     MsgBox "有錯誤請通知電腦中心(Err01)", vbExclamation
'                  End If
'               '復職日
'               ElseIf m_rs.Fields("sc03") = "02" Then
'                  strStarDate_sub = m_rs.Fields(0)
'                  If intRow = m_rs.RecordCount Then
'                     strEndDate_sub = CStr(Val(strYV01) - 1) & "1231"
'                  Else
'                     '下一筆離職日
'                     If m_rs.Fields("sc03") = "04" Then
'                        m_rs.MoveNext: intRow = intRow + 1
'                        strEndDate_sub = DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(m_rs.Fields(0))))
'                     Else
'                       MsgBox "有錯誤請通知電腦中心(Err02)", vbExclamation
'                     End If
'                  End If
'               End If
'               If strStarDate_sub <> "" And strEndDate_sub <> "" Then
'                  '起算日大於迄止日
'                  If Mid(strCountDate, 5) > Mid(strEndDate_sub, 5) Then
'                     intCountDay1 = DateDiff("d", ChangeWStringToWDateString(strStarDate_sub), ChangeWStringToWDateString(strEndDate_sub)) + 1
'                     m_DaysNew = m_DaysNew + Round(m_Days * (intCountDay1 / intYearDay1), 3)
'                     strDateNote = strDateNote & IIf(strDateNote <> "", vbCrLf, "") & ChangeWStringToWDateString(strStarDate_sub) & " ~ " & ChangeWStringToWDateString(strEndDate_sub)
'                     If m_Formulation = "" Then
'                        m_Formulation = "計算公式 = (" & m_Days & " * (" & intCountDay1 & " / " & intYearDay1 & "))"
'                     Else
'                        m_Formulation = m_Formulation & " + (" & m_Days & " * (" & intCountDay1 & " / " & intYearDay1 & "))"
'                     End If
'                  '起算日小於起始日
'                  ElseIf Mid(strCountDate, 5) < Mid(strStarDate_sub, 5) Then
'                     intCountDay2 = DateDiff("d", ChangeWStringToWDateString(strStarDate_sub), ChangeWStringToWDateString(strEndDate_sub)) + 1
'                     m_DaysNew = m_DaysNew + Round(m_Days2 * (intCountDay2 / intYearDay2), 3)
'                     strDateNote = strDateNote & IIf(strDateNote <> "", vbCrLf, "") & ChangeWStringToWDateString(strStarDate_sub) & " ~ " & ChangeWStringToWDateString(strEndDate_sub)
'                     If m_Formulation = "" Then
'                        m_Formulation = "計算公式 = (" & m_Days2 & " * (" & intCountDay2 & " / " & intYearDay2 & "))"
'                     Else
'                        m_Formulation = m_Formulation & " + (" & m_Days2 & " * (" & intCountDay2 & " / " & intYearDay2 & "))"
'                     End If
'                  Else
'                     '起始日~起算日-1
'                     intCountDay1 = DateDiff("d", ChangeWStringToWDateString(strStarDate_sub), CStr(strYV01) & Mid(DateAdd("d", -1, ChangeWStringToWDateString(strCountDate)), 5)) + 1
'                     m_DaysNew = m_DaysNew + Round(m_Days * (intCountDay1 / intYearDay1), 3)
'                     strDateNote = strDateNote & IIf(strDateNote <> "", vbCrLf, "") & ChangeWStringToWDateString(strStarDate_sub) & " ~ " & CStr(strYV01) & Mid(DateAdd("d", -1, ChangeWStringToWDateString(strCountDate)), 5)
'                     If m_Formulation = "" Then
'                        m_Formulation = "計算公式 = (" & m_Days & " * (" & intCountDay1 & " / " & intYearDay1 & "))"
'                     Else
'                        m_Formulation = m_Formulation & " + (" & m_Days & " * (" & intCountDay1 & " / " & intYearDay1 & "))"
'                     End If
'                     '起算日~迄止日
'                     intCountDay2 = DateDiff("d", CStr(strYV01) & Mid(ChangeWStringToWDateString(strCountDate), 5), ChangeWStringToWDateString(strEndDate_sub))
'                     m_DaysNew = m_DaysNew + Round(m_Days2 * (intCountDay2 / intYearDay2), 3)
'                     strDateNote = strDateNote & IIf(strDateNote <> "", vbCrLf, "") & CStr(strYV01) & Mid(ChangeWStringToWDateString(strCountDate), 5) & " ~ " & ChangeWStringToWDateString(strEndDate_sub)
'                     m_Formulation = m_Formulation & " + (" & m_Days2 & " * (" & intCountDay2 & " / " & intYearDay2 & "))"
'                  End If
'               Else
'                  MsgBox "有錯誤請通知電腦中心(Err03)", vbExclamation
'               End If
'               m_rs.MoveNext
'            Loop
'         End If
'      Else
         '1/1~(到職日-1日)
         'ex: strCountDate=19840301 DateAdd("d", -1, ChangeWStringToWDateString(strCountDate))=1984/2/29 =>無此日期
         'Modified by Morgan 2024/1/11 修正日期轉字串會因顯示格式而不同問題
         'strTempDate = CStr(strYV01) & Mid(DateAdd("d", -1, ChangeWStringToWDateString(strCountDate)), 5)
         strTempDate = CStr(strYV01) & Format(DateAdd("d", -1, ChangeWStringToWDateString(strCountDate)), "/MM/DD")
         strTempDate = Replace(strTempDate, "-", "/") 'Added by Morgan 2025/3/18
         'end 2024/1/11
         If IsDate(strTempDate) = False And Mid(strTempDate, 5) = "/2/29" Then
            strTempDate = Mid(strTempDate, 1, 4) & "/2/28"
         End If
         intCountDay1 = DateDiff("d", strStarDate, strTempDate) + 1
         strDateNote = strDateNote & IIf(strDateNote <> "", vbCrLf, "") & strStarDate & " ~ " & strTempDate
         '到職日~12/31
         intCountDay2 = DateDiff("d", CStr(strYV01) & Mid(ChangeWStringToWDateString(strCountDate), 5), strEndDate) + 1
         'Modify By Sindy 2019/12/30 檢查若年總天數為366則天數減1
'         If intYearDay2 = 366 Then
'            intCountDay2 = intCountDay2 - 1
'         End If
         intCountDay2 = 365 - intCountDay1
         '2019/12/30 END
'         If strST01 = "A5011" Then
'            MsgBox strST01
'         End If

         strDateNote = strDateNote & IIf(strDateNote <> "", vbCrLf, "") & CStr(strYV01) & Mid(ChangeWStringToWDateString(strCountDate), 5) & " ~ " & strEndDate
         If Val(m_Year) < 1 Then '年資未滿1年者
'            '當年滿1年者
'            If Val(Left(DateAdd("m", 12, ChangeWStringToWDateString(strCountDate)), 4)) = Val(strYV01) Then
'               '到職日滿半年落在計算當年者,特別假要累計
'               If Val(Left(DateAdd("m", 6, ChangeWStringToWDateString(strCountDate)), 4)) = Val(strYV01) Then
'                  m_DaysNew = 3
'                  m_Formulation = "計算公式 = 3"
'               End If
'               m_DaysNew = m_DaysNew + Round(m_Days2 * (intCountDay2 / intYearDay2), 3)
'               If m_Formulation <> "" Then
'                  m_Formulation = m_Formulation & " + "
'               Else
'                  m_Formulation = "計算公式 = "
'               End If
'               m_Formulation = m_Formulation & "(" & m_Days2 & " * (" & intCountDay2 & " / " & intYearDay2 & "))"
'            '當年滿半年者
'            ElseIf Val(Left(DateAdd("m", 6, ChangeWStringToWDateString(strCountDate)), 4)) = Val(strYV01) Then
'               m_DaysNew = 3
'               m_Formulation = "計算公式 = 6個月以上一年未滿者，3日"
'            End If
         '年資滿一年以上
         Else
            'Add By Sindy 2019/7/8
            '檢查是否有已滿該年資 ex.A6001(20170103)計算108年度特休
            'Modified by Morgan 2024/1/11 修正日期轉字串會因顯示格式而不同問題
            'If CStr(Val(CStr(Val(strYV01) - 1) & "1231")) < DateAdd("yyyy", m_Year, ChangeWStringToWDateString(strCountDate)) Then
            If CStr(Val(CStr(Val(strYV01) - 1) & "1231")) < DBDATE(DateAdd("yyyy", m_Year, ChangeWStringToWDateString(strCountDate))) Then
            'end 2024/1/11
               m_DaysNew = m_Days '未滿,直接給符合之特休
               strDateNote = strDateNote & IIf(strDateNote <> "", vbCrLf, "") & "未滿,直接給符合之特休" & m_Days
               m_Formulation = ""
            Else
            '2019/7/8 END
               'Modify By Sindy 2019/12/30 劉經理說固定365計算
'               m_DaysNew = Round(m_Days * (intCountDay1 / intYearDay1), 3) + Round(m_Days2 * (intCountDay2 / intYearDay2), 3)
'               m_Formulation = "計算公式 = (" & m_Days & " * (" & intCountDay1 & " / " & intYearDay1 & ")) + (" & m_Days2 & " * (" & intCountDay2 & " / " & intYearDay2 & "))"
               m_DaysNew = Round(m_Days * (intCountDay1 / 365), 3) + Round(m_Days2 * (intCountDay2 / 365), 3)
               m_Formulation = "計算公式 = (" & m_Days & " * (" & intCountDay1 & " / " & 365 & ")) + (" & m_Days2 & " * (" & intCountDay2 & " / " & 365 & "))"
            End If
         End If
'      End If
      '有復職日,再依上年度工作比例計算
      '*****
      If strBackDt <> "" Then
         m_DaysNew = Round(m_DaysNew * (LongWorkDay / m_YearDay), 2)
         m_Formulation = Replace(m_Formulation, "計算公式 = ", "計算公式 = (") & ") * (" & LongWorkDay & " / " & m_YearDay & ")"
      Else
      '***** END
         If InStr(m_DaysNew, ".") > 0 Then m_DaysNew = Round(m_DaysNew, 2)
      End If
      If InStr(m_DaysNew, ".") > 0 Then
         varArr = Split(m_DaysNew, ".")
         '滿半日不滿一日者,以一日加計
         If Val("0." & CStr(varArr(1))) > 0.5 Then
            m_DaysNew = Int(m_DaysNew) + 1
         '不滿0.5日者,以0.5日加計
         ElseIf Val("0." & CStr(varArr(1))) < 0.5 Then
            m_DaysNew = Int(m_DaysNew) + 0.5
         End If
      End If
      
      'A.曆年制 有復職狀況,依比例給假
      '有復職日
      If strBackDt <> "" Then
         m_Days = Round(m_Days * (LongWorkDay / m_YearDay), 1)
      End If
      
      strTemp = "到職日為 " & Val(Left(strST13, 4)) - 1911 & " 年 " & Mid(strST13, 5, 2) & " 月 " & Mid(strST13, 7, 2) & " 日，年資 " & m_Year & " 年" & vbCrLf
      If Val(strBackTaieDate) > 0 Then
         strTemp = strTemp & strNote
         strTemp = strTemp & "留職停薪後的起算日為 " & Val(Left(strBackTaieDate, 4)) - 1911 & " 年 " & Mid(strBackTaieDate, 5, 2) & " 月 " & Mid(strBackTaieDate, 7, 2) & " 日" & vbCrLf
      End If
      
      If m_DaysNew > m_Days Or m_Days = 30 Then
         m_Type = "B" '週年制
         PUB_GetSeniorityYearVacation = m_DaysNew
         strTemp = strTemp & _
                  "特別休假日數 " & PUB_GetSeniorityYearVacation & " 日" & vbCrLf & _
                  Trim(strDateNote) & vbCrLf & Trim(m_Formulation)
      Else
         If m_Days > 0 Then m_Type = "A" '曆年制
         PUB_GetSeniorityYearVacation = m_Days
         If Val(strBackTaieDate) > 0 And strFormulationA = "" Then '有留職停薪
            strTemp = strTemp & _
                     "特別休假日數 " & PUB_GetSeniorityYearVacation & " 日" & vbCrLf & _
                     Trim(strDateNote) & vbCrLf & Trim(m_Formulation) & "=" & m_DaysNew
         Else
            m_Formulation = strFormulationA '*****
            strTemp = strTemp & _
                     "特別休假日數 " & PUB_GetSeniorityYearVacation & " 日" & vbCrLf & _
                     Trim(strDateNote) & vbCrLf & Trim(strFormulationA)
         End If
      End If
      '***** 稽核用
      
      'Add By Sindy 2019/12/30 特別假最多只有30天
      If PUB_GetSeniorityYearVacation > 30 Then PUB_GetSeniorityYearVacation = 30
      
      'm_Formulation = strTemp & vbCrLf
   End If
End Function

'Add By Sindy 2019/6/25 計算留職停薪後特休起算日
'strText : 留職停薪：自   年  月  日至  年  月  日
Public Function Pub_BackTaieToDate(ByVal m_st01 As String, ByVal strYV01 As String, _
   Optional ByRef m_Text As String) As String
   
Dim m_rs2 As New ADODB.Recordset
Dim m_str2 As String
Dim int_i As Integer
Dim strStarDate As String
Dim strEndDate As String
Dim LngReturnDay As Long
Dim strST13 As String
   
   Pub_BackTaieToDate = "": m_Text = ""
   LngReturnDay = 0
   If Len(strYV01) < 4 Then
      strYV01 = Val(strYV01) + 1911
   End If
   '任職時間
   '02.復職 03.離職 04.留職停薪 08.退休 09.撤職 10.資遣
   'Mark:and sc02< " & Val(strYV01) - 1 & "0101
   m_str2 = "select sc02,sc03,st13 " & _
                  "from staff_change,staff " & _
                  "where sc03 in ('02','03','04','08','09','10')  " & _
                  "and sc01='" & m_st01 & "' " & _
                  "and sc01=st01(+) " & _
                  " order by sc02 asc"
   If m_rs2.State = 1 Then m_rs2.Close
   m_rs2.CursorLocation = adUseClient
   m_rs2.Open m_str2, cnnConnection, adOpenStatic, adLockReadOnly
   If m_rs2.RecordCount > 0 Then
      If m_rs2.RecordCount Mod 2 <> 0 Then Exit Function
      m_rs2.MoveFirst
      int_i = 0
      strST13 = m_rs2.Fields("st13")
      Do While Not m_rs2.EOF
         int_i = int_i + 1
         
         '離開日
         If int_i Mod 2 <> 0 Then
            strStarDate = m_rs2.Fields(0)
            strEndDate = ""
         '復職日
         Else
            If m_rs2.Fields(1) = "02" Then '復職
               'strEndDate = m_rs2.Fields(0)
               strEndDate = DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(m_rs2.Fields(0))))
            End If
         End If
         If strStarDate <= strEndDate And Val(strStarDate) > 0 And Val(strEndDate) > 0 Then
            LngReturnDay = LngReturnDay + DateDiff("d", ChangeWStringToWDateString(strStarDate), ChangeWStringToWDateString(strEndDate))
            If m_Text = "" Then
               m_Text = "留職停薪：" & vbCrLf
'            Else
'               m_Text = m_Text & "　　　　　"
            End If
            m_Text = m_Text & "　　　　自 " & Val(Left(strStarDate, 4)) - 1911 & " 年 " & Mid(strStarDate, 5, 2) & " 月 " & Mid(strStarDate, 7, 2) & " 日至 " & Val(Left(strEndDate, 4)) - 1911 & " 年 " & Mid(strEndDate, 5, 2) & " 月 " & Mid(strEndDate, 7, 2) & " 日" & vbCrLf
            strStarDate = ""
            strEndDate = ""
         End If
         
         m_rs2.MoveNext
      Loop
      If LngReturnDay > 0 And Val(strST13) > 0 Then
         '隔日起算(+1)
         Pub_BackTaieToDate = DBDATE(DateAdd("d", LngReturnDay + 1, ChangeWStringToWDateString(strST13)))
      End If
   End If
   
   Set m_rs2 = Nothing
End Function

'Add By Sindy 2019/8/16 人員刪除時要檢查的資料
Public Function StaffLeaveChkData(strSC01 As String, strSC03 As String) As String
Dim rsTmp As New ADODB.Recordset
   
   StaffLeaveChkData = ""
   '2011/10/3 modify by sonia 復職也不改
   'Modified by Morgan 2014/12/22 留職停薪補充保費要改為兼職計算收文部門要改為F51
   'If Left(Trim(strSC03), 2) <> "04" And Left(Trim(strSC03), 2) <> "02" Then '2011/7/1 add by sonia 辜說留職停薪不改入帳類別及外翻編號的部門
   If Left(Trim(strSC03), 2) <> "02" Then
   'end 2014/12/22
      '先檢查是否有外翻編號
      strSql = "SELECT sim02 FROM staff_idmap " & _
               "WHERE sim01='" & strSC01 & "'  "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         'Added by Morgan 2014/12/22
         If Left(Trim(strSC03), 2) = "04" Then
            StaffLeaveChkData = "請電腦中心於隔月發薪後調整該員工外翻編號 " & rsTmp.Fields(0) & " 的３個部門代號為 'F51'！"
         Else
         'end 2014/12/22
            StaffLeaveChkData = "請財務處同仁同時調整該員工外翻編號 " & rsTmp.Fields(0) & " 的入帳類別！"
            '2010/5/4 ADD BY SONIA
            StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "請電腦中心於隔月發薪後調整該員工外翻編號 " & rsTmp.Fields(0) & " 的３個部門代為 'F51'！"
         End If 'Added by Morgan 2014/12/22
      End If
   End If   '2011/7/1 end
   
   '2010/12/14 ADD BY SONIA 留職停薪SD02='S'後直接離職者SD02編制要改來,否則人事處每年統計資料會錯誤
   If Left(Trim(strSC03), 2) = "03" Or Left(Trim(strSC03), 2) = "08" Or _
      Left(Trim(strSC03), 2) = "09" Or Left(Trim(strSC03), 2) = "10" Then
      strSql = "SELECT SD02 FROM SALARYDATA WHERE SD01='" & strSC01 & "' AND SD02='S' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁為留職停薪後直接離職，請電腦中心將編制SD02改回'R'正式員工！"
      End If
      
      'Add By Sindy 2012/2/4 國外部員工離職時,加入下列提醒
      strSql = "SELECT st01 FROM staff WHERE st01='" & strSC01 & "' AND substr(st03,1,1)='F' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該員工為國外部同仁，請電腦中心依規定修改客戶檔及下一程序檔智權人員！"
      End If
      '2012/2/4 End
      
      'Add By Sindy 2016/8/25 CFP程序人員離職時,加入下列提醒
      strSql = "SELECT st01 FROM staff WHERE st01='" & strSC01 & "' AND (st05='83' or st05='85')"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該員工為CFP程序人員，請電腦中心檢查下一程序相關期限資料！"
      End If
      '2016/8/25 End
   End If
   '2010/12/14 END
   
   '2011/5/10 ADD BY SONIA 異動原因為03離職、04留職停薪、08退休、09撤職、10資遣時若為特殊人員或帶人主管
   If Left(Trim(strSC03), 2) = "03" Or _
      Left(Trim(strSC03), 2) = "04" Or _
      Left(Trim(strSC03), 2) = "08" Or _
      Left(Trim(strSC03), 2) = "09" Or _
      Left(Trim(strSC03), 2) = "10" Then
      'modify by sonia 2019/5/20 剔除'非法務部律師名單',是否含離職人員由程式控制
      'modify by sonia 2021/12/1 再剔除'MCTMember'
      strSql = "select * from SetSpecMan where instr(oMan,'" & strSC01 & "')>0 and oCode not in ('非法務部律師名單','MCTMember') "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁為特殊人員, 請電腦中心詢問主管如何調整設定！"
      End If
      strSql = "select * from staff where st04='1' and (st52='" & strSC01 & "' or st53='" & strSC01 & "' or st54='" & strSC01 & "' or st55='" & strSC01 & "') "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁為帶人主管, 請電腦中心詢問主管如何調整設定！"
      End If
      
      'Add By Sindy 2019/11/4
      strSql = "select * from staff where st04='1' and instr(st14,'" & strSC01 & "')>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁為內部郵件收件人員, 請電腦中心詢問主管如何調整設定！"
      End If
      'Add By Sindy 2025/4/18
      strSql = "select * from staff where st04='1' and instr(st59,'" & strSC01 & "')>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁為內部郵件收件人員２(適用於打卡異常通知), 請電腦中心詢問主管如何調整設定！"
      End If
      '2025/4/18 END
      '2019/11/4 END
      
      '2011/5/18 ADD BY SONIA 再加檢查是否為部門主管
      strSql = "select * from ACC090 where A0908='" & strSC01 & "' or A0909='" & strSC01 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁為部門主管, 請電腦中心詢問上級如何調整設定！"
      End If
      '2011/5/18 END
      'Added by Lydia 2024/01/04
      If strSrvDate(1) >= 新部門啟用日 And InStr("," & StaffLeaveChkData, "部門主管") = 0 Then
         strSql = "select * from ACC090NEW where A0924='" & strSC01 & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁為部門主管, 請電腦中心詢問上級如何調整設定！"
         End If
      End If
      
      'Add By Sindy 2016/1/8
      'Modify By Sindy 2016/3/1 +ST62,ST63,ST67
      'Modify By Sindy 2022/4/1 + and st04='1'
      'modify by sonia 2022/4/15 +A0916~A0918  2023/3/8取消A0917
      'Modify By Sindy 2023/12/20 +ACC090NEW
      strSql = "select 'A0912 試用期滿通知人員' from ACC090 where instr(A0912,'" & strSC01 & "')>0 " & _
               " Union select 'A0913 薪資查詢權限主管' from ACC090 where instr(A0913,'" & strSC01 & "')>0 " & _
               " Union select 'A0914 業績工作職務代理人' from ACC090 where instr(A0914,'" & strSC01 & "')>0 " & _
               " Union select 'A0915 假日加班簽核主管' from ACC090 where instr(A0915,'" & strSC01 & "')>0 " & _
               " Union select 'A0916 CFP程序管制人' from ACC090 where instr(A0916,'" & strSC01 & "')>0 " & _
               " Union select 'A0918 每月點數輸入確認主管' from ACC090 where instr(A0918,'" & strSC01 & "')>0 " & _
               " Union select 'ST62 草圖核稿人' from staff where instr(ST62,'" & strSC01 & "')>0 and st04='1'" & _
               " Union select 'ST63 繪圖判發人' from staff where instr(ST63,'" & strSC01 & "')>0 and st04='1'" & _
               " Union select 'ST67 離職後之後續管制人' from staff where instr(ST67,'" & strSC01 & "')>0 and st04='1'" & _
               " Union select 'A0924 部門主管(NEW)' from ACC090NEW where instr(A0924,'" & strSC01 & "')>0 " & _
               " Union select 'A0926 試用期滿通知人員(NEW)' from ACC090NEW where instr(A0926,'" & strSC01 & "')>0 " & _
               " Union select 'A0927 薪資查詢權限主管(NEW)' from ACC090NEW where instr(A0927,'" & strSC01 & "')>0 " & _
               " Union select 'A0928 假日加班簽核主管(NEW)' from ACC090NEW where instr(A0928,'" & strSC01 & "')>0 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         strSql = ""
         Do While Not RsTemp.EOF
            strSql = strSql & ";" & RsTemp.Fields(0)
            RsTemp.MoveNext
         Loop
         StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁有設定" & Mid(strSql, 2) & ", 請電腦中心詢問上級如何調整設定！"
      End If
      '2016/1/8 END
      
      '2013/5/28 ADD BY SONIA 再加檢查是否為核稿人
      '2013/9/24 modify by sonia 加判發人pp05
      strSql = "select PromoterProofreader.*,st02 from PromoterProofreader,staff where (pp04='" & strSC01 & "' or pp05='" & strSC01 & "') and pp02=st01(+) and st04='1' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁為核稿人或判發人, 請電腦中心詢問該部門主管如何調整設定！"
      End If
      '2013/5/28 END
      
      'Add By Sindy 2016/4/28 案件表單簽核人員設定
      'Modify By Sindy 2016/5/16 (f0101='" & strSC01 & "' or)取消
      strSql = "select f0101 from flow001,staff where (f0103='" & strSC01 & "'" & _
               " or f0104='" & strSC01 & "' or f0105='" & strSC01 & "' or f0106='" & strSC01 & "'" & _
               " or f0107='" & strSC01 & "' or f0108='" & strSC01 & "')" & _
               " and f0101=st01(+) and st04='1'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁有設定〔案件表單簽核人員資料〕, 請電腦中心詢問上級如何調整設定！"
         End If
      End If
      '2016/4/28 END
      
      'Add By Sindy 2016/5/27
      strSql = "select * from ipdeptkeyword where instr(LK04,'" & strSC01 & "')>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁為〔郵件分信關鍵字對照表〕的收受者, 請電腦中心詢問上級如何調整設定！"
         End If
      End If
      '2016/5/27 END
      
      'Add by Amy 2020/12/22 研討會聯絡人檔
      strSql = "Select 'SC04 教育訓練聯絡人檔知會主管' From SeminarContact Where InStr(SC04,'" & strSC01 & "')>0 " & _
        "Union Select 'SC05 教育訓練聯絡人檔副本收受者' From SeminarContact Where InStr(SC05,'" & strSC01 & "')>0 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            RsTemp.MoveFirst
            strSql = ""
            Do While Not RsTemp.EOF
               strSql = strSql & ";" & RsTemp.Fields(0)
               RsTemp.MoveNext
            Loop
            StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁有設定" & Mid(strSql, 2) & ", 請電腦中心詢問上級如何調整設定！"
         End If
      End If
      'end 2020/12/22
      
      'Add by Amy 2022/09/28 承辦人責任區分配檔DutyZoneAssign
      'Modify By Sindy 2023/2/13 此處智權人員離職不需檢查 mark:" Union Select 'DZA02 承辦人責任區分配人員' From DutyZoneAssign Where InStr(DZA02,'" & strSC01 & "')>0 And length(DZA02)=5"
      strSql = "Select 'DZA01 承辦人責任區分配人員' From DutyZoneAssign Where InStr(DZA01,'" & strSC01 & "')>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strSql = ""
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
             strSql = strSql & ";" & RsTemp.Fields(0)
             RsTemp.MoveNext
         Loop
         StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁有設定" & Mid(strSql, 2) & ", 請電腦中心詢問上級如何調整設定！"
      End If
      'end 2022/09/28
      
      'Added by Morgan 2022/8/25 程序人員核判表
      strSql = "select * from LETTERREVIEWER where LR07='" & strSC01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁為專利處〔程序人員核判表〕的判發人, 請電腦中心詢問程序主管如何調整設定！"
         End If
      End If
      'end 2022/8/25
      
      'Add by Sindy 2024/10/30 請假特殊規則簽核設定檔
      strSql = "select * from ABS003 where B0304='" & strSC01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁有設定〔請假特殊規則簽核設定檔〕的最後簽核人員, 請電腦中心詢問上級如何調整設定！"
         End If
      End If
      '2024/10/30 END
      
      'Added by Morgan 2023/5/3
      '客戶承辦工程師對照檔CustEngMap
      'modify by sonia 2024/6/26 +cem06
      strSql = "select * from custengmap where cem03='" & strSC01 & "' or cem06='" & strSC01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            'modify by sonia 2024/6/26 +密件副本收受人
            StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁為專利處〔客戶承辦工程師對照檔〕的承辦工程師或密件副本收受人, 請電腦中心詢問程序主管如何調整設定！"
         End If
      End If
      'end 2023/5/3
      
      'Add By Sindy 2023/5/15 增加檢查信件未沖銷但已離職的人員
      strSql = "select ir04,st14,count(*) cnt from(" & _
               "select ir04,ir01,ir03,ir22 from tminput,inputrecord,staff" & _
               " Where nvl(ti16, 0) = 0 And ti01 = iR01 And ti03 = iR03 And iR04 = st01" & _
               " and ir08=0" & _
               " Union All" & _
               " select ir04,ir01,ir03,ir22 from ipdeptinput,inputrecord,staff" & _
               " Where nvl(ii16, 0) = 0 And ii01 = iR01 And ii03 = iR03 And iR04 = st01" & _
               " and ir08=0" & _
               " Union All" & _
               " select ir04,ir01,ir03,ir22 from patentinput,inputrecord,staff" & _
               " Where nvl(pi16, 0) = 0 And pi01 = iR01 And pi03 = iR03 And iR04 = st01" & _
               " and ir08=0" & _
               ") a,staff where a.ir04=st01 and a.ir04='" & strSC01 & "' group by ir04,st14"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            'Modify By Sindy 2023/6/7
            If Val("" & RsTemp.Fields("cnt")) > 0 Then
            '2023/6/7 END
               'Modify By Sindy 2023/11/1
               'StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & "該同仁尚有信件未沖銷, 請電腦中心詢問主管如何調整[內部郵件收件員工編號]的設定(後續可處理的信件人員)！(若為MCTF人員除外，可不用設定)"
               StaffLeaveChkData = StaffLeaveChkData & vbCrLf & vbCrLf & _
                                   "該同仁尚有信件未沖銷，請詢問主管後續需設定給那一位同仁處理該信件（修改離職人員的內部郵件收件員工編號）" & vbCrLf & _
                                   "若為 MCTMember 人員不用修改。"
            End If
         End If
      End If
      '2023/5/15 END
   End If
   
   Set rsTmp = Nothing
End Function

'Added by Morgan 2020/3/24
'檢查是否為法律所人員
Public Function PUB_ChkLCompStaff(pUserNo As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   
   stSQL = "select st15 from staff where st01='" & pUserNo & "' and st03 like 'L%'"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      PUB_ChkLCompStaff = True
   End If
   Set RsQ = Nothing
End Function

'Add By Sindy 2021/2/18
Public Sub PUB_CallScMailTOM13(pUserNo As String, pSCDate As String)
Dim stSQL As String, intQ As Integer
Dim RsQ As ADODB.Recordset
Dim strTo As String 'Add By Sindy 2023/8/11
Dim strContext As String
   
   'Modify By Sindy 2023/12/20
   If strSrvDate(1) >= 新部門啟用日 Then
      stSQL = "select staff_change.*,a0922,a2.ac03 SC05_2,a3.ac03 SC06_2,a4.ac03 SC03_2" & _
              " from staff_change,acc090NEW a1,allcode a2,allcode a3,allcode a4" & _
              " where sc01='" & pUserNo & "' and sc02=" & DBDATE(pSCDate) & _
              " and SC04=a1.a0921(+) and '01'=a2.ac01(+) and SC05=a2.ac02(+)" & _
              " and '02'=a3.ac01(+) and SC06=a3.ac02(+) and '05'=a4.ac01(+) and SC03=a4.ac02(+)"
   Else
   '2023/12/20 END
      stSQL = "select staff_change.*,a0902,a2.ac03 SC05_2,a3.ac03 SC06_2,a4.ac03 SC03_2" & _
              " from staff_change,acc090 a1,allcode a2,allcode a3,allcode a4" & _
              " where sc01='" & pUserNo & "' and sc02=" & DBDATE(pSCDate) & _
              " and SC04=a1.a0901(+) and '01'=a2.ac01(+) and SC05=a2.ac02(+)" & _
              " and '02'=a3.ac01(+) and SC06=a3.ac02(+) and '05'=a4.ac01(+) and SC03=a4.ac02(+)"
   End If
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      '同仁異動原因為01新進,02復職,03離職,07調職時,請系統發通知給77047謝經理及73035陳麗��
      'Modify By Sindy 2021/3/3 + 08退休
      'Modify By Sindy 2023/8/11 + 04留職停薪,09撤職,10資遣
      If RsQ.Fields("sc03") = "01" Or _
         RsQ.Fields("sc03") = "02" Or _
         RsQ.Fields("sc03") = "03" Or _
         RsQ.Fields("sc03") = "07" Or _
         RsQ.Fields("sc03") = "08" Or _
         RsQ.Fields("sc03") = "04" Or _
         RsQ.Fields("sc03") = "09" Or _
         RsQ.Fields("sc03") = "10" Then
         
         'Modify By Sindy 2023/8/11
         strTo = Pub_GetSpecMan("北所分機設定通知")
         If PUB_GetST06(pUserNo) = "2" Then '中所
            strTo = strTo & ";" & Pub_GetSpecMan("中所分機設定通知")
         ElseIf PUB_GetST06(pUserNo) = "3" Then '南所
            strTo = strTo & ";" & Pub_GetSpecMan("南所分機設定通知")
         ElseIf PUB_GetST06(pUserNo) = "4" Then '高所
            strTo = strTo & ";" & Pub_GetSpecMan("高所分機設定通知")
         End If
         '2023/8/11 END
         'Modify By Sindy 2023/12/20 部門調整改抓ST93
         strContext = "員工編號：" + RsQ.Fields("sc01") + " " + GetPrjSalesNM(RsQ.Fields("sc01")) & vbCrLf & _
                      "異動原因：" + RsQ.Fields("SC03_2") & vbCrLf & _
                      "異動日期：" + ChangeWStringToTDateString(RsQ.Fields("sc02")) & vbCrLf & vbCrLf
         If strSrvDate(1) >= 新部門啟用日 Then
            strContext = strContext & "部　門：" + RsQ.Fields("a0922") & vbCrLf
         Else
            strContext = strContext & "部　門：" + RsQ.Fields("a0902") & vbCrLf
         End If
         strContext = strContext & _
         "職　位：" + "" & RsQ.Fields("SC06_2") & vbCrLf & _
         "職　稱：" + "" & RsQ.Fields("SC05_2") & vbCrLf & _
         "職稱說明：" + "" & RsQ.Fields("sc07") & vbCrLf & _
         "所　別：" + IIf("" & RsQ.Fields("sc14") = "1", "北所", IIf("" & RsQ.Fields("sc14") = "2", "中所", IIf("" & RsQ.Fields("sc14") = "3", "南所", IIf("" & RsQ.Fields("sc14") = "4", "高所", "其他")))) & vbCrLf & vbCrLf
         PUB_SendMail strUserNum, strTo, "", "人事異動【" & IIf(RsQ.Fields("sc03") = "01", "新進", IIf(RsQ.Fields("sc03") = "02", "復職", "離職")) & "】通知，請調整電話分機號碼！", strContext, , , , , , , , , , True
      End If
   End If
   Set RsQ = Nothing
End Sub

'Add by Amy 2022/08/04 從basPublic搬過來
'Added by Lydia 2019/12/26 產生共同查詢對造案件暫存檔R100102_1
'Modify by Amy 2021/08/13 +bolDelRelevantP '刪除其他相關人
'Modify by Amy 2021/08/16 +bolShowCPField 顯示CP查到欄位
Public Sub Pub_ProcR100102_1(ByVal pIDname As String, ByVal pSql01 As String, ByVal pSql02 As String, ByVal pSql03 As String, ByVal pSql04 As String, ByVal pSql05 As String, _
                                               ByVal pTp03 As String, ByVal pCheckWay As String, Optional ByVal bolDelRelevantP As Boolean = False, Optional ByVal bolShowCPField As Boolean = False)
'pIDname : strUserNum+@+表單名稱
Dim strMainSql As String
Dim strTemp1 As String
Dim intR As Integer
Dim mChk01 As String '對造名稱: 中文-CP40,英文-CP41,日文-CP42
Dim mChk02 As String '被授權人/被再授權人: 中文-CP50,英文-CP51,日文-CP52
Dim mSubSQL1 As String, mSubSQL2 As String
Dim mSwhSQL1 As String, mSwhSQL2 As String
Dim rsRD As New ADODB.Recordset
Dim b_Sys As String, intCnt As Integer
'Add by Amy 2020/08/27
Dim bShowF As Boolean '顯示欄位
Dim stFN01 As String, stFN02 As String '顯示欄名
Dim strApplyF1 As String, strApplyF2 As String 'Add by Amy 2022/04/11
Dim mChk01_O As String, mChk02_O As String 'Add by Amy 2023/06/27
Dim stAddSign As String 'Add by Amy 2024/01/23 加符號查

    cnnConnection.Execute "delete from R100102_1 where id='" & pIDname & "' "
    
    'Modify by Amy 2022/04/11 +申請人1~5 英文(r021026~30) 申請人1~5日文(r021031~35),查PROTEK MANUFACTURING CORP. TF-000660-1-02此對造名稱與申請人的英文名稱相同應刪除
    strMainSql = "Insert Into R100102_1 (r021001,r021002,r021003,r021004,r021005,r021006,r021007,r021008,r021009,r021010,r021011,r021012,r021013,r021014,r021015,r021016,r021017,r021018,ID,R021020,R021021,R021022,R021023,R021024,R021025" & _
                        ",r021026,r021027,r021028,r021029,r021030,r021031,r021032,r021033,r021034,r021035) "
            
    'Modify by Amy 2020/08/27 +顯示欄位(frm140401用)
    'Modify by Amy 2021/08/16+bolShowCPField顯示CP查到欄位
    If UCase(Mid(pIDname, (InStr(pIDname, "@")) + 1)) = "FRM140401" Or bolShowCPField = True Then
        bShowF = True
    End If
    'Modify by Amy 2024/01/23 進度資料會以☆區隔,比對相當=
    If UCase(Mid(pIDname, (InStr(pIDname, "@")) + 1)) = "FRM12040163" Then
        stAddSign = "☆"
    End If
    
    For intR = 1 To 3
        stFN01 = "": stFN02 = ""
        'Modify by Amy 2023/06/26 加mChk01_O,改新欄位
        Select Case intR
             Case 1: '中文
                  mChk01_O = "CP40"
                  mChk02_O = "CP50"
                  mChk01 = "CP169"
                  mChk02 = "CP172"
                  If bShowF = True Then
                        stFN01 = "@@CP40"
                        stFN02 = "@@CP50"
                  End If
             Case 2: '英文
                  mChk01_O = "CP41"
                  mChk02_O = "CP51"
                  mChk01 = "CP170"
                  mChk02 = "CP173"
                  If bShowF = True Then
                        stFN01 = "@@CP41"
                        stFN02 = "@@CP51"
                  End If
             Case 3: '日文
                  mChk01_O = "CP42"
                  mChk02_O = "CP52"
                  mChk01 = "CP171"
                  mChk02 = "CP174"
                  If bShowF = True Then
                        stFN01 = "@@CP42"
                        stFN02 = "@@CP52"
                  End If
        End Select
        'end 2023/06/26
        'end 2020/08/04

        'Add by Amy 2023/01/04 傳入的字串也要一致Repalce ex:「客戶檔」只跑此函數要一致
        'Modify by Amy 2023/01/07 取代改共用函數
        'Modify by Amy 2023/06/26 改抓ReplaceSign DB函數
        'pTp03 = Pub_ReplaceSign(False, pTp03)
        'Modify by Amy 2024/01/23 風險檢查比對進度資料會以☆區隔(比對相當=)
        If InStr(pTp03, "☆") = 0 And stAddSign <> MsgText(601) Then
            pTp03 = Pub_GetField("Dual", "1=1", "ReplaceSign(TO_MULTI_BYTE(Upper('" & stAddSign & ChgSQL(pTp03) & stAddSign & "')))")
        Else
            pTp03 = Pub_GetField("Dual", "1=1", "ReplaceSign(TO_MULTI_BYTE(Upper('" & ChgSQL(pTp03) & "')))")
        End If
        
        'Add by Amy 2023/06/08 傳入要查之字串先將數字、英文變全型 ex:DB是ＬＧ(全型)用LG(半型)會查不到
        'Mark by Amy 2023/06/13 ex:Y55074 J-star 會查不到,使用PUB_ChangeZIPToSir轉-和TO_MULTI_BYTE轉的不一致
        'pTp03 = PUB_ChangeZIPToSir(pTp03)
        
        'Modify by Amy 2021/08/13 pCheckWay為=
        'Modify by Amy 2022/08/04 去除 , 、，、 . 、．、。、半型空白、全型空白 再查 ex:李,政義
        'Modify by Amy 2023/06/08 +TO_MULTI_BYTE 數字、英文變全型再抓, ex:DB是ＬＧ(全型)用LG(半型)會查不到
        'Modify by Amy 2023/06/13 ex:Y55074 J-star 會查不到,使用PUB_ChangeZIPToSir轉-和TO_MULTI_BYTE轉的不一致
        'Modify by Amy 2023/06/26 改抓新欄位且用ReplaceSign DB函數取代符號
        If pCheckWay = "=" Then
            'mSubSQL1 = " And Replace(Replace(Replace(Replace(Replace(Replace(Replace(TO_MULTI_BYTE(Upper(" & mChk01 & ")),',',''),'，',''),'.',''),'．',''),'。',''),' ',''),'　','')=TO_MULTI_BYTE('" & _
                                           ChgSQL(Replace(Replace(Replace(Replace(Replace(Replace(Replace(UCase(pTp03), ",", ""), "，", ""), ".", ""), "．", ""), "。", ""), " ", ""), "　", "")) & "') "
            If stAddSign <> MsgText(601) Then
               mSubSQL1 = " And '☆'||" & mChk01 & "||'☆' ='" & pTp03 & "' "
               mSubSQL2 = " And '☆'||" & mChk02 & "||'☆' ='" & pTp03 & "' "
            Else
               mSubSQL1 = " And " & mChk01 & "='" & pTp03 & "' "
               mSubSQL2 = " And " & mChk02 & "='" & pTp03 & "' "
            End If
        Else
            If stAddSign <> MsgText(601) Then
               mSubSQL1 = " And InStr('☆'||" & mChk01 & "||'☆','" & pTp03 & "') " & pCheckWay
               mSubSQL2 = " And InStr('☆'||" & mChk02 & "||'☆','" & pTp03 & "') " & pCheckWay
            Else
               'mSubSQL1 = " And InStr(Replace(Replace(Replace(Replace(Replace(Replace(Replace(TO_MULTI_BYTE(Upper(" & mChk01 & ")),',',''),'，',''),'.',''),'．',''),'。',''),' ',''),'　',''),TO_MULTI_BYTE('" & ChgSQL(UCase(pTp03)) & "')) " & pCheckWay
               mSubSQL1 = " And InStr(" & mChk01 & ",'" & pTp03 & "') " & pCheckWay
               mSubSQL2 = " And InStr(" & mChk02 & ",'" & pTp03 & "') " & pCheckWay
            End If
        End If
        'end 2024/01/23
        'end 2023/06/26
        'end 2023/06/13
        'end 2023/06/08
        'end 2022/08/04
        'end 2021/08/13
'** End Memo 此處有修改要確認 GetSearchNameSql 函數 是否也要改 **

        mSwhSQL1 = mChk01 & " >' ' "
        mSwhSQL2 = mChk02 & " >' ' "

        'Modify by Amy 2020/08/27 名稱欄資料內加原始欄位
        'Modify by Amy 2022/04/11 原:"NVL(C1.CU04,NVL(C1.CU05||C1.CU88||C1.CU89||C1.CU90,C1.CU06)) AS 申請人1,NVL(C2.CU04,NVL(C2.CU05||C2.CU88||C2.CU89||C2.CU90,C2.CU06)) AS 申請人2,NVL(C3.CU04,NVL(C3.CU05||C3.CU88||C3.CU89||C3.CU90,C3.CU06)) AS 申請人3," & _
                                                           "NVL(C4.CU04,NVL(C4.CU05||C4.CU88||C4.CU89||C4.CU90,C4.CU06)) AS 申請人4,NVL(C5.CU04,NVL(C5.CU05||C5.CU88||C5.CU89||C5.CU90,C5.CU06)) AS 申請人5
        strApplyF1 = "C1.CU04,C2.CU04,C3.CU04,C4.CU04,C5.CU04" '申請人1~5 中文
        strApplyF2 = "C1.CU05||C1.CU88||C1.CU89||C1.CU90,C2.CU05||C1.CU88||C1.CU89||C1.CU90,C3.CU05||C1.CU88||C1.CU89||C1.CU90,C4.CU05||C1.CU88||C1.CU89||C1.CU90,C5.CU05||C1.CU88||C1.CU89||C1.CU90" '申請人1~5 英文
        strApplyF2 = strApplyF2 & ",C1.CU06,C2.CU06,C3.CU06,C4.CU06,C5.CU06" '申請人1~5 日文
        '1-商標
        strMainSql = strMainSql & IIf(intR > 1, " Union ", "") & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, " & mChk01_O & "||'" & stFN01 & "' as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,''||cp05 as 收文日," & _
                        strApplyF1 & ",CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & pIDname & "',tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno," & strApplyF2 & " " & _
                        "From (Select * From CaseProgress Where " & mSwhSQL1 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+) " & _
                        "and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+) " & _
                        "and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & pSql01 & mSubSQL1
        'Modified by Lydia 2020/06/10 '1' as 狀態 =>'2' as 狀態
        strMainSql = strMainSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號, " & mChk02_O & "||'" & stFN02 & "' as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(TM10,'000',CPM03,CPM04),CP10) AS 案件性質,''||cp05 as 收文日," & _
                        strApplyF1 & ",CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & pIDname & "',tm23 as cust01,tm78 as cust02,tm79 as cust03,tm80 as cust04,tm81 as cust05,tm44 as fcno," & strApplyF2 & " " & _
                        "From (Select * From CaseProgress Where " & mSwhSQL2 & "),TradeMark,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(TM23,1,8) = c1.CU01(+) and Decode(Substr(TM23,9,1),null,'0',Substr(TM23,9,1)) = c1.CU02(+) " & _
                        "and Substr(tm78,1,8)=c2.cu01(+) and Decode(Substr(tm78,9,1),null,'0',Substr(tm78,9,1))=c2.cu02(+) and Substr(tm79,1,8)=c3.cu01(+) and Decode(Substr(tm79,9,1),null,'0',Substr(tm79,9,1))=c3.cu02(+) and Substr(tm80,1,8)=c4.cu01(+) and Decode(Substr(tm80,9,1),null,'0',Substr(tm80,9,1))=c4.cu02(+) " & _
                        "and Substr(tm81,1,8)=c5.cu01(+) and Decode(Substr(tm81,9,1),null,'0',Substr(tm81,9,1))=c5.cu02(+) " & pSql01 & mSubSQL2
        '2-專利
        strMainSql = strMainSql & " Union " & _
                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號," & mChk01_O & "||'" & stFN01 & "' as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,''||cp05 as 收文日, " & _
                         strApplyF1 & ",CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & pIDname & "',pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno," & strApplyF2 & " " & _
                         "From (Select * From CaseProgress Where " & mSwhSQL1 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
                         "and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
                         "and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & pSql02 & mSubSQL1
        'Modified by Lydia 2020/06/10 '1' as 狀態 =>'2' as 狀態
        strMainSql = strMainSql & " Union " & _
                         "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號," & mChk02_O & "||'" & stFN02 & "' as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(PA09,'000',CPM03,CPM04),CP10) AS 案件性質,''||cp05 as 收文日, " & _
                         strApplyF1 & ",CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & pIDname & "',pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno," & strApplyF2 & " " & _
                         "From (Select * From CaseProgress Where " & mSwhSQL2 & "),Patent,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                         "Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(pa26,1,8)=c1.cu01(+) and Decode(Substr(pa26,9,1),null,'0',Substr(pa26,9,1))=c1.cu02(+) " & _
                         "and Substr(pa27,1,8)=c2.cu01(+) and Decode(Substr(pa27,9,1),null,'0',Substr(pa27,9,1))=c2.cu02(+) and Substr(pa28,1,8)=c3.cu01(+) and Decode(Substr(pa28,9,1),null,'0',Substr(pa28,9,1))=c3.cu02(+) and Substr(pa29,1,8)=c4.cu01(+) and Decode(Substr(pa29,9,1),null,'0',Substr(pa29,9,1))=c4.cu02(+) " & _
                         "and Substr(pa30,1,8)=c5.cu01(+) and Decode(Substr(pa30,9,1),null,'0',Substr(pa30,9,1))=c5.cu02(+) " & pSql02 & mSubSQL2
        '3-法務
        strMainSql = strMainSql & " Union " & _
                       "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號," & mChk01_O & "||'" & stFN01 & "' as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,''||cp05 as 收文日, " & _
                       strApplyF1 & ",CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & pIDname & "',lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno," & strApplyF2 & " " & _
                       "From (Select * From CaseProgress Where " & mSwhSQL1 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                       "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
                       "and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
                       "and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & pSql03 & mSubSQL1
        'Modified by Lydia 2020/06/10 '1' as 狀態 =>'2' as 狀態
        strMainSql = strMainSql & " Union " & _
                       "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號," & mChk02_O & "||'" & stFN02 & "' as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(LC15,'000',CPM03,CPM04),CP10) AS 案件性質,''||cp05 as 收文日, " & _
                       strApplyF1 & ",CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & pIDname & "',lc11 as cust01,lc43 as cust02,lc44 as cust03,lc45 as cust04,lc46 as cust05,lc22 as fcno," & strApplyF2 & " " & _
                       "From (Select * From CaseProgress Where " & mSwhSQL2 & "),LawCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                       "Where CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(LC11,1,8)=c1.CU01(+) AND Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1)) = c1.cu02(+) " & _
                       "and Substr(lc43,1,8)=c2.cu01(+) AND Decode(Substr(lc43,9,1),null,'0',Substr(lc43,9,1))=c2.cu02(+) and Substr(lc44,1,8)=c3.cu01(+) and Decode(Substr(lc44,9,1),null,'0',Substr(lc44,9,1))=c3.cu02(+) and Substr(lc45,1,8)=c4.cu01(+) and Decode(Substr(lc45,9,1),null,'0',Substr(lc45,9,1))=c4.cu02(+) " & _
                       "and Substr(lc46,1,8)=c5.cu01(+) and Decode(Substr(lc46,9,1),null,'0',Substr(lc46,9,1))=c5.cu02(+) " & pSql03 & mSubSQL2
        '4.顧問
        strMainSql = strMainSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號," & mChk01_O & "||'" & stFN01 & "'  as 名稱,' ' as 智權人,'1' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,''||cp05 as 收文日," & _
                        strApplyF1 & ",CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & pIDname & "',hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno," & strApplyF2 & " " & _
                        "From (Select * From CaseProgress Where " & mSwhSQL1 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
                        "and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
                        "and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & pSql04 & mSubSQL1
        'Modified by Lydia 2020/06/10 '1' as 狀態 =>'2' as 狀態
        strMainSql = strMainSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號," & mChk02_O & "||'" & stFN02 & "'  as 名稱,' ' as 智權人,'2' as 狀態,' ' as 申請國家,CP09 as 總收文號,NVL(Decode(CPM03,null,CPM04,CPM03),CP10) AS 案件性質,''||cp05 as 收文日," & _
                        strApplyF1 & ",CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & pIDname & "',hc05 as cust01,hc24 as cust02,hc25 as cust03,hc26 as cust04,hc27 as cust05,'' as fcno," & strApplyF2 & " " & _
                        "From (Select * From CaseProgress Where " & mSwhSQL2 & "),HireCase,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND Substr(HC05,1,8)=c1.cu01(+) AND Decode(Substr(HC05,9,1),null,'0',Substr(HC05,9,1))=c1.cu02(+) " & _
                        "and Substr(hc24,1,8)=c2.cu01(+) AND Decode(Substr(hc24,9,1),null,'0',Substr(hc24,9,1))=c2.cu02(+) and Substr(hc25,1,8)=c3.cu01(+) and Decode(Substr(hc25,9,1),null,'0',Substr(hc25,9,1))=c3.cu02(+) and Substr(hc26,1,8)=c4.cu01(+) and Decode(Substr(hc26,9,1),null,'0',Substr(hc26,9,1))=c4.cu02(+) " & _
                        "and Substr(hc27,1,8)=c5.cu01(+) AND Decode(Substr(hc27,9,1),null,'0',Substr(hc27,9,1))=c5.cu02(+) " & pSql04 & mSubSQL2
        '5-服務
        strMainSql = strMainSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號," & mChk01_O & "||'" & stFN01 & "' as 名稱,' ' as 智權人,'1' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,''||cp05 as 收文日," & _
                        strApplyF1 & ",CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & pIDname & "',sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno," & strApplyF2 & " " & _
                        "From (Select * From CaseProgress Where " & mSwhSQL1 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
                        "and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
                        "and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & pSql05 & mSubSQL1
        'Modified by Lydia 2020/06/10 '1' as 狀態 =>'2' as 狀態
        strMainSql = strMainSql & " Union " & _
                        "Select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 編號," & mChk02_O & "||'" & stFN02 & "' as 名稱,' ' as 智權人,'2' as 狀態,SP09 as申請國家,CP09 as 總收文號,NVL(Decode(SP09,'000',CPM03,CPM04),CP10) AS 案件性質,''||cp05 as 收文日," & _
                        strApplyF1 & ",CP01,CP02,CP03,CP04,CP10 AS 案件性質編號,'" & pIDname & "',sp08 as cust01,sp58 as cust02,sp59 as cust03,sp65 as cust04,sp66 as cust05,sp26 as fcno," & strApplyF2 & " " & _
                        "From (Select * From CaseProgress Where " & mSwhSQL2 & "),ServicePractice,CasePropertyMap,customer c1,customer c2,customer c3,customer c4,customer c5 " & _
                        "Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND SUBSTR(SP08,1,8)=C1.CU01(+) AND Decode(Substr(sp08,9,1),null,'0',Substr(sp08,9,1))=c1.cu02(+) " & _
                        "and Substr(sp58,1,8)=c2.cu01(+) AND Decode(Substr(sp58,9,1),null,'0',Substr(sp58,9,1))=c2.cu02(+) and Substr(sp59,1,8)=c3.cu01(+) AND Decode(Substr(sp59,9,1),null,'0',Substr(sp59,9,1))=c3.cu02(+) and Substr(sp65,1,8)=c4.cu01(+) and Decode(Substr(sp65,9,1),null,'0',Substr(sp65,9,1))=c4.cu02(+) " & _
                        "and Substr(sp66,1,8)=c5.cu01(+) and Decode(Substr(sp66,9,1),null,'0',Substr(sp66,9,1))=c5.cu02(+) " & pSql05 & mSubSQL2
        'end 2022/04/11
    Next intR
    'end 2020/08/04
    
On Error GoTo ErrHandle

    cnnConnection.Execute strMainSql, intR
    
    '刪除對造與案件申請人相同資料
    'Modify by Amy 2020/09/16 R021002 對造名稱欄會寫入@@cp欄位,應過濾
    'Modify by Amy 2022/04/11 +申請人英日
    If bShowF = True Then
        strTemp1 = "Delete From R100102_1 Where ID='" & pIDname & "' And (" & _
                        "SubStr(ltrim(rtrim(R021002)),1,InStr(R021002,'@@')-1)=ltrim(rtrim(R021009)) Or SubStr(ltrim(rtrim(R021002)),1,InStr(r021002,'@@')-1)=ltrim(rtrim(R021010)) Or SubStr(ltrim(rtrim(R021002)),1,InStr(r021002,'@@')-1)=ltrim(rtrim(R021011)) Or SubStr(ltrim(rtrim(R021002)),1,InStr(r021002,'@@')-1)=ltrim(rtrim(R021012)) Or SubStr(ltrim(rtrim(R021002)),1,InStr(r021002,'@@')-1)=ltrim(rtrim(R021013)) " & _
                        "Or SubStr(ltrim(rtrim(R021002)),1,InStr(R021002,'@@')-1)=ltrim(rtrim(R021026)) Or SubStr(ltrim(rtrim(R021002)),1,InStr(r021002,'@@')-1)=ltrim(rtrim(R021027)) Or SubStr(ltrim(rtrim(R021002)),1,InStr(r021002,'@@')-1)=ltrim(rtrim(R021028)) Or SubStr(ltrim(rtrim(R021002)),1,InStr(r021002,'@@')-1)=ltrim(rtrim(R021029)) Or SubStr(ltrim(rtrim(R021002)),1,InStr(r021002,'@@')-1)=ltrim(rtrim(R021030)) " & _
                        "Or SubStr(ltrim(rtrim(R021002)),1,InStr(R021002,'@@')-1)=ltrim(rtrim(R021031)) Or SubStr(ltrim(rtrim(R021002)),1,InStr(r021002,'@@')-1)=ltrim(rtrim(R021032)) Or SubStr(ltrim(rtrim(R021002)),1,InStr(r021002,'@@')-1)=ltrim(rtrim(R021033)) Or SubStr(ltrim(rtrim(R021002)),1,InStr(r021002,'@@')-1)=ltrim(rtrim(R021034)) Or SubStr(ltrim(rtrim(R021002)),1,InStr(r021002,'@@')-1)=ltrim(rtrim(R021035)) " & _
                        ") "
    Else
    
        strTemp1 = "Delete From R100102_1 Where ID='" & pIDname & "' And (" & _
                        "ltrim(rtrim(R021002))=ltrim(rtrim(R021009)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021010)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021011)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021012)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021013))" & _
                        "Or ltrim(rtrim(R021002))=ltrim(rtrim(R021026)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021027)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021028)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021029)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021030))" & _
                        "Or ltrim(rtrim(R021002))=ltrim(rtrim(R021031)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021032)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021033)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021034)) Or ltrim(rtrim(R021002))=ltrim(rtrim(R021035))" & _
                        ") "
    End If
    'end 2022/04/11
    cnnConnection.Execute strTemp1, intR

'****** 其他相關人******
   Call Pub_ChkRelevantPeople(2)  'Modify by Amy 2025/01/22 原程式合併至Pub_ChkRelevantPeople 避免有未改到
    'Add by Amy 2021/08/13 +刪除其他相關人
    If bolDelRelevantP = True Then
        strTemp1 = "Delete From R100102_1 Where ID='" & pIDname & "' And R021004='2' "
        cnnConnection.Execute strTemp1, intR
    End If
'****** End 其他相關人 (以上有修改需確認 Pub_ChkRelevantPeople 是否也要改) ******
               
    'Add by Amy 2020/08/27
    'Modify by Amy 2023/12/28 +風險檢查資料維護 (frm12040163)
    If UCase(Mid(pIDname, Val(InStr(pIDname, "@")) + 1)) = "FRM140401" Or UCase(Mid(pIDname, Val(InStr(pIDname, "@")) + 1)) = UCase("frm12040163") Then Exit Sub
     
    '利益衝突案件: 逐案號判斷
    If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
        b_Sys = GetAllSysKind(, "ALL")
        strTemp1 = "select R021001,R021020,R021021,R021022,R021023,R021024,R021025 from R100102_1 where id = '" & pIDname & "' order by 1 "
        intR = 1
        Set rsRD = ClsLawReadRstMsg(intR, strTemp1)
        If intR = 1 Then
            rsRD.MoveFirst
            Do While Not rsRD.EOF
                If PUB_ChkCufaByCase(pIDname, b_Sys, "" & rsRD.Fields("R021001"), "" & rsRD.Fields("R021020") & "," & rsRD.Fields("R021021") & "," & rsRD.Fields("R021022") & "," & rsRD.Fields("R021023") & "," & rsRD.Fields("R021024"), "" & rsRD.Fields("R021025")) = False Then
                    intCnt = intCnt + 1
                    cnnConnection.Execute "delete from R100102_1 where id='" & pIDname & "' and r021001='" & rsRD.Fields("R021001") & "' "
                End If
                rsRD.MoveNext
            Loop
            If intCnt > 0 Then
                MsgBox "為限閱案件之對造，請通知電腦中心協助確認！", vbInformation, MsgText(1110)
            End If
        End If
        Set rsRD = Nothing
    End If
    
    Exit Sub
    
ErrHandle:
    If Err.Number <> 0 Then
         MsgBox Err.Description, vbCritical, MsgText(1110)
    End If
End Sub

'Added by Lydia 2024/01/12 利益衝突案件：輸入案號
'Modified by Lydia 2024/04/10 +pType提示狀態
Public Function PUB_ChkCufaByCaseNo(ByVal pUserNo As String, ByVal pFrmName As String, ByVal pCaseNo As String, ByVal pType As String) As Boolean
Dim strBCase(0 To 4) As String
Dim intR As Integer
Dim rsR1 As New ADODB.Recordset
'Added by Lydia 2024/04/10
Dim IntF As Integer, arrQ As Variant
Dim strChkNo As String, intA As Integer
'end 2024/04/10

   PUB_ChkCufaByCaseNo = True '可查詢
   If pCaseNo = "" Then Exit Function
   strBCase(0) = pCaseNo
   Call ChgCaseNo(strBCase(0), strBCase)
   
   If strBCase(1) <> "" And strBCase(2) <> "" Then
      strBCase(0) = " select pa01 as c01, pa02 as c02, pa03 as c03, pa04 as c04, pa26 as app01, pa27 as app02, pa28 as app03, pa29 as app04, pa30 as app05, pa75 as fcno" & _
                    " from patent where pa01='" & strBCase(1) & "' and pa02='" & strBCase(2) & "' and pa03='" & strBCase(3) & "' and pa04='" & strBCase(4) & "' "
      strBCase(0) = strBCase(0) & "union select sp01 as c01, sp02 as c02, sp03 as c03, sp04 as c04, sp08 as app01, sp58 as app02, sp59 as app03, sp65 as app04, sp66 as app05, sp26 as fcno" & _
                    " from servicepractice where sp01='" & strBCase(1) & "' and sp02='" & strBCase(2) & "' and sp03='" & strBCase(3) & "' and sp04='" & strBCase(4) & "' "
      '保留
      'strBCase(0) = strBCase(0) & "union select tm01 as c01, tm02 as c02, tm03 as c03, tm04 as c04, tm23 as app01, tm78 as app02, tm79 as app03, tm80 as app04, tm81 as app05, tm44 as fcno" & _
                    " from trademark where tm01='" & strBCase(1) & "' and tm02='" & strBCase(2) & "' and tm03='" & strBCase(3) & "' and tm04='" & strBCase(4) & "' "
      'strBCase(0) = strBCase(0) & "union select lc01 as c01, lc02 as c02, lc03 as c03, lc04 as c04,lc11 as app01,lc43 as app02,lc44 as app03,lc45 as app04,lc46 as app05, lc22 as fcno" & _
                    " from lawcase where lc01='" & strBCase(1) & "' and lc02='" & strBCase(2) & "' and lc03='" & strBCase(3) & "' and lc04='" & strBCase(4) & "' "
      'strBCase(0) = strBCase(0) & "union select hc01 as c01, hc02 as c02, hc03 as c03, hc04 as c04,hc05 as app01,hc24 as app02,hc25 as app03,hc26 as app04,hc27 as app05, '' as fcno" & _
                    " from hirecase where hc01='" & strBCase(1) & "' and hc02='" & strBCase(2) & "' and hc03='" & strBCase(3) & "' and hc04='" & strBCase(4) & "' "
      intR = 1
      Set rsR1 = ClsLawReadRstMsg(intR, strBCase(0))
      If intR = 1 Then
         If PUB_ChkCufaByCase(pFrmName, rsR1.Fields("c01"), pCaseNo, rsR1.Fields("App01") & "," & rsR1.Fields("App02") & "," & rsR1.Fields("App03") & "," & rsR1.Fields("App04") & "," & rsR1.Fields("App05"), "" & rsR1.Fields("fcno"), pUserNo) = False Then
            PUB_ChkCufaByCaseNo = False 'Memo by Lydia 2024/04/18 因為PUB_ChkCufaByCase=False
            'Added by Lydia 2024/04/10
            If pType = "2" Then  '列出有權限的工程師
               arrQ = Split(rsR1.Fields("App01") & "," & rsR1.Fields("App02") & "," & rsR1.Fields("App03") & "," & rsR1.Fields("App04") & "," & rsR1.Fields("App05") & "," & rsR1.Fields("fcno"), ",")
               For IntF = 0 To UBound(arrQ)
                  If Trim("" & arrQ(IntF)) <> "" Then
                     strChkNo = Left(ChangeCustomerL("" & arrQ(IntF)), 8)
                     If InStr(XY特殊權限範圍, strChkNo) > 0 Then
                        strBCase(0) = "select cfr07 from cufa_right where cfr01='" & strChkNo & "' and cfr03='M51' and cfr02='" & strBCase(1) & "' "
                        intR = 1
                        Set rsR1 = ClsLawReadRstMsg(intR, strBCase(0))
                        If intR = 1 Then
                           '分案和工作進度維護點選不可查閱工程師需要彈訊息
                           If "" & rsR1.Fields("cfr07") = "Y" Then
                              strBCase(0) = "select cfr03||' '||NVL(B1.St02,D1.A0902) as 工程師" & _
                                          " From (Select Cfr01,Decode(Cu01,Null,Nvl(Fa04,Nvl(Fa05,Fa06)),Nvl(Cu04,Nvl(Cu05,Cu06))) As Cfr_Name,Cfr03" & _
                                          " ,Listagg(Cfr02,',') Within Group (Order By Cfr02) As Syslist From Cufa_Right,Customer,Fagent" & _
                                          " Where Cfr01=Cu01(+) And '0'=Cu02(+) And Cfr01=Fa01(+) And '0'=Fa02(+) AND CFR01='" & strChkNo & "'" & _
                                          " Group By Cfr01,Cfr03,Decode(Cu01,Null,Nvl(Fa04,Nvl(Fa05,Fa06)),Nvl(Cu04,Nvl(Cu05,Cu06)))) Vt1, staff b1, staff_level c1 ,acc090 d1" & _
                                          " where cfr03=b1.st01(+) and cfr03=c1.sl01(+) and cfr03=d1.a0901(+) and b1.st04='1' and b1.st03 in ('P11','F21')"
                              intR = 1
                              Set rsR1 = ClsLawReadRstMsg(intR, strBCase(0))
                              If intR = 1 Then
                                 MsgBox "本案為限閱案件，請輸入下方有權限的工程師：" & vbCrLf & vbCrLf & rsR1.GetString(adClipString, , , vbCrLf), vbOKOnly + vbInformation, MsgText(1110)
                              End If
                           'Added by Lydia 2024/04/22 與Wilison確認只有Photon Control(X88787000)、Advanced Energy(Y48904000, X48904000)、VAT(X84599000)，才需要控制工程師沒有權限不能分案給該工程師
                           Else
                              PUB_ChkCufaByCaseNo = True '沒有限閱權限也可以分案
                           'end 2024/04/22
                           End If
                        End If
                        Exit For
                     End If
                  End If
               Next IntF
               
            Else '---預設只提醒限閱案件
            'end 2024/04/10
                MsgBox "限閱案件", vbInformation, MsgText(1110)
            End If
         End If
      End If
   End If
   
   Set rsR1 = Nothing
End Function

'Added by Lydia 2019/11/01 利益衝突案件：逐案號，檢查申請人1~5和FC代理人
'Modified by Lydia 2023/08/14 +傳入人員代號
Public Function PUB_ChkCufaByCase(ByVal pFrmName As String, ByVal pAllSys As String, ByVal pCaseNo As String, ByVal pCustList As String, ByVal pFCno As String, Optional ByVal pUserNo As String) As Boolean
'pFrmName :表單名稱
'pAllSys: 預設可使用的全部系統別
'pCaseNo : 傳入串接的本所案號
'pCustList:串接申請人1~5
'pFCno: FC代理人
Dim IntF As Integer, arrQ As Variant
Dim pNo(1 To 4) As String
Dim strChkNo As String, intA As Integer
Dim strA1 As String
Dim rsA1 As New ADODB.Recordset
Dim TmpArea As String, tmpRight As String
Dim tmpConPA As String, tmpConSP As String
Dim strExcept As String 'Added by Lydia 2022/03/16
Dim pUserSt03 As String 'Added by Lydia 2023/08/14 傳入人員代號部門

    PUB_ChkCufaByCase = True '可查詢
    If Not (strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "") Then Exit Function
    
    '如果沒有X,Y編號
    If Trim(Replace(pCustList & "," & pFCno, ",", "")) = "" Then
         Exit Function
    End If
    
    'Added by Lydia 2023/08/14 傳入人員代號
    If pUserNo = "" Then
       pUserNo = strUserNum
    End If
    pUserSt03 = PUB_GetST03(pUserNo)
    'end 2023/08/14
    
    arrQ = Split(pCustList & "," & pFCno, ",")
    'Modified by Lydia 2023/08/14 +電腦名稱@pub_HostName
    cnnConnection.Execute "delete from R100102_2 where R02201='" & strUserNum & "@" & pub_HostName & "' and R02202='" & pFrmName & "' "   '清空暫存檔
'Added by Lydia 2022/03/16 例外處理------------------------
    strChkNo = Replace(Pub_RplStr(pCaseNo), "-", "")
    'Mark by Lydia 2022/06/17 長庚財團法人集團(X69365000、X69365010、X69365020、X69365050、X69365060、X75299020)管制：專利案件FCP, FG, P, PS,CFP,CPS
    'If strChkNo = "P125723000" Then
    '    strExcept = "P125723000"
    '    'P-125723開放下列人可以查詢
    '    '林總94007、王副總71011、專利國內部P程序(等級73及75)、柯昱安A7010、顧服組(等級W0及W2)、電腦中心(部門M51)
    '    '3/16+ 簡玉如=> 另外設特殊群組「 限閱案件例外-P125723」
    '    strA1 = Pub_GetSpecMan("限閱案件例外-P125723")
    '    If InStr(strA1, strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Or InStr("W0,W2,73,75", Pub_strUserST05) > 0 Then
    '       PUB_ChkCufaByCase = True
    '    Else
    '       PUB_ChkCufaByCase = False
    '    End If
    'End If
    'end 2022/06/17
    'Added by Lydia 2024/02/06 針對客戶X83049 Infineon Technologies LLC案件之限閱設定：開放下列案件
    If InStr(pCustList, "X8304900") > 0 And strChkNo <> "" Then
       'Modified by Lydia 2025/03/24 +FCP-073175
       'Modified by Lydia 2025/03/25 +FCP-071295
       'Modified by Lydia 2025/04/21 +FCP-072560
       'Modified by Lydia 2025/11/14 +FCP-070851
       If InStr("FCP064219000,FCP067920000,FCP064220000,FCP067921000,FCP064571000,FCP067922000,FCP064570000,FCP068009000,FCP066827000,FCP068148000,FCP067733000,FCP069615000,FCP073175000,FCP071295000,FCP072560000,FCP070851000", strChkNo) > 0 Then
          strExcept = "限閱案件例外-X83049"
          PUB_ChkCufaByCase = True
       End If
    End If
    'end 2024/02/06
    'Added by Lydia 2024/03/21 開放FCP-063755個案的查詢權限給A4099
    'Modified by Lydia 2024/03/25 改成A4023
    'Mark by Lydia 2024/04/18 已開放申請人權限給特助
    'If strChkNo = "FCP063755000" And pUserNo = "A4023" Then
    '    strExcept = "限閱案件例外-FCP063755000"
    '    PUB_ChkCufaByCase = True
    'End If
    'end 2024/03/21
'------------------------------------------------------------------------
    If strExcept = "" Then
    'end 2022/03/16
        For IntF = 0 To UBound(arrQ)
            If Trim("" & arrQ(IntF)) <> "" Then
                strChkNo = Left(ChangeCustomerL("" & arrQ(IntF)), 8)
                If InStr(XY特殊權限範圍, strChkNo) > 0 Then
                    'Modified by Lydia 2023/08/14
                    If PUB_ChkCuFa_Right(pFrmName, strChkNo, pAllSys, tmpRight, TmpArea, pUserNo) = True Then
                    End If
                    'Added by Lydia 2020/04/16 如果案件代理人不屬於利益衝突管制(限閱案件)，開放權限給負責該區代理人的FCP承辦人員;
                    'ex.4/6 FCP56057住友電木X4885100(限閱案件)，代理人為Y52694000，相關請款仍是由英文組承辦人員負責。
                    'Modified by Lydia 2023/08/14 傳入人員代號
                    'If tmpRight = "" And Pub_StrUserSt03 = "F23" And Left(pFCno, 1) = "Y" And InStr(XY特殊權限範圍, pFCno) = 0 Then
                    '    strA1 = "select na51 from fagent, nation where fa01||fa02='" & ChangeCustomerL(pFCno) & "' and fa10=na01(+) and na51='" & strUserNum & "' "
                    If tmpRight = "" And pUserSt03 = "F23" And Left(pFCno, 1) = "Y" And InStr(XY特殊權限範圍, pFCno) = 0 Then
                        strA1 = "select na51 from fagent, nation where fa01||fa02='" & ChangeCustomerL(pFCno) & "' and fa10=na01(+) and na51='" & pUserNo & "' "
                    'end 2023/08/14
                        intA = 1
                        Set rsA1 = ClsLawReadRstMsg(intA, strA1)
                        If intA = 1 Then
                            tmpRight = TmpArea
                        End If
                    End If
                    'end 2020/04/16
                    '有管制系統別=>組合SQL條件
                    If TmpArea <> "" Then
                        'Added by Lydia 2023/08/15 排除frm000001->利益衝突權限檢查
                        If pCaseNo = "FCP000001000" Then
                           If tmpRight <> "" Then
                              PUB_ChkCufaByCase = True
                           Else
                              PUB_ChkCufaByCase = False
                           End If
                        Else
                        'end 2023/08/15
                           strA1 = Replace(Pub_RplStr(pCaseNo), "-", "")
                           Call ChgCaseNo(strA1, pNo)
                           'Modified by Lydia 2019/12/23 2019/12/23 判斷是否為管制的系統號 ( ex.FCL-010722代理人Y53715只管制專利和服務)
                           'If TmpArea <> "" Then
                           If TmpArea <> "" And InStr(TmpArea, pNo(1)) > 0 Then
                               tmpConPA = Pub_CufaConSQL(pFrmName, "PA", strChkNo, tmpRight, TmpArea)
                               tmpConSP = Pub_CufaConSQL(pFrmName, "SP", strChkNo, tmpRight, TmpArea)
                               
                               strA1 = "select pa01,pa02,pa03,pa04 from patent where pa01='" & pNo(1) & "' and pa02='" & pNo(2) & "' and pa03='" & pNo(3) & "' and pa04='" & pNo(4) & "' " & tmpConPA
                               strA1 = strA1 & " union all select sp01,sp02,sp03,sp04 from servicepractice where sp01='" & pNo(1) & "' and sp02='" & pNo(2) & "' and sp03='" & pNo(3) & "' and sp04='" & pNo(4) & "' " & tmpConSP
                               intA = 1
                               Set rsA1 = ClsLawReadRstMsg(intA, strA1)
                               If intA = 1 Then '有符合權限，就不再判斷
                                   PUB_ChkCufaByCase = True
                                   Exit For
                               Else
                                   PUB_ChkCufaByCase = False
                               End If
                           End If
                        End If 'Added by Lydia 2023/08/15
                        'Added by Lydia 2022/08/16 例外處理：針對美國Y52694000 Promerus+X48851000住友電木的案件，開放給案件工程師和承辦主管
                        If PUB_ChkCufaByCase = False And pNo(1) <> "" And pNo(2) <> "" And InStr("FCP046011000,FCP047138000,FCP047234000,FCP049176000,FCP049961000,FCP050582000,FCP050583000,FCP052852000,FCP053721000,FCP053951000,FCP053952000,FCP054464000,FCP055248000,FCP055880000,FCP056057000,FCP056926000,FCP057165000,FCP058128000,FCP058553000,FCP060410000,FCP061100000,FCP061351000,FCP062454000,FCP062831000,FCP063235000", pNo(1) & pNo(2) & pNo(3) & pNo(4)) > 0 Then
                             strA1 = Pub_GetSpecMan("限閱案件例外-住友")
                             If InStr(strA1, strUserNum) > 0 Then
                                PUB_ChkCufaByCase = True
                                Exit For
                             Else
                                PUB_ChkCufaByCase = False
                             End If
                        End If
                        'end 2023/08/16
                    End If
                'Added by Lydia 2020/10/21 固定控制
                'Mark by Lydia 2020/10/29 因為後來修改為列出開放限閱工程師,所以先保留
                'ElseIf strSrvDate(1) >= "20221111" Then  'Memo by Lydia 2020/10/21 先隱藏，等確定上線日期
                '    '德國 Fresenius(X81804000,X19893020)的開放工程師不可看對手的案件
                '    If InStr("X4838300,Y5474500", strChkNo) > 0 Or InStr("X55340,Y21199,X72396, X80572,X49045", Left(strChkNo, 6)) > 0 Then
                '        '對手編號：B. Braun Melsungen AG(X48383000,Y54745000)、Baxter (X55340Y21199)、Nikkiso(X72396,X80572)、Nipro Corporation(X49045)
                '        If InStr("88003,94007", strUserNum) = 0 And Pub_StrUserSt03 = "F21" Then '開放工程師：排除主管
                '            strA1 = "Select * From Cufa_Right Where Instr('X8180400,X1989302',CFR01) > 0 And CFR03='" & strUserNum & "' "
                '            intA = 1
                '            Set rsA1 = ClsLawReadRstMsg(intA, strA1)
                '            If intA = 1 Then
                '                 PUB_ChkCufaByCase = False
                '            End If
                '        End If
                '    End If
                ''end 2020/10/21
                'end 2020/10/29
                End If
            End If
        Next IntF
    End If 'Added by Lydia 2022/03/16
ExitSUB01:
    Set rsA1 = Nothing
End Function

'Added by Lydia 2019/11/01 利益衝突案件：檢查利益衝突案件之權限(XY特殊權限範圍)
'Modified by Lydia 2023/08/14 +傳入人員代號
Public Function PUB_ChkCuFa_Right(ByVal pFrmName As String, ByVal pNo As String, ByVal pSys As String, ByRef outRight As String, ByRef outArea As String, Optional ByVal pUserNo As String) As Boolean
'pNo: 檢查X/Y編號
'pSys: 檢查系統別: 空白,ALL=>FCP, FG, CFP, PS, CPS
'outRight : 可使用的範圍
'outArea : X/Y編號的管制系統類別
'===============================
'檢查方式:
' 1.以X/Y編號＋系統類別＋操作員工編號檢查
' 2.以X/Y編號＋系統類別＋操作者部門檢查
' 3.檢查該案件是為"初審未審定(限PA16 is null or <>'1' )"階段，開放給案件的所有承辦工程師，案件已准則不開放；
              '若初審被駁之案件在再審(107)階段未審定 ，開放給案件的所有承辦工程師。
'前3項其中之一符合即可
'===============================
Dim strB1 As String
Dim strChkNo As String 'X/Y編號，取8碼
Dim strConB As String, StrSqlB As String
Dim intJ As Integer
Dim rsB As New ADODB.Recordset
Dim pUserSt03 As String, pUserSt05 As String 'Added by Lydia 2023/08/14 傳入人員代號部門,等級

On Error GoTo ExitProc

    PUB_ChkCuFa_Right = False
    
    'X/Y編號
    If InStr("X,Y", Left(pNo, 1)) > 0 Then
        strChkNo = Left(ChangeCustomerL(pNo), 8)
    End If
    'Added by Lydia 2023/08/14 傳入人員代號
    If pUserNo = "" Then
       pUserNo = strUserNum
    End If
    pUserSt03 = PUB_GetST03(pUserNo)
    pUserSt05 = PUB_GetST05(pUserNo)
    'end 2023/08/14
    
    outRight = ""
    outArea = ""
    'Modified by Lydia 2021/03/31 +CFR04
    'Modified by Lydia 2025/08/01 增加關聯企業(CFR01=前6碼)的判斷
    'strConB = "SELECT distinct(CFR02) AS CFRArea FROM CUFA_RIGHT WHERE CFR01='" & strChkNo & "' AND CFR04='0' "
    strConB = "SELECT distinct(CFR02) AS CFRArea FROM CUFA_RIGHT WHERE (CFR01='" & strChkNo & "' OR CFR01='" & Mid(strChkNo, 1, 6) & "') AND CFR04='0' "
    '先取得X/Y編號的管制系統類別(用M51為條件)
    intJ = 1
    Set rsB = ClsLawReadRstMsg(intJ, strConB & "and CFR03='M51' order by 1")
    If intJ = 1 Then
         outArea = "" & rsB.GetString(adClipString, , , ",")
         If Right(outArea, 1) = "," Then outArea = Mid(outArea, 1, Len(outArea) - 1)
    End If
    If outArea = "" Then
         '不管制
         PUB_ChkCuFa_Right = True
         GoTo ExitProc
    End If
    '+系統類別
    If InStr(pSys, "'") = 0 Then pSys = GetAddStr(pSys)
    strConB = strConB & IIf(pSys = "" Or pSys = "ALL", "", " AND CFR02 IN (" & pSys & ") ")

' 1.以X/Y編號＋系統類別＋操作員工編號檢查；Memo by Lydia 2025/08/07 若條件有異動，請一併調整PUB_SaveCUFA_Staff_Log
    'Modifed by Lydia 2023/08/14 傳入人員代號
    'strB1 = strConB & " and cfr03='" & strUserNum & "' "
    strB1 = strConB & " and cfr03='" & pUserNo & "' "
    intJ = 1
    Set rsB = ClsLawReadRstMsg(intJ, strB1 & " order by 1")
    If intJ = 1 Then
        outRight = "" & rsB.GetString(adClipString, , , ",")
    End If
    If outRight = "" Then
' 2.以X/Y編號＋系統類別＋操作者部門檢查；Memo by Lydia 2025/08/07 若條件有異動，請一併調整PUB_SaveCUFA_Staff_Log
        'Modified by Lydia 2023/08/14 傳入人員代號
        'strB1 = strConB & " and cfr03='" & Pub_StrUserSt03 & "' "
        strB1 = strConB & " and cfr03='" & pUserSt03 & "' "
        intJ = 1
        Set rsB = ClsLawReadRstMsg(intJ, strB1 & " order by 1")
        If intJ = 1 Then
            outRight = "" & rsB.GetString(adClipString, , , ",")
        End If
    End If
    'Added by Lydia 2019/12/23
    If outRight = "" Then
' 3.以X/Y編號＋系統類別＋操作者等級檢查；Memo by Lydia 2025/08/07 若條件有異動，請一併調整PUB_SaveCUFA_Staff_Log
        'Modified by Lydia 2023/08/14 傳入人員代號
        'strB1 = strConB & " and cfr03='" & Pub_strUserST05 & "' "
        strB1 = strConB & " and cfr03='" & pUserSt05 & "' "
        intJ = 1
        Set rsB = ClsLawReadRstMsg(intJ, strB1 & " order by 1")
        If intJ = 1 Then
            outRight = "" & rsB.GetString(adClipString, , , ",")
        End If
    End If
    'end 2019/12/23
    
    'Added by Lydia 2020/10/29
    If outRight = "" Then
' 4.以X/Y編號＋系統類別＋CFR03=操作者部門(ST03)+工程師組別(ST16)檢查；Memo by Lydia 2025/08/07 若條件有異動，請一併調整PUB_SaveCUFA_Staff_Log
       '=> 同一X/Y編號(前8碼)工程師用不同組別設定開放權限 ex. 德國 Fresenius Medical Care (X81804000,X19893020)設定
        'Modified by Lydia 2023/08/14 傳入人員代號
        'strB1 = strConB & " AND CFR03 IN (SELECT ST03||ST16 FROM STAFF WHERE ST01='" & strUserNum & "')"
        'Modified by Lydia 2025/08/01 增加CFR03=操作者部門(ST03)+FCP工程師組別(ST16)+日本部組別(ST70,114/8/1新增)檢查
        'strB1 = strConB & " AND CFR03 IN (SELECT ST03||ST16 FROM STAFF WHERE ST01='" & pUserNo & "')"
        strB1 = strConB & " AND (CFR03 IN (SELECT ST03||ST16 FROM STAFF WHERE ST01='" & pUserNo & "') OR CFR03 IN (SELECT ST03||ST16||ST70 FROM STAFF WHERE ST01='" & pUserNo & "'))"
        intJ = 1
        Set rsB = ClsLawReadRstMsg(intJ, strB1 & " order by 1")
        If intJ = 1 Then
            outRight = "" & rsB.GetString(adClipString, , , ",")
        End If
    End If
    'end 2020/10/29
    
    'Modified by Lydia 2019/12/23 限外專工程師
    'Modified by Lydia 2023/08/14 傳入人員代號
    'If outRight = "" And Pub_StrUserSt03 = "F21" Then
    If outRight = "" And pUserSt03 = "F21" Then
         'Added by Lydia 2021/03/31
'5.以X/Y編號＋系統類別＋操作員工編號＋CFR04=開放工程師組別檢查；Memo by Lydia 2025/08/07 因為與條件1.相同，不用調整PUB_SaveCUFA_Staff_Log
         '=>因為無法區分X/Y編號第9碼，所以判斷同一X/Y編號(前8碼)+符合開放組別之案號，暫存在利益衝突個案權限檔R100102_2
         'Modified by Lydia 2023/08/14 傳入人員代號
         'strB1 = "SELECT A1.* FROM CUFA_RIGHT A1, SYSTEMKIND A2 WHERE CFR03='" & strUserNum & "' AND CFR04<>'0' AND CFR02=SK01(+) AND SK02='1' ORDER BY CFR01,CFR02 "
         strB1 = "SELECT A1.* FROM CUFA_RIGHT A1, SYSTEMKIND A2 WHERE CFR03='" & pUserNo & "' AND CFR04<>'0' AND CFR02=SK01(+) AND SK02='1' ORDER BY CFR01,CFR02 "
         intJ = 1
         Set rsB = ClsLawReadRstMsg(intJ, strB1)
         If intJ = 1 Then
             rsB.MoveFirst
             strConB = ""
             Do While Not rsB.EOF
                 'Modified by Lydia 2023/08/14 +電腦名稱@pub_HostName
                 strConB = strConB & "Union select '" & strUserNum & "@" & pub_HostName & "', '" & pFrmName & "', '" & strChkNo & "', pa01,pa02,pa03,pa04 from patent where " & _
                               IIf(Left("" & rsB.Fields("cfr01"), 1) = "X", " instr(pa26||','||pa27||','||pa28||','||pa29||','||pa30,'" & rsB.Fields("cfr01") & "') > 0 ", " pa75 like '" & rsB.Fields("cfr01") & "%' ") & _
                               " and pa01= '" & rsB.Fields("cfr02") & "' and pa150='" & rsB.Fields("cfr04") & "' "
                 rsB.MoveNext
             Loop
             If strConB <> "" Then
                 cnnConnection.Execute " insert into R100102_2 (R02201,R02202,R02203,R02204,R02205,R02206,R02207) " & Mid(strConB, 6), intJ
             End If
         End If
         'end 2021/03/31
         
'6 分階段檢查該案件的承辦工程師可看的案件；
         'Modified by Lydia 2023/08/14 傳入人員代號
         'strConB = "select cp01,cp02,cp03,cp04 from caseprogress,staff where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp159=0 and cp14='" & strUserNum & "' and cp14=st01(+) and st03='F21' "
         strConB = "select cp01,cp02,cp03,cp04 from caseprogress,staff where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp159=0 and cp14='" & pUserNo & "' and cp14=st01(+) and st03='F21' "
         '----專利案的承辦人
         'Memo by Lydia 2019/12/23 重整：
         '０. 初審未審定(限PA16 is null or <>'1' )階段：開放給案件的所有承辦工程師
         '１. 案件已准(PA16='1')階段：預設不開放，若已核准但核對已准專利未發文則開放給926承辦工程師；
         '２. 若案件在再審(107)階段未審定(CP24=null)，開放給再審後(含再審)的承辦工程師。
         '　閉卷 / 銷卷則不開放給承辦工程師
         'Modified by Lydia 2019/12/23  閉卷/銷卷則不開放給承辦工程師(pa57||pa108 is null)
         'Modified by Lydia 2019/12/23 +已核准但核對已准專利未發文的承辦工程師; 僅開放給再審後(含再審)的承辦工程師
         'strB1 = "select '" & strUserNum & "', '" & pFrmName & "', '" & strChkNo & "', pa01,pa02,pa03,pa04 from patent where pa01 In ('FCP','CFP','P') and pa57||pa108 is null and " & _
                      IIf(Left(strChkNo, 1) = "X", " instr(pa26||','||pa27||','||pa28||','||pa29||','||pa30,'" & strChkNo & "') > 0 ", " substr(pa75,1,8)='" & strChkNo & "' ") & _
                     "and (pa01,pa02,pa03,pa04) in (" & strConB & ") " & _
                     "and (nvl(pa16,'0')='0' or (pa16='2' and (pa01,pa02,pa03,pa04) in (select cp01,cp02,cp03,cp04 from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp09 like 'A%' and cp10='107' and cp158>0 and cp159=0 and cp24 is null)) ) "
         'Modified by Lydia 2023/08/14 +電腦名稱@pub_HostName
         strB1 = "select '" & strUserNum & "@" & pub_HostName & "', '" & pFrmName & "', '" & strChkNo & "', pa01,pa02,pa03,pa04 from patent where pa01 In ('FCP','CFP','P') and pa57||pa108 is null and " & _
                     IIf(Left(strChkNo, 1) = "X", " instr(pa26||','||pa27||','||pa28||','||pa29||','||pa30,'" & strChkNo & "') > 0 ", " substr(pa75,1,8)='" & strChkNo & "' ")
         'Remove by Lydia 2020/05/11 簡化判斷;  109/05/08增加開放未發文進度之承辦工程師能有權限查詢(ex.P-117747李正揚)
         'strB1 = strB1 & "and (pa01,pa02,pa03,pa04) in (" & strConB & " and cp05>=(select nvl(max(cp05),'19221111') mdate from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp09 like 'A%' and cp10='107' and cp159=0 ) ) "
         'strB1 = strB1 & "and (nvl(pa16,'0')='0' or (pa16='2' and (pa01,pa02,pa03,pa04) in (select cp01,cp02,cp03,cp04 from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp09 like 'A%' and cp10='107' and cp158>0 and cp159=0 and cp24 is null)) "
         'strB1 = strB1 & "or (pa16='1' and (pa01,pa02,pa03,pa04) in (select cp01,cp02,cp03,cp04 from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='926' and cp158=0 and cp159=0 and cp14='" & strUserNum & "')) )"
         
         'Added by Lydia 2020/05/11 簡化判斷(109/05/12以後)
         '1. 案件未有准駁 (PA16 Is Null): 開放給案件的所有承辦工程師
         '2. 案件有准駁 (PA16 Is Not Null): 開放給案件的所有未發文承辦工程師
         '3. 若案件在再審(107)階段未審定(CP24=null)，開放給再審後(含再審)的承辦工程師。
         '4. 閉卷/銷卷則不開放
         StrSqlB = strB1 & "and ( " & _
                    "(pa01,pa02,pa03,pa04) in (" & strConB & " and cp05>=(select nvl(max(cp05),'99999999') mdate from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp09 like 'A%' and cp10='107' and cp159=0 and cp24 is null))"  '再審階段
                    StrSqlB = StrSqlB & " or (nvl(pa16,'0')='0' and (pa01,pa02,pa03,pa04) in (" & strConB & "))"  '案件未有准駁
                    StrSqlB = StrSqlB & " or (nvl(pa16,'0')<>'0' and (pa01,pa02,pa03,pa04) in (" & strConB & " and cp158=0))"  '案件有准駁
         StrSqlB = StrSqlB & " ) "
         'Added by Lydia 2023/11/22 增加新案翻譯的核稿人權限=>一般工程師承辦; ex.FCP-070443
         strConB = "select cp01,cp02,cp03,cp04 from caseprogress,staff,engineerprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp159=0 and cp10 in ('201') and cp09=ep02(+) and ep04=st01(+) and ep04='" & pUserNo & "' and st03='F21' "
         StrSqlB = StrSqlB & " Union " & strB1 & "and ( " & _
                    "(pa01,pa02,pa03,pa04) in (" & strConB & " and cp05>=(select nvl(max(cp05),'99999999') mdate from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp09 like 'A%' and cp10='107' and cp159=0 and cp24 is null))"  '再審階段
                    StrSqlB = StrSqlB & " or (nvl(pa16,'0')='0' and (pa01,pa02,pa03,pa04) in (" & strConB & "))"  '案件未有准駁
                    StrSqlB = StrSqlB & " or (nvl(pa16,'0')<>'0' and (pa01,pa02,pa03,pa04) in (" & strConB & " and cp158=0))"  '案件有准駁
         StrSqlB = StrSqlB & " ) "
         'end 2023/11/22
         
         '----服務案的承辦人
         'Modified by Lydia 2019/12/23  閉卷/銷卷則不開放給承辦工程師(sp15||sp61 is null)
         'Modified by Lydia 2023/08/14 +電腦名稱@pub_HostName
         StrSqlB = StrSqlB & "union select '" & strUserNum & "@" & pub_HostName & "', '" & pFrmName & "', '" & strChkNo & "', sp01,sp02,sp03,sp04 from servicepractice where sp01 in ('FG','PS','CPS') and sp15||sp61 is null and " & _
                     IIf(Left(strChkNo, 1) = "X", " instr(sp08||','||sp58||','||sp59||','||sp65||','||sp66,'" & strChkNo & "') > 0 ", " substr(sp26,1,8)='" & strChkNo & "' ") & _
                     "and (sp01,sp02,sp03,sp04) in (" & Replace(strConB, "pa", "sp") & ") "
         cnnConnection.Execute " insert into R100102_2 (R02201,R02202,R02203,R02204,R02205,R02206,R02207) " & StrSqlB, intJ
    End If

    If outRight <> "" Then
       If Right(outRight, 1) = "," Then outRight = Mid(outRight, 1, Len(outRight) - 1)
       PUB_ChkCuFa_Right = True
    End If
        
ExitProc:
    Set rsB = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, MsgText(1110)
        Resume Next
    End If
    Exit Function
End Function


'Added by Lydia 2019/11/01 利益衝突案件：輸入單一X/Y編號，組合SQL條件 (by各程式)
Public Function Pub_CufaConSQL(ByVal iFrmName As String, ByVal iKind As String, ByVal iChkNo As String, ByVal iRight As String, ByVal iArea As String) As String
Dim arrCon As Variant
Dim intR As Integer
Dim strMidCon  As String
Dim strChkNo As String 'X/Y編號，取8碼
Dim rsR1 As New ADODB.Recordset

    Pub_CufaConSQL = ""
    If iRight = "" And iArea = "" Then Exit Function
        
    strChkNo = Left(ChangeCustomerL(iChkNo), 8)
    'Modified by Lydia 2023/08/17 +電腦名稱@pub_HostName
    strMidCon = "select R02204||R02205||R02206||R02207 as caseno from R100102_2 " & _
                       "where R02201='" & strUserNum & "@" & pub_HostName & "' and R02202='" & iFrmName & "' and R02203='" & strChkNo & "' group by R02204||R02205||R02206||R02207 "
    intR = 1
    Set rsR1 = ClsLawReadRstMsg(intR, strMidCon)
    If intR = 1 Then
        strMidCon = GetAddStr(rsR1.GetString(adClipString, , , ","))
    Else
        strMidCon = ""
    End If
    Set rsR1 = Nothing

    If strMidCon = "" Then
        arrCon = Split(iArea, ",")
        For intR = 0 To UBound(arrCon)
           If Trim(arrCon(intR)) <> "" Then
              If CUFA_CheckRight1(arrCon(intR), iRight) = False Then
                 Select Case iKind
                     Case "PA" '專利
                         If Left(strChkNo, 1) = "X" Then
                            strMidCon = strMidCon & " OR (instr(PA26||','||PA27||','||PA28||','||PA29||','||PA30, '" & strChkNo & "') > 0 and PA01='" & arrCon(intR) & "')"
                         ElseIf Left(strChkNo, 1) = "Y" Then
                            strMidCon = strMidCon & " OR (instr(PA75, '" & strChkNo & "') > 0 and PA01='" & arrCon(intR) & "')"
                         End If
                     Case "SP" '服務業務
                         If Left(strChkNo, 1) = "X" Then
                            strMidCon = strMidCon & " OR (instr(SP08||','||SP58||','||SP59||','||SP65||','||SP66, '" & strChkNo & "') > 0 and SP01='" & arrCon(intR) & "')"
                         ElseIf Left(strChkNo, 1) = "Y" Then
                            strMidCon = strMidCon & " OR (instr(SP26, '" & strChkNo & "') > 0 and SP01='" & arrCon(intR) & "')"
                         End If
                 End Select
              End If
           End If
        Next intR
        If strMidCon <> "" Then strMidCon = " AND NOT(" & Mid(strMidCon, 4) & ") "
        
    Else  '本所案號
        If iKind = "PA" Then
              strMidCon = " AND (PA01||PA02||PA03||PA04) IN (" & strMidCon & ") "
        ElseIf iKind = "SP" Then
              strMidCon = " AND (SP01||SP02||SP03||SP04) IN (" & strMidCon & ") "
        End If
    End If
    Pub_CufaConSQL = strMidCon
    Set rsR1 = Nothing
End Function

'Added by Lydia 2019/11/01 利益衝突案件：從控制的系統別，比對是否有權限
Private Function CUFA_CheckRight1(ByVal pArea As String, ByVal pRights As String) As Boolean
'pArea : 控制的系統別
'pRights: 所有的權限
Dim arrRight As Variant
Dim intA As Integer

    CUFA_CheckRight1 = False
    
    If pRights = "" Then
         '全部-無權限
    Else
         '有全部權限和部份權限
         arrRight = Split(pRights, ",")
         For intA = 0 To UBound(arrRight)
              If Trim(arrRight(intA)) <> "" And Trim(arrRight(intA)) = pArea Then
                  CUFA_CheckRight1 = True '逐一比對,有權限
                  Exit For
              End If
         Next intA
    End If
    Exit Function
     
End Function
'end 2022/08/04 從basPublic搬過來

'Add by Amy 2022/09/01 從frmacc1127搬過來
Public Function GetA4112() As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    GetA4112 = ""
    
    strQ = "Select Max(A4112) From acc410 Where A4112 is not null"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        GetA4112 = ChangeTStringToTDateString(RsQ.Fields(0))
    End If
    RsQ.Close
    Set RsQ = Nothing
End Function

'Add by Amy 2022/09/02 確認是否為「其他相關人」
'intChoose:0-傳入系統別及案件性質/1-傳入CP09總收文號/2-申請人查詢更新暫存檔
'Modify by Amy 2025/01/22 原:stNo As String 改為 Optional ByVal
Public Function Pub_ChkRelevantPeople(intChoose As Integer, Optional ByVal stNo As String, Optional ByVal stCP10 As String = "") As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer, stSysKind As String
    Dim stCaseNo1 As String, strTemp1 As String, intR As Integer 'Add by Amy 2025/01/22
    
    Pub_ChkRelevantPeople = False
    
    'Add by Amy 2025/01/22 避免程式有未改到將 [申請人查詢更新暫存檔] 搬至此
    If intChoose = 2 Then
'*** !!! [申請人查詢更新暫存檔] 此有修改,下方程式也要修改 !!! ***

      '第1句
      'Add by Amy 2022/09/26 FCT/T 審查報告1201/其他來函1706/核駁1002 者狀態改為 其他相關人
      'modify by sonia 2023/12/27 加入大陸案之1205部分核駁 T-225197(胡緯民)
      strTemp1 = "Update R100102_1 Set R021004='2' Where R021014 in('FCT','T') And (R021018='1002' or R021018='1201' or R021018='1205' or R021018='1706')"
      cnnConnection.Execute strTemp1, intR
      
      '第2句
      '將所有商標案InStr(R021014,'T')且案件性質為1202(核駁前先行通知)者狀態改為 其他相關人
      '增加商標案(CFC/S) 案件性質 申請意見書(202)及延期(303)者 狀態改為 其他相關人
      'modify by sonia 2022/3/24 +305催審,703不續辦,704閉卷,201補正,1705取消催審期限
      'Modify by Amy 2025/01/22 +「通知出具同意書 1719」,「出具同意書 723 」-林承慧
      strTemp1 = "Update R100102_1 Set R021004='2' Where (InStr(R021014,'T')>0 or R021014='CFC' or R021014='S') " & _
                             "And (R021018='201' or R021018='202' or R021018='303' or R021018='305' or R021018='703' or R021018='704' or R021018='723' or R021018='1202' or R021018='1705' or R021018='1719' )"
      cnnConnection.Execute strTemp1, intR
      
      '第3句
      '所有專利案件性質404(延期) 者狀態改為 其他相關人
      'modify by sonia 2022/3/24 +411催審,907不續辦,913閉卷
      strTemp1 = "Update R100102_1 Set R021004='2' Where (InStr(R021014,'P')>0 or R021014='FG') And (R021018='404' or R021018='411' or R021018='907' or R021018='913') "
      cnnConnection.Execute strTemp1, intR
      
      '第4句
      'Add by Amy 2023/10/19 +L-888888(為客戶欠款發律師函用,故對造欄一定有值)之對造者狀態改為 其他相關人
      strTemp1 = "Update R100102_1 Set R021004='2' Where R021004='1'  And R021014||R021015='L888888' "
      cnnConnection.Execute strTemp1, intR
      
      '[相關總收文號]-第1句
      'FCT/T 相關總收文號案件性質為 審查報告1201/其他來函1706/核駁1002 者狀態改為 其他相關人
      'modify by sonia 2023/12/27 加入大陸案之1205部分核駁 T-225197(胡緯民)
      strTemp1 = "Update R100102_1 Set R021004='2' Where R021006 in (select R021006 from R100102_1,caseprogress c1,caseprogress c2 where R021014 in('FCT','T') And R021006=c1.cp09(+) and c1.cp43=c2.cp09(+) and c2.cp10 in('1002','1201','1205','1706'))"
      cnnConnection.Execute strTemp1, intR
      'end 2022/09/29
      
      '[相關總收文號]-第2句
      'add by sonia 2022/3/24
      '將所有商標案InStr(R021014,'T')且相關總收文號案件性質為 核駁前先行通知1202 者狀態改為 其他相關人
      strTemp1 = "Update R100102_1 Set R021004='2' Where R021006 in (select R021006 from R100102_1,caseprogress c1,caseprogress c2 where InStr(R021014,'T')>0 And R021006=c1.cp09(+) and c1.cp43=c2.cp09(+) and '1202'=c2.cp10)"
      cnnConnection.Execute strTemp1, intR
      'end 2022/3/24
      
      '[特殊]-第1句
      'add by sonia 2023/12/27 大陸案部分核駁1205後之復審401的勝訴1003或敗訴1004或部分勝部分敗1006，狀態也要改為 其他相關人，T-225197(胡緯民)
      strTemp1 = "Update R100102_1 Set R021004='2' Where R021006 in (select R021006 from R100102_1,caseprogress c1,caseprogress c2,caseprogress c3 where R021014='T' And R021006=c1.cp09(+) and c1.cp43=c2.cp09(+) and c2.cp10='401' and c2.cp43=c3.cp09(+) and c3.cp10='1205')"
      cnnConnection.Execute strTemp1, intR
      'end 2023/12/27
      Exit Function
'*** !!! [申請人查詢更新暫存檔] 上方有修改,下方程式也要修改 !!! ***
    '案號
    ElseIf intChoose = 0 Then
        stSysKind = stNo
    '傳入CP09總收文號
    ElseIf intChoose = 1 Then
        'Modify by Amy 2025/01/22 +cp02
         strQ = "Select o.cp01,o.cp02,o.cp10,b.cp10 as mCP10 From CaseProgress o,CaseProgress b " & _
                     "Where o.cp09='" & stNo & "' And o.cp43=b.cp09(+) "
         intQ = 1
         Set RsQ = ClsLawReadRstMsg(intQ, strQ)
         If intQ = 1 Then
            stSysKind = "" & RsQ.Fields("cp01")
            stCP10 = "" & RsQ.Fields("cp10")
            stCaseNo1 = "" & RsQ.Fields("cp02") 'Modify by Amy 2025/01/22 +cp02
                
            '[相關總收文號]-第1句
            'Add by Amy 2022/09/26 +FCT/T 相關總收文號有 核駁1002/審查報告1201/其他來函1706
            'Modify by Amy 2025/01/22 補 sonia 2023/12/27 加入大陸案之1205部分核駁 T-225197(胡緯民)
            If (stSysKind = "FCT" Or stSysKind = "T") And ("" & RsQ.Fields("mCP10") = "1002" Or "" & RsQ.Fields("mCP10") = "1201" Or "" & RsQ.Fields("mCP10") = "1205" Or "" & RsQ.Fields("mCP10") = "1706") Then
               Pub_ChkRelevantPeople = True
            '[相關總收文號]-第2句
            '商標案 相關總收文號有 核駁前先行通知 1202 者
            ElseIf InStr(stSysKind, "T") > 0 And "" & RsQ.Fields("mCP10") = "1202" Then
               Pub_ChkRelevantPeople = True
            End If
         End If
         Set RsQ = Nothing
         If Pub_ChkRelevantPeople = True Then Exit Function
    End If
    
    '第1句
    'Add by Amy 2022/09/26 +FCT/T 核駁1002/審查報告1201/其他來函1706
    'Modify by Amy 2025/01/22 補sonia 2023/12/27 加入大陸案之1205部分核駁 T-225197(胡緯民) 未加到
    If (stSysKind = "FCT" Or stSysKind = "T") And (stCP10 = "1002" Or stCP10 = "1201" Or stCP10 = "1205" Or stCP10 = "1706") Then
        Pub_ChkRelevantPeople = True
        Exit Function
    '第2句
    ElseIf InStr(stSysKind, "T") > 0 Or stSysKind = "CFC" Or stSysKind = "S" Then
        'Modify by Amy 2025/01/22 +「通知出具同意書 1719」-林承慧
        '核駁前先行通知          取消催審期限                 通知出具同意書
        If stCP10 = "1202" Or stCP10 = "1705" Or stCP10 = "1719" Then
            Pub_ChkRelevantPeople = True
            Exit Function
        End If
        'Modify by Amy 2025/01/23 +「出具同意書 723 」-林承慧
        '補正                                 申請意見書                      延期                       催審                      不續辦                          閉卷                         出具同意書
        If stCP10 = "201" Or stCP10 = "202" Or stCP10 = "303" Or stCP10 = "305" Or stCP10 = "703" Or stCP10 = "704" Or stCP10 = "723" Then
            Pub_ChkRelevantPeople = True
            Exit Function
        End If
    '第3句
    ElseIf InStr(stSysKind, "P") > 0 Or stSysKind = "FG" Then
        '                   延期                       催審                   不續辦                        閉卷
        If stCP10 = "404" Or stCP10 = "411" Or stCP10 = "907" Or stCP10 = "913" Then
            Pub_ChkRelevantPeople = True
            Exit Function
        End If
    '第4句
    'Add by Amy 補Amy 2023/10/19 +L-888888(為客戶欠款發律師函用,故對造欄一定有值) 其他相關人 未加到
    ElseIf stSysKind = "L" And stCaseNo1 = "888888" Then
         Pub_ChkRelevantPeople = True
         Exit Function
    End If
    
    '[特殊]-第1句 : intChoose = 1 or  intChoose = 2 要判斷
   'Add by Amy 2025/01/22 補 sonia 2023/12/27 T大陸案部分核駁1205後之復審401的勝訴1003或敗訴1004或部分勝部分敗1006,狀態也要改為 其他相關人，T-225197(胡緯民)
   If stSysKind = "T" Then
      strQ = "Select c1.cp09,c1.cp01,c1.cp02,c1.cp03,c1.cp04 From CaseProgress c1,CaseProgress c2,CaseProgress c3 " & _
               "Where c1.cp43=c2.cp09(+) And c2.cp10='401' And c2.cp43=c3.cp09(+) And c3.cp10='1205'" & _
               "And c1.cp01='" & stSysKind & "' And c1.cp02='" & stCaseNo1 & "' And c1.cp10='" & stCP10 & "' "
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, strQ)
      If intQ = 1 Then
         Pub_ChkRelevantPeople = True
      End If
   End If
   Set RsQ = Nothing
End Function

'Add by Amy 2022/09/05 從basQuery搬過來
'Memo by Amy 2023/10/26 業績點數語法(目前使用:frmacc41g0/41h0/43c0/44j0/frm210107/113/117/150/152)
'intChoose:0-全部 / 1-只算實績與結餘 / 2-抓人員 / 3-只抓人員目標 / 4-只抓實績 轉撥/結餘 轉撥 (傳票)
'                 1.1-只抓期初實績保留/1.2-期初結餘保留/1.3-當月實績/1.4-當月結餘/1.5-期末實績保留/1.6-期末結餘保留 'Add by Amy 2020/10/19 原intChoose 為整數,改Double
'stDate1-點數結算「起始」民國年月/stDate2-點數結算「截止」民國年月
'stSalesArea1-起始業務區 (*NS=抓非S部門) / stSalesArea2-截止業務區
'stEmpNo:員工編號(F41XX=投資法務-部門要下"*NS")
'bolPoint:是否除1000
'stFormN:表單名稱
'bolFNo:False-不包含F員編 'Add by Amy 2016/03/11
'Add by Amy 2016/03/14
'stCmp:公司別 1/J/空白(全部)
'stSysKind:系統別(M0100用-條件 stSalesArea1="TOT"/stEmpNo="M0100"/stSysKind=系統別)
Public Function GetPoint(intChoose As Double, stDate1 As String, stDate2 As String, Optional ByVal stSalesArea1 As String = "", Optional ByVal stSalesArea2 As String = "", _
                                        Optional ByVal stEmpNo As String = "", Optional ByVal bolPoint As Boolean = True, Optional ByVal stFormN As String = "", Optional ByVal bolFNo As Boolean = False, _
                                        Optional ByVal stCmp As String = "", Optional ByVal stSysKind As String = "") As String
    Dim stCon As String, stConST As String, stConST1 As String, stConR1 As String, stConR2 As String, stConPE As String '傳票/SalesPoint/員工///目標
    Dim stVTB0 As String, stVTB1 As String, stVTB2 As String, stVTB3 As String, stVTB4 As String
    Dim stVTB5 As String, stVTB6 As String, stVTB7 As String, stVTB8 As String, stVTB9 As String
    Dim stT01 As String, stT02 As String '轉撥用
    Dim stConCu As String, stConCP As String, stConM As String, stSQL As String, stWhere As String, stSPTB As String, stSPField As String
    Dim intDivisor As Integer
    Dim bolReport As Boolean, strRep As String 'Add by Amy 2018/04/19 工作報告/重覆之Where 語法
    Dim stVTB9_L As String, stConCP_L As String 'Add by Sindy 2020/10/29
    Dim strF As String, strGroup As String 'Add by Amy 2021/04/27
    Dim bolM0101 As Boolean, stM0101W As String, stM0101R, stTP As String 'Add by Amy 2023/03/24
    Dim ii As Integer, arrCon 'Add by Amy 2023/10/26
    
    intDivisor = 1000
    If bolPoint = False Then intDivisor = 1
    '*** Memo by Amy 2023/03/24 若有增加表單請於下方列示,方便知道修改影響之程式 ***
    Select Case UCase(stFormN)
        Case "FRMACC43C0"
    '*** frmacc43C0 智權點數關閉/開放 ***
            bolM0101 = True
    '*** End frmacc43C0 智權點數關閉/開放 ***
        Case "FRMACC44J0"
    '*** frmacc44J0 智權點數實績與結餘分析表 ***
           stSPTB = ",SalesPoint ": stSPField = ",SP48 "
            '業務區
            If stSalesArea1 <> MsgText(601) Then
                If stSalesArea1 = "*NS" Then
                    bolM0101 = True
                    stConST = stConST & " AND Decode(SP48,null,SubStr(ST15,1,1),SubStr(SP48,1,1))<>'S' And SP01(+)=" & IIf(Val(Mid(stDate1, 1, 5)) <= 10412, 10412, Val(Mid(stDate1, 1, 5))) + 191100 & _
                                    " AND ST01=SP02(+) "
                ElseIf Left(stSalesArea1, 1) = "S" Then
                    stConST = stConST & " AND Decode(SP48,null,SubStr(ST15,1,1),SubStr(SP48,1,1))='S' And SP01(+)=" & IIf(Val(Mid(stDate1, 1, 5)) <= 10412, 10412, Val(Mid(stDate1, 1, 5))) + 191100 & _
                                    " AND ST01=SP02(+)  "
                Else
                    stConST = stConST & " AND Decode(SP48,null,ST15,SP48)='" & stSalesArea1 & "' And SP01(+)=" & IIf(Val(Mid(stDate1, 1, 5)) <= 10412, 10412, Val(Mid(stDate1, 1, 5))) + 191100 & _
                                    " AND ST01=SP02(+)  "
                End If
            End If
            '智權人員
            If stEmpNo <> "" Then
                stCon = stCon & " AND ax209='" & stEmpNo & "' "
                stConST = stConST & " AND SP02(+)='" & stEmpNo & "' "
            'Mark by Amy 2016/05/06
    '        ElseIf stEmpNo = "" And stSalesArea1 = "*NS" Then
    '            stCon = stCon & " AND ax209 Not in ('F4101', 'F4102', 'F4103','M0100') "
            End If
            '點數結算日
            If stDate1 <> "" Then
                stCon = stCon & " AND A0205 >= " & Val(stDate1)
                '上月保留及上月結餘 計算stDate1當月
                stConR1 = "  AND A0205 >= " & Val(stDate1) & " AND A0205 <= " & Val(Mid(stDate1, 1, 5)) & "31"
                'Add by Amy 2023/03/24 stACSDate1 小於當月
                stM0101R = "  AND A0205 >= " & Val(stDate1) & " AND A0205 <= " & Val(Mid(stDate1, 1, 5)) & "31" '記錄要取代資料
                stM0101W = " AND A0205 <" & Val(Mid(stDate1, 1, 5)) & "01"
            End If
            If stDate2 <> "" Then
                stCon = stCon & " AND A0205 <= " & Val(stDate2)
                '保留餘額及結餘餘額 計算stDate2當月
                stConR2 = "  AND A0205 >= " & Val(Mid(stDate2, 1, 5)) & "01 AND A0205 <= " & Val(stDate2)
            End If
            'Add by Amy 2020/04/01 +公司別
            If stCmp <> MsgText(601) Then
                'Modify by Amy 2020/04/16 +組合公司
                If InStr(stCmp, "+") > 0 Then
                    stCon = stCon & " And A0201 In ('" & Replace(stCmp, "+", "','") & "') "
                    stConR1 = stConR1 & " And a0201 In ('" & Replace(stCmp, "+", "','") & "') "
                    stConR2 = stConR2 & " And a0201 In ('" & Replace(stCmp, "+", "','") & "') "
                Else
                    stCon = stCon & " And A0201='" & stCmp & "' "
                    stConR1 = stConR1 & " And a0201='" & stCmp & "' "
                    stConR2 = stConR2 & " And a0201='" & stCmp & "' "
                End If
            End If
            'Modfiy by Amy 2016/05/06 M0100用
            'Modfiy by Amy 2019/01/07 CFP及CPS歸入P(原只有P及PS),CFT及TF歸入T(原只有T及T字頭)
            If stEmpNo = "M0100" And stSysKind <> MsgText(601) Then
                If stSysKind = "P" Then
                    stConM = " And (((SubStr(ax214, 1, length(ax214) - 9)||'-'||SubStr(ax214,  length(ax214) - 8, length(ax214)) like 'P-%' Or SubStr(ax214, 1, length(ax214) - 9)||'-'||SubStr(ax214,  length(ax214) - 8, length(ax214)) like 'CFP-%') " & _
                    "And Exists(Select * From Patent,Fagent,Customer " & _
                    "Where pa01=SubStr(ax214, 1, length(ax214) - 9) And pa02=SubStr(ax214, length(ax214)- 8, 6) And pa03=SubStr(ax214, length(ax214)- 2,1) And pa04=SubStr(ax214, length(ax214)- 1,length(ax214)) " & _
                    "And fa01(+)=SubStr(pa75,1,8) And fa02(+)=SubStr(pa75,9) And cu01(+)=SubStr(pa26,1,8) And cu02(+)=SubStr(pa26,9) And nvl(fa10,cu10)>'009') " & _
                    " Or ((SubStr(ax214, 1, length(ax214) - 9)||'-'||SubStr(ax214,  length(ax214) - 8, length(ax214)) like 'PS-%' Or SubStr(ax214, 1, length(ax214) - 9)||'-'||SubStr(ax214,  length(ax214) - 8, length(ax214)) like 'CPS-%' ) " & _
                    "And Exists(Select * From ServicePractice,Fagent,Customer " & _
                    "Where sp01=SubStr(ax214, 1, length(ax214) - 9) And sp02=SubStr(ax214, length(ax214)- 8, 6) And sp03=SubStr(ax214, length(ax214)- 2,1) And sp04=SubStr(ax214, length(ax214)- 1,length(ax214)) " & _
                    "And fa01(+)=SubStr(sp26,1,8) And fa02(+)=SubStr(sp26,9) And cu01(+)=SubStr(sp08,1,8) and cu02(+)=SubStr(sp08,9) and nvl(fa10,cu10)>'009') ) ))"
                ElseIf stSysKind = "T" Then
                    stConM = " And (((SubStr(ax214, 1, length(ax214) - 9)||'-'||SubStr(ax214,  length(ax214) - 8, length(ax214)) like 'T-%' Or SubStr(ax214, 1, length(ax214) - 9)||'-'||SubStr(ax214,  length(ax214) - 8, length(ax214)) like 'CFT-%' ) " & _
                        "And Exists (Select * From Trademark,Fagent,Customer " & _
                        "Where tm01=SubStr(ax214, 1, length(ax214) - 9) And tm02=SubStr(ax214, length(ax214)- 8, 6) And tm03=SubStr(ax214, length(ax214)- 2,1) And tm04=SubStr(ax214, length(ax214)- 1,length(ax214)) " & _
                        "And fa01(+)=SubStr(tm44,1,8) And fa02(+)=SubStr(tm44,9) And cu01(+)=SubStr(tm23,1,8) And cu02(+)=SubStr(tm23,9) And nvl(fa10,cu10)>'009') " & _
                        "Or (SubStr(ax214, 1, length(ax214) - 9)||'-'||SubStr(ax214,  length(ax214) - 8, length(ax214)) like 'T%' And Exists(Select * From ServicePractice,Fagent,Customer " & _
                        "Where sp01=SubStr(ax214, 1, length(ax214)-9) And sp02=SubStr(ax214, length(ax214)- 8, 6) And sp03=SubStr(ax214, length(ax214)- 2,1) And sp04=SubStr(ax214, length(ax214)- 1,length(ax214)) " & _
                        "And fa01(+)=SubStr(sp26,1,8) And fa02(+)=SubStr(sp26,9) And cu01(+)=SubStr(sp08,1,8) And cu02(+)=SubStr(sp08,9) And nvl(fa10,cu10)>'009') ) )) "
                End If
            End If
            'end 2019/01/07
    '*** End frmacc44J0 智權點數實績與結餘分析表 ***
        Case "FRM210104"
    '*** frm210104 業績點數查詢 ***
    
    '*** End frm210104 業績點數查詢 ***
        Case "FRM210107"
    '*** FRM210107 業績達成日報表 ***
            '點數結算日
            If stDate1 <> MsgText(61) Then
               stCon = stCon & " AND SubStr(A0205+19110000,1,6)=" & Mid(Val(stDate1) + 19110000, 1, 6)
               stConPE = stConPE & " AND PE03 = " & Mid(Val(stDate1) + 19110000, 1, 6)
            End If
            '業務區
            If stSalesArea1 <> MsgText(601) And Left(stSalesArea1, 1) = "S" Then
               stSPTB = ",Acc090"
               '員編要為在職且員編為Sxx為目標設定,避免被抓出故需判斷st01>'6' and st01<'F';[中區其他]固定顯示,故排除;
               stConST = stConST & " And Substr(st15,1,1)='S' And st04='1' And st01>'6' And st15<>'S29' And st15=a0901(+) "
               stConST1 = stConST1 & " And Substr(st15,1,1)='S' And st04='1' And st01>'6' And st15<>'S29' And st15=a0901(+) "
            End If
            '人員-抓[有]目標 或 [有]點數
            If stEmpNo <> MsgText(601) Then
               If InStr(stEmpNo, ";") > 0 Then
                  stCon = stCon & " AND ax209 in('" & Replace(stEmpNo, ";", "','") & "') "
                  stConPE = stConPE & " AND PE01 in('" & Replace(stEmpNo, ";", "','") & "') "
               Else
                  stCon = stCon & " AND ax209='" & stEmpNo & "' "
                  stConPE = stConPE & " AND PE01='" & stEmpNo & "'"
               End If
            End If
            stCon = stCon & " AND ax207-ax206<>0 " '只抓有點數
            stConPE = stConPE & " AND Nvl(PE04,0)>0 " '只抓PE04>0
    '*** End FRM210107 業績達成日報表 ***
        Case "FRM210113"
    '*** frn210113 各區業務工作報告統計 ***
            bolReport = True
    '*** End frm210113 各區業務工作報告統計 ***
        Case "FRM210117"
    '*** frm210117 各所業務工作報告統計 ***
            bolReport = True
            
    '*** End frm210117 各所業務工作報告統計 ***
        Case "FRM210150" 'Add by Amy 2017/07/18 frm210150只抓結餘點數,排除F部門
    '*** frm210150 智權工作報告-總所 ***
            'Modify by Amy 2020/10/19 +intChoose
            If intChoose = 1.41 Then
                stSPTB = ",SalesPoint "
                'Modify by Amy 2021/05/28 11004月,因其他包含L的資料,故將其他不顯示(只顯示智權部)-簡協理
                stConST = stConST & " AND SubStr(Decode(SP48,null,ST15,SP48),1,1)='S' And SP01(+)=" & IIf(Val(Mid(stDate1, 1, 5)) <= 10412, 10412, Val(Mid(stDate1, 1, 5))) + 191100 & _
                                " AND ST01=SP02(+)  "
                '點數結算日
                If stDate1 <> "" Then
                    stCon = stCon & " AND A0205 >= " & stDate1 & "01"
                End If
        
                If stDate2 <> "" Then
                    stCon = stCon & " AND A0205 <= " & stDate2 & "31"
                End If
            End If
            'end 2017/07/18
    '*** End frm210150 智權工作報告-總所 ***
        Case "FRM210152"
    '*** frm210152 每月(智權)點數輸入 ***
            bolM0101 = True
    
    '*** End frm210152 每月(智權)點數輸入 ***
        Case UCase("Pub_GetAccRecePayAmt") 'Add by Amy 2021/06/03
    '*** Pub_GetAccRecePayAmt 函數 ***
            stSPTB = ",SalesPoint "
            '點數結算日
            If stDate1 <> "" Then
                stCon = stCon & " AND A0205 >= " & stDate1
            End If
    
            If stDate2 <> "" Then
                stCon = stCon & " AND A0205 <= " & stDate2
            End If
            stConST = stConST & " AND Decode(SP48,null,ST15,SP48)>='" & stSalesArea1 & "' AND ST01=SP02(+) " & _
                            " And SP01(+)=" & IIf(Val(Mid(stDate1, 1, 5)) <= 10412, 10412, Val(Mid(stDate1, 1, 5))) + 191100 & _
                            " And Decode(SP48,null,ST15,SP48)<='" & stSalesArea2 & "' "
            
    '*** End Pub_GetAccRecePayAmt 函數 ***
        Case UCase("HasActualP")
    '*** HasActualP 函數 ***
    '*** End HasActualP 函數 ***
    End Select
    'end 2023/03/24
    
    If bolFNo = False Then stConST = stConST & " And ST01<'F' "
    'Modify by Amy 2023/10/26 +frm210107
    If Not (UCase(stFormN) = "FRMACC44J0" Or (UCase(stFormN) = "FRM210150" And intChoose = 1.41) Or UCase(stFormN) = "FRM210107" _
      Or UCase(stFormN) = UCase("Pub_GetAccRecePayAmt")) Then
        '業務區
        If stSalesArea1 <> MsgText(601) Then
            'Modify by Amy 2017/10/17 中一高國碩/陳頌恩轉中二,除frmacc43c0抓st15部門,其餘報表抓sp48-秀玲
            If UCase(stFormN) = "FRMACC43C0" Then
               stConST = stConST & " AND ST15>='" & stSalesArea1 & "'"
               stConST1 = stConST1 & " AND ST15>='" & stSalesArea1 & "'" '目標
            'Add by Amy 2018/04/19 改抓暫存檔
            ElseIf UCase(stFormN) = "FRM210113" Then
                stSPTB = ""
                strRep = " And R02='" & stSalesArea1 & "' "
                stConST1 = " And R02='" & stSalesArea1 & "' "
            ElseIf UCase(stFormN) = "FRM210117" Then
                stSPTB = ""
                strRep = " And SubStr(R02,1,2)='" & stSalesArea1 & "' "
                stConST1 = " And SubStr(R02,1,2)='" & stSalesArea1 & "' "
            'end 2018/04/19
            Else
                stSPTB = ",SalesPoint "
                stConST = stConST & " AND Decode(SP48,null,ST15,SP48)>='" & stSalesArea1 & "' AND ST01=SP02(+) " & _
                        "And SP01(+)=" & IIf(Val(Mid(stDate1, 1, 5)) <= 10412, 10412, Val(Mid(stDate1, 1, 5))) + 191100
                stConST1 = stConST1 & " AND Decode(SP48,null,ST15,SP48)>='" & stSalesArea1 & "' AND ST01=SP02(+) " & _
                        "And SP01(+)=" & IIf(Val(Mid(stDate1, 1, 5)) <= 10412, 10412, Val(Mid(stDate1, 1, 5))) + 191100 '目標
            End If
            'end 2017/10/17
            'Mark by Amy 因中一高國碩/陳頌恩轉中二,CU12也改為新部門,造成抓舊資料會有問題,改以人員串
            'stConCu = stConCu & " AND CU12>='" & stSalesArea1 & "'"
            'Add by Amy 2018/04/19  改抓暫存檔
            If UCase(stFormN) = "FRM210113" Then
                stConCP = stConCP & " AND CP12='" & stSalesArea1 & "'"
                stConCP_L = stConCP_L & " AND st15='" & stSalesArea1 & "'" 'Add by Sindy 2020/10/29
            ElseIf UCase(stFormN) = "FRM210117" Then
                stConCP = stConCP & " AND SubStr(CP12,1,2)='" & stSalesArea1 & "'"
                stConCP_L = stConCP_L & " AND SubStr(st15,1,2)='" & stSalesArea1 & "'" 'Add by Sindy 2020/10/29
            Else
                stConCP = stConCP & " AND CP12>='" & stSalesArea1 & "'"
                stConCP_L = stConCP_L & " AND st15>='" & stSalesArea1 & "'" 'Add by Sindy 2020/10/29
            End If
        End If
        
        If stSalesArea2 <> MsgText(601) Then
            'Modify by Amy 2017/10/17 中一高國碩/陳頌恩轉中二,除frmacc43c0抓st15部門,其餘報表抓sp48-秀玲
            If UCase(stFormN) = "FRMACC43C0" Then
                stConST = stConST & " AND ST15<='" & stSalesArea2 & "'"
                stConST1 = stConST1 & " AND ST15<='" & stSalesArea2 & "'" '目標
            Else
                stConST = stConST & " AND Decode(SP48,null,ST15,SP48)<='" & stSalesArea2 & "'"
                stConST1 = stConST1 & " AND Decode(SP48,null,ST15,SP48)<='" & stSalesArea2 & "'" '目標
            End If
            'end 2017/10/17
            'Mark by Amy 因中一高國碩/陳頌恩轉中二,CU12也改為新部門,造成抓舊資料會有問題,改以人員串
            'stConCu = stConCu & " AND CU12<='" & stSalesArea2 & "'"
            stConCP = stConCP & " AND CP12<='" & stSalesArea2 & "'"
            stConCP_L = stConCP_L & " AND st15<='" & stSalesArea2 & "'" 'Add by Sindy 2020/10/29
        End If
        
        '智權人員
        If stEmpNo <> MsgText(601) Then
            stCon = stCon & " AND ax209='" & stEmpNo & "' "
            stConST = " AND ST01='" & stEmpNo & "' "
            stConST1 = " AND ST01='" & stEmpNo & "' "  '目標
            stConR1 = " AND ax209='" & stEmpNo & "' "
            stConR2 = " AND ax209='" & stEmpNo & "' "
        End If
    
        '點數結算日
        If stDate1 <> MsgText(61) Then
            'Modify by Amy 2022/05/17 +if 點數輸入作業及查詢日期需抓畫面上起迄
            'ex:下智權 87052 或 83004 點選「點數統計」日期1110401-0430 會有實績點數 與1110101-0430 實績點數沒累計,因只抓1月
            If UCase(stFormN) = UCase("frm210104") Then
                stCon = stCon & " AND A0205 >= " & stDate1
                stConR1 = "  AND A0205 >= " & stDate1 & " AND A0205 <= " & stDate2
                stConPE = stConPE & " AND PE03(+) >= " & Mid(Val(stDate1) + 19110000, 1, 6)
                stConCu = stConCu & " AND CU14>=" & TransDate(stDate1, 2)
                stConCP = stConCP & " AND CP05>=" & TransDate(stDate1, 2)
                stConCP_L = stConCP_L & " AND CP05>=" & TransDate(stDate1, 2)
            Else
                stCon = stCon & " AND A0205 >= " & stDate1 & "01"
                '上月保留及上月結餘 計算stDate1當月
                stConR1 = "  AND A0205 >= " & stDate1 & "01 AND A0205 <= " & stDate1 & "31"
                stConPE = stConPE & " AND PE03(+) >= 191100+" & stDate1
                stConCu = stConCu & " AND CU14>=" & TransDate(stDate1 & "01", 2)
                stConCP = stConCP & " AND CP05>=" & TransDate(stDate1 & "01", 2)
                stConCP_L = stConCP_L & " AND CP05>=" & TransDate(stDate1 & "01", 2) 'Add by Sindy 2020/10/29
                'Add by Amy 2023/03/28 stACSDate1 小於當月
                stM0101R = "  AND A0205 >= " & stDate1 & "01 AND A0205 <= " & stDate1 & "31" '記錄要取代資料
                stM0101W = " AND A0205 <" & Val(Mid(stDate1, 1, 5)) & "01"
            End If
        End If
        
        If stDate2 <> MsgText(61) Then
            'Modify by Amy 2022/05/17 +if 點數輸入作業及查詢日期需抓畫面上起迄
            If UCase(stFormN) = UCase("frm210104") Then
                stCon = stCon & " AND A0205 <= " & stDate2
                stConR2 = "  AND A0205 >= " & stDate1 & " AND A0205 <= " & stDate2
                stConPE = stConPE & " AND PE03(+) <= " & Mid(Val(stDate2) + 19110000, 1, 6)
                stConCu = stConCu & " AND CU14<=" & TransDate(stDate2, 2)
                stConCP = stConCP & " AND CP05<=" & TransDate(stDate2, 2)
                stConCP_L = stConCP_L & " AND CP05<=" & TransDate(stDate2, 2)
            Else
                stCon = stCon & " AND A0205 <= " & stDate2 & "31"
                '保留餘額及結餘餘額 計算stDate2當月
                stConR2 = "  AND A0205 >= " & stDate2 & "01 AND A0205 <= " & stDate2 & "31"
                stConPE = stConPE & " AND PE03(+) <= 191100+" & stDate2
                stConCu = stConCu & " AND CU14<=" & TransDate(stDate2 & "31", 2)
                stConCP = stConCP & " AND CP05<=" & TransDate(stDate2 & "31", 2)
                stConCP_L = stConCP_L & " AND CP05<=" & TransDate(stDate2 & "31", 2) 'Add by Sindy 2020/10/29
            End If
        End If
    End If
    
   '*** 只抓人員目標用 ***
    'Modify by Amy 2024/02/22 開放給frm210107-業績達成日報表用,原PE01(+)=ST01改為以[有目標]才出現
    If intChoose = 3 Then
      If UCase(stFormN) = "FRM210107" Then
         '智權部
         If stSalesArea1 <> MsgText(601) And Left(stSalesArea1, 1) = "S" Then
            strF = "'" & strUserNum & "',StateNo,a0902 as DepN,ST15 as SP48,st15||'ZZ' as ax209 "
         '[非]智權部
         Else
            strF = "'" & strUserNum & "',StateNo,st02 as DepN,st15 as SP48,st01 "
         End If
         GetPoint = "Select Distinct " & strF & " From Staff,PerFormance " & stSPTB & _
                     " Where PE01=ST01(+) And PE02='TOT'" & stConST1 & stConPE & strGroup
         Exit Function
'    'Mark by Amy 2018/04/19 改存暫存檔因中一高國碩/陳頌恩轉中二 下S21 10609~10610目標值不正確(frm210113/117用)
'    'Modify by Amy 2017/10/17 中一高國碩/陳頌恩轉中二,除frmacc43c0抓st15部門,其餘報表抓sp48-秀玲
'      Else
'         GetPoint = "Select Nvl(Sum(PE04),0) PE04,st15 as SP48,st01 From Staff,PerFormance " & stSPTB & _
'                     " Where PE01=ST01(+) And PE02='TOT'" & stConST1 & stConPE & _
'                     " Group by st15,st01 "
'         Exit Function
      End If
    End If
    '*** End 只抓人員目標用 ***
    
    If intChoose < 3 Then
        If UCase(stFormN) = "FRMACC44J0" Then
            '不抓目標因某此人會抓不到資料ex:D105022706 林銘洲結餘轉撥中三(20031-此員編沒目標抓不到)其他
            'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:(substr(ax205, 1, 2) = '41'
            'Modify by Amy 2020/04/01 +7129
            stVTB0 = "Select Distinct ST01||ST02 as StName,ax209 as ST01,Decode(SP48,null,ST15,SP48) as SP48 From Acc021, Acc020,Staff,SalesPoint" & _
                " Where ax201(+) = a0201 And ax202(+) = a0202 And ax209 is not null And (substr(ax205, 1, 1) = '4' Or ax205='7121' Or ax205='7129')" & _
                " And ax209=st01(+) " & stConST & stCon & IIf(stConM <> "", stConM, "")
            'Add by Amy 2023/05/05 M0101有餘額一律顯示
            stVTB0 = stVTB0 & " Union " & GetACSData("9.1", "GetPoint", "", ",acc020,Staff,SalesPoint", Replace(stConR1, stM0101R, stM0101W) & stConST & IIf(stConM <> "", stConM, ""))
        ''Add by Amy 2018/04/19 改抓暫存檔
        ElseIf UCase(stFormN) = "FRM210113" Or UCase(stFormN) = "FRM210117" Then
            '因frm210113 排序需抓Staff
            'Modify by Amy 2024/04/19 FRM210113 +R05
            stVTB0 = "Select R01 as ST01,R02" & IIf(UCase(stFormN) = "FRM210113", ",ST02,ST04,ST05,R05", "") & ",Sum(Nvl(PE04,0)) PE04 From R210113,PerFormance" & IIf(UCase(stFormN) = "FRM210113", ",Staff", "") & _
                " Where ID='" & strUserNum & "' And R04='" & IIf(UCase(stFormN) = "FRM210117", 2, 1) & "' And R01=PE01(+) And R03=PE03(+) And PE02(+)='TOT' " & _
                 stConST1 & stConPE & IIf(UCase(stFormN) = "FRM210113", " And R01=ST01(+) Group by R01,R02,ST02,ST04,ST05,R05 ", " Group by R01,R02 ")
        'Add by Amy 2023/10/26 +frm210107
        ElseIf UCase(stFormN) <> "FRM210107" Then
            '目標
            '2015/10/13 取消ST01>'60'條件,因D104093822做20011/取消ST01<'F'條件,因為2015/10的S212員工編號有掛目標
            'st05 For 最後的排序與目標無關
            'Modify by Amy 2017/10/17 中一高國碩/陳頌恩轉中二,除frmacc43c0抓st15部門,其餘報表抓sp48-秀玲
            stVTB0 = "Select ST01,ST02,ST04,ST05,sum(PE04) PE04" & _
                " From Staff,PerFormance" & stSPTB & _
                " Where PE01(+)=ST01 And PE02(+)='TOT'" & stConST1 & stConPE & _
                " Group by ST01,ST02,ST04,ST05 "
        End If
        
        'Modify by Amy 2018/04/19 改抓暫存檔
        If bolReport = True Then
            strRep = " And ID='" & strUserNum & "' And R04='" & IIf(UCase(stFormN) = "FRM210117", 2, 1) & "' " & _
                    "And R03=A0205(+) " & strRep
        End If
        
        'Modify by Amy 2020/10/19 +if intChoose = 0 Or intChoose = 1 Or intChoose = 1.x
        'Modify by Amy 2020/11/02 抓人員 也要抓傳票資料
        If intChoose = 0 Or intChoose = 1 Or intChoose = 1.1 Or intChoose = 2 Then
            '上月保留:點數結算「起始」當月4191+4192貸方(期初實績保留)
            'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
            'Modify by Amy 2023/03/24 M0101(專案小組)不需轉期初保留,改抓2492,故拆開
            stTP = ""
            If bolM0101 = True Then stTP = " Having ax209<>'M0101' "
            
            stVTB1 = "Select ax209 V10, Sum(ax207) V11" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6) as a0205", "") & _
                " From acc020, acc021,Staff" & stSPTB & _
                " Where ax201(+) = a0201  And ax202(+) = a0202" & stConR1 & _
                " And st01(+)=ax209 " & stConST & IIf(stConM <> "", stConM, "") & _
                " And (ax205= '4191' Or ax205='4192') And InStr(ax212,'轉撥')=0 " & _
                " Group by ax209" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6)", "") & stTP
            
            If stEmpNo = "M0101" Then stVTB1 = "" 'Add by Amy 2023/03/28 有限制員工編號只抓M0101,就不需抓其他資料
            If bolM0101 = True Then
                'Modify by Amy 2023/04/17 ACS期初(M0101)改抓共用Function
                If stVTB1 <> MsgText(601) Then stVTB1 = stVTB1 & " Union "
                stVTB1 = stVTB1 & GetACSData("9", "GetPoint", "", ",acc020,Staff" & stSPTB, Replace(stConR1, stM0101R, stM0101W) & stConST & IIf(stConM <> "", stConM, ""))
                If bolReport = True Then
                    stVTB1 = Replace(stVTB1, ",SubStr(a0205+19110000,1,6) as a0205", "")
                End If
                'end 2023/04/17
            End If
            'end 2023/03/24
            If bolReport = True Then
                stVTB1 = "Select V10, Sum(V11) V11,R02 From R210113,(" & stVTB1 & ") Where R01=V10(+) And V10 is not null" & strRep & " Group by V10,R02 "
            End If
            If intChoose = 1.1 Then GetPoint = stVTB1: Exit Function
        End If
     
        If intChoose = 0 Or intChoose = 1 Or intChoose = 1.2 Or intChoose = 2 Then
            '上月結餘:點數結算「起始」當月4194貸方(期初結餘保留)
            'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
            'Modify by Amy 2016/03/11 因瑞婷10502傳票輸錯莊宏宇之部門,以傳票做調整,故排除沖銷那筆傳票
            stVTB2 = "Select ax209 V20, Sum(ax207) V21" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6) as a0205", "") & _
                " From acc020, acc021,Staff" & stSPTB & _
                " Where ax201(+) = a0201  And ax202(+) = a0202" & stConR1 & IIf(stConM <> "", stConM, "") & _
                " And st01(+)=ax209 " & stConST & _
                " And ax205= '4194' And InStr(ax212,'轉撥')=0 And Not (ax201='1' and ax202='D105022517') Group by ax209" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6)", "")
            If bolReport = True Then
                stVTB2 = "Select V20, Sum(V21) V21,R02 From R210113,(" & stVTB2 & ") Where R01=V20(+) And V20 is not null " & strRep & " Group by V20,R02 "
            End If
            If intChoose = 1.2 Then GetPoint = stVTB2: Exit Function
        End If
        
        If intChoose = 0 Or intChoose = 1 Or intChoose = 1.3 Or intChoose = 2 Then
            'Add by Amy 2021/04/27 +if 傳入表單名為Pub_GetAccRecePayAmt 以項次加總
            If intChoose = 1.3 And UCase(stFormN) = UCase("Pub_GetAccRecePayAmt") Then
                strF = "ax201,ax202,ax203,Sum(ax207-ax206) as ax207,ax209,ax214,a0205"
                strGroup = "Group by ax201,ax202,ax203,ax209,ax214,a0205"
            'Add by Amy 2024/02/22 S部門抓部門加總,非S部門以人員加總
            ElseIf intChoose = 1.3 And UCase(stFormN) = "FRM210107" Then
               '智權部
               If stSalesArea1 <> MsgText(601) And Left(stSalesArea1, 1) = "S" Then
                  strF = "Distinct '" & strUserNum & "',StateNo,a0902 as DepN,ST15 as SP48,st15||'ZZ' as ax209 "
               '[非]智權部
               Else
                  strF = "Distinct '" & strUserNum & "',StateNo,st02 as DepN,st15 as SP48,ax209 "
               End If
            Else
                strF = "ax209 V30, Sum(ax207-ax206) V31" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6) as a0205", "")
                strGroup = "Group by ax209" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6)", "")
            End If
            '當月實績
            'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
            'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:SubStr(ax205, 1, 2) = '41'
            'Modify by Amy 2020/04/01 +7129
            'Memo by Amy 2023/04/10 J公司D112030094/95 ACS M0101傳票,因11003月收款時未做收入傳票,但有做 2492傳票
            '                                           並於11203月才加 420101傳票,導致11203多,會與ACS餘額不符,目前改都不符,等辜財務11203月報自行改後再看
            stVTB3 = "Select " & strF & _
                " From acc020, acc021,Staff" & stSPTB & _
                " Where ax201(+) = a0201  And ax202(+) = a0202" & stCon & IIf(stConM <> "", stConM, "") & _
                " And ST01(+)=ax209 " & stConST & _
                " And (SubStr(ax205, 1, 1) = '4' Or ((ax205='7121' Or ax205='7129') And ax209 is not null)) And Not( ax205='4191' or ax205='4192' or ax205='4194')" & _
                " And (ax213 Is Null or InStr(ax213||' ','結餘')=0) And InStr(ax212,'轉撥')=0 " & strGroup

            If bolReport = True Then
                stVTB3 = "Select V30, Sum(V31) V31,R02 From R210113,(" & stVTB3 & ") Where R01=V30(+) And V30 is not null " & strRep & " Group by V30,R02 "
            End If
            If intChoose = 1.3 Then GetPoint = stVTB3: Exit Function
        End If
        
        If intChoose = 0 Or intChoose = 1 Or intChoose = 1.4 Or intChoose = 1.41 Or intChoose = 2 Then
             'Add by Amy 2021/04/27 +if 傳入表單名為Pub_GetAccRecePayAmt 以項次加總
            If intChoose = 1.4 And UCase(stFormN) = UCase("Pub_GetAccRecePayAmt") Then
                strF = "ax201,ax202,ax203,Sum(ax207-ax206) as ax207,ax209,ax214,a0205"
                strGroup = "Group by ax201,ax202,ax203,ax209,ax214,a0205"
            Else
                strF = "ax209 V40, Sum(ax207-ax206) V41" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6) as a0205", "")
                strGroup = "Group by ax209" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6)", "")
            End If
            '當月結餘
            'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
            stVTB4 = "Select " & strF & _
                " From acc020, acc021,staff" & stSPTB & _
                " Where ax201(+) = a0201  And ax202(+) = a0202" & stCon & IIf(stConM <> "", stConM, "") & _
                " And ST01(+)=ax209 " & stConST & _
                " And (SubStr(ax205, 1, 1) = '4' Or ((ax205='7121' Or ax205='7129') And ax209 is not null)) And Not( ax205='4191' or ax205='4192' or ax205='4194')" & _
                " And InStr(ax213||' ','結餘')>0 And InStr(ax212,'轉撥')=0 " & strGroup
            If bolReport = True Then
                stVTB4 = "Select V40, Sum(V41) V41,R02 From R210113,(" & stVTB4 & ") Where R01=V40(+) And V40 is not null " & strRep & " Group by V40,R02 "
            End If
            If intChoose = 1.4 Then GetPoint = stVTB4: Exit Function
        End If
        'end 2019/08/01
        
        'Add by Amy 2017/07/18 +if frm210150只抓結餘點數
        'Modify  by Amy 2020/10/19 +intChoose/Having
        If UCase(stFormN) = "FRM210150" And intChoose = 1.41 Then
            GetPoint = "Select Sum(V41)/" & intDivisor & " From (" & stVTB4 & ") Having Sum(V41)/" & intDivisor & "<>0 "
            Exit Function
        End If
          
        If intChoose = 0 Or intChoose = 1 Or intChoose = 1.5 Or intChoose = 2 Then
            '保留餘額:點數結算「迄月」當月4191+4192借方(期末實績保留)
            'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
            stVTB5 = "Select ax209 V50, Sum(ax206) V51" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6) as a0205", "") & _
                " From acc020, acc021,Staff" & stSPTB & _
                " Where ax201(+) = a0201  And ax202(+) = a0202" & stConR2 & IIf(stConM <> "", stConM, "") & _
                " And ST01(+)=ax209 " & stConST & _
                " And ( ax205='4191' or ax205='4192') And InStr(ax212,'轉撥')=0 Group by ax209" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6)", "")
            If bolReport = True Then
                stVTB5 = "Select V50, Sum(V51) V51,R02 From R210113,(" & stVTB5 & ") Where R01=V50(+) And V50 is not null " & strRep & " Group by V50,R02 "
            End If
            If intChoose = 1.5 Then GetPoint = stVTB5: Exit Function
        End If
        
        If intChoose = 0 Or intChoose = 1 Or intChoose = 1.6 Or intChoose = 2 Then
            '結餘餘額:點數結算「迄月」當月4194借方(期末結餘保留)
            'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
            'Modify by Amy 2016/03/11 因瑞婷10502傳票輸錯莊宏宇之部門,以傳票做調整,故排除沖銷那筆傳票
            stVTB6 = "Select ax209 V60, Sum(ax206) V61" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6) as a0205", "") & _
                " From acc020, acc021,Staff" & stSPTB & _
                " Where ax201(+) = a0201  And ax202(+) = a0202" & stConR2 & IIf(stConM <> "", stConM, "") & _
                " And ST01(+)=ax209 " & stConST & _
                " And ax205='4194' And InStr(ax212,'轉撥')=0 And Not (ax201='1' and ax202='D105022517') Group by ax209" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6)", "")
            If bolReport = True Then
                stVTB6 = "Select V60, Sum(V61) V61,R02 From R210113,(" & stVTB6 & ") Where R01=V60(+) And V60 is not null " & strRep & " Group by V60,R02 "
            End If
            If intChoose = 1.6 Then GetPoint = stVTB6: Exit Function
        End If
        'end 2020/10/19
        'end 2020/11/02
        
        'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:(substr(ax205, 1, 2) = '41'
        If UCase(stFormN) = "FRMACC44J0" Then
            '實績轉撥
             stT01 = "Select ax209 T10, Sum(ax207-ax206) T11" & _
                " From acc020, acc021,Staff" & stSPTB & _
                " Where ax201(+) = a0201  And ax202(+) = a0202" & stCon & IIf(stConM <> "", stConM, "") & _
                " And ST01(+)=ax209 " & stConST & _
                " And (SubStr(ax205, 1, 1) = '4' Or ((ax205='7121' Or ax205='7129') And ax209 is not null)) And Not( ax205='4191' or ax205='4192' or ax205='4194')" & _
                " And (ax213 Is Null or InStr(ax213||' ','結餘')=0) And InStr(ax212,'轉撥')>0 Group by ax209"
            
            '結餘轉撥
            stT02 = "Select ax209 T20, Sum(ax207-ax206) T21" & _
                " From acc020, acc021,staff" & stSPTB & _
                " Where ax201(+) = a0201  And ax202(+) = a0202" & stCon & IIf(stConM <> "", stConM, "") & _
                " And ST01(+)=ax209 " & stConST & _
                " And (SubStr(ax205, 1, 1) = '4' Or ((ax205='7121' Or ax205='7129') And ax209 is not null)) And Not( ax205='4191' or ax205='4192' or ax205='4194')" & _
                " And InStr(ax213||' ','結餘')>0 And InStr(ax212,'轉撥')>0 Group by ax209"
        End If
        'end 2020/04/01
        'end 2019/08/01
    End If
    
    If UCase(stFormN) = "FRM210113" Or UCase(stFormN) = "FRM210117" Then
        '每月新增客戶數
        'Modify by Amy 2018/04/19  抓部門別,中一高國碩/陳頌恩轉中二若下區間需顯示於目前部門
        stVTB7 = " And ST15='" & stSalesArea1 & "' "
        If UCase(stFormN) = "FRM210117" Then stVTB7 = " And SubStr(ST15,1,2)='" & stSalesArea1 & "' "
        stVTB7 = "Select cu13 V70,count(*) V71 " & IIf(bolReport = True, ",ST15 as R02", "") & _
            " From customer " & IIf(bolReport = True, ",Staff", "") & _
            " Where cu02='0'" & IIf(bolReport = True, " And CU13=ST01(+) ", "") & stConCu & stVTB7 & _
            " group by cu13" & IIf(bolReport = True, ",ST15", "")
            
        '收款家數  'modify by sonia 2015/10/13 取消ST01>'60'條件,D104093822做20011
        'modify by sonia 2016/1/22 +412101,413101
        'Modify by Amy 2017/10/17 中一高國碩/陳頌恩轉中二,除frmacc43c0抓st15部門,其餘報表抓sp48-秀玲
        'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:(substr(ax205, 1, 2) = '41'
        stVTB8 = "Select ax209 V80, Count(Distinct SubStr(ax208,1,6) ) V81" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6) as a0205", "") & _
            " From acc020, acc021,Staff" & stSPTB & _
            " Where ax201(+) = a0201  And ax202(+) = a0202" & stCon & _
            " And ST01(+)=ax209 " & stConST & _
            " And SubStr(ax205, 1, 1) = '4' And ax207>0 And ax208 is not null" & _
            " And not( ax205='4191' or ax205='4192' or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='412101' or ax205='4131' or ax205='413101')" & _
            " And InStr(ax213||' ','結餘')>0)) And InStr(ax212,'轉撥')=0 Group by ax209" & IIf(bolReport = True, ",SubStr(a0205+19110000,1,6)", "")
        If bolReport = True Then
            stVTB8 = "Select V80, Sum(V81) V81,R02 From R210113,(" & stVTB8 & ") Where R01=V80(+) And V80 is not null " & strRep & " Group by V80,R02 "
        End If

        '收文案件來源分析
        'Modify by Amy 2018/04/19  抓部門別
        'modify by sonia 2019/7/31 調整分類並加ACS系統類別,原CFL,CFC列V94改CFL列V93,CFC列V92,ACS列V94
        'Modify By Sindy 2020/10/29 拿掉 ", SUM(DECODE(CP01,'L',1,'LA',1,'CFL',1,0)) V93"
        stVTB9 = "Select CP13 V90" & _
            ", SUM(DECODE(CP01,'P',1,'PS',1,'CFP',1,'CPS',1,0)) V91" & _
            ", SUM(DECODE(CP01,'T',1,'TF',1,'CFT',1,'TC',1,'CFC',1,0)) V92" & _
            ", SUM(DECODE(CP01,'ACS',1,0)) V94" & IIf(bolReport = True, ",CP12 as R02", "") & _
            " From CaseProgress Where cp09<'B' And cp13 is not null" & stConCP & _
            " And (  (cp01 in ('P','PS','CFP','CPS') )  or (cp01 in ('T','TF','CFT','TC') )" & _
            " or (cp01 in ('CFC','CFL','ACS')) or (cp01 in ('L','LA'))) Group by CP13" & IIf(bolReport = True, ",CP12", "")
         'Add By Sindy 2020/10/29 + stVTB9_L
         stVTB9_L = "Select st01 V9L0" & _
            ", SUM(DECODE(CP01,'L',1,'LA',1,'CFL',1,0)) V9L3" & _
            IIf(bolReport = True, ",st15 as R02", "") & _
            " From CaseProgress,lawofficesource,staff Where cp09<'B' And los04 is not null AND instr(los04,st01)>0" & stConCP_L & _
            " And cp162=los15(+) And los15 IS NOT NULL" & _
            " And cp01 in ('L','LA','CFL') Group by st01" & IIf(bolReport = True, ",st15", "")
         '2020/10/29 END
    End If
    
    'Add by Amy 2016/02/05 +依表單名稱傳回sql 語法
    Select Case UCase(stFormN)
        Case "FRMACC43C0"
            'Modify by Amy 2023/04/06 +Distinct M0101可能不止一筆
            stSQL = "Select Distinct " & Val(stDate1) + 191100 & ",ST01,'ZZZ'"
        'Modify by Amy 2018/04/19 改抓暫存檔
        Case "FRM210113"
            'Modify by Amy 2024/04/19 +ST01/R05/ST05 拿掉stWhere
            stSQL = "Select ST01,ST02,R05,ST05"
            'stWhere = " Order by Decode(ST05,'SM',1,'76',1,2),st01"
        Case "FRM210152"
            stSQL = "Select '" & strUserNum & "',ST01,ST05"
        Case "FRMACC44J0"
            Select Case intChoose
                Case 0
                    'Modify by Amy 2016/05/06
                    stSQL = "Select '" & strUserNum & "',SP48,StName,ST01"
                    'If stEmpNo <> "M0100" Then stWhere = " Order by SP48,ST01"
            End Select
            
        Case Else
             stSQL = "Select ST01"
    End Select
    
    If UCase(stFormN) = "FRM210117" Then
    'Add By Sindy 2020/10/29 + ,(" & stVTB9_L & ") V9L
    GetPoint = "Select V0.R02 as ST15,a0902 as NAME,Sum(PE04) as PE04,Sum(NVL(V11,0)/" & intDivisor & ") C1,Sum(NVL(V21,0)/" & intDivisor & ") C2,Sum(NVL(V31,0)/" & intDivisor & ") C3,Sum(NVL(V41,0)/" & intDivisor & ") C4" & _
                            ",Sum(NVL(V51,0)/" & intDivisor & ") C5,Sum(NVL(V61,0)/" & intDivisor & ") C6" & _
                            ",Sum(NVL(V71,0)) C7, Sum(NVL(V81,0)) C8, Sum(NVL(V91,0)) C9, Sum(NVL(V92,0)) C10,Sum(NVL(V9L3,0)) C11,Sum(NVL(V94,0)) C12,Count(VP) C13 " & _
                    "From Acc090,(" & stVTB0 & ") V0,(" & stVTB1 & ") V1,(" & stVTB2 & ") V2,(" & stVTB3 & ") V3,(" & stVTB4 & ") V4,(" & stVTB5 & ") V5,(" & stVTB6 & ") V6,(" & stVTB7 & ") V7" & _
                            ",(" & stVTB8 & ") V8,(" & stVTB9 & ") V9,(" & stVTB9_L & ") V9L,(Select Distinct R01 as VP,R02 From R210113 Where ID='" & strUserNum & "' And R04='2' And SubStr(R01,1,1)<>'S') VP " & _
                    "Where V0.R02=a0901(+) And V0.R02=V1.R02(+) And ST01=V10(+) And V0.R02=V2.R02(+) And ST01=V20(+) And V0.R02=V3.R02(+) And ST01=V30(+) " & _
                                "And V0.R02=V4.R02(+) And ST01=V40(+) And V0.R02=V5.R02(+) And ST01=V50(+) And V0.R02=V6.R02(+) And ST01=V60(+) And V0.R02=V7.R02(+) And ST01=V70(+) " & _
                                "And V0.R02=V8.R02(+) And ST01=V80(+) And V0.R02=V9.R02(+) And ST01=V90(+) And V0.R02=V9L.R02(+) And ST01=V9L0(+) And V0.R02=VP.R02(+) And ST01=VP(+) " & _
                                "And (PE04>0 or V11>0 or V21>0 or V31>0 or V41>0) Group by V0.R02,a0902"
   
    Else
    'Add By Sindy 2020/10/29 + ,(" & stVTB9_L & ") V9L
    GetPoint = stSQL & IIf(UCase(stFormN) <> "FRMACC44J0" And intChoose <> 2, ",Nvl(PE04,0) PE04", "") & _
        IIf(intChoose = 2, "", ",NVL(V11,0)/" & intDivisor & " C1,NVL(V21,0)/" & intDivisor & " C2,NVL(V31,0)/" & intDivisor & " C3,NVL(V41,0)/" & intDivisor & " C4" & _
                    ",NVL(V51,0)/" & intDivisor & " C5,NVL(V61,0)/" & intDivisor & " C6") & _
        IIf(UCase(stFormN) <> "FRMACC44J0", "", ",NVL(T11,0)/" & intDivisor & " T1,NVL(T21,0)/" & intDivisor & " T2") & _
        IIf(intChoose = 0 And UCase(stFormN) <> "FRMACC44J0", ", NVL(V71,0) C7, NVL(V81,0) C8, NVL(V91,0) C9, NVL(V92,0) C10, NVL(V9L3,0) C11, NVL(V94,0) C12", "") & _
        " From (" & stVTB0 & ") V0,(" & stVTB1 & ") V1,(" & stVTB2 & ") V2,(" & stVTB3 & ") V3,(" & stVTB4 & ") V4" & _
                    ",(" & stVTB5 & ") V5,(" & stVTB6 & ") V6" & _
        IIf(UCase(stFormN) <> "FRMACC44J0", "", ",(" & stT01 & ") T1,(" & stT02 & ") T2") & _
        IIf(intChoose = 0 And UCase(stFormN) <> "FRMACC44J0", ",(" & stVTB7 & ") V7,(" & stVTB8 & ") V8,(" & stVTB9 & ") V9,(" & stVTB9_L & ") V9L", "") & _
        " Where V10(+)=ST01 And V20(+)=ST01 And V30(+)=ST01 And V40(+)=ST01" & _
                    " And V50(+)=ST01 And V60(+)=ST01" & _
        IIf(UCase(stFormN) <> "FRMACC44J0", "", " And T10(+)=ST01 And T20(+)=ST01") & _
        IIf(intChoose = 0 And UCase(stFormN) <> "FRMACC44J0", " And V70(+)=ST01 And V80(+)=ST01 And V90(+)=ST01 And V9L0(+)=ST01", "") & _
        IIf(UCase(stFormN) <> "FRMACC44J0", " And (PE04>0 or V11>0 or V21>0 or V31>0 or V41>0)", " And (V11<>0 or V21<>0 or V31<>0 or V41<>0 or T11<>0 or T21<>0)") & _
        stWhere
    End If
    'end 2016/02/05
'end 2016/03/14
End Function

'傳入員編判斷是否為業務輸入區主管職代
'm_A0908:取得代理區主管員編(A0908)
'm_ST04:在職/離職 (需傳入Y)
'm_A0901:業務區 'Add by Amy 2021/11/11
Public Function IsAreaAgent(ByVal m_StaffNo As String, Optional ByVal bolShowName As Boolean = False, Optional ByRef m_A0908 As String = "" _
                                             , Optional ByRef m_ST04 As String = "", Optional ByRef m_A0901 As String = "") As Boolean
    Dim Rs As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    
    m_A0908 = "": IsAreaAgent = False
    m_A0901 = "" 'Add by Amy 2021/11/11
    'Modify by Amy 2016/03/30 +ST04
    'Memo by Amy 2021/11/11 目前無職代代理多個部門,故先不考慮-秀玲
    'Modify by Amy 2021/11/11 +a0918 先抓「每月點數輸入確認主管之員工編號(a0918)」再抓「部門主管(a0908)」
    stQ = "Select Nvl(A0918,A0908) as a0908,a0901" & IIf(bolShowName = True, ",ST02,ST04", "") & " From Acc090" & IIf(bolShowName = True, ",Staff", "") & _
            " Where InStr(A0914,'" & m_StaffNo & "')>0 " & IIf(bolShowName = True, "And Nvl(A0918,A0908)=ST01(+) ", "")
    intQ = 1
    Set Rs = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        IsAreaAgent = True
        m_A0908 = Rs.Fields("A0908") & IIf(bolShowName = True, " " & Rs.Fields("ST02"), "")
        If bolShowName = True And m_ST04 = "Y" Then m_ST04 = "" & Rs.Fields("ST04") 'Add by Amy 2016/03/30
        m_A0901 = "" & Rs.Fields("a0901") 'Add by Amy 2021/11/11 職代代理之區
    End If
    Set Rs = Nothing
End Function
'end 2016/01/14

'Add by Amy 2016/02/03 抓取業績點數資料
'stDate1-點數結算「起始」民國年月/stDate2-點數結算「截止」民國年月
'stSalesArea1-起始業務區 (*NS=抓非S部門)/ stSalesArea2-截止業務區
'bolPoint:是否為點數/'stFormN:表單名稱
'bolFNo:False不抓F編號 'Modify by Amy 2016/03/09
'stCon:傳入條件(目前業績達成日報表用,不同條件使用[;]區隔) 'Add by Amy 2023/10/26
Public Function GetPoint_SP(stDate1 As String, stDate2 As String, Optional ByVal stSalesArea1 As String = "", Optional ByVal stSalesArea2 As String = "", Optional ByVal stEmpNo As String = "", _
   Optional ByVal bolPoint As Boolean = True, Optional ByVal stFormN As String = "", Optional ByVal bolSum As Boolean = False, Optional ByVal bolFNo As Boolean = False, Optional ByVal stCon As String = "") As String
    Dim stSQL As String, stSql_C As String, stWhere1 As String, stWhere2 As String, stWhere3 As String, stOrderBy As String
    Dim intMultiplier As Integer
    Dim bolReport As Boolean 'Add by Amy 2018/04/19 工作報告
    Dim stTmp As String 'Add by Amy 2021/05/20
    Dim ii As Integer, arrCon 'Add by Amy 2023/10/26
    
    intMultiplier = 1
    If bolPoint = False Then intMultiplier = 1000
    'Add by Amy 2018/04/19
    If UCase(stFormN) = "FRM210113" Or UCase(stFormN) = "FRM210117" Then
        bolReport = True
    End If
    
    'Modify by Amy 2023/10/26 +業績達成日報表-智權
    If UCase(stFormN) = "FRM210107" Then
      If stDate1 <> MsgText(601) Then
         stWhere1 = stWhere1 & " And SP01>=" & Mid(stDate1 + 19110000, 1, 6)
      End If
      If stDate2 <> MsgText(601) Then
         stWhere1 = stWhere1 & " And SP01<=" & Mid(stDate2 + 19110000, 1, 6)
      End If
      stWhere2 = Replace(stWhere1, "SP01", "PE03")
      '未使用  stWhere3 = stWhere1
      If stCon <> MsgText(601) Then
         arrCon = Split(stCon, ";")
         For ii = LBound(arrCon) To UBound(arrCon)
            stTmp = UCase(arrCon(ii))
            If InStr(stTmp, " AND SP") > 0 Or InStr(stTmp, " AND SUBSTR(SP") > 0 Then
               stWhere1 = stWhere1 & stTmp
            ElseIf InStr(stTmp, " AND ST") > 0 Or InStr(stTmp, " AND SUBSTR(ST") > 0 Then
               stWhere2 = stWhere2 & stTmp
            End If
         Next ii
      End If
    Else
      '點數結算日
      '期末實績/結餘會結轉所以只抓止月當月資料,轉撥跨月需加總
      If stDate1 <> stDate2 And bolSum = False Then bolSum = True
      If stDate1 <> "" Then
           stWhere1 = stWhere1 & " And SP01>=" & stDate1 + 191100
      End If
      If stDate2 <> "" Then
          stWhere1 = stWhere1 & " And SP01<=" & stDate2 + 191100
          stWhere2 = stWhere2 & " And SP01=" & stDate2 + 191100
      End If
    End If
    
    
    '業務區
    'Modify by Amy 2024/02/26 +業績達成日報表-智權
    'Modify by Amy 2016/07/07 +stWhere3
    If UCase(stFormN) = "FRM210107" Then
      If stSalesArea1 = "*NS" Then
         '抓傳入的部門 及 人員,因不同年月顯示不同資料
      Else
         stWhere1 = stWhere1 & " And SubStr(SP48,1,1)='S'"
      End If
    'Modify by Amy 2018/02/06 實績與結餘分析表不需判斷業務區
    '因frmacc44j0實績與結餘分析表10610中一部分人員調中二,若下10609~10610 限制區別會抓不到期末保留及結餘(期未抓止月)
    ElseIf UCase(stFormN) <> "FRMACC44J0" Then
      If stSalesArea1 <> "" Then
            If stSalesArea1 = "*NS" Then
                stWhere1 = stWhere1 & " And SubStr(SP48,1,1)<>'S'"
                stWhere2 = stWhere2 & " And SubStr(" & IIf(bolSum = True, "O.SP48", "SP48") & " ,1,1)<>'S'"
                'If UCase(stFormN) = "FRMACC44J0" Then stWhere3 = stWhere3 & " And SubStr(R001,1,1)<>'S'"
            'Add by Amy 2018/04/19 若下S21 10609~10610 高國碩、陳頌恩期末資料會抓不到,因調區
            ElseIf UCase(stFormN) = "FRM210113" Then
                stWhere1 = stWhere1 & " And SP48='" & stSalesArea1 & "'"
                stWhere2 = stWhere2 & " And SP48='" & stSalesArea1 & "'"
            ElseIf UCase(stFormN) = "FRM210117" Then
                stWhere1 = stWhere1 & " And SubStr(SP48,1,2)='" & stSalesArea1 & "'"
                stWhere2 = stWhere2 & " And SubStr(SP48,1,2)='" & stSalesArea1 & "'"
            'end 2018/04/19
            Else
                stWhere1 = stWhere1 & " And SP48>='" & stSalesArea1 & "'"
                stWhere2 = stWhere2 & " And " & IIf(bolSum = True, "O.SP48", "SP48") & ">='" & stSalesArea1 & "'"
                'If UCase(stFormN) = "FRMACC44J0" Then stWhere3 = stWhere3 & " And R001>='" & stSalesArea1 & "'"
            End If
        End If
        
        If stSalesArea2 <> "" Then
            stWhere1 = stWhere1 & " And SP48<='" & stSalesArea2 & "'"
            stWhere2 = stWhere2 & " And " & IIf(bolSum = True, "O.SP48", "SP48") & "<='" & stSalesArea2 & "'"
            'If UCase(stFormN) = "FRMACC44J0" Then stWhere3 = stWhere3 & " And R001<='" & stSalesArea2 & "'"
        End If
    End If
    'end 2018/01/25
    
    'Modify by Amy 2016/07/08 +SP19/SP40/F41XX
    If stEmpNo = "SP19" Or stEmpNo = "SP40" Then
        If stEmpNo = "SP19" Then
            stWhere2 = stWhere2 & " And SP19<>0 "
        Else
            stWhere2 = stWhere2 & " And SP40<>0 "
        End If
    'Add by Amy 2021/05/20 +特殊員編(非智權非外國部)
    ElseIf UCase(stEmpNo) = UCase("SpecNo") Then
        stTmp = Replace(Replace(智權點數實績與結餘特殊員編, "F4102;F4103;", ""), ";20091;F4104;F4105;F4106;F4107", "")
        stWhere2 = stWhere2 & " And " & IIf(bolSum = True, "O.SP02", "SP02") & " in ('" & Replace(stTmp, ";", "','") & "') " & _
                          IIf(UCase(stFormN) = "FRMACC41G0", "", "And " & IIf(bolSum = True, "O.SP15", "SP15") & "<>0")
    ElseIf stEmpNo = "F41XX" Then
        'Modify by Amy 2019/09/09 F41XX 改由王文安及陳鳳英輸個人欄位,FRMACC41G0轉傳票只抓SP15會無資料
        'modify by sonia 2021/1/21 +F4104~F4107
        stWhere2 = stWhere2 & " And " & IIf(bolSum = True, "O.SP02", "SP02") & " in ('F4101', 'F4102', 'F4103', 'F4104', 'F4105', 'F4106', 'F4107') " & _
                           IIf(UCase(stFormN) = "FRMACC41G0", "", "And " & IIf(bolSum = True, "O.SP15", "SP15") & "<>0")
    'end 2016/07/08
    'Add by Amy 2018/01/24 frmacc44j0實績與結餘分析表,M0100 會拆各部門故不抓/F4101 105年開始不使用故不抓
    ElseIf UCase(stFormN) = "FRMACC44J0" And stEmpNo = "" Then
        stWhere1 = stWhere1 & " And SP02 Not in ('F4101','M0100') "
        stWhere3 = stWhere3 & " And R003 Not in ('F4101','M0100') "
    Else
        'Modify by Amy 2024/02/26 目前業績達成日報表(FRM210107)-非智權 會傳入stEmpNo = "" And stSalesArea1 = "*NS"
        If UCase(stFormN) = "FRM210107" Then
            '抓傳入的部門 及 人員,因不同年月顯示不同資料
        ElseIf stEmpNo <> "" Then
            stWhere1 = stWhere1 & " And SP02='" & stEmpNo & "'"
            stWhere2 = stWhere2 & " And " & IIf(bolSum = True, "O.SP02", "SP02") & "='" & stEmpNo & "'"
            If UCase(stFormN) = "FRMACC44J0" Then stWhere3 = stWhere3 & " And R003='" & stEmpNo & "'"
        ElseIf stEmpNo = "" And stSalesArea1 = "*NS" Then
            'modify by sonia 2021/1/21 +F4104~F4107
            stWhere1 = stWhere1 & " And SP02 Not in ('F4101', 'F4102', 'F4103', 'F4104', 'F4105', 'F4106', 'F4107', 'M0100') "
        End If
    End If
    'end 2016/07/07
    'Add by Amy 2016/03/09 是否抓取F編號
    If bolFNo = False Then
        stWhere1 = stWhere1 & " And SubStr(SP02,1,1)<'F'"
        'Modify by Amy 2018/04/19
        stWhere2 = stWhere2 & " And SubStr(" & IIf(bolSum = True And bolReport = False, "O.SP02", "SP02") & ",1,1)<'F'"
    End If
    
    'Add by Amy 2024/02/22 +業績達成日報表-智權
    If UCase(stFormN) = "FRM210107" Then
      '[中區其他] 及[國外部] 11001前(不含),不再此抓因要[固定顯示],故條件使用stCon傳入

      'Memo 抓SalesPoint[有資料]([中區其他]並不一定有資料) 或 [有目標](新人可能有目標後一開始無點數)
      '              [客服組] W10 201906~202206有資料 202207(含)以後不顯示->SalesPoint及目標檔 202207後 沒資料(若為 固定顯示 需考慮SalesPoint是否一定有資料)
      '              [智權部]不需去掉員編大於6字頭且小於字頭之人員部門,因都是要出現的部門
      '              [開發組]併入[顧服組],因目前無資料未告知要如何顯示
      If stSalesArea1 = "*NS" Then
         stSQL = "Select Distinct '" & strUserNum & "',StateNo,ST02 as DepN,SP48,ST01 From SalesPoint,Staff Where SP02=ST01(+)" & stWhere1 & " "
      Else
         stSQL = "Select Distinct '" & strUserNum & "',StateNo,a0902 as DepN,SP48,SP48||'ZZ'  From SalesPoint,Acc090 Where SP48=a0901(+)" & stWhere1 & " "
      End If
    'end 2024/02/26
      GetPoint_SP = stSQL
      Exit Function
    End If
    
    '加總
    If bolSum = True Then
        '財務表單(自動轉傳票)才需判斷業務區,因10610中一人員調中二
        If Left(UCase(stFormN), 6) = "FRMACC" And UCase(stFormN) <> "FRMACC44J0" Then
            stSql_C = "Select SP02,SP48,Sum(Nvl(SP19,0)) as SP19,Sum(Nvl(SP40,0)) as SP40 From SalesPoint " & _
                            "Where " & Mid(stWhere1, 5) & " Group by SP02,SP48 "
        'Modify by Amy 2018/04/19 +if
        ElseIf UCase(stFormN) <> "FRM210113" And UCase(stFormN) <> "FRM210117" Then
            stSql_C = "Select SP02,Sum(Nvl(SP19,0)) as SP19,Sum(Nvl(SP40,0)) as SP40 From SalesPoint " & _
                            "Where " & Mid(stWhere1, 5) & " Group by SP02 "
        End If
        
        'Add by Amy 2016/07/07
        If UCase(stFormN) = "FRMACC44J0" Then
            stSQL = ",Sum(Decode(SP15,null,Decode(SP11,null,Decode(SP07,null,Nvl(SP03,0),SP07),SP11),SP15)) as SP15" & _
                        ",Sum(Decode(SP36,null,Decode(SP32,null,Decode(SP28,null,Nvl(SP24,0),SP28),SP32),SP36)) as SP36"
        'Add by Amy 2016/07/08
        ElseIf UCase(stFormN) = "FRMACC41G0" Then
            If stEmpNo = "SP19" Then
                stSQL = ",Sum(S.SP19) as SP19"
            Else
                 stSQL = ",Sum(Decode(SP15,null,Decode(SP11,null,Decode(SP07,null,Nvl(SP03,0),SP07),SP11),SP15)) as SP15"
            End If
        'Modify by Amy 2018/04/19 +if
        ElseIf UCase(stFormN) <> "FRM210113" And UCase(stFormN) <> "FRM210117" Then
            stSQL = ",Sum(Decode(SP15,null,Decode(SP11,null,Decode(SP07,null,Nvl(SP03,0),SP07),SP11),SP15)) as SP15, Sum(S.SP19) as SP19" & _
                        ",Sum(Decode(SP36,null,Decode(SP32,null,Decode(SP28,null,Nvl(SP24,0),SP28),SP32),SP36)) as SP36,Sum(S.SP40) as SP40"
        End If
    '非加總
    Else
        If UCase(stFormN) = "FRMACC41H0" Then
            stSQL = ",Nvl(SP40,0)*" & intMultiplier & " SP40,SP41"
        Else
            stSQL = ",Decode(SP15,null,Decode(SP11,null,Decode(SP07,null,Nvl(SP03,0),SP07),SP11),SP15) as SP15,Nvl(SP19,0) as SP19,SP20" & _
                        ",Decode(SP36,null,Decode(SP32,null,Decode(SP28,null,Nvl(SP24,0),SP28),SP32),SP36) as SP36,Nvl(SP40,0) as SP40,SP41"
        End If
    End If
    
    Select Case UCase(stFormN)
        Case "FRM210113" '各區業務工作報告統計
            'Modify by Amy 2024/04/19 [各區業務工作報告統計]語法改成一句,增加欄位別名(mSt0x),拿掉stOrderBy至各區業務工作報告統計再加
            'stOrderBy = " Order by Decode(ST05,'SM',1,'76',1,2),R01"
            'Modify by Amy 2018/04/19 若下S21 10609~10610 高國碩、陳頌恩資料會歸錯,因調區(期末抓止日,仍不會有期末實績、結餘資料)
            If bolSum = True Then
                '期末實績/結餘抓止月,轉撥抓畫面起迄 (畫面起迄日期 不同用)
                stSQL = "Select St02 as mSt02,Nvl(SP15,0) as SP15,Nvl(SP36,0) as SP36,Nvl(SP19,0) as SP19,Nvl(SP40,0) as SP40,R01 as mSt01 From Staff," & _
                             "(Select Distinct R01 From R210113 Where ID='" & strUserNum & "' And R04='1')," & _
                             "(Select R01 as SNo1,Decode(SP15,null,Decode(SP11,null,Decode(SP07,null,Nvl(SP03,0),SP07),SP11),SP15) as SP15 " & _
                                    ",Decode(SP36,null,Decode(SP32,null,Decode(SP28,null,Nvl(SP24,0),SP28),SP32),SP36) as SP36 " & _
                                    "From R210113,SalesPoint Where ID='" & strUserNum & "' And R04='1' And R01=SP02(+) And R03=SP01(+) " & stWhere2 & " )," & _
                              "(Select R01 as SNo2,SP48,Sum(Nvl(SP19,0)) as SP19,Sum(Nvl(SP40,0)) as SP40 From R210113,SalesPoint " & _
                                    "Where ID='" & strUserNum & "' And R04='1' And R01=SP02(+)  And R03=SP01(+) " & stWhere1 & " Group by R01,SP48 ) S " & _
                              "Where R01=SNo1(+) And R01=SNo2(+) And R01=ST01(+) " & stOrderBy
            Else
                stSQL = "Select R01 as mSt01,St02  as mSt02" & stSQL & " From R210113,SalesPoint, Staff Where ID='" & strUserNum & "' And R04='1' " & _
                            "And R01=SP02 And R03=SP01(+) And R01=ST01(+) " & stWhere2 & stOrderBy
            End If
            'end 2024/04/19
            GetPoint_SP = stSQL
            Exit Function
            
        Case "FRM210117"
            'Modify by Amy 2018/04/19 若下中區 10609~10610 高國碩、陳頌恩資料會歸錯,因調區
            stOrderBy = " Order by R02"
            If bolSum = True Then
                '期末實績/結餘抓止月,轉撥抓畫面起迄
                stSQL = "Select R02,Nvl(SP15,0) as SP15,Nvl(SP36,0) as SP36,Nvl(SP19,0) as SP19,Nvl(SP40,0) as SP40 From " & _
                             "(Select Distinct R02 From R210113 Where ID='" & strUserNum & "' And R04='2')," & _
                             "(Select R02 as SNo1,Sum(Decode(SP15,null,Decode(SP11,null,Decode(SP07,null,Nvl(SP03,0),SP07),SP11),SP15)) as SP15 " & _
                                    ",Sum(Decode(SP36,null,Decode(SP32,null,Decode(SP28,null,Nvl(SP24,0),SP28),SP32),SP36)) as SP36 " & _
                                    "From R210113,SalesPoint Where ID='" & strUserNum & "' And R04='2' And R01=SP02(+) And R03=SP01(+) " & stWhere2 & _
                                    " Group by R02 )," & _
                              "(Select R02 as SNo2,Sum(Nvl(SP19,0)) as SP19,Sum(Nvl(SP40,0)) as SP40 From R210113,SalesPoint " & _
                                    "Where ID='" & strUserNum & "' And R04='2' And R01=SP02(+)  And R03=SP01(+) " & stWhere1 & " Group by R02 ) S " & _
                              "Where R02=SNo1(+) And R02=SNo2(+) " & stOrderBy
            Else
                stSQL = "Select R02" & stSQL & "From R210113,SalesPoint Where ID='" & strUserNum & "' And R04='2' " & _
                              "And R01=SP02 And R03=SP01(+) " & stWhere2 & " Group by R02 " & stOrderBy
            End If
            GetPoint_SP = stSQL
            Exit Function
        'Add by Amy 2016/02/22
        Case "FRMACC44J0"
            'Modify by Amy 2016/07/07 跨月人員抓法有誤
'            stSQL = "Select O.SP02,O.SP48" & stSQL
'            stOrderBy = " Having Sum(Decode(SP15,null,Decode(SP11,null,Decode(SP07,null,Nvl(SP03,0),SP07),SP11),SP15))<>0 or Sum(S.SP19) <>0 " & _
'                            "or Sum(Decode(SP36,null,Decode(SP32,null,Decode(SP28,null,Nvl(SP24,0),SP28),SP32),SP36))<>0 or Sum(S.SP40)<>0 " & _
'                            "Group by O.SP48,O.SP02 Order by O.SP48,O.SP02"
            'Modify by Amy 2018/01/25 改不考慮區別,因10610中一部分人員調中二,若下10609~10610 限制區別會抓不到期末保留及結餘(期未抓止月)
            stSQL = "Select R003 as SP02,Nvl(SP15,0)*" & intMultiplier & " as SP15, Nvl(SP36,0)*" & intMultiplier & " as SP36,Nvl(SP19,0)*" & intMultiplier & " as SP19,Nvl(SP40,0)*" & intMultiplier & " as SP40 From Accrpt44j0_2" & _
                        ",(Select SP02 " & stSQL & " From SalesPoint O Where " & Mid(stWhere2, 5) & " Group by O.SP02) O, (" & stSql_C & ") S " & _
                        "Where ID='" & strUserNum & "' And R003=O.SP02(+) And R003=S.SP02(+) " & stWhere3 & _
                        " Order by R003"
            GetPoint_SP = stSQL
            Exit Function
        'Add by Amy 2016/07/08
        Case "FRMACC41G0"
            If bolSum = True Then
                If stEmpNo = "F41XX" Then
                    'Modify by Amy 2019/09/09
                    stOrderBy = " Group by O.SP01 Having Sum(Decode(SP15,null,Decode(SP11,null,Decode(SP07,null,Nvl(SP03,0),SP07),SP11),SP15)) <>0"
                    GetPoint_SP = "Select O.SP01" & stSQL & " From SalesPoint O, Staff, (" & stSql_C & ") S " & _
                                            "Where O.SP02=S.SP02(+) And O.SP02=ST01(+) " & stWhere2 & stOrderBy
                    Exit Function
                Else
                    stSQL = "Select SubStr(O.SP48,1,2) as Dept" & stSQL
                    stOrderBy = " Group by SubStr(O.SP48,1,2)"
                End If
            Else
                If stEmpNo = "F41XX" Then
                    'Modify by Amy 2019/09/09
                    GetPoint_SP = "Select * From " & _
                                                "(Select SP01,SP02" & stSQL & " From SalesPoint, Staff Where SP02=ST01(+) " & stWhere2 & ") " & _
                                           " Where SP15<>0  Order by SP02"
                    Exit Function
                ElseIf stEmpNo = "SP19" Or stEmpNo = "SP40" Then
                    stSQL = "Select SP01,SP02,ST02" & stSQL
                    stOrderBy = " Order by Decode(SubStr(" & IIf(stEmpNo = "SP19", "SP19", "SP40") & ",1,1),'-',1,2),SP48,SP02"
                Else
                    'Modify by Amy 2017/05/10 +實績<>0
                    stSQL = "Select SP01,SP02,SP48" & stSQL
                    stWhere2 = stWhere2 & " And Decode(SP15,null,Decode(SP11,null,Decode(SP07,null,Nvl(SP03,0),SP07),SP11),SP15)<>0 "
                    stOrderBy = " Order by SP48,SP02"
                End If
            End If
        'Add by Amy 2017/04/18
        Case "FRMACC41H0"
            stSQL = "Select SP01,SP02,ST02" & stSQL
            stOrderBy = " Order by Decode(Sign(SP40),-1,1,2),SP48,SP02"
        Case Else
            stSQL = "Select SP01,SP02" & stSQL
    End Select
    'Memo by Amy 此有改要確認 Case "FRMACC41G0" 是否也改
    If bolSum = True Then
        stSQL = stSQL & " From SalesPoint O, Staff, (" & stSql_C & ") S " & _
                    "Where O.SP02=S.SP02(+) And O.SP02=ST01(+) " & stWhere2 & stOrderBy
    Else
        stSQL = stSQL & " From SalesPoint, Staff " & _
                    "Where SP02=ST01(+) " & stWhere2 & stOrderBy
    End If
    GetPoint_SP = stSQL
    'end 2016/03/09
End Function

'Add by Amy 2017/09/26 抓「期初結餘保留/本月結餘放出」語法
'Modify by Amy 2022/06/10 +intChoose=0
'intChoose:0-個人當月是否有傳票資料/1-期初保留/2-本月放出/3-個人T/4-個人P/5-個人CFT/6-個人CFP
'stAx209:業務編號
Public Function GetBalanceSQL(ByVal intChoose As Integer, ByVal stDate_S As String, ByVal stDate_E As String, Optional ByVal stAx209 As String = "") As String
    Dim stField As String, stField2 As String, stSQL(1 To 2) As String, stWhere As String
    
    stField = "Dept,ST01,Sum(Decode(AX204,'T',AX207,0)) T,Sum(Decode(AX204,'P',AX207,0)) P,Sum(Decode(AX204,'CFT',AX207,0)) CFT,Sum(Decode(AX204,'CFP',AX207,0)) CFP, " & _
                "Sum(Decode(AX204,'T',AX207,0)+Decode(AX204,'P',AX207,0)+Decode(AX204,'CFT',AX207,0)+Decode(AX204,'CFP',AX207,0)) Tol "
    
    If stAx209 <> MsgText(601) Then stWhere = " And Ax209='" & stAx209 & "' "
   
    If intChoose >= 3 Then
        stField = "'" & intChoose & "' Ord"
        Select Case intChoose
            Case 3
                stField = stField & ",Sum(Decode(AX204,'T',AX207,0)) StVal "
                stWhere = stWhere & " And Ax204='T' "
            Case 4
                stField = stField & ",Sum(Decode(AX204,'P',AX207,0)) StVal "
                stWhere = stWhere & " And Ax204='P' "
            Case 5
                stField = stField & ",Sum(Decode(AX204,'CFT',AX207,0)) StVal "
                stWhere = stWhere & " And Ax204='CFT' "
            Case 6
                stField = stField & ",Sum(Decode(AX204,'CFP',AX207,0)) StVal "
                stWhere = stWhere & " And Ax204='CFP' "
        End Select
    End If
    
    '期初保留
    stSQL(1) = "Select SB02 as Dept,SB03 as ST01,AX209,AX204,AX207 From Acc021,SalesBalance Where SB01=" & stDate_S & " And SB03=ax209(+) " & _
              stWhere & "And ax202>='D'||" & stDate_S & " And ax202<'D'||" & Val(stDate_E) & " And ax205(+)='4194' And ax207>0 "
    '本月放出(當月)
    'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:(substr(ax205, 1, 2) = '41'
    stSQL(2) = "Select SB02 as Dept,SB03 as ST01,AX209,AX204,AX207 From Acc021,SalesBalance Where SB01=" & stDate_S & " And ax209=SB03(+) " & _
              "And ax202>='D'||" & stDate_S & " And ax202<'D'||" & Val(stDate_E) & " And ax207>0 And ax209 Is Not Null " & _
              stWhere & "And InStr(ax213,'結餘')>0 And SubStr(ax205,1,1)='4' And ax205<>'4194' "
       
    'Modify by Amy 2022/06/10 +if intchoose=0
    '個人當月是否有傳票資料
    If intChoose = 0 Then
        GetBalanceSQL = "Select * From (" & Replace(stSQL(1), ",SalesBalance Where SB01=" & stDate_S & " And SB03=ax209(+) ", " Where 1=1 ") & _
                                           " Union All " & Replace(stSQL(2), ",SalesBalance Where SB01=" & stDate_S & " And ax209=SB03(+) ", " Where 1=1 ") & _
                                                                ")"
        GetBalanceSQL = Replace(GetBalanceSQL, "SB02 as Dept,SB03 as ST01,", "")
    'end 2022/06/10
    '個人轉撥
    ElseIf intChoose >= 3 Then
        GetBalanceSQL = "Select " & stField & _
                        "From (" & stSQL(1) & " Union All " & stSQL(2) & ") "
    
    '期初保留/本月放出
    Else
        GetBalanceSQL = "Select " & stField & _
                        "From (" & stSQL(intChoose) & ") " & _
                        "Group by Dept,ST01 "
    End If
End Function

'Add by Amy 2018/09/18 傳入員編代出職稱
'intChoose: 0:不傳員工名/1:只傳姓+職稱(沒職稱帶全名)/2:傳全名
'bolAbbreviation:職稱縮寫
Public Function GetJobTitle(ByVal stST01 As String, Optional ByVal intChoose As Integer = 0, Optional ByVal bolAbbreviation As Boolean = True) As String
    Dim RsQ As ADODB.Recordset, strQ As String, intQ As Integer, stName As String, stJobTitle As String
    
    GetJobTitle = ""
    'Modify by Amy 2022/09/08 原:And '01'=AC01(+) A5020會串出多筆
    strQ = "Select st02,ac03 From Staff,AllCode Where ST01='" & stST01 & "' And ST20=AC02(+) And AC01='01' "
    'Added by Lydia 2025/08/15 A7010會無法回傳員工名
    strQ = strQ & " Union select st02,'' as ac03 from staff Where ST01='" & stST01 & "' And ST20 is null "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        'Modify by Amy 2018/11/05 沒職稱帶全名
        stName = "" & RsQ.Fields("st02")
        If intChoose = 0 Then stName = ""
        stJobTitle = "" & RsQ.Fields("ac03")
        If stJobTitle <> MsgText(601) And intChoose = 1 Then stName = Left(stName, 1)
        'end 2018/11/05
        If bolAbbreviation = True Then
            If stJobTitle = "副總經理" Then stJobTitle = "副總"
            If stJobTitle = "主任祕書" Then stJobTitle = "主秘" 'Add by Amy 2018/10/23
        End If
        'Modify by Amy 2024/04/19 職稱第一個字為[代] 拿掉 ex:李柏翰 代經理
        If Left(stJobTitle, 1) = "代" Then
            stJobTitle = Mid(stJobTitle, 2)
        End If
        'end 2024/04/19
        GetJobTitle = stName & stJobTitle & IIf(intChoose = 2, "　", "")
    End If
    RsQ.Close
    Set RsQ = Nothing
End Function

'Add by Amy 2022/09/29 商申承辦人責任業務區分配人員確認
'Modify by Amy 2022/10/05 + stReturnMailMsg
Public Function ChkDutyZoneAssign(ByVal stFormN As String, ByVal stNo As String, bolMsg As Boolean, bolMail As Boolean, Optional ByRef stMsg As String, Optional ByRef stReturnMailMsg As String) As Boolean
    Dim RsQ As ADODB.Recordset, strQ As String, intQ As Integer, stStateN As String, stTP(2) As String
    Dim stToNo As String, stToName As String, stSubject As String, stContent As String, stToManage As String
    
    ChkDutyZoneAssign = False: stMsg = ""
    stToManage = Pub_GetSpecMan("程式管理人員")
    stToNo = Pub_GetSpecMan("I")
    If stToNo = MsgText(601) Then
        stToNo = stToManage
    End If
    stToName = GetPrjSalesNM(stToNo)
    
    If Left(stNo, 1) = 客戶編號 Then
        stStateN = "客戶"
        '8碼
        strQ = "Select Dza01,St02 as WP,Dza02,Decode(cu04,null,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) as SName,1 as Count,'1' as State " & _
                  "From DutyZoneAssign,Staff,Customer Where Dza02='" & stNo & "' And Length(Dza02)=8 And Dza01=St01(+) " & _
                  "And Dza02=Cu01(+) And Cu02(+)='0' And Cu01 is not null "
        '6碼顯示8碼00名稱
        'Modify by Amy 2024/11/29 多寫了Dza02=Cu01(+)拿掉,避免6碼會抓不到 ex:X68546
        strQ = strQ & " Union " & _
                  "Select Dza01,WP,Dza02,Decode(cu04,null,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) as SName,Count,'2' as State From " & _
                  "(Select Dza01,St02 as WP,Dza02 From DutyZoneAssign,Staff Where Dza02='" & Left(stNo, 6) & "' And Length(Dza02)=6 And Dza01=St01(+) )," & _
                  "(Select SubStr(Cu01,1,6) as CuNo,Count(*) as Count From Customer Where SubStr(Cu01,1,6)='" & Left(stNo, 6) & "' And cu02='0' Group by SubStr(Cu01,1,6) ) " & _
                  ",Customer Where Dza02||'00'=Cu01(+) And Cu02(+)='0' And Cu01 is not null Order by State,Dza01"
    Else
        stStateN = "員工"
        '智權
        strQ = "Select Dza01,W.St02 as WP,Dza02,S.St02 as SName,0 as Count,'3' as State From DutyZoneAssign,Staff W,Staff S " & _
                    "Where Dza02='" & stNo & "' And Dza01=W.St01(+) And Dza02=S.St01(+) And S.St01 is not null "
        '承辦人
        strQ = strQ & " Union " & _
                   "Select Dza01,W.St02 as WP,Dza02,S.St02 as SName,0 as Count,'4' as State From DutyZoneAssign,Staff W,Staff S " & _
                   "Where Dza01='" & stNo & "' And Dza01=W.St01(+) And Dza02=S.St01(+) And W.St01 is not null Order by State,Decode(State,3,Dza01,Dza02)"
    End If
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        ChkDutyZoneAssign = True
        RsQ.MoveFirst
        If stTP(0) = MsgText(601) Then
            stTP(0) = stTP(0) & stStateN & "編號  "
            If "" & RsQ.Fields("State") = "4" Then
                stTP(0) = stTP(0) & RsQ.Fields("Dza01")
            Else
                stTP(0) = stTP(0) & RsQ.Fields("Dza02")
            End If
            If "" & RsQ.Fields("State") = "2" Then stTP(0) = stTP(0) & "（6碼）"
            stTP(0) = stTP(0) & "[名稱：" & RsQ.Fields("SName") & "]"
            stSubject = stTP(0) & " 已於「商申承辦人責任業務區分配」設定，請調整設定 ！"
        End If
                
        '客戶 8碼 / 6碼 只有一筆
        If "" & RsQ.Fields("State") = "1" Or ("" & RsQ.Fields("State") = "2" And Val("" & RsQ.Fields("Count")) = 1) Then
            stMsg = "此" & stStateN & "已於「商申承辦人責任業務區分配」設定" & vbCrLf
            If bolMail = True Then stMsg = stMsg & "系統已自動發信通知 " & stToName & " 處理！" & vbCrLf
            stContent = "您好," & vbCrLf & vbCrLf & stTP(0) & vbCrLf & "已設定「商申承辦人責任業務區分配」如下：" & vbCrLf & vbCrLf & _
                               "承  辦  人：" & RsQ.Fields("Dza01") & " " & RsQ.Fields("WP") & vbCrLf & vbCrLf
        '客戶 6碼 /智權/承辦
        Else
            Do While Not RsQ.EOF
                stTP(1) = ""
                '客戶 6 碼 /智權
                If "" & RsQ.Fields("State") = "2" Or "" & RsQ.Fields("State") = "3" Then
                    stTP(1) = stTP(1) & "承  辦  人：" & RsQ.Fields("Dza01") & " " & RsQ.Fields("WP") & vbCrLf
                '承辦人
                Else
                    stTP(1) = stTP(1) & "智權人員：" & RsQ.Fields("Dza02") & " " & RsQ.Fields("SName") & vbCrLf
                End If
                RsQ.MoveNext
                stMsg = stMsg & stTP(1)
            Loop
            If stMsg <> MsgText(601) Then
                stContent = "您好," & vbCrLf & vbCrLf & stTP(0) & vbCrLf & "已設定「商申承辦人責任業務區分配」如下：" & vbCrLf & vbCrLf & _
                                   stMsg & vbCrLf & vbCrLf
            End If
            '客戶基本檔維護-客戶 6 碼 不需 Mail及彈訊息
            If UCase(stFormN) = UCase("frm140401") Then
                ChkDutyZoneAssign = False
                bolMsg = False: bolMail = False
            End If
        End If
        If stMsg <> MsgText(601) Then
            '客戶基本檔維護
            If UCase(stFormN) = UCase("frm140401") Then
                stMsg = stMsg & "待人員調整完回覆後再行處理！"
                stContent = stContent & vbCrLf & "＊調整後，請回Mail通知" & strUserNum & "(" & GetPrjSalesNM(strUserNum) & ")！！"
            End If
        End If
        If stContent <> MsgText(601) Then stReturnMailMsg = stToNo & ";" & stSubject & ";" & stContent 'Add by Amy 2022/10/05
        If bolMail = True And stContent <> MsgText(601) Then
            PUB_SendMail strUserNum, stToNo, "", stSubject, stContent, , , , , , , , , , True, , , , False
        End If
        If bolMsg = True And stMsg <> MsgText(601) Then
            MsgBox stMsg, vbExclamation + vbOKOnly
        End If
    End If
    Set RsQ = Nothing
End Function

'Add by Amy 2022/09/29 取得cp36,若有「,」及「;」取,或; 前字元
Public Function GetCP36(ByVal stTxt As String) As String
    If InStr(stTxt, ",") = 0 And InStr(stTxt, ";") = 0 Then GetCP36 = stTxt: Exit Function
    
    If InStr(stTxt, ",") > 0 Then
        GetCP36 = Mid(stTxt, 1, InStr(stTxt, ",") - 1)
    Else
        GetCP36 = Mid(stTxt, 1, InStr(stTxt, ";") - 1)
    End If
End Function

'Add By Sindy 2022/12/2 是否留職停薪中
Public Function PUB_StaffChange04(StrST01 As String) As Boolean
   Dim RsQ As New ADODB.Recordset
   Dim strQ As String, intQ As Integer
   
   PUB_StaffChange04 = False
   
   strQ = "select sc01,sc03 from Staff_Change where sc01='" & StrST01 & "' and sc02=(select max(sc02) from Staff_Change where sc01='" & StrST01 & "')" & _
          " and sc03='04'"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      PUB_StaffChange04 = True
   End If
   RsQ.Close
   Set RsQ = Nothing
End Function

'Add by Amy 2023/09/01 身份證字號/統一編號與其他客戶(CU02='0')資料相同(從frm140401搬過來)
'stMsg:彈訊息用
'stBackData:發mail用 (因客戶檔檢查 客戶為學校,不檢查證號相同,但身份證/統編 重覆發信仍發)
'Modify by Amy 2024/06/13 +stFormN
'Modify by Amy 2024/08/30 +傳入要檢查客戶的cu15(0:個人/1:公司/2:學校/3:特殊機構)-目前客戶檔用
Public Function ChkCU11Same(ByVal stCU01 As String, ByVal stCU02 As String, ByVal stCU11 As String, ByRef stMsg As String, ByVal m_EditMode As Integer, _
  Optional ByVal stFormN As String = "", Optional ByVal stCU15 As String = "", Optional ByRef stBackData As String) As Boolean
    Dim RsQ As New ADODB.Recordset, intQ As Integer, strQ As String, strWhere As String
    
    'Add by Amy 2024/06/13
    Select Case UCase(stFormN)
      Case "FRM090801_NEW", "FRM090801" '接洽記錄單
      Case "FRM12040163" '風險檢查資料維護
         strWhere = "And cu80<>'設為對造' " '設為對造 要可存檔
      Case "FRM140401" '客戶檔維護
    End Select

    '修改
    If m_EditMode = 2 Then
        'Modify by Amy 2023/08/28 +前8碼同不檢查 ex:X87157001
        strWhere = strWhere & "And cu01||cu02<>'" & stCU01 & stCU02 & "' And cu01<>'" & stCU01 & "' "
    End If
    strWhere = strWhere & "And cu11='" & stCU11 & "' "
    'Modify by Amy 2024/07/03 +66666666,有加排除的編號,要確認 FRM140401.ChkShowMsg 是否也改
    strQ = "Select * From Customer Where CU02='0' And (cu80<>'不再使用' or cu80 is null) And cu11 is not null " & _
                  "And cu11<>'00000000' And cu11<>'66666666' " & strWhere
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        ChkCU11Same = True
        Do While RsQ.EOF = False
            'Modify by 2024/08/30 +if
            If UCase(stFormN) = "FRM140401" And stCU15 = "2" And stCU15 = RsQ.Fields("cu15") Then
               '傳入的客戶為學校,且檢查到相同的證號也為學校,不需彈訊息-秀玲
            Else
               stMsg = stMsg & "," & RsQ.Fields("cu01")
            End If
            stBackData = stBackData & "," & RsQ.Fields("cu01") '客戶檔發信用
            'end 2024/08/30
            RsQ.MoveNext
        Loop
        If stMsg <> MsgText(601) Then
            stMsg = Mid(stMsg, 2)
        End If
        'Add by Amy 2024/08/30 客戶檔發信用
        If stBackData <> MsgText(601) Then
         stBackData = Mid(stBackData, 2)
        End If
    End If
    Set RsQ = Nothing
End Function

'Add By Sindy 2023/10/20
'取得有權限查看工作評價的同仁清單
'strGrade=1=B0130.主管1
'         2=B0131.主管2
'         3=B0132.主管3
Public Function PUB_GetSTJLimitEmp(Optional ByVal strGrade As String = "") As String
   PUB_GetSTJLimitEmp = ""
   
   '人事處,01=總經理,所長,副所長
   If Pub_StrUserSt03 = "M21" Or Pub_strUserST05 = "01" Then
      PUB_GetSTJLimitEmp = "ALL"
      Exit Function
   End If
   
   strSql = "select B0101,B0130,B0131,B0132 from ABS001 where B0130='" & strUserNum & "' or B0131='" & strUserNum & "' or B0132='" & strUserNum & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      While Not RsTemp.EOF
         If strGrade = "1" And _
            "" & RsTemp.Fields("B0130") = strUserNum Then
            PUB_GetSTJLimitEmp = PUB_GetSTJLimitEmp & "," & Trim(RsTemp.Fields("B0101").Value)
         ElseIf strGrade = "2" And _
                "" & RsTemp.Fields("B0131") = strUserNum Then
            PUB_GetSTJLimitEmp = PUB_GetSTJLimitEmp & "," & Trim(RsTemp.Fields("B0101").Value)
         ElseIf strGrade = "3" And _
                "" & RsTemp.Fields("B0132") = strUserNum Then
            PUB_GetSTJLimitEmp = PUB_GetSTJLimitEmp & "," & Trim(RsTemp.Fields("B0101").Value)
         Else
            PUB_GetSTJLimitEmp = PUB_GetSTJLimitEmp & "," & Trim(RsTemp.Fields("B0101").Value)
         End If
         RsTemp.MoveNext
      Wend
   End If
   If PUB_GetSTJLimitEmp <> "" Then
      PUB_GetSTJLimitEmp = Mid(PUB_GetSTJLimitEmp, 2)
   End If
End Function

'Add By Sindy 2023/10/25
Public Sub PUB_OpenFrm160016(oForm As Object)
Dim strSTJLimitEmp As String
   
   strSTJLimitEmp = PUB_GetSTJLimitEmp '取得有權限查看工作評價的同仁清單
   If InStr(strSTJLimitEmp, "") > 0 Or strSTJLimitEmp = "ALL" Then
      oForm.Show
   Else
      MsgBox "無此使用權限...", , "警告!!"
      Exit Sub
   End If
End Sub

'Added by Morgan 2023/11/7
'全勤名單語法
Public Function PUB_GetFullAttendanceStaff(pDateFrom As String, pDateTo As String) As String
   PUB_GetFullAttendanceStaff = "select * from staff " & _
      " where ST04='1' and st13<=" & pDateFrom & " and st01>'6' and st01<'F' and substr(ST01,4,1)<>'9' and substr(st03,1,1)<>'R'" & _
      " and ST01 not in('67001','68007','86026','68091','68092','94099','97099','99998','60000','99997','99099','68099','99999','96029','96030','68096','73029','99029')" & _
      " and nvl(st20,' ') not in ('01','02','21','22','11','12','15') and st01 not in('94015','79037') " & _
      " and not exists(select * from Staff_Assist where SA02 between " & pDateFrom & " and " & pDateTo & " and (SA04>0 or SA05>0 or SA06>0) and sa01=st01)" & _
      " and not exists(select * from Staff_Absence where SA02 between " & pDateFrom & " and " & pDateTo & " and SA06 in ('05','06') and sa01=st01)" & _
      " and not exists(select * from Staff_Change where SC02 between " & pDateFrom & " and " & pDateTo & " and SC03='02' and sc01=st01) "
End Function

'Add By Sindy 2023/11/29
'傳入:日期,時分,加減分鐘(例如:-1 或 1)
'回傳:hhmm=時分 ex:0759
Public Function PUB_DTtoDateAdd(strDate As String, strMMHH As String, intdiff As Integer) As String
Dim strDTtime As String
   
   strMMHH = Format(strMMHH, "00:00")
   strDTtime = Format(DateAdd("n", intdiff, ChangeWStringToWDateString(DBDATE(strDate)) & " " & strMMHH), "hhmm")
   PUB_DTtoDateAdd = strDTtime
End Function

'Add By Sindy 2023/12/19
'設定新部門 ST93 的下拉式選單 - 依在職人員列出全部部門(ex:出缺勤)
Public Sub SetST93Combo(ByRef cboTemp As Object)
Dim Rs As New ADODB.Recordset
   
   cboTemp.Clear
   cboTemp.AddItem ""
   
   '2014/2/11 modify by sonia 除電腦中心及人事處外,其他人只能看到有在職員工的部門(王副總提需求江總同意)
   'rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' Order By A0901", _
            cnnConnection, adOpenStatic, adLockReadOnly
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M21" Then
      strSql = "select A0921,A0922 From ACC090New order by A0921 asc"
   Else
      strSql = "select A0921,A0922 From ACC090New where A0921 in(" & _
               "select A0921 From staff,ACC090New " & _
               "where substr(st01,1,1) in (" & ST01CodeNum1 & ") " & _
               "and st04='1' " & _
               "and substr(st01,4,1)<>'9' " & _
               "and st01 not in('60000','96029','96030') " & _
               "and ST93=A0921(+) and A0921 is not null group by A0921) " & _
               "order by A0921 asc "
   End If
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   While Not Rs.EOF
      cboTemp.AddItem Left(Rs.Fields(0).Value & Space(5), 5) & Rs.Fields(1).Value
      Rs.MoveNext
   Wend
   If Rs.State <> adStateClosed Then Rs.Close
   Set Rs = Nothing
   If cboTemp.ListCount > 0 Then cboTemp.ListIndex = 0
End Sub

'Add By Sindy 2023/12/19
'設定新部門 A0921 的下拉式選單 - ACC090New全部部門(ex:人事系統)
Public Sub SetA0921Combo(ByRef cboTemp As Object)
Dim Rs As New ADODB.Recordset
   
   cboTemp.Clear
   cboTemp.AddItem ""
   Rs.CursorLocation = adUseClient
   strSql = "select A0921,A0922 " & _
            "From ACC090New " & _
            "order by A0921 asc "
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   While Not Rs.EOF
      cboTemp.AddItem Left(Rs.Fields(0).Value & Space(5), 5) & Rs.Fields(1).Value
      Rs.MoveNext
   Wend
   If Rs.State <> adStateClosed Then Rs.Close
   Set Rs = Nothing
   If cboTemp.ListCount > 0 Then cboTemp.ListIndex = 0
End Sub

'Add By Sindy 2025/3/19 調整為共用函數
'出缺勤查詢畫面欄位預設值
'  frm180301 - 出缺勤查詢
'  frm180303 - 打卡資料查詢
'  frm160018 - 下班逾30分鐘原因確認 Add By Sindy 2025/10/15
'm_strEmp=所屬簽核的人員
'strFormNm=傳入Form Name
Public Sub PUB_SetQFormCol_ABS(ByRef m_IsAbsBossST03 As String, ByRef m_strEmp As String, ByVal strFormNm As String, _
   ByRef oTxtDate_0 As Object, ByRef oTxtDate_1 As Object, _
   ByRef oTxtB1003_0 As Object, ByRef oTxtB1003_1 As Object, _
   ByRef oCboDept_0 As Object, ByRef oCboDept_1 As Object, _
   ByRef oTxtDept_0 As Object, ByRef oTxtDept_1 As Object, _
   ByRef otxtST06_0 As Object, ByRef otxtST06_1 As Object, Label5 As Object)

'宣告變數
Dim Rs As New ADODB.Recordset
Dim ii As Integer
   
   '設定起迄部門
   oCboDept_0.Clear
   oCboDept_1.Clear
   Rs.CursorLocation = adUseClient
   '2014/2/11 modify by sonia 除電腦中心及人事處外,其他人只能看到有在職員工的部門(王副總提需求江總同意)
   'rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' Order By A0901", _
            cnnConnection, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2023/12/19
   If strSrvDate(1) >= 新部門啟用日 Then
      Call SetST93Combo(oCboDept_0)
      Call SetST93Combo(oCboDept_1)
   Else
      If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M21" Then
         Rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' Order By A0901", _
                  cnnConnection, adOpenStatic, adLockReadOnly
      Else
         Rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' and a0901<>'P29' and a0901 in (select distinct st03 from staff where st04='1' and st01>'6' and substr(st01,1,1)<'G' and substr(st01,4,1)<>'9') Order By A0901", _
                  cnnConnection, adOpenStatic, adLockReadOnly
      End If
      '2014/2/11 end
      oCboDept_0.AddItem ""
      oCboDept_1.AddItem ""
      While Not Rs.EOF
         oCboDept_0.AddItem Left(Rs.Fields(0).Value & Space(5), 5) & Rs.Fields(1).Value
         oCboDept_1.AddItem Left(Rs.Fields(0).Value & Space(5), 5) & Rs.Fields(1).Value
         Rs.MoveNext
      Wend
      If Rs.State <> adStateClosed Then Rs.Close
      Set Rs = Nothing
   End If
   'Modify By Sindy 2023/12/19
   If strSrvDate(1) >= 新部門啟用日 Then
      oTxtDept_0 = Pub_StrUserSt93
      oTxtDept_1 = Pub_StrUserSt93
   Else
      oTxtDept_0 = Pub_StrUserSt03
      oTxtDept_1 = Pub_StrUserSt03
   End If
   'Add By Sindy 2021/12/21
   If InStr(Pub_GetSpecMan("專利處出缺勤可查詢權限"), strUserNum) > 0 Then
      'Modify By Sindy 2023/12/19
      If strSrvDate(1) >= 新部門啟用日 Then
'         oTxtDept_0 = "P00"
'         oTxtDept_1 = "P41" 'Modify By Sindy 2024/1/30 "P99"
         'Modify By Sindy 2025/3/5 用SQL抓出起迄部門別
         strSql = "select A0921,A0922 From ACC090New where A0921 in(" & _
                  "select A0921 From staff,ACC090New " & _
                  "where substr(st01,1,1) in (" & ST01CodeNum1 & ") " & _
                  "and st04='1' and substr(st93,1,1)='P'" & _
                  "and substr(st01,4,1)<>'9' " & _
                  "and st01 not in('60000','96029','96030') " & _
                  "and ST93=A0921(+) and A0921 is not null group by A0921) " & _
                  "order by A0921 asc "
         Rs.CursorLocation = adUseClient
         Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If Rs.RecordCount > 0 Then
            Rs.MoveFirst
            oTxtDept_0 = Rs.Fields("A0921")
            Rs.MoveLast
            oTxtDept_1 = Rs.Fields("A0921")
         End If
         '2025/3/5 END
      Else
      '2023/12/19 END
         oTxtDept_0 = "P10"
         oTxtDept_1 = "P14"
      End If
   End If
   For ii = 1 To oCboDept_0.ListCount - 1
      If Left(oCboDept_0.List(ii), Len(oTxtDept_0)) = oTxtDept_0 Then
         oCboDept_0.ListIndex = ii
         Exit For
      End If
   Next ii
   For ii = 1 To oCboDept_1.ListCount - 1
      If Left(oCboDept_1.List(ii), Len(oTxtDept_1)) = oTxtDept_1 Then
         oCboDept_1.ListIndex = ii
         Exit For
      End If
   Next ii
   '2021/12/21 END
   
   If UCase(strFormNm) = "FRM180303" Then
      '當月1日
      oTxtDate_0 = Left(ChangeWStringToTString(strSrvDate(1)), 5) & "01"
      oTxtDate_1 = strSrvDate(2)
   'Add By Sindy 2025/10/15
   ElseIf UCase(strFormNm) = "FRM160018" Then
      '上個月1日
      oTxtDate_0 = Left(ChangeWStringToTString(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##")))), 5) & "01"
      oTxtDate_1 = strSrvDate(2)
   '2025/10/15 END
   Else
      oTxtDate_0 = strSrvDate(2) 'CStr((Val(Left(strSrvDate(1), 4)) - 1911)) & "0101"
      oTxtDate_1 = strSrvDate(2)
   End If
'   oTxtDept_0 = Pub_StrUserSt03
'   oTxtDept_1 = Pub_StrUserSt03
   oTxtB1003_0 = strUserNum
   oTxtB1003_1 = strUserNum
   
   Label5.Visible = False
   '取得審核主管的部門別權限
   'Modify By Sindy 2023/12/19
   If strSrvDate(1) >= 新部門啟用日 Then
      m_IsAbsBossST03 = GetIsAbsBossST93(strUserNum, m_strEmp)
   Else
   '2023/12/19 END
      m_IsAbsBossST03 = GetIsAbsBossST03(strUserNum, m_strEmp)
   End If
   
   '電腦中心,M21人事處,部份決策會人員開放可以查所有員工的資料,不要鎖
   '76012桂齊恆
   '68001江丕群
   '67001蘇國安
   '81040閰�K泰
   '94007林景郁
   'Modify By Sindy 2011/10/11 開放M71管理部分所權限
   'Modify By Sindy 2015/11/2 +68009.何主秘開放查詢權限
   'Modify By Sindy 2025/5/12 改用系統特殊設定
'   If Pub_StrUserSt03 = "M71" Or Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M21" Or _
'      strUserNum = "76012" Or strUserNum = "68001" Or strUserNum = "68009" Or strUserNum = "67001" Or _
'      strUserNum = "81040" Or strUserNum = "94007" Then
   If Pub_StrUserSt03 = "M71" Or Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M21" Or _
      InStr(Pub_GetSpecMan("全所出缺勤可查詢權限"), strUserNum) > 0 Then
      m_IsAbsBossST03 = ""
      'Label5 = Label5.Caption & IIf(m_IsAbsBossST03 = "", "ALL", Replace(m_IsAbsBossST03, "'", ""))
      'Label5.Visible = True
      If Pub_StrUserSt03 = "M71" Then
         otxtST06_0 = PUB_GetST06(strUserNum)
         otxtST06_1 = PUB_GetST06(strUserNum)
         otxtST06_0.Enabled = False
         otxtST06_1.Enabled = False
      End If
   Else
      '非審核主管身份者,只能查看自己的表單
      If ChkIsAbsBoss(strUserNum) = False Then
         oTxtDept_0.Enabled = False
         oTxtDept_1.Enabled = False
         oCboDept_0.Enabled = False
         oCboDept_1.Enabled = False
         oTxtB1003_0.Enabled = False
         oTxtB1003_1.Enabled = False
      Else
         '只有自己部門的權限
         If Replace(m_IsAbsBossST03, "'", "") = IIf(strSrvDate(1) >= 新部門啟用日, Pub_StrUserSt93, Pub_StrUserSt03) Then
            oTxtDept_0.Enabled = False
            oTxtDept_1.Enabled = False
            oCboDept_0.Enabled = False
            oCboDept_1.Enabled = False
'         Else
'            'Add By Sindy 2011/10/14 林副理權責權限的部門別只有P21,但自己又是P20,所以加入自己所屬部門別
'            If m_IsAbsBossST03 <> "" And InStr(m_IsAbsBossST03, Pub_StrUserSt03) <= 0 Then
'               m_IsAbsBossST03 = m_IsAbsBossST03 & ",'" & Pub_StrUserSt03 & "'"
'            End If
         End If
         Label5 = Label5.Caption & IIf(m_IsAbsBossST03 = "", "ALL", Replace(m_IsAbsBossST03, "'", "")) & " (只可查看您簽核的人員)"
         Label5.Visible = True
      End If
   End If
   
   '固定權限
   'Modify By Sindy 2022/7/19
   If InStr(Pub_GetSpecMan("專利處出缺勤可查詢權限"), strUserNum) > 0 Then
      'oTxtDate_0 = strSrvDate(2)
      'Modify By Sindy 2025/3/31 柏翰:請協助修改成起迄日期都預設當日
'      '前一工作天
'      oTxtDate_0 = ChangeWStringToTString(CompWorkDay(2, strSrvDate(1), 1)) 'ChangeWStringToTString(DBDATE(DateAdd("d", -1, Format(strSrvDate(1), "####/##/##"))))
'      oTxtDate_1 = ChangeWStringToTString(DBDATE(DateAdd("m", 1, Format(strSrvDate(1), "####/##/##"))))
      oTxtB1003_0 = ""
      oTxtB1003_1 = ""
   Else
   '2022/7/19 END
      'Add By Sindy 2012/2/10
      Select Case strUserNum
'         Case "71011", "99050" '王副總,李柏翰
'            oTxtDate_0 = strSrvDate(2)
'            oTxtDate_1 = ChangeWStringToTString(DBDATE(DateAdd("m", 1, Format(strSrvDate(1), "####/##/##"))))
'      '      oTxtDept_0 = "P10"
'      '      oTxtDept_1 = "P14"
'            oTxtB1003_0 = ""
'            oTxtB1003_1 = ""
         '2012/2/10 End
         'Add By Sindy 2022/5/3
         Case "82026" '林柄佑
            otxtST06_0 = PUB_GetST06(strUserNum)
            otxtST06_1 = PUB_GetST06(strUserNum)
            otxtST06_0.Enabled = False
            otxtST06_1.Enabled = False
         Case Else
      End Select
   End If
End Sub
 
'Added by Lydia 2025/08/14 利益衝突案件：檢查利益衝突案件之權限，若人事有異動留下相關記錄---參考PUB_ChkCuFa_Right
Public Sub PUB_SaveCUFA_Staff_Log(ByVal pBolTrans As Boolean, ByVal pUserNo As String, ByVal pKind As String, ByVal pFrmName As String, Optional ByVal pHostName As String, Optional ByVal pNewST03 As String, Optional ByVal pOldST03 As String, _
      Optional ByVal pNewST05 As String, Optional ByVal pOldSt05 As String, Optional ByVal pNewST16 As String, Optional ByVal pOldSt16 As String, Optional ByVal pNewST70 As String, Optional ByVal pOldSt70 As String)
'pBolMail: 是否發Email
'pUserNo: 傳入修改的員工編號/部門ST03/等級ST05
'pKind:調整方式
'pFrmName、pHostName:傳入操作的程式名稱、電腦名稱
Dim intQ As Integer, intJ As Integer
Dim strConB As String, strConA As String, strQ1 As String
Dim strMemo As String, strKindA As String, strKindB As String, strNowNo As String
Dim tmpArr As Variant
Dim rsQD As New ADODB.Recordset
Dim rsBD As New ADODB.Recordset
Dim bolConn As Boolean

'參考PUB_ChkCuFa_Right的檢查方式
'1.以X/Y編號＋系統類別＋操作員工編號檢查
'2.以X/Y編號＋系統類別＋操作者部門檢查
'3.以X/Y編號＋系統類別＋操作者等級檢查
'4.以X/Y編號＋系統類別＋CFR03=操作者部門(ST03)+工程師組別(ST16)檢查／操作者部門(ST03)+FCP工程師組別(ST16)+日本部組別(ST70,114/8/1新增)檢查
'5.以X/Y編號＋系統類別＋操作員工編號＋CFR04=開放工程師組別檢查 ---同1.不用再次檢查
 '--------------------------------------

   If pUserNo = "" Or pFrmName = "" Then Exit Sub
   If Mid(pUserNo, 4, 1) = "9" Then Exit Sub '虛建帳戶不用記錄
   
On Error GoTo EXITSUB
   If Trim(pNewST03) = "" Or Trim(pNewST05) = "" Or Trim(pNewST16) = "" Or Trim(pNewST70) = "" Then
      strQ1 = "select st03,st05,st16,st70 from staff where st01='" & pUserNo & "' "
      intQ = 1
      Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
      If intQ = 1 Then
         If Trim(pNewST03) = "" Then pNewST03 = "" & rsQD.Fields("st03")
         If Trim(pNewST05) = "" Then pNewST05 = "" & rsQD.Fields("st05")
         If Trim(pNewST16) = "" Then pNewST16 = "" & rsQD.Fields("st16")
         If Trim(pNewST70) = "" Then pNewST70 = "" & rsQD.Fields("st70")
      End If
   End If

   If pHostName = "" Then
      pHostName = pub_HostName
   End If

   strConB = "SELECT '10' AS ord1,'" & pUserNo & "' AS grpno,cfr01 FROM cufa_right WHERE cfr03='" & pUserNo & "' " & "AND cfr01 IN (SELECT cfr01 FROM cufa_right WHERE cfr03='M51' and cfr08='Y' GROUP BY cfr01) group by cfr01 "
   strConB = strConB & " Union SELECT '2A' AS ord1,'" & pNewST03 & "' AS grpno,cfr01 FROM cufa_right WHERE cfr03='" & pNewST03 & "' " & "AND cfr01 IN (SELECT cfr01 FROM cufa_right WHERE cfr03='M51' and cfr08='Y' GROUP BY cfr01) group by cfr01"
   If pOldST03 <> "" Then
      strConB = strConB & " Union SELECT '2B' AS ord1,'" & pOldST03 & "' AS grpno,cfr01 FROM cufa_right WHERE cfr03='" & pOldST03 & "' " & "AND cfr01 IN (SELECT cfr01 FROM cufa_right WHERE cfr03='M51' and cfr08='Y' GROUP BY cfr01) group by cfr01"
   End If
   strConB = strConB & " Union SELECT '3A' AS ord1,'" & pNewST05 & "' AS grpno,cfr01 FROM cufa_right WHERE cfr03='" & pNewST05 & "' AND cfr01 IN (SELECT cfr01 FROM cufa_right WHERE cfr03='M51' and cfr08='Y' GROUP BY cfr01) group by cfr01"
   If pOldSt05 <> "" Then
      strConB = strConB & " Union SELECT '3B' AS ord1,'" & pOldSt05 & "' AS grpno,cfr01 FROM cufa_right WHERE cfr03='" & pOldSt05 & "' " & "AND cfr01 IN (SELECT cfr01 FROM cufa_right WHERE cfr03='M51' and cfr08='Y' GROUP BY cfr01) group by cfr01"
   End If
   If InStr(pNewST03 & "," & pOldST03, "F21") > 0 Then
      strConB = strConB & " Union SELECT '4A' AS ord1,'" & pNewST03 & pNewST16 & "' AS grpno,cfr01 FROM cufa_right WHERE CFR03 LIKE 'F21%' AND cfr03='" & pNewST03 & pNewST16 & "' " & "AND cfr01 IN (SELECT cfr01 FROM cufa_right WHERE cfr03='M51' and cfr08='Y' GROUP BY cfr01) group by cfr01"
      strConB = strConB & " Union SELECT '4B' AS ord1,'" & pOldST03 & pOldSt16 & "' AS grpno,cfr01 FROM cufa_right WHERE CFR03 LIKE 'F21%' AND cfr03='" & pOldST03 & pOldSt16 & "' " & "AND cfr01 IN (SELECT cfr01 FROM cufa_right WHERE cfr03='M51' and cfr08='Y' GROUP BY cfr01) group by cfr01"
      If pNewST70 & pOldSt70 <> "" Then
         strConB = strConB & " Union SELECT '5A' AS ord1,'" & pNewST03 & pNewST16 & pNewST70 & "' AS grpno,cfr01 FROM cufa_right WHERE CFR03 LIKE 'F21%' AND cfr03='" & pNewST03 & pNewST16 & pNewST70 & "' " & "AND cfr01 IN (SELECT cfr01 FROM cufa_right WHERE cfr03='M51' and cfr08='Y' GROUP BY cfr01) group by cfr01"
         strConB = strConB & " Union SELECT '5B' AS ord1,'" & pOldST03 & pOldSt16 & pOldSt70 & "' AS grpno,cfr01 FROM cufa_right WHERE CFR03 LIKE 'F21%' AND cfr03='" & pOldST03 & pOldSt16 & pOldSt70 & "' " & "AND cfr01 IN (SELECT cfr01 FROM cufa_right WHERE cfr03='M51' and cfr08='Y' GROUP BY cfr01) group by cfr01"
      End If
   End If
   
   strMemo = ""
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strConB)
   If intQ = 1 Then
      '人事異動資料
      If UCase(pFrmName) = "FRM160007" Then
         If pKind = "02" Then '02復職
            strMemo = "復職"
         ElseIf InStr("03、04、08、09、10", pKind) > 0 Then '03離職、04留職停薪、08退休、09撤職、10資遣
            strMemo = "離職"
         End If
         If strMemo <> "" Then
            If pBolTrans = True And bolConn = False Then cnnConnection.BeginTrans
            bolConn = True
            strQ1 = "INSERT INTO CUFA_Staff_Log (csl01,csl02,csl03,csl04,csl05,csl06,csl07,csl08,csl09) " & _
                    "SELECT cfr01 AS csl01, '" & strUserNum & "' AS csl02, to_char(SYSDATE,'yyyymmdd') AS csl03, to_char(sysdate,'hh24miss') AS csl04, " & _
                    "'" & pFrmName & "' AS CSL05,'" & pUserNo & "' AS CSL06,SYS_CONTEXT('USERENV','IP_ADDRESS') AS csl07,'" & pHostName & "' AS csl08, " & _
                    "'" & strMemo & "' AS memo1 from (" & strConB & ") group by cfr01 "
            cnnConnection.Execute strQ1
         Else
            GoTo EXITSUB
         End If
      End If
      '電腦中心之員工檔維護frm12040105
      If UCase(pFrmName) = "FRM12040105" Then
         If pKind = "01" Then '01新建
            strMemo = "到職"
         ElseIf pKind = "02" Then '02離職
            strMemo = "離職"
         End If
         If strMemo <> "" Then
JumpToExcept: 'Added by Lydia 2025/08/18
            If pBolTrans = True And bolConn = False Then cnnConnection.BeginTrans
            bolConn = True
            strQ1 = "INSERT INTO CUFA_Staff_Log (csl01,csl02,csl03,csl04,csl05,csl06,csl07,csl08,csl09) " & _
                    "SELECT cfr01 AS csl01, '" & strUserNum & "' AS csl02, to_char(SYSDATE,'yyyymmdd') AS csl03, to_char(sysdate,'hh24miss') AS csl04, " & _
                    "'" & pFrmName & "' AS CSL05,'" & pUserNo & "' AS CSL06,SYS_CONTEXT('USERENV','IP_ADDRESS') AS csl07,'" & pHostName & "' AS csl08, " & _
                    "'" & strMemo & "' AS memo1 from (" & strConB & ") group by cfr01 "
            cnnConnection.Execute strQ1
         Else
            rsQD.MoveFirst
            Do While Not rsQD.EOF
               If "" & rsQD.Fields("ord1") <> "10" Then
                  strConA = "select * from (" & strConB & ") where ord1='" & Mid("" & rsQD.Fields("ord1"), 1, 1) & IIf(Mid("" & rsQD.Fields("ord1"), 2, 1) = "A", "B", "A") & "' " & _
                            "and grpno='" & rsQD.Fields("grpno") & "' and cfr01='" & rsQD.Fields("cfr01") & "' "
                  intJ = 1
                  Set rsBD = ClsLawReadRstMsg(intJ, strConA)
                  If intJ = 0 Then
                     '舊部門ST03、等級ST05及組別ST16+ST70有修改，且修改前或修改後其中之一符合條件就要寫Log：
                     '由有權限改為無權限記錄: 取消權限
                     '由無權限改為有權限記錄: 開放權限
                     If Mid("" & rsQD.Fields("ord1"), 2, 1) = "A" Then
                        If InStr(strKindA & ";", rsQD.Fields("cfr01")) = 0 Then
                           strKindA = strKindA & Mid("" & rsQD.Fields("ord1"), 1, 1) & "開放權限" & "，" & rsQD.Fields("cfr01") & ";"
                        End If
                     Else
                        If InStr(strKindB & ";", rsQD.Fields("cfr01")) = 0 Then
                           strKindB = strKindB & Mid("" & rsQD.Fields("ord1"), 1, 1) & "取消權限" & "，" & rsQD.Fields("cfr01") & ";"
                        End If
                     End If
                  End If
               End If
               rsQD.MoveNext
            Loop
            '讀取所有變更，再分析最後結果
            If strKindB <> "" Then
               tmpArr = Empty
               tmpArr = Split(strKindB, ";")
               For intQ = 0 To UBound(tmpArr)
                   strQ1 = Trim(tmpArr(intQ))
                   If strQ1 <> "" Then
                      strNowNo = Mid(strQ1, InStr(strQ1, "，") + 1)
                      strConA = "select * from (" & strConB & ") where substr(ord1,1,1)<>'" & Mid(strQ1, 1, 1) & "' and (substr(ord1,2,1)='A' or substr(ord1,2,1)='0') and cfr01='" & strNowNo & "' "
                      intJ = 1
                      Set rsBD = ClsLawReadRstMsg(intJ, strConA)
                      If intJ = 0 Then
                         If InStr(strKindA, strNowNo) = 0 Then  '排除仍有權限
                            If pBolTrans = True And bolConn = False Then cnnConnection.BeginTrans
                            bolConn = True
                            strMemo = "取消權限，" & IIf(pOldST03 <> pNewST03, "修改部門：" & pOldST03 & "=>" & pNewST03 & "，", "") & _
                                       IIf(pOldSt05 <> pNewST05, "修改等級：" & pOldSt05 & "=>" & pNewST05 & "，", "") & _
                                       IIf(pOldSt16 <> pNewST16, "修改工程師組別：" & pOldSt16 & "=>" & pNewST16 & "，", "") & _
                                       IIf(pOldSt70 <> pNewST70, "修改國外部小組：" & pOldSt70 & "=>" & pNewST70 & "，", "")
                            strQ1 = "INSERT INTO CUFA_Staff_Log (csl01,csl02,csl03,csl04,csl05,csl06,csl07,csl08,csl09) values (" & _
                                   "'" & strNowNo & "', '" & strUserNum & "', to_char(SYSDATE,'yyyymmdd'),to_char(sysdate,'hh24miss'), " & _
                                   "'" & pFrmName & "','" & pUserNo & "',SYS_CONTEXT('USERENV','IP_ADDRESS'),'" & pHostName & "','" & strMemo & "') "
                            cnnConnection.Execute strQ1
                         End If
                      End If
                   End If
               Next intQ
            End If
            '-------------------
            If strKindA <> "" Then
               tmpArr = Empty
               tmpArr = Split(strKindA, ";")
               For intQ = 0 To UBound(tmpArr)
                   strQ1 = Trim(tmpArr(intQ))
                   If strQ1 <> "" Then
                      strNowNo = Mid(strQ1, InStr(strQ1, "，") + 1)
                      strConA = "select * from (" & strConB & ") where substr(ord1,1,1)<>'" & Mid(strQ1, 1, 1) & "' and (substr(ord1,2,1)='B' or substr(ord1,2,1)='0') and cfr01='" & strNowNo & "' "
                      intJ = 1
                      Set rsBD = ClsLawReadRstMsg(intJ, strConA)
                      If intJ = 0 Then
                         'Added by Lydia 2025/08/18 判斷「部門ST03和等級ST05」同時由空白修改，改成到職
                         If Trim(pOldST03 & pOldSt05) = "" Then
                            strMemo = "到職"
                            GoTo JumpToExcept
                         Else
                         'end 2025/08/15
                            If pBolTrans = True And bolConn = False Then cnnConnection.BeginTrans
                            bolConn = True
                            strMemo = "開放權限，" & IIf(pOldST03 <> pNewST03, "修改部門：" & pOldST03 & "=>" & pNewST03 & "，", "") & _
                                       IIf(pOldSt05 <> pNewST05, "修改等級：" & pOldSt05 & "=>" & pNewST05 & "，", "") & _
                                       IIf(pOldSt16 <> pNewST16, "修改工程師組別：" & pOldSt16 & "=>" & pNewST16 & "，", "") & _
                                       IIf(pOldSt70 <> pNewST70, "修改國外部小組：" & pOldSt70 & "=>" & pNewST70 & "，", "")
                            strQ1 = "INSERT INTO CUFA_Staff_Log (csl01,csl02,csl03,csl04,csl05,csl06,csl07,csl08,csl09) values (" & _
                                   "'" & strNowNo & "', '" & strUserNum & "', to_char(SYSDATE,'yyyymmdd'),to_char(sysdate,'hh24miss'), " & _
                                   "'" & pFrmName & "','" & pUserNo & "',SYS_CONTEXT('USERENV','IP_ADDRESS'),'" & pHostName & "','" & Mid(strMemo, 1, Len(strMemo) - 1) & "') "
                            cnnConnection.Execute strQ1
                         End If
                      End If
                   End If
               Next intQ
            End If
            '--end--讀取所有變更，再分析最後結果
         End If
      End If
   End If

   If strMemo <> "" Then
      strQ1 = Pub_GetSpecMan("程式管理人員")
      If strQ1 <> "" Then
         strConA = "員工編號：" & pUserNo & vbCrLf & "姓　　名：" & GetStaffName(pUserNo, True)
         strConA = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08) values(" & _
                  "'" & strUserNum & "','" & strQ1 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                  ",'有資料寫入限閱編號有權限人員異動記錄檔CUFA_Staff_Log，請確認內容。','" & ChgSQL(strConA) & "')"
         cnnConnection.Execute strConA
      End If
      If pBolTrans = True Then
          cnnConnection.CommitTrans
      End If
      bolConn = False
   End If

EXITSUB:
   If bolConn = True And pBolTrans = True And Err.Number <> 0 Then
      MsgBox "寫入失敗:" & Err.Description
      cnnConnection.RollbackTrans
   End If
   Set rsQD = Nothing
   Set rsBD = Nothing
End Sub


