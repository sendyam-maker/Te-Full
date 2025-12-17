VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAutoBatch 
   BorderStyle     =   0  '沒有框線
   Caption         =   "自動批次作業"
   ClientHeight    =   1670
   ClientLeft      =   4800
   ClientTop       =   1800
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1670
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   900
      Top             =   630
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   4980
      Top             =   2820
      _ExtentX        =   953
      _ExtentY        =   953
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   3420
      Top             =   2880
      _ExtentX        =   953
      _ExtentY        =   953
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   494
      _ExtentY        =   494
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "無畫面"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   80.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1665
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5025
   End
End
Attribute VB_Name = "frmAutoBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/12 智權人員欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
Option Explicit

'Add By Sindy 2009/06/04
Const TextPath = "\textfile\"
'Add by Morgan 2009/9/30
Dim fso As New FileSystemObject
Dim strEngDate As String '系統日英文格式
Dim Mailing As Boolean
Dim Result$, Sec%
'Modified by Morgan 2012/9/28
'Const Server$ = "exchange"
Const Server$ = "192.168.1.10"

Const Domain$ = "Taie"
Const TimeOut% = 20
Const MailBefore$ = "IMCEAEX-_O=TAIE_OU=DOMAIN_CN=RECIPIENTS_CN="
Const MailAfter$ = "@taie.com.tw"
Const MailReplace$ = "EX:/O=TAIE/OU=DOMAIN/CN=RECIPIENTS/CN="
Dim sChoice As String 'Modify by Amy 2020/01/03 從Form_Initialize搬過來


Private Sub Form_Initialize()
   Dim tmpnext As String
   
   strUserNum = "QPGMR" 'Add By Sindy 2021/12/15
   bolMailFailNoAlert = True 'Added by Morgan 2014/1/23 寄信都不要彈錯誤訊息
   g_strWriteSysLogFilePath = App.path & "\autobatchlog.log" '欲記錄Log的完整路徑及檔名 Add By Sindy 2018/5/28
   
   Set cnnConnection = New ADODB.Connection
   
   'Modified by Morgan 2019/1/11 執行VB時改可選擇連線
   'cnnConnection.ConnectionString = CnStr
   'Modified by Morgan 2022/3/23 修正執行檔Provider未換新版問題
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") > 0 Then
      sChoice = Trim(InputBox("1.【 正式 】資料庫" & vbCrLf & vbCrLf & "2.【 測試 】資料庫", "請選擇連線", "2"))
      If sChoice = "" Then End
   End If

   cnnConnection.ConnectionTimeout = 60
   cnnConnection.Provider = IIf(strProvider <> "", strProvider, cProvider)
   If sChoice = "" Then
      cnnConnection.Properties("Data Source").Value = "m51con"
   Else
      cnnConnection.Properties("Data Source").Value = IIf(sChoice = "1", "Live", "Test")
   End If
   cnnConnection.Properties("User ID").Value = UserName
   cnnConnection.Properties("Password").Value = Password
   'end 2019/1/11
   
   cnnConnection.Open
   
   Debug.Print PUB_GetDbTerminal
   
   strSrvDate(1) = Format(ServerDate)
   strSrvDate(2) = strSrvDate(1) - 19110000 'Add By Sindy 2018/2/12
   
   PUB_SetSystemVar 'Add By Sindy 2017/9/6 設定系統變數
   PUB_SetUserData 'Added by Morgan 2024/7/31 設定使用者資料庫參數
   
   
   '已發文未輸入會稿完成日， mail 給王協理
   WLog "開始  已發文未輸入會稿完成日", 1
   StrMenu
   WLog "完畢  已發文未輸入會稿完成日", 1
   
   '每月目次重編
   WLog "開始  每月目次重編", 1
   StrMenu2
   WLog "完畢  每月目次重編", 1
   
   '申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細
   WLog "開始  申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細", 1
   StrMenu3 'Modify by Amy 2021/07/28 改寫至暫存檔 原:StrMenu3改至StrMenu3_Old
   WLog "完畢  申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細", 1
   
   'Add By Sindy 2009/05/26 專利補(分割案及國內外關聯案)代表圖
   WLog "開始  補專利(分割案及國內外關聯案)代表圖(Sindy)", 1
   StrMenu4
   WLog "完畢  補專利(分割案及國內外關聯案)代表圖", 1
   
   'Add By Sindy 2009/06/11 申復案清單
   WLog "開始  申復案清單(Sindy)", 1
   StrMenu5
   WLog "完畢  申復案清單", 1
   
   'Add By Sindy 2010/10/15 每個月複製Performance專業件數上月資料至本月
   WLog "開始  複製Performance專業件數上月資料至本月(Sindy)", 1
   StrMenu9
   WLog "完畢  複製Performance專業件數上月資料至本月", 1
   
   '2012/8/30 ADD BY SONIA PS,CPS非台灣案一年內無進度上可結餘日期
   WLog "開始  PS,CPS非台灣案一年內無進度上可結餘日期", 1
   StrMenu10
   WLog "完畢  PS,CPS非台灣案一年內無進度上可結餘日期", 1
   '2012/8/30 END
   
   'Add By Amy 2013/05/03 刪除三年內無往來記錄之國內潛在客戶(排除D01及P12)
   WLog "開始  刪除三年內無往來記錄之國內潛在客戶(Amy)", 1
   StrMenu11
   WLog "完畢  刪除三年內無往來記錄之國內潛在客戶", 1
   '2013/05/03 End
   
   'Added by Lydia 2015/04/09 每月顧問到期通知
   WLog "開始  每月顧問到期通知(Lydia)", 1
   StrMenu13
   WLog "完畢  每月顧問到期通知", 1
   'end 2015/04/09

  'Add by Amy   2016/06/15 上個月個人客戶資料修改通知
   WLog "開始  上個月個人客戶資料修改通知(Amy)", 1
   StrMenu14
   WLog "完畢  上個月個人客戶資料修改通知", 1
   'end 2016/06/15
    
   'Added by Lydia 2016/10/05 CFT可辦期限管制表
   WLog "開始  CFT可辦期限管制表(Lydia)", 1
   StrMenu15
   WLog "完畢  CFT可辦期限管制表", 1
   'end 2016/10/05
   
   'Added by Lydia 2017/02/22 國內新客戶清單
   WLog "開始  國內新客戶清單(Lydia)", 1
   StrMenu16
   WLog "完畢  國內新客戶清單", 1
   'end 2017/02/22
   
   'Added by Lydia 2017/02/22 國內專利收文未發文明細表
   WLog "開始  國內專利收文未發文明細表(Lydia)", 1
   StrMenu17
   WLog "完畢  國內專利收文未發文明細表", 1
   'end 2017/02/22
   
   'Added by Lydia 2017/06/27 FCT案超過2個月尚未列印核准定稿通知
   'Mark by Lydia 2023/08/01 取消管制: 因為現在FCT所有定稿(催延展除外)在產生於定稿作業維護同時，會另將定稿儲存於FCT-workflow
                                       '所以程序人員都在FCT -workflow做修改或列印的動作, 不會每件都從定稿作業維護列印定稿了
   'WLog "開始  FCT案超過2個月尚未列印核准定稿通知(Lydia)", 1
   'StrMenu18
   'WLog "完畢  FCT案超過2個月尚未列印核准定稿通知", 1
   ''end 2017/06/27
   'end 2023/08/01
   
   'Add By Sindy 2018/2/12 每年3/1及10/1日寄發勞工健康檢查通知函
   WLog "開始  每年3/1及10/1日寄發勞工健康檢查通知函(Sindy)", 1
   StrMenu19
   WLog "完畢  每年3/1及10/1日寄發勞工健康檢查通知函", 1
   '2018/2/12 END
   
   'Added by Morgan 2018/5/21
   WLog "開始  FMP案已發文未輸完稿日報表(Morgan)", 1
   StrMenu20
   WLog "完畢  FMP案已發文未輸完稿日報表", 1
   'end 2018/5/21
   
   'Added by Lydia 2018/08/22
   'Mark by Lydia 2025/07/08 改成每日批次frmAutoBatchDay.StrMenu143
   'WLog "開始  應收帳款逾付款週期管制表(Lydia)", 1
   'StrMenu21
   'WLog "完畢  應收帳款逾付款週期管制表", 1
   ''end 2018/08/22
   'end 2025/07/08
   
   'Added by Morgan 2018/11/30
   WLog "開始  刪除3個月前之商標註冊費繳費單pdf檔(Morgan)", 1
   StrMenu22
   WLog "完畢  刪除3個月前之商標註冊費繳費單pdf檔", 1
   'end 2018/11/30
   
   'Added by Morgan 2019/1/22
   WLog "開始  優先權期限資料提供(Morgan)", 1
   StrMenu23
   WLog "完畢  優先權期限資料提供", 1
   'end 2019/1/22
   
   'Added by Morgan 2019/6/17
   WLog "開始  刪除3個月前上傳之台一網站資料夾[網頁提供國內專利公報資訊](Morgan)", 1
   StrMenu24
   WLog "完畢  刪除3個月前上傳之台一網站資料夾[網頁提供國內專利公報資訊]", 1
   'end 2019/6/17
   
   'Added by Lydia 2019/07/31
   WLog "開始  FCP寄證書後年費不續辦(Lydia)", 1
   StrMenu25
   WLog "完畢  FCP寄證書後年費不續辦", 1
   'end 2019/07/31
   
   'Add by Amy 2020/01/03 刪除M51-APP 電子發票 部分資料夾三個月前資料
   WLog "開始  刪除電子發票 部分資料夾三個月前資料(Amy)", 1
   StrMenu26
   WLog "完畢  刪除電子發票 部分資料夾三個月前資料", 1
   
   'Added by Lydia 2020/09/01 結餘單流水號檢查: 未結算逾2個月結餘單寄給財務處總帳人員
   WLog "開始  結餘單流水號檢查(Lydia)", 1
   StrMenu27
   WLog "完畢  結餘單流水號檢查", 1
   'end 2020/09/01
   
   'Add by Sindy 2021/12/15
   WLog "開始  外商FCT案註冊已滿三年案件管制表(Sindy)", 1
   StrMenu28
   WLog "完畢  外商FCT案註冊已滿三年案件管制表", 1
   '2021/12/15 END
   
   'Add by Sindy 2024/3/14
   WLog "開始  外商收文未發文清單(Sindy)", 1
   StrMenu32
   WLog "完畢  外商收文未發文清單", 1
   WLog "開始  外商催審表(Sindy)", 1
   StrMenu33
   WLog "完畢  外商催審表", 1
   WLog "開始  外商FCT延展管制表(Sindy)", 1
   StrMenu34
   WLog "完畢  外商FCT延展管制表", 1
   '2024/3/14 END
   
   'Add by Sindy 2024/6/11
   WLog "開始  針對客戶編號ID仍為66666666(8個6)的通知(Sindy)", 1
   StrMenu35
   WLog "完畢  針對客戶編號ID仍為66666666(8個6)的通知", 1
   '2024/6/11 END
   
   'Add by Amy 2023/08/16 智權部新客戶身份證號統一編號與舊客戶相同清單,寄給「全所智權部主管」
   WLog "開始  智權部新客戶身份證號統一編號與舊客戶相同清單(Amy)", 1
   StrMenu29
   WLog "完畢  智權部新客戶身份證號統一編號與舊客戶相同清單", 1
   
   'Added by Lydia 2023/09/14 自動轉入待活化客戶：每年之1月1日及7月1日，將已超過12年未收文但尚未設為活化客戶的客戶資訊，整合進入活化區域。
   If strSrvDate(1) >= "20240701" Then 'Memo by Lydia 2023/10/05 下次調整時間為113.7.1
      If Val(Mid(strSrvDate(1), 5, 2)) = 1 Or Val(Mid(strSrvDate(1), 5, 2)) = 7 Then
         WLog "開始  自動轉入待活化客戶(Lydia)", 1
         StrMenu30
         WLog "完畢  自動轉入待活化客戶", 1
      End If
   End If
   'end 2023/09/14

   'Added by Lydia 2024/01/15
   WLog "開始  顧問聘任期間檢查Email設定(Lydia)", 1
   StrMenu31
   WLog "完畢  顧問聘任期間檢查Email設定", 1
   'end 2024/01/15
   
   'Add By Sindy 2020/9/3
   '每月第1個工作天執行
   'If strSrvDate(1) = GetMonthStdDay(Left(strSrvDate(1), 6)) Then 'Removed by Morgan 2024/7/31 有跟Sindy確認每月1號跑不必限制工作天
      strDate = DBDATE(DateAdd("yyyy", -3, Format(strSrvDate(1), "####/##/##")))
      WLog "開始  刪除電子收據紀錄檔(Sindy), 繳費時間<=" & strDate, 1
      strSql = "delete from ereceipt where to_char(er04,'yyyymmdd')<=" & strDate
      cnnConnection.Execute strSql, intI
      WLog "完畢  刪除電子收據紀錄檔" & intI & "筆", 1
   'End If
   '2020/9/3 END
   
   'Added by Morgan 2024/7/31
   WLog "開始  批次新增上個月天災不給薪假單(Morgan)", 1
   StrMenu36
   WLog "完畢  批次新增上個月天災不給薪假單", 1
   'end 2024/7/31
   
   'Added by Morgan 2024/8/19
   WLog "開始  BASF待請款之案件(Morgan)", 1
   StrMenu37
   WLog "完畢  BASF待請款之案件", 1
   'end 2024/8/19
   
   'Added by Sindy 2025/6/20 每月列出電郵主旨中有主旨標籤的電郵資料(以連結新案變化以分析電郵開拓成效)
   WLog "開始  每月列出電郵主旨中有主旨標籤的電郵資料(Sindy)", 1
   StrMenu38
   WLog "完畢  每月列出電郵主旨中有主旨標籤的電郵資料", 1
   'end 2025/6/20
   
   'Added by Sindy 2025/6/30 財務信箱定期檢查
   WLog "開始  財務信箱定期檢查(Sindy)", 1
   StrMenu39
   WLog "完畢  財務信箱定期檢查", 1
   'end 2025/6/30
      
   'Added by Lydia 2025/09/16
   WLog "開始  批次閉卷T台灣案陳述意見書(Lydia)", 1
   StrMenu41
   WLog "完畢  批次閉卷T台灣案陳述意見書", 1
   'end 2025/09/16
   
   'Mark by Amy 2015/0810 搬至AutoBatchDay 改每月1日及16日發
'   'Add by Amy 2015/05/26 FCT,T,TB,TC,TD,TF,TM,TR,TS,TT每月自動發催審表給承辦人
'   WLog "開始  每月自動發催審表給承辦人", 1
'   StrMenu14
'   WLog "完畢  每月自動發催審表給承辦人", 1
'   'end 2015/05/26
   
'   'Add By Sindy 2014/6/30 複檢卷宗區檔案狀況
'   WLog "開始  複檢卷宗區檔案狀況", 1
'   StrMenu12
'   WLog "完畢  複檢卷宗區檔案狀況", 1
'   '2014/6/30 End
   
   ''Add By Sindy 2009/08/18 延展案於期滿六個月前一個工作天發mail提醒承辦人
   ''If ChkWorkDay(Format(Now, "YYYYMMDD")) = True Then
   '   WLog "開始   FCT延展案於期滿六個月的前一個工作天發mail提醒承辦人", 1
   '   StrMenu6
   '   WLog "完畢   FCT延展案於期滿六個月的前一個工作天發mail提醒承辦人", 1
   ''End If
   ''2009/08/18 END
   
   
   '2010/7/5 cancel by sonia DAVID說現IPO不限制時間故不必再發函
   ''Add by Morgan 2009/9/30
   'If InStr("0101,0401,0701,1001", Mid(strSrvDate(1), 5)) > 0 Then
   '   WLog "開始  FCP終止辦理詢問函及清單", 1
   '   'Modify by Morgan 2010/3/30 改新規則
   '   'StrMenu7
   '   StrMenu8
   '   WLog "完畢  FCP終止辦理詢問函及清單", 1
   'End If
   '2010/7/5 END
   
   'Modify by Morgan 2008/6/2 不發Mail改寫Log
   'tmpNext = "準備發 mail..."
   'MAPISession1.LogonUI = False
   'MAPISession1.UserName = "administrator"
   'tmpNext = "準備登入郵件伺服器..."
   'MAPISession1.SignOn
   'tmpNext = "登入郵件伺服器..."
   'MAPIMessages1.SessionID = MAPISession1.SessionID
   'MAPIMessages1.MsgIndex = -1
   'tmpNext = "建立郵件..."
   'MAPIMessages1.Compose
   'MAPIMessages1.MsgSubject = "每月批次作業---------" & Format(Now, "YYYY/MM/DD")
   'MAPIMessages1.MsgNoteText = "Dear Nickc：" & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的批次已經完成    " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "                                                        自動批次中心"
   'MAPIMessages1.RecipIndex = 0
   'MAPIMessages1.RecipDisplayName = "93013"
   'MAPIMessages1.ResolveName
   'tmpNext = "準備存入郵件..."
   'MAPIMessages1.Send
   'tmpNext = "發信..."
   'MAPISession1.SignOff
   'tmpNext = "登出..."
   WLog "民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的批次已經完成" & vbCrLf & vbCrLf, 1  'Modified by Lydia 2024/07/01 加上跳行符號
   'end 2008/6/2
   
   'Add By Sindy 2011/9/20
   If cnnConnection.State <> adStateClosed Then cnnConnection.Close
   Set cnnConnection = Nothing
   '2011/9/20 End
   
   End
End Sub

'已發文未輸入會稿完成日， mail 給王協理
Sub StrMenu()
Dim ff As Integer
Dim i As Integer
Dim A01 As String
Dim A02 As String
Dim A03 As String
Dim A04 As String
Dim A05 As String
Dim iCount As Long
Dim TempFileName As String
Dim tmpnext As String

 On Error GoTo DebugErr
   'tmpnext = "準備連資料庫..."
   'Set cnnConnection = New ADODB.Connection
   'cnnConnection.ConnectionString = CnStr
   'cnnConnection.Open
   'edit by nick 2004/12/07  加條件 2005/12/2再加413,429 2005/12/27再加505,506
   'StrSQL = "select rpad(nvl(st02,' '),10,' '),rpad(" & SQLDate("cp27") & ",12,' '),rpad(cp01||'-'||cp02||'-'||cp03||'-'||cp04,18,' '),rpad(nvl(pa05,nvl(pa06,pa07)),52,' '),rpad(nvl(cpm03,cpm04),42,' ') " & _
                  " from caseprogress,engineerprogress,patent,casepropertymap,staff " & _
                  " where cp01 in ('P','CFP') and CP10<>'1101' and cp27>=(select substr(to_char(max(wd01)),1,6)||'01' from workday where wd01<to_char(sysdate, 'YYYYMMDD') )" & _
                  " and cp27<=(select substr(to_char(max(wd01)),1,6)||'31' from workday where wd01<to_char(sysdate, 'YYYYMMDD') )" & _
                  " and cp09=ep02(+) and ep08 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and st03<>'P12' and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C'))  " & _
                  " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  order by cp14,2 "
   strSql = "select rpad(nvl(st02,' '),10,' '),rpad(" & SQLDate("cp27") & ",12,' '),rpad(cp01||'-'||cp02||'-'||cp03||'-'||cp04,18,' '),rpad(nvl(pa05,nvl(pa06,pa07)),52,' '),rpad(decode(pa09,'000',cpm03,cpm04),42,' ') " & _
                  " from caseprogress,engineerprogress,patent,casepropertymap,staff " & _
                  " where cp01 in ('P','CFP') and CP10<>'1101' and cp27>=(select substr(to_char(max(wd01)),1,6)||'01' from workday where wd01<to_char(sysdate, 'YYYYMMDD') )" & _
                  " and cp27<=(select substr(to_char(max(wd01)),1,6)||'31' from workday where wd01<to_char(sysdate, 'YYYYMMDD') )" & _
                  " and cp09=ep02(+) and ep08 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C'))  " & _
                  " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506') order by cp14,cp27 "
   
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   tmpnext = "已取得資料..."
   With Rs
        ff = FreeFile
        'Modify By Sindy 2009/06/04
        'TempFileName = "c:\" & Format(Now, "YYYYMMDDhhmmss") & ".txt"
        TempFileName = App.path & TextPath & Format(Now, "YYYYMMDDhhmmss") & ".txt"
        '2009/06/04 End
        Open TempFileName For Output As ff

        tmpnext = "開檔中..."
        'Added by Lydia 2016/10/19 加報表抬頭
        Print #ff, Space(30) & "已發文未輸入會稿完成日"
        Print #ff, ""
        'end 2016/10/19
        Print #ff, "承辦人      發文日       本所案號　　       案件名稱　　　　                                　　 案件性質　 "
        Print #ff, "=========== ============ ================== ==================================================== ================="
        
        If .RecordCount > 0 And .RecordCount <> 0 Then
               iCount = 1
               .MoveFirst
               tmpnext = "寫檔..."
               Do While Not .EOF
                    A01 = "" & (.Fields(0).Value)
                    A02 = "" & (.Fields(1).Value)
                    A03 = "" & (.Fields(2).Value)
                    A04 = "" & (.Fields(3).Value)
                    A05 = "" & (.Fields(4).Value)
                    Print #ff, A01 & "  " & A02 & " " & A03 & " " & A04 & " " & A05
                    .MoveNext
               Loop
        Else
             TempFileName = ""
        End If
   End With
   Close ff
   tmpnext = "關檔..."
   
   'Modify By Sindy 2011/10/27
   'modify by sonia 2016/10/3 取消職稱
   'SendMAPIMail "71011", "已發文未輸入會稿完成日---------" & Format(Now, "YYYY/MM/DD"), "Dear 王協理：" & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的    已發文未輸入會稿完成日   " & IIf(TempFileName = "", "資料庫找不到資料", "資料如附件") & "！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心", TempFileName
   'Added by Lydia 2023/04/24 修改王副總退休之相關控制
   'Modified by Morgan 2025/2/19
   'strExc(0) ="73022"
   pub_PMan = Pub_GetSpecMan("專利處特定編號") 'Added by Morgan 2025/3/3
   strExc(0) = pub_PMan
   'end 2025/2/19
   
   'Modified by Lydia 2023/04/24 "71011" => strExc(0)
   SendMAPIMail strExc(0), "已發文未輸入會稿完成日---------" & Format(Now, "YYYY/MM/DD"), "Dear Sirs," & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的    已發文未輸入會稿完成日   " & IIf(TempFileName = "", "資料庫找不到資料", "資料如附件") & "！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心", TempFileName
   
   ''發 mail
   'tmpnext = "準備發 mail..."
   'MAPISession1.LogonUI = False
   'MAPISession1.UserName = "administrator"
   'tmpnext = "準備登入郵件伺服器..."
   'MAPISession1.SignOn
   'tmpnext = "登入郵件伺服器..."
   'MAPIMessages1.SessionID = MAPISession1.SessionID
   'MAPIMessages1.MsgIndex = -1
   'tmpnext = "建立郵件..."
   'MAPIMessages1.Compose
   'MAPIMessages1.MsgSubject = "已發文未輸入會稿完成日---------" & Format(Now, "YYYY/MM/DD")
   'MAPIMessages1.MsgNoteText = "Dear 王協理：" & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的    已發文未輸入會稿完成日   " & IIf(TempFileName = "", "資料庫找不到資料", "資料如附件") & "！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
   'If TempFileName <> "" Then
   '    MAPIMessages1.AttachmentPosition = 120
   '    MAPIMessages1.AttachmentPathName = TempFileName
   'End If
   'MAPIMessages1.RecipIndex = 0
   'MAPIMessages1.RecipDisplayName = "71011"
   'MAPIMessages1.ResolveName
   'tmpnext = "準備存入郵件..."
   'MAPIMessages1.Send
   'tmpnext = "發信..."
   'MAPISession1.SignOff
   'tmpnext = "登出..."
   If TempFileName <> "" Then
       Kill TempFileName
   End If
   tmpnext = "清除暫存檔..."
   'Shell "net send /domain:taient2 '每月1日自動郵件資料已送出，請清除郵件備份' ", vbNormalNoFocus
   Set Rs = Nothing
   'Set cnnConnection = Nothing
   Exit Sub
DebugErr:
        MsgBox tmpnext & " " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmAutoBatch = Nothing
End Sub

Sub StrMenu2()
Dim i As Integer, k As Integer, strTemp3 As String, s As Integer
Dim tmpnext As String

   On Error GoTo CheckingErr
   'tmpnext = "準備連資料庫..."
   'Set cnnConnection = New ADODB.Connection
   'cnnConnection.ConnectionString = CnStr
   'cnnConnection.Open
   strSrvDate(1) = Format(ServerDate)
   'Modified by Lydia 2016/09/14
   'strSql = "SELECT EP01,EP02,ep05,EP10,CP05 FROM ENGINEERPROGRESS,CASEPROGRESS WHERE cp09=ep02(+) AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP27<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp57 is null ) or (CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP57<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp27 is null))) AND EP05 IS NOT NULL and cp05>=19980101 "
   strSql = "SELECT EP01,EP02,ep05,EP10,CP05 FROM ENGINEERPROGRESS,CASEPROGRESS WHERE cp09=ep02(+) " & _
            "AND ((CP158=0 AND CP159=0) " & _
            "OR ((CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP27<=" & Mid(strSrvDate(1), 1, 6) & "31 and CP159=0)" & _
            "OR (CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP57<=" & Mid(strSrvDate(1), 1, 6) & "31 and CP158=0))" & _
            ") AND EP05 IS NOT NULL and cp05>=19980101 "
   strSql = strSql & " ORDER BY EP05,ep01,cp05 "
   CheckOC
   i = 0
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           .MoveFirst
           strTemp3 = CheckStr(.Fields(2))
           i = 0
           Do While .EOF = False
               If strTemp3 <> CheckStr(.Fields(2)) Then
                   i = 1
                   strTemp3 = CheckStr(.Fields(2))
               Else
                   i = i + 1
               End If
                cnnConnection.Execute "UPDATE ENGINEERPROGRESS SET EP01=" & i & " WHERE EP02='" & CheckStr(.Fields(1)) & "' "
               .MoveNext
               DoEvents
           Loop
       End If
   End With
   CheckOC
        Exit Sub
CheckingErr:
   'Modify By Sindy 2011/10/27
   'Modify By Sindy 2024/3/18 改抓 程式管理人員
   SendMAPIMail Pub_GetSpecMan("程式管理人員"), "每月目次重編---------" & Format(Now, "YYYY/MM/DD"), "Dear Sirs：" & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的目次重編發生錯誤(" & Err.Description & ")" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
      
'      tmpnext = "準備發 mail..."
'      MAPISession1.LogonUI = False
'      MAPISession1.UserName = "administrator"
'      tmpnext = "準備登入郵件伺服器..."
'      MAPISession1.SignOn
'      tmpnext = "登入郵件伺服器..."
'      MAPIMessages1.SessionID = MAPISession1.SessionID
'      MAPIMessages1.MsgIndex = -1
'      tmpnext = "建立郵件..."
'      MAPIMessages1.Compose
'      MAPIMessages1.MsgSubject = "每月目次重編---------" & Format(Now, "YYYY/MM/DD")
'      MAPIMessages1.MsgNoteText = "Dear Sirs：" & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的目次重編發生錯誤(" & Err.Description & ")" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
'      MAPIMessages1.RecipIndex = 0
'      MAPIMessages1.RecipDisplayName = Pub_GetSpecMan("程式管理人員")
'      MAPIMessages1.ResolveName
'      tmpnext = "準備存入郵件..."
'      MAPIMessages1.Send
'      tmpnext = "發信..."
'      MAPISession1.SignOff
'      tmpnext = "登出..."
End Sub

'add by nick 2005/01/07
'申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細
Sub StrMenu3_Old()
Dim ff As Integer
Dim i As Integer
Dim A01 As String
Dim A02 As String
Dim A03 As String
Dim A04 As String
Dim A05 As String
Dim A06 As String
Dim A07 As String
Dim A08 As String
Dim A09 As String
Dim SeekA02 As String
Dim iCount As Long
Dim TempFileName As String
Dim TempFileNameT As String
Dim TempFileNameFCT As String
Dim ArrMail As Variant
Dim tmpnext As String
Dim StrMenu3I As Integer
Dim strSubject As String, strContent As String 'Add By Sindy 2011/10/27

On Error GoTo DebugErr
   'tmpnext = "準備連資料庫..."
   'Set cnnConnection = New ADODB.Connection
   'cnnConnection.ConnectionString = CnStr
   'cnnConnection.Open
   'edit by nick 2005/01/31 有暫緩審理(310,1401)不列，6 個月內有來文(C 類 6個月內收文)或覆出(A,B 類 6 個月內發文但不續辦 703 和閉卷 704 不算)
   'StrSQL = "select substr(nvl(S1.st03,s2.st03),1,1),rpad(nvl(S1.st02,' '),6,' '),rpad(tm01||'-'||tm02||'-'||tm03||'-'||tm04,15,' '),rpad(tm05||tm06||tm07,20,' '),rpad('申請',10,' ')," & _
                   "rpad(DECODE(cp27,'','',SUBSTR(cp27,1,4)-1911||'/'||SUBSTR(cp27,5,2)||'/'||SUBSTR(cp27,7,2)),10,' ')," & _
                   "rpad(tm23,10,' '),rpad(cu04,12,' ') from ( select tm01,tm02,tm03,tm04,tm05,tm06,tm07,tm23 " & _
                   " From trademark where tm11>=20000000 and tm01 in ('T','FCT') and tm10='000' and tm29 is null " & _
                   " and tm16 is null and tm28='1'  and tm11 <=to_number(to_char(add_months(sysdate,-20),'yyyymmdd')) " & _
                   " and not exists(select * from caseprogress C1 where C1.cp01=tm01 and C1.cp02=tm02  and C1.cp03=tm03 and C1.cp04=tm04 " & _
                   " and ( (C1.cp10<>'1101' and C1.cp09>'C') or (C1.cp10='306' and C1.cp27 is not null and exists (select * from caseprogress C3 where c1.cp43=C3.cp09 and C3.cp10='101') )) " & _
                   " and C1.cp05>=20000000)) ,caseprogress C2,staff S1,customer,staff s2  where tm01=C2.cp01(+) and tm02=C2.cp02(+) and tm03=C2.cp03(+) and tm04=C2.cp04(+) and C2.cp10='101' and C2.cp14=s1.st01(+) and C2.cp65=s2.st01(+)  and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & _
                   " union all select substr(nvl(S1.st03,s2.st03),1,1),rpad(nvl(S1.st02,' '),6,' '),rpad(tm01||'-'||tm02||'-'||tm03||'-'||tm04,15,' '),rpad(tm05||tm06||tm07,20,' '),rpad('核駁前先行',10,' '), " & _
                   " rpad(DECODE(cp27,'','',SUBSTR(cp27,1,4)-1911||'/'||SUBSTR(cp27,5,2)||'/'||SUBSTR(cp27,7,2)),10,' '),rpad(tm23,10,' '),rpad(cu04,12,' ') " & _
                   " From trademark, caseprogress c2, staff s1, customer,staff s2  where c2.cp05>=20000000 and tm01 in ('T','FCT') and tm10='000' and tm29 is null and tm16 is null and tm01=c2.cp01(+)  and tm02=c2.cp02(+) and tm03=c2.cp03(+) and tm04=c2.cp04(+) and c2.cp10='1202' " & _
                   " and not exists(select * from caseprogress C1 where C1.cp01=tm01 and C1.cp02=tm02  and C1.cp03=tm03 and C1.cp04=tm04 " & _
                   " and ( (C1.cp10='306' and C1.cp27 is not null and not exists (select * from caseprogress C3 where c1.cp43=C3.cp09(+) and C3.cp10='101') )) " & _
                   " and C1.cp05>=20000000) and c2.cp14=s1.st01(+) and c2.cp65=s2.st01(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+)  and c2.cp05 <=to_number(to_char(add_months(sysdate,-12),'yyyymmdd') ) order by 1,2,5,6,3 "
   'edit by nickc 2005/06/03 加欄位
   'StrSQL = "select substr(nvl(S1.st03,s2.st03),1,1),rpad(nvl(S1.st02,' '),6,' '),rpad(tm01||'-'||tm02||'-'||tm03||'-'||tm04,15,' '),rpad(tm05||tm06||tm07,20,' '),rpad('申請',10,' '), " & _
            " rpad(DECODE(cp27,'','',SUBSTR(cp27,1,4)-1911||'/'||SUBSTR(cp27,5,2)||'/'||SUBSTR(cp27,7,2)),10,' '),rpad(tm23,10,' '),rpad(cu04,12,' ') from ( select tm01,tm02,tm03,tm04,tm05,tm06,tm07,tm23 " & _
            " From trademark where tm11>=20000000 and tm01 in ('T','FCT') and tm10='000' and tm29 is null and tm16 is null and tm28='1'  and tm11 <=to_number(to_char(add_months(sysdate,-20),'yyyymmdd')) " & _
            " and not exists(select * from caseprogress C1 where C1.cp01=tm01 and C1.cp02=tm02  and C1.cp03=tm03 and C1.cp04=tm04 and ((c1.cp10 not in ('703','704') and c1.cp09<'C' and c1.cp27>to_number(to_char(add_months(sysdate,-6),'yyyymmdd'))) " & _
            " or (c1.cp09>'C' and c1.cp05>to_number(to_char(add_months(sysdate,-6),'yyyymmdd'))) or (C1.cp10<>'1101' and C1.cp09>'C') or (c1.cp10='310' and c1.cp27>to_number(to_char(add_months(sysdate,-24),'yyyymmdd'))) " & _
            " or (c1.cp10='1401' and c1.cp05>to_number(to_char(add_months(sysdate,-24),'yyyymmdd'))) or (C1.cp10='306' and C1.cp27 is not null and exists (select * from caseprogress C3 where c1.cp43=C3.cp09 and C3.cp10='101') ))  and C1.cp05>=20000000)) " & _
            " ,caseprogress C2,staff S1,customer,staff s2  where tm01=C2.cp01(+) and tm02=C2.cp02(+) and tm03=C2.cp03(+) and tm04=C2.cp04(+) and C2.cp10='101' and C2.cp14=s1.st01(+) and C2.cp65=s2.st01(+)  and substr(tm23,1,8)=cu01(+) " & _
            " and substr(tm23,9,1)=cu02(+) " & _
            " union all select substr(nvl(S1.st03,s2.st03),1,1),rpad(nvl(S1.st02,' '),6,' '), rpad(tm01||'-'||tm02||'-'||tm03||'-'||tm04,15,' '),rpad(tm05||tm06||tm07,20,' '),rpad('核駁前先行',10,' '), " & _
            " rpad(DECODE(cp27,'','',SUBSTR(cp27,1,4)-1911||'/'||SUBSTR(cp27,5,2)||'/'||SUBSTR(cp27,7,2)),10,' '), rpad(tm23,10,' '),rpad(cu04,12,' ')  From trademark, caseprogress c2, staff s1, customer,staff s2 " & _
            " where c2.cp05>=20000000 and tm01 in ('T','FCT') and tm10='000' and tm29 is null and tm16 is null and tm01=c2.cp01(+)  and tm02=c2.cp02(+) and tm03=c2.cp03(+) and tm04=c2.cp04(+) and c2.cp10='1202' " & _
            " and not exists(select * from caseprogress C1 where C1.cp01=tm01 and C1.cp02=tm02  and C1.cp03=tm03 and C1.cp04=tm04 and ((c1.cp10 not in ('703','704') and c1.cp09<'C' and c1.cp27>to_number(to_char(add_months(sysdate,-6),'yyyymmdd'))) " & _
            " or (c1.cp09>'C' and c1.cp05>to_number(to_char(add_months(sysdate,-6),'yyyymmdd'))) or (c1.cp10='310' and c1.cp27>to_number(to_char(add_months(sysdate,-24),'yyyymmdd'))) " & _
            " or (c1.cp10='1401' and c1.cp05>to_number(to_char(add_months(sysdate,-24),'yyyymmdd'))) or (C1.cp10='306' and C1.cp27 is not null and not exists (select * from caseprogress C3 where c1.cp43=C3.cp09(+) and C3.cp10='101') )) " & _
            " and C1.cp05>=20000000) and c2.cp14=s1.st01(+) and c2.cp65=s2.st01(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+)  and c2.cp05 <=to_number(to_char(add_months(sysdate,-12),'yyyymmdd') ) " & _
            " order by 1,2,5,6,3 "
   'Modify By Sindy 2011/12/26 +tm01,tm02,tm03,tm04
   '2012/5/2 modify by sonia 核駁前先行通知加 or (c1.cp10='1002' and c1.cp64='申請駁回') T-152552
   'Modify By Sindy 2012/12/4 若為外商substr(CP12,1,1)='F'者抓智權人員,其他才抓承辦人
   'Modify By Sindy 2015/9/17 再判斷該案號之101申請程序,若已有1724通知已轉他所的來函,則不出現在清單上:and not exists (select C4.cp09 from caseprogress C4,caseprogress C5 where C4.cp01=tm01 and C4.cp02=tm02 and C4.cp03=tm03 and C4.cp04=tm04 and C4.cp10='101' and C4.cp09=C5.cp43(+) and C5.cp10='1724')
   'Modified by Lydia 2016/07/18 T台灣案內商審查報告輸入為核駁前先行通知 時,會更新該案號申請101那一道的下一程序催審305期限(NP06 IS NULL) 為法定期限+8個月(2015/3/21上線請作單)
   'strSql = "select substr(nvl(S1.st03,s2.st03),1,1),decode(substr(nvl(S1.st03,s2.st03),1,1),'F',rpad(nvl(s3.st02,' '),6,' '),rpad(nvl(S1.st02,' '),6,' ')),rpad(nvl(nvl(tm15,tm12),' '),20,' '),rpad(tm01||'-'||tm02||'-'||tm03||'-'||tm04,15,' '),rpad(tm05||tm06||tm07,20,' '),rpad('申請',10,' '), " & _
            " rpad(DECODE(cp27,'','',SUBSTR(cp27,1,4)-1911||'/'||SUBSTR(cp27,5,2)||'/'||SUBSTR(cp27,7,2)),10,' '),rpad(tm23,10,' '),rpad(cu04,12,' '),nvl(cp27,0),tm01,tm02,tm03,tm04 from ( select tm01,tm02,tm03,tm04,tm05,tm06,tm07,tm23,tm15,tm12 " & _
            " From trademark where tm11>=20000000 and tm01 in ('T','FCT') and tm10='000' and tm29 is null and tm16 is null and tm28='1'  and tm11 <=to_number(to_char(add_months(sysdate,-20),'yyyymmdd')) " & _
            " and not exists(select * from caseprogress C1 where C1.cp01=tm01 and C1.cp02=tm02  and C1.cp03=tm03 and C1.cp04=tm04 and ((c1.cp10 not in ('703','704') and c1.cp09<'C' and c1.cp27>to_number(to_char(add_months(sysdate,-6),'yyyymmdd'))) " & _
            " or (c1.cp09>'C' and c1.cp05>to_number(to_char(add_months(sysdate,-6),'yyyymmdd'))) or (C1.cp10<>'1101' and C1.cp09>'C') or (c1.cp10='310' and c1.cp27>to_number(to_char(add_months(sysdate,-24),'yyyymmdd'))) " & _
            " or (c1.cp10='1401' and c1.cp05>to_number(to_char(add_months(sysdate,-24),'yyyymmdd'))) or (C1.cp10='306' and C1.cp27 is not null and exists (select * from caseprogress C3 where c1.cp43=C3.cp09 and C3.cp10='101') ))  and C1.cp05>=20000000)) " & _
            " ,caseprogress C2,staff S1,customer,staff s2,staff s3 where tm01=C2.cp01(+) and tm02=C2.cp02(+) and tm03=C2.cp03(+) and tm04=C2.cp04(+) and C2.cp10='101' and C2.cp14=s1.st01(+) and C2.cp65=s2.st01(+) and C2.cp13=s3.st01(+) and substr(tm23,1,8)=cu01(+) " & _
            " and substr(tm23,9,1)=cu02(+) " & _
            " and not exists (select C4.cp09 from caseprogress C4,caseprogress C5 where C4.cp01=tm01 and C4.cp02=tm02 and C4.cp03=tm03 and C4.cp04=tm04 and C4.cp10='101' and C4.cp09=C5.cp43(+) and C5.cp10='1724')" & _
            " union all select substr(nvl(S1.st03,s2.st03),1,1),decode(substr(nvl(S1.st03,s2.st03),1,1),'F',rpad(nvl(s3.st02,' '),6,' '),rpad(nvl(S1.st02,' '),6,' ')),rpad(nvl(nvl(tm15,tm12),' '),20,' '),rpad(tm01||'-'||tm02||'-'||tm03||'-'||tm04,15,' '),rpad(tm05||tm06||tm07,20,' '),rpad('核駁前先行',10,' '), " & _
            " rpad(DECODE(cp27,'','',SUBSTR(cp27,1,4)-1911||'/'||SUBSTR(cp27,5,2)||'/'||SUBSTR(cp27,7,2)),10,' '), rpad(tm23,10,' '),rpad(cu04,12,' '),nvl(cp27,0),tm01,tm02,tm03,tm04 From trademark, caseprogress c2, staff s1, customer,staff s2,staff s3 " & _
            " where c2.cp05>=20000000 and tm01 in ('T','FCT') and tm10='000' and tm29 is null and tm16 is null and tm01=c2.cp01(+)  and tm02=c2.cp02(+) and tm03=c2.cp03(+) and tm04=c2.cp04(+) and c2.cp10='1202' " & _
            " and not exists(select * from caseprogress C1 where C1.cp01=tm01 and C1.cp02=tm02  and C1.cp03=tm03 and C1.cp04=tm04 and ((c1.cp10 not in ('703','704') and c1.cp09<'C' and c1.cp27>to_number(to_char(add_months(sysdate,-6),'yyyymmdd'))) " & _
            " or (c1.cp09>'C' and c1.cp05>to_number(to_char(add_months(sysdate,-6),'yyyymmdd'))) or (c1.cp10='310' and c1.cp27>to_number(to_char(add_months(sysdate,-24),'yyyymmdd'))) " & _
            " or (c1.cp10='1401' and c1.cp05>to_number(to_char(add_months(sysdate,-24),'yyyymmdd'))) or (c1.cp10='1002' and c1.cp64='申請駁回') or (C1.cp10='306' and C1.cp27 is not null and not exists (select * from caseprogress C3 where c1.cp43=C3.cp09(+) and C3.cp10='101') )) " & _
            " and C1.cp05>=20000000) and c2.cp14=s1.st01(+) and c2.cp65=s2.st01(+) and C2.cp13=s3.st01(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and c2.cp05 <=to_number(to_char(add_months(sysdate,-12),'yyyymmdd') ) " & _
            " and not exists (select C4.cp09 from caseprogress C4,caseprogress C5 where C4.cp01=tm01 and C4.cp02=tm02 and C4.cp03=tm03 and C4.cp04=tm04 and C4.cp10='101' and C4.cp09=C5.cp43(+) and C5.cp10='1724')" & _
            " order by 1,2,6,10,4 "
   strSql = "select substr(nvl(S1.st03,s2.st03),1,1),decode(substr(nvl(S1.st03,s2.st03),1,1),'F',rpad(nvl(s3.st02,' '),6,' '),rpad(nvl(S1.st02,' '),6,' ')),rpad(nvl(nvl(tm15,tm12),' '),20,' '),rpad(tm01||'-'||tm02||'-'||tm03||'-'||tm04,15,' '),rpad(tm05||tm06||tm07,20,' '),rpad('申請',10,' '), " & _
            " rpad(DECODE(cp27,'','',SUBSTR(cp27,1,4)-1911||'/'||SUBSTR(cp27,5,2)||'/'||SUBSTR(cp27,7,2)),10,' '),rpad(tm23,10,' '),rpad(cu04,12,' '),nvl(cp27,0),tm01,tm02,tm03,tm04 from ( select tm01,tm02,tm03,tm04,tm05,tm06,tm07,tm23,tm15,tm12 " & _
            " From trademark where tm11>=20000000 and tm01 in ('FCT') and tm10='000' and tm29 is null and tm16 is null and tm28='1'  and tm11 <=to_number(to_char(add_months(sysdate,-20),'yyyymmdd')) " & _
            " and not exists(select * from caseprogress C1 where C1.cp01=tm01 and C1.cp02=tm02  and C1.cp03=tm03 and C1.cp04=tm04 and ((c1.cp10 not in ('703','704') and c1.cp09<'C' and c1.cp27>to_number(to_char(add_months(sysdate,-6),'yyyymmdd'))) " & _
            " or (c1.cp09>'C' and c1.cp05>to_number(to_char(add_months(sysdate,-6),'yyyymmdd'))) or (C1.cp10<>'1101' and C1.cp09>'C') or (c1.cp10='310' and c1.cp27>to_number(to_char(add_months(sysdate,-24),'yyyymmdd'))) " & _
            " or (c1.cp10='1401' and c1.cp05>to_number(to_char(add_months(sysdate,-24),'yyyymmdd'))) or (C1.cp10='306' and C1.cp27 is not null and exists (select * from caseprogress C3 where c1.cp43=C3.cp09 and C3.cp10='101') ))  and C1.cp05>=20000000)) " & _
            " ,caseprogress C2,staff S1,customer,staff s2,staff s3 where tm01=C2.cp01(+) and tm02=C2.cp02(+) and tm03=C2.cp03(+) and tm04=C2.cp04(+) and C2.cp10='101' and C2.cp14=s1.st01(+) and C2.cp65=s2.st01(+) and C2.cp13=s3.st01(+) and substr(tm23,1,8)=cu01(+) " & _
            " and substr(tm23,9,1)=cu02(+) " & _
            " and not exists (select C4.cp09 from caseprogress C4,caseprogress C5 where C4.cp01=tm01 and C4.cp02=tm02 and C4.cp03=tm03 and C4.cp04=tm04 and C4.cp10='101' and C4.cp09=C5.cp43(+) and C5.cp10='1724')" & _
            " union all select substr(nvl(S1.st03,s2.st03),1,1),decode(substr(nvl(S1.st03,s2.st03),1,1),'F',rpad(nvl(s3.st02,' '),6,' '),rpad(nvl(S1.st02,' '),6,' ')),rpad(nvl(nvl(tm15,tm12),' '),20,' '),rpad(tm01||'-'||tm02||'-'||tm03||'-'||tm04,15,' '),rpad(tm05||tm06||tm07,20,' '),rpad('核駁前先行',10,' '), " & _
            " rpad(DECODE(cp27,'','',SUBSTR(cp27,1,4)-1911||'/'||SUBSTR(cp27,5,2)||'/'||SUBSTR(cp27,7,2)),10,' '), rpad(tm23,10,' '),rpad(cu04,12,' '),nvl(cp27,0),tm01,tm02,tm03,tm04 From trademark, caseprogress c2, staff s1, customer,staff s2,staff s3 " & _
            " where c2.cp05>=20000000 and tm01 in ('FCT') and tm10='000' and tm29 is null and tm16 is null and tm01=c2.cp01(+)  and tm02=c2.cp02(+) and tm03=c2.cp03(+) and tm04=c2.cp04(+) and c2.cp10='1202' " & _
            " and not exists(select * from caseprogress C1 where C1.cp01=tm01 and C1.cp02=tm02  and C1.cp03=tm03 and C1.cp04=tm04 and ((c1.cp10 not in ('703','704') and c1.cp09<'C' and c1.cp27>to_number(to_char(add_months(sysdate,-6),'yyyymmdd'))) " & _
            " or (c1.cp09>'C' and c1.cp05>to_number(to_char(add_months(sysdate,-6),'yyyymmdd'))) or (c1.cp10='310' and c1.cp27>to_number(to_char(add_months(sysdate,-24),'yyyymmdd'))) " & _
            " or (c1.cp10='1401' and c1.cp05>to_number(to_char(add_months(sysdate,-24),'yyyymmdd'))) or (c1.cp10='1002' and c1.cp64='申請駁回') or (C1.cp10='306' and C1.cp27 is not null and not exists (select * from caseprogress C3 where c1.cp43=C3.cp09(+) and C3.cp10='101') )) " & _
            " and C1.cp05>=20000000) and c2.cp14=s1.st01(+) and c2.cp65=s2.st01(+) and C2.cp13=s3.st01(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and c2.cp05 <=to_number(to_char(add_months(sysdate,-12),'yyyymmdd') ) " & _
            " and not exists (select C4.cp09 from caseprogress C4,caseprogress C5 where C4.cp01=tm01 and C4.cp02=tm02 and C4.cp03=tm03 and C4.cp04=tm04 and C4.cp10='101' and C4.cp09=C5.cp43(+) and C5.cp10='1724')" & _
            " order by 1,2,6,10,4 "
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   tmpnext = "已取得資料..."
   With Rs
        TempFileNameT = ""
        TempFileNameFCT = ""
        SeekA02 = ""
        If .RecordCount > 0 And .RecordCount <> 0 Then
               .MoveFirst
               Do While Not .EOF
                    A01 = "" & (.Fields(0).Value)
                    A02 = "" & (.Fields(1).Value)
                    A03 = "" & (.Fields(2).Value)
                    A04 = "" & (.Fields(3).Value)
                    A05 = "" & (.Fields(4).Value)
                    A06 = "" & (.Fields(5).Value)
                    A07 = "" & (.Fields(6).Value)
                    A08 = "" & (.Fields(7).Value)
                    A09 = "" & (.Fields(8).Value)
                    'Modify By Sindy 2011/12/26 只判斷FCT的
                    '若該收文號的下一程序檔有掛305.催審期限而NP08.本所期限尚未到期尚不列出(FCT-028812),
                    '或該催審期限的np06 is not null也不列出(FCT-030066).
                    '若下一程序檔無催審期限則仍然要列出.
                    strSql = "select * from nextprogress,caseprogress where np01=cp09 and cp10='101' " & _
                             "and np02='" & "" & (.Fields("tm01").Value) & "' " & _
                             "and np03='" & "" & (.Fields("tm02").Value) & "' " & _
                             "and np04='" & "" & (.Fields("tm03").Value) & "' " & _
                             "and np05='" & "" & (.Fields("tm04").Value) & "' " & _
                             "and np07='305' " & _
                             "and (np08>" & strSrvDate(1) & " or np06 is not null) "
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                    If UCase(A01) = "F" And (intI = 1 And RsTemp.RecordCount > 0) Then
                        '不列出
                        'MsgBox (.Fields("tm01").Value) & "-" & (.Fields("tm02").Value) & "-" & (.Fields("tm03").Value) & "-" & (.Fields("tm04").Value)
                    Else
                    '2011/12/26 End
                       If SeekA02 <> A02 Then
                           If ff > 0 Then Close #ff
                            ff = FreeFile
                            'TempFileName = "c:\" & Trim(A02) & Format(Now, "YYYYMMDDhhmmss") & ".txt"
                            'Modify By Sindy 2009/06/04
                            'TempFileName = "c:\" & IIf(Trim(A02) = "", "沒承辦人" & IIf(A01 = "F", "外商", "內商"), Trim(A02)) & ".txt"
                            TempFileName = App.path & TextPath & IIf(Trim(A02) = "", "沒承辦人" & IIf(A01 = "F", "外商", "內商"), Trim(A02)) & ".txt"
                            TempFileName = PUB_UniToBIG5(TempFileName, "F") 'Added by Lydia 2022/03/28 員工名稱有Unicode
                            '2009/06/04 End
                            If UCase(A01) = "F" Then
                               TempFileNameFCT = TempFileNameFCT & TempFileName & ";"
                            Else
                               TempFileNameT = TempFileNameT & TempFileName & ";"
                            End If
                            Open TempFileName For Output As ff
   'add by nickc 2005/06/03 加欄位
   '                         Print #ff, "承辦人  本所案號　　　　案件名稱　　　　　　 案件性質　 發文日　　 申請人編號 申請人名稱　"
   '                         Print #ff, "======= =============== ==================== ========== ========== ========== ==========="
                            'Added by Lydia 2016/10/19 加報表抬頭
                            Print #ff, Space(20) & "申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細"
                            Print #ff, ""
                            'end 2016/10/19
                            Print #ff, "承辦人  申請案號             本所案號　　　　案件名稱　　　　　　 案件性質　 發文日　　 申請人編號 申請人名稱　"
                            Print #ff, "======= ==================== =============== ==================== ========== ========== ========== ==========="
                            tmpnext = "開檔中..."
                            SeekA02 = A02
                       End If
                       tmpnext = "寫檔..."
                       'Print #ff, A02 & "  " & A03 & " " & A04 & " " & A05 & " " & A06 & " " & A07 & " " & A08
                       Print #ff, A02 & "  " & A03 & " " & A04 & " " & A05 & " " & A06 & " " & A07 & " " & A08 & " " & A09
                    End If
                    .MoveNext
               Loop
        End If
   End With
   Close ff
   tmpnext = "關檔..."
   If TempFileNameT = "" And TempFileNameFCT = "" Then Exit Sub
   
   'Modify By Sindy 2011/10/27
   strSubject = "申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細---------" & Format(Now, "YYYY/MM/DD")
   'modify by sonia 2016/10/3 取消稱謂
   'strContent = "Dear 葉經理：" & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的    申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細   " & IIf(TempFileNameT = "", "資料庫找不到資料", "資料如附件") & "！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "請橫印！" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
   strContent = "Dear Sirs," & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的    申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細   " & IIf(TempFileNameT = "", "資料庫找不到資料", "資料如附件") & "！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "請橫印！" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
   'Modified by Lydia 2016/10/19 改發為林經理，並同時副本通知葉特助。
   'SendMAPIMail "67002", strSubject, strContent, TempFileNameT
   'Modify by Amy 2020/05/05 副本取消通知葉特助
   'SendMAPIMail "69008", strSubject, strContent, TempFileNameT, , "67002"
   'SendMAPIMail "69008", strSubject, strContent, TempFileNameT
   SendMAPIMail "A2004", strSubject, strContent, TempFileNameT
   ''發 mail
   'tmpnext = "準備發 mail..."
   'MAPISession1.LogonUI = False
   'MAPISession1.UserName = "administrator"
   'tmpnext = "準備登入郵件伺服器..."
   'MAPISession1.SignOn
   'tmpnext = "登入郵件伺服器..."
   'MAPIMessages1.SessionID = MAPISession1.SessionID
   ''發葉經理
   'MAPIMessages1.MsgIndex = -1
   'tmpnext = "建立郵件..."
   'MAPIMessages1.Compose
   'MAPIMessages1.MsgSubject = "申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細---------" & Format(Now, "YYYY/MM/DD")
   'MAPIMessages1.MsgNoteText = "Dear 葉經理：" & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的    申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細   " & IIf(TempFileNameT = "", "資料庫找不到資料", "資料如附件") & "！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "請橫印！" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
   'If TempFileNameT <> "" Then
   '    ArrMail = Split(TempFileNameT, ";")
   '    For StrMenu3I = 0 To UBound(ArrMail) - 1
   '        MAPIMessages1.AttachmentIndex = StrMenu3I
   '        MAPIMessages1.AttachmentPosition = 120 + (StrMenu3I * 5)
   '        MAPIMessages1.AttachmentPathName = ArrMail(StrMenu3I)
   '    Next StrMenu3I
   'End If
   'MAPIMessages1.RecipIndex = 0
   'MAPIMessages1.RecipDisplayName = "67002"
   ''MAPIMessages1.RecipDisplayName = "93013"
   'MAPIMessages1.ResolveName
   'tmpnext = "準備存入郵件..."
   'MAPIMessages1.Send
   
   'Modify By Sindy 2011/10/27
   strSubject = "申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細---------" & Format(Now, "YYYY/MM/DD")
   'modify by sonia 2016/10/3 取消稱謂
   strContent = "Dear Sirs," & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的    申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細   " & IIf(TempFileNameFCT = "", "資料庫找不到資料", "資料如附件") & "！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "請橫印！" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
   'SendMAPIMail "68005", strSubject, strContent, TempFileNameFCT
   SendMAPIMail "A2004", strSubject, strContent, TempFileNameFCT
   ''發陳經理
   'MAPIMessages1.MsgIndex = -1
   'tmpnext = "建立郵件..."
   'MAPIMessages1.Compose
   'MAPIMessages1.MsgSubject = "申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細---------" & Format(Now, "YYYY/MM/DD")
   'MAPIMessages1.MsgNoteText = "Dear 陳經理：" & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的    申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細   " & IIf(TempFileNameFCT = "", "資料庫找不到資料", "資料如附件") & "！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "請橫印！" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
   'If TempFileNameFCT <> "" Then
   '    ArrMail = Split(TempFileNameFCT, ";")
   '    For StrMenu3I = 0 To UBound(ArrMail) - 1
   '        MAPIMessages1.AttachmentIndex = StrMenu3I
   '        MAPIMessages1.AttachmentPosition = 120 + (StrMenu3I * 5)
   '        MAPIMessages1.AttachmentPathName = ArrMail(StrMenu3I)
   '    Next StrMenu3I
   'End If
   'MAPIMessages1.RecipIndex = 0
   'MAPIMessages1.RecipDisplayName = "68005"
   ''MAPIMessages1.RecipDisplayName = "93013"
   'MAPIMessages1.ResolveName
   'tmpnext = "準備存入郵件..."
   'MAPIMessages1.Send
   'tmpnext = "發信..."
   'MAPISession1.SignOff
   'tmpnext = "登出..."
   If TempFileNameT <> "" Then
       ArrMail = Split(TempFileNameT, ";")
       For StrMenu3I = 0 To UBound(ArrMail) - 1
           Kill ArrMail(StrMenu3I)
       Next StrMenu3I
   End If
   If TempFileNameFCT <> "" Then
       ArrMail = Split(TempFileNameFCT, ";")
       For StrMenu3I = 0 To UBound(ArrMail) - 1
           Kill ArrMail(StrMenu3I)
       Next StrMenu3I
   End If
   tmpnext = "清除暫存檔..."
   'Shell "net send /domain:taient2 '每月1日自動郵件資料已送出，請清除郵件備份' ", vbNormalNoFocus
   Set Rs = Nothing
   'Set cnnConnection = Nothing
   Exit Sub
DebugErr:
   MsgBox tmpnext & " " & Err.Description
End Sub

'Add By Sindy 2009/05/26
'補專利(分割案及國內外關聯案)代表圖
Sub StrMenu4()
Dim strSqlT As String
Dim tmpnext As String
Dim p_fPA(4) As String
Dim p_tPA(4) As String

On Error GoTo DebugErr
   'tmpnext = "準備連資料庫..."
   'Set cnnConnection = New ADODB.Connection
   'cnnConnection.ConnectionString = CnStr
   'cnnConnection.Open
   'Modify by Amy 2018/07/23 +彩色代表圖 原:and ibf05='1'
   '分割案 : A有圖B無圖則複製A圖到B案
   'Modified by Morgan 2021/3/10  +有新申請案發文且繪圖人員非 9999(不繪圖)--玲玲
   strSql = "select Distinct dc01,dc02,dc03,dc04,dc05,dc06,dc07,dc08 from divisioncase,imgbytefile " & _
                  "where dc01=ibf01 and dc02=ibf02 and dc03=ibf03 and dc04=ibf04 and ibf05 in('1','2') " & _
                  "and dc01 in ('P','CFP') and dc05 in ('P','CFP') " & _
                  "and not exists (select * from imgbytefile where dc05=ibf01 and dc06=ibf02 and dc07=ibf03 and dc08=ibf04 and ibf05 in('1','2')) " & _
                  "and exists (select * from caseprogress where cp01=dc05 and cp02=dc06 and cp03=dc07 and cp04=dc08 and cp10 in (" & NewCasePtyList & ") and cp27>0 and nvl(cp29,'X')<>'99999') "

                  
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   tmpnext = "已取得資料(1)..."
   With Rs
      If .RecordCount > 0 And .RecordCount <> 0 Then
         .MoveFirst
         Do While Not .EOF
            p_fPA(1) = "" & Rs.Fields("dc01")
            p_fPA(2) = "" & Rs.Fields("dc02")
            p_fPA(3) = "" & Rs.Fields("dc03")
            p_fPA(4) = "" & Rs.Fields("dc04")
            p_tPA(1) = "" & Rs.Fields("dc05")
            p_tPA(2) = "" & Rs.Fields("dc06")
            p_tPA(3) = "" & Rs.Fields("dc07")
            p_tPA(4) = "" & Rs.Fields("dc08")
            'Modify By Sindy 2014/7/24 不管是分割案或國內外關聯案時，都要控制該案之新申請案那一道是否已發文
            '，已發文的案件才可以補圖
            'Modified by Lydia 2016/09/14 nvl(cp27,0)>0 => CP158>0
            strSql = "select cp09 from caseprogress" & _
                     " where cp01='" & p_tPA(1) & "'" & _
                       " and cp02='" & p_tPA(2) & "'" & _
                       " and cp03='" & p_tPA(3) & "'" & _
                       " and cp04='" & p_tPA(4) & "'" & _
                       " and cp10 in(" & NewCasePtyList & ")" & _
                       " and CP158>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
            '2014/7/24 END
               Call PUB_CopyImgFile(p_fPA(), p_tPA())
            End If
            .MoveNext
            DoEvents
         Loop
      End If
   End With
   'CheckOC
   
   '分割案 : B有圖A無圖則複製B圖到A案
   'Modified by Morgan 2021/3/10  +有新申請案發文且繪圖人員非 9999(不繪圖)--玲玲
   strSql = "select Distinct dc01,dc02,dc03,dc04,dc05,dc06,dc07,dc08 from divisioncase,imgbytefile " & _
                  "where dc05=ibf01 and dc06=ibf02 and dc07=ibf03 and dc08=ibf04 and ibf05 in('1','2') " & _
                  "and dc01 in ('P','CFP') and dc05 in ('P','CFP') " & _
                  "and not exists (select * from imgbytefile where dc01=ibf01 and dc02=ibf02 and dc03=ibf03 and dc04=ibf04 and ibf05 in('1','2')) " & _
                  "and exists (select * from caseprogress where cp01=dc01 and cp02=dc02 and cp03=dc03 and cp04=dc04 and cp10 in (" & NewCasePtyList & ") and cp27>0 and nvl(cp29,'X')<>'99999') "
                  
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   tmpnext = "已取得資料(2)..."
   With Rs
      If .RecordCount > 0 And .RecordCount <> 0 Then
         .MoveFirst
         Do While Not .EOF
            p_fPA(1) = "" & Rs.Fields("dc05")
            p_fPA(2) = "" & Rs.Fields("dc06")
            p_fPA(3) = "" & Rs.Fields("dc07")
            p_fPA(4) = "" & Rs.Fields("dc08")
            p_tPA(1) = "" & Rs.Fields("dc01")
            p_tPA(2) = "" & Rs.Fields("dc02")
            p_tPA(3) = "" & Rs.Fields("dc03")
            p_tPA(4) = "" & Rs.Fields("dc04")
            'Modify By Sindy 2014/7/24 不管是分割案或國內外關聯案時，都要控制該案之新申請案那一道是否已發文
            '，已發文的案件才可以補圖
            'Modified by Lydia 2016/09/14 nvl(cp27,0)>0 => CP158>0
            strSql = "select cp09 from caseprogress" & _
                     " where cp01='" & p_tPA(1) & "'" & _
                       " and cp02='" & p_tPA(2) & "'" & _
                       " and cp03='" & p_tPA(3) & "'" & _
                       " and cp04='" & p_tPA(4) & "'" & _
                       " and cp10 in(" & NewCasePtyList & ")" & _
                       " and CP158>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
            '2014/7/24 END
               Call PUB_CopyImgFile(p_fPA(), p_tPA())
            End If
            .MoveNext
            DoEvents
         Loop
      End If
   End With
   
   '國內外關聯案 : A有圖B無圖則複製A圖到B案
   'Modified by Morgan 2021/3/10  +有新申請案發文且繪圖人員非 9999(不繪圖)--玲玲
   strSql = "select Distinct cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08 from CaseMap,imgbytefile " & _
                  "where cm01=ibf01 and cm02=ibf02 and cm03=ibf03 and cm04=ibf04 and ibf05 in('1','2') " & _
                  "and cm01 in ('P','CFP') and cm05 in ('P','CFP') " & _
                  "and cm10 in ('0','4') " & _
                  "and not exists (select * from imgbytefile where cm05=ibf01 and cm06=ibf02 and cm07=ibf03 and cm08=ibf04 and ibf05 in('1','2')) " & _
                  "and exists (select * from caseprogress where cp01=cm05 and cp02=cm06 and cp03=cm07 and cp04=cm08 and cp10 in (" & NewCasePtyList & ") and cp27>0 and nvl(cp29,'X')<>'99999') "
                  
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   tmpnext = "已取得資料(3)..."
   With Rs
      If .RecordCount > 0 And .RecordCount <> 0 Then
         .MoveFirst
         Do While Not .EOF
            p_fPA(1) = "" & Rs.Fields("cm01")
            p_fPA(2) = "" & Rs.Fields("cm02")
            p_fPA(3) = "" & Rs.Fields("cm03")
            p_fPA(4) = "" & Rs.Fields("cm04")
            p_tPA(1) = "" & Rs.Fields("cm05")
            p_tPA(2) = "" & Rs.Fields("cm06")
            p_tPA(3) = "" & Rs.Fields("cm07")
            p_tPA(4) = "" & Rs.Fields("cm08")
            'Modify By Sindy 2014/7/24 不管是分割案或國內外關聯案時，都要控制該案之新申請案那一道是否已發文
            '，已發文的案件才可以補圖
            '國內外關聯案時，要剔除該案之新申請案那一道的案件性質為113或114或115的資料
            'Modified by Lydia 2016/09/14 nvl(cp27,0)>0 => CP158>0
            strSql = "select cp09 from caseprogress" & _
                     " where cp01='" & p_tPA(1) & "'" & _
                       " and cp02='" & p_tPA(2) & "'" & _
                       " and cp03='" & p_tPA(3) & "'" & _
                       " and cp04='" & p_tPA(4) & "'" & _
                       " and cp10 in(" & NewCasePtyList & ")" & _
                       " and cp10 not in(113,114,115)" & _
                       " and CP158>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
            '2014/7/24 END
               Call PUB_CopyImgFile(p_fPA(), p_tPA())
            End If
            .MoveNext
            DoEvents
         Loop
      End If
   End With
   
   '國內外關聯案 : B有圖A無圖則複製B圖到A案
   'Modified by Morgan 2021/3/10 +有新申請案發文且繪圖人員非 9999(不繪圖)--玲玲
   strSql = "select Distinct cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08 from CaseMap,imgbytefile " & _
                  "where cm05=ibf01 and cm06=ibf02 and cm07=ibf03 and cm08=ibf04 and ibf05 in('1','2') " & _
                  "and cm01 in ('P','CFP') and cm05 in ('P','CFP') " & _
                  "and cm10 in ('0','4') " & _
                  "and not exists (select * from imgbytefile where cm01=ibf01 and cm02=ibf02 and cm03=ibf03 and cm04=ibf04 and ibf05 in('1','2')) " & _
                  "and exists (select * from caseprogress where cp01=cm01 and cp02=cm02 and cp03=cm03 and cp04=cm04 and cp10 in (" & NewCasePtyList & ") and cp27>0 and nvl(cp29,'X')<>'99999') "
   'end 2018/07/23
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   tmpnext = "已取得資料(4)..."
   With Rs
      If .RecordCount > 0 And .RecordCount <> 0 Then
         .MoveFirst
         Do While Not .EOF
            p_fPA(1) = "" & Rs.Fields("cm05")
            p_fPA(2) = "" & Rs.Fields("cm06")
            p_fPA(3) = "" & Rs.Fields("cm07")
            p_fPA(4) = "" & Rs.Fields("cm08")
            p_tPA(1) = "" & Rs.Fields("cm01")
            p_tPA(2) = "" & Rs.Fields("cm02")
            p_tPA(3) = "" & Rs.Fields("cm03")
            p_tPA(4) = "" & Rs.Fields("cm04")
            'Modify By Sindy 2014/7/24 不管是分割案或國內外關聯案時，都要控制該案之新申請案那一道是否已發文
            '，已發文的案件才可以補圖
            '國內外關聯案時，要剔除該案之新申請案那一道的案件性質為113或114或115的資料
            'Modified by Lydia 2016/09/14 nvl(cp27,0)>0 => CP158>0
            strSql = "select cp09 from caseprogress" & _
                     " where cp01='" & p_tPA(1) & "'" & _
                       " and cp02='" & p_tPA(2) & "'" & _
                       " and cp03='" & p_tPA(3) & "'" & _
                       " and cp04='" & p_tPA(4) & "'" & _
                       " and cp10 in(" & NewCasePtyList & ")" & _
                       " and cp10 not in(113,114,115)" & _
                       " and cp158>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
            '2014/7/24 END
               Call PUB_CopyImgFile(p_fPA(), p_tPA())
            End If
            .MoveNext
            DoEvents
         Loop
      End If
   End With
   
   'Shell "net send /domain:taient2 '每月1日自動郵件資料已送出，請清除郵件備份' ", vbNormalNoFocus
   Set Rs = Nothing
   'Set cnnConnection = Nothing
   Exit Sub
DebugErr:
        MsgBox tmpnext & " " & Err.Description
End Sub

'申復案清單
Sub StrMenu5()
Dim ff As Integer
Dim i As Integer
Dim A01 As String
Dim A02 As String
Dim A03 As String
Dim A04 As String
Dim A05 As String
Dim iCount As Long
Dim TempFileName As String
Dim tmpnext As String
Dim strCntDate As String, strSysDate As String
Dim strSubject As String, strContent As String 'Add By Sindy 2011/10/27

On Error GoTo DebugErr
   'tmpnext = "準備連資料庫..."
   'Set cnnConnection = New ADODB.Connection
   'cnnConnection.ConnectionString = CnStr
   'cnnConnection.Open
   
   'Modify By Sindy 2012/8/24 原只抓案件性質202申請意見書的資料, 但100/11加入210陳述意見書後, 此程式未修改, 故補加入; 同時加入FCT案件
   strSql = "select TM01||'-'||TM02||'-'||TM03||'-'||TM04,TM12,CP36,CP37,CP40,CP27 " & _
                  "From caseprogress, trademark " & _
                  "Where CP01 in('T','FCT') and CP10 in('202','210') and CP24 is null and not CP27 is null " & _
                  "and CP01=TM01(+) and CP02=TM02(+) and CP03=TM03(+) and CP04=TM04(+) " & _
                  "and TM28<>'1' and TM29 is null and TM10='000' " & _
                  "order by TM01,TM02,TM03,TM04 "
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   tmpnext = "已取得資料..."
   iCount = 0
   With Rs
         ff = FreeFile
         TempFileName = App.path & TextPath & Format(Now, "YYYYMMDD") & "申復案清單.txt"
         Open TempFileName For Output As ff
         tmpnext = "開檔中..."
         'Added by Lydia 2016/10/19 加報表抬頭
         Print #ff, Space(30) & Format(Now, "YYYYMMDD") & "申復案清單"
         Print #ff, ""
         'end 2016/10/19
         Print #ff, "本所案號       申請案號       對造號數       對造案件名稱　　                對造中文"
         Print #ff, "============   ===========    ===========    ============================    ==========================="
         
         If .RecordCount > 0 And .RecordCount <> 0 Then
            .MoveFirst
            strSysDate = ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(Format(Now, "YYYYMMDD"))))
            Do While Not .EOF
               strCntDate = ChangeWDateStringToWString(DateAdd("m", 3, ChangeWStringToWDateString(.Fields("CP27").Value)))
               Do While Left(strCntDate, 6) <= Left(strSysDate, 6)
                  If Left(strCntDate, 6) = Left(strSysDate, 6) Then
                        iCount = iCount + 1
                        tmpnext = "寫檔..."
                        A01 = Left("" & Trim(.Fields(0).Value) & "               ", 15)
                        A02 = Left("" & Trim(.Fields(1).Value) & "               ", 15)
                        A03 = Left("" & Trim(.Fields(2).Value) & "               ", 15)
                        A04 = Left("" & Trim(.Fields(3).Value) & "                              ", 30)
                        A05 = "" & Trim(.Fields(4).Value)
                        Print #ff, A01 & A02 & A03 & A04 & A05
                  End If
                  strCntDate = ChangeWDateStringToWString(DateAdd("m", 3, ChangeWStringToWDateString(strCntDate)))
               Loop
               .MoveNext
            Loop
         End If
         If iCount = 0 Then
            TempFileName = ""
         End If
   End With
   Close ff
   tmpnext = "關檔..."
   
   'Modify By Sindy 2011/10/27
   strSubject = "申復案清單---------" & Format(Now, "YYYY/MM/DD")
   strContent = "Dear Sirs," & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的    申復案清單   " & IIf(TempFileName = "", "資料庫找不到資料", "資料如附件") & "！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
   SendMAPIMail Pub_GetSpecMan("P1"), strSubject, strContent, TempFileName
   
   ''發 mail
   'tmpnext = "準備發 mail..."
   'MAPISession1.LogonUI = False
   'MAPISession1.UserName = "administrator"
   'tmpnext = "準備登入郵件伺服器..."
   'MAPISession1.SignOn
   'tmpnext = "登入郵件伺服器..."
   'MAPIMessages1.SessionID = MAPISession1.SessionID
   'MAPIMessages1.MsgIndex = -1
   'tmpnext = "建立郵件..."
   'MAPIMessages1.Compose
   'MAPIMessages1.MsgSubject = "申復案清單---------" & Format(Now, "YYYY/MM/DD")
   ''Modify By Sindy 2010/6/30
   ''MAPIMessages1.MsgNoteText = "Dear 桂英：" & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的    申復案清單   " & IIf(TempFileName = "", "資料庫找不到資料", "資料如附件") & "！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
   'MAPIMessages1.MsgNoteText = "Dear Sirs," & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的    申復案清單   " & IIf(TempFileName = "", "資料庫找不到資料", "資料如附件") & "！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心"
   'If TempFileName <> "" Then
   '    MAPIMessages1.AttachmentPosition = 100
   '    MAPIMessages1.AttachmentPathName = TempFileName
   'End If
   'MAPIMessages1.RecipIndex = 0
   ''Modify By Sindy 2010/6/30
   ''MAPIMessages1.RecipDisplayName = "79041"
   'MAPIMessages1.RecipDisplayName = Pub_GetSpecMan("P1")
   'MAPIMessages1.ResolveName
   'tmpnext = "準備存入郵件..."
   'MAPIMessages1.Send
   'tmpnext = "發信..."
   'MAPISession1.SignOff
   'tmpnext = "登出..."
   If TempFileName <> "" Then
       Kill TempFileName
   End If
   tmpnext = "清除暫存檔..."
   Set Rs = Nothing
   'Set cnnConnection = Nothing
   Exit Sub
DebugErr:
        MsgBox tmpnext & " " & Err.Description
End Sub

''Add By Sindy 2009/08/18 延展案於期滿六個月前一個工作天發mail提醒承辦人
'Sub StrMenu6()
'Dim ff As Integer
'Dim i As Integer
'Dim A01 As String
'Dim A02 As String
'Dim A03 As String
'Dim A04 As String
'Dim A05 As String
'Dim A06 As String
'Dim A07 As String
'Dim iCount As Long
'Dim TempFileName As String
'Dim TempFileNameSeek As String
'Dim ThisFileName As Variant
'Dim StrSQL6 As String
'Dim T_Date As String
'Dim tmpnext As String
'
'On Error GoTo DebugErr
'   'tmpnext = "準備連資料庫..."
'   'Set cnnConnection = New ADODB.Connection
'   'cnnConnection.ConnectionString = CnStr
'   'cnnConnection.Open
'   TempFileName = ""
'   T_Date = CompWorkDay(2, Format(Now, "YYYYMMDD"), 0)           '後一個工作天
'
'   CheckOC
'   strSql = "SELECT nvl(CP14,'68005') as C14,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as sortA,TM15,TM05,TM09,Cu04,ST02,(cp07-19110000) as cp" & _
'                  " FROM CaseProgress,Customer,Staff,TradeMark " & _
'                  " WHERE cp01='FCT' and  CP10='102'  and  cp07 > 20000000 and (cp27 is null) and (cp57 is null) and (TM29 is null) and TO_CHAR(ADD_MONTHS(TO_date(cp07,'YYYYMMDD'),-6),'YYYYMMDD') < " & T_Date & _
'                  " and cp05<=TO_CHAR(ADD_MONTHS(TO_date(cp07,'YYYYMMDD'),-6),'YYYYMMDD') and cp13=st01(+) and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+)  order by nvl(CP14,'68005'),sortA "
'
'   Set rs = New ADODB.Recordset
'   rs.CursorLocation = adUseClient
'   rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   tmpnext = "已取得資料..."
'   TempFileNameSeek = ""
'
'   With rs
'       TempFileName = ""
'       If .RecordCount > 0 And .RecordCount <> 0 Then
'           .MoveFirst
'           ff = FreeFile
'           Do While Not .EOF
'               If InStr(1, TempFileNameSeek, CheckStr(.Fields("C14"))) = 0 Then
'                   TempFileName = CheckStr(.Fields("C14"))
'                   TempFileNameSeek = TempFileNameSeek & "," & CheckStr(.Fields("C14"))
'                   If ff > 0 Then Close #ff
'                   ff = FreeFile
'                   'Modify By Sindy 2009/06/04
'                   'Open "c:\" & TempFileName & ".txt" For Output As ff
'                   Open App.path & TextPath & TempFileName & ".txt" For Output As ff
'                   '2009/06/04 End
'                   'Add By Sindy 2009/07/20
'                   Print #ff, "　　　　　　　　　　　　　　　　　　　　　延展案期滿前六個月可辦案件明細"
'                   Print #ff, ""
'                   '2009/07/20 End
'                   Print #ff, "本所案號   註冊號                商標名稱             類別                 客戶名稱             智權人員     法定期限    "
'                   Print #ff, "========== ===================== ==================== ==================== ==================== ============ ========"
'               End If
'               tmpnext = "開檔中..."
'               A01 = "" & convForm(CheckStr(.Fields(1).Value), 10)
'               A02 = "" & convForm(CheckStr(.Fields(2).Value), 20)
'               A03 = "" & convForm(CheckStr(.Fields(3).Value), 20)
'               A04 = "" & convForm(CheckStr(.Fields(4).Value), 20)
'               A05 = "" & convForm(CheckStr(.Fields(5).Value), 20)
'               A06 = "" & convForm(CheckStr(.Fields(6).Value), 12)
'               A07 = "" & convForm(CheckStr(.Fields(7).Value), 8)
'               tmpnext = "寫檔..."
'
'               Print #ff, A01 & " " & A02 & " " & " " & A03 & " " & A04 & " " & A05 & " " & A06 & " " & A07 & " "
'               .MoveNext
'           Loop
'       End If
'   End With
'   Close ff
'   tmpnext = "關檔..."
'
'   If TempFileNameSeek <> "" Then
'      ThisFileName = Split(TempFileNameSeek, ",")
'      For iCount = 0 To UBound(ThisFileName)
'         If ThisFileName(iCount) <> "" Then
'            TempFileName = ThisFileName(iCount)
'            '發 mail
'            'Modify By Sindy 2009/06/04
'            'SendMAPIMail TempFileName, "延展案期滿六個月，案件明細", "Dear Sirs," & vbCrLf & "          逾承辦期限有完稿日無會稿日 資料如附件！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "          請橫印！" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心", "c:\" & TempFileName & ".txt"
'            'Modify By Sindy 2009/07/20
'            'SendMAPIMail TempFileName, "延展案期滿六個月，案件明細", "Dear Sirs," & vbCrLf & "          逾承辦期限有完稿日無會稿日 資料如附件！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "          請橫印！" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心", App.Path & TextPath & TempFileName & ".txt"
'            SendMAPIMail TempFileName, "延展案期滿前六個月可辦案件明細", "Dear Sirs," & vbCrLf & "          延展案期滿前七個月可辦案件明細 資料如附件！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "          請橫印！" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心", App.path & TextPath & TempFileName & ".txt"
'            '2009/06/04 End
'         End If
'      Next iCount
'   'edit by nickc 2006/03/20 因為 taient4 的服務預設關閉
'   '   Shell "net send /domain:taient4 '每日自動郵件資料已送出，請清除郵件備份' ", vbNormalNoFocus
'   End If
'   Set rs = Nothing
'   'Set cnnConnection = Nothing
'   Exit Sub
'DebugErr:
'   'MsgBox tmpnext & " " & Err.Description
'   WLog tmpnext & " " & Err.Description, 1
'End Sub

'Add by Morgan 2005/10/13
'限定字串長度
'Remove by Lydia 2018/08/24 與basQuery重複
'Private Function convForm(ByVal p_InStr As String, ByVal p_Num As Integer, Optional ByVal p_Char As String = " ") As String
'   convForm = StrConv(LeftB(StrConv(p_InStr & String(p_Num, p_Char), vbFromUnicode), p_Num), vbUnicode)
'End Function

'Add by Morgan 2005/10/13
'發Mail
'Modified by Lydia 2016/10/07 與frmAutoBatchDay一致,增加副本收件人(CCr$)、 密件副本(p_BCCr$)、p_HTML
'Private Function SendMAPIMail(p_RecID$, p_Sub$, p_Text$, Optional p_AttPath$, Optional bolAbsenceSys As Boolean = False) As Boolean
Private Function SendMAPIMail(ByVal p_RecID$, ByVal p_Sub$, ByVal p_Text$, Optional p_AttPath$, Optional bolAbsenceSys As Boolean = False, Optional CCr$, Optional p_BCCr$, Optional p_HTML As Boolean = False) As Boolean


'Added by Morgan 2019/9/10
'統一改 HTML 格式
If p_HTML = False Then
   p_Text = Replace(p_Text & " ", " ", "&nbsp;")  '為避免後面的字型設定無效，必需手動轉空白(因tag內的空白不可轉否則無效),最後面多加一個空白避免內文沒有
   p_Text = "<DIV style=""FONT: 12pt 細明體"">" & p_Text & "</DIV>"
   p_HTML = True
End If
'end 2019/9/10

'Added by Morgan 2014/1/2
bolMailSendOk = False
'Modify By Sindy 2014/4/10
'PUB_SendMail strUserNum, p_RecID, "", p_Sub, p_Text, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", Replace(p_AttPath, ";", "*"), False, , , , "QPGMR", "電腦中心", , bolAbsenceSys, False
'Modified by Lydia 2016/10/07
'PUB_SendMail strUserNum, p_RecID, "", p_Sub, p_Text, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", Replace(p_AttPath, ";", "*"), False, , , , "QPGMR", "系統管理員", , bolAbsenceSys, False
'Modified by Morgan 2019/10/25 +bolShowErrMsg=False
PUB_SendMail strUserNum, p_RecID, "", p_Sub, p_Text, vbCrLf & vbCrLf & "***此信件為系統自動寄出，請勿直接回覆。***", Replace(p_AttPath, ";", "*"), p_HTML, , , CCr$, "QPGMR", "系統管理員", , bolAbsenceSys, False, p_BCCr$, , False
'2014/4/10 END

'Added by Morgan 2014/1/23
SendMAPIMail = bolMailSendOk
If bolMailSendOk = False Then
   WLog "寄信失敗!! -->主旨:" & p_Sub & "  內容:" & p_Text & " 附件:" & p_AttPath, 1
End If
'end 2014/1/23
End Function

'Add By Sindy 2010/10/15 複製Performance專業件數上月資料至本月
Sub StrMenu9()
Dim i As Integer, k As Integer, strTemp3 As String, s As Integer
Dim tmpnext As String, dTemp As Date, strTemp As String

On Error GoTo DebugErr

   'tmpnext = "準備連資料庫..."
   'Set cnnConnection = New ADODB.Connection
   'cnnConnection.ConnectionString = CnStr
   'cnnConnection.Open
   
   strSrvDate(1) = Format(ServerDate)
   'Modified by Morgan 2024/1/11 修正日期轉字串會因顯示格式而不同問題
   'dTemp = DateAdd("m", -1, CDate(Left(strSrvDate(1), 4) & "-" & Mid(strSrvDate(1), 5, 2) & "-" & Right(strSrvDate(1), 2)))
   'strTemp = Format(dTemp, "YYYYMMDD")
   strTemp = Format(DateAdd("m", -1, CDate(Left(strSrvDate(1), 4) & "-" & Mid(strSrvDate(1), 5, 2) & "-" & Right(strSrvDate(1), 2))), "YYYYMMDD")
   'end 2024/1/11
   
   cnnConnection.BeginTrans
   'cnnConnection.Execute "delete performance where pe02='T' and pe05>0 and pe03='" & Left(Trim(strSrvDate(1)), 6) & "' "
   'modify by sonia 2019/12/12 加內商人員專業點數目標,但PE07~PE29不可複製
   'cnnConnection.Execute "insert into performance " & _
   '"select PE01,PE02,'" & Left(Trim(strSrvDate(1)), 6) & "',PE04,PE05,PE06,PE07,PE08,PE09,PE10,PE11,PE12,PE13,PE14,PE15,PE16,PE17,PE18,PE19,PE20,PE21,PE22,PE23,PE24,PE25,PE26,PE27,PE28,PE29 " & _
   '"From performance " & _
   '"where pe02='T' and pe05>0 and pe03='" & Left(strTemp, 6) & "' "
   cnnConnection.Execute "insert into performance (PE01,PE02,PE03,PE04,PE05,PE06) " & _
   "select PE01,PE02,'" & Left(Trim(strSrvDate(1)), 6) & "',PE04,PE05,PE06 " & _
   "From performance " & _
   "where pe02='T' and (pe05>0 or pe06>0) and pe03='" & Left(strTemp, 6) & "' "
   cnnConnection.CommitTrans
   
   'Set cnnConnection = Nothing
   Exit Sub

DebugErr:
   cnnConnection.RollbackTrans
   MsgBox tmpnext & " " & Err.Description
End Sub

'2012/8/30 ADD BY SONIA PS,CPS非台灣案一年內無進度上可結餘日期
Sub StrMenu10()

On Error GoTo CheckingErr
   cnnConnection.BeginTrans
   strSql = "SELECT DISTINCT CP01,CP02,CP03,CP04 FROM CASEPROGRESS," & _
            "(SELECT SP01 NO01,SP02 NO02,SP03 NO03,SP04 NO04,MAX(CP05) MAXDT FROM CASEPROGRESS,SERVICEPRACTICE WHERE SP01 IN ('PS','CPS') AND SP09<>'000' " & _
            "AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) GROUP BY SP01,SP02,SP03,SP04) X WHERE MAXDT<" & strSrvDate(1) - 10000 & _
            " AND NO01=CP01(+) AND NO02=CP02(+) AND NO03=CP03(+) AND NO04=CP04(+) AND CP27 IS NOT NULL AND CP59||CP109 IS NULL "
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 And .RecordCount > 0 Then
         .MoveFirst
         Do While .EOF = False
            bolEndModCash = True
            Pub_UpdateEndModCash CheckStr(.Fields(0)), CheckStr(.Fields(1)), CheckStr(.Fields(2)), CheckStr(.Fields(3))
            .MoveNext
            DoEvents
         Loop
      End If
   End With
   CheckOC
   cnnConnection.CommitTrans
   Exit Sub

CheckingErr:
   cnnConnection.RollbackTrans
End Sub
'2012/8/30 END

'Add By Amy 2013/05/03 刪除三年內無往來記錄之國內潛在客戶(排除D01及P12)
Sub StrMenu11()
Dim StrSQLa As String, StrSqlB As String
Dim rsTmp As New ADODB.Recordset
Dim strExclude As String 'Add by Amy 2015/04/27 +控制不可刪除編號
Dim strTp As String 'Add by Amy 2023/07/20

  strExclude = "'R0794500' "
On Error GoTo CheckErr
  'Add by Amy 2023/07/20 客戶代理人來源資料檔之介紹者R編號也不可刪
  strSql = "Select Distinct XYS02 From Xynosource Where SubStr(XYS02,1,1)='R' "
  CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         .MoveFirst
         Do While .EOF = False
            strTp = strTp & ",'" & .Fields("XYS02") & "'"
            .MoveNext
         Loop
      End If
   End With
   If strTp <> MsgText(601) Then strExclude = strExclude & strTp
   'end 2023/07/20
  cnnConnection.BeginTrans
  
  '抓取國內潛在客戶檔之開發日期三年以上(排除研發處D01及專利處P12輸入)及該客戶3年內無往來記錄
  'Modify by Amy 2015/04/27 +不可刪除編號
  'Modify by Amy 2021/10/14 +POC13='75033',因夏慧珠由研發處調至智權部,導致在研發處建的資料會被刪除
  strSql = "SELECT POC01,POC02,POC12,POC13,ST03,MaxCOR02,COR03 From PotCustomer1,STAFF, " & _
                "(Select Max(COR02) As MaxCOR02,SubStr(COR03,1,8) COR03 From  ContactRecord1,PotCustomer1 " & _
                "Where POC02='0' AND POC12 <  " & strSrvDate(1) - 30000 & " And POC01= SubStr (COR03,1,8) Group By  SubStr (COR03,1,8)) " & _
                "Where POC02='0' AND POC13=ST01(+) And ST03 NOT IN ('D01','P12') And POC13<>'75033' And POC12 < " & strSrvDate(1) - 30000 & _
                " And POC01=COR03(+) And (MaxCOR02 is null or MaxCOR02 < " & strSrvDate(1) - 30000 & ")" & _
                IIf(strExclude <> "", " And POC01 not in(" & strExclude & ")", "")
                
  CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 And .RecordCount > 0 Then
         .MoveFirst
         Do While .EOF = False
            '讀取國內潛在客戶逐筆刪除國內潛在客戶檔
            StrSqlB = "Select POC01,POC02,POC03 From PotCustomer1 Where POC01='" & .Fields("POC01") & "'"
            If RsTemp.State <> 0 Then RsTemp.Close
            RsTemp.CursorLocation = adUseClient
            RsTemp.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
            If RsTemp.RecordCount <> 0 And RsTemp.RecordCount > 0 Then
                RsTemp.MoveFirst
                Do While RsTemp.EOF = False
                    StrSQLa = "Delete From PotCustomer1 Where POC01='" & RsTemp.Fields("POC01") & "' and POC02='" & RsTemp.Fields("POC02") & "'"
                    Pub_SeekTbLog StrSQLa, "QPGMR" 'Modify by Amy 2017/09/04 +QPGMR
                    cnnConnection.Execute StrSQLa
                
                    RsTemp.MoveNext
                    DoEvents
                Loop
            End If
            
            
            '讀取國內往來記錄逐筆刪除國內往來記錄
            StrSqlB = "Select COR01,COR02,COR03 From ContactRecord1 Where SubStr(COR03,1,8)='" & .Fields("POC01") & "'"
            If RsTemp.State <> 0 Then RsTemp.Close
            RsTemp.CursorLocation = adUseClient
            RsTemp.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
            
            If RsTemp.RecordCount <> 0 And RsTemp.RecordCount > 0 Then
                RsTemp.MoveFirst
                Do While RsTemp.EOF = False
                    StrSQLa = "Delete From ContactRecord1 Where COR01='" & RsTemp.Fields("COR01") & "'"
                    Pub_SeekTbLog StrSQLa
                    cnnConnection.Execute StrSQLa
                   
                    RsTemp.MoveNext
                    DoEvents
                Loop
            End If
            
            .MoveNext
            DoEvents
         Loop
      End If
   End With
   CheckOC
   cnnConnection.CommitTrans
   Exit Sub


CheckErr:
  cnnConnection.RollbackTrans
End Sub
'2013/05/03 END

'Add By Sindy 2014/6/30 複檢卷宗區檔案狀況
Sub StrMenu12()
Dim tmpnext As String
'Dim ChkDate As String
'Dim errFile As String
'Dim iFiles As Integer
   
On Error GoTo DebugErr
   
   strDate = DBDATE(DateAdd("d", -2, Format(strSrvDate(1), "####/##/##")))
   
   '1.若為電子送件的程序,卷宗區是否還有存放.dwg.pdf
   strSql = "select cp09,cpp02,cp01,cp02,cp03,cp04,cp118" & _
            " From caseprogress,casepaperpdf" & _
            " Where cp27<=" & strDate & " and cp27>=20130801" & _
            " and cp01='P' and cp118 is not null" & _
            " and cp09=cpp01(+)" & _
            " and instr(upper(cpp02),upper('.dwg.pdf'))>0"
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   tmpnext = "已取得資料 SQL..."
   With Rs
      If .RecordCount > 0 Then
         SendMAPIMail "97038", "為電子送件的程序,卷宗區還有存放.dwg.pdf", _
                               "本所案號：" & .Fields("cp01") & "-" & .Fields("cp02") & "-" & .Fields("cp03") & "-" & .Fields("cp04") & vbCrLf & _
                               "總收文號：" & .Fields("cp09") & vbCrLf & _
                               "檔　　名：" & .Fields("cpp02") & vbCrLf & _
                               "電子送件狀態：" & .Fields("cp118"), , True
      End If
   End With
   
   '2.P大陸新案若有.dwg.pdf時,檢查是否有國內案,若有,則P大陸新案的.dwg.pdf刪除
   strSql = "select cp09,cpp02,cp01,cp02,cp03,cp04,cp118,pa09,cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08,cm10,cp10" & _
            " From caseprogress, casepaperpdf, patent, casemap" & _
            " Where cp27<=" & strDate & " and cp27>=20130801 and cp01='P' and cp10 in(" & NewCasePtyList & ")" & _
            " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09='020'" & _
            " and cp09=cpp01(+) and instr(upper(cpp02),upper('.dwg.pdf'))>0" & _
            " and pa01=cm01(+) and pa02=cm02(+) and pa03=cm03(+) and pa04=cm04(+) and cm10='0'"
   Set Rs = New ADODB.Recordset
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   tmpnext = "已取得資料 SQL..."
   With Rs
      If .RecordCount > 0 Then
         SendMAPIMail "97038", "P大陸新案有.dwg.pdf檔存在且有國內案", _
                               "本所案號：" & .Fields("cp01") & "-" & .Fields("cp02") & "-" & .Fields("cp03") & "-" & .Fields("cp04") & vbCrLf & _
                               "總收文號：" & .Fields("cp09") & vbCrLf & _
                               "案件性質：" & .Fields("cp10") & vbCrLf & _
                               "檔　　名：" & .Fields("cpp02") & vbCrLf & _
                               "相關案號：" & .Fields("cm05") & "-" & .Fields("cm06") & "-" & .Fields("cm07") & "-" & .Fields("cm08"), , True
      End If
   End With
   
   Set Rs = Nothing
   Exit Sub
   
DebugErr:
   Set Rs = Nothing
   WLog tmpnext & " " & Err.Description
End Sub

'Private Function SendHTMLMail(stToMails As String, stRef As String, stAttPath As String, bolJp As Boolean) As Boolean
'Dim stSamplePath As String
'Dim stSQL As String
'Dim stSubject As String
'Dim arrToMail, stToMail As String, stToName As String, stMIME As String
'Dim adoRst As New ADODB.Recordset, lngRec As Long
'Dim lngSize As Long
'Dim iFileNo As Integer
'Dim bytes() As Byte
'Dim iErrNo As Integer
'
'   Static stMimeHead As String
'   Static stBoundryTag As String
'   Static stSubjectLead As String
'   Static stFromMail As String
'   Static strFromName As String
'
'   Static stMimeHeadJp As String
'   Static stBoundryTagJp As String
'   Static stSubjectLeadJp As String
'   Static stFromMailJp As String
'   Static strFromNameJp As String
'
'   arrToMail = Split(Trim(Replace(stToMails, vbCrLf, "")), ";") '去除前後的空白和跳行符號
'   stToMail = arrToMail(0) '只寄送第一個信箱
'
'   If bolJp = True Then
'      If stMimeHeadJp = "" Or stBoundryTagJp = "" Then
'         stSamplePath = App.Path & TextPath & "sample.eml"
'         stSQL = "SELECT * FROM MailSchedule,MailScheduleTemplet WHERE ms01=12 and mst01(+)=ms01"
'         If adoRst.State <> adStateClosed Then adoRst.Close
'         With adoRst
'         .CursorLocation = adUseClient
'         .Open stSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
'         If .RecordCount > 0 Then
'            strFromNameJp = "" & .Fields("ms14") '寄件名稱
'            stFromMailJp = "" & .Fields("ms03") '寄件信箱
'            'Modify by Morgan 2009/9/29
'            'stSubjectLeadJp = "" & .Fields("ms02") '主旨
'            stSubjectLeadJp = "Quarterly Report on Refund of Official Fees for Patent Applications in Taiwan" '主旨
'
'            lngSize = Val(.Fields("mst02").Value) '樣本郵件大小
'
'            ReDim bytes(lngSize)
'            bytes() = .Fields("mst03").GetChunk(lngSize)
'            iFileNo = FreeFile
'            If fso.FileExists(stSamplePath) Then
'               Kill stSamplePath
'            End If
'            Open stSamplePath For Binary Access Write As #iFileNo
'            Put #iFileNo, , bytes()
'            Close #iFileNo
'            stMimeHeadJp = GetMimeHead(stSamplePath, stBoundryTagJp)
'         End If
'         End With
'         Set adoRst = Nothing
'      End If
'      stSubject = stSubjectLeadJp & " ( O/Ref: " & stRef & " ) " & strEngDate
'      stMIME = stMimeHeadJp & GetAttMime(stAttPath, stBoundryTagJp)
'   Else
'      If stMimeHead = "" Or stBoundryTag = "" Then
'         stSamplePath = App.Path & TextPath & "sample.eml"
'         stSQL = "SELECT * FROM MailSchedule,MailScheduleTemplet WHERE ms01=11 and mst01(+)=ms01"
'         If adoRst.State <> adStateClosed Then adoRst.Close
'         With adoRst
'         .CursorLocation = adUseClient
'         .Open stSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
'         If .RecordCount > 0 Then
'            strFromName = "" & .Fields("ms14") '寄件名稱
'            stFromMail = "" & .Fields("ms03") '寄件信箱
'            'Modify by Morgan 2009/9/29
'            'stSubjectLead = "" & .Fields("ms02") '主旨
'            stSubjectLead = "Quarterly Report on Refund of Official Fees for Patent Applications in Taiwan"
'
'            lngSize = Val(.Fields("mst02").Value) '樣本郵件大小
'
'            ReDim bytes(lngSize)
'            bytes() = .Fields("mst03").GetChunk(lngSize)
'            iFileNo = FreeFile
'            If fso.FileExists(stSamplePath) Then
'               Kill stSamplePath
'            End If
'            Open stSamplePath For Binary Access Write As #iFileNo
'            Put #iFileNo, , bytes()
'            Close #iFileNo
'            stMimeHead = GetMimeHead(stSamplePath, stBoundryTag)
'         End If
'         End With
'         Set adoRst = Nothing
'      End If
'      stSubject = stSubjectLead & " ( O/Ref: " & stRef & " ) " & strEngDate
'      stMIME = stMimeHead & GetAttMime(stAttPath, stBoundryTag)
'   End If
'
'   stToName = stToMail
'   SendHTMLMail = SendXMail(strFromName, stFromMail, stToName, stToMail, stSubject, stMIME, iErrNo)
'
'End Function
'
'Private Function SendXMail(FromName$, FromMail$, ToName$, ToMail$, Subj$, strMime$, Optional iErrCode As Integer) As Boolean
'Dim strData(0 To 9) As String
'Dim DateNow As String
'Dim SMTP As String
'Dim iRetry As Integer
'Dim stBas64 As String
'
'   '測試
'   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") Then
'      SendXMail = True
'      Exit Function
'      ToMail = "morganljh@hotmail.com"
'   End If
'
'   Call Sleep(2000) '等2秒
'
'On Error GoTo ErrHnd
'
'   iErrCode = 0
'   Result = ""
'   DoEvents
'
'   If InStr(ToMail, "@") = 0 Then
'      ToMail = ToMail & "@taie.com.tw"
'   End If
'
''Modified by Morgan 2012/9/28 統一寄往 192.168.1.15
''   SMTP = "192.168.1.10"
''   'Added by Morgan 2012/9/5 外部信箱改寄往 192.168.1.15
''   If InStr(LCase(ToMail), "@taie.com.tw") = 0 Then
''      SMTP = "192.168.1.15"
''   End If
''   'end 2012/9/5
'   SMTP = "192.168.1.15"
''end 2012/9/28
'
'   strData(1) = "mail from:" + Chr(32) + FromMail + vbCrLf
'   strData(2) = "rcpt to:" + Chr(32) + ToMail + vbCrLf
'
'   stBas64 = ConvertToBase64(FromName, False, False)
'   strData(3) = "From: =?Big5?B?" & stBas64 & "?= <" & FromMail & ">" & vbCrLf
'
'   stBas64 = ConvertToBase64(ToName, False, False)
'   strData(4) = "To: =?Big5?B?" & stBas64 & "?= <" & ToMail & ">" & vbCrLf
'
'   stBas64 = ConvertToBase64(Subj, False, False)
'   strData(5) = "Subject: =?Big5?B?" & stBas64 & "?=" & vbCrLf
'
'   DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " +0800"
'
'   stBas64 = ConvertToBase64(FromName, False, False)
'
'   'strData(6) = "Date:" + Chr(32) + DateNow + vbCrLf & _
'      "Importance: high" & vbCrLf & _
'      "X-Priority: 1" & vbCrLf & _
'      "Return-Receipt-To: =?Big5?B?" & stBas64 & "?= <" & FromMail & ">" & vbCrLf
'
'   strData(6) = "Date:" + Chr(32) + DateNow + vbCrLf
'
'   strData(0) = strData(3) + strData(4) + strData(5) + strData(6)
'
'   strData(9) = strMime
'   If strData(9) = "" Then
'      strData(7) = "MIME-Version: 1.0" & vbCrLf & _
'                   "Content-Type: text/plain;" + vbCrLf & _
'                   "   charset=""big5""" + vbCrLf
'      strData(8) = "testing..." + vbCrLf
'      strData(9) = strData(7) & strData(8)
'   End If
'
'   strData(0) = strData(0) + strData(9)
'
'RetryPoint:
'
'   If Winsock1.State <> sckClosed Then Winsock1.Close
'
'   Winsock1.LocalPort = 0
'   Winsock1.Protocol = sckTCPProtocol
'   Winsock1.RemoteHost = SMTP
'   Winsock1.RemotePort = 25
'   DoEvents
'
'   Winsock1.Connect
'   If Not Response("220") Then
'      Winsock1.Close
'      iErrCode = 1
'      If SMTP <> Server Then SMTP = Server '若連線失敗改連另一台 Added by Morgan 2012/9/28
'      GoTo ERRORMail
'   End If
'
'   DoEvents
'   Winsock1.SendData ("HELO msa.hinet.net" + vbCrLf)
'   If Not Response("250") Then
'      iErrCode = 2
'      GoTo ERRORMail
'   End If
'
'   DoEvents
'   Winsock1.SendData (strData(1))
'   If Not Response("250") Then
'      iErrCode = 3
'      GoTo ERRORMail
'   End If
'
'   DoEvents
'   Winsock1.SendData (strData(2))
'   If Not Response("250") Then
'      iErrCode = 4
'      GoTo ERRORMail
'   End If
'
'   DoEvents
'   Winsock1.SendData ("data" + vbCrLf)
'   If Not Response("354") Then
'      iErrCode = 5
'      GoTo ERRORMail
'   End If
'
'   DoEvents
'   Winsock1.SendData (strData(0) & vbCrLf & "." & vbCrLf)
'   If Not Response("250") Then
'      iErrCode = 6
'      GoTo ERRORMail
'   End If
'
'   DoEvents
'   Winsock1.SendData ("quit" + vbCrLf)
'   If Not Response("221") Then
'      iErrCode = 7
'      GoTo ERRORMail
'   End If
'   Winsock1.Close
'   SendXMail = True
'   Exit Function
'
'ERRORMail:
'   iRetry = iRetry + 1
'   If iRetry < 3 Then
'      GoTo RetryPoint
'   End If
'
'ErrHnd:
'
'End Function
'
'Private Function GetMimeHead(stSamplePath As String, Optional stOutBoundryTag As String) As String
'Dim ts As TextStream
'Dim strLine As String
'Dim bStart As Boolean
'Dim stMIME As String
'Dim iPos As Integer
'
'   stMIME = ""
'   stOutBoundryTag = ""
'   If fso.FileExists(stSamplePath) Then
'      Set ts = fso.OpenTextFile(stSamplePath)
'      bStart = False
'      Do While Not ts.AtEndOfStream
'         strLine = ts.ReadLine
'         If bStart = False Then
'            '開始
'            If InStr(UCase(strLine), UCase("MIME-Version: 1.0")) > 0 Then
'               bStart = True
'            End If
'         End If
'
'         If stOutBoundryTag = "" Then
'            iPos = InStr(UCase(strLine), UCase("boundary="))
'            If iPos > 0 Then
'               '扣除前後的雙引號
'               stOutBoundryTag = Mid(strLine, iPos + 10)
'               stOutBoundryTag = Left(stOutBoundryTag, Len(stOutBoundryTag) - 1)
'            End If
'         End If
'
'         If bStart = True Then
'            '結束
'            If InStr(UCase(strLine), UCase("Content-Disposition: attachment;")) > 0 Then
'               stMIME = stMIME & strLine & vbCrLf
'               Exit Do
'            Else
'               stMIME = stMIME & strLine & vbCrLf
'            End If
'         End If
'      Loop
'      ts.Close
'   End If
'   GetMimeHead = stMIME
'End Function
'
'Private Function GetAttMime(stAttPath As String, stBoundryTag As String) As String
'Dim stMIME As String, sFil64 As String
'Dim ff As Integer, stLine As String
'
'   stMIME = ""
'   If LCase(Right(stAttPath, 4)) = ".txt" Then
'      'stMIME = stMIME & _
'         "Content-Type: text/plain;" & vbCrLf & _
'         "  name=""explanation.txt""" & vbCrLf & _
'         "Content-Disposition: attachment;" & vbCrLf & _
'         "  filename=""explanation.txt""" & vbCrLf
'
'      stMIME = stMIME & "  filename=""Cases for Refund.txt""" & vbCrLf
'      ff = FreeFile()
'      sFil64 = ""
'      Open stAttPath For Input As ff
'      Do While Not EOF(ff)
'         Line Input #ff, stLine
'         sFil64 = sFil64 & stLine & vbCrLf
'      Loop
'      If ff > 0 Then Close #ff
'      stMIME = stMIME & sFil64 & vbCrLf & vbCrLf & cDASH2 & stBoundryTag & cDASH2 & vbCrLf & vbCrLf
'   Else
'      stMIME = stMIME & _
'         "Content-Type: application/octet-stream;" & vbCrLf & _
'         "  name=""explanation.pdf""" & vbCrLf & _
'         "Content-Transfer-Encoding: base64" & vbCrLf & _
'         "Content-Disposition: attachment;" & vbCrLf & _
'         "  filename=""explanation.pdf""" & vbCrLf
'      sFil64 = ConvertToBase64(stAttPath, True, True)
'      stMIME = stMIME & sFil64 & vbCrLf & vbCrLf & cDASH2 & stBoundryTag & cDASH2 & vbCrLf & vbCrLf
'   End If
'   GetAttMime = stMIME
'End Function
'
'Private Function Response(RCode$) As Boolean
'  Sec = 0
'  Timer1.Interval = 200
'  Timer1.Enabled = True
'  Response = True
'
'  Do While Left$(Result, 3) <> RCode
'    DoEvents
'    If Sec > 10 Then
'      If Len(Result) Then
'        'MsgBox "伺服器錯誤！", vbCritical
'        WLog "Mail,伺服器錯誤！", 1
'      Else
'        'MsgBox "伺服器逾時！", vbCritical
'        WLog "Mail,伺服器逾時！", 1
'      End If
'      Response = False
'      Exit Do
'    End If
'  Loop
'
'  Result = ""
'  Timer1.Enabled = False
'End Function

Private Sub Timer1_Timer()
   Sec = Sec + 1
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
   Winsock1.GetData Result
End Sub

'Add by Lydia 2015/04/09 每月顧問到期通知　（若有修改，請檢查Law.frm073008）
'顧問到期通知直接mail給個人
'每月1日寄發個人, 到期日期條件: 當月1日至隔月月底
Private Sub StrMenu13()
   Dim rsA As New ADODB.Recordset
   Dim stSQL As String, stDate As String, stDate2 As String
   Dim ff13 As Integer
   Dim TempFileName As String, strTemp(8) As String, strTFName As String
   Dim strTo As String, strApatch As String
On Error GoTo ErrHandle

   Call PUB_KillTempFile(Mid(TextPath & "顧問到期明細表*.*", 2))    'Added by Lydia 2020/08/19 清除舊檔

   stDate = strSrvDate(1): stDate2 = stDate
   '每月1日
   stDate = Mid(stDate2, 1, 6) & "01"
   '隔月月底
   stDate2 = GetLastDay(CompDate(1, 1, stDate2))
   strTFName = "顧問到期明細表" & ChangeWStringToTString(stDate) & "-" & ChangeWStringToTString(stDate2)
   'Modified by Lydia 2015/09/01 移除 bug
   'Modified by Lydia 2016/09/13 CP57 IS NULL=> CP159=0 ; cp27 is null=> CP159
   'Memo by Lydia 2021/09/01 還未有案源資料
   'Modified by Lydia 2021/09/01 +列出CP12
   'Modified by Lydia 2021/11/02 改成判斷沒案源檔抓CP13為收件人; ex.LA-999999
   'stSQL = "SELECT s1.st01 as B00,DECODE(CP12,A0901,A0902) B01,DECODE(CP13,s1.ST01,s1.ST02) B02,HC01||'-'||HC02||DECODE(HC03,'0','','-'||HC03)||DECODE(HC04,'00','','-'||HC04) B03" & _
           ",SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2) B04,SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2) B05" & _
           ",SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2) B06,CP16,HC05,DECODE(HC05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))) B07" & _
           ",C1.CU16 as CU16_1,C1.CU79 as CU79_1, HC01, HC02, HC03, HC04,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))) B08,C2.CU16 as CU16_2,C2.CU79 as CU79_2,CP12, CP13" & _
           " FROM HIRECASE,CASEPROGRESS,STAFF s1,CUSTOMER C1,CUSTOMER C2,ACC090" & _
           " WHERE HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 AND CP13=s1.st01(+) AND CP12=A0901(+) AND (SUBSTR(HC05,1,8)=C1.CU01(+) AND SUBSTR(HC05,9,1)=C1.CU02(+))" & _
           " AND (SUBSTR(HC24,1,8)=C2.CU01(+) AND SUBSTR(HC24,9,1)=C2.CU02(+))" & _
           " AND (CP54 BETWEEN " & stDate & " AND " & stDate2 & ") AND CP10='0' and substr(CP12,1,1)='S'" & _
           " AND CP158=0 AND CP159=0"
   stSQL = "SELECT s1.st01 as B00,DECODE(CP12,A0901,A0902) B01,DECODE(CP13,s1.ST01,s1.ST02) B02,HC01||'-'||HC02||DECODE(HC03,'0','','-'||HC03)||DECODE(HC04,'00','','-'||HC04) B03" & _
           ",SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2) B04,SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2) B05" & _
           ",SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2) B06,CP16,HC05,DECODE(HC05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))) B07" & _
           ",C1.CU16 as CU16_1,C1.CU79 as CU79_1, HC01, HC02, HC03, HC04,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))) B08,C2.CU16 as CU16_2,C2.CU79 as CU79_2,CP12, CP13" & _
           " FROM HIRECASE,CASEPROGRESS,STAFF s1,CUSTOMER C1,CUSTOMER C2,ACC090" & _
           " WHERE HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 AND CP13=s1.st01(+) AND CP12=A0901(+) AND (SUBSTR(HC05,1,8)=C1.CU01(+) AND SUBSTR(HC05,9,1)=C1.CU02(+))" & _
           " AND (SUBSTR(HC24,1,8)=C2.CU01(+) AND SUBSTR(HC24,9,1)=C2.CU02(+))" & _
           " AND (CP54 BETWEEN " & stDate & " AND " & stDate2 & ") AND CP10='0' and cp09 not in (Select Los06 From Lawofficesource where los07 is null and los06 is not null) " & _
           " AND CP158=0 AND CP159=0"
   'Added by Lydia 2021/09/01 案源的資料應改發給介紹人
   'Modified by Lydia 2021/11/02  判斷有案源
   'stSQL = stSQL & "Union SELECT S1.ST01 AS B00,A0902 B01,S1.ST02 B02,HC01||'-'||HC02||DECODE(HC03,'0','','-'||HC03)||DECODE(HC04,'00','','-'||HC04) B03" & _
           ",SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2) B04,SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2) B05" & _
           ",SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2) B06,CP16,HC05,DECODE(HC05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))) B07" & _
           ",C1.CU16 as CU16_1,C1.CU79 as CU79_1, HC01, HC02, HC03, HC04,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))) B08,C2.CU16 as CU16_2,C2.CU79 as CU79_2,s1.ST15 as CP12,substr(los04,1,5) as CP13" & _
           " FROM HIRECASE,CASEPROGRESS,STAFF s1,CUSTOMER C1,CUSTOMER C2,ACC090,LAWOFFICESOURCE" & _
           " WHERE HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 AND (SUBSTR(HC05,1,8)=C1.CU01(+) AND SUBSTR(HC05,9,1)=C1.CU02(+))" & _
           " AND (SUBSTR(HC24,1,8)=C2.CU01(+) AND SUBSTR(HC24,9,1)=C2.CU02(+))" & _
           " AND (CP54 BETWEEN " & stDate & " AND " & stDate2 & ") AND CP10='0' and substr(CP12,1,1)<>'S' " & _
           " AND CP158=0 AND CP159=0 AND CP09=LOS06(+) AND SUBSTR(LOS04,1,5)=S1.ST01(+) AND S1.ST15=A0901(+)"
   stSQL = stSQL & "Union SELECT S1.ST01 AS B00,A0902 B01,S1.ST02 B02,HC01||'-'||HC02||DECODE(HC03,'0','','-'||HC03)||DECODE(HC04,'00','','-'||HC04) B03" & _
           ",SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2) B04,SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2) B05" & _
           ",SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2) B06,CP16,HC05,DECODE(HC05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))) B07" & _
           ",C1.CU16 as CU16_1,C1.CU79 as CU79_1, HC01, HC02, HC03, HC04,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))) B08,C2.CU16 as CU16_2,C2.CU79 as CU79_2,s1.ST15 as CP12,substr(los04,1,5) as CP13" & _
           " FROM HIRECASE,CASEPROGRESS,STAFF s1,CUSTOMER C1,CUSTOMER C2,ACC090,LAWOFFICESOURCE" & _
           " WHERE HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04 AND (SUBSTR(HC05,1,8)=C1.CU01(+) AND SUBSTR(HC05,9,1)=C1.CU02(+))" & _
           " AND (SUBSTR(HC24,1,8)=C2.CU01(+) AND SUBSTR(HC24,9,1)=C2.CU02(+))" & _
           " AND (CP54 BETWEEN " & stDate & " AND " & stDate2 & ") AND CP10='0' and los06 is not null " & _
           " AND CP158=0 AND CP159=0 AND CP09=LOS06(+) AND SUBSTR(LOS04,1,5)=S1.ST01(+) AND S1.ST15=A0901(+)"
   'Modified by Lydia 2021/09/01
   'stSQL = stSQL & " ORDER BY 2,1,CP01||CP02||CP03||CP04"
   stSQL = stSQL & " ORDER BY CP12,CP13,B03"
   If rsA.State <> adStateClosed Then rsA.Close

    Set rsA = New ADODB.Recordset
    rsA.CursorLocation = adUseClient
    rsA.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
    With rsA
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If TempFileName <> strTFName & "_" & Trim(.Fields("B00")) & Trim(.Fields("B02")) Then
                    strExc(10) = TempFileName: strApatch = ""
                    TempFileName = strTFName & "_" & Trim(.Fields("B00")) & Trim(.Fields("B02")) '+員工編號，姓名
                    TempFileName = PUB_UniToBIG5(TempFileName, "F") 'Added by Lydia 2022/03/28 員工名稱有Unicode
                    If ff13 > 0 Then
                       Close #ff13
                       If Len(Trim(strTemp(0))) > 0 Then
                          strTo = strTemp(0) '前一筆記錄的收信人
                          strApatch = App.path & TextPath & strExc(10) & ".txt"  '前一個業務區附件
                          SendMAPIMail strTo, strTFName, vbCrLf & vbCrLf & "***詳情請見附件***" & vbCrLf & vbCrLf, strApatch
                       End If
                    End If
                    ff13 = FreeFile
                    Open App.path & TextPath & TempFileName & ".txt" For Output As ff13
                    strExc(0) = convForm(" ", 30)
                    strExc(1) = convForm(" ", 90)
                    Print #ff13, strExc(0) & strTFName
                    Print #ff13, "列印日期：" & ChangeWStringToTString(stDate)
                    Print #ff13, "智權人員 顧問案號        顧問期間              收文日期        金額 客戶編號  當事人名稱                     備      註"
                    Print #ff13, "======== =============== ===================== ========= ========== ========= ============================== =============================="
                End If
                strTemp(0) = "" & .Fields("B00")
                strTemp(1) = convForm("" & .Fields("B02"), 8)
                strTemp(2) = convForm("" & .Fields("B03"), 15)
                strTemp(3) = convForm(Trim("" & .Fields("B04")) & " - " & Trim("" & .Fields("B05")), 21)
                strTemp(4) = convForm(Trim("" & .Fields("B06")), 9)
                strTemp(5) = PUB_StrToStr(.Fields("cp16"), 10, True, True)
                strTemp(6) = convForm("" & .Fields("hc05"), 9)
                strTemp(7) = convForm(PUB_StrToStr(CheckStr("" & .Fields("B07")), 30), 30)
                strTemp(8) = convForm(PUB_StrToStr(CheckStr("" & .Fields("CU79_1")), 30), 30)
                
                If Left(Trim(.Fields("hc05")), 6) = "X65299" Then
                   strTemp(2) = convForm(PUB_StrToStr(Trim(strTemp(2)) & "（謝）", 15), 15) '顧問案號
                   '改用第２當事人
                   strTemp(6) = convForm("" & .Fields("hc24"), 9) '客戶代號
                   strTemp(7) = convForm(PUB_StrToStr(CheckStr("" & .Fields("B08")), 30), 30) '客戶名稱
                   strTemp(8) = convForm(PUB_StrToStr(CheckStr("" & .Fields("CU79_2")), 30), 30) '備註
                End If
            
                Print #ff13, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & _
                      " " & strTemp(5) & " " & strTemp(6) & " " & strTemp(7) & " " & strTemp(8)
                
                .MoveNext
            Loop
        End If
        If TempFileName <> "" Then
            Close ff13
            strTo = strTemp(0) '最後一筆記錄的收信人
            strApatch = App.path & TextPath & TempFileName & ".txt"
            SendMAPIMail strTo, strTFName, vbCrLf & vbCrLf & "***詳情請見附件***" & vbCrLf & vbCrLf, strApatch
        End If
    End With
    
    TempFileName = ""
    Set rsA = Nothing
ErrHandle:
   If Err.Number <> 0 Then
      WLog "每月顧問到期通知:" & Err.Description
   End If
   Set rsA = Nothing
End Sub

'Add By Amy 2016/06/15 上個月個人客戶資料修改通知
'業務區為F字頭者不抓；智權部非區主管(非A0908)資料，寄給各區區主管A0908；
'各區區主管資料寄給特殊設定「全所智權部主管」-Add by Amy 1091110;
'L01、L02、P31資料寄給桂所長(112年退休前)；業務區為P1字頭者寄給王副總；業務區為P2字頭者寄給林純貞(110年退休前)；
'非智權部區主管A0908資料、桂所長、王副總、林純貞、全所智權部主管(Add by Amy 1091110)及其他非上述人員資料、皆寄給總經理/楊監察人(112年退休前)；
Sub StrMenu14()
    Dim strQ As String
    Dim RsQ As New ADODB.Recordset
    Dim strApatch As String, strFileN As String, strAttnF As String
    Dim strTemp(3) As String, strTemp2(2) As String
    Dim OldSt15 As String, OldCU13 As String, stDate As String, strFieldN As String
    Dim OldDeptMan As String '部門主管
    Dim strTo As String, strContent As String
    Dim Txt_Head As Boolean, bolData1 As Boolean, bolData2 As Boolean, bolData3 As Boolean 'F1/2/3是否有資料
    Dim F1 As Integer, F2 As Integer, F3 As Integer, ii As Integer, jj As Integer, kk As Integer
    Dim intInStr_S As Integer, intGetStr As Integer
    Dim strCU_Not() As String '不需列的欄位
    '需列欄位
    Dim strCU As Variant
    Const CUFields As String = "16,17,18,19,20,22,30,31,103,116,117,118,125,127"
    Dim bolData4 As Boolean, strAllSMan As String, F4 As Integer 'Add by Amy 2020/11/10 F4是否有資料(智權部區主管資料)/全所智權部主管/檔案4
    
On Error GoTo ErrHand
    stDate = ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(DBDATE(Left(strSrvDate(1), 6) & "01"))))
    strApatch = App.path & TextPath
    strFileN = Left(stDate, 4) & "年" & Mid(stDate, 5, 2) & "月個人客戶資料修改通知.txt"
    strAllSMan = Pub_GetSpecMan("全所智權部主管") 'Add by Amy 2020/11/10
    
    strCU = Split(CUFields, ",")
    strQ = "Select column_name From user_tab_columns Where table_name = 'CUSTOMER' " & _
                "Order by  to_number(substr(column_name,3))"
    If RsQ.State = adStateOpen Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        ReDim strCU_Not(RsQ.RecordCount - UBound(strCU) - 2)
        jj = LBound(strCU): kk = 0
        RsQ.MoveFirst
        For ii = 0 To RsQ.RecordCount - 1
            If jj <= UBound(strCU) Then
                If UCase(RsQ.Fields("column_name")) = "CU" & strCU(jj) Then
                    jj = jj + 1
                Else
                    strCU_Not(kk) = Replace(UCase(RsQ.Fields("column_name")), "CU", "")
                    kk = kk + 1
                End If
            Else
                strCU_Not(kk) = Replace(UCase(RsQ.Fields("column_name")), "CU", "")
                kk = kk + 1
            End If
            RsQ.MoveNext
        Next ii
    End If
    RsQ.Close
    
    strContent = "資料如附件！" & vbCrLf & vbCrLf & "列印注意事項：" & vbCrLf & _
         vbCrLf & String(4, "　") & "1.利用筆記本開啟附件" & _
         vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
         vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
         vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
         vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
         vbCrLf & String(4, "　") & "6.選擇<橫印>"
    'Modified by Morgan 2019/9/19 68009 -> 68006
    'modify by sonia 2020/3/4 68006->94007;69009
    'Modify by Amy 2021/11/15 林經理(69008)退休,其發信改抓PUB_GetManP2,寫入暫存檔
'     strQ = "Select A0902,S.st02,M.st02||' '||SqlDatet(DL07),NVL(CU04,Decode(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as CUName,DL09" & _
'                ",SubStr(DL09,(InStr(UPPER(DL09),'CU01')+6),8) CU1,SubStr(DL09,(InStr(UPPER(DL09),'CU02')+6),1) CU2,CU13,S.ST15" & _
'                ",Decode(SubStr(S.ST15,1,1), 'S',A0908,Decode(SubStr(S.ST15,1,2), 'P1', '71011', 'P2', '69008',Decode(S.ST15, 'P31', '76012', 'L01', '76012','L02', '76012','94007'))) as DeptManNo " & _
'                "From DML_LOG,Staff S,Staff M,Customer,Acc090 " & _
'                "Where DL07>=" & stDate & " And DL07<=" & Left(stDate, 6) & "31 And UPPER(DL10)='CUSTOMER' And SubStr(DL09,2,2)='修改' " & _
'                "And S.ST15=A0901(+) And CU13=S.st01(+) And DL06=M.st01(+) And (M.st03<>'M51' Or (M.st03='M51' and M.st01='74001' )) " & _
'                "And SubStr(S.ST15,1,1)<>'F' And CU82<>CU85 And (DL12='個人客戶資料修改(frm210101_1)' Or DL12='客戶基本資料維護(frm140401)') " & _
'                "And SubStr(DL09,(InStr(UPPER(DL09),'CU01')+6),8)=CU01(+) And SubStr(DL09,(InStr(UPPER(DL09),'CU02')+6),1)=CU02(+) " & _
'                "Order by S.ST15,S.ST06,CU13, SubStr(DL09,(InStr(UPPER(DL09),'CU01')+6),8)|| SubStr(DL09,(InStr(UPPER(DL09),'CU02')+6),1),DL07,DL08 "
    strQ = "Delete Rab14 "
    cnnConnection.Execute strQ
    'Modify by Amy 2022/12/30 原:DL12='個人客戶資料修改(frm210101_1)' Or DL12='客戶基本資料維護(frm140401)' ,由於 2019年有改 frm210101_1表單名稱拿掉「個人」造成資料都沒抓到
    'Added by Lydia 2023/04/24 修改王副總退休之相關控制
    If strSrvDate(1) >= "20230511" Then
        'Modified by Morgan 2025/6/11
        'strExc(0) = "73022"
        strExc(0) = Pub_GetSpecMan("專利處特定編號")
        'end 2025/6/11
    Else
        strExc(0) = "71011"
    End If
    'end 2023/04/24
    'Modified by Lydia 2023/04/24 '71011'=> " & CNULL(strExc(0)) & "
    'Modify by Amy 2023/07/10 桂所長退休(原:Decode(S.ST15, 'P31', '76012', 'L01', '76012','L02', '76012','94007')) 改抓a0908
    strQ = "Insert into Rab14 (SalesName,UpdName,DL07,CuName,DL09,cu01,cu02,cu13,SalesDept,DeptManNo,DL06,DL08,SalesSt06,SalesSt05) " & _
                "Select S.st02,M.st02,DL07,NVL(CU04,Decode(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) as CUName,DL09" & _
                ",SubStr(DL09,(InStr(UPPER(DL09),'CU01')+6),8) CU1,SubStr(DL09,(InStr(UPPER(DL09),'CU02')+6),1) CU2,CU13,S.ST15" & _
                ",Decode(SubStr(S.ST15,1,1), 'S',A0908,Decode(SubStr(S.ST15,1,2), 'P1', " & CNULL(strExc(0)) & " , 'P2', '69008',Decode(S.ST15, 'P31', a0908, 'L01', a0908,'L02', a0908,'94007'))) as DeptManNo " & _
                ",DL06,DL08,S.ST06,S.St05 " & _
                "From DML_LOG,Staff S,Staff M,Customer,Acc090 " & _
                "Where DL07>=" & stDate & " And DL07<=" & Left(stDate, 6) & "31 And UPPER(DL10)='CUSTOMER' And SubStr(DL09,2,2)='修改' " & _
                "And S.ST15=A0901(+) And CU13=S.st01(+) And DL06=M.st01(+) And (M.st03<>'M51' Or (M.st03='M51' and M.st01='74001' )) " & _
                "And SubStr(S.ST15,1,1)<>'F' And CU82<>CU85 And (instr(Upper(dl12),Upper('frm210101_1'))>0 Or instr(Upper(dl12),Upper('frm140401'))>0) " & _
                "And SubStr(DL09,(InStr(UPPER(DL09),'CU01')+6),8)=CU01(+) And SubStr(DL09,(InStr(UPPER(DL09),'CU02')+6),1)=CU02(+) " & _
                "Order by S.ST15,S.ST06,CU13, SubStr(DL09,(InStr(UPPER(DL09),'CU01')+6),8)|| SubStr(DL09,(InStr(UPPER(DL09),'CU02')+6),1),DL07,DL08 "
    cnnConnection.Execute strQ
    '抓取林經理(69008) 資料,更新欲發信人員
    strQ = "Select cu13,SalesSt05,cu01,cu02 From Rab14 Where DeptManNo='69008' "
    If RsQ.State = adStateOpen Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    With RsQ
         If .RecordCount > 0 Then
            RsQ.MoveFirst
            Do While .EOF = False
                If Left("" & RsQ.Fields("cu13"), 3) = "MCT" Then
                    strTemp(0) = "96003" '沈佳穎
                Else
                    Select Case "" & RsQ.Fields("SalesSt05")
                        Case "93"
                            strTemp(0) = "79041" '林桂英
                        Case "95"
                            strTemp(0) = "87027" '林嘉雯
                        Case "97"
                            strTemp(0) = "86048" '林承慧
                        Case Else
                           'Modify by Amy 2024/04/01 原:A2004,林純真經理退休後改程式,怕未有例外未列示,故先寄A2004,因後來智權人員增加P2006
                           '1130401跑3月資料時,因新光金融控股股份有限公司(X82973020) 智權人員為P2006 (部門:P20) 其st05為SA,無主管可寄,秀玲說改寄江協理
                            strTemp(0) = "98020"
                    End Select
                End If
                strTemp(1) = "Update Rab14 Set DeptManNo='" & strTemp(0) & "' Where cu01='" & RsQ.Fields("cu01") & "' And cu02='" & RsQ.Fields("cu02") & "' "
                cnnConnection.Execute strTemp(1)
                RsQ.MoveNext
            Loop
         End If
    End With
    'P2部門,且業務人員=發信人員,則發給江協理
    strQ = "Update Rab14 Set DeptManNo='98020' Where SubStr(SalesDept,1,2)='P2' And CU13=DeptManNo "
    cnnConnection.Execute strQ
    
    '讀取資料
    strQ = "Select A0902,SalesName,UpdName||' '||SqlDatet(DL07),CUName,DL09,CU01 as CU1,CU02 as CU2,CU13,SalesDept as ST15,DeptManNo " & _
                "From Rab14,Acc090 " & _
                "Where SalesDept=A0901(+) " & _
                "Order by SalesDept,SalesSt06,DeptManNo,CU13,CU01||CU02,DL07,DL08 "
    'end 2021/11/15
    If RsQ.State = adStateOpen Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    With RsQ
         If .RecordCount > 0 Then
            '*** To 林總/112年起 B1015(林岱嫻特助) (原:112年前楊監察人/何主秘)檔案
            F2 = FreeFile
            Open strApatch & Left(stDate, 4) & "年" & Mid(stDate, 5, 2) & "月客戶資料修改通知.txt" For Output As F2
            Print #F2, "                                             " & Mid(strFileN, 1, Len(strFileN) - 4)
            Print #F2, "列印日期：" & ChangeWStringToTString(strSrvDate(1))
            Print #F2, "業務區       智權人員   修改人員及日期       客戶名稱              欄位名稱                修改後內容 "
            Print #F2, "============ ========== ==================== ==================== ======================== ================================================== "
            '*** End To 林總/112年起 B1015(林岱嫻特助) (原:112年前楊監察人/何主秘)檔案
            '*** To 桂所長檔案
            F3 = FreeFile
            Open strApatch & Left(stDate, 4) & "年" & Mid(stDate, 5, 2) & "月客戶資料修改通知-法務.txt" For Output As F3
            Print #F3, "                                             " & Mid(strFileN, 1, Len(strFileN) - 4)
            Print #F3, "列印日期：" & ChangeWStringToTString(strSrvDate(1))
            Print #F3, "業務區       智權人員   修改人員及日期       客戶名稱              欄位名稱                修改後內容 "
            Print #F3, "============ ========== ==================== ==================== ======================== ================================================== "
            '*** End To桂所長檔案
            '*** To 全所智權部主管檔案
            F4 = FreeFile
            Open strApatch & Left(stDate, 4) & "年" & Mid(stDate, 5, 2) & "月客戶資料修改通知-智權.txt" For Output As F4
            Print #F4, "                                             " & Mid(strFileN, 1, Len(strFileN) - 4)
            Print #F4, "列印日期：" & ChangeWStringToTString(strSrvDate(1))
            Print #F4, "業務區       智權人員   修改人員及日期       客戶名稱              欄位名稱                修改後內容 "
            Print #F4, "============ ========== ==================== ==================== ======================== ================================================== "
            '*** End To 全所智權部主管檔案
            .MoveFirst
            
            Do While .EOF = False
                '產生資料寄信-區主管
                If (OldSt15 <> "" & .Fields("st15") And Left(OldSt15, 1) = "S" And OldCU13 <> OldDeptMan) Or ((Left(OldSt15, 2) = "P1" Or Left(OldSt15, 2) = "P2") And OldDeptMan <> .Fields("DeptManNo")) Then
                    If F1 > 0 Then Close #F1
                    strAttnF = strApatch & strFileN
                    'Added by Lydia 2023/04/24 修改王副總退休之相關控制
                    If strSrvDate(1) < "20230511" And OldDeptMan = "71011" Then
                         strTo = OldDeptMan & ";73022"
                    Else
                    'end 2023/04/24
                         strTo = OldDeptMan
                    End If 'Added by Lydia 2023/04/24
                    If bolData1 = True Then
                        SendMAPIMail strTo, Mid(strFileN, 1, Len(strFileN) - 4), strContent, strAttnF
                    End If
                    Txt_Head = False: bolData1 = False
                End If
                If Txt_Head = False Then
                    F1 = FreeFile
                    Open strApatch & strFileN For Output As F1
                    Print #F1, "                                             " & Mid(strFileN, 1, Len(strFileN) - 4)
                    Print #F1, "列印日期：" & ChangeWStringToTString(strSrvDate(1))
                    Print #F1, "業務區       智權人員   修改人員及日期       客戶名稱              欄位名稱                修改後內容 "
                    Print #F1, "============ ========== ==================== ==================== ======================== ================================================== "
                    Txt_Head = True
                End If
                
                For ii = 0 To UBound(strTemp)
                    strTemp(ii) = "" & .Fields(ii)
                Next ii
                '修改內容
                For ii = 0 To UBound(strTemp2)
                    If ii = 0 Then
                        strTemp2(ii) = "" & .Fields("DL09")
                    Else
                        strTemp2(ii) = ""
                    End If
                Next ii
                strTemp(0) = convForm(CheckStr(strTemp(0)), 12) '業務區
                strTemp(1) = convForm(CheckStr(strTemp(1)), 10) '智權人員
                strTemp(2) = convForm(CheckStr(strTemp(2)), 20) '修改人員及日期
                strTemp(3) = convForm(CheckStr(strTemp(3)), 20) '客戶名稱
                '解析修改內容
                strTemp2(0) = Mid(strTemp2(0), InStr(strTemp2(0), "；") + 1)
                '取代不需列的欄位
                For ii = LBound(strCU_Not) To UBound(strCU_Not)
                    If InStr(UCase(strTemp2(0)), "CU" & strCU_Not(ii) & "[") > 0 Then
                        intInStr_S = InStr(UCase(strTemp2(0)), "CU" & strCU_Not(ii) & "[")
                        intGetStr = InStr(Mid(strTemp2(0), intInStr_S), "];") + 1
                        strTemp2(0) = Replace(strTemp2(0), Mid(strTemp2(0), intInStr_S, intGetStr), "")
                    End If
                    If strTemp2(0) = MsgText(601) Then Exit For
                Next ii
                
                If strTemp2(0) <> MsgText(601) Then
                    For ii = LBound(strCU) To UBound(strCU)
                        strFieldN = ""
                        If strTemp2(0) = MsgText(601) Then Exit For
                        '需列的欄位
                        If InStr(UCase(strTemp2(0)), "CU" & strCU(ii) & "[") > 0 Then
                            Select Case Val(strCU(ii))
                                '電話
                                Case 16 To 17
                                    strFieldN = "電話" & Val(strCU(ii)) Mod 15
                                 '傳真
                                Case 18 To 19
                                    strFieldN = "傳真" & Val(strCU(ii)) Mod 17
                                '信箱
                                Case 20, 116 To 118
                                    strFieldN = "E-MAIL代表信箱"
                                    If Val(strCU(ii)) <> 20 Then
                                        strFieldN = "E-MAIL其他信箱" & Val(strCU(ii)) Mod 115
                                    End If
                                'MOBILE PHONE
                                Case 22
                                    strFieldN = "MOBILE PHONE"
                                '聯絡地址
                                Case 30 To 31
                                    strFieldN = "聯絡地址"
                                    If Val(strCU(ii)) = 30 Then strFieldN = strFieldN & "郵遞區號"
                                '公司負責人英文名稱
                                Case 103
                                    strFieldN = "公司負責人英文名稱"
                                '業務備註
                                Case 125
                                    strFieldN = "業務備註"
                                '預設接洽人編號
                                Case 127
                                    strFieldN = "預設接洽人"
                            End Select
                            
                            intInStr_S = InStr(strTemp2(0), "=>") + 2
                            intGetStr = InStr(strTemp2(0), "];") - intInStr_S
                            strTemp2(1) = Replace(Replace(Mid(strTemp2(0), intInStr_S, intGetStr), "NULL", "取消"), "null", "取消")
                            '預設接洽人編號
                            If Val(strCU(ii)) = 127 Then
                                strTemp2(2) = strTemp2(1)
                                strTemp2(1) = strTemp2(1) & PUB_GetContact("" & .Fields("CU1"), strTemp2(2))
                            End If
                                
                            strFieldN = convForm(CheckStr(strFieldN), 24)
                            strTemp2(1) = convForm(CheckStr(strTemp2(1)), 50)
                            
                            'Add by Amy 2020/11/10 +智權部區主管資料且不是全所智權部主管自己的資料,寄給全所智權部主管(系統特殊設定)
                            'Modify by Amy 2022/12/30 +Or "" & .Fields("cu13") < "6",排除區編號,因區編號(區主管能改)要寄全所智權部主管
                            If Left("" & .Fields("st15"), 1) = "S" And ("" & .Fields("cu13") = "" & .Fields("DeptManNo") Or "" & .Fields("cu13") < "6") And "" & .Fields("cu13") <> strAllSMan Then
                            '寄給「全所智權部主管」
                                If bolData4 = False Then bolData4 = True 'F4 有資料
                                Print #F4, strTemp(0) & " " & strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strFieldN & _
                                            " " & strTemp2(1)
                            'Modify by Amy 2023/07/10 桂所長退休,改寄a0908
                            '寄給「桂所長76012」
                            'ElseIf "" & .Fields("cu13") <> "76012" And "" & .Fields("DeptManNo") = "76012" Then
                            ElseIf "" & .Fields("cu13") <> .Fields("DeptManNo") And "" & .Fields("DeptManNo") = GetDeptMan("L01") Then
                                If bolData3 = False Then bolData3 = True 'F3 有資料
                                Print #F3, strTemp(0) & " " & strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strFieldN & _
                                            " " & strTemp2(1)
                            'Modified by Morgan 2019/9/19 68009->68006
                            'modify by sonia 2020/3/4 68006->94007;69009
                            'Modify by Amy 2020/11/09 +全所智權部主管
                            ElseIf "" & .Fields("cu13") = "" & .Fields("DeptManNo") Or "" & .Fields("cu13") = strAllSMan Or "" & .Fields("DeptManNo") = Pub_GetSpecMan("總經理員工編號") Then
                            '寄給「總經理」
                                If bolData2 = False Then bolData2 = True 'F2 有資料
                                Print #F2, strTemp(0) & " " & strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strFieldN & _
                                            " " & strTemp2(1)
                            Else
                            '寄給「區主管」
                                'Memo by Amy 2022/12/30 發現 20221129 林青祺修改智權人員邱崴靖的客戶資料,發給林青祺(規則不變以cu13為主)-秀玲
                                If bolData1 = False Then bolData1 = True 'F1 有資料
                                Print #F1, strTemp(0) & " " & strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strFieldN & _
                                            " " & strTemp2(1)
                            End If
                            intInStr_S = InStr(UCase(strTemp2(0)), "CU" & strCU(ii) & "[")
                            intGetStr = InStr(Mid(strTemp2(0), intInStr_S), "];") + 2 - intInStr_S
                            strTemp2(0) = Replace(strTemp2(0), Mid(strTemp2(0), intInStr_S, intGetStr), "")
                        End If
                    Next ii
                End If
                
                OldSt15 = "" & .Fields("ST15")
                OldCU13 = "" & .Fields("CU13")
                OldDeptMan = "" & .Fields("DeptManNo")
                .MoveNext
            Loop
        End If
    End With
    If F1 > 0 Then Close #F1
    If F2 > 0 Then Close #F2
    If F3 > 0 Then Close #F3
    If F4 > 0 Then Close #F4 'Add by Amy 2020/11/10
    If bolData1 = True Then
        'Added by Lydia 2023/04/24 修改王副總退休之相關控制
        If strSrvDate(1) < "20230511" And OldDeptMan = "71011" Then
             strTo = OldDeptMan & ";73022"
        Else
        'end 2023/04/24
             strTo = OldDeptMan
        End If 'Added by Lydia 2023/04/24
        SendMAPIMail strTo, Mid(strFileN, 1, Len(strFileN) - 4), strContent, strAttnF
    End If
    If bolData2 = True Then
        strAttnF = strApatch & Left(stDate, 4) & "年" & Mid(stDate, 5, 2) & "月客戶資料修改通知.txt"
        'Modified by Morgan 2019/9/19 68009->68006
        'modify by sonia 2020/3/4 68006->94007(林總);69009(楊監察人)
        'Modify by Amy 69009->B1015(林岱嫻特助)
        strTo = "94007;B1015"
        SendMAPIMail strTo, Left(stDate, 4) & "年" & Mid(stDate, 5, 2) & "月客戶資料修改通知", strContent, strAttnF
    End If
    If bolData3 = True Then
        strAttnF = strApatch & Left(stDate, 4) & "年" & Mid(stDate, 5, 2) & "月客戶資料修改通知-法務.txt"
        strTo = GetDeptMan("L01") 'Modify by Amy 2023/07/10 原:"76012" '桂所長
        SendMAPIMail strTo, Left(stDate, 4) & "年" & Mid(stDate, 5, 2) & "月客戶資料修改通知", strContent, strAttnF
    End If
    'Add by Amy 2020/11/10 區主管資料寄給「全所智權部主管」
    If bolData4 = True Then
        strAttnF = strApatch & Left(stDate, 4) & "年" & Mid(stDate, 5, 2) & "月客戶資料修改通知-智權.txt"
        strTo = strAllSMan
        SendMAPIMail strTo, Left(stDate, 4) & "年" & Mid(stDate, 5, 2) & "月客戶資料修改通知", strContent, strAttnF
    End If
        
    Set RsQ = Nothing
    Exit Sub
    
ErrHand:
    WLog "上個月個人客戶資料修改通知:" & Err.Description, 1
    If F1 > 0 Then Close #F1
    If F2 > 0 Then Close #F2
    If F3 > 0 Then Close #F3
    If F4 > 0 Then Close #F4 'Add by Amy 2021/11/15
    Set RsQ = Nothing
End Sub

'Mark by Amy 2015/0810 搬至AutoBatchDay 改每月1日、16日發
''Add by Amy 2015/05/26 FCT,T,TB,TC,TD,TF,TM,TR,TS,TT每月自動發催審表給承辦人
'Private Sub StrMenu14()
'    Dim RsQ As New ADODB.Recordset
'    Dim strQ As String, strQ1 As String
'    Dim ff14 As Integer
'    Dim strFileN As String, strTemp(11) As String, strDate(1) As String
'    Dim Is1705or310 As Boolean
'    Dim Txt_Head As Boolean
'    Dim strTo As String, strApatch As String
'    Dim ii As Integer
'
'On Error GoTo ErrHand
'    strFileN = Val(Left(strSrvDate(1), 6)) - 191100 & "催審表"
'    strApatch = App.path & TextPath & strFileN & ".txt"
'
'    strDate(0) = Left(strSrvDate(1), 6) & "01"
'    strDate(1) = Left(strSrvDate(1), 6) & "31"
'
'    strQ = "Select np01,sqldatet(np08) NP08,NP10,tm01 TM01,tm02 TM02,tm03 TM03,tm04 TM04,tm10 TM10,Nvl(tm15,tm12) TM15,tm23 TM23,Nvl(tm05,Nvl(tm06,tm07)) CaseN " & _
'            "From NextProgress,TradeMark Where NP02 in ('TF','T','FCT',' ') And NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) " & _
'            "And tm29 IS NULL And NP08>=" & strDate(0) & " And NP08<=" & strDate(1) & " And NP07=305 And NP06 Is Null And Not Exists" & _
'            "(Select CP01 From CaseProgress Where CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And (CP10='1202' or cp10='1201') And CP09>'C') "
'    strQ = strQ & " Union All " & _
'            "Select np01,sqldatet(np08) NP0,NP10,sp01 TM01,sp02 TM02,sp03 TM03,sp04 TM04,sp09 TM10,sp11 TM15,sp08 TM23,Nvl(sp05,Nvl(sp06,sp07)) CaseN " & _
'            "From NextProgress,ServicePractice Where NP02 in ('TT','TS','TR','TM','TD','TC','TB',' ') And NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) " & _
'            "And SP29 IS NULL And NP08>=" & strDate(0) & " And NP08<=" & strDate(1) & " And NP07=305 And NP06 Is Null And Not Exists" & _
'            "(Select CP01 From CaseProgress Where CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And (CP10='1202' or cp10='1201') And CP09>'C') "
'    strQ = strQ & " Union All " & _
'            "Select np01,sqldatet(np08) NP0,NP10,tm01 TM01,tm02 TM02,tm03 TM03,tm04 TM04,tm10 TM10,Nvl(tm15,tm12) TM15,tm23 TM23,Nvl(tm05,Nvl(tm06,tm07)) CaseN " & _
'            "From NextProgress,TradeMark Where NP02 in ('TF','T','FCT',' ') And NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) " & _
'            "And tm29 IS NULL And NP08>=" & strDate(0) & " And NP08<=" & strDate(1) & " And NP07=305 And NP06 Is Null " & _
'            "And Exists(Select CP01 From CaseProgress Where CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And (CP10='1202' or cp10='1201') And CP09>'C') " & _
'            "And Exists(Select CP01 From CaseProgress Where CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And (CP10='203' or cp10='201' or cp10='301' or cp10='302') And CP09<'C' And (CP27='' OR CP27 Is Null) )"
'    strQ = strQ & " Union All " & _
'            "Select np01,sqldatet(np08) NP0,NP10,sp01 TM01,sp02 TM02,sp03 TM03,sp04 TM04,sp09 TM10,sp11 TM15,sp08 TM23,Nvl(sp05,Nvl(sp06,sp07)) CaseN " & _
'            "From NextProgress,ServicePractice Where NP02 in ('TT','TS','TR','TM','TD','TC','TB',' ') And NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+) " & _
'            "And SP29 IS NULL And NP08>=" & strDate(0) & " And NP08<=" & strDate(1) & " And NP07=305 And NP06 Is Null " & _
'            "And Exists(Select CP01 From CaseProgress Where CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And (CP10='1202' or cp10='1201') And CP09>'C') " & _
'            "And Exists(Select CP01 From CaseProgress Where CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And (CP10='203' or cp10='201' or cp10='301' or cp10='302') And CP09<'C' And (CP27='' OR CP27 Is Null) )"
'
'    strQ = "Select np10,sqldatet(cp27) CP27,NP08,TM15, TM01||'-'||TM02||'-'||TM03||'-'||TM04 CaseNo,CaseN,Nvl(Decode(TM10,'000',cpm03,cpm04),cp10),Nvl(cu04,Nvl(cu05||cu88||cu89||cu90,cu06)),S1.st02,S3.st02," & _
'            "Decode(TM10,'000',cp22,fa04) CP44,Nvl(na03,na04) NationN,np01 From (" & strQ & "),CaseProgress,Staff S1,Staff S2,Staff S3,CaseProPertyMap,Customer,Nation,Fagent " & _
'            "Where NP01=CP09(+) And CP14=S1.ST01(+) And NP10=S2.ST01(+) And SubStr(S2.ST03,1,2)='P2' And CP13=S3.ST01(+) And CP01=CPM01(+) And CP10=CPM02(+) And TM10=NA01(+) " & _
'            "And SubStr(TM23,1,8)=CU01(+) And Decode(SubStr(TM23,9,1),'','0',SubStr(TM23,9,1))=CU02(+) And SubStr(CP44,1,8)=FA01(+) And Decode(SubStr(CP44,9,1),'','0',SubStr(CP44,9,1))=FA02(+) Order by NP10,CP27,CP10"
'    RsQ.CursorLocation = adUseClient
'    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
'    With RsQ
'        If .RecordCount > 0 Then
'            .MoveFirst
'            Do While Not .EOF
'
'                If strTemp(0) <> .Fields("np10") Then
'                    If strTemp(0) <> "" Then
'                        strTo = strTemp(0) '前一筆記錄的收信人
'                        If ff14 > 0 Then Close #ff14
'                        SendMAPIMail strTo, strFileN, vbCrLf & vbCrLf & "***詳情請見附件***" & vbCrLf & vbCrLf, strApatch
'                    End If
'                    Txt_Head = False
'                End If
'                If Txt_Head = False Then
'                    ff14 = FreeFile
'                    Open strApatch For Output As ff14
'                    Print #ff14, "                                                                     內商催審函/催審表"
'                    Print #ff14, "列印日期：" & ChangeWStringToTString(strSrvDate(1))
'                    Print #ff14, "發文日     催審期限   申請案號/審定號  本所案號          案件名稱             案件性質 申請人               承辦人   智權人員 是否出名/代理人  申請國家 "
'                    Print #ff14, "========== ========== ================ ================= ==================== ======== ==================== ======== ======== ================ ========== "
'                    Txt_Head = True
'                End If
'
'                Is1705or310 = False
'                For ii = 1 To 11
'                    If ii = 4 Then
'                        '本所案號欄位-檢查有無暫緩或取消催審,有顯示△
'                        strQ1 = "Select * From CaseProgress Where CP43='" & CheckStr(.Fields("np01")) & "' and cp10 in ('1705','310') "
'                        CheckOC2
'                        adoRecordset1.CursorLocation = adUseClient
'                        adoRecordset1.Open strQ1, cnnConnection, adOpenStatic, adLockReadOnly
'                        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                            Is1705or310 = True
'                        End If
'                        strTemp(ii) = IIf(Is1705or310 = True, "△" & "" & .Fields(ii), "" & .Fields(ii))
'                    Else
'                        strTemp(ii) = "" & .Fields(ii)
'                    End If
'                Next ii
'
'                strTemp(1) = convForm(CheckStr(strTemp(1)), 10) '發文日
'                strTemp(2) = convForm(CheckStr(strTemp(2)), 10) '催審期限
'                strTemp(3) = convForm(CheckStr(strTemp(3)), 16) '申請案號/審定號
'                strTemp(4) = convForm(CheckStr(strTemp(4)), 17) '本所案號
'                strTemp(5) = convForm(CheckStr(strTemp(5)), 20) '案件名稱
'
'                strTemp(6) = convForm(CheckStr(strTemp(6)), 8) '案件性質
'                strTemp(7) = convForm(CheckStr(strTemp(7)), 20) '申請人
'                strTemp(8) = convForm(CheckStr(strTemp(8)), 8) '承辦人
'                strTemp(9) = convForm(CheckStr(strTemp(9)), 8) '智權人員
'                strTemp(10) = convForm(CheckStr(strTemp(10)), 16) '是否出名/代理人
'                strTemp(11) = convForm(CheckStr(strTemp(11)), 10) '申請國家
'
'                Print #ff14, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & _
'                          " " & strTemp(6) & " " & strTemp(7) & " " & strTemp(8) & " " & strTemp(9) & " " & strTemp(10) & _
'                          " " & strTemp(11)
'
'                strTemp(0) = .Fields("np10")
'                .MoveNext
'            Loop
'
'            If ff14 > 0 Then
'                Close ff14
'                strTo = strTemp(0) '最後一筆記錄的收信人
'                SendMAPIMail strTo, strFileN, vbCrLf & vbCrLf & "***詳情請見附件***" & vbCrLf & vbCrLf, strApatch
'            End If
'        End If
'    End With
'    RsQ.Close
'    Exit Sub
'
'ErrHand:
'    If Err.Number <> 0 Then
'        WLog "每月自動發催審表給承辦人:" & Err.Description
'    End If
'    Set RsQ = Nothing
'End Sub

'Added by Lydia 2016/10/05 CFT可辦期限管制表 & 提醒上月未通知之可辦期限管制表
'a.每月1日以e-mail分別通知各個CFT承辦組人員處理(副本寄陳經理、陳金蓮);主旨:CFT承辦組-可辦期限管制表
'b.在下一個月1日針對尚未發函(即1717通知延展及1723本所通知使用宣誓)之案件再發e-mail提醒(副本寄陳經理、陳金蓮);主旨:CFT承辦組-提醒上月未通知之可辦期限管制表
Private Sub StrMenu15()
'**************Memo by Lydia 2025/07/04
'可辦期限=法定期限-提前X個月，依案件性質和國家會有不同的提期X個月，整理如下：
   '1. 案件性質為「102延展、109緩審延展」的期限：一律用"法定期限-各國延展時間(月)"
   '2. 案件性質為「105使用宣誓」的期限：
   '   非特定國家並且已有專用期間：提前12個月，另外「112 波多黎各的105使用宣誓並且有專用期間」就照一般規則=可辦期限 = 法定期限- 12個月)。
   '   104墨西哥已有專用期間的使用宣誓可辦期限 = 法定期限-3個月
   '   318莫三比克共和國(不論有無專用期間)的可辦期限 = 法定期限 - 5個月；5個月再加上提前3個月(列印當月+３個月)，剛好是Ellie提出的"前8個月發出提醒"。
   '   030菲律賓的105使用宣誓 (不論有無專用期間) ，可辦期限 = 法定期限-12個月(2025/7/4 調整)
   '   112波多黎各的105使用宣誓沒有專用期間，可辦期限 = 法定期限-6個月。
'**************
   Dim rsB As New ADODB.Recordset
   Dim rsChk As New ADODB.Recordset
   Dim stDate As String, stDate2 As String
   Dim strSql As String, strSQL1 As String
   Dim strTemp(11) As String
   Dim ff15 As Integer
   Dim TempFileName As String, strTFName As String
   Dim strTo As String, strApatch As String
   Dim intCnt As Integer
   Dim callR As Integer
   Dim tmpTitle As String, tmpTitle2 As String, tmpTitle3 As String '報表抬頭設定
   Dim tmpLine As Integer
   Dim stCC As String
   Dim tmpArr As Variant
   Dim tmpArr2 As Variant
   Dim strGrp As String '報表分檔
   Dim strGrp2 As String '記錄承辦人
   Dim tmpTyp As String '案件狀態
   Dim strContent As String 'Added by Lydia 2016/11/23
   Dim mESeqNo As String 'Added by Lydia 2017/05/12 暫存檔序號
   Dim m_Memo As String 'Added by Lydia 2018/01/12 已通知
   'Dim stSpecNa01 As String 'Added b Lydia 2022/11/28 特定國家 'Mark by Lydia 2023/03/24
   
On Error GoTo ErrHandle
   
   'Remove by Lydia 2021/07/26 改成變數
   'stCC = "68005;72012" '副本
   
   'Added by Lydia 2020/08/19 清除舊檔
   Call PUB_KillTempFile(Mid(TextPath & "CFT承辦組-可辦期限管制表*.*", 2))
   'Modified by Lydia 2022/03/25 更名「提醒上月未通知」=>「提醒上月以前未通知」
   Call PUB_KillTempFile(Mid(TextPath & "CFT承辦組-提醒上月以前未通知之可辦期限管制表*.*", 2))
   'end 2020/08/19
   
   For callR = 1 To 2
       If callR = 1 Then
          '列印當月+３個月
          stDate = Mid(CompDate(1, 3, strSrvDate(1)), 1, 6) & "01"
          '月底
          stDate2 = GetLastDay(stDate)
          strTFName = "CFT承辦組-可辦期限管制表_" & ChangeWStringToTString(stDate) & "-" & ChangeWStringToTString(stDate2)
       Else
          '列印當月+２個月
          'Modifed by Lydia 2022/03/25  請修改為提醒所有上月以前未通知可辦期限之案件，即一直提醒到輸入通知期限為止
          'stDate = Mid(CompDate(1, 2, strSrvDate(1)), 1, 6) & "01"
          stDate = "19221111"
          '月底
          'Modified by Lydia 2022/03/25
          'stDate2 = GetLastDay(stDate)
          stDate2 = GetLastDay(Mid(CompDate(1, 2, strSrvDate(1)), 1, 6) & "01")
          'Modified by Lydia 2022/03/25 更名「提醒上月未通知」=>「提醒上月以前未通知」
          'strTFName = "CFT承辦組-提醒上月未通知之可辦期限管制表_" & ChangeWStringToTString(stDate) & "-" & ChangeWStringToTString(stDate2)
          strTFName = "CFT承辦組-提醒上月以前未通知之可辦期限管制表_" & Mid(ChangeWStringToTString(stDate2), 1, 5) & "以前(含)"
       End If
       
       'Added by Lydia 2016/11/23 +列印說明
        strContent = "資料如附件！" & vbCrLf & vbCrLf & "列印注意事項：" & vbCrLf & _
             vbCrLf & String(4, "　") & "1.利用筆記本開啟附件" & _
             vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
             vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
             vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
             vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
             vbCrLf & String(4, "　") & "6.選擇<橫印>"
         
       strGrp = "":  strGrp2 = ""
       strApatch = "": TempFileName = ""
       strSql = "":  strSQL1 = ""
       intCnt = 0: ff15 = 0
                   
       strSQL1 = strSQL1 & " AND NP02 IN (" & SQLGrpStr("CFT", 2) & ")"
       'Modified by Lydia 2019/08/29 +109緩審延展
       'strSQL1 = strSQL1 & " AND ( NP07=102 OR  (tm21 is not null and NP07=105 and tm10<>'030' and tm10<>'112') OR (NP07=105 and tm10 in ('030','112') ) or  NP07=0)"
       'strSQL1 = strSQL1 & " AND ((to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and np07=102) or (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105 and tm21 is not null ) or (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='030')" & _
                           " or (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),6 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),6 ),'YYYYMMDD') and np07=105 and tm21 is null and tm10='112'))"
       'Modified by Lydai 2022/11/28 使用宣誓105的特定國家改用變數控制
       'strSQL1 = strSQL1 & " AND ( NP07=102 OR NP07=109 OR (tm21 is not null and NP07=105 and tm10<>'030' and tm10<>'112') OR (NP07=105 and tm10 in ('030','112') ) or  NP07=0)"
       
       'stSpecNa01 = " '030','112','104','318' " 'Mark by Lydia 2023/03/24 不用變數
       'Modified by Lydia 2023/03/24 debug:不用變數，改在日期條件句判斷
       'strSQL1 = strSQL1 & " AND ( NP07=102 OR NP07=109 OR (tm21 is not null and NP07=105 and tm10 not in (" & stSpecNa01 & ")) OR (NP07=105 and tm10 in (" & stSpecNa01 & ")) or  NP07=0)"
       strSQL1 = strSQL1 & " AND ( NP07=102 OR NP07=109 OR NP07=105 OR NP07=0)"
       'end 2022/11/28
       strSQL1 = strSQL1 & " AND ("
       '109緩審延展的期限算法=102延展
       'Modified by Lydia 2021/06/09 固定月底為31日 ,加上 substr(... ,1,6)||'31'
       strSQL1 = strSQL1 & " (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD') and to_char(NP09)<=substr(to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),n1.na15 ),'YYYYMMDD'),1,6)||'31' and (np07=102 or np07=109)) "
       'modify by sonia 2021/6/9 CFT-020426墨西哥104使用宣誓(可辦期限 = 法定期限-3個月)
       'strSQL1 = strSQL1 & " or (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and to_char(NP09)<=to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and np07=105 and tm21 is not null ) "
       'Modified by Lydia 2022/11/28 排除特定國家
       'strSQL1 = strSQL1 & " or (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and to_char(NP09)<=substr(to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),12 ),'YYYYMMDD'),1,6)||'31' and np07=105 and tm21 is not null and tm10<>'104') "
       'Modified by Lydia 2023/03/24 debug: 030菲律賓,112 波多黎各有專用期間就照一般規則=使用宣誓(可辦期限 = 法定期限- 1年)
       'strSQL1 = strSQL1 & " or (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and to_char(NP09)<=substr(to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),12 ),'YYYYMMDD'),1,6)||'31' and np07=105 and tm21 is not null and tm10 not in (" & stSpecNa01 & ")) "
       'Modified by Lydia 2025/07/04 030菲律賓的105使用宣誓 (不論有無專用期間) ，可辦期限 = 法定期限-12個月=> + and tm10<>'030' 直接排除國家，在更下方才獨立判斷
       strSQL1 = strSQL1 & " or (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and to_char(NP09)<=substr(to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),12 ),'YYYYMMDD'),1,6)||'31' and np07=105 and tm21 is not null and tm10<>'104' and tm10<>'318' and tm10<>'030') "
       '---申請國家104墨西哥 'Memo by Lydia 2023/03/24 參考frm030403:modify by sonia 2021/6/2 CFT-020426墨西哥104使用宣誓(可辦期限 = 法定期限-3個月
       strSQL1 = strSQL1 & " or (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),3 ),'YYYYMMDD') and to_char(NP09)<=substr(to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),3 ),'YYYYMMDD'),1,6)||'31' and np07=105 and tm21 is not null and tm10='104') "
       'end 2021/6/9
       'Added by Lydia 2022/11/28 申請國家318莫三比克共和國的105使用宣誓期限，可辦期限 = 法定期限－5個月，注意不必管TM21有沒有值都要提醒。5個月再加上提前3個月(列印當月+３個月)，剛好是Ellie提出的"前8個月發出提醒"。
       strSQL1 = strSQL1 & " or (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),5 ),'YYYYMMDD') and to_char(NP09)<=substr(to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),5 ),'YYYYMMDD'),1,6)||'31' and np07=105 and tm10='318') "
       'end 2022/11/28
       '---030菲律賓 'Memo by Lydia 2023/03/24 參考frm030403: Modify By Sindy 2013/11/15 菲律賓無專用期105 要抓 14個月
       'Modified by Lydia 2025/07/04 030菲律賓的105使用宣誓 (不論有無專用期間) ，可辦期限 = 法定期限-12個月
       'strSQL1 = strSQL1 & " or (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),14 ),'YYYYMMDD') and to_char(NP09)<=substr(to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),14 ),'YYYYMMDD'),1,6)||'31' and np07=105 and tm21 is null and tm10='030') "
       strSQL1 = strSQL1 & " or (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),12 ),'YYYYMMDD') and to_char(NP09)<=substr(to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),12 ),'YYYYMMDD'),1,6)||'31' and np07=105 and tm10='030') "
       '---112 波多黎各 'Memo by Lydia 2023/03/24 參考frm030403: modify by sonia 2014/11/4 CFT-013059波多黎各112要抓6個月
       strSQL1 = strSQL1 & " or (to_char(NP09)>=to_char(add_months(to_date(" & stDate & " ,'YYYYMMDD'),6 ),'YYYYMMDD') and to_char(NP09)<=substr(to_char(add_months(to_date(" & stDate2 & " ,'YYYYMMDD'),6 ),'YYYYMMDD'),1,6)||'31' and np07=105 and tm21 is null and tm10='112') "
       strSQL1 = strSQL1 & ")"
       'end ----   'Modified by Lydia 2021/06/09 固定月底為31日 ,加上 substr(... ,1,6)||'31'
       'Added by Lydia 2022/03/25 限可辦案件=> 法限>系統日
       If callR = 2 Then strSQL1 = strSQL1 & " and np09>=" & strSrvDate(1)
       
       '排除閉卷
       strSQL1 = strSQL1 & " and (tm29 is null or tm29 <> 'Y' ) "
       
       'Added by Lydia 2017/05/15 先丟暫存檔,再抓承辦人
       cnnConnection.BeginTrans
         cnnConnection.Execute "delete from RAB15 "
         'Modified by Lydia 2019/08/29 +109緩審延展
         'strSql = "INSERT INTO RAB15 " & _
                  "SELECT NP01,NP02,NP03,NP04,NP05,DECODE(NP07,'102','NA69','105','NA69',CP14) NA69,NP22,TM10,CP10,CP13,CP27 " & _
                  "FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,NATION N1 WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND TM10=N1.NA01(+) " & strSQL1 & " AND (NP07<>'997' And NP07<>'998') AND (NP06 IS NULL OR NP06='')"
         strSql = "INSERT INTO RAB15 " & _
                  "SELECT NP01,NP02,NP03,NP04,NP05,DECODE(NP07,'102','NA69','105','NA69','109','NA69',CP14) NA69,NP22,TM10,CP10,CP13,CP27 " & _
                  "FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,NATION N1 WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND TM10=N1.NA01(+) " & strSQL1 & " AND (NP07<>'997' And NP07<>'998') AND (NP06 IS NULL OR NP06='')"
         cnnConnection.Execute strSql, intI
            
         strSql = "SELECT * FROM RAB15 WHERE NA69='NA69' ORDER BY 2,3,4,5 "
         Set rsB = New ADODB.Recordset
         With rsB
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
               .MoveFirst
               Do While Not .EOF
                  'Modified by Lydia 2017/06/03 debug .Fields("CP10") => .Fields("CP13")
                  Call GetNA69("", .Fields("TM10"), "" & .Fields("CP13"), strExc(1), .Fields("NP02"), .Fields("NP03"), .Fields("NP04"), .Fields("NP05"))
                  strSql = "UPDATE RAB15 SET NA69='" & strExc(1) & "' WHERE NP01='" & .Fields("NP01") & "' AND NP22=" & .Fields("NP22")
                  cnnConnection.Execute strSql
                  .MoveNext
               Loop
            End If
         End With
         cnnConnection.CommitTrans
       'end 2017/05/15
       
       'Modified by Lydia 2017/05/15 從暫存檔抓資料
       ' strSql = "SELECT S1.ST03, NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05 as CASENO,NVL(TM05,NVL(TM06,TM07)) as CASENAME,TM09,NVL(TM15,TM12) as TM1512,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) as CPM03," & _
                 " NP15,DECODE(NP07,'102',N1.NA69,'105',N1.NA69,S2.ST01) NA69,NVL(CU1.CU04,NVL(CU1.CU05||CU1.CU88||CU1.CU89||CU1.CU90,CU1.CU06)) as CUSTNAME,NVL(N1.NA03,N1.NA04) as N1NA03,TM22,np07,NP01,CP27,TM01,TM02,TM03,TM04,S1.ST15 as ST15" & _
                 " FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,CASEPROPERTYMAP,CUSTOMER CU1" & _
                 " WHERE NP01=CP09(+) AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND CP14=S2.ST01(+)" & _
                 " AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU1.CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU1.CU02(+)"
       ' strSql = strSql & strSQL1 & " AND (NP07<>'997' And NP07<>'998') AND (NP06 IS NULL OR NP06='')"
        'Modifie by Lydia 2020/05/04 +報表別 ORD1
        'strSql = "SELECT S1.ST03, NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05 as CASENO,NVL(TM05,NVL(TM06,TM07)) as CASENAME,TM09,NVL(TM15,TM12) as TM1512,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) as CPM03," & _
                 " NP15,X.NA69,NVL(CU1.CU04,NVL(CU1.CU05||CU1.CU88||CU1.CU89||CU1.CU90,CU1.CU06)) as CUSTNAME,NVL(N1.NA03,N1.NA04) as N1NA03,TM22,np07,NP01,CP27,TM01,TM02,TM03,TM04,S1.ST15 as ST15" & _
                 " FROM (SELECT B.*,A.NA69,A.CP10,A.CP13,A.CP27 FROM RAB15 A, NEXTPROGRESS B WHERE A.NP01=B.NP01(+) AND A.NP22=B.NP22(+)) X,TRADEMARK,STAFF S1,NATION N1,CASEPROPERTYMAP,CUSTOMER CU1" & _
                 " WHERE NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU1.CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU1.CU02(+)"
        ''end 2017/05/15
        'Modified by Lydia 2021/07/26 副本改寄給「收件人除ST55以外的最高主管+72012」=>NVL(NVL(ST54,ST53),ST52)＋72012
        'strSql = "SELECT DECODE(NP07,'102','1','109','1','105','2',NP07) AS ORD1, S1.ST03, NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05 as CASENO,NVL(TM05,NVL(TM06,TM07)) as CASENAME,TM09,NVL(TM15,TM12) as TM1512,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) as CPM03," & _
                 " NP15,X.NA69,NVL(CU1.CU04,NVL(CU1.CU05||CU1.CU88||CU1.CU89||CU1.CU90,CU1.CU06)) as CUSTNAME,NVL(N1.NA03,N1.NA04) as N1NA03,TM22,np07,NP01,CP27,TM01,TM02,TM03,TM04,S1.ST15 as ST15" & _
                 " FROM (SELECT B.*,A.NA69,A.CP10,A.CP13,A.CP27 FROM RAB15 A, NEXTPROGRESS B WHERE A.NP01=B.NP01(+) AND A.NP22=B.NP22(+)) X,TRADEMARK,STAFF S1,NATION N1,CASEPROPERTYMAP,CUSTOMER CU1" & _
                 " WHERE NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU1.CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU1.CU02(+)"
        'Modified by Lydia 2024/05/02 (2024/2/1)阿蓮回覆退休後不必再發副本給程序。拿掉 ||';72012;' AS NA69M
        strSql = "SELECT DECODE(NP07,'102','1','109','1','105','2',NP07) AS ORD1, S1.ST03, NP10,NP08,NP09,NP02||'-'||NP03||'-'||NP04||'-'||NP05 as CASENO,NVL(TM05,NVL(TM06,TM07)) as CASENAME,TM09,NVL(TM15,TM12) as TM1512,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) as CPM03," & _
                 " NP15,X.NA69,NVL(CU1.CU04,NVL(CU1.CU05||CU1.CU88||CU1.CU89||CU1.CU90,CU1.CU06)) as CUSTNAME,NVL(N1.NA03,N1.NA04) as N1NA03,TM22,np07,NP01,CP27,TM01,TM02,TM03,TM04,S1.ST15 as ST15,X.NA69M" & _
                 " FROM (SELECT B.*,A.NA69,A.CP10,A.CP13,A.CP27,NVL(NVL(S2.ST54,S2.ST53),S2.ST52) AS NA69M FROM RAB15 A, NEXTPROGRESS B,STAFF S2 WHERE A.NP01=B.NP01(+) AND A.NP22=B.NP22(+) AND A.NA69=S2.ST01(+)) X, TRADEMARK,STAFF S1,NATION N1,CASEPROPERTYMAP,CUSTOMER CU1" & _
                 " WHERE NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+) AND NP10=S1.ST01(+) AND TM10=N1.NA01(+) AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+)) AND SUBSTR(TM23,1,8)=CU1.CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU1.CU02(+)"
        'Modified by Lydia 2020/05/04
        'strSql = strSql & " ORDER BY NA69,NP07,CASENO"
        strSql = strSql & " ORDER BY NA69,ORD1,CASENO"
        Set rsB = New ADODB.Recordset
        With rsB
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    '在下一個月1日針對尚未發函(即1717通知延展及1723本所通知使用宣誓)之案件再發e-mail提醒
                    If callR = 2 Then
                       'Modified by Lydia 2020/05/04 +緩審延展109
                       'strSql = "select cp09 from caseprogress where cp01='" & .Fields("TM01") & "' and cp02='" & .Fields("TM02") & "' and cp03='" & .Fields("TM03") & "' and cp04='" & .Fields("TM04") & "'" & _
                                " and cp43='" & .Fields("NP01") & "' and cp10='" & IIf(.Fields("NP07") = "102", "1717", "1723") & "' and cp159=0"
                       strSql = "select cp09 from caseprogress where cp01='" & .Fields("TM01") & "' and cp02='" & .Fields("TM02") & "' and cp03='" & .Fields("TM03") & "' and cp04='" & .Fields("TM04") & "'" & _
                                " and cp43='" & .Fields("NP01") & "' and cp10='" & IIf(.Fields("NP07") = "102" Or .Fields("NP07") = "109", "1717", "1723") & "' and cp159=0"
                       intI = 1
                       Set rsChk = ClsLawReadRstMsg(intI, strSql)
                       If intI = 1 Then GoTo JumpToNext
                    End If
                    
                    '分割檔案
                    'Modified by Lydia 2020/05/04 改成承辦人和報表別
                    'If strGrp <> "" & .Fields("NA69") & .Fields("NP07") Then
                    If strGrp <> "" & .Fields("NA69") & .Fields("ORD1") Then
                       If ff15 > 0 Then
                          Print #ff15, String(tmpLine, "-")
                          Print #ff15, "　共 " & intCnt & " 筆"
                          Close ff15
                       End If
                       '寄給承辦人
                       If strGrp2 <> "" And strGrp2 <> "" & .Fields("NA69") Then
                           strTo = strGrp2 '前一筆記錄的收信人
                           'Added by Lydia 2021/07/26 副本改寄給「收件人除ST55以外的最高主管」
                           stCC = Replace(stCC, strTo & ";", "")
                           If Mid(stCC, 1, 1) = ";" Then stCC = Mid(stCC, 2)
                           'end 2021/07/26
                           
                           'Modified by Lydia 2016/11/23
                           'SendMAPIMail strTo, strTFName, vbCrLf & vbCrLf & "***詳情請見附件***" & vbCrLf & vbCrLf, strApatch, , stCC
                           SendMAPIMail strTo, strTFName, strContent, strApatch, , stCC
                           strApatch = ""
                       End If

                       If "" & .Fields("NP07") = "105" Then
                          tmpTitle = "外商使用宣誓管制表"
                          'Modified by Lydia 2018/01/12
                          'tmpTitle2 = "智權人員,本所期限,法定期限,本所案號,案件名稱,商品類別,申請案號/審定號,備　註,申請人,申請國家"
                          'tmpTitle3 = "4,5,5,8,6,8,8,5,8,4"
                          'tmpLine = 135
                          tmpTitle2 = "智權人員,本所期限,法定期限,本所案號,案件名稱,商品類別,申請案號/審定號,備　註,申請人,申請國家,兩年內"
                          tmpTitle3 = "4,5,5,8,6,8,8,5,8,4,3"
                          tmpLine = 143
                          'end 2018/01/12
                       'Modified by Lydia 2020/05/04 +緩審延展109
                       ElseIf "" & .Fields("NP07") = "102" Or "" & .Fields("NP07") = "109" Then
                          tmpTitle = "外商CFT延展管制表"
                          'Modified by Lydia 2018/01/12
                          'tmpTitle2 = "智權人員,申請人,本所案號,案件名稱,商品類別,審定號,申請國家,本所期限,專用期止日"
                          'tmpTitle3 = "4,8,8,8,8,8,4,5,6"
                          'tmpLine = 125
                          tmpTitle2 = "智權人員,申請人,本所案號,案件名稱,商品類別,審定號,申請國家,本所期限,專用期止日,兩年內"
                          tmpTitle3 = "4,8,8,8,8,8,4,5,6,3"
                          tmpLine = 133
                          'end 2018/01/12
                       End If
                       tmpArr = Empty: tmpArr2 = Empty
                       tmpArr = Split(tmpTitle2, ",")
                       tmpArr2 = Split(tmpTitle3, ",")

                       TempFileName = App.path & TextPath & tmpTitle & ".txt"
                       
                       If Dir(TempFileName) <> "" Then
                          Kill TempFileName
                       End If
                       strApatch = strApatch & IIf(strApatch <> "", "*", "") & TempFileName '郵件附件(報表)
                       
                       intCnt = 0
                       ff15 = FreeFile
                       Open TempFileName For Output As ff15
                       
                       Print #ff15, Space(54) & tmpTitle
                       'Added by Lydia 2022/03/25 區別
                       If callR = 2 Then
                           Print #ff15, "承辦人：" & GetStaffName(.Fields("NA69"), True) & Space(30) & "可辦期限：" & ChangeWStringToTDateString(stDate2) & "以前(含)" & Space(30) & "列印日期：" & ChangeWStringToTDateString(strSrvDate(1))
                       Else
                       'end 2022/03/25
                           Print #ff15, "承辦人：" & GetStaffName(.Fields("NA69"), True) & Space(30) & "可辦期限：" & ChangeWStringToTDateString(stDate) & "-" & ChangeWStringToTDateString(stDate2) & Space(30) & "列印日期：" & ChangeWStringToTDateString(strSrvDate(1))
                       End If 'Added by Lydia 2022/03/25
                       Print #ff15, ""
                       If "" & .Fields("NP07") = "105" Then
                          Print #ff15, "# 表示承辦人未通知主管機關來函 , x 表示不催延展"
                       'Modified by Lydia 2020/05/04 + 緩審延展109
                       'ElseIf "" & .Fields("NP07") = "102" Or "" Then
                       '   Print #ff15, "x 表示不催延展"
                       ElseIf "" & .Fields("NP07") = "102" Or "" & .Fields("NP07") = "109" Then
                          Print #ff15, "x 表示不催延展, * 表示緩審延展"
                       End If
                       
                       strExc(1) = ""
                       For intI = 0 To UBound(tmpArr)
                          If Trim(tmpArr(intI)) <> "" Then
                             'Added by Lydia 2018/01/12 兩年內=>空白
                             If Trim(tmpArr(intI)) = "兩年內" Then
                                 strExc(1) = strExc(1) & convForm("　", Val(tmpArr2(intI)) * 2) & " "
                             Else
                             'end 2018/01/12
                                 strExc(1) = strExc(1) & convForm(Trim(tmpArr(intI)), Val(tmpArr2(intI)) * 2) & " "
                             End If 'end 2018/01/12
                          End If
                       Next
                       
                       Print #ff15, String(tmpLine, "-")
                       Print #ff15, strExc(1)
                       Print #ff15, String(tmpLine, "-")
                    End If
     
                    '案件狀態
                    tmpTyp = ""
                    If PUB_ChkCaseIsNoticeScale(.Fields("TM01"), .Fields("TM02"), .Fields("TM03"), .Fields("TM04")) = False Then
                       tmpTyp = "x" '不催延展者
                    Else
                        If "" & .Fields("NP07") = "105" Then
                            If Mid(CheckStr("" & .Fields("NP01")), 1, 1) = "C" And Len(CheckStr("" & .Fields("CP27"))) = 0 Then
                                tmpTyp = "#"
                            End If
                        End If
                    End If
                    'Added by Lydia 2020/05/04 緩審延展
                    If "" & .Fields("NP07") = "109" Then
                         tmpTyp = "*"
                    End If
                    'end 2020/05/04
                    
                    'Added by Lydia 2018/01/12 檢查是否已通知
                    m_Memo = ""
                    If "" & .Fields("NP09") <> "" Then
                       '外商使用宣誓管制表：若法定期限前二年內有1723進度則報表加註'已通知'
                       '外商CFT延展管制表：若法定期限前二年內有1717進度則報表加註'已通知'
                       strSql = "select cp09 from caseprogress where cp01='" & .Fields("TM01") & "' and cp02='" & .Fields("TM02") & "' and cp03='" & .Fields("TM03") & "' and cp04='" & .Fields("TM04") & "'" & _
                                " and cp05>=" & CompDate(0, -2, .Fields("NP09")) & " and cp10='" & IIf(.Fields("NP07") = "102", "1717", "1723") & "' and cp159=0"
                       intI = 1
                       Set rsChk = ClsLawReadRstMsg(intI, strSql)
                       If intI = 1 Then
                           m_Memo = "已通知"
                       End If
                    End If
                    'end 2018/01/12
                    
                    '檢查若智權人員離職時, 需要重新取得目前承辦智權人員
                    strTemp(0) = GetStaffName("" & .Fields("NP10"))
                    If strTemp(0) = "" Then
                       strTemp(0) = GetStaffName(PUB_GetAKindSalesNo(.Fields("TM01"), .Fields("TM02"), .Fields("TM03"), .Fields("TM04")))
                    End If
                    strTemp(0) = PUB_StrToStr(strTemp(0), Val(tmpArr2(0)) * 2, True)
                    
                    If "" & .Fields("NP07") = "105" Then '使用宣誓
                        strTemp(1) = convForm(ChangeWStringToTDateString("" & .Fields("NP08")), Val(tmpArr2(1)) * 2)
                        strTemp(2) = convForm(ChangeWStringToTDateString("" & .Fields("NP09")), Val(tmpArr2(2)) * 2)
                        strTemp(3) = convForm(tmpTyp & .Fields("CASENO"), Val(tmpArr2(3)) * 2)
                        strTemp(4) = convForm(PUB_StringFilter("" & .Fields("CASENAME")), Val(tmpArr2(4)) * 2)
                        strTemp(5) = convForm("" & .Fields("TM09"), Val(tmpArr2(5)) * 2)
                        strTemp(6) = convForm("" & .Fields("TM1512"), Val(tmpArr2(6)) * 2)
                        strTemp(7) = convForm("" & .Fields("NP15"), Val(tmpArr2(7)) * 2)
                        strTemp(8) = convForm(PUB_StringFilter("" & .Fields("CUSTNAME")), Val(tmpArr2(8)) * 2)
                        strTemp(9) = convForm("" & .Fields("N1NA03"), Val(tmpArr2(9)) * 2)
                        strTemp(10) = convForm(m_Memo, Val(tmpArr2(10)) * 2) 'Added by Lydia 2018/01/12
                    'Modified by Lydia 2020/05/04 + 緩審延展109
                    ElseIf "" & .Fields("NP07") = "102" Or "" & .Fields("NP07") = "109" Then '延展
                        strTemp(1) = convForm(PUB_StringFilter("" & .Fields("CUSTNAME")), Val(tmpArr2(1)) * 2)
                        strTemp(2) = convForm(tmpTyp & .Fields("CASENO"), Val(tmpArr2(2)) * 2)
                        strTemp(3) = convForm(PUB_StringFilter("" & .Fields("CASENAME")), Val(tmpArr2(3)) * 2)
                        strTemp(4) = convForm("" & .Fields("TM09"), Val(tmpArr2(4)) * 2)
                        strTemp(5) = convForm("" & .Fields("TM1512"), Val(tmpArr2(5)) * 2)
                        strTemp(6) = convForm("" & .Fields("N1NA03"), Val(tmpArr2(6)) * 2)
                        strTemp(7) = convForm(ChangeWStringToTDateString("" & .Fields("NP08")), Val(tmpArr2(7)) * 2)
                        strTemp(8) = convForm(ChangeWStringToTDateString("" & .Fields("TM22")), Val(tmpArr2(8)) * 2)
                        strTemp(9) = convForm(m_Memo, Val(tmpArr2(9)) * 2) 'Added by Lydia 2018/01/12
                    End If
                    
                    strExc(1) = ""
                    For intI = 0 To UBound(tmpArr)
                       If Trim(tmpArr(intI)) <> "" Then
                          strExc(1) = strExc(1) & strTemp(intI) & " "
                       End If
                    Next
                
                    Print #ff15, strExc(1)
                    intCnt = intCnt + 1
                    'Modified by Lydia 2020/05/04 改成承辦人和報表別
                    'strGrp = "" & .Fields("NA69") & .Fields("NP07")
                    strGrp = "" & .Fields("NA69") & .Fields("ORD1")
                    strGrp2 = "" & .Fields("NA69")
                    stCC = "" & .Fields("NA69M")  'Added by Lydia 2021/07/26 副本改寄給「收件人除ST55以外的最高主管」
JumpToNext:
                    .MoveNext
                Loop
            End If
            If ff15 > 0 Then
                 '寄給承辦人
                Print #ff15, String(tmpLine, "-")
                Print #ff15, "　共 " & intCnt & " 筆"
                Close ff15
                strTo = strGrp2
                'Added by Lydia 2021/07/26 副本改寄給「收件人除ST55以外的最高主管」
                stCC = Replace(stCC, strTo & ";", "")
                If Mid(stCC, 1, 1) = ";" Then stCC = Mid(stCC, 2)
                'end 2021/07/26
                'Modified by Lydia 2016/11/23
                'SendMAPIMail strTo, strTFName, vbCrLf & vbCrLf & "***詳情請見附件***" & vbCrLf & vbCrLf, strApatch, , stCC
                SendMAPIMail strTo, strTFName, strContent, strApatch, , stCC
            End If
        End With
JumpToNoData: 'Added by Lydia 2017/05/15

    Next callR
    
    Set rsB = Nothing
    Set rsChk = Nothing
ErrHandle:
   If Err.Number <> 0 Then
      WLog "CFT可辦期限管制表:" & Err.Description
   End If
   Set rsB = Nothing
   Set rsChk = Nothing
End Sub

'Added by Lydia 2017/02/22 國內新客戶清單(電腦中心->定期作業->新客戶清單frm12040141)
Private Sub StrMenu16()
Dim rsRD As New ADODB.Recordset
Dim inR As Integer
Dim strDate_S As String, strDate_E As String
Dim PLeft(0 To 4) As Integer
Dim TempFileName As String
Dim ff16 As Integer
Dim A01 As String, A02 As String, A03 As String, A04 As String, A05 As String

Dim iRow As Integer, j As Integer
Dim strNo As String '所別
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strTo As String
Dim strPS As String

'系統日之前一個月
strDate_S = Mid(CompDate(1, -1, strSrvDate(1)), 1, 6) & "01"
strDate_E = GetLastDay(strDate_S)
PLeft(0) = 12
PLeft(1) = 42
PLeft(2) = 12
PLeft(3) = 12
PLeft(4) = 13
TempFileName = "新客戶清單"
   
Call PUB_KillTempFile(Mid(TextPath & TempFileName & "*.*", 2))    'Added by Lydia 2020/08/19 清除舊檔

strPS = vbCrLf & vbCrLf & vbCrLf & String(2, "　") & "列印注意事項：" & vbCrLf & _
      vbCrLf & String(4, "　") & "1.利用筆記本開啟附件" & _
      vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
      vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
      vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
      vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
      vbCrLf & String(4, "　") & "6.選擇<橫印>   "
      
On Error GoTo ErrHandle

   '限國內,以e -mail寄給智權人員
   strSql = "SELECT CU01||CU02 C01,SUBSTR(CU04,1,30) C02,SUBSTR(CU07,1,10) C03,SUBSTR(NA03,1,14) C04,CU16,CU30,SUBSTR(NVL(CU31,CU23),1,65) C07,CU18," & _
            "CU32,CU80,CU13,ST02,CU12,CU23,CU11,ST06,DECODE(ST06,'1','北所','2','中所','3','南所','4','高所','其他')  AREA,ST04,A0908 " & _
            "FROM CUSTOMER,NATION,STAFF,ACC090 Where  CU10=NA01(+) AND CU13=ST01(+)  AND CU14 BETWEEN " & strDate_S & " AND " & strDate_E & "  And ST06>='1' " & _
            "And ST06<='5'  AND (ST15<'F' OR ST15>'F99')  AND CU12>='S' AND CU12<='S99' AND ST15=A0901(+) ORDER BY CU12,CU13,CU01,CU02 "
   inR = 1
   Set rsRD = ClsLawReadRstMsg(inR, strSql)
       
   If inR = 1 Then
        rsRD.MoveFirst
        iRow = 0
        '若智權人員離職則改發部門主管(ACC090之A0908)
        strTo = IIf(rsRD.Fields("ST04") = "1", rsRD.Fields(10).Value, IIf("" & rsRD.Fields("A0908") <> "", rsRD.Fields("A0908"), rsRD.Fields(10).Value))
        strNo = rsRD.Fields(12).Value & "　" & rsRD.Fields(11).Value
        StrMenu16_Title strNo, TempFileName, strDate_S, strDate_E, ff16, PLeft, A01, A02, A03, A04, A05
        With rsRD
           Do While Not .EOF
              iRow = iRow + 1
              For j = 0 To 4
                 If j = 0 Then
                    If Not IsNull(rsRD.Fields(9).Value) Then
                       If Not IsNull(rsRD.Fields(8).Value) Then
                          A01 = convForm("＊" & Left(.Fields(j) & "000", 9) & " N", PLeft(0))
                       Else
                          A01 = convForm("＊" & Left(.Fields(j) & "000", 9), PLeft(0))
                       End If
                    Else
                       If Not IsNull(rsRD.Fields(8).Value) Then
                          A01 = convForm(" " & Left(.Fields(j) & "000", 9) & " N", PLeft(0))
                       Else
                          A01 = convForm("" & Left(.Fields(j) & "000", 9), PLeft(0))
                       End If
                    End If
                 ElseIf j = 1 Then
                    A02 = convForm("" & .Fields(j), PLeft(1))
                 '若列印負責人
                 ElseIf j = 2 Then
                    A03 = convForm(Left("" & .Fields(j), 7), PLeft(2))
                 ElseIf j = 3 Then
                    A04 = convForm("" & .Fields(j), PLeft(3))
                 ElseIf j = 4 Then
                    A05 = convForm("" & .Fields(j), PLeft(4))
                 End If
              Next j
              Print #ff16, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
              
              For j = 5 To 7
                 If j = 5 Then
                    A01 = convForm("" & .Fields(j), PLeft(0))
                 ElseIf j = 6 Then
                    A02 = convForm("" & .Fields(j), PLeft(1))
                    A03 = convForm(" ", PLeft(2))
                    A04 = convForm(" ", PLeft(3))
                 ElseIf j = 7 Then
                    A05 = convForm("" & .Fields(j), PLeft(4))
                 End If
              Next j
              Print #ff16, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
              
              A01 = convForm(" ", PLeft(0))
              A02 = convForm("" & .Fields(13), PLeft(1))
              A03 = convForm(" ", PLeft(2))
              A04 = convForm("" & .Fields(11), PLeft(3))
              A05 = convForm("" & .Fields(14), PLeft(4))
              Print #ff16, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
              
              '若為子公司
              If Right("" & .Fields(0).Value, 3) <> "000" Then
                 StrSQLa = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,CU30,SUBSTR(NVL(CU31,CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,ST06,DECODE(ST06,'1','北所','2','中所','3','南所','4','高所','其他') FROM CUSTOMER,NATION,STAFF Where  CU10=NA01(+) AND CU13=ST01(+) AND CU01='" & Left(.Fields(0).Value, 6) & "00" & "' AND CU02='0' "
                 rsA.CursorLocation = adUseClient
                 rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                 If rsA.RecordCount > 0 Then
                    Print #ff16, "母公司資料："
                    For j = 0 To 4
                       If j = 0 Then
                          If Not IsNull(rsA.Fields(9).Value) Then
                             If Not IsNull(rsA.Fields(8).Value) Then
                                A01 = convForm("＊" & Left(rsA.Fields(j) & "000", 9) & " N", PLeft(0))
                             Else
                                A01 = convForm("＊" & Left(rsA.Fields(j) & "000", 9), PLeft(0))
                             End If
                          Else
                             If Not IsNull(rsA.Fields(8).Value) Then
                                A01 = convForm(" " & Left(rsA.Fields(j) & "000", 9) & " N", PLeft(0))
                             Else
                                A01 = convForm("" & Left(rsA.Fields(j) & "000", 9), PLeft(0))
                             End If
                          End If
                       ElseIf j = 1 Then
                          A02 = convForm("" & rsA.Fields(j), PLeft(1))
                       '若列印負責人
                       ElseIf j = 2 Then
                          A03 = convForm(Left("" & rsA.Fields(j), 7), PLeft(2))
                       ElseIf j = 3 Then
                          A04 = convForm("" & rsA.Fields(j), PLeft(3))
                       ElseIf j = 4 Then
                          A05 = convForm("" & rsA.Fields(j), PLeft(4))
                       End If
                    Next j
                    Print #ff16, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
                    
                    For j = 5 To 7
                       If j = 5 Then
                          A01 = convForm("" & rsA.Fields(j), PLeft(0))
                       ElseIf j = 6 Then
                          A02 = convForm("" & rsA.Fields(j), PLeft(1))
                          A03 = convForm(" ", PLeft(2))
                          A04 = convForm(" ", PLeft(3))
                       ElseIf j = 7 Then
                          A05 = convForm("" & rsA.Fields(j), PLeft(4))
                       End If
                    Next j
                    Print #ff16, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
                    
                    A01 = convForm(" ", PLeft(0))
                    A02 = convForm("" & rsA.Fields(13), PLeft(1))
                    A03 = convForm(" ", PLeft(2))
                    A04 = convForm("" & rsA.Fields(11), PLeft(3))
                    A05 = convForm("" & rsA.Fields(14), PLeft(4))
                    Print #ff16, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
                 End If
                 If rsA.State <> adStateClosed Then rsA.Close
                 Set rsA = Nothing
              End If
              Print #ff16, "-----------------------------------------------------------------------------------------------"
              
              .MoveNext
              If rsRD.EOF Then Exit Do
              If rsRD.Fields(12).Value & "　" & rsRD.Fields(11).Value <> strNo Then
                 Print #ff16, "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
                 Print #ff16, ""
                 Print #ff16, "共 " & iRow & " 筆"
                 Close ff16
                 SendMAPIMail strTo, ChangeTStringToTDateString(TransDate(strDate_S, 1)) & " - " & ChangeTStringToTDateString(TransDate(strDate_E, 1)) & TempFileName, strPS, App.path & TextPath & "\" & TempFileName & ".txt"
                 iRow = 0
                 '若智權人員離職則改發部門主管(ACC090之A0908)
                 strTo = IIf(rsRD.Fields("ST04") = "1", rsRD.Fields(10).Value, IIf("" & rsRD.Fields("A0908") <> "", rsRD.Fields("A0908"), rsRD.Fields(10).Value))
                 strNo = rsRD.Fields(12).Value & "　" & rsRD.Fields(11).Value
                 StrMenu16_Title strNo, TempFileName, strDate_S, strDate_E, ff16, PLeft, A01, A02, A03, A04, A05
              End If
           Loop
        End With
        Print #ff16, "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
        Print #ff16, ""
        Print #ff16, "共 " & iRow & " 筆"
        Close ff16
        SendMAPIMail strTo, ChangeTStringToTDateString(TransDate(strDate_S, 1)) & " - " & ChangeTStringToTDateString(TransDate(strDate_E, 1)) & TempFileName, strPS, App.path & TextPath & "\" & TempFileName & ".txt"
   End If
   
ErrHandle:
   If Err.Number <> 0 Then
      WLog "國內新客戶清單:" & Err.Description
   End If
   Set rsA = Nothing
   Set rsRD = Nothing
End Sub

'Added by Lydia 2017/02/22 國內新客戶清單-表單
Private Sub StrMenu16_Title(ByVal strSNo As String, ByVal strFN As String, ByVal sDate1 As String, ByVal sDate2 As String, ByRef rF1 As Integer, ByRef rLeft() As Integer, ByRef rA01 As String, ByRef rA02 As String, ByRef rA03 As String, ByRef rA04 As String, ByRef rA05 As String)
'strSNo : 業務區別+智權人員 或 所別

   rF1 = FreeFile
   If rF1 > 0 Then Close #rF1
   rF1 = FreeFile

   Open App.path & TextPath & "\" & strFN & ".txt" For Output As rF1
   Print #rF1, "                                      新　客　戶　清　單                                      "
   Print #rF1, "智權人員 : " & strSNo & "　　　　　　　　　　　　　　　　　　　　          列印日期 : " & ChangeTStringToTDateString(TransDate(strSrvDate(1), 1))
   Print #rF1, "開發日期 : " & ChangeTStringToTDateString(TransDate(sDate1, 1)) & " - " & ChangeTStringToTDateString(TransDate(sDate2, 1))
   Print #rF1, "-----------------------------------------------------------------------------------------------"
   
   rA01 = convForm("編號", rLeft(0))
   rA02 = convForm("公司名稱", rLeft(1))
   rA03 = convForm("負責人", rLeft(2))
   rA04 = convForm("國籍", rLeft(3))
   rA05 = convForm("電話", rLeft(4))
   Print #rF1, rA01 & " " & rA02 & " " & rA03 & " " & rA04 & " " & rA05
   
   rA01 = convForm("郵遞區號", rLeft(0))
   rA02 = convForm("聯絡地址", rLeft(1))
   rA03 = convForm(" ", rLeft(2))
   rA04 = convForm(" ", rLeft(3))
   rA05 = convForm("傳真", rLeft(4))
   Print #rF1, rA01 & " " & rA02 & " " & rA03 & " " & rA04 & " " & rA05
   
   rA01 = convForm(" ", rLeft(0))
   rA02 = convForm("中文地址", rLeft(1))
   rA03 = convForm(" ", rLeft(2))
   rA04 = convForm("智權人員", rLeft(3))
   rA05 = convForm("統一編號", rLeft(4))
   Print #rF1, rA01 & " " & rA02 & " " & rA03 & " " & rA04 & " " & rA05
   Print #rF1, "-----------------------------------------------------------------------------------------------"
End Sub
'end 2017/02/22

'Added by Lydia 2017/02/22 國內專利收文未發文明細表(電腦中心->定期作業->收文未發文明細表frm12040142)
Sub StrMenu17()
Dim strSQLSkipP As String, StrTest98 As String
Dim StrTest1 As String, StrTest4 As String
Dim strDate_S As String, strDate_E As String
Dim strTemp3(0 To 15) As String
Dim strTempName As String '代理人名稱
Dim strSalesGrp As String '業務區
Dim strSalesGrpName As String '業務區名稱
Dim TmpArea As String '業務區＆部門
Dim strSG As String, strSGN As String, strSName As String
Dim eFile As Integer, eFilename As String
Dim strPath As String
Dim strFileN As String
Dim mailFList As String, m_SpecMan As String
Dim St As String, iK As Integer, iTatle As Integer
Dim iKK As Integer '業務區合計
Dim iCall As Integer  '管制時段
Dim iR As Integer, j As Integer
Dim rsAD As New ADODB.Recordset
Dim strFTitle As String
Dim strAddFCP As String, strAddFCT  As String 'Added by Lydia 2020/03/04 若有外專F2x／外商資料F1x，則增加通知王協理(外專)和江協理(外商)
Dim strAddFCPcc As String 'Added by Lydia 2022/11/25 增加判斷專利國外部及日本部收文案件CC

strDate_S = TransDate("850101", 2)
strPath = App.path & TextPath

Call PUB_KillTempFile(Mid(TextPath & "*專利收文逾*.*", 2))     'Added by Lydia 2020/08/19 清除舊檔

For iCall = 1 To 3 '管制時段
'----------查詢條件
    StrTest1 = " AND cc1.cp01 IN ('CFP','P') "        '專利
    StrTest4 = " cc1.cp01 IN ('CPS','PS') "      '服務專利
    StrTest98 = " and cc2.cp14 IS NOT NULL AND cc2.cp158=0 AND cc2.cp159=0 and substr(cc2.cp09,1,1) <> 'D' "
    
    Select Case iCall '管制時段-收文日期
        Case 1  '智權人員
            '保留
            'If Val(Right(strSrvDate(1), 2)) > 20 Then
            '   strDate_E = CompDate(2, -1, CompDate(1, -5, Left(strSrvDate(1), 6) & "01"))
            'Else
               strDate_E = CompDate(2, -1, CompDate(1, -6, Left(strSrvDate(1), 6) & "01"))
               strFTitle = "專利收文逾六個月未發文明細表"
            'End If
        Case 2  '部門主管
            '保留
            'If Val(Right(strSrvDate(1), 2)) > 20 Then
            '   strDate_E = CompDate(2, -1, CompDate(1, -6, Left(strSrvDate(1), 6) & "01"))
            'Else
               strDate_E = CompDate(2, -1, CompDate(1, -7, Left(strSrvDate(1), 6) & "01"))
               strFTitle = "專利收文逾七個月未發文明細表"
            'End If
        Case 3  '副總
            '保留
            'If Val(Right(strSrvDate(1), 2)) > 20 Then
            '   strDate_E = CompDate(2, -1, CompDate(1, -7, Left(strSrvDate(1), 6) & "01"))
            'Else
               strDate_E = CompDate(2, -1, CompDate(1, -8, Left(strSrvDate(1), 6) & "01"))
               strFTitle = "專利收文逾八個月未發文明細表"
            'End If
    End Select

    '專利-收文日期
    StrTest1 = StrTest1 + " AND cc1.cp05>='" & strDate_S & "' AND cc1.cp05<='" & strDate_E & "'"
    StrTest4 = StrTest4 + " AND cc1.cp05>='" & strDate_S & "' AND cc1.cp05<='" & strDate_E & "'"
    StrTest98 = StrTest98 & " AND cc2.cp05>='" & strDate_S & "' AND cc2.cp05<='" & strDate_E & "'"
    
    '加未到期控制條件
    StrTest1 = StrTest1 & " AND ((substr(cc1.cp12,1,1)='F' and (cc1.cp06 Is Null Or cc1.cp06<=" & strSrvDate(1) & ")) or (substr(cc1.cp12,1,1)<>'F' and ((cc1.cp06 Is Null Or cc1.cp06<=" & strSrvDate(1) & ") or (cc1.cp10 in (" & CaseMapIn & ") OR CC1.CP10='107'))))"
    StrTest98 = StrTest98 & " AND ( cc2.cp06 Is Null Or cc2.cp06<=" & strSrvDate(1) & " ) "

    '排除特定案件
    strSQLSkipP = " And ((Not (cc1.cp01||cc1.cp10 in ('CFP215','CFP416','CFP207','P421','P416','P211','P212','P111') and (cc1.cp06 is  null or cc1.cp06='')) ) or (cc1.cp01 in (select cc2.cp01 from CASEPROGRESS cc2 where cc2.cp01=cc1.cp01 and cc2.cp02=cc1.cp02 and cc2.cp03=cc1.cp03 and cc2.cp04=cc1.cp04 and cc2.cp31='Y'  " & StrTest98 & " ))) "
    
    '加延遲過系統日或無延遲
    StrTest1 = StrTest1 & " and (cc1.cp108 is null or cc1.cp108<=" & strSrvDate(1) & ") "
    StrTest4 = StrTest4 & " and (cc1.cp108 is null or cc1.cp108<=" & strSrvDate(1) & ") " '等於StrTest7
    StrTest98 = StrTest98 & " and (cc2.cp108 is null or cc2.cp108<=" & strSrvDate(1) & ") "

    '國外部收的澳門大陸案關聯,香港案110(無期限,不顯示) ,以及P,FCP之分割案無期限,不顯示
    strExc(5) = ",(select cp01 v1c1,cp02 v1c2,cp03 v1c3,cp04 v1c4,cp06 v1c6,cp07 v1c7,cp12 v1c8 from casemap,caseprogress where cm10 in ('4','5') and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and ((cm10='4' and cp10='110') or (cm10='5' and cp10 in (" & CaseMapIn & "))) ) VT1 " & _
                ",(select cp01 v2c1,cp02 v2c2,cp03 v2c3,cp04 v2c4,cp06 v2c6,cp07 v2c7,cp12 v2c8 from divisioncase,caseprogress where dc01 in ('P','FCP') and dc01=cp01(+) and dc02=cp02(+) and dc03=cp03(+) and dc04=cp04(+) and cp10 = '307' ) VT2 "
    '判斷條件
    strExc(6) = "and decode(v1c1||v2c1,null,1,decode(substr(v1c6||v1c8,1,1),'F',0,decode(substr(v2c6||v2c8,1,1),'F',0,1)))=1 "
    '專利案431(PPH)無期限不出現
    strExc(6) = strExc(6) & " and decode(cc1.cp10||cc1.cp06,'431',0,1)=1 "
    
    '專利
    'Modified by Lydia 2022/03/28 若總收文號為B類，則在案件性質前加B，例B補收款; + decode(substr(cc1.cp09,1,1),'B','B',null)||decode(substr(cc1.cp09,1,1),'B','B',null)||
    'Modified by Lydia 2022/11/25 增加判斷專利國外部及日本部收文案件; S2.ST15 AS D => S2.ST15||DECODE(S2.ST15,'F23',S2.ST16,'') AS D
    strSql = "SELECT S2.ST01 AS A,cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(PA05,NVL(PA06,PA07)) casename,'' c05,'' c06,cc1.cp06,cc1.cp07," & _
       "DECODE(PA09,'000',PTM03,PTM04) ptm03,decode(substr(cc1.cp09,1,1),'B','B',null)||decode(pa09,'000',cpm03,cpm04) cpm03,cc1.cp14 as F,NA03,PA26,PA27,PA28,PA29,PA30," & _
       "PA75,cc1.cp44,S2.ST15||DECODE(S2.ST15,'F23',S2.ST16,'') AS D,AA.A0902 as E, cc1.cp79,s1.st03 as G,BB.A0902 as H,PA47 NO " & _
       "FROM CASEPROGRESS cc1,PATENT,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP,NATION,ACC090 AA,Acc090 BB " & _
       strExc(5) & "WHERE  cc1.cp14 IS NOT NULL AND cc1.cp158=0 AND cc1.cp159=0 and substr(cc1.cp09,1,1) <> 'D' AND cc1.cp01=PA01(+) AND cc1.cp02=PA02(+) AND cc1.cp03=PA03(+) AND cc1.cp04=PA04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) " & _
       "AND cc1.cp01=cpm01(+) AND cc1.cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND (PA57<>'Y' OR PA57 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) " & _
       " and cp01=v1c1(+) and cp02=v1c2(+) and cp03=v1c3(+) and cp04=v1c4(+) and cp01=v2c1(+) and cp02=v2c2(+) and cp03=v2c3(+) and cp04=v2c4(+) " & _
       strExc(6) & StrTest1
    strSql = strSql & strSQLSkipP

    '服務業務
    'Modified by Lydia 2022/03/28 若總收文號為B類，則在案件性質前加B，例B補收款; + decode(substr(cc1.cp09,1,1),'B','B',null)||decode(substr(cc1.cp09,1,1),'B','B',null)||
    'Modified by Lydia 2022/11/25 增加判斷專利國外部及日本部收文案件; S2.ST15 AS D => S2.ST15||DECODE(S2.ST15,'F23',S2.ST16,'') AS D
    strSql = strSql + " union all select S2.ST01 AS A,cc1.cp05 AS B,cc1.cp01||'-'||cc1.cp02||'-'||cc1.cp03||'-'||cc1.cp04 AS C,NVL(SP05,NVL(SP06,SP07))," & _
       "'','',cc1.cp06,cc1.cp07,'',decode(substr(cc1.cp09,1,1),'B','B',null)||decode(sp09,'000',cpm03,cpm04) cpm03,cc1.cp14 as F,NA03,SP08,SP58,SP59,SP65,SP66," & _
       "SP26,cc1.cp44,S2.ST15||DECODE(S2.ST15,'F23',S2.ST16,'') AS D,AA.A0902 as E, cc1.cp79,s1.st03 as G,BB.A0902 as H,SP28 NO " & _
       "FROM CASEPROGRESS cc1,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2,NATION,ACC090 AA,Acc090 BB " & _
       "WHERE cc1.cp14 IS NOT NULL AND cc1.cp158=0 AND cc1.cp159=0 and substr(cc1.cp09,1,1) <> 'D' AND " & _
       "cc1.cp01=SP01(+) AND cc1.cp02=SP02(+) AND cc1.cp03=SP03(+) AND cc1.cp04=SP04(+) AND cc1.cp14=S1.ST01(+) AND cc1.cp13=S2.ST01(+) AND cc1.cp01=cpm01(+) AND " & _
       "cc1.cp10=cpm02(+) AND SP09=NA01(+) AND (SP15<>'Y' OR SP15 IS NULL) AND S2.ST15=AA.A0901(+) and s1.st03=BB.A0901(+) and " & StrTest4
    
    strSql = strSql + " ORDER BY D,A,B,C "
'----------查詢條件
    iR = 1
    Set rsAD = ClsLawReadRstMsg(iR, strSql)
    If iR = 1 Then
        eFilename = "": eFile = 0: mailFList = "": m_SpecMan = ""
        '清除舊檔
        strExc(7) = Dir(strPath & "*" & strFileN & "*")
        Do While strExc(7) <> ""
           strExc(8) = Mid(strExc(7), 1, 7)
           If CheckIsTaiwanDate(strExc(8), False) Then
              '保留前一個月及當月的檔案
              strExc(8) = ChangeWStringToTString(ChangeWDateStringToWString(strExc(8)))
              If (Left(strExc(8), 5) = Left(TransDate(strSrvDate(1), 1), 5)) Or (Val(Left(strExc(8), 5)) = Val(Left(TransDate(strSrvDate(1), 1), 5)) - 1) Or (Val(Left(strExc(8), 5)) = Val(Left(TransDate(strSrvDate(1), 1), 5)) - 89) Then
                 Exit Do
              Else
                 Kill strPath & strExc(7)
              End If
           Else
              Exit Do
           End If
           strExc(7) = Dir(strPath & "*" & strFileN & "*")
        Loop
        '檔名前+日期
        strFileN = TransDate(strSrvDate(1), 1) & "_" & strFTitle

        With rsAD
            .MoveFirst
            TmpArea = "業務區：" & .Fields("E").Value
            St = .Fields("A")
            iTatle = 0       ' 總筆數
            iKK = 0           ' 合計
            iK = 0           ' 小計
            'Added by Lydia 2020/03/04
            strAddFCP = ""
            strAddFCT = ""
            strAddFCPcc = "" 'Added by Lydia 2022/11/25
            'end 2020/03/04
            Do While .EOF = False
                For j = 0 To 11
                    If Not IsNull(.Fields(j)) Then
                        strTemp3(j) = .Fields(j)
                    Else
                        strTemp3(j) = ""
                    End If
                Next j
                strTemp3(0) = "" & .Fields("A").Value '智權人
                strTemp3(10) = StrToStr(GetStaffName("" & .Fields("F").Value, True), 4) '承辦人-名稱
                If Not IsNull(.Fields(12)) Then '申請人1
                    strTemp3(4) = .Fields(12)
                Else
                    If Not IsNull(.Fields(13)) Then
                        strTemp3(4) = .Fields(13)
                    Else
                        If Not IsNull(.Fields(14)) Then
                            strTemp3(4) = .Fields(14)
                        Else
                            If Not IsNull(.Fields(15)) Then
                                strTemp3(4) = .Fields(15)
                            Else
                                If Not IsNull(.Fields(16)) Then
                                    strTemp3(4) = .Fields(16)
                                End If
                            End If
                        End If
                    End If
                End If
                strTemp3(4) = GetPrjPeople1(strTemp3(4))
                If Not IsNull(.Fields(17)) Then 'FC代理人
                    strTemp3(5) = .Fields(17)
                Else
                    If Not IsNull(.Fields(18)) Then 'CP44代理人
                        strTemp3(5) = .Fields(18)
                    End If
                End If
        
              '代理人名稱
              If PUB_GetAgentName(SystemNumber(strTemp3(2), 1), strTemp3(5), strTempName) = True Then
                   strTemp3(5) = strTempName
              Else
                   strTemp3(5) = ""
              End If
              strTemp3(12) = "" & .Fields("NO").Value '加分所號
              
                '記錄業務區
                strSalesGrp = "" & .Fields("D").Value
                strSalesGrpName = "" & .Fields("E").Value
                strSName = GetStaffName(strTemp3(0), True)

                iK = iK + 1
                '記錄業務區筆數
                iKK = iKK + 1
                iTatle = iTatle + 1
                If Len(strTemp3(1)) > 7 Then
                    strTemp3(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(1)))
                End If
                If Len(strTemp3(6)) > 7 Then
                    strTemp3(6) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(6)))
                End If
                If Len(strTemp3(7)) > 7 Then
                    strTemp3(7) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(7)))
                End If

                '新增E-MAIL.TXT
                If St <> "" & .Fields("A") And eFile > 0 Then  '不同業務
                   Print #eFile, String(135, "-")
                   Print #eFile, String(10, " ") & "小計： " & iK - 1 & " 筆"
                   Print #eFile, String(135, "-")
                   iK = 1
                End If
                'Modified by Lydia 2022/11/25
                'If eFilename <> "" And eFile > 0 And ((iCall = "1" And St <> "" & .Fields("A")) Or (InStr("2,3", iCall) > 0 And eFilename <> IIf(InStr("2,3", iCall) > 0, strFileN & "_" & .Fields("E").Value, strFileN & "_" & .Fields("E").Value & "_" & strSName))) Then
                If iCall = 2 Or iCall = 3 Then
                   strExc(10) = strFileN & "_" & .Fields("E").Value & IIf("" & .Fields("D").Value = "F231", "英文組", IIf("" & .Fields("D").Value = "F232", "日文組", ""))
                Else
                   strExc(10) = ""
                End If
                If eFilename <> "" And eFile > 0 And ((iCall = "1" And St <> "" & .Fields("A")) Or (InStr("2,3", iCall) > 0 And eFilename <> strExc(10))) Then
                'end 2022/11/25
                   If iCall <> "1" Then
                      Print #eFile, strSGN & "合計： " & iKK - 1 & " 筆"
                      Print #eFile, String(135, "-")
                   End If
                   iKK = 1
                   Close eFile
                   'Modified by Lydia 2022/11/25 若有外專F2x／外商資料F1x，則增加通知
                   'Call StrMenu17_Mailto("1", strSG, St, eFilename, strSalesGrp, Trim(iCall), mailFList)
                   Call StrMenu17_Mailto("1", strSG, St, eFilename, strSalesGrp, Trim(iCall), mailFList, strAddFCP & strAddFCT, strAddFCPcc)
                   eFilename = ""
                End If
                If eFilename = "" Then
                   'Added by Lydia 2022/11/25
                   If iCall <> 3 Then
                      strAddFCP = ""
                      strAddFCT = ""
                      strAddFCPcc = ""
                   End If
                   'end 2022/11/25
                   'Modified by Lydia 2022/11/25
                   'eFilename = IIf(InStr("2,3", iCall) > 0, strFileN & "_" & .Fields("E").Value, strFileN & "_" & .Fields("E").Value & "_" & strSName)
                   If iCall = 1 Then
                       eFilename = strFileN & "_" & .Fields("E").Value & "_" & strSName
                   Else
                       eFilename = strFileN & "_" & .Fields("E").Value & IIf("" & .Fields("D").Value = "F231", "英文組", IIf("" & .Fields("D").Value = "F232", "日文組", ""))
                   End If
                   TmpArea = "業務區：" & .Fields("E").Value & IIf("" & .Fields("D").Value = "F231", "英文組", IIf("" & .Fields("D").Value = "F232", "日文組", ""))
                   'end 2022/11/25
                   eFilename = PUB_UniToBIG5(eFilename, "F") 'Added by Lydia 2022/03/28 員工名稱有Unicode
                   eFile = FreeFile
                   If eFile > 0 Then Close #eFile
                   eFile = FreeFile
                   Open strPath & eFilename & ".txt" For Output As eFile
                   Select Case iCall
                        Case 1:  strExc(3) = "管制時段：個人"
                        Case 2:  strExc(3) = "管制時段：部門主管"
                        Case 3:  strExc(3) = "管制時段：副總"
                        Case Else: strExc(3) = ""
                   End Select

                   Print #eFile, String(40, " ") & strFTitle
                   Print #eFile, "列印順序：" & convForm("智權人員", 80) & "列印日期：" & ChangeTStringToTDateString(TransDate(strSrvDate(1), 1))
                   Print #eFile, TmpArea '業務區＆部門
                   Print #eFile, "專利收文日期：" & ChangeTStringToTDateString(TransDate(strDate_S, 1)) & "-" & ChangeTStringToTDateString(TransDate(strDate_E, 1))

                   '列印順序：智權人員
                   Print #eFile, "智權人員 收文日　  本所案號　　    案件名稱 申請人       本所期限  法定期限  種類     案件性質 承辦人   申請國家 未收金額 分所號　　  "
                   Print #eFile, "======== ========= =============== ======== ============ ========= ========= ======== ======== ======== ======== ======== ============"
                End If
                
                '0~3
                strExc(1) = convForm(strSName, 8) & " " & convForm(strTemp3(1), 9) & " " & convForm(strTemp3(2), 15) & " " & convForm(strTemp3(3), 8)
                '4~8 ,strTemp3(5)=代理人
                strExc(1) = strExc(1) & " " & convForm(strTemp3(4), 12) & " " & convForm(strTemp3(6), 9) & " " & convForm(strTemp3(7), 9) & " " & convForm(strTemp3(8), 8)
                '9~12
                strExc(1) = strExc(1) & " " & convForm(strTemp3(9), 8) & " " & convForm(strTemp3(10), 8) & " " & convForm(strTemp3(11), 8) & " " & PUB_StrToStr(Format(.Fields(21).Value, "#,###"), 8, True, True) & " " & convForm(strTemp3(12), 12)
                Print #eFile, strExc(1)
    
                St = "" & .Fields("A"): strSG = strSalesGrp: strSGN = strSalesGrpName

                'Added by Lydia 2022/11/25 專利國外部：逾七個月副本給總經理、逾八個月給總經理。正本不變。
                '專利日本部：逾七個月副本給簡偉倫經理、逾八個月給簡偉倫經理。
                If iCall = 2 And strSrvDate(1) >= "20221130" And strAddFCPcc = "" Then   '逾七個月才有副本
                   If strSalesGrp = "F231" Then '國外部
                       strAddFCPcc = ";94007"
                   ElseIf strSalesGrp = "F232" Then
                       strAddFCPcc = ";" & Pub_GetSpecMan("S")
                   End If
                End If
                'end 2022/11/25
                'Move by Lydia 2022/11/25 從strSName = GetStaffName( 下方移過來
                'Added by Lydia 2020/03/04 若有外專F2x／外商資料F1x，則增加通知王協理(外專)和江協理(外商)
                If iCall = 3 Then '逾八個月
                    'Modified by Lydia 2022/11/25
                    'If strAddFCP = "" And Left(strSalesGrp, 2) = "F2" Then
                    If Left(strSalesGrp, 2) = "F2" Then
                        'Added by Lydia 2022/11/25 專利國外部：逾七個月副本給總經理、逾八個月給總經理。正本不變。
                        '專利日本部：逾七個月副本給簡偉倫經理、逾八個月給簡偉倫經理。
                        If strSrvDate(1) >= "20221130" Then
                             '國外部
                            If strSalesGrp = "F231" And InStr(strAddFCP & ",", "94007") = 0 Then
                               strAddFCP = strAddFCP & ";94007"
                            '日本部
                            ElseIf strSalesGrp = "F232" And InStr(strAddFCP & ",", Pub_GetSpecMan("S")) = 0 Then
                               strAddFCP = strAddFCP & ";" & Pub_GetSpecMan("S")
                            End If
                        Else
                        'end 2022/11/25
                            strAddFCP = ";88003"  '王協理(王文安)
                        End If 'Added by Lydia 2022/11/25
                    End If
                    If strAddFCT = "" And Left(strSalesGrp, 2) = "F1" Then
                        strAddFCT = ";98020" '江協理(江郁仁)
                    End If
                End If
                'end 2020/03/04
                'end --- Move by Lydia 2022/11/25 從strSName = GetStaffName( 下方移過來
                .MoveNext
            Loop
        End With

        '新增E-MAIL.TXT
        If St <> "" And eFile > 0 Then
           Print #eFile, String(135, "-")
           Print #eFile, String(10, " ") & "小計： " & iK & " 筆"
           Print #eFile, String(135, "-")
        End If
        If InStr("2,3", iCall) > 0 And eFile > 0 Then
           Print #eFile, strSGN & "合計： " & iKK & " 筆"
           Print #eFile, String(135, "-")
        End If
        Close eFile
        'Modified by Lydia 2020/03/04 若有外專F2x／外商資料F1x，則增加通知王協理(外專)和江協理(外商)
        'Call StrMenu17_Mailto("2", strSG, St, eFilename, strSalesGrp, Trim(iCall), mailFList)
        'Modified by Lydia 2022/11/25 增加另外CC => strAddFCPcc
        Call StrMenu17_Mailto("2", strSG, St, eFilename, strSalesGrp, Trim(iCall), mailFList, strAddFCP & strAddFCT, strAddFCPcc)

    End If
Next iCall

ErrHandle:
   If Err.Number <> 0 Then
      WLog "國內專利收文未發文明細表:" & Err.Description
   End If
   Set rsAD = Nothing
End Sub

'Added by Lydia 2017/02/22 國內專利收文未發文明細表-判斷收件人和發mail
'sKind=1 讀檔中;    sKind=2 最後一筆
'stArea=部門代號
'pList 多個附加檔寄出,區隔用*
'Modified by Lydia 2020/03/04 +pAddTo 另外通知
'Modified by Lydia 2022/11/25 +pAddCC 另外CC
Private Sub StrMenu17_Mailto(ByVal sKind As String, ByVal stArea As String, ByVal stTO As String, ByVal fName As String, ByVal nArea As String, ByVal iTyp As String, ByRef pList As String, Optional ByVal pAddTo As String, Optional ByVal pAddCC As String)
Dim stCC As String, Str01 As String, Str02 As String
Dim strPS As String

  
strPS = vbCrLf & vbCrLf & vbCrLf & String(2, "　") & "列印注意事項：" & vbCrLf & _
      vbCrLf & String(4, "　") & "1.利用記事本開啟附件" & _
      vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
      vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
      vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
      vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
      vbCrLf & String(4, "　") & "6.選擇<橫印>  "
      
    stCC = ""
    Select Case iTyp
        Case "1" '逾六個月
              Str01 = PUB_GetST03(stTO)
        Case "2" '逾七個月
              'Modified by Lydia 2022/11/25 增加判斷專利國外部及日本部收文案件
              'Str01 = GetDeptA09(stArea, "08")
              Str01 = GetDeptA09(Mid(stArea, 1, 3), "08")
              If Len(Str01) > 0 Then stTO = Str01
              'Mark by Lydia 2022/11/25 已在StrMenu17設定好
              'If Left(stArea, 1) = "F" Then
              '   stCC = Pub_GetSpecMan("O")
              'End If
              'end 2022/11/25
              
        Case "3" '逾八個月: 統一整合只發一封mail
             'Modified by Morgan 2019/9/19 68009->68006
             'modify by sonia 2020/1/9 +69005
             'modify by sonia 2020/3/4 -68006退休且已加入69005
             'modify by sonia 2022/2/7 +82026
             'Modified by Lydia 2022/05/03 簡協理69005改為抓系統特殊設定「全所智權部主管」、林協理82026改為抓系統特殊設定「中所智權部主管」
             'stTO = "94007;69005;82026"
             stTO = "94007;" & Pub_GetSpecMan("全所智權部主管") & ";" & Pub_GetSpecMan("中所智權部主管")
             stTO = Replace(stTO, ";;", ";")
             'end 2022/05/03
        Case Else
             stTO = ""
    End Select
    
    'Added by Lydia 2020/03/04 另外通知若有外專F2x／外商資料F1x，則增加通知王協理(外專)和江協理(外商)
    If pAddTo <> "" Then
       stTO = stTO & pAddTo
    End If
    'end 2020/03/04
    'Added by Lydia 2022/11/25 另外CC
    If pAddCC <> "" Then
        stCC = stCC & pAddCC
        If Left(stCC, 1) = ";" Then stCC = Mid(stCC, 2)
    End If
    'end 2022/11/25
    
    If iTyp = "3" Or (iTyp = "2" And stCC <> "") Then
        pList = pList & "*" & App.path & TextPath & fName & ".txt" '多個附件的區隔用*
    End If
    
    'Added by Lydia 2022/05/25 每月、每日批次發通知無收受者時改發「程式管理人員」，以利查覺異常情形，才能儘早修改系統設定
    If stTO = "" Then
        stTO = Pub_GetSpecMan("程式管理人員")
    End If
    'end 2022/05/25

    If iTyp = "1" Or iTyp = "2" Then
        'Modified by Lydia 2022/11/25
        'If fName <> "" And stTO <> "" And stCC = "" Then
        '   SendMAPIMail stTO, fName, strPS, App.path & TextPath & fName & ".txt"
        'End If
        ''F開頭部門寄給系統特殊人員("O")大寫,但依部門之前二碼為部門檔案控制
        'If stCC <> "" And (sKind = "2" Or (Left(stArea, 2) <> Left(nArea, 2))) Then
        '   SendMAPIMail stTO, fName, strPS, IIf(Left(pList, 1) = "*", Mid(pList, 2, Len(pList) - 1), pList), , stCC
        '   pList = ""
        'End If
        If fName <> "" And stTO <> "" Then
            SendMAPIMail stTO, fName, strPS, App.path & TextPath & fName & ".txt", , stCC
        End If
        'end 2022/11/25
    ElseIf iTyp = "3" And sKind = "2" Then '逾八個月: 統一整合只發一封mail
           'Modified by Lydioa 2022/11/25 +CC
           'SendMAPIMail stTO, TransDate(strSrvDate(1), 1) & "_專利收文逾八個月未發文明細表", strPS, IIf(Left(pList, 1) = "*", Mid(pList, 2, Len(pList) - 1), pList)
           SendMAPIMail stTO, TransDate(strSrvDate(1), 1) & "_專利收文逾八個月未發文明細表", strPS, IIf(Left(pList, 1) = "*", Mid(pList, 2, Len(pList) - 1), pList), , stCC
    End If
End Sub

'Added by Morgan 2018/5/21
'FMP案已發文未輸完稿日報表
Private Sub StrMenu20()

   Dim strDate As String, strRptDate As String
   Dim strStaffNo As String, strStaffName As String
   Dim strRptName As String, strFileName As String
   Dim strTitle As String, strTitleLine As String
   Dim ff As Integer, iRecs As Integer
   Dim strPS As String, strText As String, strContent As String
      
On Error GoTo ErrHandle

   '6個月內發文者
   strDate = CompDate(1, -6, strSrvDate(1))
   strDate = Left(strDate, 6) & "01"
   If strDate < "20180101" Then strDate = "20180101"
   strRptDate = Format(strSrvDate(2), "@@@/@@/@@")
   
   '1.案件性質表有設定外專程序考核點數 + 2.外專工程師承辦 + 3.非外專程序發文 的FMP案AB類已發文進度
   'Modified by Morgan 2025/10/2 排除有同日發文"陳述意見205"或"復審申請107"的B類"補正204"--敏莉
   strExc(0) = "select na16,s3.st02 SName1,sqldatet(cp27) DDate,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNo" & _
      ",cp14||' '||s2.st02 SName2,cp10,cpm04" & _
      " from caseprogress a,staff s1,engineerprogress,staff s2,casepropertymap,patent,fagent,nation,staff s3" & _
      " Where cp27 >= " & strDate & " and cp01 in ('P','PS','CFP','CPS') and cp09<'C'" & _
      " and cp12 like 'F%' and s1.st01(+)=cp83 and s1.st03<>'F22'" & _
      " and not exists(select * from caseprogress b where cp10 in ('205','107') and cp27=a.cp27 and a.cp10='204' and a.cp09>'B')" & _
      " and ep02(+)=cp09 and ep09 is null" & _
      " and s2.st01(+)=cp14 and s2.st03 like 'F%'" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and cpm31>0" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)" & _
      " and na01(+)=fa10 and s3.st01(+)=na16" & _
      " order by 1,2,3,4"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strPS = vbCrLf & vbCrLf & vbCrLf & String(2, "　") & "列印注意事項：" & vbCrLf & _
      vbCrLf & String(4, "　") & "1.利用記事本開啟附件" & _
      vbCrLf & String(4, "　") & "2.取消<自動換行>設定" & _
      vbCrLf & String(4, "　") & "3.將視窗展開到可顯示所有資料"

      strRptName = "FMP案已發文未輸完稿日報表"
          strTitle = "發文日    本所案號        案件性質                 承辦人"
      strTitleLine = "========= =============== ======================== ============== "
      
      strStaffNo = "" & RsTemp("na16")
      strStaffName = "" & RsTemp("SName1")
      strFileName = App.path & TextPath & strRptName & "(" & strStaffName & ").txt"
      strFileName = PUB_UniToBIG5(strFileName, "F") 'Added by Lydia 2022/03/28 員工名稱有Unicode
      strText = "敬閱者：" & vbCrLf & vbCrLf & String(2, "　") & _
                     strRptName & "(" & strRptDate & ") 資料如附件！" & _
                     vbCrLf & "請管制人去〔FMP案完稿日/核稿完成日輸入〕補上完稿日及時數。" & _
                     vbCrLf & vbCrLf & String(25, "　") & "電腦中心" & strPS
                     
      ff = FreeFile
      Open strFileName For Output As ff
            
      Print #ff, ""
      Print #ff, "日期：" & strRptDate
      Print #ff, ""
      Print #ff, strTitle
      Print #ff, strTitleLine
      
      With RsTemp
      iRecs = 0
      Do While Not .EOF
         If strStaffNo <> "" & RsTemp("na16") Then
            Print #ff, strTitleLine
            Print #ff, "共" & iRecs & "筆"
            Close ff
            SendMAPIMail strStaffNo, strRptName & "(" & strStaffName & ")", strText, strFileName
            
            strStaffNo = "" & RsTemp("na16")
            strStaffName = "" & RsTemp("SName1")
            strFileName = App.path & TextPath & strRptName & "(" & strStaffName & ").txt"
            strFileName = PUB_UniToBIG5(strFileName, "F") 'Added by Lydia 2022/03/28 員工名稱有Unicode
            ff = FreeFile
            Open strFileName For Output As ff
                        
            Print #ff, ""
            Print #ff, "日期：" & strRptDate
            Print #ff, ""
            Print #ff, strTitle
            Print #ff, strTitleLine
            iRecs = 0
         End If
         
         iRecs = iRecs + 1
         strContent = convForm("" & .Fields("DDate"), 9)   '發文日
         strContent = strContent & " " & convForm("" & .Fields("CNo"), 15) '本所案號
         strContent = strContent & " " & convForm("" & .Fields("cpm04"), 24) '案件性質
         strContent = strContent & " " & convForm("" & .Fields("SName2"), 14)  '承辦人
         Print #ff, strContent
         
         .MoveNext
      Loop
      Print #ff, strTitleLine
      Print #ff, "共" & iRecs & "筆"
      Close ff
      SendMAPIMail strStaffNo, strRptName & "(" & strStaffName & ")", strText, strFileName
      
      End With
   End If
   
ErrHandle:
   If Err.Number <> 0 Then
      WLog strFileName & " " & Err.Description
   End If
   
End Sub

'Added by Lydia 2017/06/27 FCT案超過2個月尚未列印核准定稿通知
'Mark by Lydia 2023/08/01 取消管制
'Private Sub StrMenu18()
'Dim rsRd As New ADODB.Recordset
'Dim intR As Integer
'Dim TempFileName As String
'Dim ff18 As Integer
'Dim strTemp(0 To 6) As String
'Dim strDate As String
'Dim strTitle As String, strTitleLine As String
'
''系統日之前2個月
'strDate = CompDate(1, -2, strSrvDate(1))
'
'TempFileName = "FCT案超過2個月尚未列印核准定稿"
'    strTitle = "智權人員 承辦人   本所案號        收文日    案件名稱             案件性質 客戶名稱            "
'strTitleLine = "======== ======== =============== ========= ==================== ======== ===================="
'If Dir(App.path & TextPath & TempFileName & ".txt") <> "" Then
'   Kill App.path & TextPath & TempFileName & ".txt"
'End If
'
'On Error GoTo ErrHandle
'
'   'FCT延展、移轉、變更核准定稿,超過2個月未列印,發通知信
'   strSql = " select st02,c2.cp01,c2.cp02,c2.cp03,c2.cp04,sqldatet(c2.cp05) cp05,c2.cp09,cpm03,nvl(tm05,nvl(tm06,tm07)) casename,tm23 cu01,nvl(cu05,nvl(cu06,cu07)) cu01n" & _
'            " from letterdemand,trademark,customer,caseprogress c1,caseprogress c2,casepropertymap,staff" & _
'            " where ld02='99999999' and ld05='FCT' and ld04=c1.cp09(+) and c1.cp10 in ('102','501','301')" & _
'            " and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and c1.cp09=c2.cp43(+)" & _
'            " and c1.cp01=cpm01(+) and c1.cp10=cpm02(+) and c2.cp14=st01(+) and c1.cp01=tm01(+) and c1.cp02=tm02(+) and c1.cp03=tm03(+) and c1.cp04=tm04(+)" & _
'            " and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and c2.cp10='1001' and c2.cp05<=" & strDate & _
'            " group by c2.cp01,c2.cp02,c2.cp03,c2.cp04,sqldatet(c2.cp05),c2.cp09,cpm03,st02,nvl(tm05,nvl(tm06,tm07)),tm23,nvl(cu05,nvl(cu06,cu07))" & _
'            " order by 6,2,3,4,5"
'
'   rsRd.CursorLocation = adUseClient
'   rsRd.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsRd.RecordCount > 0 Then
'      ff18 = 0
'      rsRd.MoveFirst
'      ff18 = FreeFile
'      Open App.path & TextPath & TempFileName & ".txt" For Output As ff18
'      Print #ff18, String(25, " ") & TempFileName & "明細表"
'      Print #ff18, ""
'      Print #ff18, strTitle
'      Print #ff18, strTitleLine
'
'      With rsRd
'         Do While Not .EOF
'            strTemp(0) = convForm(GetStaffName(PUB_GetFCTSalesNo(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"))), 8)
'            strTemp(1) = convForm("" & .Fields("st02"), 8)
'            strTemp(2) = convForm(.Fields("cp01") & "-" & .Fields("cp02") & "-" & .Fields("cp03") & "-" & .Fields("cp04"), 15)
'            strTemp(3) = convForm("" & .Fields("cp05"), 9)
'            strTemp(4) = convForm("" & .Fields("casename"), 20)
'            strTemp(5) = convForm("" & .Fields("cpm03"), 8)
'            strTemp(6) = convForm("" & .Fields("cu01n"), 20)
'
'            Print #ff18, strTemp(0) & " " & strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & " " & strTemp(6)
'            intR = intR + 1
'            .MoveNext
'         Loop
'      End With
'
'      If ff18 > 0 Then
'         Print #ff18, String(Len(strTitleLine), "=")
'         Print #ff18, "共 " & intR & " 筆"
'         Close ff18
'
'         strExc(1) = vbCrLf & vbCrLf & vbCrLf & String(2, "　") & "列印注意事項：" & vbCrLf & _
'               vbCrLf & String(4, "　") & "1.利用筆記本開啟附件" & _
'               vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
'               vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
'               vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
'               vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
'               vbCrLf & String(4, "　") & "6.選擇<橫印>   "
'         '寄信給陳金蓮72012
'         SendMAPIMail "72012", ChangeTStringToTDateString(TransDate(strSrvDate(1), 1)) & "_" & TempFileName & "通知", strExc(1), App.path & TextPath & TempFileName & ".txt"
'      End If
'   End If
'
'ErrHandle:
'   If Err.Number <> 0 Then
'      WLog "FCT案超過2個月尚未列印核准定稿通知:" & Err.Description
'   End If
'   Set rsRd = Nothing
'End Sub
'end---'Mark by Lydia 2023/08/01 取消管制

'Add By Sindy 2018/2/12 勞工健康檢查通知函
'於每年3/1及10/1日由系統自動發送通知函給該年度應實施勞工健檢的同仁,並將通知函的名單寄給劉經理
Private Sub StrMenu19()
Dim rsRD As New ADODB.Recordset
Dim intR As Integer
Dim TempFileName As String
Dim ff18 As Integer
Dim strTemp(0 To 6) As String
Dim strYear As String
Dim strTitle As String, strTitleLine As String
Dim tmpnext As String
Dim strContext As String, strTo As String
   
   'Add By Sindy 2021/2/17 劉經理: 系統原訂3/1會發email給110年度須要做勞工健康檢查的通知函，因實施上有改變，暫停3/1通知函發送
   'Modify By Sindy 2021/9/6  劉經理: 取消10/1將由系統發出的勞工健康檢查通知, 今年因疫情不通知
   'Modify By Sindy 2022/1/25  劉經理: 請再取消今年3/1將由系統發出的勞工健康檢查通知, 除了疫情外, 今年規劃全體員工實施勞工健檢
   'Modify By Sindy 2022/9/8 劉經理: 請再取消今年10/1將由系統發出的勞工健康檢查通知, 因今年已請全體員工實施勞工健檢
   If Val(Left(strSrvDate(2), 5)) = 11003 Or Val(Left(strSrvDate(2), 5)) = 11010 _
      Or Val(Left(strSrvDate(2), 5)) = 11103 Or Val(Left(strSrvDate(2), 5)) = 11110 Then
      Exit Sub
   End If
   '2021/2/17 END
   
   '系統日之年度
   strYear = Left(strSrvDate(2), 3)
   '每年3/1及10/1日由系統自動發送通知函
   If Val(Mid(strSrvDate(2), 4, 2)) <> 3 And Val(Mid(strSrvDate(2), 4, 2)) <> 10 Then Exit Sub
   
   TempFileName = strYear & "年應繳健檢報告清單"
       strTitle = "部門       員工編號 姓名         目前年齡 上次健檢日期 下次應繳年度"
   strTitleLine = "========== ======== ============ ======== ============ ============"
   
   'Modify By Sindy 2020/9/16 + 依職業安全法第20條、第46條規定，勞工對於健康檢查，有接受的義務，因此不得拒絕檢查。違反者，勞工將處3,000元以下罰鍰。
'   strContext = "各位同仁大家好:" & vbCrLf & vbCrLf & _
'                "依員工健康檢查實施辦法，您今年應實施勞工健康檢查，並繳交健康檢查報告，" & vbCrLf & _
'                "請重視自已的健康，前往勞工健檢醫療院所實施健檢，將健檢報告擲回人事處存查，" & vbCrLf & _
'                "並以單據辦理補助或核銷事宜。" & vbCrLf & vbCrLf & _
'                "依職業安全法第20條、第46條規定，勞工對於健康檢查，有接受的義務，因此不得拒絕檢查。" & vbCrLf & _
'                "違反者，勞工將處3,000元以下罰鍰。" & vbCrLf & vbCrLf & _
'                "謝謝。" & vbCrLf & vbCrLf & _
'                "人事處啟"
   'Modify By Sindy 2021/1/11 人事處調整內容
   strContext = "各位同仁大家好：" & vbCrLf & vbCrLf & _
                "依員工健康檢查實施辦法，您今年應實施勞工健康檢查，並繳交健康檢查報告，" & vbCrLf & _
                "請重視自已的健康，前往勞工健檢醫療院所實施健檢，將健檢報告擲回人事處存查，" & vbCrLf & _
                "並以單據辦理補助或核銷事宜。" & vbCrLf & vbCrLf & _
                "說明：" & vbCrLf & _
                "依勞工健康保護規則第15條規定" & vbCrLf & _
                "雇主對在職勞工，應依下列規定，定期實施一般健康檢查：" & vbCrLf & _
                "一、年滿六十五歲者，每年檢查一次。" & vbCrLf & _
                "二、四十歲以上未滿六十五歲者，每三年檢查一次。" & vbCrLf & _
                "三、未滿四十歲者，每五年檢查一次。" & vbCrLf & vbCrLf & _
                "職業安全衛生法第20條第6項規定，勞工對於健康檢查，有接受之義務。違反者，依同法第46條規定，可處新臺幣3千元以下罰鍰。" & vbCrLf & _
                "敬請同仁留意，謝謝。" & vbCrLf & vbCrLf & _
                "人事處啟" & vbCrLf
                
   strExc(10) = vbCrLf & vbCrLf & vbCrLf & String(2, "　") & "列印注意事項：" & vbCrLf & _
                vbCrLf & String(4, "　") & "1.利用筆記本開啟附件" & _
                vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
                vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
                vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
                vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
                vbCrLf & String(4, "　") & "6.選擇<橫印>   "
   
On Error GoTo ErrHandle
   'Added by Lydia 2023/12/20
'   If strSrvDate(1) >= 新部門啟用日 Then
      'Modify By Sindy 2024/3/11 有關勞工健檢通知函，請協助排除以下人員。
      '一、智慧所：
      '1. 63001林晉章
      '2. 67004林大熙
      '3. 81040閻啟泰
      '4. 94007林景郁
      '二、法律所：
      '1. 79037 蔣文正
      '2. 98020 江郁仁
      '三、當年度年滿65歲同仁。
      strSql = "select nvl(a0922,a0902) a0902,st01,st02,decode(nvl(st23,''),'','',to_char(sysdate,'YYYY')-substr(st23,1,4)),sqldatet(sh02),decode(nvl(st68,0),'','',st68-1911)" & _
               " FROM staff,(select sh01,max(sh02) sh02 from staff_health group by sh01),acc090,acc090new" & _
               " where substr(st01,1,1) in(" & ST01CodeNum1 & ")" & _
               " and st04='1' and substr(st03,1,1)<>'R'" & _
               " and substr(st01,4,1)<>'9' and st01 not in('60000','96029','96030','63001','67004','81040','94007','79037','98020')" & _
               " and st01=sh01(+) and (ST68 between " & Val(strYear) + 1911 & " and " & Val(strYear) + 1911 & " or ST68 is null)" & _
               " and a0901=st03(+) AND ST93=A0921(+)" & _
               " and " & Val(strYear) + 1911 & " - substr(st23,1,4)<>65" & _
               " order by st03,st01 asc"
'   Else
'   'end 2023/12/20
'      'Modify By Sindy 2019/3/4 健檢通知名單請取消台一投資公司 ==> + and substr(st03,1,1)<>'R'
'      strSql = "select a0902,st01,st02,decode(nvl(st23,''),'','',to_char(sysdate,'YYYY')-substr(st23,1,4)),sqldatet(sh02),decode(nvl(st68,0),'','',st68-1911)" & _
'               " FROM staff,(select sh01,max(sh02) sh02 from staff_health group by sh01),acc090" & _
'               " where substr(st01,1,1) in(" & ST01CodeNum1 & ")" & _
'               " and st04='1' and substr(st03,1,1)<>'R'" & _
'               " and substr(st01,4,1)<>'9' and st01 not in('60000','96029','96030')" & _
'               " and st01=sh01(+) and (ST68 between " & Val(strYear) + 1911 & " and " & Val(strYear) + 1911 & " or ST68 is null)" & _
'               " and a0901=st03(+)" & _
'               " order by st03,st01 asc"
'   End If 'Added by Lydia 2023/12/20
   rsRD.CursorLocation = adUseClient
   rsRD.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsRD.RecordCount > 0 Then
      tmpnext = "已取得資料..."
      ff18 = 0
      rsRD.MoveFirst
      
      If Dir(App.path & TextPath & TempFileName & ".txt") <> "" Then
         Kill App.path & TextPath & TempFileName & ".txt"
      End If
      
      ff18 = FreeFile
      Open App.path & TextPath & TempFileName & ".txt" For Output As ff18
      Print #ff18, String(25, " ") & TempFileName
      Print #ff18, ""
      Print #ff18, strTitle
      Print #ff18, strTitleLine
      
      With rsRD
         Do While Not .EOF
            tmpnext = "讀取明細資料...(" & .Fields(1) & ")"
            strTemp(0) = convForm(CheckStr(.Fields(0)), 10)
            strTo = strTo & ";" & .Fields(1)
            strTemp(1) = convForm(CheckStr(.Fields(1)), 8)
            strTemp(2) = convForm(CheckStr(.Fields(2)), 12)
            strTemp(3) = convForm(CheckStr(.Fields(3)), 8)
            strTemp(4) = convForm(CheckStr(.Fields(4)), 12)
            strTemp(5) = convForm(CheckStr(.Fields(5)), 12)
            
            Print #ff18, strTemp(0) & " " & strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5)
            intR = intR + 1
            .MoveNext
         Loop
      End With
      tmpnext = "進入結束..."
      
      If ff18 > 0 Then
         Print #ff18, String(Len(strTitleLine), "=")
         Print #ff18, "共 " & intR & " 筆"
         Close ff18
         
         tmpnext = "寄通知信給同仁"
         '寄通知信給同仁
         strTo = Mid(strTo, 2)
         '收件者不能空白所以掛劉經理,同仁們放至密件副本(個資法)
         SendMAPIMail Pub_GetSpecMan("國外大陸出差通知人員"), "勞工健康檢查通知函" & IIf(Mid(strSrvDate(2), 4, 2) = 10, "　(第二次通知)", ""), strContext, , , , strTo
         
         tmpnext = "結束,寄通知信給劉經理"
         '寄信給劉經理
         'Modify By Sindy 2023/3/1 劉經理說:此份通知函請加通知嘉渝(A8034)
         SendMAPIMail Pub_GetSpecMan("國外大陸出差通知人員"), TempFileName & " 通知", strExc(10), App.path & TextPath & TempFileName & ".txt"
      End If
   End If
   
ErrHandle:
   If Err.Number <> 0 Then
      tmpnext = tmpnext & " " & Err.Description
      WLog tmpnext
      
      If ff18 > 0 Then
         Print #ff18, String(Len(strTitleLine), "=")
         Print #ff18, "共 " & intR & " 筆　(尚未通知完畢...)"
         Print #ff18, tmpnext
         Close ff18
         
         '寄信給97038
         SendMAPIMail "97038", ChangeTStringToTDateString(TransDate(strSrvDate(1), 1)) & " " & TempFileName & " 通知　(尚未通知完畢...)", tmpnext & vbCrLf & vbCrLf & vbCrLf & strExc(10), App.path & TextPath & TempFileName & ".txt"
      End If
   End If
   Set rsRD = Nothing
End Sub

'Added by Lydia 2018/08/22 應收帳款逾付款週期管制表
'Mark by Lydia 2025/07/08 改成每日批次frmAutoBatchDay.StrMenu143
'Private Sub StrMenu21()
'Dim rsRd As New ADODB.Recordset
'Dim inR As Integer, iRow As Integer
'Dim dTotAmt As Double
''Modified by Lydia 2018/09/19
''Dim strTemp(0 To 7) As String
''Modified by Lydia 2018/11/05 +最後收文日
''Dim strTemp(0 To 8) As String
'Dim strTemp(0 To 9) As String
'Dim arrTmp1 As Variant, arrTmp2 As Variant
'Dim TempFileName As String
'Dim ff21 As Integer
'Dim strGrp As String
'Dim strTo As String, strPS As String
'Dim strTitle1 As String, strTitle2 As String
'Dim intW As Integer
'Dim iRound As Integer  'Added by Lydia 2018/09/19 分兩種報表(1-各區,2-全部)
'
''判斷每季第一個月1號通知
'If Val(Mid(strSrvDate(1), 5, 2)) Mod 3 <> 1 Then
'     Exit Sub
'End If
'
'On Error GoTo ErrHandle 'Move by Lydia 2018/09/19 從下方移上來
'
'TempFileName = "應收帳款逾付款週期管制表"
'If Dir(App.path & TextPath & TempFileName & ".txt") <> "" Then
'    Kill App.path & TextPath & TempFileName & ".txt"
'End If
'
'strPS = vbCrLf & vbCrLf & vbCrLf & String(2, "　") & "列印注意事項：" & vbCrLf & _
'      vbCrLf & String(4, "　") & "1.利用筆記本開啟附件" & _
'      vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
'      vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
'      vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
'      vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
'      vbCrLf & String(4, "　") & "6.選擇<橫印>   "
'
''Added by Lydia 2018/09/19 分兩種報表(1-各區,2-全部)
'For iRound = 1 To 2
'    strGrp = ""
'    strTo = ""
'    strTitle1 = ""
'    strTitle2 = ""
''end 2018/09/19
'    '----欄位抬頭
'    'Modified by Lydia 2018/11/05 最後收文日
'    'strExc(1) = "業務區,智權人員,本所案號,案件名稱,客戶名稱,案件性質,收文日,發文日,應收金額"
'    'strExc(2) = "12,8,15,26,26,10,10,10,10'"
'    strExc(1) = "業務區,智權人員,本所案號,案件名稱,客戶名稱,案件性質,收文日,發文日,應收金額,最後收文日"
'    strExc(2) = "12,8,15,26,26,10,10,10,10,10'"
'    'end 2018/11/05
'    arrTmp1 = Split(strExc(1), ",")
'    arrTmp2 = Split(strExc(2), ",")
'
'    'Modified by Lydia 2018/09/19
'    'For inR = 0 To UBound(arrTmp1)
'    For inR = IIf(iRound = 1, 1, 0) To UBound(arrTmp1)
'       If Trim(arrTmp1(inR)) <> "" Then
'           strTitle1 = strTitle1 & convForm(arrTmp1(inR), Val(arrTmp2(inR))) & " "
'           strTitle2 = strTitle2 & String(Val(arrTmp2(inR)), "=") & " "
'       End If
'    Next
'    strTitle1 = Mid(strTitle1, 1, Len(strTitle1) - 1)
'    strTitle2 = Mid(strTitle2, 1, Len(strTitle2) - 1)
'    intW = GetTextLength(strTitle1)
'   '已發文且已列印收據之國內應收帳款:超過付款週期(若客戶未設定特殊付款週期CU175，則歸到一般付款週期預設2個月)
'   'Modified by Lydia 2018/09/25 改抓a0k20的業務員代號和收文部門st15
'   'Modified by Lydia 2018/11/05 +增加欄位(sk02,a0k03 as custno)
'   'Modified by Lydia 2025/06/10 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
'   strExc(0) = " select cp01,cp02,cp03,cp04,cp05,cp09,cp10,st15 as cp12,a0k20 as cp13,st02 as sname, cp27,nvl(cu04,nvl(cu06,cu07)) custname,a0k01 as billno, a0k02 as billdate, sum(nvl(a0j09,0)+nvl(a0j10,0)-nvl(a1u04,0)-nvl(a1u05,0)-nvl(a1u07,0)-nvl(a1u09,0)+nvl(a1u08,0)+nvl(a1u10,0)) as A1" & _
'                    " ,sk02,a0k03 as custno From acc0k0, caseprogress, systemkind, acc0j0,customer,staff," & _
'                         " (select a1u02,a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07,sum(a1u09) a1u09,sum(a1u08) a1u08,sum(a1u10) a1u10 From acc1u0 where a1u03" & _
'                         " in ( select distinct a1u03 From acc0k0, caseprogress, acc1u0, acc0j0,customer Where geta0k32type(a0k01)='1' And nvl(a0k09, 0) = 0 And (a0k06 + a0k07) > (nvl(a0k17, 0) + nvl(a0k18, 0))" & _
'                               " and substr(a0k03,1,8)=cu01 and substr(a0k03,9,1)=cu02 and cp27<to_char(add_months(sysdate-1,nvl(cu175,2) * -1),'YYYYMMDD')" & _
'                               " and a0k01=a0j13(+) and a0j01=cp09(+) and nvl(cp27,0)>0 and cp79>0 and a1u03(+)=cp09" & _
'                               " ) group by a1u02,a1u03) X" & _
'                    " Where geta0k32type(a0k01)='1' And nvl(a0k09, 0) = 0 And (a0k06 + a0k07) > (nvl(a0k17, 0) + nvl(a0k18, 0)) and substr(a0k03,1,8)=cu01 and substr(a0k03,9,1)=cu02" & _
'                    " and cp27<to_char(add_months(sysdate-1,nvl(cu175,2) * -1),'YYYYMMDD') and a0k20=st01(+)" & _
'                    " and a0k01=a0j13(+) and a0j01=cp09(+) and nvl(cp27,0)>0 and cp79>0 and X.a1u02(+)=a0j13 and X.a1u03(+)=a0j01 and sk01(+)=cp01" & _
'                    " group by cp01,cp02,cp03,cp04,cp05,cp09,cp10,st15,a0k20,st02,cp27,nvl(cu04,nvl(cu06,cu07)),a0k01, a0k02,sk02,a0k03 "
'   strSql = " select aa.*,cp10,decode(pa01,null,decode(tm01,null,decode(sp01,null,cpm03,decode(sp09,'000',cpm03,cpm04)),decode(tm10,'000',cpm03,cpm04)),decode(pa09,'000',cpm03,cpm04)) as cp10n," & _
'               " nvl(pa05,nvl(pa06,pa07))||nvl(tm05,nvl(tm06,tm07))||nvl(sp05,nvl(sp06,sp07))||nvl(lc05,nvl(lc06,lc07)) as casename,a0902 as deptname,a0908 as deptman" & _
'               " from (" & strExc(0) & " ) AA,patent,trademark,servicepractice,lawcase ,acc090, casepropertymap c1" & _
'               " where pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
'               " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
'               " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
'               " and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04" & _
'               " and a0901(+)=cp12 and cpm01(+)=cp01 and cpm02(+)=cp10"
'   strSql = strSql & " order by cp12,cp13,cp27,cp09"
'   inR = 1
'   Set rsRd = ClsLawReadRstMsg(inR, strSql)
'   If inR = 1 Then
'        rsRd.MoveFirst
'        With rsRd
'           Do While Not .EOF
'              'Modified by Lydia 2018/09/19 改判斷
'              'If strGrp <> "" & .Fields("cp12") Then
'              If (iRound = 1 And strGrp <> "" & .Fields("cp12")) Or _
'                   (iRound = 2 And strGrp = "") Then
'                  If strTo <> "" Then
'                      '寄給各區主管
'                      Print #ff21, String(intW, "=")
'                      strExc(1) = "共 " & iRow & " 筆"
'                      strExc(2) = "小計：" & PUB_StrToStr(Format("" & dTotAmt, DDollar2), Val(arrTmp2(7)), True, True)
'                      Print #ff21, strExc(1) & String(intW - GetTextLength(strExc(1)) - GetTextLength(strExc(2)), " ") & strExc(2)
'                      Close ff21
'                      SendMAPIMail strTo, TempFileName, strPS, App.path & TextPath & TempFileName & ".txt"
'                  End If
'                  ff21 = FreeFile
'                  Open App.path & TextPath & TempFileName & ".txt" For Output As ff21
'                  Print #ff21, String(50, " ") & TempFileName
'                  Print #ff21, ""
'                  If iRound = 1 Then 'Added by Lydia 2018/09/19
'                      Print #ff21, "業務區：" & "" & .Fields("deptname") & String(85, " ") & "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'                   'Added by Lydia 2018/09/19 全部
'                  Else
'                      Print #ff21, "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'                  End If
'                  'end 2018/09/19
'                  'Added by Lydia 2018/11/07 增加註解
'                  Print #ff21, "註1：""最後收文日""指同一客戶編號之最後收文日"
'                  Print #ff21, "註2：""最後收文日""顯示""*******""表示同本案收文日"
'                  'end 2018/11/07
'                  Print #ff21, strTitle1
'                  Print #ff21, strTitle2
'                  dTotAmt = 0
'                  iRow = 0
'              End If
'
'              'Added by Lydia 2018/09/19 業務區
'              strTemp(0) = convForm("" & .Fields("deptname"), Val(arrTmp2(0)))
'              '智權人員
'              strTemp(1) = convForm("" & .Fields("sname"), Val(arrTmp2(1)))
'              '本所案號
'              strTemp(2) = convForm(.Fields("cp01") & "-" & .Fields("cp02") & "-" & .Fields("cp03") & "-" & .Fields("cp04"), Val(arrTmp2(2)))
'              '案件名稱
'              'Modified by Lydia 2018/09/19 去除跳行符號(ex.CFP-023183)
'              'strTemp(3) = convForm("" & .Fields("casename"), Val(arrTmp2(3)))
'              strTemp(3) = convForm(PUB_StringFilter("" & .Fields("casename")), Val(arrTmp2(3)))
'              '客戶名稱
'              strTemp(4) = convForm("" & .Fields("custname"), Val(arrTmp2(4)))
'              '案件性質
'              strTemp(5) = convForm("" & .Fields("cp10n"), Val(arrTmp2(5)))
'              '收文日
'              strTemp(6) = convForm(ChangeTStringToTDateString(TransDate("" & .Fields("cp05"), 1)), Val(arrTmp2(6)))
'              '發文日
'              strTemp(7) = convForm(ChangeTStringToTDateString(TransDate("" & .Fields("cp27"), 1)), Val(arrTmp2(7)))
'              '應收金額
'              strTemp(8) = PUB_StrToStr(Format("" & .Fields("a1"), DDollar2), Val(arrTmp2(8)), True, True)
'              'Added by Lydia 2018/11/05 最後收文日
'              '抓該案件申請人1之A類最大收文日期,但專利 , 商標, 法務顧問分開處理
'              strSql = ""
'              Select Case "" & .Fields("sk02")
'                   Case "1", "5" '專利
'                          strSql = "select pa01 as pno1,pa02 as pno2,pa03 as pno3,pa04 as pno4 from patent,systemkind where pa26='" & .Fields("custno") & "'  and pa01=sk01(+) and sk02='1' " & _
'                                      "union select sp01 as pno1,sp02 as pno2,sp03 as pno3,sp04 as pno4 from servicepractice,systemkind where sp08='" & .Fields("custno") & "' and sp01=sk01(+) and sk02='5' "
'                   Case "2", "6" '商標
'                          strSql = "select tm01 as pno1,tm02 as pno2,tm03 as pno3,tm04 as pno4 from trademark,systemkind where tm23='" & .Fields("custno") & "'  and tm01=sk01(+) and sk02='2' " & _
'                                      "union select sp01 as pno1,sp02 as pno2,sp03 as pno3,sp04 as pno4 from servicepractice,systemkind where sp08='" & .Fields("custno") & "' and sp01=sk01(+) and sk02='6' "
'                   Case "3", "4", "7", "8" '法務顧問
'                          strSql = "select lc01 as pno1,lc02 as pno2,lc03 as pno3,lc04 as pno4 from lawcase,systemkind where lc11='" & .Fields("custno") & "'  and lc01=sk01(+) and sk02 in ('3','4','7','8') " & _
'                                      "union select hc01 as pno1,hc02 as pno2,hc03 as pno3,hc04 as pno4 from hirecase,systemkind where hc05='" & .Fields("custno") & "' and hc01=sk01(+) and sk02 in ('3','4','7','8') "
'              End Select
'              strTemp(9) = convForm("*******", Val(arrTmp2(9))) '預設-表示同本案收文日
'              If strSql <> "" Then
'                   strSql = "select max(cp05) as mdate from caseprogress where cp05>" & .Fields("cp05") & " and substr(cp09,1,1)='A' and cp159=0 " & _
'                                "and (cp01,cp02,cp03,cp04) in (" & strSql & ") "
'                   intI = 1
'                   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                   If Val("" & RsTemp.Fields("mdate")) > 0 Then
'                       strTemp(9) = convForm(ChangeTStringToTDateString(TransDate("" & RsTemp.Fields("mdate"), 1)), Val(arrTmp2(9)))
'                   End If
'              End If
'              'end 2018/11/05
'
'              'Modified by Lydia 2018/09/19
'              'Print #ff21, strTemp(0) & " "; strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & " " & strTemp(6) & " " & strTemp(7)
'              If iRound = 1 Then   '分區
'                    'Modified by Lydia 2018/11/05 +strTemp(9)
'                    Print #ff21, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & " " & strTemp(6) & " " & strTemp(7) & " " & strTemp(8) & " " & strTemp(9)
'              Else     '全部-明細列有業務區
'                    'Modified by Lydia 2018/11/05 +strTemp(9)
'                    Print #ff21, strTemp(0) & " "; strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & " " & strTemp(6) & " " & strTemp(7) & " " & strTemp(8) & " " & strTemp(9)
'              End If
'              'end 2018/09/19
'
'              dTotAmt = dTotAmt + Val("" & .Fields("a1"))
'              iRow = iRow + 1
'              If iRound = 1 Then 'Added by Lydia 2018/09/19 分區
'                  strGrp = "" & .Fields("cp12")
'                  strTo = "" & .Fields("deptman")
'                 'Added by Lydia 2022/05/25 每月、每日批次發通知無收受者時改發「程式管理人員」，以利查覺異常情形，才能儘早修改系統設定
'                 If strTo = "" Then
'                     strTo = Pub_GetSpecMan("程式管理人員")
'                 End If
'                 'end 2022/05/25
'              'Added by Lydia 2018/09/19 全部
'              Else
'                   strGrp = "M01"
'              End If
'              'end 2018/09/19
'
'              .MoveNext
'           Loop
'        End With
'        Print #ff21, String(intW, "=")
'        strExc(1) = "共 " & iRow & " 筆"
'        strExc(2) = "小計：" & PUB_StrToStr(Format("" & dTotAmt, DDollar2), Val(arrTmp2(8)), True, True)
'        Print #ff21, strExc(1) & String(intW - GetTextLength(strExc(1)) - GetTextLength(strExc(2)), " ") & strExc(2)
'        Close ff21
'        'Added by Lydia 2018/09/19 增加完整報表通知總經理和何主秘
'        If iRound = 2 Then
'              strTo = Pub_GetSpecMan("總經理員工編號")
'              'Modified by Morgan 2019/9/19 68009->68006
'              'modify by sonia 2020/1/9 +69005
'              'modify by sonia 2020/3/4 -68006退休,且已加入69005
'              'Modified by Lydia 2022/05/03 簡協理69005改為抓系統特殊設定「全所智權部主管」
'              'strTo = strTo & IIf(strTo <> "", ";", "") & "69005"
'              strTo = strTo & IIf(strTo <> "", ";", "") & Pub_GetSpecMan("全所智權部主管")
'        End If
'        'end 2018/09/19
'        If strTo <> "" Then
'            SendMAPIMail strTo, TempFileName, strPS, App.path & TextPath & TempFileName & ".txt"
'        End If
'   End If
'Next iRound 'end 2018/09/19
'
'ErrHandle:
'   If Err.Number <> 0 Then
'      WLog "應收帳款逾付款週期管制表:" & Err.Description
'   End If
'   Set rsRd = Nothing
'End Sub
'end 2025/07/08

'Added by Morgan 2018/11/30
'刪除3個月前之商標註冊費繳費單pdf檔
'刪除4個月前之專利領證費/延緩公告繳費單pdf檔 'Added by Morgan 2025/4/23
Private Sub StrMenu22()
   Dim oFolder As Folder
   Dim oFile As File
   Dim iFlag As Integer
   
On Error GoTo ErrHnd

   '商標
   iFlag = 1
   Set oFolder = fso.GetFolder(strTFeeForm)
   For Each oFile In oFolder.files
      If oFile.DateLastModified < DateAdd("m", -3, Now) Then
         oFile.Delete True
      End If
   Next
   
   'Added by Morgan 2025/4/23
   '專利
   iFlag = 2
   Set oFolder = fso.GetFolder("\\" & strPat1Path & "\Fee_Form")
   For Each oFile In oFolder.files
      If oFile.DateLastModified < DateAdd("m", -4, Now) Then
         oFile.Delete True
      End If
   Next
   'end 2025/4/23
   
ErrHnd:
   If Err.Number <> 0 Then
      If iFlag = 1 Then
         WLog "刪除3個月前之商標註冊費繳費單pdf檔:" & Err.Description
      Else
         WLog "刪除4個月前之專利領證費/延緩公告繳費單pdf檔:" & Err.Description
      End If
   End If
   Set oFolder = Nothing
End Sub

'Added by Morgan 2019/1/10
'優先權期限資料提供:
'1.八個月前當月發文的發明、新型申請案 (不論是否已經申請過其他國家)
'2.四個月前當月發文的設計申請案 (不論是否已經申請過其他國家)
'3.傳送對象如下:北所邱素蓮(20190412改莊敏惠,20190515再改智權委辦區ip_transfer)、中所陳家欣、南所鄭鈺華、高所謝秀珠。副本傳送北所簡協理、中所林柄佑經理、南所杜經理、高所楊經理(20190430改簡國靜)。
'Modified by Morgan 2023/3/3
'1.EXCEL檔加代表圖
'2.改發個人、部門主管A0908(該部門所有人)、中所智權部主管(限中所智權部資料)、全所智權部主管(限全所智權部資料)。
Private Sub StrMenu23()
   Dim xlsReport As New Excel.Application
   Dim wksReport As New Worksheet
   Dim rsQuery As ADODB.Recordset
   Dim intQ As Integer, stSQL As String, stVTB As String
   Dim xlsFileName As String
   Dim iCol As Integer, iRow As Integer
   Dim arrTmp
   'Modified by Morgan 2023/3/3
   'Dim stZone As String, iZone As String
   'Dim arrRec(4) As String, arrCC(4) As String
   Dim stKeyID As String, stKeyCol As String, stSubj As String, stRcvr As String, IntF As Integer
   'end 2023/3/3
   Dim iMn1 As Integer, iMn2 As Integer
   'Added by Morgan 2023/3/1
   Dim strNo1 As String, strNo2 As String, strNo3 As String, strNo4 As String, stFileName As String
   Dim oWidth As Single, oHeight As Single, wkWidth As Single
   Dim oShape
   'end 2023/3/1
   
   'modify by sonia 2019/4/12 改邱素蓮74028為莊敏惠73017,20190515再改智權委辦區ip_transfer
   'Modified by Lydia 2022/05/03 簡協理69005改為抓系統特殊設定「全所智權部主管」、林協理82026改為抓系統特殊設定「中所智權部主管」
   'arrRec(1) = "ip_transfer": arrCC(1) = "69005"
   'arrRec(2) = "A7016": arrCC(2) = "82026"
   'Removed by Morgan 2023/3/3 改發個人
   'arrRec(1) = "ip_transfer": arrCC(1) = Pub_GetSpecMan("全所智權部主管")
   'arrRec(2) = "A7016": arrCC(2) = Pub_GetSpecMan("中所智權部主管")
   'end 2023/3/3
   'end 2022/05/03
   'Modified by Lydia 2022/05/06 原通知鄭鈺華A5005的部分，請調整通知蘇嫄媛79053--- 杜經理
   'arrRec(3) = "79053" 'Removed by Morgan 2023/3/3 改發個人
   'Modified by Morgan 2023/2/2 南高所的副本改發ACC090之A0908
   'arrCC(3) = "74018"
   'arrCC(3) = GetDeptMan("S31") 'Removed by Morgan 2023/3/3 改發個人
   'end 2023/2/2
   
   'arrRec(4) = "89047" 'Removed by Morgan 2023/3/3 改發個人
   'modify by sonia 2019/4/30 改楊家復為簡國靜87052
   'Modified by Morgan 2023/2/2 南高所的副本改發ACC090之A0908
   'arrCC(4) = "87052"
   'arrCC(4) = GetDeptMan("S41") 'Removed by Morgan 2023/3/3 改發個人
   'end 2023/2/2
   
On Error GoTo ErrHnd
   
   If Left(strSrvDate(1), 6) = "201902" Then
      iMn1 = 11
      iMn2 = 5
   Else
      iMn1 = 8
      iMn2 = 4
   End If
   
   '發明新型案留4個月開拓,設計案留2個月開拓
   '基礎案:1.無申請日更早之國內案或多國案且未主張優先權 2.主張非本所案優先權
   '已申請國家:基礎案之國外案或多國案(依關聯檔)、主張該案優先權或該案主張之非本所案件優先權的申請國
   '發明新型
   stVTB = "select '1' Flg,pa01,pa02,pa03,pa04,pa05,pa08,pa09,pa10,pa11,pa26 from patent,caseprogress" & _
      " where pa10>=to_char(add_months(sysdate,-" & iMn1 & "),'yyyymm')||'01' and pa10<to_char(add_months(sysdate,-7),'yyyymm')||'01'" & _
      " and pa23='1' and pa08 in ('1','2')" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10 in ('101','102') and pa158>0 and substr(CP12,1,1)<>'F'" & _
      " and not exists(select * from pridate where pd01=pa01 and pd02=pa02 and pd03=pa03 and pd04=pa04)"
   '設計
   stVTB = stVTB & " union select '1' Flg,pa01,pa02,pa03,pa04,pa05,pa08,pa09,pa10,pa11,pa26 from patent,caseprogress" & _
      " where pa10>=to_char(add_months(sysdate,-" & iMn2 & "),'yyyymm')||'01' and pa10<to_char(add_months(sysdate,-3),'yyyymm')||'01'" & _
      " and pa23='1' and pa08='3'" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='103' and pa158>0 and substr(CP12,1,1)<>'F'" & _
      " and not exists(select * from pridate where pd01=pa01 and pd02=pa02 and pd03=pa03 and pd04=pa04)"
   '非本所案優先權(發明新型)
   stVTB = stVTB & " union select '2' Flg,pa01,pa02,pa03,pa04,pa05,pa08,pa09,pa10,pa11,pa26" & _
      " from pridate a,patent b,caseprogress where pd01 in ('P','CFP')" & _
      " and pd05>=to_char(add_months(sysdate,-" & iMn1 & "),'yyyymm')||'01'" & _
      " and pd05<to_char(add_months(sysdate,-7),'yyyymm')||'01'" & _
      " and not exists(select * from patent x where x.pa11=pd06 and x.pa09=pd07)" & _
      " and not exists(select * from pridate x where x.pd01=a.pd01 and x.pd02=a.pd02" & _
      " and x.pd03=a.pd03 and x.pd04=a.pd04 and x.pd05<a.pd05)" & _
      " and pa01(+)=pd01 and pa02(+)=pd02 and pa03(+)=pd03 and pa04(+)=pd04 and pa08 in ('1','2')" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10 in ('101','102') and pa158>0 and substr(CP12,1,1)<>'F'"
   '非本所案優先權(設計)
   stVTB = stVTB & " union select '2' Flg,pa01,pa02,pa03,pa04,pa05,pa08,pa09,pa10,pa11,pa26" & _
      " from pridate a,patent b,caseprogress where pd01 in ('P','CFP')" & _
      " and pd05>=to_char(add_months(sysdate,-" & iMn2 & "),'yyyymm')||'01'" & _
      " and pd05<to_char(add_months(sysdate,-3),'yyyymm')||'01'" & _
      " and not exists(select * from patent x where x.pa11=pd06 and x.pa09=pd07)" & _
      " and not exists(select * from pridate x where x.pd01=a.pd01 and x.pd02=a.pd02" & _
      " and x.pd03=a.pd03 and x.pd04=a.pd04 and x.pd05<a.pd05)" & _
      " and pa01(+)=pd01 and pa02(+)=pd02 and pa03(+)=pd03 and pa04(+)=pd04 and pa08='3'" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='103' and pa158>0 and substr(CP12,1,1)<>'F'"
      
   'Modified by Morgan 2023/3/1 +X1,pa01,pa02,pa03,pa04
   'Modified by Morgan 2023/3/3 +a0908
   'Modified by Morgan 2023/5/2 申請人改抓 中->英->日
   'Modified by Morgan 2025/4/8 區別改顯示中文--秀玲
   stSQL = "select a0902,st02 C1,NVL(CU04,DECODE(CU05,NULL,CU06,rtrim(CU05||' '||CU88||' '||CU89||' '||CU90))) C2,'' X1" & _
      ",pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C3" & _
      ",decode(pa08,'1','發明','2','新型','3','設計') C4" & _
      ",pa05 C5,NA03 C6,sqldatet(pa10) C7" & _
      ",GETOTHERAPPCOUNTRY(PA01,PA02,PA03,PA04) C8" & _
      ",decode(st06,'1','北','2','中','3','南','4','高') C9,st06,cu13,pa10,pa01,pa02,pa03,pa04,a0901,cu12,a0908" & _
      ",decode(substr(cu12,1,2),'S2',1,0) S2,decode(substr(cu12,1,1),'S',1,0) S" & _
      " from (" & stVTB & ") x,customer,staff,nation,acc090" & _
      " where cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and st01(+)=cu13 and na01(+)=x.pa09 and a0901(+)=cu12" & _
      " and NOT EXISTS(select * from casemap,patent b where cm01=x.pa01 and cm02=x.pa02" & _
      " and cm03=x.pa03 and cm04=x.pa04 and cm10='0' and b.pa01(+)=cm05 and b.pa02(+)=cm06" & _
      " and b.pa03(+)=cm07 and b.pa04(+)=cm08 and b.pa10<=x.pa10)" & _
      " and NOT EXISTS(select * from casemap,patent b where cm05=x.pa01 and cm06=x.pa02" & _
      " and cm07=x.pa03 and cm08=x.pa04 and cm10='0' and b.pa01(+)=cm01 and b.pa02(+)=cm02" & _
      " and b.pa03(+)=cm03 and b.pa04(+)=cm04 and b.pa10<x.pa10)" & _
      " and NOT EXISTS(select * from caserelation,patent b where cr01=x.pa01 and cr02=x.pa02" & _
      " and b.pa01(+)=cr05 and b.pa02(+)=cr06 and b.pa03(+)=cr07 and b.pa04(+)=cr08 and b.pa10<x.pa10)" & _
      " order by st06,cu12,cu13,C7,C3"
                  
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      
      PUB_KillTempFile "$$*.*" '刪代表圖檔 'Added by Morgan 2023/3/1
      'Modified by Morgan 2023/3/1 +代表圖
      arrTmp = Array("區別", "智權人員", "申請人", "代表圖", "本所案號", "專利種類", "案件名稱", "申請國家", "申請日", "已申請其他國家")
      With rsQuery
      For IntF = 1 To 4
         If IntF = 1 Then '個人
            .MoveFirst
            stKeyCol = "cu13"
            stKeyID = "" & .Fields(stKeyCol)
            stRcvr = stKeyID
            stSubj = "" & .Fields("C1")
         ElseIf IntF = 2 Then '部門主管
            .MoveFirst
            stKeyCol = "cu12"
            stKeyID = "" & .Fields(stKeyCol)
            stRcvr = "" & .Fields("a0908")
            stSubj = "" & .Fields("a0902")
         ElseIf IntF = 3 Then '中所智權部主管
            .MoveFirst
            .Sort = "S2"
            .Find "S2=1"
            stRcvr = Pub_GetSpecMan("中所智權部主管")
            stSubj = "中所智權部"
            stKeyID = stSubj
         ElseIf IntF = 4 Then '全所智權部主管
            .MoveFirst
            .Sort = "S"
            .Find "S=1"
            stRcvr = Pub_GetSpecMan("全所智權部主管")
            stSubj = "全所智權部"
            stKeyID = stSubj
         End If
         
         iRow = 0
         Do While Not .EOF
            'Modified by Morgan 2023/3/3
            'If stZone <> rsQuery.Fields("C9") Then
            If IntF < 3 And stKeyID <> .Fields(stKeyCol) Then '個人/部門主管
            'end 2023/3/3
               If xlsFileName <> "" Then
                  wksReport.Range("A1", Chr(UBound(arrTmp) + 65) & "1").Font.Bold = True
                  wksReport.Columns(Chr(LBound(arrTmp) + 65) & ":" & Chr(UBound(arrTmp) + 65)).EntireColumn.Font.Name = "標楷體"
                  wksReport.Columns(Chr(LBound(arrTmp) + 65) & ":" & Chr(UBound(arrTmp) + 65)).EntireColumn.AutoFit
                  wksReport.Columns("D:D").ColumnWidth = 16.63 'Added by Morgan 2024/3/11 設定代表圖寬度與總簿同
                  If Val(xlsReport.Version) < 12 Then
                     xlsReport.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=-4143
                  Else
                     xlsReport.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=56
                  End If
                  xlsReport.Workbooks.Close
                  xlsReport.Quit
                  'Modified by Morgan 2023/3/3
                  'SendMAPIMail arrRec(iZone), xlsFileName & "(請轉發相關智權人員)", "如旨", App.path & TextPath & xlsFileName, , arrCC(iZone)
                  SendMAPIMail stRcvr, stSubj, "如旨", App.path & TextPath & xlsFileName
                  'end 2023/3/3
               End If
               iRow = 0
               'Modified by Morgan 2023/3/3
               'stZone = rsQuery.Fields("C9") '所別
               'iZone = Val("" & rsQuery.Fields("st06"))
               stKeyID = "" & .Fields(stKeyCol)
               If IntF = 2 Then
                  stRcvr = "" & .Fields("a0908")
                  stSubj = "" & .Fields("a0902")
               Else
                  stRcvr = stKeyID
                 stSubj = "" & .Fields("C1")
               End If
               'end 2023/3/3
            End If
            If iRow = 0 Then
               stSubj = "優先權期限資料(" & stSubj & ")-" & Left(strSrvDate(2), 5)
               xlsFileName = "優先權期限資料(" & stKeyID & ")-" & Left(strSrvDate(2), 5) & ".xls"
               
               If Dir(App.path & TextPath & xlsFileName) <> "" Then
                  Kill App.path & TextPath & xlsFileName
               End If
               
               xlsReport.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/12 預設工作表數目
               xlsReport.Workbooks.add
               xlsReport.Application.WindowState = xlMinimized
               Set wksReport = xlsReport.Worksheets(1)
               
               '設定欄位名稱及欄寬
               iRow = iRow + 1
               For iCol = LBound(arrTmp) To UBound(arrTmp)
                   wksReport.Range(Chr(iCol + 65) & iRow).Value = arrTmp(iCol)
                   wksReport.Range(Chr(iCol + 65) & iRow).HorizontalAlignment = xlCenter
               Next
               wksReport.Columns("D:D").ColumnWidth = 16.63 'Added by Morgan 2024/3/11 設定代表圖寬度與總簿同
               
            End If
            
            iRow = iRow + 1
            For iCol = LBound(arrTmp) To UBound(arrTmp)
               'Added by Morgan 2023/3/1
               '代表圖
               If iCol = 3 Then
                  strNo1 = .Fields("pa01")
                  strNo2 = .Fields("pa02")
                  strNo3 = .Fields("pa03")
                  strNo4 = .Fields("pa04")
                  stFileName = Dir(App.path & "\$$" & strNo1 & strNo2 & strNo3 & strNo4 & ".*")
                  If stFileName = "" Then
                     stFileName = App.path & "\$$" & strNo1 & strNo2 & strNo3 & strNo4
                     If GetImgByteFile_Case(strNo1, strNo2, strNo3, strNo4, stFileName, 0) = False Then
                        stFileName = ""
                     End If
                  Else
                     stFileName = App.path & "\" & stFileName
                  End If
                  If stFileName <> "" Then
                      Set oShape = wksReport.Shapes.AddPicture(FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True, Left:=100, Top:=100, Width:=-1, Height:=-1)
                      'Modified by Morgan 2024/3/11 +縮放比例不可小於0.014或負值，改為設大小 Ex:P-131589
                      'wksReport.Rows(iRow).RowHeight = 110
                      'oWidth = wksReport.Range(Chr(iCol + 65) & iRow).Width / oShape.Width
                      'oHeight = wksReport.Rows(iRow).RowHeight / oShape.Height
                      'oShape.Select
                      'If oWidth > oHeight Then
                      '  xlsReport.Selection.ShapeRange.ScaleWidth Round(oHeight, 2) - 0.02, True, 0  '等比例縮放
                      'Else
                      '  xlsReport.Selection.ShapeRange.ScaleWidth Round(oWidth, 2) - 0.02, True, 0 '-0.02避免覆蓋格線上
                      'End If
                      wksReport.Rows(iRow).RowHeight = 110
                      oHeight = wksReport.Rows(iRow).RowHeight
                      oWidth = wksReport.Range(Chr(iCol + 65) & iRow).Width
                      oShape.LockAspectRatio = -1 'msoTrue
                      If oShape.Height / oHeight >= oShape.Width / oWidth Then
                        oShape.Height = oHeight - 10
                      Else
                        oShape.Width = oWidth - 10
                      End If
                      'end 2024/3/11
                      oShape.Top = wksReport.Range(Chr(iCol + 65) & iRow).Top + 5
                      oShape.Left = wksReport.Columns(Chr(iCol + 65)).Left + 5
                  End If
               Else
               'end 2023/3/1
                  wksReport.Range(Chr(iCol + 65) & iRow).Value = "" & .Fields(iCol)
               End If
            Next
            .MoveNext
         Loop
         
         If xlsFileName <> "" Then
            wksReport.Range("A1", Chr(UBound(arrTmp) + 65) & "1").Font.Bold = True
            wksReport.Columns(Chr(LBound(arrTmp) + 65) & ":" & Chr(UBound(arrTmp) + 65)).EntireColumn.Font.Name = "標楷體"
            wksReport.Columns(Chr(LBound(arrTmp) + 65) & ":" & Chr(UBound(arrTmp) + 65)).EntireColumn.AutoFit
            wksReport.Columns("D:D").ColumnWidth = 16.63 'Added by Morgan 2024/3/11 設定代表圖寬度與總簿同
            
            If Val(xlsReport.Version) < 12 Then
               xlsReport.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=-4143
            Else
               xlsReport.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=56
            End If
            
            xlsReport.Workbooks.Close
            xlsReport.Quit
            'Modified by Morgan 2023/3/3
            'SendMAPIMail arrRec(iZone), xlsFileName & "(請轉發相關智權人員)", "如旨", App.path & TextPath & xlsFileName, , arrCC(iZone)
            SendMAPIMail stRcvr, stSubj, "如旨", App.path & TextPath & xlsFileName
            'end 2023/3/3
         End If
      Next
      End With
   End If
   
   PUB_KillTempFile "$$*.*" '刪代表圖檔 'Added by Morgan 2023/3/1
   Set rsQuery = Nothing
   Exit Sub
    
ErrHnd:
    WLog Err.Description
'    'Modify by Amy 2021/06/22 改與上面一致
'    If Val(xlsReport.Version) < 12 Then
'       xlsReport.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
'    Else
'       xlsReport.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
'    End If
    If Val(xlsReport.Version) < 12 Then
         xlsReport.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=-4143
    Else
         xlsReport.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=56
    End If
    'end 2021/06/22
    xlsReport.Workbooks.Close
    xlsReport.Quit
    Set xlsReport = Nothing
    Set rsQuery = Nothing
End Sub

'Added by Morgan 2019/6/17
'刪除3個月前上傳之台一網站資料夾[網頁提供國內專利公報資訊]
Private Sub StrMenu24()
   Dim rsQuery As ADODB.Recordset
   Dim intQ As Integer, stSQL As String
   Dim stPS01 As String
   
On Error GoTo ErrHnd

   stSQL = "select ps01 from PatentSearch where ps10>add_months(sysdate,-5) and ps10<add_months(sysdate,-3) and ps11 is not null"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      Do While Not .EOF
         stPS01 = .Fields("ps01")
         If PUB_DeleteWWW(stPS01) = True Then
            cnnConnection.Execute "update PatentSearch set ps11=null where ps01=" & stPS01
         End If
         .MoveNext
      Loop
      End With
   End If
   
ErrHnd:
   If Err.Number <> 0 Then
      WLog "刪除3個月前上傳之台一網站資料夾(流水號:" & stPS01 & "):" & Err.Description
   End If
   
End Sub

'Added by Lydia 2019/06/21 FCP寄證書後年費不續辦
Private Sub StrMenu25()
Dim intB As Integer, strB1 As String
Dim strDate As String
Dim strMsg As String
Dim strBCP09 As String
Dim strBCP12 As String
Dim strBCP13 As String
Dim strBCP14 As String
Dim rsB As New ADODB.Recordset
'Added by Lydia 2020/03/16
Dim strPassSQL As String
Dim tmpArr As Variant

   '在代理人(FA41)、申請人(CU74)以及個案(PA70)KEY已經設定"FCP年費自動代繳：N (Y：自動代繳 / N：寄證書後年費不續辦)"的案件
                                                                                         '，於寄證書發文日起算一個月系統自動上"年費不續辦"。
   'Modified by Lydia 2019/12/04 改為系統日前第2個月發文;
                          '因為三星鑽石指示FMP案的期限不足一年則先辦繳費,所以多等一個月工作時間處理案件;承辦內部協調全部改成前第2個月
   'strDate = Left(CompDate(1, -1, strSrvDate(1)), 6)
   strDate = Left(CompDate(1, -2, strSrvDate(1)), 6)
   
   '判斷順序PA70>FA41>CU74
   'Modified by Lydia 2019/11/26 個案PA70=N寄證書後年費不續辦，改為PA156(FCP年費特殊管制)
   'Modified by Lydia 2019/11/27 +P案
   'Modified by Lydia 2019/12/11 FMP案要排除香港案013和澳門案044 ; 因為不是每年繳費(by 潘子微)
   'strB1 = "SELECT CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) AS CASENO,CP01,CP02,CP03,CP04,CP158,PA08,PA09,FA10,NP01,NP22,NP07 " & _
               "From CASEPROGRESS, PATENT, CUSTOMER, FAGENT, NEXTPROGRESS " & _
                "WHERE CP01 in ('FCP','P') AND CP10='1603' AND CP158>='" & strDate & "01' AND CP158<='" & strDate & "31' AND CP159=0 " & _
                "AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
                "AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) " & _
                "AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & _
                "AND NVL(PA156,NVL(FA41,CU74))='N' " & _
                "AND CP01=NP02(+) AND CP02=NP03(+) AND CP03=NP04(+) AND CP04=NP05(+) AND NP06 IS NULL AND NP07='605' "
   'Modified by Lydia 2019/12/24 增加判斷PA70=Y個案年費代繳
   'Modified by Lydia 2020/03/16 FMP案件不自動上年費不續辦，改發清單給程序，由各區程序逐筆產生定稿通知大陸代理人
   'strB1 = "SELECT CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) AS CASENO,CP01,CP02,CP03,CP04,CP158,PA08,PA09,FA10,NP01,NP22,NP07 " & _
               "From CASEPROGRESS, PATENT, CUSTOMER, FAGENT, NEXTPROGRESS " & _
                "WHERE CP10='1603' AND CP158>='" & strDate & "01' AND CP158<='" & strDate & "31' AND CP159=0 " & _
                "AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
                "AND (PA01='FCP' OR (PA01='P' AND PA09<>'013' AND PA09<>'044')) " & _
                "AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) " & _
                "AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & _
                "AND NVL(PA70,NVL(PA156,NVL(FA41,CU74)))='N' " & _
                "AND CP01=NP02(+) AND CP02=NP03(+) AND CP03=NP04(+) AND CP04=NP05(+) AND NP06 IS NULL AND NP07='605' "
   strB1 = "SELECT CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) AS CASENO,CP01,CP02,CP03,CP04,CP158,PA08,PA09,FA10,NP01,NP22,NP07 " & _
               "From CASEPROGRESS, PATENT, CUSTOMER, FAGENT, NEXTPROGRESS " & _
                "WHERE CP10='1603' AND CP158>='" & strDate & "01' AND CP158<='" & strDate & "31' AND CP159=0 " & _
                "AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
                "AND PA01='FCP' " & _
                "AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) " & _
                "AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & _
                "AND NVL(PA70,NVL(PA156,NVL(FA41,CU74)))='N' " & _
                "AND CP01=NP02(+) AND CP02=NP03(+) AND CP03=NP04(+) AND CP04=NP05(+) AND NP06 IS NULL AND NP07='605' "
   'Added by Lydia 2020/03/16 抓FMP案的範圍
   strPassSQL = "SELECT PATENT.*, NEXTPROGRESS.* " & _
               "From CASEPROGRESS, PATENT, CUSTOMER, FAGENT, NEXTPROGRESS " & _
                "WHERE CP10='1603' AND CP158>='" & strDate & "01' AND CP158<='" & strDate & "31' AND CP159=0 " & _
                "AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
                "AND (PA01='P' AND PA09<>'013' AND PA09<>'044') AND SUBSTR(CP12,1,1)='F' " & _
                "AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) " & _
                "AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & _
                "AND NVL(PA70,NVL(PA156,NVL(FA41,CU74)))='N' " & _
                "AND CP01=NP02(+) AND CP02=NP03(+) AND CP03=NP04(+) AND CP04=NP05(+) AND NP06 IS NULL AND NP07='605' "
                
   strB1 = strB1 & "ORDER BY FA10,CASENO "
   intB = 1
   Set rsB = ClsLawReadRstMsg(intB, strB1)
   If intB = 1 Then
       With rsB
           strBCP14 = "F4102"
           '設定資料庫變數
           strSql = "BEGIN user_data.user_num:='QPGMR'; END;"
           cnnConnection.Execute strSql
           .MoveFirst
           Do While Not .EOF
               strBCP13 = PUB_GetAKindSalesNo(.Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04"))
               strBCP12 = GetSalesArea(strBCP13)
               '下一程序年費自動上解除期限, 解除原因=02已轉由他所續辦
               strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & strSrvDate(1) & ",NP12='02' WHERE NP01='" & .Fields("NP01") & "' AND NP22='" & .Fields("NP22") & "' "
               cnnConnection.Execute strSql, intI
               '產生的B類進度不續辦907的承辦人預設為F4102
               strBCP09 = AutoNo("B", 6)
               strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp27,cp30,cp32,cp43,cp57,cp58) values (" & _
                            CNULL(.Fields("CP01")) & ", " & CNULL(.Fields("CP02")) & ", " & CNULL(.Fields("CP03")) & ", " & CNULL(.Fields("CP04")) & ", " & strSrvDate(1) & ", " & CNULL(strBCP09) & ", '907', " & _
                            CNULL(strBCP12) & ", " & CNULL(strBCP13) & ", " & CNULL(strBCP14) & ", 'N', 'N', " & strSrvDate(1) & ", " & .Fields("NP22") & ", 'N', " & CNULL(.Fields("NP01")) & ", " & strSrvDate(1) & ", '02') "
               cnnConnection.Execute strSql, intI
              
               '更新c類的代理人及彼所案號，要在新增c類之後
               Pub_UpdateFromMaxCP27 .Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04")
               
               'FCP台灣新型年費解除期限一案兩請提醒
               strMsg = Pub_ChkFCPDualCaseBYcancel(.Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04"), .Fields("PA08"), .Fields("PA09"), .Fields("NP07"), False)
               If strMsg <> "" Then '發Email通知FCP程序
                   strB1 = PUB_GetFCPHandler(.Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04"))
                   If strB1 <> "" Then
                       SendMAPIMail strB1, .Fields("CASENO") & " 寄證書後年費不續辦", strMsg
                   End If
               End If
               .MoveNext
           Loop
       End With
   End If
     
    'Added by Lydia 2020/03/16 FMP案件不自動上年費不續辦，改發清單給程序，由各區程序逐筆產生定稿通知大陸代理人
    If PUB_GetP605Email("2", strPassSQL, strB1) = True Then
        '每月批次-先放在MailCache,以mc11儲存檔案路徑
        strSql = "select mc01,mc02,mc03,mc04,mc07,mc08,mc11 from mailcache where mc01='QPGMR' and mc03=" & strSrvDate(1) & " and mc05 is null order by mc04 "
        intB = 1
        Set rsB = ClsLawReadRstMsg(intB, strSql)
        If intB = 1 Then
            rsB.MoveFirst
            Do While Not rsB.EOF
                If "" & rsB.Fields("mc02") <> "" Then
                     SendMAPIMail rsB.Fields("mc02"), "" & rsB.Fields("mc07"), vbCrLf & "" & rsB.Fields("mc08"), "" & rsB.Fields("mc11")
                     strSql = "update mailcache set  mc05=to_char(sysdate,'yyyymmdd'), mc06=to_char(sysdate,'hh24miss') where mc01='QPGMR' and mc03=" & strSrvDate(1) & " and mc04=" & rsB.Fields("mc04")
                     cnnConnection.Execute strSql, intI
                End If
                rsB.MoveNext
            Loop
        End If
    ElseIf strB1 <> "" Then
            WLog "FCP寄證書後年費不續辦:" & strB1
    End If
    'end 2020/03/16
   Set rsB = Nothing
   
ErrHnd:
   If Err.Number <> 0 Then
      WLog "FCP寄證書後年費不續辦:" & Err.Description
   End If
   
End Sub

'Add by Amy 2020/01/03 刪除電子發票 部分資料夾三個月前資料
Private Sub StrMenu26()
    Dim ii As Integer, jj As Integer, strPath As String, strFolder As String, stYear1 As String, stMonth1 As String, stDate1 As String
    Dim kk As Integer 'Add by Amy 2024/10/25
    
On Error GoTo ErrHand
    'Modify by Amy 2024/01/18 改抓系統特殊設定
    strPath = "\\" & Pub_GetSpecMan("分信主機名稱") & "\c$\551cron\Error\"
    'Modify by Amy 2024/02/01 +不是空
    If sChoice = "2" And sChoice <> MsgText(601) Then
      '測式用
      'Modify by Amy 2024/10/24 原:\\AA2004-\
      strPath = "\\" & PUB_ReadHostName & "\c$\551cron\Error\"
    End If
    stDate1 = DBDATE(DateAdd("m", -3, Format(strSrvDate(1), "####/##/##")))
    stYear1 = Mid(stDate1, 1, 4)
    stMonth1 = Mid(stDate1, 5, 2)
    
    'Modify by Amy 2024/10/25 上傳後資料改移至History,增加History資料夾的刪除
    For kk = 1 To 2
       strFolder = strPath & stYear1
       If kk = 2 Then strFolder = strPath & "History\" & stYear1
       For ii = 1 To Val(stMonth1)
         '判斷月資料夾
         If Dir(strFolder & "\" & Format(ii, "00"), vbDirectory) <> MsgText(601) Then
            For jj = 1 To 31
               '判斷日資料夾
               If Dir(strFolder & "\" & Format(ii, "00") & "\" & Format(jj, "00"), vbDirectory) <> MsgText(601) Then
                  '日資料夾內有資料刪除資料
                  If Dir(strFolder & "\" & Format(ii, "00") & "\" & Format(jj, "00") & "\*.*") <> MsgText(601) Then
                     Kill strFolder & "\" & Format(ii, "00") & "\" & Format(jj, "00") & "\*.*"
                  End If
                  '刪除日資料夾
                  Call RmDir(strFolder & "\" & Format(ii, "00") & "\" & Format(jj, "00"))
               End If
            Next jj
            '刪除月資料夾
            Call RmDir(strFolder & "\" & Format(ii, "00"))
            '12月資料夾刪完,刪除年資料夾
            If ii = 12 Then
               '刪除年資料夾
               Call RmDir(strFolder)
            End If
         End If
       Next ii
    Next kk
    'end 2024/10/25
    
ErrHand:
    If Err.Number <> 0 Then
       'Modify by Amy 2024/02/01 +1,寫錯log
        WLog "刪除電子發票 部分資料夾三個月前資料:" & Err.Description, 1
    End If
End Sub

'Added by Lydia 2020/09/01 結餘單流水號檢查: 未結算逾2個月結餘單寄給財務處總帳人員
Private Sub StrMenu27()
   Dim xlsReport27 As New Excel.Application
   Dim wksReport27 As New Worksheet
   Dim intQ As Integer, stSQL As String
   Dim rsQuery As New ADODB.Recordset
   Dim strGrp As String, nPages As Integer
   Dim nRows As Integer
   Dim arrTmp As Variant, arrTmpW As Variant
   Dim stDate As String
   Dim xlsFileName As String
   Dim strTo As String

On Error GoTo ErrHnd

   '未結算逾2個月結餘單寄給財務處總帳人員
   stDate = Left(TransDate(CompDate(1, -2, strSrvDate(1)), 1), 5)
   
   Call PUB_KillTempFile(Mid(TextPath & "*未結算結餘單.*", 2))
   xlsFileName = stDate & "未結算結餘單.xls"
    
   'Mark by Lydia 2025/08/12 改在迴圈內
   ''欄位抬頭
   'stSQL = "本所案號,結餘單號,結餘日期"
   'arrTmp = Split(stSQL, ",")
   'stSQL = "15,12,10"
   'arrTmpW = Split(stSQL, ",")
   'end 2025/08/12
   
   'Modified by Lydia 2025/08/12
   'stSQL = "SELECT A240005, A240005||A240006||A240007||A240008 本所案號,A240002 結餘單號,A240001 結餘日期 FROM ACC240 " & _
             "WHERE substr(A240001,1,5) < " & stDate & " AND NVL(A240003,0)=0 AND NVL(A240015,0)=0 ORDER BY A240005 desc,A240002"
   stSQL = "SELECT A240005, A240005||A240006||A240007||A240008 本所案號,A240002 結餘單號,A240001 結餘日期 " & _
           ",A240006,A240007,A240008, NVL(PA09,NVL(TM10,NVL(SP09,NVL(LC15,'000')))) AS NA01 " & _
           "From ACC240, PATENT, TRADEMARK, SERVICEPRACTICE, LAWCASE " & _
           "Where SUBSTR(A240001, 1, 5) < " & stDate & " And NVL(A240003, 0) = 0 And NVL(A240015, 0) = 0 " & _
           "AND A240005=PA01(+) AND A240006=PA02(+) AND A240007=PA03(+) AND A240008=PA04(+) " & _
           "AND A240005=TM01(+) AND A240006=TM02(+) AND A240007=TM03(+) AND A240008=TM04(+) " & _
           "AND A240005=SP01(+) AND A240006=SP02(+) AND A240007=SP03(+) AND A240008=SP04(+) " & _
           "AND A240005=LC01(+) AND A240006=LC02(+) AND A240007=LC03(+) AND A240008=LC04(+) " & _
           "ORDER BY A240005 DESC,A240002 "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
       rsQuery.MoveFirst
       Do While Not rsQuery.EOF
            If strGrp <> "" & rsQuery.Fields(0) Then
                 nPages = nPages + 1
                 If strGrp = "" Then
                     xlsReport27.SheetsInNewWorkbook = 1 '預設工作表數目
                     xlsReport27.Workbooks.add
                     xlsReport27.Application.WindowState = xlMinimized
                 Else
                     xlsReport27.Worksheets.add  '插入sheet
                 End If
                 xlsReport27.Worksheets("工作表" & nPages).Select
                 xlsReport27.Worksheets("工作表" & nPages).Name = "" & rsQuery.Fields(0)
                 Set wksReport27 = xlsReport27.Worksheets("" & rsQuery.Fields(0))
                 'Added by Lydia 2025/08/12 欄位抬頭
                 arrTmp = Empty: arrTmpW = Empty
                 If InStr(",CFT,CFC,S,", "," & rsQuery.Fields(0) & ",") > 0 Then
                     stSQL = "本所案號,結餘單號,結餘日期,承辦人"
                     arrTmp = Split(stSQL, ",")
                     stSQL = "15,12,10,10"
                     arrTmpW = Split(stSQL, ",")
                 Else
                     stSQL = "本所案號,結餘單號,結餘日期"
                     arrTmp = Split(stSQL, ",")
                     stSQL = "15,12,10"
                     arrTmpW = Split(stSQL, ",")
                 End If
                 'end 2025/08/12
                 '設定欄位名稱及欄寬
                 nRows = 1
                 For intQ = 1 To UBound(arrTmp) + 1
                     wksReport27.Range(Chr(intQ + 64) & nRows).Value = arrTmp(intQ - 1)
                     wksReport27.Range(Chr(intQ + 64) & ":" & Chr(intQ + 64)).ColumnWidth = Val(arrTmpW(intQ - 1))
                     wksReport27.Range(Chr(intQ + 64) & nRows).HorizontalAlignment = xlCenter
                 Next
                 nRows = nRows + 1
                 strGrp = "" & rsQuery.Fields(0)
            End If
            For intQ = 1 To UBound(arrTmp) + 1
                With wksReport27.Range(Chr(intQ + 64) & nRows)
                    'Added by Lydia 2025/08/12
                    If intQ = 4 Then
                        Call GetNA69("", rsQuery.Fields("NA01"), "", strExc(1), rsQuery.Fields("A240005"), rsQuery.Fields("A240006"), rsQuery.Fields("A240007"), rsQuery.Fields("A240008"))
                        .Value = GetStaffName(strExc(1))
                    Else
                    'end 2025/08/12
                        .Value = "" & rsQuery.Fields(intQ)
                    End If
                    .NumberFormatLocal = "@"
                    'Modified by Lydia 2025/08/12
                    'If intQ < 3 Then
                    If intQ <> 3 Then
                       wksReport27.Range(Chr(intQ + 64) & nRows).HorizontalAlignment = xlLeft
                    Else
                       wksReport27.Range(Chr(intQ + 64) & nRows).HorizontalAlignment = xlRight
                    End If
                End With
            Next intQ
            nRows = nRows + 1
            rsQuery.MoveNext
       Loop
       If Val(xlsReport27.Version) < 12 Then
          xlsReport27.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=-4143
       Else
          xlsReport27.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=56
       End If
       xlsReport27.Workbooks.Close
       xlsReport27.Quit
       
       strTo = Pub_GetSpecMan("財務處總帳人員")
       If strTo <> "" Then
          SendMAPIMail strTo, "結餘單流水號檢查", "同主旨", App.path & TextPath & xlsFileName
       End If
   'Added by Lydia 2021/11/24 若當月沒資料也要發mail通知：本月無未結算逾2個月之結餘單。
   Else
       strTo = Pub_GetSpecMan("財務處總帳人員")
       If strTo <> "" Then
          SendMAPIMail strTo, "本月無未結算逾2個月之結餘單", "同主旨"
       End If
   'end 2021/11/24
   End If

   Set rsQuery = Nothing
   Set xlsReport27 = Nothing
   Exit Sub
    
ErrHnd:
    
    WLog "結餘單流水號檢查:" & Err.Description
    If Val(xlsReport27.Version) < 12 Then
       xlsReport27.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=-4143
    Else
       xlsReport27.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=56
    End If
    xlsReport27.Workbooks.Close
    xlsReport27.Quit
    Set xlsReport27 = Nothing
    Set rsQuery = Nothing
End Sub

'Modify by Amy 2021/07/28 寫入暫存檔-FCT案原全部MAIL給68005,改逐案以本所案號呼叫PUB_GetFCTSalesNo抓出負責的人,再抓該員的NVL(NVL(ST54,ST53),ST52),再將清單MAIL給抓出的主管
'申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細
Sub StrMenu3()
    Dim RsQ As New ADODB.Recordset, rsA As New ADODB.Recordset
    Dim ff As Integer, i As Integer, intQ As Integer
    Dim TempFileNameAll As String, TempFileName As String, TempDelFileName As String
    Dim ArrMail As Variant
    Dim strOldStN As String, TpNext As String, strOldTo As String, strTo As String, strSubject As String, strContentFix As String, strContent As String
    Dim strTmp(8) As String

On Error GoTo DebugErr
    '刪除暫存檔(與 AutoBatchDay-StrMenu18共用暫存檔)
    strSql = "Delete RAB3 Where State='1' "
    cnnConnection.Execute strSql
    
    '申請
    strSql = "Select '1',cp01,cp02,cp03,cp04,cp09,cp10,cp05,cp07,cp13,cp14,cp27,Rpad('申請',10,' '),cp65 " & _
                "From (Select tm01,tm02,tm03,tm04,tm05,tm06,tm07,tm23,tm15,tm12 " & _
                        " From Trademark Where tm11>=20000000 And tm01 in ('FCT') And tm10='000' And tm29 is null And tm16 is null And tm28='1'  And tm11 <=To_Number(To_Char(Add_Months(sysdate,-20),'yyyymmdd')) " & _
                        " And Not Exists(Select * From CaseProgress C1 Where C1.cp01=tm01 And C1.cp02=tm02  And C1.cp03=tm03 And C1.cp04=tm04 And ((c1.cp10 not in ('703','704') And c1.cp09<'C' And c1.cp27>To_Number(To_Char(Add_Months(sysdate,-6),'yyyymmdd'))) " & _
                        " Or (c1.cp09>'C' And c1.cp05>To_Number(To_Char(Add_Months(sysdate,-6),'yyyymmdd'))) Or (C1.cp10<>'1101' And C1.cp09>'C') Or (c1.cp10='310' And c1.cp27>To_Number(To_Char(Add_Months(sysdate,-24),'yyyymmdd'))) " & _
                        " Or (c1.cp10='1401' And c1.cp05>To_Number(To_Char(Add_Months(sysdate,-24),'yyyymmdd'))) Or (C1.cp10='306' And C1.cp27 is not null And Exists (Select * From CaseProgress C3 Where c1.cp43=C3.cp09 And C3.cp10='101') ))  And C1.cp05>=20000000) " & _
                "),CaseProgress C2 Where tm01=C2.cp01(+) And tm02=C2.cp02(+) And tm03=C2.cp03(+) And tm04=C2.cp04(+) And C2.cp10='101' " & _
                " And Not Exists (Select C4.cp09 From CaseProgress C4,CaseProgress C5 Where C4.cp01=tm01 And C4.cp02=tm02 And C4.cp03=tm03 And C4.cp04=tm04 And C4.cp10='101' And C4.cp09=C5.cp43(+) And C5.cp10='1724') "
    '核駁前先行
    strSql = strSql & " Union All " & _
                "Select '1',cp01,cp02,cp03,cp04,cp09,cp10,cp05,cp07,cp13,cp14,cp27,Rpad('核駁前先行',10,' '),cp65 " & _
                "From Trademark, CaseProgress c2 " & _
                " Where c2.cp05>=20000000 And tm01 in ('FCT') And tm10='000' And tm29 is null And tm16 is null And tm01=c2.cp01(+)  And tm02=c2.cp02(+) And tm03=c2.cp03(+) And tm04=c2.cp04(+) And c2.cp10='1202' " & _
                " And Not Exists(Select * From CaseProgress C1 Where C1.cp01=tm01 And C1.cp02=tm02  And C1.cp03=tm03 And C1.cp04=tm04 And ((c1.cp10 not in ('703','704') And c1.cp09<'C' And c1.cp27>To_Number(To_Char(Add_Months(sysdate,-6),'yyyymmdd'))) " & _
                " Or (c1.cp09>'C' And c1.cp05>To_Number(To_Char(Add_Months(sysdate,-6),'yyyymmdd'))) Or (c1.cp10='310' And c1.cp27>To_Number(To_Char(Add_Months(sysdate,-24),'yyyymmdd'))) " & _
                " Or (c1.cp10='1401' And c1.cp05>To_Number(To_Char(Add_Months(sysdate,-24),'yyyymmdd'))) Or (c1.cp10='1002' And c1.cp64='申請駁回') Or (C1.cp10='306' And C1.cp27 is not null And Not Exists (Select * From CaseProgress C3 Where c1.cp43=C3.cp09(+) And C3.cp10='101') )) " & _
                " And C1.cp05>=20000000) And c2.cp05 <=To_Number(To_Char(Add_Months(sysdate,-12),'yyyymmdd') ) " & _
                " And Not Exists (Select C4.cp09 From CaseProgress C4,CaseProgress C5 Where C4.cp01=tm01 And C4.cp02=tm02 And C4.cp03=tm03 And C4.cp04=tm04 And C4.cp10='101' And C4.cp09=C5.cp43(+) And C5.cp10='1724') "
    
    strSql = "Insert Into RAB3 (State,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012,R013) " & strSql
    cnnConnection.Execute strSql, intI
    If intI = 0 Then Exit Sub
    
    TpNext = "資料已新增至RAB3..."
    
    '若該收文號的下一程序檔有掛305.催審期限而NP08.本所期限尚未到期尚不列出(FCT-028812),
    '或該催審期限的np06 is not null也不列出(FCT-030066).
    '若下一程序檔無催審期限則仍然要列出.
    strSql = "Delete From RAB3 Where State='1' And R001='FCT' " & _
                "And R001||R002||R003||R004 In (Select np02||np03||np04||np05  From NextProgress,CaseProgress,Staff S1,Staff S2 " & _
                                                                    "Where np01=cp09 And cp10='101' And np07='305' And (np08>" & strSrvDate(1) & " Or np06 is not null) " & _
                                                                    "And R001='FCT' And R001=np02 And R002=np03 And R003=np04 And R004=np05 " & _
                                                                    "And SubStr(Nvl(S1.st03,S2.st03),1,1)='F' And R010=S1.st01 And R013=S2.st01 ) "
    cnnConnection.Execute strSql
    TpNext = "條件不符已刪除..."
  
    TpNext = "寄件人員更新開始..."
    'Mark by Amy 2022/11/01 已沒抓T(看 strMenu3_Old)
'    'T案統一寄給林純真經理
'    'Modify by Amy 2021/11/10 原:69008改為84027(林嘉雯)及79041(林桂英)
'    strSql = "Update RAB3 Set MailTo='84027;79041' Where State='1' And R001='T' "
'    cnnConnection.Execute strSql
    'end 2022/11/01
    
    '更新寄件人員:以本所案號 PUB_GetFCTSalesNo 抓出負責的人，再抓該員的NVL(NVL(ST54,ST53),ST52)
    strSql = "Select Distinct R001 as cp01,R002 as cp02,R003 as cp03,R004 as cp04 From RAB3 Where State='1' And R001='FCT' "
    intI = 1
    Set RsQ = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            strTmp(0) = PUB_GetFCTSalesNo(RsQ.Fields("cp01"), RsQ.Fields("cp02"), RsQ.Fields("cp03"), RsQ.Fields("cp04"))
            'Add by Amy 2022/11/01 FCT案不抓承辦人改抓目前智權人員PUB_GetFCTSalesNo-陳金蓮
            If strTmp(0) <> MsgText(601) And "" & RsQ.Fields("cp01") = "FCT" Then
                strSql = "Update RAB3 Set R009='" & strTmp(0) & "' Where State='1' " & _
                            "And R001='" & RsQ.Fields("cp01") & "' And R002='" & RsQ.Fields("cp02") & "' And  R003='" & RsQ.Fields("cp03") & "' And R004='" & RsQ.Fields("cp04") & "' "
                cnnConnection.Execute strSql
            End If
            'end 2022/11/01
            'Mark by Amy 2024/09/13 目前只有FCT案,改後面更新MailTo對象
'            If strTmp(0) <> MsgText(601) Then
'                strTmp(1) = "Select Nvl(Nvl(ST54,ST53),ST52) as MailTo From Staff Where st01='" & strTmp(0) & "' "
'                intQ = 1
'                Set rsA = ClsLawReadRstMsg(intQ, strTmp(1))
'                If intQ = 1 Then
'                    strTmp(2) = "" & rsA.Fields("MailTo")
'                     If strTmp(2) <> MsgText(601) Then
'                        strTmp(3) = "Update RAB3 Set MailTo='" & strTmp(2) & "' Where State='1' " & _
'                                            "And R001='" & RsQ.Fields("cp01") & "' And R002='" & RsQ.Fields("cp02") & "' " & _
'                                            "And R003='" & RsQ.Fields("cp03") & "' And R004='" & RsQ.Fields("cp04") & "' "
'                        cnnConnection.Execute strTmp(3)
'                     End If
'                End If
'            End If
            RsQ.MoveNext
        Loop
    End If
    'Add by Amy 2024/09/13 秀玲:抓FCT案目前智權人員(R009)之st93的a0924(部門主管-發MAIL用)-洪琬姿
    strSql = "Update RAB3 Set MailTo=(Select A0924 From Acc090New,Staff Where st01=R009 And st93=A0921)" & _
                   "Where State='1' And R009 is not null "
    cnnConnection.Execute strSql
    'end 2024/09/13
    'Add by Amy 2022/06/17 無收件人寄給Pub_GetSpecMan("程式管理人員")
    strSql = "Update RAB3 Set MailTo='" & Pub_GetSpecMan("程式管理人員") & "' Where State='1' And MailTo Is Null "
    cnnConnection.Execute strSql
    'end 2022/06/17
    TpNext = "寄件人員更新完成..."
    
    
    strSubject = "申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細---------" & Format(Now, "YYYY/MM/DD")
    strContentFix = "Dear Sirs," & vbCrLf & "       民國 " & Trim(Year(Now) - 1911) & "  年 " & Trim(Month(Now)) & " 月 的    申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細   "

    TpNext = "開始讀取資料..."
    strTmp(0) = "Select Distinct MailTo From RAB3 Where State='1' Order by MailTo "
    intI = 1
    Set rsA = ClsLawReadRstMsg(intI, strTmp(0))
    If intI = 1 Then
        rsA.MoveFirst
        strOldStN = ""
        Do While rsA.EOF = False
            '寄信
            If strOldTo <> MsgText(601) Then
                If ff > 0 Then Close #ff
                strContent = strContentFix & IIf(TempFileNameAll = "", "資料庫找不到資料", "資料如附件") & "！" & _
                                    vbCrLf & vbCrLf & vbCrLf & vbCrLf & "請橫印！" & vbCrLf & vbCrLf & vbCrLf & _
                                    "                                                        電腦中心"
                'SendMAPIMail "A2004", strSubject, strContent, TempFileNameAll
                SendMAPIMail strOldTo, strSubject, strContent, TempFileNameAll
                TempFileNameAll = "": strOldStN = ""
            End If
            
            '*** 抓取資料,依每個人員(stName)產生一個檔案 ***
            strSql = "Select SubStr(nvl(S1.st03,s2.st03),1,1) as Dept,Decode(SubStr(Nvl(S1.st03,S2.st03),1,1),'F',Rpad(Nvl(S3.st02,' '),6,' '),Rpad(Nvl(S1.st02,' '),6,' ')) as stName" & _
                                ",Rpad(Nvl(Nvl(tm15,tm12),' '),20,' ') as tm15,Rpad(R001||'-'||R002||'-'||R003||'-'||R004,15,' ') as caseNo,Rpad(tm05||tm06||tm07,20,' ') as caseN,R012" & _
                                ",Rpad(SqlDatet(R011),10,' ') as cp27,Rpad(tm23,10,' '),Rpad(cu04,12,' '),Nvl(R011,0),R001,R002,R003,R004 " & _
                        "From RAB3,TradeMark,Customer,Staff S1,Staff S2,Staff S3 " & _
                        "Where MailTo='" & rsA.Fields("MailTo") & "' And R001=tm01(+) And R002=tm02(+) And R003=tm03(+) And R004=tm04(+) " & _
                         "And R010=S1.st01(+) And R013=S2.st01(+) And R009=S3.st01(+) And SubStr(tm23,1,8)=cu01(+) And SubStr(tm23,9,1)=cu02(+) " & _
                         "Order by 1,2,6,10,4 "
            intQ = 1
            Set RsQ = ClsLawReadRstMsg(intQ, strSql)
            If intQ = 1 Then
                RsQ.MoveFirst
                Do While RsQ.EOF = False
                    For i = LBound(strTmp) To UBound(strTmp)
                        strTmp(i) = "" & RsQ.Fields(i)
                    Next i
                    If strOldStN <> "" & RsQ.Fields("stName") Then
                        If ff > 0 Then Close #ff
                        ff = FreeFile
                        TempFileName = App.path & TextPath & IIf(Trim("" & RsQ.Fields("stName")) = "", "沒承辦人" & IIf("" & RsQ.Fields("Dept") = "F", "外商", "內商"), Trim("" & RsQ.Fields("stName"))) & ".txt"
                        TempFileName = PUB_UniToBIG5(TempFileName, "F") 'Added by Lydia 2022/03/28 員工名稱有Unicode
                        TempFileNameAll = TempFileNameAll & TempFileName & ";"
                        TempDelFileName = TempDelFileName & TempFileName & ";" 'Add by Amy 2021/07/30
                        Open TempFileName For Output As ff
                        Print #ff, Space(20) & "申請案發文 20 個月無來函或收到核駁先行通知後 6 個月未接獲審定之案件明細"
                        Print #ff, ""
                        Print #ff, "承辦人  申請案號             本所案號　　　　案件名稱　　　　　　 案件性質　 發文日　　 申請人編號 申請人名稱　"
                        Print #ff, "======= ==================== =============== ==================== ========== ========== ========== ==========="
                        TpNext = "開啟" & RsQ.Fields("stName") & " 檔..."
                        strOldStN = "" & RsQ.Fields("stName")
                    End If
                    TpNext = "正在寫" & RsQ.Fields("stName") & " 檔..."
                    Print #ff, strTmp(1) & "  " & strTmp(2) & " " & strTmp(3) & " " & strTmp(4) & " " & strTmp(5) & " " & strTmp(6) & " " & strTmp(7) & " " & strTmp(8)
                    
                    RsQ.MoveNext
                Loop
            End If
            '*** End 抓取資料,依每個人員(stName)產生一個檔案 ***
            strOldTo = "" & rsA.Fields("MailTo")
            rsA.MoveNext
        Loop
    End If
    '寄最後一封信
    If ff > 0 Then Close #ff
    strContent = strContentFix & IIf(TempFileNameAll = "", "資料庫找不到資料", "資料如附件") & "！" & _
                        vbCrLf & vbCrLf & vbCrLf & vbCrLf & "請橫印！" & vbCrLf & vbCrLf & vbCrLf & _
                        "                                                        電腦中心"
                
    'SendMAPIMail "A2004", strSubject, strContent, TempFileNameAll
    SendMAPIMail strOldTo, strSubject, strContent, TempFileNameAll
    
    'Add by Amy 2021/07/30 刪除檔案
    If TempDelFileName <> "" Then
        'Modify by Amy 2021/08/02 +Mid,否則多run 一次會出現找不到檔案
        ArrMail = Split(Mid(TempDelFileName, 1, Len(TempDelFileName) - 1), ";")
        For i = LBound(ArrMail) To UBound(ArrMail)
            Kill ArrMail(i)
        Next i
    End If
    
    Set rsA = Nothing
    Set RsQ = Nothing
    Exit Sub
    
DebugErr:
    'Modify by Amy 2021/08/02 +1,寫錯log
    WLog TpNext & " " & Err.Description, 1
End Sub

'Add by Sindy 2021/12/15
'每月批次增加外商FCT案註冊已滿三年案件管制表
Private Sub StrMenu28()
   Dim rsB As New ADODB.Recordset
   Dim stDate As String, stDate2 As String
   Dim strTemp(11) As String
   Dim ff1 As Integer
   Dim TempFileName As String
   Dim strEmp As String
   Dim tmpTitle As String '報表抬頭設定
   Dim strContent As String
   Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String
   Dim jj As Integer
   
On Error GoTo ErrHandle
   
   tmpTitle = "外商FCT案註冊已滿三年案件管制表"
   '清除舊檔及資料
   Call PUB_KillTempFile(Mid(TextPath & "*" & tmpTitle & "*.*", 2))
   cnnConnection.Execute "DELETE FROM R020302 WHERE ID='" & strUserNum & "'"
   
   '列印當月+３個月(再減三年)
   stDate = (Mid(CompDate(1, 3, strSrvDate(1)), 1, 6) & "01") - 30000
   '月底
   stDate2 = GetLastDay(stDate)
   
'本所案號 R055001
'審定號 R055007
'案件名稱 R055008 140
'商品類別 R055006
'專用期止日 R055003
'申請人 R055009  80
'代理人 R055010  80
'R055013: PUB_GetFCTSalesNo
'ID: strUserNum

   '商標註冊屆滿3年的3個月前：抓出未閉卷未銷卷且TM21＋3年的年月＝系統月＋3個月的年月之FCT案件
   '，以清單方式通知PUB_GetFCTSalesNo(正本)，副本給其主管 (即主任及副理) (副本主管功能已做在PUB_SendMail)
   '例：2021/10跑2022/1/1~2022/1/31，即TM21為2019/1/1~2019/1/31的資料
   '只要英文組案件，即FC代理人國籍非日本的案件
   strSql = "insert into R020302(R055001,R055007,R055008,R055006,R055003,R055009,R055010,ID)" & _
            " SELECT TM01||'-'||TM02||'-'||TM03||'-'||TM04,TM15,TM05,TM09,sqldatet(TM22),substr(TM23||NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),1,80),substr(TM44||NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),1,80),'" & strUserNum & "'" & _
            " FROM TradeMark,Nation,Staff s1,Customer,fagent" & _
            " WHERE CU13=s1.ST01(+)" & _
            " AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+)" & _
            " AND SUBSTR(TM44,1,8)=FA01(+) AND decode(SUBSTR(TM44,9,1),'','0',SUBSTR(TM44,9,1))=FA02(+)" & _
            " AND substr(FA10,1,3)<>'011'" & _
            " AND TM10=NA01(+)" & _
            " And TM21 is not null  and TM01 in ('FCT')" & _
            " And to_char(to_date(add_months(to_date(tm21,'YYYYMMDD'),3))+1,'YYYYMMDD')>=" & stDate & _
            " And to_char(to_date(add_months(to_date(tm21,'YYYYMMDD'),3))+1,'YYYYMMDD')<=" & stDate2 & _
            " And TM29||TM57 is null And Nvl(tm14,0)<>0 And Nvl(tm21,0)<>0"
   cnnConnection.Execute strSql, intI
   
   '更新目前智權人員
   strSql = "SELECT * FROM R020302 WHERE ID='" & strUserNum & "'"
   Set rsB = New ADODB.Recordset
   With rsB
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
            Str01 = SystemNumber(.Fields("R055001"), 1)
            Str02 = SystemNumber(.Fields("R055001"), 2)
            Str03 = SystemNumber(.Fields("R055001"), 3)
            Str04 = SystemNumber(.Fields("R055001"), 4)
            strEmp = PUB_GetFCTSalesNo(Str01, Str02, Str03, Str04)
            strSql = "UPDATE R020302 SET R055013='" & strEmp & "' WHERE ID='" & strUserNum & "' AND R055001='" & .Fields("R055001") & "'"
            cnnConnection.Execute strSql
            .MoveNext
         Loop
      End If
   End With
   
   '+列印說明
    strContent = "資料如附件！" & vbCrLf & vbCrLf & "列印注意事項：" & vbCrLf & _
         vbCrLf & String(4, "　") & "1.利用筆記本開啟附件" & _
         vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
         vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
         vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
         vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
         vbCrLf & String(4, "　") & "6.選擇<橫印>"
   strEmp = ""
   '產出清單
   strSql = "SELECT * FROM R020302 WHERE ID='" & strUserNum & "' ORDER BY R055013 asc,R055001 asc"
   Set rsB = New ADODB.Recordset
   With rsB
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
            If strEmp = "" Or strEmp <> Trim("" & .Fields("R055013")) Then
               If strEmp <> "" Then
                  If Dir(TempFileName) <> "" Then
                     Close ff1
                     SendMAPIMail strEmp, tmpTitle & "_" & ChangeWStringToTString(stDate) & "-" & ChangeWStringToTString(stDate2), strContent, TempFileName
                  End If
               End If
               
               strEmp = Trim("" & .Fields("R055013"))
               
               TempFileName = strEmp & tmpTitle & "_" & ChangeWStringToTString(stDate) & "-" & ChangeWStringToTString(stDate2)
               TempFileName = App.path & TextPath & TempFileName & ".txt"
               TempFileName = PUB_UniToBIG5(TempFileName, "F") 'Added by Lydia 2022/03/28 員工名稱有Unicode
               If Dir(TempFileName) <> "" Then Kill TempFileName
               
               ff1 = FreeFile
               Open TempFileName For Output As ff1
               Print #ff1, Space(47) & tmpTitle
               Print #ff1, "承辦人：" & GetStaffName(strEmp, True) & Space(30) & "註冊屆滿三年日期：" & ChangeWStringToTDateString(stDate) & "-" & ChangeWStringToTDateString(stDate2) & Space(30) & "列印日期：" & ChangeWStringToTDateString(strSrvDate(1))
               Print #ff1, ""
               Print #ff1, "------------------------------------------------------------------------------------------------------------------------------------"
               Print #ff1, "本所案號        審定號          案件名稱             商品類別           申請人               代理人               專用期止日"
               Print #ff1, "------------------------------------------------------------------------------------------------------------------------------------"
            
               '以防沒有抓到智權人員
               If strEmp = "" Then
                  strEmp = "97038"
               End If
            End If
            
            For jj = 1 To 7
               strTemp(jj) = ""
            Next jj
            strTemp(1) = Trim("" & .Fields("R055001")) '本所案號
            strTemp(2) = Trim("" & .Fields("R055007")) '審定號
            strTemp(3) = Trim("" & .Fields("R055008")) '案件名稱
            strTemp(4) = Trim("" & .Fields("R055006")) '商品類別
            strTemp(5) = Trim("" & .Fields("R055009")) '申請人
            strTemp(6) = Trim("" & .Fields("R055010")) '代理人
            strTemp(7) = Trim("" & .Fields("R055003")) '專用期止日
            
            strTemp(1) = convForm(CheckStr(strTemp(1)), 15)
            strTemp(2) = convForm(CheckStr(strTemp(2)), 15)
            strTemp(3) = convForm(CheckStr(strTemp(3)), 20)
            strTemp(4) = convForm(CheckStr(strTemp(4)), 18)
            strTemp(5) = convForm(CheckStr(strTemp(5)), 20)
            strTemp(6) = convForm(CheckStr(strTemp(6)), 20)
            strTemp(7) = convForm(CheckStr(strTemp(7)), 10)
            
            Print #ff1, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & " " & strTemp(6) & " " & strTemp(7)
            .MoveNext
         Loop
         If strEmp <> "" Then
            If Dir(TempFileName) <> "" Then
               Close ff1
               SendMAPIMail strEmp, tmpTitle & "_" & ChangeWStringToTString(stDate) & "-" & ChangeWStringToTString(stDate2), strContent, TempFileName
            End If
         End If
      End If
   End With
   
   cnnConnection.Execute "DELETE FROM R020302 WHERE ID='" & strUserNum & "' "
   Set rsB = Nothing
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      WLog tmpTitle & ":" & Err.Description
   End If
   Set rsB = Nothing
End Sub

'Add by Amy 2023/08/16 檢查智權部新客戶與舊客戶統編相同清單,寄給「全所智權部主管」
Private Sub StrMenu29()
   Dim RsQ As New ADODB.Recordset, intQ As Integer, strQ As String, ii As Integer
   Dim strFieldName As String, strTo As String, strText_Fix As String, strText As String, strTp(5) As String
On Err GoTo ErrHnd
   
   strFieldName = Mid(strSrvDate(1), 1, 6) & "月智權部新客戶身份證號統一編號與舊客戶相同"
   strFieldName = App.path & TextPath & strFieldName & ".txt"
   If Dir(strFieldName) <> "" Then Kill strFieldName
   
   strText_Fix = "備註：改字型Fixedsys標準11號字以直式上下左右各10MM列印" & vbCrLf & vbCrLf
   strText_Fix = strText_Fix & "統一編號   客戶編號   申請人名稱           智權人員   開發日期  參考備註   " & vbCrLf
   strText_Fix = strText_Fix & "========== ========== ==================== ========== ========  ====================================" & vbCrLf
   
   strQ = "Select cu11,cu01||cu02 as CusNo,cu04,st02,sqldatet(cu14) as Cu14,cu79 " & _
                  "From Customer,Staff Where cu11 in " & _
                    "(Select cu11 From Customer Where cu11 in " & _
                           "(Select Distinct cu11 From Customer " & _
                           "Where Substr(cu14,1,6)=to_char(add_months(sysdate,-1),'yyyymm') And (cu80<>'不再使用' or cu80 is null) " & _
                           "And cu11 is not null And cu11<>'00000000' And cu12 like 'S%') " & _
                    "Group by cu11 having count(*)>1) " & _
                  "And cu13=st01(+) Order by cu11,cu01"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      RsQ.MoveFirst
      Do While Not RsQ.EOF
         For ii = LBound(strTp) To UBound(strTp)
            strTp(ii) = "" & RsQ.Fields(ii)
         Next ii
         strTp(0) = PUB_StrToStr(strTp(0), 10, True) '統一編號
         strTp(1) = PUB_StrToStr(strTp(1), 10, True) '客戶編號
         strTp(2) = PUB_StrToStr(strTp(2), 20, True) '申請人名稱
         strTp(3) = PUB_StrToStr(strTp(3), 10, True) '智權人員
         strTp(4) = PUB_StrToStr(strTp(4), 10, True) '智權人員
         strTp(5) = PUB_StrToStr(strTp(5), 30, True) '參考備註
         strText = strText & strTp(0) & " " & strTp(1) & " " & strTp(2) & " " & strTp(3) & " " & strTp(4) & strTp(5) & " " & vbCrLf
         
         RsQ.MoveNext
      Loop
      If strText <> MsgText(6001) Then
         Call PUB_SaveTextAsUTF8(strFieldName, strText_Fix & strText)
         If Dir(strFieldName) <> MsgText(601) Then
            strTo = Pub_GetSpecMan("全所智權部主管")
            SendMAPIMail strTo, Mid(strSrvDate(1), 1, 6) & "月智權部新客戶身份證號/統一編號與舊客戶相同清單，請檢視！", "Dear Sirs," & vbCrLf & Mid(strSrvDate(1), 1, 6) & "月智權部新客戶身份證號/統一編號與舊客戶相同清單，資料如附件，請橫印！" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心", strFieldName
         End If
      End If
   End If
                  
   Set RsQ = Nothing
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
      WLog "智權部新客戶身份證號統一編號與舊客戶相同清單:" & Err.Description
   End If
   Set RsQ = Nothing
End Sub

'Added by Lydia 2023/09/14 自動轉入待活化客戶：每年之1月1日及7月1日，將已超過12年未收文但尚未設為活化客戶的客戶資訊，整合進入活化區域。
Private Sub StrMenu30()
Dim intB As Integer, strB1 As String, strB2 As String
Dim strDate0 As String, strCond As String
Dim strTmp(1 To 5) As String
   
   '-------Memo by Lydia 2024/07/04
   '2024/07/01 發現20231201匯入資料有語法錯誤，補上143筆(OCU02=20240630)
   '2024/07/01~2024/07/03 檢查資料，增加對簽訂合約(Contract)的客戶以及移轉/變更申請人(CP56)，新增記錄1547筆(OCU02=20240704)
   '-------
   strDate0 = Val(CompDate(0, -12, strSrvDate(1))) - 1 '12年內=>重整待活化客戶：預設以101/1/1後未收文，且符合以下條件者，即列入待活化客戶。另請設定每半年調整一次，下次調整時間為113.7.1。
   'Modified by Lydia 2025/01/14 改為系統特殊設定
   'strCond = "解散、廢止、撤銷、死亡"
   strCond = Pub_GetSpecMan("待活化客戶-無效狀態設定")
   
On Error GoTo ErrHnd
   strSql = "TRUNCATE TABLE RAB30_SALESNO "
   cnnConnection.Execute strSql
   '排除已存在的活化客戶
   'Modified by Lydia 2023/10/19 排除:1.MCTF開頭之智權人員客戶；2.96029~96032；3.郭雅娟經理79075客戶；
   'Modified by Lydia 2023/11/30 排除:LXX(即L01、L02)部門的非台灣客戶請剔除。因為這些國外客戶都是FCL案件的客戶，跟外專外商一樣所以要剔除。
                                '排除CU197,CU198固定不列入待活化客戶的設定
   'Modified by Lydia 2025/03/07 + NVL(OCU03,0)=0
   strSql = "INSERT INTO RAB30_SALESNO (CU01) " & _
            "SELECT CU01 FROM CUSTOMER WHERE (CU01 BETWEEN 'X0000000' AND 'XZZZZZZZ') AND CU02='0' " & _
            "AND CU01 NOT IN (SELECT OCU01 FROM OLDCUSTOMER WHERE NVL(OCU03,0)=0 ) " & _
            "and substr(cu12,1,1) <> 'F' and (instr(" & CNULL(strCond) & ",cu80)=0 or cu80 is null) AND CU13 NOT LIKE 'MCTF%' AND NOT (CU13>='96029' AND CU13<='96032') AND CU13<>'79075' " & _
            "AND NOT (CU12 LIKE 'L%' AND CU10>'010') AND CU197 IS NULL "
   cnnConnection.Execute strSql
   
   strSql = "TRUNCATE TABLE RAB30_SALESNOC "
   cnnConnection.Execute strSql
   strSql = "Insert Into RAB30_SALESNOC " & _
            "(Select Distinct B.Cu01,C.Cu01 From RAB30_SALESNO B,Customer C " & _
            "WHERE Substr(B.Cu01,1,6)=Substr(C.Cu01,1,6) And C.Cu01 Is Not Null And B.Cu01<>C.Cu01)"
   cnnConnection.Execute strSql

   strSql = "TRUNCATE TABLE RAB30_SALESNOA "
   cnnConnection.Execute strSql
   strB2 = GetAllSysKind(, "ALL")
   '專利
   strTmp(1) = Replace(SQLGrpStr(strB2, 1), ",' '", "")
   'Modified by Lydia 2023/11/30 (Email:112/11/13) "CP158=0 AND CP159=0"拿掉" AND CP159=0"
            '若於12年內有收文 , 但最後仍以銷案處理掉的客戶, 因狀況太多元, 為避免客戶關係有所影響, 仍視為有收文之資訊, 不列入待活化客戶
            '待活化的收文條件原剔除了已取消收文的案件及法律所案源被放棄的條件，請再修改只要有收文或是有填寫案源接洽單都視為有往來。
   'Modified by Lydia 2024/11/20 整理語法，拿掉所有AND (CP27>0 OR (CP158=0))
   strExc(1) = "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(PA26,1,8) PA26 FROM CASEPROGRESS,PATENT,RAB30_SALESNO " & _
                 "WHERE SUBSTR(PA26,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 " & _
                 "AND CP01 IN (" & strTmp(1) & ") AND CP09<'B'    " & _
               "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(PA27,1,8) FROM CASEPROGRESS,PATENT,RAB30_SALESNO " & _
                 "WHERE SUBSTR(PA27,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 " & _
                 "AND CP01 IN (" & strTmp(1) & ") AND CP09<'B'    " & _
               "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(PA28,1,8) FROM CASEPROGRESS,PATENT,RAB30_SALESNO " & _
                 "WHERE SUBSTR(PA28,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 " & _
                 "AND CP01 IN (" & strTmp(1) & ") AND CP09<'B'    " & _
               "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(PA29,1,8) FROM CASEPROGRESS,PATENT,RAB30_SALESNO " & _
                 "WHERE SUBSTR(PA29,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 " & _
                 "AND CP01 IN (" & strTmp(1) & ") AND CP09<'B'    " & _
               "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(PA30,1,8) FROM CASEPROGRESS,PATENT,RAB30_SALESNO " & _
                 "WHERE SUBSTR(PA30,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 " & _
                 "AND CP01 IN (" & strTmp(1) & ") AND CP09<'B'    "
   '商標
   strTmp(2) = Replace(SQLGrpStr(strB2, 2), ",' '", "")
   strExc(2) = "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(TM23,1,8) FROM CASEPROGRESS,TRADEMARK,RAB30_SALESNO " & _
                 "WHERE SUBSTR(TM23,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 " & _
                 "AND CP01 IN (" & strTmp(2) & ") AND CP09<'B'    " & _
               "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(TM78,1,8) FROM CASEPROGRESS,TRADEMARK,RAB30_SALESNO " & _
                 "WHERE SUBSTR(TM78,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 " & _
                 "AND CP01 IN (" & strTmp(2) & ") AND CP09<'B'    " & _
               "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(TM79,1,8) FROM CASEPROGRESS,TRADEMARK,RAB30_SALESNO " & _
                 "WHERE SUBSTR(TM79,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 " & _
                 "AND CP01 IN (" & strTmp(2) & ") AND CP09<'B'    " & _
               "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(TM80,1,8) FROM CASEPROGRESS,TRADEMARK,RAB30_SALESNO " & _
                 "WHERE SUBSTR(TM80,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 " & _
                 "AND CP01 IN (" & strTmp(2) & ") AND CP09<'B'    " & _
               "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(TM81,1,8) FROM CASEPROGRESS,TRADEMARK,RAB30_SALESNO " & _
                 "WHERE SUBSTR(TM81,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 " & _
                 "AND CP01 IN (" & strTmp(2) & ") AND CP09<'B'    "
   '法務
   strTmp(3) = Replace(SQLGrpStr(strB2, 3), ",' '", "")
   strExc(3) = "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(LC11,1,8) FROM CASEPROGRESS,LAWCASE,RAB30_SALESNO " & _
               "WHERE SUBSTR(LC11,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 " & _
                 "AND CP01 IN (" & strTmp(3) & ") AND CP09<'B'    " & _
                 "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(LC43,1,8) FROM CASEPROGRESS,LAWCASE,RAB30_SALESNO " & _
               "WHERE SUBSTR(LC43,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 " & _
                 "AND CP01 IN (" & strTmp(3) & ") AND CP09<'B'    " & _
                 "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(LC44,1,8) FROM CASEPROGRESS,LAWCASE,RAB30_SALESNO " & _
               "WHERE SUBSTR(LC44,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 " & _
                 "AND CP01 IN (" & strTmp(3) & ") AND CP09<'B'    " & _
                 "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(LC45,1,8) FROM CASEPROGRESS,LAWCASE,RAB30_SALESNO " & _
               "WHERE SUBSTR(LC45,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 " & _
                 "AND CP01 IN (" & strTmp(3) & ") AND CP09<'B'    " & _
                 "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(LC46,1,8) FROM CASEPROGRESS,LAWCASE,RAB30_SALESNO " & _
               "WHERE SUBSTR(LC46,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 " & _
                 "AND CP01 IN (" & strTmp(3) & ") AND CP09<'B'    "
   '顧問
   strTmp(4) = Replace(SQLGrpStr(strB2, 4), ",' '", "")
   strExc(4) = "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(HC05,1,8) FROM CASEPROGRESS,HIRECASE,RAB30_SALESNO " & _
               "WHERE SUBSTR(HC05,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 " & _
                 "AND CP01 IN (" & strTmp(4) & ") AND CP09<'B'    " & _
                 "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(HC24,1,8) FROM CASEPROGRESS,HIRECASE,RAB30_SALESNO " & _
               "WHERE SUBSTR(HC24,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 " & _
                 "AND CP01 IN (" & strTmp(4) & ") AND CP09<'B'    " & _
                 "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(HC25,1,8) FROM CASEPROGRESS,HIRECASE,RAB30_SALESNO " & _
               "WHERE SUBSTR(HC25,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 " & _
                 "AND CP01 IN (" & strTmp(4) & ") AND CP09<'B'    " & _
                 "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(HC26,1,8) FROM CASEPROGRESS,HIRECASE,RAB30_SALESNO " & _
               "WHERE SUBSTR(HC26,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 " & _
                 "AND CP01 IN (" & strTmp(4) & ") AND CP09<'B'    " & _
                 "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(HC27,1,8) FROM CASEPROGRESS,HIRECASE,RAB30_SALESNO " & _
               "WHERE SUBSTR(HC27,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 AND CP04(+)=HC04 " & _
                 "and CP01 IN (" & strTmp(4) & ") and cp09<'B'    "
   '服務
   strTmp(5) = Replace(SQLGrpStr(strB2, 5), ",' '", "")
   strExc(5) = "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(SP08,1,8) FROM CASEPROGRESS,SERVICEPRACTICE,RAB30_SALESNO " & _
               "WHERE SUBSTR(SP08,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 " & _
                 "AND CP01 IN (" & strTmp(5) & ") AND CP09<'B'    " & _
                 "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(SP58,1,8) FROM CASEPROGRESS,SERVICEPRACTICE,RAB30_SALESNO " & _
               "WHERE SUBSTR(SP58,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 " & _
                 "AND CP01 IN (" & strTmp(5) & ") AND CP09<'B'    " & _
                 "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(SP59,1,8) FROM CASEPROGRESS,SERVICEPRACTICE,RAB30_SALESNO " & _
               "WHERE SUBSTR(SP59,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 " & _
                 "AND CP01 IN (" & strTmp(5) & ") AND CP09<'B'    " & _
                 "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(SP65,1,8) FROM CASEPROGRESS,SERVICEPRACTICE,RAB30_SALESNO " & _
               "WHERE SUBSTR(SP65,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 " & _
                 "AND CP01 IN (" & strTmp(5) & ") AND CP09<'B'    " & _
                 "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(SP66,1,8) FROM CASEPROGRESS,SERVICEPRACTICE,RAB30_SALESNO " & _
               "WHERE SUBSTR(SP66,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 " & _
                 "AND CP01 IN (" & strTmp(5) & ") AND CP09<'B'    "
   '移轉人(讓與人)
   strExc(6) = "UNION SELECT cp01,cp02,cp03,cp04,cp05,cp10,cp12,cp13,cp57,substr(cp55||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
                 "WHERE SUBSTR(CP55||'00000000',1,8)=CU01(+) AND CU01 IS NOT NULL and cp09<'B'    " & _
               "UNION SELECT cp01,cp02,cp03,cp04,cp05,cp10,cp12,cp13,cp57,substr(cp93||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
                 "WHERE SUBSTR(cp93||'00000000',1,8)=cu01(+) and cu01 is not null and cp09<'B'    " & _
               "UNION SELECT cp01,cp02,cp03,cp04,cp05,cp10,cp12,cp13,cp57,substr(cp94||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
                 "WHERE SUBSTR(cp94||'00000000',1,8)=cu01(+) and cu01 is not null and cp09<'B'    " & _
               "UNION SELECT cp01,cp02,cp03,cp04,cp05,cp10,cp12,cp13,cp57,substr(cp95||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
                 "WHERE SUBSTR(CP95||'00000000',1,8)=CU01(+) AND CU01 IS NOT NULL and cp09<'B'    " & _
               "UNION SELECT cp01,cp02,cp03,cp04,cp05,cp10,cp12,cp13,cp57,substr(cp96||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
                 "WHERE SUBSTR(cp96||'00000000',1,8)=cu01(+) and cu01 is not null and cp09<'B'    "
   'Added by Lydia 2024/07/03 移轉申請人(讓與申請人)
   strExc(8) = "UNION SELECT cp01,cp02,cp03,cp04,cp05,cp10,cp12,cp13,cp57,substr(cp56||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
                 "WHERE SUBSTR(CP56||'00000000',1,8)=CU01(+) AND CU01 IS NOT NULL and cp09<'B'    " & _
               "UNION SELECT cp01,cp02,cp03,cp04,cp05,cp10,cp12,cp13,cp57,substr(cp89||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
                 "WHERE SUBSTR(cp89||'00000000',1,8)=cu01(+) and cu01 is not null and cp09<'B'    " & _
               "UNION SELECT cp01,cp02,cp03,cp04,cp05,cp10,cp12,cp13,cp57,substr(cp90||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
                 "WHERE SUBSTR(cp90||'00000000',1,8)=cu01(+) and cu01 is not null and cp09<'B'    " & _
               "UNION SELECT cp01,cp02,cp03,cp04,cp05,cp10,cp12,cp13,cp57,substr(cp91||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
                 "WHERE SUBSTR(CP91||'00000000',1,8)=CU01(+) AND CU01 IS NOT NULL and cp09<'B'    " & _
               "UNION SELECT cp01,cp02,cp03,cp04,cp05,cp10,cp12,cp13,cp57,substr(cp92||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
                 "WHERE SUBSTR(cp92||'00000000',1,8)=cu01(+) and cu01 is not null and cp09<'B'    "
   'Added by Lydia 2023/11/07 案源: 抓收文資料也要抓法律所案源資料(但要剔除已放棄的案源)
   'Modified by Lydia 2023/11/30 (Email:112/11/13) 拿掉" AND LOS07 IS NULL"
            '若於12年內有收文 , 但最後仍以銷案處理掉的客戶, 因狀況太多元, 為避免客戶關係有所影響, 仍視為有收文之資訊, 不列入待活化客戶
            '待活化的收文條件原剔除了已取消收文的案件及法律所案源被放棄的條件，請再修改只要有收文或是有填寫案源接洽單都視為有往來。
   strExc(7) = "UNION SELECT CP01,CP02,CP03,CP04,CP05,CP10,CP12,CP13,CP57,SUBSTR(LOS05,1,8) PA26 FROM LAWOFFICESOURCE,CASEPROGRESS,RAB30_SALESNO " & _
               "WHERE SUBSTR(LOS05,1,8)=CU01(+) AND CU01 IS NOT NULL AND LOS01=CP09(+) "
   'Added by Lydia 2024/07/03 增加合約資料
   strExc(9) = "UNION SELECT 'L' AS CP01,'888888' AS CP02,'0' AS CP03,'00' AS CP04,CT04 AS CP05,'001',B.CU12,B.CU13,NULL,SUBSTR(B.CU01,1,8) AS PA26 FROM Contract,RAB30_SALESNO A,CUSTOMER B " & _
               "WHERE SUBSTR(A.CU01,1,8)=CT02(+) AND CT02 IS NOT NULL AND A.CU01=B.CU01(+) AND '0'=B.CU02(+) GROUP BY CT04,B.CU12,B.CU13,SUBSTR(B.CU01,1,8) "
   'Modified by Lydia 2023/11/07 +strExc(7)
   'Modified by Lydia 2024/07/03 +strExc(8),strExc(9)
   strSql = "INSERT INTO RAB30_SALESNOA (SELECT PA26,MAX(CP05) CP05 FROM ( " & _
            Mid(strExc(1), 6) & " " & strExc(2) & " " & strExc(3) & " " & strExc(4) & " " & strExc(5) & " " & strExc(6) & " " & strExc(8) & " " & strExc(7) & " " & strExc(9) & _
            ") GROUP BY PA26)"
   
   cnnConnection.Execute strSql
   
   '更新客戶的最大收文日
   strSql = "UPDATE RAB30_SALESNO B SET B.MAXCP05=(SELECT A.MAXCP05 FROM RAB30_SALESNOA A WHERE B.CU01=A.CU01)"
   cnnConnection.Execute strSql
   strSql = "UPDATE RAB30_SALESNO B SET B.GPMAXCP05=(SELECT MAX(A.MAXCP05) FROM RAB30_SALESNOA A WHERE SUBSTR(B.CU01,1,6)=SUBSTR(A.CU01,1,6)) "
   cnnConnection.Execute strSql
   strSql = "Update RAB30_SALESNO Set Gpcu=(Select Count(*) From RAB30_SALESNOC Where CU01=CUNO group by cuno) "
   cnnConnection.Execute strSql
   
   '更新各系統的統計量
   strSql = "update RAB30_SALESNO set pcase='Y' where cu01 in (select distinct pa26 from ( " & _
            "SELECT distinct SUBSTR(PA26,1,8) PA26 FROM CASEPROGRESS,PATENT,RAB30_SALESNO " & _
              "WHERE SUBSTR(PA26,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 " & _
              "And Cp01 In (" & strTmp(1) & ") And Cp09<'B'    and cp05>" & strDate0 & " " & _
            "UNION SELECT distinct SUBSTR(PA27,1,8) FROM CASEPROGRESS,PATENT,RAB30_SALESNO " & _
              "WHERE SUBSTR(PA27,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 " & _
              "And Cp01 In (" & strTmp(1) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
            "UNION SELECT distinct SUBSTR(PA28,1,8) FROM CASEPROGRESS,PATENT,RAB30_SALESNO " & _
              "WHERE SUBSTR(PA28,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 " & _
              "And Cp01 In (" & strTmp(1) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
            "UNION SELECT distinct SUBSTR(PA29,1,8) FROM CASEPROGRESS,PATENT,RAB30_SALESNO " & _
              "WHERE SUBSTR(Pa29,1,8)=Cu01(+) And Cu01 Is Not Null And Cp01(+)=Pa01 And Cp02(+)=Pa02 And Cp03(+)=Pa03 And CP04(+)=PA04 " & _
              "And Cp01 In (" & strTmp(1) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
            "UNION SELECT distinct SUBSTR(PA30,1,8) From Caseprogress,Patent,RAB30_SALESNO " & _
              "WHERE SUBSTR(PA30,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 " & _
              "And Cp01 In (" & strTmp(1) & ") And Cp09<'B'    and cp05>" & strDate0 & ")) "
   cnnConnection.Execute strSql
   strSql = "update RAB30_SALESNO Set Pcase='Y' Where Cu01 In (Select Distinct Cp55 from ( " & _
            "SELECT distinct SUBSTR(cp55||'00000000',1,8) cp55 from caseprogress,RAB30_SALESNO " & _
              "WHERE SUBSTR(Cp55||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " and Cp01 In (" & strTmp(1) & ")" & _
            "UNION SELECT distinct SUBSTR(cp93||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
              "WHERE SUBSTR(Cp93||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " and Cp01 In (" & strTmp(1) & ") " & _
            "UNION SELECT distinct SUBSTR(cp94||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
              "WHERE SUBSTR(Cp94||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " and Cp01 In (" & strTmp(1) & ") " & _
            "UNION SELECT distinct SUBSTR(cp95||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
              "WHERE SUBSTR(Cp95||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " and Cp01 In (" & strTmp(1) & ")" & _
            "UNION SELECT distinct SUBSTR(cp96||'00000000',1,8) From Caseprogress,RAB30_SALESNO " & _
              "WHERE SUBSTR(Cp96||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " and Cp01 In (" & strTmp(1) & "))) "
   cnnConnection.Execute strSql
   'Added by Lydia 2024/07/03 移轉申請人(讓與申請人)
   strSql = "update RAB30_SALESNO Set Pcase='Y' Where Cu01 In (Select Distinct Cp56 from ( " & _
            "SELECT distinct SUBSTR(cp56||'00000000',1,8) cp56 from caseprogress,RAB30_SALESNO " & _
              "WHERE SUBSTR(cp56||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " and Cp01 In (" & strTmp(1) & ")" & _
            "UNION SELECT distinct SUBSTR(cp89||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
              "WHERE SUBSTR(cp89||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " and Cp01 In (" & strTmp(1) & ") " & _
            "UNION SELECT distinct SUBSTR(cp90||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
              "WHERE SUBSTR(cp90||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " and Cp01 In (" & strTmp(1) & ") " & _
            "UNION SELECT distinct SUBSTR(cp91||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
              "WHERE SUBSTR(cp91||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " and Cp01 In (" & strTmp(1) & ")" & _
            "UNION SELECT distinct SUBSTR(cp92||'00000000',1,8) From Caseprogress,RAB30_SALESNO " & _
              "WHERE SUBSTR(cp92||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " and Cp01 In (" & strTmp(1) & "))) "
   cnnConnection.Execute strSql
   'end 2024/07/03
   strSql = "update RAB30_SALESNO Set Tcase='Y' Where Cu01 In (Select Distinct tm23 from ( " & _
            "SELECT distinct SUBSTR(TM23,1,8) tm23 FROM CASEPROGRESS,TRADEMARK,RAB30_SALESNO " & _
              "WHERE SUBSTR(TM23,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 " & _
               "AND CP01 IN (" & strTmp(2) & ") AND CP09<'B'    and cp05>" & strDate0 & _
            "UNION SELECT distinct SUBSTR(TM78,1,8) FROM CASEPROGRESS,TRADEMARK,RAB30_SALESNO " & _
               "WHERE SUBSTR(TM78,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 " & _
               "AND CP01 IN (" & strTmp(2) & ") AND CP09<'B'    and cp05>" & strDate0 & _
            "UNION SELECT distinct SUBSTR(TM79,1,8) From Caseprogress,Trademark,RAB30_SALESNO " & _
               "WHERE SUBSTR(TM79,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 " & _
               "AND CP01 IN (" & strTmp(2) & ") AND CP09<'B'    and cp05>" & strDate0 & _
            "UNION SELECT distinct SUBSTR(TM80,1,8) FROM CASEPROGRESS,TRADEMARK,RAB30_SALESNO " & _
               "WHERE SUBSTR(TM80,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 " & _
               "AND CP01 IN (" & strTmp(2) & ") AND CP09<'B'    and cp05>" & strDate0 & _
            "UNION SELECT distinct SUBSTR(TM81,1,8) From Caseprogress,Trademark,RAB30_SALESNO " & _
               "WHERE SUBSTR(TM81,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=TM01 AND CP02(+)=TM02 AND CP03(+)=TM03 AND CP04(+)=TM04 " & _
               "AND CP01 IN (" & strTmp(2) & ") And Cp09<'B'    and cp05>" & strDate0 & " ))"
   cnnConnection.Execute strSql
   strSql = "update RAB30_SALESNO Set Tcase='Y' Where Cu01 In (Select Distinct Cp55 from ( " & _
            "SELECT distinct SUBSTR(cp55||'00000000',1,8) cp55 from caseprogress,RAB30_SALESNO " & _
               "WHERE SUBSTR(Cp55||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " AND CP01 IN (" & strTmp(2) & ") " & _
            "UNION SELECT distinct SUBSTR(cp93||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
               "WHERE SUBSTR(Cp93||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " AND CP01 IN (" & strTmp(2) & ") " & _
            "UNION SELECT distinct SUBSTR(cp94||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
               "WHERE SUBSTR(Cp94||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " AND CP01 IN (" & strTmp(2) & ") " & _
            "UNION SELECT distinct SUBSTR(cp95||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
               "WHERE SUBSTR(Cp95||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " AND CP01 IN (" & strTmp(2) & ") " & _
            "UNION SELECT distinct SUBSTR(cp96||'00000000',1,8) From Caseprogress,RAB30_SALESNO " & _
               "WHERE SUBSTR(Cp96||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " and CP01 IN (" & strTmp(2) & ")))"
   cnnConnection.Execute strSql
   'Added by Lydia 2024/07/03 移轉申請人(讓與申請人)
   strSql = "update RAB30_SALESNO Set Tcase='Y' Where Cu01 In (Select Distinct Cp56 from ( " & _
            "SELECT distinct SUBSTR(cp56||'00000000',1,8) cp56 from caseprogress,RAB30_SALESNO " & _
               "WHERE SUBSTR(cp56||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " AND CP01 IN (" & strTmp(2) & ") " & _
            "UNION SELECT distinct SUBSTR(cp89||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
               "WHERE SUBSTR(cp89||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " AND CP01 IN (" & strTmp(2) & ") " & _
            "UNION SELECT distinct SUBSTR(cp90||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
               "WHERE SUBSTR(cp90||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " AND CP01 IN (" & strTmp(2) & ") " & _
            "UNION SELECT distinct SUBSTR(cp91||'00000000',1,8) from caseprogress,RAB30_SALESNO " & _
               "WHERE SUBSTR(cp91||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " AND CP01 IN (" & strTmp(2) & ") " & _
            "UNION SELECT distinct SUBSTR(cp92||'00000000',1,8) From Caseprogress,RAB30_SALESNO " & _
               "WHERE SUBSTR(cp92||'00000000',1,8)=Cu01(+) And Cu01 Is Not Null And Cp09<'B'    and cp05>" & strDate0 & " and CP01 IN (" & strTmp(2) & ")))"
   cnnConnection.Execute strSql
   'end 2024/07/03
   strExc(3) = "UNION SELECT distinct SUBSTR(LC11,1,8) lc11 FROM CASEPROGRESS,LAWCASE,RAB30_SALESNO " & _
                 "WHERE SUBSTR(LC11,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 " & _
                 "And Cp01 In (" & strTmp(3) & ") And Cp09<'B'    and cp05>" & strDate0 & " " & _
               "UNION SELECT distinct SUBSTR(LC43,1,8) FROM CASEPROGRESS,LAWCASE,RAB30_SALESNO " & _
                 "WHERE SUBSTR(LC43,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 " & _
                 "And Cp01 In (" & strTmp(3) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
               "UNION SELECT distinct SUBSTR(LC44,1,8) FROM CASEPROGRESS,LAWCASE,RAB30_SALESNO " & _
                 "WHERE SUBSTR(LC44,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 " & _
                 "And Cp01 In (" & strTmp(3) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
               "UNION SELECT distinct SUBSTR(LC45,1,8) FROM CASEPROGRESS,LAWCASE,RAB30_SALESNO " & _
                 "WHERE SUBSTR(LC45,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 " & _
                 "And Cp01 In (" & strTmp(3) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
               "UNION SELECT distinct SUBSTR(LC46,1,8) FROM CASEPROGRESS,LAWCASE,RAB30_SALESNO " & _
                 "WHERE SUBSTR(LC46,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=LC01 AND CP02(+)=LC02 AND CP03(+)=LC03 AND CP04(+)=LC04 " & _
                  "And Cp01 In (" & strTmp(3) & ") AND CP09<'B'    and cp05>" & strDate0 & " "
   strExc(5) = "UNION SELECT distinct SUBSTR(sp08,1,8) From Caseprogress,Servicepractice,RAB30_SALESNO " & _
                 "WHERE SUBSTR(SP08,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 " & _
                 "AND CP01 IN (" & strTmp(5) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
               "UNION SELECT distinct SUBSTR(sp58,1,8) FROM CASEPROGRESS,SERVICEPRACTICE,RAB30_SALESNO " & _
                 "WHERE SUBSTR(SP58,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 " & _
                 "AND CP01 IN (" & strTmp(5) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
               "UNION SELECT distinct SUBSTR(sp59,1,8) From Caseprogress,Servicepractice,RAB30_SALESNO " & _
                 "WHERE SUBSTR(SP59,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 " & _
                 "AND CP01 IN (" & strTmp(5) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
               "UNION SELECT distinct SUBSTR(sp65,1,8) FROM CASEPROGRESS,SERVICEPRACTICE,RAB30_SALESNO " & _
                 "WHERE SUBSTR(SP65,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 " & _
                 "AND CP01 IN (" & strTmp(5) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
               "UNION SELECT distinct SUBSTR(sp66,1,8) From Caseprogress,Servicepractice,RAB30_SALESNO " & _
                 "WHERE SUBSTR(SP66,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=SP01 AND CP02(+)=SP02 AND CP03(+)=SP03 AND CP04(+)=SP04 " & _
                 "AND CP01 IN (" & strTmp(5) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
               "UNION SELECT distinct SUBSTR(hc05,1,8) FROM CASEPROGRESS,HIRECASE,RAB30_SALESNO " & _
                 "WHERE SUBSTR(HC05,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 And Cp04(+)=Hc04 " & _
                 "AND CP01 IN (" & strTmp(4) & ") AND CP09<'B'    and cp05>" & strDate0 & " "
   strExc(4) = "UNION SELECT distinct SUBSTR(hc24,1,8) FROM CASEPROGRESS,HIRECASE,RAB30_SALESNO " & _
                 "WHERE SUBSTR(HC24,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 And Cp04(+)=Hc04 " & _
                 "AND CP01 IN (" & strTmp(4) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
               "UNION SELECT distinct SUBSTR(hc25,1,8) FROM CASEPROGRESS,HIRECASE,RAB30_SALESNO " & _
                 "WHERE SUBSTR(HC25,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 And Cp04(+)=Hc04 " & _
                 "AND CP01 IN (" & strTmp(4) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
               "UNION SELECT distinct SUBSTR(hc26,1,8) FROM CASEPROGRESS,HIRECASE,RAB30_SALESNO " & _
                 "WHERE SUBSTR(HC26,1,8)=CU01(+) AND CU01 IS NOT NULL AND CP01(+)=HC01 AND CP02(+)=HC02 AND CP03(+)=HC03 And Cp04(+)=Hc04 " & _
                 "AND CP01 IN (" & strTmp(4) & ") AND CP09<'B'    and cp05>" & strDate0 & " " & _
               "UNION SELECT distinct SUBSTR(hc27,1,8) FROM CASEPROGRESS,HIRECASE,RAB30_SALESNO " & _
                 "WHERE SUBSTR(Hc27,1,8)=Cu01(+) And Cu01 Is Not Null And Cp01(+)=Hc01 And Cp02(+)=Hc02 And Cp03(+)=Hc03 And Cp04(+)=Hc04 " & _
                 "And CP01 IN (" & strTmp(4) & ") And Cp09<'B'    and cp05>" & strDate0 & " "
   'Added by Lydia 2023/11/07 案源: 抓收文資料也要抓法律所案源資料(但要剔除已放棄的案源)
   'Modified by Lydia 2023/11/30 (Email:112/11/13) 拿掉" AND LOS07 IS NULL"
            '若於12年內有收文 , 但最後仍以銷案處理掉的客戶, 因狀況太多元, 為避免客戶關係有所影響, 仍視為有收文之資訊, 不列入待活化客戶
            '待活化的收文條件原剔除了已取消收文的案件及法律所案源被放棄的條件，請再修改只要有收文或是有填寫案源接洽單都視為有往來。
   strExc(6) = "UNION SELECT DISTINCT(SUBSTR(LOS05,1,8)) FROM LAWOFFICESOURCE,CASEPROGRESS,RAB30_SALESNO " & _
                  "WHERE LOS12>=" & strDate0 & " AND SUBSTR(LOS05,1,8)=CU01(+) AND CU01 IS NOT NULL AND LOS01=CP09(+) "
   'Added by Lydia 2024/07/03 增加合約資料
   strExc(7) = "UNION SELECT DISTINCT(SUBSTR(CU01,1,8)) FROM Contract,RAB30_SALESNO " & _
               "WHERE SUBSTR(CU01,1,8)=CT02(+) AND CT02 IS NOT NULL "
   'Modified by Lydia 2023/11/07 +strExc(6)
   'Modified by Lydia 2024/07/03 +strExc(7)
   strSql = "update RAB30_SALESNO Set Ocase='Y' Where Cu01 In (Select Distinct Lc11 from ( " & _
            Mid(strExc(3), 6) & " " & strExc(5) & " " & strExc(4) & " " & strExc(6) & " " & strExc(7) & " ))"
   cnnConnection.Execute strSql
   '尚未設為活化客戶，並且符合以下兩條件:
   'A. １２年內未有收文的案件，且無關係企業之客戶。
   'B. １２年內未有收文的案件，雖有關係企業，但母公司與關係企業均未有收文。
   'Modified by Lydia 2024/07/01 debug: GROUP BY SUBSTR(CU01,1,6),DECODE(PCASE,'Y',1,0),DECODE(TCASE,'Y',1,0),DECODE(OCASE,'Y',1,0) => GROUP BY SUBSTR(CU01,1,6)
   strSql = "INSERT INTO OLDCUSTOMER (OCU01,OCU02) " & _
            "SELECT CU01, TO_CHAR(SYSDATE,'YYYYMMDD') AS SDATE FROM RAB30_SALESNO WHERE CU01 NOT IN (SELECT OCU01 FROM OLDCUSTOMER) " & _
            "AND SUBSTR(CU01,1,6) IN (SELECT GRPNO FROM (" & _
            "SELECT SUBSTR(CU01,1,6) GRPNO,SUM(DECODE(PCASE,'Y',1,0)) P1,SUM(DECODE(TCASE,'Y',1,0)) T1,SUM(DECODE(OCASE,'Y',1,0)) O1 " & _
            "FROM RAB30_SALESNO GROUP BY SUBSTR(CU01,1,6) " & _
            ") WHERE NVL(P1,0)+NVL(T1,0)+NVL(O1,0)=0) "
   cnnConnection.Execute strSql, intI

ErrHnd:
   If Err.Number <> 0 Then
      WLog "自動轉入待活化客戶:" & Err.Description
   End If
   
End Sub

'Added by Lydia 2024/01/15 顧問聘任期間檢查Email設定：在顧問聘任期間並且「顧問專用信箱」=空白，發清單通知智權人員。
Private Sub StrMenu31()
   Dim xlsReport31 As New Excel.Application
   Dim wksReport31 As New Worksheet
   Dim intQ As Integer, stSQL As String, strTmp(1 To 4) As String
   Dim rsQuery As New ADODB.Recordset
   Dim strGrp As String, strGrpMan As String
   Dim nRows As Integer
   Dim arrTmp As Variant, arrTmpW As Variant
   Dim stDate1 As String
   Dim xlsFileName As String
   Dim stCC1 As String
   
On Error GoTo ErrHnd
   
   stDate1 = strSrvDate(1)
   
   Call PUB_KillTempFile(Mid(TextPath & "*顧問專用信箱.*", 2))
   xlsFileName = "顧問專用信箱.xls"
    
   '欄位抬頭
   stSQL = "本所案號,案件名稱,申請人名稱,備註"
   arrTmp = Split(stSQL, ",")
   stSQL = "15,25,25,25"
   arrTmpW = Split(stSQL, ",")
   
   stSQL = Pub_GetChkCU199("2", stDate1, False)  '從共用模組取得SQL
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      rsQuery.MoveFirst
      stCC1 = Pub_GetSpecMan("全所智權部主管") '智權部人員時副本加發「全所智權部主管」
      Do While Not rsQuery.EOF
           If strGrp <> "" & rsQuery.Fields("cp13") Then
               If strGrp <> "" Then
                  If Val(xlsReport31.Version) < 12 Then
                     xlsReport31.Workbooks(1).SaveAs FileName:=App.path & TextPath & strGrp & "_" & xlsFileName, FileFormat:=-4143
                  Else
                     xlsReport31.Workbooks(1).SaveAs FileName:=App.path & TextPath & strGrp & "_" & xlsFileName, FileFormat:=56
                  End If
                  xlsReport31.Workbooks.Close
                  xlsReport31.Quit
                  
                  SendMAPIMail strGrp, TransDate(stDate1, 1) & "_顧問聘任期間檢查Email設定", vbCrLf & vbCrLf & "請進入承辦人系統->智權部->其他->客戶資料修改，選擇其他頁籤在顧問專用信箱輸入Email或無信箱。" & vbCrLf & vbCrLf, App.path & TextPath & strGrp & "_" & xlsFileName, , strGrpMan
               End If
               xlsReport31.SheetsInNewWorkbook = 1 '預設工作表數目
               xlsReport31.Workbooks.add
               xlsReport31.Application.WindowState = xlMinimized
               Set wksReport31 = xlsReport31.Worksheets(1)
               '設定欄位名稱及欄寬
               nRows = 1
               For intQ = 1 To UBound(arrTmp) + 1
                   wksReport31.Range(Chr(intQ + 64) & nRows).Value = arrTmp(intQ - 1)
                   wksReport31.Range(Chr(intQ + 64) & ":" & Chr(intQ + 64)).ColumnWidth = Val(arrTmpW(intQ - 1))
                   wksReport31.Range(Chr(intQ + 64) & nRows).HorizontalAlignment = xlCenter
               Next
               nRows = nRows + 1
               strGrp = "" & rsQuery.Fields("cp13")
               If Left("" & rsQuery.Fields("st15"), 1) = "S" Then
                  strGrpMan = stCC1
               Else
                  strGrpMan = "" & rsQuery.Fields("a0908")
               End If
           End If
           strTmp(1) = rsQuery.Fields("cp01") & "-" & rsQuery.Fields("cp02") & IIf(rsQuery.Fields("cp03") & rsQuery.Fields("cp04") = "000", "", "-" & rsQuery.Fields("cp03") & rsQuery.Fields("cp04"))
           strTmp(2) = "" & rsQuery.Fields("casename")
           If "" & rsQuery.Fields("c01") = "" Then
               strTmp(3) = "新客戶，尚待建檔"
           Else
               strTmp(3) = rsQuery.Fields("c01") & " " & GetCustomerName(rsQuery.Fields("c01"))
           End If
           strTmp(4) = ""
           If "" & rsQuery.Fields("c01") <> "" And "" & rsQuery.Fields("c01a") = "WA" Then
              strTmp(4) = strTmp(4) & "；申請人1沒有設定顧問電子信箱"
           End If
           If "" & rsQuery.Fields("c02") <> "" And "" & rsQuery.Fields("c02a") = "WA" Then
              strTmp(4) = strTmp(4) & "；申請人2沒有設定顧問電子信箱"
           End If
           If "" & rsQuery.Fields("c03") <> "" And "" & rsQuery.Fields("c03a") = "WA" Then
              strTmp(4) = strTmp(4) & "；申請人3沒有設定顧問電子信箱"
           End If
           If "" & rsQuery.Fields("c04") <> "" And "" & rsQuery.Fields("c04a") = "WA" Then
              strTmp(4) = strTmp(4) & "；申請人4沒有設定顧問電子信箱"
           End If
           If "" & rsQuery.Fields("c05") <> "" And "" & rsQuery.Fields("c05a") = "WA" Then
              strTmp(4) = strTmp(4) & "；申請人5沒有設定顧問電子信箱"
           End If
           If strTmp(4) <> "" Then
              strTmp(4) = Mid(strTmp(4), 2)
           End If

           For intQ = 1 To UBound(arrTmp) + 1
               With wksReport31.Range(Chr(intQ + 64) & nRows)
                   .Value = "" & strTmp(intQ)
                   .NumberFormatLocal = "@"
                   wksReport31.Range(Chr(intQ + 64) & nRows).HorizontalAlignment = xlLeft
               End With
           Next intQ
           nRows = nRows + 1
           rsQuery.MoveNext
      Loop

      If strGrp <> "" Then
         If Val(xlsReport31.Version) < 12 Then
            xlsReport31.Workbooks(1).SaveAs FileName:=App.path & TextPath & strGrp & "_" & xlsFileName, FileFormat:=-4143
         Else
            xlsReport31.Workbooks(1).SaveAs FileName:=App.path & TextPath & strGrp & "_" & xlsFileName, FileFormat:=56
         End If
         xlsReport31.Workbooks.Close
         xlsReport31.Quit
         
         SendMAPIMail strGrp, TransDate(stDate1, 1) & "_顧問聘任期間檢查Email設定", vbCrLf & vbCrLf & "請進入承辦人系統->智權部->其他->客戶資料修改，選擇其他頁籤在顧問專用信箱輸入Email或無信箱。" & vbCrLf & vbCrLf, App.path & TextPath & strGrp & "_" & xlsFileName, , strGrpMan
      End If
   End If
      
   Set rsQuery = Nothing
   Set xlsReport31 = Nothing
   Exit Sub
    
ErrHnd:
    
    WLog "顧問聘任期間檢查Email設定:" & Err.Description
    If Val(xlsReport31.Version) < 12 Then
       xlsReport31.Workbooks(1).SaveAs FileName:=App.path & TextPath & strGrp & "_" & xlsFileName, FileFormat:=-4143
    Else
       xlsReport31.Workbooks(1).SaveAs FileName:=App.path & TextPath & strGrp & "_" & xlsFileName, FileFormat:=56
    End If
    xlsReport31.Workbooks.Close
    xlsReport31.Quit
    Set xlsReport31 = Nothing
    Set rsQuery = Nothing
End Sub

'Add By Sindy 2024/3/14 外商收文未發文清單
'外商->報表列印->收文未發文 每月固定報表,修改為每月1日由系統自動發E-Mail通知
Private Sub StrMenu32()
Dim RsQ As New ADODB.Recordset
Dim strTemp(1 To 15) As String
Dim TempFileName As String
Dim strDate1 As String, StrDate2 As String
Dim strTo As String
Dim jj As Integer
Dim RptCnt As Integer, intCnt As Integer
Dim strCon As String, strEmp As String
Dim strText As String
Dim strTempName As String, strFileName As String

On Err GoTo ErrHandle
   
   strDate1 = "20160101" '固定
   StrDate2 = Left(CompDate(1, -1, strSrvDate(1)), 6) & "31" '上個月最後一天
   
   For RptCnt = 1 To 2
      '1=日文組
      If RptCnt = 1 Then
         strCon = " And S2.ST16='4'"
         strSql = "select st01,st02 from staff where st03='F11' and st04='1' and st16='4' and st01<'F'"
      '2=英文組
      Else
         strCon = " And S2.ST16='2'"
         strSql = "select st01,st02 from staff where st03='F11' and st04='1' and st16='2' and st01<'F'"
      End If
      intI = 1: strTo = ""
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            strTo = strTo & ";" & RsTemp.Fields("st01")
            RsTemp.MoveNext
         Loop
      End If
      
      If strTo <> "" Then
         strTo = Mid(strTo, 2)
         strFileName = "收文未發文明細表-" & IIf(RptCnt = 1, "日文組", "英文組")
         TempFileName = App.path & TextPath & strFileName & ".txt"
         If Dir(TempFileName) <> "" Then Kill TempFileName
         
         strSql = "select S2.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)),'','',CP06,CP07,PTM03,CPM03,S1.ST02,NA03,TM23,'','','','',TM44,CP44,CP12,A0902,S2.ST15 AS D" & _
                  " FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,STAFF S1,STAFF S2,PATENTTRADEMARKMAP,NATION,ACC090,ENGINEERPROGRESS" & _
                  " WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & strDate1 & " AND " & StrDate2 & ") AND CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'" & _
                  " AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+)" & _
                  " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)" & _
                  " AND cp01=cpm01(+) AND CP10=CPM02(+)" & _
                  " AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND (TM29<>'Y' OR TM29 IS NULL) AND CP12=A0901(+)" & _
                  " AND CP01 IN ('TF','T','FCT','CFT',' ') AND S2.ST15>='F10' AND S2.ST15<='F19'" & strCon & _
                  " Union All " & _
                  "select S2.ST01 AS A,CP05 AS B,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07)),'','',CP06,CP07,'',CPM03,S1.ST02,NA03,SP08,SP58,SP59,'','',SP26,CP44,CP12,A0902,S2.ST15 AS D" & _
                  " FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2,NATION,ACC090,ENGINEERPROGRESS" & _
                  " WHERE EP02(+)=CP09 AND (CP05 BETWEEN " & strDate1 & " AND " & StrDate2 & ") AND CP158=0 AND CP159=0 AND SUBSTR(CP09,1,1) <> 'D'" & _
                  " AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+)" & _
                  " AND CP14=S1.ST01(+) AND CP13=S2.ST01(+)" & _
                  " AND cp01=cpm01(+) AND CP10=CPM02(+)" & _
                  " AND SP09=NA01(+) AND (SP15<>'Y' OR SP15 IS NULL) AND CP12=A0901(+)" & _
                  " AND CP01 IN ('S','CFC',' ') AND S2.ST15>='F10' AND S2.ST15<='F19'" & strCon & _
                  " ORDER BY D,A,B,C"
         intI = 1
         Set RsQ = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsQ.RecordCount > 0 Then
               With RsQ
                  .MoveFirst
                  
                  strText = String(50, " ") & "收文未發文明細表" & vbCrLf
                  strText = strText & "列印順序：" & convForm("智權人員", 85) & "列印日期：" & ChangeTStringToTDateString(TransDate(strSrvDate(1), 1)) & vbCrLf
                  strText = strText & "業 務 區：外商承辦" & vbCrLf
                  strText = strText & "收文日期：" & ChangeTStringToTDateString(TransDate(strDate1, 1)) & "-" & ChangeTStringToTDateString(TransDate(GetLastDay(StrDate2), 1)) & vbCrLf & vbCrLf
                  strText = strText & "智權人員 收文日　  本所案號　　    案件名稱 申請人       代理人       本所期限  法定期限  種類     案件性質 承辦人   申請國家" & vbCrLf
                  strText = strText & "======== ========= =============== ======== ============ ============ ========= ========= ======== ======== ======== ========" & vbCrLf
                  strEmp = "": intCnt = 0
                  Do While Not .EOF
                     '換智權人員時
                     If strEmp <> "" And strEmp <> Trim(.Fields(0)) Then
                        strText = strText & "-----------------------------------------------------------------------------------------------------------------------------" & vbCrLf
                        strText = strText & "          小計： " & intCnt & " 筆" & vbCrLf
                        strText = strText & "-----------------------------------------------------------------------------------------------------------------------------" & vbCrLf
                        intCnt = 0
                     End If
                     '明細資料
                     For jj = 1 To 12
                        strTemp(jj) = ""
                     Next jj
                     '智權人員
                     strEmp = Trim(.Fields(0)): intCnt = intCnt + 1
                     strTemp(1) = convForm(GetPrjSalesNM(strEmp), 8)
                     strTemp(2) = convForm(ChangeWStringToTDateString(.Fields("B")), 9) '收文日
                     strTemp(3) = convForm(Trim(.Fields("C")), 15) '本所案號
                     strTemp(4) = convForm(Trim(.Fields(3)), 8) '案件名稱
                     strTemp(5) = convForm(GetPrjPeople1(Trim("" & .Fields("TM23"))), 12) '申請人
                     '代理人
                     If Not IsNull(Trim("" & .Fields("TM44"))) Then
                        strTemp(6) = Trim("" & .Fields("TM44"))
                     Else
                        If Not IsNull(Trim("" & .Fields("CP44"))) Then
                           strTemp(6) = Trim("" & .Fields("CP44"))
                        End If
                     End If
                     '若系統種類對照檔SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
                     'strTemp(6) = GetPrjName1(strTemp(6))
                     If PUB_GetAgentName(SystemNumber(strTemp(3), 1), strTemp(6), strTempName) = True Then
                        strTemp(6) = convForm(strTempName, 12)
                     Else
                        strTemp(6) = ""
                     End If
                     strTemp(7) = convForm(ChangeWStringToTDateString("" & .Fields("CP06")), 9) '本所期限
                     strTemp(8) = convForm(ChangeWStringToTDateString("" & .Fields("CP07")), 9) '法定期限
                     strTemp(9) = convForm(Trim("" & .Fields("PTM03")), 8) '種類
                     strTemp(10) = convForm(Trim("" & .Fields("CPM03")), 8) '案件性質
                     strTemp(11) = convForm(Trim("" & .Fields(10)), 8) '承辦人
                     strTemp(12) = convForm(Trim("" & .Fields("NA03")), 8) '申請國家
                     
                     strText = strText & strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & " " & strTemp(6) & " " & strTemp(7) & " " & strTemp(8) & " " & strTemp(9) & " " & strTemp(10) & " " & strTemp(11) & " " & strTemp(12) & vbCrLf
                     .MoveNext
                  Loop
                  '結束
                  strText = strText & "-----------------------------------------------------------------------------------------------------------------------------" & vbCrLf
                  strText = strText & "          小計： " & intCnt & " 筆" & vbCrLf
                  strText = strText & "-----------------------------------------------------------------------------------------------------------------------------" & vbCrLf
                  strText = strText & "  外商承辦合計： " & RsQ.RecordCount & " 筆" & vbCrLf
                  strText = strText & "-----------------------------------------------------------------------------------------------------------------------------" & vbCrLf
               End With
            End If
            Call PUB_SaveTextAsUTF8(TempFileName, strText)
            If Dir(TempFileName) <> "" Then
               strExc(1) = vbCrLf & vbCrLf & vbCrLf & String(2, "　") & "列印注意事項：" & vbCrLf & _
                           vbCrLf & String(4, "　") & "1.利用筆記本開啟附件" & _
                           vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
                           vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
                           vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
                           vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
                           vbCrLf & String(4, "　") & "6.選擇<橫印>   "
               SendMAPIMail strTo, strFileName & "，請檢視！", strExc(1), TempFileName, , Pub_GetSpecMan("V2") '商標處第五級管制人
            End If
         End If
         Set RsQ = Nothing
      End If
   Next RptCnt
   
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      WLog "收文未發文明細表:" & Err.Description
   End If
End Sub

'Add By Sindy 2024/3/15 外商催審表
'外商->報表列印->催審表 每月固定報表,修改為每月1日由系統自動發E-Mail通知
Private Sub StrMenu33()
Dim RsQ As New ADODB.Recordset
Dim strTemp(1 To 15) As String
Dim TempFileName As String
Dim strDate1 As String, StrDate2 As String
Dim strTo As String
Dim jj As Integer
Dim strText As String
Dim strTempName As String, strFileName As String

On Err GoTo ErrHandle
   
   strDate1 = "20210101" '固定; 因阿蓮說她在操作報表時,都沒有更新下一程序的期限
   StrDate2 = Left(CompDate(1, -1, strSrvDate(1)), 6) & "31" '上個月最後一天
   
   strSql = "select st01,st02 from staff where st03='F11' and st04='1' and st16 in('2','4') and st01<'F'"
   intI = 1: strTo = ""
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         strTo = strTo & ";" & RsTemp.Fields("st01")
         RsTemp.MoveNext
      Loop
   End If
   
   If strTo <> "" Then
      strTo = Mid(strTo, 2)
      strFileName = "外商催審表"
      TempFileName = App.path & TextPath & strFileName & ".txt"
      If Dir(TempFileName) <> "" Then Kill TempFileName
      
      strSql = "SELECT NP08,CP27,NVL(TM15,TM12),CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(TM05,NVL(TM06,TM07)),NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,NVL(NA03,NA04)" & _
               " FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION" & _
               " WHERE NP01=CP09(+)" & _
               " AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+)" & _
               " AND CP14=S1.ST01(+) AND NP10=S2.ST01(+)" & _
               " AND SUBSTR(TM23,1,8)=CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+)" & _
               " AND CP01=CPM01(+) AND CP10=CPM02(+)" & _
               " AND TM10=NA01(+)" & _
               " AND NP02 IN ('FCT',' ')" & _
               " AND (TM29 IS NULL OR TM29='') AND NP08>=" & strDate1 & " AND NP08<=" & StrDate2 & " AND NP07=305 AND (NP06 IS NULL OR NP06='')" & _
               " Union All " & _
               "select NP08,CP27,SP11,CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(SP05,NVL(SP06,SP07)),NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)),S1.ST02,S2.ST02,CP22,NP01,NP07,NP22,NP09,NVL(NA03,NA04)" & _
               " FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,CASEPROPERTYMAP,CUSTOMER,NATION" & _
               " WHERE NP01=CP09(+)" & _
               " AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)" & _
               " AND CP14=S1.ST01(+) AND NP10=S2.ST01(+)" & _
               " AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+)" & _
               " AND CP01=CPM01(+) AND CP10=CPM02(+)" & _
               " AND SP09=NA01(+)" & _
               " AND NP02 IN (' ')" & _
               " AND (SP15 IS NULL OR SP15='') AND NP08>=" & strDate1 & " AND NP08<=" & StrDate2 & " AND NP07=305 AND (NP06 IS NULL OR NP06='')"
      intI = 1
      Set RsQ = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsQ.RecordCount > 0 Then
            With RsQ
               .MoveFirst
               
               strText = String(50, " ") & "外商催審表" & vbCrLf
               strText = strText & "催審期限：" & convForm(ChangeTStringToTDateString(TransDate(strDate1, 1)) & "-" & _
                           ChangeTStringToTDateString(TransDate(GetLastDay(StrDate2), 1)), 80) & _
                           "列印日期：" & ChangeTStringToTDateString(TransDate(strSrvDate(1), 1)) & vbCrLf & vbCrLf
               strText = strText & "催審期限  發文日　  申請案號/審定號      本所案號　　    案件名稱             案件性質   申請人          承辦人    智權人員 " & vbCrLf
               strText = strText & "========= ========= ==================== =============== ==================== ========== =============== ========= =========" & vbCrLf
               Do While Not .EOF
                  '明細資料
                  For jj = 1 To 9
                     strTemp(jj) = ""
                  Next jj
                  strTemp(1) = convForm(ChangeWStringToTDateString(.Fields(0)), 9) '催審期限
                  strTemp(2) = convForm(ChangeWStringToTDateString(.Fields(1)), 9) '發文日
                  strTemp(3) = convForm(Trim("" & .Fields(2)), 20) '申請案號/審定號
                  strTemp(4) = convForm(Trim(.Fields(3)), 15) '本所案號
                  strTemp(5) = convForm(Trim(.Fields(4)), 20) '案件名稱
                  strTemp(6) = convForm(Trim("" & .Fields(5)), 10) '案件性質
                  strTemp(7) = convForm(Trim("" & .Fields(6)), 15) '申請人
                  strTemp(8) = convForm(Trim("" & .Fields(7)), 9) '承辦人
                  strTemp(9) = convForm(Trim(.Fields(8)), 9) '智權人員
                  
                  strText = strText & strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & " " & strTemp(6) & " " & strTemp(7) & " " & strTemp(8) & " " & strTemp(9) & vbCrLf
                  .MoveNext
               Loop
            End With
         End If
         Call PUB_SaveTextAsUTF8(TempFileName, strText)
         If Dir(TempFileName) <> "" Then
            strExc(1) = vbCrLf & vbCrLf & vbCrLf & String(2, "　") & "列印注意事項：" & vbCrLf & _
                        vbCrLf & String(4, "　") & "1.利用筆記本開啟附件" & _
                        vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
                        vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
                        vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
                        vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
                        vbCrLf & String(4, "　") & "6.選擇<橫印>   "
      
            SendMAPIMail strTo, strFileName & "，請檢視！", strExc(1), TempFileName, , Pub_GetSpecMan("V2") '商標處第五級管制人
         End If
      End If
      Set RsQ = Nothing
   End If
   
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      WLog "外商催審表:" & Err.Description
   End If
End Sub

'Add By Sindy 2024/3/18 外商FCT延展管制表
'外商->報表列印->期限管制表 每月固定報表,修改為每月1日由系統自動發E-Mail通知
Private Sub StrMenu34()
Dim RsQ As New ADODB.Recordset
Dim strTemp(1 To 15) As String
Dim TempFileName As String
Dim strDate1 As String, StrDate2 As String
Dim strTo As String
Dim jj As Integer
Dim strText As String
Dim strTempName As String, strFileName As String
Dim RptCnt As Integer, strCon As String
Dim eFile As String
Dim strEType As String
Dim strTM01 As String, strTM02 As String, strTM03 As String, strTM04 As String
Dim strWhere_TM As String, strWhere_SP As String, strOrder As String
   
On Err GoTo ErrHandle
   
   eFile = ""
   For RptCnt = 1 To 3
      If RptCnt = 3 Then
         '第2次催FCT延展管制表
         strDate1 = Left(CompDate(1, 1, strSrvDate(1)), 6) & "01" '當月+1個月第一天
         StrDate2 = Left(CompDate(1, 1, strSrvDate(1)), 6) & "31" '當月+1個月最後一天
         strFileName = "第2次FCT延展管制表(英文組)"
         strCon = " AND NOT(NP02='FCT' AND substr(FA1.FA10,1,3)='011')"
         strOrder = " order by sort_1, sort_2, R093003"
      Else
         '第1次催FCT延展管制表
         strDate1 = Left(CompDate(1, 8, strSrvDate(1)), 6) & "01" '當月+8個月第一天
         StrDate2 = Left(CompDate(1, 8, strSrvDate(1)), 6) & "31" '當月+8個月最後一天
         If RptCnt = 2 Then
            strFileName = "第1次FCT延展管制表(日文組)"
            strCon = " AND (NP02='FCT' AND substr(FA1.FA10,1,3)='011')"
            strOrder = " order by sort_1, sort_2, R093003"
         Else
            strFileName = "第1次FCT延展管制表(英文組)"
            strCon = " AND NOT(NP02='FCT' AND substr(FA1.FA10,1,3)='011')"
            strOrder = " order by Etype, sort_1, sort_2, R093003"
         End If
      End If
      
      strSql = "Delete from R030403_2 where ID='" & strUserNum & "'"
      cnnConnection.Execute strSql
      
      strWhere_TM = " WHERE NP01=CP09(+)" & _
               " AND NP02=TM01(+) AND NP03=TM02(+) AND NP04=TM03(+) AND NP05=TM04(+)" & _
               " AND NP10=S1.ST01(+) AND CP14=S2.ST01(+)" & _
               " AND Fa2.FA10=N4.na01(+) and FA1.FA10=N2.NA01(+)" & _
               " AND TM10=N1.NA01(+)" & _
               " AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+))" & _
               " AND SUBSTR(TM23,1,8)=CU1.CU01(+) AND decode(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU1.CU02(+)" & _
               " AND SUBSTR(TM44,1,8)=FA1.FA01(+) AND SUBSTR(TM44,9,1)=FA1.FA02(+)" & _
               " AND CU1.CU10=N3.NA01(+) And CU2.CU10=N5.NA01(+)" & _
               " AND SUBSTR(TM33,1,8)=FA2.FA01(+) AND SUBSTR(TM33,9,1)=FA2.FA02(+)" & _
               " AND SUBSTR(TM33,1,8)=CU2.CU01(+) AND SUBSTR(TM33,9,1)=CU2.CU02(+)" & _
               " AND NP02 IN ('FCT',' ') AND ( NP07=102 OR NP07=0) and (tm29 is null or tm29 <> 'Y' )" & strCon & _
               " AND decode(np02,'CFT','FCTY',np02||decode(np07,716,tm17,102,tm17,'Y'))='FCTY'" & _
               " AND (NP07<>'997' And NP07<>'998') AND (NP06 IS NULL OR NP06='')" & _
               " AND NP09>=" & strDate1 & " AND NP09<=" & StrDate2
      
      strWhere_SP = " WHERE NP01=CP09(+)" & _
               " AND NP02=SP01(+) AND NP03=SP02(+) AND NP04=SP03(+) AND NP05=SP04(+)" & _
               " AND NP10=S1.ST01(+) AND CP14=S2.ST01(+)" & _
               " AND FA10=N2.NA01(+) AND SP09=N1.NA01(+)" & _
               " AND NP02=CPM01(+) AND NP07=TO_NUMBER(CPM02(+))" & _
               " AND SUBSTR(SP08,1,8)=CU01(+) AND decode(SUBSTR(SP08,9,1),'','0',substr(sp08,9,1))=CU02(+)" & _
               " AND SUBSTR(SP26,1,8)=FA01(+) AND SUBSTR(SP26,9,1)=FA02(+)" & _
               " AND CU10=N3.NA01(+)" & _
               " AND NP02 IN (' ') AND ( NP07=102 OR  NP07=0) AND (SP15 IS NULL OR SP15 <> 'Y' )" & Replace(strCon, "FA1.FA10", "FA10") & _
               " AND (NP07<>'997' And NP07<>'998') AND (NP06 IS NULL OR NP06='')" & _
               " AND NP09>=" & strDate1 & " AND NP09<=" & StrDate2
      
      strSql = "INSERT INTO R030403_2(ID,R093003,R093006,R093004,R093008,R093009,R093010) " & _
               "SELECT '" & strUserNum & "',NP02||'-'||NP03||'-'||NP04||'-'||NP05,NVL(TM15,TM12),NP01,NP08,NP09,CP27" & _
               " FROM NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT FA1,FAGENT FA2,CUSTOMER CU1,CUSTOMER CU2, Nation N3,nation N4,nation N5" & _
               strWhere_TM
      strSql = strSql & " Union All " & _
               "select '" & strUserNum & "',NP02||'-'||NP03||'-'||NP04||'-'||NP05,SP11,NP01,NP08,NP09,CP27" & _
               " FROM NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT,CUSTOMER, Nation N3" & _
               strWhere_SP
      cnnConnection.Execute strSql
      
      strSql = "select * from R030403_2 where ID='" & strUserNum & "'"
      intI = 1
      Set RsQ = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsQ.RecordCount > 0 Then
            With RsQ
               .MoveFirst
               Do While Not .EOF
                  strExc(9) = "" & .Fields("R093008") '本所期限
                  strExc(8) = "" & .Fields("R093003") '本所案號
                  strTM01 = "" & SystemNumber(strExc(8), 1)
                  strTM02 = "" & SystemNumber(strExc(8), 2)
                  strTM03 = "" & SystemNumber(strExc(8), 3)
                  strTM04 = "" & SystemNumber(strExc(8), 4)
                  'Add By Sindy 2013/8/16 是否為不催延展者
                  If PUB_ChkCaseIsNoticeScale(strTM01, strTM02, strTM03, strTM04) = False Then
                     strExc(10) = "x" & strExc(8)
                  '2013/8/16 END
                  Else
                     If Val(strExc(9)) < Val(strSrvDate(1)) Then
                         strExc(10) = "*" & strExc(8)
                     Else
                         If Val(strExc(9)) = Val(strSrvDate(1)) Then
                             strExc(10) = "V" & strExc(8)
                         Else
                             'C類未發文
                             If Mid(CheckStr("" & .Fields("R093004")), 1, 1) = "C" And Len(CheckStr("" & .Fields("R093010"))) = 0 Then
                                 strExc(10) = "#" & strExc(8)
                             'Add By Sindy 2021/3/19 為排序
                             Else
                                 strExc(10) = " " & strExc(8)
                             '2021/3/19 END
                             End If
                         End If
                     End If
                  End If
                  
                  strSql = "update R030403_2 set R093003='" & strExc(10) & "'" & _
                           " where ID='" & strUserNum & "' AND R093006='" & .Fields("R093006") & "' AND R093003='" & strExc(8) & "'"
                  cnnConnection.Execute strSql, intI
                  
                  .MoveNext
               Loop
            End With
         End If
      End If
      'NP02||'-'||NP03||'-'||NP04||'-'||NP05
      strSql = "SELECT S1.ST03,NP10,NP08,NP09,R093003,NVL(TM05,NVL(TM06,TM07)),TM09,NVL(TM15,TM12)" & _
               ",NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),NP15,S2.ST01,NVL(CU1.CU04,NVL(CU1.CU05||CU1.CU88||CU1.CU89||CU1.CU90,CU1.CU06))" & _
               ",NVL(N1.NA03,N1.NA04),decode(np07,'102',decode(tm33,null,NVL(FA1.FA04,NVL(FA1.FA05||FA1.FA63||FA1.FA64||FA1.FA65,FA1.FA06))" & _
               ",NVL(NVL(FA2.FA04,NVL(FA2.FA05||FA2.FA63||FA2.FA64||FA2.FA65,FA2.FA06)),NVL(CU2.CU04,NVL(CU2.CU05||CU2.CU88||CU2.CU89||CU2.CU90,CU2.CU06))))" & _
               ",NVL(FA1.FA04,NVL(FA1.FA05||FA1.FA63||FA1.FA64||FA1.FA65,FA1.FA06)))" & _
               ",decode(np07,'102',decode(tm33,null,NVL(N2.NA03,N2.NA04),NVL(NVL(N4.NA03,N4.NA04),NVL(N5.NA03,N5.NA04))),NVL(N2.NA03,N2.NA04))" & _
               ",TM22,np07,NP01,CP27,TM01,TM02,TM03,TM04,S1.ST15 as ST15,nvl(upper(GetEMailFlag(cp09)),'P') as Etype,decode(substr(R093003,1,1),'x',1,0) as sort_1,Decode(substr(R093003,Length(R093003)-3,1),'T','T','0') as sort_2" & _
               " FROM R030403_2,NEXTPROGRESS,CASEPROGRESS,TRADEMARK,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT FA1,FAGENT FA2,CUSTOMER CU1,CUSTOMER CU2, Nation N3,nation N4,nation N5" & _
               strWhere_TM & " AND ID='" & strUserNum & "' AND NVL(TM15,TM12)=R093006 AND instr(R093003,NP02||'-'||NP03||'-'||NP04||'-'||NP05)>0"
      strSql = strSql & " Union All " & _
               "select S1.ST03,NP10,NP08,NP09,R093003,NVL(SP05,NVL(SP06,SP07)),'',SP11" & _
               ",NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),NP15,S2.ST01,NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06))" & _
               ",NVL(N1.NA03,N1.NA04),NVL(FA04,NVL(FA05||FA63||FA64||FA65,FA06))" & _
               ",NVL(N2.NA03,N2.NA04),SP21,NP07,NP01,CP27,SP01,SP02,SP03,SP04,S1.ST15 as ST15,nvl(upper(GetEMailFlag(cp09)),'P') as Etype,decode(substr(R093003,1,1),'x',1,0) as sort_1,Decode(substr(R093003,Length(R093003)-3,1),'T','T','0') as sort_2" & _
               " FROM R030403_2,NEXTPROGRESS,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,NATION N1,NATION N2,CASEPROPERTYMAP,FAGENT,CUSTOMER, Nation N3" & _
               strWhere_SP & " AND ID='" & strUserNum & "' AND SP11=R093006 AND instr(R093003,NP02||'-'||NP03||'-'||NP04||'-'||NP05)>0"
      strSql = strSql & strOrder
      intI = 1
      Set RsQ = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsQ.RecordCount > 0 Then
            With RsQ
               strEType = "" '預設值
               .MoveFirst
               Do While Not .EOF
                  '要依通知方式分Email通知或紙本通知
                  If InStr(UCase(strOrder), UCase("EType")) > 0 Then '要區分紙本,Email
                     If strEType = "" Or strEType <> "" & .Fields("Etype") Then
                        If strEType <> "" Then
                           strText = strText & "-----------------------------------------------------------------------------------------------------------------------------------" & vbCrLf
                           strText = strText & "* 表示逾本所期限 , V 表示當日本所期限 , # 表示承辦人未通知主管機關來函 , x 表示不催延展" & vbCrLf
                           Call PUB_SaveTextAsUTF8(TempFileName, strText)
                           If eFile <> "" Then eFile = eFile & ";"
                           eFile = eFile & TempFileName
                           strEType = ""
                        End If
                        TempFileName = App.path & TextPath & strFileName & IIf("" & .Fields("Etype") = "E", "-Email", "-紙本") & "通知.txt"
                        If Dir(TempFileName) <> "" Then Kill TempFileName
                        
                        '標題
                        strText = String(50, " ") & strFileName & vbCrLf
                        strText = strText & "法定期限：" & convForm(ChangeTStringToTDateString(TransDate(strDate1, 1)) & "-" & _
                                    ChangeTStringToTDateString(TransDate(GetLastDay(StrDate2), 1)), 80) & _
                                    "列印日期：" & ChangeTStringToTDateString(TransDate(strSrvDate(1), 1)) & vbCrLf & vbCrLf
                        
                        strText = strText & "代理人          申請人　        本所案號　　    案件名稱             商品類別 審定號     代理人國籍 申請國家   本所期限  專用期止日" & vbCrLf
                        strText = strText & "=============== =============== =============== ==================== ======== ========== ========== ========== ========= ==========" & vbCrLf
                     End If
                  ElseIf strEType = "" Then
                     strEType = "不區分"
                     TempFileName = App.path & TextPath & strFileName & ".txt"
                     If Dir(TempFileName) <> "" Then Kill TempFileName
                     
                     '標題
                     strText = String(50, " ") & strFileName & vbCrLf
                     strText = strText & "法定期限：" & convForm(ChangeTStringToTDateString(TransDate(strDate1, 1)) & "-" & _
                                 ChangeTStringToTDateString(TransDate(GetLastDay(StrDate2), 1)), 80) & _
                                 "列印日期：" & ChangeTStringToTDateString(TransDate(strSrvDate(1), 1)) & vbCrLf & vbCrLf
                     
                     strText = strText & "代理人          申請人　        本所案號　　    案件名稱             商品類別 審定號     代理人國籍 申請國家   本所期限  專用期止日" & vbCrLf
                     strText = strText & "=============== =============== =============== ==================== ======== ========== ========== ========== ========= ==========" & vbCrLf
                  End If
                  '明細資料
                  For jj = 1 To 10
                     strTemp(jj) = ""
                  Next jj
                  strTemp(1) = convForm(.Fields(13), 15) '代理人
                  strTemp(2) = convForm(.Fields(11), 15) '申請人
                  strTemp(3) = convForm(.Fields(4), 15) '本所案號
                  strTemp(4) = convForm(Trim(.Fields(5)), 20) '案件名稱
                  strTemp(5) = convForm(Trim(.Fields(6)), 8) '商品類別
                  strTemp(6) = convForm(Trim("" & .Fields(7)), 10) '審定號
                  strTemp(7) = convForm(Trim("" & .Fields(14)), 10) '代理人國籍
                  strTemp(8) = convForm(Trim("" & .Fields(12)), 10) '申請國家
                  strTemp(9) = convForm(ChangeWStringToTDateString(Trim(.Fields(2))), 9) '本所期限
                  strTemp(10) = convForm(ChangeWStringToTDateString(Trim(.Fields(15))), 10) '專用期止日
                  
                  strText = strText & strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & " " & strTemp(6) & " " & strTemp(7) & " " & strTemp(8) & " " & strTemp(9) & " " & strTemp(10) & vbCrLf
                  
                  strEType = "" & .Fields("Etype")
                  .MoveNext
               Loop
            End With
         End If
         strText = strText & "-----------------------------------------------------------------------------------------------------------------------------------" & vbCrLf
         strText = strText & "* 表示逾本所期限 , V 表示當日本所期限 , # 表示承辦人未通知主管機關來函 , x 表示不催延展" & vbCrLf
         Call PUB_SaveTextAsUTF8(TempFileName, strText)
         If eFile <> "" Then eFile = eFile & ";"
         eFile = eFile & TempFileName
      End If
      Set RsQ = Nothing
   Next RptCnt
   If eFile <> "" Then
      strExc(1) = vbCrLf & vbCrLf & vbCrLf & String(2, "　") & "列印注意事項：" & vbCrLf & _
                  vbCrLf & String(4, "　") & "1.利用筆記本開啟附件" & _
                  vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
                  vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
                  vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
                  vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
                  vbCrLf & String(4, "　") & "6.選擇<橫印>   "
      If Val(Mid(strSrvDate(1), 5, 2)) Mod 2 = 0 Then '雙月
         strTo = Right(Pub_GetSpecMan("每月外商延展管制表收件者"), 5) '第2位
      Else '單月
         strTo = Left(Pub_GetSpecMan("每月外商延展管制表收件者"), 5) '第1位
      End If
      SendMAPIMail strTo, Left(strSrvDate(2), 3) & "年" & Mid(strSrvDate(2), 4, 2) & "月外商FCT延展管制表，請檢視！", strExc(1), eFile, , Pub_GetSpecMan("D")
   End If
   
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      WLog "外商FCT延展管制表:" & Err.Description
   End If
End Sub

'Add By Sindy 2024/5/31 針對客戶編號ID仍為66666666(8個6)的通知
Private Sub StrMenu35()
Dim RsQ As New ADODB.Recordset
Dim strTemp(1 To 15) As String
Dim TempFileName As String
Dim strDate1 As String, StrDate2 As String
Dim jj As Integer
Dim strText As String
Dim strFileName As String
Dim RptCnt As Integer
Dim strEmp As String, strCC As String

On Err GoTo ErrHandle

   For RptCnt = 1 To 2
      strExc(10) = "敬啟者:" & vbCrLf & vbCrLf & _
                   "附表中之客戶ID仍設定為8碼之6資訊，若客戶是以下狀況，為目前無須提供ID的客戶，請通知檔案室異動客戶之ID狀況為【不提供ID】：" & vbCrLf & _
                   "1. 客戶目前是委辦非台灣之其他國家案件。" & vbCrLf & _
                   "2. 客戶目前是委辦介紹於法律所之案源。" & vbCrLf & _
                   "3. 客戶目前僅是委辦代管期限之案件。" & vbCrLf & _
                   "4. 其他目前暫無須提供ID的案件。" & vbCrLf
      strExc(10) = strExc(10) & vbCrLf & vbCrLf & String(2, "　") & "列印注意事項：" & _
                   vbCrLf & String(4, "　") & "1.利用筆記本開啟附件" & _
                   vbCrLf & String(4, "　") & "2.將視窗展開到最大" & _
                   vbCrLf & String(4, "　") & "3.取消<自動換行>設定" & _
                   vbCrLf & String(4, "　") & "4.<字型>設定為<細明體 標準 11>" & _
                   vbCrLf & String(4, "　") & "5.左右邊界分別設<10mm 0mm>" & _
                   vbCrLf & String(4, "　") & "6.選擇<橫印>   "

      If RptCnt = 1 Then
         '每個月針對客戶編號ID仍為66666666(8個6)通知智權人員及其區主管
         strDate1 = Left(CompDate(1, -1, strSrvDate(1)), 6) & "01" '當月-1個月第一天
         StrDate2 = Left(CompDate(1, -1, strSrvDate(1)), 6) & "31" '當月-1個月最後一天
         strFileName = "通知客戶ID尚還是8碼6的清單"
         '欄位中包含，客戶編號.客戶名稱.客戶國籍
         strSql = "select cu01||cu02,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cuname" & _
                  ",na03,a0924 as mailCC,cu13 as mailTo,st02" & _
                  " from Customer,nation,staff,acc090new" & _
                  " where CU11='66666666' and (CU182<>'Y' and CU182 is not null) and CU14>=" & strDate1 & " and CU14<=" & StrDate2 & _
                  " and cu10=na01(+) and cu13=st01(+) and st93=a0921(+)" & _
                  " order by cu13 asc,cu01||cu02 asc"
      Else
         '3個月針對客戶編號ID仍為66666666(8個6)通知智權主管:
         '於三個月出報表予智權主管，以確認是否有應處理而未處理的狀況。
         strDate1 = Left(CompDate(1, -3, strSrvDate(1)), 6) & "01" '當月-3個月第一天
         StrDate2 = Left(CompDate(1, -3, strSrvDate(1)), 6) & "31" '當月-3個月最後一天
         strFileName = "通知客戶ID尚還是8碼6的清單_已3個月"
         'Modify By Sindy 2024/10/1 and CU14>=" & strDate1 & " and CU14<=" & StrDate2
         '                          改為 and CU14<=" & StrDate2
         strSql = "select cu01||cu02,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cuname" & _
                  ",na03,a0924 as mailTo,cu13,st02,'' as mailCC" & _
                  " from Customer,nation,staff,acc090new" & _
                  " where CU11='66666666' and (CU182<>'Y' and CU182 is not null) and CU14<=" & StrDate2 & _
                  " and cu10=na01(+) and cu13=st01(+) and st93=a0921(+)" & _
                  " order by a0924 asc,cu01||cu02 asc"
      End If
      intI = 1
      Set RsQ = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strEmp = "": strCC = "" '預設值
         If RsQ.RecordCount > 0 Then
            With RsQ
               .MoveFirst
               Do While Not .EOF
                  If strEmp = "" Or strEmp <> "" & .Fields("mailTo") Then
                     If strEmp <> "" Then
                        TempFileName = App.path & TextPath & strFileName & "_" & strEmp & ".txt"
                        If Dir(TempFileName) <> "" Then Kill TempFileName

                        strText = strText & "-------------------------------------------------------------------------" & vbCrLf
                        Call PUB_SaveTextAsUTF8(TempFileName, strText)
                        SendMAPIMail strEmp, strFileName, strExc(10), TempFileName, , strCC, "97038"
                        strEmp = "": strCC = "" '預設值
                     End If
                     '標題
                     strText = "客戶編號   客戶名稱                       客戶國籍             智權人員" & vbCrLf
                     strText = strText & "========== ============================== ==================== ==========" & vbCrLf
                  'Add By Sindy 2025/5/5
                  Else
                     strText = strText & vbCrLf
                  '2025/5/5 END
                  End If
                  '明細資料
                  For jj = 1 To 4
                     strTemp(jj) = ""
                  Next jj
                  strTemp(1) = convForm(.Fields(0), 10) '客戶編號
                  strTemp(2) = convForm(.Fields(1), 30) '客戶名稱
                  strTemp(3) = convForm(.Fields(2), 20) '客戶國籍
                  strTemp(4) = convForm(Trim(.Fields("st02")), 10) '智權人員
                  strText = strText & strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & vbCrLf

                  '客戶委辦且已發文之案件資訊，包含國家,類型及案件性質，以讓智權人員快速知悉
                  strTemp(1) = Trim(strTemp(1))
                  strSql = "select cp01,cp02,cp03,cp04,cp05,cp10,decode(pa09,'000',cpm03,cpm04) as cp10Nm,na03,decode(pa09,'000',ptm03,ptm04) PTM03,cp27,cp09" & _
                           " from patent,caseprogress,casepropertymap,nation,PATENTTRADEMARKMAP" & _
                           " where (pa26='" & strTemp(1) & "' or pa27='" & strTemp(1) & "' or pa28='" & strTemp(1) & "' or pa29='" & strTemp(1) & "' or pa30='" & strTemp(1) & "')" & _
                           " and pa57 is null and pa108 is null" & _
                           " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp158>0 and cp159=0" & _
                           " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                           " and pa09=na01(+) and PTM01(+)='1' AND PTM02(+)=PA08"
                  strSql = strSql & " union " & _
                           "select cp01,cp02,cp03,cp04,cp05,cp10,decode(tm10,'000',cpm03,cpm04) as cp10Nm,na03,decode(tm10,'000',ptm03,ptm04) PTM03,cp27,cp09" & _
                           " from trademark,caseprogress,casepropertymap,nation,PATENTTRADEMARKMAP" & _
                           " where (tm23='" & strTemp(1) & "' or tm78='" & strTemp(1) & "' or tm79='" & strTemp(1) & "' or tm80='" & strTemp(1) & "' or tm81='" & strTemp(1) & "')" & _
                           " and tm30 is null and tm57 is null" & _
                           " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04 and cp158>0 and cp159=0" & _
                           " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                           " and tm10=na01(+) and '2'=PTM01(+) AND tm08=PTM02(+)"
                  strSql = strSql & " union " & _
                           "select cp01,cp02,cp03,cp04,cp05,cp10,decode(sp09,'000',cpm03,cpm04) as cp10Nm,na03,'' PTM03,cp27,cp09" & _
                           " from servicepractice,caseprogress,casepropertymap,nation" & _
                           " where (sp08='" & strTemp(1) & "' or sp58='" & strTemp(1) & "' or sp59='" & strTemp(1) & "' or sp65='" & strTemp(1) & "' or sp66='" & strTemp(1) & "')" & _
                           " and sp16 is null and sp61 is null" & _
                           " and cp01(+)=sp01 and cp02(+)=sp02 and cp03(+)=sp03 and cp04(+)=sp04 and cp158>0 and cp159=0" & _
                           " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                           " and sp09=na01(+)"
                  strSql = strSql & " union " & _
                           "select cp01,cp02,cp03,cp04,cp05,cp10,decode(lc15,'000',cpm03,cpm04) as cp10Nm,na03,'' PTM03,cp27,cp09" & _
                           " from Lawcase,caseprogress,casepropertymap,nation" & _
                           " where (lc11='" & strTemp(1) & "' or lc43='" & strTemp(1) & "' or lc44='" & strTemp(1) & "' or lc45='" & strTemp(1) & "' or lc46='" & strTemp(1) & "')" & _
                           " and lc09 is null and lc34 is null" & _
                           " and cp01(+)=lc01 and cp02(+)=lc02 and cp03(+)=lc03 and cp04(+)=lc04 and cp158>0 and cp159=0" & _
                           " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                           " and lc15=na01(+)"
                  strSql = strSql & " union " & _
                           "select cp01,cp02,cp03,cp04,cp05,cp10,cpm03 as cp10Nm,'' na03,'' PTM03,cp27,cp09" & _
                           " from Hirecase,caseprogress,casepropertymap" & _
                           " where (hc05='" & strTemp(1) & "' or hc24='" & strTemp(1) & "' or hc25='" & strTemp(1) & "' or hc26='" & strTemp(1) & "' or hc27='" & strTemp(1) & "')" & _
                           " and hc10 is null" & _
                           " and cp01(+)=hc01 and cp02(+)=hc02 and cp03(+)=hc03 and cp04(+)=hc04 and cp158>0 and cp159=0" & _
                           " and cp01=cpm01(+) and cp10=cpm02(+)"
                  strSql = strSql & " order by cp01,cp02,cp03,cp04,cp09 "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     RsTemp.MoveFirst
                     strText = strText & vbCrLf & "   客戶委辦且已發文之案件資訊：" & vbCrLf
                     strText = strText & "   本所案號     國家       類型       案件性質        發文日期" & vbCrLf
                     strText = strText & "   ------------ ---------- ---------- --------------- ----------" & vbCrLf
                     Do While Not RsTemp.EOF
                        For jj = 1 To 5
                           strTemp(jj) = ""
                        Next jj
                        strTemp(1) = convForm(RsTemp.Fields("cp01") & RsTemp.Fields("cp02") & RsTemp.Fields("cp03") & RsTemp.Fields("cp04"), 12) '本所案號
                        strTemp(2) = convForm(RsTemp.Fields("na03"), 10) '國家
                        strTemp(3) = convForm("" & RsTemp.Fields("PTM03"), 10) '類型
                        strTemp(4) = convForm(Trim(RsTemp.Fields("cp10Nm")), 15) '案件性質
                        strTemp(5) = convForm(ChangeWStringToTDateString(Trim(RsTemp.Fields("cp27"))), 10) '發文日期
                        strText = strText & "   " & strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & vbCrLf & vbCrLf
                        RsTemp.MoveNext
                     Loop
                  End If
                  
                  strEmp = "" & .Fields("mailTo")
                  strCC = "" & .Fields("mailCC")
                  'Add By Sindy 2024/10/1 3個月的副本要加發智權主管
                  If RptCnt = 2 Then
                     strCC = IIf(strCC <> "", strCC & ";", "") & Pub_GetSpecMan("全所智權部主管")
                  End If
                  '2024/10/1 END
                  .MoveNext
               Loop
            End With
         End If
         If strEmp <> "" Then
            'Add By Sindy 2025/7/1
            TempFileName = App.path & TextPath & strFileName & "_" & strEmp & ".txt"
            If Dir(TempFileName) <> "" Then Kill TempFileName
            '2025/7/1 END
            
            strText = strText & "-------------------------------------------------------------------------" & vbCrLf
            Call PUB_SaveTextAsUTF8(TempFileName, strText)
            SendMAPIMail strEmp, strFileName, strExc(10), TempFileName, , strCC, "97038"
         End If
      End If
      Set RsQ = Nothing
   Next RptCnt

   Exit Sub

ErrHandle:
   If Err.Number <> 0 Then
      WLog "針對客戶編號ID仍為66666666(8個6)的通知:" & Err.Description
   End If
End Sub

'Added by Morgan 2024/7/31
'批次新增上個月天災不給薪假單
Private Sub StrMenu36()
   Dim stSQL As String, stSQLM As String, intQ As Integer, stLstMon As String
   Dim rsQuery As ADODB.Recordset
   Dim stText As String, stTO As String
   
On Err GoTo ErrHandle

   stLstMon = CompDate(1, -1, strSrvDate(1))
   stSQL = "select * from workday where wd01>=" & stLstMon & " and wd02||wd03||wd04||wd05 is not null"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      stText = ""
      With rsQuery
      Do While Not .EOF
         'Modified by Morgan 2024/8/1 排除颱風假前最後的人事異動為 留職停薪04 或 離職03 者
         'Modified by Morgan 2024/11/25 +排除國外出差者--小魚
         stSQLM = "insert into Staff_Absence(sa01,sa02,sa03,sa04,sa05,sa06,sa07,sa08)" & _
            " select st01 sa01,wd01 sa02,800 sa03,wd01 sa04,1700 sa05,'25' sa06,1 sa07,0 sa08" & _
            " from workday a,staff s where wd01=" & .Fields("wd01") & _
            " and st01<'F' and st13<wd01 and (st51 is null or st51>wd01) and substr(st01,-2)<'9' and st01 not in ('63001','67004') and st03<>'R04'" & _
            " and exists(select * from salarydata where sd01=st01)" & _
            " and not exists(select max(sc02||sc03) from staff_change where sc01=st01 and sc02<wd01 having substr(max(sc02||sc03),-2) in ('04','03'))" & _
            " and not exists(select * from Staff_Absence where sa01=st01 and sa02<=wd01 and sa04>=wd01) " & _
            " and not exists(select * from Staff_Busi_Trip where sb01=st01 and sb02<=wd01 and sb04>=wd01 and sb08 in ('3','4'))"
         '北所
         If .Fields("wd02") = "Y" Then
            stSQL = stSQLM & " and st06='1' and exists(select max(wd01||wd02) from workday b where wd01<a.wd01 having substr(max(wd01||wd02),-1)=a.wd02)"
            cnnConnection.Execute stSQL, intQ
            If intQ > 0 Then
               stText = stText & "已新增 " & ChangeWStringToTDateString(.Fields("wd01")) & " 北所天災不給薪假單 " & intQ & " 筆。" & vbCrLf
            End If
         End If
         '中所
         If .Fields("wd03") = "Y" Then
            stSQL = stSQLM & " and st06='2' and exists(select max(wd01||wd03) from workday b where wd01<a.wd01 having substr(max(wd01||wd03),-1)=a.wd03)"
            cnnConnection.Execute stSQL, intQ
            If intQ > 0 Then
               stText = stText & "已新增 " & ChangeWStringToTDateString(.Fields("wd01")) & " 中所天災不給薪假單 " & intQ & " 筆。" & vbCrLf
            End If
         End If
         '南所
         If .Fields("wd04") = "Y" Then
            stSQL = stSQLM & " and st06='3' and exists(select max(wd01||wd04) from workday b where wd01<a.wd01 having substr(max(wd01||wd04),-1)=a.wd04)"
            cnnConnection.Execute stSQL, intQ
            If intQ > 0 Then
               stText = stText & "已新增 " & ChangeWStringToTDateString(.Fields("wd01")) & " 南所天災不給薪假單 " & intQ & " 筆。" & vbCrLf
            End If
         End If
         '高所
         If .Fields("wd05") = "Y" Then
            stSQL = stSQLM & " and st06='4' and exists(select max(wd01||wd05) from workday b where wd01<a.wd01 having substr(max(wd01||wd05),-1)=a.wd05)"
            cnnConnection.Execute stSQL, intQ
            If intQ > 0 Then
               stText = stText & "已新增 " & ChangeWStringToTDateString(.Fields("wd01")) & " 高所天災不給薪假單 " & intQ & " 筆。" & vbCrLf
            End If
         End If
         .MoveNext
      Loop
      End With
      If stText <> "" Then
         'Modified by Morgan 2024/11/4 新增收件人余佳叡B3028--嘉渝
         'stTO = Pub_GetSpecMan("人事室出缺勤電子簽核")
         stTO = Pub_GetSpecMan("天災相關通知人事人員")
         'end 2024/11/4
         stText = "各所新增筆數如下：" & vbCrLf & vbCrLf & stText & vbCrLf & vbCrLf & "明細資料請執行「一般作業 -> 出缺勤作業 -> 查詢 -> 出缺勤查詢」，清單可選擇「列印輸出。"
         SendMAPIMail stTO, "已批次新增" & Left(ChangeWStringToTDateString(stLstMon), 6) & "月天災不給薪假單！", stText
      End If
   End If
   Set rsQuery = Nothing
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      WLog Err.Description
   End If
   Set rsQuery = Nothing
End Sub

'Added by Morgan 2024/8/19
'每月1日RUN請款對象為德國Y45814010 BASF SE及美國Y33268010 BASF Corporation之外專案件資訊報表，以excel形式列出，email通知尚未請款的項目
'通知對象:承辦人(及其主管)+ 程序人員(及其主管)，新案翻譯則固定通知程序人員(及其主管)
Private Sub StrMenu37()
   Dim stSQL As String, stSQLM As String, intQ As Integer, stLstMon As String
   Dim rsQuery As ADODB.Recordset
   Dim stText As String, stTO As String, strCC  As String, strCC2  As String, strCC3  As String
   Dim xlsReport As New Excel.Application
   Dim wksReport As New Worksheet
   Dim xlsFileName As String
   Dim iCol As Integer, iRow As Integer
   Dim arrTmp
   
On Err GoTo ErrHandle
   
   stLstMon = Left(CompDate(1, -1, strSrvDate(1)), 6)
   'Modified by Morgan 2024/8/30 增加一併通知程序人員及其主管，但新案翻譯(201)固定只通知程序人員及其主管。--Izumi
   'Modified by Morgan 2024/11/4 +408面詢(可能是B類收文)--Izumi/Bobbie
   'Modified by Morgan 2025/4/8 +FG案--Franny
   'Modified by Morgan 2025/4/8 +Y45697000 BASF Schweiz AG --Franny
   stSQL = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) C01" & _
      ",pa75||' '||rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) C02" & _
      ",pa26||' '||rtrim(c1.cu05||' '||c1.cu88||' '||c1.cu89||' '||c1.cu90) C03" & _
      ",decode(pa09,'000',cpm03,cpm04) C04,s1.st02 C05,s2.st02 C06,cp14,decode(cp10,'201',na16,cp14) RCV,s1.st52 CC,na16 CC2,s2.st52 CC3" & _
      " from (select cp01,cp02,cp03,cp04,cp10,cp14,nvl(pa09,sp09) pa09,nvl(pa26,sp08) pa26,nvl(pa75,sp26) pa75,nvl(pa88,sp37) pa88" & _
      " from caseprogress,patent,servicepractice" & _
      " where cp27>=" & stLstMon & "01 and (substr(cp09,1,1) in ('A','C') or cp10='408') and cp20||cp60 is null and cp10 not in('901','902')" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      ") X,customer c1,fagent,casepropertymap,staff s1,nation,staff s2" & _
      " where cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)" & _
      " and Nvl(PA88, Nvl(FA30, Nvl(CU57, PA75))) in ('Y45814010','Y33268010','Y45697000')" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and s1.st01(+)=cp14 and na01(+)=fa10 and s2.st01(+)=na16" & _
      " order by RCV,1"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      stText = "附件表列為BASF上個月已發文但未請款之案件，請盡速於本月20日前完成請款；若已確定無須請款，則可以忽略此通知。"
      arrTmp = Array("本所案號", "代理人", "申請人", "案件性質", "承辦人", "程序人員")
      With rsQuery
      iRow = 0
      Do While Not .EOF
         If .Fields("RCV") <> stTO Then
            If xlsFileName <> "" Then
               wksReport.Range("A1", Chr(UBound(arrTmp) + 65) & "1").Font.Bold = True
               wksReport.Columns(Chr(LBound(arrTmp) + 65) & ":" & Chr(UBound(arrTmp) + 65)).EntireColumn.Font.Name = "標楷體"
               wksReport.Columns(Chr(LBound(arrTmp) + 65) & ":" & Chr(UBound(arrTmp) + 65)).EntireColumn.AutoFit
               If Val(xlsReport.Version) < 12 Then
                  xlsReport.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=-4143
               Else
                  xlsReport.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=56
               End If
               xlsReport.Workbooks.Close
               xlsReport.Quit
               SendMAPIMail stTO, "BASF待請款之案件，請盡速於本月20日前完成請款", stText, App.path & TextPath & xlsFileName, , strCC
            End If
            iRow = 0
            
            stTO = .Fields("RCV")
            If stTO = .Fields("CC2") Then
               strCC = "" & .Fields("CC3")
            Else
               strCC = "" & .Fields("CC")
            End If
            xlsFileName = "BASF已發文未請款案件清單(" & stTO & ")-" & stLstMon & ".xls"
            If Dir(App.path & TextPath & xlsFileName) <> "" Then
               Kill App.path & TextPath & xlsFileName
            End If
            xlsReport.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/12 預設工作表數目
            xlsReport.Workbooks.add
            xlsReport.Application.WindowState = xlMinimized
            Set wksReport = xlsReport.Worksheets(1)
            '設定欄位名稱及欄寬
            iRow = iRow + 1
            For iCol = LBound(arrTmp) To UBound(arrTmp)
                wksReport.Range(Chr(iCol + 65) & iRow).Value = arrTmp(iCol)
                wksReport.Range(Chr(iCol + 65) & iRow).HorizontalAlignment = xlCenter
            Next
         End If
         
         '案件可能會有不同的程序人員,要個案判斷並都寄發
         If Not IsNull(.Fields("CC2")) Then
            If InStr(stTO & ";" & strCC, .Fields("CC2")) = 0 Then
               strCC = strCC & ";" & .Fields("CC2")
            End If
            If Not IsNull(.Fields("CC3")) Then
               If InStr(stTO & ";" & strCC, .Fields("CC3")) = 0 Then
                  strCC = strCC & ";" & .Fields("CC3")
               End If
            End If
         End If
            
         iRow = iRow + 1
         For iCol = LBound(arrTmp) To UBound(arrTmp)
            wksReport.Range(Chr(iCol + 65) & iRow).Value = "" & .Fields(iCol)
         Next
         .MoveNext
      Loop
      
      If xlsFileName <> "" Then
         wksReport.Range("A1", Chr(UBound(arrTmp) + 65) & "1").Font.Bold = True
         wksReport.Columns(Chr(LBound(arrTmp) + 65) & ":" & Chr(UBound(arrTmp) + 65)).EntireColumn.Font.Name = "標楷體"
         wksReport.Columns(Chr(LBound(arrTmp) + 65) & ":" & Chr(UBound(arrTmp) + 65)).EntireColumn.AutoFit
         If Val(xlsReport.Version) < 12 Then
            xlsReport.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=-4143
         Else
            xlsReport.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=56
         End If
         xlsReport.Workbooks.Close
         xlsReport.Quit
         SendMAPIMail stTO, "BASF待請款之案件，請盡速於本月20日前完成請款", stText, App.path & TextPath & xlsFileName, , strCC
      End If
      End With
   End If
   Set rsQuery = Nothing
   Set xlsReport = Nothing
   Set rsQuery = Nothing
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      WLog Err.Description
   End If
   Set rsQuery = Nothing
   Set xlsReport = Nothing
   Set rsQuery = Nothing
End Sub

'Added by Sindy 2025/6/20 每月列出電郵主旨中有主旨標籤的電郵資料(以連結新案變化以分析電郵開拓成效)
Private Sub StrMenu38()
Dim xlsReport38 As New Excel.Application
Dim wksReport38 As New Worksheet
Dim intQ As Integer, stSQL As String
Dim rsQuery As New ADODB.Recordset
Dim nRows As Integer
Dim arrTmp As Variant, arrTmpW As Variant
Dim stDate As String
Dim xlsFileName As String
Dim strTo As String
Dim strTitle As String
   
On Error GoTo ErrHnd

   '上個月
   stDate = Left(TransDate(CompDate(1, -1, strSrvDate(1)), 1), 5)
   
   strTitle = "(往來記綠)列出電郵主旨中有主旨標籤的電郵資料"
   Call PUB_KillTempFile(Mid(TextPath & "*" & strTitle & ".*", 2))
   xlsFileName = stDate & strTitle & ".xls"
    
   '欄位抬頭
   stSQL = "寄/收,寄/收件日期,寄/收件時間,寄件人,員工名稱,部門,組別,客戶編號,客戶名稱,客戶國籍,往來類別代碼,往來類別說明,主旨,往來記錄編號"
   arrTmp = Split(stSQL, ",")
   stSQL = "10,10,10,10,10,10,10,10,10,10,15,15,15,10"
   arrTmpW = Split(stSQL, ",")
   
   stSQL = "SELECT decode(cf09,'T','寄','R','收') AS 寄收," & _
           "sqldateT(cf11) AS 寄收日期,sqltime(cf12) AS 寄收時間," & _
           "cf10,st02,a0902,decode(substr(st03,1,2)||st16,'F12','英文組','F14','日文組',decode(st16,'1','英文組','2','日文組',st16)) st16" & _
           ",cr03,NVL(fa04,Decode(fa05,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),na03" & _
           ",cr05,ac03,cf13,cf01" & _
           " From contactfile, contactrecord, allcode, fagent, nation, staff, acc090" & _
           " Where cf04 >= " & Val(stDate) + 191100 & "01 And cf04 <= " & Val(stDate) + 191100 & "31 And cf09 Is Not Null" & _
           " AND cf01=cr01 AND ac01='11' AND ac02(+)=cr05" & _
           " AND substr(cr03,1,8)=fa01 AND substr(cr03,9,1)=fa02 AND fa10=na01(+)" & _
           " AND decode(instr(cf10,'@taie.com.tw'),0,cf10,cf14)=st01(+) AND a0901(+)=st03" & _
           " AND cf10<>'QPGMR'"
   stSQL = stSQL & " Union All" & _
           " SELECT decode(cf09,'T','寄','R','收') AS 寄收," & _
           "sqldateT(cf11) AS 寄收日期,sqltime(cf12) AS 寄收時間," & _
           "cf10,st02,a0902,decode(substr(st03,1,2)||st16,'F12','英文組','F14','日文組',decode(st16,'1','英文組','2','日文組',st16)) st16" & _
           ",cr03,NVL(CU04,Decode(cu05,NULL,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),na03" & _
           ",cr05,ac03,cf13,cf01" & _
           " From contactfile, contactrecord, allcode, customer, nation, staff, acc090" & _
           " Where cf04 >= " & Val(stDate) + 191100 & "01 And cf04 <= " & Val(stDate) + 191100 & "31 And cf09 Is Not Null" & _
           " AND cf01=cr01 AND ac01='11' AND ac02(+)=cr05" & _
           " AND substr(cr03,1,8)=cu01 AND substr(cr03,9,1)=cu02 AND cu10=na01(+)" & _
           " AND decode(instr(cf10,'@taie.com.tw'),0,cf10,cf14)=st01(+) AND a0901(+)=st03" & _
           " AND cf10<>'QPGMR'"
   stSQL = stSQL & " Union All" & _
           " SELECT decode(cf09,'T','寄','R','收') AS 寄收," & _
           "sqldateT(cf11) AS 寄收日期,sqltime(cf12) AS 寄收時間," & _
           "cf10,st02,a0902,decode(substr(st03,1,2)||st16,'F12','英文組','F14','日文組',decode(st16,'1','英文組','2','日文組',st16)) st16" & _
           ",cr03,NVL(PCU08,Decode(PCU03,null,PCU07,RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06))),na03" & _
           ",cr05,ac03,cf13,cf01" & _
           " From contactfile, contactrecord, allcode, potcustomer, nation, staff, acc090" & _
           " Where cf04 >= " & Val(stDate) + 191100 & "01 And cf04 <= " & Val(stDate) + 191100 & "31 And cf09 Is Not Null" & _
           " AND cf01=cr01 AND ac01='11' AND ac02(+)=cr05" & _
           " AND substr(cr03,1,8)=pcu01 AND substr(cr03,9,1)=pcu02 AND pcu09=na01(+)" & _
           " AND decode(instr(cf10,'@taie.com.tw'),0,cf10,cf14)=st01(+) AND a0901(+)=st03" & _
           " AND cf10<>'QPGMR'"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      rsQuery.MoveFirst
      xlsReport38.SheetsInNewWorkbook = 1 '預設工作表數目
      xlsReport38.Workbooks.add
      xlsReport38.Application.WindowState = xlMaximized
      Set wksReport38 = xlsReport38.Worksheets(1)
      '設定欄位名稱及欄寬
      nRows = 1
      For intQ = 1 To UBound(arrTmp) + 1
          wksReport38.Range(Chr(intQ + 64) & nRows).Value = arrTmp(intQ - 1)
          wksReport38.Range(Chr(intQ + 64) & ":" & Chr(intQ + 64)).ColumnWidth = Val(arrTmpW(intQ - 1))
          wksReport38.Range(Chr(intQ + 64) & nRows).HorizontalAlignment = xlCenter
      Next
      intQ = intQ - 1
      wksReport38.Range("A1:" & Chr(intQ + 64) & "1").Select
      xlsReport38.Selection.Font.Bold = True
      '凍結窗格 Start
      xlsReport38.ActiveSheet.Range("A2").Select
      xlsReport38.ActiveWindow.FreezePanes = True
      '凍結窗格 End
      
      nRows = nRows + 1
      Do While Not rsQuery.EOF
         For intQ = 1 To UBound(arrTmp) + 1
            With wksReport38.Range(Chr(intQ + 64) & nRows)
               .Value = "" & rsQuery.Fields(intQ - 1)
               .NumberFormatLocal = "@" '文字
               wksReport38.Range(Chr(intQ + 64) & nRows).HorizontalAlignment = xlLeft
            End With
         Next intQ
         nRows = nRows + 1
         rsQuery.MoveNext
       Loop
       If Val(xlsReport38.Version) < 12 Then
          xlsReport38.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=-4143
       Else
          xlsReport38.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=56
       End If
       xlsReport38.Workbooks.Close
       xlsReport38.Quit
       
       strTo = Pub_GetSpecMan("兼職業務拓展處人員")
       If strTo <> "" Then
          SendMAPIMail strTo, stDate & strTitle, "同主旨", App.path & TextPath & xlsFileName
       End If
   '若當月沒資料也發mail通知
   Else
       strTo = Pub_GetSpecMan("兼職業務拓展處人員")
       If strTo <> "" Then
          SendMAPIMail strTo, "無資料! " & stDate & strTitle, "同主旨"
       End If
   End If

   Set rsQuery = Nothing
   Set xlsReport38 = Nothing
   Exit Sub
   
ErrHnd:
   
   WLog strTitle & ":" & Err.Description
   If Val(xlsReport38.Version) < 12 Then
      xlsReport38.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=-4143
   Else
      xlsReport38.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=56
   End If
   xlsReport38.Workbooks.Close
   xlsReport38.Quit
   Set xlsReport38 = Nothing
   Set rsQuery = Nothing
End Sub

'Add By Sindy 2025/6/30 財務信箱定期檢查
Private Sub StrMenu39()
Dim xlsReport39 As New Excel.Application
Dim wksReport39 As New Worksheet
Dim intQ As Integer, stSQL As String
Dim rsQuery As New ADODB.Recordset
Dim nRows As Integer, ii As Integer
Dim arrTmp As Variant, arrTmpW As Variant
Dim strYear As String
Dim xlsFileName As String
Dim strCC As String, strBCC As String, strEmp As String
Dim strTitle As String
Dim strShowSheets As String
Dim strLastDay As String
   
On Error GoTo ErrHnd
   
   'Modify By Sindy 2025/10/8 改每月第一工作天由系統提供給智權同仁
'   '定期在每年9月1日發信給當年有開立收據的智權同仁(退休或離職者發信給該區主管)
'   If Mid(strSrvDate(2), 4, 2) <> "09" Then Exit Sub
   
   strYear = Left(strSrvDate(2), 3)
   Call PUB_Frmacc44x0(Me.Name, "2", strYear, strSrvDate(2)) '讀取統計資料
   strBCC = "taieacc@taie.com.tw"
   
   strTitle = "財務信箱定期檢查"
   Call PUB_KillTempFile(Mid(TextPath & "*" & strTitle & ".*", 2))
   xlsFileName = strTitle & ".xls"
   
   stSQL = "select max(wd01) from workday where wd01 between " & Left(strSrvDate(1), 6) & "01 and " & Left(strSrvDate(1), 6) & "31"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      strLastDay = "" & RsTemp.Fields(0)
   End If
   
   stSQL = "select T22,T23" & _
            " From ACCTMP44q0" & _
            " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
            " and (T06||T16||T20 IS NULL or T06||T16||T20='X')" & _
            " and T23 is not null" & _
            " group by T22,T23" & _
            " order by T22,T23"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      RsTemp.MoveFirst
      
      'T05.Form Name
      'T14.UserID
      'T22.業務區
      'T23=智權員編
      'T02=客戶編號 ==> 改用"T29"收據抬據抓出來的客戶編號
      'T15=收據抬頭
      'T17=電話
      'T18=傳真
      'T06=E-Mail(代表)
      'T16=財務信箱
      'T19=會計師姓名
      'T20=會計師信箱
      'T27=會計師事務所
      'T28=會計師電話
      'T29.依收據抬頭讀到的客戶編號
      Do While Not RsTemp.EOF
         strShowSheets = ""
         xlsReport39.Application.WindowState = xlMaximized
         xlsReport39.SheetsInNewWorkbook = 2 '預設工作表數目
         xlsReport39.Workbooks.add
         xlsReport39.Visible = True
         For ii = 1 To 2 '2個工作表
            Set wksReport39 = xlsReport39.Worksheets(ii)
            wksReport39.Activate
            If ii = 1 Then '本所客戶
               xlsReport39.Sheets(ii).Name = "本所客戶"
               '欄位抬頭
               stSQL = "智權員編,智權人員,客戶編號,收據抬頭,電話,財務電話,傳真,財務信箱"
               arrTmp = Split(stSQL, ",")
               stSQL = "10,10,10,30,15,15,15,20,10,10,10,10"
               arrTmpW = Split(stSQL, ",")
               stSQL = "select distinct t23,st02,nvl(t29,t02),t15,t17,'',t18,t16,decode(t06,'X','',t06),t19,t20,t27,t28,T29" & _
                         " From ACCTMP44q0, staff" & _
                         " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                         " and T23='" & RsTemp.Fields("t23") & "' and T23=st01" & _
                         " and (T29<>'T' or T29 is null)" & _
                         " and (T06||T16||T20 IS NULL or T06||T16||T20='X')"
            Else '特殊收據抬頭
               xlsReport39.Sheets(ii).Name = "特殊收據抬頭"
               '欄位抬頭
               stSQL = "智權員編,智權人員,收據抬頭,電話,手機,財務電話,傳真,代表/財務信箱"
               arrTmp = Split(stSQL, ",")
               stSQL = "10,10,30,15,15,15,15,20,10,10"
               arrTmpW = Split(stSQL, ",")
               
               stSQL = "select distinct t23,st02,t15,t17,'','',t18,t16,t19,t20,t27,t28,T29" & _
                         " From ACCTMP44q0, staff" & _
                         " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                         " and T23='" & RsTemp.Fields("t23") & "' and T23=st01" & _
                         " and T29='T'" & _
                         " and T16||T20 IS NULL"
            End If
            '設定欄位名稱及欄寬
            nRows = 1
            For intQ = 1 To UBound(arrTmp) + 1
                wksReport39.Range(Chr(intQ + 64) & nRows).Value = arrTmp(intQ - 1)
                wksReport39.Range(Chr(intQ + 64) & ":" & Chr(intQ + 64)).ColumnWidth = Val(arrTmpW(intQ - 1))
                wksReport39.Range(Chr(intQ + 64) & nRows).HorizontalAlignment = xlCenter
            Next
            intQ = intQ - 1
            wksReport39.Range("A1:" & Chr(intQ + 64) & "1").Select
            xlsReport39.Selection.Font.Bold = True
            '凍結窗格 Start
            xlsReport39.ActiveSheet.Range("A2").Select
            xlsReport39.ActiveWindow.FreezePanes = True
            '凍結窗格 End
            
            intQ = 1
            Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
            If intQ = 1 Then
               strShowSheets = strShowSheets & "," & ii
               rsQuery.MoveFirst
               nRows = nRows + 1
               Do While Not rsQuery.EOF
                  For intQ = 1 To UBound(arrTmp) + 1
                     With wksReport39.Range(Chr(intQ + 64) & nRows)
                        .NumberFormatLocal = "@" '文字
                        .Value = "" & rsQuery.Fields(intQ - 1)
                        wksReport39.Range(Chr(intQ + 64) & nRows).HorizontalAlignment = xlLeft
                        If "" & rsQuery.Fields(intQ - 1) = "" Then
                           .Select
                           '.Style = "輔色4"
                           .Font.ColorIndex = 30
                           .Interior.ColorIndex = 27
                        End If
                     End With
                  Next intQ
                  nRows = nRows + 1
                  rsQuery.MoveNext
               Loop
            End If
         Next ii
         '切換到有資料的工作區
         If InStr(strShowSheets, ",1") > 0 Then
            xlsReport39.Worksheets(1).Activate
         Else
            xlsReport39.Worksheets(2).Activate
         End If
         
         strEmp = RsTemp.Fields("T23") '& GetPrjSalesNM(RsTemp.Fields("T23"))
         If Val(xlsReport39.Version) < 12 Then
            xlsReport39.Workbooks(1).SaveAs FileName:=App.path & TextPath & strEmp & "_" & xlsFileName, FileFormat:=-4143
         Else
            xlsReport39.Workbooks(1).SaveAs FileName:=App.path & TextPath & strEmp & "_" & xlsFileName, FileFormat:=56
         End If
         xlsReport39.Workbooks.Close
         xlsReport39.Quit
         
'         If RsTemp.Fields("T23") = "A3013" Or RsTemp.Fields("T23") = "A9017" Then
'            MsgBox RsTemp.Fields("T23")
'         End If
         'Modify By Sindy 2025/10/8 扣繳的名單仍請依目前(密件副本給瑞婷,分所則改提供該所會計同仁)
         strExc(9) = PUB_GetST06(RsTemp.Fields("T23"))
         If strExc(9) = "2" Then
            strCC = Pub_GetSpecMan("出納人員-中所")
            strExc(10) = strCC
         ElseIf strExc(9) = "3" Then
            strCC = Pub_GetSpecMan("出納人員-南所")
            strExc(10) = strCC
         ElseIf strExc(9) = "4" Then
            strCC = Pub_GetSpecMan("出納人員-高所")
            strExc(10) = strCC
         Else
            strCC = strBCC
            strExc(10) = Left(Pub_GetSpecMan("財務處應收處理人員"), 5)
         End If
         'Modify By Sindy 2025/11/13 瑞婷有與協理討論信內容日期原為每月最後一個工作日,改成每月固定15號
         '   因為現在每月在檢查,不會像一開始數量很多,確實須要時間,協理同意改成每月15日
         '"/" & Right(strLastDay, 2) & " 前回覆 財務處 " => "/15 前回覆 財務處 "
         strExc(8) = "敬啟者：" & vbCrLf & vbCrLf & _
                     "為因應公司客戶每年均有扣繳問題須處理，請提供客戶連絡資訊" & vbCrLf & _
                     "A.若客戶有專屬的財務信箱，請直接提供財務信箱。" & vbCrLf & _
                     "B.若客戶沒有財務信箱，但有其他可連絡的信箱，請提供此可連絡的信箱。" & vbCrLf & _
                     "C.若客戶均沒有信箱，請提供客戶之傳真。" & vbCrLf & _
                     "D.若客戶無信箱也沒有傳真，最低的底限為——市內電話及手機擇一" & vbCrLf & vbCrLf & _
                     "* 代表信箱及客戶巿話/手機請自行更新建檔" & vbCrLf & _
                     "* 財務資訊請填入EXCEL表格內，信件請於 " & Mid(strLastDay, 5, 2) & "/15 前回覆 財務處 " & GetPrjSalesNM(strExc(10))
         SendMAPIMail RsTemp.Fields("T23"), "為核對扣繳使用,敬請連絡客戶提供連絡資訊。(注意!Excel檔中, 有2個工作表)", _
            strExc(8), App.path & TextPath & strEmp & "_" & xlsFileName, True, strCC, strBCC
         
         RsTemp.MoveNext
      Loop
   '若沒資料,也發mail通知財務處
   Else
      If strBCC <> "" Then
         SendMAPIMail strBCC, "無資料! 為核對扣繳使用,敬請連絡客戶提供連絡資訊。", "同主旨", , True
      End If
   End If
   
   Set rsQuery = Nothing
   Set xlsReport39 = Nothing
   Exit Sub
   
ErrHnd:
   
   WLog strTitle & ":" & Err.Description
   If Val(xlsReport39.Version) < 12 Then
      xlsReport39.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=-4143
   Else
      xlsReport39.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName, FileFormat:=56
   End If
   xlsReport39.Workbooks.Close
   xlsReport39.Quit
   Set xlsReport39 = Nothing
   Set rsQuery = Nothing
End Sub

'Added by Lydia 2025/08/?? 不得代理和風險檢查名單提醒---先上傳,不上線
Private Sub StrMenu40()
Dim strBDate As String, strCond As String, strANO As String, strANoList As String
Dim intB As Integer, strB1 As String
Dim intCounter As Integer
Dim strSub As String, strTo As String
Dim rsBD As New ADODB.Recordset
Dim tmpArr As Variant

   '來源1:不得代理NotAgent
   'strB1 = "SELECT '1' AS ord1,nt25 AS cdate,nt35 AS ano,nvl(st93,st03) AS cdept,nvl(a0924,a0908) AS cdeptman,nt18 AS st01,st04 " & _
           "FROM notagent,staff,acc090,acc090new WHERE nvl(nt21,0)=0 AND nt18=st01(+)  AND nvl(nt35,'N')<>'N' AND st03=a0901(+) AND st93=a0921(+) " & _
           "group by nt25,nt35,nvl(st93,st03),nvl(a0924,a0908),nt18,st04 "
   '來源2:風險檢查對象RiskCheckList
   'strB1 = strB1 & "UNION SELECT '2' AS ord1,rcl28 AS cdate,rcl18 AS ano,nvl(st93,st03) AS cdept,nvl(a0924,a0908) as cdeptman,rcl22 AS st01,st04 " & _
           "FROM riskchecklist,staff,acc090,acc090new WHERE nvl(rcl24,0)=0 AND rcl22=st01(+)  AND nvl(rcl18,'N')<>'N' AND st03=a0901(+) AND st93=a0921(+) " & _
           "GROUP BY rcl28,rcl18,nvl(st93,st03),nvl(a0924,a0908),rcl22,st04 "
   '1234
   '來源1:不得代理NotAgent
   strB1 = "SELECT '1' AS ord1,nt25 AS cdate,nt35 AS ano,nt18 AS cuser FROM notagent WHERE nvl(nt21,0)=0 AND nvl(nt35,'N')<>'N' and instr(nt35||',','R')=0 group by nt25,nt35,nt18 "
   '來源2:風險檢查對象RiskCheckList
   strB1 = strB1 & "UNION SELECT '2' AS ord1,rcl28 AS cdate,rcl18 AS ano,rcl22 AS cuser FROM riskchecklist WHERE nvl(rcl24,0)=0 AND nvl(rcl18,'N')<>'N' and instr(rcl18||',','R')=0 GROUP BY rcl28,rcl18,rcl22 "
   strB1 = strB1 & "ORDER BY ord1,cdate,ano "
   intB = 1
   Set rsBD = ClsLawReadRstMsg(intB, strB1)
   If intB = 1 Then
      cnnConnection.Execute "delete from rdatafactory where formname = 'StrMenu40' and id ='QPGMR' "
      rsBD.MoveFirst
      Do While Not rsBD.EOF
         If strSrvDate(1) >= Mid(CompDate(1, 6, "" & rsBD.Fields("cdate")), 1, 6) & "01" Then '不得代理/風險檢查之建檔日+6個月
            tmpArr = Empty
            tmpArr = Split("" & rsBD.Fields("ano"), ",")
            For intB = 0 To UBound(tmpArr)
               If Left(Trim(tmpArr(intB)), 1) = "X" Or Left(Trim(tmpArr(intB)), 1) = "Y" Then
                  strExc(0) = "": strExc(1) = ""
                  strANO = Mid(Trim(tmpArr(intB)) & String(8, "0"), 1, 8)
                  If InStr(strANoList & ",", strANO) = 0 Then
                     If Left(Trim(tmpArr(intB)), 1) = "X" Then
                        strExc(0) = "select cu14 as bdate from customer where cu01='" & strANO & "' and cu02='0' "
                     ElseIf Left(Trim(tmpArr(intB)), 1) = "Y" Then
                        strExc(0) = "select fa11 as bdate from fagent where fa01='" & strANO & "' and fa02='0' "
                     End If
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strBDate = "" & RsTemp.Fields("bdate")
                     End If
                     If Trim(strBDate) = "" Then strBDate = "" & rsBD.Fields("cdate")
                     '將所有基本檔+收文日的資料存在rdatafactory
                     'R001=X/Y編號, R002~R005=PA01~PA04, R006=是否閉卷/銷卷,R007=申請國家, R008=新案收文號, R009=新案收文性質, R010=新案收文日, R011=收文人員, R012=已轉案至他所收文
                     If Left(strANO, 1) = "X" Then
                        '***專利***
                        strSql = "INSERT into rdatafactory (formname,ID,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007) " & _
                                 "SELECT 'StrMenu40' AS formname,'QPGMR' AS ID," & intCounter & " AS seqno,ROWNUM,'" & strANO & "' as ano,pa01,pa02,pa03,pa04,decode(pa57||pa108,NULL,NULL,'Y') AS cstatus,pa09 " & _
                                 "From patent WHERE (pa01 IN ('FCP','P') or pa01||pa04='CFP00') and instr(pa26||pa27||pa28||pa29||pa30,'" & strANO & "') > 0 "
                        cnnConnection.Execute strSql, intI
                        '***商標***
                        strSql = "INSERT into rdatafactory (formname,ID,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007) " & _
                                 "SELECT 'StrMenu40' AS formname,'QPGMR' AS ID," & intCounter & " AS seqno,ROWNUM,'" & strANO & "' as ano,tm01,tm02,tm03,tm04,decode(tm29||tm57,NULL,NULL,'Y') AS cstatus,tm10 " & _
                                 "From trademark WHERE (tm01 IN ('FCT','T') or tm01||tm04='CFT00') and instr(tm23||tm78||tm79||tm80||tm81,'" & strANO & "') > 0 "
                        cnnConnection.Execute strSql, intI
                     Else
                        '***專利***
                        strSql = "INSERT into rdatafactory (formname,ID,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007) " & _
                                 "SELECT 'StrMenu40' AS formname,'QPGMR' AS ID," & intCounter & " AS seqno,ROWNUM,'" & strANO & "' as ano,pa01,pa02,pa03,pa04,decode(pa57||pa108,NULL,NULL,'Y') AS cstatus,pa09 " & _
                                 "FROM patent,caseprogress WHERE (pa01 IN ('FCP','P') OR pa01||pa04='CFP00') AND pa01=cp01(+) AND pa02=cp02(+) AND pa03=cp03(+) AND pa04=cp04(+) AND cp31='Y' " & _
                                 "AND ((pa09='000' AND substr(pa75,1,8)='" & strANO & "') OR (pa09<>'000' AND substr(cp44,1,8)='" & strANO & "') )"
                        cnnConnection.Execute strSql, intI
                        '***商標***
                        strSql = "INSERT into rdatafactory (formname,ID,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007) " & _
                                 "SELECT 'StrMenu40' AS formname,'QPGMR' AS ID," & intCounter & " AS seqno,ROWNUM,'" & strANO & "' as ano,tm01,tm02,tm03,tm04,decode(tm29||tm57,NULL,NULL,'Y') AS cstatus,tm10 " & _
                                 "From trademark,caseprogress WHERE (tm01 IN ('FCT','T') or tm01||tm04='CFT00') AND tm01=cp01(+) AND tm02=cp02(+) AND tm03=cp03(+) AND tm04=cp04(+) AND cp31='Y' " & _
                                 "AND ((tm10='000' AND substr(tm44,1,8)='" & strANO & "') OR (tm10<>'000' AND substr(cp44,1,8)='" & strANO & "') )"
                        cnnConnection.Execute strSql, intI
                     End If
                     '***收文性質,收文日/發文日,是否已轉它所***
                     'P案的台灣案判斷PA75+收文日，非台灣案判斷CP44+發文日---參考frmAutoBatchDay.StrMenu93
                     strSql = "UPDATE rdatafactory A SET A.R008=(SELECT cp09 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              ",A.R009=(SELECT cp10 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              ",A.R010=(SELECT cp05 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              ",A.R011=(SELECT cp13 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              "WHERE A.formname = 'StrMenu40' AND A.ID ='QPGMR' AND A.r001='" & strANO & "' and (r002='FCP' or (r002='P' and r007='000')) "
                     cnnConnection.Execute strSql, intI
                     strSql = "UPDATE rdatafactory A SET A.R008=(SELECT cp09 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              ",A.R009=(SELECT cp10 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              ",A.R010=(SELECT cp27 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              ",A.R011=(SELECT cp13 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              "WHERE A.formname = 'StrMenu40' AND A.ID ='QPGMR' AND A.r001='" & strANO & "' and (r002='CFP' or (r002='P' and r007<>'000')) "
                     cnnConnection.Execute strSql, intI
                     strSql = "UPDATE rdatafactory A SET A.R012=(SELECT count(*) cnt FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp159=0 AND cp01 IN ('FCP','P','CFP') " & _
                              "AND cp10 IN ('929','1929')) WHERE A.formname = 'StrMenu40' AND A.ID ='QPGMR' AND A.r001='" & strANO & "' "
                     cnnConnection.Execute strSql, intI
                     'FCT,T案的判斷收文日，CFT判斷發文日---參考frmAutoBatchDay.StrMenu93
                     strSql = "UPDATE rdatafactory A SET A.R008=(SELECT cp09 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              ",A.R009=(SELECT cp10 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              ",A.R010=(SELECT cp05 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              ",A.R011=(SELECT cp13 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              "WHERE A.formname = 'StrMenu40' AND A.ID ='QPGMR' AND A.r001='" & strANO & "' and r002 in ('FCT','T') "
                     cnnConnection.Execute strSql, intI
                     strSql = "UPDATE rdatafactory A SET A.R008=(SELECT cp09 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              ",A.R009=(SELECT cp10 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              ",A.R010=(SELECT cp27 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              ",A.R011=(SELECT cp13 FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp31='Y' AND cp159=0) " & _
                              "WHERE A.formname = 'StrMenu40' AND A.ID ='QPGMR' AND A.r001='" & strANO & "' and r002 in ('CFT') "
                     cnnConnection.Execute strSql, intI
                     strSql = "UPDATE rdatafactory A SET A.R012=(SELECT count(*) cnt FROM caseprogress WHERE cp01=A.r002 AND cp02=A.r003 AND cp03=A.r004 AND cp04=A.r005 AND cp159=0 AND cp01 IN ('FCT','T','CFT') " & _
                              "AND cp10 IN ('728','1724')) WHERE A.formname = 'StrMenu40' AND A.ID ='QPGMR' AND A.r001='" & strANO & "' "
                     cnnConnection.Execute strSql, intI
                     '***收文性質,收文日/發文日,是否已轉它所***
                     strANoList = strANoList & strANO & ","
                     intCounter = intCounter + 1
                  End If
                  '1234

                  '已轉案至他所
                  strExc(0) = "SELECT count(*) cnt1,sum(decode(nvl(R012,'0'),'0',0,1)) cnt2 FROM rdatafactory WHERE formname = 'StrMenu40' AND ID ='QPGMR' AND r001='" & strANO & "' "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                      '1234
                     If RsTemp.Fields("cnt1") = RsTemp.Fields("cnt2") Then
                        Call StrMenu40_Sub("2", "" & rsBD.Fields("ord1"), strANO, "" & rsBD.Fields("cuser"))
                     End If
                  End If

                  intB = Abs(DateDiff("d", AFDate(strBDate), AFDate("" & rsBD.Fields("cdate"))))
                  If intB >= 365 * 3 Then '3年內未收新案
                     strExc(2) = " and r010>=" & Mid(CompDate(0, -3, strSrvDate(1)), 1, 6) & "01"
                     strExc(5) = "3"
                  Else
                     '1年
                     strExc(2) = " and r010>=" & Mid(CompDate(0, -1, strSrvDate(1)), 1, 6) & "01"
                     strExc(5) = "1"
                  End If
                  strExc(0) = "SELECT R002,R003,R004,R005 FROM rdatafactory WHERE formname='StrMenu40' AND ID='QPGMR' AND r001='" & strANO & "'" & strExc(2) & _
                              " and ((r002 in ('FCP','P','CFP')  and r009 in ('101','102','103')) or (r002 in ('FCT','T','CFT') and r009 in ('101')))"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 0 Then
                      '1234
                      '1年、近3年已無新案委辦
                      Call StrMenu40_Sub(strExc(5), "" & rsBD.Fields("ord1"), strANO, "" & rsBD.Fields("cuser"))
                  End If
            
               End If
            Next intB
         End If
         rsBD.MoveNext
      Loop
   End If
   
   Set rsBD = Nothing
End Sub

'Added by Lydia 2025/08/??
Private Sub StrMenu40_Sub(ByVal pKind As String, ByVal pOrd As String, ByVal pAno As String, ByVal pUserNo As String)
Dim intQ As Integer, strQ1 As String
Dim rsQD As New ADODB.Recordset
Dim strORD As String, strSub As String, strTo As String, strContent As String
Dim strST15 As String, xlsFileName1 As String, intCounter As Integer, strAns2 As String
Dim xlsRpt40 As New Excel.Application
Dim wksRPT40 As New Worksheet

   If pOrd = "1" Then
      strORD = "不得代理"
   ElseIf pOrd = "2" Then
      strORD = "風險檢查"
   End If
   
   If pKind = "1" Or pKind = "3" Then
      strContent = pAno & "此編號設有" & strORD & "名單，" & IIf(pKind = "1", "一年內皆無", "但近三年已無") & "專利(FCP,P,CFP)/商標(FCT,T,CFT)新案委辦，請檢視合約，追蹤聯繫客戶，內部檢視或重新評估是否仍需維持設定。"
   ElseIf pKind = "2" Then
      strContent = pAno & "此編號設有" & strORD & "名單，但底下專利(FCP,P,CFP)/商標(FCT,T,CFT)案件，皆已經轉案至他所，請檢視合約，追蹤聯繫客戶，內部檢視或重新評估是否仍需維持設定。"
   End If
   If strContent = "" Then Exit Sub
   
   strSub = "專利／商標設有" & strORD & "名單之提醒_" & pAno
   strQ1 = "SELECT st01,st02,st15,st04,nvl(st93,st03) AS cdept,nvl(a0924,a0908) cdeptman,st52,st53,st54,st55 " & _
           "From staff, acc090, acc090new where st01='" & pUserNo & "' AND st03=a0901(+) AND st93=a0921(+) "
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      If "" & rsQD.Fields("st04") <> "1" Or "" & rsQD.Fields("cdeptman") = "" & rsQD.Fields("st01") Then
         strTo = "" & rsQD.Fields("cdeptman")
      Else
         strTo = IIf("" & rsQD.Fields("st52") <> "", ";" & rsQD.Fields("st52"), "")
         If InStr(strTo, "" & rsQD.Fields("cdeptman")) = 0 Then strTo = strTo & ";" & rsQD.Fields("cdeptman")
         strTo = pUserNo & strTo
         strST15 = "" & rsQD.Fields("st15")
      End If
   End If
   If strTo = "" Then strTo = Pub_GetSpecMan("程式管理人員")
   
   '附件:續存案件清單
   strQ1 = "select r002||'-'||r003||decode(r004||r005,'000',null,'-'||r004||'-'||r005) as caseno,r002,st15,nvl(a0922,a0902) as deptname " & _
           "from rdatafactory,staff,acc090,acc090new " & _
           "where formname='StrMenu40' and id ='QPGMR' and r001='" & pAno & "' and r006 is null and nvl(r012,'0')='0' " & _
           "and r011=st01(+) AND st03=a0901(+) AND st93=a0921(+) order by 1"
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      xlsFileName1 = strORD & "續存案件清單.xls"
      If Dir(App.path & TextPath & xlsFileName1) <> "" Then
         Kill App.path & TextPath & xlsFileName1
      End If
      rsQD.MoveFirst
      xlsRpt40.SheetsInNewWorkbook = 1 '預設工作表數目
      xlsRpt40.Workbooks.add
      xlsRpt40.Application.WindowState = xlMinimized
      xlsRpt40.Application.Visible = False
      Set wksRPT40 = xlsRpt40.Worksheets(1)
      wksRPT40.Range("A1").Value = "本所案號"
      wksRPT40.Columns("A:A").ColumnWidth = 12
      wksRPT40.Range("A1").HorizontalAlignment = xlCenter
      intCounter = 2
      
      Do While Not rsQD.EOF
         wksRPT40.Range("A" & intCounter).Value = "" & rsQD.Fields("caseno")
         If strST15 <> "" & rsQD.Fields("st15") Then
            If InStr(strAns2 & "、", "" & rsQD.Fields("deptname")) = 0 Then
               strAns2 = strAns2 & "、" & rsQD.Fields("deptname")
            End If
         End If
         intCounter = intCounter + 1
         rsQD.MoveNext
      Loop
      xlsRpt40.Sheets(1).Select '選擇工作表
      If Val(xlsRpt40.Version) < 12 Then
         xlsRpt40.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName1, FileFormat:=-4143
      Else
         xlsRpt40.Workbooks(1).SaveAs FileName:=App.path & TextPath & xlsFileName1, FileFormat:=56
      End If
      xlsRpt40.Workbooks.Close
      xlsRpt40.Quit
      Set xlsRpt40 = Nothing
      Set wksRPT40 = Nothing
   End If
   If strTo <> "" Then
      strTo = "A3034"
      SendMAPIMail strTo, strSub, vbCrLf & vbCrLf & strContent & IIf(strAns2 = "", "", vbCrLf & vbCrLf & "有其他案件之部門：" & Mid(strAns2, 2) & vbCrLf), IIf(xlsFileName1 <> "", App.path & TextPath & xlsFileName1, "")
   End If

End Sub

'Added by Lydia 2025/09/16 批次閉卷T台灣案陳述意見書的勝訴
Private Sub StrMenu41()
Dim strR1 As String, intR As Integer, bolConn As Boolean
Dim rsRD As New ADODB.Recordset
   
On Error GoTo ErrHandle

   'T的陳述意見書(210)之CP24=1且CP25小於等於系統日-2年未閉卷者上閉卷。
   strR1 = "SELECT cp01,cp02,cp03,cp04,cp05,cp24,cp25 " & _
           "From caseprogress, trademark " & _
           "WHERE cp01='T' AND cp10='210' AND cp24='1' AND nvl(cp25,99999999)<=to_char(sysdate,'yyyymmdd')-20000 " & _
           "AND cp01=tm01(+) AND cp02=tm02(+) AND cp03=tm03(+) AND cp04=tm04(+) AND tm10='000' AND tm29 IS NULL " & _
           "order by cp25 "
   intR = 1
   Set rsRD = ClsLawReadRstMsg(intR, strR1)
   If intR = 1 Then
      rsRD.MoveFirst
      cnnConnection.BeginTrans
      bolConn = True
      Do While Not rsRD.EOF
         strSql = "Update Trademark set tm29='Y', tm30=" & strSrvDate(1) & ", tm31='99', tm58=tm58||';陳述意見書勝訴' where tm01='" & rsRD.Fields("cp01") & "' and tm02='" & rsRD.Fields("cp02") & "' and tm03='" & rsRD.Fields("cp03") & "' and tm04='" & rsRD.Fields("cp04") & "' "
         Pub_SeekTbLog strSql, "QPGMR", , , "每月批次(frmAutoBatch)"
         cnnConnection.Execute strSql
         rsRD.MoveNext
      Loop
      cnnConnection.CommitTrans
      bolConn = False
   End If
   Set rsRD = Nothing
   
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      If bolConn = True Then cnnConnection.RollbackTrans
      WLog Err.Description
   End If
   Set rsRD = Nothing

End Sub
