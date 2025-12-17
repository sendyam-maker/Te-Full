Attribute VB_Name = "basConst"
'Memo By Sonia 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Public Pub_strLogText As String 'Add By Sindy 2016/5/11 要抓莫名其妙的Bug
Public Const strAccMailBox As String = "account@taie.com.tw" '財務寄信帳號 Added by Morgan 2011/10/12
'Add by Morgan 2011/9/14 常數整合
Public Const lngWidth As Long = 9600
Public Const lngHeight As Long = 6000
Public Const intSleep As Integer = 10
Public Const ADFormat As String = "####/##/##"
Public Const DFormat As String = "###/##/##"
Public Const Tformat As String = "##:##:##"
Public Const FDollar As String = "###,###,###,###.00"
Public Const DDollar As String = "###,###,###,###"
Public Const DDollar2 As String = "###,###,###,##0" 'Added by Morgan 2011/12/19
Public Const FAmount As String = "0.00"
Public Const DAmount As String = "0"
Public Const strPercent As String = "###.00"
Public pub_SaveCoRec As Boolean 'Add By Sindy 2022/6/17 記錄是否有儲存往來記錄

Public Const DriveShield = "x:"
'Public Const MDBPathOut = "x:\taie\mdb\分所下載"
'Public Const MDBPathIn = "x:\taie\mdb\分所上傳"
'**************************************************
'900905 修改，根據新文件，於90年09月05日，與邱秀玲和薛德璟共同討論並核對無誤，予以修改。
Public Const DOCPathOut = DriveShield & "\taie\文件檔案\文件"
Public Const DOCPathIn = DriveShield & "\taie\文件檔案\分所上傳"
Public Const SMPPath = DriveShield & "\taie\範本"
Public Const DocTempPath = "c:\windows\temp"    '申請書暫放檔  不存檔
Public Const DOCLOGPath = DriveShield & "\taie\文件檔案\LOG"   '有問題
'**************************************************
'權限檢核變數
Public Const strAdd As String = "A"
Public Const strEdit As String = "E"
Public Const strDel As String = "D"
Public Const strFind As String = "F"
Public Const strPrint As String = "P"
Public Const strExec As String = "X"
Public Const strCrossDept As String = "Y"  '20080926 add by Toni 跨部門權限檢核變數
Public Const strBranch As String = "B"  'Add By Sindy 2010/12/30 跨所別權限檢核變數

'Public Const strAdoConnect As String = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=M51CON;Persist Security Info=True" 'Removed by Morgan 2017/4/20 沒用了
Public Const strPicPath As String = "c:\pics\"
Public Const strRptPath As String = "c:\work\vb60\rpt\taie\"
Public strExcelPath As String, strExcelPathN As String 'Modify by Amy 2021/06/21 原:Public Const strExcelPath As String ="c:\xls\" 避免居家刪到別人的檔案
'Modified by Lydia 2024/07/22 改成變數
'Public Const strDocImportPath As String = "\\Pat1\OA_SCAN" '公文來函電子檔匯入預設路徑 Added by Morgan 2014/5/20
'Public Const strTFeeForm As String = "\\SALE1\XFER\FEE_FORM" 'T案繳費單存放路徑 Added by Morgan 2018/11/30
'Public Const strTApp1CasePath As String = "\\Sale1\商標客戶專區" 'T案商標客戶專區路徑 Add By Sindy 2023/3/28
Public strDocImportPath As String
Public strTFeeForm As String
Public strTApp1CasePath As String
Public strTyping2Path As String 'Added by Lydia 2024/07/22 Typing2的DNS名稱
Public strSale1Path As String 'Added by Lydia 2024/07/22 SALE1的DNS名稱
Public strPat1Path As String 'Added by Lydia 2024/07/22 PAT1的DNS名稱
'end 2024/07/22
Public Const strT000Sale1CPMList = "101,102,103,201,206,208,211,301,302,303,304,306,307,308,309,310,313,501,502,503,504,505,506,507,717,725,729"
Public Const strTMCppFilePath As String = "c:\TM" 'TM信件程序人員從信件切檔案出來,讓系統自動歸入卷宗區的存放路徑 Add by Sindy 2019/5/7
'Modify By Sindy 2022/10/25
'Modified by Lydia 2024/07/22 改成變數
'Public Const str_T_OrderPath = "\\SALE1\tm_order_scan"
'Public Const str_FCT_OrderPath = "\\SALE1\FCT_Order_SCAN"
'Public Const str_ACS_OrderPath = "\\SALE1\ACS01_Order_SCAN"
'Public Const str_P_OrderPath = "\\Pat1\Order_SCAN"
'Public Const str_CFP_OrderPath = "\\Pat1\CFP_Order_SCAN"定
'Public Const str_P_台灣電子送件檔案路徑 = "\\PAT1\Te資料暫存區\(勿刪)台灣電子送件檔案"
'2022/10/25 END
Public str_T_OrderPath As String
Public str_FCT_OrderPath As String
Public str_ACS_OrderPath As String
Public str_P_OrderPath As String
Public str_CFP_OrderPath As String
Public str_P_台灣電子送件檔案路徑 As String
'end 2024/07/22
Public Const strIcoPath As String = strPicPath & "bmw.ico"
Public Const strBackPicPath1 As String = strPicPath & "background014.jpg"
Public Const strBackPicPath2 As String = strPicPath & "background015.jpg"
Public Const strBackPicPath3 As String = strPicPath & "background016.jpg"
Public Const strBackPicPath4 As String = strPicPath & "background040.jpg"
Public Const strBackPicPath5 As String = strPicPath & "background030.jpg"
Public Const strBackPicPath6 As String = strPicPath & "background019.jpg"
Public Const strBackPicPath8 As String = strPicPath & "background090.jpg"
Public Const intFontSize As Integer = 12
Public Const intMax As Integer = 50
'end 2011/9/14

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

'Add By Sindy 2017/7/6
'Modify By Sindy 2024/5/15 + ,'05','LAbackup'
Public Const MRL01CName As String = "decode(MRL01,'01','IPDept_inbound','02','IPDept_backup','03','Patent','04','TM','05','LAbackup',MRL01)"
Public Const MRL09CName As String = "decode(MRL09,'Y','執行中','E','已完成','F','失敗','B','中斷','A','人工啟動',MRL09)"
Public Const MRL01CName2 As String = "01.IPDept_inbound 02.IPDept_backup 03.Patent 04.TM 05.LAbackup"

'Add By Sindy 2022/4/6
Public Const WM_全型符號表 As String = "，、。！：；˙‥‧‵〃〝〞﹁﹂﹃﹒＃＄％＆＊．？＠～…（）＜＞｛｝〈〉《》「」『』【】〔〕﹙﹚﹛﹜﹝﹞﹤﹥"
'Add By Sindy 2022/4/18
Public Const WM_半型符號表 As String = ",./\;:'""~!@#$%^&*(){}[]<>"

'Add By Sindy 2021/6/28
Public Const 國外潛在客戶類別 As String = "'1','廠商','2','事務所','3','個人','4','平台','5','供應商','6','媒體','7','協會','8','其他'"

'Add By Sindy 2017/3/23
Public Const IPDept收件匣 = "01.國外部IPDept收信郵件"
Public Const IPDept寄件匣 = "02.國外部IPDept寄信郵件"
Public Const Patent收件匣 = "03.專利處Patent收信郵件"
Public Const TM收件匣 = "04.商標處TM收信郵件"
Public Const LAbackup寄件匣 = "05.法律所寄信郵件" 'Add By Sindy 2024/5/14
Public Const 國外部收件信箱 As String = "inbound@taie.com.tw" '"inbound"
Public Const 國外部寄件信箱 As String = "backup@taie.com.tw" '"backup"
Public Const 專利處收件信箱 As String = "PATENT@taie.com.tw" '"PATENT" '專利處信箱(patent)
Public Const 商標處收件信箱 As String = "tm@taie.com.tw" '"商標處公用信箱(tm)"
Public Const 法律所寄件信箱 As String = "LAbackup@taie.com.tw" 'Add By Sindy 2024/5/14
'Modified by Morgan 2020/8/13 +7已確認, 8退回2
Public Const 信件處理狀態 As String = "'1','輸入','2','不處理','3','退回','4','歸卷','5','已處理','6','待歸檔','7','已確認','8','退回2','9','回信'"
'2017/3/23 END
Public Const 外專信件處理結果 As String = "'1','舊案收文','2','多案收文','3','新案命名','4','客戶提供文件','5','不續辦或閉卷','6','往來記錄','7','承辦作業','8','不處理','9','輸入','10','回信','11','歸卷','12','已處理'" 'Add By Sindy 2022/6/22
Public Const OL_PatMailCC = "79075" 'Patent信箱的副本收受者 Add By Sindy 2019/7/17
'Modify By Sindy 2021/11/12 取消林純貞不增加人
'Public Const OL_TmMailCC = "98020;69008"
Public Const OL_TmMailCC = "98020" 'TM信箱的副本收受者 Add By Sindy 2019/7/17
Public Const OL_TmMail需排除的收受者 = "96029,96030,TM4,MCTM,P2001,P2002,P2003,P2004" 'TM信箱需排除的收受者 Add By Sindy 2020/11/4

Public Const 客戶編號 = "X"
Public Const 代理人編號 = "Y"

Public Const 接洽記錄單 = "A"
Public Const 內部收文 = "B"
Public Const 主管機關來函 = "C"
Public Const 政府機關來函 = "D"

'TF為馬德里案
Public Const 馬德里案 = "TF"
'TC為內商著作權
Public Const 內商著作權 = "TC"

Public Const 大陸國家代號 = "020"
Public Const 台灣國家代號 = "000"
Public Const 美國國家代號 = "101"

Public Const 專利 = 1
Public Const 商標 = 2
Public Const 法務 = 3
Public Const 顧問 = 4

Public Const 國內 = 0
Public Const 國外_CF = 1
Public Const 國外_FC = 2

'商標案件性質
Public Const 申請 = "101"
Public Const 移轉 = "501"
Public Const 異議 = "601"
Public Const 異議答辯_商 = "602"
Public Const 評定 = "603"
Public Const 評定答辯_商 = "604"
Public Const 廢止 = "605"
Public Const 廢止答辯_商 = "606"

'專利案件性質

Public Const 發明申請 = "101"
Public Const 新型申請 = "102"
Public Const 設計申請 = "103"
Public Const 追加申請 = "104"
Public Const 聯合申請 = "105"
Public Const 主張優先權 = "106"
Public Const 答辯 = "107"
Public Const 申請寄存 = "108"
Public Const PCT申請 = "109"
Public Const 記錄請求_標準專利 = "110"
Public Const 批准記錄請求_標準專利 = "111"
Public Const 短期專利申請 = "112"
Public Const CIP申請 = "113"
Public Const CPA申請 = "114"
Public Const 再發行 = "115"
'Add By Cheng 2002/07/30
Public Const 美國暫時申請 = "118"
'Add By Sindy 2012/5/31
Public Const 植物新品種保護 = "120"
Public Const CA申請 = "122" 'Added by Morgan 2018/3/31
Public Const 衍生設計 = "125" 'Added by Morgan 2012/10/8
Public Const 翻譯 = "201"
Public Const 補文件 = "202"
Public Const 主動修正 = "203"
Public Const 修正 = "204"
Public Const 申復 = "205"
Public Const 補充說明 = "206"
Public Const 提供前案資料 = "207"
Public Const 選取 = "208"
'Add By Cheng 2002/12/11
Public Const 檢視中說 = "209"
Public Const 核對中說格式 = "235"
Public Const 製作中說 = "210"
'Add By Cheng 2002/03/05
Public Const 準備程序 = "211"
Public Const 言詞辯論 = "212"
'Add By Cheng 2002/03/07
Public Const 公開費 = "217"

Public Const 改請發明 = "301"
Public Const 改請新型 = "302"
Public Const 改請設計 = "303"
Public Const 改請追加 = "304"
Public Const 改請聯合 = "305"
Public Const 改請獨立 = "306"
Public Const 分割 = "307"
Public Const 改請衍生設計 = "308" 'Added by Morgan 2012/12/20
Public Const 改請部分設計 = "309" 'Added by Morgan 2013/1/14
Public Const 變更 = "401"
Public Const 更正 = "402"
Public Const 更改 = "403"
Public Const 延期 = "404"
Public Const 申請優先權證明 = "405"
Public Const 申請英文證明 = "406"
Public Const 請求面詢 = "407"
Public Const 面詢 = "408"
Public Const 請求閱卷 = "409"
Public Const 閱卷 = "410"
Public Const 催審 = "411"
Public Const 延緩公告 = "412"
Public Const 自請撤回 = "413"
Public Const 申請復活 = "414"
Public Const 專利權延長 = "415"
Public Const 實體審查 = "416"
Public Const 提早公開 = "417"
Public Const 請求公告 = "418"
Public Const 訴願 = "501"
Public Const 再訴願 = "502"
Public Const 行政訴訟 = "503"
Public Const 行政再審 = "504"
Public Const 行政訴訟上訴 = "507"
'Add By Cheng 2002/12/28
Public Const 參加訴訟 = "506"
Public Const 領證及繳年費 = "601"
Public Const 加註追加 = "602"
Public Const 加註聯合 = "603"
Public Const 補換發證書 = "604"
Public Const 年費 = "605"
Public Const 維持費 = "606"
Public Const 延展費 = "607"
'Add By Cheng 2002/07/03
Public Const 加註專用權延長 = "608"

Public Const 讓與 = "701"
Public Const 合併 = "702"
Public Const 繼承 = "703"
Public Const 授權 = "704"
Public Const 終止授權 = "705"
Public Const 設定質權 = "706"
Public Const 終止設定質權 = "707"
Public Const 專利權讓與 = "708"
Public Const 異議_專 = "801"
Public Const 異議答辯 = "802"
Public Const 舉發 = "803"
Public Const 舉發答辯 = "804"
Public Const 復審 = "805"
Public Const 告知代理人 = "901"
Public Const 回覆代理人 = "902"
Public Const 專利調查 = "903"
Public Const 調卷 = "904"
Public Const 列印專利資料 = "905"
Public Const 鑑定報告 = "906"
Public Const 不續辦 = "907"
Public Const 退費 = "908"
Public Const 後金 = "909"
Public Const 其他 = "910"
Public Const 補收款 = "911"
Public Const 收達 = "997"
Public Const 提申 = "998"
Public Const 公開 = "999"
Public Const 減免退費 = "919"

'專利主管機關來函案件性質
Public Const 核准 = "1001"
Public Const 核駁 = "1002"
Public Const 通知補文件 = "1003"
Public Const 延期受理 = "1004"
Public Const 通知申請案號 = "1101"
Public Const 通知申請日 = "1102" 'Add by Morgan 2016/5/30
Public Const 通知修正 = "1201"
'Public Const 通知申復 = "1202" 'Mark by Lydia 2025/03/18 不再使用常數
Public Const 通知補充說明 = "1203"
Public Const 通知實審日 = "1204"
Public Const 通知提供前案 = "1205"
Public Const 通知要求選取 = "1206"
Public Const 通知公開 = "1207"
Public Const 通知公告 = "1208"
Public Const 檢索報告 = "1209"
Public Const 通知改請發明 = "1301"
Public Const 通知改請新型 = "1302"
Public Const 通知改請設計 = "1303"
Public Const 通知改請追加 = "1304"
Public Const 通知改請聯合 = "1305"
Public Const 通知改請獨立 = "1306"
Public Const 通知分割 = "1307"
Public Const 通知面詢 = "1401"
Public Const 通知閱卷 = "1402"
Public Const 通知變更 = "1403"
Public Const 延長審查時間 = "1501"
Public Const 撤銷原處分 = "1502"
Public Const 改變原處分 = "1503"
Public Const 通知參加訴願 = "1504"
Public Const 通知參加訴訟 = "1505"
Public Const 通知智慧局答辯函 = "1506"
'Add By Cheng 2002/10/30
Public Const 通知行政上訴答辯 = "1507"

Public Const 通知領證 = "1601"
Public Const 通知證書號數 = "1602"
Public Const 專利證書 = "1603"
Public Const 專利權消滅 = "1604"
Public Const 被異議理由 = "1801"
Public Const 被舉發理由 = "1802"
Public Const 爭議受理 = "1803"
Public Const 對方延期 = "1804"
Public Const 發回補理由 = "1805"
Public Const 發回補答辯 = "1806"
Public Const 對方補充說明 = "1807"
Public Const 對方撤回 = "1808"
Public Const 通知退費 = "1901"
Public Const 其他來函 = "1902"
Public Const 所外鑑定報告結果 = "1903"
'Add By Cheng 2002/05/29
Public Const 通知審查中 = "1905"
Public Const 准予延緩公告 = "1906"
'Add By Cheng 2002/02/20
Public Const 通知退證註銷 = "1907"

'顧問案件性質
Public Const 顧問聘任 = "0"
Public Const 通知開庭 = "9001"

'分案用－CFP指定國家
Public Const EPC指定國家 = "221"

'國內外案件資料維護資料
Public Const 國內外案件 = "0"
'美國IDS資料
Public Const 美國IDS = "1"

'年費起算日
Public Const 收文日 = 1
Public Const 申請日 = 2
Public Const 發文日 = 3
Public Const 准駁日 = 4
Public Const 公告日 = 5
Public Const 發證日 = 6
Public Const 公開日 = 7


'會計科目代號所代表之各系統別
Public Const AcctgL1 = "414101"  'L
Public Const AcctgL2 = "418101"  'L 律師
Public Const AcctgLAT = "410102" 'T 類的顧問聘任
Public Const AcctgLAP = "411102" 'P 類的顧問聘任
Public Const AcctgLA1 = "414102"  '其他 類的顧問聘任
Public Const AcctgLA2 = "418102"
Public Const AcctgFCL = "416101" 'FCL 類的顧問聘任
Public Const AcctgCFL = "416102" 'CFL 類的顧問聘任
'刑事
Public Const 告訴_發明 = "2101"
Public Const 告訴_新型 = "2102"
Public Const 告訴_設計 = "2103"
Public Const 告訴_著作權 = "2104"
Public Const 告訴_商標權 = "2105"
Public Const 告訴_服務標章 = "2106"
Public Const 告訴_其他 = "2108"
Public Const 告訴_補充告訴 = "2109"
Public Const 地檢_答辯 = "2121"
Public Const 地檢_辯護意旨 = "2122"
Public Const 法院_答辯 = "2123"
Public Const 法院_辯護意旨 = "2124"
Public Const 告訴_再審 = "2139"
Public Const 告訴_非常上訴 = "2140"
Public Const 自訴_發明 = "2201"
Public Const 自訴_新型 = "2202"
Public Const 自訴_設計 = "2203"
Public Const 自訴_著作權 = "2204"
Public Const 自訴_商標權 = "2205"
Public Const 自訴_服務標章 = "2206"
Public Const 自訴_其他 = "2207"
Public Const 自訴_補充自訴 = "2208"
Public Const 自訴_答辯 = "2221"
Public Const 自訴_辯護意旨 = "2222"
Public Const 自訴_補充上訴 = "2235"
Public Const 自訴_再審 = "2236"
Public Const 自訴_非常上訴 = "2237"
'民事
Public Const 起訴_給付 = "1111"
Public Const 起訴_確認 = "1112"
Public Const 起訴_形成 = "1113"
Public Const 一審_答辯 = "1115"
Public Const 二審_上訴理由 = "1122"
Public Const 二審_答辯 = "1124"
Public Const 三審_上訴理由 = "1132"
Public Const 三審_再審 = "1133"
Public Const 三審_答辯 = "1134"
'強制
Public Const 強制執行聲請 = "1301"
Public Const 假執行聲請 = "1304"
Public Const 動產_假扣押 = "1311"
Public Const 動產_假處分 = "1312"
Public Const 不動產_假扣押 = "1313"
Public Const 不動產_假處分 = "1314"

'Add by Morgan 2005/8/24
'Modify by Morgan 2008/11/18 改21天
'Public Const 預估公告天數 As Integer = 30
Public Const 預估公告天數 As Integer = 21
'Add by Morgan 2008/4/29 含當天
Public Const 期限通知天數 As Integer = 3
'Add by Morgan 2008/4/29 含當天
Public Const 報價確認天數 As Integer = 3

'Add By Sindy 2022/5/11
Public m_strContactSheetA4 As String '記錄外專簡易聯絡單內容
'2022/5/11 END

'Modify By Sindy 2024/9/26 改為常變數
'C類只取 被異議, 被異議(理由), 被評定, 被評定(理由), 被廢止, 被廢止(理由), 通知參加訴願, 通知參加訴訟, 通知言詞辯論, 通知準備程序, 通知行政上訴答辯
'2007/5/29 ADD BY SONIA 增1205部分核駁,1609對方補充理由,1612發回補理由,1613發回補答辯,1616通知復審答辯,1618對方答辯,1619被部分廢止,1620被部分廢止(理由),1621對造分割
'Modify By Sindy 2024/9/26 增1623被部分異議, 1624 被部分異議 (理由), 1625 被部分評定, 1626 被部分評定 (理由)
Public Const 商爭審查來函案件性質 = "1601,1602,1603,1604,1605,1606,1404,1405,1203,1204,1406" & _
                                   ",1205,1609,1612,1613,1616,1618,1619,1620,1621" & _
                                   ",1623,1624,1625,1626"
'2024/9/26 END

'Add by Morgan 2006/4/14 可建國內外案的案件性質
'Modify by Morgan 2006/4/14 加CIP申請(113)
'Modify by Morgan 2006/5/5 加CA申請(122)
'2007/8/31 modify by SONIA 加再發行(115),美國暫時申請(118)
'Modified by Morgan 2012/10/8 +衍生設計(125)
Public Const CaseMapIn As String = "101,102,103,104,105,109,110,112,113,114,115,118,122,125,201,307"
Public Const CaseMapOut As String = "101,102,103,104,105,109,110,112,113,114,115,118,122,201,307"
'Add by Morgan 2009/8/12 相同案要更新期限的案件性質
'Modify by Morgan 2010/3/29 取消集體設計105(管制設計申請103就好)
Public Const SameCaseProperty4Update As String = "101,102,103,104,109,112,118"
'Add by Morgan 2010/2/9 來函准駁要回寫基本檔的案件性質
'Modified by Morgan 2012/10/8 +衍生設計(125)
'Modified by Morgan 2012/12/20 +改請衍生設計(308)
'Modified by Morgan 2013/1/14 +改請部分設計(309)
'Modified by Morgan 2025/8/7 +植物新品種保護(120)
Public Const UpdateCaseResultCP10List As String = "101,102,103,104,105,107,113,114,120,122,125,301,302,303,304,305,306,307,308,309,424,503,504,802,804"

'Add By Sindy 2023/11/30 FMP非寰華案,有開放FCP程序人員操作發文的案件性質
'Modified by Lydia 2024/06/07 +待客戶最終指示970
Public Const FMPtoFCPSendCasePtyList As String = "901,902,903,904,924,927,937,949,969,970"

'Add by Morgan 2008/4/17 CFP統計給案數的案件性質--慧汶
'Modify by Morgan 2008/8/26 改為專利共用的新案案件性質--秀玲
'Public Const NewCasePtyList As String = "101,102,103,104,113,114,118,122,307"
'Modified by Morgan 2012/10/8 +衍生設計(125)
Public Const NewCasePtyList As String = "101,102,103,104,105,109,110,112,113,114,115,118,120,122,125,307"
'Added by Morgan 2021/9/2 CFP要輸入約定期限的案件性質
'Modified by Morgan 2021/10/20 +501訴願
'Modified by Morgan 2022/10/7 +204修正,206 補充說明,207 提供前案資料,401 面詢,802 異議答辯,804 舉發答辯--陳玫音
Public Const CFPAppDatePtyList As String = "107,204,206,207,208,218,401,421,424,438,501,802,804"

'Add by Morgan 2008/9/3 FCP非以收文日計算的案件性質
Public Const SkipCasePtyList As String = "209,210"
'Add by Morgan 2006/5/29 非英語系國家
'Modify by Morgan 2006/11/3 加德國231--甄妮
'Modify by Morgan 2007/8/24 加越南042--郭
'Modify by Morgan 2009/4/16 改抓國家檔設定(NA59)
'Public Const NoneEngCountry As String = "012,203,205,209,226,217,211,204,222,218,019,231,042"
Public NoneEngCountry As String

'Added by Morgan 2023/3/7 UP會員國
'Modified by Morgan 2024/9/30 +228羅馬尼亞
Public Const UPMember As String = "203,204,206,207,208,209,213,214,216,217,226,228,231,232,236,240,241,242"

'Add by Morgan 2006/6/27 多國案主案設定順序
Public Const MultiCountryPriority As String = "101,011,231,201,012"

'add by sonia 2015/8/24 專利案核准1001時,顯示名稱核准改為核發的案件性質
'Modified by Morgan 2016/9/2 +426新穎性調查--玲玲
'Modified by Morgan 2017/3/8 +436申請優先權證明存取碼
'modify by sonia 2021/9/13 +443申請紙本專利證書
Public Const Patent1001Display As String = "405,406,421,423,426,436,604,807,443"

'Added by Morgan 2021/10/4 專利OA來函性質
Public Const PatentOAPtyList As String = "1002,1006,1202,1209,1206,1220,1227,1221,1810,1802,1807"

'Added by Morgan 2022/12/21
'專利有證書的案件性質
Public Const PACertPtyList As String = "601,604,708,703"
'商標有證書的案件性質
Public Const TMCertPtyList As String = "717,729,103,308"
'end 2022/12/21

'Add by Morgan 2008/10/2 剔除下一程序為程序管控之案件性質語法
'MODIFY BY SONIA 2010/10/14 專利加994順稿
'MODIFY BY SONIA 2015/5/28 商標加1711通知使用宣誓

'***2011/5/12 加註:修改案件性質時,下一程序維護frm075007_2的txtValidate也要改
'Memo by Morgan 2011/6/10 修改案件性質時 frm100123.doQuery 也要改
'Memo by Lydia 2016/10/20 修改商標案件性質時,trigger的NATION_BEFORE也要改
'***請注意未剔除NP06是否續辦條件
'Modified by Lydia 2016/09/07 TC案增加994陸代申請書
'Modified by Lydia 2016/09/12 商標案+註冊證1701
'Modified by Lydia 2022/08/10 CFT商標案+1101通知申請案號
Public Const TMnp07NotIn As String = "'994','997','998','995','996','999','305','1403','312','1701','1711','1101'" 'Add By Sindy 2018/4/19
'Modified by Morgan 2025/9/10+1238
Public Const PAnp07NotIn As String = "'997','998','994','995','996','999','411','1204','1209','1238','1503','1603'" 'Add By Sindy 2018/4/23
'Modify By Sindy 2018/4/19 改用全域變數(TMnp07NotIn)
'modify by sonia 2019/7/30 +ACS系統類別(法務)
Public Const strNpSqlOfNoSalesDuty As String = _
   " and not (np02 in ('L','FCL','CFL','LA','LIN','ACS') and np07='6001')" & _
   " and not (np02 in ('P','PS','CFP','CPS','FCP','FG') and np07 IN (" & PAnp07NotIn & "))" & _
   " and not (np02 NOT in ('L','FCL','CFL','LA','P','PS','CFP','CPS','FCP','FG','LIN','ACS') and np07 IN (" & TMnp07NotIn & "))"

'宣告欄位內容結構
Public Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type

Public Const strFMPNum As String = "F4102"

'諾華公司(商標案)
Public Const strTmNovartisCust As String = "'X45799','X61957','X63125','X63977','X62688','X48639','X47117','X45867','X57979','X60758','X30910','X34106','X52580','Y52795','Y45799','Y52781','Y49575','Y53205','Y51936','Y52347','Y52212','Y28343','Y52341','Y53057','Y52212'"
'add by sonia 2015/11/24 旅狐國際及部分關係企業(商標案)
Public Const strTmTRAVEL_FOXCust As String = "'X0700006','X0700008','X0700015','X0700016','X0700019'"

'Modify By Sindy 2024/7/8  鑑定報告 => 比對分析報告
Public Const P核判分類 As String = "'A','A 發明','P','P 新型','B','B 設計','C','C 再審查','D','D 救濟案','E','E 爭議案','F','F 比對分析報告','G','G 申請案','H','H 設計案','I','I 答辯','J','J 爭議救濟','K','K 比對分析報告','L','L 對外信函','M','M 對內信函'"
Public Const T核判分類 As String = "'A','A 商申申請','B','B 商申延期','C','C 商爭智慧財產局','D','D 商爭經濟部','E','E CMT變更','F','F 商申分析','G','G 商爭分析','H','H CMT申請'"
Public Const CF核判分類 As String = "'A','A 申請','B','B 非申請','C','C 審定爭議來函','D','D 其他來函','E','E 提申'" 'Add Sindy 2024/9/16

'電子承辦單簽辦流程狀態
Public Const EMP_聯絡 = "00"
Public Const EMP_草完 = "01"
Public Const EMP_送標號 = "02"
Public Const EMP_標號 = "03"
Public Const EMP_送核 = "04"
Public Const EMP_送英核 = "05"
Public Const EMP_核修 = "06"
Public Const EMP_核完 = "07"
Public Const EMP_送會 = "08"
Public Const EMP_會修 = "09"
Public Const EMP_會完 = "10"
Public Const EMP_上墨 = "11"
Public Const EMP_墨完 = "12"
Public Const EMP_繪圖判發 = "13"
Public Const EMP_送判 = "14"
Public Const EMP_判發 = "15"
Public Const EMP_退回 = "16"
Public Const EMP_轉回 = "17" 'Add By Sindy 2013/8/22
Public Const EMP_退件 = "18" 'Add By Sindy 2013/9/10
Public Const EMP_退件重送 = "19" 'Add By Sindy 2013/9/10
Public Const EMP_不自動更新會完日 = "20" 'Add By Sindy 2013/9/17
Public Const EMP_附加流程 = "21" 'Add By Sindy 2013/9/23
Public Const EMP_系統上墨 = "22" 'Add By Sindy 2013/11/14
Public Const EMP_草核 = "23" 'Add By Sindy 2015/4/20
Public Const EMP_草修 = "24" 'Add By Sindy 2015/4/20
Public Const EMP_草核完 = "25" 'Add By Sindy 2015/4/20
'Add By Sindy 2016/2/23
Public Const EMP_會圖 = "26" 'Modify By Sindy 2019/1/18 改名稱為”會(圖/文)”
Public Const EMP_圖修 = "27" 'Modify By Sindy 2019/1/18 改名稱為”(圖/文)修”
Public Const EMP_圖完 = "28" 'Modify By Sindy 2019/1/18 改名稱為”(圖/文)完”
Public Const EMP_會完重修 = "29"
Public Const EMP_修改圖式 = "30"
Public Const EMP_不自動更新齊備日 = "31"
'2016/2/23 END
Public Const EMP_准許先會 = "32" 'Add By Sindy 2016/12/21
Public Const EMP_送件 = "33" 'Add By Sindy 2018/3/1
Public Const EMP_發文歸檔 = "34" 'Add By Sindy 2018/7/9
Public Const EMP_客戶會稿 = "35" 'Add By Sindy 2018/8/28
Public Const EMP_查名 = "36" 'Add By Sindy 2019/6/27
Public Const EMP_查名結果 = "37" 'Add By Sindy 2019/6/27
'Add By Sindy 2023/9/27
Public Const EMP_翻譯交稿 = "38"
Public Const EMP_交辦 = "39"
Public Const EMP_送排版 = "40"
Public Const EMP_排版完成 = "41"
Public Const EMP_核稿分案 = "42"
Public Const EMP_送轉檔 = "43"
Public Const EMP_轉檔完成 = "44"
Public Const EMP_送核稿分案 = "45"
Public Const EMP_程序送判 = "46"
Public Const EMP_程序退回 = "47"
'2023/9/27 END
'***************************************
'Add By Sindy 2025/7/31 僅為當下操作流程時使用,不會新增歷程資訊至DB
Public Const EMP_收文分析 = "A1"
'***************************************

'Added by Morgan 2025/7/31
'稽核日誌類別常數
Public Const AL_登入 = "01"
Public Const AL_登出 = "02"
Public Const AL_上傳 = "03"
Public Const AL_下載 = "04"
Public Const AL_刪除 = "05"
'end 2025/7/31

'Add By Sindy 2013/9/17
'Modify By Sindy 2016/2/23 +,'31'
Public Const EMP_流程控制除外的狀態 = "'00','20','21','22','31'"
'Modify By Sindy 2021/10/15 +,'19'
Public Const EMP_待辦歷程查詢除外的狀態 = "'00','19','20','21','31','33','34','35'"
'2013/9/17 END
'Add By Sindy 2015/12/18
'Modify By Sindy 2016/2/23 +,'26'
'Modify By Sindy 2023/9/28 +,'38','40','41','43','45','46'
Public Const EMP_需等待回覆的狀態 = "'04','05','08','12','14','17','23','26','38','40','41','43','45','46'"
'2015/12/18 END
'Add By Sindy 2013/11/28
'Modify By Sindy 2016/2/23
Public Const EMP_收受者為承辦人 = "'00','01','03','06','07','09','10','13','16','17','18','25','27','28','29','32','36','37','41','42'"
Public Const EMP_收受者為繪圖人員 = "'00','02','11','16','17','22','24','30'"
Public Const EMP_收受者為智權人員 = "'00','08','17','26'"
'2016/2/23 END
'2013/11/28 END
Public Const EMP_收受者為核判或繪圖主管 = "'04','05','12','14','23'" 'Add By Sindy 2018/12/24
'***** 電子化,電子檔名的保留字 EFileCaption *****
Public Const EMP_承辦單 = "WorkSheet"
Public Const EMP_多案承辦單 = "WrkSht" 'Add By Sindy 2020/9/26
Public Const EMP_回覆單 = "Reply"
Public Const EMP_接洽單 = "Order"
Public Const EMP_Email = "Email" 'Add By Sindy 2015/9/9
Public Const EMP_結案單 = "Close"
Public Const EMP_客戶資料 = "Case"
Public Const EMP_存卷資料 = "Info"
Public Const EMP_通知函 = "Cus"     'Add By Sindy 2015/7/27
Public Const EMP_銷案銷帳單 = "Off" 'Add By Sindy 2015/7/27
'***********************************
Public Const EfileNameFCP_04 = "Notice of Allowance with translation.pdf" 'Add By Sindy 2015/6/18 核准函
Public Const EfileNameFCP_14 = "pre-grant gazette.pdf" 'Add By Sindy 2015/6/16 公開通知函:公開公報PDF檔
Public EfileNameFCP_08 As String 'Add By Sindy 2015/7/9 請款單(Debit Note)電子檔

'Add By Sindy 2025/11/3
Public Tmpfrm060209 As Form
Public Tmpfrm180201 As Form
Public Tmpfrm180101 As Form
Public Tmpfrm180203_1 As Form
Public Tmpfrm160102 As Form
Public Tmpfrm160018 As Form
Public Tmpfrm010035_2 As Form
'2025/11/3 END
Public Tmpfrm210147 As Form
Public Tmpfrm210148 As Form
Public Tmpfrm06010616 As Form 'Add By Sindy 2022/10/17
Public Tmpfrm090401 As Form 'Add By Sindy 2015/7/15
'Add By Sindy 2015/10/20
Public Tmpfrm071004 As Form
Public Tmpfrm071005 As Form
Public Tmpfrm1103_2 As Form
'2015/10/20 END

'Add By Sindy 2021/5/28 財務L收據,抓智慧所智權人員
Public Const strLOSSalesDuty As String = _
   "SELECT a0j13" & _
   " FROM LawOfficeSource, CaseProgress C1, CaseProgress C2, acc0j0, lawcase, staff" & _
   " WHERE a0j13=a0k01 AND LOS06(+)=a0j01 AND LOS06=C1.cp09(+)" & _
   " AND C1.cp01=LC01(+) AND C1.cp02=LC02(+) AND C1.cp03=LC03(+) AND C1.cp04=LC04(+)" & _
   " AND LC01 IS NOT NULL AND LOS02 IS NOT NULL AND LOS02<>'C'" & _
   " AND LOS10=C2.cp09(+) AND LOS10 is not null and C2.cp13=st01(+) \#ST06SQL#\" & _
   " Union All " & _
   "SELECT a0j13" & _
   " FROM LawOfficeSource, CaseProgress C1, CaseProgress C2, acc0j0, hirecase, staff" & _
   " WHERE a0j13=a0k01 AND LOS06(+)=a0j01 AND LOS06=C1.cp09(+)" & _
   " AND C1.cp01=HC01(+) AND C1.cp02=HC02(+) AND C1.cp03=HC03(+) AND C1.cp04=HC04(+)" & _
   " AND HC01 IS NOT NULL AND LOS02 IS NOT NULL AND LOS02<>'C'" & _
   " AND LOS10=C2.cp09(+) AND LOS10 is not null and C2.cp13=st01(+) \#ST06SQL#\"

'Added by Lydia 2015/11/16 查名電子化
'Modified by Lydia 2016/03/21 +相同本所案(相同△)
'Public Const TMQ_結果清單 = "1 近似△,2 相同,3 近似,4 稍近似,5 無,9 不查"
'Public Const TMQ_結果查詢 = "'1','近似△','2','相同','3','近似','4','稍近似','5','無','9','不查',''"
'Modified by Lydia 2016/06/02 查名結果改成 PUB_GetTMQans
Public Const TMQ_AkindPic = "0"     '在明細檔、附件檔的圖形類別
Public Const TMQ_AkindWord1 = "1"   '在明細檔、附件檔的文字1類別
Public Const TMQ_AkindWord2 = "2"   '在明細檔、附件檔的文字2類別
Public Const TMQ_CtrRead = False     '是否控制結果已讀才能收文 '2016/03/28 針對先收文的情況,預設全部不控制
Public Const TMQ_附件F02 = "0"      '預設查名內容的附件編號
Public Const TMQ_附件F04 = "00"     '預設查名內容的附件編號
Public Const TMQ_查名作業 = "TS"    '查名結果附件的副檔名
Public Const TMQ電子化啟用日 = "20160420" 'Modified by Lydia 2016/04/20 下午開始上查名單電子化
Public Const TMQ_ReApp = True 'Added by Lydia 2016/04/06 智權人員可否重複收文(申請案)
Public Const TMQ_T案 = "101,737" 'Added by Lydia 2016/04/25 T案申請(查名單對應案件進度) 'Modified by Lydia 2021/11/19 增加737智財協作之T案
Public Const TMQ_TS案 = "001" 'Added by Lydia 2016/04/25 TS案申請(查名單對應案件進度)
Public Const TMQFileFTP = "20160701" 'Added by Lydia 2016/06/23 查名單附件改放在FTP的啟用日,並且同時開啟核可案狀態

'2015/12/18 薪資查詢共用參數 ADD
Public Const Pub_StartYM As String = 201601 '啟用資料年月
Public Pub_MaxSMYM As String                '可查詢薪資資料最大年月
Public Pub_MaxYBYear As String              '可查詢年終資料最大年度
Public Pub_MaxOB1Year As String             '可查詢端午資料最大年度
Public Pub_MaxOB2Year As String             '可查詢中秋資料最大年度
Public Pub_StaffList As String              '可查詢薪資資料人員名單
Public Pub_StaffBonusList As String         '可查詢年終獎金資料人員名單  add by sonia 2025/2/5與Pub_StaffList區分權限
'2015/12/18

'Add by Sindy 2025/6/27 'or cu80='其他' or cu80='業務自行處理' or cu80='國內同業' or cu80='解除對造' or cu80='不得代理專利' or cu80='不得代理商標' or cu80='設為對造'
Public Const 客戶及代理人可讀取的狀態 = "其他,業務自行處理,國內同業,解除對造,不得代理專利,不得代理商標,不得代理,設為對造"

'Modify By Sindy 2016/6/15 + ,'8','開拓'
Public Const Show國外部信件分類 = "'1','個案','2','外商','3','外專','4','專利處','5','外法','6','新知','7','財務','Z','其他','8','開拓'" 'Modify by Sindy 2016/1/14
'Add by Sindy 2016/1/14 + Show專利處信件分類
'Modify By Sindy 2018/6/21
'Public Const Show專利處信件分類 = "'1','P程序1','2','P程序2','3','亞洲','4','歐洲','5','美洋非(單)','6','美洋非(雙)','7','其他','8','垃圾信箱','9','國外部匯入'"
'Add by Sindy 2019/4/2 國外部匯入==>其他信箱匯入
Public Const Show專利處信件分類 = "'1','P程序1','2','P程序2','3','美日(單)','4','美日(雙)','5','美日外(單)','6','美日外(雙)','7','其他','8','垃圾信箱','9','其他信箱匯入','A','亞洲','B','歐洲','C','美洋非(單)','D','美洋非(雙)'"
'2018/6/21 END
'Add by Sindy 2019/4/2 + Show商標處信件分類
Public Const Show商標處信件分類 = "'1','MCTF','2','大陸案','3','個人','4','非大陸案','5','其他','6','其他信箱匯入'"
'Added by Lydia 2016/02/15 每日批次(frmAutoBatchDay)與業務目標及達成通知日報表(frm210107)共用
'Modified by Lydia 2016/05/16 去掉中區其他S29
'Public Const AutoBatchSalesArea = "DECODE(SUBSTR(ST15,1,1),'F','Z',DECODE(ST15,'P29','X',DECODE(INSTR('S11,S13,S14,S15,S21,S23,S24,S29,S31,S41',ST15),0,'W',ST15))"
'Modified by Lydia 2017/10/12 (10/1起) S22成立
'Public Const AutoBatchSalesArea = "DECODE(SUBSTR(ST15,1,1),'F','Z',DECODE(ST15,'P29','X',DECODE(INSTR('S11,S13,S14,S15,S21,S23,S24,S31,S41',ST15),0,'W',ST15))"
'Modified by Lydia 2022/06/07 因為林炳佑歸S29,所以S29再加回來,另外去掉S23; P.S.因為frm210107部門列出更多,現在已不共用
'Modified by Lydia 2023/12/28 +S10台北所
'Public Const AutoBatchSalesArea = "DECODE(SUBSTR(ST15,1,1),'F','Z',DECODE(ST15,'P29','X',DECODE(INSTR('S10,S11,S13,S14,S15,S21,S22,S24,S29,S31,S41',ST15),0,'W',ST15))" 'Mark by Lydia 2025/01/20 每日批次StrMenu64不再使用
'Added by Lydia 2016/05/17 CFP控制年費.延展費及維持費智權同仁可加的點數
Public Const CFP_dg605 As Single = 8
Public Const CFP_dg606 As Single = 10
Public Const CFP_dg607 As Single = 10
Public Tmpfrm880004_4 As Form '輸入提高點數簽核主管及簽核點數,設定在mdiMain的PrintLetter報價轉定稿
'Added by Lydia 2016/09/10 代表人中文名稱長度和代表人英文名稱長度
Public Const Pub_MaxCEL10 As Integer = 50 '代表人中文名稱長度
Public Const Pub_MaxCEL11 As Integer = 80 '代表人英文名稱長度
'end 2016/09/10
'Modified by Morgan 2024/6/7 +221EPC
Public Const CFP_ChkEntity As String = "101,102,203,040,030,221" 'Added by Lydia 2016/09/13 需要檢查個案和客戶減免設定的國別
Public Const 專利客戶案號max As Integer = 120 'Added by Lydia 2017/06/14 專利案的客戶
'Mark by Lydia 2025/03/13 改用模組取得>>Pub_SetF51Order
''Public Const 外翻Y編號 As String = "Y53541000,Y52268000,Y54868000" 'Added by Lydia 2017/10/16 FCP案外翻(捷恩凱Y53541,舜禹Y52268,迅達Y54868)的Y編號,用在對外聯絡(FCP)和付款(account)
''Added by Lydia 2018/01/04 外翻員工編號
'Public Const 外翻_舜禹 As String = "F5588"
'Public Const 外翻_捷恩凱 As String = "F5653"
'Public Const 外翻_迅達 As String = "F5698"
'end 2025/03/13
'Move by Lydia 2018/01/09
Public Const FCPHaveEP04 = "201,209,210,235,927" 'Added by Lydia 2016/06/21 需要抓核稿人的案件性質
Public Const FCPHaveEP09 = "'201','927'" 'Added by Lydia 2016/06/21 翻譯完稿輸入的案件性質
'Memo by Lydia 2018/05/15 外專命名作業
Public Const FcpTctPtys = "201,209,210,242,235,942" 'Added by Lydia 2017/11/14預設的中說案件性質; Modified by Lydia 2018/05/10 +942 檢視PCT公開本與FCP相異處
Public Const FCP命名記錄 = "RCD.Menu"  'Added by Lydia 2017/11/23 卷宗區-電子檔副檔名
Public Const FcpTcnFKey01 = ".msg" 'Added by Lydia 2017/12/04 檢查命名追蹤的必要上傳檔案
Public Const FcpTcnFKey02 = ".ORI.PDF" 'Added by Lydia 2017/12/11 外文原文本
'end 2018/05/15
Public Const FCP提供文件 = "CASE.Menu" 'Added by Lydia 2018/03/06 卷宗區-電子檔副檔名
'Public Const Taie_Jpn_Title = "台一�罈痡M利法律事務所" 'Mark by Amy 2020/04/06 不使用 'Added by Lydia 2018/03/16 日文抬頭
'Added by Lydia 2020/01/17 專利案件和English_Vers檔案：上傳到原始檔區所掛的案件性質
Public Const cnt專利案件 = "991"
Public Const cntEnglish_Vers = "992"
'Add by Amy 2020/08/05
Public Const 智慧局送件開票銀行 = "011010075"
Public Const 智慧局送件開票帳號 = "1756650"
Public Const cntAutoQueryInterval = 30 '自動跑語法時間(分鐘),解決跨網段會自動斷線問題

'Added by Morgan 2021/4/13
'自訂提示視窗用(替換ToolTipText功能)
Private Type TOOLINFO
    lSize       As Long
    lFlags      As Long
    hWnd        As Long
    lId         As Long
    '
    'lpRect      As RECT
    Left        As Long
    Top         As Long
    Right       As Long ' This is +1 (right - left = width)
    Bottom      As Long ' This is +1 (bottom - top = height)
    '
    hInstance   As Long
    lpStr       As String
    lParam      As Long
End Type
Private Declare Function SendMessageLong Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
'
Private Const WM_USER               As Long = &H400&
Private Const CW_USEDEFAULT         As Long = &H80000000
'
Private Const TTM_ACTIVATE          As Long = WM_USER + 1&
'Private Const TTM_ADDTOOLA          As Long = WM_USER + 4&
Private Const TTM_ADDTOOLW          As Long = WM_USER + 50&
Private Const TTM_SETDELAYTIME      As Long = WM_USER + 3&
'Private Const TTM_UPDATETIPTEXTA    As Long = WM_USER + 12&
Private Const TTM_UPDATETIPTEXTW    As Long = WM_USER + 57&
Private Const TTM_SETTIPBKCOLOR     As Long = WM_USER + 19&
Private Const TTM_SETTIPTEXTCOLOR   As Long = WM_USER + 20&
Private Const TTM_SETMAXTIPWIDTH    As Long = WM_USER + 24&
'Private Const TTM_SETTITLEA         As Long = WM_USER + 32&
Private Const TTM_SETTITLEW         As Long = WM_USER + 33&
'
Private Const TTS_NOPREFIX          As Long = &H2&
Private Const TTS_BALLOON           As Long = &H40&
Private Const TTS_ALWAYSTIP         As Long = &H1&
'
Private Const TTF_CENTERTIP         As Long = &H2&
Private Const TTF_IDISHWND          As Long = &H1&
Private Const TTF_SUBCLASS          As Long = &H10&
Private Const TTF_TRANSPARENT       As Long = &H100&
'
Private Const TTDT_AUTOPOP          As Long = 2&
Private Const TTDT_INITIAL          As Long = 3&
'
Private Const TOOLTIPS_CLASS        As String = "tooltips_class32"
'
Public Enum ttIconType
    TTNoIcon
    TTIconInfo
    TTIconWarning
    TTIconError
End Enum
#If False Then ' Intellisense fix.
    Public TTNoIcon, TTIconInfo, TTIconWarning, TTIconError
#End If
'
Private hwndTT As Long ' hwnd of the tooltip
'end 2021/4/13

'Added by Morgan 2021/4/14
'Unicode Msgbox
Private Const MB_USERICON = &H80&

Private Type MsgBoxParams
    cbSize As Long
    hwndOwner As Long
    hInstance As Long
    lpszText As Long
    lpszCaption As Long
    dwStyle As Long
    lpszIcon As Long
    dwContextHelpId As Long
    lpfnMsgBoxCallback As Long
    dwLanguageId As Long
End Type

Private Declare Function MessageBoxIndirectW Lib "user32" (lpMsgBoxParams As MsgBoxParams) As Long
'end 2021/4/14
'Added by Lydia 2021/04/14 從acc_var(Finance)移過來
Public strTrackMode  As String 'Added by Lydia 2021/03/18 Form2.0 記錄鍵盤傳入順序
Private Const cntTrackRun = "KeyUp:" 'Added by Lydia 2021/03/18  Form2.0 記錄鍵盤傳入順序:確定執行階段
'end 2021/04/14
Public Const ACS_PFrateStart = "20210722" 'Added by Lydia 2021/04/29 ACS智財顧問專業分配比例管制：啟用日
Public Const 外專新案認領啟用日 = "20230508" 'Added by Lydia 2023/02/15
Public Const 新部門啟用日 = "20240101" 'Added by Lydia 2023/12/14 'Memo by Lydia 2023/12/28 先延至1/10
Public Const 風險警示啟用日 = 20991231 'Modify by Amy 2023/12/15
Public Const CaseLawerPtyList As String = "220113" 'Added by Lydia 2024/04/23 須限制案件性質表規費科目才可點出庭律師。
Public Const 財務拆總帳出納國內應收啟用日 = "20240520" 'Modify by Amy 2024/05/17

'從 Service_Const1 及 aacc_msg 移來(合併) Morgan 2013/8/12
Public Function MsgText(strMsg As Integer) As String
   Select Case strMsg
      '財務訊息
      Case 1
         MsgText = "遞增"
      Case 2
         MsgText = "遞減"
      Case 3
         MsgText = "A"
      Case 4
         MsgText = "E"
      Case 5
         MsgText = "警告!!"
      Case 6
         MsgText = "確定刪除?"
      Case 7
         MsgText = "已經是第一筆資料..."
      Case 8
         MsgText = "已經是最後一筆資料..."
      Case 9
         MsgText = "此筆資料已存在..."
      Case 10
         MsgText = "必要欄位，請輸入資料..."
      Case 11
         MsgText = "借貸方金額不平衡，請再確認..."
      Case 12
         MsgText = "00"
      Case 13
         MsgText = "31"
      Case 14
         MsgText = "已過帳之傳票，無法變更內容..."
      Case 15
         MsgText = "過帳完成..."
      Case 16
         MsgText = "001"
      Case 17
         MsgText = "存檔完成..."
      Case 18
         MsgText = "R"   '應收票據'
      Case 19
         MsgText = "P"   '應付票據"
      Case 20
         MsgText = "刪除完成..."
      Case 21
         MsgText = "注意!!"
      Case 22
         MsgText = "兌現作業完成..."
      Case 23
         MsgText = "傳票轉外帳作業完成..."
      Case 24
         MsgText = "ZZ"  '總所代號'
      Case 25
         MsgText = "年度結轉作業完成..."
      Case 26
         MsgText = "處理中，請稍候..."
      Case 27
         MsgText = "收據給號作業完成..."
      Case 28
         MsgText = "查無此資料..."
      Case 29
         MsgText = "___/__/__"
      Case 30
         MsgText = "金額不符，請再確認..."
      Case 31
         MsgText = "全部"
      Case 32
         MsgText = "請檢核流水號是否有誤..."
      Case 33
         MsgText = "此筆資料不存在..."
      Case 34
         MsgText = "此筆資料已抵帳或已付款或未審核或已作廢或資料不存在..."
      Case 35
         MsgText = " / "
      Case 36
         MsgText = "請確認是否新增此本所案號"
      Case 37
         MsgText = "本所案號之流水號不可大於目前流水號"
      Case 38
         MsgText = "申請國家不可輸入 001 - 008"
      Case 39
         MsgText = "或此票據已作廢..."
      Case 40
         MsgText = "應收"
      Case 41
         MsgText = "應付"
      Case 42
         MsgText = "此張收據已開立過..."
      Case 43
         MsgText = ".xls"
      Case 44
         MsgText = "借/貸方別:只可為1--借或2--貸 ..."
      Case 45
         MsgText = "查無此"
      Case 46
         MsgText = "..."
      Case 47
         MsgText = "借/貸方金額不可同時輸入"
      Case 48
         MsgText = "每月傳票日不可為 0, 亦不可大於 28..."
      Case 49
         MsgText = "分攤總比例不可大於100%..."
      Case 50
         MsgText = "未輸入資料無法存檔..."
      Case 51
         MsgText = "分攤總比例未分配至100%..."
      Case 52
         MsgText = "不可為空白..."
      Case 53
         MsgText = "限於1至4..."
      Case 54
         MsgText = "必需為 Y 或 N ..."
      Case 55
         MsgText = "TOT"
      Case 56
         MsgText = "限於1至28..."
      Case 57
         MsgText = "到期日不可小於收/開票日..."
      Case 58
         MsgText = "金額不可為0..."
      Case 59
         MsgText = "金額不符..."
      Case 60
         MsgText = "此張票據無法託收..."
      Case 61
         MsgText = "此張票據無法貼現..."
      Case 62
         MsgText = "此張票據無法作廢..."
      Case 63
         MsgText = "輸入錯誤..."
      Case 64
         MsgText = "輸入字數過長..."
      Case 65
         MsgText = "輸入字數限制為 "
      Case 66
         MsgText = "原資料已變更，是否存檔?"
      Case 67
         MsgText = "收票作業轉檔..."
      Case 68
         MsgText = "開票作業轉檔..."
      Case 69
         MsgText = "即期票存入作業轉檔..."
      Case 70
         MsgText = "銀行調節作業轉檔..."
      Case 71
         MsgText = "票據轉出作業轉檔..."
      Case 72
         MsgText = "票據兌現作業轉檔(應收)..."
      Case 73
         MsgText = "票據兌現作業轉檔(應付)..."
      Case 74
         MsgText = "每月固定傳票轉檔..."
      Case 75
         MsgText = "M" '管理部門
      Case 76
         MsgText = "結算完成..."
      Case 77
         MsgText = "月結日期不能早於上次月結日，請檢查月結日期..."
      Case 78
         MsgText = "處理完成..."
      Case 79
         MsgText = "應收/付資料轉檔..."
      Case 80
         MsgText = "已轉暫收款, 無法存檔..."
      Case 81
         MsgText = "收款金額不可大於應收金額..."
      Case 82
         MsgText = "溢收金額不可小於零..."
      Case 83
         MsgText = "已收規費不可大於應收規費..."
      Case 84
         MsgText = "本次收款總額不等於實際收款額..."
      Case 85
         MsgText = "請輸入日期..."
      Case 86
         MsgText = "已過帳, 不可修改..."
      Case 87
         MsgText = "已過帳, 不可刪除..."
      Case 88
         MsgText = "此往來對象今日已付款..."
      Case 89
         MsgText = "溢收金額不可小於零..."
      Case 90
         MsgText = "退費匯差"
      Case 91
         MsgText = "一張收據最多選至兩筆收文..."
      Case 92
         MsgText = "兩個以上之智權人員不可開成同一張收據..."
      Case 93
         MsgText = "是否開此公司別 ??"
      Case 94
         MsgText = "請放入專利收據紙張..."
      Case 95
         MsgText = "請放入商標收據紙張..."
      Case 96
         MsgText = "請放入律師收據紙張..."
      Case 97
         MsgText = "請先點選欲開收據之收文..."
      Case 98
         MsgText = "按 F12 查詢"
      Case 99
         'Modify By Sindy 2015/12/9
         'MsgText = "* 以上帳款如有任何疑問，請洽 (02)2506-1023 分機 545 楊小姐"
         'Modify by Amy 2024/05/13 財務拆成總帳、出納、國內應收,故拿掉分機顯示,於各程式中再加
         MsgText = "* 以上帳款如有任何疑問，請洽 (02)2506-1023" '& Pub_EMPTelNumbrandName(Pub_GetSpecMan("財務處總帳人員"), "#") 'Modify By Sindy 2021/6/9 IIf(Pub_GetSpecMan("財務處總帳人員") = "71006", "#545 楊小姐", "#543 辜小姐")
         '2015/12/9 END
      Case 100
         MsgText = "請更換套表"
      Case 101
         MsgText = "請更換小表"
      Case 102
         MsgText = "請更換大表"
      Case 103
         MsgText = "請輸入部門別..."
      Case 104
         MsgText = "     啟"
      Case 105
         MsgText = "產生傳票之貸方科目中..."
      Case 106
         MsgText = "開啟本所案號金額輸入畫面中..."
      Case 107
         MsgText = "若為部份收款時, 請按 F12 依收文號分配收款金額"
      Case 108
         MsgText = "銷帳金額不可大於應收金額"
      Case 109
         MsgText = "退費金額不可大於已收金額"
      Case 110
         MsgText = "未收款之收據不可退費"
      Case 111
         MsgText = "查無此票據"
      Case 112
         MsgText = "暫收款金額不符"
      Case 113
         MsgText = "會計科目不符"
      Case 114
         MsgText = "依客戶抬頭跳頁"
      Case 115
         MsgText = "請確認退費金額"
      Case 116
         MsgText = "本次扣繳額不等於實際扣繳額..."
      Case 117
         MsgText = "本次退服務總額不等於實際已退服務費總額..."
      Case 118
         MsgText = "本次退規費總額不等於實際已退規費總額..."
      Case 119
         MsgText = "本次銷帳總額不等於實際已銷帳總額..."
      Case 120
         MsgText = "扣單金額與實際補扣繳金額不符"
      Case 121
         MsgText = "本次退費金額 $"
      Case 122
         MsgText = "未扣繳+已扣繳不可大於應扣繳金額"
      Case 123
         MsgText = "請按轉出單號"
      Case 124
         MsgText = "請輸入貸方金額"
      Case 125
         MsgText = "本次扣繳款項收回金額 $"
      Case 126
         MsgText = "已收款不可變更"
      Case 127
         MsgText = "結匯"
      Case 128
         MsgText = "資料輸入不正確"
      Case 129
         MsgText = "應收/付傳票更新..."
      Case 130
         MsgText = "必須為數字..."
      Case 131
         MsgText = "是否存檔..."
      Case 132
         MsgText = "請將溢收轉暫收款..."
      Case 133
         MsgText = "按 Insert 鍵儲存資料"
      Case 134
         MsgText = "按 F12 鍵查詢資料"
      Case 135
         MsgText = "按 F12 鍵可調出結匯的幣別"
      Case 136
         MsgText = "無應付款資料..."
      Case 137
         MsgText = "已兌現支票..."
      Case 138
         MsgText = "已兌領支票..."
      Case 139
         MsgText = "請先列印當月份綜合損益表"
      Case 140
         MsgText = "月結算年度輸入錯誤"
      Case 141
         MsgText = "月結算月份不可為12月"
      Case 142
         MsgText = "此對象無應付款資料..."
      Case 143
         MsgText = "或此票據未到期..."
      Case 144
         MsgText = "年度結轉必需為12月..."
      Case 145
         MsgText = "此張票據未託收, 已兌現, 已貼現或退票..."
      Case 146
         MsgText = "列印類別, 請輸入 1 or 2"
      Case 147
         MsgText = "已收款, 不可列印..."
      Case 148
         MsgText = "此張票據已存在, 不可重複輸入..."
      Case 149
         MsgText = "無代理人, 請檢核..."
      Case 150
         MsgText = "此查詢約一分鐘..."
      Case 151
         MsgText = "請更換中一刀報表"
      Case 152
         MsgText = "或此票據已調節..."
      Case 153
         MsgText = "或帳號不符..."
      Case 154
         MsgText = "此張票據不存在, 已轉出, 已兌現或退票..."
      Case 155
         MsgText = "傳票已過帳, 不可變更原始資料..."
      Case 156
         MsgText = "國籍空白約五分鐘"
      Case 157
         MsgText = "按 Tab 鍵調出補扣繳資料"
      Case 158
         MsgText = "已出傳票, 不可刪除原始資料..."
      Case 159
         MsgText = "此張票據"
      Case 160
         MsgText = "是否重複列印?"
      Case 161
         MsgText = "請輸入往來類別..."
      Case 162
         MsgText = "請輸入"
      Case 163
         MsgText = "請更換"
      Case 164
         MsgText = "公司收據"
      Case 165
         MsgText = "此客戶本日已無收文案件了..."
      Case 166
         MsgText = "暫收款不可退費二次以上..."
      Case 167
         MsgText = "此筆資料不符..."
      Case 168
         MsgText = "前月損益轉入"
      Case 169
         MsgText = "此張收據已作全額銷帳, 無法列印"
      Case 170
         MsgText = "請輸入公司別..."
      Case 171
         MsgText = "一張收據不可選取超過兩筆以上的收文..."
      Case 172
         MsgText = "需輸入六碼..."
      Case 173
         MsgText = "無"
      Case 174
         MsgText = "無客戶國籍, 請補輸入後再開立..."
      Case 175
         MsgText = "未發文或已收文取消..."
      Case 176
         MsgText = "此收文號已開過五次帳單..."
      Case 177
         '2007/8/9 MODIFY BY SONIA
         'MsgText = "本收據已收款..."
         MsgText = "本收據已收款, 若傳票已過帳則請於備註註明修改金額原因"
      Case 178
         MsgText = "此收文之代理人不符合..."
      Case 179
         MsgText = "請先輸入借方資料..."
      Case 180
         MsgText = "兌現傳票已過帳, 不可變更原始資料..."
      Case 181
         MsgText = "請輸入條件後, 再按F12查詢..."
      Case 182
         MsgText = "已開發票, 不可退費..."
      Case 183
         MsgText = "外帳1公司傳票轉入其他公司別完成..."
      Case 184
         MsgText = "傳票號碼重編完成..."
      Case 185
         MsgText = "處理失敗..."
      Case 186
         MsgText = "請輸入帳單金額..."
      Case 187
         MsgText = "請主管核示是否付款..."
      Case 188
         MsgText = "查無此"
      Case 189
         MsgText = "扣單金額不可為零..."
      Case 190
         MsgText = "主管未審核，無法結匯..."
      Case 191
         MsgText = "無此收據的收文號資料..."
      Case 192
         MsgText = "查詢中..."
      Case 193
         MsgText = "此票據已處理，不可修改，如欲修改，請通知票據作業人員..."
      Case 194
         MsgText = "本張收據已收款請於收款作業更改摘要【重新收款】"
      Case 195
         MsgText = "請輸入條件後, 再按列印..."
      Case 196
         MsgText = "此張票據已存在, 請確認..."
      Case 197
         MsgText = "分錄傳送中, 請稍待..."
      Case 198
         MsgText = "請輸入非總所之部門別..."
      Case 199
         MsgText = "請輸入數字1~3..."
      Case 200
         MsgText = "請輸入數字1~2..."
      Case 201
         MsgText = "代理人或金額不符合..."
      Case 202
         MsgText = "請輸入代理人編號..."
      Case 203
         MsgText = "已收款不可作廢..."
      Case 204
         MsgText = "上期餘額"
      Case 205
         MsgText = "此收文之帳單, 已超過五次..."
      Case 206
         MsgText = "修改幣別時, 請重新輸入明細資料, 以重新計算正確損益..."
      Case 207
         MsgText = "列印完成..."
      Case 208
         MsgText = "此帳單無收文紀錄..."
      Case 209
         MsgText = "借貸方金額不平衡 --> "
      Case 210
         MsgText = "傳票號碼迄號必需輸入..."
      Case 211
         MsgText = "請放入地址條..."
      Case 212
         MsgText = "請放入大報表..."
      Case 213
         MsgText = "請放入中一刀報表..."
      Case 214
         MsgText = "手開收據或開立發票, 不可列印此收據 !"
      Case 215
         MsgText = "匯票號碼不可重複..."
      Case 216
         MsgText = "無結匯資料, 不可存檔..."
      Case 217
         MsgText = "本次扣繳退費不等於實際已扣繳退費..."
      Case 218
         MsgText = "已結匯, 不可作廢..."
      Case 501
         MsgText = "="
      Case 502
         MsgText = ">"
      Case 503
         MsgText = "<"
      Case 504
         MsgText = ">="
      Case 505
         MsgText = "<="
      Case 506
         MsgText = "<>"
      Case 507
         MsgText = "and"
      Case 508
         MsgText = "or"
      Case 601
         MsgText = ""
      Case 602
         MsgText = "Y"
      Case 603
         MsgText = "N"
      Case 801
         MsgText = "D"   '傳票編號-1公司 /819:傳票編號-J公司
      Case 802
         MsgText = "E"   '國內收據號碼'
      Case 803
         MsgText = "F"   '國內收款單號'
      Case 804
         MsgText = "G"   '國內應付單號'
      Case 805
         MsgText = "I"   '國內銷帳退費單號'
      Case 806
         MsgText = "J"   '國內暫收款單號'
      Case 807
         MsgText = "K"   '扣繳憑單編號'
      Case 808
         MsgText = "M"   '國外收款單號'
      Case 809
         MsgText = "N"   '國外暫收款單號'
      Case 810
         MsgText = "O"   '國外暫收款退費單號'
      Case 811
         MsgText = "Q"   '國外銷帳單號'
      Case 812
         MsgText = "U"   '國外帳單編號'
      Case 813
         MsgText = "V"   '國外抵帳單編號'
      Case 814
         MsgText = "W"   '國外付款單號'
      Case 815
         MsgText = "X"   '國外請款編號'
      Case 816
         MsgText = "YY"  '貼現序號'
      Case 817
         MsgText = "Z"   'D/N No.抵帳編號'
      Case 818
         MsgText = "G"  '國內付款單號'
      'Add by Amy 2013/12/18
      Case 819
        MsgText = "JD"  '傳票編號-J公司 /801:傳票編號-1公司
      'Add by Amy 2020/03/17
      Case 820
        MsgText = "LD" '傳票編號-L公司
      Case 901
         MsgText = "專利商標"
      Case 902
         MsgText = "專利法律"
      Case 903
         MsgText = "桂律師"
      Case 904
         MsgText = "蔣律師"
      Case 905
         MsgText = "詹律師"
      Case 906
         MsgText = "唐律師"
      'Add By Sindy 2013/12/18
      Case 907
         MsgText = "智權"     'modify by sonia 2023/4/11 原為"台一智權"
      'Add By Sindy 2013/12/19
      Case 908
         MsgText = "開發"
      '收文訊息
      Case 1501
         MsgText = "讀取案件進度檔時發生錯誤"
      Case 1502
         MsgText = "找不到此收文號在案件進度檔之資料"
      Case 1503
         MsgText = "找不到此本所案號在客戶基本檔之資料"
      Case 1504
         MsgText = "找不到此本所案號在商標基本檔之資料"
      Case 1505
         MsgText = "找不到此收文號在來函記錄檔之資料"
      Case 1506
         MsgText = "錯誤之系統代號"
      Case 1507
         MsgText = "讀取案件名稱及申請人時失敗"
      Case 1508
         MsgText = "本所案號輸入錯誤：案號不符"
      Case 1509
         MsgText = "本所案號輸入錯誤：長度不符"
      Case 1510
         MsgText = "本所案號輸入錯誤：案號不存在或重複"
      Case 1511
         MsgText = "本所案號輸入錯誤：無流水號"
      Case 1512
         MsgText = "本所案號輸入錯誤：無追加號"
      Case 1513
         MsgText = "設定標準價及底價時，發生錯誤"
      Case 1514
         MsgText = "從下一程序檔取回機關文號、對造名稱時發生錯誤"
      Case 1515
         MsgText = "從下一程序檔取回本所期限、法定期限時發生錯誤"
      Case 1516
         MsgText = "001"
      Case 1517
         MsgText = "存檔完成..."
      
      '專利訊息
      Case 1001
         MsgText = "讀取CaseMap檔案失敗"
      Case 1002
         MsgText = "無發明人資料"
      Case 1003
         MsgText = "無法新增指定國家資料"
      Case 1004
         MsgText = "無法新增資料至PermitRecord檔"
      Case 1005
         MsgText = "新增至CaseMap時失敗"
      Case 1006
         MsgText = "新增至CaseMap時失敗"
      Case 1007
         MsgText = "國內外案件關聯不存在"
      Case 1008
         MsgText = "讀取檔案失敗"
      Case 1009
         MsgText = "案件之本所案號之案件性質不為申請案"
      Case 1010
         MsgText = "案件之本所案號不存在"
      Case 1011
         MsgText = "案號已發文"
      Case 1012
         MsgText = "案號之案件性質為檢索報告，此案號錯誤"
      Case 1013
         MsgText = "內部收文不可新增新案號"
      Case 1014
         MsgText = "請輸入案件性質"
      Case 1015
         MsgText = "請輸入申請人"
      Case 1016
         MsgText = "案件性質輸入長度不符"
      Case 1017
         MsgText = "本所案號輸入長度不符"
      Case 1018
         MsgText = "顧問之非聘任案件，不得新增"
      Case 1019
         MsgText = "新申請案或異議、舉發、評定、廢止案，必須為新本所案號"
      Case 1020
         MsgText = "期限：必須選擇一種格式輸入"
      Case 1021
         MsgText = "你必須輸入正確之本所案號"
      Case 1023
         MsgText = "此筆之收件號"
      Case 1024
         MsgText = "系統種類代碼不存在"
      Case 1025
         MsgText = "期限起算日代碼不存在"
      Case 1026
         MsgText = "期限必須為數字"
      Case 1027
         MsgText = "請先輸入期限起算日"
      Case 1028
         MsgText = "政府機關代碼不存在"
      Case 1029
         MsgText = "讀取資料時發生錯誤"
      Case 1030
         MsgText = "找尋不到資料"
      Case 1031
         MsgText = "案件之中英日文名稱，至少需輸入一項"
      Case 1032
         MsgText = "本所期限小於系統日"
      Case 1033
         MsgText = "本所期限必須≦法定期限"
      Case 1034
         MsgText = "費用必須輸入"
      Case 1035
         MsgText = "點數必須輸入"
      Case 1036
         MsgText = "點數不符"
      Case 1037
         MsgText = "點數必須空白"
      Case 1038
         MsgText = "請輸入N或不輸入任何字"
      Case 1039
         MsgText = "規費必須空白"
      Case 1040
         MsgText = "規費必須輸入"
      Case 1041
         MsgText = "案件之名稱，必需輸入"
      Case 1042
         MsgText = "第二個日期需大於第一個日期"
      Case 1043
         MsgText = "範圍不符"
      Case 1044
         MsgText = "此欄位可不輸入，但如要輸入請輸滿八碼"
      Case 1045
         MsgText = "報表對象不在範圍內"
      Case 1046
         MsgText = "此次共刪除"
      Case 1047
         MsgText = " 筆資料"
      Case 1048
         MsgText = "結束日期應大於開始日期"
      Case 1049
         MsgText = "有下一程序且來函性質有定義工作天數時，承辦期限不可空白"
      Case 1050
         MsgText = "來函收文日不可大於系統日"
      Case 1051
         MsgText = "子案全都無領證"
      Case 1052
         MsgText = "准駁通知日不可大於系統日"
      Case 1053
         MsgText = "因為來函性質為核准，所以此欄位不可空白"
      Case 1054
         MsgText = "郵寄方式種類代碼不存在"
      Case 1055
         MsgText = "說明書不可空白"
      Case 1056
         MsgText = "輸入之系統類別不符,請依照子系統輸入允許之系統類別"
      Case 1057
         MsgText = "無申請案號時不可空白"
      Case 1058
         MsgText = "專利號數不可空白"
      Case 1059
         MsgText = "專利期間輸入錯誤"
      Case 1060
         MsgText = "因為為美國案，所以此欄位不可空白"
      Case 1061
         MsgText = "請至少輸入一項"
      Case 1062
         MsgText = "只可為 P 或 PS 案件"
      Case 1063
         MsgText = "只可為 P 案件"
      Case 1064
         MsgText = "此案號已有申請案號, 請以申請案號輸入"
      Case 1101
         MsgText = "請輸入客戶資料(X..)"
      Case 1102
         MsgText = "本代理人編號已存在."
      Case 1103
         MsgText = "本收文號不存在於案件進度檔!"
      Case 1104
         MsgText = "請輸入代理人資料(Y..)"
      Case 1105
         MsgText = "請檢查費用資料."
      Case 1106
         MsgText = "請檢查輸入資料;對造號數、對造案件名稱、對造名稱三者需全部輸入或全部空白"
      Case 1107
         MsgText = "請檢查系統類別."
      Case 1108
         MsgText = "與櫃台之來函收文記錄不符, 請確認!!"
      'Added by Lydia 2019/12/23 利益衝突案件：提示
      Case 1109
         MsgText = "另有限閱案件"
      Case 1110
         MsgText = "檢查利益衝突案件之權限"
      'Added by Lydia 2020/02/17 案件名稱有特殊字：各式申請書的提示
      Case 1111
         MsgText = "名稱有特殊字，已開啟正確名稱維護檔，請Word自行處理。"
      
      '共同程序
      Case 8001
         MsgText = "讀取ReasonOfRelief檔時，發生錯誤"
      Case 8002
         MsgText = "解除期限日不可大於系統日"
      Case 8003
         MsgText = "閉卷日期不可大於系統日"
      Case 8004
         MsgText = "本所案號與相關本所案號相同"
      Case 8005
         MsgText = "相關本所案號重複"
      Case 8006
         MsgText = "請先選擇要移除之項目"
      '共用訊息
      Case 9001
         MsgText = "警告!!"
      Case 9002
         MsgText = "輸入資料錯誤!"
      '2012/5/14 ADD BY SONIA
      Case 9003
         MsgText = "日期格式錯誤!"
      '2012/5/14 END
      Case 9004
         MsgText = "存檔失敗,請洽系統管理者!"
      Case 9005
         MsgText = "無此權限!"
      Case 9006
         MsgText = "不可為空白!"
      Case 9007
         MsgText = "查無此資料..."
      Case 9008
         MsgText = "已到第一筆..."
      Case 9009
         MsgText = "已到最末筆..."
      Case 9010
         MsgText = "無資料..."
      Case 9011
         MsgText = "此筆資料不存在!"
      Case 9012
         MsgText = "變更資料成功..."
      Case 9013
         MsgText = "變更資料失敗..."
      Case 9014
         MsgText = "請檢查輸入資料..."
      Case 9015
         MsgText = "輸入欄位不可留空白..."
      Case 9016
         MsgText = "請檢查輸入範圍..."
      Case 9017
         MsgText = "刪除資料成功!"
      Case 9018
         MsgText = "刪除資料失敗!"
      Case 9019
         MsgText = "新增資料成功!"
      Case 9020
         MsgText = "新增資料失敗!"
      Case 9021
         MsgText = "不可小於系統日!"
      Case 9022
         MsgText = "不可大於系統日!"
      Case 9023
         MsgText = "輸入日期大於系統日!"
      Case 9024
         MsgText = "輸入日期小於系統日!"
      Case 9025
         MsgText = "輸入日期小於本所期限!"
      Case 9026
         MsgText = "輸入日期大於法定期限!"
      Case 9027
         MsgText = "請洽系統管理員(使用權限)!"
      Case 9028
         MsgText = "請檢查輸入資料(1,2)..."
      Case 9029
         MsgText = "請檢查專利/商標種類代號資料..."
      Case 9030
         MsgText = "Y"
      Case 9031
         MsgText = "N"
      Case 9032
         MsgText = "資料未變動."
      Case 9033
         MsgText = "請輸入Y或N."
      Case 9034
         MsgText = "請輸入Y或空白."
      Case 9035
         MsgText = "請輸入1-100%或空白"
      Case 9036
         MsgText = "請輸入1-3或空白."
      Case 9037
         MsgText = "請檢查輸入資料;至少輸入一種語文別名稱"
      Case 9038
         MsgText = "請輸入1-2或空白."
      Case 9039
         MsgText = "請檢查輸入資料;至少輸入一筆名稱"
      Case 9040
         MsgText = "輸入資料重複"
      Case 9041
         MsgText = "該系統類別非服務業務"
      Case 9042
         MsgText = "請輸入6位數字"
      Case 9043
         MsgText = "請輸入2位數字"
      Case 9044
         MsgText = "請輸入N或空白."
      Case 9045
         MsgText = "請輸入數字"
      Case 9046
         MsgText = "請輸入1-5或空白."
      Case 9047
         MsgText = "法定期限必須大於或等於本所期限."
      Case 9048
         MsgText = "請輸入Y或N或空白."
      Case 9049
         
      Case 9050
         MsgText = "解除期限原因代碼錯誤."
      Case 9051
         MsgText = "代理人代碼錯誤!!"
      Case 9052
         MsgText = "自動編號錯誤!!"
      Case 9053
         MsgText = "業務區代碼錯誤!!"
      Case 9054
         MsgText = "欲輸入資料已存在!!"
      Case 9055
         MsgText = "使用者權限有問題,請洽系統管理者!"
      Case 9056
         MsgText = "專利/商標種類代碼錯誤!!"
      Case 9057
         MsgText = "報表已列印出."
      Case 9058
         MsgText = "請點選欲維護之資料!!"
      Case 9059
         MsgText = "請輸入1-5"
      Case 9060
         MsgText = "請輸入1-2"
      Case 9101
         MsgText = "你沒有使用任何系統類別的權利"
      Case 9102
         MsgText = "分析系統類別字串時發生錯誤"
      Case 9103
         MsgText = "無法新增或更新申請人國外ID對照檔"
      Case 9104
         MsgText = "此國家沒有此項專利種類"
      Case 9105
         MsgText = "讀取國家基本檔時發生錯誤"
      Case 9106
         MsgText = "讀取Patent檔時發生錯誤"
      Case 9107
         MsgText = "讀取Trademark檔時發生錯誤"
      Case 9108
         MsgText = "讀取LawCase檔時發生錯誤"
      Case 9109
         MsgText = "讀取HireCase檔時發生錯誤"
      Case 9110
         MsgText = "讀取ServicePractice檔時發生錯誤"
      Case 9111
         MsgText = "讀取CaseProgress檔時發生錯誤"
      Case 9112
         MsgText = "儲存Patent檔時發生錯誤"
      Case 9113
         MsgText = "儲存Trademark檔時發生錯誤"
      Case 9114
         MsgText = "儲存LawCase檔時發生錯誤"
      Case 9115
         MsgText = "儲存HireCase檔時發生錯誤"
      Case 9116
         MsgText = "儲存ServicePractice檔時發生錯誤"
      Case 9117
         MsgText = "儲存CaseProgress檔時發生錯誤"
      Case 9118
         MsgText = "讀取案件進度檔之代理人之彼所案號時發生錯誤"
      Case 9119
         MsgText = "讀取案件進度檔之代理人時發生錯誤"
      Case 9120
         MsgText = "讀取案件收費表時發生錯誤"
      Case 9121
         MsgText = "無法讀取優先權資料"
      Case 9122
         MsgText = "無法刪除優先權資料"
      Case 9123
         MsgText = "無法新增優先權資料"
      Case 9124
         MsgText = "無法讀取指定國家資料"
      Case 9125
         MsgText = "無法刪除指定國家資料"
      Case 9126
         MsgText = "無法新增指定國家資料"
      Case 9127
         MsgText = "此本所案號已無任何收文號資料"
      Case 9128
         MsgText = "無法轉本所案號"
      Case 9129
         MsgText = "無法新增資料至ChangeEvent基本檔"
      Case 9130
         MsgText = "無法新增資料至ServicePractice基本檔"
      Case 9131
         MsgText = "無法新增資料至Patent基本檔"
      Case 9132
         MsgText = "無法新增資料至Trademark基本檔"
      Case 9133
         MsgText = "新增資料至CaseProgress進度檔時,無法產生流水號"
      Case 9134
         MsgText = "無法新增資料至CaseProgress進度檔"
      Case 9135
         MsgText = "無法讀取下一程序檔"
      Case 9136
         MsgText = "無法更新下一程序之本所期限及法定期限"
      Case 9137
         MsgText = "無法更新下一程序之是否續辦欄"
      Case 9138
         MsgText = "無法新增資料至NextProgress檔"
      Case 9139
         MsgText = "找不到此本所案號之收文號"
      Case 9140
         MsgText = "ServicePractice無法存檔"
      Case 9141
         MsgText = "找不到此本所案號之資料"
      Case 9142
         MsgText = "CaseProgress無法存檔"
      Case 9143
         MsgText = "Patent無法存檔"
      Case 9144
         MsgText = "LawCase無法存檔"
      Case 9145
         MsgText = "HireCase無法存檔"
      Case 9146
         MsgText = "Trademark無法存檔"
      Case 9147
         MsgText = "找不到此收文號之資料"
      Case 9148
         MsgText = "找不到此系統類別"
      Case 9149
         MsgText = "員工已離職"
      Case 9150
         MsgText = "員工代碼錯誤"
      Case 9151
         MsgText = "申請人代碼錯誤"
      Case 9152
         MsgText = "代理人代碼錯誤"
      Case 9153
         MsgText = "國籍代碼錯誤"
      Case 9154
         MsgText = "此系統代號不存在"
      Case 9155
         MsgText = "找不到此代號的流水號"
      Case 9156
         MsgText = "該國家無此專利種類"
      Case 9157
         MsgText = "在Nation中無此國家代號"
      Case 9158
         MsgText = "種類代碼錯誤"
      Case 9159
         MsgText = "本所案號錯誤"
      Case 9160
         MsgText = "案件性質代碼錯誤"
      Case 9161
         MsgText = "收文號不存在"
      Case 9162
         MsgText = "讀取申請人國外ID對照時，發生錯誤"
      Case 9163
         MsgText = "案件來源代碼錯誤"
      Case 9164
         MsgText = "費用資料與檔案不符"
      Case 9165
         MsgText = "使用者已離職，請重新登入網域"
      Case 9166
         MsgText = "無效之使用者，請重新登入網域"
      Case 9167
         MsgText = "無法取得NT網域之登入使用者名稱，請重新登入網域"
      Case 9168
         MsgText = "找不到案件起算日"
      Case 9169
         MsgText = "當開始日輸入時，結束日一定要輸入"
      Case 9170
         MsgText = "開始日必須小於結束日"
      Case 9171
         MsgText = "輸入之系統類別不符,請依照子系統輸入允許之系統類別"
      Case 9172
         MsgText = "本所案號輸入長度不符"
      Case 9173
         MsgText = "本所案號必須為同系統類別"
      Case 9174
         MsgText = "請輸入Y或不輸入任何字"
      Case 9175
         MsgText = "卷宗性質代碼不存在"
      Case 9176
         MsgText = "此收文號之本所案號不存在"
      Case 9177
         MsgText = "請輸入Y、N或不輸入任何字"
      Case 9178
         MsgText = "原本所案號為"
      Case 9179
         MsgText = "請自行去更新原本所案號之下一程序資料內容"
      Case 9180
         MsgText = "指定國家尚未輸入"
      Case 9181
         MsgText = "轉本所案號與原本之本所案號相同"
      Case 9182
         MsgText = "流水號不可大於自動編號"
      Case 9183
         MsgText = "此案號為新本所案號"
      Case 9184
         MsgText = "與國內案號相同之欄位未輸入"
      Case 9185
         MsgText = "本所案號必須為國內系統類別"
      Case 9186
         MsgText = "因為轉本所案號，所以無法再做指定國家"
      Case 9187
         MsgText = "並沒有選擇變更事項，無法存檔"
      Case 9188
         MsgText = "請輸入Y、N"
      Case 9189
         MsgText = "請輸入數字或不輸入任何字"
      Case 9190
         MsgText = "你必須輸入申請優先權證明書"
      Case 9191
         MsgText = "並沒有輸入指定國家之費用，無法存檔"
      Case 9192
         MsgText = "因為為美國案且列印指示信，所以此欄位不可空白"
      Case 9193
         MsgText = "子案全都無證書號數"
      Case 9194
         MsgText = "年（次）錯誤"
      Case 9195
         MsgText = "請輸入年（次），必須為數字"
      Case 9196
         MsgText = "請輸入1或2"
      Case 9197
         MsgText = "如輸入期間，則必須全部輸入"
      Case 9198
         MsgText = "功能代號不存在"
      Case 9199
         MsgText = "重複"
      Case 9200
         MsgText = "優先權號資料重複"
      Case 9201
         MsgText = "再次輸入之值與原本不符"
      Case 9202
         MsgText = "你選擇的發明人超過了10個"
      Case 9203
         MsgText = "你已勾選此項變更，此欄位一定要輸入"
      Case 9204
         MsgText = "找不到案件起算日"
      Case 9205
         MsgText = "輸入之資料過長，超過"
      Case 9206
         MsgText = "個字（註：中文算兩個字)"
      Case 9207
         MsgText = "錯誤號碼："
      Case 9208
         MsgText = "錯誤敘述："
      Case 9209
         MsgText = "國內案件已送達"
      Case 9210
         MsgText = "延期後本所期限必須≦延期後法定期限"
      Case 9211
         MsgText = "資料庫無資料"
      Case 9212
         MsgText = "無此機關代號資料"
      Case 9213
         MsgText = "部門代號錯誤"
      Case 9214
         MsgText = "輸入之流水號大於目前最大之流水號"
      Case 9215
         MsgText = "此本所案號"
      Case 9216
         MsgText = "已存在"
      Case 9217
         MsgText = "不存在"
      Case 9218
         MsgText = "員工代號不存在"
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/30 +ACS系統類別
      Case 9219
         MsgText = "FCP,FG 或 FCT 或 FCL,LIN,ACS 案件, 申請國家必須為 000 !"
      Case 9220
        MsgText = "不可輸入 EPC！"
      'Add By Sindy 2013/9/6
      Case 9221
         MsgText = " 檔案插入有誤，因檔案大小為 0 KB！" & vbCrLf & _
                   "(請檢查該檔案是否正使用於其他應用程式，請關閉後，" & vbCrLf & _
                   " 再重新插入檔案)"
   End Select
End Function

'Added by Morgan 2021/4/13
Public Function GetHWndForToolTip(ByVal ctl As Object) As Long
    ' This returns the control's hWnd, or it keeps crawling up
    ' into containers until it finds one with a valid hWnd.
    ' Ultimately, the top-level form always has a hWnd.
    On Error Resume Next
        Do
            GetHWndForToolTip = ctl.hWnd
            If Err = 0 Then Exit Do
            Err.Clear
            Set ctl = ctl.Container                 ' This will still work for controls nested on a container on a UC.
            If Err <> 0 Then                        ' Apparently it's directly on a UC.  At this point, Container will be the UC (if Err).
                Err.Clear
                Set ctl = ctl.Extender.Container    ' Extender must be exposed for this to work.
                If Err <> 0 Then
                    On Error GoTo 0
                    Error 438                       ' Apparently it's on a UC and the Extender wasn't exposed.
                    Exit Function
                End If
            End If
        Loop
    On Error GoTo 0
End Function

Public Sub CreateToolTip(ByVal ParentHwnd As Long, _
                         ByVal TipText As String, _
                         Optional ByVal uIcon As ttIconType = TTNoIcon, _
                         Optional ByVal sTitle As String, _
                         Optional ByVal lForeColor As Long = -1&, _
                         Optional ByVal lBackColor As Long = -1&, _
                         Optional ByVal bCentered As Boolean, _
                         Optional ByVal bBalloon As Boolean, _
                         Optional ByVal lWrapTextLength As Long = 50&, _
                         Optional ByVal lDelayTime As Long = 200&, _
                         Optional ByVal lVisibleTime As Long = 5000&)
    '
    ' If lWrapTextLength = 0 then there will be no wrap.
    ' Also, lWrapTextLength = 40 is a minimum value.
    ' The max for lVisibleTime is 32767.
    '
    Static bCommonControlsInitialized As Boolean
    Dim lWinStyle As Long
    Dim ti As TOOLINFO
    Static PrevParentHwnd As Long
    Static PrevTipText As String
    Static PrevTitle As String
    '
    ' Don't do anything unless we need to.
    If hwndTT <> 0 And ParentHwnd = PrevParentHwnd And TipText = PrevTipText And sTitle = PrevTitle Then Exit Sub
    '
    If Not bCommonControlsInitialized Then
        InitCommonControls
        bCommonControlsInitialized = True
    End If
    '
    ' Destroy any previous tooltip.
    If hwndTT <> 0 Then DestroyWindow hwndTT
    '
    ' Format the text.
    FormatTooltipText TipText, lWrapTextLength
    '
    ' Initial style settings.
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    If bBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON ' Create baloon style if desired.
    ' Set the style.
    hwndTT = CreateWindowExW(0&, StrPtr(TOOLTIPS_CLASS), 0&, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0&, 0&, App.hInstance, 0&)
    '
    ' Setup our tooltip info structure.
    ti.lFlags = TTF_SUBCLASS Or TTF_IDISHWND
    If bCentered Then ti.lFlags = ti.lFlags Or TTF_CENTERTIP
    ' Set the hwnd prop to our parent control's hwnd.
    ti.hWnd = ParentHwnd
    ti.lId = ParentHwnd
    ti.hInstance = App.hInstance
    ti.lpStr = TipText
    ti.lSize = LenB(ti)
    ' Set the tooltip structure
    SendMessageLong hwndTT, TTM_ADDTOOLW, 0&, VarPtr(ti)
    SendMessageLong hwndTT, TTM_UPDATETIPTEXTW, 0&, VarPtr(ti)
    '
    ' Colors.
    If lForeColor <> -1 Then SendMessage hwndTT, TTM_SETTIPTEXTCOLOR, lForeColor, 0&
    If lBackColor <> -1 Then SendMessage hwndTT, TTM_SETTIPBKCOLOR, lBackColor, 0&
    '
    ' Title or icon.
    If uIcon <> TTNoIcon Or sTitle <> vbNullString Then SendMessageLong hwndTT, TTM_SETTITLEW, CLng(uIcon), StrPtr(sTitle)
    '
    SendMessageLong hwndTT, TTM_SETDELAYTIME, TTDT_AUTOPOP, lVisibleTime
    SendMessageLong hwndTT, TTM_SETDELAYTIME, TTDT_INITIAL, lDelayTime
    '
    PrevParentHwnd = ParentHwnd
    PrevTipText = TipText
    PrevTitle = sTitle
End Sub

Public Sub DestroyToolTip()
    ' It's not a bad idea to put this in the Form_Unload event just to make sure.
    If hwndTT <> 0 Then DestroyWindow hwndTT
    hwndTT = 0
End Sub

Private Sub FormatTooltipText(TipText As String, LLen As Long)
    Dim s As String
    Dim i As Long
    '
    ' Make sure we need to do anything.
    If LLen = 0 Then Exit Sub
    If LLen < 40 Then LLen = 40
    If Len(TipText) <= LLen Then Exit Sub
    '
    Do
        i = InStrRev(TipText, " ", LLen + 1)
        If i = 0 Then
            s = s & Left$(TipText, LLen) & vbCrLf ' Build "s" and trim from TipText.
            TipText = Mid$(TipText, LLen + 1)
        Else
            s = s & Left$(TipText, i - 1) & vbCrLf ' Build "s" and trim from TipText.
            TipText = Mid$(TipText, i + 1)
        End If
        If Len(TipText) <= LLen Then
            TipText = s & TipText ' Place "s" back into TipText and get out.
            Exit Sub
        End If
    Loop
End Sub
'end 2021/4/13

'Added by Morgan 2021/4/14
'可顯示 Unicode 的對話框(可替代 MsgBox)
Public Function UniMsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String, Optional ByVal ResourceIcon As String, Optional ByVal hwndOwner As Long) As VbMsgBoxResult
    Dim udtMsgBox As MsgBoxParams
    ' if no owner is specified, try to use the active form
    If hwndOwner = 0 Then If Not Screen.ActiveForm Is Nothing Then hwndOwner = Screen.ActiveForm.hWnd
    With udtMsgBox
        .cbSize = Len(udtMsgBox)
        ' important to set owner to get behavior similar to the native MsgBox
        .hwndOwner = hwndOwner
        .hInstance = App.hInstance
        ' set the message
        .lpszText = StrPtr(Prompt)
        ' if no title is given, use the application title like the native MsgBox
        If LenB(Title) = 0 Then Title = App.Title
        .lpszCaption = StrPtr(Title)
        ' thought this would be a nice feature addition
        If LenB(ResourceIcon) = 0& Then
            .dwStyle = Buttons
        Else
            .dwStyle = (Buttons Or MB_USERICON) And Not (&H70&)
            .lpszIcon = StrPtr(ResourceIcon)
        End If
    End With
    ' show the message box
    UniMsgBox = MessageBoxIndirectW(udtMsgBox)
End Function
'Added by Morgan 2022/8/3
Public Function MsgBoxU(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String, Optional ByVal ResourceIcon As String, Optional ByVal hwndOwner As Long) As VbMsgBoxResult
   MsgBoxU = UniMsgBox(Prompt, Buttons, Title, ResourceIcon, hwndOwner)
End Function


'Added by Morgan 2021/4/14
'Form 2.0 上線前已改好的表單要加此檢查(上線後可直接回傳True後跳離)
'檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
'pForm:待檢查的表單, pShowMsg:是否顯示錯誤訊息, pAskUser:是否詢問後可繼續, pTypeName：指定要檢查的物件類別名稱
'Modified by Lydia 2021/09/30 +不列入檢查的名稱
'Modified by Sindy 2021/11/2 + , Optional pIsReplace As Boolean = False: 直接要用?取代Unicode文字(信件主旨)
Public Function PUB_ChkUniText(pForm As Form, Optional pShowMsg As Boolean = True, _
   Optional pAskUser As Boolean = False, Optional pTypeName As String = "", _
   Optional pJumpName As String, Optional pIsReplace As Boolean = False) As Boolean
   
   Dim oCtrl As Control
   Dim strText As String
   Dim strType As String
   Dim strMsg As String
     
   PUB_ChkUniText = True
   
   'Added by Morgan 2022/5/5
   '為避免資料已轉Unicode但物件漏改造成回存成 ?，增加 ? 檢查。
   For Each oCtrl In pForm.Controls
      strType = TypeName(oCtrl)
      If (strType = "TextBox" Or strType = "ComboBox") And (strType = pTypeName Or pTypeName = "") And _
               (pJumpName = "" Or (pJumpName <> "" And InStr(UCase(pJumpName), UCase(oCtrl.Name) = 0))) Then
         If oCtrl.Enabled = True And oCtrl.Locked = False Then
            If InStr(oCtrl.Text, "?") > 0 Then
               strMsg = "欄位中有？號，若非原始資料，請通知電腦中心人員！" & vbCrLf & vbCrLf & "欄位內容：" & oCtrl.Text & vbCrLf & vbCrLf & "【 是:繼續存檔   否:回畫面 】"
               If UniMsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  PUB_ChkUniText = False
                  Exit For
               End If
            End If
         End If
      End If
   Next
   'end 2022/5/5
   
   If strSrvDate(1) >= Form20上線日 Then Exit Function 'Added by Morgan 2022/3/2
   
   'Memo by Lydia 2021/12/27 在\\LINUX\PolyCOM\TaieNew\定稿\User\ModMenu 有另外寫同樣的模組，若要變更設定請一併修改
   
   'Added by Lydia 2021/08/20 因為O8的Provider在存檔時，不能接受Unicode(ex. Account存檔使用Recordset.Value = TextBox會程式出錯)
                                          '所以判斷Provider，現在只對M51=O12Provider有詢問的功能
   'Mark by Lydia 2021/09/02 改成詢問後,自動更換
   'If strProvider = cProvider Then
   '    pAskUser = False
   'End If
   'end 2021/08/20
   
   For Each oCtrl In pForm.Controls
      strType = TypeName(oCtrl)
      'Modified by Lydia 2021/09/30 增加判斷不列入檢查的名稱
      'If (strType = "TextBox" Or strType = "ComboBox") And (strType = pTypeName Or pTypeName = "") Then
      If (strType = "TextBox" Or strType = "ComboBox") And (strType = pTypeName Or pTypeName = "") And _
               (pJumpName = "" Or (pJumpName <> "" And InStr(UCase(pJumpName), UCase(oCtrl.Name) = 0))) Then
         If oCtrl.Enabled = True And oCtrl.Locked = False Then
            strText = oCtrl.Text
            strText = StrConv(StrConv(strText, vbFromUnicode), vbUnicode)
            If strText <> oCtrl.Text Then
               'Add By Sindy 2021/11/2
               If pIsReplace = True Then
                  oCtrl.Text = strText '自動更換
               '2021/11/2 END
               ElseIf pAskUser = True Then
                  strMsg = "欄位含Unicode字元，存檔時將會以 ? 號取代！" & vbCrLf & vbCrLf & "原文字：" & oCtrl.Text & vbCrLf & "存檔後：" & strText & vbCrLf & vbCrLf & "是否要繼續？"
                  oCtrl.Text = strText 'Added by Lydia 2021/09/02 改成詢問後,自動更換
                  If UniMsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                     PUB_ChkUniText = False
                     Exit For
                  End If
               Else
                  If pShowMsg = True Then
                     strMsg = "欄位含Unicode字元，不可存檔！" & vbCrLf & vbCrLf & "內容：" & oCtrl.Text
                     UniMsgBox strMsg, vbCritical
                  End If
                  PUB_ChkUniText = False
                  Exit For
               End If
            End If
         End If
      End If
   Next
End Function

'Added by Lydia 2018/11/21 Form 2.0的TextBox按Enter換行,下一行會被隱藏,要重新focus才會出現
'Move by Lydia 2021/04/14 從basUpdate移過來
Public Sub PUB_HandleForm2TextBox(ByRef pTb01, ByRef pSetPass, ByVal pKeyCode As Integer, ByVal pShift As Integer)
    If pShift = 2 And pKeyCode = 90 Then
         MsgBox "在新版輸入畫面中，Ctrl鍵+Z無效 !", vbInformation, "Form 2.0 輸入"
         
    ElseIf pTb01.MultiLine = True Then
    '當TextBox.MultiLine = True，請記得TextBox.EnterKeyBehavior=True，才能用Enter鍵換行；
    '為了避免重新focus的全選反白造成使用者的疑問，請記得在原程式的全選反白排除該TextBox
        'BackSpace鍵=8, Enter鍵=13, Delete鍵=46,Ctrl+C=(shift=2 and Keycode=67), Ctrl+V,Ctrl+X
        If pKeyCode = 8 Or pKeyCode = 13 Or pKeyCode = 46 Or (pShift = 2 And (pKeyCode = 67 Or pKeyCode = 86 Or pKeyCode = 88)) Then
             If pSetPass.Enabled = True Then pSetPass.SetFocus
             pTb01.SetFocus
        End If
    End If
End Sub

'Added by Lydia 2018/11/21 Form 2.0的TextBox點選滑鼠右鍵,呼叫快捷選單無效,改成彈訊息
'Move by Lydia 2021/04/14 從basUpdate移過來
Public Sub PUB_HandleForm2TextBoxR(ByVal pButton As Integer, ByVal pShift As Integer, Optional ByRef bMsg As Boolean = False)
Dim strDesc As String

    If pButton = 2 And bMsg = False Then
       strDesc = "在新版輸入畫面中，用滑鼠呼叫選單功能失效 !" & vbCrLf
       strDesc = strDesc & "請使用下列快速鍵組合：" & vbCrLf
       strDesc = strDesc & "1. 複製：Ctrl+C鍵" & vbCrLf
       strDesc = strDesc & "2. 貼上：Ctrl+V鍵" & vbCrLf
       strDesc = strDesc & "3. 剪下：Ctrl+X鍵" & vbCrLf
       strDesc = strDesc & "4. 全選：Ctrl+A鍵" & vbCrLf
       bMsg = True  '只彈一次
       MsgBox strDesc, vbInformation, "Form 2.0 輸入"
    End If
End Sub

'Added by Lydia 2021/03/18 Form2.0 控制Function鍵：記錄鍵盤傳入順序
'Move by Lydia 2021/04/14 從acc_var(Finance)移過來
Public Sub PUB_SaveTrackMode(ByVal inMode As Integer, ByVal inCode As Integer)
'P.S. 注意Form_Unload 要清空「記錄鍵盤傳入順序」，另外在用frmacc000最上方Toolbar的功能按鈕也要清空

'inMode: 0-Form_KeyDown階段, 1-Form_KeyUp階段(確定執行)
'inCode: KeyCode值
   If inCode = vbKeyF5 Or inCode = vbKeyF2 Or inCode = vbKeyF3 Or inCode = vbKeyF4 Or inCode = vbKeyF9 Or inCode = vbKeyF10 Or inCode = vbKeyF12 Or _
          inCode = vbKeyEscape Or inCode = vbKeyPageUp Or inCode = vbKeyPageDown Or inCode = vbKeyHome Or inCode = vbKeyEnd Or inCode = vbKeyInsert Then
        If inMode = 0 Then 'Form_KeyDown階段
            strTrackMode = "KeyDown"
        ElseIf inMode = 1 Then 'Form_KeyUp階段
            If InStr(strTrackMode & ";", cntTrackRun & inCode & ";") = 0 Then
                 strTrackMode = cntTrackRun & inCode & ";"  'Form2.0 記錄鍵盤傳入順序
            Else
                 strTrackMode = "KeyPass" 'Form2.0 記錄鍵盤傳入順序(重複->不執行)
            End If
        End If
   Else
        strTrackMode = ""
   End If
End Sub

'Added by Lydia 2021/03/18 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
'Move by Lydia 2021/04/14 從acc_var(Finance)移過來
Public Function PUB_ChkTrackMode() As Boolean
'KeyCode的觸發階段順序為 Form_KeyCode -> Form_KeyUp -> Form_KeyPress；
'Form1.0只會單純觸發程式撰寫的Form_KeyUp階段，
'不知為何Form2.0的表單在Form_KeyCode階段和Form_KeyUp階段皆會觸發KeyEnter，所以Form2.0的表單要在Form_KeyCode階段和Form_KeyUp階段分別記錄鍵盤傳入順序，當執行共同模組KeyEnter時，判斷在Form_KeyUp階段才執行。
'P.S 依帳務->財產目錄作業之觀察有時傳入KeyCode的次數不只一次，但是KeyEnter原本就有一些判斷所以不會每個動作都可以感覺到重複操作，目前比較有可能影響的是vbKeyInsert和vbKeyF5
        
    If strTrackMode <> "" And InStr(strTrackMode & ";", cntTrackRun) = 0 Then
        Exit Function
    End If
    
    PUB_ChkTrackMode = True
End Function


'Added by Lydia 2021/10/20 (案件系統使用)Form2.0 控制Function鍵：記錄鍵盤傳入順序
Public Sub PUB_SaveMeTrackMode(ByRef MeTrackMode As String, ByVal inMode As Integer, ByVal inCode As Integer)
'MeTrackMode: 表單變數
'inMode: 0-Form_KeyDown階段, 1-Form_KeyUp階段(確定執行)
'inCode: KeyCode值
   If inCode = vbKeyF5 Or inCode = vbKeyF2 Or inCode = vbKeyF3 Or inCode = vbKeyF4 Or inCode = vbKeyF9 Or inCode = vbKeyF10 Or inCode = vbKeyF12 Or _
          inCode = vbKeyEscape Or inCode = vbKeyPageUp Or inCode = vbKeyPageDown Or inCode = vbKeyHome Or inCode = vbKeyEnd Or inCode = vbKeyInsert Then
        If inMode = 0 Then 'Form_KeyDown階段
            MeTrackMode = "KeyDown"
        ElseIf inMode = 1 Then 'Form_KeyUp階段
            If InStr(MeTrackMode & ";", cntTrackRun & inCode & ";") = 0 Then
                 MeTrackMode = cntTrackRun & inCode & ";"  'Form2.0 記錄鍵盤傳入順序
            Else
                 MeTrackMode = "KeyPass" 'Form2.0 記錄鍵盤傳入順序(重複->不執行)
            End If
        End If
   Else
        MeTrackMode = ""
   End If
End Sub

'Added by Lydia 2021/10/20 (案件系統使用)Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
Public Function PUB_ChkMeTrackMode(ByRef MeTrackMode As String) As Boolean
'MeTrackMode: 表單變數
    If MeTrackMode <> "" And InStr(MeTrackMode & ";", cntTrackRun) = 0 Then
        Exit Function
    End If
    PUB_ChkMeTrackMode = True
End Function

'Added by Lydia 2021/10/20 (案件系統使用)Form2.0 控制ToolBar：記錄鍵盤傳入順序
Public Function Pub_SaveMeToolBar(ByRef MeTrackMode As String, ByRef MeToolbar As Control, ByVal pIndex As Integer)
    '若有交錯使用Function鍵和Toolbar鍵會失去記錄造成無法判斷，所以ToolBar鍵另外記錄
    Dim strType As String
    Dim intKeyCode As Integer
    
     strType = TypeName(MeToolbar)
     If UCase(strType) = UCase("ToolBar") Then
         Select Case pIndex
             Case 1 '新增
                 intKeyCode = vbKeyF2
             Case 2 '修改
                 intKeyCode = vbKeyF3
             Case 3 '刪除
                 intKeyCode = vbKeyF5
             Case 4 '查詢
                 intKeyCode = vbKeyF4
             Case 6 '第一筆
                 intKeyCode = vbKeyHome
             Case 7 '前一筆
                 intKeyCode = vbKeyPageUp
             Case 8 '後一筆
                 intKeyCode = vbKeyPageDown
             Case 9 '末筆
                 intKeyCode = vbKeyEnd
             Case 11 '確定
                 intKeyCode = vbKeyF9
             Case 12 '取消
                 intKeyCode = vbKeyEnd
         End Select
         If intKeyCode <> 0 Then
             Call PUB_SaveMeTrackMode(MeTrackMode, 1, intKeyCode)
         End If
     End If
End Function

'Added by Morgan 2021/4/22
'取得指定index的陣列字串值(逗號區隔)
Public Function PUB_GetItemData(ByVal pItemDataList As String, ByVal pIndex As Integer) As String
   Dim tmpArr() As String
     
   PUB_GetItemData = ""
   If pItemDataList = "" Or pIndex < 0 Then Exit Function
   
   tmpArr = Split(pItemDataList, ",")
   If pIndex <= UBound(tmpArr) Then
       PUB_GetItemData = tmpArr(pIndex)
   End If
End Function

'Added by Lydia 2022/01/22 去掉不可為檔名的符號
Public Function Pub_RepFileName(ByVal pOldName As String) As String
Dim pTempText As String
Dim intP As Integer, ICode As Integer, m_Text As String 'Added by Lydia 2022/01/26
    
    'Unicode會轉?
    pTempText = StrConv(StrConv(pOldName, vbFromUnicode), vbUnicode)
    '去掉不可為檔名的符號 ? \ / : < > | "=chr(34)
    'Modified by Lydia 2022/01/25 去掉*
    pTempText = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(pTempText, "?", ""), "/", ""), "\", ""), ":", ""), "<", ""), ">", ""), "|", ""), Chr(34), ""), "*", "")
    'Modified by Lydia 2022/01/26
    'Pub_RepFileName = pTempText
    If Len(pTempText) = 0 Then
        Pub_RepFileName = ""
    Else
        For intP = 1 To Len(pTempText)
            m_Text = Mid(pTempText, intP, 1)
            If m_Text <> "" Then
                ICode = Asc(m_Text)
                '非中文: 只保留英數字, 中文字保留
                If ICode > 0 And ICode <= 255 Then
                    If (ICode >= 48 And ICode <= 57) Or (ICode >= 65 And ICode <= 90) Or (ICode >= 97 And ICode <= 122) _
                             Or ICode = 45 Or ICode = 46 Or ICode = 95 Or ICode = 32 Then
                        Pub_RepFileName = Pub_RepFileName & m_Text
                    End If
                ElseIf ICode < 0 Then
                    Pub_RepFileName = Pub_RepFileName & m_Text
                End If
            End If
        Next intP
    End If
    'end 2022/01/26
    
End Function

'Added by Lydia 2022/03/28 將Unicode文字轉為BIG5可支援的文字
Public Function PUB_UniToBIG5(ByVal strText As String, Optional ByVal pType As String) As String
Dim strTemp As String

    strTemp = StrConv(StrConv(strText, vbFromUnicode), vbUnicode)
    If pType = "F" Then '檔名處理=>去掉?
        strTemp = Replace(strTemp, "?", "")
    End If
    PUB_UniToBIG5 = strTemp

End Function

'Added by Morgan 2022/7/4
'讀取程式內UniCode文字
Public Function PUB_GetUniText(pFormName As String, pVarName As String) As String
   Dim strSql As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   intQ = 1
   strSql = "select UTL03,1 Srt from UniTextList where upper(UTL01)=upper('" & pFormName & "')  and UTL02='" & pVarName & "'"
   strSql = strSql & " union select UTL03,2 Srt from UniTextList where UTL01='共用'  and UTL02='" & pVarName & "' order by Srt"
   Set rsQuery = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
      PUB_GetUniText = "" & rsQuery("UTL03")
   End If
   
   Set rsQuery = Nothing
End Function

'Added by Lydia 2022/07/15 模擬送出鍵盤的指令
Public Sub PUB_SendSKey(ByVal KeyName As String)
    Select Case UCase(KeyName)
        Case "KEYINSERT"
             '舉例:Finance的傳票輸入Frmacc4120 編輯明細要按Insert才會更新資料，但是Form2.0元件支援Insert鍵會切換”新增/覆寫模式”Insert/OverType
                      '在按下Insert鍵時先重送Insert鍵於第2次才執行更新明細
             keybd_event vbKeyInsert, 0, 0, 0  '模擬KeyPress
             'keybd_event vbKeyInsert, 0, &H0, 0  '模擬KeyDown   ---不需要
             keybd_event vbKeyInsert, 0, &H2, 0   '模擬KeyUp
    End Select
End Sub

'Added by Lydia 2022/10/12 特殊情況之指定職代
'Modified by Lydia 2022/10/13
Public Function PUB_GetStateForMan(ByVal pOldNo As String, Optional ByVal pType As String) As String

   PUB_GetStateForMan = pOldNo

   If strSrvDate(1) >= "20221013" And strSrvDate(1) <= "20221019" Then
       'Added by Lydia 2022/10/13 取得被代理之主管編號
       If pType = "A" Then
           If pOldNo = "A1034" Then
               PUB_GetStateForMan = "88003"
           End If
           If pOldNo = "96008" Then
               PUB_GetStateForMan = "94012"
           End If
       Else
       'end 2022/10/13
          '111/10/13~111/10/19 日本部王協理88003與簡偉倫經理99037、林軒吉副理94012將於赴韓國參加APAA年會
          '日本部新案命名依下述方式設定系統處理:
          '蕭人瑄副理A1034暫代平常王協理88003的角色，化學類案件由其（代理簡偉倫）分至命名同仁，機電類則由其轉至何立中96008主任（代理林軒吉），然後分至命名同仁。
          If InStr(pOldNo, "88003") > 0 Or InStr(pOldNo, "99037") > 0 Then
               PUB_GetStateForMan = Replace(PUB_GetStateForMan, "88003", "A1034")
               PUB_GetStateForMan = Replace(PUB_GetStateForMan, "99037", "A1034")
          End If
          If InStr(pOldNo, "94012") > 0 Then
               PUB_GetStateForMan = Replace(PUB_GetStateForMan, "94012", "96008")
          End If
       End If 'Added by Lydia 2022/10/13
   End If
   'Added by Lydia 2022/11/02 因簡偉倫與王協理88003一起出差(11/6～11/19)，請以林軒吉94012為我的職務代理人，特別是新案命名通知，一定要寄給他。
   If strSrvDate(1) >= "20221106" And strSrvDate(1) <= "20221119" Then
       If pType = "A" Then
           If pOldNo = "94012" Then
               PUB_GetStateForMan = "88003"
           End If
       Else
          If InStr(pOldNo, "88003") > 0 Then
               PUB_GetStateForMan = Replace(PUB_GetStateForMan, "88003", "94012")
          End If
       End If
   End If
   'end 2022/11/02
   'Added by Lydia 2024/02/27 外專機械設計組人員異動調整程式：新案認領組別，請取消機械設計組，案件歸電子組; 由Wilson代機械組主管T1
   If strSrvDate(1) >= "20240229" Then
      If pOldNo = "89020" Then
         PUB_GetStateForMan = "87003"
      End If
   End If
   'end 2024/02/27
End Function

'Move by Lydia 2023/04/27 從basQuery搬過來
'Add by Morgan 2008/5/27
'讀取電子檔存放路徑
'Modified by Morgan 2024/9/3 +p_UserSt03:指定使用者部門
Public Function PUB_GetEFilePath(ByVal p_System As String, Optional p_UserSt03 As String) As String
   
   Dim strST03 As String
   If p_UserSt03 <> "" Then
      strST03 = p_UserSt03
   Else
      strST03 = Pub_StrUserSt03
   End If
      
   'Modified by Morgan 2023/3/23 測試或VB也不可放正式資料夾
   'Modified by Morgan 2024/9/3 有指定使用者部門時要抓該部們實際會存放的位置(請款單月報表複製檔案要用)
   If p_UserSt03 = "" And (strST03 = "M51" Or UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0) Then
      PUB_GetEFilePath = PUB_Getdesktop
   Else
'      Select Case p_System
'         'FCP
'         Case "FCP", "FG", "P", "PS", "CFP", "CPS"
'            PUB_GetEFilePath = EFilePath
'         'FCL
'         Case "FCL", "LIN"
'            PUB_GetEFilePath = FCLeFilePath
'         'Add By Sindy 2012/1/11
'         Case "T", "TF"
'            PUB_GetEFilePath = PUB_Getdesktop & TeFilePath
'         '2012/1/11 End
'         'FCT
'         Case Else
'            PUB_GetEFilePath = FCTeFilePath
'      End Select
      'Modify By Sindy 2012/2/16 原以判斷系統別丟檔案,改以操作人員的部門別決定檔案存放位置
      If Left(strST03, 2) = "F2" Then
         PUB_GetEFilePath = EFilePath
      ElseIf Left(strST03, 2) = "F1" Then
         PUB_GetEFilePath = FCTeFilePath
      ElseIf Left(strST03, 1) = "F" Then
         PUB_GetEFilePath = FCLeFilePath
      Else
         PUB_GetEFilePath = PUB_Getdesktop & TeFilePath
         If Dir(PUB_GetEFilePath, vbDirectory) = "" Then
            MkDir PUB_GetEFilePath
         End If
      End If
   End If
End Function

'Added by Lydia 2023/04/27 建立預設共用資料夾的電子檔存放路徑
Public Function Pub_GetEFilePath_All(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String) As String
Dim oFileSys
Dim strDefPath  As String

On Error GoTo ErrHandle
   Set oFileSys = CreateObject("Scripting.FileSystemObject")
   strDefPath = PUB_GetEFilePath(pCP01) & "\" & pCP01
   If Not oFileSys.FolderExists(strDefPath) Then
      MkDir strDefPath
   End If
   strDefPath = strDefPath & "\" & Left(pCP02, 3)
   If Not oFileSys.FolderExists(strDefPath) Then
      MkDir strDefPath
   End If
   strDefPath = strDefPath & "\" & pCP01 & pCP02 & IIf(pCP03 & pCP04 = "000", "", pCP03 & pCP04)
   If Not oFileSys.FolderExists(strDefPath) Then
      MkDir strDefPath
   End If
   Pub_GetEFilePath_All = strDefPath
   
ErrHandle:
   If Err.Number <> 0 Then
       MsgBox Err.Description, vbCritical
   End If
   Set oFileSys = Nothing
End Function

'Added by Lydia 2023/04/27 檢查存放路徑有存在相同檔名，直接加上系統日期和時間
Public Function Pub_GetEFileName(ByVal ToFilePath As String, ByVal sFileName As String) As String
'ToFilePath: (目的)存放路徑，
'SFileName: (來源)原始檔名
Dim oFileSys
Dim strNowName As String

On Error GoTo ErrHandle

   Set oFileSys = CreateObject("Scripting.FileSystemObject")
   strNowName = sFileName
   If Not oFileSys.FolderExists(ToFilePath) Then
   Else
       If oFileSys.FileExists(ToFilePath & "\" & sFileName & IIf(InStr(sFileName, ".") = 0, ".*", "")) = True Then
          If InStr(sFileName, ".") = 0 Then
             strNowName = sFileName & "." & strSrvDate(1) & Format(ServerTime, "000000")
          Else
             strNowName = Mid(sFileName, 1, InStrRev(sFileName, ".")) & strSrvDate(1) & Format(ServerTime, "000000") & Mid(sFileName, InStrRev(sFileName, "."))
          End If
       End If
   End If
   Pub_GetEFileName = strNowName
   
ErrHandle:
   If Err.Number <> 0 Then
       MsgBox Err.Description, vbCritical
   End If
   Set oFileSys = Nothing
End Function

'Added by Lydia 2023/04/27 取得FCT案定稿存至存至FCT_WORKFLOW的檔名
'Modified by Lydia 2023/09/04 +pFN04, pFN05
Public Function Pub_GetFCTeFileName(ByVal pFilePath As String, ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pCP10 As String, _
             Optional ByVal pCP43toCP10 As String, Optional ByRef pFN01 As String, Optional ByRef pFN02 As String, Optional pFN03 As String, Optional ByRef pFN04 As String, Optional pFN05 As String) As Boolean
'Memo: 註冊證1701與核准輸入(補換發證書103、更正302)的處理相同，若規則有變更，請一併修改。
'pCP01~pCP04,pCP10: 目前進度的案號和案件性質
'pcp43tocp10: 相關收文號的案件性質
'pFN01: 定稿名稱、 pFN02: 譯文、 pFN03: 官方來函; pFN04,pFN05=>日文組之核准-更正(延展核准函)另外產生定稿
Dim strB01 As String
Dim mStrLang As String
   
   pFN01 = "": pFN02 = "": pFN03 = ""
   pFN04 = "": pFN05 = "" 'Added by Lydia 2023/09/04
   Pub_GetFCTeFileName = False
   mStrLang = GetLetterLanguage(pCP01, pCP02, pCP03, pCP04)
   strB01 = pCP01 & pCP02 & IIf(pCP03 & pCP04 = "000", "", pCP03 & pCP04) & "." & pCP10 & IIf(pCP43toCP10 <> "", "." & pCP43toCP10, "")
   If mStrLang = "3" Then  '日文定稿檔案名稱
       'If pCP10 <> "1701" Then '核准 'Mark by Lydia 2024/11/14 因日本代理人特別要求，需將通知信函與譯文等分開，並且統一名稱如下
          'Modified by Lydia 2023/09/04
          'If "" & pCP43toCP10 = "102" Then
          If "" & pCP43toCP10 = "302102" Then '核准-更正(延展核准函)
             pFN01 = "信.doc"
             pFN02 = "延展核准譯文.doc"
             pFN03 = "更正核准譯文.doc"
             pFN04 = "延展核准函.pdf"
             pFN05 = "更正核准函.pdf"
          ElseIf Left("" & pCP43toCP10, 3) = "102" Then
          'end 2023/09/04
             pFN01 = "書簡（更新）.doc"
             pFN02 = "和" & PUB_GetUniText("共用", "譯") & ".doc"  '和譯
             pFN03 = "更新許可書.pdf"
'Modified by Lydia 2024/11/14 統一名稱如下：
'          ElseIf InStr("103,302", pCP43toCP10) > 0 And pCP43toCP10 <> "" Then
'             '比照「註冊證輸入1701」的規則，在輸入「核准-補換發證書103」、「核准-更正302」的同時，將通知函、譯文和證書一併產生檔案
'             GoTo JumpToJP
'          Else
'             '日文組目前未對延展核准102之外的案件性質通知定稿命名, 故除了延展外, 其他案件性質之命名請同英文組
'             GoTo JumpToEN
'          End If
'       ElseIf pCP10 = "1701" Then
'JumpToJP:
'          pFN01 = "書簡.doc"
          Else
             'Added by Lydia 2024/11/19 註冊證不用+letter
             If pCP10 = "1701" Then
                pFN01 = "書簡.doc"
             Else
             'end 2024/11/19
                pFN01 = "書簡.letter.doc"
             End If
'end 2024/11/14
             'Modified by Lydia 2023/07/28 debug: PDF檔1=>Word檔
             'pFN02 = "和" & PUB_GetUniText("共用", "譯") & ".pdf"  '和譯
             pFN02 = "和" & PUB_GetUniText("共用", "譯") & ".doc"
             pFN03 = "証書.pdf"
          End If 'Added by Lydia 2024/11/14
       'End If '----If pCP10 <> "1701" Then '核准 'Mark by Lydia 2024/11/14 因日本代理人特別要求，需將通知信函與譯文等分開，並且統一名稱
   Else  '英文組
'JumpToEN: 'Mark by Lydia 2024/11/14
       pFN01 = strB01 & ".LTR.doc"
       pFN02 = strB01 & ".TRANS.doc"
       If pCP10 = "1701" Or (pCP10 = "1001" And InStr("103,302", pCP43toCP10) > 0 And pCP43toCP10 <> "") Then
          pFN03 = strB01 & ".CERT.pdf"
       Else
          pFN03 = strB01 & ".APVL.pdf"
       End If
   End If
   
   If pFN01 <> "" Then
      pFN01 = Pub_GetEFileName(pFilePath, pFN01)
   End If
   If pFN02 <> "" Then
      pFN02 = Pub_GetEFileName(pFilePath, pFN02)
   End If
   If pFN03 <> "" Then
      pFN03 = Pub_GetEFileName(pFilePath, pFN03)
   End If
   'Added by Lydia 2023/09/04
   If pFN04 <> "" Then
      pFN04 = Pub_GetEFileName(pFilePath, pFN04)
   End If
   If pFN05 <> "" Then
      pFN05 = Pub_GetEFileName(pFilePath, pFN05)
   End If
   'end 2023/09/04
   Pub_GetFCTeFileName = True
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbOKOnly, "取得FCT案定稿存至存至FCT_WORKFLOW的檔名"
   End If
End Function

'Added by Lydia 2023/09/07 互惠代理人案件統計表(frm050408):符合條件的代理人互惠設定檔和相對的專利/商標設定檔
'Modified by Lydia 2025/06/06 共同查詢=>新增「互惠期間統計」  ; +, Optional ByVal pSpecNo As String
Public Sub Pub_GetFCfrm050408(ByVal pUserNo As String, ByVal pKind As String, ByVal pOpt As String, ByVal pYear As String, ByVal pPeriod As String, ByVal p_bolByAgent As Boolean, Optional ByVal pFCno As String, Optional ByVal pSpecNo As String)
'pKind: 案件類別1-專利,2=商標 ; pOpt: 0-互惠代理人1, 1-關聯企業
'pYear: 統計年度 ;             pPeriode: 統計區間
'p_bolByAgent: 統計對象是否為代理人
Dim intA As Integer, strA1 As String, strTmp As String
Dim rsAD As New ADODB.Recordset
Dim intB As Integer, strFrmName As String
   
   strFrmName = "frm050408"
   If pOpt = "0" Then
      cnnConnection.Execute "delete from rdatafactory where FORMNAME='" & strFrmName & "' and ID=" & CNULL(pUserNo)
      '排序依照報表排列,區分非關聯客戶
      'Modified by Lydia 2025/06/06 共同查詢=>新增「互惠期間統計」
      'strA1 = "SELECT FC01,FC02,FC03,FC04,FC05,FC06,FC07,FC08,FC16,FC17,'0' AS KIND FROM FAGENTCONFIG,FAGENT, NATION " & _
                  "WHERE FC06=" & CNULL(IIf(pKind = "1", "CFP", "CFT")) & " AND FC04 = " & IIf(Val(pYear) > 1911, Val(pYear) - 1911, pYear) & " AND FC05 = " & pPeriod & _
                  IIf(pFCno <> "", " AND FC01||FC02='" & ChangeCustomerL(pFCno) & "'", "") & " AND FC01=FA01(+) AND FC02=FA02(+) AND FA10=NA01(+) ORDER BY SUBSTR(NA01,1,3) ASC, FC07 DESC,NVL(FA05,NVL(FA06,FA04)) ASC, FC01 ASC, FC03 ASC "
      strA1 = "SELECT FC01,FC02,FC03,FC04,FC05,FC06,FC07,FC08,FC16,FC17,'0' AS KIND,SUBSTR(NA01,1,3) AS FNA01,NVL(FA05,NVL(FA06,FA04)) AS FNAME" & _
              " From FAGENTCONFIG, FAGENT, NATION WHERE FC06='" & IIf(pKind = "1", "CFP", "CFT") & "' AND FC04=" & IIf(Val(pYear) > 1911, Val(pYear) - 1911, pYear) & " AND FC05='" & pPeriod & "'" & _
              " AND FC01=FA01(+) AND FC02=FA02(+) AND FA10=NA01(+) " & IIf(pFCno <> "", " AND FC01||FC02='" & ChangeCustomerL(pFCno) & "'", "")
      If pSpecNo <> "" Then  '未設定互惠的代理人
         pSpecNo = ChangeCustomerL(pSpecNo)
         strA1 = strA1 & " UNION SELECT FA01,FA02,'' AS FC03," & IIf(Val(pYear) > 1911, Val(pYear) - 1911, pYear) & " AS FC04,'" & pPeriod & "' AS FC05,'" & IIf(pKind = "1", "CFP", "CFT") & "' AS FC06,0 AS FC07,'' AS FC08,'' AS FC16,'' AS FC17,'0' AS KIND,SUBSTR(NA01,1,3) AS FNA01,NVL(FA05,NVL(FA06,FA04)) AS FNAME" & _
                " FROM FAGENT,NATION WHERE FA01='" & Mid(pSpecNo, 1, 8) & "' AND FA02='" & Mid(pSpecNo, 9, 1) & "' AND FA10=NA01(+)" & _
                " AND FA01||FA02 NOT IN (SELECT FC01||FC02 AS MNO FROM FAGENTCONFIG WHERE FC01||FC02='" & pSpecNo & "' AND FC06='" & IIf(pKind = "1", "CFP", "CFT") & "' AND FC04=" & IIf(Val(pYear) > 1911, Val(pYear) - 1911, pYear) & " AND FC05='" & pPeriod & "')"
      End If
      strA1 = strA1 & " ORDER BY FNA01 ASC, FC07 DESC, FNAME ASC, FC01 ASC, FC03 ASC"
      'end 2025/06/06
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strA1)
      If intA = 1 Then
         Set rsAD = PUB_CreateRecordset(rsAD, , , , strFrmName) '先暫存:符合條件的代理人互惠設定檔
      End If
      'Added by Lydia 2025/06/06
   Else
      cnnConnection.Execute "delete from rdatafactory where FORMNAME='" & strFrmName & "' and ID=" & CNULL(pUserNo) & " and r011='1' "
   End If

   strA1 = "select formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008,r009,r010,r011 from rdatafactory where FORMNAME='" & strFrmName & "' and ID=" & CNULL(pUserNo) & " and seqno =1 order by rowseq "
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strA1)
   If intA = 1 Then
      rsAD.MoveFirst
      Do While Not rsAD.EOF
         If pOpt = "0" Then '新增互惠代理人對應的案件類別
            strTmp = "INSERT INTO rdatafactory (formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008,r009,r010,r011) " & _
                        "SELECT '" & strFrmName & "' as frmname, " & CNULL(pUserNo) & " as id, 2 as seqno, " & rsAD.Fields("rowseq") & " as rowseqno, " & _
                        "FC01,FC02," & IIf(p_bolByAgent = True, " '' as FC03", "FC03") & ",FC04,FC05,FC06,FC07,FC08,FC16,FC17,'0' as CKind " & _
                        "FROM FAGENTCONFIG WHERE FC06=" & CNULL(IIf(pKind = "1", "CFT", "CFP")) & " AND FC04 = " & IIf(Val(pYear) > 1911, Val(pYear) - 1911, pYear) & " AND FC05 = " & pPeriod & _
                        " AND FC01=" & CNULL(rsAD.Fields("R001")) & " AND FC02=" & CNULL(rsAD.Fields("R002"))
            cnnConnection.Execute strTmp, intB
            If intB = 0 Then
                strTmp = "INSERT INTO rdatafactory (formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008,r009,r010,r011) " & _
                            "values ('" & strFrmName & "','" & pUserNo & "', '2' , '" & rsAD.Fields("rowseq") & "', '" & rsAD.Fields("r001") & "', '" & rsAD.Fields("r002") & "', NULL, '" & rsAD.Fields("r004") & "', '" & rsAD.Fields("r005") & "'," & _
                            CNULL(IIf(pKind = "1", "CFT", "CFP")) & ", '0', '','','','0' ) "
                cnnConnection.Execute strTmp, intB
            End If
         ElseIf pOpt = "1" Then '取得關聯企業
            If "" & rsAD.Fields("R001") <> "" Then
              intA = PUB_GetR100114_1(True, "frm050408_1", "" & rsAD.Fields("R001"), pUserNo)
              '1.畫面輸入案件類別(專利/商標,2.另一種案件類別(商標/專利)
              'Modified by Lydia 2024/07/04  另外抓的關聯企業也要抓自己的建議給案量; ex.Y5588700 (日本), Y5588701 (大陸)
              'strTmp = "INSERT INTO rdatafactory (formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008,r009,r010,r011) " & _
                       "select '" & strFrmName & "' as formname,'" & pUserNo & "' as id , '3' as seqno, '" & rsAD.Fields("rowseq") & "' as rowseq, ano as r001, '0' as r002, null as r003, '" & rsAD.Fields("r004") & "' as r004, '" & rsAD.Fields("r005") & "' as r005, " & _
                       CNULL(IIf(pKind = "1", "CFP", "CFT")) & " as r006, '0' as r007, '' as r008,'" & rsAD.Fields("r009") & "' as r009,'" & rsAD.Fields("r010") & "' as r010,'1' as r011 " & _
                       "from (select substr(r11402,1,8) ano from r100114_1 where id='" & pUserNo & "' and formid='FRM050408_1' and substr(r11402,1,8)<>'" & Left("" & rsAD.Fields("R001"), 8) & "' group by substr(r11402,1,8)) "
              'Memo by Lydia 2025/06/06 未設定互惠的代理人=>已檢查不用修改
              strTmp = "INSERT INTO rdatafactory (formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008,r009,r010,r011) " & _
                       "select '" & strFrmName & "' as formname,'" & pUserNo & "' as id , '3' as seqno, '" & rsAD.Fields("rowseq") & "' as rowseq, ano as r001, '0' as r002, null as r003, '" & rsAD.Fields("r004") & "' as r004, '" & rsAD.Fields("r005") & "' as r005, " & _
                       CNULL(IIf(pKind = "1", "CFP", "CFT")) & " as r006, nvl(fc07,'0') as r007, '' as r008,'" & rsAD.Fields("r009") & "' as r009,'" & rsAD.Fields("r010") & "' as r010,'1' as r011 " & _
                       "from (select substr(r11402,1,8) ano from r100114_1 where id='" & pUserNo & "' and formid='FRM050408_1' and substr(r11402,1,8)<>'" & Left("" & rsAD.Fields("R001"), 8) & "' group by substr(r11402,1,8)) " & _
                       ", fagentconfig where fc01(+)=ano and fc02(+)='0' and fc06(+)=" & CNULL(IIf(pKind = "1", "CFP", "CFT")) & " and fc04(+)=" & IIf(Val(pYear) > 1911, Val(pYear) - 1911, pYear) & " and fc05(+)=" & pPeriod
              cnnConnection.Execute strTmp, intB
              strTmp = "INSERT INTO rdatafactory (formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008,r009,r010,r011) " & _
                       "select '" & strFrmName & "' as formname,'" & pUserNo & "' as id , '4' as seqno, '" & rsAD.Fields("rowseq") & "' as rowseq, ano as r001, '0' as r002, null as r003, '" & rsAD.Fields("r004") & "' as r004, '" & rsAD.Fields("r005") & "' as r005, " & _
                       CNULL(IIf(pKind = "1", "CFT", "CFP")) & " as r006, '0' as r007, '' as r008,'" & rsAD.Fields("r009") & "' as r009,'" & rsAD.Fields("r010") & "' as r010,'1' as r011 " & _
                       "from (select substr(r11402,1,8) ano from r100114_1 where id='" & pUserNo & "' and formid='FRM050408_1' and substr(r11402,1,8)<>'" & Left("" & rsAD.Fields("R001"), 8) & "' group by substr(r11402,1,8)) "
              cnnConnection.Execute strTmp, intB
            End If
         End If
         rsAD.MoveNext
      Loop
   End If
   Set rsAD = Nothing
  
End Sub

'Added by Lydia 2023/09/07 互惠代理人案件統計表(" & strfrmname & "):統計語法與frmAutoBatchDay共用
'Modified by Lydia 2023/10/06 是否為Excel語法=> bolExcel
Public Sub Pub_GetSqlfrm050408(ByVal pUserNo As String, ByVal pKind As String, ByVal pOpt As String, ByVal pYear As String, ByVal pPeriod As String, ByVal p_bolByAgent As Boolean, _
       ByRef pConFirst As String, ByRef pConSecond As String, Optional ByVal pStDate0 As String, Optional ByVal pTDate1 As String, Optional ByVal pTDate2 As String, Optional ByVal bolExcel As Boolean = True)
'pKind: 案件類別1-專利,2=商標 ; pOpt: 0-互惠代理人1, 1-關聯企業
'pYear: 統計年度 ;             pPeriod: 統計區間
'pConFirst案件類別的主SQL(專利/商標), pConSecond相對的SQL(商標/專利)
'p_bolByAgent: 統計對象是否為代理人
Dim stVTB1 As String, stVTB2 As String
Dim stDate1 As String, stDate2 As String, bExtra As Boolean
Dim strFC06_P As String, strFC06_T As String
Dim strTmp(0 To 3) As String, intQ As Integer
'分別存放專利、商標SQL語法
Dim tmpPA As String, tmpTM As String, tmpPA2 As String, tmpTM2 As String
Dim strFrmName As String
      
   strFrmName = "frm050408"
   If pTDate1 <> "" Or pTDate2 <> "" Then  '指定日期區間
      bExtra = True
      If pTDate1 <> "" Then
         stDate1 = DBDATE(pTDate1)
      Else
         stDate1 = 0
      End If
      If pTDate1 <> "" Then
         stDate2 = DBDATE(pTDate2)
      Else
         stDate2 = strSrvDate(1)
      End If
   End If
   
   If Val(pStDate0) = 0 Then  '預設統計日期(起)
      pStDate0 = strSrvDate(1)
   End If
   
   pConFirst = "":   pConSecond = ""
   
   If Val(pYear) < 1911 Then
      pYear = Val(pYear) + 1911
   End If
   
   strFC06_P = "CFP"
   strFC06_T = "CFT"

   'CF案件統計
   stVTB1 = "select FC01||FC03 A2,count(*) CF_TOT" & _
      ",sum(decode(substr(cp27,1,4)," & pYear & "-2,1)) CF_L2" & _
      ",sum(decode(substr(cp27,1,4)," & pYear & "-1,1)) CF_L1" & _
      ",sum(decode(substr(cp27,1,4)," & pYear & ",decode(sign(substr(cp27,5,2)-6),1,0,1))) CF_C1" & _
      ",sum(decode(substr(cp27,1,4)," & pYear & ",decode(sign(substr(cp27,5,2)-6),1,1,0))) CF_C2 "
   If bExtra = True Then
      stVTB1 = stVTB1 & ",sum(decode(sign(cp27-" & stDate1 & "+1),1,decode(sign(cp27-" & stDate2 & "-1),-1,1))) CF_X "
   End If
   
    '專利: 原始新案定義為廣義新案數
    'Modified by Lydia 2024/01/30 (Widen)因應目前互惠的策略已擴大到事務所層級，再麻煩進行程式的調整到所有國家地區>> 加抓P案 ; and cp01||cp04='CFP00' 改為 and instr('CFP00,P00',cp01||cp04)>0
    tmpPA = stVTB1 & " From (SELECT DISTINCT R001 as FC01, " & IIf(p_bolByAgent = False, "R003", "''") & "  as FC03 FROM rdatafactory " & _
       " Where Formname='" & strFrmName & "' And Id='" & pUserNo & "' and R006='" & strFC06_P & "' AND R004 = '" & (pYear - 1911) & "' And R005 = '" & pPeriod & "' " & _
       " and R011='" & pOpt & "') VTB1,caseprogress WHERE CP44=FC01||'0'" & IIf(p_bolByAgent = False, " AND NVL(CP116,'0')=NVL(FC03,'0')", "") & _
       " and instr('CFP00,P00',cp01||cp04)>0 and cp27+0>19221111 and cp57 is null and cp27+0<=" & pStDate0 & _
       " and (instr('" & NewCasePtyList & "', cp10) > 0 or cp10 like '3%') and cp09<'B'" & _
       " group by FC01,FC03"
    '商標
    'Modified by Lydia 2024/01/30 加抓T案; and cp01||cp04='CFT00' 改為 and instr('CFT00,T00',cp01||cp04)>0
    tmpTM = stVTB1 & " From (SELECT DISTINCT R001 as FC01, " & IIf(p_bolByAgent = False, "R003", "''") & "  as FC03 FROM rdatafactory " & _
       " Where Formname='" & strFrmName & "' And Id='" & pUserNo & "' and R006='" & strFC06_T & "' AND R004 = '" & (pYear - 1911) & "' And R005 = '" & pPeriod & "' " & _
       " and R011='" & pOpt & "') VTB1,caseprogress WHERE CP44=FC01||'0'" & IIf(p_bolByAgent = False, " AND NVL(CP116,'0')=NVL(FC03,'0')", "") & _
       " and instr('CFT00,T00',cp01||cp04)>0 and cp27+0>19221111 and cp57 is null and cp27+0<=" & pStDate0 & _
       " and instr('101',cp10)>0 and cp09<'B'" & _
       " group by FC01,FC03"

   'FC案件統計
   stVTB2 = "select FC01||FC03 B2,count(*) FC_TOT" & _
      ",sum(decode(substr(cp05,1,4)," & pYear & "-2,1)) FC_L2" & _
      ",sum(decode(substr(cp05,1,4)," & pYear & "-1,1)) FC_L1" & _
      ",sum(decode(substr(cp05,1,4)," & pYear & ",decode(sign(substr(cp05,5,2)-6),1,0,1))) FC_C1" & _
      ",sum(decode(substr(cp05,1,4)," & pYear & ",decode(sign(substr(cp05,5,2)-6),1,1,0))) FC_C2 "
   If bExtra = True Then
      stVTB2 = stVTB2 & _
         ",sum(decode(sign(cp05-" & stDate1 & "+1),1,decode(sign(cp05-" & stDate2 & "+1),-1,1))) FC_X "
   End If
   
    '專利
    'Modified by Lydia 2024/01/30 (Widen)因應目前互惠的策略已擴大到事務所層級，再麻煩進行程式的調整到所有國家地區>> 加抓P案 pa01||''='FCP' 改為 (pa01||''='FCP' or pa01||''='P')
    tmpPA2 = stVTB2 & " from (SELECT DISTINCT R001 as FC01, " & IIf(p_bolByAgent = False, "R003", "''") & "  as FC03 FROM rdatafactory " & _
         " Where Formname='" & strFrmName & "' And Id='" & pUserNo & "' and R006='" & strFC06_P & "' AND R004 = '" & (pYear - 1911) & "' And R005 = '" & pPeriod & "' " & _
         " and R011='" & pOpt & "') VTB2,patent,caseprogress" & _
         " where pa75=FC01||'0' AND (pa01||''='FCP' or pa01||''='P')" & IIf(p_bolByAgent = False, " AND NVL(pa144,'0')=NVL(FC03,'0')", "") & _
         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp09<'B'" & _
         " and (instr('" & NewCasePtyList & "', cp10) > 0 or cp10 like '3%') and cp57 is null and cp05+0<=" & pStDate0 & _
         " group by FC01,FC03"
    '商標
    'Modified by Lydia 2024/01/30 加抓T案; tm01||''='FCT' 改為 (tm01||''='FCT' or tm01||''='T')
    tmpTM2 = stVTB2 & " from(SELECT DISTINCT R001 as FC01, " & IIf(p_bolByAgent = False, "R003", "''") & "  as FC03 FROM rdatafactory " & _
         " Where Formname='" & strFrmName & "' And Id='" & pUserNo & "' and R006='" & strFC06_T & "' AND R004 = '" & (pYear - 1911) & "' And R005 = '" & pPeriod & "' " & _
         " and R011='" & pOpt & "') VTB2,trademark,caseprogress" & _
         " where tm44=FC01||'0' AND (tm01||''='FCT' or tm01||''='T')" & IIf(p_bolByAgent = False, " AND NVL(tm119,'0')=NVL(FC03,'0')", "") & _
         " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04 and cp09<'B'" & _
         " and instr('101',cp10)>0 and cp57 is null and cp05+0<=" & pStDate0 & _
         " group by FC01,FC03"
   '加上欄位名稱
   strTmp(0) = "SELECT NVL(FA05,NVL(FA06,FA04)) C1,FC01||DECODE(FC03,NULL,'','-'||FC03) C2" & _
        ",NVL(PCC03,NVL(PCC04,PCC05)) C3,NA03,NVL(FC_TOT,0) FC_TOT,NVL(CF_TOT,0) CF_TOT,NVL(FC_L2,0) FC_L2,NVL(CF_L2,0) CF_L2 " & _
        ",NVL(FC_L1,0) FC_L1,NVL(CF_L1,0) CF_L1,NVL(FC_C1,0) FC_C1,NVL(CF_C1,0) CF_C1,NVL(FC_C1,0)-NVL(CF_C1,0) DIFF01,NVL(FC_C2,0) FC_C2" & _
        ",NVL(CF_C2,0) CF_C2,NVL(FC_C2,0)-NVL(CF_C2,0) DIFF02,FC07,FC08 "
   If bExtra = True Then
      strTmp(0) = strTmp(0) & ",nvl(FC_X,0) FC_x,nvl(CF_X,0) CF_X,nvl(FC_X,0)-nvl(CF_X,0) DIFF03"
   End If
   
   '加以代理人統計,分別存放專利、商標SQL語法
   For intQ = 1 To 2
      'Added by Lydia 2023/12/26
      If strSrvDate(1) >= 新部門啟用日 Then
         If p_bolByAgent = True Then
            strTmp(intQ) = strTmp(0) & ",FC16,NVL(A0922,A0902) AS FC17DEPT,ST02 AS FC17NAME from (SELECT R001 as FC01,'' FC03,'' PCC03,'' PCC04,'' PCC05,SUM(R007) FC07,MIN(R008) FC08,R009 as FC16, R010 as FC17, ROWSEQ " & _
               " FROM Rdatafactory where Formname='" & strFrmName & "' And Id='" & pUserNo & "' and R006='" & IIf(intQ = 1, strFC06_P, strFC06_T) & "' AND R004 = '" & (pYear - 1911) & "' And R005 = '" & pPeriod & "' and R011='" & pOpt & "' " & _
               " GROUP BY R001, R009, R010,ROWSEQ) X,(" & IIf(intQ = 1, tmpPA, tmpTM) & ") A,(" & IIf(intQ = 1, tmpPA2, tmpTM2) & ") B,FAGENT,NATION,STAFF,ACC090,ACC090NEW where A2(+)=FC01||FC03" & _
               " and B2(+)=FC01||FC03 AND FA01(+)=FC01 AND FA02(+)='0' AND NA01(+)=SUBSTR(FA10,1,3) AND FC17=ST01(+) AND ST03=A0901(+) AND ST93=A0921(+)" & _
               " ORDER BY NA01 ASC,FC07 DESC,C1 ASC,FC01 ASC"
         Else
            strTmp(intQ) = strTmp(0) & ",FC16,NVL(A0922,A0902) AS FC17DEPT,ST02 AS FC17NAME from (SELECT R001 AS FC01, R002 AS FC02, R003 AS FC03,R007 AS FC07, R008 AS FC08,R009 as FC16, R010 as FC17, ROWSEQ " & _
                " from rdatafactory where Formname='" & strFrmName & "' And Id='" & pUserNo & "' and R006='" & IIf(intQ = 1, strFC06_P, strFC06_T) & "' AND R004 = '" & (pYear - 1911) & "' And R005 = '" & pPeriod & "' and R011='" & pOpt & "' " & ") X, " & _
                " (" & IIf(intQ = 1, tmpPA, tmpTM) & ") A,(" & IIf(intQ = 1, tmpPA2, tmpTM2) & ") B,FAGENT,NATION,POTCUSTCONT,STAFF,ACC090,ACC090NEW" & _
               " WHERE A2(+)=FC01||FC03 AND B2(+)=FC01||FC03 AND FA01(+)=FC01 AND FA02(+)='0' AND NA01(+)=SUBSTR(FA10,1,3) AND PCC01(+)=FC01 AND PCC02(+)=FC03 AND FC17=ST01(+) AND ST03=A0901(+) AND ST93=A0921(+)" & _
               " ORDER BY NA01 ASC,FC07 DESC,C1,FC01 ASC,C1 ASC,FC03 ASC"
         End If
      Else
      'end 2023/12/26
         If p_bolByAgent = True Then
            strTmp(intQ) = strTmp(0) & ",FC16,A0902 AS FC17DEPT,ST02 AS FC17NAME from (SELECT R001 as FC01,'' FC03,'' PCC03,'' PCC04,'' PCC05,SUM(R007) FC07,MIN(R008) FC08,R009 as FC16, R010 as FC17, ROWSEQ " & _
               " FROM Rdatafactory where Formname='" & strFrmName & "' And Id='" & pUserNo & "' and R006='" & IIf(intQ = 1, strFC06_P, strFC06_T) & "' AND R004 = '" & (pYear - 1911) & "' And R005 = '" & pPeriod & "' and R011='" & pOpt & "' " & _
               " GROUP BY R001, R009, R010,ROWSEQ) X,(" & IIf(intQ = 1, tmpPA, tmpTM) & ") A,(" & IIf(intQ = 1, tmpPA2, tmpTM2) & ") B,FAGENT,NATION,STAFF,ACC090 where A2(+)=FC01||FC03" & _
               " and B2(+)=FC01||FC03 AND FA01(+)=FC01 AND FA02(+)='0' AND NA01(+)=SUBSTR(FA10,1,3) AND FC17=ST01(+) AND ST03=A0901(+)" & _
               " ORDER BY NA01 ASC,FC07 DESC,C1 ASC,FC01 ASC"
         Else
            strTmp(intQ) = strTmp(0) & ",FC16,A0902 AS FC17DEPT,ST02 AS FC17NAME from (SELECT R001 AS FC01, R002 AS FC02, R003 AS FC03,R007 AS FC07, R008 AS FC08,R009 as FC16, R010 as FC17, ROWSEQ " & _
                " from rdatafactory where Formname='" & strFrmName & "' And Id='" & pUserNo & "' and R006='" & IIf(intQ = 1, strFC06_P, strFC06_T) & "' AND R004 = '" & (pYear - 1911) & "' And R005 = '" & pPeriod & "' and R011='" & pOpt & "' " & ") X, " & _
                " (" & IIf(intQ = 1, tmpPA, tmpTM) & ") A,(" & IIf(intQ = 1, tmpPA2, tmpTM2) & ") B,FAGENT,NATION,POTCUSTCONT,STAFF,ACC090" & _
               " WHERE A2(+)=FC01||FC03 AND B2(+)=FC01||FC03 AND FA01(+)=FC01 AND FA02(+)='0' AND NA01(+)=SUBSTR(FA10,1,3) AND PCC01(+)=FC01 AND PCC02(+)=FC03 AND FC17=ST01(+) AND ST03=A0901(+)" & _
               " ORDER BY NA01 ASC,FC07 DESC,C1,FC01 ASC,C1 ASC,FC03 ASC"
         End If
      End If
   Next intQ
   
   If pKind = "1" Then '案件類別:專利
       pConFirst = strTmp(1): pConSecond = strTmp(2)
   Else
       pConFirst = strTmp(2): pConSecond = strTmp(1)
   End If

   'Modified by Lydia 2023/10/06 +bolExcel
   If bExtra = True And bolExcel = True Then
       pConFirst = Replace(Replace(Replace(pConFirst, ",sum(decode(sign(cp27-" & stDate1 & "+1),1,decode(sign(cp27-" & stDate2 & "-1),-1,1))) CF_X", ""), ",sum(decode(sign(cp05-" & stDate1 & "+1),1,decode(sign(cp05-" & stDate2 & "+1),-1,1))) FC_X", ""), ",nvl(FC_X,0) FC_x,nvl(CF_X,0) CF_X,nvl(FC_X,0)-nvl(CF_X,0) DIFF03", "")
       pConSecond = Replace(Replace(Replace(pConSecond, ",sum(decode(sign(cp27-" & stDate1 & "+1),1,decode(sign(cp27-" & stDate2 & "-1),-1,1))) CF_X", ""), ",sum(decode(sign(cp05-" & stDate1 & "+1),1,decode(sign(cp05-" & stDate2 & "+1),-1,1))) FC_X", ""), ",nvl(FC_X,0) FC_x,nvl(CF_X,0) CF_X,nvl(FC_X,0)-nvl(CF_X,0) DIFF03", "")
   End If
End Sub

'Add by Morgan 2010/5/24
'Modify By Sindy 2016/3/21 + , Optional bolInOneSetTwo As Boolean = True : 是否含一案兩請
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function PUB_GetRefCaseMapSQL(PField() As String, Optional bolInOneSetTwo As Boolean = True) As String
   Dim stVTable As String
   Dim strSQLCon As String 'Add By Sindy 2016/3/21
   
   'Add By Sindy 2016/3/21
   If bolInOneSetTwo = True Then
      strSQLCon = ",'3'"
   End If
   '2016/3/21 END
   'Modified by Morgan 2015/9/16 +CM10='0'
   'Mofidied by Morgan 2015/12/1 +一案兩請 (,'3')
   '國內案
   stVTable = " SELECT CM05 C01,CM06 C02,CM07 C03,CM08 C04 FROM CASEMAP" & _
      " WHERE CM01='" & PField(1) & "' AND CM02='" & PField(2) & "' AND CM03='" & PField(3) & "' AND CM04='" & PField(4) & "' AND CM10 IN ('0'" & strSQLCon & ")"
   '國內案的其他國外案
   stVTable = stVTable & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
      " (SELECT CM05,CM06,CM07,CM08 FROM CASEMAP" & _
      " WHERE CM01='" & PField(1) & "' AND CM02='" & PField(2) & "' AND CM03='" & PField(3) & "' AND CM04='" & PField(4) & "' AND CM10 IN ('0'" & strSQLCon & ")) AND CM10 IN ('0'" & strSQLCon & ")"
   '國內案的其他國外案的國外案
   stVTable = stVTable & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
      " (SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
      " (SELECT CM05,CM06,CM07,CM08 FROM CASEMAP" & _
      " WHERE CM01='" & PField(1) & "' AND CM02='" & PField(2) & "' AND CM03='" & PField(3) & "' AND CM04='" & PField(4) & "' AND CM10='0') AND CM10 IN ('0'" & strSQLCon & ")) AND CM10 IN ('0'" & strSQLCon & ")"
   
   '國外案
   stVTable = stVTable & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP" & _
      " WHERE CM05='" & PField(1) & "' AND CM06='" & PField(2) & "' AND CM07='" & PField(3) & "' AND CM08='" & PField(4) & "' AND CM10 IN ('0'" & strSQLCon & ")"
   '國外案的其他國外案
   stVTable = stVTable & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
      " (SELECT CM01,CM02,CM03,CM04 FROM CASEMAP" & _
      " WHERE CM05='" & PField(1) & "' AND CM06='" & PField(2) & "' AND CM07='" & PField(3) & "' AND CM08='" & PField(4) & "' AND CM10 IN ('0'" & strSQLCon & ")) AND CM10 IN ('0'" & strSQLCon & ")"
      
   PUB_GetRefCaseMapSQL = stVTable
End Function

'Added by Lydia 2024/05/10 傳入客戶/代理人編號+聯絡人編號，回傳Key
Public Function Pub_GetPCCtoIBF(ByVal pKey01 As String, ByVal pKey02 As String, ByVal pPos As String) As String
Dim strMid As String
    
   'Modified by Lydia 2024/05/14 +潛在客戶R
   If pKey01 = "" Or pKey02 = "" Or (Left(pKey01, 1) <> "X" And Left(pKey01, 1) <> "Y" And Left(pKey01, 1) <> "R") Then
      Exit Function
   Else
      strMid = Left(Trim(pKey01) & String(8, "0"), 8) & Right("00" & Trim(pKey02), 2)
      Select Case pPos
         Case "1": Pub_GetPCCtoIBF = Mid(strMid, 1, 3)
         Case "2": Pub_GetPCCtoIBF = Mid(strMid, 4, 6)
         Case "3": Pub_GetPCCtoIBF = Mid(strMid, 10, 1)
         Case "4": Pub_GetPCCtoIBF = "00"
         Case "5": Pub_GetPCCtoIBF = "3"
      End Select
   End If
   
End Function

'Added by Lydia 2024/05/10 傳入客戶/代理人編號+聯絡人編號，檢查是否有相片
Public Function Pub_GetPCCtoIBF_2(ByVal pKey01 As String, ByVal pKey02 As String, Optional ByRef pOobj As CommandButton) As Boolean
Dim intQ As Integer, strQ1 As String
Dim rsQD As New ADODB.Recordset
   
   If Pub_GetPCCtoIBF(pKey01, pKey02, "1") <> "" Then
      strQ1 = "select * from imgbytefile where ibf01='" & Pub_GetPCCtoIBF(pKey01, pKey02, "1") & "' " & _
               "and ibf02='" & Pub_GetPCCtoIBF(pKey01, pKey02, "2") & "' and ibf03='" & Pub_GetPCCtoIBF(pKey01, pKey02, "3") & "' " & _
               "and ibf04='" & Pub_GetPCCtoIBF(pKey01, pKey02, "4") & "' and ibf05='" & Pub_GetPCCtoIBF(pKey01, pKey02, "5") & "' "
      intQ = 1
      Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
      If intQ = 1 Then
         Pub_GetPCCtoIBF_2 = True
         GoTo EXITSUB
      End If
   End If
   Pub_GetPCCtoIBF_2 = False
   
EXITSUB:
   Set rsQD = Nothing
   If UCase(pOobj.Name) <> "" Then
      If Pub_GetPCCtoIBF_2 = True Then
         pOobj.Caption = "已有相片"
         pOobj.BackColor = &H80FF80    '綠色
      Else
         pOobj.Caption = "上傳相片"
         pOobj.BackColor = &H8080FF     '紅色
      End If
   End If
End Function

'Added by Lydia 2024/09/30 取得出庭費明細表的語法(frm075013_2)
Public Function PUB_GetFrm075013toSQL(ByVal pKind As String, ByVal pArea As String, ByVal pCL06 As String, ByVal pDDate1 As String, ByVal pDDate2 As String, Optional ByVal pNoList As String = "") As String
'pKind : 1-Grid用(frm075013_2), 2-Excel用(frm075013_2),4-每年1月1號通知
Dim strCon1 As String, strCon2 As String
Dim strQ1 As String
   
   If Trim(pArea) <> "" Then
      If Left(pArea, 1) >= "6" And Left(pArea, 1) < "F" Then
         strCon1 = strCon1 & " and a.cl02='" & Trim(Left(pArea, 6)) & "' "
      Else
         strCon2 = strCon2 & " and c2.cp01='" & Trim(Left(pArea, 4)) & "' "
      End If
   End If
   
   'Added by Lydia 2025/04/07 只顯示不領取確認(CL09) 'Memo by Lydia 2025/04/17 從A改K
   If pCL06 = "K" Then
      strCon1 = strCon1 & " and NVL(A.CL09,0) > 0 "
   Else
   'end 2025/04/07
      '財務:Y：已發放  N：未發放  空白：不限制 'Modified by Lydia 2025/04/07 a.cl06->a.cl06||a.cl09
      If pCL06 = "Y" Then
         'Modified by Lydia 2025/08/20 debug: 財務不領取確認為單獨顯示
         'strCon1 = strCon1 & " and a.cl06||a.cl09 is not null "
         strCon1 = strCon1 & " and a.cl06 is not null "
      ElseIf pCL06 = "N" Then
         strCon1 = strCon1 & " and a.cl06||a.cl09 is null "
      'Added by Lydia 2025/04/07 排除不領取確認
      ElseIf pCL06 = "" Then
         strCon1 = strCon1 & " and a.cl09 is null "
      End If
   End If 'Added by Lydia 2025/04/07
   
   If pKind = "4" Then  '每日批次:每年第1個工作天寄發未領取清單
      If pDDate1 <> "" Then
         strCon2 = strCon2 & " and c1.cp158>='" & DBDATE(pDDate1) & "' "
      End If
      If pDDate2 <> "" Then
         strCon2 = strCon2 & " and c1.cp158<='" & DBDATE(pDDate2) & "' "
      End If
   Else
      '律師:確認日期(領取/不領取)
      If pDDate1 <> "" Then
         strCon1 = strCon1 & " and nvl(a.cl04,a.cl05)>='" & DBDATE(pDDate1) & "' "
      End If
      If pDDate2 <> "" Then
         strCon1 = strCon1 & " and nvl(a.cl04,a.cl05)<='" & DBDATE(pDDate2) & "' "
      End If
   End If
   If pNoList <> "" Then '出庭費發放通知:傳入收文號+員工編號
      strCon1 = strCon1 & " and instr ('" & pNoList & "',a.cl01||a.cl02)>0 "
   End If
   
   'Modified by Lydia 2024/11/21 修改收據的判斷
   'strQ1 = "SELECT CL01,CL06,Y01,Y02,Y03,LISTAGG(X01,',') WITHIN GROUP (ORDER BY CL01) AS X01,LISTAGG(X02,',') WITHIN GROUP (ORDER BY CL01) AS X02,SUM(X03) AS X03 " & _
           "FROM (SELECT A.CL01,A.CL06,A.CL02 AS Y01,S1.ST02 AS Y02, A.CL03 AS Y03,B.CL02 AS X01,S2.ST02 AS X02, B.CL03 AS X03 " & _
                 "FROM CASELAWER A,STAFF S1,CASELAWER B,STAFF S2 WHERE A.CL02=S1.ST01(+) " & strCon1 & _
                 "AND A.CL01=B.CL01(+) AND A.CL02<>B.CL02 AND B.CL02=S2.ST01(+) AND NVL(B.CL03,0)>0 " & _
                 "UNION SELECT A.CL01,A.CL06,A.CL02 AS Y01,S1.ST02 AS Y02,A.CL03 AS Y03 ,'' AS X01, '' AS X02,NULL AS X03 " & _
                 "FROM CASELAWER A, STAFF S1 WHERE A.CL02=S1.ST01(+) AND NVL(A.CL03,0)>0 " & strCon1 & _
                  ") GROUP BY CL01,CL06,Y01,Y02,Y03"
   'Modified by Lydia 2024/11/05 因為還有特殊案件性質，不能限制會計科目CPM12=> 拿掉AND INSTR('," & CaseLawerPtyList & ",',','||CPM12||',') > 0
   'PUB_GetFrm075013toSQL = "SELECT " & IIf(pKind = "1", " '' as v,", "") & " LC01||'-'||LC02||'-'||LC03||'-'||LC04 AS CASENO, DECODE(C2.CP01,NULL,NULL,'TT',NULL,C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04) AS PCASE " & _
               ", Y01||' '||Y02 AS Y01NAME, Y03,SQLDATET(NVL(B2.CL04,B2.CL05)) AS CHKDATE,DECODE(B2.CL04,NULL,DECODE(B2.CL05,NULL,NULL,'不領取'),'領取') AS CHKTYPE," & _
               "SQLDATET(B1.CL06) AS CL06T, SQLDATET(C1.CP27) AS CP27T, SQLDATET(A0L02) AS A0L02T, X01, X02, X03, Y01, B1.CL01 " & _
               "FROM (" & strQ1 & ") B1, CASEPROGRESS C1, LAWCASE, CASEPROPERTYMAP, LAWOFFICESOURCE, CASEPROGRESS C2,CASELAWER B2, " & _
               "(SELECT A0M02,MIN(A0L02) A0L02 FROM ACC0M0,ACC0L0 WHERE A0M01=A0L01(+) AND A0M02 IN (SELECT CP60 FROM CASELAWER A, CASEPROGRESS WHERE A.CL01=CP09(+) " & strCon1 & ") GROUP BY A0M02) VTB1 " & _
               "WHERE B1.CL01=C1.CP09(+) AND C1.CP159=0 AND C1.CP01=LC01(+) AND C1.CP02=LC02(+) AND C1.CP03=LC03(+) AND C1.CP04=LC04(+) " & _
               "AND C1.CP01=CPM01(+) AND C1.CP10=CPM02(+) AND C1.CP162=LOS15(+) AND LOS01=C2.CP09(+) " & _
               "AND B1.CL01=B2.CL01(+) AND B1.Y01=B2.CL02(+) AND C1.CP60=A0M02(+) " & strCon2
   strQ1 = "SELECT LC01||'-'||LC02||'-'||LC03||'-'||LC04 AS CASENO, DECODE(C2.CP01,NULL,NULL,'TT',NULL,C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04) AS PCASE " & _
               ", Y01||' '||Y02 AS Y01NAME, Y03,SQLDATET(NVL(B2.CL04,B2.CL05)) AS CHKDATE,DECODE(B2.CL04,NULL,DECODE(B2.CL05,NULL,NULL,'不領取'),'領取') AS CHKTYPE," & _
               "SQLDATET(B1.CL06) AS CL06T, SQLDATET(C1.CP27) AS CP27T, NVL(C1.CP60,C2.CP60) AS CCP60, X01, X02, X03, Y01, B1.CL01 " & _
               "FROM (SELECT CL01,CL06,Y01,Y02,Y03,LISTAGG(X01,',') WITHIN GROUP (ORDER BY CL01) AS X01,LISTAGG(X02,',') WITHIN GROUP (ORDER BY CL01) AS X02,SUM(X03) AS X03 " & _
           "FROM (SELECT A.CL01,A.CL06,A.CL02 AS Y01,S1.ST02 AS Y02, A.CL03 AS Y03,B.CL02 AS X01,S2.ST02 AS X02, B.CL03 AS X03 " & _
                 "FROM CASELAWER A,STAFF S1,CASELAWER B,STAFF S2 WHERE A.CL02=S1.ST01(+) " & strCon1 & _
                 "AND A.CL01=B.CL01(+) AND A.CL02<>B.CL02 AND B.CL02=S2.ST01(+) AND NVL(B.CL03,0)>0 " & _
                 "UNION SELECT A.CL01,A.CL06,A.CL02 AS Y01,S1.ST02 AS Y02,A.CL03 AS Y03 ,'' AS X01, '' AS X02,NULL AS X03 " & _
                 "FROM CASELAWER A, STAFF S1 WHERE A.CL02=S1.ST01(+) AND NVL(A.CL03,0)>0 " & strCon1 & _
                  ") GROUP BY CL01,CL06,Y01,Y02,Y03 ) B1, CASEPROGRESS C1, LAWCASE, CASEPROPERTYMAP, LAWOFFICESOURCE, CASEPROGRESS C2,CASELAWER B2 " & _
               "WHERE B1.CL01=C1.CP09(+) AND C1.CP159=0 AND C1.CP01=LC01(+) AND C1.CP02=LC02(+) AND C1.CP03=LC03(+) AND C1.CP04=LC04(+) " & _
               "AND C1.CP01=CPM01(+) AND C1.CP10=CPM02(+) AND C1.CP162=LOS15(+) AND LOS01=C2.CP09(+) " & _
               "AND B1.CL01=B2.CL01(+) AND B1.Y01=B2.CL02(+) " & strCon2
   PUB_GetFrm075013toSQL = "SELECT " & IIf(pKind = "1", " '' as v,", "") & " CASENO,PCASE,Y01NAME,Y03,CHKDATE,CHKTYPE,CL06T,CP27T,SQLDATET(MIN(NVL(A0L02,A0Y02))) AS A0L02T,X01,X02,X03,Y01,CL01 " & _
           "FROM (" & strQ1 & "), ACC0M0,ACC0L0,ACC0Z0,ACC0Y0 " & _
           "WHERE CCP60=A0M02(+) AND A0M01=A0L01(+) AND CCP60=A0Z02(+) AND A0Z01=A0Y01(+) " & _
           "GROUP BY CASENO,PCASE,Y01NAME,Y03,CHKDATE,CHKTYPE,CL06T,CP27T,X01,X02,X03,Y01,CL01 "
   'end 2024/11/21
   'P、T、FCP、FCT(傳票作業才分給律師，會計科目2201): 因為語法較慢，有需要才加入
   'Modified by Lydia 2025/04/10 排除「只顯示不領取確認(CL09)」 'Memo by Lydia 2025/04/17 從A改K
   If (pKind = "1" Or pKind = "2") And pCL06 <> "K" And pNoList = "" And (Trim(pArea) = "" Or (InStr(",P,T,FCP,FCT,", "," & Trim(Left(pArea, 4)) & ",") > 0)) Then
      '國內收款
      strQ1 = "SELECT " & IIf(pKind = "1", " '' as v,", "") & " C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04 AS CASENO,GETCASECODE(A1P17,'9') AS PCASENO,SUBSTR(A1P17, 1, LENGTH(A1P17) - 9) AS Y01NAME,A1P08 AS Y03, " & _
              "NULL AS CHKDAKTE,NULL AS CHKTYPE,NULL AS CL06T,NULL AS CP27T,SQLDATET(A1P18+19110000) AS A0L02T, " & _
              "NULL AS X01,NULL AS X02,NULL AS X03,SUBSTR(A1P17, 1, LENGTH(A1P17) - 9) AS Y01,C2.CP09 AS CL01 " & _
              "From ACC1P0, LAWOFFICESOURCE, ACC010, ACC0M0, ACC0J0,CASEPROGRESS C2 " & _
              "WHERE A1P01='1' AND A1P02='A' AND A1P05=A0101(+) " & IIf(pDDate1 <> "", " AND A1P18>=" & TransDate(pDDate1, 1), "") & IIf(pDDate2 <> "", " AND A1P18<=" & TransDate(pDDate2, 1), "") & _
              " AND SUBSTR(A1P05,1,4)='2201' AND A1P08=5000 AND A1P04=A0M01(+) AND A0M02=A0J13(+) AND INSTR(A0J02,'L')>0 AND A0J01=LOS06(+) AND LOS02>'B' AND LOS06=C2.CP09(+) " & _
              IIf(Trim(Left(pArea, 4)) <> "", " AND SUBSTR(A1P17,1," & Len(Trim(Left(pArea, 4))) & ")='" & Trim(Left(pArea, 4)) & "' ", "") & _
              "GROUP BY C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,GETCASECODE(A1P17,'9'),SUBSTR(A1P17, 1, LENGTH(A1P17) - 9),A1P08,SQLDATET(A1P18+19110000),C2.CP09 "
      '國外收款或抵帳
      strQ1 = strQ1 & " UNION SELECT " & IIf(pKind = "1", " '' as v,", "") & " C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04 AS CASENO,GETCASECODE(A1P17,'9') AS PCASENO,SUBSTR(A1P17, 1, LENGTH(A1P17) - 9) AS Y01NAME,A1P08 AS Y03, " & _
              "NULL AS CHKDAKTE,NULL AS CHKTYPE,NULL AS CL06T,NULL AS CP27T,SQLDATET(A1P18+19110000) AS A0L02T, " & _
              "NULL AS X01,NULL AS X02,NULL AS X03,SUBSTR(A1P17, 1, LENGTH(A1P17) - 9) AS Y01,C2.CP09 AS CL01 " & _
              "FROM ACC1P0,ACC010,ACC1K0,LAWOFFICESOURCE,CASEPROGRESS C1,CASEPROGRESS C2 " & _
              "WHERE A1P01='1' AND A1P02 IN ('F','K') AND A1P05=A0101(+) " & IIf(pDDate1 <> "", " AND A1P18>=" & TransDate(pDDate1, 1), "") & IIf(pDDate2 <> "", " AND A1P18<=" & TransDate(pDDate2, 1), "") & _
              " AND SUBSTR(A1P05,1,4)='2201' AND A1P08=5000 AND A1P23=A1K01(+) AND A1K01=C1.CP60(+) AND C1.CP09=LOS01(+) AND LOS02>'B' AND LOS06=C2.CP09(+) " & _
              IIf(Trim(Left(pArea, 4)) <> "", " AND SUBSTR(A1P17,1," & Len(Trim(Left(pArea, 4))) & ")='" & Trim(Left(pArea, 4)) & "' ", "") & _
              "GROUP BY C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04,GETCASECODE(A1P17,'9'),SUBSTR(A1P17, 1, LENGTH(A1P17) - 9),A1P08,SQLDATET(A1P18+19110000),C2.CP09 "
      PUB_GetFrm075013toSQL = PUB_GetFrm075013toSQL & " UNION " & strQ1
   End If
   
   'Added by Lydia 2025/04/07 只顯示不領取確認(CL09) 'Memo by Lydia 2025/04/17 從A改K
   If pCL06 = "K" Then
      PUB_GetFrm075013toSQL = Replace(UCase(PUB_GetFrm075013toSQL), "CL06", "CL09")
      'Added by Lydia 2025/08/20 debug: 條件不可變更
      PUB_GetFrm075013toSQL = Replace(UCase(PUB_GetFrm075013toSQL), "NVL(A.CL06,0)", "NVL(A.CL09,0)")
   End If
   'end 2025/04/07
   If pKind = "1" Then
      PUB_GetFrm075013toSQL = PUB_GetFrm075013toSQL & " order by cp27t, Y01"
   Else
      PUB_GetFrm075013toSQL = PUB_GetFrm075013toSQL & " order by Y01, cp27t"
   End If


End Function

'Added by Lydia 2024/09/30 取得出庭費明細表的語法(frm075013_2)
Public Function PUB_GetFrm075013toXls(ByVal pFrmName As String, ByVal pArea As String, ByVal pCL06 As String, ByVal pDDate1 As String, ByVal pDDate2 As String, ByRef pFilePath As String, Optional ByVal pNoList As String = "", Optional ByVal pBolMsg As Boolean = False) As Boolean
Dim intP As Integer, strP1 As String, intQ As Integer
Dim rsRD As New ADODB.Recordset
Dim xlsReport
Dim wksrpt
Dim strGrp As String, nRow As Integer, intPage As Integer
Dim tmpTitle As Variant, tmpArr As Variant
Dim tmpTitleW As Variant, tmpArr2 As Variant
Dim strTitleName As String, strTitleW As String
Dim strTitle1 As String
Dim pKind As String, tKind As String

   PUB_GetFrm075013toXls = False
   
   If UCase(pFrmName) = UCase("Frmacc42d0") Then
      strTitle1 = "出庭費發放通知"
   ElseIf UCase(pFrmName) = UCase("frmAutoBatchDay") Then
      strTitle1 = "律師出庭費未領取清單"
   Else
      strTitle1 = "出庭費清單"
   End If
   '---來源已刪除檔案,所以不做檢查
   
'Excel來源:
'1. 出庭費發放通知(Frmacc42d0)：給個人+領取
'2. 出庭費查詢(frm075013_2)：依輸入條件，選擇全部律師則以頁籤方式呈現
'3. 年底產出未領取清單(frmAutoBatchDay)：全部律師則以頁籤方式呈現
   If UCase(pFrmName) = UCase("frmAutoBatchDay") Then
      pKind = "4"
   Else
      pKind = "2"
   End If
   strP1 = PUB_GetFrm075013toSQL(pKind, pArea, pCL06, pDDate1, pDDate2, pNoList)
   intP = 1
   Set rsRD = ClsLawReadRstMsg(intP, strP1)
   If intP = 1 Then
   
On Error GoTo ErrHandle
      If UCase(pFrmName) = UCase("Frmacc42d0") Then
         strTitleName = "律所案號,智慧所案號,承辦律師,出庭費,發放日期,發文日,收款日"
         strTitleW = "15,15,15,11,11,11,11"
         tKind = "1" 'Added by Lydia 2024/11/06 不含其他出庭律師
      ElseIf UCase(pFrmName) = UCase("frm075013_2") Then
         'Modified by Lydia 2024/11/06 +其他出庭律師,出庭費總額
         'Modified by Lydia 2025/04/07 只顯示不領取確認 'Memo by Lydia 2025/04/17 從A改K
         'strTitleName = "律所案號,智慧所案號,承辦律師,出庭費,確認日期,確認結果,發放日期,發文日,收款日,其他出庭律師,出庭費總額"
         'strTitleW = "15,15,15,11,11,11,11,11,11,15,13"
         strTitleName = "律所案號,智慧所案號,承辦律師,出庭費,確認日期,確認結果," & IIf(pCL06 = "K", "財務不領取確認日期", "發放日期") & ",發文日,收款日,其他出庭律師,出庭費總額"
         strTitleW = "15,15,15,11,11,11," & IIf(pCL06 = "K", 15, 11) & ",11,11,15,13"
         'end 2025/04/07
         tKind = "2" 'Added by Lydia 2024/11/06 含其他出庭律師
      ElseIf UCase(pFrmName) = UCase("frmAutoBatchDay") Then
         'Modified by Lydia 2024/11/06 +其他出庭律師,出庭費總額
         strTitleName = "律所案號,智慧所案號,承辦律師,出庭費,確認日期,確認結果,發文日,收款日,其他出庭律師,出庭費總額"
         strTitleW = "15,15,15,11,11,11,11,11,15,13"
         tKind = "2" 'Added by Lydia 2024/11/06 含其他出庭律師
      End If
      tmpTitle = Split(strTitleName, ",")
      ReDim tmpArr(0 To UBound(tmpTitle))
      tmpTitleW = Split(strTitleW, ",")
      ReDim tmpArr2(0 To UBound(tmpTitleW))
      
      rsRD.MoveFirst
      Do While Not rsRD.EOF
         If strGrp <> "" & rsRD.Fields("Y01") Then
            If strGrp = "" Then
               Set xlsReport = CreateObject("Excel.Application")
               xlsReport.SheetsInNewWorkbook = 3
               xlsReport.Workbooks.add
               xlsReport.Application.Visible = False
            Else
               wksrpt.Range("A5").Select
               xlsReport.ActiveWindow.FreezePanes = True '凍結窗格
               If intPage + 1 > 3 Then
                  xlsReport.Worksheets.add After:=wksrpt
               End If
            End If
            intPage = intPage + 1
            Set wksrpt = xlsReport.Worksheets(intPage)
            xlsReport.Sheets(intPage).Select
            nRow = 1
            wksrpt.Range("C" & nRow).Value = strTitle1
            wksrpt.Range("C" & nRow).Font.Size = 16
            wksrpt.Range("C" & nRow).Font.Bold = True
            wksrpt.Range(nRow & ":" & nRow).RowHeight = 25
            'Added by Lydia 2024/11/06
            If tKind = "2" Then
                xlsReport.Range("C" & nRow & ":" & "G" & nRow).Select
            Else
            'end 2024/11/06
                xlsReport.Range("C" & nRow & ":" & "E" & nRow).Select
            End If
            xlsReport.Selection.Cells.Merge
            xlsReport.Selection.HorizontalAlignment = xlCenter
            nRow = nRow + 1
            If pNoList = "" Then
               wksrpt.Range("A" & nRow).Value = IIf(UCase(pFrmName) = UCase("frmAutoBatchDay"), "發文期間：", "確認期間：") & IIf(pDDate1 & pDDate2 = "", "　～　", ChangeTStringToTDateString(TransDate(pDDate1, 1)) & "～" & ChangeTStringToTDateString(TransDate(pDDate2, 1)))
               xlsReport.Range("A" & nRow & ":" & "B" & nRow).Select
               xlsReport.Selection.Cells.Merge
               xlsReport.Selection.HorizontalAlignment = xlLeft
               'Added by Lydia 2024/11/21
               If UCase(pFrmName) = UCase("frm075013_2") Then
                  'Added by Lydia 2025/04/07 只顯示不領取確認 'Memo by Lydia 2025/04/17 從A改K
                  If pCL06 = "K" Then
                     wksrpt.Range("D" & nRow).Value = "只顯示不領取確認"
                  Else
                  'end 2025/04/07
                     wksrpt.Range("D" & nRow).Value = "是否已發放：" & IIf(pCL06 = "Y", "已發放", IIf(pCL06 = "N", "未發放", "不限制"))
                  End If
                  xlsReport.Range("D" & nRow & ":" & "E" & nRow).Select
                  xlsReport.Selection.Cells.Merge
                  xlsReport.Selection.HorizontalAlignment = xlLeft
               End If
               'end 2024/11/21
            End If
            'Added by Lydia 2024/11/06
            If tKind = "2" Then '含其他出庭律師,出庭費總額
                wksrpt.Range("I" & nRow).Value = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
                xlsReport.Range("I" & nRow & ":" & "J" & nRow).Select
            Else
            'end 2024/11/06
                wksrpt.Range("G" & nRow).Value = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
                xlsReport.Range("G" & nRow & ":" & "H" & nRow).Select
            End If  'Added by Lydia 2024/11/06
            xlsReport.Selection.Cells.Merge
            xlsReport.Selection.HorizontalAlignment = xlLeft
            nRow = nRow + 2
            wksrpt.Range("A" & nRow & ":" & Chr(65 + UBound(tmpTitle)) & nRow).Value = tmpTitle
            wksrpt.Range(nRow & ":" & nRow).Font.Bold = True
            For intP = 0 To UBound(tmpTitleW)
               wksrpt.Range(Chr(65 + intP) & ":" & Chr(65 + intP)).ColumnWidth = Val("" & tmpTitleW(intP))
            Next intP
            nRow = nRow + 1
         End If
         '律所案號
         tmpArr(0) = "" & rsRD.Fields("caseno")
         '智慧所案號
         tmpArr(1) = "" & rsRD.Fields("pcase")
         '承辦律師
         tmpArr(2) = "" & rsRD.Fields("Y01NAME")
         '出庭費
         tmpArr(3) = Format("" & rsRD.Fields("Y03"), "##,##0")
         intQ = 4
         If pNoList = "" Then
            '確認日期
            tmpArr(intQ) = "" & rsRD.Fields("CHKDATE")
            '確認結果
            tmpArr(intQ + 1) = "" & rsRD.Fields("CHKTYPE")
            intQ = intQ + 2
         End If
         If UCase(pFrmName) <> UCase("frmAutoBatchDay") Then
            'Added by Lydia 2025/04/07 只顯示不領取確認 'Memo by Lydia 2025/04/17 從A改K
            If pCL06 = "K" Then
               tmpArr(intQ) = "" & rsRD.Fields("CL09T")
            Else
            'end 2025/04/07
            '發放日期
               tmpArr(intQ) = "" & rsRD.Fields("CL06T")
            End If
            intQ = intQ + 1
         End If
         '發文日
         tmpArr(intQ) = "" & rsRD.Fields("CP27T")
         intQ = intQ + 1
         '收款日
         tmpArr(intQ) = "" & rsRD.Fields("A0L02T")
         intQ = intQ + 1
         'Added by Lydia 2024/11/06
         If tKind = "2" Then
            tmpArr(intQ) = "" & rsRD.Fields("X02")  '其他出庭律師
            intQ = intQ + 1
            tmpArr(intQ) = "" & rsRD.Fields("X03")  '出庭費總額
            intQ = intQ + 1
         End If
         'end 2024/11/06
         wksrpt.Range("A" & nRow & ":" & Chr(65 + UBound(tmpTitle)) & nRow).Value = tmpArr
         If Len(Trim("" & rsRD.Fields("y01name"))) >= 6 Then
            wksrpt.Name = Trim(Mid("" & rsRD.Fields("y01name"), 7))
         Else
            wksrpt.Name = Trim("" & rsRD.Fields("y01name"))
         End If
         strGrp = "" & rsRD.Fields("Y01")
         nRow = nRow + 1
         rsRD.MoveNext
      Loop
      
      If intPage > 0 Then
         'For intP = 0 To UBound(tmpTitle) '調整為能使文字全部顯示之欄寬(目前工作表)
         '   wksrpt.Columns(Chr(65 + intP) & ":" & Chr(65 + intP)).EntireColumn.AutoFit
         'Next
         wksrpt.Range("A5").Select
         xlsReport.ActiveWindow.FreezePanes = True '凍結窗格
                  
         Set wksrpt = xlsReport.Worksheets(1)
         xlsReport.Sheets(1).Select
         '判斷版本
         If Val(xlsReport.Version) < 12 Then
            xlsReport.Workbooks(1).SaveAs FileName:=pFilePath, FileFormat:=-4143
         Else
            xlsReport.Workbooks(1).SaveAs FileName:=pFilePath, FileFormat:=56
         End If
      
         xlsReport.Workbooks.Close
         xlsReport.Quit
      End If
      Set wksrpt = Nothing
      Set xlsReport = Nothing
   End If
   
   Set rsRD = Nothing
   PUB_GetFrm075013toXls = True
   
   Exit Function
      
ErrHandle:
   If Err.Number <> 0 Then
      If pBolMsg = True Then MsgBox "Excel檔案產生失敗：" & vbCrLf & Err.Description
   End If
End Function

'Added by Lydia 2025/03/13 取得國外譯者的順序和員工編號/Y編號;
'(114/3/13)新增國外翻譯社要記得到外專工程師「外專新案命名作業frm090902_2,frm090903_1」需要人工增加Chk27的選項。,寫入TableSchema.案件命名記錄檔 TransCaseTitle.TCT27
'(114/10/14)新增國外翻譯社一併調整翻譯人員顯示「外專新案命名作業(待分案/待確認,待命名)frm090902,frm090903」
               'Compile: Patpro1, Patpro, Promoter, Account
Public Function Pub_SetF51Order(ByVal pType As String, ByVal pVAL01 As String) As String
Dim strMid As String

   If pVAL01 = "" And pType = "" Then Exit Function

   If pType = "Y" Then '所有Y編號,用在對外聯絡(FCP)和付款(account)
      '外翻_舜禹Y52268000,外翻_捷恩凱Y53541000,外翻_迅達Y54868000,外翻_百靈Y56151000(114/3/13新增),外翻_湃傳思Y56216000(114/9/16新增，原來捷恩凱翻譯公司，以湃傳思信息技術有限公司進行運作)
      strMid = "Y53541000,Y52268000,Y54868000,Y56151000,Y56216000"
   ElseIf pType = "F" Then  '所有F編號：外翻_舜禹F5588,外翻_捷恩凱F5653,外翻_迅達F5698,外翻_百靈F5726(114/3/13新增),外翻_湃傳思F5730(114/9/16新增)
      If pVAL01 = "2" Then
         strMid = "F5588舜禹,F5653捷恩凱,F5698迅達,F5726百靈,F5730湃傳思"
      ElseIf pVAL01 = "3" Then
         'Modified by Lydia 2025/10/23 目前僅用在外專翻譯分案不需要顯示捷恩凱，已與Sharon溝通後面的編號不變動(因為命名記錄TCT27=2尚有2筆)
         'strMid = "1: 舜禹 2: 捷恩凱 3: 迅達 4:百靈 5:湃傳思"
         strMid = "1: 舜禹 3: 迅達 4:百靈 5:湃傳思"
      Else
         strMid = "F5588,F5653,F5698,F5726,F5730"
      End If
   ElseIf pType = "T" Then '翻譯社Title
      strMid = pVAL01
      Select Case pVAL01
         Case "F5588"
            strMid = "江蘇舜禹翻譯"
         Case "F5653"
            strMid = "南京捷恩凱信息技術"
         Case "F5698"
            strMid = "迅達翻譯"
         Case "F5726"
            strMid = "百靈翻譯"
         'Added by Lydia 2025/09/16
         Case "F5730"
            strMid = "湃傳思信息技術"
         'end 2025/09/16
      End Select
   Else
      strMid = pVAL01
      Select Case pVAL01
         Case "1" '外翻_舜禹F5588
             strMid = "F5588"
         Case "F5588"
            strMid = "1"
         Case "2"  '外翻_捷恩凱F5653'Memo by Lydia 2025/03/13 Sharon:已不再給案>>保留順序
            strMid = "F5653"
         Case "F5653"
            strMid = "2"
         Case "3" '外翻_迅達F5698
            strMid = "F5698"
         Case "F5698"
            strMid = "3"
         Case "4"  '外翻_百靈F5726
            strMid = "F5726"
         Case "F5726"
            strMid = "4"
         'Added by Lydia 2025/09/16
         Case "5"  '外翻_湃傳思F5730
            strMid = "F5730"
         Case "F5730"
            strMid = "5"
         'end 2025/09/16
      End Select
   End If
   Pub_SetF51Order = strMid
   
End Function

'Added by Lydia 2025/06/27 原本在frm050408_1.GetStatistic案件統計改成共用模組，統計數先存rdatafactory
'＊＊若表單frm050408_1的欄位有變動，呼叫Pub_Frm050408_GetStatistic也要變動＊＊
Public Sub Pub_Frm050408_GetStatistic(ByVal pKind As String, ByVal pYear As String, ByVal pPeriod As String, ByVal p_bolByAgent As Boolean, ByVal pStrFC As String, ByVal pStrCF As String, ByVal pCol As Integer, _
       Optional ByVal pFCno As String, Optional ByRef pLblCond As String, Optional ByVal pStDate1 As String, Optional ByVal pStDate2 As String)
Dim stCon As String, iSys As Integer, iPos As Integer, stCP44 As String, stCP116 As String
Dim stVTB0 As String, stVTB1 As String, stVTB2 As String, stVTB3 As String
Dim stDate As String, iYear As Integer
Dim stAgentNo As String
Dim intA As Integer, strA1 As String
Dim intR As Integer, strTmp As String
Dim rsAD As New ADODB.Recordset
Dim tmpArr As Variant
   
   stDate = strSrvDate(1)  '案件盈虧只能從系統日起算
   iYear = Left(stDate, 4)
   pLblCond = ""
   If pFCno = "ALL" Then  '全部代理人
      strA1 = "SELECT FC01,FC02,FC03,FC04,FC05,FC06,FC07,FC08,FC16,FC17,'0' AS KIND,SUBSTR(NA01,1,3) AS FNA01,NVL(FA05,NVL(FA06,FA04)) AS FNAME" & _
              " From FAGENTCONFIG, FAGENT, NATION WHERE FC06='" & IIf(pKind = "1", "CFP", "CFT") & "' AND FC04=" & IIf(Val(pYear) > 1911, Val(pYear) - 1911, pYear) & " AND FC05='" & pPeriod & "'" & _
              " AND FC01=FA01(+) AND FC02=FA02(+) AND FA10=NA01(+) "
      strA1 = strA1 & " ORDER BY FNA01 ASC, FC07 DESC, FNAME ASC, FC01 ASC, FC03 ASC"
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strA1)
      If intA = 1 Then
         rsAD.MoveFirst
         Do While Not rsAD.EOF
            strTmp = strTmp & "," & rsAD.Fields("fc01") & rsAD.Fields("fc02")
            rsAD.MoveNext
         Loop
      End If
   Else
      cnnConnection.Execute "delete from rdatafactory where FORMNAME like 'frm050408_2%' and ID=" & CNULL(strUserNum)
      strTmp = "," & pFCno
   End If
   tmpArr = Split(Mid(strTmp, 2), ",")
   
   For intR = 0 To UBound(tmpArr)
      stAgentNo = Trim(tmpArr(intR))
      If stAgentNo <> "" Then
         iPos = InStr(stAgentNo, "-")
         If iPos = 0 Then
            stCP44 = Left(stAgentNo & "000", 9)
            stCP116 = ""
         Else
            stCP44 = Left(Left(stAgentNo, iPos - 1) & "000", 9)
            stCP116 = Right("00" & Mid(stAgentNo, iPos + 1), 2)
         End If
'----------------------------------
         stCon = ""
         Select Case pCol
            Case 4, 6, 8, 10, 13, 18 'FC
               iSys = 1
               If pStrFC = "FCP" Then
                  stCon = stCon & " and (instr('" & NewCasePtyList & "', cp10) > 0 or cp10 like '3%') and cp09<'B'"
                  'Modified by Lydia 2024/01/30 (Widen)因應目前互惠的策略已擴大到事務所層級，再麻煩進行程式的調整到所有國家地區>> 加抓P案
                  stCon = stCon & " and (pa01||''='FCP' or pa01||''='P') and pa75='" & stCP44 & "'"
                  '加以代理人統計
                  If p_bolByAgent = False Then
                     If stCP116 = "" Then
                        stCon = stCon & " and pa144 is null"
                     Else
                        stCon = stCon & " and pa144='" & stCP116 & "'"
                     End If
                  End If
               ElseIf pStrFC = "FCT" Then
                  stCon = stCon & " and instr('101',cp10)>0 and cp09<'B'"
                  'Modified by Lydia 2024/01/30 (Widen)因應目前互惠的策略已擴大到事務所層級，再麻煩進行程式的調整到所有國家地區>> 加抓T案
                  stCon = stCon & " and (tm01||''='FCT' or tm01||''='T') and tm44='" & stCP44 & "'"
                  '加以代理人統計
                  If p_bolByAgent = False Then
                     If stCP116 = "" Then
                        stCon = stCon & " and tm119 is null"
                     Else
                        stCon = stCon & " and tm119='" & stCP116 & "'"
                     End If
                  End If
               End If
               '2013/5/24 End
               stCon = stCon & " and cp05<=" & stDate
               If pCol = 6 Then '前年
                  stCon = stCon & " and cp05 between " & (iYear - 2) & "0101 and " & (iYear - 2) & "1231"
                  pLblCond = pStrFC & " " & (iYear - 1911 - 2) & " 年"
               ElseIf pCol = 8 Then '去年
                  stCon = stCon & " and cp05 between " & (iYear - 1) & "0101 and " & (iYear - 1) & "1231"
                  pLblCond = pStrFC & " " & (iYear - 1911 - 1) & " 年"
               ElseIf pCol = 10 Then '當年1-6月
                  stCon = stCon & " and cp05 between " & iYear & "0101 and " & iYear & "0630"
                  pLblCond = pStrFC & " " & (iYear - 1911) & " 年 1-6 月"
               ElseIf pCol = 13 Then '當年7-12月
                  stCon = stCon & " and cp05 between " & iYear & "0701 and " & iYear & "1231"
                  pLblCond = pStrFC & " " & (iYear - 1911) & " 年 7-12 月"
               ElseIf pCol = 18 Then '指定區間
                  stCon = stCon & " and cp05 between " & pStDate1 & " And " & pStDate2
                  pLblCond = pStrFC & " " & (pStDate1 - 19110000) & " － " & (pStDate2 - 19110000)
               Else
                  pLblCond = pStrFC & " 全部"
               End If
               
            Case 5, 7, 9, 11, 14, 19 'CF
               iSys = 2
               If pStrCF = "CFP" Then
                  'Modified by Lydia 2024/01/30 (Widen)因應目前互惠的策略已擴大到事務所層級，再麻煩進行程式的調整到所有國家地區>> 加抓P案
                  stCon = stCon & " and instr('CFP00,P00',cp01||cp04)>0  and (instr('" & NewCasePtyList & "', cp10) > 0 or cp10 like '3%') and cp09<'B'"
               ElseIf pStrCF = "CFT" Then
                  'Modified by Lydia 2024/01/30 (Widen)因應目前互惠的策略已擴大到事務所層級，再麻煩進行程式的調整到所有國家地區>> 加抓T案
                  stCon = stCon & " and instr('CFT00,T00',cp01||cp04)>0 and instr('101',cp10)>0 and cp09<'B'"
               End If
               '2013/5/24 End
               stCon = stCon & " and cp44='" & stCP44 & "'"
               '加以代理人統計
               If p_bolByAgent = False Then
                  If stCP116 = "" Then
                     stCon = stCon & " and cp116 is null"
                  Else
                     stCon = stCon & " and cp116='" & stCP116 & "'"
                  End If
               End If
               stCon = stCon & " and cp27<=" & stDate
               
               If pCol = 7 Then '前年
                  stCon = stCon & " and cp27 between " & (iYear - 2) & "0101 and " & (iYear - 2) & "1231"
                  pLblCond = pStrCF & " " & (iYear - 1911 - 2) & " 年"
               ElseIf pCol = 9 Then '去年
                  stCon = stCon & " and cp27 between " & (iYear - 1) & "0101 and " & (iYear - 1) & "1231"
                  pLblCond = pStrCF & " " & (iYear - 1911 - 1) & " 年"
               ElseIf pCol = 11 Then '當年1-6月
                  stCon = stCon & " and cp27 between " & iYear & "0101 and " & iYear & "0630"
                  pLblCond = pStrCF & " " & (iYear - 1911) & " 年 1-6 月"
               ElseIf pCol = 14 Then '當年7-12月
                  stCon = stCon & " and cp27 between " & iYear & "0701 and " & iYear & "1231"
                  pLblCond = pStrCF & " " & (iYear - 1911) & " 年 7-12 月"
               ElseIf pCol = 19 Then '指定區間
                  stCon = stCon & " and cp27 between " & pStDate1 & " And " & pStDate2
                  pLblCond = pStrCF & " " & (pStDate1 - 19110000) & " － " & (pStDate2 - 19110000)
               Else
                  pLblCond = pStrCF & " 全部"
               End If
         End Select
         If stCon = "" Then Exit Sub
         
         If iSys = 1 Then
            If pStrFC = "FCP" Then
               stVTB0 = "SELECT distinct PA01,PA02,PA03,PA04 FROM PATENT,CASEPROGRESS" & _
                  " where cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & stCon & _
                  " AND cp57 is null" & stCon
            ElseIf pStrFC = "FCT" Then
               stVTB0 = "SELECT distinct TM01,TM02,TM03,TM04 FROM Trademark,CASEPROGRESS" & _
                  " where cp01(+)=TM01 and cp02(+)=TM02 and cp03(+)=TM03 and cp04(+)=TM04" & stCon & _
                  " AND cp57 is null" & stCon
            End If
            
            'FCP只顯示案件盈虧
            If pStrFC = "FCP" Then
               stVTB1 = "SELECT distinct cp01,cp02,cp03,cp04,cp60 FROM PATENT,CASEPROGRESS" & _
                  " where cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
                  " AND cp57 is null and cp60 is not null" & stCon
            ElseIf pStrFC = "FCT" Then
               stVTB1 = "SELECT distinct cp01,cp02,cp03,cp04,cp60 FROM Trademark,CASEPROGRESS" & _
                  " where cp01(+)=TM01 and cp02(+)=TM02 and cp03(+)=TM03 and cp04(+)=TM04" & _
                  " AND cp57 is null and cp60 is not null" & stCon
            End If
               
            '--案件盈虧=請款金額-折讓-規費-作業失誤
            '請款金額-折讓-規費
            stVTB2 = "select cp01,cp02,cp03,cp04,sum(nvl(a1k11,0)-nvl(a1k06,0)-nvl(a1k09,0)) A1" & _
               " From (" & stVTB1 & ") X,acc1k0 a Where a1k01(+)=cp60" & _
               " group by X.cp01,X.cp02,X.cp03,X.cp04"
            
            If pStrFC = "FCP" Then
               '作業失誤=cp17(cp16=0 & cp17>0)
               stVTB3 = "select cp01,cp02,cp03,cp04,sum(cp17) B1" & _
                  " From (" & stVTB0 & ") X,caseprogress a" & _
                  " Where a.cp01(+)=X.pa01 and a.cp02(+)=X.pa02 and a.cp03(+)=X.pa03 and a.cp04(+)=X.pa04" & _
                  " and cp16=0 and cp17>0" & _
                  " group by cp01,cp02,cp03,cp04"
                  
               strA1 = "select X.pa01||'-'||X.pa02||decode(X.pa03||X.pa04,'000','','-'||X.pa03||'-'||X.pa04) C01" & _
                  ",Y.pa05 C02,to_char(nvl(A1,0)-nvl(B1,0),'9,999,999.00') C03,' ' AS SF, 0 AS SFV, A.CP01, A.CP02, A.CP03, A.CP04" & _
                  " from (" & stVTB0 & ") X,(" & stVTB2 & ") A,(" & stVTB3 & ") B,patent Y" & _
                  " where A.cp01(+)=X.pa01 and A.cp02(+)=X.pa02 and A.cp03(+)=X.pa03 and A.cp04(+)=X.pa04" & _
                  " and B.cp01(+)=X.pa01 and B.cp02(+)=X.pa02 and B.cp03(+)=X.pa03 and B.cp04(+)=X.pa04" & _
                  " and Y.pa01(+)=X.pa01 and Y.pa02(+)=X.pa02 and Y.pa03(+)=X.pa03 and Y.pa04(+)=X.pa04 order by 1,2"
            ElseIf pStrFC = "FCT" Then
               '作業失誤=cp17(cp16=0 & cp17>0)
               stVTB3 = "select cp01,cp02,cp03,cp04,sum(cp17) B1" & _
                  " From (" & stVTB0 & ") X,caseprogress a" & _
                  " Where a.cp01(+)=X.tm01 and a.cp02(+)=X.tm02 and a.cp03(+)=X.tm03 and a.cp04(+)=X.tm04" & _
                  " and cp16=0 and cp17>0" & _
                  " group by cp01,cp02,cp03,cp04"
                  
               strA1 = "select X.tm01||'-'||X.tm02||decode(X.tm03||X.tm04,'000','','-'||X.tm03||'-'||X.tm04) C01" & _
                  ",Y.tm05 C02,to_char(nvl(A1,0)-nvl(B1,0),'9,999,999.00') C03,' ' AS SF, 0 AS SFV, A.CP01, A.CP02, A.CP03, A.CP04" & _
                  " from (" & stVTB0 & ") X,(" & stVTB2 & ") A,(" & stVTB3 & ") B,Trademark Y" & _
                  " where A.cp01(+)=X.tm01 and A.cp02(+)=X.tm02 and A.cp03(+)=X.tm03 and A.cp04(+)=X.tm04" & _
                  " and B.cp01(+)=X.tm01 and B.cp02(+)=X.tm02 and B.cp03(+)=X.tm03 and B.cp04(+)=X.tm04" & _
                  " and Y.tm01(+)=X.tm01 and Y.tm02(+)=X.tm02 and Y.tm03(+)=X.tm03 and Y.tm04(+)=X.tm04 order by 1,2"
            End If
         Else
            stVTB0 = "SELECT CP01,CP02,CP03,CP04,max(CP31) SF FROM CASEPROGRESS" & _
               " WHERE cp57 is null" & stCon & " group by cp01,cp02,cp03,cp04"
               
            '--案件收(不扣安全基金)(若有銷規費時會錯，案件盈虧查詢也是)
            stVTB1 = "select X.cp01,X.cp02,X.cp03,X.cp04,sum(decode(nvl(cp16,0)-nvl(cp77,0),0,0,nvl(cp16,0)-nvl(CP18,0)*1000)) A1" & _
               " From (" & stVTB0 & ") X,caseprogress a Where a.cp01(+)=X.cp01 and a.cp02(+)=X.cp02 and a.cp03(+)=X.cp03 and a.cp04(+)=X.cp04" & _
               " and (cp16>0 or cp61 is not null)" & _
               " group by X.cp01,X.cp02,X.cp03,X.cp04"
      
            '--案件付(只能用本所號抓因為舊系統沒串到收文號)
            stVTB2 = "select cp01,cp02,cp03,cp04,sum(decode(A1G01,null,decode(a1901,null,AXF15,AXF04*A1906),AXF04*A1G03)) B1" & _
               " From (" & stVTB0 & ") X,acc151,acc150,acc190,acc1g0" & _
               " where axf03(+)=cp01||cp02||cp03||cp04 and a1501(+)=axf01 and a1507 is null" & _
               " and A1902(+)=AXF01 AND A1G01(+)=A1512" & _
               " group by cp01,cp02,cp03,cp04"
           
            If pStrFC = "FCP" Then
               strA1 = "select X.cp01||'-'||X.cp02||decode(X.cp03||X.cp04,'000','','-'||X.cp03||'-'||X.cp04) C01" & _
                  ",pa05 C02,to_char(nvl(A1,0)-nvl(B1,0),'9,999,999.00') C03,X.SF,0 SFV,X.cp01,X.cp02,X.cp03,X.cp04" & _
                  " from (" & stVTB0 & ") X,(" & stVTB1 & ") A,(" & stVTB2 & ") B,patent" & _
                  " where A.cp01(+)=X.cp01 and A.cp02(+)=X.cp02 and A.cp03(+)=X.cp03 and A.cp04(+)=X.cp04" & _
                  " and B.cp01(+)=X.cp01 and B.cp02(+)=X.cp02 and B.cp03(+)=X.cp03 and B.cp04(+)=X.cp04" & _
                  " and pa01(+)=X.cp01 and pa02(+)=X.cp02 and pa03(+)=X.cp03 and pa04(+)=X.cp04 order by 1,2"
            ElseIf pStrFC = "FCT" Then
               strA1 = "select X.cp01||'-'||X.cp02||decode(X.cp03||X.cp04,'000','','-'||X.cp03||'-'||X.cp04) C01" & _
                  ",tm05 C02,to_char(nvl(A1,0)-nvl(B1,0),'9,999,999.00') C03,X.SF,0 SFV,X.cp01,X.cp02,X.cp03,X.cp04" & _
                  " from (" & stVTB0 & ") X,(" & stVTB1 & ") A,(" & stVTB2 & ") B,Trademark" & _
                  " where A.cp01(+)=X.cp01 and A.cp02(+)=X.cp02 and A.cp03(+)=X.cp03 and A.cp04(+)=X.cp04" & _
                  " and B.cp01(+)=X.cp01 and B.cp02(+)=X.cp02 and B.cp03(+)=X.cp03 and B.cp04(+)=X.cp04" & _
                  " and tm01(+)=X.cp01 and tm02(+)=X.cp02 and tm03(+)=X.cp03 and tm04(+)=X.cp04 order by 1,2"
            End If
         End If
'--------------------------------
      End If
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strA1)
      If intA = 1 Then
         rsAD.MoveFirst
         Do While Not rsAD.EOF
            iPos = 0
            If Left("" & rsAD.Fields("c01"), 3) = "CFP" And "" & rsAD.Fields("sf") = "Y" Then
               iPos = GetFloatPrepareCase("" & rsAD.Fields("cp01"), "" & rsAD.Fields("cp02"), "" & rsAD.Fields("cp03"), "" & rsAD.Fields("cp04"))
            End If
            'R004欄位長度500，專門放案件名稱
            strTmp = "INSERT INTO rdatafactory (formname,id,seqno,rowseq,r001,r004,r003,r002,r005,r006,r007,r008,r009,r010) " & _
                     "values ('frm050408_2" & IIf(pFCno = "ALL", "-" & IIf(pKind = "1", "CFP", "CFT"), "") & "','" & strUserNum & "'," & intR & ", " & rsAD.AbsolutePosition & ", '" & rsAD.Fields("c01") & "', '" & ChgSQL(rsAD.Fields("c02")) & "', " & _
                     "'" & Format(Format(rsAD.Fields("c03")) - iPos, "#,###.00") & "', '" & rsAD.Fields("sf") & "','" & iPos & "', '" & rsAD.Fields("cp01") & "', '" & rsAD.Fields("cp02") & "', '" & rsAD.Fields("cp03") & "', '" & rsAD.Fields("cp04") & "', '" & stCP44 & "') "
            cnnConnection.Execute strTmp
            rsAD.MoveNext
         Loop
      End If
   Next intR
   Set rsAD = Nothing
   
End Sub

'Added by Lydia 2025/07/25 財務系統：國內應收帳款相關查詢 □排除未達客戶付款週期之應收帳款 的選項
Public Sub PUB_ProcAcctmp08(ByVal pFrmName As String, ByVal pUserId As String)
Dim strEx As String, intS As Integer
   
   If Trim(pFrmName) = "" Or Trim(pUserId) = "" Then
       Exit Sub
   Else
      '排除未達客戶付款週期之應收帳款;
      '國內應收帳款週期起算日：計算方式以發文日當月的翌月1日起計----參考PUB_GetBillDataAll
      strEx = "delete FROM acctmp08 WHERE T05='" & pFrmName & "' AND T14='" & pUserId & "' AND (t01,t02,t05,t06) IN (" & _
              "select t01,t02,t05,t06 From acctmp08, acc0k0, caseprogress, customer " & _
              "WHERE T05='" & pFrmName & "' AND T14='" & pUserId & "' AND t02=cp09(+) AND t01=a0k01(+) AND substr(a0k03,1,8)=cu01(+) AND substr(a0k03,9,1)=cu02(+) " & _
              "AND nvl(cp27,0)>0 AND to_char(add_months(to_date(substr(to_char(add_months(to_date(cp27,'yyyymmdd'),1),'yyyymmdd'),1,6)||'01','yyyymmdd'),nvl(cu175,2)),'yyyymmdd') >= to_char(SYSDATE,'yyyymmdd'))"
      cnnConnection.Execute strEx, intS
      '刪除未發文的帳款
      strEx = "delete FROM acctmp08 WHERE T05='" & pFrmName & "' AND T14='" & pUserId & "' AND (t01,t02,t05,t06) IN (" & _
              "select t01,t02,t05,t06 From acctmp08, caseprogress WHERE T05='" & pFrmName & "' AND T14='" & pUserId & "' AND t02=cp09(+) AND cp158=0)"
      cnnConnection.Execute strEx, intS
   End If
End Sub

'Added by Morgan 2025/7/31
'新增稽核日誌
'Modified by Morgan 2025/8/5 +pTableName,pKeyNo,pFileName
Public Function PUB_AddAuditLog(pType As String, Optional ByVal pDesc As String, Optional pOldValue As String, Optional pNewValue As String, Optional pTableName As String, Optional pKeyNo As String, Optional pFileName As String) As Boolean
   Dim stSQL As String, intR As Integer
   
   On Error GoTo ErrHnd
   
   If pType = "01" Then
      If pDesc = "" Then
         pDesc = "登入:" & App.EXEName
      End If
   ElseIf pType = "02" Then
      If pDesc = "" Then
         pDesc = "登出:" & App.EXEName
      End If
   ElseIf pType = "03" Then
      pDesc = "上傳:" & pDesc
   ElseIf pType = "04" Then
      pDesc = "下載:" & pDesc
   ElseIf pType = "05" Then
      pDesc = "刪除:" & pDesc
   End If
   
   stSQL = "insert into AuditLog(AL01,AL02,AL03,AL04,AL05,AL06,AL07,AL08,AL09,AL10,AL11)" & _
      " values(Systimestamp,'" & strUserNum & "',SYS_CONTEXT('USERENV','IP_ADDRESS')" & _
      ",'" & pType & "','" & ChgSQL(pDesc) & "','" & ChgSQL(pOldValue) & "','" & ChgSQL(pNewValue) & "'" & _
      ",SYS_CONTEXT('USERENV','TERMINAL'),'" & pTableName & "','" & pKeyNo & "','" & ChgSQL(pFileName) & "')"
      
   cnnConnection.Execute stSQL, intR
   PUB_AddAuditLog = True
   Exit Function
   
ErrHnd:
   
   If Pub_StrUserSt15 = "M51" Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

'Added by Morgan 2025/7/31
'Modified by Morgan 2025/8/4 原放實體檔名改為放DB顯示的檔名及收文號才比較可讀(Ex:電子公文收文後當日的實體仍是公文號，隔日才會轉檔名)
'Modified by Morgan 2025/8/5 改說明放FTP路徑,增加欄位放Table,單號,DB檔名
'Modified by Morgan 2025/8/6 +pLocalFilePath
Public Function PUB_AddFTPAuditLog(pType As String, pFtpFilePath As String, Optional pLocalFilePath As String) As Boolean
   Dim stFTPFileName As String, stTable As String, stFileName As String, stKeyNo As String
   Dim intP As Integer, stSQL As String, stSQL2 As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stDbPath As String
   Dim arrTMP1() As String, arrTmp2() As String
   
On Error GoTo ErrHnd
   If strUserNum <> "" And strUserNum <> "QPGMR" Then
      stFTPFileName = Mid(pFtpFilePath, InStrRev(pFtpFilePath, "/") + 1)
      'Added by Morgan 2025/8/29
      If pLocalFilePath <> "" Then
         stFileName = Mid(pLocalFilePath, InStrRev(pLocalFilePath, "\") + 1)
      Else
         stFileName = stFTPFileName
      End If
      'end 2025/8/29
      
      'Added by Morgan 2025/8/6 上傳時DB可能還沒有紀錄,先解析FTP檔名預設，再檢查是否有對應的紀錄
      intP = InStr(stFTPFileName, ".")
      If intP > 0 Then
         stKeyNo = Left(stFTPFileName, intP - 1)
         'stFileName = Mid(stFTPFileName, intP + 1) 'Removed by Morgan 2025/8/29 移到上面
      End If
      'end 2025/8/6
      
      intP = InStr(UCase(pFtpFilePath), "CASEPAPERPDF")
      If intP > 0 Then
         stTable = "CASEPAPERPDF"
         stDbPath = Mid(pFtpFilePath, intP + Len(stTable) + 1)
         'cpp14 可能會有 "\" 開頭
         stSQL = "select cpp01,cpp02 from CASEPAPERPDF where (cpp14='" & stDbPath & "' or cpp14='\" & stDbPath & "')"
         stSQL2 = "select cp09 from caseprogress where cp09='" & stKeyNo & "'"
      End If
      
      If intP = 0 Then
         intP = InStr(UCase(pFtpFilePath), "CASEPAPERFILE")
         If intP > 0 Then
            stTable = "CASEPAPERFILE"
            stDbPath = Mid(pFtpFilePath, intP + Len(stTable) + 1)
            stSQL = "select cpf01,cpf02 from CASEPAPERFILE where cpf13='" & stDbPath & "'"
            stSQL2 = "select cp09 from caseprogress where cp09='" & stKeyNo & "'"
         End If
      End If
      
      'Added by Morgan 2025/8/6
      If intP = 0 Then
         intP = InStr(UCase(pFtpFilePath), "CONTACTFILE")
         If intP > 0 Then
            stTable = "CONTACTFILE"
            stDbPath = Mid(pFtpFilePath, intP + Len(stTable) + 1)
            stSQL = "select cf01,cf02 from CONTACTFILE where cf06='" & stDbPath & "'"
            stKeyNo = Left(stKeyNo, 9)
            'Removed by Morgan 2025/8/29 移到上面
            'If pLocalFilePath <> "" Then
            '   stFileName = Mid(pLocalFilePath, InStrRev(pLocalFilePath, "\") + 1)
            'End If
            stSQL2 = "select cr01 from Contactrecord where cr01='" & stKeyNo & "' union select cor01 from Contactrecord1 where cor01='" & stKeyNo & "'"
         End If
      End If
      'end 2025/8/6
      
      'Added by Morgan 2025/8/29
      If intP = 0 Then
         intP = InStr(UCase(pFtpFilePath), "CONTRACT")
         If intP > 0 Then
            stTable = "CONTRACT"
            stDbPath = Mid(pFtpFilePath, intP + Len(stTable) + 1)
            stSQL = "select ct01,ct08,ct09 from CONTRACT where instr(ct09,'" & stDbPath & "')>0"
            stKeyNo = Left(stKeyNo, 6)
            stSQL2 = "select ct01 from CONTRACT where ct01='" & stKeyNo & "'"
         End If
      End If
      'end 2025/8/29
      
      If stTable <> "" Then
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            stKeyNo = rsQuery(0)
            If stTable = "CONTRACT" Then
               arrTMP1 = Split(rsQuery(1), ",")
               arrTmp2 = Split(rsQuery(2), ",")
               For intP = LBound(arrTmp2) To UBound(arrTmp2)
                  If arrTmp2(intP) = stDbPath Then
                     stFileName = arrTMP1(intP)
                     Exit For
                  End If
               Next
            Else
               stFileName = rsQuery(1)
            End If
         'Added by Morgan 2025/8/6 上傳時DB可能還沒有紀錄,解析FTP檔名內的編號檢查是否有對應的紀錄
         Else
            intQ = 1
            Set rsQuery = ClsLawReadRstMsg(intQ, stSQL2)
            If intQ <> 1 Then '單號不存在時清除變數
               stKeyNo = ""
               'stFileName = "" 'Removed by Morgan 2025/8/29
            End If
         'end 2025/8/6
         End If
         'Modified by Morgan 2025/8/29
         'PUB_AddFTPAuditLog = PUB_AddAuditLog(pType, stFTPFileName, , , stTable, stKeyNo, stFileName)
         PUB_AddFTPAuditLog = PUB_AddAuditLog(pType, stFileName, , , stTable, stKeyNo, stFTPFileName)
      End If
   End If
   Set rsQuery = Nothing
   Exit Function
   
ErrHnd:
   If Pub_StrUserSt15 = "M51" Then
      MsgBox Err.Description, vbCritical
   End If
   Set rsQuery = Nothing
End Function

'Added by Lydia 2025/08/08 國外往來記錄的維護及查詢限制
Public Function Pub_GetCRExceptNo(ByVal pFrmName As String) As String
Dim strQ1 As String, intQ As Integer, strMid As String
Dim rsAD As New ADODB.Recordset
   
   Pub_GetCRExceptNo = ""
   If UCase(pFrmName) = "FRM140404" Then
      '維護時，KB4000830只有建檔人李道昀B0004及其各級主管可以看到及維護此筆記錄，電腦中心人員也不能維護。
      strQ1 = "select cr01 from contactrecord,staff where cr01='KB4000830' and cr12=st01(+) and instr(cr12||','||st52||','||st53||','||st54,'" & strUserNum & "') > 0 "
      intQ = 1
      Set rsAD = ClsLawReadRstMsg(intQ, strQ1)
      If intQ = 0 Then
         Pub_GetCRExceptNo = "KB4000830"
      End If
   Else
      '共同查詢時，往來記錄區之KB4000830只有下列人員可以開附件、可以看往來記錄內容，電腦中心人員也不能看。
      '所長(81040)、總經理(94007)、文雄特助(A4023)、岱嫻特助(B1015)
      '工程師: Wilson(87003), Stellar(97031), Alina(99025)及Red(A0022)
      '承辦: David(77015) , Anny(A4011), Lisa(A6035), Tim(A5023) & Kahn(B0004)
      If InStr("81040,94007,A4023,B1015,87003,97031,99025,A0022,77015,A4011,A6035,A5023,B0004", strUserNum) = 0 Then
         Pub_GetCRExceptNo = "KB4000830"
      End If
   End If
   Set rsAD = Nothing
End Function

'Added by Morgan 2025/9/11
'大陸案要主張台灣優先權提醒
Public Function PUB_CNPriorityMsg() As Integer
   PUB_CNPriorityMsg = MsgBox("若要主張台灣優先權，則台灣案及大陸案的申請人中，均須包含有國籍為台灣的申請人，若有多位申請人，則第一申請人必須是台灣籍！" & vbCrLf & vbCrLf & "是否仍要繼續？", vbExclamation + vbYesNo + vbDefaultButton2, "大陸案要主張台灣優先權提醒")
End Function

'Added by Morgan 2025/9/11
'大陸案主張台灣優先權檢查:第1申請人要是台灣籍且被主張的台灣案要有台灣籍的申請人
Public Function PUB_ChkCNTWPriority(ByRef pCaseNo() As String, Optional ByRef pCountryList As String, Optional ByRef pAppNoList As String) As Boolean
   Dim arrCountry() As String, arrAppNo() As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim ii As Integer
   
   '檢查第一申請人是否台灣籍
   stSQL = "select cu10 from patent,customer where pa01='" & pCaseNo(1) & "' and pa02='" & pCaseNo(2) & "' and pa03='" & pCaseNo(3) & "' and pa04='" & pCaseNo(4) & "'" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and cu10<'010'"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      If pCountryList = "" Then
         stSQL = "select pd06,pd07 from pridate where pd01='" & pCaseNo(2) & "' and pd01='" & pCaseNo(2) & "' and pd03='" & pCaseNo(3) & "' and pd04='" & pCaseNo(4) & "' and pd07='000'"
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            With rsQuery
            Do While Not .EOF
               pCountryList = pCountryList & .Fields("pd07") & "，"
               pAppNoList = pAppNoList & .Fields("pd06") & "，"
               .MoveNext
            Loop
            End With
         End If
      End If
      If pCountryList <> "" Then
         arrCountry = Split(pCountryList, "，")
         arrAppNo = Split(pAppNoList, "，")
         For ii = LBound(arrCountry) To UBound(arrCountry)
            If arrCountry(ii) = "000" Then
               stSQL = "select pa26 from patent where pa09='000' and pa11='" & arrAppNo(ii) & "' and exists(select * from customer where instr(pa26||pa27||pa28||pa29||pa30,cu01||cu02)>0 and cu10<'010')"
               intQ = 1
               Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
               If intQ = 1 Then
                  PUB_ChkCNTWPriority = True
               Else
                  PUB_ChkCNTWPriority = False
                  GoTo ExitFlag
               End If
            End If
         Next
      End If
   End If
   
ExitFlag:
   Set rsQuery = Nothing
End Function

'Added by Morgan 2025/9/11
'台灣優先權證明書用於中國大陸案提醒
Public Function PUB_TWPriCertMsg() As Integer
   PUB_TWPriCertMsg = MsgBox("因全部申請人均為外國籍無法用於主張中國大陸申請案，請確認此優先權證明書是否用於中國大陸？" & vbCrLf & vbCrLf & "「是」取消，「否」繼續。", vbQuestion + vbYesNo + vbDefaultButton1, "台灣優先權證明書用於中國大陸案提醒")
End Function

'Added by Morgan 2025/9/12
'檢查案件申請人是否無台灣籍
Public Function PUB_ChkNoTWApp(ByRef pCaseNo() As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select * from patent where pa01='" & pCaseNo(1) & "' and pa02='" & pCaseNo(2) & "' and pa03='" & pCaseNo(3) & "' and  pa04='" & pCaseNo(4) & "'" & _
      " and not exists(select * from customer where cu01||cu02 in (pa26,pa27,pa28,pa29,pa30) and cu10<'010')"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      PUB_ChkNoTWApp = True
   End If
   
   Set rsQuery = Nothing
End Function

'Added by Lydia 2025/09/19 取得共同查詢案件統計之新申請案性質(是否只統計新申請案)的SQL
Public Function PUB_GetForNewCaseSql(ByVal pKind As String) As String
   PUB_GetForNewCaseSql = ""
   Select Case pKind
      Case "1" '專利
         PUB_GetForNewCaseSql = " (INSTR('" & NewCasePtyList & "',CP10)>0 OR SUBSTR(CP10,1,1)='3') "
      Case "2" '商標
         PUB_GetForNewCaseSql = " AND CP10='101' "
      Case "5" '服務
         'Memo by Lydia 2024/12/02 若統計類別=5的條件有異動，請一併變更frm090642-收文量、發文量
         PUB_GetForNewCaseSql = " AND INSTR('801,802,805,806',CP10)>0 "
   End Select
End Function

'Added by Lydia 2018/06/13 對外翻譯(提申本)：取得原文/翻譯語種
'Move by Lydia 2022/09/08 從basPublic搬過來
'Modified by Morgan 2025/11/10 從service1搬過來(財務也要用)
Public Function Pub_GetTransFeeL(ByVal pType As String, ByVal pKind As String) As String
'pType =1.原文語種 2.翻譯語種
Dim strDesc As String

    If pType = "1" Then
           '原文語種
           Select Case pKind
                Case "1": strDesc = "英文"
                Case "2": strDesc = "日文"
                Case "3": strDesc = "德文"
                Case "4": strDesc = "韓文" 'Added by Lydia 2024/02/21
                Case "5": strDesc = "中文" 'Added by Morgan 2025/11/10 CFP會用
                Case Else: strDesc = ""
           End Select
    ElseIf pType = "2" Then
           '翻譯語種
           Select Case pKind
                Case "1": strDesc = "繁體中文"
                Case "2": strDesc = "簡體中文"
                Case "3": strDesc = "日文" 'Added by Morgan 2025/11/10 CFP會用
                Case "4": strDesc = "德文" 'Added by Morgan 2025/11/10 CFP會用
                Case Else: strDesc = ""
           End Select
    End If
    Pub_GetTransFeeL = strDesc
End Function
