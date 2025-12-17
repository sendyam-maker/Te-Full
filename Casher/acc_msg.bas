Attribute VB_Name = "acc_msg"
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

'*************************************************
'  共用 Combo 內碼對照函式
'
'*************************************************
Public Function ComboItem(InputIndex As Integer) As String
   Select Case InputIndex
      Case 1
         ComboItem = "1--借方"
      Case 2
         ComboItem = "2--貸方"
      Case 11
         ComboItem = "1--支票"
      Case 12
         ComboItem = "2--本票"
      Case 13
         ComboItem = "3--匯票"
      Case 21
         ComboItem = "1--在途中"
      Case 22
         ComboItem = "2--已兌現"
      Case 31
         ComboItem = "1--繼續追蹤"
      Case 32
         ComboItem = "2--不管制"
      Case 41
         ComboItem = "1--轉應付款"
      Case 42
         ComboItem = "2--轉暫收款"
      Case 43
         ComboItem = "3--本所繳納"
      Case 51
         ComboItem = "1--IR"
      Case 52
         ComboItem = "2--CB"
      Case 53
         ComboItem = "3--託收"
      Case 54
         ComboItem = "4--其他銀行"
      Case 55
         ComboItem = "5--其他"
      Case 61
         ComboItem = "110204"
      Case 62
         ComboItem = "110205"
      Case 63
         ComboItem = "110208"
      Case 64
         ComboItem = "110216"
      Case 65
         ComboItem = "110222"
      Case 66
         ComboItem = "113001"
      Case 67
         ComboItem = "113002"
      Case 68
         ComboItem = "1202"
      Case 69
         ComboItem = "2401"
      Case 70
         ComboItem = "611301"
      Case 71
         ComboItem = "1--票匯"
      Case 72
         ComboItem = "2--電匯"
      Case 73
         ComboItem = "3--旅行支票"
      Case 74
         ComboItem = "4--現金"
      Case 75
         ComboItem = "5--商務卡"
      Case 76
         ComboItem = "6--其他"
      Case 81
         ComboItem = "1--簽收"
      Case 82
         ComboItem = "2--寄出"
      Case 83
         ComboItem = "3--寄分所"
      Case 84
         ComboItem = "4--其它"
      Case 85
         ComboItem = "5--寄出不印名條"
      Case 91
         ComboItem = "1--廠商"
      Case 92
         ComboItem = "2--客戶"
      Case 93
         ComboItem = "3--員工"
      Case 101
         ComboItem = "1--應付款"
      Case 102
         ComboItem = "2--銷退轉入"
      Case 103
         ComboItem = "3--補扣繳退費"
      Case 111
         ComboItem = "1--核駁退費"
      Case 112
         ComboItem = "2--溢收款"
      Case 113
         ComboItem = "3--案件未辦退費"
      Case 114
         ComboItem = "4--扣繳"
      Case 115
         ComboItem = "5--貨款"
      Case 116
         ComboItem = "6--稅款繳款書"
      Case 117
         ComboItem = "7--其他"
      Case 121
         ComboItem = "中"
      Case 122
         ComboItem = "英"
      Case 123
         ComboItem = "日"
      Case 131
         ComboItem = "1--客戶"
      Case 132
         ComboItem = "2--廠商"
      Case 133
         ComboItem = "3--員工"
      Case 134
         ComboItem = "4--其他"
      Case 141
         ComboItem = "1--未沖平"
      Case 142
         ComboItem = "2--沖平"
      Case 143
         ComboItem = "3--全部"
      Case 151
         ComboItem = "1--上半年"
      Case 152
         ComboItem = "2--下半年"
      Case 161
         ComboItem = "1--商標"
      Case 162
         ComboItem = "2--專利"
      Case 163
         ComboItem = "3--桂律師"
      Case 164
         ComboItem = "5--詹律師"
      Case 165
         ComboItem = "7--蔣律師"
      Case 166
         ComboItem = "8--唐律師"
      Case 171
         ComboItem = "FC"
      Case 172
         ComboItem = "CF"
      Case 173
         ComboItem = "FCT"
      Case 174
         ComboItem = "FCP"
      Case 175
         ComboItem = "FCL"
      Case 176
         ComboItem = "T"
      Case 177
         ComboItem = "P"
      Case 181
         ComboItem = "R--應收"
      Case 182
         ComboItem = "P--應付"
      Case 191
         ComboItem = "1--台中"
      Case 192
         ComboItem = "2--台南"
      Case 193
         ComboItem = "3--高雄"
      Case 201
         ComboItem = "1--暫收款"
      Case 202
         ComboItem = "2--溢收轉入"
      Case 203
         ComboItem = "3--退費轉入"
      Case 211
         ComboItem = "1-轉應付款"
      Case 212
         ComboItem = "2-轉暫收款"
      Case 221
         ComboItem = "1--沖暫收款"
      Case 222
         ComboItem = "2--其它"
      Case 231
         ComboItem = "1--本所案號"
      Case 232
         ComboItem = "2--客戶/廠商/員工"
      Case 233
         ComboItem = "3--智權人員"
      Case 234
         ComboItem = "4--其它"
      Case 241
         ComboItem = "結餘明細"
      Case 242
         ComboItem = "點數明細"
      Case 251
         ComboItem = "結匯明細表"
      Case 252
         ComboItem = "匯票函件明細表"
      Case 253
         ComboItem = "付款明細表"
   End Select
End Function

'*************************************************
'  報表內碼對照函式
'
'*************************************************
Public Function ReportTitle(InputIndex As Integer) As String
   Select Case InputIndex
      Case 101
         ReportTitle = "***  收文與收據資料檢核表  ***"
      Case 103
         ReportTitle = "***  付款工作底稿  ***"
      Case 1041
         ReportTitle = "***  客戶往來對帳單  ***"
      Case 1042
         ReportTitle = "***  客戶應收帳款對帳單  ***"
      Case 105
         ReportTitle = "***  客戶帳款明細表  ***"
      Case 106
         ReportTitle = "***  智權人員帳款明細表  ***"
      Case 107
         ReportTitle = "***  智權人員應收規費明細表  ***"
      Case 1081
         ReportTitle = "***  客戶帳齡分析表  ***"
      Case 1082
         ReportTitle = "***  智權人員帳齡分析表  ***"
      Case 1083
         ReportTitle = "***  智權人員客戶帳齡分析表  ***"
      Case 109
         ReportTitle = "***  銷帳退費明細表  ***"
      Case 110
         ReportTitle = "***  暫收款明細表  ***"
      Case 111
         ReportTitle = "***  支票寄出明細清單  ***"
      Case 1111
         ReportTitle = "***  付款通知單  ***"
      Case 1112
         ReportTitle = "***  付款簽收簿  ***"
      Case 1113
         ReportTitle = "***  地址條  ***"
      Case 1114
         ReportTitle = "***  票據受領收據  ***"
      Case 112
         ReportTitle = "***  應付款統計表  ***"
      Case 113
         ReportTitle = "***  國內付款明細表  ***"
      Case 114
         ReportTitle = "***  智權人員別客戶扣繳稅款明細表  ***"
      Case 1141
         ReportTitle = "***  扣繳催收明細進度表  ***"
      Case 115
         ReportTitle = "***  回　執　單  ***"
      Case 1151
         ReportTitle = "***  繳款書寄出明細  ***"
      Case 116
         ReportTitle = "***  收據作廢明細表  ***"
      Case 204
         ReportTitle = "***  結匯明細表  ***"
      Case 209
         ReportTitle = "***  代理人對帳單  ***"
      Case 210
         ReportTitle = "***  國外帳齡分析表  ***"
      Case 211
         ReportTitle = "***  代理人帳目排名  ***"
      Case 2121
         ReportTitle = "***  智權人員業績請款點數統計表  ***"
      Case 2122
         ReportTitle = "***  智權人員業績收款點數統計表  ***"
      Case 213
         ReportTitle = "***  代理人FC帳款明細表  ***"
      'Add By Cheng 2002/09/02
      Case 2131
         ReportTitle = "***  國外FC帳款明細表  ***"
      Case 214
         ReportTitle = "***  國外應收規費及服務費分析表  ***"
      Case 215
         ReportTitle = "***  代理人逾期帳款分析表  ***"
      Case 216
         ReportTitle = "***  國內未收款明細表  ***"
      Case 217
         ReportTitle = "***  代理人未收未付對照表  ***"
      Case 218
         ReportTitle = "***  付款明細草稿  ***"
      Case 219
         ReportTitle = "***  匯票函件明細表  ***"
      Case 225
         ReportTitle = "台銀外幣兌現明細表"
      Case 301
         ReportTitle = "***  應收票據資料表  ***"
      Case 302
         ReportTitle = "***  應付票據資料表  ***"
      Case 303
         ReportTitle = "***  託收票據資料表  ***"
      Case 304
         ReportTitle = "***  銀行帳號別票據彙總表  ***"
      Case 305
         ReportTitle = "***  銀行帳號別票據明細表  ***"
      Case 306
         ReportTitle = "***  兌現日別資金流動彙總表  ***"
      Case 307
         ReportTitle = "***  兌現日別票據明細表  ***"
      Case 308
         ReportTitle = "***  往來對象別票據彙總表  ***"
      Case 309
         ReportTitle = "***  往來對象別票據明細表  ***"
      Case 310
         ReportTitle = "***  退票資料表  ***"
      Case 311
         ReportTitle = "***  抽票資料表  ***"
      Case 312
         ReportTitle = "***  票據貼現資料檢核表  ***"
      Case 313
         ReportTitle = "***  銀行帳號別資金流動表  ***"
      Case 314
         ReportTitle = "***  日期別資金流動預測表  ***"
      Case 315
         ReportTitle = "***  銀行調節資料表  ***"
      Case 316
         ReportTitle = "***  銀行別資料表  ***"
      Case 317
         ReportTitle = "***  甲存支票未兌領明細表  ***"
      Case 401
         ReportTitle = "***  日計表  ***"
      Case 402
         ReportTitle = "***  會計科目代號對照表  ***"
      Case 403
         ReportTitle = "***  科目餘額表  ***"
      Case 404
         ReportTitle = "***  科目明細表(對沖)  ***"
      Case 405
         ReportTitle = "***  試算表  ***"
      Case 406
         ReportTitle = "***  科目分類帳  ***"
      Case 407
         ReportTitle = "***  損益表  ***"
      Case 408
         ReportTitle = "***  損益比較表  ***"
      Case 409
         ReportTitle = "***  資產負債表  ***"
      Case 410
         ReportTitle = "***  預算實績比較表  ***"
      Case 411
         ReportTitle = "***  部門費用統計表  ***"
      Case 412
         ReportTitle = "***  部門損益表(子科目)  ***"
      Case 413
         ReportTitle = "***  年度損益統計表  ***"
      Case 414
         ReportTitle = "***  年度部門損益統計表  ***"
      Case 415
         ReportTitle = "***  資產負債比較表  ***"
      Case 416
         ReportTitle = "***  部門損益表  ***"
      Case 417
         ReportTitle = "***  智權人員點數明細表  ***"
      Case 418
         ReportTitle = "***  預算資料表  ***"
      Case 419
         ReportTitle = "***  費用科目分攤比率表  ***"
      Case 420
         ReportTitle = "扣繳憑單核對表"
      Case 421
         ReportTitle = "***  扣繳憑單明細表  ***"
      Case 422
         ReportTitle = "***  客戶扣繳明細核對表  ***"
      Case 423
         ReportTitle = "***  智權人員結餘點數總表  ***"
      Case 424
         ReportTitle = "月份專業點數明細表"
   End Select
End Function

'*************************************************
'  電話區域內碼對照函式
'
'*************************************************
Public Function TelLocalNo(InputIndex As Integer) As String
   Select Case InputIndex
      Case 1
         TelLocalNo = "02"
      Case 2
         TelLocalNo = "03"
      Case 3
         TelLocalNo = "035"
      Case 4
         TelLocalNo = "037"
      Case 5
         TelLocalNo = "038"
      Case 6
         TelLocalNo = "039"
      Case 7
         TelLocalNo = "04"
      Case 8
         TelLocalNo = "049"
      Case 9
         TelLocalNo = "05"
      Case 10
         TelLocalNo = "06"
      Case 11
         TelLocalNo = "07"
      Case 12
         TelLocalNo = "08"
      Case 13
         TelLocalNo = "0823"
      Case 14
         TelLocalNo = "089"
   End Select
End Function

'Modify By Cheng 2003/02/13
'修改傳入參數型態
'Public Function ReportSum(InputIndex As Integer) As String
Public Function ReportSum(InputIndex As Double) As String
   Select Case InputIndex
      Case 1
         ReportSum = "營業收入:"
      Case 2
         ReportSum = "營業支出:"
      Case 3
         ReportSum = "營業損益:"
      Case 4
         ReportSum = "－－－－－－"
      Case 5
         ReportSum = "營業外收入:"
      Case 6
         ReportSum = "營業外支出:"
      Case 7
         ReportSum = "稅前淨損益:"
      Case 8
         ReportSum = "＝＝＝＝＝＝"
      Case 9
         ReportSum = "資產總額:"
      'Add By Cheng 2002/01/18
      Case 9001
         ReportSum = "*** 資產總額 ***"
      Case 10
         ReportSum = "負債小計:"
      'Add By Cheng 2002/01/18
      Case 10001
         ReportSum = "*** 負債小計 ***"
      Case 11
         ReportSum = "本期損益:"
      Case 12
         ReportSum = "股東權益小計:"
      'Add By Cheng 2002/01/18
      Case 12001
         ReportSum = "*** 股東權益小計 ****"
      Case 13
         ReportSum = "負債總額:"
      'Add By Cheng 2002/01/18
      Case 13001
         ReportSum = "*** 負債與股東權益總額 ***"
      Case 14
         ReportSum = "實際營業收入:"
      Case 15
         ReportSum = "費用合計:"
      Case 16
         ReportSum = "部門損益:"
      Case 17
         ReportSum = "分攤費用:"
      Case 18
         ReportSum = "各部門營業損益:"
      Case 19
         ReportSum = "營業外收支:"
      Case 20
         ReportSum = "全所損益:"
      Case 21
         ReportSum = "部門經營損益:"
      Case 22
         ReportSum = "資產合計:"
      'Add By Cheng 2002/01/18
      Case 22001
         ReportSum = "*** 資產合計 ***"
      Case 23
         ReportSum = "負債與股東權益合計:"
      'Add By Cheng 2002/01/18
      Case 23001
         ReportSum = "*** 負債與股東權益合計 ***"
      Case 24
         ReportSum = "小計:"
      Case 25
         ReportSum = "合計:"
      Case 26
         ReportSum = "筆"
      Case 27
         ReportSum = "統計日期:"
      Case 28
         ReportSum = "至"
      Case 29
         ReportSum = "對沖代號(業)"
      Case 30
         ReportSum = "傳票編號"
      Case 31
         ReportSum = "對沖代號(客)"
      Case 32
         ReportSum = "對沖代號(本)"
      Case 33
         ReportSum = "摘要"
      Case 34
         ReportSum = "金額"
      Case 35
         ReportSum = "製表日期: "
      Case 36
         ReportSum = "頁　　次: "
      Case 37
         ReportSum = "付款行庫: "
      Case 38
         ReportSum = "付款帳號: "
      Case 39
         ReportSum = "支票號碼: "
      Case 40
         ReportSum = "到 期 日:   "
      Case 41
         ReportSum = "金　　額: "
      Case 42
         ReportSum = "備　　註: "
      Case 43
         ReportSum = "     台 鑒:"
      Case 44
         ReportSum = "茲 寄 上 應 付    台 端  ( 貴 公 司 )  之 票 據  ( 詳 述 如 下 )  ， 並 將 票 據 受 領 收 據"
      Case 45
         ReportSum = "填 妥 寄 回 ， 謝 謝 您 的 支 持 與 合 作 。"
      Case 46
         ReportSum = "特 別 說 明 : "
      Case 47
         ReportSum = "( 一 ) 敬 請 於 票 據 受 領 收 據 上 簽 蓋    台 端  ( 貴 公 司 )  之 收 款 章 ！ ！"
      Case 48
         ReportSum = "( 二 ) 票 據 受 領 收 據 若 未 寄 回 者 ， 以 後    台 端  ( 貴 公 司 )  之 款 項 ，"
      Case 49
         ReportSum = "　　  恕 不 再 郵 寄 送 達 ！ ！"
      Case 50
         ReportSum = "請 延 此 虛 線 撕 下 寄 回"
      Case 51
         ReportSum = "茲 收 到    貴 事 務 所 寄 來 之 票 據  ( 詳 述 如 下 )  ， 一 切 無 誤 ， 特 此 證 明 。"
      Case 52
         ReportSum = "茲 寄 上    貴 公 司 之 各 類 所 得 稅 扣 繳 稅 款 繳 款 書 　　　　 份 ， 金 額 共 計"
      Case 53
         ReportSum = "元 整 ，請 查 收 ， 並 回 執 單 上 蓋 章 後 寄 回 本 事 務 所 ， 謝 謝 您 的 合 作 。"
      Case 54
         ReportSum = "茲 收 到    貴 事 務 所 寄 來 之 各 類 所 得 稅 扣 繳 稅 款 繳 款 書 共 　　　　 份 ， 金 額 共 計"
      Case 55
         ReportSum = "元 整 ， 一 切 無 誤 。 特 此 證 明 。"
      Case 56
         ReportSum = "簽　　收　　人　："
      Case 57
         ReportSum = "智權人員"
      Case 58
         ReportSum = "業務達成點數"
      Case 59
         ReportSum = "加轉撥點數"
      Case 60
         ReportSum = "減轉撥點數"
      Case 61
         ReportSum = "保留點數"
      Case 62
         ReportSum = "實際達成點數"
      Case 63
         ReportSum = "台北所"
      Case 644
         ReportSum = "其它"
      Case 65
         ReportSum = "國內"
      Case 66
         ReportSum = "全所"
      Case 67
         ReportSum = "FCP"
      Case 68
         ReportSum = "FCT"
      Case 69
         ReportSum = "FCL"
      Case 70
         ReportSum = "國外"
      Case 71
         ReportSum = "Name of Bank: Bank of Taiwan, Head Office Foreign Department"
        'Add By Cheng 2003/02/13
      Case 71001
         ReportSum = "Name of Bank: Bank of Taiwan, Head Office"
      Case 72
         ReportSum = "Address: 120, Sec. 1, Chungking S. Rd., Taipei, Taiwan, R.O.C."
      Case 73
         ReportSum = "S.W.I.F.T. Address: BKTW TWTP"
        'Add By Cheng 2003/02/13
      Case 73001
         ReportSum = "S.W.I.F.T. Code: BKTW TWTP"
      Case 74 '美金帳戶
        'Modify By Cheng 2003/07/25
'         ReportSum = "Account No.: 006007052643 (for US currency)"
         ReportSum = "Account No.: 003007052646 (for US currency)"
      Case 75
         ReportSum = "Currency Rate: USD1.00=NTD"
      Case 76
         ReportSum = "扣繳年度:"
      Case 77
         ReportSum = "已扣金額"
      Case 78
         ReportSum = "已收扣單"
      Case 79
         ReportSum = "已收現金"
      Case 80
         ReportSum = "列呆帳"
      Case 81
         ReportSum = "催收中"
      Case 82
         ReportSum = "轉列下年度"
      Case 83
        'Modify By Cheng 2003/03/06
'         ReportSum = "To our professional service charges for:"
         ReportSum = "To our professional service charges for "
      Case 84
         ReportSum = "Re: Taiwanese "
        'Add By Cheng 2003/03/27
      Case 84001
         ReportSum = "Re: China "
      Case 85
         ReportSum = "Account Name: Tai E International Patent & Law Office"
      Case 86
         ReportSum = "PS: Please return copy of invoice(s) or indicate invoice number(s) paid with remittance"
        'Add By Cheng 2003/05/19
      Case 86001
         ReportSum = "PS: Please return a copy of the invoice(s) or indicate the invoice number(s) paid with remittance"
      Case 87
         ReportSum = "Gentlemen:"
      Case 88
         ReportSum = "We are sending you the attached bank draft(s) in cover of your debit note(s)"
      Case 89
         ReportSum = "detailed hereunder."
      Case 90
         ReportSum = "Please acknowledge safe receipt of the above-mentioned payment. It would"
      Case 91
         ReportSum = "be appreciated if you could mention our reference number while sending as your"
      Case 92
         ReportSum = "debit notes or statements."
      Case 93
         ReportSum = "With best regards."
      Case 94
         ReportSum = "Sincerely yours,"
      Case 95
         ReportSum = "Tai E International"
      Case 96
         ReportSum = "Patent & Law Office"
      Case 97
         ReportSum = "A remittance has been effected through our bank, to settle your debit"
      Case 97001
         ReportSum = "We inform you that we duly transferred the amounts listed below to your "
      Case 98
         ReportSum = "notes(invoices) as follows :"
      Case 98001
         ReportSum = "bank account, i.,e. "
      Case 99
         ReportSum = "Please acknowledges safe receipt thereof, and we remain."
      Case 100
         ReportSum = "We acknowledge with thanks receipt of your payment as identified below:"
      Case 101
         ReportSum = "合計"
      Case 102
         ReportSum = "地址: "
      Case 103
         ReportSum = "電話: "
      Case 104
         ReportSum = "台北市中山區長安東路二段112號9樓"
      Case 105
         ReportSum = "台北所合計:"
      Case 106
         ReportSum = "台中所合計:"
      Case 107
         ReportSum = "台南所合計:"
      Case 108
         ReportSum = "高雄所合計:"
      Case 109
         ReportSum = "We reimburse the redundant payment to you. Please find enclosed our Credit"
      Case 110
         ReportSum = " Note No. "
      Case 111
         ReportSum = ". Please acknowledge receipt of this Credit Note."
      Case 112
         ReportSum = "If you have any questions concerning this matter, please do not hesitate to"
      Case 113
         ReportSum = " contact us."
      Case 114
         'Modify by Morgan 2006/7/6
         'ReportSum = "I-Chu Lin"
         ReportSum = "Fred C. T. Yen"
      Case 115
         ReportSum = "Patent Attorney"
      Case 116
         ReportSum = "Tai E International Patent & Law Office"
      Case 117
         'Modify by Morgan 2006/7/6
         'ReportSum = "ICL/dy"
         ReportSum = "CTY/dy"
      Case 118
         ReportSum = "Encl."
      Case 119
         ReportSum = "Reimbursing the redundant payment to you"
      Case 120
         ReportSum = "Total"
      Case 121
         ReportSum = "Account No.: 003001305688 (for Taiwan currency)"
      Case 122
         ReportSum = "銀行: 中國工商銀行上海徐匯支行  天鈅橋路儲蓄所"
      Case 123
         ReportSum = "賬戶名稱: 汪家翰 (人民幣個人賬戶)"
      Case 124
         ReportSum = "賬號: 47271010301*0"
      Case 125
         ReportSum = "※ 貴所可將款項匯至本所上海或台灣之銀行賬戶，惟於匯款後請"
      Case 126
         ReportSum = "     務必知匯台北總所，並告知匯款金額。"
      Case 127
         ReportSum = "廣東所合計:"
        'Add By Cheng 2003/02/07
        '可以歐元支付
      Case 128
         ReportSum = "Payment by EURO is acceptable"
        'Add By Cheng 2003/02/13
      Case 129 '歐元帳戶
        'Modify By Cheng 2003/07/25
'         ReportSum = "Account No.: 006007085124 (for EURO currency)"
         ReportSum = "Account No.: 003007085127 (for EURO currency)"
   End Select
End Function

'*************************************************
'  程式訊息內碼對照函式(中文)
'
'*************************************************
Public Function ShowNumberWord(InputNumber As Long) As String
   Select Case InputNumber
      Case 0
         ShowNumberWord = "零"
      Case 1
         ShowNumberWord = "壹"
      Case 2
         ShowNumberWord = "貳"
      Case 3
         ShowNumberWord = "參"
      Case 4
         ShowNumberWord = "肆"
      Case 5
         ShowNumberWord = "伍"
      Case 6
         ShowNumberWord = "陸"
      Case 7
         ShowNumberWord = "柒"
      Case 8
         ShowNumberWord = "捌"
      Case 9
         ShowNumberWord = "玖"
      Case 10
         ShowNumberWord = "拾"
      Case 11
         ShowNumberWord = "佰"
      Case 12
         ShowNumberWord = "仟"
      Case 13
         ShowNumberWord = "萬"
      Case 14
         ShowNumberWord = "億"
      Case 20
         ShowNumberWord = "元整"
   End Select
End Function

'*************************************************
'  程式訊息內碼對照函式(英文)
'
'*************************************************
Public Function ShowNumber(InputNumber As Long) As String
   Select Case InputNumber
      Case 0
         ShowNumber = "ZEROS"
      Case 1
         ShowNumber = "ONE"
      Case 2
         ShowNumber = "TWO"
      Case 3
         ShowNumber = "THREE"
      Case 4
         ShowNumber = "FOUR"
      Case 5
         ShowNumber = "FIVE"
      Case 6
         ShowNumber = "SIX"
      Case 7
         ShowNumber = "SEVEN"
      Case 8
         ShowNumber = "EIGHT"
      Case 9
         ShowNumber = "NINE"
      Case 10
         ShowNumber = "TEN"
      Case 11
         ShowNumber = "ELEVEN"
      Case 12
         ShowNumber = "TWELVE"
      Case 13
         ShowNumber = "THIRTEEN"
      Case 14
         ShowNumber = "FOURTEEN"
      Case 15
         ShowNumber = "FIFTEEN"
      Case 16
         ShowNumber = "SIXTEEN"
      Case 17
         ShowNumber = "SEVENTEEN"
      Case 18
         ShowNumber = "EIGHTEEN"
      Case 19
         ShowNumber = "NINTEEN"
      Case 20
         ShowNumber = "TWENTY"
      Case 30
         ShowNumber = "THIRTY"
      Case 40
         ShowNumber = "FORTY"
      Case 50
         ShowNumber = "FIFTY"
      Case 60
         ShowNumber = "SIXTY"
      Case 70
         ShowNumber = "SEVENTY"
      Case 80
         ShowNumber = "EIGHTY"
      Case 90
         ShowNumber = "NINETY"
      Case 99
         ShowNumber = "CENTS"
      Case 100
         ShowNumber = "HUNDRED"
      Case 101
         ShowNumber = "THOUSAND"
      Case 102
         ShowNumber = "MILLION"
      Case 103
         ShowNumber = "BILLION"
      Case 104
         ShowNumber = "TRILLION"
      Case 105
         ShowNumber = "AND"
      Case 106
         ShowNumber = "POINT"
      Case 107
         ShowNumber = "ONLY."
      Case 108
         ShowNumber = "DOLLARS"
   End Select
End Function

