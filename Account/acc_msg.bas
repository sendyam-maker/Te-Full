Attribute VB_Name = "aacc_msg"
'Memo by Morgan 2022/8/17 ¤é¤å¤w§ï§ìTable
'Memo by Morgan2010/8/19 ¤é´ÁÄæ¤w­×§ï
Option Explicit


'*************************************************
'  ¦@¥Î Combo ¤º½X¹ï·Ó¨ç¦¡
'
'*************************************************
Public Function ComboItem(InputIndex As Integer) As String
   Select Case InputIndex
      Case 1
         ComboItem = "1--­É¤è"
      Case 2
         ComboItem = "2--¶U¤è"
      Case 11
         ComboItem = "1--¤ä²¼"
      Case 12
         ComboItem = "2--¥»²¼"
      Case 13
         ComboItem = "3--¶×²¼"
      Case 21
         ComboItem = "1--¦b³~¤¤"
      Case 22
         ComboItem = "2--¤w§I²{"
      Case 31
         ComboItem = "1--Ä~Äò°lÂÜ"
      Case 32
         ComboItem = "2--¤£ºÞ¨î"
      Case 41
         ComboItem = "1--ÂàÀ³¥I´Ú"
      Case 42
         ComboItem = "2--Âà¼È¦¬´Ú"
      Case 43
         ComboItem = "3--¥»©ÒÃº¯Ç"
      Case 51
         ComboItem = "1--IR"
      Case 52
         ComboItem = "2--CB"
      Case 53
         ComboItem = "3--°U¦¬"
      Case 54
         ComboItem = "4--¨ä¥L»È¦æ"
      Case 55
         ComboItem = "5--¨ä¥L"
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
         ComboItem = "1--²¼¶×"
      Case 72
         'Modified by Lydia 2017/10/11 ¹q¶×§ó¦W"¾ã§å¹q¶×"
         ComboItem = "2--¾ã§å¹q¶×"
      'Added by Lydia 2015/04/17 ¶×´Ú¤è¦¡¼W¥[3
      Case 77
         ComboItem = "3--¥x»È¹q¶×¯È¥»"
      'Added by Lydia 2017/09/06 ¶×´Ú¤è¦¡¼W¥[4
      Case 78
         ComboItem = "4--µØ»È¹q¶×¯È¥»"
      'Added by Lydia 2017/09/22 ¶×´Ú¤è¦¡¼W¥[5
      Case 79
         'Modified by Lydia 2017/10/11 ¦X­pµ²¶×§ó¦W"¥x»È¦X¨Öµ²¶×"
         ComboItem = "5--¥x»È¦X¨Öµ²¶×"
      'Added by Lydia 2024/09/03 ¶×´Ú¤è¦¡¼W¥[6
      Case 80
         ComboItem = "6--©è±b"
      'end 2024/09/03
      Case 73
         ComboItem = "3--®È¦æ¤ä²¼"
      Case 74
         ComboItem = "4--²{ª÷"
      Case 75
         ComboItem = "5--°Ó°È¥d"
      Case 76
         ComboItem = "6--¨ä¥L"
      Case 81
         ComboItem = "1--Ã±¦¬"
      Case 82
         ComboItem = "2--±H¥X"
      Case 83
         ComboItem = "3--±H¤À©Ò"
      Case 84
         ComboItem = "4--¨ä¥¦"
      Case 85
         ComboItem = "5--±H¥X¤£¦L¦W±ø"
      Case 91
         ComboItem = "1--¼t°Ó"
      Case 92
         ComboItem = "2--«È¤á"
      Case 93
         ComboItem = "3--­û¤u"
      Case 101
         ComboItem = "1--À³¥I´Ú"
      Case 102
         ComboItem = "2--¾P°hÂà¤J"
      Case 103
         ComboItem = "3--¸É¦©Ãº°h¶O"
      Case 111
         ComboItem = "1--®Ö»é°h¶O"
      Case 112
         ComboItem = "2--·¸¦¬´Ú"
      Case 113
         ComboItem = "3--®×¥ó¥¼¿ì°h¶O"
      Case 114
         ComboItem = "4--¦©Ãº"
      Case 115
         ComboItem = "5--³f´Ú"
      Case 116
         ComboItem = "6--µ|´ÚÃº´Ú®Ñ"
      Case 117
         ComboItem = "7--¨ä¥L"
      Case 121
         ComboItem = "¤¤"
      Case 122
         ComboItem = "­^"
      Case 123
         ComboItem = "¤é"
      Case 131
         ComboItem = "1--«È¤á"
      Case 132
         ComboItem = "2--¼t°Ó"
      Case 133
         ComboItem = "3--­û¤u"
      Case 134
         ComboItem = "4--¨ä¥L"
      Case 141
         ComboItem = "1--¥¼¨R¥­"
      Case 142
         ComboItem = "2--¨R¥­"
      Case 143
         ComboItem = "3--¥þ³¡"
      Case 151
         ComboItem = "1--¤W¥b¦~"
      Case 152
         ComboItem = "2--¤U¥b¦~"
      Case 161
         ComboItem = "1--°Ó¼Ð"
      Case 162
         ComboItem = "2--±M§Q"
      Case 163
         ComboItem = "3--®Û«ß®v"
      Case 164
         ComboItem = "5--¸â«ß®v"
      Case 165
         ComboItem = "7--½±«ß®v"
      Case 166
         ComboItem = "8--­ð«ß®v"
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
         ComboItem = "R--À³¦¬"
      Case 182
         ComboItem = "P--À³¥I"
      Case 191
         ComboItem = "1--¥x¤¤"
      Case 192
         ComboItem = "2--¥x«n"
      Case 193
         ComboItem = "3--°ª¶¯"
      Case 201
         ComboItem = "1--¼È¦¬´Ú"
      Case 202
         ComboItem = "2--·¸¦¬Âà¤J"
      Case 203
         ComboItem = "3--°h¶OÂà¤J"
      Case 211
         ComboItem = "1-ÂàÀ³¥I´Ú"
      Case 212
         ComboItem = "2-Âà¼È¦¬´Ú"
      Case 221
         ComboItem = "1--¨R¼È¦¬´Ú"
      Case 222
         ComboItem = "2--¨ä¥¦"
      Case 231
         ComboItem = "1--¥»©Ò®×¸¹"
      Case 232
         ComboItem = "2--«È¤á/¼t°Ó/­û¤u"
      Case 233
         ComboItem = "3--´¼Åv¤H­û"
      Case 234
         ComboItem = "4--¨ä¥¦"
      Case 241
         ComboItem = "µ²¾l©ú²Ó"
      Case 242
         ComboItem = "ÂI¼Æ©ú²Ó"
      Case 251
         ComboItem = "µ²¶×©ú²Óªí"
      Case 252
         ComboItem = "¶×²¼¨ç¥ó©ú²Óªí"
      Case 253
         ComboItem = "¥I´Ú©ú²Óªí"
      Case 254
        ComboItem = "1--¨Æ°È©Ò"
      Case 255
        ComboItem = "J--´¼Åv¤½¥q"
   End Select
End Function

'*************************************************
'  ³øªí¤º½X¹ï·Ó¨ç¦¡
'
'*************************************************
Public Function ReportTitle(InputIndex As Integer) As String
   Select Case InputIndex
      Case 101
         ReportTitle = "***  ¦¬¤å»P¦¬¾Ú¸ê®ÆÀË®Öªí  ***"
      Case 103
         ReportTitle = "***  ¥I´Ú¤u§@©³½Z  ***"
      Case 1041
         ReportTitle = "***  «È¤á©¹¨Ó¹ï±b³æ  ***"
      Case 1042
         ReportTitle = "***  «È¤áÀ³¦¬±b´Ú¹ï±b³æ  ***"
      Case 105
         ReportTitle = "***  «È¤á±b´Ú©ú²Óªí  ***"
      Case 106
         ReportTitle = "***  ´¼Åv¤H­û±b´Ú©ú²Óªí  ***"
      Case 107
         ReportTitle = "***  ´¼Åv¤H­ûÀ³¦¬³W¶O©ú²Óªí  ***"
      Case 1081
         ReportTitle = "***  «È¤á±bÄÖ¤ÀªRªí  ***"
      Case 1082
         ReportTitle = "***  ´¼Åv¤H­û±bÄÖ¤ÀªRªí  ***"
      Case 1083
         ReportTitle = "***  ´¼Åv¤H­û«È¤á±bÄÖ¤ÀªRªí  ***"
      Case 109
         ReportTitle = "***  ¾P±b°h¶O©ú²Óªí  ***"
      Case 110
         ReportTitle = "***  ¼È¦¬´Ú©ú²Óªí  ***"
      Case 111
         ReportTitle = "***  ¤ä²¼±H¥X©ú²Ó²M³æ  ***"
      Case 1111
         ReportTitle = "***  ¥I´Ú³qª¾³æ  ***"
      Case 1112
         ReportTitle = "***  ¥I´ÚÃ±¦¬Ã¯  ***"
      Case 1113
         ReportTitle = "***  ¦a§}±ø  ***"
      Case 1114
         ReportTitle = "***  ²¼¾Ú¨ü»â¦¬¾Ú  ***"
      Case 112
         ReportTitle = "***  À³¥I´Ú²Î­pªí  ***"
      Case 113
         ReportTitle = "***  °ê¤º¥I´Ú©ú²Óªí  ***"
      Case 114
         ReportTitle = "***  ´¼Åv¤H­û§O«È¤á¦©Ãºµ|´Ú©ú²Óªí  ***"
      Case 1141
         ReportTitle = "***  ¦©Ãº¶Ê¦¬©ú²Ó¶i«×ªí  ***"
      Case 115
         ReportTitle = "***  ¦^¡@°õ¡@³æ  ***"
      Case 1151
         ReportTitle = "***  Ãº´Ú®Ñ±H¥X©ú²Ó  ***"
      Case 116
         ReportTitle = "***  ¦¬¾Ú§@¼o©ú²Óªí  ***"
      Case 204
         ReportTitle = "***  µ²¶×©ú²Óªí  ***"
      Case 209
         ReportTitle = "***  ¥N²z¤H¹ï±b³æ  ***"
      Case 210
         ReportTitle = "***  °ê¥~±bÄÖ¤ÀªRªí  ***"
      Case 211
         ReportTitle = "***  ¥N²z¤H±b¥Ø±Æ¦W  ***"
      Case 2121
         ReportTitle = "***  ´¼Åv¤H­û·~ÁZ½Ð´ÚÂI¼Æ²Î­pªí  ***"
      Case 2122
         ReportTitle = "***  ´¼Åv¤H­û·~ÁZ¦¬´ÚÂI¼Æ²Î­pªí  ***"
      Case 213
         ReportTitle = "***  ¥N²z¤HFC±b´Ú©ú²Óªí  ***"
      'Add By Cheng 2002/09/02
      Case 2131
         ReportTitle = "***  °ê¥~FC±b´Ú©ú²Óªí  ***"
      Case 214
         ReportTitle = "***  °ê¥~À³¦¬³W¶O¤ÎªA°È¶O¤ÀªRªí  ***"
      Case 215
         ReportTitle = "***  ¥N²z¤H¹O´Á±b´Ú¤ÀªRªí  ***"
      Case 216
         ReportTitle = "***  °ê¤º¥¼¦¬´Ú©ú²Óªí  ***"
      Case 217
         ReportTitle = "***  ¥N²z¤H¥¼¦¬¥¼¥I¹ï·Óªí  ***"
      Case 218
         ReportTitle = "***  ¥I´Ú©ú²Ó¯ó½Z  ***"
      Case 219
         ReportTitle = "***  ¶×²¼¨ç¥ó©ú²Óªí  ***"
      Case 225
         ReportTitle = "¥x»È¥~¹ô§I²{©ú²Óªí"
      Case 301
         ReportTitle = "***  À³¦¬²¼¾Ú¸ê®Æªí  ***"
      Case 302
         ReportTitle = "***  À³¥I²¼¾Ú¸ê®Æªí  ***"
      Case 303
         ReportTitle = "***  °U¦¬²¼¾Ú¸ê®Æªí  ***"
      Case 304
         ReportTitle = "***  »È¦æ±b¸¹§O²¼¾Ú·JÁ`ªí  ***"
      Case 305
         ReportTitle = "***  »È¦æ±b¸¹§O²¼¾Ú©ú²Óªí  ***"
      Case 306
         ReportTitle = "***  §I²{¤é§O¸êª÷¬y°Ê·JÁ`ªí  ***"
      Case 307
         ReportTitle = "***  §I²{¤é§O²¼¾Ú©ú²Óªí  ***"
      Case 308
         ReportTitle = "***  ©¹¨Ó¹ï¶H§O²¼¾Ú·JÁ`ªí  ***"
      Case 309
         ReportTitle = "***  ©¹¨Ó¹ï¶H§O²¼¾Ú©ú²Óªí  ***"
      Case 310
         ReportTitle = "***  °h²¼¸ê®Æªí  ***"
      Case 311
         ReportTitle = "***  ©â²¼¸ê®Æªí  ***"
      Case 312
         ReportTitle = "***  ²¼¾Ú¶K²{¸ê®ÆÀË®Öªí  ***"
      Case 313
         ReportTitle = "***  »È¦æ±b¸¹§O¸êª÷¬y°Êªí  ***"
      Case 314
         ReportTitle = "***  ¤é´Á§O¸êª÷¬y°Ê¹w´úªí  ***"
      Case 315
         ReportTitle = "***  »È¦æ½Õ¸`¸ê®Æªí  ***"
      Case 316
         ReportTitle = "***  »È¦æ§O¸ê®Æªí  ***"
      Case 317
         ReportTitle = "***  ¥Ò¦s¤ä²¼¥¼§I»â©ú²Óªí  ***"
      Case 401
         ReportTitle = "***  ¤é­pªí  ***"
      Case 402
         ReportTitle = "***  ·|­p¬ì¥Ø¥N¸¹¹ï·Óªí  ***"
      Case 403
         ReportTitle = "***  ¬ì¥Ø¾lÃBªí  ***"
      Case 404
         ReportTitle = "***  ¬ì¥Ø©ú²Óªí(¹ï¨R)  ***"
      Case 405
         ReportTitle = "***  ¸Õºâªí  ***"
      Case 406
         ReportTitle = "***  ¬ì¥Ø¤ÀÃþ±b  ***"
      Case 407
         ReportTitle = "***  ºî¦X·l¯qªí  ***"
      Case 408
         ReportTitle = "***  ºî¦X·l¯q¤ñ¸ûªí  ***"
      Case 409
         ReportTitle = "***  ¸ê²£­t¶Åªí  ***"
      Case 410
         ReportTitle = "***  ¹wºâ¹êÁZ¤ñ¸ûªí  ***"
      Case 411
         ReportTitle = "***  ³¡ªù¶O¥Î²Î­pªí  ***"
      Case 412
         ReportTitle = "***  ³¡ªùºî¦X·l¯qªí(¤l¬ì¥Ø)  ***"
      Case 413
         ReportTitle = "***  ¦~«×ºî¦X·l¯q²Î­pªí  ***"
      Case 414
         ReportTitle = "***  ¦~«×³¡ªùºî¦X·l¯q²Î­pªí  ***"
      Case 415
         ReportTitle = "***  ¸ê²£­t¶Å¤ñ¸ûªí  ***"
      Case 416
         ReportTitle = "***  ³¡ªùºî¦X·l¯qªí  ***"
      Case 417
         ReportTitle = "***  ´¼Åv¤H­ûÂI¼Æ©ú²Óªí  ***"
      Case 4171
         ReportTitle = "***  ´¼Åv¤H­ûÂI¼Æ¤ÀªRªí  ***"
      Case 418
         ReportTitle = "***  ¹wºâ¸ê®Æªí  ***"
      Case 419
         ReportTitle = "***  ¶O¥Î¬ì¥Ø¤ÀÅu¤ñ²vªí  ***"
      Case 420
         ReportTitle = "¦©Ãº¾Ì³æ®Ö¹ïªí"
      Case 421
         ReportTitle = "***  ¦©Ãº¾Ì³æ©ú²Óªí  ***"
      Case 422
         ReportTitle = "***  ¦~«×¦©Ãº©ú²Ó®Ö¹ïªí  ***" 'Modify By Sindy 2021/12/3 ***  «È¤á¦©Ãº©ú²Ó®Ö¹ïªí  ***
      Case 423
         ReportTitle = "***  ´¼Åv¤H­ûµ²¾lÂI¼ÆÁ`ªí  ***"
      Case 424
         ReportTitle = "¤ë¥÷±M·~ÂI¼Æ©ú²Óªí"

   End Select
End Function

'*************************************************
'  ¹q¸Ü°Ï°ì¤º½X¹ï·Ó¨ç¦¡
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
'­×§ï¶Ç¤J°Ñ¼Æ«¬ºA
'Public Function ReportSum(InputIndex As Integer) As String
Public Function ReportSum(InputIndex As Double) As String
   Select Case InputIndex
      Case 1
         ReportSum = "Àç·~¦¬¤J:"
      Case 2
         ReportSum = "Àç·~¤ä¥X:"
      Case 3
         ReportSum = "Àç·~·l¯q:"
      Case 4
         ReportSum = "¡Ð¡Ð¡Ð¡Ð¡Ð¡Ð"
      Case 5
         ReportSum = "Àç·~¥~¦¬¤J:"
      Case 6
         ReportSum = "Àç·~¥~¤ä¥X:"
      Case 7
         ReportSum = "µ|«e²b·l¯q:"
      Case 8
         ReportSum = "¡×¡×¡×¡×¡×¡×"
      Case 9
         ReportSum = "¸ê²£Á`ÃB:"
      'Add By Cheng 2002/01/18
      Case 9001
         ReportSum = "*** ¸ê²£Á`ÃB ***"
      Case 10
         ReportSum = "­t¶Å¤p­p:"
      'Add By Cheng 2002/01/18
      Case 10001
         ReportSum = "*** ­t¶Å¤p­p ***"
      Case 11
         ReportSum = "¥»´Á·l¯q:"
      Case 12
         ReportSum = "ªÑªFÅv¯q¤p­p:"
      'Add By Cheng 2002/01/18
      Case 12001
         ReportSum = "*** ªÑªFÅv¯q¤p­p ****"
      Case 13
         ReportSum = "­t¶ÅÁ`ÃB:"
      'Add By Cheng 2002/01/18
      Case 13001
         ReportSum = "*** ­t¶Å»PªÑªFÅv¯qÁ`ÃB ***"
      Case 14
         ReportSum = "¹ê»ÚÀç·~¦¬¤J:"
      Case 15
         ReportSum = "¶O¥Î¦X­p:"
      Case 16
         ReportSum = "³¡ªù·l¯q:"
      Case 17
         ReportSum = "¤ÀÅu¶O¥Î:"
      Case 18
         ReportSum = "¦U³¡ªùÀç·~·l¯q:"
      Case 19
         ReportSum = "Àç·~¥~¦¬¤ä:"
      Case 20
         ReportSum = "¥þ©Ò·l¯q:"
      Case 21
         ReportSum = "³¡ªù¸gÀç·l¯q:"
      Case 22
         ReportSum = "¸ê²£¦X­p:"
      'Add By Cheng 2002/01/18
      Case 22001
         ReportSum = "*** ¸ê²£¦X­p ***"
      Case 23
         ReportSum = "­t¶Å»PªÑªFÅv¯q¦X­p:"
      'Add By Cheng 2002/01/18
      Case 23001
         ReportSum = "*** ­t¶Å»PªÑªFÅv¯q¦X­p ***"
      Case 24
         ReportSum = "¤p­p:"
      Case 25
         ReportSum = "¦X­p:"
      Case 26
         ReportSum = "µ§"
      Case 27
         ReportSum = "²Î­p¤é´Á:"
      Case 28
         ReportSum = "¦Ü"
      Case 29
         ReportSum = "¹ï¨R¥N¸¹(·~)"
      Case 30
         ReportSum = "¶Ç²¼½s¸¹"
      Case 31
         ReportSum = "¹ï¨R¥N¸¹(«È)"
      Case 32
         ReportSum = "¹ï¨R¥N¸¹(¥»)"
      Case 33
         ReportSum = "ºK­n"
      Case 34
         ReportSum = "ª÷ÃB"
      Case 35
         ReportSum = "»sªí¤é´Á: "
      Case 36
         ReportSum = "­¶¡@¡@¦¸: "
      Case 37
         ReportSum = "¥I´Ú¦æ®w: "
      Case 38
         ReportSum = "¥I´Ú±b¸¹: "
      Case 39
         ReportSum = "¤ä²¼¸¹½X: "
      Case 40
         ReportSum = "¨ì ´Á ¤é:   "
      Case 41
         ReportSum = "ª÷¡@¡@ÃB: "
      Case 42
         ReportSum = "³Æ¡@¡@µù: "
      Case 43
         ReportSum = "     ¥x Å³:"
      Case 44
         ReportSum = "¯÷ ±H ¤W À³ ¥I    ¥x ºÝ  ( ¶Q ¤½ ¥q )  ¤§ ²¼ ¾Ú  ( ¸Ô ­z ¦p ¤U )  ¡A ¨Ã ±N ²¼ ¾Ú ¨ü »â ¦¬ ¾Ú"
      Case 45
         ReportSum = "¶ñ §´ ±H ¦^ ¡A ÁÂ ÁÂ ±z ªº ¤ä «ù »P ¦X §@ ¡C"
      Case 46
         ReportSum = "¯S §O »¡ ©ú : "
      Case 47
         ReportSum = "( ¤@ ) ·q ½Ð ©ó ²¼ ¾Ú ¨ü »â ¦¬ ¾Ú ¤W Ã± »\    ¥x ºÝ  ( ¶Q ¤½ ¥q )  ¤§ ¦¬ ´Ú ³¹ ¡I ¡I"
      Case 48
         ReportSum = "( ¤G ) ²¼ ¾Ú ¨ü »â ¦¬ ¾Ú ­Y ¥¼ ±H ¦^ ªÌ ¡A ¥H «á    ¥x ºÝ  ( ¶Q ¤½ ¥q )  ¤§ ´Ú ¶µ ¡A"
      Case 49
         ReportSum = "¡@¡@  ®¤ ¤£ ¦A ¶l ±H °e ¹F ¡I ¡I"
      Case 50
         ReportSum = "½Ð ªu ¦¹ µê ½u ¼¹ ¤U ±H ¦^"
      Case 51
         ReportSum = "¯÷ ¦¬ ¨ì    ¶Q ¨Æ °È ©Ò ±H ¨Ó ¤§ ²¼ ¾Ú  ( ¸Ô ­z ¦p ¤U )  ¡A ¤@ ¤Á µL »~ ¡A ¯S ¦¹ ÃÒ ©ú ¡C"
      Case 52
         ReportSum = "¯÷ ±H ¤W    ¶Q ¤½ ¥q ¤§ ¦U Ãþ ©Ò ±o µ| ¦© Ãº µ| ´Ú Ãº ´Ú ®Ñ ¡@¡@¡@¡@ ¥÷ ¡A ª÷ ÃB ¦@ ­p"
      Case 53
         ReportSum = "¤¸ ¾ã ¡A½Ð ¬d ¦¬ ¡A ¨Ã ¦^ °õ ³æ ¤W »\ ³¹ «á ±H ¦^ ¥» ¨Æ °È ©Ò ¡A ÁÂ ÁÂ ±z ªº ¦X §@ ¡C"
      Case 54
         ReportSum = "¯÷ ¦¬ ¨ì    ¶Q ¨Æ °È ©Ò ±H ¨Ó ¤§ ¦U Ãþ ©Ò ±o µ| ¦© Ãº µ| ´Ú Ãº ´Ú ®Ñ ¦@ ¡@¡@¡@¡@ ¥÷ ¡A ª÷ ÃB ¦@ ­p"
      Case 55
         ReportSum = "¤¸ ¾ã ¡A ¤@ ¤Á µL »~ ¡C ¯S ¦¹ ÃÒ ©ú ¡C"
      Case 56
         ReportSum = "Ã±¡@¡@¦¬¡@¡@¤H¡@¡G"
      Case 57
         ReportSum = "´¼Åv¤H­û"
      Case 58
         ReportSum = "·~°È¹F¦¨ÂI¼Æ"
      Case 59
         ReportSum = "¥[Âà¼·ÂI¼Æ"
      Case 60
         ReportSum = "´îÂà¼·ÂI¼Æ"
      Case 61
         ReportSum = "«O¯dÂI¼Æ"
      Case 62
         ReportSum = "¹ê»Ú¹F¦¨ÂI¼Æ"
      Case 63
         ReportSum = "¥x¥_©Ò"
      Case 64
         ReportSum = "¨ä¥¦"
      Case 65
         ReportSum = "°ê¤º"
      Case 66
         ReportSum = "¥þ©Ò"
      Case 67
         ReportSum = "FCP"
      Case 68
         ReportSum = "FCT"
      Case 69
         ReportSum = "FCL"
      Case 70
         ReportSum = "°ê¥~"
      Case 71
         ReportSum = "Name of Bank: Bank of Taiwan, Head Office Foreign Department"
        'Add By Cheng 2003/02/13
      Case 71001
         'modify by sonia 2018/12/6 ÔÑÞ±³qª¾­×§ï
         'ReportSum = "Name of Bank: Bank of Taiwan, Head Office"
         ReportSum = "Name of Bank: Bank of Taiwan DEPT. OF BUSINESS"
      Case 71002
         'modify by sonia 2018/12/6 ÔÑÞ±³qª¾­×§ï
         'ReportSum = "¨ú¤Þ»È¦æ¡G Bank of Taiwan, Head Office"
         ReportSum = "¨ú¤Þ»È¦æ¡G Bank of Taiwan DEPT. OF BUSINESS"
      Case 72
         'Modified by Morgan 2013/12/17--ÔÑÞ±
         'ReportSum = "Address: 120, Sec. 1, Chungking S. Rd., Taipei, Taiwan, R.O.C."
         ReportSum = "Address: No.120, Sec.1 Chong-Qing S. Rd., Taipei, Taiwan, R.O.C."
      'Add by Morgan 2013/12/18
      Case 72001
         ReportSum = "»È¦æ¦í©Ò¡GNo.120, Sec.1 Chong-Qing S. Rd., Taipei, Taiwan, R.O.C."
      Case 73
         'ReportSum = "S.W.I.F.T. Address: BKTW TWTP"
         ReportSum = "SWIFT Address: BKTW TWTP"
        'Add By Cheng 2003/02/13
      Case 73001
         'Modified by Morgan 2013/12/17--ÔÑÞ±
         'ReportSum = "S.W.I.F.T. Code: BKTW TWTP"
         ReportSum = "SWIFT Code: BKTWTWTP"
      Case 74 '¬üª÷±b¤á
        'Modify By Cheng 2003/07/25
'         ReportSum = "Account No.: 006007052643 (for US currency)"
         'Modified by Morgan 2013/12/17--ÔÑÞ±
         'ReportSum = "Account No.: 003007052646 (for US currency)"
         ReportSum = "Account No.: 003007052646 (Multi-Currency Account)"
      Case 74001 '¬üª÷±b¤á
         'Modified by Morgan 2013/12/18
         'ReportSum = "¤f®yµf†A¡G 003007052646 (USÇÅÇçÇRÇoÇrÆú°eª÷ÇU³õ¦X)"
         'Modified by Morgan 2022/10/26
         'ReportSum = "¤f®yµf†A¡G 003007052646 (USÇÅÇçÇeþòÇV’AÇRÇoÇrþç°eª÷ÇU³õ¦X)"
         ReportSum = PUB_GetUniText("ReportSum", "74001")
      Case 75
         ReportSum = "Currency Rate: USD1.00=NTD"
      Case 75001
         'Modified by Morgan 2022/10/26
         'ReportSum = "²{¦bÇU¢ã¢áÇÅÇçÇR’cÇ@ÇrNTÇÅÇçÇUÇè¡ÐÇÄÇVUSD1.00=NTD"
         ReportSum = PUB_GetUniText("ReportSum", "75001")
      Case 76
         ReportSum = "¦©Ãº¦~«×:"
      Case 77
         ReportSum = "¤w¦©ª÷ÃB"
      Case 78
         ReportSum = "¤w¦¬¦©³æ"
      Case 79
         ReportSum = "¤w¦¬²{ª÷"
      Case 80
         ReportSum = "¦C§b±b"
      Case 81
         ReportSum = "¶Ê¦¬¤¤"
      Case 82
         ReportSum = "Âà¦C¤U¦~«×"
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
         'Modified by Morgan 2013/12/20
         'ReportSum = "Account Name: Tai E International Patent & Law Office"
         ReportSum = "Account Name: Tai E International Patent and Law Office"
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
         ReportSum = "be appreciated if you could mention our reference number while sending us your"
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
      'Add by Morgan 2008/4/17
      Case 991
         'Modify by Morgan 2008/4/24
         'ReportSum = "If your account information have any change, please inform us to renew it."
         'Modify by Morgan 2008/4/29
         'ReportSum = "If your account information has been changed, please inform us accordingly and we shall update our records."
         ReportSum = "If your account information has been changed, please inform us accordingly"
      'Add by Morgan 2008/4/29
      Case 99101
         ReportSum = "and we shall update our records."
      Case 100
         ReportSum = "We acknowledge with thanks receipt of your payment as identified below:"
      Case 101
         ReportSum = "¦X­p"
      Case 102
         ReportSum = "¦a§}: "
      Case 103
         ReportSum = "¹q¸Ü: "
      Case 104
         ReportSum = "¥x¥_¥«¤¤¤s°Ïªø¦wªF¸ô¤G¬q112¸¹9¼Ó"
      Case 105
         ReportSum = "¥x¥_©Ò¦X­p:"
      Case 106
         ReportSum = "¥x¤¤©Ò¦X­p:"
      Case 107
         ReportSum = "¥x«n©Ò¦X­p:"
      Case 108
         ReportSum = "°ª¶¯©Ò¦X­p:"
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
      'Added by Morgan 2013/12/20
      Case 11601
         ReportSum = "¤f®y¦W¸q¡G Tai E International Patent and Law Office"
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
      Case 12101
         'Modified by Morgan 2013/12/20
         'ReportSum = "¤f®yµf†A¡G 003001305688 (¥x“øÇÅÇçÇRÇoÇrÆú°eª÷ÇU³õ¦X)"
         'Modified by Morgan 2022/10/26
         'ReportSum = "¤f®yµf†A¡G 003001305688 (¥x“øÇÅÇçÇRÇoÇrþç°eª÷ÇU³õ¦X)"
         ReportSum = PUB_GetUniText("ReportSum", "12101")
      Case 122
         ReportSum = "»È¦æ: ¤¤°ê¤u°Ó»È¦æ¤W®ü®}¶×¤ä¦æ  ¤ÑÜv¾ô¸ôÀx»W©Ò"
      Case 123
         ReportSum = "½ã¤á¦WºÙ: ¨L®a¿« (¤H¥Á¹ô­Ó¤H½ã¤á)"
      Case 124
         ReportSum = "½ã¸¹: 47271010301*0"
      Case 125
         ReportSum = "¡° ¶Q©Ò¥i±N´Ú¶µ¶×¦Ü¥»©Ò¤W®ü©Î¥xÆW¤§»È¦æ½ã¤á¡A±©©ó¶×´Ú«á½Ð"
      Case 126
         ReportSum = "     °È¥²ª¾¶×¥x¥_Á`©Ò¡A¨Ã§iª¾¶×´Úª÷ÃB¡C"
      Case 127
         ReportSum = "¼sªF©Ò¦X­p:"
        'Add By Cheng 2003/02/07
        '¥i¥H¼Ú¤¸¤ä¥I
      Case 128
         ReportSum = "Payment by EURO is acceptable"
        'Add By Cheng 2003/02/13
      Case 129 '¼Ú¤¸±b¤á
        'Modify By Cheng 2003/07/25
'         ReportSum = "Account No.: 006007085124 (for EURO currency)"
         ReportSum = "Account No.: 003007085127 (for EURO currency)"
        'Add By Cheng 2004/05/07
        Case "130"
            'Modified by Morgan 2022/10/26
            'ReportSum = "¤U°OÇU³qÇq½Ð¨D¥ÓÆý¤WÆøÇeÇ@¡C"
            ReportSum = PUB_GetUniText("ReportSum", "130")
   End Select
End Function

'*************************************************
'  µ{¦¡°T®§¤º½X¹ï·Ó¨ç¦¡(¤¤¤å)
'
'*************************************************
Public Function ShowNumberWord(InputNumber As Long) As String
   Select Case InputNumber
      Case 0
         ShowNumberWord = "¹s"
      Case 1
         ShowNumberWord = "³ü"
      Case 2
         ShowNumberWord = "¶L"
      Case 3
         ShowNumberWord = "°Ñ"
      Case 4
         ShowNumberWord = "¸v"
      Case 5
         ShowNumberWord = "¥î"
      Case 6
         ShowNumberWord = "³°"
      Case 7
         ShowNumberWord = "¬m"
      Case 8
         ShowNumberWord = "®Ã"
      Case 9
         ShowNumberWord = "¨h"
      Case 10
         ShowNumberWord = "¬B"
      Case 11
         ShowNumberWord = "¨Õ"
      Case 12
         ShowNumberWord = "¥a"
      Case 13
         ShowNumberWord = "¸U"
      Case 14
         ShowNumberWord = "»õ"
      Case 20
         ShowNumberWord = "¤¸¾ã"
   End Select
End Function

'*************************************************
'  µ{¦¡°T®§¤º½X¹ï·Ó¨ç¦¡(­^¤å)
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

