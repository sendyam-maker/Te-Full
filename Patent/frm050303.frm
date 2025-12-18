VERSION 5.00
Begin VB.Form frm050303 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人案件收達/提申管制表"
   ClientHeight    =   5280
   ClientLeft      =   1620
   ClientTop       =   1500
   ClientWidth     =   4488
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4488
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   1290
      MaxLength       =   6
      TabIndex        =   12
      Top             =   2958
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   7
      Top             =   2040
      Width           =   800
   End
   Begin VB.ListBox List1 
      Height          =   768
      Index           =   1
      ItemData        =   "frm050303.frx":0000
      Left            =   1440
      List            =   "frm050303.frx":0002
      TabIndex        =   32
      Top             =   4380
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "刪除"
      Height          =   400
      Index           =   1
      Left            =   3420
      TabIndex        =   17
      Top             =   4470
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "新增"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3420
      TabIndex        =   16
      Top             =   4050
      Width           =   600
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   1
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   31
      Text            =   "P"
      Top             =   4065
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   2
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   13
      Top             =   4065
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   3
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   14
      Top             =   4065
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   4
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   15
      Top             =   4065
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   6
      Top             =   2040
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3420
      TabIndex        =   19
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   2610
      TabIndex        =   18
      Top             =   15
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1290
      MaxLength       =   1
      TabIndex        =   1
      Top             =   810
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   2895
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2640
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   1290
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2640
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2580
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2340
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1290
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2340
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1740
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1290
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1740
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2580
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1440
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1440
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1290
      TabIndex        =   0
      Top             =   465
      Width           =   2100
   End
   Begin VB.Label Label14 
      Height          =   180
      Left            =   2160
      TabIndex        =   35
      Top             =   3000
      Width           =   1365
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "管制人："
      Height          =   180
      Left            =   480
      TabIndex        =   34
      Top             =   3000
      Width           =   720
   End
   Begin VB.Line Line5 
      X1              =   2220
      X2              =   2460
      Y1              =   2165
      Y2              =   2165
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "不管制案號："
      Height          =   180
      Left            =   300
      TabIndex        =   33
      Top             =   4080
      Width           =   1080
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "     2.管制別3,4,5僅內專適用!!"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   315
      TabIndex        =   30
      Top             =   3810
      Width           =   2280
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "管制日："
      Height          =   180
      Left            =   480
      TabIndex        =   29
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "                 b.案件性質為領證,超項費,其他不管制"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   300
      TabIndex        =   28
      Top             =   3570
      Width           =   3690
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "PS:1.內商 a.期限半年內未到期不出現"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   300
      TabIndex        =   27
      Top             =   3330
      Width           =   2865
   End
   Begin VB.Label Label7 
      Caption         =   "( 1. 未收達　　 2. 未提申　　      3. 收達管制函 4. 提申管制函      5.公開管制函)"
      Height          =   630
      Left            =   1620
      TabIndex        =   26
      Top             =   810
      Width           =   2460
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "管制別："
      Height          =   180
      Left            =   480
      TabIndex        =   25
      Top             =   810
      Width           =   720
   End
   Begin VB.Line Line4 
      X1              =   2580
      X2              =   2820
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   480
      TabIndex        =   24
      Top             =   2640
      Width           =   720
   End
   Begin VB.Line Line3 
      X1              =   2220
      X2              =   2460
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   300
      TabIndex        =   23
      Top             =   2340
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   1860
      X2              =   2100
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   300
      TabIndex        =   22
      Top             =   1740
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2220
      X2              =   2460
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Left            =   480
      TabIndex        =   21
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   300
      TabIndex        =   20
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "frm050303"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo by Morgan2010/12/28 申請案號欄已修改
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'整理 by Morgan 2005/7/11
Option Explicit

Dim strSql As String, strTemp1 As Variant, strTemp2 As Variant, StrTest1 As String, StrTest2 As String, i As Integer, j As Integer, s As Integer
Dim PLeft(0 To 10) As Integer, k As Integer, TmpArea As String, iLine As Integer, Page As Integer
Dim strTemp3(0 To 10) As String, iPrint As Integer
Dim StrTest3 As String
Dim m_strET02 As String, m_strCaseName As String, m_strCaseNo1 As String, m_strCaseNo2 As String, m_strCaseNo3 As String, m_bolShowNo3 As Boolean
Dim mOpenSql As String 'Added by Lydia 2015/10/26 因為basLetter簽名檔判斷會造成原表單的FMP2openSQL變空白,先保留
Dim m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String, m_strCP09 As String, m_strCP10 As String, m_Subject As String 'Added by Morgan 2016/5/13
Dim mSeqNo As String 'Added by Lydia 2022/09/20
Dim stVTBX As String 'Added by Morgan 2025/1/21

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If Len(txt1(0)) = 0 Then
            s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
            txt1(0).SetFocus
            Exit Sub
         End If
         If Len(txt1(9)) = 0 Then
             s = MsgBox("管制別不可空白!!", , "USER 輸入錯誤")
             txt1(9).SetFocus
             Exit Sub
         End If
         'Modify by Morgan 2006/5/23 管制函不可輸
         'Added by Lydia 2015/09/09 +5
         If txt1(9) <> "3" And txt1(9) <> "4" And txt1(9) <> "5" Then
            If Len(txt1(2)) = 0 Then
                s = MsgBox("發文日不可空白!!", , "USER 輸入錯誤")
                txt1(1).SetFocus
                txt1_GotFocus (1)
                Exit Sub
            End If
            
            If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
               Me.txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
               Me.txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
         Else
            If txt1(10) = "" Then
               s = MsgBox("管制日不可空白!!", , "USER 輸入錯誤")
                txt1(10).SetFocus
                txt1_GotFocus (10)
                Exit Sub
            End If
            
            'Added by Morgan 2016/5/23
            If txt1(9) <> "5" Then
               If txt1(11) = "" Then
                  s = MsgBox("管制迄日不可空白!!", , "USER 輸入錯誤")
                  txt1(11).SetFocus
                  txt1_GotFocus (11)
                  Exit Sub
               End If
            End If
            'end 2016/5/23
            
         End If
         
         If Len(Trim(txt1(7))) <> 0 Or Len(Trim(txt1(8))) <> 0 Then
             If Left(txt1(7), 6) <> Left(txt1(8), 6) Then
                 s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
                  txt1(7).SetFocus
                  txt1_GotFocus (7)
                 Exit Sub
             End If
         End If
         
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/9/30 清除查詢印表記錄檔欄位
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         'Modified by Lydia 2015/09/09
'         If txt1(9) = "4" Then
'            pub_QL05 = pub_QL05 & ";" & Label6 & "提申管制函" 'Add By Sindy 2010/9/30
'            StrMenu1
'         ElseIf txt1(9) = "3" Then
'            pub_QL05 = pub_QL05 & ";" & Label6 & "收達管制函" 'Add By Sindy 2010/9/30
'            StrMenu2
'         Else
'            StrMenu
'         End If

         FMP2openSQL = mOpenSql 'Added by Lydia 2015/10/26
         Select Case txt1(9)
            Case "3"
                pub_QL05 = pub_QL05 & ";" & Label6 & "收達管制函"
                StrMenu2
            Case "4"
                pub_QL05 = pub_QL05 & ";" & Label6 & "提申管制函"
                StrMenu1
            Case "5"
                pub_QL05 = pub_QL05 & ";" & Label6 & "公開管制函"
                StrMenu3
            Case Else
                StrMenu
         End Select
         'end 2015/09/09
         Me.Enabled = True
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
      Case Else
   End Select
End Sub

'Add by Morgan 2006/5/19
'提申管制函
Sub StrMenu1()
   Dim strCon As String, stVTB As String, iSNo As Integer
   Dim strAgent As String, bolEdit As Boolean
   Dim strDate As String, StrDate2 As String
   'Dim strNewCase As String 'Add by Morgan 2010/5/31 是否新申請案
   Dim strV1 As String '指定提申 1,其他 0
   Dim strV2 As String 'Added by Morgan 2016/5/16
   Dim strCP09 As String, strCP12 As String, strCP13 As String, strCP64 As String, bolInTrans As Boolean 'Added by Morgan 2016/5/13
   Dim strAF06 As String 'Added by Morgan 2016/5/18
   Dim strCatchNA16 As String 'Added by Lydia 2016/06/14
   
   strCon = ""
   
   '系統類別
   'Added by Morgan 2024/3/14
   If Len(txt1(0)) <> 0 Then
      strCon = strCon & " AND C1.CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0)
   End If
   'end 2024/3/14

   If txt1(5) <> "" Then
      strCon = strCon & " And PA09||''>='" & txt1(5) & "'"
   End If
   If txt1(6) <> "" Then
      strCon = strCon & " And PA09||''<='" & txt1(6) & "'"
   End If
   If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label4 & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/9/30
   End If
   
   'Added by Morgan 2016/5/17
   If List1(1).ListCount > 0 Then
      pub_QL05 = pub_QL05 & ";不管制案號" & List1(1).List(0)
      strCon = strCon & " and pa01||pa02||pa03||pa04 not in ('" & List1(1).List(0) & "'"
      For intI = 1 To List1(1).ListCount - 1
         strCon = strCon & ",'" & List1(1).List(intI) & "'"
         pub_QL05 = pub_QL05 & "," & List1(1).List(intI)
      Next
      strCon = strCon & ")"
   End If
   'end 2016/5/17
   
   If txt1(7) <> "" Then
      strCon = strCon & " And C1.CP44||''>='" & txt1(7) & "'"
   Else
      strCon = strCon & " And C1.CP44||''>='Y'"
   End If
   If txt1(8) <> "" Then
      strCon = strCon & " And C1.CP44||''<='" & txt1(8) & "'"
      If txt1(7) <> "" Then
         bolEdit = True
      End If
   End If
   If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label5 & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/9/30
   End If
   strDate = TransDate(txt1(10), 2)
   StrDate2 = TransDate(txt1(11), 2) 'Added by Morgan 2016/5/23
   
   If FMP2open Then 'Added by Morgan 2025/1/21 開放內專也可輸管制人且規則不同
   
      'Added by Lydia 2016/06/14 判斷FCP管制人 strCatchNa16
      strCatchNA16 = " and (cp01,cp02,cp03,cp04) in (SELECT x1.PA01,x1.PA02,x1.PA03,x1.PA04 FROM PATENT x1,FAGENT x2,NATION x3 WHERE x1.PA01=CP01 and x1.PA02=CP02 and x1.PA03=CP03 and x1.PA04=CP04 and SUBSTR(x1.PA75,1,8)=x2.FA01(+) AND SUBSTR(x1.PA75,9,1)=x2.FA02(+) AND x2.FA10=x3.NA01(+) "
      'Modified by Lydia 2017/02/13 +FMP管制人
      If strSrvDate(1) < FMP管制人啟用日 Then
          If txt1(12).Text <> "" Then strCatchNA16 = strCatchNA16 & "and x3.na16='" & Trim(txt1(12).Text) & "' "
      Else
          If txt1(12).Text <> "" Then strCatchNA16 = strCatchNA16 & "and decode(x1.PA01,'P',nvl(x3.na79,x3.na16),x3.na16)='" & Trim(txt1(12).Text) & "' "
      End If
      'end 2017/02/13
      
      strCatchNA16 = strCatchNA16 & ") "
      'end 2016/06/14
    
   End If 'Added by Morgan 2025/1/21
   
On Error GoTo ErrHandle

'Modify by Morgan 2010/7/27 +只管控已發文程序--玲玲

   '1 系統日=提申管制日(已收達)
   '2 系統日=提申管制日+N*提申管制週期(已收達)
   '3 系統日>=發文日+管制底限天數(已收達)
   '4 系統日+3工作天<=指定(最終)提申日 'Modify by Morgan 2011/7/15 改+3個工作天(要和AutoBatchDay同步),原來+3日曆天
'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    strCon = strCon & Replace(FMP2openSQL, "f0", "C1")
   '1
   stVTB = " select C1.CP09 KNo,0 V1,0 V2" & _
      " from nextprogress NP,caseprogress C1,patent,nation" & _
      " where np02='P' and np06 is null and np07='998' and np08>to_char(sysdate-180,'yyyymmdd')" & _
      " and cp09(+)=np01 and cp46>0 and cp57 is null and cp27>0" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa57 is null and pa01 is not null" & _
      " and na01(+)=pa09 and np08>=" & strDate & " and np08<=" & StrDate2 & strCon & _
      " AND NOT EXISTS(SELECT * FROM CASEPROGRESS C4 WHERE C4.CP01=C1.CP01 AND C4.CP02=C1.CP02 AND C4.CP03=C1.CP03 AND C4.CP04=C1.CP04 AND C4.CP10='936' AND C4.CP27 IS NULL and C4.cp57 is null)"
   '2
   stVTB = stVTB & " Union" & _
      " select C1.CP09 KNo,0 V1,0 V2" & _
      " from nextprogress NP,caseprogress C1,patent,nation" & _
      " where np02='P' and np06 is null and np07='998' and np08>to_char(sysdate-180,'yyyymmdd')" & _
      " and cp09(+)=np01 and cp46>0 and cp57 is null and cp27>0" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa57 is null and pa01 is not null" & _
      " and na01(+)=pa09 and np08<" & StrDate2 & strCon & _
      " and (MOD(TO_DATE(" & strDate & ",'YYYYMMDD')-TO_DATE(NP08,'YYYYMMDD'),NVL(NA61,15))=0 or MOD(TO_DATE(" & StrDate2 & ",'YYYYMMDD')-TO_DATE(NP08,'YYYYMMDD'),NVL(NA61,15))=0)" & _
      " and TO_DATE(" & StrDate2 & ",'YYYYMMDD')<TO_DATE(CP27,'YYYYMMDD')+NVL(NA62,90)" & _
      " AND NOT EXISTS(SELECT * FROM CASEPROGRESS C4 WHERE C4.CP01=C1.CP01 AND C4.CP02=C1.CP02 AND C4.CP03=C1.CP03 AND C4.CP04=C1.CP04 AND C4.CP10='936' AND C4.CP27 IS NULL and C4.cp57 is null)"
   '3
   stVTB = stVTB & " Union" & _
      " select C1.CP09 KNo,0 V1,0 V2" & _
      " from nextprogress NP,caseprogress C1,patent,nation" & _
      " where np02='P' and np06 is null and np07='998' and np08>to_char(sysdate-180,'yyyymmdd')" & _
      " and cp09(+)=np01 and cp46>0 and cp57 is null and cp27>0" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa57 is null and pa01 is not null" & strCon & _
      " and na01(+)=pa09 and TO_DATE(" & StrDate2 & ",'YYYYMMDD')>=TO_DATE(CP27,'YYYYMMDD')+NVL(NA62,90)" & _
      " AND NOT EXISTS(SELECT * FROM CASEPROGRESS C4 WHERE C4.CP01=C1.CP01 AND C4.CP02=C1.CP02 AND C4.CP03=C1.CP03 AND C4.CP04=C1.CP04 AND C4.CP10='936' AND C4.CP27 IS NULL and C4.cp57 is null)"
   '4
   'Modify by Morgan 2011/7/15 改+3個工作天(要和AutoBatchDay同步)
   'Modified by Morgan 2012/2/15 指定提申日改單獨
   stVTB = stVTB & " Union" & _
      " select C1.CP09 KNo,0 V1,0 V2" & _
      " from (select WORKDAYADD(4," & StrDate2 & ") dd from dual) X,nextprogress NP,caseprogress C1,patent,nation" & _
      " where np08>to_char(sysdate-180,'yyyymmdd') and np08<=dd and np02='P' and np06 is null and np07='996'" & _
      " and cp09(+)=np01 and cp57 is null and cp27>0" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa57 is null and pa01 is not null" & strCon & _
      " and na01(+)=pa09 "
      
   'Added by Morgan 2012/2/15 +指定提申日抓前一個工作日
   '5
   stVTB = stVTB & " Union" & _
      " select C1.CP09 KNo,1 V1,np08 V2" & _
      " from (select WORKDAYADD(2," & StrDate2 & ") dd from dual) X,nextprogress NP,caseprogress C1,patent,nation" & _
      " where np08>to_char(sysdate-180,'yyyymmdd') and np08<=dd and np02='P' and np06 is null and np07='995'" & _
      " and cp09(+)=np01 and cp57 is null and cp27>0" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa57 is null and pa01 is not null" & strCon & _
      " and na01(+)=pa09 "
   
   'Modify by Morgan 2010/5/31 +是否新案
   'Modify by Morgan 2010/6/10 改回一代理人一指示信
   'strSql = "SELECT CP44 C00,CP09 C01,CP45 C02,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04) C03" & _
      ",NVL(PA05,NVL(PA06,PA07)) C04,decode(instr('101,102,103',cp10),0,0,1) C05,CP10 C06" & _
      " FROM (" & stVTB & ") X,caseprogress,patent" & _
      " where CP09=KNo and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " ORDER BY 6,1,2,3,4"
   'Modified by Lydia 2016/06/14 判斷FCP管制人 strCatchNa16
   'Modified by Morgan 2018/7/30 +CP10,PA09
   'Modified by Morgan 2025/1/21 +PID
   strSql = "SELECT CP44 C00,CP09 C01,CP45 C02,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04) C03" & _
      ",NVL(PA05,NVL(PA06,PA07)) C04,decode(CP10,'202','補'||cp64,cpm04) C05,PA11 C06,V1,V2,CP01,CP02,CP03,CP04,CP10,PA09,'' PID" & _
      " FROM (" & stVTB & ") X,caseprogress,patent,casepropertymap" & _
      " where CP09=KNo and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10" & strCatchNA16 & _
      " ORDER BY 1,V2,4,3,2"
      
'end 2009/7/3
      
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   
   'Added by Morgan 2025/1/21
   If FMP2open = False And adoRecordset.RecordCount > 0 And strSrvDate(1) >= P業務區劃分啟用日 And txt1(12).Text <> "" Then
      Set RsTemp = PUB_CreateRecordset(adoRecordset, , , 300, Me.Name, mSeqNo)
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            If Left(.Fields("C03"), 2) = "P-" Then
               .Fields("PID") = PUB_GetPHandler(.Fields("C03"))
            End If
            .MoveNext
         Loop
         .UpdateBatch
         
         stVTBX = "select R001 as " & .Fields(0).Name
         For intI = 2 To .Fields.Count
            stVTBX = stVTBX & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
         Next
         stVTBX = stVTBX & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
      End With
      strSql = "Select X.* From (" & stVTBX & ") X where PID='" & txt1(12) & "' ORDER BY 1,V2,4,3,2"
      intI = 1
      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   End If
   'end 2025/1/21
   
   With adoRecordset
      If .RecordCount > 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/9/30
         iSNo = 0
         strAgent = .Fields(0)
         m_strET02 = "" & .Fields(1)
         strV1 = .Fields("V1") 'Added by Morgan 2012/2/15
         strV2 = .Fields("V2") 'Added by Morgan 2016/5/16
         Do While Not .EOF
            'Modified by Morgan 2016/5/16 指定日期不同也要分開
            If strAgent <> "" & .Fields(0) Or strV1 <> .Fields("V1") Or strV2 <> .Fields("V2") Then
               If strV1 = 1 Then
                  Export2Word "02", "14", bolEdit
               Else
                  Export2Word "02", "07", bolEdit
               End If
               iSNo = 0
               strAgent = .Fields(0)
               m_strET02 = "" & .Fields(1)
               strV1 = .Fields("V1") 'Added by Morgan 2012/2/15
               strV2 = .Fields("V2") 'Added by Morgan 2016/5/16
            End If
            
            'Added by Morgan 2016/5/13
            '指示信電子化
            If Left(Pub_StrUserSt03, 1) <> "F" Then
               'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
               'If strAF06 = "" Then strAF06 = Pub_GetSpecMan("PS4") 'P案指示信判發人
               strAF06 = PUB_GetLetterJudgeNew("2", "P", "953", .Fields("PA09"), .Fields("CP10"))
               'END 2018/7/30
               
               cnnConnection.BeginTrans
               bolInTrans = True
               '新增"催提申"進度
               m_strCP10 = "953"
               strCP09 = AutoNo("B", 6)
               If iSNo = 0 Then
                  m_strCP01 = .Fields("CP01")
                  m_strCP02 = .Fields("CP02")
                  m_strCP03 = .Fields("CP03")
                  m_strCP04 = .Fields("CP04")
                  m_strCP09 = strCP09
                  strCP64 = "指示信存於" & .Fields("C03") & "案(" & m_strCP09 & ")卷宗區"
               End If
               
               strCP13 = PUB_GetAKindSalesNo(.Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04"))
               strCP12 = GetSalesArea(strCP13)
               strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
                  "CP12,CP13,CP14,CP20,CP26,CP28,CP32,CP43,CP44,CP45,CP64) VALUES ('" & .Fields("CP01") & "','" & .Fields("CP02") & "'" & _
                  ",'" & .Fields("CP03") & "','" & .Fields("CP04") & "'," & strSrvDate(1) & ",'" & strCP09 & "','" & m_strCP10 & "','" & strCP12 & "'" & _
                  ",'" & strCP13 & "','" & strUserNum & "','N','N','" & m_strCP09 & "','N','" & .Fields("C01") & "','" & .Fields("C00") & "','" & .Fields("C02") & "','" & IIf(iSNo = 0, "", strCP64) & "')"
               cnnConnection.Execute strSql, intI
               
               If iSNo = 0 Then
                  m_Subject = "請回覆提申進度，Y/R：" & IIf(Trim("" & .Fields("C02")) = "", "(請提供)", "" & .Fields("C02")) & "，O/R：" & .Fields("C03") & "，謝謝。"                  '
                  PUB_AddAppForm m_strCP09, True, strAF06, m_Subject  '自行判發,不轉檔(一定要先看過)
               Else
                  'Modified by Morgan 2016/6/22--蕭茹曣
                  'm_Subject = "請回覆提申進度。"
                  If iSNo = 1 Then
                     m_Subject = Replace(m_Subject, "，O/R：", "等...案，O/R：")
                     m_Subject = Replace(m_Subject, "，謝謝。", "等...案，謝謝。")
                  End If
                  strSql = "update appform set af13='" & m_Subject & "' where af01='" & m_strCP09 & "'"
                  cnnConnection.Execute strSql, intI
               End If
               cnnConnection.CommitTrans
               bolInTrans = False
            End If
            'end 2016/5/13
            
            iSNo = iSNo + 1
            If iSNo = 1 Then
               m_strCaseNo1 = "" & .Fields(2)
               m_strCaseNo2 = "" & .Fields(3) & "(" & .Fields(5) & ")"
               m_strCaseName = Left("" & .Fields(4) & Space(40), 40)
               'Add by Morgan 2010/8/18 +申請案號
               m_strCaseNo3 = "" & .Fields(6)
               'Added by Morgan 2012/9/26
               If IsNull(.Fields(6)) Then
                  m_bolShowNo3 = False
               Else
                  m_bolShowNo3 = True
               End If
            ElseIf iSNo = 2 Then
               m_strCaseNo1 = "1. " & m_strCaseNo1 & vbCrLf & String(5, "　") & "2. " & .Fields(2)
               m_strCaseNo2 = "1. " & m_strCaseNo2 & vbCrLf & String(5, "　") & "2. " & .Fields(3) & "(" & .Fields(5) & ")"
               m_strCaseName = "1. " & m_strCaseName & vbCrLf & String(5, "　") & "2. " & Left("" & .Fields(4) & Space(40), 40)
               'Add by Morgan 2010/8/18 +申請案號
               m_strCaseNo3 = "1. " & m_strCaseNo3 & vbCrLf & String(5, "　") & "2. " & .Fields(6)
               If Not IsNull(.Fields(6)) Then m_bolShowNo3 = True 'Added by Morgan 2012/9/26
            Else
               m_strCaseNo1 = m_strCaseNo1 & vbCrLf & String(5, "　") & iSNo & ". " & .Fields(2)
               m_strCaseNo2 = m_strCaseNo2 & vbCrLf & String(5, "　") & iSNo & ". " & .Fields(3) & "(" & .Fields(5) & ")"
               m_strCaseName = m_strCaseName & vbCrLf & String(5, "　") & iSNo & ". " & Left("" & .Fields(4) & Space(40), 40)
               'Add by Morgan 2010/8/18 +申請案號
               m_strCaseNo3 = m_strCaseNo3 & vbCrLf & String(5, "　") & iSNo & ". " & .Fields(6)
               If Not IsNull(.Fields(6)) Then m_bolShowNo3 = True 'Added by Morgan 2012/9/26
            End If
            .MoveNext
         Loop
         
         'If strNewCase = "1" Then
            'Modified by Morgan 2012/2/15 指定提申定稿另外
            If strV1 = 1 Then
               Export2Word "02", "14", bolEdit
            Else
               Export2Word "02", "07", bolEdit
            End If
         'End If
         If bolEdit = False Then
            MsgBox "列印完畢！"
            'Removed by Morgan 2016/5/23 指示信改都要開啟維護畫面,因上傳後系統會通知判發人故此處不必再通知
            'If strAF06 <> "" Then PUB_SendMail strUserNum, strAF06, "", "請判發催提申指示信！", "如旨"
         End If
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/9/30
         MsgBox "無符合資料！"
      End If
         End With
   
ErrHandle:
   If bolInTrans Then cnnConnection.RollbackTrans 'Added by Morgan 2016/5/13
   If Err.Number <> 0 Then MsgBox Err.Description

End Sub

'Add by Morgan 2006/5/23
'收達管制函
Sub StrMenu2()
   Dim strCon As String, stVTB As String, iSNo As Integer
   Dim strAgent As String, bolEdit As Boolean
   Dim strDate As String, StrDate2 As String
   'Dim strNewCase As String 'Add by Morgan 2010/5/31 是否新申請案
   Dim strCP09 As String, strCP12 As String, strCP13 As String, strCP64 As String, bolInTrans As Boolean 'Added by Morgan 2016/5/13
   Dim strAF06 As String 'Added by Morgan 2016/5/18
   Dim strCatchNA16 As String 'Added by Lydia 2016/06/14
   
   strCon = ""
   
   '系統類別
   'Added by Morgan 2024/3/14
   If Len(txt1(0)) <> 0 Then
      strCon = strCon & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0)
   End If
   'end 2024/3/14
   
   If txt1(5) <> "" Then
      strCon = strCon & " And PA09||''>='" & txt1(5) & "'"
   End If
   If txt1(6) <> "" Then
      strCon = strCon & " And PA09||''<='" & txt1(6) & "'"
   End If
   If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label4 & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/9/30
   End If
   
   'Added by Morgan 2016/5/17
   If List1(1).ListCount > 0 Then
      pub_QL05 = pub_QL05 & ";不管制案號" & List1(1).List(0)
      strCon = strCon & " and pa01||pa02||pa03||pa04 not in ('" & List1(1).List(0) & "'"
      For intI = 1 To List1(1).ListCount - 1
         strCon = strCon & ",'" & List1(1).List(intI) & "'"
         pub_QL05 = pub_QL05 & "," & List1(1).List(intI)
      Next
      strCon = strCon & ")"
   End If
   'end 2016/5/17
   
   If txt1(7) <> "" Then
      strCon = strCon & " And CP44||''>='" & txt1(7) & "'"
   Else
      strCon = strCon & " And CP44||''>='Y'"
   End If
   If txt1(8) <> "" Then
      strCon = strCon & " And CP44||''<='" & txt1(8) & "'"
      If txt1(7) <> "" Then
         bolEdit = True
      End If
   End If
   If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label5 & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/9/30
   End If
   strDate = TransDate(txt1(10), 2)
   StrDate2 = TransDate(txt1(11), 2) 'Added by Morgan 2016/5/23
   
   If FMP2open Then 'Added by Morgan 2025/1/21 開放內專也可輸管制人且規則不同
   
      'Added by Lydia 2016/06/14 判斷FCP管制人 strCatchNa16
      strCatchNA16 = " and (f0.cp01,f0.cp02,f0.cp03,f0.cp04) in (SELECT x1.PA01,x1.PA02,x1.PA03,x1.PA04 FROM PATENT x1,FAGENT x2,NATION x3 WHERE x1.PA01=f0.CP01 and x1.PA02=f0.CP02 and x1.PA03=f0.CP03 and x1.PA04=f0.CP04 and SUBSTR(x1.PA75,1,8)=x2.FA01(+) AND SUBSTR(x1.PA75,9,1)=x2.FA02(+) AND x2.FA10=x3.NA01(+) "
      'Modified by Lydia 2017/02/13 +FMP管制人
      If strSrvDate(1) < FMP管制人啟用日 Then
          If txt1(12).Text <> "" Then strCatchNA16 = strCatchNA16 & "and x3.na16='" & Trim(txt1(12).Text) & "' "
      Else
          If txt1(12).Text <> "" Then strCatchNA16 = strCatchNA16 & "and decode(x1.PA01,'P',nvl(x3.na79,x3.na16),x3.na16)='" & Trim(txt1(12).Text) & "' "
      End If
      'end 2017/02/13
      
      strCatchNA16 = strCatchNA16 & ") "
      'end 2016/06/14
      
   End If 'Added by Morgan 2025/1/21

On Error GoTo ErrHandle
   
   'Modify by Morgan 2010/7/27 +只管控已發文程序--玲玲
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件。  +別名f0,FMP2openSQL
   'Modified by Lydia 2016/06/14 判斷FCP管制人 strCatchNa16
   'Modified by Morgan 2018/7/30 +CP10,PA09
   'Modified by Morgan 2025/1/21 +PID
   strSql = "SELECT  CP44 C00,CP09 C01,CP45 C02,CP01||'-'||CP02||DECODE(CP03||CP04,'000',NULL,'-'||CP03||'-'||CP04) C03,NVL(PA05,NVL(PA06,PA07)) C04" & _
      ",decode(CP10,'202','補'||cp64,cpm04) C05,PA11 C06,CP01,CP02,CP03,CP04,CP10,PA09,'' PID FROM nextprogress,caseprogress f0,patent,casepropertymap" & _
      " where NP02 in ('P','PS') AND NP06 IS NULL AND NP07='997'" & _
      " and NP02 in (" & SQLGrpStr(txt1(0), 1) & "," & SQLGrpStr(txt1(0), 5) & ")" & _
      " AND CP09(+)=NP01" & _
      " AND CP24 IS NULL AND CP46 IS NULL AND CP47 IS NULL AND CP57 IS NULL and cp27>0" & _
      " AND (TO_DATE(NP08,'YYYYMMDD')=TRUNC(TO_DATE(" & strDate & ",'YYYYMMDD')) OR TO_DATE(NP08,'YYYYMMDD')=TRUNC(TO_DATE(" & StrDate2 & ",'YYYYMMDD')) OR TO_DATE(CP27,'YYYYMMDD')<=TRUNC(TO_DATE(" & StrDate2 & ",'YYYYMMDD')-7))" & _
      " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
      " AND PA57||PA108 IS NULL" & strCon & FMP2openSQL & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10" & strCatchNA16 & _
      " ORDER BY 1,4,3,2"

   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   
   'Added by Morgan 2025/1/21
   If FMP2open = False And adoRecordset.RecordCount > 0 And strSrvDate(1) >= P業務區劃分啟用日 And txt1(12).Text <> "" Then
      Set RsTemp = PUB_CreateRecordset(adoRecordset, , , 300, Me.Name, mSeqNo)
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            If Left(.Fields("C03"), 2) = "P-" Then
               .Fields("PID") = PUB_GetPHandler(.Fields("C03"))
            End If
            .MoveNext
         Loop
         .UpdateBatch
         
         stVTBX = "select R001 as " & .Fields(0).Name
         For intI = 2 To .Fields.Count
            stVTBX = stVTBX & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
         Next
         stVTBX = stVTBX & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
      End With
      strSql = "Select X.* From (" & stVTBX & ") X where PID='" & txt1(12) & "' ORDER BY 1,4,3,2"
      intI = 1
      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   End If
   'end 2025/1/21
   
   With adoRecordset
      If .RecordCount > 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/9/30
         iSNo = 0
         strAgent = .Fields(0)
         m_strET02 = "" & .Fields(1)
         Do While Not .EOF
            If strAgent <> "" & .Fields(0) Then
               Export2Word "02", "08", bolEdit
               iSNo = 0
               strAgent = .Fields(0)
               m_strET02 = "" & .Fields(1)
            End If
               
            'Added by Morgan 2016/5/13
            '指示信電子化
            If Left(Pub_StrUserSt03, 1) <> "F" Then
               'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
               'If strAF06 = "" Then strAF06 = Pub_GetSpecMan("PS4") 'P案指示信判發人
               strAF06 = PUB_GetLetterJudgeNew("2", "P", "952", .Fields("PA09"), .Fields("CP10"))
               'END 2018/7/30
               cnnConnection.BeginTrans
               bolInTrans = True
               '新增"催收達"進度
               m_strCP10 = "952"
               strCP09 = AutoNo("B", 6)
               If iSNo = 0 Then
                  m_strCP01 = .Fields("CP01")
                  m_strCP02 = .Fields("CP02")
                  m_strCP03 = .Fields("CP03")
                  m_strCP04 = .Fields("CP04")
                  m_strCP09 = strCP09
                  strCP64 = "指示信存於" & .Fields("C03") & "案(" & m_strCP09 & ")卷宗區"
               End If
               
               strCP13 = PUB_GetAKindSalesNo(.Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04"))
               strCP12 = GetSalesArea(strCP13)
               strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
                  "CP12,CP13,CP14,CP20,CP26,CP28,CP32,CP43,CP44,CP45,CP64) VALUES ('" & .Fields("CP01") & "','" & .Fields("CP02") & "'" & _
                  ",'" & .Fields("CP03") & "','" & .Fields("CP04") & "'," & strSrvDate(1) & ",'" & strCP09 & "','" & m_strCP10 & "','" & strCP12 & "'" & _
                  ",'" & strCP13 & "','" & strUserNum & "','N','N','" & m_strCP09 & "','N','" & .Fields("C01") & "','" & .Fields("C00") & "','" & .Fields("C02") & "','" & IIf(iSNo = 0, "", strCP64) & "')"
               cnnConnection.Execute strSql, intI
               
               If iSNo = 0 Then
                  m_Subject = "請確認是否收達，Y/R：" & IIf(Trim("" & .Fields("C02")) = "", "(請提供)", "" & .Fields("C02")) & "，O/R：" & .Fields("C03") & "，謝謝。"
                  PUB_AddAppForm m_strCP09, True, strAF06, m_Subject '自行判發,不轉檔(一定要先看過)
               Else
                  'Modified by Morgan 2016/6/22--蕭茹曣
                  'm_Subject = "請確認是否收達。"
                  If iSNo = 1 Then
                     m_Subject = Replace(m_Subject, "，O/R：", "等...案，O/R：")
                     m_Subject = Replace(m_Subject, "，謝謝。", "等...案，謝謝。")
                  End If
                  'end 2016/6/22
                  strSql = "update appform set af13='" & m_Subject & "' where af01='" & m_strCP09 & "'"
                  cnnConnection.Execute strSql, intI
               End If
               cnnConnection.CommitTrans
               bolInTrans = False
            End If
            'end 2016/5/13
               
            iSNo = iSNo + 1
            If iSNo = 1 Then
               m_strCaseNo1 = "" & .Fields(2)
               m_strCaseNo2 = "" & .Fields(3) & "(" & .Fields(5) & ")"
               m_strCaseName = Left("" & .Fields(4) & Space(40), 40)
               'Add by Morgan 2010/8/18 +申請案號
               m_strCaseNo3 = "" & .Fields(6)
               'Added by Morgan 2012/9/26
               If IsNull(.Fields(6)) Then
                  m_bolShowNo3 = False
               Else
                  m_bolShowNo3 = True
               End If
            ElseIf iSNo = 2 Then
               m_strCaseNo1 = "1. " & m_strCaseNo1 & vbCrLf & String(5, "　") & "2. " & .Fields(2)
               m_strCaseNo2 = "1. " & m_strCaseNo2 & vbCrLf & String(5, "　") & "2. " & .Fields(3) & "(" & .Fields(5) & ")"
               m_strCaseName = "1. " & m_strCaseName & vbCr & String(5, "　") & "2. " & Left("" & .Fields(4) & Space(40), 40)
               'Add by Morgan 2010/8/18 +申請案號
               m_strCaseNo3 = "1. " & m_strCaseNo3 & vbCrLf & String(5, "　") & "2. " & .Fields(6)
               If Not IsNull(.Fields(6)) Then m_bolShowNo3 = True 'Added by Morgan 2012/9/26
            Else
               m_strCaseNo1 = m_strCaseNo1 & vbCrLf & String(5, "　") & iSNo & ". " & .Fields(2)
               m_strCaseNo2 = m_strCaseNo2 & vbCrLf & String(5, "　") & iSNo & ". " & .Fields(3) & "(" & .Fields(5) & ")"
               m_strCaseName = m_strCaseName & vbCrLf & String(5, "　") & iSNo & ". " & Left("" & .Fields(4) & Space(40), 40)
               'Add by Morgan 2010/8/18 +申請案號
               m_strCaseNo3 = m_strCaseNo3 & vbCrLf & String(5, "　") & iSNo & ". " & .Fields(6)
               If Not IsNull(.Fields(6)) Then m_bolShowNo3 = True 'Added by Morgan 2012/9/26
            End If
            .MoveNext
         Loop
         Export2Word "02", "08", bolEdit
         
         If bolEdit = False Then
            MsgBox "列印完畢！"
            'Removed by Morgan 2016/5/23 指示信改都要開啟維護畫面,因上傳後系統會通知判發人故此處不必再通知
            'If strAF06 <> "" Then PUB_SendMail strUserNum, strAF06, "", "請判發催收達指示信！", "如旨"
         End If
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/9/30
         MsgBox "無符合資料！"
      End If
   End With
   
ErrHandle:
   If bolInTrans Then cnnConnection.RollbackTrans 'Added by Morgan 2016/5/13
   If Err.Number <> 0 Then MsgBox Err.Description
   
End Sub

Private Sub Export2Word(ByVal ET01 As String, ByVal ET03 As String, ByVal p_bEdit As Boolean)

   Dim bolShow As Boolean

   StartLetter ET01, ET03
   
On Error GoTo ErrHnd
   bolShow = g_WordAp.Visible
   
   'Added byMorgan 2016/5/13
   '非臺灣案電子化
   If Left(Pub_StrUserSt03, 1) = "F" Then
   'end 2016/5/13
      NowPrint m_strET02, ET01, ET03, True, strUserNum
      If Not p_bEdit Then
         '列印定稿
         'Modified by Morgan 2015/12/3 改只要印一份--玲玲
         g_WordAp.PrintOut Copies:=1, Collate:=True: DoEvents
         g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
         g_WordAp.Visible = bolShow
      End If
      
   'Added byMorgan 2016/5/13
   '非臺灣案電子化
   Else
      NowPrint m_strET02, ET01, ET03, p_bEdit, strUserNum, , , , , 1, , , , , , , , m_strCP09
      If p_bEdit Then
         strExc(1) = PUB_CaseNo2FileName(m_strCP01, m_strCP02, m_strCP03, m_strCP04)
         frm1105_1.m_RecNo = m_strCP09
         frm1105_1.m_PdfName = strExc(1) & "." & m_strCP10 & ".DATA.PDF"
         frm1105_1.m_Subject = m_Subject
         frm1105_1.Show
      End If
      m_strCP09 = ""
   End If
   'end 2016/5/13

ErrHnd:
   If Err.Number = 91 Or Err.Number = 462 Then
      bolShow = False
      Resume Next
   ElseIf Err.Number <> 0 Then
      Err.Raise Err.Number
   End If
   
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(1 To 5) As String
   Dim ii As Integer
   EndLetter ET01, m_strET02, ET03, strUserNum
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & m_strET02 & "','" & ET03 & "','" & strUserNum & "','專利名稱s','" & m_strCaseName & "')"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & m_strET02 & "','" & ET03 & "','" & strUserNum & "','貴方卷號s','" & m_strCaseNo1 & "')"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & m_strET02 & "','" & ET03 & "','" & strUserNum & "','我方案號s','" & m_strCaseNo2 & "')"
      
   If m_bolShowNo3 = True Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & m_strET02 & "','" & ET03 & "','" & strUserNum & "','申請案號s','" & m_strCaseNo3 & "')"
   End If
   'Removed by Morgan 2012/8/31 改抓共用例外欄位
   'ii = ii + 1
   'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
   '   "('" & ET01 & "','" & m_strET02 & "','" & ET03 & "','" & strUserNum & "','指定提申日','" & TransDate(txt1(10), 2) & "')"
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(ii, strTxt) Then
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Sub StrMenu()
'Add By Cheng 2002/07/01
Dim StrSQLa As String
Dim strCatchNA16 As String 'Added by Lydia 2016/06/14
Dim strConLast As String, strNA69 As String 'Added by Lydia 2022/09/20

Screen.MousePointer = vbHourglass
StrTest1 = ""      '專利

'Morgan 2003/11/21
'StrTest2 = ""      '商標
'StrTest3 = ""      '服務
Dim strExceptCFT As String
'modify by sonia 2017/8/30 +623,624
'Modified by Lydia 2024/07/31 增加下列案件性質：S-001-查名,CFT-201-補正,CFT-202-申請意見書（CFT,TF:答辯）,CFT-203-修正,CFT-204-準備程序,CFT-205-言詞辯論,CFT-207-聲明參訟,CFT-211-檢送同意書,CFT-305-催審,CFT-306-自請撤回,CFT-307-自請拋棄商標權,CFT-308-分割,CFT-310-暫緩審理,CFT-311-加速審查,CFT-313-減縮商品,CFT-402-再訴願,CFT-403-行政訴訟(CFT：訴訟),CFT-404-再審之訴(CFT：再審),CFT-406-參加訴願,CFT-407-參加訴訟,CFT-408-行政訴訟上訴,CFT-410-行政上訴答辯,CFT-505-廢止再授權,CFT-611-補理由書,CFT-612-補充理由,CFT-613-補充答辯,CFT-620-行政查處,CFT-623-部分廢止,CFT-624-部分廢止答辯,CFT-627-部分異議,CFT-628-部分異議答辯,CFT-629-部分評定,CFT-630-部分評定答辯,S-707-調查,CFT-707-調查,S-711-文件公／簽證,CFT-730-海關登記,TF-202-答辯,TF-401-復審,TF-701-領證,TF-105-使用宣誓
'strExceptCFT = " AND CP10 NOT IN ('101','102','103','105','202','301','401','501','502','503','504','506','507','601','602','603','604','605','606','623','624','701','702','706','708','709')"
'Modified by Lydia 2024/08/01 CFT不用711文件公/簽證; 區分不同變數
'strExceptCFT = " AND CP10 NOT IN ('001','101','102','103','105','201','202','203','204','205','207','211', " & _
               "'301','305','306','307','308','310','311','313','401','402','403','404','406','407','408','410'," & _
               "'501','502','503','504','505','506','507','601','602','603','604','605','606','611','612','613','620','623','624','627','628','629','630'," & _
               "'701','702','706','708','707','709','711','730')"
Dim strExceptS As String
'原本的範圍，追加本次列出的CFT,T,TF
'Modified by Lydia 2025/08/13 改成不同變數；目前設定=>101,102,103,105,201,202,203,204,205,207,211,301,305,306,307,308,310,311,313,401,402,403,404,406,407,408,410,501,502,503,504,505,506,507,601,602,603,604,605,606,611,612,613,620,623,624,627,628,629,630,701,702,706,707,708,709,730
'strExceptCFT = " AND CP10 NOT IN ('101','102','103','105','201','202','203','204','205','207','211','301','305','306','307','308','310','311','313','401','402','403','404','406','407','408','410','501','502','503','504','505','506','507','601','602','603','604','605','606','611','612','613','620','623','624','627','628','629','630','701','702','706','707','708','709','730')"
strExceptCFT = " AND CP10 NOT IN (" & GetAddStr(Replace(Pub_GetSpecMan("CFT提申管制_CFT"), ";", ",")) & ")"
'原本的範圍，追加本次列出的CFC,S
'Modified by Lydia 2025/08/13 改成不同變數；目前設定=>001,103,202,501,701,706,707,709,711
'strExceptS = " AND CP10 NOT IN ('001','101','102','103','105','202','301','401','501','502','503','504','506','507','601','602','603','604','605','606','623','624','701','702','706','708','707','709','711')"
strExceptS = " AND CP10 NOT IN (" & GetAddStr(Replace(Pub_GetSpecMan("CFT提申管制_CFC_S"), ";", ",")) & ")"
'end 2024/08/01

StrTest2 = " AND NOT( CP01='CFT' " & strExceptCFT & " )"                '商標
'Modified by Lydia 2024/08/01
'StrTest3 = " AND NOT( CP01 IN ('CFC','S') " & strExceptCFT & " )"       '服務
StrTest3 = " AND NOT( CP01 IN ('CFC','S') " & strExceptS & " )"       '服務
'--- end
'2005/12/26 ADD BY SONIA 內商此類案件性質不管制
Dim strExceptT As String
If intPCaseKind = 商標 And intPWhere = 國內 Then
   'Modify By Sindy 2020/3/5 + 705.補收款 => 不管制
   strExceptT = " AND CP10 IN ('701','706','714','705')"
   StrTest2 = StrTest2 & " AND NOT( CP01 IN ('T','TF') " & strExceptT & " )"                '商標
   StrTest3 = StrTest3 & " AND NOT( CP01 IN ('TB','TC','TD','TM','TR','TT','TS') " & strExceptT & " )"       '服務
End If
'2005/12/26 END
'Add By Cheng 2002/07/01
StrSQLa = "" '代理人名稱

'Exit Sub

'系統類別
If Len(txt1(0)) <> 0 Then
   StrTest1 = StrTest1 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
   StrTest2 = StrTest2 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
   StrTest3 = StrTest3 & " AND CP01 IN (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/9/30
End If
'Add By Cheng 2003/04/14
'發文日
If Len(txt1(1)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP27>='" & ChangeTStringToWString(txt1(1)) & "' "
    StrTest2 = StrTest2 + " AND CP27>='" & ChangeTStringToWString(txt1(1)) & "' "
    StrTest3 = StrTest3 + " AND CP27>='" & ChangeTStringToWString(txt1(1)) & "' "
End If
If Len(txt1(2)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP27<='" & ChangeTStringToWString(txt1(2)) & "' "
    StrTest2 = StrTest2 + " AND CP27<='" & ChangeTStringToWString(txt1(2)) & "' "
    StrTest3 = StrTest3 + " AND CP27<='" & ChangeTStringToWString(txt1(2)) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/9/30
End If
'案件性質
If Len(txt1(3)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP10>='" & txt1(3) & "' "
    StrTest2 = StrTest2 + " AND CP10>='" & txt1(3) & "' "
    StrTest3 = StrTest3 + " AND CP10>='" & txt1(3) & "' "
End If
If Len(txt1(4)) <> 0 Then
    StrTest1 = StrTest1 + " AND CP10<='" & txt1(4) & "' "
    StrTest2 = StrTest2 + " AND CP10<='" & txt1(4) & "' "
    StrTest3 = StrTest3 + " AND CP10<='" & txt1(4) & "' "
End If
If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/9/30
End If
'申請國家
If Len(txt1(5)) <> 0 Then
    StrTest3 = StrTest3 + " AND SP09>='" & txt1(5) & "' "
    StrTest2 = StrTest2 + " AND TM10>='" & txt1(5) & "' "
    StrTest1 = StrTest1 + " AND PA09>='" & txt1(5) & "' "
End If
If Len(txt1(6)) <> 0 Then
    StrTest3 = StrTest3 + " AND SP09<='" & txt1(6) & "' "
    StrTest2 = StrTest2 + " AND TM10<='" & txt1(6) & "' "
    StrTest1 = StrTest1 + " AND PA09<='" & txt1(6) & "' "
End If

'2011/7/1 MODIFY BY SONIA 剔除申請國家為台灣者,因為T-171766不知為何會有CP44
StrTest1 = StrTest1 + " AND PA09<>'000' "
StrTest2 = StrTest2 + " AND TM10<>'000' "
StrTest3 = StrTest3 + " AND SP09<>'000' "
'2011/7/1 END

If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label4 & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/9/30
End If
'代理人
If Len(Trim(txt1(7))) <> 0 And Len(Trim(txt1(8))) <> 0 Then
    StrTest1 = StrTest1 + " AND (CP44>='" & GetNewFagent(txt1(7)) & "' AND CP44<='" & GetNewFagent(txt1(8)) & "') "
    StrTest2 = StrTest2 + " AND (CP44>='" & GetNewFagent(txt1(7)) & "' AND CP44<='" & GetNewFagent(txt1(8)) & "') "
    StrTest3 = StrTest3 + " AND (CP44>='" & GetNewFagent(txt1(7)) & "' AND CP44<='" & GetNewFagent(txt1(8)) & "') "
Else
    If Len(Trim(txt1(7))) <> 0 And Len(Trim(txt1(8))) = 0 Then
        StrTest1 = StrTest1 + " AND (CP44>='" & GetNewFagent(txt1(7)) & "') "
        StrTest2 = StrTest2 + " AND (CP44>='" & GetNewFagent(txt1(7)) & "') "
        StrTest3 = StrTest3 + " AND (CP44>='" & GetNewFagent(txt1(7)) & "') "
    Else
        If Len(Trim(txt1(7))) = 0 And Len(Trim(txt1(8))) <> 0 Then
            StrTest1 = StrTest1 + " AND (CP44<='" & GetNewFagent(txt1(8)) & "') "
            StrTest2 = StrTest2 + " AND (CP44<='" & GetNewFagent(txt1(8)) & "') "
            StrTest3 = StrTest3 + " AND (CP44<='" & GetNewFagent(txt1(8)) & "') "
        End If
    End If
End If
If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label5 & txt1(7) & "-" & txt1(8) 'Add By Sindy 2010/9/30
End If
'Add By Cheng 2002/12/17
StrTest1 = StrTest1 & " AND (CP24 IS NULL OR CP24 = '') "
StrTest2 = StrTest2 & " AND (CP24 IS NULL OR CP24 = '') "
StrTest3 = StrTest3 & " AND (CP24 IS NULL OR CP24 = '') "
'貝爾查名不抓Y49579000 2006/2/7也不抓Y00000000
StrTest1 = StrTest1 & " AND CP44<>'Y49579000' AND CP44<>'Y00000000' "
StrTest2 = StrTest2 & " AND CP44<>'Y49579000' AND CP44<>'Y00000000' "
StrTest3 = StrTest3 & " AND CP44<>'Y49579000' AND CP44<>'Y00000000' "

'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
StrTest1 = StrTest1 & FMP2openSQL
StrTest2 = StrTest2 & FMP2openSQL
StrTest3 = StrTest3 & FMP2openSQL

If FMP2open Then 'Added by Morgan 2025/1/21 開放內專也可輸管制人且規則不同

   'Added by Lydia 2016/06/14 判斷FCP管制人 strCatchNa16
   strCatchNA16 = " and (f0.cp01,f0.cp02,f0.cp03,f0.cp04) in (SELECT x1.PA01,x1.PA02,x1.PA03,x1.PA04 FROM PATENT x1,FAGENT x2,NATION x3 WHERE x1.PA01=f0.CP01 and x1.PA02=f0.CP02 and x1.PA03=f0.CP03 and x1.PA04=f0.CP04 and SUBSTR(x1.PA75,1,8)=x2.FA01(+) AND SUBSTR(x1.PA75,9,1)=x2.FA02(+) AND x2.FA10=x3.NA01(+) "
   'Modified by Lydia 2017/02/13 +FMP管制人
   If strSrvDate(1) < FMP管制人啟用日 Then
       If txt1(12).Text <> "" Then strCatchNA16 = strCatchNA16 & "and x3.na16='" & Trim(txt1(12).Text) & "' "
   Else
       If txt1(12).Text <> "" Then strCatchNA16 = strCatchNA16 & "and decode(x1.PA01,'P',nvl(x3.na79,x3.na16),x3.na16)='" & Trim(txt1(12).Text) & "' "
   End If
   'end 2017/02/13
   
   strCatchNA16 = strCatchNA16 & ") "
   'end 2016/06/14
   
End If 'Added by Morgan 2025/1/21

strConLast = ", CP13, CP01, CP02, CP03, CP04, NA01" 'Added by Lydia 2022/09/20 加在顯示欄位的最後

'Modify By Cheng 2002/07/05
'若系統種類對照檔SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
''Add By Cheng 2002/07/01
''若由內專或內商進入, 代理人名稱抓中-->英-->日
'If (intPCaseKind = 專利 And intPWhere = 0) Or (intPCaseKind = 商標 And intPWhere = 0) Then
'   strSQLA = "NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),"
''其他代理人名稱抓英-->中-->日
'Else
'   strSQLA = "DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65),"
'End If
StrSQLa = "DECODE(SK03,0,(NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65))), " & _
         "(NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)))) AS FANAME,"
'管制別為未收達
If Val(txt1(9)) = 1 Then
   pub_QL05 = pub_QL05 & ";" & Label6 & "未收達" 'Add By Sindy 2010/9/30
   '申請國家為台灣或C類收文號不抓
   '910819 Sieg 307
   'Add by Lydia 2014/10/31 CASEPROGRESS +別名f0
   If intPCaseKind = 專利 And intPWhere = 國外_CF Then
      'Modified by Lydia 2022/09/20 +strConLast
      strSql = "SELECT ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)) AS CASENAME," & _
         "PTM03,NA03,DECODE(PA09,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,PATENT,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP46 IS NULL AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)  AND CP14=ST01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND pa09=na01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) and (PA57<>'Y' OR PA57 IS NULL) AND (CP47= 0 OR CP47 IS NULL) AND CP57 IS NULL " & StrTest1 & " And '000'<>PA09(+) And CP09<'C' "
      strSql = strSql + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)) AS CASENAME,DECODE(TM10,'000',PTM03,PTM04) AS CASETYPE,NA03,DECODE(TM10,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,TRADEMARK,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP46 IS NULL AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)  AND CP14=ST01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND '2'=ptm01(+) AND TM08=PTM02(+)  AND TM10=NA01(+)  AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (TM29<>'Y' OR TM29 IS NULL) AND (CP47=0 OR CP47 IS NULL) AND CP57 IS NULL " & StrTest2 & " And '000'<>TM10(+) And CP09<'C' "
      strSql = strSql + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07)) AS CASENAME,'',NA03,DECODE(SP09,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,SERVICEPRACTICE,CASEPROPERTYMAP,NATION,FAGENT,STAFF ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP46 IS NULL AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)  AND CP14=ST01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND sp09=na01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (SP15<>'Y' OR SP15 IS NULL) AND (CP47=0 OR CP47 IS NULL) AND CP57 IS NULL " & StrTest3 & " And '000'<>SP09(+) And CP09<'C' "
      'strSql = strSql + " ORDER BY A,B,C " 'Removed by Morgan 2025/1/21 統一改最後再加
      
    'Add By Cheng 2002/12/17
    ElseIf intPCaseKind = 商標 And intPWhere = 國內 Then
      'Modified by Lydia 2022/09/20 +strConLast
      strSql = "SELECT ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)) AS CASENAME,DECODE(PA09,'000',PTM03,PTM04) AS CASETYPE,NA03,DECODE(PA09,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,PATENT,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP46 IS NULL AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)  AND CP14=ST01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND pa09=na01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) and (PA57<>'Y' OR PA57 IS NULL) AND (CP47= 0 OR CP47 IS NULL) AND CP57 IS NULL " & StrTest1 & " And '000'<>PA09(+) And CP09<'B' "
      strSql = strSql + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)) AS CASENAME,DECODE(TM10,'000',PTM03,PTM04) AS CASETYPE,NA03,DECODE(TM10,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,TRADEMARK,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP46 IS NULL AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)  AND CP14=ST01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND '2'=ptm01(+) AND TM08=PTM02(+)  AND TM10=NA01(+)  AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (TM29<>'Y' OR TM29 IS NULL) AND (CP47=0 OR CP47 IS NULL) AND CP57 IS NULL " & StrTest2 & " And '000'<>TM10(+) And CP09<'B' "
      strSql = strSql + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07)) AS CASENAME,'',NA03,DECODE(SP09,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,SERVICEPRACTICE,CASEPROPERTYMAP,NATION,FAGENT,STAFF ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP46 IS NULL AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)  AND CP14=ST01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND sp09=na01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (SP15<>'Y' OR SP15 IS NULL) AND (CP47=0 OR CP47 IS NULL) AND CP57 IS NULL " & StrTest3 & " And '000'<>SP09(+) And CP09<'B' "
      'strSql = strSql + " ORDER BY A,B,C " 'Removed by Morgan 2025/1/21 統一改最後再加
   
   Else
      'Modified by Lydia 2022/09/20 +strConLast
      strSql = "SELECT ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)) AS CASENAME,DECODE(PA09,'000',PTM03,PTM04) AS CASETYPE,NA03,DECODE(PA09,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,PATENT,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP46 IS NULL AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)  AND CP14=ST01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND pa09=na01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) and (PA57<>'Y' OR PA57 IS NULL) AND (CP47= 0 OR CP47 IS NULL) AND CP57 IS NULL " & StrTest1 & " And '000'<>PA09(+) And CP09<'C' "
      strSql = strSql + strCatchNA16 'Added by Lydia 2016/06/14 判斷FCP管制人
      '2006/5/9 MODIFY BY SONIA CFT之B類其他706不管制
      'strSQL = strSQL + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)),DECODE(TM10,'000',PTM03,PTM04),NA03,DECODE(TM10,'000',CPM03,CPM04)," & StrSQLa & "CP46 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP46 IS NULL AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)  AND CP14=ST01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND '2'=ptm01(+) AND TM08=PTM02(+)  AND TM10=NA01(+)  AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (TM29<>'Y' OR TM29 IS NULL) AND (CP47=0 OR CP47 IS NULL) AND CP57 IS NULL " & StrTest2 & " And '000'<>TM10(+) And CP09<'C' "
      strSql = strSql + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)) AS CASENAME,DECODE(TM10,'000',PTM03,PTM04) AS CASETYPE,NA03,DECODE(TM10,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,TRADEMARK,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP46 IS NULL AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)  AND CP14=ST01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND '2'=ptm01(+) AND TM08=PTM02(+)  AND TM10=NA01(+)  AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (TM29<>'Y' OR TM29 IS NULL) AND (CP47=0 OR CP47 IS NULL) AND CP57 IS NULL " & StrTest2 & " And '000'<>TM10(+) And CP09<'C' AND NOT (CP09>'A' AND CP01='CFT' AND CP10='706') "
      '2006/5/9 END
      strSql = strSql + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07)) AS CASENAME,'',NA03,DECODE(SP09,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,SERVICEPRACTICE,CASEPROPERTYMAP,NATION,FAGENT,STAFF ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP46 IS NULL AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)  AND CP14=ST01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) AND sp09=na01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (SP15<>'Y' OR SP15 IS NULL) AND (CP47=0 OR CP47 IS NULL) AND CP57 IS NULL " & StrTest3 & " And '000'<>SP09(+) And CP09<'C' "
      'strSql = strSql + " ORDER BY A,B,C " 'Removed by Morgan 2025/1/21 統一改最後再加
   End If
'管制別為未提申
Else
   pub_QL05 = pub_QL05 & ";" & Label6 & "未提申" 'Add By Sindy 2010/9/30
   '申請國家為台灣或C類收文號不抓
   '910819 Sieg 307
   If intPCaseKind = 專利 And intPWhere = 國外_CF Then
      'Modified by Lydia 2022/09/20 +strConLast
      strSql = "SELECT ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)) AS CASENAME," & _
         "PTM03,NA03,DECODE(PA09,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,PATENT,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP47 IS NULL AND CP57 IS NULL AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)  AND CP14=ST01(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1 & " And '000'<>PA09(+) And CP09<'C' "
      strSql = strSql + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)) AS CASENAME,DECODE(TM10,'000',PTM03,PTM04) AS CASETYPE,NA03,DECODE(TM10,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,TRADEMARK,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP47 IS NULL AND CP57 IS NULL AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)  AND CP14=ST01(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (TM29<>'Y' OR TM29 IS NULL) " & StrTest2 & " And '000'<>TM10(+) And CP09<'C' "
      strSql = strSql + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07)) AS CASENAME,'',NA03,DECODE(SP09,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,SERVICEPRACTICE,CASEPROPERTYMAP,NATION,FAGENT,STAFF ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP47 IS NULL AND CP57 IS NULL AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)  AND CP14=ST01(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND SP09=NA01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (SP15<>'Y' OR SP15 IS NULL) " & StrTest3 & " And '000'<>SP09(+) And CP09<'C' "
      'strSql = strSql + " ORDER BY A,B,C " 'Removed by Morgan 2025/1/21 統一改最後再加
    'Add By Cheng 2002/12/17
    ElseIf intPCaseKind = 商標 And intPWhere = 國內 Then
      'Modified by Lydia 2022/09/20 +strConLast
      strSql = "SELECT ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)) AS CASENAME,DECODE(PA09,'000',PTM03,PTM04) AS CASETYPE,NA03,DECODE(PA09,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,PATENT,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP47 IS NULL AND CP57 IS NULL AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)  AND CP14=ST01(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1 & " And '000'<>PA09(+) And CP09<'B' "
      '2005/12/26 MODIFY BY SONIA 有期限但期限在6個月內未到期者不出現
      'strSQL = strSQL + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)),DECODE(TM10,'000',PTM03,PTM04),NA03,DECODE(TM10,'000',CPM03,CPM04)," & StrSQLa & "CP46 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP47 IS NULL AND CP57 IS NULL AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)  AND CP14=ST01(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (TM29<>'Y' OR TM29 IS NULL) " & StrTest2 & " And '000'<>TM10(+) And CP09<'B' "
      strSql = strSql + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)) AS CASENAME,DECODE(TM10,'000',PTM03,PTM04) AS CASETYPE,NA03,DECODE(TM10,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,TRADEMARK,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP47 IS NULL AND CP57 IS NULL AND (CP06 IS NULL OR (CP06 IS NOT NULL AND CP06<= " & DBDATE(DateAdd("m", 6, ChangeWStringToWDateString(ServerDate))) & ")) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)  AND CP14=ST01(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (TM29<>'Y' OR TM29 IS NULL) " & StrTest2 & " And '000'<>TM10(+) And CP09<'B' "
      '2005/12/26 END
      strSql = strSql + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07)) AS CASENAME,'',NA03,DECODE(SP09,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,SERVICEPRACTICE,CASEPROPERTYMAP,NATION,FAGENT,STAFF ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP47 IS NULL AND CP57 IS NULL AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)  AND CP14=ST01(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND SP09=NA01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (SP15<>'Y' OR SP15 IS NULL) " & StrTest3 & " And '000'<>SP09(+) And CP09<'B' "
      'strSql = strSql + " ORDER BY A,B,C " 'Removed by Morgan 2025/1/21 統一改最後再加
   
   Else
      'Modified by Lydia 2022/09/20 +strConLast
      strSql = "SELECT ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(PA05,NVL(PA06,PA07)) AS CASENAME,DECODE(PA09,'000',PTM03,PTM04) AS CASETYPE,NA03,DECODE(PA09,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,PATENT,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP47 IS NULL AND CP57 IS NULL AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)  AND CP14=ST01(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (PA57<>'Y' OR PA57 IS NULL) " & StrTest1 & " And '000'<>PA09(+) And CP09<'C' "
      strSql = strSql + strCatchNA16 'Added by Lydia 2016/06/14 判斷FCP管制人
      '2006/5/9 MODIFY BY SONIA CFT之B類其他706不管制
      'strSQL = strSQL + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)),DECODE(TM10,'000',PTM03,PTM04),NA03,DECODE(TM10,'000',CPM03,CPM04)," & StrSQLa & "CP46 FROM CASEPROGRESS,TRADEMARK,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP47 IS NULL AND CP57 IS NULL AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)  AND CP14=ST01(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (TM29<>'Y' OR TM29 IS NULL) " & StrTest2 & " And '000'<>TM10(+) And CP09<'C' "
      strSql = strSql + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)) AS CASENAME,DECODE(TM10,'000',PTM03,PTM04) AS CASETYPE,NA03,DECODE(TM10,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,TRADEMARK,CASEPROPERTYMAP,NATION,FAGENT,STAFF,PATENTTRADEMARKMAP ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP47 IS NULL AND CP57 IS NULL AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)  AND CP14=ST01(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND '2'=ptm01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (TM29<>'Y' OR TM29 IS NULL) " & StrTest2 & " And '000'<>TM10(+) And CP09<'C' AND NOT (CP09>'A' AND CP01='CFT' AND CP10='706') "
      '2006/5/9 END
      'Added by Lydia 2024/07/31 T,TF抓發文期間收文號(CP47+CP24=null)對應的下一程序提申998尚未上NP06，承辦人帶NP10
      strSql = strSql + " union all SELECT S2.ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(TM05,NVL(TM06,TM07)) AS CASENAME,DECODE(TM10,'000',PTM03,PTM04) AS CASETYPE,NA03,DECODE(TM10,'000',M2.CPM03,M2.CPM04)||'-'||DECODE(TM10,'000',M1.CPM03,M1.CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & _
             " FROM CASEPROGRESS F0,TRADEMARK,CASEPROPERTYMAP M1,NATION,FAGENT,STAFF S1,PATENTTRADEMARKMAP ,SYSTEMKIND,NEXTPROGRESS,STAFF S2,CASEPROPERTYMAP M2 WHERE CP01=SK01(+) " & _
             " AND CP09 < 'C' AND CP47 IS NULL AND CP57 IS NULL AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+)  AND CP14=S1.ST01(+) AND CP01=M1.CPM01(+) AND CP10=M1.CPM02(+) " & _
             " AND '2'=PTM01(+) AND TM08=PTM02(+) AND TM10=NA01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (TM29<>'Y' OR TM29 IS NULL) " & StrTest2 & " AND CP01 IN ('T','TF') AND TM10<>'000'  AND (CP24 IS NULL OR CP24 = '') " & _
             " AND '000'<>TM10(+) AND CP09<'C' AND NOT (CP09>'A' AND CP01='CFT' AND CP10='706') AND CP09=NP01(+) AND CP01=NP02(+) AND CP02=NP03(+) AND CP03=NP04(+) AND CP04=NP05(+) AND NP07='998' AND NP06 IS NULL AND NP10=S2.ST01(+) AND NP02=M2.CPM01(+) AND NP07=M2.CPM02(+) "
      'end 2024/07/31
      strSql = strSql + " union all select ST02 AS A,CP27 AS B,CP09,CP05,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS C,NVL(SP05,NVL(SP06,SP07)) AS CASENAME,'',NA03,DECODE(SP09,'000',CPM03,CPM04) AS CPMNAME," & StrSQLa & "CP46 " & strConLast & " FROM CASEPROGRESS f0,SERVICEPRACTICE,CASEPROPERTYMAP,NATION,FAGENT,STAFF ,SYSTEMKIND WHERE CP01=SK01(+) AND CP09 < 'C' AND CP47 IS NULL AND CP57 IS NULL AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+)  AND CP14=ST01(+) AND cp01=cpm01(+) AND CP10=CPM02(+) AND SP09=NA01(+) AND SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND (SP15<>'Y' OR SP15 IS NULL) " & StrTest3 & " And '000'<>SP09(+) And CP09<'C' "
      'strSql = strSql + " ORDER BY A,B,C " 'Removed by Morgan 2025/1/21 統一改最後再加
   End If
End If

strSql = "select X.*,'' PID from (" & strSql & ") X ORDER BY A,B,C" 'Added by Morgan 2025/1/21 +PID

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/9/30
    'Added by Lydia 2022/09/20 先放在暫存檔
    Set RsTemp = PUB_CreateRecordset(adoRecordset, , , , Me.Name, mSeqNo)
    
    'Added by Morgan 2025/1/21
    If FMP2open = False And strSrvDate(1) >= P業務區劃分啟用日 And txt1(12).Text <> "" Then
      With RsTemp
      .MoveFirst
      Do While Not .EOF
         If Left(.Fields("C"), 2) = "P-" Then
            .Fields("PID") = PUB_GetPHandler(.Fields("C"))
         End If
         .MoveNext
      Loop
      .UpdateBatch
      End With
    End If
    'end 2025/1/21
    
    strExc(0) = "select ROWSEQ,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012,R013,R014,R015,R016,R017 " & _
                     "From RDataFactory where FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno = " & mSeqNo & _
                     " order by R017,R013,R014,R015,R016"
    '依國籍+本所案號排序
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        RsTemp.MoveFirst
        Do While Not RsTemp.EOF
            'CFT、CFC、S的未收達、未提申報表，承辦人改依外商規定GetNa69
            'Modified by Lydia 2022/10/27 debug: 排除T案
            If InStr("CFT、CFC、S", "" & RsTemp.Fields("R013")) > 0 And "" & RsTemp.Fields("R013") <> "T" And Val("" & RsTemp.Fields("R017")) > 10 And "" & RsTemp.Fields("R013") <> "" And "" & RsTemp.Fields("R014") <> "" Then
                 Call GetNA69("", "" & RsTemp.Fields("R017"), "" & RsTemp.Fields("R012"), strNA69, "" & RsTemp.Fields("R013"), "" & RsTemp.Fields("R014"), "" & RsTemp.Fields("R015"), "" & RsTemp.Fields("R016"))
                 strExc(1) = GetStaffName(strNA69)
                 If strExc(1) <> "" & RsTemp.Fields("R001") Then
                     strSql = "Update  RDataFactory set R001=" & CNULL(ChgSQL(strExc(1))) & " where FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno = " & mSeqNo & " and rowseq= " & RsTemp.Fields("ROWSEQ")
                     cnnConnection.Execute strSql
                 End If
            End If
            RsTemp.MoveNext
        Loop
        '重新讀取資料
        CheckOC
        
        strSql = "select R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012,R013,R014,R015,R016,R017 " & _
                         "From RDataFactory where FormName='" & Me.Name & "' And ID='" & strUserNum & "' and seqno = " & mSeqNo
                         
        'Added by Morgan 2025/1/21
        If FMP2open = False And strSrvDate(1) >= P業務區劃分啟用日 And txt1(12).Text <> "" Then
            strSql = strSql & " and R018='" & txt1(12) & "'"
        End If
        'end 2025/1/21
                         
        strSql = strSql & " order by R001,R002,R005"
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    End If
    'end 2022/09/20
End If

If adoRecordset.RecordCount > 0 Then
    StrPrintDoc       '列印主程式
    CheckOC
    'Add By Cheng 2002/12/17
    ShowPrintOk
Else
    InsertQueryLog (0) 'Add By Sindy 2010/9/30
    ShowNoData
    CheckOC
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Screen.MousePointer = vbDefault
End Sub
Sub StrPrintDoc()

GetPrintLeft
iLine = 1
Page = 1
If Val(txt1(9)) = 1 Then
    TmpArea = "未收達"
Else
    TmpArea = "未提申"
End If
StrPrintTital TmpArea, str(Page)
iPrint = 2700
StrTest1 = ""
With adoRecordset
    .MoveFirst
    Do While .EOF = False
        For j = 0 To 10
            If Not IsNull(.Fields(j)) Then
                strTemp3(j) = .Fields(j)
            Else
                strTemp3(j) = ""
            End If
        Next j
        
        If Len(strTemp3(1)) > 7 Then
            strTemp3(1) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(1)))
        End If
        If Len(strTemp3(3)) > 7 Then
            strTemp3(3) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(3)))
        End If
        If Len(strTemp3(10)) > 7 Then
            strTemp3(10) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp3(10)))
        End If
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = iPrint
        If StrTest1 <> strTemp3(0) Then
            Printer.Print Format(StrConv(MidB(StrConv(strTemp3(0), vbFromUnicode), 1, 8), vbUnicode), "!@@@@@@@@")
            StrTest1 = strTemp3(0)
        Else
            Printer.Print ""
        End If
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(1)
        Printer.CurrentX = PLeft(2)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(3)
        Printer.CurrentX = PLeft(3)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(2)
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = iPrint
        Printer.Print strTemp3(4)
        Printer.CurrentX = PLeft(5)
        Printer.CurrentY = iPrint
        Printer.Print Format(StrConv(MidB(StrConv(strTemp3(5), vbFromUnicode), 1, 24), vbUnicode), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
        Printer.CurrentX = PLeft(6)
        Printer.CurrentY = iPrint
        Printer.Print Format(StrConv(MidB(StrConv(strTemp3(6), vbFromUnicode), 1, 8), vbUnicode), "!@@@@@@@@")
        Printer.CurrentX = PLeft(7)
        Printer.CurrentY = iPrint
        Printer.Print Format(StrConv(MidB(StrConv(strTemp3(7), vbFromUnicode), 1, 8), vbUnicode), "!@@@@@@@@")
        Printer.CurrentX = PLeft(8)
        Printer.CurrentY = iPrint
        Printer.Print Format(StrConv(MidB(StrConv(strTemp3(8), vbFromUnicode), 1, 8), vbUnicode), "!@@@@@@@@")
        Printer.CurrentX = PLeft(9)
        Printer.CurrentY = iPrint
        Printer.Print Format(StrConv(MidB(StrConv(strTemp3(9), vbFromUnicode), 1, 10), vbUnicode), "!@@@@@@@@@@")
        If Val(txt1(9)) = 2 Then
            Printer.CurrentX = PLeft(10)
            Printer.CurrentY = iPrint
            Printer.Print strTemp3(10)
        End If
        .MoveNext
        If .EOF = False Then
            If (iLine Mod 26 = 0) Then
                StrPrintEnd
                Printer.NewPage
                Page = Page + 1
                StrPrintTital TmpArea, str(Page)
                iPrint = 2400
                iLine = 0
            End If
            iLine = iLine + 1
            iPrint = iPrint + 300
        End If
    Loop
End With
StrPrintEnd
Printer.EndDoc
CheckOC


End Sub
Sub StrPrintTital(ByRef Area As String, ByRef Page As String)
GetPrintLeft
k = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5000
Printer.CurrentY = i
Printer.Print "代理人案件收達/提申管制表"
Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = 6000
Printer.CurrentY = k + 500
Printer.Print "發文日期：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "-" & ChangeTStringToTDateString(txt1(2))
Printer.Font.Bold = False
Printer.CurrentX = 500
Printer.CurrentY = k + 800
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = k + 800
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.CurrentX = 500
Printer.CurrentY = k + 1100
Printer.Print "列印別：" & Area
Printer.CurrentX = 13000
Printer.CurrentY = k + 1100
Printer.Print "頁    次：" & Page
Printer.CurrentX = 500
Printer.CurrentY = k + 1400
Printer.Print String(200, "-")
Printer.Font.Underline = True
Printer.CurrentX = PLeft(0)
Printer.CurrentY = k + 1700
Printer.Print "承辦人"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = k + 1700
Printer.Print "發文日"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = k + 1700
Printer.Print "收文日"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = k + 1700
Printer.Print "收文號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = k + 1700
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = k + 1700
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = k + 1700
Printer.Print "種  類"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = k + 1700
Printer.Print "申請國家"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = k + 1700
Printer.Print "案件性質"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = k + 1700
Printer.Print "代理人"
If Val(txt1(9)) = 2 Then
    Printer.CurrentX = PLeft(10)
    Printer.CurrentY = k + 1700
    Printer.Print "收達日"
End If
Printer.Font.Underline = False
Printer.CurrentX = 500
Printer.CurrentY = k + 2000
Printer.Print String(200, "-")
End Sub

Sub StrPrintEnd()
Printer.CurrentX = 500
Printer.CurrentY = iPrint + 300
Printer.Print String(200, "-")
End Sub

Sub GetPrintLeft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1500
PLeft(2) = 2700
PLeft(3) = 3900
PLeft(4) = 5100
PLeft(5) = 7000
PLeft(6) = 10100
PLeft(7) = 11100
PLeft(8) = 12200
PLeft(9) = 13300
PLeft(10) = 14700
End Sub

Private Sub Command2_Click(Index As Integer)
   Dim strTmp As String
   If Index = 0 And Text2(2).Text <> "" Then
      strTmp = Text2(1) & Text2(2)
      If Text2(3).Text = "" Then
         strTmp = strTmp & "0"
      Else
         strTmp = strTmp & Text2(3).Text
      End If
      If Text2(4).Text = "" Then
         strTmp = strTmp & "00"
      Else
         strTmp = strTmp & Text2(4).Text
      End If
      intI = 1
      strExc(0) = "SELECT 1 FROM PATENT WHERE " & ChgPatent(strTmp)
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         List1(1).AddItem strTmp
         Text2(2).Text = ""
      Else
         MsgBox "案號不存在，請重新輸入 !", vbCritical
      End If
      Text2(2).SetFocus
   Else
      If List1(1).ListIndex > -1 Then List1(1).RemoveItem List1(1).ListIndex
   End If
End Sub

Private Sub Form_Load()
MoveFormToCenter Me

'Added by Lydia 2016/06/14 限外專使用
If Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 1) = "F" Then
   Label13.Visible = True: Label14.Visible = True
   txt1(12).Visible = True
   If Pub_StrUserSt03 <> "M51" Then
      txt1(12).Text = strUserNum
      Label14.Caption = strUserName
   End If
Else
   Label13.Visible = False: Label14.Visible = False
   txt1(12).Visible = False
End If
'end 2016/06/14

'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
'Modified by Lydia 2014/12/29 非FMP寰華權限,不可看寰華案=>回傳SQL
'FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05, "INVERSE_SQL")
If Pub_StrUserSt03 = "M51" Then
   If MsgBox("電腦中心人員請注意你現在是要看FMP寰華案嗎?", vbYesNo + vbInformation + vbDefaultButton1) = vbYes Then
      FMP2openSQL = Replace(FMP2openSQL, "not", "")
      FMP2open = True 'Added by Lydia 2016/06/14
   Else
      MsgBox "現在本報表不可查FMP寰華案"
      'Added by Lydia 2016/06/14
      Label13.Visible = False: Label14.Visible = False
      txt1(12).Visible = False
   End If
End If

mOpenSql = FMP2openSQL 'Added by Lydia 2015/10/26

If FMP2open = True Then
   txt1(0) = "P,PS,"
Else
   txt1(0) = UCase(GetSystemKindByNick)
End If

'Added by Morgan 2025/1/21
If FMP2open = False And strSrvDate(1) >= P業務區劃分啟用日 Then
   Label13.Visible = True: Label14.Visible = True
   txt1(12).Visible = True
   If Pub_StrUserSt03 <> "M51" Then
      txt1(12).Text = strUserNum
      Label14.Caption = strUserName
   End If
End If
'end 2025/1/15
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050303 = Nothing
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_Change(Index As Integer)
   If Index = 9 Then List1(1).Clear 'Added by Morgan 2016/5/17
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txt1(Index).IMEMode = 2
   CloseIme
   'Add by Morgan 2006/5/22
   If Index = 6 Or Index = 8 Then
      txt1(Index) = txt1(Index - 1)
   End If
   'end 2006/5/22
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   'Add by Morgan 2006/5/19
   If Index = 9 Then
      'Modified by Lydia 2015/09/09 +5.公開管制函
      'If KeyAscii <> 8 And (KeyAscii < Asc("1") Or KeyAscii > Asc("4")) Then
      If KeyAscii <> 8 And (KeyAscii < Asc("1") Or KeyAscii > Asc("5")) Then
         Beep
         KeyAscii = 0
         Exit Sub
      'ElseIf (KeyAscii = Asc("3") Or KeyAscii = Asc("4")) Then
      ElseIf (KeyAscii = Asc("3") Or KeyAscii = Asc("4") Or KeyAscii = Asc("5")) Then
         'Added by Morgan 2024/3/14
         If InStr(txt1(0), "P") = 0 Then
            Beep
            KeyAscii = 0
            MsgBox "非專利案無此管制別！", vbCritical
            Exit Sub
         End If
         'end 2024/3/14
      
         txt1(1) = "": txt1(1).Enabled = False
         txt1(2) = "": txt1(2).Enabled = False
         txt1(3) = "": txt1(3).Enabled = False
         txt1(4) = "": txt1(4).Enabled = False
         txt1(10).Enabled = True
         If txt1(10) = "" Then txt1(10) = strSrvDate(2)
         
         'Added by Morgan 2016/5/23
         If KeyAscii <> Asc("5") Then
            txt1(11).Visible = True: Line5.Visible = True
            If txt1(11) = "" Then txt1(11) = txt1(10)
         Else
            txt1(11) = "": txt1(11).Visible = False: Line5.Visible = False
         End If
         'end 2016/5/23
         
         Command2(0).Enabled = True 'Added by Morgan 2016/5/17
      Else
         txt1(1).Enabled = True
         txt1(2).Enabled = True
         txt1(3).Enabled = True
         txt1(4).Enabled = True
         txt1(10) = "": txt1(10).Enabled = False
         txt1(11) = "": txt1(11).Visible = False: Line5.Visible = False 'Added by Morgan 2016/5/23
         'Added by Morgan 2016/5/17
         List1(1).Clear
         Command2(0).Enabled = False
         'end 2016/5/17
      End If
   End If
   'end 2006/5/19
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
     If FMP2open = True Then
        strTemp1 = Split("P,PS,", ",")
     Else
        strTemp1 = Split(UCase(GetSystemKindByNick), ",")
     End If

     strTemp2 = Split(UCase(txt1(0)), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
    Next i
Case 7
   If Len(txt1(Index)) = 6 Then
      txt1(Index) = txt1(Index) & "000"
   End If
   
Case 8
     'If Trim(txt1(7)) <> "" And Trim(txt1(8)) <> "" Then
        If Left(txt1(7), 6) <> Left(txt1(8), 6) Then
            s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
            txt1(7).SetFocus
            txt1_GotFocus (7)
            Exit Sub
        End If
      'End If
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
      If Len(txt1(Index)) = 6 Then
         txt1(Index) = txt1(Index) & "000"
      End If
Case 9
     Select Case Val(txt1(9))
     Case 1, 2
      'Add By Cheng 2002/05/16
      '若管制別為1未收達時, 發文日止日為民國年之系統日-14天
      If Val(Me.txt1(9).Text) = 1 Then
         Me.txt1(2).Text = ChangeWDateStringToTString(DateAdd("d", -14, ChangeWStringToWDateString(ServerDate)))
         Me.txt1(1).Text = ""
'      '若管制別為2未提申時, 發文日止日為民國年之系統日-3個月
      ElseIf Val(Me.txt1(9).Text) = 2 Then
'Modify by Morgan 2005/7/11 CFP未提申止日預設系統日-15日,起日預設止日-3月
'         Me.txt1(2).Text = ChangeWDateStringToTString(DateAdd("m", -3, ChangeWStringToWDateString(ServerDate)))

         'Modify by Morgan 2005/8/23 加判斷類別
         If InStr(txt1(0), "CFP") > 0 Then
            Me.txt1(2).Text = TransDate(CompDate(2, -15, strSrvDate(1)), 1)
            Me.txt1(1).Text = TransDate(CompDate(1, -3, txt1(2)), 1)
         Else
            '2007/12/7 MODIFY BY SONIA 有CFT或CFC或S者,止日預設為系統日-2個月又15天
            'Me.txt1(2).Text = ChangeWDateStringToTString(DateAdd("m", -3, ChangeWStringToWDateString(ServerDate)))
            If InStr(txt1(0), "CFT") > 0 Or InStr(txt1(0), "CFC") > 0 Or InStr(txt1(0), "S") > 0 Then
               'Modified by Lydia 2024/07/31 CFT、S案件各國各項程序統一設定管制天數為30天，TF案管制天數不變動
               'Me.txt1(2).Text = ChangeWDateStringToTString(DateAdd("d", -15, DateAdd("m", -2, ChangeWStringToWDateString(ServerDate))))
               Me.txt1(2).Text = ChangeWDateStringToTString(DateAdd("d", -30, ChangeWStringToWDateString(ServerDate)))
            Else
               Me.txt1(2).Text = ChangeWDateStringToTString(DateAdd("m", -3, ChangeWStringToWDateString(ServerDate)))
            End If
            '2007/12/7 end
            Me.txt1(1).Text = ""
         End If
         
      Else
         Me.txt1(2).Text = ""
      End If
     'Add by Morgan 2006/5/19
     'Added by Lydia 2015/09/09 +5
     Case 3, 4, 5
      
     Case Else
         If txt1(9) <> "" Then 'Added by Morgan 2012/9/16 有輸才檢查否則前一跳離的欄位若有錯則會造成循環
            'Added by Lydia 2015/09/09 +5
            s = MsgBox("管制別只能輸入 1, 2, 3, 4 或 5 !!", , "USER 輸入錯誤")
            txt1(9).SetFocus
            txt1(9).SelStart = 0
            txt1(9).SelLength = Len(txt1(9))
            Exit Sub
         End If
     End Select
Case 1, 2, 10, 11
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 2 Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   End If
Case 4, 6
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
'Added by Lydia 2016/06/14
Case 12
      Label14.Caption = ""
      If txt1(Index).Text <> "" Then
         Label14.Caption = GetStaffName(txt1(Index), True)
         If Label14.Caption = "" Then
            txt1(Index).SetFocus
            txt1_GotFocus Index
            Exit Sub
         End If
      End If
      
End Select

End Sub

'Added by Lydia 2015/09/09
'公開管制函
Private Sub StrMenu3()
   Dim strCon As String, stVTB As String, iSNo As Integer
   Dim strAgent As String, bolEdit As Boolean
   Dim strDate As String, StrDate2 As String
   Dim strCP09 As String, strCP12 As String, strCP13 As String, strCP64 As String, bolInTrans As Boolean 'Added by Morgan 2016/5/13
   Dim strCatchNA16 As String 'Added by Lydia 2016/06/14
   
   strCon = ""
   
   '系統類別
   'Added by Morgan 2024/3/14
   If Len(txt1(0)) <> 0 Then
      strCon = strCon & " AND C1.CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0)
   End If
   'end 2024/3/14
   
   If txt1(5) <> "" Then
      strCon = strCon & " And PA09||''>='" & txt1(5) & "'"
   End If
   If txt1(6) <> "" Then
      strCon = strCon & " And PA09||''<='" & txt1(6) & "'"
   End If
   If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label4 & txt1(5) & "-" & txt1(6)
   End If
   
   'Added by Morgan 2016/5/17
   If List1(1).ListCount > 0 Then
      pub_QL05 = pub_QL05 & ";不管制案號" & List1(1).List(0)
      strCon = strCon & " and pa01||pa02||pa03||pa04 not in ('" & List1(1).List(0) & "'"
      For intI = 1 To List1(1).ListCount - 1
         strCon = strCon & ",'" & List1(1).List(intI) & "'"
         pub_QL05 = pub_QL05 & "," & List1(1).List(intI)
      Next
      strCon = strCon & ")"
   End If
   'end 2016/5/17
   
   
   'Modified by Lydia 2016/1/13 因為大陸發明案輸入申請案號時,下一程序公開期限999掛通知申請案號的進度(NP01=通知申請案號之CP09),這道案件進度可能沒掛代理人
   '                            所以改抓通知申請案號(np01)之相關總收文號(cp43)的代理人做判斷
'   If txt1(7) <> "" Then
'      strCon = strCon & " And C1.CP44||''>='" & txt1(7) & "'"
'   Else
'      strCon = strCon & " And C1.CP44||''>='Y'"
'   End If
'   If txt1(8) <> "" Then
'      strCon = strCon & " And C1.CP44||''<='" & txt1(8) & "'"
'      If txt1(7) <> "" Then
'         bolEdit = True
'      End If
'   End If
   If txt1(7) <> "" Then
      strCon = strCon & " AND NVL(C2.CP44,C1.CP44)||''>='" & txt1(7) & "'"
   Else
      strCon = strCon & " AND NVL(C2.CP44,C1.CP44)||''>='Y'"
   End If
   If txt1(8) <> "" Then
      strCon = strCon & " AND NVL(C2.CP44,C1.CP44)||''<='" & txt1(8) & "'"
      If txt1(7) <> "" Then
         bolEdit = True
      End If
   End If
   'end 2016/1/13
   
   If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label5 & txt1(7) & "-" & txt1(8)
   End If
   If txt1(10) <> "" Then
      strDate = TransDate(txt1(10), 2)
   End If
   
   If FMP2open Then 'Added by Morgan 2025/1/21 開放內專也可輸管制人且規則不同
   
      'Added by Lydia 2016/06/14 判斷FCP管制人 strCatchNa16
      strCatchNA16 = " and (C1.cp01,C1.cp02,C1.cp03,C1.cp04) in (SELECT x1.PA01,x1.PA02,x1.PA03,x1.PA04 FROM PATENT x1,FAGENT x2,NATION x3 WHERE x1.PA01=C1.CP01 and x1.PA02=C1.CP02 and x1.PA03=C1.CP03 and x1.PA04=C1.CP04 and SUBSTR(x1.PA75,1,8)=x2.FA01(+) AND SUBSTR(x1.PA75,9,1)=x2.FA02(+) AND x2.FA10=x3.NA01(+) "
      'Modified by Lydia 2017/02/13 +FMP管制人
      If strSrvDate(1) < FMP管制人啟用日 Then
          If txt1(12).Text <> "" Then strCatchNA16 = strCatchNA16 & "and x3.na16='" & Trim(txt1(12).Text) & "' "
      Else
          If txt1(12).Text <> "" Then strCatchNA16 = strCatchNA16 & "and decode(x1.PA01,'P',nvl(x3.na79,x3.na16),x3.na16)='" & Trim(txt1(12).Text) & "' "
      End If
      'end 2017/02/13
      
      strCatchNA16 = strCatchNA16 & ") "
      'end 2016/06/14
      
   End If 'Added by Morgan 2025/1/21
      
   'Added by Morgan 2020/1/15
   '寰華案排除尚未收到初步審查合格通知書之案件
   If FMP2open = True Then
     strCon = strCon & " and exists(select * from caseprogress x where x.cp01=pa01 and x.cp02=pa02 and x.cp03=pa03 and x.cp04=pa04 and x.cp10='1213') "
   End If
   'end 2020/1/15
   
On Error GoTo ErrHandle

   strCon = strCon & Replace(FMP2openSQL, "f0", "C1")
   'Modified by Lydia 2016/1/13  改抓np01之cp43的CP44判斷
   'strSql = "SELECT c1.CP44 C00,c1.CP09 C01,c1.CP45 C02,c1.CP01||'-'||c1.CP02||DECODE(c1.CP03||c1.CP04,'000',NULL,'-'||c1.CP03||'-'||c1.CP04) C03," & _
            "NVL(PA05,NVL(PA06,PA07)) C04,cpm04 C05,PA11 C06 from caseprogress c1,patent,nextprogress,casepropertymap" & _
            " where np02='P' and np06 is null and np07='999' and np09<=" & strDate & _
            " and c1.CP09(+)=np01 and c1.cp57 is null and c1.CP01=pa01(+) and c1.CP02=pa02(+) and c1.CP03=pa03(+) and c1.CP04=pa04(+) and pa09='020' and pa08='1'and pa23='1'" & _
            " and pa16||pa57||pa12||pa21||pa108||pa136 is null and c1.CP01=cpm01(+) and c1.CP10=cpm02(+) " & strCon
   'Modified by Morgan 2016/5/17 +指示信電子化相關欄位
   'Modified by Lydia 2016/06/14 判斷FCP管制人
   'Modified by Morgan 2025/1/21 +PID
   strSql = "SELECT NVL(C2.CP44,C1.CP44) C00,NVL(C2.CP09,C1.CP09) C01,NVL(C2.CP45,C1.CP45) C02,C1.CP01||'-'||C1.CP02||DECODE(C1.CP03||C1.CP04,'000',NULL,'-'||C1.CP03||'-'||C1.CP04) C03," & _
            "NVL(PA05,NVL(PA06,PA07)) C04,NVL(D2.CPM04,D1.CPM04) C05,PA11 C06,C1.CP01,C1.CP02,C1.CP03,C1.CP04,'' PID" & _
            " FROM CASEPROGRESS C1,PATENT,NEXTPROGRESS,CASEPROPERTYMAP D1,CASEPROGRESS C2,CASEPROPERTYMAP D2 " & _
            "WHERE NP02='P' AND NP06 IS NULL AND NP07='999' and np09<=" & strDate & " AND C1.CP09(+)=NP01 AND C1.CP43=C2.CP09(+) AND C1.CP57 IS NULL " & _
            "AND C1.CP01=PA01(+) AND C1.CP02=PA02(+) AND C1.CP03=PA03(+) AND C1.CP04=PA04(+) AND PA09='020' AND PA08='1'AND PA23='1' " & _
            "AND PA16||PA57||PA12||PA21||PA108||PA136 IS NULL AND C1.CP01=D1.CPM01(+) AND C1.CP10=D1.CPM02(+) AND C2.CP01=D2.CPM01(+) AND C2.CP10=D2.CPM02(+) " & strCon & strCatchNA16
   strSql = strSql & "order by 1, 4"
      
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   
   'Added by Morgan 2025/1/21
   If FMP2open = False And adoRecordset.RecordCount > 0 And strSrvDate(1) >= P業務區劃分啟用日 And txt1(12).Text <> "" Then
      Set RsTemp = PUB_CreateRecordset(adoRecordset, , , 300, Me.Name, mSeqNo)
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            If Left(.Fields("C03"), 2) = "P-" Then
               .Fields("PID") = PUB_GetPHandler(.Fields("C03"))
            End If
            .MoveNext
         Loop
         .UpdateBatch
         
         stVTBX = "select R001 as " & .Fields(0).Name
         For intI = 2 To .Fields.Count
            stVTBX = stVTBX & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
         Next
         stVTBX = stVTBX & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
      End With
      strSql = "Select X.* From (" & stVTBX & ") X where PID='" & txt1(12) & "' order by 1,4"
      intI = 1
      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   End If
   'end 2025/1/21
         
   With adoRecordset
      If .RecordCount > 0 Then
         InsertQueryLog (.RecordCount)
         iSNo = 0
         strAgent = .Fields(0)
         m_strET02 = "" & .Fields(1)
         Do While Not .EOF
            If strAgent <> "" & .Fields(0) Then
               Export2Word "18", "04", bolEdit
               iSNo = 0
               strAgent = .Fields(0)
               m_strET02 = "" & .Fields(1)
            End If
            
            'Added by Morgan 2016/5/17
            '指示信電子化
            'If Left(Pub_StrUserSt03, 1) <> "F" Then 'Removed by Morgan 2017/8/18 寰華案也要新增954(催公開)並於非第1案的進度備註加註
               cnnConnection.BeginTrans
               bolInTrans = True
               '新增"催公開"進度
               m_strCP10 = "954"
               strCP09 = AutoNo("B", 6)
               If iSNo = 0 Then
                  m_strCP01 = .Fields("CP01")
                  m_strCP02 = .Fields("CP02")
                  m_strCP03 = .Fields("CP03")
                  m_strCP04 = .Fields("CP04")
                  m_strCP09 = strCP09
                  'Modified by Morgan 2017/8/18 寰華案也要新增954(催公開)並於非第1案的進度備註加註
                  If Left(Pub_StrUserSt03, 1) = "F" Then
                     strCP64 = "寄件備份存於" & .Fields("C03") & "案(" & m_strCP09 & ")卷宗區"
                  Else
                     strCP64 = "指示信存於" & .Fields("C03") & "案(" & m_strCP09 & ")卷宗區"
                  End If
                  'end 2017/8/18
               End If
               strCP13 = PUB_GetAKindSalesNo(.Fields("CP01"), .Fields("CP02"), .Fields("CP03"), .Fields("CP04"))
               strCP12 = GetSalesArea(strCP13)
               strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
                  "CP12,CP13,CP14,CP20,CP26,CP28,CP32,CP43,CP44,CP45,CP64) VALUES ('" & .Fields("CP01") & "','" & .Fields("CP02") & "'" & _
                  ",'" & .Fields("CP03") & "','" & .Fields("CP04") & "'," & strSrvDate(1) & ",'" & strCP09 & "','" & m_strCP10 & "','" & strCP12 & "'" & _
                  ",'" & strCP13 & "','" & strUserNum & "','N','N','" & m_strCP09 & "','N','" & .Fields("C01") & "','" & .Fields("C00") & "','" & .Fields("C02") & "','" & IIf(iSNo = 0, "", strCP64) & "')"
               cnnConnection.Execute strSql, intI
               
            If Left(Pub_StrUserSt03, 1) <> "F" Then 'Added by Morgan 2017/8/29 寰華案除外
               If iSNo = 0 Then
                  m_Subject = "請代為查詢公佈通知書是否發出，Y/R：" & IIf(Trim("" & .Fields("C02")) = "", "(請提供)", "" & .Fields("C02")) & "，O/R：" & .Fields("C03") & "，謝謝。"
                  'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
                  strExc(2) = PUB_GetLetterJudgeNew("2", "P", m_strCP10, "020")
                  PUB_AddAppForm m_strCP09, True, strExc(2), m_Subject '自行判發,不轉檔(一定要先看過)
               Else
                  'Modified by Morgan 2016/6/22--蕭茹曣
                  'm_Subject = "請代為查詢公佈通知書是否發出，謝謝。"
                  If iSNo = 1 Then
                     m_Subject = Replace(m_Subject, "，O/R：", "等...案，O/R：")
                     m_Subject = Replace(m_Subject, "，謝謝。", "等...案，謝謝。")
                  End If
                  'end 2016/6/22
                  strSql = "update appform set af13='" & m_Subject & "' where af01='" & m_strCP09 & "'"
                  cnnConnection.Execute strSql, intI
               End If
            End If
            
               cnnConnection.CommitTrans
               bolInTrans = False
               
            'End If 'Removed by Morgan 2017/8/18 寰華案也要新增954(催公開)並於非第1案的進度備註加註
            'end 2016/5/17
               
            iSNo = iSNo + 1
            If iSNo = 1 Then
               m_strCaseNo1 = "" & .Fields(2) 'C02=>貴方卷號
               m_strCaseNo2 = "" & .Fields(3) & "(" & .Fields(5) & ")" '本所案號(案件性質)
               m_strCaseName = Left("" & .Fields(4) & Space(40), 40) '案件名稱
               m_strCaseNo3 = "" & .Fields(6) '申請案號
               If IsNull(.Fields(6)) Then
                  m_bolShowNo3 = False
               Else
                  m_bolShowNo3 = True
               End If
            ElseIf iSNo = 2 Then
               m_strCaseNo1 = "1. " & m_strCaseNo1 & vbCrLf & String(5, "　") & "2. " & .Fields(2)
               m_strCaseNo2 = "1. " & m_strCaseNo2 & vbCrLf & String(5, "　") & "2. " & .Fields(3) & "(" & .Fields(5) & ")"
               m_strCaseName = "1. " & m_strCaseName & vbCrLf & String(5, "　") & "2. " & Left("" & .Fields(4) & Space(40), 40)
               m_strCaseNo3 = "1. " & m_strCaseNo3 & vbCrLf & String(5, "　") & "2. " & .Fields(6)
               If Not IsNull(.Fields(6)) Then m_bolShowNo3 = True
            Else
               m_strCaseNo1 = m_strCaseNo1 & vbCrLf & String(5, "　") & iSNo & ". " & .Fields(2)
               m_strCaseNo2 = m_strCaseNo2 & vbCrLf & String(5, "　") & iSNo & ". " & .Fields(3) & "(" & .Fields(5) & ")"
               m_strCaseName = m_strCaseName & vbCrLf & String(5, "　") & iSNo & ". " & Left("" & .Fields(4) & Space(40), 40)
               m_strCaseNo3 = m_strCaseNo3 & vbCrLf & String(5, "　") & iSNo & ". " & .Fields(6)
               If Not IsNull(.Fields(6)) Then m_bolShowNo3 = True

            End If
            .MoveNext
         Loop
         
         Export2Word "18", "04", bolEdit

          MsgBox "列印完畢！"
      Else
         InsertQueryLog (0)
         MsgBox "無符合資料！"
      End If
   End With
   
ErrHandle:
   If bolInTrans Then cnnConnection.RollbackTrans 'Added by Morgan 2016/5/17
   If Err.Number <> 0 Then MsgBox Err.Description

End Sub

