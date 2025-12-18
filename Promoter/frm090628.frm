VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090628 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件逾期及異常查詢"
   ClientHeight    =   5715
   ClientLeft      =   3180
   ClientTop       =   2205
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4815
      Left            =   0
      TabIndex        =   5
      Top             =   900
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   1
      Left            =   1995
      MaxLength       =   7
      TabIndex        =   1
      Top             =   585
      Width           =   810
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   0
      Left            =   1110
      MaxLength       =   7
      TabIndex        =   0
      Top             =   585
      Width           =   810
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7260
      TabIndex        =   2
      Top             =   30
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   8040
      TabIndex        =   3
      Top             =   30
      Width           =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "註:已發文未輸會稿完成日查詢"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   270
      TabIndex        =   6
      Top             =   150
      Width           =   2685
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   2490
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   630
      Width           =   720
   End
End
Attribute VB_Name = "frm090628"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Private Sub SetDataListWidth()
'add by nickc  2006/02/10
'edit by nickc 2006/03/07 加兩個欄位
'grdDataList.Cols = 6
grdDataList.Cols = 8
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "承辦人"
grdDataList.ColWidth(0) = 650
grdDataList.CellAlignment = flexAlignCenterCenter
'edit by nickc 2006/03/07 加兩個欄位
'grdDataList.col = 1: grdDataList.Text = "發文日"
'grdDataList.ColWidth(1) = 800
'grdDataList.CellAlignment = flexAlignCenterCenter
'grdDataList.col = 2: grdDataList.Text = "本所案號"
'grdDataList.ColWidth(2) = 1400
'grdDataList.CellAlignment = flexAlignCenterCenter
'grdDataList.col = 3: grdDataList.Text = "案件名稱"
'grdDataList.ColWidth(3) = 2000
'grdDataList.CellAlignment = flexAlignCenterCenter
'grdDataList.col = 4: grdDataList.Text = "案件性質"
'grdDataList.ColWidth(4) = 1100
'grdDataList.CellAlignment = flexAlignCenterCenter
''add by nickc 2006/02/10
'grdDataList.col = 5: grdDataList.Text = "來  源"
'grdDataList.ColWidth(5) = 2200
'grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "收文日"
grdDataList.ColWidth(1) = 700
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "齊備日"
grdDataList.ColWidth(2) = 700
grdDataList.CellAlignment = flexAlignCenterCenter
'edit by nickc 2006/09/14
'grdDataList.col = 3: grdDataList.Text = "發文日"
grdDataList.col = 3: grdDataList.Text = "發文"
grdDataList.ColWidth(3) = 300
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "本所案號"
grdDataList.ColWidth(4) = 1400
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 5: grdDataList.Text = "案件名稱"
grdDataList.ColWidth(5) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 6: grdDataList.Text = "案件性質"
grdDataList.ColWidth(6) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 7: grdDataList.Text = "來  源"
grdDataList.ColWidth(7) = 3000
grdDataList.CellAlignment = flexAlignCenterCenter

End Sub


Private Sub cmdOK_Click(Index As Integer)
Dim CheckOk As Boolean
Select Case Index
Case 0
        '發文日是一定要輸入的
        If Trim(txt1(0).Text) = "" Or Trim(txt1(1).Text) = "" Then
            MsgBox "發文日一定要輸入！", , "警告！"
            If Trim(txt1(0).Text) = "" Then
                'edit by nickc 2006/03/07
                'txt1_GotFocus (0)
                txt1(0).SetFocus
                Exit Sub
            Else
                'edit by nickc 2006/03/07
                'txt1_GotFocus (1)
                txt1(1).SetFocus
                Exit Sub
            End If
        End If
        CheckOk = False
        txt1_Validate 0, CheckOk
        If CheckOk = True Then
            Exit Sub
        End If
        txt1_Validate 1, CheckOk
        If CheckOk = True Then
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        'add by nick 2005/01/03
        Me.grdDataList.MousePointer = flexHourglass
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/20 清除查詢印表記錄檔欄位
        StrMenu
        'add by nick 2005/01/03
        Me.grdDataList.MousePointer = flexDefault
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
Case Else
End Select
End Sub

Sub StrMenu()
Dim strSql As String
'add by nickc 2006/02/10
Dim p_Date As String
Dim CFP_date As String
'add by nickc 2006/10/11
Dim p_Date2 As String
Dim CFP_date2 As String

   'edit by nickc 2006/09/05 改成超過4天，就是第 5 天
   'P_date = CompWorkDay(2, strSrvDate(1), 1)
   'edit by nickc 2006/10/11 與協理柄佑開過會，確認工作天算法
   'P_date = CompWorkDay(5, strSrvDate(1), 1)
   p_Date = CompWorkDay(3, strSrvDate(1), 1)
   'edit by nickc 2006/07/28 協理有發 mail  來要修改
   'CFP_date = CompWorkDay(2, strSrvDate(1), 1)
   'edit by nickc 2006/09/05 改成超過 6 天就是第 7 天
   'CFP_date = CompWorkDay(3, strSrvDate(1), 1)
   'edit by nickc 2006/10/11 與協理柄佑開過會，確認工作天算法
   'CFP_date = CompWorkDay(7, strSrvDate(1), 1)
   CFP_date = CompWorkDay(4, strSrvDate(1), 1)
   'add by nickc 2006/10/11
   p_Date2 = CompWorkDay(4, strSrvDate(1), 1)
   CFP_date2 = CompWorkDay(6, strSrvDate(1), 1)
   
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) & "-" & txt1(1) 'Add By Sindy 2010/12/20
   
   '**************************************************************************
   '**************************************************************************
   '**************************************************************************
   '下面程式再加新的案件性質時，要一併動 Autobatchday 的 strmenu8 內的案件性質
   '下面程式再加新的案件性質時，要一併動 Autobatchday 的 strmenu8 內的案件性質
   '下面程式再加新的案件性質時，要一併動 Autobatchday 的 strmenu8 內的案件性質
   '**************************************************************************
   '**************************************************************************
   '**************************************************************************
   
'因為發文日是一定要輸入的，
'edit by nick 2004/09/02 王協理增加條件
'StrSQL = "select st02," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),nvl(cpm03,cpm04) from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01 in ('P','CFP') and cp27>=" & Val(DBDATE(Txt1(0).Text)) & " and cp27<=" & Val(DBDATE(Txt1(1).Text)) & " and cp09=ep02(+) and ep08 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  order by cp14 "
'edit by nick 2004/12/07 加條件 2005/12/2再加413,429 2005/12/27再加505,506 2006/5/30再加915
'StrSql = "select st02," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),nvl(cpm03,cpm04) from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01 in ('P','CFP') AND CP10<>'1101' and cp27>=" & Val(DBDATE(Txt1(0).Text)) & " and cp27<=" & Val(DBDATE(Txt1(1).Text)) & " and cp09=ep02(+) and ep08 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and st03<>'P12' and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C'))    order by cp14 "
'edit by nickc 2006/03/07 加收文日、齊備日
'strSQL = "select st02," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'已發文未輸會稿完成',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01 in ('P','CFP') AND CP10<>'1101' and cp27>=" & Val(DBDATE(Txt1(0).Text)) & " and cp27<=" & Val(DBDATE(Txt1(1).Text)) & " and cp09=ep02(+) and ep08 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506') "
'edit by nickc 2006/09/14 發文日改已發文
'strSQL = "select st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'已發文未輸會稿完成',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01 in ('P','CFP') AND CP10<>'1101' and cp27>=" & Val(DBDATE(txt1(0).Text)) & " and cp27<=" & Val(DBDATE(txt1(1).Text)) & " and cp09=ep02(+) and ep08 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915') "
'Modify by Morgan 2007/8/31 加807
'strSQL = "select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'已發文未輸會稿完成',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01 in ('P','CFP') AND CP10<>'1101' and cp27>=" & Val(DBDATE(txt1(0).Text)) & " and cp27<=" & Val(DBDATE(txt1(1).Text)) & " and cp09=ep02(+) and ep08 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915') "
'Modify by Morgan 2008/10/30 加指定索引(因為有用Union所以每一句都要指定才會有效)
'2009/4/23 modify by sonia 加案件性質922檢還樣品證據P-085588
'2009/7/1 MODIFY BY SONIA 加案件性質933覆函P-083558
'2010/1/6 modify by sonia 加938超頁費,939超項費

'Modify by Morgan 2010/7/20 改寫暫存以便過濾資料
'strSql = "select /*+index(caseprogress IDXCP275710)*/ st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'已發文未輸會稿完成',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01 in ('P','CFP') AND CP10<>'1101' and cp27>=" & Val(DBDATE(Txt1(0).Text)) & " and cp27<=" & Val(DBDATE(Txt1(1).Text)) & " and cp09=ep02(+) and ep08 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','938','939','922','933','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915','807') "
strSql = "select /*+index(caseprogress IDXCP275710)*/ '" & strUserNum & "',cp09,'1' from caseprogress,engineerprogress,staff" & _
   " where cp01 in ('P','CFP') AND CP10<>'1101' and cp27>=" & Val(DBDATE(txt1(0).Text)) & " and cp27<=" & Val(DBDATE(txt1(1).Text)) & _
   " and cp09=ep02(+) and ep08 is null and cp14=st01(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51')" & _
   " and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C'))" & _
   " and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909'" & _
   ",'910','911','916','917','938','939','922','933','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206'" & _
   ",'401','413','429','505','506','915','807') "
'end 2010/7/20

'add by nickc 2006/02/10
'edit by nickc 2006/03/07 加入未輸完稿日 才出來，加收文日、齊備日
'strSQL = strSQL & " union select st02," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾預定會稿日未輸會稿日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01 in ('P','CFP') AND CP10<>'1101' and cp112='Y' and ep28<=" & strSrvDate(1) & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and ep08 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506') "
'strSQL = strSQL & " union select st02," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾承辦期限之次日未輸會稿日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01='P' AND CP10<>'1101' and cp112='Y' and cp48<=" & P_date & " and cp27 is null and cp57 is null and cp09=ep02(+) and ep08 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506')  "
'strSQL = strSQL & " union select st02," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾承辦期限第二日未輸會稿日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01='CFP' AND CP10<>'1101' and cp112='Y' and cp48<=" & CFP_date & " and cp27 is null and cp57 is null and cp09=ep02(+) and ep08 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506') order by cp14 "
'edit by nickc 2006/09/14 發文日改已發文
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾預定會稿日未輸會稿日及未完稿',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01 in ('P','CFP') AND CP10<>'1101' and cp112='Y' and ep28<" & strSrvDate(1) & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915') and ep09 is null "
'Modify by Morgan 2007/8/31 加807
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾預定會稿日未輸會稿日及未完稿',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01 in ('P','CFP') AND CP10<>'1101' and cp112='Y' and ep28<" & strSrvDate(1) & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915') and ep09 is null "
'Modify by Morgan 2008/10/29 多的條件要拿掉否則乙規則的性質會抓不到,順便調語法
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾預定會稿日未輸會稿日及未完稿',cp14 from caseprogress c1,engineerprogress e1,patent,casepropertymap,staff where cp01 in ('P','CFP') AND CP10<>'1101' and cp112='Y' and ep28<" & strSrvDate(1) & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915','807') and ep09 is null "
'Modify by Morgan 2008/12/3 不必再控制適用會稿加乘註記
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾預定會稿日未輸會稿日及未完稿',cp14 from engineerprogress e1,caseprogress c1,patent,casepropertymap,staff where ep28<" & strSrvDate(1) & " and ep07 is null  and ep09 is null and cp09(+)=ep02 and cp01 in ('P','CFP') AND cp112='Y' and cp27 is null and cp57 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51')"

'Modify by Morgan 2010/7/20 改寫暫存以便過濾資料
'strSql = strSql & " union select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾預定會稿日未輸會稿日及未完稿',cp14 from engineerprogress e1,caseprogress c1,patent,casepropertymap,staff where ep28<" & strSrvDate(1) & " and ep07 is null  and ep09 is null and cp09(+)=ep02 and cp01 in ('P','CFP') and cp27 is null and cp57 is null and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51')"
''Add by Morgan 2008/9/22 若為多國子案且母案已輸入會稿日則不顯示
'strSql = strSql & " and NOT (cp01='CFP' and cp21='Y' and instr('" & CaseMapOut & "',cp10)>0 and exists(select * from caserelation,caseprogress c2,engineerprogress e2" & _
' " Where cr01 = C1.CP01 And cr02 = C1.cp02 And cr03 = C1.cp03 And cr04 = C1.cp04 and c2.cp01(+)=cr05 and c2.cp02(+)=cr06 and c2.cp03(+)=cr07 and c2.cp04(+)=cr08" & _
' " and c2.cp14=c1.cp14 and c2.cp21 is null and instr('" & CaseMapOut & "',c2.cp10)>0 and e2.ep02(+)=c2.cp09 and e2.ep07>0))"
''end 2008/9/22
strSql = strSql & " union select '" & strUserNum & "',cp09,'2' from engineerprogress e1,caseprogress c1,staff where ep28<" & strSrvDate(1) & " and ep07 is null  and ep09 is null and cp09(+)=ep02 and cp01 in ('P','CFP') and cp27 is null and cp57 is null and cp14=st01(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51')"
'end 2010/7/20

'edit by nickc 2006/09/05 改有完稿的
'edit by nickc 2006/09/14 發文日改已發文
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾承辦期限之次日未輸會稿日及未完稿',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01='P' AND CP10<>'1101' and cp112='Y' and cp48<=" & P_date & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915') and ep09 is null  and ep28 is null "
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾承辦期限第二日未輸會稿日及未完稿',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01='CFP' AND CP10<>'1101' and cp112='Y' and cp48<=" & CFP_date & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915') and ep09 is null and ep28 is null "
'Modify by Morgan 2007/8/31 加807
'Modify by Morgan 2008/10/29 多的條件要拿掉否則乙規則的性質會抓不到,順便調語法
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾承辦期限第一日未輸會稿日及未完稿',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where      cp01='P'   AND CP10<>'1101' and cp112='Y' and cp48<=" & p_Date & "   and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915','807') and ep09 is null and ep28 is null "
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾承辦期限第二日未輸會稿日及未完稿',cp14 from caseprogress c1,engineerprogress e1,patent,casepropertymap,staff where cp01='CFP' AND CP10<>'1101' and cp112='Y' and cp48<=" & CFP_date & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915','807') and ep09 is null and ep28 is null "

'Modify by Morgan 2008/12/3 改判斷只要有定設規則的就要不管適用會稿加乘註記與否
'strSQL = strSQL & " union select /*+index(caseprogress IDXCP0114262757)*/ st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾承辦期限第一日未輸會稿日及未完稿',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where       cp01='P'   and cp112='Y' and cp48<=" & p_Date & "   and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep09 is null and ep28 is null "
'strSQL = strSQL & " union select /*+index(caseprogress IDXCP0114262757)*/ st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾承辦期限第二日未輸會稿日及未完稿',cp14 from caseprogress c1,engineerprogress e1,patent,casepropertymap,staff where cp01='CFP' and cp112='Y' and cp48<=" & CFP_date & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep09 is null and ep28 is null "
'2009/10/12 modify by sonia 王協理說要取消完稿日條件CFP-022500
'strSQL = strSQL & " union select /*+index(caseprogress IDXCP0114262757)*/ st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾承辦期限第一日未輸會稿日及未完稿',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where       cp01='P'   and cpm05 is not null and cp48<=" & p_Date & "   and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep09 is null and ep28 is null "
'strSQL = strSQL & " union select /*+index(caseprogress IDXCP0114262757)*/ st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾承辦期限第二日未輸會稿日及未完稿',cp14 from caseprogress c1,engineerprogress e1,patent,casepropertymap,staff where cp01='CFP' and cpm05 is not null and cp48<=" & CFP_date & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep09 is null and ep28 is null "

'Modify by Morgan 2010/7/20 改寫暫存以便過濾資料
'strSql = strSql & " union select /*+index(caseprogress IDXCP0114262757)*/ st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾承辦期限第一日未輸會稿日及未完稿',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where       cp01='P'   and cpm05 is not null and cp48<=" & p_Date & "   and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep28 is null "
'strSql = strSql & " union select /*+index(caseprogress IDXCP0114262757)*/ st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'逾承辦期限第二日未輸會稿日及未完稿',cp14 from caseprogress c1,engineerprogress e1,patent,casepropertymap,staff where cp01='CFP' and cpm05 is not null and cp48<=" & CFP_date & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep28 is null "
''2009/10/12 END
''end 2008/12/3
''Add by Morgan 2008/9/22 若為多國子案且母案已輸入會稿日則不顯示
'strSql = strSql & " and NOT (cp01='CFP' and cp21='Y' and instr('" & CaseMapOut & "',cp10)>0 and exists(select * from caserelation,caseprogress c2,engineerprogress e2" & _
' " Where cr01 = C1.CP01 And cr02 = C1.cp02 And cr03 = C1.cp03 And cr04 = C1.cp04 and c2.cp01(+)=cr05 and c2.cp02(+)=cr06 and c2.cp03(+)=cr07 and c2.cp04(+)=cr08" & _
' " and c2.cp14=c1.cp14 and c2.cp21 is null and instr('" & CaseMapOut & "',c2.cp10)>0 and e2.ep02(+)=c2.cp09 and e2.ep07>0))"
''end 2008/9/22
''end 2007/8/31
strSql = strSql & " union select /*+index(caseprogress IDXCP0114262757)*/ '" & strUserNum & "',cp09,'3' from caseprogress,engineerprogress,casepropertymap,staff where       cp01='P'   and cpm05 is not null and cp48<=" & p_Date & "   and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep28 is null "
strSql = strSql & " union select /*+index(caseprogress IDXCP0114262757)*/ '" & strUserNum & "',cp09,'4' from caseprogress c1,engineerprogress e1,casepropertymap,staff where cp01='CFP' and cpm05 is not null and cp48<=" & CFP_date & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep28 is null "
'end 2010/7/20


'edit by nickc 2006/09/14 拆成 2 次且發文日改已發文
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'超預會日已完稿逾承辦期限 4 日未輸會稿日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01='P' AND CP10<>'1101' and cp112='Y' and cp48<=" & P_date & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915')   and ep28<" & strSrvDate(1) & " and ep09 is not null "
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'超預會日已完稿逾承辦期限 6 日未輸會稿日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01='CFP' AND CP10<>'1101' and cp112='Y' and cp48<=" & CFP_date & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915')  and ep28<" & strSrvDate(1) & " and ep09 is not null "
'edit by nickc 2006/10/11 改回原來
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'已完稿逾承辦期限 4 日未輸會稿日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01='P' AND CP10<>'1101' and cp112='Y' and cp48<=" & P_date2 & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915') and ep09 is not null "
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'已完稿逾承辦期限 6 日未輸會稿日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01='CFP' AND CP10<>'1101' and cp112='Y' and cp48<=" & CFP_date2 & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915')  and ep09 is not null "
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'已完稿逾預定會稿日 4 日未輸會稿日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01='P' AND CP10<>'1101' and cp112='Y' and ep28<=" & P_date2 & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915')  and ep09 is not null "
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'已完稿逾預定會稿日 6 日未輸會稿日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp01='CFP' AND CP10<>'1101' and cp112='Y' and ep28<=" & CFP_date2 & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915') and ep09 is not null "

'Modify by Morgan 2007/8/31
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'超預會日已完稿逾承辦期限 4 日未輸會稿日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff" & _
   " where cp01='P' AND CP10<>'1101' and cp112='Y' and cp48<=" & P_date2 & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915')   and ep28<" & strSrvDate(1) & " and ep09 is not null "
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'超預會日已完稿逾承辦期限 6 日未輸會稿日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff" & _
   " where cp01='CFP' AND CP10<>'1101' and cp112='Y' and cp48<=" & CFP_date2 & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915')  and ep28<" & strSrvDate(1) & " and ep09 is not null "
'Modify by Morgan 2008/10/29 多的條件要拿掉否則乙規則的性質會抓不到,順便調語法
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'超預會日已完稿逾承辦期限 4 日未輸會稿日',cp14 from caseprogress c1,engineerprogress e1,patent,casepropertymap,staff" & _
'   " where cp01='P' AND CP10<>'1101' and cp112='Y' and cp48<=" & P_date2 & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915','807')   and ep28<" & strSrvDate(1) & " and ep09 is not null "
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'超預會日已完稿逾承辦期限 6 日未輸會稿日',cp14 from caseprogress c1,engineerprogress e1,patent,casepropertymap,staff" & _
'   " where cp01='CFP' AND CP10<>'1101' and cp112='Y' and cp48<=" & CFP_date2 & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ((cp04 = '00' and substr(cp09,1,1)='B') or substr(cp09,1,1) in ('A','C')) and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','915','807')  and ep28<" & strSrvDate(1) & " and ep09 is not null "

'Modify by Morgan 2008/12/3 改判斷只要有定設規則的就要不管適用會稿加乘註記與否
'strSQL = strSQL & " union select /*+index(caseprogress IDXCP0114262757)*/ st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'超預會日已完稿逾承辦期限 4 日未輸會稿日',cp14 from caseprogress c1,engineerprogress e1,patent,casepropertymap,staff" & _
'   " where cp01='P'   and cp112='Y' and cp48<=" & P_date2 & "   and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep28<" & strSrvDate(1) & " and ep09 is not null "
'strSQL = strSQL & " union select /*+index(caseprogress IDXCP0114262757)*/ st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'超預會日已完稿逾承辦期限 6 日未輸會稿日',cp14 from caseprogress c1,engineerprogress e1,patent,casepropertymap,staff" & _
'   " where cp01='CFP' and cp112='Y' and cp48<=" & CFP_date2 & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep28<" & strSrvDate(1) & " and ep09 is not null "

'Modify by Morgan 2010/7/20 改寫暫存以便過濾資料
'strSql = strSql & " union select /*+index(caseprogress IDXCP0114262757)*/ st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'超預會日已完稿逾承辦期限 4 日未輸會稿日',cp14 from caseprogress c1,engineerprogress e1,patent,casepropertymap,staff" & _
'   " where cp01='P'   and cpm05 is not null and cp48<=" & P_date2 & "   and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep28<" & strSrvDate(1) & " and ep09 is not null "
'strSql = strSql & " union select /*+index(caseprogress IDXCP0114262757)*/ st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'超預會日已完稿逾承辦期限 6 日未輸會稿日',cp14 from caseprogress c1,engineerprogress e1,patent,casepropertymap,staff" & _
'   " where cp01='CFP' and cpm05 is not null and cp48<=" & CFP_date2 & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep28<" & strSrvDate(1) & " and ep09 is not null "
''end 2008/12/3
''Add by Morgan 2008/9/22 若為多國子案且母案已輸入會稿日則不顯示
'strSql = strSql & " and NOT (cp01='CFP' and cp21='Y' and instr('" & CaseMapOut & "',cp10)>0 and exists(select * from caserelation,caseprogress c2,engineerprogress e2" & _
' " Where cr01 = C1.CP01 And cr02 = C1.cp02 And cr03 = C1.cp03 And cr04 = C1.cp04 and c2.cp01(+)=cr05 and c2.cp02(+)=cr06 and c2.cp03(+)=cr07 and c2.cp04(+)=cr08" & _
' " and c2.cp14=c1.cp14 and c2.cp21 is null and instr('" & CaseMapOut & "',c2.cp10)>0 and e2.ep02(+)=c2.cp09 and e2.ep07>0))"
''end 2008/9/22
''end 2007/8/31
strSql = strSql & " union select /*+index(caseprogress IDXCP0114262757)*/ '" & strUserNum & "',cp09,'5' from caseprogress c1,engineerprogress e1,casepropertymap,staff" & _
   " where cp01='P'   and cpm05 is not null and cp48<=" & p_Date2 & "   and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep28<" & strSrvDate(1) & " and ep09 is not null "
strSql = strSql & " union select /*+index(caseprogress IDXCP0114262757)*/ '" & strUserNum & "',cp09,'6' from caseprogress c1,engineerprogress e1,casepropertymap,staff" & _
   " where cp01='CFP' and cpm05 is not null and cp48<=" & CFP_date2 & " and ep07 is null and cp27 is null and cp57 is null and cp09=ep02(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ((st03<>'P12' and st03>='P' and st03<='P19') or st03='F51') and ep28<" & strSrvDate(1) & " and ep09 is not null "
'end 2010/7/20



'add by nickc 2006/08/24 加入會稿日自動上的紀錄
'edit by nickc 2006/09/14 發文日改已發文
'strSQL = strSQL & " union select st02,sqldateT(cp05),sqldatet(ep06)," & SQLDate("cp27") & ",cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'系統自動上會稿完成為發文日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp27>=" & Val(DBDATE(txt1(0).Text)) & " and cp27<=" & Val(DBDATE(txt1(1).Text)) & " and cp09=ep02(+) and cp09 in (select pl01 from pat_log) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) order by cp14  "
'Modify by Morgan 2010/7/20 改寫暫存以便過濾資料
'strSql = strSql & " union select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04,nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),'系統自動上會稿完成為發文日',cp14 from caseprogress,engineerprogress,patent,casepropertymap,staff where cp27>=" & Val(DBDATE(Txt1(0).Text)) & " and cp27<=" & Val(DBDATE(Txt1(1).Text)) & " and cp09=ep02(+) and cp09 in (select pl01 from pat_log) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+)  and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) order by cp14  "
strSql = strSql & " union select '" & strUserNum & "',cp09,'7' from caseprogress,engineerprogress,staff where cp27>=" & Val(DBDATE(txt1(0).Text)) & " and cp27<=" & Val(DBDATE(txt1(1).Text)) & " and cp09=ep02(+) and cp09 in (select pl01 from pat_log) and cp14=st01(+)"
'end 2010/7/20


'Add by Morgan 2010/7/20
cnnConnection.Execute "delete R090628 where ID='" & strUserNum & "'"
cnnConnection.Execute "insert into R090628 (ID,R01,R02) " & strSql, intI
'多國案控制美日德案都顯示,其餘國家若同一工程師承辦則只要顯示一筆(非美日德且無美日德案時顯示收文號最小的案件)
strSql = "delete from R090628 a where ID='" & strUserNum & "'" & _
   " and exists(select * from caseprogress c,patent e" & _
   " where cp09=r01 and cp01='CFP' and cp10 in (" & NewCasePtyList & ")" & _
   " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09 not in('011','101','231')" & _
   " and c.cp09>(select min(decode(instr('011,101,231',f.pa09),0,'',' ')||d.cp09)" & _
   " from caserelation,caseprogress d,R090628 b,patent f" & _
   " Where cr01 = c.CP01 And cr02 = c.cp02 And cr03 = c.cp03 And cr04 = c.cp04" & _
   " and d.cp01(+)=cr05 and d.cp02(+)=cr06 and d.cp03(+)=cr07 and d.cp04(+)=cr08" & _
   " and d.cp10 in (" & NewCasePtyList & ") and d.cp14=c.cp14" & _
   " and f.pa01(+)=d.cp01 and f.pa02(+)=d.cp02 and f.pa03(+)=d.cp03 and f.pa04(+)=d.cp04" & _
   " and b.R01(+)=d.cp09 and b.R02=a.R02))"
cnnConnection.Execute strSql, intI

strSql = "select st02,sqldateT(cp05),sqldatet(ep06),decode(cp27,null,'','Y'),cp01||'-'||cp02||'-'||cp03||'-'||cp04" & _
   ",nvl(pa05,nvl(pa06,pa07)),decode(pa09,'000',cpm03,cpm04),decode(R02,'1','已發文未輸會稿完成'" & _
   ",'2','逾預定會稿日未輸會稿日及未完稿','3','逾承辦期限第一日未輸會稿日及未完稿'" & _
   ",'4','逾承辦期限第二日未輸會稿日及未完稿','5','超預會日已完稿逾承辦期限 4 日未輸會稿日'" & _
   ",'6','超預會日已完稿逾承辦期限 6 日未輸會稿日','7','系統自動上會稿完成為發文日'),cp14" & _
   " from R090628,caseprogress,engineerprogress,patent,casepropertymap,staff" & _
   " where ID='" & strUserNum & "' and cp09(+)=R01 and ep02(+)=cp09" & _
   " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
   " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp14" & _
   " order by cp14,cp09"
'end 2010/7/20

CheckOC
'Set adoRecordset = New ADODB.Recordset
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/20
        Set grdDataList.Recordset = adoRecordset
    '94.1.5 add by sonia
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/20
        MsgBox "無符合條件之資料 !", vbInformation
        grdDataList.Clear
    '94.1.5 end
    End If
End With
CheckOC
SetDataListWidth
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
grdDataList.Clear
SetDataListWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090628 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
        If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
           Me.txt1(Index).SetFocus
           txt1_GotFocus Index
           Cancel = True
           Exit Sub
        End If
        If Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
                txt1(Index - 1).SetFocus
                txt1_GotFocus (Index - 1)
                Cancel = True
                Exit Sub
            End If
        End If
End Sub
