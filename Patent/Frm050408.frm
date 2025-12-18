VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050408 
   BorderStyle     =   1  '單線固定
   Caption         =   "互惠代理人案件統計表"
   ClientHeight    =   3780
   ClientLeft      =   3048
   ClientTop       =   1512
   ClientWidth     =   3900
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3900
   Begin VB.CommandButton cmdStatistic 
      BackColor       =   &H00C0FFFF&
      Caption         =   "互惠代理人案件盈虧統計表"
      Height          =   492
      Left            =   120
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   72
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtKind 
      Height          =   300
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1740
      Width           =   345
   End
   Begin VB.ComboBox cboTarget 
      Height          =   300
      ItemData        =   "frm050408.frx":0000
      Left            =   1410
      List            =   "frm050408.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   1380
      Width           =   1410
   End
   Begin VB.ComboBox cboPeriod 
      Height          =   300
      Left            =   1410
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   1020
      Width           =   1410
   End
   Begin VB.TextBox txtYear 
      Height          =   300
      Left            =   1410
      MaxLength       =   3
      TabIndex        =   0
      Top             =   660
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "指定日期區間(請輸民國年月日)"
      Height          =   645
      Left            =   270
      TabIndex        =   10
      Top             =   2190
      Width           =   3345
      Begin VB.TextBox Text2 
         Height          =   264
         Left            =   1800
         MaxLength       =   7
         TabIndex        =   5
         Top             =   240
         Width           =   1290
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Left            =   270
         MaxLength       =   7
         TabIndex        =   4
         Top             =   240
         Width           =   1290
      End
      Begin VB.Line Line1 
         X1              =   1530
         X2              =   1860
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   2880
      TabIndex        =   7
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   2100
      TabIndex        =   6
      Top             =   60
      Width           =   756
   End
   Begin MSForms.Label lblAppNo 
      Height          =   372
      Index           =   1
      Left            =   288
      TabIndex        =   17
      Top             =   3264
      Width           =   3372
      ForeColor       =   16711680
      Caption         =   "代理人名稱"
      Size            =   "5948;656"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAppNo 
      Height          =   228
      Index           =   0
      Left            =   288
      TabIndex        =   16
      Top             =   2976
      Width           =   2220
      ForeColor       =   16711680
      Caption         =   "代理人編號："
      Size            =   "3916;402"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "（1.專利 2.商標）"
      Height          =   180
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      Top             =   1770
      Width           =   1395
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "案件類別："
      Height          =   180
      Index           =   0
      Left            =   420
      TabIndex        =   14
      Top             =   1770
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "統計對象："
      Height          =   180
      Left            =   420
      TabIndex        =   13
      Top             =   1410
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "統計年度："
      Height          =   180
      Left            =   420
      TabIndex        =   12
      Top             =   690
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "統計區間："
      Height          =   180
      Left            =   420
      TabIndex        =   11
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "(1. 管制表 2. 定稿)"
      Height          =   180
      Left            =   5955
      TabIndex        =   9
      Top             =   5325
      Width           =   1425
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "列印格式:"
      Height          =   180
      Left            =   4170
      TabIndex        =   8
      Top             =   5340
      Width           =   765
   End
End
Attribute VB_Name = "frm050408"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/11 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2008/3/4
Option Explicit

Public m_bolByAgent As Boolean  '是否以代理人統計
Dim m_bQuery As Boolean '查詢
Dim m_bPrint As Boolean '列印
'Added by Lydia 2020/09/22 另外抓語法
Public strConFirst As String '畫面上指定案件類別
Public strConSecond As String '另外抓額外的案件類別: ex. First=專利, Second=商標
'Added by Lydia 2025/06/06
Dim m_PrevForm As Form '前一畫面
Dim m_AppNo As String  '傳入代理人編號
Public m_strQL05 As String 'Added by Lydia 2025/07/30

'Added by Lydia 2025/06/06
Public Sub SetParent(ByRef fm As Form, ByVal pNo As String)
   Set m_PrevForm = fm
   m_AppNo = ChangeCustomerL(pNo)
End Sub

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 1
         Screen.MousePointer = vbHourglass
         If TxtValidate Then
            Process
         End If
         Screen.MousePointer = vbDefault
      Case 2
         Unload Me
   End Select
End Sub

'Memo by Lydia 2020/09/22 原模組：不可產生統計表excel
'Mark by Lydia 2025/06/06 確定不使用
'Private Sub Process_old()
'   Dim stVTB1 As String, stVTB2 As String
'   Dim stDate As String, iYear As Integer, iPeriod As Integer
'   Dim stDate1 As String, stDate2 As String, bExtra As Boolean
'   Dim strFC06 As String 'Add By Sindy 2013/5/23
'
'   If cboTarget.ListIndex = 0 Then
'      m_bolByAgent = True
'   Else
'      m_bolByAgent = False
'   End If
'
'   'Add by Morgan 2008/6/24
'   If Text1 <> "" Or Text2 <> "" Then
'      bExtra = True
'      If Text1 <> "" Then
'         stDate1 = DBDATE(Text1)
'      Else
'         stDate1 = 0
'      End If
'      If Text1 <> "" Then
'         stDate2 = DBDATE(Text2)
'      Else
'         stDate2 = strSrvDate(1)
'      End If
'   End If
'
'   stDate = strSrvDate(1)
'
'   'Modif by Morgan 2008/7/22 統計年度&區間改抓畫面輸入
'   'iYear = Left(stDate, 4)
'   'If Val(Mid(stDate, 5, 2)) > 6 Then
'   '   iPeriod = 2
'   'Else
'   '   iPeriod = 1
'   'End If
'   iYear = Val(txtYear) + 1911
'   iPeriod = cboPeriod.ListIndex + 1
'   'end 2008/7/22
'
'   'Add By Sindy 2013/5/23
'   If txtKind = "1" Then '專利
'      strFC06 = "CFP"
'   ElseIf txtKind = "2" Then '商標
'      strFC06 = "CFT"
'   End If
'   '2013/5/23 End
'
'   'CF案件統計
'   stVTB1 = "select FC01||FC03 A2,count(*) CF_TOT" & _
'      ",sum(decode(substr(cp27,1,4)," & iYear & "-2,1)) CF_L2" & _
'      ",sum(decode(substr(cp27,1,4)," & iYear & "-1,1)) CF_L1" & _
'      ",sum(decode(substr(cp27,1,4)," & iYear & ",decode(sign(substr(cp27,5,2)-6),1,0,1))) CF_C1" & _
'      ",sum(decode(substr(cp27,1,4)," & iYear & ",decode(sign(substr(cp27,5,2)-6),1,1,0))) CF_C2"
'   'Add by Morgan 2008/6/24
'   If bExtra = True Then
'      'Modified by Morgan 2015/12/1
'      'stVTB1 = stVTB1 & ",sum(decode(sign(cp27-" & stDate1 & "+1),1,decode(sign(cp27-" & stDate2 & "+1),-1,1))) CF_X"
'      stVTB1 = stVTB1 & ",sum(decode(sign(cp27-" & stDate1 & "+1),1,decode(sign(cp27-" & stDate2 & "-1),-1,1))) CF_X"
'   End If
'
'   If txtKind = "1" Then '專利
'      'Modify by Morgan 2008/7/22 加以代理人統計
'      If m_bolByAgent = True Then
'         stVTB1 = stVTB1 & " From (SELECT DISTINCT FC01,'' FC03 FROM FAGENTCONFIG" & _
'            " Where FC06='" & strFC06 & "' AND FC04 = " & (iYear - 1911) & " And FC05 = " & iPeriod & _
'            " ) VTB1,caseprogress WHERE CP44=FC01||'0'" & _
'            " and cp01||cp04='CFP00' and cp27+0>19221111 and cp57 is null and cp27+0<=" & stDate & _
'            " and instr('" & NewCasePtyList & "',cp10)>0 and cp09<'B'" & _
'            " group by FC01,FC03"
'      Else
'         stVTB1 = stVTB1 & " From FAGENTCONFIG, caseprogress" & _
'            " Where FC06='" & strFC06 & "' AND FC04 = " & (iYear - 1911) & " And FC05 = " & iPeriod & _
'            " AND CP44=FC01||'0' AND NVL(CP116,'0')=NVL(FC03,'0')" & _
'            " and cp01||cp04='CFP00' and cp27+0>19221111 and cp57 is null and cp27+0<=" & stDate & _
'            " and instr('" & NewCasePtyList & "',cp10)>0 and cp09<'B'" & _
'            " group by FC01,FC03"
'      End If
'   'Add By Sindy 2013/5/23
'   ElseIf txtKind = "2" Then '商標
'      '加以代理人統計
'      If m_bolByAgent = True Then
'         stVTB1 = stVTB1 & " From (SELECT DISTINCT FC01,'' FC03 FROM FAGENTCONFIG" & _
'            " Where FC06='" & strFC06 & "' AND FC04 = " & (iYear - 1911) & " And FC05 = " & iPeriod & _
'            " ) VTB1,caseprogress WHERE CP44=FC01||'0'" & _
'            " and cp01||cp04='CFT00' and cp27+0>19221111 and cp57 is null and cp27+0<=" & stDate & _
'            " and instr('101',cp10)>0 and cp09<'B'" & _
'            " group by FC01,FC03"
'      Else
'         stVTB1 = stVTB1 & " From FAGENTCONFIG, caseprogress" & _
'            " Where FC06='" & strFC06 & "' AND FC04 = " & (iYear - 1911) & " And FC05 = " & iPeriod & _
'            " AND CP44=FC01||'0' AND NVL(CP116,'0')=NVL(FC03,'0')" & _
'            " and cp01||cp04='CFT00' and cp27+0>19221111 and cp57 is null and cp27+0<=" & stDate & _
'            " and instr('101',cp10)>0 and cp09<'B'" & _
'            " group by FC01,FC03"
'      End If
'   End If
'   '2013/5/23 End
'
'   'FC案件統計
'   stVTB2 = "select FC01||FC03 B2,count(*) FC_TOT" & _
'      ",sum(decode(substr(cp05,1,4)," & iYear & "-2,1)) FC_L2" & _
'      ",sum(decode(substr(cp05,1,4)," & iYear & "-1,1)) FC_L1" & _
'      ",sum(decode(substr(cp05,1,4)," & iYear & ",decode(sign(substr(cp05,5,2)-6),1,0,1))) FC_C1" & _
'      ",sum(decode(substr(cp05,1,4)," & iYear & ",decode(sign(substr(cp05,5,2)-6),1,1,0))) FC_C2"
'   'Add by Morgan 2008/6/24
'   If bExtra = True Then
'      stVTB2 = stVTB2 & _
'         ",sum(decode(sign(cp05-" & stDate1 & "+1),1,decode(sign(cp05-" & stDate2 & "+1),-1,1))) FC_X"
'   End If
'
'   If txtKind = "1" Then '專利
'      'Modify by Morgan 2008/7/22 加以代理人統計
'      If m_bolByAgent = True Then
'         'Modified by Lydia 2018/06/21 改新申請案性質instr('101,102,103,104,105,307',cp10)>0 => instr('" & NewCasePtyList & "',cp10)>0
'         stVTB2 = stVTB2 & " from (select DISTINCT FC01,'' FC03 from FAGENTCONFIG where FC06='" & strFC06 & "' AND FC04=" & (iYear - 1911) & " AND FC05=" & iPeriod & _
'            ") VTB2,patent,caseprogress" & _
'            " where pa75=FC01||'0' AND pa01||''='FCP'" & _
'            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp09<'B'" & _
'            " and instr('" & NewCasePtyList & "',cp10)>0 and cp57 is null and cp05+0<=" & stDate & _
'            " group by FC01,FC03"
'      Else
'         'Modified by Lydia 2018/06/21 改新申請案性質instr('101,102,103,104,105,307',cp10)>0 => instr('" & NewCasePtyList & "',cp10)>0
'         stVTB2 = stVTB2 & " from (select FC01,FC03 from FAGENTCONFIG where FC06='" & strFC06 & "' AND FC04=" & (iYear - 1911) & " AND FC05=" & iPeriod & _
'            ") X,patent,caseprogress" & _
'            " where pa75=FC01||'0' AND NVL(pa144,'0')=NVL(FC03,'0') AND pa01||''='FCP'" & _
'            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp09<'B'" & _
'            " and instr('" & NewCasePtyList & "',cp10)>0 and cp57 is null and cp05+0<=" & stDate & _
'            " group by FC01,FC03"
'      End If
'   'Add By Sindy 2013/5/23
'   ElseIf txtKind = "2" Then '商標
'      '加以代理人統計
'      If m_bolByAgent = True Then
'         stVTB2 = stVTB2 & " from (select DISTINCT FC01,'' FC03 from FAGENTCONFIG where FC06='" & strFC06 & "' AND FC04=" & (iYear - 1911) & " AND FC05=" & iPeriod & _
'            ") VTB2,trademark,caseprogress" & _
'            " where tm44=FC01||'0' AND tm01||''='FCT'" & _
'            " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04 and cp09<'B'" & _
'            " and instr('101',cp10)>0 and cp57 is null and cp05+0<=" & stDate & _
'            " group by FC01,FC03"
'      Else
'         stVTB2 = stVTB2 & " from (select FC01,FC03 from FAGENTCONFIG where FC06='" & strFC06 & "' AND FC04=" & (iYear - 1911) & " AND FC05=" & iPeriod & _
'            ") X,trademark,caseprogress" & _
'            " where tm44=FC01||'0' AND NVL(tm119,'0')=NVL(FC03,'0') AND tm01||''='FCT'" & _
'            " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04 and cp09<'B'" & _
'            " and instr('101',cp10)>0 and cp57 is null and cp05+0<=" & stDate & _
'            " group by FC01,FC03"
'      End If
'   End If
'   '2013/5/23 End
'
'   strExc(0) = "select NVL(FA05,NVL(FA06,FA04)) C1,FC01||DECODE(FC03,NULL,'','-'||FC03) C2" & _
'      ",NVL(PCC03,NVL(PCC04,PCC05)) C3,NA03,nvl(FC_TOT,0),nvl(CF_TOT,0),nvl(FC_L2,0),nvl(CF_L2,0)" & _
'      ",nvl(FC_L1,0),nvl(CF_L1,0),nvl(FC_C1,0),nvl(CF_C1,0),nvl(FC_C1,0)-nvl(CF_C1,0),nvl(FC_C2,0)" & _
'      ",nvl(CF_C2,0),nvl(FC_C2,0)-nvl(CF_C2,0),FC07,FC08"
'   'Add by Morgan 2008/6/24
'   If bExtra = True Then
'      strExc(0) = strExc(0) & ",nvl(FC_X,0),nvl(CF_X,0),nvl(FC_X,0)-nvl(CF_X,0)"
'   End If
'
'   'Modify by Morgan 2008/7/22 加以代理人統計
'   If m_bolByAgent = True Then
'      strExc(1) = strExc(0) & " from (SELECT FC01,'' FC03,'' PCC03,'' PCC04,'' PCC05,SUM(FC07) FC07,MIN(FC08) FC08" & _
'         " FROM FAGENTCONFIG where FC06='" & strFC06 & "' AND FC04=" & (iYear - 1911) & " AND FC05=" & iPeriod & _
'         " GROUP BY FC01) X,(" & stVTB1 & ") A,(" & stVTB2 & ") B,FAGENT,NATION where A2(+)=FC01||FC03" & _
'         " and B2(+)=FC01||FC03 AND FA01(+)=FC01 AND FA02(+)='0' AND NA01(+)=SUBSTR(FA10,1,3)" & _
'         " ORDER BY NA01 ASC,FC07 DESC,C1,FC01 ASC"
'   Else
'      strExc(1) = strExc(0) & " from FAGENTCONFIG,(" & stVTB1 & ") A,(" & stVTB2 & ") B,FAGENT,NATION,POTCUSTCONT" & _
'         " where FC06='" & strFC06 & "' AND FC04=" & (iYear - 1911) & " AND FC05=" & iPeriod & " and A2(+)=FC01||FC03" & _
'         " and B2(+)=FC01||FC03 AND FA01(+)=FC01 AND FA02(+)='0' AND NA01(+)=SUBSTR(FA10,1,3) AND PCC01(+)=FC01 AND PCC02(+)=FC03" & _
'         " ORDER BY NA01 ASC,FC07 DESC,C1,FC01 ASC,FC03 ASC"
'   End If
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
'   If intI = 1 Then
'      With frm050408_1
'         .Show
'         .Caption = Me.Caption & "(" & txtYear & "年" & cboPeriod & ")"
'         .txt1(0) = txtYear & "0101"
'         .txt1(1) = IIf(cboPeriod.ListIndex = 0, txtYear & "0630", txtYear & "1231")
'         'Add by Morgan 2008/6/24
'         If bExtra = True Then
'            .lblExtra = "指定日期區間：" & Text1 & " － " & Text2
'            .m_stDate1 = stDate1
'            .m_stDate2 = stDate2
'            .cmdok(4).Visible = True
'            .cmdok(5).Visible = True
'         End If
'         Set .m_adoRst = RsTemp
'         .InitGrid
'         If m_bPrint = False Then
'            .cmdok(2).Visible = False
'            .cmdok(3).Visible = False
'            .cmdok(4).Visible = False
'            .cmdok(5).Visible = False
'         End If
'      End With
'      Me.Hide
'   Else
'      MsgBox "無資料！"
'   End If
'End Sub
'end 2025/06/06

'Memo by Lydia 2020/09/22 可產生統計表excel
Private Sub Process()
   Dim stVTB1 As String, stVTB2 As String
   Dim stDate As String, iYear As Integer, iPeriod As Integer
   Dim stDate1 As String, stDate2 As String, bExtra As Boolean
   'Modified by Lydia 2020/09/22
   'Dim strFC06 As String 'Add By Sindy 2013/5/23
   Dim strFC06_P As String, strFC06_T As String
   'Added by Lydia 2020/09/22 分別存放專利、商標SQL語法
   Dim tmpPA As String, tmpTM As String, tmpPA2 As String, tmpTM2 As String
   Dim rsAD As New ADODB.Recordset
   Dim rsRD1 As New ADODB.Recordset, intJ As Integer 'Added by Lydia 2020/12/22
   
   ClearQueryLog (Me.Name) 'Added by Lydia 2025/07/30
   
   
   If cboTarget.ListIndex = 0 Then
      m_bolByAgent = True
   Else
      m_bolByAgent = False
   End If
   
   'Add by Morgan 2008/6/24
   If Text1 <> "" Or Text2 <> "" Then
      bExtra = True
      If Text1 <> "" Then
         stDate1 = DBDATE(Text1)
      Else
         stDate1 = 0
      End If
      If Text1 <> "" Then
         stDate2 = DBDATE(Text2)
      Else
         stDate2 = strSrvDate(1)
      End If
   End If
   
   stDate = strSrvDate(1)
   
   strConFirst = "":   strConSecond = "" 'Added by Lydia 2020/09/22
   
   'Modif by Morgan 2008/7/22 統計年度&區間改抓畫面輸入
   'iYear = Left(stDate, 4)
   'If Val(Mid(stDate, 5, 2)) > 6 Then
   '   iPeriod = 2
   'Else
   '   iPeriod = 1
   'End If
   iYear = Val(txtYear) + 1911
   iPeriod = cboPeriod.ListIndex + 1
   'end 2008/7/22
   
   'Add By Sindy 2013/5/23
   'Modified by Lydia 2020/09/22
   'If txtKind = "1" Then '專利
   '   strFC06 = "CFP"
   'ElseIf txtKind = "2" Then '商標
   '   strFC06 = "CFT"
   'End If
   '2013/5/23 End
   strFC06_P = "CFP"
   strFC06_T = "CFT"
   'end 2020/09/22
   
   'Added by Lydia 2020/09/21 暫存: 符合條件的代理人互惠設定檔和相對的專利/商標設定檔
   'Modified by Lydia 2023/09/12 排序依照報表排列,區分非關聯客戶=>模組化
'   cnnConnection.Execute "delete from rdatafactory where FORMNAME=" & CNULL(Me.Name) & " and ID=" & CNULL(strUserNum)
'   'Modified by Lydia 2022/05/12 +FC16,FC17
'   strExc(0) = "SELECT FC01,FC02,FC03,FC04,FC05,FC06,FC07,FC08,FC16,FC17 FROM FAGENTCONFIG WHERE FC06=" & CNULL(IIf(txtKind = "1", strFC06_P, strFC06_T)) & " AND FC04 = " & (iYear - 1911) & " AND FC05 = " & iPeriod & " order by fc01,fc02,fc03 "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      Set rsAD = PUB_CreateRecordset(RsTemp, , , , Me.Name) '先暫存:符合條件的代理人互惠設定檔
'      'Modified by Lydia 2022/05/12 + r009,r010
'      strExc(0) = "select formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008,r009,r010 from rdatafactory where FORMNAME=" & CNULL(Me.Name) & " and ID=" & CNULL(strUserNum) & " and seqno =1 order by rowseq "
'      intI = 1
'      intJ = 1 'Added by Lydia 2020/12/22
'      Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
'      rsAD.MoveFirst
'      Do While Not rsAD.EOF
'          'Mark by Lydia 2020/12/22 抓所有聯絡人並產生設定檔; 原本程式只抓有設定的聯絡人
''          If m_bolByAgent = False Then
''                strExc(1) = "select r001 from rdatafactory where FORMNAME=" & CNULL(Me.Name) & " and ID=" & CNULL(strUserNum) & " and seqno =2 and r001=" & CNULL(rsAD.Fields("R001")) & "  and r006=" & CNULL(IIf(txtKind = "1", strFC06_P, strFC06_T)) & ""
''                intI = 1
''                Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
''                If intI = 0 Then
''                    strExc(1) = "Select " & CNULL(rsAD.Fields("R001")) & " As A01,'0' As A02,Null As A03 From Dual " & _
''                                     "Union Select Pcc01,'0' As A02,Pcc02 From Potcustcont Where Pcc01=" & CNULL(rsAD.Fields("R001")) & " order by 1 "
''                    intI = 1
''                    Set rsRD1 = ClsLawReadRstMsg(intI, strExc(1))
''                    If intI = 1 Then
''                        rsRD1.MoveFirst
''                        Do While Not rsRD1.EOF
''                            strSql = "INSERT INTO rdatafactory (formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008) " & _
''                                        "SELECT " & CNULL(Me.Name) & " as frmname, " & CNULL(strUserNum) & " as id, 2 as seqno, " & intJ & " as rowseqno, " & _
''                                        "FC01,FC02,FC03,FC04,FC05,FC06,FC07,FC08 FROM FAGENTCONFIG WHERE FC06=" & CNULL(IIf(txtKind = "2", strFC06_T, strFC06_P)) & " AND FC04 = " & (iYear - 1911) & " AND FC05 = " & iPeriod & _
''                                        " AND FC01=" & CNULL(rsAD.Fields("R001")) & " AND FC02=" & CNULL(rsAD.Fields("R002")) & IIf("" & rsRD1.Fields("A03") = "", " AND FC03 IS NULL ", " AND FC03='" & rsRD1.Fields("A03") & "' ")
''                            cnnConnection.Execute strSql, intI
''                            If intI = 0 Then
''                                strSql = "INSERT INTO rdatafactory (formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008) " & _
''                                            "values ('" & Me.Name & "','" & strUserNum & "', '2' , '" & intJ & "', '" & rsAD.Fields("r001") & "', '" & rsAD.Fields("r002") & "', " & CNULL("" & rsRD1.Fields("A03")) & ", '" & iYear - 1911 & "', '" & iPeriod & "'," & _
''                                            CNULL(IIf(txtKind = "2", strFC06_T, strFC06_P)) & ", '0', '' ) "
''                                cnnConnection.Execute strSql, intI
''                            End If
''                            intJ = intJ + 1
''                            '相對的專利/商標設定檔
''                            strSql = "INSERT INTO rdatafactory (formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008) " & _
''                                        "SELECT " & CNULL(Me.Name) & " as frmname, " & CNULL(strUserNum) & " as id, 2 as seqno, " & intJ & " as rowseqno, " & _
''                                        "FC01,FC02,FC03,FC04,FC05,FC06,FC07,FC08 FROM FAGENTCONFIG WHERE FC06=" & CNULL(IIf(txtKind = "1", strFC06_T, strFC06_P)) & " AND FC04 = " & (iYear - 1911) & " AND FC05 = " & iPeriod & _
''                                        " AND FC01=" & CNULL(rsAD.Fields("R001")) & " AND FC02=" & CNULL(rsAD.Fields("R002")) & IIf("" & rsRD1.Fields("A03") = "", " AND FC03 IS NULL ", " AND FC03='" & rsRD1.Fields("A03") & "' ")
''                            cnnConnection.Execute strSql, intI
''                            If intI = 0 Then
''                                strSql = "INSERT INTO rdatafactory (formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008) " & _
''                                            "values ('" & Me.Name & "','" & strUserNum & "', '2' , '" & intJ & "', '" & rsAD.Fields("r001") & "', '" & rsAD.Fields("r002") & "', " & CNULL("" & rsRD1.Fields("A03")) & ", '" & iYear - 1911 & "', '" & iPeriod & "'," & _
''                                            CNULL(IIf(txtKind = "1", strFC06_T, strFC06_P)) & ", '0', '' ) "
''                                cnnConnection.Execute strSql, intI
''                            End If
''                            intJ = intJ + 1
''                            rsRD1.MoveNext
''                        Loop
''                    End If
''                End If
''                Set rsRD1 = Nothing
''          Else
''          'end 2020/12/22
'                'Modified by Lydia 2022/05/12 + r009,r010; + FC16,FC17
'                strSql = "INSERT INTO rdatafactory (formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008,r009,r010) " & _
'                            "SELECT " & CNULL(Me.Name) & " as frmname, " & CNULL(strUserNum) & " as id, 2 as seqno, " & rsAD.Fields("rowseq") & " as rowseqno, " & _
'                            "FC01,FC02," & IIf(m_bolByAgent = True, " '' as FC03", "FC03") & ",FC04,FC05,FC06,FC07,FC08,FC16,FC17 " & _
'                            "FROM FAGENTCONFIG WHERE FC06=" & CNULL(IIf(txtKind = "1", strFC06_T, strFC06_P)) & " AND FC04 = " & (iYear - 1911) & " AND FC05 = " & iPeriod & _
'                            " AND FC01=" & CNULL(rsAD.Fields("R001")) & " AND FC02=" & CNULL(rsAD.Fields("R002"))
'                cnnConnection.Execute strSql, intI
'                If intI = 0 Then
'                    'Modified by Lydia 2022/05/12 + r009,r010
'                    strSql = "INSERT INTO rdatafactory (formname,id,seqno,rowseq,r001,r002,r003,r004,r005,r006,r007,r008,r009,r010) " & _
'                                "values ('" & Me.Name & "','" & strUserNum & "', '2' , '" & rsAD.Fields("rowseq") & "', '" & rsAD.Fields("r001") & "', '" & rsAD.Fields("r002") & "', NULL, '" & rsAD.Fields("r004") & "', '" & rsAD.Fields("r005") & "'," & _
'                                CNULL(IIf(txtKind = "1", strFC06_T, strFC06_P)) & ", '0', '','','') "
'                    cnnConnection.Execute strSql, intI
'                End If
''          End If 'Added by Lydia 2020/12/22
'          rsAD.MoveNext
'      Loop
'   End If
'   'end 2020/09/21
   'Modified by Lydia 2025/06/06
   'Call Pub_GetFCfrm050408(strUserNum, Me.txtKind, "0", Me.txtYear, "" & iPeriod, m_bolByAgent)
   Call Pub_GetFCfrm050408(strUserNum, Me.txtKind, "0", Me.txtYear, "" & iPeriod, m_bolByAgent, m_AppNo, m_AppNo)
   'end 2023/09/12
   
   'Modified by Lydia 2023/09/12 改成模組Pub_GetSqlfrm050408取得SQL
'   'CF案件統計
'   stVTB1 = "select FC01||FC03 A2,count(*) CF_TOT" & _
'      ",sum(decode(substr(cp27,1,4)," & iYear & "-2,1)) CF_L2" & _
'      ",sum(decode(substr(cp27,1,4)," & iYear & "-1,1)) CF_L1" & _
'      ",sum(decode(substr(cp27,1,4)," & iYear & ",decode(sign(substr(cp27,5,2)-6),1,0,1))) CF_C1" & _
'      ",sum(decode(substr(cp27,1,4)," & iYear & ",decode(sign(substr(cp27,5,2)-6),1,1,0))) CF_C2 "
'   'Add by Morgan 2008/6/24
'   If bExtra = True Then
'      'Modified by Morgan 2015/12/1
'      'stVTB1 = stVTB1 & ",sum(decode(sign(cp27-" & stDate1 & "+1),1,decode(sign(cp27-" & stDate2 & "+1),-1,1))) CF_X"
'      stVTB1 = stVTB1 & ",sum(decode(sign(cp27-" & stDate1 & "+1),1,decode(sign(cp27-" & stDate2 & "-1),-1,1))) CF_X "
'   End If
'
'   'Modified by Lydia 2020/09/22 代理人互惠設定檔改用暫存；專利、商標案件語法分別暫存
''   If txtKind = "1" Then '專利
''      'Modify by Morgan 2008/7/22 加以代理人統計
''      If m_bolByAgent = True Then
''         stVTB1 = stVTB1 & " From (SELECT DISTINCT FC01,'' FC03 FROM FAGENTCONFIG" & _
''            " Where FC06='" & strFC06 & "' AND FC04 = " & (iYear - 1911) & " And FC05 = " & iPeriod & _
''            " ) VTB1,caseprogress WHERE CP44=FC01||'0'" & _
''            " and cp01||cp04='CFP00' and cp27+0>19221111 and cp57 is null and cp27+0<=" & stDate & _
''            " and instr('" & NewCasePtyList & "',cp10)>0 and cp09<'B'" & _
''            " group by FC01,FC03"
''      Else
''         stVTB1 = stVTB1 & " From FAGENTCONFIG, caseprogress" & _
''            " Where FC06='" & strFC06 & "' AND FC04 = " & (iYear - 1911) & " And FC05 = " & iPeriod & _
''            " AND CP44=FC01||'0' AND NVL(CP116,'0')=NVL(FC03,'0')" & _
''            " and cp01||cp04='CFP00' and cp27+0>19221111 and cp57 is null and cp27+0<=" & stDate & _
''            " and instr('" & NewCasePtyList & "',cp10)>0 and cp09<'B'" & _
''            " group by FC01,FC03"
''      End If
''   'Add By Sindy 2013/5/23
''   ElseIf txtKind = "2" Then '商標
''      '加以代理人統計
''      If m_bolByAgent = True Then
''         stVTB1 = stVTB1 & " From (SELECT DISTINCT FC01,'' FC03 FROM FAGENTCONFIG" & _
''            " Where FC06='" & strFC06 & "' AND FC04 = " & (iYear - 1911) & " And FC05 = " & iPeriod & _
''            " ) VTB1,caseprogress WHERE CP44=FC01||'0'" & _
''            " and cp01||cp04='CFT00' and cp27+0>19221111 and cp57 is null and cp27+0<=" & stDate & _
''            " and instr('101',cp10)>0 and cp09<'B'" & _
''            " group by FC01,FC03"
''      Else
''         stVTB1 = stVTB1 & " From FAGENTCONFIG, caseprogress" & _
''            " Where FC06='" & strFC06 & "' AND FC04 = " & (iYear - 1911) & " And FC05 = " & iPeriod & _
''            " AND CP44=FC01||'0' AND NVL(CP116,'0')=NVL(FC03,'0')" & _
''            " and cp01||cp04='CFT00' and cp27+0>19221111 and cp57 is null and cp27+0<=" & stDate & _
''            " and instr('101',cp10)>0 and cp09<'B'" & _
''            " group by FC01,FC03"
''      End If
''   End If
''   '2013/5/23 End
'    '專利
'    tmpPA = stVTB1 & " From (SELECT DISTINCT R001 as FC01, " & IIf(m_bolByAgent = False, "R003", "''") & "  as FC03 FROM rdatafactory " & _
'       " Where Formname='" & Me.Name & "' And Id='" & strUserNum & "' and R006='" & strFC06_P & "' AND R004 = '" & (iYear - 1911) & "' And R005 = '" & iPeriod & "' " & _
'       " ) VTB1,caseprogress WHERE CP44=FC01||'0'" & IIf(m_bolByAgent = False, " AND NVL(CP116,'0')=NVL(FC03,'0')", "") & _
'       " and cp01||cp04='CFP00' and cp27+0>19221111 and cp57 is null and cp27+0<=" & stDate & _
'       " and instr('" & NewCasePtyList & "',cp10)>0 and cp09<'B'" & _
'       " group by FC01,FC03"
'    '商標
'    tmpTM = stVTB1 & " From (SELECT DISTINCT R001 as FC01, " & IIf(m_bolByAgent = False, "R003", "''") & "  as FC03 FROM rdatafactory " & _
'       " Where Formname='" & Me.Name & "' And Id='" & strUserNum & "' and R006='" & strFC06_T & "' AND R004 = '" & (iYear - 1911) & "' And R005 = '" & iPeriod & "' " & _
'       " ) VTB1,caseprogress WHERE CP44=FC01||'0'" & IIf(m_bolByAgent = False, " AND NVL(CP116,'0')=NVL(FC03,'0')", "") & _
'       " and cp01||cp04='CFT00' and cp27+0>19221111 and cp57 is null and cp27+0<=" & stDate & _
'       " and instr('101',cp10)>0 and cp09<'B'" & _
'       " group by FC01,FC03"
'    'end 2020/09/22
'
'   'FC案件統計
'   stVTB2 = "select FC01||FC03 B2,count(*) FC_TOT" & _
'      ",sum(decode(substr(cp05,1,4)," & iYear & "-2,1)) FC_L2" & _
'      ",sum(decode(substr(cp05,1,4)," & iYear & "-1,1)) FC_L1" & _
'      ",sum(decode(substr(cp05,1,4)," & iYear & ",decode(sign(substr(cp05,5,2)-6),1,0,1))) FC_C1" & _
'      ",sum(decode(substr(cp05,1,4)," & iYear & ",decode(sign(substr(cp05,5,2)-6),1,1,0))) FC_C2 "
'   'Add by Morgan 2008/6/24
'   If bExtra = True Then
'      stVTB2 = stVTB2 & _
'         ",sum(decode(sign(cp05-" & stDate1 & "+1),1,decode(sign(cp05-" & stDate2 & "+1),-1,1))) FC_X "
'   End If
'
'   'Modified by Lydia 2020/09/22 代理人互惠設定檔改用暫存；專利、商標案件語法分別暫存
''   If txtKind = "1" Then '專利
''      'Modify by Morgan 2008/7/22 加以代理人統計
''      If m_bolByAgent = True Then
''         'Modified by Lydia 2018/06/21 改新申請案性質instr('101,102,103,104,105,307',cp10)>0 => instr('" & NewCasePtyList & "',cp10)>0
''         stVTB2 = stVTB2 & " from (select DISTINCT FC01,'' FC03 from FAGENTCONFIG where FC06='" & strFC06 & "' AND FC04=" & (iYear - 1911) & " AND FC05=" & iPeriod & _
''            ") VTB2,patent,caseprogress" & _
''            " where pa75=FC01||'0' AND pa01||''='FCP'" & _
''            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp09<'B'" & _
''            " and instr('" & NewCasePtyList & "',cp10)>0 and cp57 is null and cp05+0<=" & stDate & _
''            " group by FC01,FC03"
''      Else
''         'Modified by Lydia 2018/06/21 改新申請案性質instr('101,102,103,104,105,307',cp10)>0 => instr('" & NewCasePtyList & "',cp10)>0
''         stVTB2 = stVTB2 & " from (select FC01,FC03 from FAGENTCONFIG where FC06='" & strFC06 & "' AND FC04=" & (iYear - 1911) & " AND FC05=" & iPeriod & _
''            ") X,patent,caseprogress" & _
''            " where pa75=FC01||'0' AND NVL(pa144,'0')=NVL(FC03,'0') AND pa01||''='FCP'" & _
''            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp09<'B'" & _
''            " and instr('" & NewCasePtyList & "',cp10)>0 and cp57 is null and cp05+0<=" & stDate & _
''            " group by FC01,FC03"
''      End If
''   'Add By Sindy 2013/5/23
''   ElseIf txtKind = "2" Then '商標
''      '加以代理人統計
''      If m_bolByAgent = True Then
''         stVTB2 = stVTB2 & " from (select DISTINCT FC01,'' FC03 from FAGENTCONFIG where FC06='" & strFC06 & "' AND FC04=" & (iYear - 1911) & " AND FC05=" & iPeriod & _
''            ") VTB2,trademark,caseprogress" & _
''            " where tm44=FC01||'0' AND tm01||''='FCT'" & _
''            " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04 and cp09<'B'" & _
''            " and instr('101',cp10)>0 and cp57 is null and cp05+0<=" & stDate & _
''            " group by FC01,FC03"
''      Else
''         stVTB2 = stVTB2 & " from (select FC01,FC03 from FAGENTCONFIG where FC06='" & strFC06 & "' AND FC04=" & (iYear - 1911) & " AND FC05=" & iPeriod & _
''            ") X,trademark,caseprogress" & _
''            " where tm44=FC01||'0' AND NVL(tm119,'0')=NVL(FC03,'0') AND tm01||''='FCT'" & _
''            " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04 and cp09<'B'" & _
''            " and instr('101',cp10)>0 and cp57 is null and cp05+0<=" & stDate & _
''            " group by FC01,FC03"
''      End If
''   End If
''   '2013/5/23 End
'    '專利
'    tmpPA2 = stVTB2 & " from (SELECT DISTINCT R001 as FC01, " & IIf(m_bolByAgent = False, "R003", "''") & "  as FC03 FROM rdatafactory " & _
'         " Where Formname='" & Me.Name & "' And Id='" & strUserNum & "' and R006='" & strFC06_P & "' AND R004 = '" & (iYear - 1911) & "' And R005 = '" & iPeriod & "' " & _
'         ") VTB2,patent,caseprogress" & _
'         " where pa75=FC01||'0' AND pa01||''='FCP'" & IIf(m_bolByAgent = False, " AND NVL(pa144,'0')=NVL(FC03,'0')", "") & _
'         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp09<'B'" & _
'         " and instr('" & NewCasePtyList & "',cp10)>0 and cp57 is null and cp05+0<=" & stDate & _
'         " group by FC01,FC03"
'    '商標
'    tmpTM2 = stVTB2 & " from(SELECT DISTINCT R001 as FC01, " & IIf(m_bolByAgent = False, "R003", "''") & "  as FC03 FROM rdatafactory " & _
'         " Where Formname='" & Me.Name & "' And Id='" & strUserNum & "' and R006='" & strFC06_T & "' AND R004 = '" & (iYear - 1911) & "' And R005 = '" & iPeriod & "' " & _
'         ") VTB2,trademark,caseprogress" & _
'         " where tm44=FC01||'0' AND tm01||''='FCT'" & IIf(m_bolByAgent = False, " AND NVL(tm119,'0')=NVL(FC03,'0')", "") & _
'         " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04 and cp09<'B'" & _
'         " and instr('101',cp10)>0 and cp57 is null and cp05+0<=" & stDate & _
'         " group by FC01,FC03"
'    'end 2020/09/22
'
'   'Modified by Lydia 2020/09/22 加上欄位名稱
''   strExc(0) = "select NVL(FA05,NVL(FA06,FA04)) C1,FC01||DECODE(FC03,NULL,'','-'||FC03) C2" & _
''      ",NVL(PCC03,NVL(PCC04,PCC05)) C3,NA03,nvl(FC_TOT,0),nvl(CF_TOT,0),nvl(FC_L2,0),nvl(CF_L2,0)" & _
''      ",nvl(FC_L1,0),nvl(CF_L1,0),nvl(FC_C1,0),nvl(CF_C1,0),nvl(FC_C1,0)-nvl(CF_C1,0),nvl(FC_C2,0)" & _
''      ",nvl(CF_C2,0),nvl(FC_C2,0)-nvl(CF_C2,0),FC07,FC08"
''   'Add by Morgan 2008/6/24
''   If bExtra = True Then
''      strExc(0) = strExc(0) & ",nvl(FC_X,0),nvl(CF_X,0),nvl(FC_X,0)-nvl(CF_X,0)"
''   End If
'   strExc(0) = "select NVL(FA05,NVL(FA06,FA04)) C1,FC01||DECODE(FC03,NULL,'','-'||FC03) C2" & _
'        ",NVL(PCC03,NVL(PCC04,PCC05)) C3,NA03,nvl(FC_TOT,0) FC_TOT,nvl(CF_TOT,0) CF_TOT,nvl(FC_L2,0) FC_L2,nvl(CF_L2,0) CF_L2 " & _
'        ",nvl(FC_L1,0) FC_L1,nvl(CF_L1,0) CF_L1,nvl(FC_C1,0) FC_C1,nvl(CF_C1,0) CF_C1,nvl(FC_C1,0)-nvl(CF_C1,0) DIFF01,nvl(FC_C2,0) FC_C2" & _
'        ",nvl(CF_C2,0) CF_C2,nvl(FC_C2,0)-nvl(CF_C2,0) DIFF02,FC07,FC08 "
'
'   'Add by Morgan 2008/6/24
'   If bExtra = True Then
'      strExc(0) = strExc(0) & ",nvl(FC_X,0) FC_x,nvl(CF_X,0) CF_X,nvl(FC_X,0)-nvl(CF_X,0) DIFF03"
'   End If
'   'end 2020/09/22
'
'   'Modify by Morgan 2008/7/22 加以代理人統計
'   'Modified by Lydia 2020/09/22 分別存放專利、商標SQL語法
''   If m_bolByAgent = True Then
''      strExc(1) = strExc(0) & " from (SELECT FC01,'' FC03,'' PCC03,'' PCC04,'' PCC05,SUM(FC07) FC07,MIN(FC08) FC08" & _
''         " FROM FAGENTCONFIG where FC06='" & strFC06 & "' AND FC04=" & (iYear - 1911) & " AND FC05=" & iPeriod & _
''         " GROUP BY FC01) X,(" & stVTB1 & ") A,(" & stVTB2 & ") B,FAGENT,NATION where A2(+)=FC01||FC03" & _
''         " and B2(+)=FC01||FC03 AND FA01(+)=FC01 AND FA02(+)='0' AND NA01(+)=SUBSTR(FA10,1,3)" & _
''         " ORDER BY NA01 ASC,FC07 DESC,C1,FC01 ASC"
''   Else
''      strExc(1) = strExc(0) & " from FAGENTCONFIG,(" & stVTB1 & ") A,(" & stVTB2 & ") B,FAGENT,NATION,POTCUSTCONT" & _
''         " where FC06='" & strFC06 & "' AND FC04=" & (iYear - 1911) & " AND FC05=" & iPeriod & " and A2(+)=FC01||FC03" & _
''         " and B2(+)=FC01||FC03 AND FA01(+)=FC01 AND FA02(+)='0' AND NA01(+)=SUBSTR(FA10,1,3) AND PCC01(+)=FC01 AND PCC02(+)=FC03" & _
''         " ORDER BY NA01 ASC,FC07 DESC,C1,FC01 ASC,FC03 ASC"
''   End If
'   For intI = 1 To 2
'        If m_bolByAgent = True Then
'           'Modified by Lydia 2022/05/12 +提出年度、提出部門、提出人員
'           'strExc(intI) = strExc(0) & " from (SELECT R001 as FC01,'' FC03,'' PCC03,'' PCC04,'' PCC05,SUM(R007) FC07,MIN(R008) FC08,R009 as FC16" & _
'              " FROM Rdatafactory where Formname='" & Me.Name & "' And Id='" & strUserNum & "' and R006='" & IIf(intI = 1, strFC06_P, strFC06_T) & "' AND R004 = '" & (iYear - 1911) & "' And R005 = '" & iPeriod & "' " & _
'              " GROUP BY R001) X,(" & IIf(intI = 1, tmpPA, tmpTM) & ") A,(" & IIf(intI = 1, tmpPA2, tmpTM2) & ") B,FAGENT,NATION where A2(+)=FC01||FC03" & _
'              " and B2(+)=FC01||FC03 AND FA01(+)=FC01 AND FA02(+)='0' AND NA01(+)=SUBSTR(FA10,1,3)" & _
'              " ORDER BY NA01 ASC,FC07 DESC,C1,FC01 ASC"
'           strExc(intI) = strExc(0) & ",FC16,A0902 AS FC17DEPT,ST02 AS FC17NAME from (SELECT R001 as FC01,'' FC03,'' PCC03,'' PCC04,'' PCC05,SUM(R007) FC07,MIN(R008) FC08,R009 as FC16, R010 as FC17" & _
'              " FROM Rdatafactory where Formname='" & Me.Name & "' And Id='" & strUserNum & "' and R006='" & IIf(intI = 1, strFC06_P, strFC06_T) & "' AND R004 = '" & (iYear - 1911) & "' And R005 = '" & iPeriod & "' " & _
'              " GROUP BY R001, R009, R010) X,(" & IIf(intI = 1, tmpPA, tmpTM) & ") A,(" & IIf(intI = 1, tmpPA2, tmpTM2) & ") B,FAGENT,NATION,STAFF,ACC090 where A2(+)=FC01||FC03" & _
'              " and B2(+)=FC01||FC03 AND FA01(+)=FC01 AND FA02(+)='0' AND NA01(+)=SUBSTR(FA10,1,3) AND FC17=ST01(+) AND ST03=A0901(+)" & _
'              " ORDER BY NA01 ASC,FC07 DESC,C1,FC01 ASC"
'        Else
'           'Modified by Lydia 2022/05/12 +提出年度、提出部門、提出人員
'           'strExc(intI) = strExc(0) & " from (SELECT R001 AS FC01, R002 AS FC02, R003 AS FC03,R007 AS FC07, R008 AS FC08 from rdatafactory " & _
'               " where Formname='" & Me.Name & "' And Id='" & strUserNum & "' and R006='" & IIf(intI = 1, strFC06_P, strFC06_T) & "' AND R004 = '" & (iYear - 1911) & "' And R005 = '" & iPeriod & "' ) X, " & _
'               " (" & IIf(intI = 1, tmpPA, tmpTM) & ") A,(" & IIf(intI = 1, tmpPA2, tmpTM2) & ") B,FAGENT,NATION,POTCUSTCONT" & _
'              " WHERE A2(+)=FC01||FC03 AND B2(+)=FC01||FC03 AND FA01(+)=FC01 AND FA02(+)='0' AND NA01(+)=SUBSTR(FA10,1,3) AND PCC01(+)=FC01 AND PCC02(+)=FC03" & _
'              " ORDER BY NA01 ASC,FC07 DESC,C1,FC01 ASC,FC03 ASC"
'           strExc(intI) = strExc(0) & ",FC16,A0902 AS FC17DEPT,ST02 AS FC17NAME from (SELECT R001 AS FC01, R002 AS FC02, R003 AS FC03,R007 AS FC07, R008 AS FC08,R009 as FC16, R010 as FC17 from rdatafactory " & _
'               " where Formname='" & Me.Name & "' And Id='" & strUserNum & "' and R006='" & IIf(intI = 1, strFC06_P, strFC06_T) & "' AND R004 = '" & (iYear - 1911) & "' And R005 = '" & iPeriod & "' ) X, " & _
'               " (" & IIf(intI = 1, tmpPA, tmpTM) & ") A,(" & IIf(intI = 1, tmpPA2, tmpTM2) & ") B,FAGENT,NATION,POTCUSTCONT,STAFF,ACC090" & _
'              " WHERE A2(+)=FC01||FC03 AND B2(+)=FC01||FC03 AND FA01(+)=FC01 AND FA02(+)='0' AND NA01(+)=SUBSTR(FA10,1,3) AND PCC01(+)=FC01 AND PCC02(+)=FC03 AND FC17=ST01(+) AND ST03=A0901(+)" & _
'              " ORDER BY NA01 ASC,FC07 DESC,C1,FC01 ASC,FC03 ASC"
'        End If
'   Next intI
'
'   'Mark by Lydia 2020/12/22
''   strExc(1) = Replace(UCase(strExc(1)), "FC07,FC08 ", "FC07,FC08,NA01,FC01 ")
''   strExc(2) = Replace(UCase(strExc(2)), "FC07,FC08 ", "FC07,FC08,NA01,FC01 ")
''   strExc(1) = "SELECT * FROM (" & strExc(1) & " ) GROUP BY C1,C2,C3,NA03,FC_TOT,CF_TOT,FC_L2,CF_L2,FC_L1,CF_L1,FC_C1,CF_C1,DIFF01,FC_C2,CF_C2,DIFF02,FC07,FC08,NA01,FC01,FC07 ORDER BY NA01 ASC,FC01 ASC,FC07 DESC"
''   strExc(2) = "SELECT * FROM (" & strExc(2) & " ) GROUP BY C1,C2,C3,NA03,FC_TOT,CF_TOT,FC_L2,CF_L2,FC_L1,CF_L1,FC_C1,CF_C1,DIFF01,FC_C2,CF_C2,DIFF02,FC07,FC08,NA01,FC01,FC07 ORDER BY NA01 ASC,FC01 ASC,FC07 DESC"
'   'end 2020/12/22
'   If txtKind = "1" Then
'       strConFirst = strExc(1): strConSecond = strExc(2)
'   Else
'       strConFirst = strExc(2): strConSecond = strExc(1)
'   End If
'   'end 2020/09/22
   'Modified by Lydia 2023/10/06 +bolExcel=False
   Call Pub_GetSqlfrm050408(strUserNum, Me.txtKind, "0", Me.txtYear, "" & iPeriod, m_bolByAgent, strConFirst, strConSecond, strSrvDate(1), Me.Text1, Me.Text2, False)
   'end 2023/09/12
   
   intI = 1
   
   'Added by Lydia 2025/07/30 查詢印表記錄檔欄位
   ClearQueryLog (Me.Name)
   m_strQL05 = ""
   m_strQL05 = m_strQL05 & ""
   m_strQL05 = m_strQL05 & ";統計年度:" & txtYear & "年"
   m_strQL05 = m_strQL05 & ";統計區間:" & cboPeriod.Text
   m_strQL05 = m_strQL05 & ";統計對象:" & cboTarget.Text
   m_strQL05 = m_strQL05 & ";案件類別:" & IIf(txtKind = "1", "專利", "商標")
   If bExtra = True Then
      m_strQL05 = m_strQL05 & ";指定日期區間:" & Text1 & " － " & Text2
   End If
   pub_QL05 = m_strQL05
   'end 2025/07/30
   
   'Modified by Lydia 2020/09/22 strExc(1) =>strConFirst
   Set RsTemp = ClsLawReadRstMsg(intI, strConFirst)
   If intI = 1 Then
      'Added by Lydia 2020/09/22 去掉發文的SQL
      'Mark by Lydia 2023/09/12 併入模組Pub_GetSqlfrm050408
      'If bExtra = True Then
      '    strConFirst = Replace(Replace(Replace(strConFirst, ",sum(decode(sign(cp27-" & stDate1 & "+1),1,decode(sign(cp27-" & stDate2 & "-1),-1,1))) CF_X", ""), ",sum(decode(sign(cp05-" & stDate1 & "+1),1,decode(sign(cp05-" & stDate2 & "+1),-1,1))) FC_X", ""), ",nvl(FC_X,0) FC_x,nvl(CF_X,0) CF_X,nvl(FC_X,0)-nvl(CF_X,0) DIFF03", "")
      '    strConSecond = Replace(Replace(Replace(strConSecond, ",sum(decode(sign(cp27-" & stDate1 & "+1),1,decode(sign(cp27-" & stDate2 & "-1),-1,1))) CF_X", ""), ",sum(decode(sign(cp05-" & stDate1 & "+1),1,decode(sign(cp05-" & stDate2 & "+1),-1,1))) FC_X", ""), ",nvl(FC_X,0) FC_x,nvl(CF_X,0) CF_X,nvl(FC_X,0)-nvl(CF_X,0) DIFF03", "")
      'End If
      ''end 2020/09/22
      'end 2023/09/12
      InsertQueryLog (RsTemp.RecordCount) 'Added by Lydia 2025/07/30
      With frm050408_1
         .Show
         .Caption = Me.Caption & "(" & txtYear & "年" & cboPeriod & ")"
         .txt1(0) = txtYear & "0101"
         .txt1(1) = IIf(cboPeriod.ListIndex = 0, txtYear & "0630", txtYear & "1231")
         'Add by Morgan 2008/6/24
         If bExtra = True Then
            .lblExtra = "指定日期區間：" & Text1 & " － " & Text2
            .m_stDate1 = stDate1
            .m_stDate2 = stDate2
            .cmdok(4).Visible = True
            .cmdok(5).Visible = True
         End If
         Set .m_adoRst = RsTemp
         .InitGrid
         If m_bPrint = False Then
            .cmdok(2).Visible = False
            .cmdok(3).Visible = False
            .cmdok(4).Visible = False
            .cmdok(5).Visible = False
         End If
      End With
      Me.Hide
   Else
      InsertQueryLog (0) 'Added by Lydia 2025/07/30
      MsgBox "無資料！"
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   'Modify by Morgan 2008/7/22 改輸入統計年度&區間
   If txtYear = "" Then
      MsgBox "統計年度不可空白！"
      txtYear.SetFocus
      Exit Function
   End If
   txtYear_Validate bCancel
   If bCancel = True Then
      txtYear_GotFocus
      txtYear.SetFocus
      Exit Function
   End If
   'end 2008/7/22
   Text1_Validate bCancel
   If bCancel = True Then
      Text1_GotFocus
      Text1.SetFocus
      Exit Function
   End If
   Text2_Validate bCancel
   If bCancel = True Then
      Text2_GotFocus
      Text2.SetFocus
      Exit Function
   End If
   
   If Text1 <> "" Or Text2 <> "" Then
      If Text1 = "" Then
         MsgBox "請輸入日期區間起日！"
         Text1_GotFocus
         Text1.SetFocus
         Exit Function
      End If
      If Text2 = "" Then
         MsgBox "請輸入日期區間迄日！"
         Text2_GotFocus
         Text2.SetFocus
         Exit Function
      End If
      If Val(Text2) < Val(Text1) Then
         MsgBox "日期區間輸入錯誤！"
         Text1_GotFocus
         Text1.SetFocus
         Exit Function
      End If
   End If
   
   'Add By Sindy 2013/5/23
   If txtKind = "" Then
      MsgBox "案件類別不可空白！"
      txtKind.SetFocus
      Exit Function
   End If
   txtKind_Validate bCancel
   If bCancel = True Then
      txtkind_GotFocus
      txtKind.SetFocus
      Exit Function
   End If
   '2013/5/23 End
   
   TxtValidate = True
End Function

Private Sub Form_Load()
   'Add by Morgan 2008/8/15 加權限控管
   m_bQuery = IsUserHasRightOfFunction("frm050408", strFind, False)
   m_bPrint = IsUserHasRightOfFunction("frm050408", strPrint, False)
   
   'Added by Lydia 2025/06/06
   If m_AppNo <> "" Then
      Me.Caption = "互惠期間統計表"
      Me.Height = 4200
      lblAppNo(0).Visible = True
      lblAppNo(1).Visible = True
      lblAppNo(0).Caption = "代理人編號：" & m_AppNo
      lblAppNo(1).Caption = GetFAgentName(m_AppNo)
      Me.Tag = m_AppNo
   Else
      Me.Height = 3540
      lblAppNo(0).Visible = False
      lblAppNo(1).Visible = False
      Me.Tag = ""
   End If
   'end 2025/06/06
   
   MoveFormToCenter Me
   cboPeriod.Clear
   cboPeriod.AddItem "上半年", 0
   cboPeriod.AddItem "下半年", 1
   txtYear = strSrvDate(2) \ 10000
   If Mid(strSrvDate(1), 5, 2) < 7 Then
      cboPeriod.ListIndex = 0
   Else
      cboPeriod.ListIndex = 1
   End If
   cboTarget.Clear
   cboTarget.AddItem "代理人", 0
   cboTarget.AddItem "聯絡人", 1
   cboTarget.ListIndex = 0
   
   'Added by Lydia 2025/06/27
   If Pub_StrUserSt03 = "M51" Then
      cmdStatistic.Visible = True
   End If
   'end 2025/06/27
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   'Added by Lydia 2025/06/06
   If TypeName(m_PrevForm) <> "Nothing" Then
      m_PrevForm.Show
   End If
   'end 2025/06/06
   
   Set frm050408 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And (KeyAscii > Asc(9) Or KeyAscii < Asc(0)) Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "" Then
      If ChkDate(Text1) = False Then
         Cancel = True
      End If
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And (KeyAscii > Asc(9) Or KeyAscii < Asc(0)) Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 <> "" Then
      If ChkDate(Text2) = False Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtYear_GotFocus()
   TextInverse txtYear
End Sub

Private Sub txtYear_Validate(Cancel As Boolean)
   If txtYear <> "" Then
      If Not IsNumeric(txtYear) Then
         MsgBox "年度輸入錯誤！"
         txtYear.SetFocus
         Cancel = True
      End If
   End If
End Sub

'Add By Sindy 2013/5/23
Private Sub txtkind_GotFocus()
   TextInverse txtKind
End Sub

'Add By Sindy 2013/5/23
Private Sub txtKind_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
      KeyAscii = 0
   End If
End Sub

'Add By Sindy 2013/5/23
Private Sub txtKind_Validate(Cancel As Boolean)
   If txtKind <> "" Then
      If Not IsNumeric(txtKind) Then
         MsgBox "案件類別輸入錯誤！"
         txtKind.SetFocus
         Cancel = True
      End If
      If txtKind <> "1" And txtKind <> "2" Then
         MsgBox "案件類別只可輸入1或2！"
         txtKind.SetFocus
         Cancel = True
      End If
   Else
      MsgBox "請輸入案件類別！", vbCritical
      txtKind.SetFocus
      Cancel = True
   End If
End Sub

'Added by Lydia 2025/06/27 互惠代理人案件盈虧統計表：目前只開放給電腦中心使用
Private Sub cmdStatistic_Click()
Dim xlsAgentPoint As New Excel.Application
Dim wksrpt As New Worksheet
Dim xlsFileName1 As String
Dim stTmp, intWidth
Dim tmpArr1 As Variant
Dim intCounter As Integer, intB As Integer

   If TxtValidate Then
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      '附件名稱
      'Modify By Sindy 2025/7/14
'      xlsFileName1 = "互惠代理人案件盈虧統計表-當年度1-6月"
      xlsFileName1 = IIf(txtKind = "2", "(商標)", "專利") & "互惠代理人案件盈虧統計表-當年度1-6月"
      '2025/7/14 END
      cnnConnection.Execute "delete from rdatafactory where FORMNAME like 'frm050408_2%' and ID=" & CNULL(strUserNum)
      '＊＊若表單frm050408_1的欄位有變動，呼叫Pub_Frm050408_GetStatistic也要變動＊＊
      'Modify By Sindy 2025/7/14
      If txtKind = "2" Then '商標
         '當年1-6月FCT => 10
         Call Pub_Frm050408_GetStatistic(txtKind, txtYear, cboPeriod.ListIndex + 1, IIf(cboTarget.ListIndex = 0, True, False), "FCT", "CFT", 10, "ALL")
         '當年1-6月CFT => 11
         Call Pub_Frm050408_GetStatistic(txtKind, txtYear, cboPeriod.ListIndex + 1, IIf(cboTarget.ListIndex = 0, True, False), "FCT", "CFT", 11, "ALL")
      Else
      '2025/7/14 END
         '當年1-6月FCP => 10
         Call Pub_Frm050408_GetStatistic(txtKind, txtYear, cboPeriod.ListIndex + 1, IIf(cboTarget.ListIndex = 0, True, False), "FCP", "CFP", 10, "ALL")
         '當年1-6月CFP => 11
         Call Pub_Frm050408_GetStatistic(txtKind, txtYear, cboPeriod.ListIndex + 1, IIf(cboTarget.ListIndex = 0, True, False), "FCP", "CFP", 11, "ALL")
      End If
      'R004欄位長度500，專門放案件名稱
      strExc(0) = "SELECT substr(na01,1,3)||na03 AS 代理人國籍,rtrim(decode(fa05,NULL,nvl(fa04,fa06),fa05||' '||fa63||' '||fa64||' '||fa65)) AS 代理人名稱" & _
                  " ,fa01||fa02 as 代理人編號 ,r001 as 本所案號,r004 as 案件名稱,r003 as 案件盈虧 FROM rdatafactory,fagent,nation" & _
                  " WHERE formname like 'frm050408_2%' AND ID='" & strUserNum & "' AND substr(r010,1,8)=fa01(+) AND substr(r010,9,1)=fa02(+) AND fa10=na01(+)"
      '排除沒有案件盈虧 and r003 <> '.00'
      strExc(0) = strExc(0) & " ORDER BY 1,r010,r001"
      intI = 0
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         stTmp = Array("代理人國籍", "代理人名稱", "代理人編號", "本所案號", "案件名稱", "案件盈虧")
         intWidth = Array(11, 35, 11, 12, 35, 11)
         ReDim tmpArr1(0 To UBound(stTmp))
         intCounter = 1
         xlsAgentPoint.SheetsInNewWorkbook = 1 '預設工作表數目
         xlsAgentPoint.Workbooks.add
         xlsAgentPoint.Application.WindowState = xlMinimized
         xlsAgentPoint.Application.Visible = False
         Set wksrpt = xlsAgentPoint.Worksheets(1)
         xlsAgentPoint.Sheets(1).Select

         '設定欄位名稱及欄寬
         For intB = LBound(stTmp) To UBound(stTmp)
             wksrpt.Range(Chr(intB + 65) & intCounter).Value = stTmp(intB)
             wksrpt.Columns(Chr(intB + 65) & ":" & Chr(intB + 65)).ColumnWidth = intWidth(intB)
             wksrpt.Range(Chr(intB + 65) & intCounter).HorizontalAlignment = xlCenter
         Next intB
         
         intCounter = 2
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            For intB = 0 To UBound(tmpArr1)
               tmpArr1(intB) = "" & RsTemp.Fields(intB)
            Next intB
            wksrpt.Range("A" & intCounter & ":" & Chr(UBound(stTmp) + 65) & intCounter).Value = tmpArr1
            intCounter = intCounter + 1
            RsTemp.MoveNext
         Loop
         wksrpt.Range(Chr(UBound(stTmp) + 65) & ":" & Chr(UBound(stTmp) + 65)).HorizontalAlignment = xlRight
         xlsAgentPoint.ActiveWindow.WindowState = xlMaximized  '讓下面的ActiveWindow.FreezePanes可使用
         xlsAgentPoint.ActiveSheet.Range("A2").Select
         xlsAgentPoint.ActiveWindow.FreezePanes = True '凍結窗格
         
         xlsFileName1 = xlsFileName1 & "_" & strSrvDate(1) & Format(ServerTime, "000000") & MsgText(43)
         If Val(xlsAgentPoint.Version) < 12 Then
            xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName1, FileFormat:=-4143
         Else
            xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName1, FileFormat:=56
         End If
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         MsgBox "Excel檔案產生完成！" & vbCrLf & "檔案位置：" & strExcelPathN & " " & xlsFileName1
         xlsAgentPoint.Workbooks.Close
         xlsAgentPoint.Quit
         Set xlsAgentPoint = Nothing
         Set wksrpt = Nothing
      End If
   End If
   
   Exit Sub
   
ErrHandle:
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   If Err.Number <> 0 Then
      xlsAgentPoint.Workbooks.Close
      xlsAgentPoint.Quit
      Set xlsAgentPoint = Nothing
      Set wksrpt = Nothing
   End If
End Sub

