VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210103 
   BorderStyle     =   1  '單線固定
   Caption         =   "每日點數輸入"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9432
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9432
   Begin VB.TextBox txtSalesArea 
      Height          =   285
      Left            =   810
      MaxLength       =   3
      TabIndex        =   2
      Top             =   576
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Left            =   4455
      MaxLength       =   7
      TabIndex        =   1
      Top             =   576
      Width           =   1140
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  '平面
      Height          =   375
      Left            =   315
      TabIndex        =   0
      Text            =   "Text3"
      Top             =   3000
      Width           =   1635
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3400
      Left            =   24
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   876
      Width           =   9372
      _ExtentX        =   16531
      _ExtentY        =   6011
      _Version        =   393216
      BackColor       =   -2147483624
      Rows            =   3
      Cols            =   7
      FixedRows       =   2
      FixedCols       =   0
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6930
      Top             =   30
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210103.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210103.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210103.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210103.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210103.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210103.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210103.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210103.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210103.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210103.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210103.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   528
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9432
      _ExtentX        =   16637
      _ExtentY        =   931
      ButtonWidth     =   1076
      ButtonHeight    =   882
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "新增"
            Key             =   "keyInsert"
            Object.Tag             =   "F2"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "輸入"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   $"frm210103.frx":20F4
      Height          =   1440
      Left            =   105
      TabIndex        =   7
      Top             =   4305
      Width           =   9135
   End
   Begin VB.Label lblSalesArea 
      Height          =   180
      Left            =   1770
      TabIndex        =   6
      Top             =   660
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "業務區"
      Height          =   180
      Left            =   135
      TabIndex        =   5
      Top             =   660
      Width           =   540
   End
   Begin VB.Label Label2 
      Caption         =   "點數結算日"
      Height          =   180
      Left            =   3420
      TabIndex        =   4
      Top             =   660
      Width           =   900
   End
End
Attribute VB_Name = "frm210103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/20 改成Form2.0 (grdDataList)
'Memo by Lydia 2021/07/27 表單名稱「每日業績點數輸入」=>更名為「每日點數輸入」'Memo by Lydia 2021/08/27 上線
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
'2005/7/5整理
Option Explicit

Dim iRow As Integer '本次點選列數
Dim iCol As Integer '智權人員名稱欄位
Dim iLstRow As Integer '前次點選列數
Dim ii As Integer, jj As Integer '迴圈共用
Dim stSQL As String
Dim iCurState As Integer '目前狀態
Dim stET03s As String '有權限業務區
Dim MaxDF02 As String, MinDF02 As String 'Add by Amy 2021/02/04
Dim m_PrevForm As Form 'Added by Lydia 2021/07/27 前一畫面

'Added by Lydia 2021/07/27 外部呼叫使用
Public Sub SetParent(ByVal pForm As Form)
   Set m_PrevForm = pForm
End Sub

Private Function doQuery() As Boolean
   Dim stDate0 As String, stDDate As String
   Dim stCon As String
   Dim stVTable1 As String, stVTable2 As String, stVTable3 As String, stVTable4 As String, stVTable5 As String
   Dim strField As String 'Add by Amy 2018/11/28
   Dim strTime As String

   strTime = ServerTime
On Error GoTo ErrHnd

   stDDate = Val(txtCloseDate)
   stDate0 = Val(stDDate) \ 100 & "01"
   
   'Modify by Morgan 2010/3/15 已離職當月分仍要顯示,否則加總會錯
   'Modify by Morgan 2010/4/2 改判斷離職前一天的月份要顯示(離職日為1號則當月不用)
   'stCon = " AND (st04='1' OR ST51> " & DBDATE(stDDate) & ") and st15='" & txtSalesArea & "'"
   'Modify by Amy 2018/11/28 10801開始國外部分為FCP及FCT拆出
   stCon = " AND (st04='1' OR substr( decode(ST51,null,'',to_char(to_date(st51,'yyyymmdd')-1,'yyyymmdd')) ,1,6)>= " & Left(DBDATE(stDDate), 6) & ") "
   If (Val(stDDate) > Val(每日業務點數FCPFCT啟用日) And txtSalesArea <> "F41") Or Val(stDDate) < Val(每日業務點數FCPFCT啟用日) Then
      stCon = stCon & "and st15='" & txtSalesArea & "' "
   End If
   'Add by Morgan 2005/4/25 加國外部
   If txtSalesArea = "F41" Then
      'Add by Amy 2018/11/28 10801開始國外部分為FCP及FCT拆出
      If Val(stDDate) > Val(每日業務點數FCPFCT啟用日) Then
        stCon = stCon & " and st01 In ('F4102','F4103')"
      Else
        stCon = stCon & " and st01='F4100'"
      End If
   'Mark by Amy 2024/10/01 林柄佑協理(82026) st15改為s29後 11309月,有收入導致不會出現,原S29員編目前只用於目標設定
   '                        20091-目前也都沒在輸每日點數,故抓法改與S部門相同-秀玲
   '2012/8/23 add by sonia 再加S29
'   ElseIf txtSalesArea = "S29" Then
'      stCon = stCon & " and st01='S29'"
   '2012/8/23 END
   'Added by Morgan 2020/9/22 客服組會用 W1001
   ElseIf txtSalesArea = "W10" Then
   'end 2020/9/22
   Else
      '2011/4/1 MODIFY BY SONIA
      'stCon = stCon & " and st01>'60000' and st01<'999999'"
      'modify by sonia 2024/8/6 +30015
      'stCon = stCon & " and st01>'60000' and st01<'F'"
      stCon = stCon & "and (st01='30015' or st01>'60000' and st01<'F')"
   End If
   
   '已入帳資料 國外部要抓所有部門為F41的員工 2009/8/3因F4102(F21),F4103(F11)改部門故同時抓F21,F31部門員工資料
   '2014/1/21 modif by sonia 取消a0201='1'條件
   'modify by sonia 2015/4/23 加不含'4194'科目,不含結餘傳票不要加科目限制
   'stVTable1 = "select DECODE(ST15,'F41','F4100','F21','F4100','F11','F4100',ax209) V1C0,ROUND(sum(decode(a0205," & stDDate & ", ax207))/1000,2) V1C1" & _
      ",ROUND(sum(ax207)/1000,2) V1C2" & _
      " From acc020, acc021,STAFF" & _
      " Where a0205 >= " & stDate0 & " And a0205 <= " & stDDate & _
      " and ax201(+) = a0201  and ax202(+) = a0202" & _
      " and ax209 Is Not Null and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
      " and ax207>0 and not( ax205='4191' or ax205='4192'" & _
      " or ((ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0))" & _
      " AND ST01(+)=AX209" & _
      " GROUP BY DECODE(ST15,'F41','F4100','F21','F4100','F11','F4100',AX209)"
   'Modify by Amy 2018/11/28 10801開始國外部分為FCP及FCT拆出
   strField = "DECODE(ST15,'F41','F4100','F21','F4100','F11','F4100',ax209)"
   If Val(stDDate) > Val(每日業務點數FCPFCT啟用日) Then
      strField = "ax209"
   End If
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   stVTable1 = "select " & strField & " V1C0,ROUND(sum(decode(a0205," & stDDate & ", ax207))/1000,2) V1C1" & _
      ",ROUND(sum(ax207)/1000,2) V1C2" & _
      " From acc020, acc021,STAFF" & _
      " Where a0205 >= " & stDate0 & " And a0205 <= " & stDDate & _
      " and ax201(+) = a0201  and ax202(+) = a0202" & _
      " and ax209 Is Not Null and (substr(ax205, 1, 1) = '4' Or ax205 = '7121')" & _
      " and ax207>0 and not( ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
      " AND ST01(+)=AX209" & _
      " GROUP BY " & strField
   'end 2018/11/28
      
   '已收款資料
   stVTable2 = "select df01 V2C0,sum(decode(df02," & stDDate & ",0,df03)) V2C1" & _
      ",sum(decode(df02," & stDDate & ",df03,0)) V2C2" & _
      ",decode(Max(df02)," & stDDate & ",'*',null) V2C3 From STAFF,DailyFeat" & _
      " Where DF01(+)=ST01 AND DF02>=" & stDate0 & " AND df02 <= " & stDDate & stCon & " group by df01"
      
   '已簽約資料
   stVTable3 = "select df01 V3C0,sum(decode(df02," & stDDate & ",0,NVL(df04,0)-NVL(DF03,0))) V3C1" & _
      ",sum(decode(df02," & stDDate & ",df04,0)) V3C2" & _
      ",decode(Max(df02)," & stDDate & ",'*',null) V3C3 From STAFF,DailyFeat" & _
      " Where DF01(+)=ST01 AND df02 <= " & stDDate & stCon & " group by df01"
      
   '銷帳資料 94/6以後才要 國外部要抓所有部門為F41的員工 2009/8/3因F4102(F21),F4103(F11)改部門故同時抓F21,F31部門員工資料
   'Modify by Amy 2018/11/28 10801開始國外部分為FCP及FCT拆出
   strField = "DECODE(ST15,'F41','F4100','F21','F4100','F11','F4100',A0K20)"
   If Val(stDDate) > Val(每日業務點數FCPFCT啟用日) Then
      strField = "A0K20"
   End If
   stVTable4 = "SELECT " & strField & " V4C0,ROUND(SUM(DECODE(A0S03," & stDDate & ",0,A1U07))/1000,2) V4C1" & _
      ",ROUND(SUM(DECODE(A0S03," & stDDate & ",A1U07,0))/1000,2) V4C2" & _
      " From ACC0S0, ACC0K0, ACC1U0, STAFF" & _
      " WHERE A0S04='1' AND A0S03>=940601 AND A0S03 <= " & stDDate & _
      " AND A0K01(+)=A0S02 AND A0K20 IS NOT NULL AND A1U01(+)=A0S01 AND ST01(+)=A0K20" & _
      " GROUP BY " & strField
   'end 2018/11/28
   
   'Modify by Morgan 2005/7/29 國外部要抓所有部門為F41的員工 2009/8/3因F4102(F21),F4103(F11)改部門故同時抓F21,F31部門員工資料
   '扣點數資料 94/6以後才要
   'Modify by Morgan 2005/9/2 條件改與點數一致,只差抓借方>0
   '2014/1/21 modif by sonia 取消a0201='1'條件
   'modify by sonia 2015/4/23 加不含'4194'科目,不含結餘傳票不要加科目限制
   'stVTable5 = "select DECODE(ST15,'F41','F4100','F21','F4100','F11','F4100',ax209) V5C0,ROUND(sum(DECODE(A0205," & stDDate & ",0,ax206))/1000,2) V5C1" & _
      ",ROUND(sum(DECODE(A0205," & stDDate & ",ax206,0))/1000,2) V5C2" & _
      " from acc020, acc021,STAFF where ax201(+) = a0201  and ax202(+) = a0202" & _
      " AND A0205>=940601 AND A0205>=" & stDate0 & " AND A0205<=" & stDDate & _
      " and ax209 is not null and ax206>0 and (substr(ax205, 1, 2) = '41' Or ax205 = '7121')" & _
      " and not (  (ax205='4191' or ax205='4192' or ax205='4194')" & _
      " or ( (ax205='410103' or ax205='411103' or ax205='4121' or ax205='4131') and instr(ax213||' ','結餘')>0 ) )" & _
      " AND ST01(+)=AX209" & _
      " group by DECODE(ST15,'F41','F4100','F21','F4100','F11','F4100',ax209)"
   'Modify by Amy 2018/11/28 10801開始國外部分為FCP及FCT拆出
   strField = "DECODE(ST15,'F41','F4100','F21','F4100','F11','F4100',ax209)"
   If Val(stDDate) > Val(每日業務點數FCPFCT啟用日) Then
      strField = "ax209"
   End If
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101  原:substr(ax205, 1, 2) = '41'
   stVTable5 = "select " & strField & " V5C0,ROUND(sum(DECODE(A0205," & stDDate & ",0,ax206))/1000,2) V5C1" & _
      ",ROUND(sum(DECODE(A0205," & stDDate & ",ax206,0))/1000,2) V5C2" & _
      " from acc020, acc021,STAFF where ax201(+) = a0201  and ax202(+) = a0202" & _
      " AND A0205>=940601 AND A0205>=" & stDate0 & " AND A0205<=" & stDDate & _
      " and ax209 is not null and ax206>0 and (substr(ax205, 1, 1) = '4' Or ax205 = '7121')" & _
      " and not (ax205='4191' or ax205='4192' or ax205='4194' or instr(ax213||' ','結餘')>0)" & _
      " AND ST01(+)=AX209" & _
      " group by " & strField
   
   stSQL = "select ST02||decode(st51,null,'',decode(sign(st51-" & DBDATE(stDDate) & "),1,'','*')) 智權人員" & _
      ",NVL(V1C1,0)-NVL(V5C2,0) 本日入賬,NVL(V1C2,0)-NVL(V5C1,0)-NVL(V5C2,0) 本月入賬" & _
      ",NVL(V2C1,0)-NVL(V5C1,0) 前日收款,NVL(V2C2,0) 本日收款,NVL(V5C2,0) 扣點數" & _
      ",NVL(V2C1,0)-NVL(V5C1,0)+NVL(V2C2,0)-NVL(V5C2,0) 本日收款累計" & _
      ",NVL(V3C1,0)-NVL(V4C1,0) 前日簽約未收,NVL(V3C2,0) 本日簽約,NVL(V4C2,0) 本日銷帳" & _
      ",NVL(V3C1,0)-NVL(V4C1,0)+NVL(V3C2,0)-NVL(V2C2,0)-NVL(V4C2,0) 簽約未收累計" & _
      ",ST01 C11,V2C3 C12" & _
      " from staff,(" & stVTable1 & ") VT1,(" & stVTable2 & ") VT2,(" & stVTable3 & ") VT3" & _
      ",(" & stVTable4 & ") VT4,(" & stVTable5 & ") VT5" & _
      " where V1C0(+)=st01 and V2C0(+)=ST01 and V3C0(+)=ST01 AND V4C0(+)=ST01 and V5C0(+)=ST01" & stCon
      
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         If SetRow(AdoRecordSet3, stDate0, stDDate) = True Then
            doQuery = True
         End If
      Else
         MsgBox "查無資料！", vbInformation
         SetDataListWidth
      End If
   End With
   
   'MsgBox strTime & " ~ " & ServerTime
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub ReCalSum(Optional ByRef p_oValue As Double = 0#)
   Dim arrSum(3 To 10) As Double
   With grdDataList
      .Visible = False
      '收款
      If iCol = 4 Then
         .TextMatrix(iRow, 6) = Format(Val(.TextMatrix(iRow, 6)) - p_oValue + Val(.TextMatrix(iRow, 4)), "#.00")
         
      'Removed by Morgan 2020/9/23
      '   '簽約
      '   .TextMatrix(iRow, 10) = Format(Val(.TextMatrix(iRow, 10)) + p_oValue - Val(.TextMatrix(iRow, 4)), "#.00")
      ''簽約
      'Else
      '   .TextMatrix(iRow, 10) = Format(Val(.TextMatrix(iRow, 10)) - p_oValue + Val(.TextMatrix(iRow, 8)), "#.00")
      'end 2020/9/23
      End If
         
      For ii = 3 To 10
         arrSum(ii) = 0
         For jj = 2 To .Rows - 3
            arrSum(ii) = arrSum(ii) + Val(.TextMatrix(jj, ii))
         Next
      Next
         
      For ii = 3 To 10
         .TextMatrix(.Rows - 1, ii) = Format("" & arrSum(ii), "#.00")
      Next ii
      .Visible = True
   End With
End Sub

'Modify By Sindy 2020/8/31 + ByVal strDateS As String, ByVal strDateE As String
Private Function SetRow(ByRef p_adoRst As ADODB.Recordset, ByVal strDateS As String, _
   ByVal strDateE As String) As Boolean
   
Dim arrSum(1 To 10) As Double
Dim strID As String
   
On Error GoTo ErrHnd
   
   'Add By Sindy 2020/8/31
   Load frmpic002
   frmpic002.Label1.Caption = "資料計算中...請稍候..."
   frmpic002.Show
   frmpic002.ZOrder 0: DoEvents
'   If PUB_IsFormExist("frm210137") = True Then
'      Unload frm210137
'   End If
'   If PUB_IsFormExist("frm210141") = True Then
'      Unload frm210141
'   End If
   '2020/8/31 END
   
   With grdDataList
      .Visible = False
      .Rows = 3
      Call ClearRow(.Rows - 1)
      p_adoRst.MoveFirst
      While Not p_adoRst.EOF
         .Rows = .Rows + 1: .row = .Rows - 2
         .col = 0: .Text = "" & p_adoRst.Fields(.col)
         .CellAlignment = flexAlignCenterCenter
         For ii = 1 To 6 '10
            .col = ii
            If "" & p_adoRst.Fields(ii) <> "" Then
               .Text = Format("" & p_adoRst.Fields(ii), "#.00"): arrSum(.col) = arrSum(.col) + Val(.Text)
            End If
            If ii Mod 4 = 2 Then
               .CellBackColor = &H7FFFD4
            End If
         Next ii
         For ii = 11 To 12
            .col = ii: .Text = "" & p_adoRst.Fields(.col)
         Next ii
         p_adoRst.MoveNext
      Wend
      .row = .Rows - 1: .RowHeight(.Rows - 1) = 220 'Add by Amy 2013/08/01 +列高
      For ii = 1 To 10
         .col = ii
         If ii Mod 4 = 2 Then
            .CellBackColor = &H7FFFD4
         End If
      Next ii
      
      'Add By Sindy 2020/8/31
      For jj = 2 To .Rows - 1 '=grdDataList
         .row = jj
         .col = 11: strID = .Text '智權人員ID
         If strID <> "" Then
            For ii = 7 To 8
               .col = ii
               '已收文點數
               If ii = 7 Then
                  .Text = Format(PUB_CountCP18(0, strDateS, strDateE, , , strID), "#.00")
                  arrSum(.col) = arrSum(.col) + Val(.Text)
'                  frm210137.Hide
'                  frm210137.txtSalesArea = "" '業務區(起)
'                  frm210137.txtSalesArea1 = "" '業務區(迄)
'                  frm210137.txtSales = strID '智權人員ID
'                  frm210137.txtCloseDate(0) = strDateS '點數結算日(起)
'                  frm210137.txtCloseDate(1) = strDateE '點數結算日(迄)
'                  frm210137.cmdSearch_Click
'                  If frm210137.grdDataList.Rows > 1 Then
'                     If Trim(frm210137.grdDataList.TextMatrix(1, 5)) <> "" Then
'                        .Text = Format(frm210137.grdDataList.TextMatrix(1, 5), "#.00")
'                        arrSum(.col) = arrSum(.col) + Val(.Text)
'                     End If
'                  End If
               '未收款點數
               Else
                   'Modified by Lydia 2021/07/30 (總計)未收款點數，請不要傳入日期止日；因為不管收據日期只要是未收款都列入計算，但不改模組，怕將來有其他需求
                  '.Text = Format(PUB_CountCP18(1, "", strDateE, , , strID), "#.00")
                  .Text = Format(PUB_CountCP18(1, "", "", , , strID), "#.00")
                  arrSum(.col) = arrSum(.col) + Val(.Text)
'                  frm210141.Hide
'                  frm210141.txtSales = strID '智權人員ID
'                  frm210141.txtDate(0) = "" '點數結算日(起)
'                  frm210141.txtDate(1) = strDateE '點數結算日(迄)
'                  frm210141.cmdok_Click (1)
'                  If Trim(frm210141.txtTot(0)) <> "" Then
'                     .Text = Format(frm210141.txtTot(0), "#.00")
'                     arrSum(.col) = arrSum(.col) + Val(.Text)
'                  End If
               End If
               If ii Mod 4 = 2 Then
                  .CellBackColor = &H7FFFD4
               End If
            Next ii
         End If
      Next jj
'      If PUB_IsFormExist("frm210137") = True Then
'         Unload frm210137
'      End If
'      If PUB_IsFormExist("frm210141") = True Then
'         Unload frm210141
'      End If
      '2020/8/31 END
      
      .Rows = .Rows + 1
      .row = .Rows - 1
      .col = 0: .Text = "總計"
      .CellAlignment = flexAlignRightCenter: .CellFontBold = True
      For ii = 1 To 8 '10
         .col = ii
         .Text = Format("" & arrSum(.col), "#.00")
         If ii Mod 4 = 2 Then
            .CellBackColor = &H7FFFD4
         End If
      Next ii
      .Visible = True
   End With
   
   Unload frmpic002 'Add By Sindy 2020/8/31
   
   SetRow = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   Call SetToolBar(0)
   '預設為查詢
   Call SetToolBar(4)
   iCurState = 4
   SetDataListWidth
   txtSalesArea = PUB_GetStaffST15(strUserNum, 1)
   stET03s = txtSalesArea
   '2005/7/5 羅經理72009自7/1起只負責中二區
   'Add by Morgan 2005/4/25 小真輸國外部資料
   'modify by sonia 2014/6/9 +美珍77027
   If strUserNum = "65001" Or strUserNum = "77027" Then
      stET03s = "F41"
      txtSalesArea = "F41"
   '2010/4/1 ADD BY SONIA 開放林協理可處理S22及S29
   ElseIf strUserNum = "71003" Then
      stET03s = stET03s & ",S22,S29"
   '2010/9/8 ADD BY Sindy 開放陳淑芳可處理所有業務區 S2 字頭 的權限
   ElseIf strUserNum = "87027" Then
      stET03s = stET03s & ",S20,S21,S22,S23,S24,S29"
   '2015/5/27 ADD BY SONIA 再開放中所M71部門人員可處理中所各區資枓
   ElseIf PUB_GetST03(strUserNum) = "M71" And PUB_GetST06(strUserNum) = "2" Then
      stET03s = stET03s & ",S20,S21,S22,S23,S24,S29"
   '2015/5/27 END
   'ADD BY SONIA 2016/6/2 開放蘇嫄媛79053可處理S31
   ElseIf strUserNum = "79053" Then
      txtSalesArea = "S31"
      stET03s = "S31"
   End If
   'Add by Amy 2021/02/04 取得最大,最小日期
   MinDF02 = "Y": MaxDF02 = "Y"
   Call GetCloseDate(6, MinDF02)
   Call GetCloseDate(9, MaxDF02)
   'end 2021/02/04
   'Modify by Morgan 2005/5/3 不預設--秀玲
   'txtCloseDate = GetDDate(txtSalesArea)
   txtInput.Visible = False
   Me.grdDataList.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Added by Lydia 2021/07/27 回前一畫面
   If TypeName(m_PrevForm) <> "Nothing" Then
       m_PrevForm.Show
   End If
   'end 2021/07/27
   
   Set frm210103 = Nothing
End Sub

Private Sub GrdDataList_Click()
'Add by Amy 2013/08/02 +if 判斷
If iCurState = 2 Then
   With grdDataList
      .row = .MouseRow
      .col = .MouseCol
      If .CellBackColor = vbWhite Then
         SetBox
      End If
   End With
End If
End Sub

Private Sub SetBox()
   
   Dim lngLeft As Long, lngTop As Long
   
   With grdDataList
      'Modified by Morgan 2020/9/23 已取消簽約點數
      'If .row > 1 And (.col = 4 Or .col = 8) Then
      If .row > 1 And (.col = 4) Then
      'end 2020/9/23
         If .TextMatrix(.row, 11) <> "" Then
            txtInput.FontName = .CellFontName
            txtInput.FontSize = .CellFontSize
            txtInput.Alignment = .CellAlignment \ 5
            txtInput.Text = .TextMatrix(.row, .col)
            txtInput.Tag = txtInput.Text
            txtInput.Width = .ColWidth(.col)
            txtInput.Height = .RowHeight(.row)
            iRow = .row: iCol = .col
            txtInput.Visible = True
            txtInput.SetFocus
            TextInverse txtInput
            lngLeft = .Left + 25
            lngTop = .Top + .RowHeight(0) + .RowHeight(1) + 25
            For ii = 0 To .col - 1
               lngLeft = lngLeft + .ColWidth(ii)
            Next
            For ii = .TopRow To .row - 1
               lngTop = lngTop + .RowHeight(ii)
            Next
            txtInput.Left = lngLeft: txtInput.Top = lngTop
         End If
      End If
   End With
End Sub

Private Sub SetDataListWidth()
   
   With grdDataList
      .Visible = False
      .Clear
      .Rows = 3: .Cols = 13
      
      .row = 0: .RowHeight(0) = 250
      .col = 0: .ColWidth(.col) = 900
      
      'Modified by Lydia 2021/07/27 「已入帳」更名「已入帳點數」
      .col = .col + 1: .ColWidth(.col) = 750: .Text = "已入帳點數": .CellFontBold = True: .CellFontSize = 10 'Add by Amy 2013/08/01 +FontSize
      .col = .col + 1: .ColWidth(.col) = 900: .Text = "已入帳點數"
      
      'Modified by Lydia 2021/07/27 「已收款」更名「已收款點數」
      .col = .col + 1: .ColWidth(.col) = 900: .Text = "已收款點數": .CellFontBold = True: .CellFontSize = 10 'Add by Amy 2013/08/01 +FontSize
      .col = .col + 1: .ColWidth(.col) = 750: .Text = "已收款點數"
      .col = .col + 1: .ColWidth(.col) = 750: .Text = "已收款點數"
      .col = .col + 1: .ColWidth(.col) = 900: .Text = "已收款點數"
      
      
      .col = .col + 1: .ColWidth(.col) = 1400: .Text = "" 'Modify by Sindy 2020/8/31 "已簽約": .CellFontBold = True: .CellFontSize = 10 'Add by Amy 2013/08/01 +FontSize
      .col = .col + 1: .ColWidth(.col) = 1400: .Text = "" 'Modify by Sindy 2020/8/31 "已簽約"
      .col = .col + 1: .ColWidth(.col) = 0: .Text = "" 'Modify by Sindy 2020/8/31 "已簽約" 750
      .col = .col + 1: .ColWidth(.col) = 0: .Text = "" 'Modify by Sindy 2020/8/31 "已簽約" 900
      
      '控制欄位
      .col = .col + 1: .ColWidth(.col) = 0 '員工編號
      .col = .col + 1: .ColWidth(.col) = 0 '是否修改
      
      .row = 1
      .col = 0: .Text = "姓　名": .CellFontBold = True
      .col = .col + 1: .CellFontSize = 10: .Text = "本　日"
      .col = .col + 1: .CellFontSize = 10: .Text = "本月累計"
      
      .col = .col + 1: .CellFontSize = 10: .Text = "前日累計"
      .col = .col + 1: .CellFontSize = 10: .Text = "本　日"
      .col = .col + 1: .CellFontSize = 10: .Text = "扣點數"
      .col = .col + 1: .CellFontSize = 10: .Text = "本月累計"
      
      .col = .col + 1: .CellFontSize = 10: .Text = "已收文點數" 'Modify by Sindy 2020/8/31 "前日累計"
      .col = .col + 1: .CellFontSize = 10: .Text = "未收款點數" 'Modify by Sindy 2020/8/31 "本　日"
'      .col = .col + 1: .CellFontSize = 10: .Text = "銷　帳"
'      .col = .col + 1: .CellFontSize = 10: .Text = "未收累計"
      
      .MergeRow(0) = True
      .MergeCells = flexMergeRestrictRows
      .ColAlignmentFixed = flexAlignCenterCenter
      
      .Visible = True
   End With
End Sub

Private Sub ClearRow(ByRef p_iRow As Integer)
   For ii = 0 To grdDataList.Cols - 1
      grdDataList.TextMatrix(p_iRow, ii) = ""
   Next ii
End Sub

Private Sub grdDataList_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim iNextRow As Integer, iNextCol As Integer
   If KeyCode = 13 Or (Shift = 0 And KeyCode >= 37 And KeyCode <= 40) Then
      With grdDataList
         iNextRow = .row
         iNextCol = .col
         Select Case KeyCode
            Case 13
               SetBox
            Case 38 '上
               iNextRow = .row - 1
            Case 40 '下
               iNextRow = .row + 1
            Case 37 '左
               iNextCol = .col - 1
            Case 39 '右
               iNextCol = .col + 1
         End Select
         If iNextRow > 1 And iNextRow < .Rows And iNextCol > 0 And iNextCol < .Cols - 1 Then
'            .Row = iNextRow:
            .col = iNextCol
         End If
      End With
   End If
End Sub

Private Sub SetGridColor(p_Status As Integer)
   Dim lngColor As Long, iRow As Integer, bolVis As Boolean
   
   If p_Status = 2 Then
      lngColor = vbWhite
   Else
      lngColor = grdDataList.BackColor
   End If
   
   With grdDataList
      bolVis = .Visible
      .Visible = False
      For iRow = 2 To .Rows - 3
         'Add by Morgan 2010/3/15 已離職不可輸入
         'cancel by sonia 2020/4/27 南所A8021魏祥恩4/11離職,4/27有收款,杜經理說仍應輸魏祥恩
         'If Right(.TextMatrix(iRow, 0), 1) <> "*" Then
            .row = iRow
            .col = 4: .CellBackColor = lngColor
            '.col = 8: .CellBackColor = lngColor 'Removed by Morgan 2020/9/23 已取消簽約點數
         'End If
      Next
      .Visible = bolVis
   End With
End Sub

Private Sub SetInputs(Optional ByVal p_Status As Integer = 0)
   
   Select Case p_Status
      
      Case 2
      '修改
         txtSalesArea.Enabled = False
         txtCloseDate.Enabled = False
         'grdDataList.Enabled = True 'Modify by Amy 2013/08/02 搬至下面
         
      Case 4
      '查詢
         txtSalesArea.Enabled = True
         txtCloseDate.Enabled = True
         'grdDataList.Enabled = False

      Case Else
      '其他
         txtSalesArea.Enabled = False
         txtCloseDate.Enabled = False
         'grdDataList.Enabled = False
         
   End Select
   grdDataList.Enabled = True 'Add by Amy 2013/08/02
   SetGridColor p_Status
End Sub
Private Function SaveData() As Boolean
   With grdDataList
      For ii = 2 To .Rows - 3
         If .TextMatrix(ii, 12) <> "" Then
            If UpdateData(ii) = False Then
               MsgBox "更新[" & .TextMatrix(ii, 0) & "]點數資料失敗！", vbCritical
               Exit For
            End If
         Else
            If insertdata(ii) = False Then
               MsgBox "新增[" & .TextMatrix(ii, 0) & "]點數資料失敗！", vbCritical
               Exit For
            End If
         End If
      Next ii
      If ii > .Rows - 3 Then
         SaveData = True
      End If
   End With
   
End Function

Private Function insertdata(ByRef p_iRow As Integer) As Boolean

   Dim stCols As String, stValues As String
   
   With grdDataList
         stCols = "DF01,DF02"
         stValues = "'" & .TextMatrix(p_iRow, 11) & "'," & txtCloseDate
         If (.TextMatrix(p_iRow, 4) <> "") Then
            stCols = stCols & ",DF03"
            stValues = stValues & "," & .TextMatrix(p_iRow, 4)
         End If
         'Removed by Morgan 2020/9/23 已取消簽約點數
         'If (.TextMatrix(p_iRow, 8) <> "") Then
         '   stCols = stCols & ",DF04"
         '   stValues = stValues & "," & .TextMatrix(p_iRow, 8)
         'End If
         'end 2020/9/23
         stCols = stCols & ",DF06,DF07,DF08"
         stValues = stValues & ",'" & strUserNum & "'," & strSrvDate(1) & ",TO_NUMBER(TO_CHAR(SYSDATE,'HH24MI'))"
   End With
   
   strSql = "Insert Into DailyFeat (" & stCols & ") Values (" & stValues & ")"
   cnnConnection.Execute strSql
   insertdata = True
   
ErrHand:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Function UpdateData(ByRef p_iRow As Integer) As Boolean

   Dim stCols As String, stValues As String
   
   strSql = "Update DailyFeat Set DF09='" & strUserNum & "',DF10=" & strSrvDate(1) & ",DF11=TO_NUMBER(TO_CHAR(SYSDATE,'HH24MI'))"
   
   With grdDataList
         strSql = strSql & ",DF03=" & IIf(.TextMatrix(p_iRow, 4) = "", "NULL", .TextMatrix(p_iRow, 4))
         'strSql = strSql & ",DF04=" & IIf(.TextMatrix(p_iRow, 8) = "", "NULL", .TextMatrix(p_iRow, 8))'Removed by Morgan 2020/9/23 已取消簽約點數
         strSql = strSql & " Where DF01='" & .TextMatrix(p_iRow, 11) & "' AND DF02=" & txtCloseDate
   End With
   
   cnnConnection.Execute strSql
   UpdateData = True
   
ErrHand:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)

   Screen.MousePointer = vbHourglass
   
   Dim bolCancel As Boolean, stNextDate As String
   
   Select Case Button.Index
      Case 2 '修改
        'Add by Amy 2018/11/28 業務區為F字頭時不可修改
        If Left(txtSalesArea, 1) = "F" Then
            MsgBox "國外部資料由系統自動產生，不可異動！", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
         Call SetToolBar(2)
         iCurState = 2
         Call SetInputs(iCurState)
         grdDataList.row = 2
         grdDataList.col = 4
         SetBox
      Case 4 '查詢
         Call SetToolBar(4)
         iCurState = 4
         Call SetInputs(iCurState)
         txtSalesArea.SetFocus
         txtSalesArea_GotFocus
         SetDataListWidth
         
      Case 6, 7, 8, 9
         DoSearch Button.Index
      
      Case 11 '確定
         bolCancel = True
         '查詢
         If iCurState = 4 Then
            If txtSalesArea = "" Then
               MsgBox "請輸入業務區！", vbExclamation
               txtSalesArea.SetFocus
            ElseIf lblSalesArea = "" Then
               MsgBox "業務區輸入錯誤！", vbExclamation
               txtSalesArea_GotFocus
               txtSalesArea.SetFocus
            'Modify By Sindy 2010/9/9 開放杜副總可以輸入所有的業務區
            'ElseIf InStr(stET03s, txtSalesArea) = 0 And Pub_StrUserSt03 <> "M51" Then
            '2012/8/23 MODIFY BY SONIA 開放小真可以輸入S29
            'ElseIf InStr(stET03s, txtSalesArea) = 0 And Pub_StrUserSt03 <> "M51" And _
            '   strUserNum <> "68006" Then
            '   MsgBox "無該業務區權限！", vbExclamation
            '   txtSalesArea_GotFocus
            '   txtSalesArea.SetFocus
            'modify by sonia 2014/6/9 +美珍77027
            ElseIf (strUserNum = "65001" Or strUserNum = "77027") And txtSalesArea <> "S29" And txtSalesArea <> "F41" Then
               MsgBox "無該業務區權限！", vbExclamation
               txtSalesArea_GotFocus
               txtSalesArea.SetFocus
            '2012/8/23 END
            'modify by sonia 2014/6/9 +美珍77027
            'modify by sonia 2020/1/9 +69005
            'Modified by Lydia 2022/05/03 簡協理69005改為抓系統特殊設定「全所智權部主管」
            ElseIf InStr(stET03s, txtSalesArea) = 0 And Pub_StrUserSt03 <> "M51" And _
               strUserNum <> "68006" And strUserNum <> "65001" And strUserNum <> "77027" And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) = 0 Then
               MsgBox "無該業務區權限！", vbExclamation
               txtSalesArea_GotFocus
               txtSalesArea.SetFocus
            'Add by Morgan 2005/9/8 控制只可輸入智權部及國外部資料
            'Modified by Morgan 2020/9/22 +W10客服組
            ElseIf Not (txtSalesArea = "W10" Or txtSalesArea = "F41" Or Left(txtSalesArea, 1) = "S") Then
               MsgBox "業務區只可為智權部或國外部！", vbExclamation
               txtSalesArea_GotFocus
               txtSalesArea.SetFocus
            'ADD BY SONIA 2016/6/2 開放蘇嫄媛79053可處理S31
            ElseIf strUserNum = "79053" And txtSalesArea <> "S31" Then
               MsgBox "業務區只可為台南所S31！", vbExclamation
               txtSalesArea_GotFocus
               txtSalesArea.SetFocus
            ElseIf txtCloseDate = "" Then
               MsgBox "請輸入點數結算日！", vbExclamation
               txtCloseDate.SetFocus
            '日期格式
            ElseIf ChkDate(txtCloseDate) = False Then
               txtCloseDate.SetFocus
               txtCloseDate_GotFocus
            'Modify by Morgan 2010/8/19 百年蟲
            'ElseIf txtCloseDate > strSrvDate(2) Then
            ElseIf Val(txtCloseDate) > Val(strSrvDate(2)) Then
               MsgBox "點數結算日不可大於系統日！", vbExclamation
               txtCloseDate.SetFocus
               txtCloseDate_GotFocus
            '2012/1/5 add by sonia 因為1010104輸成1000104
            ElseIf Val(Mid(DBDATE(txtCloseDate), 1, 6)) <= Val(Mid(ChangeWDateStringToWString(DateAdd("M", -2, ChangeWStringToWDateString(strSrvDate(1)))), 1, 6)) Then
               MsgBox "點數結算日不可小於系統日前２個月！", vbExclamation
               txtCloseDate.SetFocus
               txtCloseDate_GotFocus
            '2012/1/5 end
            '工作天
            ElseIf ChkWorkDay(TransDate(txtCloseDate, 2)) = False Then
               MsgBox "請輸入工作天！", vbExclamation
               txtCloseDate.SetFocus
               txtCloseDate_GotFocus
            Else
               bolCancel = False
            End If
         '修改
         ElseIf iCurState = 2 Then
            If SaveData() = True Then
               txtInput.Visible = False
               bolCancel = False
               MsgBox "存檔成功！", vbInformation
            End If
         End If
         If bolCancel = False Then
            DoSearch
         End If
         
      Case 12
      '取消
         bolCancel = True
         If iCurState = 4 Then
           If txtSalesArea.Tag = "" Or txtCloseDate.Tag = "" Then
               MsgBox "無前次查詢紀錄，不可取消！", vbCritical
           Else
               bolCancel = False
           End If
         ElseIf iCurState = 2 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
               txtInput.Visible = False
               bolCancel = False
            End If
         End If
         If bolCancel = False Then
            txtSalesArea = txtSalesArea.Tag
            txtCloseDate = txtCloseDate.Tag
            DoSearch
         End If
      Case 14
      '結束
         If iCurState = 2 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
               Unload Me
            End If
         Else
            Unload Me
         End If
   End Select
   
   Screen.MousePointer = vbDefault
   
End Sub

Private Function GetCloseDate(p_Action As Integer, Optional ByRef stDF02 As String) As Boolean
   Dim stMsg As String
   Dim bolSet As Boolean 'Add by Amy 2021/02/04
   
On Error GoTo ErrHnd

   stMsg = "無法讀取工作日！"
   If stDF02 <> MsgText(601) Then bolSet = True: stDF02 = "" 'Add by Amy 2021/02/04
   Select Case p_Action
      Case 6
         strSql = "SELECT MIN(DF02) FROM STAFF,DAILYFEAT WHERE st04='1' and st15='" & txtSalesArea & "' AND DF01=ST01"
      Case 7
         strSql = "SELECT MAX(DF02) FROM STAFF,DAILYFEAT WHERE st04='1' and st15='" & txtSalesArea & "' AND DF01=ST01 AND DF02<" & txtCloseDate
         stMsg = "已經是第一筆！"
      Case 8
         strSql = "SELECT MIN(DF02) FROM STAFF,DAILYFEAT WHERE st04='1' and st15='" & txtSalesArea & "' AND DF01=ST01 AND DF02>" & txtCloseDate & " And DF02 is not null "
         stMsg = "已經是最後一筆！"
      Case 9
         strSql = "SELECT MAX(DF02) FROM STAFF,DAILYFEAT WHERE st04='1' and st15='" & txtSalesArea & "' AND DF01=ST01"
   End Select
   
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If Not .EOF Then
         'Modify by Amy 2021/02/04 若大於系統日的前一天無資料,再按下一筆(txtCloseDate被清空)會錯
         If bolSet = False Then
            txtCloseDate = "" & .Fields(0)
            If txtCloseDate = MsgText(601) And (p_Action = 7 Or p_Action = 8) Then
                If p_Action = 7 Then
                    txtCloseDate = MinDF02
                Else
                    txtCloseDate = MaxDF02
                End If
                MsgBox stMsg
            End If
         Else
            stDF02 = "" & .Fields(0)
         End If
         'end 2021/02/04
         GetCloseDate = True
      Else
         MsgBox stMsg
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub DoSearch(Optional p_Action As Integer = 0)
   If p_Action <> 0 Then
      If GetCloseDate(p_Action) = False Then Exit Sub
   End If
   If doQuery = True Then
      txtSalesArea.Tag = txtSalesArea
      txtCloseDate.Tag = txtCloseDate
      Call SetToolBar(0)
      iCurState = 0
      Call SetInputs(iCurState)
   End If
End Sub

Private Sub txtCloseDate_GotFocus()
   TextInverse txtCloseDate
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtCloseDate.IMEMode = 2
   CloseIme
End Sub

Private Sub GoNext()
   With grdDataList
      'Modified by Morgan 2020/9/23 已取消簽約點數
      'If .col = 4 Then
      '   .col = 8
      'Else
         .col = 4
         If .row < .Rows - 3 Then
            .row = .row + 1
         Else
            .row = 2
         End If
      'End If
      'end 2020/9/23
      If .CellBackColor = vbWhite Then
         SetBox
      Else
         GoNext
      End If
   End With
End Sub

Private Sub txtCloseDate_Validate(Cancel As Boolean)
   If txtCloseDate <> "" Then
      If ChkDate(txtCloseDate) = False Then
         Cancel = True
         txtCloseDate.SetFocus
         txtCloseDate_GotFocus
      End If
   End If
   'Added by Lydia 2024/01/31
   If strUserNum = "75007" Then
      If DBDATE(txtCloseDate) >= "20240201" Then
         MsgBox "不可查詢113年2月後的資料！", vbExclamation
         Cancel = True
         txtCloseDate.SetFocus
         txtCloseDate_GotFocus
      End If
   End If
   'end 2024/01/31
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = Asc(".") Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
      Beep
   Else
      If KeyAscii = vbKeyReturn Then
         grdDataList.TextMatrix(iRow, iCol) = Format(txtInput.Text, "#.00")
         Call ReCalSum(Val(txtInput.Tag))
         txtInput.Tag = txtInput
         GoNext
      ElseIf KeyAscii = vbKeyEscape Then
         txtInput = txtInput.Tag
         TextInverse txtInput
      End If
   End If
End Sub

Private Function GetDDate(ByRef p_ST03, Optional ByRef p_ST01 = "") As String
   stSQL = "SELECT MAX(DF02) FROM STAFF,DAILYFEAT WHERE ST03='" & p_ST03 & "' AND DF01=ST01"
   If p_ST01 <> "" Then
      stSQL = stSQL & " AND ST01='" & p_ST01 & "'"
   End If
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If Not (.EOF And .BOF) Then
         GetDDate = CompDate("2", 1, "" & .Fields(0))
         GetDDate = ChangeWStringToTString(PUB_GetWorkDay1(GetDDate, False))
      End If
   End With
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF3
      '修改
         If tlbar.Buttons(2).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(2))
         End If
      Case vbKeyF4
      '查詢
         If tlbar.Buttons(4).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(4))
         End If
      Case vbKeyF9
      '確定
         If tlbar.Buttons(11).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(11))
         End If
      Case vbKeyReturn
         If iCurState = 4 Then
            If tlbar.Buttons(11).Enabled = True Then
               Call tlbar_ButtonClick(tlbar.Buttons(11))
            End If
         End If
      Case vbKeyF10
      '取消
         If tlbar.Buttons(12).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(12))
         End If
      Case vbKeyEscape
      '結束
         If tlbar.Buttons(14).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(14))
         End If
    End Select
End Sub

'工具列控制
Private Sub SetToolBar(iStatus As Integer)

   Dim i As Integer
   For i = 1 To 13
      tlbar.Buttons(i).Enabled = False
   Next
   tlbar.Buttons(14).Enabled = True
   
   Select Case iStatus
   
      Case 0
      '瀏覽
         tlbar.Buttons(2).Enabled = True
         tlbar.Buttons(4).Enabled = True
         tlbar.Buttons(14).Enabled = True
         tlbar.Buttons(6).Enabled = True
         tlbar.Buttons(7).Enabled = True
         tlbar.Buttons(8).Enabled = True
         tlbar.Buttons(9).Enabled = True
      Case 1
      '新增
      Case 2
      '修改
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         tlbar.Buttons(14).Enabled = False
      Case 3
      '刪除
      Case 4
      '查詢
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         tlbar.Buttons(14).Enabled = True
      Case Else
      
   End Select
   
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
   'Added by Morgan 2020/9/23
   grdDataList.TextMatrix(iRow, iCol) = Format(txtInput.Text, "#.00")
   Call ReCalSum(Val(txtInput.Tag))
   txtInput.Tag = txtInput
   'end 2020/9/23
End Sub

Private Sub txtSalesArea_Change()
   If txtSalesArea = "" Then
      lblSalesArea = ""
   Else
      lblSalesArea = A0902Query(txtSalesArea)
   End If
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtSalesArea.IMEMode = 2
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
