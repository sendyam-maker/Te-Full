VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文"
   ClientHeight    =   5748
   ClientLeft      =   168
   ClientTop       =   960
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9348
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3048
      TabIndex        =   4
      Top             =   696
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060104_1.frx":0000
      Left            =   960
      List            =   "frm060104_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   1110
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8388
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "延期(&D)"
      Height          =   400
      Index           =   0
      Left            =   6732
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7560
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "FCP"
      Top             =   756
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   1
      Top             =   756
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   2
      Top             =   756
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   3
      Top             =   756
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4125
      Left            =   30
      TabIndex        =   12
      Top             =   1530
      Width           =   9255
      _ExtentX        =   16341
      _ExtentY        =   7260
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   9180
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   9180
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   780
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   1110
      Width           =   768
   End
   Begin MSForms.Label Label8 
      Height          =   285
      Left            =   1620
      TabIndex        =   9
      Top             =   1110
      Width           =   7560
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "13335;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm060104_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/12 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim pa(0 To 10) As String
Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer
Dim m_pa17 As String 'Add by Morgan 2008/5/5
Dim m_bolFMP As Boolean '是否外對外案件 Added by Morgan 2012/5/16
Dim mPA57 As String 'Add by Lydia 2014/12/24
Public bolIsEMPFlow As Boolean 'Add By Sindy 2023/11/9 是否為電子承辦簽核
Public m_EEP01 As String 'Add By Sindy 2023/11/27
Dim bolFirst As Boolean 'Add By Sindy 2024/1/8


'Add by Morgan 2007/10/18 重新委任之補文件檢查
Private Function CheckNeedAsign() As Boolean
   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 10) = "" And (pa(10) = 讓與 Or pa(10) = 合併 Or pa(10) = 繼承 Or pa(10) = 變更) Then
      strSql = "select np01 from nextprogress,caseprogress" & _
         " where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'" & _
         " and np06 is null and np07='202' and cp09(+)=np01 and cp10='928'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         CheckNeedAsign = True
         MsgBox "請至分案畫面輸入承辦人！"
      End If
   End If
End Function

Private Sub cmdOK_Click(Index As Integer)
   Dim i As Integer, bolChk As Boolean
   ' 90.10.18 modify by louis (記錄總收文號)
   Dim strCP09 As String
   'Add by Morgan 2004/8/18
   Dim bolNoFeeAlert As Boolean '是否提示規費檢查

   Select Case Index
      Case 0 '延期
         With MSHFlexGrid1
            For i = 1 To .Rows - 1
               If .TextMatrix(i, 0) = "v" Then
                  bolChk = True
                  Me.Tag = .TextMatrix(i, 2)
                  ' 90.10.18 modify by louis (記錄總收文號)
                  strCP09 = .TextMatrix(i, 2)
                  strExc(3) = .TextMatrix(i, 7) 'Added by Lydia 2018/02/01 案件性質
                  strExc(4) = .TextMatrix(i, 8)
                  strExc(5) = .TextMatrix(i, 9)
                  Exit For
               End If
            Next
         End With
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
            Exit Sub
         End If
         Me.Tag = Me.Tag & "0"
         'Add By Cheng 2002/07/15
         '所點選的案件性質不可為"延期"
         If PUB_CPKindDelay(strCP09, "P") Then
            Exit Sub
         End If
         'Add By Cheng 2002/07/12
         '已閉卷案件不可發文
         If PUB_CaseClosedCP09(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.row, 2)) Then
            Exit Sub
         End If
         
         'Added by Lydia 2018/02/01 排除D類客戶提供文件1920
         If Left(strCP09, 1) = "D" And (pa(10) = "1920" Or strExc(3) = "1920") Then
               MsgBox "客戶提供文件不可延期，請改到客戶提供文件處理 ! ", vbCritical
               Exit Sub
         End If
         'end 2018/0/2/01
         
         'Add By Sindy 2023/11/13
         '檢查是否有承辦歷程是否有產生承辦單可以發文
         If PUB_IsEmpFlowIsSend(strCP09) = False Then
            Exit Sub
         End If
         '2023/11/13 END
         
         'Add by Morgan 2007/7/19
         If Len(pa(10)) = 4 Then
            Exit Sub
         End If
         'end 2007/7/19
          
         'Add By Cheng 2002/06/20
         '延期記錄資料來源為案件進度檔
         frm060104_2.m_str_DL05 = "1"
         frm060104_2.Show
         frm060104_2.StrSales1 = strExc(4)
         frm060104_2.StrSales2 = strExc(5)
         Command1.SetFocus
         Me.Hide
      Case 1 '確定
         With MSHFlexGrid1
            For i = 1 To .Rows - 1
               If .TextMatrix(i, 0) = "v" Then
                  bolChk = True
                  Me.Tag = .TextMatrix(i, 2)
                  ' 90.10.18 modify by louis (記錄總收文號)
                  strCP09 = .TextMatrix(i, 2)
                  pa(10) = .TextMatrix(i, 7)
                  strExc(4) = .TextMatrix(i, 8)
                  strExc(5) = .TextMatrix(i, 9)
                  Exit For
               End If
            Next
         End With
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
            Exit Sub
         End If
         
         'Added by Lydia 2018/02/01 排除D類客戶提供文件1920
         If Left(strCP09, 1) = "D" And pa(10) = "1920" Then
               MsgBox "客戶提供文件不可直接發文，請改到客戶提供文件處理 ! ", vbCritical
               Exit Sub
         End If
         'end 2018/0/2/01
         
         'Add By Sindy 2023/11/13
         '檢查是否有承辦歷程是否有產生承辦單可以發文
         If PUB_IsEmpFlowIsSend(strCP09) = False Then
            Exit Sub
         End If
         '2023/11/13 END
         
         'Added by Morgan 2012/5/16
         If m_bolFMP Then
            frm060104_g.Show
            bolNoFeeAlert = True
         Else
         'end 2012/5/16
         
             'Added by Lydia 2018/04/20 控制新申請案發文要在中說之前(為了計算提申後告代/主動修正)
             If InStr("201,209,235,210", pa(10)) > 0 Then
                    strExc(0) = "select cp10 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'  and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' " & _
                                     " and cp10 in (101,102,103,125) and cp158=0 and cp159=0 "
                    intI = 1 'Add By Sindy 2024/1/11
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                    If intI = 1 Then
                       MsgBox "尚有新申請案未發文！", vbCritical
                       Exit Sub
                    End If
             End If
             'end 2018/04/20
             
            'Modify by Morgan 2009/9/8 實審或再審的退費不必檢查是否閉卷
            strExc(1) = ""
            If pa(10) = "908" Then
               'Modified by Morgan 2013/6/6 +檢查再審延期
               'strExc(0) = "select 1 from caseprogress a,caseprogress b where a.cp09='" & strCP09 & "' and b.cp09(+)=a.cp43 and b.cp10 in ('416','107')"
               'modify by sonia 2021/10/4  再加續行母案再審435(FCP-064306)
               strExc(0) = "select 1 from caseprogress a,caseprogress b where a.cp09='" & strCP09 & "' and b.cp09(+)=a.cp43 and b.cp10 in ('416','107','435')" & _
                  " union select 2 from  caseprogress a,caseprogress b,nextprogress where a.cp09='" & strCP09 & "' and b.cp09(+)=a.cp43 and b.cp10='404' and np01(+)=b.cp43 and np07 in ('107','435')" & _
                  " union select 3 from  caseprogress a,caseprogress b,caseprogress c where a.cp09='" & strCP09 & "' and b.cp09(+)=a.cp43 and b.cp10='404' and c.cp09(+)=b.cp43 and c.cp10 in ('107','435')"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  'Add by Morgan 2009/10/1 若退費收文後有機關來函則不可發文
                  'modify by sonia 2014/12/19 剔除通知實審日1204的來函 FCP-039303
                  strExc(0) = "select 1 from caseprogress a,caseprogress b where a.cp09='" & strCP09 & "' and b.cp01(+)=a.cp01 and b.cp02(+)=a.cp02 and b.cp03(+)=a.cp03 and b.cp04(+)=a.cp04 and b.cp05>=a.cp05 and b.cp09>'C' and b.cp10<>'1204' "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     MsgBox "已有收文機關來函不可發文!!"
                     Exit Sub
                  End If
                  'end 2009/10/1
                  strExc(1) = "N"
               'Add by Lydia 2014/12/24 代辦退費發文時檢查該案是否已閉卷,若無則提醒"本案尚未閉卷不能發文",且不能發文
                    If Not (mPA57 = "Y") Then
                         MsgBox "本案尚未閉卷不能發文!!"
                         Exit Sub
                    End If
               End If
            End If
            If strExc(1) = "" Then
               'Add By Cheng 2002/07/12
               '已閉卷案件不可發文
               'Modified by Morgan 2020/2/7 413 自請撤回 除外--何淑華
               If pa(10) <> "413" Then
                  If PUB_CaseClosedCP09(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.row, 2)) Then
                     Exit Sub
                  End If
               End If
            End If
            'Add By Cheng 2003/09/01
            '若為領證或年費
            If pa(10) = "601" Or pa(10) = "605" Then
               If PUB_ChkNP605(Me.Text1.Text & Me.Text2.Text & Me.Text3.Text & Me.Text4.Text) Then
                   MsgBox "下一程序有<年費>期限不可發文!!!", vbExclamation + vbOKOnly
                  Exit Sub
               End If
            End If
            'Add by Morgan 2007/10/18 無承辦人需分案檢查
            If CheckNeedAsign = True Then
               Exit Sub
            End If
            
            'Add by Morgan 2010/3/24
            'Modified by Morgan 2024/11/18 +477再審查加速審查並改用專用模組判斷
            'If pa(10) = "422" Then
            '   strExc(0) = "select * from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'" & _
            '      " and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='1204'" & _
            '      " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03" & _
            '      " and b.cp04=a.cp04 and b.cp10='107' and b.cp05>a.cp05 and b.cp57 is null)"
            '   intI = 1
            '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            '   If intI = 0 Then
            '      MsgBox "本案尚未有通知實審日程序，不可發文！"
            If pa(10) = "422" Or pa(10) = "447" Then
               If PUB_Chk1204(pa) = False Then
                  MsgBox "本案尚未接獲通知實審函，不可發文！", vbCritical
            'end 2024/11/14
                  Exit Sub
               End If
            End If
            
            'Add by Morgan 2012/12/4
            If pa(10) = "1001" Then
               'Modified by Morgan 2019/10/9 新法不限定發明初審
               'strExc(0) = "select * from caseprogress a,engineerprogress where cp09='" & strCP09 & "'" & _
                  " and ep02(+)=cp09 and ep09 is null" & _
                  " and exists(select * from caseprogress b where b.cp09=a.cp43 and b.cp10='101')"
               'intI = 1
               'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               'If intI = 1 Then
               '   MsgBox "發明初審核准尚未輸入完稿日，不可發文！"
               '   Exit Sub
               'End If
               strExc(0) = "select dst05,ep09 from caseprogress,patent,divsugtext,engineerprogress" & _
                  " where cp09='" & strCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
                  " and pa162='Y' and ep02(+)=cp09 and dst09(+)=cp09"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If IsNull(RsTemp("dst05")) Then
                     MsgBox "工程師尚未於核准函輸入分割建議，不可發文！", vbExclamation
                     Exit Sub
                  ElseIf IsNull(RsTemp("ep09")) Then
                     MsgBox "工程師主管尚未確認完稿，不可發文！", vbExclamation
                     Exit Sub
                  End If
               End If
               'end 2019/10/7
            End If
            'end 2012/12/4
            
            '20140318START ADD By eric
            If pa(10) = "414" Then
               strExc(0) = "select * from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                        " and (cp10='601' or cp10='605') and cp57 is null and cp27 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "先發文領證或年費！"
                  Exit Sub
               End If
            End If
            '20140318END
            
            
            'Add by Sindy 2011/3/11 若為新法, 416實審及203主動修正尚未發文, 先發文實審時請提示訊息
            'Remove by Lydia 2021/05/12 因流程已有改變，故請刪除彈提醒
'            If Chk99NewCase(pa(1), pa(2), pa(3), pa(4)) = True And pa(10) = "416" Then
'               strExc(0) = "select * from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'" & _
'                  " and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in('203')" & _
'                  " and cp27 is null"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  '2011/4/27 modify by sonia 改詢問方式
'                  'MsgBox "請先發文主動修正！"
'                  'Exit Sub
'                  If MsgBox("此案有主動修正未發文, 若為同日發文請先發文主動修正 ! 是否繼續發文 ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
'                     Exit Sub
'                  End If
'                  '2011/4/27 end
'               End If
'            End If
'
'            '2011/4/27 add by sonia 加201新案翻譯及203主動修正的控制,但實審未收文或未發文則不檢查
'            If Chk99NewCase(pa(1), pa(2), pa(3), pa(4)) = True And pa(10) = "201" Then
'               strExc(0) = "select a.cp27,b.cp27 from caseprogress a,caseprogress b where a.cp01='" & pa(1) & "' and a.cp02='" & pa(2) & "'" & _
'                  " and a.cp03='" & pa(3) & "' and a.cp04='" & pa(4) & "' and a.cp10='203' and a.cp27 is null" & _
'                  " and a.cp01=b.cp01(+) and a.cp02=b.cp02(+) and a.cp03=b.cp03(+) and a.cp04=b.cp04(+) and '416'=b.cp10(+)"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  If Not IsNull(RsTemp.Fields(1)) Then
'                     If MsgBox("此案有主動修正未發文, 若為同日發文請先發文主動修正 ! 是否繼續發文 ?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
'                        Exit Sub
'                     End If
'                  End If
'               End If
'            End If
'            '2011/4/27
            'end 2021/05/12
            
            'Added by Morgan 2012/7/24
            '發文申請依職權修正時若有修正未發文則提醒--靜芳101/6/25請作單
            If pa(10) = "227" Then
               If PUB_ChkCPExist(pa, "204", 1) Then
                  MsgBox "此案尚有修正未發文, 是否要取消修正收文!!", vbInformation
               End If
            End If
            'end 2012/7/24
            
            'Added by Morgan 2013/7/10
            '一案兩請檢查
            'Modified by Morgan 2013/11/6 +235核對中說格式
            If pa(10) = "101" Or pa(10) = "102" Then
               strExc(0) = "select cm05||'-'||cm06||decode(cm07||cm08,'000','','-'||cm07||'-'||cm08) p2 from casemap" & _
                  " where cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "' and cm10='3'" & _
                  " union select cm01||'-'||cm02||decode(cm03||cm04,'000','','-'||cm03||'-'||cm04) p2 from casemap" & _
                  " where cm05='" & pa(1) & "' and cm06='" & pa(2) & "' and cm07='" & pa(3) & "' and cm08='" & pa(4) & "' and cm10='3'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "本案與 " & RsTemp(0) & " 案為一案兩請，請檢查二案是否同時送件！", vbExclamation
                  If pa(10) = "102" Then
                     If PUB_ChkCPExist(pa, "209") = False And PUB_ChkCPExist(pa, "235") = False Then
                        MsgBox "本案為一案兩請之新型案，要收文檢視中說/核對中說格式!!", vbExclamation
                     End If
                  End If
               End If
            End If
            'end 2013/7/10
            
            Select Case pa(10)
             Case 延期
                Me.Tag = Me.Tag & "1"
                'Add By Cheng 2002/06/20
                '延期記錄資料來源為下一程序檔
                frm060104_2.m_str_DL05 = "2"
                
                frm060104_2.Show
                frm060104_2.StrSales1 = strExc(4)
                frm060104_2.StrSales2 = strExc(5)
             Case 授權, 終止授權
                frm060104_6.Show
             Case 領證及繳年費
                 'Modify by Morgan 2004/3/19
                 '加基本檔核准檢查
                 'frm060104_7.Show
                 If PUB_ApproveCheck(strCP09) Then
                     frm060104_7.Show
                 Else
                     Exit Sub
                 End If
                'Add by Morgan 2004/8/18
                bolNoFeeAlert = True
             Case 設定質權, 終止設定質權
                frm060104_8.Show
             Case 異議_專, 舉發
                frm060104_9.Show
             Case 年費
                 'Modify by Morgan 2004/3/19
                 '加基本檔核准檢查
                 'frm060104_a.Show
                 If PUB_ApproveCheck(strCP09) Then
                     frm060104_a.Show
                 Else
                     Exit Sub
                 End If
             'Add by Morgan 2004/8/18
             bolNoFeeAlert = True
             Case 讓與, 合併
                frm060104_c.Show
                'Add by Morgan 2006/6/8
                If frm060104_c.StopMe = True Then
                   Unload frm060104_c
                   Exit Sub
                End If
                'end 2006/6/8
                If pa(10) = 合併 Then
                   frm060104_c.Caption = "外專發文-合併"
                End If
                
             Case 繼承
                frm060104_c.Show
                frm060104_c.Caption = "外專發文-繼承"
             
             Case Else
                'Add by Morgan 2007/7/19 來函發文
                If Len(pa(10)) = 4 Then
                   frm060104_e.Show
                   'frm060104_e.txtCP113.SetFocus 'Remove by Morgan 2007/8/27 取消預設--靜芳
                   bolNoFeeAlert = True
                Else
                'End 2007/7/19
                
                   'Add by Morgan 2004/6/3 台灣改請案,延緩公告檢查
                   Select Case Val(pa(10))
                      Case 301 To 306
                         If PUB_Check301(pa(1), pa(2), pa(3), pa(4)) = False Then Exit Sub
                      Case 412
                         MsgBox "延緩公告不可單獨發文！", vbInformation: Exit Sub
                      Case 421 '技術報告
                         If PUB_Check421(pa(1), pa(2), pa(3), pa(4)) = False Then Exit Sub
                         
                      'Add by Morgan 2004/8/13
                      '更正案發文時,要檢查是否領証已發文
                      Case 更正
                         'Modify by Morgan 2008/5/5 '加判斷無專利權 --Susan FCP-36817
                         If m_pa17 = "" Then
                            If PUB_ChkCPExist(pa, 領證及繳年費, 2) = False Then
                               MsgBox "領証未發文，更正不可發文！", vbInformation: Exit Sub
                            End If
                         End If
                   End Select
                   'Add end
                   
                   'Add by Morgan 2005/12/6 有國外案時提醒
                   If pa(10) = "201" Then
                      strExc(0) = "select cm01||cm02||cm03||cm04,ST02,nvl(pa05,nvl(pa06,pa07)) FROM " & _
                          "CASEMAP,CASEPROGRESS,PATENT,STAFF WHERE CM05='" & pa(1) & "' AND CM06='" & pa(2) & "' AND " & _
                          "CM07='" & pa(3) & "' AND CM08='" & pa(4) & "' AND CM10='0' AND " & _
                          "cm01=pa01 and cm02=pa02 and cm03=pa03 and cm04=pa04 AND " & _
                          "cm01=cp01 and cm02=cp02 and cm03=cp03 and cm04=cp04 AND " & _
                          "CP27 IS NULL and CP57 IS NULL and cp10 in (" + CNULL(發明申請) + "," + CNULL(新型申請) + "," + CNULL(設計申請) + "," + CNULL(追加申請) + "," + CNULL(聯合申請) + "," + CNULL(翻譯) + ")" & _
                          " and cp14=st01(+) ORDER BY cm01,cm02,cm03,CM04"
                      intI = 1
                      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                      If intI = 1 Then
                         frm060104_1_1.Show vbModal
                      End If
                      '2007/8/3 add by sonia 新案翻譯未上完稿日者不可發文
                      strExc(0) = "select ep09 FROM ENGINEERPROGRESS WHERE EP02='" & strCP09 & "'"
                      intI = 1
                      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                      If intI = 1 Then
                         If IsNull(RsTemp.Fields("EP09")) Then
                            MsgBox "新案翻譯未上完稿日，不可發文！", vbInformation: Exit Sub
                         End If
                      Else
                         MsgBox "新案翻譯未上完稿日，不可發文！", vbInformation: Exit Sub
                      End If
                      '2007/8/3 end
                   End If
                   '2005/12/6 end
                   frm060104_3.Show
                End If
            End Select
            
         End If 'Added by Morgan 2012/5/16
         
         Command1.SetFocus
         Me.Hide
         
'cancel by sonia 2017/11/22與敏莉確認,取消此項提醒
'         If bolNoFeeAlert = False Then
'            MsgBox "發文前請檢查發文日及規費是否正確！", vbExclamation, "注意"
'         Else
'            MsgBox "發文前請檢查發文日是否正確！", vbExclamation, "注意"
'         End If
'end 2017/11/22
         
         ' 顯示專利基本檔
         'Modify By Cheng 2002/01/24
         '暫時先註解起來
'         ShowMaintainForm strCP09

      Case 2
         Unload Me
   End Select
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label8 = pa(5)
      Case "英"
         Label8 = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label8 = pa(7)
   End Select
End Sub

Public Sub Command1_Click()
   If CheckCP02 = False Then Exit Sub 'Add by Morgan 2004/10/21
   Dim i As Integer
   Dim stCon As String, stNation As String
   Dim intQRow As Integer 'Add By Sindy 2023/11/27
   
   Label8 = ""
   MSHFlexGrid1.Clear
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   'Modified by Morgan 2012/5/16 +P,CFP,PS,CPS
   'If Text1 = "FCP" Then
   'Add by Lydia 2014/12/24 +是否閉卷(PA57,SP15)
   If Text1 = "FCP" Or Text1 = "P" Or Text1 = "CFP" Then
      strExc(0) = "SELECT PA05,PA06,PA07,PA23,PA17,PA09,PA57 FROM PATENT WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   ElseIf Text1 = "FG" Or Text1 = "PS" Then
      strExc(0) = "SELECT SP05,SP06,SP07,'','',SP09,SP15 FROM SERVICEPRACTICE WHERE " & ChgService(pa(1) & pa(2) & pa(3) & pa(4))
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         For i = 0 To 2
            If IsNull(.Fields(i)) = False Then pa(i + 5) = .Fields(i)
         Next
         Label8 = pa(5)
         m_pa17 = "" & .Fields(4) 'Add by Morgan 2008/5/5
         stNation = "" & .Fields(5) 'Added by Morgan 2012/5/16
         mPA57 = "" & .Fields(6)  'Add by Lydia 2014/12/24
      End With
   End If
   
   'Added by Morgan 2012/5/16
   m_bolFMP = False
   cmdOK(0).Enabled = True
   If Text1 = "P" Or Text1 = "CFP" Or Text1 = "PS" Or Text1 = "CPS" Then
      m_bolFMP = True
      cmdOK(0).Enabled = False
      'Modified by Morgan 2012/7/11 +924
      '2012/8/3 modify by sonia +937
      'Modified by Morgan 2012/10/1 +903,904
      'Modified by Morgan 2013/10/4 +927
      'MODIFY BY SONIA 2014/6/23 +949寄中說
      'Modified by Morgan 2023/4/27 +969提供本所意見 Ex:P129271
      'Modify By Sindy 2023/11/30 改用常變數 FMPtoFCPSendCasePtyList
      'stCon = " and cp12 like 'F%' and (cp09>'C' or cp10='901' or cp10='902' or cp10='903' or cp10='904' or cp10='924' or cp10='927' or cp10='937' or cp10='949' or cp10='969') and cp14 is not null"
      stCon = " and cp12 like 'F%' and (cp09>'C' or cp10 in(" & FMPtoFCPSendCasePtyList & ")) and cp14 is not null"
   End If
   'end 2012/5/16
   
   'Modify by Morgan 2007/7/18 來函也要可以發文
   'strExc(0) = "select ''," & SQLDate("CP05") & ",cp09,cpm03,staff.st02 as st1," & _
      "staff1.st02 as st2,cp64,cp10,cp12,cp13 from caseprogress, casepropertymap," & _
      "staff,staff staff1 where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND CP27 IS NULL AND CP57 IS NULL AND ( CP09<'C' )" & _
      " AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)"
   strExc(0) = "select ''," & SQLDate("CP05") & ",cp09,decode(cp01,'P',decode('" & stNation & "','000',cpm03,cpm04),cpm03) cpm03,staff.st02 as st1," & _
      "staff1.st02 as st2,cp64,cp10,cp12,cp13,cp14 from caseprogress, casepropertymap," & _
      "staff,staff staff1 where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND CP27 IS NULL AND CP57 IS NULL" & stCon & _
      " AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)"
      
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   'Add By Sindy 2024/1/8 控管多筆未發文時,後面的資料不能直接進入發文作業
   If bolFirst = True Then
      bolIsEMPFlow = False
   End If
   bolFirst = True
   '2024/1/8 END
   
   'Add By Cheng 2002/05/10
   '若只搜尋到一筆時直接勾選
'   'Modify By Sindy 2023/11/27 + Or (bolIsEMPFlow = True And m_EEP01 <> "")
   If Me.MSHFlexGrid1.Rows = 2 Or (bolIsEMPFlow = True And m_EEP01 <> "") Then
      'Modify By Sindy 2023/11/9
      'MSHFlexGrid1_Click
      '若有資料游標停在第一筆
      'Add By Sindy 2023/11/27
      If (bolIsEMPFlow = True And m_EEP01 <> "") Then
         For i = 1 To MSHFlexGrid1.Rows - 1
            If MSHFlexGrid1.TextMatrix(i, 2) = m_EEP01 Then
               intQRow = i
               Exit For
            End If
         Next i
      Else
      '2023/11/27 END
        intQRow = 1
      End If
      If intQRow > 0 Then
         MSHFlexGrid1.Visible = False
         MSHFlexGrid1.col = 0
         MSHFlexGrid1.row = intQRow
         'If RsTemp.RecordCount = 1 Then
            MSHFlexGrid1.Text = "v"
            For i = 0 To MSHFlexGrid1.Cols - 1
               MSHFlexGrid1.col = i
               MSHFlexGrid1.CellBackColor = &HFFC0C0
            Next i
         'End If
         MSHFlexGrid1.Visible = True
         If bolIsEMPFlow = True Then Call cmdOK_Click(1)
      End If
      '2023/11/9 END
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   Combo1.ListIndex = 0
   InitGrid 11, MSHFlexGrid1
   GridHead
   'Add By Cheng 2002/12/11
   SendKeys "{Tab}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   bolIsEMPFlow = False 'Add By Sindy 2023/11/9
   Set frm060104_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK(1).SetFocus
End Sub

Private Sub Text1_Change()
   MSHFlexGrid1.Clear
End Sub

Private Sub Text1_GotFocus()
  TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   'Modified by Morgan 2012/5/16 +P,PS,CFP,CPS
   'If Text1 <> "FCP" And Text1 <> "FG" Then
   If Text1 <> "FCP" And Text1 <> "FG" And Text1 <> "P" And Text1 <> "PS" And Text1 <> "CFP" And Text1 <> "CPS" Then
      MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
      TextInverse Text1
      Cancel = True
   End If
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1200: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1400: .Text = "進度備註"
      For i = 7 To 10
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Public Sub ReQuery()
   Command1_Click
End Sub

Public Sub Clear()
    'Modify By Cheng 2002/12/11
    '保留原輸入的系統類別
'   Text1 = Empty
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   'Combo1.Clear
   Label8 = ""
   InitGrid 11, MSHFlexGrid1
   GridHead
    'Modify By Cheng 2002/12/19
    '控制游標停在本所案號第二欄
   'Add By Cheng 2002/12/11
   Me.Text1.SetFocus: DoEvents
   Me.Text2.SetFocus: DoEvents
End Sub

Private Sub Text2_Change()
   MSHFlexGrid1.Clear
End Sub

Private Sub Text2_GotFocus()
  TextInverse Text2
End Sub

Private Sub Text3_Change()
   MSHFlexGrid1.Clear
End Sub

Private Sub Text3_GotFocus()
  TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Change()
   MSHFlexGrid1.Clear
End Sub

Private Sub Text4_GotFocus()
  TextInverse Text4
End Sub
'Add by Morgan 2004/10/21 檢查本所號
Private Function CheckCP02() As Boolean
   If Len(Text2.Text) <> 6 Then
      MsgBox "本所案號輸入錯誤！"
      Text2.SetFocus
      Text2_GotFocus
      CheckCP02 = False
      Exit Function
   End If
   CheckCP02 = True
End Function

