VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060103_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書"
   ClientHeight    =   5736
   ClientLeft      =   144
   ClientTop       =   2424
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9336
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   17
      Top             =   5310
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   3048
      TabIndex        =   16
      Top             =   528
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "FCP"
      Top             =   576
      Width           =   550
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1512
      MaxLength       =   6
      TabIndex        =   1
      Top             =   576
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2352
      MaxLength       =   1
      TabIndex        =   2
      Top             =   576
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2592
      MaxLength       =   2
      TabIndex        =   3
      Top             =   576
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   270
      Left            =   6420
      MaxLength       =   7
      TabIndex        =   6
      Top             =   576
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7464
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8316
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3600
      Left            =   120
      TabIndex        =   15
      Top             =   1656
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   6350
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "(1.詢進度 2.延期 3.電子送件 4.新案申請書)"
      Height          =   252
      Left            =   1560
      TabIndex        =   19
      Top             =   5340
      Width           =   3405
   End
   Begin VB.Label Label9 
      Caption         =   "特殊申請書:"
      Height          =   252
      Left            =   120
      TabIndex        =   18
      Top             =   5340
      Width           =   972
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   1230
      Width           =   8235
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "14520;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   576
      Width           =   768
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Left            =   5400
      TabIndex        =   13
      Top             =   576
      Visible         =   0   'False
      Width           =   948
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   936
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   180
      Left            =   960
      TabIndex        =   11
      Top             =   936
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   5400
      TabIndex        =   10
      Top             =   930
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      Height          =   180
      Left            =   6420
      TabIndex        =   9
      Top             =   936
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   1296
      Width           =   768
   End
End
Attribute VB_Name = "frm060103_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/10/12 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim intWhere As Integer
Dim intLastRow As Integer
Dim bNewCaseCp118 As Boolean 'Added by Lydia 2018/09/12 本案的新案進度是否為電子送件
'Added by Lydia 2018/09/12 案件性質(非新申請案已有電子送件申請書)
Private Const cNewFrmList As String = "202,201,209,210,235,414,404,701,702,405,413,604,401"


Public Sub cmdok_Click(Index As Integer)
Dim i As Integer, bolChk As Boolean
Dim strCP09 As String 'Add By Sindy 2022/3/4
   
   Select Case Index
      Case 0 '確定
         For i = 1 To MSHFlexGrid1.Rows - 1
            If MSHFlexGrid1.TextMatrix(i, 0) = "v" Then
               bolChk = True
               Me.Tag = MSHFlexGrid1.TextMatrix(i, 2)
               pa(10) = MSHFlexGrid1.TextMatrix(i, 8)
               strCP09 = MSHFlexGrid1.TextMatrix(i, 2)
               Exit For
            End If
         Next
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
            Exit Sub
         End If
         
         'Added by Morgan 2025/10/20--Sharon
         If pa(143) = "N" And pa(10) = "605" And Text6 = "3" Then
            MsgBox "本案繳年費本所不出名，請用大批CSV 繳費！", vbCritical
            Exit Sub
         End If
         'end 2025/10/20
         
         'Added by Lydia 2018/09/12 判斷本案的新案進度是否為電子送件
         If bNewCaseCp118 = False And Text6 = "3" Then
              If MsgBox("此案非電子送件，是否確定產生電子送件申請書？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
              End If
         End If
         'end 2018/09/12
         
         'Modify By Sindy 2024/4/25 改共用函數
         Call PUB_FCPChkCP141(strCP09)
'         'Add By Sindy 2022/3/4 若有設定指定送件日，在產生申請書時可彈提醒
'         strExc(0) = "select * from caseprogress where cp09='" & strCP09 & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If "" & RsTemp.Fields("cp142") <> "" Then '有指定送件日期
'               '程序人員在產生申請書時(僅限程序產生的申請書)，若那道進度檔有設定指定送件日"當天"or "之後"，
'               '若產生申請書的系統日，小於指定送件"當天"or "之後"的日期，則彈提醒
'               If ("" & RsTemp.Fields("cp164") = "1" Or "" & RsTemp.Fields("cp164") = "") Or "" & RsTemp.Fields("cp164") = "3" Then
'                  If RsTemp.Fields("cp142") > strSrvDate(1) Then
'                     MsgBox "此道設定指定" & ChangeWStringToTDateString(RsTemp.Fields("cp142")) & "日" & IIf("" & RsTemp.Fields("cp164") = "3", "之後", "當天") & "送件，請注意。"
'                  End If
'               End If
'            End If
'         End If
'         '2022/3/4 END
         
         '若有輸入特殊申請書
         If Text6 <> "" Then
            'Add By Sindy 2025/1/6 敏莉:申請書的管控： 在產生 "自請撤回" or "代辦退費"電子申請書時，
            '   若下一程序尚有"委任書"續辦是"空"，則彈訊息：本案尚有委任書未補呈，故無法辦理撤回，請通知承辦。
            If pa(10) = 退費 Or pa(10) = 自請撤回 Then
               strExc(0) = "select np01 from nextprogress where np02='" & pa(1) & "' AND np03='" & pa(2) & "' AND np04='" & pa(3) & "' AND np05='" & pa(4) & "'" & _
                           " AND instr(np15,'委任書')>0 AND np06 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "辦理撤回需補呈委任書，請通知承辦。", vbExclamation
               End If
            End If
            '2025/1/6 END
            
            Select Case Text6
               Case "1" '詢進度
                  frm06010309_1.Show
                  
               Case "2" '延期
                  '92.7.5 MODIFY BY SONIA
                  'frm06010308_1.Show
                  If pa(10) = 異議_專 Or pa(10) = 舉發 Then
                     frm06010311_1.Show
                  Else
                     'Add By Sindy 2018/7/4
                     If MsgBox("要出電子送件申請書嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                        frm06010308_1.m_CP118isY = True '電子送件申請書
                        frm06010308_1.Caption = "各式申請書-電子送件-延期"
                     End If
                     '2018/7/4 END
                     frm06010308_1.Show
                  End If
                  '92.7.5 END
                  
               'Added by Morgan 2013/7/11
               Case "3" '電子送件
                  'Memo by Lydia 2018/09/12 若有非新申請案的案件性質加入，請加入常數cNewFrmList
                  Select Case pa(10)
                  'Modified by Morgan 2018/1/12 +新型申請,設計申請
                  'Modify By Sindy 2018/8/7 + 衍生設計
                  Case 發明申請, 新型申請, 設計申請, 衍生設計
                     frm06010301_1.SetParent Me 'Add By Sindy 2022/5/3
                     frm06010301_1.Show
                  'Modified by Morgan 2013/8/15 +201,209,210 --靜芳
                  '202.補文件, 201.翻譯, 209.檢視中說, 210.製作中說, 235.核對中說格式
                  'Add By Sindy 2017/11/8 + 實體審查
                  'Modify By Sindy 2019/8/8 + 911.補收款
                  Case 補文件, "201", "209", "210", "235", 實體審查, "435", 補收款
                     frm06010301_2.Show
                     If pa(10) = 實體審查 Then
                        frm06010301_2.Caption = "各式申請書-電子送件-實體審查"
                     End If
                  'Add By Sindy 2018/6/13
                  Case 讓與, 合併
                     frm06010302_1.Show
                     frm06010302_1.Caption = "各式申請書-電子送件-讓與, 合併"
                  'Add By Sindy 2018/8/10
                  Case 變更
                     'Modify By Sindy 2022/6/7
                     '授權變更
                     If GetCP10(MSHFlexGrid1.TextMatrix(i, 4)) = "704" Then '相關總收文號
                        frm06010310_1.SetParent Me 'Add By Sindy 2023/2/16
                        frm06010310_1.Show
                        frm06010310_1.Caption = "各式申請書-電子送件-授權變更"
                        frm06010310_1.txtCP84 = 2000
                     Else
                     '2022/6/7 END
                        Set frm06010303_1.oParent = Me 'Add by Morgan 2011/10/5
                        frm06010303_1.m_CP118isY = "Y" '電子送件申請書
                        frm06010303_1.Caption = "各式申請書-電子送件-變更"
                        frm06010303_1.LoadMe Me.Tag, Text1, Text2, Text3, Text4, 61
                     End If
                  'Add By Sindy 2019/1/2
                  Case 退費 '=908代辦退費
                     frm06010306_1.m_CP118isY = True '電子送件申請書
                     frm06010306_1.Show
                     frm06010306_1.Caption = "各式申請書-電子送件-退費"
                  'Add By Sindy 2018/7/4
                  Case 延期
                     frm06010308_1.m_CP118isY = True '電子送件申請書
                     frm06010308_1.Show
                     frm06010308_1.Caption = "各式申請書-電子送件-延期"
                  'Add By Sindy 2019/1/2
                  Case 催審
                     frm06010309_1.m_CP118isY = True '電子送件申請書
                     frm06010309_1.Show
                     frm06010309_1.Caption = "各式申請書-電子送件-催審"
                  'Add By Sindy 2019/7/23
                  Case 更改
                     frm06010307_1.m_CP118isY = True '電子送件申請書
                     frm06010307_1.Show
                     frm06010307_1.Caption = "各式申請書-電子送件-更改"
                  'Add By Sindy 2018/8/7
                  '電子送件-其他
                  'Modify By Sindy 2018/11/2 + 領證及繳年費
                  'Modify By Sindy 2018/11/28 + 年費
                  'Modify By Sindy 2019/11/6 + 425.優先審查,704.授權,705.終止授權
                  'Modify By Sindy 2020/8/24 432.回復原狀,206.補充說明
                  'Modify By Sindy 2022/5/31 + 245.延緩審查
                  'Modify By Sindy 2022/6/2 + 439.專利權部分拋棄,440.申請權部分拋棄
                  'Modify By Sindy 2022/12/2 + 124.回復優先權主張
                  'Modified by Morgan 2022/12/27 +443 申請證書副本
                  'Modify By Sindy 2024/9/26 + 422.加速審查(再審查)
                  'Modified by Morgan 2024/11/14 422改為447再審查加速審查
                  Case 申請優先權證明, 自請撤回, 補換發證書, 領證及繳年費, 年費, _
                       申請英文證明, 繼承, 提早公開, 其他, 425, 授權, 終止授權, _
                       432, 補充說明, "245", 439, 440, 124, 443, 447
                     'Modify By Sindy 2018/11/28 加基本檔核准檢查
                     If pa(10) = 領證及繳年費 Or pa(10) = 年費 Then
                        If PUB_ApproveCheck(Me.Tag, "不可產生申請書") = False Then
                           Exit Sub
                        End If
                     End If
                     '2018/11/28 END
                     'Modify By Sindy 2024/9/26
                     'Modified by Morgan 2024/11/14 改為447再審查加速審查並檢查是否已通知實審(原檢查取消且分割案關聯的可能是分割或續行母案再審)--敏莉
                     'If pa(10) = "422" Then
                     '   If GetCP10(MSHFlexGrid1.TextMatrix(i, 4)) <> "107" Then
                     '      MsgBox "此加速審查，相關總收文號非再審查，請確認！"
                     '      Exit Sub
                     '   End If
                     '   strExc(0) = "select cp10,cp27,cp57 from caseprogress" & _
                     '               " where cp09='" & MSHFlexGrid1.TextMatrix(i, 4) & "'"
                     '   intI = 1
                     '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     '   If intI = 1 Then
                     '      If Val("" & RsTemp.Fields("cp27")) = 0 Then
                     '         MsgBox "加速審查的再審查案未發文，請確認！"
                     '         Exit Sub
                     '      End If
                     '      If Val("" & RsTemp.Fields("cp57")) > 0 Then
                     '         MsgBox "加速審查的再審查已取消收文，請確認！"
                     '         Exit Sub
                     '      End If
                     '   End If
                     If pa(10) = "447" Then
                        If PUB_Chk1204(pa) = False Then
                           MsgBox "本案尚未接獲通知實審函，請勿送件！", vbCritical
                           Exit Sub
                        End If
                     'end 2024/11/14
                     End If
                     '2024/9/26 END
                     frm06010310_1.SetParent Me 'Add By Sindy 2023/2/16
                     frm06010310_1.Show
                     frm06010310_1.Caption = "各式申請書-電子送件-其他"
                  Case Else
                     MsgBox "點選的案件性質目前尚無電子送件申請書！"
                     Exit Sub
                  End Select
               
               'Add By Sindy 2015/11/26
               Case "4" '新案申請書
                  Select Case pa(10)
                  Case 發明申請, 新型申請, 設計申請
                     frm06010301_1.Hide 'Show
                     'cmdOK(1).SetFocus
                     'Me.Hide
                     'Modify By Sindy 2018/5/24
'                     Call frm06010301_1.cmdok_Click(0)
'                     Exit Sub
                     '2018/5/24 END
                     frm06010301_1.SSTab1.TabVisible(1) = False
                     frm06010301_1.SSTab1.TabVisible(2) = False
                     frm06010301_1.SSTab1.TabVisible(3) = False
                     frm06010301_1.SSTab1.TabVisible(4) = False
                     frm06010301_1.SSTab1.TabVisible(5) = False
                     frm06010301_1.SSTab1.TabVisible(6) = False
                     frm06010301_1.SetParent Me 'Add By Sindy 2022/5/3
                     frm06010301_1.Show
                  Case Else
                     MsgBox "點選的案件性質目前尚無申請書！"
                     Exit Sub
                  End Select
               '2015/11/26 END
            End Select
            
         '若未輸入特殊申請書
         Else
            Select Case pa(10)
               Case 讓與
                  'Modified by Morgan 2016/12/14 申請書在發文畫面產生,原程式不會出申請書且,不會帶出受讓人,若輸入6碼也只會存6碼會造成後面輸核准時檢查錯誤(FCP-051619)
                  'Modify By Sindy 2018/6/13 紙本,電子送件改在此作業操作
'                  MsgBox "讓與申請書請至發文畫面產生！", vbExclamation
                  frm06010302_1.Show
'                  Exit Sub
'                  'end 2016/12/14
               Case 變更
                  Set frm06010303_1.oParent = Me 'Add by Morgan 2011/10/5
                  frm06010303_1.m_CP118isY = "N" '非電子送件
                  frm06010303_1.LoadMe Me.Tag, Text1, Text2, Text3, Text4, 61
               'Modify by Morgan 2004/8/6
               '加翻譯
               'Modify by Morgan 2004/9/23
               '加檢視中說,製作中說
               'Modified by Morgan 2013/11/6 +235核對中說格式
               Case 翻譯, 檢視中說, 製作中說, 235
                  frm06010304_1.Show
               Case 補文件
                  '重新委任
                  If GetCP10(MSHFlexGrid1.TextMatrix(i, 4)) = "928" Then  '相關總收文號
                     frm06010304_2.Show '補委任書
                  Else
                     frm06010304_1.Show
                  End If
                  
               Case 延緩公告
                  frm06010305_1.Show
               Case 退費
                  frm06010306_1.Show
               Case 更改
                  frm06010307_1.Show
               Case 延期
                  frm06010308_1.Show '維持紙本申請書
               '2007/6/21 ADD BY SONIA 重新委任
               Case "928"
                  frm06010312_1.Show
               '2007/6/21 END
               Case 自請撤回 'Add by Morgan 2007/9/13
                  '重新委任
                  If GetCP10(MSHFlexGrid1.TextMatrix(i, 4)) = "928" Then '相關總收文號
                     frm06010313_1.Show
                  '其他
                  Else
                     frm06010310_1.SetParent Me 'Add By Sindy 2023/2/16
                     frm06010310_1.Show
                  End If
               Case Else
                  frm06010310_1.SetParent Me 'Add By Sindy 2023/2/16
                  frm06010310_1.Show
            End Select
         End If
         cmdok(1).SetFocus
         Me.Hide
      Case 1 '尋找
         Label4 = ""
         Label6 = ""
'         MSHFlexGrid1.Clear
         If Text3 = "" Then Text3 = "0"
         If Text4 = "" Then Text4 = "00"
         pa(1) = Text1
         pa(2) = Text2
         pa(3) = Text3
         pa(4) = Text4
   
         If pa(1) = "FCP" Then
            If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
               Label6.Caption = pa(22)
               Label4.Caption = pa(11)
               Text5.Text = pa(10)
            End If
         ElseIf pa(1) = "FG" Then
            If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
               Text5.Text = pa(10)
               Label4.Caption = pa(11)
            End If
         End If
         
         'Added by Lydia 2018/09/12 本案的新案進度是否為電子送件
         strExc(0) = "select 1 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10 IN (" & NewCasePtyList & ") and cp118 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            bNewCaseCp118 = True
         Else
            bNewCaseCp118 = False
         End If
         'end 2018/09/12
         
         AddCboName Combo1, pa(5), pa(6), pa(7)
         
         'strExc(0) = "select ''," & SQLDate("CP05") & ",cp09,cpm03,staff.st02 as st1,staff1.st02 as st2," & _
         '   "cp64,cp10 from caseprogress, casepropertymap,staff,staff staff1 where " & _
         '   ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp27 is null and " & _
         '   "(substr(cp09,1,1)='A' or substr(cp09,1,1)='B') and cp01=cpm01(+) and " & _
         '   "cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)"
         'Modify By Cheng 2002/04/12
'         strExc(0) = "select ''," & SQLDate("CP05") & ",cp09,cpm03,CP43,staff.st02 as st1,staff1.st02 as st2," & _
'            "cp64,cp10 from caseprogress, casepropertymap,staff,staff staff1 where " & _
'            ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and " & _
'            "(substr(cp09,1,1)='A' or substr(cp09,1,1)='B') and cp01=cpm01(+) and " & _
'            "cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)"
         strExc(0) = "select ''," & SQLDate("CP05") & ",cp09,cpm03,CP43,staff.st02 as st1,staff1.st02 as st2," & _
            "cp64,cp10 from caseprogress, casepropertymap,staff,staff staff1 where " & _
            ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and " & _
            "( cp09<'C' ) and cp01=cpm01(+) and " & _
            "cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)" & _
            " order by cp66 desc,substr(cp09,2) desc"
         intI = 0
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
         GridHead
         'Add By Cheng 2002/05/10
         '若只搜尋到一筆時直接勾選
         If Me.MSHFlexGrid1.Rows = 2 Then
            MSHFlexGrid1_Click
         End If
      
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
 Dim i As Integer
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = ""
      Next
   End With
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   'Combo1.ListIndex = 0
   Label4 = ""
   Label6 = ""
   'Modified by Lydia 2018/09/12
   'InitGrid 9, MSHFlexGrid1
   InitGrid 10, MSHFlexGrid1
   GridHead
   Text5.Text = strSrvDate(2)
   'Add By Cheng 2002/12/10
   SendKeys "{Tab}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060103_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   'Added by Morgan 2013/7/11 若是電子送件的案件，在各式申請書的"特殊申請書"欄位，預設"3.電子送件"
   If MSHFlexGrid1.TextMatrix(intLastRow, 0) = "v" Then
      Text6.Text = ""
      'Modify By Sindy 2018/1/10
      'Modify By Sindy 2018/6/15 + GetCP10(MSHFlexGrid1.TextMatrix(intLastRow, 2), "CP118") = "A"
      'Modified by Lydia 2018/09/12
      'If GetCP10(MSHFlexGrid1.TextMatrix(intLastRow, 2), "CP118") = "Y" Or _
      '   GetCP10(MSHFlexGrid1.TextMatrix(intLastRow, 2), "CP118") = "A" Then
      strExc(0) = GetCP10(MSHFlexGrid1.TextMatrix(intLastRow, 2), "CP118")
      If strExc(0) <> "" Then
      'end 2018/09/12
         Text6.Text = "3" '3.電子送件
      Else
      '2018/1/10 END
         'Modify By Sindy 2015/11/26
         If MSHFlexGrid1.TextMatrix(intLastRow, 8) = "101" Or _
            MSHFlexGrid1.TextMatrix(intLastRow, 8) = "102" Or _
            MSHFlexGrid1.TextMatrix(intLastRow, 8) = "103" Then
            'Text6.Text = "3"
            Text6.Text = "4"
         '2015/11/26 END
   '      ElseIf Text6.Text = "3" Then
   '         Text6.Text = ""
         'Added by Lydia 2018/09/12 本案的新案進度為電子送件,預設其他為電子送件
         ElseIf bNewCaseCp118 = True And InStr(cNewFrmList, "" & MSHFlexGrid1.TextMatrix(intLastRow, 8)) > 0 Then
             Text6.Text = "3"
         'end 2018/09/12
         End If
      End If
   End If
   'end 2013/7/11
   cmdok(0).SetFocus
End Sub

Private Sub Text1_GotFocus()
  TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "FCP" And Text1 <> "FG" Then
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
      .col = 4: .ColWidth(4) = 1200: .Text = "相關總收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1200: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1400: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 1400: .Text = "進度備註"
      .col = 8: .ColWidth(8) = 0 'Memo by Lydia 2018/09/12 CP10
      .Visible = True
      If .Rows > 1 Then .row = 1
   End With
End Sub

Private Sub Text2_GotFocus()
  TextInverse Text2
End Sub

Private Sub Text3_GotFocus()
  TextInverse Text3
End Sub

Private Sub Text4_GotFocus()
  TextInverse Text4
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2013/7/11 +電子送件 3
   'Modify by Sindy 2015/11/26 +新案申請書 4
   'If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
   If KeyAscii <> 8 And Chr(KeyAscii) <> "1" And Chr(KeyAscii) <> "2" And Chr(KeyAscii) <> "3" _
      And Chr(KeyAscii) <> "4" Then
      KeyAscii = 0
      Beep
   End If
End Sub

Public Sub ClearForm()
    'Modify By Cheng 2002/12/10
    '保留原輸入的系統類別
'   Text1 = Empty
   
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   Text5 = Empty
   Label4 = Empty
   Label6 = Empty
   Text6 = Empty
   Combo1.Clear
   'Modified by Lydia 2018/09/12
   'InitGrid 9, MSHFlexGrid1
   InitGrid 10, MSHFlexGrid1
   GridHead
   Text5.Text = strSrvDate(2)
   If Text1.Visible = True And Text1.Enabled = True Then Text1.SetFocus
   'Add By Cheng 2002/12/10
   If Text2.Visible = True And Text2.Enabled = True Then Text2.SetFocus
End Sub

'Add by Morgan 2007/9/13
'讀取案件性質
'Add By Sindy 2018/1/10 + Optional strCol As String = "CP10"
Private Function GetCP10(p_CP09 As String, Optional strCol As String = "CP10") As String
   Dim stSQL As String, iRtn As Integer

   GetCP10 = "" 'Add By Sindy 2018/1/10
   If p_CP09 <> "" Then
      stSQL = "select " & strCol & " from caseprogress where cp09='" & p_CP09 & "'"
      iRtn = 1
      Set AdoRecordSet3 = ClsLawReadRstMsg(iRtn, stSQL)
      If iRtn = 1 Then
         GetCP10 = "" & AdoRecordSet3.Fields(0)
      End If
   End If
End Function
