VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010502_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "核准函輸入"
   ClientHeight    =   5760
   ClientLeft      =   -2976
   ClientTop       =   4596
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9336
   Begin VB.CommandButton cmdOK 
      Caption         =   "內部收文(&E)"
      Height          =   400
      Index           =   3
      Left            =   5088
      TabIndex        =   16
      Top             =   72
      Width           =   1200
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   9
      Top             =   660
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   8
      Top             =   660
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   7
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   6
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4980
      TabIndex        =   5
      Top             =   660
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "frm04010502_2.frx":0000
      Left            =   1080
      List            =   "frm04010502_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1020
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8352
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6300
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7128
      TabIndex        =   1
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   720
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "1"
      Top             =   5400
      Width           =   255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3912
      Left            =   120
      TabIndex        =   10
      Top             =   1380
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   6900
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label Label8 
      Height          =   255
      Left            =   1770
      TabIndex        =   18
      Top             =   1020
      Width           =   7395
      Caption         =   "lblFM2"
      Size            =   "13044;450"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "訴願，行政訴訟，上訴的核准請改至  一般來函輸 1502撤銷原處分！"
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   2
      Left            =   480
      TabIndex        =   17
      Top             =   150
      Width           =   3000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   660
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Index           =   0
      Left            =   3900
      TabIndex        =   14
      Top             =   660
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1020
      Width           =   768
   End
   Begin VB.Label Label3 
      Caption         =   "結果:"
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   5400
      Width           =   492
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1:核准, 2:改變原處分)"
      Height          =   180
      Left            =   1080
      TabIndex        =   11
      Top             =   5400
      Width           =   1740
   End
End
Attribute VB_Name = "frm04010502_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/18 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、Label8
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim strReceiveNo As String, strTemp As String
'edit by nickc 2007/02/02 改成動態 tf_pa
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0 '確定
         FormConfirm
      Case 1 '回前畫面
         frm04010502_1.Show
         Unload Me
      Case 2 '結束
         Unload frm04010502_1
         Unload Me
      'Add By Cheng 2002/06/21
      Case 3 '內部收文
         mdiMain.mnu1102_Click 1
   End Select
End Sub

' 確認鈕
Private Sub FormConfirm()
Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String
Dim bolExcept As Boolean 'Added by Morgan 2012/2/7 是否例外
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "v" Then
            'Add by Morgan 2007/1/18
            If InStr(CaseMapIn, .TextMatrix(i, 12)) > 0 Then
               If pa(10) = "" Then
                  MsgBox "新申請案不可無申請日"
                  Exit Sub
               End If
            End If
            'end 2007/1/8
            
            'Added by Morgan 2013/1/14
            '檢查移轉或讓與的受讓人(5個)與基本檔是否相同
            If InStr("701,702,703,708", .TextMatrix(i, 12)) > 0 Then
               If PUB_ChkAsignCaseCustNo(.TextMatrix(i, 1)) = False Then
                  Exit Sub
               End If
            End If
            'end 2013/1/14
            
            'Add by Morgan 2007/10/26 專利權讓與的核准要有專用期，繳費年度，證書號，年費期限
            If Not (pa(9) = "000" And Len(pa(11)) > 9) Then
               If .TextMatrix(i, 12) = 專利權讓與 Then
                  If Val(DBDATE(pa(25))) < strSrvDate(1) Then
                     MsgBox "專利權讓與之核准不可無專用期(或已過期)！"
                     Exit Sub
                  End If
                  If pa(22) = "" Then
                     MsgBox "專利權讓與之核准不可無證書號！"
                     Exit Sub
                  End If
                  
                  If pa(9) <> "013" Then 'Added by Morgan 2020/4/15 香港案自動發證不必檢查 Ex:P-113869
                  
                     If pa(72) = "" Then
                        MsgBox "專利權讓與之核准不可無繳費年度！"
                        Exit Sub
                     End If
                     If PUB_708FeeCheck(pa) = False Then
                        MsgBox "專利權讓與之核准必須有年費期限！"
                        Exit Sub
                     End If
                     
                  End If 'Added by Morgan 2020/4/15
                  
               'Added by Morgan 2012/1/18 非爭議程序的核准若有公告日時控制都要有年費期限--玲玲
               ElseIf Left(.TextMatrix(i, 12), 1) <> "8" And Val(pa(14)) > 0 Then
                  strExc(0) = "select 1 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='605' and cp57||cp27 is null" & _
                     " union select 1 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07='605' and np06 is null"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 0 Then
                     'Added by Morgan 2012/2/7 +排除基本檔已閉卷且最後年費期限為閉卷案件--玲玲 P-093438
                     bolExcept = False
                     If pa(57) = "Y" Then
                        'Added by Morgan 2012/2/29 +排除已通知專利權消滅的案件--玲玲 P-090774
                        strExc(0) = "select cp09 from caseprogress where cp01='" & pa(1) & "' and  cp02='" & pa(2) & "' and  cp03='" & pa(3) & "' and  cp04='" & pa(4) & "' and cp10='1604'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           bolExcept = True
                        Else
                        'end 2012/2/29
                        
                           strExc(0) = "select np06,np09 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07='605' order by np09 desc"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              If RsTemp(0) = "N" Then
                                 bolExcept = True
                              End If
                           End If
                        End If
                     End If
                     '2012/5/14 ADD BY SONIA 繳費年度已繳至專利權期滿,不必再補基本資料P-065005
                     'Modified by Morgan 2013/11/26 改為 >= ,Ex.P-74182
                     If Val(PUB_GetNextFeeDate(pa())) >= Val(DBDATE(pa(25))) Then
                        bolExcept = True
                     End If
                     '2012/5/14 END
                     'add by sonia 2014/5/8 放棄專利權429也不必補年費期限 P-096341
                     If .TextMatrix(i, 12) = "429" Then
                        bolExcept = True
                     End If
                     'end 2014/5/8
                     
                     'Added by Morgan 2024/8/27 X81780 113/8/6大批讓與後仍為年費不續辦，故移除輸入[讓與准]時提醒恢復年費之控管--陳亭妙,李道昀
                     If ChangeCustomerL(pa(26)) = "X81780000" And InStr("701,708", .TextMatrix(i, 12)) > 0 Then
                        strExc(0) = "select cp05,cp65 from caseprogress where cp09='" & .TextMatrix(i, 1) & "'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If RsTemp("cp05") = 20240806 And RsTemp("cp65") = "QPGMR" Then
                              bolExcept = True
                           End If
                        End If
                     End If
                     'end 2024/8/27
                     
                     If bolExcept = False Then
                     'end 2012/2/7
                        
                        'Added by Morgan 2024/9/26 寰華案改詢問方式--陳亭妙/何淑華
                        If Left(Pub_StrUserSt03, 2) = "F2" Then
                           If MsgBox("請先和承辦確認，本案是否需補輸基本資料並管制年費期限？", vbYesNo + vbDefaultButton1 + vbExclamation) = vbYes Then
                              Exit Sub
                           End If
                        Else
                        'end 2024/9/26
                           
                           MsgBox "請先補輸基本資料並管制年費期限！", vbExclamation
                           Exit Sub
                           
                        End If
                     End If 'Added by Morgan 2012/2/7
                  End If

               End If
            End If
            'end 2007/10/26
            
            '2011/7/11 ADD BY SONIA
            If pa(9) = "000" Then
               '台灣非新申請案或改請案檢查來函記錄檔期限
               'modify by sonia 2014/5/29 剔除再審案P-099529
               If InStr(CaseMapIn, .TextMatrix(i, 12)) = 0 And (.TextMatrix(i, 12) < "3" Or .TextMatrix(i, 12) >= "4") And .TextMatrix(i, 12) <> "107" Then
                  If ClsLawChkMRec(TransDate(frm04010502_1.Text5, 2), Text2 & Text3 & Text4 & Text5, strTmp(1), strTmp(2)) Then
                     If strTmp(1) <> "" Then
                        If MsgBox("與櫃台之來函收文記錄期限 ( " & TransDate(strTmp(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
                     End If
                  'Modified by Morgan 2014/5/5 排除無期限電子公文
                  'Else
                  ElseIf frm04010502_1.m_DocNo = "" Then
                  'end 2014/5/5
                     If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
                  End If
               '新申請案或改請案只檢查是否有來函記錄  2014/5/29+再審案P-099529
               Else
                  If ClsLawChkMRec(TransDate(frm04010502_1.Text5, 2), Text2 & Text3 & Text4 & Text5, strTmp(1), strTmp(2)) Then
                  'Modified by Morgan 2014/5/5 排除無期限電子公文
                  'Else
                  'Modified by Morgan 2014/7/21 +電子公文申請案也要檢查
                  'modify by sonia 2018/2/8 電子公文都不檢查來函記錄檔
                  'ElseIf frm04010502_1.m_DocNo = "" Or InStr(CaseMapIn, .TextMatrix(i, 12)) > 0 Then
                  ElseIf frm04010502_1.m_DocNo = "" Then
                  'end 2014/5/5
                     If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
                  End If
               End If
               
               'Added by Morgan 2014/8/4
               '申請案核准期限檢查
               If Len(pa(11)) = 9 And InStr(CaseMapIn, .TextMatrix(i, 12)) > 0 And Val(frm04010502_1.m_DeadLine) > 0 Then
                  'Modified by Morgan 2014/8/18 因有日的期限改單位也要判斷
                  If Val(frm04010502_1.m_DeadLine) <> 3 Or Right(frm04010502_1.m_DeadLine, 1) <> "月" Then
                     If MsgBox("電子公文來函期限 ( " & frm04010502_1.m_DeadLine & "個月 ) 並非 3 個月，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
                  End If
               End If
               'end 2014/8/4
            End If
            '2011/7/11 END
            
            bolChk = True
            Me.Tag = .TextMatrix(i, 1)
            'Modify By Cheng 2002/06/21
'            strExc(5) = .TextMatrix(i, 3)
            strExc(5) = .TextMatrix(i, 4)
            'Modify By Cheng 2003/01/27
            '因加對造號數欄故位移一位
'            'Add By Cheng 2002/06/21
'            frm04010502_3.m_strCP10 = "" & .TextMatrix(i, 11)
            frm04010502_3.m_strCP10 = "" & .TextMatrix(i, 12)
            Exit For
         End If
      Next
   End With
   If bolChk = False Then
      MsgBox "請選擇資料 !", vbInformation
      Exit Sub
   End If
   'Added by Morgan 2014/1/14
   frm04010502_3.m_AppNo = frm04010502_1.m_AppNo
   frm04010502_3.m_DocNo = frm04010502_1.m_DocNo
   frm04010502_3.m_DocWord = frm04010502_1.m_DocWord
   'end 2014/1/14
   'Add By Sindy 2016/10/5
   frm04010502_3.m_strIR01 = m_strIR01
   frm04010502_3.m_strIR02 = m_strIR02
   frm04010502_3.m_strIR03 = m_strIR03
   frm04010502_3.m_strIR04 = m_strIR04
   '2016/10/5 END
   frm04010502_3.Show
   Me.Hide
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label8 = pa(5)
      Case "英"
         Label8 = pa(6)
      Case "日"
         Label8 = pa(7)
   End Select
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   
   pa(1) = strExc(1)
   pa(2) = strExc(2)
   pa(3) = strExc(3)
   pa(4) = strExc(4)
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010502_1.m_strIR01
   m_strIR02 = frm04010502_1.m_strIR02
   m_strIR03 = frm04010502_1.m_strIR03
   m_strIR04 = frm04010502_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
   
   ReadPatent 1
End Sub

Private Sub ReadPatent(ByVal iSitu As Integer)
 Dim Lbl As LABEL, txt As TextBox, i As Integer
 Dim strTmp As String
   Label8 = ""
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Label8 = pa(5)
      Text1 = pa(11)
   End If
   If iSitu = 1 Then
      'Modify By Cheng 2002/04/15
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is null and " & _
'         "(substr(cp09,1,1)='A' or substr(cp09,1,1)='B' or " & _
'         "(substr(cp09,1,1)='C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))"
        'Modify By Cheng 2003/07/25
        '加案件性質
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is null and " & _
'         "( cp09<'C' or " & _
'         "( cp09>'C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))"
      '2010/11/12 MODIFY BY SONIA 訴願或行政訴訟或上訴的核准請改輸  一般來函的撤銷原處分
      'Modified by Morgan 2025/1/2 行政訴訟上訴507改只限制台灣案，因為大陸案要可以輸核准 Ex:P-122180-韻丞
      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is null and CP10<>'501' AND CP10<>'503' AND " & _
         "( cp09<'C' or " & _
         "( cp09>'C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "' Or CP10='1504' Or CP10='1505' Or CP10='1211' Or CP10='1210' )))"
         
      If pa(9) = 台灣國家代號 Then
         strExc(1) = strExc(1) & " AND CP10<>'507'"
      End If
      'end 2025/1/2
   Else
      'Modify By Cheng 2002/04/15
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is not null and " & _
'         "(substr(cp09,1,1)='A' or substr(cp09,1,1)='B' or " & _
'         "(substr(cp09,1,1)='C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))"
        'Modify By Cheng 2003/07/25
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is not null and " & _
'         "( cp09<'C' or " & _
'         "( cp09>'C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))"
      '2010/11/12 MODIFY BY SONIA 訴願或行政訴訟或上訴的核准請改輸  一般來函的撤銷原處分
      'Modified by Morgan 2025/1/2 行政訴訟上訴507改只限制台灣案，因為大陸案要可以輸核准 Ex:P-122180-韻丞
      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is not null and CP10<>'501' AND CP10<>'503' AND " & _
         "( cp09<'C' or " & _
         "( cp09>'C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "' Or CP10='1504' Or CP10='1505' Or CP10='1211' Or CP10='1210' )))"
      
      If pa(9) = 台灣國家代號 Then
         strExc(1) = strExc(1) & " AND CP10<>'507'"
      End If
      'end 2025/1/2
   End If
   
   If pa(9) = 台灣國家代號 Then
      strTmp = "CPM03"
   Else
      strTmp = "CPM04"
   End If
   
   'Modify By Cheng 2002/06/21
'   strExc(2) = "'',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP19,CP64 " & _
'      "from caseprogress,casepropertymap,CUSTOMER"
' 91.09.13 modify by louis (排序)
   'strExc(2) = "'',CP09," & strTmp & ",CP43," & _
   '   "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
   '   SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP19,CP64,CP10 " & _
   '   "from caseprogress,casepropertymap,CUSTOMER"
    'Modify By Cheng 2003/01/27
    '加對造號數
'   strExc(2) = "'',CP09," & strTmp & ",CP43," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP19,CP64,CP10 " & _
'      ", DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
'      "from caseprogress,casepropertymap,CUSTOMER"
   strExc(2) = "'',CP09," & strTmp & ",CP43," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40),CP36," & _
      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP19,CP64,CP10 " & _
      ", DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
      "from caseprogress,casepropertymap,CUSTOMER"
      
   ' 91.09.13 modify by louis (排序)
   'strExc(0) = "select " & strExc(2) & " where substr(cp10,1,1)='1' and " & strExc(1) & " and " & _
   '   "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)" & _
   '   " and (cp01,cp02,cp03,cp04) not in " & _
   '   "(select cp01,cp02,cp03,cp04 from caseprogress where substr(cp10,1,1)='3' and " & _
   '   strExc(1) & ") union " & _
   '   "select " & strExc(2) & " where substr(cp10,1,1)<>'1' and " & strExc(1) & " and " & _
   '   "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)"
   '94.2.2 modify by sonia
   'strExc(0) = "select " & strExc(2) & " where substr(cp10,1,1)='1' and " & strExc(1) & " and " & _
   '   "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)" & _
   '   " and (cp01,cp02,cp03,cp04) not in " & _
   '   "(select cp01,cp02,cp03,cp04 from caseprogress where substr(cp10,1,1)='3' and " & _
   '   strExc(1) & ") union " & _
   '   "select " & strExc(2) & " where substr(cp10,1,1)<>'1' and " & strExc(1) & " and " & _
   '   "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) " & _
   '   "ORDER BY SORTFIELD DESC "
   strExc(0) = "select " & strExc(2) & " where substr(cp10,1,1)='1' and " & strExc(1) & " and " & _
      "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)" & _
      " and (cp01,cp02,cp03,cp04) not in " & _
      "(select cp01,cp02,cp03,cp04 from caseprogress where substr(cp10,1,1)='3' and " & _
      strExc(1) & ") union " & _
      "select " & strExc(2) & " where (substr(cp10,1,1)<>'1' or cp10='107') and " & strExc(1) & " and " & _
      "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) " & _
      "ORDER BY SORTFIELD DESC "
   '94.2.2 end
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   Combo1.ListIndex = 0
   
   ' 若只有一筆資料時自動選取第一筆
   If MSHFlexGrid1.Rows = 2 Then
      MSHFlexGrid1.row = 1
      'Add by Morgan 2003/11/27
      If GridDataCheck() = False Then Exit Sub
      'End
      GridClick MSHFlexGrid1, intLastRow, 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010502_2 = Nothing
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      'Add By Cheng 2002/06/21
      '加相關總收文號
      .col = 3: .ColWidth(3) = 1000: .Text = "相關總收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1000: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
        'Add By Cheng 2003/01/27
        '加對造號數
      .col = 5: .ColWidth(5) = 1000: .Text = "對造號數"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 800: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 800: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 800: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 9: .ColWidth(9) = 600: .Text = "結果"
      .CellAlignment = flexAlignCenterCenter
      .col = 10: .ColWidth(10) = 800: .Text = "後金"
      .CellAlignment = flexAlignCenterCenter
      .col = 11: .ColWidth(11) = 1000: .Text = "進度備註"
      .CellAlignment = flexAlignCenterCenter
      'Add By Cheng 2002/06/21
      .col = 12: .ColWidth(12) = 0: .Text = "案件性質代號"
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()
   
    'Add by Morgan 2003/11/25
   If GridDataCheck() = False Then Exit Sub
   '---End
   
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK(0).SetFocus
   
End Sub
'Add by Morgan 2003/11/25
Private Function GridDataCheck() As Boolean
   
   Dim strSql As String, strTemp As String, bolRtn As Boolean
   
   bolRtn = False
   If (MSHFlexGrid1.row = 0) Then
      bolRtn = True
   ElseIf (pa(9) <> "000") Then
      bolRtn = True
   Else
      MSHFlexGrid1.Recordset.Move MSHFlexGrid1.row - 1, 1
      strTemp = MSHFlexGrid1.Recordset.Fields("CP10")
      If (Len(strTemp) = 3 And strTemp >= "101" And strTemp <= "105") Then
         strTemp = pa(11)
         If (Trim(strTemp) = Empty) Then
            bolRtn = True
         Else
            'Modify by Morgan 2004/6/16
            'Modified by Morgan 2012/12/27
            '核准無須再檢查
            'Dim stCaseNo As String
            'If PUB_ChkPriDate(strTemp, stCaseNo) Then
            '   MsgBox "此案已被 " & stCaseNo & " 主張國內優先權且自申請日起逾15個月，不可輸入准駁！", vbCritical
            'Else
            '   bolRtn = True
            'End If
            bolRtn = True
            'end 2012/12/27
            'end
         End If
      Else
         bolRtn = True
      End If
   End If
   GridDataCheck = bolRtn
   
End Function

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   
   'Moidfy by Morgan 2003/12/22
   If KeyAscii = 49 Then
         Text6 = ""
         ReadPatent 1
   ElseIf KeyAscii = 50 Then
      '2006/8/15 MODIFY BY SONIA 開放大陸案 P-65644
      'If pA(9) <> "000" Then
      If pa(9) <> "000" And pa(9) <> "020" Then
         MsgBox "申請國家不為台灣不可改變原處分！", vbCritical
         KeyAscii = 0
      Else
         Text6 = ""
         ReadPatent 2
      End If
   ElseIf KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   
End Sub
