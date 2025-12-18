VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010608_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利權消滅函輸入"
   ClientHeight    =   3780
   ClientLeft      =   156
   ClientTop       =   960
   ClientWidth     =   9012
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   9012
   Begin VB.CommandButton Cmd1 
      Caption         =   "整批"
      Height          =   345
      Left            =   6840
      TabIndex        =   29
      Top             =   510
      Width           =   885
   End
   Begin VB.TextBox Text26 
      Height          =   270
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   2
      Left            =   8355
      TabIndex        =   5
      Top             =   90
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6810
      TabIndex        =   3
      Top             =   90
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   7425
      TabIndex        =   4
      Top             =   90
      Width           =   900
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   10
      Top             =   150
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   9
      Top             =   150
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   8
      Top             =   150
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   7
      Top             =   150
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4680
      TabIndex        =   6
      Top             =   150
      Width           =   1575
   End
   Begin MSForms.TextBox Text29 
      Height          =   795
      Left            =   1500
      TabIndex        =   1
      Top             =   1935
      Width           =   6795
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "11986;1402"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text31 
      Height          =   795
      Left            =   1500
      TabIndex        =   2
      Top             =   2790
      Width           =   6795
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "11986;1402"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   330
      Left            =   1080
      TabIndex        =   15
      Top             =   420
      Width           =   5655
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "9975;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   135
      X2              =   8800
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   2
      X1              =   135
      X2              =   8800
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Label Label43 
      Caption         =   "進度備註:"
      Height          =   255
      Left            =   90
      TabIndex        =   28
      Top             =   1980
      Width           =   855
   End
   Begin VB.Label Label38 
      Caption         =   "專利權消滅日:"
      Height          =   255
      Left            =   90
      TabIndex        =   27
      Top             =   1590
      Width           =   1215
   End
   Begin VB.Label Label46 
      Caption         =   "案件備註:"
      Height          =   255
      Left            =   90
      TabIndex        =   26
      Top             =   2820
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(Y:閉卷)"
      Height          =   180
      Index           =   4
      Left            =   7920
      TabIndex        =   25
      Top             =   1050
      Width           =   645
   End
   Begin VB.Label lblPA57 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Left            =   7200
      TabIndex        =   24
      Top             =   1050
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否閉卷"
      Height          =   180
      Index           =   3
      Left            =   6240
      TabIndex        =   23
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   6
      Left            =   4680
      TabIndex        =   22
      Top             =   1050
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   5
      Left            =   1080
      TabIndex        =   21
      Top             =   1050
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日"
      Height          =   180
      Index           =   2
      Left            =   3600
      TabIndex        =   20
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "收文號"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   1050
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   2
      Left            =   4680
      TabIndex        =   18
      Top             =   810
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   1
      Left            =   1080
      TabIndex        =   17
      Top             =   810
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   810
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   3600
      TabIndex        =   13
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   450
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Left            =   3600
      TabIndex        =   11
      Top             =   810
      Width           =   585
   End
End
Attribute VB_Name = "frm06010608_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/23 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim strReceiveNo As String, strTemp As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String, cp(1 To T_CP) As String
Dim pa() As String, cp() As String

Dim intWhere As Integer, intLastRow As Integer
Public MPa9 As String

'Add By Cheng 2002/01/28
Dim m_strCP09ByCheng As String '總收文號

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim m_928Upd As Boolean '是否更新重新委任准駁
Dim m_928CP09 As String '重新委任收文號
Dim m_CP16 As String '預設請款金額
'Added by Morgan 2017/5/10 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_DocDate As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/5/10


'Added by Lydia 2019/08/26 整批匯入PDF
Private Sub Cmd1_Click()
Dim strFName As String
Dim pCP09 As String
Dim intCount As Integer
    
    'Modified by Lydia 2024/07/22 改用變數
    'strFName = Dir("\\TYPING2\Pat_cancel\FCP*.pdf")
    strFName = Dir("\\" & strTyping2Path & "\Pat_cancel\FCP*.pdf")
    
    Do While strFName <> ""
         Call ChgCaseNo(Mid(strFName, 1, InStr(strFName, ".") - 1) & "000", strExc)
         
         If strExc(1) <> "" And strExc(2) <> "" Then
             If PUB_ChkCPExist(strExc, "1604", , pCP09) = True Then
                 If AutoSavePdf_FCP(strExc(1), strExc(2), strExc(3), strExc(4), pCP09, "1604") = True Then
                    intCount = intCount + 1
                 Else
                    Exit Sub
                 End If
             Else
                 MsgBox "掃瞄檔案有問題：" & strFName, vbCritical
                 GoTo JumpToExit
             End If
         End If
         'Modified by Lydia 2024/07/22 改用變數
         'strFName = Dir("\\TYPING2\Pat_cancel\FCP*.pdf")
         strFName = Dir("\\" & strTyping2Path & "\Pat_cancel\FCP*.pdf")
    Loop
    
JumpToExit:
    If intCount > 0 Then
        MsgBox "匯入" & intCount & "筆！", vbInformation, "整批匯入完畢"
    End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim adoRst As ADODB.Recordset   'add by sonia 2016/11/22
   
   Select Case Index
      Case 0
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         'Add by Sindy 2021/11/23 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
            
         'Added by Lydia 2019/08/26
         If AutoSavePdf_FCP(pa(1), pa(2), pa(3), pa(4), m_strCP09ByCheng, "1604") = False Then
         End If
         
         'Added by Morgan 2012/11/5
         If Left(pa(75), 6) = "Y53309" Then
            MsgBox "本案需調卷轉承辦組報告並寄代！", vbInformation
         End If
         'end 2012/11/5
         
         'add by sonia 2016/11/22
         '檢查是否有收文申請復權414 FCP-034498
         strExc(0) = "select cp09 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                     " and cp10='414' and (cp159=0 or (cp27>0 and cp57>0))"
         intI = 1
         Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
             MsgBox "曾收文 申請復權, 請向承辦確認是否要通知消滅函！", vbInformation
         End If
         'end 2016/11/22
         
         Unload frm06010608_2
         Unload Me
         
         'Modified by Morgan 2017/5/10 電子公文
         'frm06010608_1.Show
         'frm06010608_1.Clear
         If m_DocNo <> "" Then
            Unload frm06010608_1
            frm060119.GoNext
         Else
            frm06010608_1.Show
            frm06010608_1.Clear
         End If
         'end 2017/5/10
         
      Case 1
         'frm06010608_2.Show
         frm06010608_1.Show
         Unload frm06010608_2
         Unload Me
      Case 2
         Unload frm06010608_2
         Unload frm06010608_1
         Unload Me
   End Select
End Sub

Private Function FormSave() As Boolean
   Dim intStep As Integer, strTxt(1 To 10) As String, strTmp As String, i As Integer
   'edit by nickc 2007/02/02
   'Dim Ncp(1 To T_CP) As String
   Dim Ncp() As String
   ReDim Ncp(1 To TF_CP) As String
   
   Dim strSql As String
   Dim strNP22 As String
   'Add By Cheng 2002/01/29
   Dim BlnCheck As Boolean '判斷是否有勾選本案期限
   Dim strDate1 As String '本所期限
   Dim strDate2 As String '法定期限
      
   m_928Upd = PUB_928Check(pa, m_928CP09) 'Add by Morgan 2007/7/17
   
 '911105 nick
 FormSave = True
 On Error GoTo CheckingErr
cnnConnection.BeginTrans
   
   'Add by Morgan 2007/7/17
   If m_928Upd = True And m_928CP09 <> "" Then
      PUB_928Update pa, m_928CP09
   End If
   'end 2007/7/17
   
   strNP22 = Empty
   'intMax = objPublicData.GetNextProgressNo 91.9.15 modify by sonia 不需要
   intStep = 1
   
   '1
      Ncp(1) = cp(1)
      Ncp(2) = cp(2)
      Ncp(3) = cp(3)
      Ncp(4) = cp(4)
      Ncp(5) = Label3(6)
'      Ncp(6) = Text14(0)
'      Ncp(7) = Text14(1)
'      Ncp(8) = Text9
      'Modify by Morgan 2011/2/24 修正百年收文號問題
      'Ncp(9) = "C" & Left(strSrvDate(2), 2)
      'Modified by Lydia 2019/05/31
      'Ncp(9) = "C" & CompAutoNumberYear(GetTaiwanThisYear)
      Ncp(9) = AutoNo("C", 6)
      
      Ncp(10) = "1604"
      '2012/10/2 MODIFY BY SONIA
      'Ncp(12) = cp(12)
      Ncp(12) = GetSalesArea(PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4)))
      '2012/10/2 END
        'Modify By Cheng 2003/04/07
        '智權人員存國家檔FCP承辦智權人員
'      Ncp(13) = cp(13)
      Ncp(13) = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
'      If Text8 <> "" Then
'         Ncp(14) = Text16
'         Ncp(48) = Text17
'      Else
         'Modified by Lydia 2019/05/31 外專程序工作大項先不上發文日(整批發文)
         'Ncp(14) = strUserNum
         Ncp(14) = Pub_GetSpecMan("外專程序-專利權消滅")
'      End If
      
      'Modify by Morgan 2007/7/24 改為輸N與欄位一致
      'If Text27(1) = "Y" Then
      '   Ncp(20) = ""
      'Else
      '   Ncp(20) = "N"
      'End If
      Ncp(20) = "N" 'Text27(1) '是否向客戶請款
      If Ncp(20) = "" Then
         Ncp(16) = Val(m_CP16)
         Ncp(17) = 0
         Ncp(18) = Val(m_CP16) / 1000
      End If
      'end 2007/7/24
      
'      If Text7 = 撤銷原處分 Then Ncp(24) = Text13
'      If "1604" = 專利權消滅 Then Ncp(25) = Text26
      Ncp(25) = Text26
'      Ncp(26) = Text18
      'If Text8 = "" Then Ncp(27) = strSrvDate(2)
      'Modified by Lydia 2019/05/31 外專程序工作大項先不上發文日(整批發文)
      'Ncp(27) = strSrvDate(2)
      Ncp(27) = Empty
      Ncp(32) = "N"
'      Ncp(36) = Text19
'      For i = 0 To 5
'         ' 91.01.22 modify by louis
'         'Ncp(i + 37) = Text20(i)
'         Ncp(i + 37) = ChgSQL(Text20(i))
'      Next
      Ncp(43) = cp(9)
      Ncp(64) = Text31
      
      ' 承辦期限
      'Modified by Lydia 2019/05/31 外專程序工作大項先不上發文日(整批發文)
      'Ncp(48) = Empty
      'Added by Lydia 2019/06/17 已上閉卷的案件，各項大批進度檔發文日請先上111111
      'Modified by Lydia 2019/10/16 排除-已銷卷,但未閉卷( ex.FCP-21504在108/6/6銷卷未閉卷,但是在108/10/3收到專利權消滅函)
      'If Trim(pa(57) & pa(108)) <> "" Then
      If Trim(pa(57)) <> "" Then
          Ncp(27) = "19221111"
          Ncp(48) = Empty
      Else
      'end 2019/06/17
          'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
          Ncp(48) = PUB_GetWorkDay1(CompDate(2, 10, strSrvDate(1)), True)
      End If 'end 2019/06/17
      
      'Add By Cheng 2002/01/28
      m_strCP09ByCheng = Empty
      
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.SaveNewCaseProgressDatabase("C", Ncp, intWhere, m_strCP09ByCheng) Then
      If Not ClsPDSaveNewCaseProgressDatabase("C", Ncp, intWhere, m_strCP09ByCheng) Then
         Exit Function
      End If

   '2
'   If Text27(0) = "Y" Then '閉卷
      strTxt(intStep) = "UPDATE PATENT SET PA57='Y',PA17='N' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      '911105 nick transation
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      
      '94.1.20 MODIFY BY SONIA
      strTxt(intStep) = "UPDATE PATENT SET PA58=" & TransDate(Label3(6), 2) & ", PA59='89' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4)) & " AND PA59 IS NULL "
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      '94.1.20 END
'   End If

   'Added by Morgan 2017/5/10 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, Ncp(9), pa(1), pa(2), pa(3), pa(4), Ncp(10)
   'Added by Morgan 2021/6/11 紙本公文--何淑華
   Else
      PUB_FCPOAInform Ncp(9), pa(1), pa(2), pa(3), pa(4), Ncp(10)
   End If
   'end 2017/5/10
   
   PUB_DualCaseInform Ncp(9) 'Added by Morgan 2022/4/6
  
   '911105 nick transation
   
   'FormSave = objLawDll.ExecSQL(intStep - 1, strTxt)
   cnnConnection.CommitTrans
   ' 列印接洽結案單
   'Modify By Cheng 2001/12/20
'   If IsEmptyText(strNP22) = False Then
'      g_PrtForm001.PrintForm strNP22, pa(1), pa(2), pa(3), pa(4)
'   End If
'911105 nick
   Exit Function
CheckingErr:
   
   cnnConnection.RollbackTrans
   FormSave = False
End Function

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
ReDim cp(1 To TF_CP) As String
End Sub

Private Sub Form_Load()
Dim ret As Long
Dim strTmp As String
   
   MoveFormToCenter Me
   
   intWhere = 國外_FC
   With frm06010608_2
      pa(1) = .Text2
      pa(2) = .Text3
      pa(3) = .Text4
      pa(4) = .Text5
      strReceiveNo = .Tag
      ReadPatent
   End With
   Combo1.ListIndex = 0
   
   'Modify By Cheng 2002/05/31
'   Text6 = strSrvDate(2)
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   '92.11.11 add by sonia
   'Modify by Morgan 2007/7/24
   'Text27(1) = ""
   'Modify by Morgan 2008/3/27 +pa75
   'Modify by Morgan 2008/4/10 +本所案號
'   Text27(1) = PUB_GetCP20(Text2, "1601", m_CP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
   '92.11.11 END
   
   'Added by Morgan 2017/5/10 電子公文
   If m_DocDate <> "" And Text26.Locked = False Then
      Text26 = TransDate(m_DocDate, 1)
   End If
   'end 2017/5/10
   
   'Added by Lydia 2019/08/26 Gill有8/13和8/15整批尚未匯入
   If Pub_StrUserSt03 <> "M51" Then
       Cmd1.Visible = False
   End If
   
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, i As Integer
Dim rsTemp1 As New ADODB.Recordset, bolTmp As Boolean
Dim adoRst As ADODB.Recordset
Dim iPos As Integer, strOldDate As String
   
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   Label3(6) = frm06010608_1.Text5
   Label3(5) = strReceiveNo
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   ' 90.06.26 modify by louis 是否閉卷
   lblPA57 = Empty
   If pa(1) = "FCP" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         MPa9 = pa(9)
         Label3(2) = pa(10)
         Text1 = pa(11)
         Text31 = pa(91)
            'Modify By Cheng 2002/12/19
            '是否閉卷預設為"Y"
'         Text27(0) = pa(57)
'            Me.Text27(0).Text = "Y"
'         If pa(71) = "" Then
'            If pa(75) = "" Then
'               strExc(0) = "SELECT CU75 FROM CUSTOMER WHERE " & ChgCustomer(pa(26))
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'               If intI = 1 And Not IsNull(RsTemp.Fields(0)) Then Text27(2) = RsTemp.Fields(0)
'            Else
'               strExc(0) = "SELECT FA42 FROM FAGENT WHERE " & ChgFagent(pa(75))
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'               If intI = 1 And Not IsNull(RsTemp.Fields(0)) Then Text27(2) = RsTemp.Fields(0)
'            End If
'         Else
'            Text27(2) = pa(71)
'         End If
         ' 90.06.26 modify by louis 是否閉卷
         lblPA57 = pa(57)
         
         'Add By Sindy 2015/10/5 預設消滅日
         If Val(strSrvDate(2)) > Val(pa(25)) And Val(pa(25)) > 0 Then '期滿者為專用期間(止日)隔一天
            Text26 = ChangeWDateStringToTString(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(pa(25)))))
         Else
            '檢查是否有一年內1605.通知年費逾期的期限
            strExc(0) = "select cp64 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                        " and cp10='1605' and cp05>=" & DBDATE(DateAdd("yyyy", -1, ChangeWStringToWDateString(strSrvDate(1))))
            intI = 1
            Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               iPos = InStr(adoRst.Fields("cp64"), "原繳費期限:")
               If iPos > 0 Then
                  strOldDate = Val(Mid(adoRst.Fields("cp64"), iPos + 6))
                  If Val(strOldDate) > 0 Then
                     Text26 = ChangeWDateStringToTString(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(strOldDate))))
                  End If
               End If
            Else
               '檢查是否有未領證601.領證及繳年費
               strExc(0) = "select np09 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'" & _
                           " and np07='601' and (np06 is null or np06='N')"
               intI = 1
               Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  Text26 = ChangeWDateStringToTString(DateAdd("d", 1, ChangeWStringToWDateString(adoRst.Fields("np09"))))
               Else
                  '檢查是否有未繳實審416.實體審查
                  strExc(0) = "select np09 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'" & _
                              " and np07='416' and (np06 is null or np06='N')"
                  intI = 1
                  Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     Text26 = ChangeWDateStringToTString(DateAdd("d", 1, ChangeWStringToWDateString(adoRst.Fields("np09"))))
                  End If
               End If
            End If
         End If
         '2015/10/5 END
      End If
   Else
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.ReadServicePracticeDatabase(pA(), intWhere) Then
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         
         Label3(2) = pa(10)
         Text1 = pa(11)
         Text31 = pa(18)
            'Modify By Cheng 2002/12/19
            '是否閉卷預設為"Y"
'         Text27(0) = pa(15)
'            Me.Text27(0).Text = "Y"
            'Add By Cheng 2002/12/19
            lblPA57 = pa(15)
      End If
   End If
   ' 90.06.26 modify by louis, 下一程序名稱帶出來
   'strExc(0) = "SELECT '',CPM03," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP13," & _
   '   "NP14," & SQLDate("NP11") & ",NP22 FROM NEXTPROGRESS,CASEPROPERTYMAP " & _
   '   "WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
   '   " AND (NP06<>'Y' OR NP06 IS NULL) AND NP01=CPM01(+) AND NP07=CPM02(+)"
   strExc(0) = "SELECT '',CPM03," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP13," & _
      "NP14," & SQLDate("NP11") & ",NP22 FROM NEXTPROGRESS,CASEPROPERTYMAP " & _
      "WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND (NP06<>'Y' OR NP06 IS NULL) AND NP02=CPM01(+) AND NP07=CPM02(+)"
   intI = 1
   
   cp(9) = strReceiveNo
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.ReadCaseProgressDatabase(cp(), intWhere) Then
   If ClsPDReadCaseProgressDatabase(cp(), intWhere) Then
      If cp(10) <> "" Then
         If pa(9) = 台灣國家代號 Then
            bolTmp = False
         Else
            bolTmp = True
         End If
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(cp(1), cp(10), strExc(0), BolTmp) Then Label3(1) = strExc(0)
         If ClsPDGetCaseProperty(cp(1), cp(10), strExc(0), bolTmp) Then Label3(1) = strExc(0)
      End If
      ' 90.06.26 modify by louis 讀檔時不須帶入承辦人, 改在輸入完來函性質後才去帶出承辦人
      'If Left(cp(10), 1) = "1" Then
      '   strExc(0) = "SELECT CP14 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10=" & 翻譯
      '   intI = 1
      '   Set rsTemp1 = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      '   If intI = 1 Then
      '      If Not IsNull(rsTemp1.Fields(0)) Then Text16 = rsTemp1.Fields(0): ChgType 16
      '   End If
      'Else
      '   If cp(14) <> "" Then Text16 = cp(14): ChgType 16
      'End If
      
      ' 90.10.09 modify by louis
      'If cp(10) = 撤銷原處分 Then
      '   Label17.Visible = True
      '   Text13.Visible = True
      '   Label16.Visible = True
      '   Label17.Visible = True
      'Else
      '   Label17.Visible = False
      '   Text13.Visible = False
      '   Label16.Visible = False
      '   Label17.Visible = False
      'End If
      
      ' 90.06.26 modify by louis 進度備註放錯位置且不帶出來
      'Text31 = cp(64)
   End If
End Sub

Private Function ChgType(i As Integer, Optional SstrKind As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 8
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(pA(1), Format(SstrKind), strTempName, False) Then
         If ClsPDGetCaseProperty(pa(1), Format(SstrKind), strTempName, False) Then
            Label3(3) = strTempName
            ChgType = True
         Else
            Label3(3) = ""
         End If
   End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
   Dim ret As Long
   'If prevWndProc <> 0 Then
   '   ret = SetWindowLong(Text8.hwnd, GWL_WNDPROC, prevWndProc)
   '   prevWndProc = 0
   'End If
   PUB_SendMailCache 'Added by Morgan 2021/6/11
   Set frm06010608_3 = Nothing
End Sub

Private Sub Text26_GotFocus()
   InverseTextBox Text26
End Sub

Private Sub Text26_Validate(Cancel As Boolean)
   If Text26 = "" Then
      MsgBox "來函性質為專利權消滅時，不可空白 !", vbCritical
      Cancel = True
   Else
      If Not ChkDate(Text26) Then
         Cancel = True
      End If
   End If
End Sub

'Private Sub Text27_GotFocus(Index As Integer)
'   InverseTextBox Text27(Index)
'End Sub
'
'Private Sub Text27_KeyPress(Index As Integer, KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'   'Modify by Morgan 2007/7/24 是否向客戶請款改輸N
'   If Index = 1 Then
'      If KeyAscii <> Asc("N") And KeyAscii <> 8 Then
'         KeyAscii = 0
'         Beep
'      End If
'   Else
'   'End 2007/7/24
'      If KeyAscii <> 89 And KeyAscii <> 8 Then
'         KeyAscii = 0
'         Beep
'      End If
'   End If
'   If Index = 0 Then
'      If KeyAscii <> 89 Then
'         MsgBox "來函性質為專利權消滅時，必須為 Y !", vbCritical
'         KeyAscii = 89
'      End If
'   End If
'End Sub

Private Sub Text31_GotFocus()
   InverseTextBox Text31
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

If Me.Text26.Enabled = True Then
   Cancel = False
   Text26_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

'Added by Lydia 2019/08/26自動將下載的.PDF檔,上傳到卷宗區
Private Function AutoSavePdf_FCP(ByVal iCP01 As String, ByVal iCP02 As String, ByVal iCP03 As String, ByVal iCP04 As String, ByVal iCp09 As String, ByVal iCP10 As String) As Boolean
Dim fs, f
Dim fType As String, fPath As String
Dim strErr As String
Dim strFileName As String, stReName As String
Dim strB01 As String, intB As Integer
Dim rsB1 As New ADODB.Recordset

      If iCP01 <> "FCP" Or iCP02 = "" Or iCp09 = "" Then
          Exit Function
      End If
      
      'fType = "DATA"
      'Modified by Lydia 2024/07/22 改用變數
      'fPath = "\\TYPING2\Pat_cancel"
      fPath = "\\" & strTyping2Path & "\Pat_cancel"

On Error GoTo JumpExit
      '可以多筆PDF上傳
JumpToRe:
      strFileName = Dir(fPath & "\" & iCP01 & "*" & Val(iCP02) & IIf(iCP03 <> "0", "*" & iCP03, "") & IIf(iCP04 <> "00", "*" & iCP04, "") & "*.pdf")
      If strFileName = "" Then
           If MsgBox(fPath & "資料夾底下無 " & iCP01 & iCP02 & ".pdf，" & vbCrLf & "請檢查檔案後，確定是否重新查詢？", vbCritical + vbYesNo + vbDefaultButton1) = vbYes Then
                GoTo JumpToRe
           End If
      Else
           Do While strFileName <> ""
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set f = fs.GetFile(fPath & "\" & strFileName)
                '檔案大小為 0 KB 有誤
                If f.Size = 0 Then
                     strErr = strErr & vbCrLf & fPath & "\" & strFileName & "，" & MsgText(9221)
                     GoTo JumpNextDir
                End If
                If PUB_ChkFileOpening(fPath & "\" & strFileName) = True Then
                    strErr = strErr & vbCrLf & fPath & "\" & strFileName & "，檔案正在使用中，請關閉或關閉檔案後間隔1分鐘，方能上傳到卷宗區。"
                    GoTo JumpNextDir
                End If
                 
                '檢查檔名規則
                If PUB_ChkEmpFlowFNMRule(iCP01 & "-" & iCP02 & "-" & iCP03 & "-" & iCP04, strFileName, "Y", iCP10, , , False, False, strErr) = False Then
                     GoTo JumpNextDir
                End If
                '更名
                If PUB_GetEmpFlowReNameFile(iCP01, iCP02, iCP03, iCP04, iCP10, strFileName, stReName, True, 1, False, strErr, , fType) = False Then
                     GoTo JumpNextDir
                End If
                'stReName = 案號.案件性質.PDF
                
                '檢查卷宗區檔案是否存在
                strB01 = "SELECT cpp01,cpp02 FROM casepaperpdf " & _
                              "WHERE cpp01 ='" & iCp09 & "' and instr(upper(cpp02),'" & UCase(stReName) & "') > 0 and instr(upper(cpp02),'PDF.DEL') = 0 "
                intB = 1
                Set rsB1 = ClsLawReadRstMsg(intB, strB01)
                If intB = 1 Then
                     strErr = strErr & vbCrLf & fPath & "\" & rsB1.Fields("cpp02") & "，卷宗區檔案已存在！"
                     GoTo JumpNextDir
                End If

                '上傳到卷宗區
                If SaveAttFile_PDF(iCp09, fPath & "\" & strFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = False Then
                     strErr = strErr & vbCrLf & fPath & "\" & strFileName & "，存檔失敗！" & vbCrLf & Err.Description
                     GoTo JumpNextDir
                End If
                fs.DeleteFile fPath & "\" & strFileName, True '刪檔

JumpNextDir:
                strFileName = Dir()
           Loop
           AutoSavePdf_FCP = True
      End If

JumpExit:

Set rsB1 = Nothing

If Err.Number <> 0 Then strErr = strErr & vbCrLf & Err.Description

If strErr <> "" Then
   MsgBox "FCP發文自動將下載的PDF檔，上傳到卷宗區作業失敗：" & strErr, vbCritical
End If
End Function
