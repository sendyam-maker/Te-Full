VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm110102_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "取消收文"
   ClientHeight    =   5424
   ClientLeft      =   6612
   ClientTop       =   5388
   ClientWidth     =   8484
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5424
   ScaleWidth      =   8484
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   3
      Left            =   5250
      MaxLength       =   1
      TabIndex        =   55
      Top             =   2010
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   8
      Left            =   6960
      TabIndex        =   10
      Top             =   5310
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox cboReason 
      Height          =   300
      ItemData        =   "frm110102_2.frx":0000
      Left            =   1440
      List            =   "frm110102_2.frx":000A
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   2310
      Width           =   6975
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   1
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2640
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   0
      Left            =   1425
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2010
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   6300
      TabIndex        =   12
      Top             =   60
      Width           =   1284
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   5490
      TabIndex        =   11
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   7620
      TabIndex        =   13
      Top             =   60
      Width           =   756
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   4
      Left            =   1488
      MaxLength       =   1
      TabIndex        =   6
      Top             =   4770
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   5
      Left            =   6048
      MaxLength       =   1
      TabIndex        =   7
      Top             =   4770
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   2
      Left            =   5100
      MaxLength       =   1
      TabIndex        =   3
      Top             =   2640
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   6
      Left            =   1488
      MaxLength       =   7
      TabIndex        =   8
      Top             =   5070
      Width           =   1212
   End
   Begin VB.TextBox txtCaseField 
      Height          =   264
      Index           =   7
      Left            =   4290
      MaxLength       =   7
      TabIndex        =   9
      Top             =   5070
      Width           =   1212
   End
   Begin VB.Label lblCP20 
      Caption         =   "取消收文之進度是否向客戶收款：　　　(N：不收)"
      Height          =   240
      Left            =   2550
      TabIndex        =   54
      Top             =   2022
      Width           =   4065
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   732
      Left            =   1044
      TabIndex        =   5
      Top             =   3990
      Width           =   7332
      VariousPropertyBits=   -1467987941
      ScrollBars      =   2
      Size            =   "12885;1244"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   276
      Left            =   1080
      TabIndex        =   14
      Top             =   720
      Width           =   7332
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14420;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtMemo 
      Height          =   705
      Left            =   1050
      TabIndex        =   4
      Top             =   2940
      Width           =   7305
      VariousPropertyBits=   -1467987941
      ScrollBars      =   2
      Size            =   "12885;1244"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件備註欄：不可銷卷案請加註 ""不銷卷"" 字樣！  與他案合併計算結餘請註明""與某案號合併計算結餘""！"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   115
      Left            =   120
      TabIndex        =   53
      Top             =   3720
      Width           =   8220
   End
   Begin VB.Label Label8 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   2
      Left            =   4290
      TabIndex        =   52
      Top             =   1524
      Width           =   900
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "審定號數："
      Height          =   180
      Index           =   5
      Left            =   4290
      TabIndex        =   51
      Top             =   1770
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label21 
      Caption         =   "人民幣費用："
      Height          =   180
      Index           =   6
      Left            =   5760
      TabIndex        =   50
      Top             =   5370
      Visible         =   0   'False
      Width           =   1160
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   9
      Left            =   5205
      TabIndex        =   49
      Top             =   1770
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   8
      Left            =   1092
      TabIndex        =   47
      Top             =   1776
      Width           =   492
   End
   Begin VB.Label lblNation 
      Height          =   180
      Left            =   1692
      TabIndex        =   46
      Top             =   1776
      Width           =   2412
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   1
      Left            =   5268
      TabIndex        =   44
      Top             =   504
      Width           =   3012
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "是否閉卷：            （Y：閉卷）"
      Height          =   180
      Left            =   60
      TabIndex        =   43
      Top             =   2640
      Width           =   2460
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "取消收文原因："
      Height          =   180
      Left            =   105
      TabIndex        =   42
      Top             =   2340
      Width           =   1260
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "取消收文日期："
      Height          =   180
      Left            =   105
      TabIndex        =   41
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "進度備註："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   33
      Top             =   3990
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "是否列印指示信：            （N：不印）"
      Height          =   180
      Left            =   45
      TabIndex        =   32
      Top             =   4770
      Width           =   3000
   End
   Begin VB.Label Label4 
      Caption         =   "是否修改指示信內容：            （Y：Word）"
      Height          =   180
      Left            =   4245
      TabIndex        =   31
      Top             =   4770
      Width           =   3495
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   0
      Left            =   1068
      TabIndex        =   29
      Top             =   504
      Width           =   3012
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   2
      Left            =   948
      TabIndex        =   28
      Top             =   1020
      Width           =   852
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   3
      Left            =   5256
      TabIndex        =   27
      Top             =   1020
      Width           =   372
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   4
      Left            =   1056
      TabIndex        =   26
      Top             =   1272
      Width           =   2892
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   5
      Left            =   5256
      TabIndex        =   25
      Top             =   1272
      Width           =   2892
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   6
      Left            =   972
      TabIndex        =   24
      Top             =   1524
      Width           =   852
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   7
      Left            =   5205
      TabIndex        =   23
      Top             =   1530
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "後續准駁簡單報告：            （Y：核准以及C類來函簡單報告）"
      Height          =   180
      Left            =   3420
      TabIndex        =   22
      Top             =   2640
      Width           =   5004
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "下次本所期限："
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   21
      Top             =   5130
      Width           =   1260
   End
   Begin VB.Label Label21 
      Caption         =   "下次法定期限："
      Height          =   180
      Index           =   4
      Left            =   2940
      TabIndex        =   20
      Top             =   5130
      Width           =   1305
   End
   Begin MSForms.Label lblSales 
      Height          =   180
      Left            =   6090
      TabIndex        =   19
      Top             =   1530
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "3625;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNextProgress 
      Height          =   180
      Left            =   5688
      TabIndex        =   18
      Top             =   1020
      Width           =   2652
   End
   Begin MSForms.Label lblPetitionName 
      Height          =   180
      Left            =   1908
      TabIndex        =   17
      Top             =   1020
      Width           =   2292
      VariousPropertyBits=   27
      Size            =   "3625;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPromoter 
      Height          =   180
      Left            =   1932
      TabIndex        =   16
      Top             =   1524
      Width           =   2292
      VariousPropertyBits=   27
      Size            =   "3625;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblChildCase 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "有子案或相關卷號"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   6690
      TabIndex        =   15
      Top             =   2040
      Width           =   1440
   End
   Begin VB.Label Label5 
      Caption         =   "申請國家："
      Height          =   180
      Left            =   132
      TabIndex        =   48
      Top             =   1776
      Width           =   972
   End
   Begin VB.Label Label11 
      Caption         =   "收文號："
      Height          =   180
      Index           =   2
      Left            =   4308
      TabIndex        =   45
      Top             =   504
      Width           =   972
   End
   Begin VB.Label Label21 
      Caption         =   "法定期限："
      Height          =   180
      Index           =   2
      Left            =   4296
      TabIndex        =   40
      Top             =   1272
      Width           =   972
   End
   Begin VB.Label Label21 
      Caption         =   "本所期限："
      Height          =   180
      Index           =   1
      Left            =   96
      TabIndex        =   39
      Top             =   1272
      Width           =   972
   End
   Begin VB.Label Label9 
      Caption         =   "案件性質："
      Height          =   180
      Left            =   4308
      TabIndex        =   38
      Top             =   1020
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   108
      TabIndex        =   37
      Top             =   504
      Width           =   972
   End
   Begin VB.Label Label15 
      Caption         =   "申請人："
      Height          =   180
      Index           =   0
      Left            =   108
      TabIndex        =   36
      Top             =   1020
      Width           =   732
   End
   Begin VB.Label Label14 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   0
      Left            =   132
      TabIndex        =   35
      Top             =   1524
      Width           =   732
   End
   Begin VB.Label Label21 
      Caption         =   "案件備註："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   34
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   756
      Width           =   972
   End
End
Attribute VB_Name = "frm110102_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/5/7 改成Form2.0(cboCaseName,lblPetitionName,lblPromoter,lblSales,txtMemo,txtCP64)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'2010/8/3 日期欄已修改 by sonia
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
'intWhere 國內,國外_CF,國外_FC
Dim intCaseKind As Integer, intWhere As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String, SCp(1 To 79) As String
'Wcp09存於有代理人之總收文號,以列印定稿及小信封
Dim Wcp09 As String
'intLeaveKind離開時，是0:結束  1:回上一畫面 2:確定
Dim intLeaveKind As Integer
'edit by nickc 2007/02/05 不用 dll 了
'Dim obj011 As New prjTaieDll011.cls011
'看卷是否已經閉卷
Dim BolFileClose As Boolean
'看卷是否確定閉卷
Dim BolFileCloseOk As Boolean
Dim strSql As String
'下次本所期限，下次法定期限
Dim Nextdate1 As String, Nextdate2 As String
Dim m_stReason As String 'Add by Morgan 2006/7/4 取消收文原因
Dim strMCaseCP09 As String 'Add by Morgan 2009/12/28 新多國主案收文號
Dim bol416NPCtrl As Boolean 'Add by Morgan 2010/3/1 實審取消收文是否恢復下一程序期限
Dim m_bolFMP As Boolean 'Add by Lydia 2015/02/10 判斷FMP案
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/09 是否為寰華案
Dim m_boleOrderLetter As Boolean 'Added by Morgan 2015/11/3 指示信電子化
'Added by Lydia 2018/03/16 是否已觸發 Form Active 事件
Dim bolActive As Boolean
Dim m_strAF01 As String 'Added by Morgan 2018/8/22
Dim m_bolPA141 As Boolean 'Added by Lydia 2022/02/10 是否更新FCP是否核對已准專利: N
Dim m_PA177 As String 'Added by Lydia 2023/07/28 FCP專利連結通知

Private Sub cboReason_Click()
   If Me.cboReason.Text <> "" Then
      m_stReason = Left(Me.cboReason.Text, 2)
   End If
End Sub

'add by nickc 2008/05/30 再次檢查原因
Private Sub cboReason_LostFocus()
Dim ii As Integer
Dim blnInput As Boolean

    'Add By Cheng 2003/04/16
    If Me.cboReason.Text <> "" Then
        blnInput = False
        For ii = 0 To Me.cboReason.ListCount - 1
            If Left(Me.cboReason.Text, 2) = Left(Me.cboReason.List(ii), 2) Then
                Me.cboReason.ListIndex = ii
                blnInput = True
            End If
        Next ii
        If blnInput = False Then
            Me.cboReason.Text = ""
        End If
    End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim stLetter As String 'Add by Morgan 2004/9/27
   Dim strTmp As String, bolChk As Boolean, i As Integer
   'Add By Cheng 2002/07/31
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim oText As TextBox 'Add by Morgan 2007/10/3
   Dim strAccNo As String 'Add By Sindy 2012/10/18
   Dim strSubject As String 'Added by Morgan 2018/8/20
   Dim strProcNp06 As String 'Added by Lydia 2021/10/29 詢問USER是否恢復下一程序原期限管控
   Dim strCP12 As String, strCP13 As String 'Added by Morgan 2023/5/16
   
    Select Case Index
       Case 0
          If txtCaseField(0) = "" Then
             MsgBox "取消收文日期不可空白 !", vbCritical
             txtCaseField(0).SetFocus
             txtCaseField_GotFocus (0)
             Exit Sub
          End If
          'add by nickc 2008/05/30 再次檢查原因
          cboReason_LostFocus
          
          If cboReason.Text = "" Then
             MsgBox "取消收文原因不可空白 !", vbCritical
             cboReason.SetFocus
             Exit Sub
          End If
          
          'Add by Morgan 2007/10/30
          For Each oText In txtCaseField
            If oText.Enabled = True Then
               If CheckKeyIn(oText.Index) = -1 Then
                  txtCaseField_GotFocus oText.Index
                  Exit Sub
               End If
            End If
          Next
          'end 2007/10/30
          
          BolFileCloseOk = False
          If txtCaseField(1) = "Y" Then
            If cp(31) = "Y" Then
               If MsgBox("是否確定閉卷？", vbInformation + vbDefaultButton1 + vbOKCancel) = vbOK Then
                  BolFileCloseOk = True
               Else
                  Exit Sub
               End If
            Else
               If CheckCloseFile Then
                  BolFileCloseOk = True
               Else
                  Exit Sub
               End If
            End If
          End If
         
         'Add by Morgan 2009/12/28
         strMCaseCP09 = ""
         If BolFileCloseOk Then
            '多國主案閉卷彈新主案國家及案號之訊息(新案未發文才要--禧佩)
            '新主案順序 美->日->德->英->韓->收文順序
            If cp(1) = "CFP" And cp(21) = "" And cp(31) = "Y" And cp(27) = "" Then
               strSql = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo" & _
                  ",na03,cp09 from caserelation,patent,caseprogress,nation where cr01='" & cp(1) & "'" & _
                  " and cr02='" & cp(2) & "' and cr03='" & cp(3) & "' and cr04='" & cp(4) & "'" & _
                  " and pa01(+)=cr05 and pa02(+)=cr06 and pa03(+)=cr07 and pa04(+)=cr08 and pa57 is null" & _
                  " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp31='Y' and cp21='Y'" & _
                  " and na01(+)=pa09" & _
                  " order by decode(pa09,'101','1','011','2','231','3','201','4','012','5',CP09) ASC"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strMCaseCP09 = RsTemp("cp09")
                  MsgBox "本案為多國主案,閉卷後將改設定【 " & RsTemp("CNo") & "(" & RsTemp("na03") & ") 】為主案!!"
               End If
            End If
         End If
         
         'cancel by sonia 2025/1/3 取消查名費請款單
'         'Add By Sindy 2012/9/4
'          If cp(1) = "T" And cp(10) = "101" And Not IsNull(cp(139)) Then
'            If field(10) = "000" And GetPrjNationNumber(cp(139)) = "020" Then
'               If Val(txtCaseField(8)) = 0 Then
'                  MsgBox "請輸入人民幣費用！", vbCritical
'                  txtCaseField(8).SetFocus
'                  Call txtCaseField_GotFocus(8)
'                  Exit Sub
'               End If
'            End If
'          End If
'          '2012/9/4 End
                     
         'add by nickc 2005/05/05 加入沒發文且為新案的不問
         If Trim(CheckStr(cp(27))) = "" And Trim(CheckStr(cp(31))) = "Y" Then
         Else   '2012/4/30 add by sonia CFT-013909發現都沒問
            'add by nickc 2005/04/22
            '2011/11/8 modify by sonia TF子案不可結餘故加傳本所案號
            'Pub_EndModCashMsg lblCaseField(8).Caption
           ' Pub_EndModCashMsg lblCaseField(8).Caption, cp(1), cp(2), cp(3), cp(4)
            'Add by Lydia 2015/02/10 P案非台灣案自動上結餘日 ->有申請日及申請案
             If m_bolFMP = False And field(9) <> 台灣國家代號 And field(1) = "P" And Val(field(10)) > 0 And Len(Trim(field(11))) > 0 Then
                 bolEndModCash = True  '自動上結餘日
             Else
                 'Modified by Lydia 2015/02/12 P案非台灣案一律上結餘日,其餘P案皆不詢問
                 'Pub_EndModCashMsg lblCaseField(8).Caption, cp(1), cp(2), cp(3), cp(4)
                 If field(1) <> "P" Then Pub_EndModCashMsg lblCaseField(8).Caption, cp(1), cp(2), cp(3), cp(4)
             End If
         End If
         
         'Added by Morgan 2015/11/3 指示信電子化
         'P非臺灣案指示信都要彈修改畫面來確認送判的內容
         m_boleOrderLetter = False
         'Modified by Morgan 2015/12/15 外專程序除外
         'Modified by Morgan 2018/8/16 +CFP電子化
         If (field(1) = "P" Or (field(1) = "CFP" And strSrvDate(1) >= CFP指示信電子化啟用日)) And field(9) <> "000" And txtCaseField(4) = "" And Left(Pub_StrUserSt03, 1) <> "F" Then
            m_boleOrderLetter = True
         End If
         'end 2015/11/3

          Screen.MousePointer = vbHourglass
          For i = 0 To 7
            If i <> 3 Then 'Added by Morgan 2021/5/7
               If txtCaseField(i).Enabled Then
                  If CheckKeyIn(i) <> 1 Then
                     txtCaseField(i).SetFocus
                     txtCaseField_GotFocus (i)
                     Exit For
                  End If
               End If
            End If
          Next
          
         'Add by Morgan 2010/3/1 已有申請日的實審取消收文時要恢復下一程序期限
         bol416NPCtrl = False
         If (cp(1) = "P" Or cp(1) = "CFP" Or cp(1) = "FCP") And cp(10) = "416" And field(10) <> "" Then
            If Val(cp(6)) > 0 And Val(cp(7)) > 0 Then
               '檢查國家檔是否有實審設定
               strExc(0) = "select na01 from nation where na01='" & field(9) & "' " & _
                  " and decode('" & field(8) & "','1',na26*na27,'2',na28*na29,'3',na30*na31)>0"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  bol416NPCtrl = True
                  MsgBox "實審取消收文後系統將恢復下一程序期限！"
               End If
            End If
         End If
         'Added by Lydia 2021/11/02 取消收文後下一程序原期限管控：詢問USER是否恢復下一程序原期限管控
         strProcNp06 = ""
         If bol416NPCtrl = False And cp(43) <> "" Then  '排除
            '比照案件進度維護
            strExc(0) = "SELECT NP22 FROM NEXTPROGRESS WHERE NP01='" & cp(43) & "' AND NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP07='" & cp(10) & "' AND NP06='Y'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                If MsgBox("此筆下一程序期限已上續辦，是否要恢復期限管制？", vbInformation + vbYesNo, "下一程序原期限管控") = vbYes Then
                    strProcNp06 = "Y"
                Else
                    strProcNp06 = "N"
                End If
            Else
                If Val(cp(7)) > 0 And Left(cp(9), 1) < "C" Then
                    MsgBox "此收文程序有期限, 但無相關之下一程序期限, 是否需補期限, 請自行決定！"
                End If
            End If
         End If
         'end 2021/11/02
         
         'Added by Lydia 2022/02/10 取消收文核對已准專利的時候，彈訊息確認是否回寫到「專利案件基本資料維護中>FCP是否核對已准專利: N」
         m_bolPA141 = False
         If field(1) = "FCP" And cp(10) = "926" Then
            If "" & field(141) <> "N" Then
                'Modified by Lydia 2022/02/11 改成直接提示
                'If MsgBox("是否設定專利案件基本檔之FCP是否核對已准專利=N ？", vbYesNo + vbDefaultButton2, "更新專利案件基本檔") = vbYes Then
                '   m_bolPA141 = True
                'End If
                MsgBox "將回存到專利基本檔FCP是否核對已准專利: N "
                m_bolPA141 = True
                'end 2022/02/11
            End If
         End If
         'end 2022/02/10
         
         CheckFCPDualCase 'Added by Morgan 2015/5/18
         
          '911106 nick transation
          On Error GoTo CheckingErr
          cnnConnection.BeginTrans
          
          'Added by Morgan 2023/5/16 新案取消收文前先抓,否則法律案會抓不到資料
          strCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
          strCP12 = GetSalesArea(strCP13)
          'end 2023/5/16
          
          '**************************************************
          ' nick 900801 改
          '收文號
          'frm110102_1.grdDataList.TextMatrix(frm110102_1.grdDataList.Row, 0)
          'frm110102_1.grdDataList.TextMatrix(frm110102_1.grdDataList.Row, 10)
          '本所案號
          'cp(1), cp(2), cp(3), cp(4)
          'UPDATE 是否續辦為 N 和解除期限日期和解除期限原因
          'Modify By Cheng 2002/01/23
          '是不算案件數(CP26)設為"N"
'          strSQL = "UPDATE CASEPROGRESS SET CP57=" & ChangeTStringToWString(txtCaseField(0)) & ",CP58='" & m_stReason & "' WHERE CP09='" & frm110102_1.grdDataList.TextMatrix(frm110102_1.grdDataList.Row, 0) & "' "
          'Added by Lydia 2023/04/12
          strExc(1) = ""
          If txtCaseField(3).Visible = True And txtCaseField(3) <> "" Then
             strExc(1) = strExc(1) & ", CP20=" & CNULL(txtCaseField(3))
          End If
          'end 2023/04/12
          'Modified by Lydia 2023/04/12 +strExc(1)
          strSql = "UPDATE CASEPROGRESS SET CP57=" & ChangeTStringToWString(txtCaseField(0)) & ",CP58='" & m_stReason & "' " & ",CP26='N' " & strExc(1) & _
                  " WHERE CP09='" & frm110102_1.grdDataList.TextMatrix(frm110102_1.grdDataList.row, 0) & "' "
          'Added by Lydia 2023/04/12
          If strExc(1) <> "" Then
             Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
          End If
          'end 2023/04/12
          cnnConnection.Execute strSql
          If BolFileCloseOk Then
               'Modify By Sindy 2014/4/28 txtMemo==>ChgSQL(txtMemo)
               Select Case Val(CheckSys(cp(1)))
               Case 1
                    strSql = "UPDATE PATENT SET PA57='" & txtCaseField(1) & "',PA58=" & ChangeTStringToWString(txtCaseField(0)) & ",PA59='" & m_stReason & "',PA89='" & txtCaseField(2) & "',PA91='" & ChgSQL(txtMemo) & "' WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "' "
                    cnnConnection.Execute strSql
               Case 2
                    strSql = "UPDATE TRADEMARK SET TM29='" & txtCaseField(1) & "',TM30=" & ChangeTStringToWString(txtCaseField(0)) & ",TM31='" & m_stReason & "',TM58='" & ChgSQL(txtMemo) & "' WHERE TM01='" & cp(1) & "' AND TM02='" & cp(2) & "' AND TM03='" & cp(3) & "' AND TM04='" & cp(4) & "' "
                    cnnConnection.Execute strSql
               Case 3
                    strSql = "UPDATE LAWCASE SET LC08='" & txtCaseField(1) & "',LC09=" & ChangeTStringToWString(txtCaseField(0)) & ",LC10='" & m_stReason & "',LC27='" & ChgSQL(txtMemo) & "' WHERE LC01='" & cp(1) & "' AND LC02='" & cp(2) & "' AND LC03='" & cp(3) & "' AND LC04='" & cp(4) & "' "
                    cnnConnection.Execute strSql
               Case 4
                    strSql = "UPDATE HIRECASE SET HC09='" & txtCaseField(1) & "',HC10=" & ChangeTStringToWString(txtCaseField(0)) & ",HC11='" & m_stReason & "',HC12='" & ChgSQL(txtMemo) & "' WHERE HC01='" & cp(1) & "' AND HC02='" & cp(2) & "' AND HC03='" & cp(3) & "' AND HC04='" & cp(4) & "' "
                    cnnConnection.Execute strSql
               Case 5, 6, 7, 8
                    strSql = "UPDATE SERVICEPRACTICE SET SP15='" & txtCaseField(1) & "',SP16=" & ChangeTStringToWString(txtCaseField(0)) & ",SP17='" & m_stReason & "',SP18='" & ChgSQL(txtMemo) & "' WHERE SP01='" & cp(1) & "' AND SP02='" & cp(2) & "' AND SP03='" & cp(3) & "' AND SP04='" & cp(4) & "' "
                    cnnConnection.Execute strSql
               Case Else
               End Select
          Else
               'Modify By Sindy 2014/4/28 txtMemo==>ChgSQL(txtMemo)
               Select Case Val(CheckSys(cp(1)))
               Case 1
                    strSql = "UPDATE PATENT SET PA89='" & txtCaseField(2) & "',PA91='" & ChgSQL(txtMemo) & "' WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "' "
                    cnnConnection.Execute strSql
               Case 2
                    strSql = "UPDATE TRADEMARK SET TM58='" & ChgSQL(txtMemo) & "' WHERE TM01='" & cp(1) & "' AND TM02='" & cp(2) & "' AND TM03='" & cp(3) & "' AND TM04='" & cp(4) & "' "
                    cnnConnection.Execute strSql
               Case 3
                    strSql = "UPDATE LAWCASE SET LC27='" & ChgSQL(txtMemo) & "' WHERE LC01='" & cp(1) & "' AND LC02='" & cp(2) & "' AND LC03='" & cp(3) & "' AND LC04='" & cp(4) & "' "
                    cnnConnection.Execute strSql
               Case 4
                    strSql = "UPDATE HIRECASE SET HC12='" & ChgSQL(txtMemo) & "' WHERE HC01='" & cp(1) & "' AND HC02='" & cp(2) & "' AND HC03='" & cp(3) & "' AND HC04='" & cp(4) & "' "
                    cnnConnection.Execute strSql
               Case 5, 6, 7, 8
                    strSql = "UPDATE SERVICEPRACTICE SET SP18='" & ChgSQL(txtMemo) & "' WHERE SP01='" & cp(1) & "' AND SP02='" & cp(2) & "' AND SP03='" & cp(3) & "' AND SP04='" & cp(4) & "' "
                    cnnConnection.Execute strSql
               Case Else
               End Select
          End If
          bolLeave = True
          'If BolFileClose = False Then
             'If BolFileCloseOk = True Then
                   Dim strAutoNum As String
                  'Modify By Cheng 2002/10/01
'                   If objPublicData.GetAutoNumber("B", strAutoNum, True, False) Then
                   'edit by nickc 2007/02/02 不用 dll 了
                   'If objPublicData.GetAutoNumber("B", strAutoNum, True, True) Then
                   If ClsPDGetAutoNumber("B", strAutoNum, True, True) Then
                        CheckOC
                        strSql = "select au01||(au02-1911) from autonumber where au01='B'"
                        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                        If Not adoRecordset.BOF Then adoRecordset.MoveFirst
                        If adoRecordset.BOF And adoRecordset.EOF Then MsgBox "自動編號錯誤", vbInformation: Exit Sub
                        'Modify By Sindy 2010/8/18 比對自動編號年度
                        'strAutoNum = CheckStr(adoRecordset.Fields(0).Value) & strAutoNum
                        strAutoNum = "B" + CompAutoNumberYear(CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911)) & strAutoNum
                        CheckOC
                        strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp12,cp13,cp14,cp15,cp16,cp17,cp18,cp19,cp20,cp21,cp22,cp23,cp24,cp25,cp26,cp27,cp28,cp29,cp30,cp31,cp32,cp33,cp34,cp35,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp43,cp44,cp45,cp46,cp47,cp48,cp49,cp50,cp51,cp52,cp53,cp54,cp55,cp56,cp57,cp58,cp59,cp60,cp61,cp62,cp63,cp64,cp71,cp72,cp73,cp74,cp75,cp76,cp77,cp78,cp79) values "
                        'Set SCp() = cp()
                        For i = 1 To 79
                           Select Case i
                           '文字null
'                           Case 8, 11, 14, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 44, 45, 49, 50, 51, 52, 55, 56, 58, 59, 60, 61, 62, 63
                           '92.1.25 MODIFY BY SONIA 取消收文日及原因要存
                           'Case 8, 11, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 44, 45, 49, 50, 51, 52, 55, 56, 58, 59, 60, 61, 62, 63
                           Case 8, 11, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 44, 45, 49, 50, 51, 52, 55, 56, 59, 60, 61, 62, 63
                           '92.1.25 END
                                'Added by Lydia 2016/02/25 CFP案取消收文同時閉卷，請閉卷程序加掛代理人
                                If cp(1) = "CFP" And i = 44 Then
                                   SCp(i) = CNULL(GetPrjFagentNumByCP(cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)))
                                Else
                                'end 2016/02/25
                                   SCp(i) = "null "
                                End If
                           'Add By Cheng 2002/07/26
                           Case 14 '承辦人代號
                                SCp(i) = "'" & strUserNum & "'"
                           '文字畫面上
                           Case 64
                                SCp(i) = "'" & Trim(txtCP64) & "'"
                           Case 1, 2, 3, 4
                                SCp(i) = "'" & Trim(ChgSQL(cp(i))) & "'"
                           '2012/10/2 add by sonia 自上面Case 1, 2, 3, 4抽出來
                           Case 12
                              'Modified by Morgan 2023/5/16
                              'SCp(i) = "'" & GetSalesArea(PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))) & "'"
                              SCp(i) = "'" & strCP12 & "'"
                              'end 2023/5/16
                           '2012/10/2 end
                           '2011/6/1 add by sonia 自上面Case 1, 2, 3, 4, 12抽出來
                           Case 13
                              'Modified by Morgan 2023/5/16
                              'SCp(i) = "'" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "'"
                              SCp(i) = "'" & strCP13 & "'"
                              'end 2023/5/16
                           '2011/6/1 end
                           Case 5, 27
                                SCp(i) = GetTodayDate
                           Case 9
                                SCp(i) = "'" & strAutoNum & "'"
                           '91.12.6 modify by sonia
                           'Case 26, 20, 32
                           '     SCp(i) = "'N'"
                           Case 20
                              If intWhere <> "2" Then
                                 SCp(i) = "'N'"
                                 '2013/8/13 add by sonia FMT要請款
                                 'Modified by Morgan 2023/5/16
                                 'If cp(1) = "T" And Left(GetSalesArea(PUB_GetAKindSalesNo(field(1), field(2), field(3), field(4))), 1) = "F" Then
                                 If cp(1) = "T" And Left(strCP12, 1) = "F" Then
                                    SCp(i) = "null "
                                 End If
                                 '2013/8/13 end
                              'Add by Morgan 2007/8/1 改抓CPM設定
                              ElseIf cp(1) = "FCP" Then
                                 SCp(i) = CNULL(PUB_GetCP20(cp(1), Replace(SCp(10), "'", "")))
                              'add by sonia 2013/12/24 FCT不請款
                              ElseIf cp(1) = "FCT" Then
                                 SCp(i) = "'N'"
                              '2013/12/24 END
                              Else
                                 SCp(i) = "null "
                              End If
                           Case 26, 32
                                SCp(i) = "'N'"
                           '91.12.6 end
                           Case 43
                                SCp(i) = "'" & frm110102_1.grdDataList.TextMatrix(frm110102_1.grdDataList.row, 0) & "'"
                           Case 10
                                If (IsNull(txtCaseField(1)) Or txtCaseField(1) = "") Then
                                    Select Case Val(CheckSys(cp(1)))
                                    Case 1, 5         'patent
                                       SCp(i) = "'925'"   '2006/9/15 MODIFY BY SONIA 907 -> 925
                                    Case 2, 6         'trademark
                                       SCp(i) = "'718'"   '2006/9/15 MODIFY BY SONIA 703 -> 718
                                    Case 3, 4, 7, 8   'lawcase & hirecase
                                       SCp(i) = "'992'"   '2006/9/15 MODIFY BY SONIA 991 -> 992
                                    Case Else
                                    End Select
                                Else
                                    Select Case Val(CheckSys(cp(1)))
                                    Case 1, 5         'patent
                                       SCp(i) = "'913'"
                                    Case 2, 6         'trademark
                                       SCp(i) = "'704'"
                                    Case 3, 4, 7, 8   'lawcase & hirecase
                                       SCp(i) = "'993'"
                                    Case Else
                                    End Select
                                End If
                           Case 65, 66, 67, 68, 69, 70
                                SCp(i) = ""
                           '92.1.25 ADD BY SONIA
                           Case 57
                                SCp(i) = ChangeTStringToWString(txtCaseField(0))
                           Case 58
                                SCp(i) = "'" & Trim(ChgSQL(m_stReason)) & "'"
                           '92.1.25 END
                           '數字
                           Case Else
                                SCp(i) = "null "
                           End Select
                        Next i
                        strSql = strSql & " ("
                        For i = 1 To 79
                            Select Case i
                            Case 65, 66, 67, 68, 69, 70
                            Case Else
                                 strSql = strSql & SCp(i)
                                 If i <> 79 Then
                                    strSql = strSql & ","
                                 End If
                            End Select
                        Next i
                        strSql = strSql & ") "
                        cnnConnection.Execute strSql
                                                
                        'Added by Morgan 2015/11/3 指示信電子化
                        If m_boleOrderLetter Then
                           'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
                           'strExc(2) = Pub_GetSpecMan("PS4") 'P案指示信判發人
                           strExc(2) = PUB_GetLetterJudgeNew("2", field(1), Replace(SCp(10), "'", ""), field(9), lblCaseField(3))
                           'Modified by Morgan 2018/8/20 +傳指示信主旨strSubject
                           strSubject = PUB_GetSubject(field(1), field(2), field(3), field(4), field(11), cp(45))
                           PUB_AddAppForm strAutoNum, True, strExc(2), strSubject
                           strSubject = ""
                           m_strAF01 = strAutoNum
                           'end 2018/8/20
                        End If
                        'end 2015/11/3
                        
                        'Add by Sindy 2013/04/12 更新c類的代理人及彼所案號，要在新增c類之後
                        Pub_UpdateFromMaxCP27 cp(1), cp(2), cp(3), cp(4)
                        
                        bolLeave = True
                        intLeaveKind = 2
                        Me.Hide
                   Else
                       MsgBox ("自動給號錯誤")
                       Me.Hide
                   End If
             'End If
          'End If
          '若有輸下次本所期限或下次法定期限
          If Len(txtCaseField(6)) <> 0 Or Len(txtCaseField(7)) <> 0 Then
            'frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.Row, 0)
            'frm110101_1.grdDataList.TextMatrix(frm110101_1.grdDataList.Row, 10)
            strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22,NP15) SELECT CP09,CP01,CP02,CP03,CP04,TO_NUMBER(CP10)," & IIf(Len(txtCaseField(6)) <> 0, ChangeTStringToWString(txtCaseField(6)), "Null") & "," & IIf(Len(txtCaseField(7)) <> 0, ChangeTStringToWString(txtCaseField(7)), "Null") & ",CP13,CP08,NVL(SUBSTR(CP40,1,60),NVL(SUBSTR(CP41,1,60),SUBSTR(CP42,1,60)))," & GetNextProgressNo & ",'" & ChgSQL(Trim(txtCP64)) & "' FROM CASEPROGRESS WHERE CP09='" & frm110102_1.grdDataList.TextMatrix(frm110102_1.grdDataList.row, 0) & "' "
            cnnConnection.Execute strSql
          End If
          '**************************************************************************************************
          
          'add by nickc 2005/04/22
          Pub_UpdateEndModCash cp(1), cp(2), cp(3), cp(4)
          
          'Add By Sindy 2023/12/13 檢查接洽單的Flow是否要結束
          'Modify By Sindy 2024/11/20 +, Me
          Call PUB_UpdateCRLFlowClose(cp(140), cp(9), Me)
          
          'add by nickc 2005/09/16 若取消原因為 14 時，清掉所有的 可結餘日期
          If m_stReason = "14" Then
               cnnConnection.Execute "update caseprogress set cp109=null where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp59 is null "
          End If
          
          'Add by Morgan 2005/10/25 若取消收文原因=14,是否閉卷=Y時,其他AB類未發文未取消的也要更新CP57,CP58
          'Modified by Morgan 2016/8/24 改與閉卷相同所有未發文CP都上取消收文日
          'If m_stReason = "14" And txtCaseField(1) = "Y" Then
          '  strSql = "Update CaseProgress Set CP57='" & DBDATE(txtCaseField(0)) & "',CP58='14' Where " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & " AND CP09<'C' AND CP27 IS NULL AND CP57 IS NULL"
          '  cnnConnection.Execute strSql
          'End If
          If txtCaseField(1) = "Y" Then
            strSql = "UPDATE CASEPROGRESS SET CP26='N',CP57=" & ChangeTStringToWString(txtCaseField(0)) & ",CP58='" & m_stReason & "' WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP57 IS NULL AND CP27 IS NULL "
            '排除FCP案的代辦退費(實審,再審和再審延期)
            If cp(1) = "FCP" Then
               'Modified by Morgan 2022/11/23 +排除續行母案再審的代辦退費 Ex:FCP-067213 --Winfrey
               strSql = strSql & "and cp09 not in (select a.cp09 from caseprogress a,caseprogress b where a.cp01='" & cp(1) & "' and a.cp02='" & cp(2) & "' and a.cp03='" & cp(3) & "' and a.cp04='" & cp(4) & "' and a.cp10='" & 退費 & "' and a.cp27||a.cp57 is null and b.cp09(+)=a.cp43 and b.cp10 in ('416','107','435') " & _
                        "union select a.cp09 from  caseprogress a,caseprogress b,nextprogress where a.cp01='" & cp(1) & "' and a.cp02='" & cp(2) & "' and a.cp03='" & cp(3) & "' and a.cp04='" & cp(4) & "' and a.cp10='" & 退費 & "' and a.cp27||a.cp57 is null and b.cp09(+)=a.cp43 and b.cp10='404' and np01(+)=b.cp43 and np07='107' " & _
                        "union select a.cp09 from  caseprogress a,caseprogress b,caseprogress c where a.cp01='" & cp(1) & "' and a.cp02='" & cp(2) & "' and a.cp03='" & cp(3) & "' and a.cp04='" & cp(4) & "' and a.cp10='" & 退費 & "' and a.cp27||a.cp57 is null and b.cp09(+)=a.cp43 and b.cp10='404' and c.cp09(+)=b.cp43 and c.cp10='107') "
            End If
            cnnConnection.Execute strSql, intI
         End If
         'end 2016/8/24
          
         'Add by Morgan 2009/12/28
         If strMCaseCP09 <> "" Then
             strSql = "UPDATE caseprogress SET cp21=null WHERE cp09='" & strMCaseCP09 & "' and cp21='Y'"
             cnnConnection.Execute strSql
         End If
          
'Removed by Morgan 2016/11/14 改列印銷帳銷案單時確認並列印於理由 --秀玲
'         'Add by Morgan 2009/10/15
'         '大陸案一案兩請:新型年費欲結案時,若該發明案尚未核准公告,則發E-MAIL告知智權同仁及其所屬區主管
'         If cp(10) = "605" And field(1) = "P" And field(9) = "020" And field(8) = "2" And Val(DBDATE(field(10))) >= 20091001 Then
'            strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1" & _
'               " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & field(1) & "' and cm02='" & field(2) & "' and cm03='" & field(3) & "' and cm04='" & field(4) & "'" & _
'               " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & field(1) & "' and cm06='" & field(2) & "' and cm07='" & field(3) & "' and cm08='" & field(4) & "') X" & _
'               ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null and (pa16 is null or pa16='2')"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strExc(1) = field(1) & "-" & field(2) & IIf(field(3) & field(4) = "000", "", "-" & field(3) & "-" & field(4))
'               strExc(1) = "提醒:" & strExc(1) & "大陸案為一案兩請,新型放棄續繳年費將同時放棄發明或實用新型間擇一選擇的權利。"
'               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                  " values ('" & strUserNum & "','" & lblCaseField(7) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(1)) & "','如旨')"
'               cnnConnection.Execute strSql, intI
'
'               strExc(0) = "select a0908 from staff,acc090 where st01='" & lblCaseField(7) & "' and a0901(+)=st15 and a0908<>'" & lblCaseField(7) & "'"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                     " values ('" & strUserNum & "','" & RsTemp(0) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(1)) & "','如旨')"
'                  cnnConnection.Execute strSql, intI
'               End If
'            End If
'         End If
'         'end 2009/10/15
'end 2016/11/14
         
         'Add by Morgan 2010/3/1 恢復下一程序實審期限
         If bol416NPCtrl Then
            intI = 0
            If cp(43) <> "" Then
               strSql = "update nextprogress set np06=null,np08=" & DBDATE(cp(6)) & ",np09=" & DBDATE(cp(7)) & _
                  " where np01='" & cp(43) & "' and np02='" & cp(1) & "' and np03='" & cp(2) & "'" & _
                  " and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np06='Y' and np07='416'"
               cnnConnection.Execute strSql, intI
            End If
            If intI = 0 Then
               'strExc(1) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) 'Removed by Morgan 2023/5/16
               'Modified by Morgan 2023/5/16 strExc(1)->strCP13
               strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
                  "NP07,NP08,NP09,NP10,NP22) VALUES (" & SCp(9) & ",'" & cp(1) & "','" & cp(2) & "'" & _
                  ",'" & cp(3) & "','" & cp(4) & "','416'," & DBDATE(cp(6)) & "," & DBDATE(cp(7)) & ",'" & strCP13 & "'" & _
                  ",GETNP22)"
               'end 2023/5/16
               cnnConnection.Execute strSql, intI
            End If
         End If
         
         'Added by Lydia 2021/11/02 取消收文後下一程序原期限管控：詢問USER是否恢復下一程序原期限管控
         If strProcNp06 = "Y" Then
             '選是，下一程序記錄原續辦(Y)清空，恢復原期限管控
             strSql = "UPDATE NEXTPROGRESS SET NP06=NULL WHERE NP01='" & cp(43) & "' AND NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP07='" & cp(10) & "' AND NP06='Y'"
             cnnConnection.Execute strSql, intI
         ElseIf strProcNp06 = "N" Then
             '選否，下一程序記錄改為不續辦(N)，原期限不管控，
             strSql = "UPDATE NEXTPROGRESS SET NP06='N' WHERE NP01='" & cp(43) & "' AND NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP07='" & cp(10) & "' AND NP06='Y'"
             cnnConnection.Execute strSql, intI
         End If
         'end 2021/11/02
         
'cancel by sonia 2025/1/3 取消查名費請款單
'         'Add By Sindy 2012/9/4 大->台有代理人案件,查名近似銷案不辦,但仍向代理人請款,因此增加出定稿及請款單,以利方便通知代理人
'          If cp(1) = "T" And cp(10) = "101" And Not IsNull(cp(139)) Then
'            If field(10) = "000" And GetPrjNationNumber(cp(139)) = "020" Then
'               'Modify By Sindy 2019/11/28 程序說作廢不出定稿
''               '定稿
''               EndLetter "17", cp(9), "01", strUserNum
''               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                        "VALUES ('" & "17" & "','" & cp(9) & "','" & "01" & "','" & strUserNum & "'," & _
''                        "'費用','" & Format(txtCaseField(8), "##,##0") & "')"
''               cnnConnection.Execute strSql
''               NowPrint cp(9), "17", "01", False, strUserNum, 0
'               '2019/11/28 END
'               '請款單
'               strAccNo = AddAccData("001")  'cancel by sonia 2025/1/3 取消查名費請款單
'            End If
'          End If
'          '2012/9/4 End
'end 2025/1/3
          
          'Added by Lydia 2022/02/10 取消收文核對已准專利的時候，彈訊息確認是否回寫到「專利案件基本資料維護中>FCP是否核對已准專利: N」
          If m_bolPA141 = True Then
              'Modified by Lydia 2022/02/11 +備註PA91
              strSql = "Update Patent Set PA141='N',PA91='" & ChangeWStringToTDateString(strSrvDate(1)) & "  取消收文核對已准;" & "'||PA91 where pa01='" & field(1) & "' and pa02='" & field(2) & "' and pa03='" & field(3) & "' and pa04='" & field(4) & "' "
              Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
              cnnConnection.Execute strSql, intI
          End If
          'end 2022/02/10
          
         'Added by Lydia 2015/08/13 取消收文主張優先權時要刪除優先權資料
         strSql = "": strExc(1) = ""
         Select Case cp(10)
            Case "106" '主張國際優先權: 刪除優先權國家<>該案申請國家的資料
                 If cp(1) = "P" Or cp(1) = "CFP" Or cp(1) = "FCP" Then
                    '若有相同案件性質的進度,詢問是否刪除
                    If PUB_ChkCPExist(cp(), "106") Then
                       If MsgBox("本案有相同案件性質的進度檔,請問是否刪除國際優先權資料?", vbCritical + vbYesNo) = vbNo Then
                          GoTo JumpPriD
                       End If
                    End If
                    strSql = "delete from pridate where pd01||pd02||pd03||pd04||pd06||pd07 in (select pd01||pd02||pd03||pd04||pd06||pd07 from pridate " & _
                             "where pd01='" & cp(1) & "' and pd02='" & cp(2) & "' and pd03='" & cp(3) & "' and pd04='" & cp(4) & "' and pd07<>'" & lblCaseField(8).Caption & "' )"
                    strExc(1) = "已刪除國際優先權資料!"
                 End If
            Case "121" '主張國內優先權: 刪除優先權國家<>該案申請國家的資料
                 If cp(1) = "P" Or cp(1) = "CFP" Or cp(1) = "FCP" Then
                    If PUB_ChkCPExist(cp(), "121") Then
                       If MsgBox("本案有相同案件性質的進度檔,請問是否刪除國內優先權資料?", vbCritical + vbYesNo) = vbNo Then
                          GoTo JumpPriD
                       End If
                    End If
                    strSql = "delete from pridate where pd01||pd02||pd03||pd04||pd06||pd07 in (select pd01||pd02||pd03||pd04||pd06||pd07 from pridate " & _
                             "where pd01='" & cp(1) & "' and pd02='" & cp(2) & "' and pd03='" & cp(3) & "' and pd04='" & cp(4) & "' and pd07='" & lblCaseField(8).Caption & "' )"
                    strExc(1) = "已刪除國內優先權資料!"
                 End If
            Case "108" '主張優先權: 刪除所有優先權資料
                 If cp(1) = "T" Or cp(1) = "CFT" Or cp(1) = "FCT" Or cp(1) = "TF" Then
                    If PUB_ChkCPExist(cp(), "108") Then
                       If MsgBox("本案有相同案件性質的進度檔,請問是否刪除優先權資料?", vbCritical + vbYesNo) = vbNo Then
                          GoTo JumpPriD
                       End If
                    End If
                    strSql = "delete from pridate where pd01='" & cp(1) & "' and pd02='" & cp(2) & "' and pd03='" & cp(3) & "' and pd04='" & cp(4) & "' "
                    strExc(1) = "已刪除優先權資料!"
                 End If
         End Select
         If strSql <> "" Then
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
            MsgBox strExc(1)
         End If
JumpPriD:
         'end 2015/08/13
         
          'Added by Lydia 2015/10/12 商標爭議案無結果掛催審期限
          'T,FCT,CFT,TF之602異答,604評答,606廢答之解除期限或取消收文,都要掛被異議(1602)、被評定(1604)、被撤銷(1606)之c類來函的下一程序305催審
          If (field(1) = "T" Or field(1) = "TF" Or field(1) = "FCT" Or field(1) = "CFT") And (lblCaseField(3) = "602" Or lblCaseField(3) = "604" Or lblCaseField(3) = "606") Then
             '管制人員為1602,1604,1606之承辦人
             strExc(0) = " select cp09,cp14 from caseprogress where cp09='" & cp(43) & "' and cp10 in (1602,1604,1606) "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
             '法限=系統日+1年,所限=法限
                strExc(9) = PUB_GetWorkDay1(CompDate(0, 1, strSrvDate(1)), True)
                strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                         "VALUES('" & RsTemp.Fields("cp09") & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','305','" & strExc(9) & "','" & strExc(9) & "','" & RsTemp.Fields("cp14") & "'," & GetNextProgressNo & ") "
                cnnConnection.Execute strSql, intI
             End If
          End If
          'END 2015/10/12
           
          'Added by Lydia 2016/10/19 新案(101,102)銷案時,取消一案兩請關聯 (判斷未發文的新案才取消關聯)
          If field(1) <> "FCP" And cp(27) = "" And m_bolFMP = False And intCaseKind = 專利 And (lblCaseField(3) = "101" Or lblCaseField(3) = "102") Then
             If PUB_DualCaseRelationExist(field) Then
                strExc(0) = field(1): strExc(1) = field(2): strExc(2) = field(3): strExc(3) = field(4)
                strExc(4) = "": strExc(5) = "": strExc(6) = "": strExc(7) = ""
                If PUB_DeleteCaseRelation(strExc, 3) Then
                End If
             End If
          End If
          'end 2016/10/19
          
          'Added by Morgan 2017/9/12
          '專利設計案取消收文檢查歐盟設計案是否可送件
          If (field(1) = "CFP" Or field(1) = "P") And (lblCaseField(3) = "103" Or lblCaseField(3) = "105" Or lblCaseField(3) = "125") Then
            chk103in239OK cp
          End If
          'end 2017/9/12
          'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：輸入閉卷913自動收文「通知資訊變更961」,發一封Email給承辦工程師
          If txtCaseField(1) = "Y" And field(1) = "FCP" And m_PA177 = "Y" Then
             'Memo by Lydia 2025/04/02 模組內已去掉SCp(10)的單引號Replace
             If PUB_GetFCPlinkMC("6", TransDate(txtCaseField(0), 2), field, strAutoNum, SCp(10)) = True Then
             End If
          End If
          'end 2023/07/28
          'Added by Lydia 2025/09/12 TF基礎案號設定：基礎案已閉卷、【703不續辦-續展】=>基礎案狀態通知Email
          If (field(1) = "T" Or field(1) = "CFT") And BolFileCloseOk = True Then
             strSql = PUB_GetTFbaseInfo(field(1), field(2), field(3), field(4), field(15), field(10), "2", field(12))
          End If
          'end 2025/09/12
          
          cnnConnection.CommitTrans
          
          'Add By Sindy 2012/10/18
          If strAccNo <> "" Then
            MsgBox "系統自動產生的請款單編號為：" & strAccNo
          End If
          '2012/10/18 End
          
            'Add By Cheng 2003/10/21
            '若系統類別為P或FCP
            If cp(1) = "P" Or cp(1) = "FCP" Then
                '若案件性質為領證
                If Me.lblCaseField(3).Caption = "601" Then
                    '檢查本案下一程序是否有年費期限
                    StrSQLa = "Select * From Nextprogress Where " & ChgNextProgress(cp(1) & cp(2) & cp(3) & cp(4)) & " And NP07='605' And NP06 Is Null "
                    rsA.CursorLocation = adUseClient
                    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                    While Not rsA.EOF
                        If MsgBox("本案下一程序有年費期限，是否確定同時刪除年費期限???" & vbCrLf & "本所期限：" & ChangeTStringToTDateString(ChangeWStringToTString("" & rsA.Fields("NP08").Value)) & vbCrLf & "法定期限：" & ChangeTStringToTDateString(ChangeWStringToTString("" & rsA.Fields("NP09").Value)), vbExclamation + vbOKCancel + vbDefaultButton2) = vbOK Then
                            StrSQLa = "Delete From Nextprogress Where NP01='" & rsA.Fields("NP01").Value & "' And NP07=" & CNULL(rsA.Fields("NP07").Value) & " And NP22=" & CNULL(rsA.Fields("NP22").Value) & " "
                            cnnConnection.Execute StrSQLa
                        End If
                        rsA.MoveNext
                    Wend
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                End If
            End If
            'End
          
          '2015/8/10 add by sonia 專利案件閉卷時有新案翻譯尚未完稿要提醒(FCP-51551)
          'modify by sonia 2015/9/4 再加cp05>20150101否則舊案無完稿也會有訊息
          If txtCaseField(1) = "Y" And intCaseKind = 專利 Then
            strExc(0) = "select cp09,nvl(ep09,0) ep09,nvl(cp27,0) cp27 from caseprogress,engineerprogress where " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & " and cp10='201' and cp09=ep02(+) and cp05>20150101"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Val("" & RsTemp("ep09")) = 0 Then
                  MsgBox "此案新案翻譯進度尚未完稿！"
               ElseIf Val("" & RsTemp("cp27")) = 0 Then
                  MsgBox "此案新案翻譯進度尚未發文！"
               End If
            End If
          End If
          '2015/8/10 end
          
          CaseMapMailCheck 'Add by Morgan 2005/5/18
          
          If txtCaseField(4) <> "N" Then '指示信
             If txtCaseField(5) = "Y" Then
                bolChk = True
             Else
                bolChk = False
             End If
             Select Case cp(1)
               Case "CFP"
                  bolChk = True
                  Select Case m_stReason
                     Case "10" '自行處理
                        strTmp = "31"
                     Case "02" '找其他代理人
                        strTmp = "32"
                     '2005/5/4 MODIFY BY SONIA
                     'Case "05", "11" '倒閉
                     '   strTmp = "33"
                     Case "05" '遷移
                        strTmp = "33"
                     Case "11" '倒閉
                        strTmp = "34"
                     '2005/5/4 END
                     Case Else
                        '一般 30
                        strTmp = "30"
                  End Select
               Case "P"
                     Select Case lblCaseField(3)
                         Case "601" '領證
                            strTmp = "31"
                         Case "605", "606" '年費,維持費
                            strTmp = "32"
                         '92.7.7 ADD BY SONIA
                         Case "408" '面詢
                            strTmp = "34"
                         '92.7.7 ENDD
                         Case Else
                            '一般 30
                            strTmp = "30"
                     End Select
                     If m_stReason = "02" Then strTmp = "32" '找其他代理人
             End Select
             GetCP09 "14", strTmp
             
             'Modify by Morgan 2004/9/27
             'CFP取消收文加印傳真封面
             'NowPrint Wcp09, "14", strTmp, bolChk, strUserNum, 0
             If cp(1) = "CFP" Then
               'Modify by Morgan 2004/10/22
               'NowPrint Wcp09, "01", "99", False, strUserNum, , , True, stLetter
               'Removed by Morgan 2018/10/22 取消傳真封面--慧汶
               'NowPrint Wcp09, "01", "89", False, strUserNum, , , True, stLetter, , , , , , , , , m_strAF01
               'If m_strAF01 <> "" Then Sleep 1000 '等1秒以確保letterdemand不會發生dupe錯誤 Added by Morgan 2018/8/20
               'end 2018/10/22
               NowPrint Wcp09, "14", strTmp, bolChk, strUserNum, , stLetter, , , , , , , , , , , m_strAF01
                  
               'Added by Morgan 2018/8/20 CFP電子化
               If bolChk = True And m_strAF01 <> "" Then
                  frm1105_1.m_RecNo = m_strAF01
                  'Modify By Sindy 2022/5/11 流水號要足6碼
                  frm1105_1.m_PdfName = field(1) & field(2) & IIf(field(3) & field(4) = "000", "", "-" & field(3) & "-" & field(4)) & "." & Mid(SCp(10), 2, 3) & ".DATA.PDF"
                  frm1105_1.Show
               End If
               'end 2018/8/20
               
            'Added by Morgan 2015/11/3 指示信電子化
            ElseIf m_boleOrderLetter Then
               NowPrint Wcp09, "14", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , strAutoNum
               If bolChk = True Then
                  frm1105_1.m_RecNo = strAutoNum
                  'Modify By Sindy 2022/5/11 流水號要足6碼
                  frm1105_1.m_PdfName = field(1) & field(2) & IIf(field(3) & field(4) = "000", "", "-" & field(3) & "-" & field(4)) & "." & Mid(SCp(10), 2, 3) & ".DATA.PDF"
                  frm1105_1.Show
               End If
            'end 2015/11/3
            
             Else
               NowPrint Wcp09, "14", strTmp, bolChk, strUserNum, 0
             End If
             '2004/9/27 end
                
            'Removed by Morgan 2012/7/12 取消--禧佩
            ''Add By Cheng 2002/07/31
            'If cp(1) = "CFP" Then
            '   StrSQLa = "Select FA04,FA05||' '||FA63||' '||FA64||' '||FA65,FA06,FA01||FA02 From CASEPROGRESS, FAGENT WHERE SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP09='" & Wcp09 & "'"
            '   rsA.CursorLocation = adUseClient
            '   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            '   If rsA.RecordCount > 0 Then
            '     If MsgBox("代理人名稱(中)：" & rsA.Fields(0).Value & Chr(10) & Chr(13) & _
            '               "　　　　　(英)：" & rsA.Fields(1).Value & Chr(10) & Chr(13) & _
            '               "　　　　　(日)：" & rsA.Fields(2).Value & Chr(10) & Chr(13) & Chr(10) & Chr(13) & _
            '               "是否列印代理人小信封？", vbExclamation + vbYesNo) = vbYes Then
            '        '列印地址條
            '        'Modify by Morgan 2006/10/17 改Call公用函數
            '        'PrintCase "" & rsA.Fields(3).Value
            '        PUB_PrintCase "" & rsA.Fields(3).Value
            '     End If
            '   End If
            '   If rsA.State <> adStateClosed Then rsA.Close
            '   Set rsA = Nothing
            'End If
          End If
          
          'Added by Lydia 2023/06/09 當寰華案在key閉卷按確認時，請判斷是否有相關香港案及澳門案未不續辦/閉卷，若有則發mail
          If m_bolFMP2 = True And txtCaseField(1) = "Y" And lblCaseField(8) = "020" Then
             'Modified by Lydia 2023/06/28 傳入案件性質SCp(10)
             'Modified by Lydia 2025/04/02 去掉案件性質SCp(10)的單引號Replace
             Call PUB_CloseMailto013044("1", field(1), field(2), field(3), field(4), Replace(SCp(10), "'", ""))
          End If
          PUB_SendMailCache
          'end 2023/06/09
          
          Unload Me
          Screen.MousePointer = vbDefault
          Exit Sub
      Case 1, 2
         If Index = 2 Then
            intLeaveKind = 0
         Else
            intLeaveKind = 1
         End If
         Unload Me
   End Select
 '911106 nick transation
   Exit Sub
   
CheckingErr:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Sub

Private Function SaveDatabase() As Boolean
cp(57) = txtCaseField(0)
cp(58) = m_stReason
Select Case intCaseKind
             Case 專利
                        field(89) = txtCaseField(2)
                        field(57) = txtCaseField(1)
                        field(58) = txtCaseField(0)
                        field(59) = m_stReason
             Case 商標
                        field(29) = txtCaseField(1)
                        field(30) = txtCaseField(0)
                        field(31) = m_stReason
             Case 法務
                        field(8) = txtCaseField(1)
                        field(9) = txtCaseField(0)
                        field(10) = m_stReason
             Case 顧問
                        field(9) = txtCaseField(1)
                        field(10) = txtCaseField(0)
                        field(11) = m_stReason
             Case Else
                        field(15) = txtCaseField(1)
                        field(16) = txtCaseField(0)
                        field(17) = m_stReason
End Select
'edit by nickc 2007/02/05 不用 dll 了
'If obj011.SaveCancelReceivedDayData(intCaseKind, intWhere, cp(), field(), txtCaseField(1), txtCaseField(6), txtCaseField(7), txtCP64) Then
If Cls011SaveCancelReceivedDayData(intCaseKind, intWhere, cp(), field(), txtCaseField(1), txtCaseField(6), txtCaseField(7), txtCP64) Then
   SaveDatabase = True
Else
   ShowMsg MsgText(9004)
End If
End Function
Private Sub ReadAllData()
Dim i As Integer, varSaveCursor, strTemp As String, strTemp1 As String, j As Integer
'Dim intCaseKind As Integer

On Error GoTo ErrHand
Nextdate1 = "": Nextdate2 = ""
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetSystemKind(frm110102_1.txtSystem, intCaseKind, , intWhere) = False Then
If ClsPDGetSystemKind(frm110102_1.txtSystem, intCaseKind, , intWhere) = False Then
   GoTo err1
End If
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.ReadAllData(frm110102_1.grdDataList.TextMatrix(frm110102_1.grdDataList.Row, 0), cp(), field(), intCaseKind, intWhere) Then
ReDim cp(TF_CP) As String
cp(9) = frm110102_1.grdDataList.TextMatrix(frm110102_1.grdDataList.row, 0)
If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
   'Add by Lydia 2015/02/10 判斷FMP案
   'Modified by Morgan 2021/2/2
   'If Left(cp(12), 1) = "F" And cp(1) = "P" And field(9) <> "000" Then
   '   m_bolFMP = True
   'Else
   '   m_bolFMP = False
   'End If
   m_bolFMP = PUB_ChkIsFMP(field(1), field(2), field(3), field(4), field(9))
   'end 2021/2/2
   'end 2015/02/10
   
   'Added by Lydia 2023/06/09 判斷寰華案
   m_bolFMP2 = False
   If m_bolFMP = True Then
      m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, field(1), field(2), field(3), field(4))
   End If
   'end 2023/06/09
   'Added by Lydia 2023/07/28 FCP專利連結通知
   If field(1) = "FCP" Then
      m_PA177 = field(177)
   Else
      m_PA177 = ""
   End If
   'end 2023/07/28
   
   lblCaseField(0) = cp(1) + " - " + cp(2) + _
      IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
      IIf(cp(4) = "00", "", " - " + cp(4))
   lblCaseField(1) = cp(9)
   'lblCaseField(3) = cp(10) 'Remove by Morgan 2007/2/8 移往下面,申請國家要先抓
   ' 91.08.07  邱小姐說全部改成民國年  nickc
   'If intWhere <> 國外_CF Then
      lblCaseField(4) = ChangeTStringToTDateString(cp(6))
      lblCaseField(5) = ChangeTStringToTDateString(cp(7))
   'Else
   '   lblCaseField(4) = ChangeWStringToWDateString(cp(6))
   '   lblCaseField(5) = ChangeWStringToWDateString(cp(7))
   'End If
   lblCaseField(6) = cp(14)
   lblCaseField(7) = cp(13)
   txtCP64 = cp(64)
   If intCaseKind = 顧問 Then
      SetNameToCombo cboCaseName, field(6), "", ""
   Else
      SetNameToCombo cboCaseName, field(5), field(6), field(7)
   End If
   Select Case intCaseKind
                Case 專利
                           lblCaseField(2) = field(26)
                           lblCaseField(8) = field(9)
                           txtMemo = field(91)
                           txtCaseField(1) = field(57)
                Case 商標
                           lblCaseField(2) = field(23)
                           lblCaseField(8) = field(10)
                           txtMemo = field(58)
                           txtCaseField(1) = field(29)
                           'add by nickc 2007/07/11 加入審定號數
                           Label21(5).Visible = True
                           lblCaseField(9).Visible = True
                           lblCaseField(9) = field(15)
                Case 法務
                           lblCaseField(2) = field(11)
                           'add by nickc 2005/07/21 相關國家要代出來
                           lblCaseField(8) = field(15)
                           txtMemo = field(27)
                           txtCaseField(1) = field(8)
                Case 顧問
                           lblCaseField(2) = field(5)
                           'add by nickc 2005/07/21 國家是台灣
                           lblCaseField(8) = "000"
                           txtMemo = field(12)
                           txtCaseField(1) = field(9)
                Case Else
                           lblCaseField(2) = field(8)
                           lblCaseField(8) = field(9)
                           txtMemo = field(18)
                           txtCaseField(1) = field(15)
   End Select
   lblCaseField(3) = cp(10) 'Add by Morgan 2007/2/8 從上面移下來,申請國家要先抓
   If txtCaseField(1) = "Y" Then txtCaseField(1).BackColor = vbRed: BolFileClose = True
   If intCaseKind = 專利 Then
      txtCaseField(2) = field(89)
   Else
      txtCaseField(2).Enabled = False
   End If
   
   If cp(31) = "Y" And cp(27) = "" Then
      txtCaseField(4) = "N"
      If Not (cp(1) = "P" And lblCaseField(8) <> "000") Then 'Added by Morgan 2025/7/31 P非台灣案開放USER可手動改為要出指示信--玲玲
         txtCaseField(4).Enabled = False
      End If
   Else
      If lblCaseField(8) < "010" Then
         txtCaseField(4) = "N"
         txtCaseField(4).Enabled = False
      Else
         If cp(1) <> "CFT" And cp(1) <> "CFP" Then
            txtCaseField(4) = "N"
            If cp(1) = "P" Then
               '2005/3/22 ADD BY SONIA
               txtCaseField(4) = ""
               '2005/3/22 END
               txtCaseField(4).Enabled = True
            Else
               txtCaseField(4).Enabled = False
            End If
         Else
            txtCaseField(4).Enabled = True
         End If
      End If
   End If
   
   '91.4.30 CANCEL BY SONIA
   ''91.12.4 add by sonia
   'txtCaseField(4) = "N"
   'txtCaseField(4).Enabled = False
   ''91.12.4 end
   '92.4.30 END
   'edit by nickc 2006/06/22 從 dll copy 出
   'Select Case obj011.CheckChildCaseOrCaseRelation(field())
   Select Case CheckChildCaseOrCaseRelation(field())
                Case 1, 2
                           lblChildCase.Visible = True
                Case 0
                           lblChildCase.Visible = False
                Case -1, -2
                           GoTo err1
   End Select
   '下次期限(本所和法定)
   If cp(7) <> "" Then
      If cp(1) <> "P" And cp(1) <> "T" Then '92.4.11 add by sonia
         Dim tmpSQL  As String
         Dim tmpRs As New ADODB.Recordset
         tmpSQL = "select cf12,cf28 from casefee where cf01='" & cp(1) & "' and cf02='" & lblCaseField(8) & "' and cf03='" & lblCaseField(3) & "' "
         Set tmpRs = New ADODB.Recordset
         'Add By Cheng 2002/12/31
         tmpRs.CursorLocation = adUseClient
         tmpRs.Open tmpSQL, cnnConnection, adOpenStatic, adLockReadOnly
           'Modify By Cheng 2002/12/31
           '若有資料有大於0判斷
   '      If tmpRs.RecordCount <> 0 Then
         If tmpRs.RecordCount > 0 Then
              If CheckStr(tmpRs.Fields(0).Value) <> "" Then
                   Nextdate2 = ChangeWStringToTString(ChangeWDateStringToWString(DateAdd("d", Val(CheckStr(tmpRs.Fields(0).Value)), ChangeWStringToWDateString(ChangeTStringToWString(cp(7))))))
              Else
                  If CheckStr(tmpRs.Fields(1).Value) <> "" Then
                       Nextdate2 = ChangeWStringToTString(ChangeWDateStringToWString(DateAdd("M", Val(CheckStr(tmpRs.Fields(1).Value)), ChangeWStringToWDateString(ChangeTStringToWString(cp(7))))))
                  Else
                       Nextdate1 = ""
                       Nextdate2 = ""
                  End If
              End If
         Else
              Nextdate1 = ""
              Nextdate2 = ""
         End If
         Dim strDate(0 To 3) As String
         If Nextdate2 <> "" Then
               strDate(1) = cp(1)     '系統別
               strDate(2) = lblCaseField(8) '國家
               strDate(3) = ChangeTStringToWString(Nextdate2) '下次法定期限
               GetCtrlDT strDate()
               Nextdate1 = ChangeWStringToTString(strDate(0))
         End If
      
         'strTemp = GetCaseFeeNextDays(cp(1), lblCaseField(8), lblCaseField(3))
         '   If strTemp <> "" Then
         '      If intWhere <> 國外_CF Then
         '         strTemp1 = ChangeWStringToWDateString(ChangeTStringToWString(cp(7)))
         '         strTemp1 = DateAdd("D", Val(strTemp), strTemp1)
         '         Nextdate2 = ChangeWDateStringToTString(strTemp1)
         '         strTemp1 = DateAdd("D", -4, strTemp1)
         '         Nextdate1 = ChangeWDateStringToTString(strTemp1)
         '      Else
         '         strTemp1 = ChangeWStringToWDateString(cp(7))
         '         strTemp1 = DateAdd("D", Val(strTemp), strTemp1)
         '         Nextdate2 = ChangeWDateStringToWString(strTemp1)
         '         strTemp1 = DateAdd("D", -4, strTemp1)
         '         Nextdate1 = ChangeWDateStringToWString(strTemp1)
         '      End If
         '   End If
      Else
      '92.4.11 P領證及年費案件不管制半年
         Nextdate1 = ""
         Nextdate2 = ""
      End If
   End If
   '94.2.18 ADD BY SONIA 全部不預設下次期限, 改由人工輸入
   Nextdate1 = ""
   Nextdate2 = ""
   '94.2.18 END
   If txtCaseField(1) <> "Y" Then
      txtCaseField(6) = Nextdate1
      txtCaseField(7) = Nextdate2
   Else
      txtCaseField(6) = ""
      txtCaseField(7) = ""
   End If

   'Add by Morgan 2006/7/4
   cboReason.Clear
   strExc(0) = "Select * From ReasonofRelief Order By ROR01 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      .MoveFirst
      Do While Not .EOF
         cboReason.AddItem "" & .Fields("ROR01").Value & "--" & .Fields("ROR02").Value
         .MoveNext
      Loop
      End With
   End If
   'end 2006/7/4
   
   'Added by Lydia 2023/04/12 點選取消收文之收文號為國外部收文(CP12 like 'F%')且有收費(CP16>0)且尚未請款(CP60 is null)時，請在取消收文日期欄右邊加「取消收文之進度是否向客戶收款」欄。在存檔時，更新點選收文號之CP20。
   txtCaseField(3) = cp(20)
   If Left(cp(12), 1) = "F" And Val(cp(16)) > 0 And Val(cp(60)) = 0 Then
      lblCP20.Visible = True
      txtCaseField(3).Visible = True
   Else
      lblCP20.Visible = False
      txtCaseField(3).Visible = False
   End If
   'end 2023/04/12
   
Else
err1:
   bolLeave = True
   intLeaveKind = 1
   Unload Me
End If
If Len(Trim(txtCaseField(0))) = 0 Then txtCaseField(0) = ChangeWStringToTString(GetTodayDate)
If txtCaseField(1) = "Y" Then
   txtCaseField(6) = ""
   txtCaseField(7) = ""
   txtCaseField(6).Enabled = False
   txtCaseField(7).Enabled = False
Else
   txtCaseField(6) = Nextdate1
   txtCaseField(7) = Nextdate2
   txtCaseField(6).Enabled = True
   txtCaseField(7).Enabled = True
End If

Screen.MousePointer = varSaveCursor
Exit Sub
ErrHand:
ErrorMsg
Screen.MousePointer = varSaveCursor
End Sub

Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String, bolIsChina As Boolean

Select Case Index
             Case 1
                   If txtCaseField(1) = "Y" Then
                     txtCaseField(6) = ""
                     txtCaseField(7) = ""
                     txtCaseField(6).Enabled = False
                     txtCaseField(7).Enabled = False
                  Else
                     txtCaseField(6) = Nextdate1
                     txtCaseField(7) = Nextdate2
                     txtCaseField(6).Enabled = True
                     txtCaseField(7).Enabled = True
                  End If
             Case 2
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetCustomer(lblCaseField(Index), strTemp) Then
                        If ClsPDGetCustomer(lblCaseField(Index), strTemp) Then
                           lblPetitionName = strTemp
                        End If
             Case 3
                       If lblCaseField(8) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
                       'edit by nickc 2007/02/02 不用 dll 了
                       'If objPublicData.GetCaseProperty(cp(1), lblCaseField(Index), strTemp, bolIsChina) Then
                       If ClsPDGetCaseProperty(cp(1), lblCaseField(Index), strTemp, bolIsChina) Then
                           lblNextProgress = strTemp
                        End If
             Case 6
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
                        If ClsPDGetStaff(lblCaseField(Index), strTemp) Then
                           lblPromoter = strTemp
                        End If
             Case 7
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
                        If ClsPDGetStaff(lblCaseField(Index), strTemp) Then
                           lblSales = strTemp
                        End If
             Case 8
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetNation(lblCaseField(Index), strTemp) Then
                        If ClsPDGetNation(lblCaseField(Index), strTemp) Then
                           lblNation.Caption = strTemp
                        End If
End Select
End Sub

Private Sub Form_Activate()
   'Added by Lydia 2018/03/16 避免重複執行
   If bolActive Then
      Exit Sub
   Else
      bolActive = True
   End If
   'end 2018/03/16
   
   BolFileClose = False
   ReadAllData
   If txtCaseField(1) = "Y" Then
      MsgBox ("此案號已閉卷！！")
      bolLeave = True
      cmdok_Click (1)
      Exit Sub
   End If
   If cp(31) = "Y" And cp(27) = "" Then
      txtCaseField(1) = "Y"
      '92.12.9 add by sonia
      'Modify by Morgan 2006/7/4*
      'm_stReason = "14"
      'CheckKeyIn (8)
      For intI = 0 To Me.cboReason.ListCount - 1
         If Left(Me.cboReason.List(intI), 2) = "14" Then
            cboReason.ListIndex = intI
         End If
      Next
      '92.12.9 end
   '2015/7/13 ADD BY SONIA FCPhapi
   ElseIf cp(1) = "FCP" And cp(10) = "926" Then
      For intI = 0 To Me.cboReason.ListCount - 1
         If Left(Me.cboReason.List(intI), 2) = "15" Then
            cboReason.ListIndex = intI
         End If
      Next
   '2015/7/13 END
   End If
   
   'Added by Morgan 2015/11/3 指示信電子化
   'P非臺灣案指示信都要彈修改畫面來確認送判的內容
   'Modified by Morgan 2015/12/15 外專程序除外
   'Modified by Morgan 2018/8/16 +CFP電子化
   If (field(1) = "P" Or (field(1) = "CFP" And strSrvDate(1) >= CFP指示信電子化啟用日)) And field(9) <> "000" And Left(Pub_StrUserSt03, 1) <> "F" Then
      txtCaseField(5) = "Y"
      txtCaseField(5).Enabled = False
   End If
   'end 2015/11/3
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   'Memo by Amy 2025/08/06 不續辦但准通知 改為 後續准駁簡單報告
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bolLeave = False Then
      If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Cancel = 1
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2009/10/15
   If intLeaveKind <> 0 Then
      frm110102_1.Show
      If intLeaveKind = 2 Then
         frm110102_1.Cleartxt
      End If
   ElseIf intLeaveKind = 0 Then
     Unload frm110102_1
   End If
   ShowEditForm 'Added by Morgan 2018/8/22
   
   Set frm110102_2 = Nothing
End Sub

Private Sub txtCaseField_Change(Index As Integer)
If Index = 1 Then
   If txtCaseField(1) = "Y" Then
      txtCaseField(6) = ""
      txtCaseField(7) = ""
      txtCaseField(6).Enabled = False
      txtCaseField(7).Enabled = False
   Else
      txtCaseField(6) = Nextdate1
      txtCaseField(7) = Nextdate2
      txtCaseField(6).Enabled = True
      txtCaseField(7).Enabled = True
   End If
End If
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   'Modified by Lydia 2023/04/12 +3
   Case 1, 2, 4, 5, 3
      KeyAscii = UpperCase(KeyAscii)
   'Add By Sindy 2012/9/4
   Case 8
      KeyAscii = Pub_NumAscii(KeyAscii)
   '2012/9/4 End
End Select
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = -1 Then
   Cancel = True
   txtCaseField_GotFocus (Index)
End If
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, strCusTemp As String

CheckKeyIn = -1
Select Case intIndex
             Case 0
                        If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                           '2010/8/3 加val
                           If Val(txtCaseField(intIndex)) <= Val(GetTaiwanTodayDate) Then
                              CheckKeyIn = 1
                           Else
                              ShowMsg MsgText(8002)
                           End If
                         End If
             Case 1
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "Y" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(9174)
                        End If
             Case 2, 5
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "Y" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(9174)
                        End If
             'Modified by Lydia 2023/04/12 +3 取消收文之進度是否向客戶收款CP20
             Case 4, 3
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
             Case 6
                        If txtCaseField(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              If CheckReKey(txtCaseField(intIndex)) Then
                                  CheckKeyIn = 1
                                  txtCaseField(intIndex).Text = TransDate(PUB_GetWorkDay1(txtCaseField(intIndex).Text, True), 1)
                              Else
                                  CheckKeyIn = 0
                              End If
                           End If
                        End If
             'end 2020/07/09
             Case 7
                        If txtCaseField(intIndex) <> "" Then
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              '2010/8/3 加val
                              If Val(txtCaseField(6)) <= Val(txtCaseField(7)) Then
                                 If CheckReKey(txtCaseField(intIndex)) Then
                                    CheckKeyIn = 1
                                 Else
                                    CheckKeyIn = 0
                                 End If
                              Else
                                 ShowMsg MsgText(1033)
                              End If
                           End If
                        ElseIf txtCaseField(6) <> "" Then
                           ShowMsg MsgText(1033)
                           CheckKeyIn = 0
                        Else
                           CheckKeyIn = 1
                        End If
             
             Case Else
                        CheckKeyIn = 1
End Select
End Function

Private Sub txtCaseField_GotFocus(Index As Integer)
   TextInverse txtCaseField(Index)
   If Index = 3 Then
      'edit by nickc 2007/06/06 切換輸入法改用API
      'txtCaseField(Index).IMEMode = 1
      OpenIme
   Else
      'edit by nickc 2007/06/06 切換輸入法改用API
      'txtCaseField(Index).IMEMode = 2
      CloseIme
   End If
End Sub

Private Sub txtMemo_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtMemo.IMEMode = 1
   OpenIme
End Sub

'92.5.1 Add By SONIA
'取得有代理人之總收文號
Private Function GetCP09(ByVal ET01 As String, ByVal ET03 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'2005/3/25 ADD BY SONIA
Dim strTxt(1 To 5) As String, strTemp As String
Dim ii As Integer
Dim bolIsChina As Boolean
Dim WCP43 As String
'2005/3/25 END

   Wcp09 = ""
   StrSQLa = "Select CP09 From Caseprogress, (Select CP01 A1, CP02 A2, CP03 A3, CP04 A4 From CaseProgress Where CP09='" & cp(9) & "' ) A Where A.A1=CP01 AND A.A2=CP02 AND A.A3=CP03 AND A.A4=CP04 AND CP09 <'C' AND CP27 IS NOT NULL AND CP57 IS NULL ORDER BY CP27 DESC, CP09 DESC "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic
   '若有資料
   If rsA.RecordCount > 0 Then
       Wcp09 = "" & rsA("CP09").Value
   '2011/6/24 ADD BY SONIA CFT-013983
   Else
      Wcp09 = cp(9)
   '2011/6/24 END
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing

'2005/3/25 ADD BY SONIA
    ii = 1
    EndLetter ET01, Wcp09, ET03, strUserNum
    
    If cp(43) > "C" And cp(1) = "P" Then
      If lblCaseField(8) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
      '取得相關總收文號之案件性質
      WCP43 = ""
      StrSQLa = "Select CP10 From Caseprogress WHERE CP09='" & cp(43) & "'"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic
      '若有資料
      If rsA.RecordCount > 0 Then
          WCP43 = "" & rsA("CP10").Value
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCaseProperty(cp(1), WCP43, strTemp, bolIsChina) Then
      If ClsPDGetCaseProperty(cp(1), WCP43, strTemp, bolIsChina) Then
      End If
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & Wcp09 & "','" & ET03 & "','" & strUserNum & _
         "','案件性質分類','" & strTemp & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & Wcp09 & "','" & ET03 & "','" & strUserNum & _
         "','下一程序名稱','" & Me.lblNextProgress.Caption & "')"
      ii = ii + 1
    End If
    
    'Added by Morgan 2016/11/17
    '年費結案代理人不是案件代理人時指示信不要帶轉寄官方文件段落
    If lblCaseField(3) = "605" And cp(1) = "P" Then
         strExc(0) = "select substr(cp27||cp44,9) from caseprogress where cp01='" & field(1) & "' and cp02='" & field(2) & "' and cp03='" & field(3) & "' and cp04='" & field(4) & "' and cp10<>'605' and cp10<>'421' and cp09<'B' and cp27>0 and cp44 is not null"
         strExc(0) = strExc(0) & " union select substr(cp27||cp44,9) from caseprogress where cp01='" & field(1) & "' and cp02='" & field(2) & "' and cp03='" & field(3) & "' and cp04='" & field(4) & "' and cp10='605' and cp27>0 and cp44 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.RecordCount > 1 Then
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & Wcp09 & "','" & ET03 & "','" & strUserNum & _
                  "','非案件代理人不印','♀')"
               ii = ii + 1
            End If
         End If
    End If
    'end 2016/11/17
    
    If ii > 1 Then
      'edit by nickc 2007/02/05 不用 dll 了
      'If Not objLawDll.ExecSQL(ii, strTxt) Then
      If Not ClsLawExecSQL(ii, strTxt) Then
         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
      End If
    End If
'2005/3/25 END
End Function
'Add by Morgan 2005/5/18
'當國內新案未發文且銷案原因為14(新案未送件不辦)時，若有國外案未取消收文則發Mail通知承辦
Private Sub CaseMapMailCheck()

   Dim stSubject As String
   
   'Modify by Morgan 2005/6/27 原因要抓畫面值
   'If Not (cp(31) = "Y" And cp(27) = "" And cp(58) = "14") Then
   'Modify by Morgan 2010/12/29 只要原因是14 就通知,不管是不是選新案取消收文--郭
   'If Not (cp(31) = "Y" And cp(27) = "" And m_stReason = "14") Then
   If Not (cp(27) = "" And m_stReason = "14") Then
      Exit Sub
   End If
   'Modified by Morgan 2019/7/22 +未取消收文
   strSql = "select cm01,cm02,cm03,cm04,cp09,cp14 from casemap,caseprogress" & _
      " where cm10='0' and cm05='" & cp(1) & "' and cm06='" & cp(2) & "'" & _
      " and cm07='" & cp(3) & "' and cm08='" & cp(4) & "'" & _
      " and cp01(+)=cm01 and cp02(+)=cm02 and cp03(+)=cm03" & _
      " and cp04(+)=cm04 and cp31='Y' and cp27 is null and cp57 is null and cp14 is not null"
   
On Error GoTo ErrHnd

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Do While Not .EOF
            stSubject = "" & .Fields("cm01") & "-" & .Fields("cm02") & .Fields("cm03") & .Fields("cm04")
            stSubject = stSubject & " 之國內案 "
            stSubject = stSubject & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
            'modify by sonia 2021/9/23 加字--郭
            'stSubject = stSubject & " 已閉卷"
            stSubject = stSubject & " 已閉卷，請更改案件關聯。"
            PUB_SendMail strUserNum, "" & .Fields("cp14"), "" & .Fields("cp09"), stSubject
            .MoveNext
         Loop
      End If
   End With
   
ErrHnd:

   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'cancel by sonia 2025/1/3 取消查名費請款單
''Add by Sindy 2012/9/3
''新增國外請款資料
'Private Function AddAccData(strCP10 As String) As String
'Dim stA1k01 As String, stA1k03 As String, strA1K27 As String, strA1K28 As String
'Dim strPrintCust As String
'Dim stA1L05 As Double, stA1L07 As Double, stA1k11 As Double
'Dim stA1k08 As Double, stA1k10 As Double
'
'   AddAccData = "" 'Add By Sindy 2012/10/18
'   '1:先以"X"抓ACC1R0之國外請款單的自動編號, 並更新其流水號
'   stA1k01 = AccAutoNo(MsgText(815), 5)
'   AccSaveAutoNo MsgText(815), Right(stA1k01, 5)
'   '2:新增ACC1K0
'   '代理人編號
'   stA1k03 = PUB_GetA1K03(cp(1), cp(2), cp(3), cp(4))
'   '列印對象
'   strA1K27 = PUB_GetA1K27(cp(1), cp(2), cp(3), cp(4), strCP10)
'   If strA1K27 = "" Then strA1K27 = stA1k03
'   '請款對象
'   strA1K28 = PUB_GetA1K28(cp(1), cp(2), cp(3), cp(4), strCP10)
'   If strA1K28 = "" Then strA1K28 = stA1k03
'   '是否列印申請人
'   strPrintCust = PUB_GetA1K04(cp(1), cp(2), cp(3), cp(4), strA1K28, strCP10)
'   stA1L07 = 0 '折扣金額
'   'modify by sonia 2016/7/27 T-202178
'   'stA1L05 = PUB_GetUSXRate_1(strSrvDate(2), "RMB") * txtCaseField(8) '請款金額
'   'stA1k11 = stA1L05 '台幣金額
'   'stA1k10 = PUB_GetUSXRate_1(strSrvDate(2), "RMB") * (1 / PUB_GetDNRate(strSrvDate(2), "RMB")) '美金對台幣匯率
'   ''美金取整數位(無條件捨去)
'   'stA1k08 = Fix(Val(stA1k11) / stA1k10) '請款美金金額
'   stA1k10 = PUB_GetUSXRate_1(strSrvDate(2), "RMB")  '請款幣別對台幣匯率
'   stA1k08 = Val(txtCaseField(8)) '請款幣別金額
'   stA1L05 = stA1k10 * txtCaseField(8) '請款金額
'   stA1k11 = stA1L05 '台幣金額
'   'end
'   strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K08,A1K03,A1K27,A1K28,A1K04,A1K21,A1K19,A1K20 ) " & _
'            " VALUES  ('" & stA1k01 & "'," & strSrvDate(2) & ",NULL,0," & stA1k10 & "," & stA1k11 & ",NULL,'" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'" & _
'            ",'RMB'," & stA1k08 & ",'" & stA1k03 & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "','" & strUserNum & "'," & strSrvDate(2) & ",to_char(sysdate,'hh24miss'))"
'   cnnConnection.Execute strSql, intI
'   '3:新增一筆ACC1L0
'   strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L07,A1L02,A1L04,A1L05,A1L10,A1L08,A1L09) " & _
'            " VALUES  ('" & stA1k01 & "','" & cp(1) & "'," & stA1L07 & ",'001','" & strCP10 & "'," & stA1L05 & ",'" & strUserNum & "'," & strSrvDate(2) & ",to_char(sysdate,'hh24miss'))"
'   cnnConnection.Execute strSql, intI
'
'   PUB_UpdateA1k08 stA1k01 'Added by Morgan 2012/11/2 更新請款單外幣金額
'
'   '4:新增ACC1W0
'   strSql = "INSERT INTO ACC1W0 VALUES  ('" & stA1k01 & "','" & cp(9) & "')"
'   cnnConnection.Execute strSql, intI
'   '5:更新新增的C類收文號
'   strSql = "UPDATE CASEPROGRESS SET CP60='" & stA1k01 & "' WHERE CP09='" & cp(9) & "'"
'   cnnConnection.Execute strSql, intI
'   '6:自動分配點數
'   PUB_PointAutoassign stA1k01, True
'   AddAccData = stA1k01 'Add By Sindy 2012/10/18
'
'   'Added by Lydia 2016/11/21 以請款對象檢查是否存在於國外固定寄催款單代理人檔(ACC225)且下次寄發日期＞系統日，若存在則顯示訊息提醒操作人員
'   If PUB_ChkAcc225MsgList(stA1k01, strA1K28, cp(1), cp(2), cp(3), cp(4)) Then
'   End If
'   'end 2016/11/21
'End Function
'end 2025/1/3

'Added by Morgan 2015/5/18
'FCP台灣新型年費解除期限一案兩請提醒
Private Sub CheckFCPDualCase()
   
   If cp(10) = "605" And field(1) = "FCP" And field(9) = "000" And field(8) = "2" Then
      '若發明案尚未審定或核駁且未閉卷時，提醒使用者
      strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)" & _
         " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & field(1) & "' and cm02='" & field(2) & "' and cm03='" & field(3) & "' and cm04='" & field(4) & "'" & _
         " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & field(1) & "' and cm06='" & field(2) & "' and cm07='" & field(3) & "' and cm08='" & field(4) & "') X" & _
         ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 and pa08='1' AND pa57 is null and (pa16 is null or pa16='2')"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(2) = "此案為一案兩請且發明案 " & RsTemp(0) & " 尚未審定，請將卷宗交業務承辦告知客戶新型專利權若因未繳年費而當然消滅者，則將不予專利！"
         MsgBox strExc(2), vbExclamation
      End If
      
   End If

End Sub


