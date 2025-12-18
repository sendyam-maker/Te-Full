VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040104_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "內專發文(延期)"
   ClientHeight    =   5748
   ClientLeft      =   4956
   ClientTop       =   3720
   ClientWidth     =   8940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   8940
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   4995
      MaxLength       =   4
      TabIndex        =   10
      Top             =   5460
      Width           =   540
   End
   Begin VB.TextBox txtCP118 
      Height          =   300
      Left            =   7770
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3450
      Width           =   255
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1665
      MaxLength       =   2
      TabIndex        =   2
      Top             =   2610
      Width           =   300
   End
   Begin VB.TextBox txtCP84 
      Height          =   285
      Left            =   7515
      TabIndex        =   1
      Top             =   2301
      Width           =   1380
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1680
      TabIndex        =   6
      Top             =   3210
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   9
      Top             =   5460
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm040104_2.frx":0000
      Left            =   1200
      List            =   "frm040104_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   14
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   8004
      TabIndex        =   13
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5952
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6780
      TabIndex        =   12
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   4995
      MaxLength       =   7
      TabIndex        =   4
      Top             =   2910
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   3
      Top             =   2910
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1680
      MaxLength       =   9
      TabIndex        =   0
      Top             =   2310
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   18
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   17
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   16
      Top             =   660
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   15
      Top             =   660
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1605
      Left            =   225
      TabIndex        =   8
      Top             =   3780
      Width           =   8535
      _ExtentX        =   15050
      _ExtentY        =   2836
      _Version        =   393216
      Cols            =   10
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
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   816
      Left            =   7512
      TabIndex        =   5
      Top             =   2580
      Width           =   1500
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;1439"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   18
      Left            =   4095
      TabIndex        =   51
      Top             =   5505
      Width           =   765
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:         (Y:是)"
      Height          =   240
      Index           =   3
      Left            =   6570
      TabIndex        =   50
      Top             =   3480
      Width           =   1995
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   240
      TabIndex        =   49
      Top             =   1980
      Width           =   765
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   9
      Left            =   1200
      TabIndex        =   48
      Top             =   1950
      Width           =   7530
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "13282;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "延期月數:"
      Height          =   180
      Left            =   240
      TabIndex        =   47
      Top             =   2655
      Width           =   765
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   6570
      TabIndex        =   46
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label lblCP84 
      AutoSize        =   -1  'True
      Caption         =   "發文規費:"
      Height          =   180
      Left            =   6570
      TabIndex        =   45
      Top             =   2355
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   180
      X2              =   8760
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   8760
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   3555
      TabIndex        =   44
      Top             =   1500
      Width           =   720
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   8
      Left            =   4395
      TabIndex        =   43
      Top             =   1470
      Width           =   1710
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3016;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   7
      Left            =   3120
      TabIndex        =   42
      Top             =   3270
      Width           =   3360
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5927;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "欲延期期限:"
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   41
      Top             =   3540
      Width           =   945
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否列印指示信         (Y)"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   40
      Top             =   5460
      Width           =   2385
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   6
      Left            =   4395
      TabIndex        =   39
      Top             =   1710
      Width           =   1710
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3016;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   5
      Left            =   4395
      TabIndex        =   38
      Top             =   1230
      Width           =   1710
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3016;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   4
      Left            =   4395
      TabIndex        =   37
      Top             =   630
      Width           =   1710
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3016;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   3
      Left            =   1200
      TabIndex        =   36
      Top             =   1740
      Width           =   1740
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3069;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   2
      Left            =   1200
      TabIndex        =   35
      Top             =   1500
      Width           =   1740
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3069;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   1
      Left            =   1200
      TabIndex        =   34
      Top             =   1260
      Width           =   1740
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3069;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "代理人:"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   33
      Top             =   3270
      Width           =   585
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "延期後法定期限:"
      Height          =   180
      Left            =   3555
      TabIndex        =   32
      Top             =   2955
      Width           =   1305
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "延期後本所期限:"
      Height          =   180
      Left            =   240
      TabIndex        =   31
      Top             =   2955
      Width           =   1305
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "上次延期日:"
      Height          =   180
      Left            =   3555
      TabIndex        =   30
      Top             =   2355
      Width           =   945
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Label15"
      Height          =   180
      Left            =   4635
      TabIndex        =   29
      Top             =   2355
      Width           =   570
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "延期日:"
      Height          =   180
      Left            =   240
      TabIndex        =   28
      Top             =   2355
      Width           =   585
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Left            =   3555
      TabIndex        =   27
      Top             =   1740
      Width           =   765
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   240
      TabIndex        =   26
      Top             =   1740
      Width           =   765
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   240
      TabIndex        =   25
      Top             =   1500
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "收文日期:"
      Height          =   180
      Left            =   3555
      TabIndex        =   24
      Top             =   1260
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   240
      TabIndex        =   23
      Top             =   1260
      Width           =   585
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   0
      Left            =   1920
      TabIndex        =   22
      Top             =   990
      Width           =   6810
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "12012;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   21
      Top             =   660
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   3555
      TabIndex        =   20
      Top             =   660
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   19
      Top             =   1020
      Width           =   765
   End
End
Attribute VB_Name = "frm040104_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (Label2,lstNameAgent)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
'整理 by Morgan 2008/2/21
Option Explicit

Dim Situ As Integer
'Modify by Morgan 2005/8/1 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String
'Modify by Morgan 2008/2/21
'Dim cp(5 To 14) As String
Dim cp() As String

Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer
Dim m_bln_keyinValidate As Boolean
Dim m_bln_ShowMsgText5 As Boolean
Dim m_bln_ShowMsgText6 As Boolean
Public m_str_DL05 As String '延期記錄檔的資料來源
Dim m_str_DL06 As String
Dim m_NP22 As String    '延期案件性質
Dim m_ET02 As String '指示信收文號
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/23 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim m_CP09 As String 'Add by Morgan 2009/12/23 點選的收文號
Dim m_bolFMP As Boolean 'Add by Morgan 2009/12/24
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/20 是否為寰華

Dim m_DelayProPerty As String 'Added by Morgan 2011/12/6 被延期的案件性質
Dim m_NP23 As String '約定期限 Added by Morgan 2012/4/17
Dim m_bolTW107Extended As Boolean 'Added by Morgan 2013/10/2 台灣再審是否已延期過
Dim m_lngFee As String, m_strCP09 As String 'Added by Morgan 2013/10/22 收文規費,要檢查規費的收文號
Dim strFileName As String 'Add by Amy 2014/10/22 上傳的檔案名稱
Dim m_Subject As String 'Added by Morgan 2016/5/12
Dim stCP12 As String, stCP13 As String 'Added by Morgan 2021/1/28
'Added by Lydia 2022/05/23
Dim m_LosCP84 As String ' 法律所案源之規費
Dim m_LOS15 As String '法律所案源單號
Dim m_LOS02 As String '法律所案源類別
Dim m_LosMemo As String  'email說明

Private Function FormSave() As Boolean
Dim strTmp(0 To 3) As String, bolChk As Boolean
Dim i As Integer, ii As Integer
Dim strCP09 As String '總收文號
Dim strPromoteDate  As String '承辦期限
Dim strUpdate As String
Dim strCP30 As String 'Add by Morgan 2011/4/22
Dim st307Msg As String '分割案提醒訊息 Added by Morgan 2011/12/6
Dim strCF03 As String 'Added by Morgan 2012/3/21
Dim strNP23 As String 'Added by Morgan 2012/4/17

'Add By Cheng 2002/11/05
On Error GoTo ErrorHandler

FormSave = True
cnnConnection.BeginTrans
   
   'Add By Cheng 2002/06/24
   If Combo2 <> "" Then
      'Modify by Morgan 2008/2/21
      'cp(44) = ChangeCustomerL(Combo2)
      intI = InStr(Combo2, "-")
      If intI > 0 Then
         cp(44) = Left(Combo2, intI - 1)
         cp(116) = Mid(Combo2, intI + 1)
      Else
         cp(44) = Combo2
         cp(116) = ""
      End If
      cp(44) = ChangeCustomerL(cp(44))
      'end 2008/2/21
      
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.GetCaseThatCode(cp) Then cp(45) = ""
      If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   Else
      cp(44) = ""
      cp(116) = ""
      cp(45) = ""
   End If
   
   Select Case Situ
      Case 0 '前畫面按 延期 鈕
         strExc(1) = "DELETE FROM DATELIMIT WHERE DL01='" & cp(9) & "' AND DL02=" & TransDate(Text7, 2)
         cnnConnection.Execute strExc(1)
         
         strExc(2) = "INSERT INTO DATELIMIT (DL01,DL02,DL03,DL04,DL05,DL06) VALUES " & _
            "('" & cp(9) & "'," & TransDate(Text7, 2) & "," & CNULL(cp(6)) & "," & CNULL(cp(7)) & "," & CNULL(m_str_DL05) & ",'" & IIf(m_str_DL05 = "1", "", m_str_DL06) & "' )"
        cnnConnection.Execute strExc(2)
         
         '取得總收文號
         strCP09 = AutoNo("B", 6)

         'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'         'Modify by Amy 2014/09/05 for 台灣案電子化 此改確認case 1 是否也要改
'         If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
'            If pa(1) = "P" And pa(9) = 台灣國家代號 Then
'               'Modify by Amy 2015/02/13 修改判斷(此沒有客戶函)
'                 '1.    電子送件且規費>0,有收據
'                 '2.非電子送件且經發文室要計件,有回執
'               strExc(1) = PUB_GetLetterJudge(pa(1), cp(10))
'               If txtCP118 = "Y" Then
'                    If Val(txtCP84) > 0 Then
'                        PUB_AddLetterProgress strCP09, 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'                    End If
'               Else
'                    If Left(m_CP123s, 1) = "Y" Then
'                        PUB_AddLetterProgress strCP09, 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'                    End If
'               End If
'               'Add by Lydia 2015/01/13 延期發文開放可電子送件並必須輸入官方收文號
''               If (Val(txtCP84) > 0 Or txtCP118 <> "Y") Then
''                  '判斷同一天 沒有其他有規費的發文,新增信函進度(因會有回執or收據)
''                    If ChkOneDayHasCP84(pa(1), pa(2), pa(3), pa(4)) = True Then
''                        strExc(1) = PUB_GetLetterJudge(pa(1), cp(10))
''                        PUB_AddLetterProgress strCP09, 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
''                    End If
''               End If
'               'end 2015/01/13
'               'end 2015/02/13
'            End If
'         End If
'         'end 2014/09/05

         'Modify by Morgan 2004/8/11 加 cp84
         'Modify by Morgan 2005/7/14 加 cp110,CP22
         'Modify by Morgan 2008/2/21 +cp116
         'Modified by Lydia 2015/01/15 +cp118,cp64
         'Modified by Morgan 2016/1/8 +cp120
         'Modified by Morgan 2020/2/4 +cp71
         'Modified by Lydia 2021/05/25 +CP113工作時數
         strExc(3) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
            "CP09,CP10,CP12,CP13,CP14,CP20,CP22,CP26,CP32,CP27,CP43,CP44,CP45,CP84,CP110," & _
            "cp116,cp118,cp64,CP120,CP71,CP113 ) VALUES " & _
            "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & _
            strSrvDate(1) & "," & CNULL(PUB_GetWorkDay1(cp(6), True)) & "," & CNULL(cp(7)) & _
            ",'" & strCP09 & "','" & 延期 & "'," & CNULL(cp(12)) & "," & CNULL(cp(13)) & _
            ",'" & strUserNum & "','N'," & CNULL(cp(22)) & ",'N','N'," & TransDate(Text7, 2) & ",'" & cp(9) & _
            "'," & CNULL(cp(44)) & "," & CNULL(cp(45)) & "," & Format(Val(txtCP84.Text)) & "," & CNULL(cp(110)) & _
            "," & CNULL(cp(116)) & "," & CNULL(txtCP118.Text) & "," & CNULL(ChgSQL(Label2(9))) & ",'" & IIf(txtCP118 <> "" And pa(9) = "000", "Y", "") & "','" & Text9 & "', " & CNULL(txtCP113, True) & ")"
        cnnConnection.Execute strExc(3)
        
        m_ET02 = strCP09
        
        strUpdate = ""
        
        If m_bolFMP Or PUB_IfSetCP48() Then 'Add by Morgan 2010/10/1 新規則改所限時Trigger會清除承辦期限,待隔日凌晨重算
        
            '93.3.6 MODIFY BY SONIA 須延期案件一般都是期限快到，而承辦期限也多會被縮減為本所期限，所以延期時需用齊備日重新計算
            'strExc(4) = "UPDATE CASEPROGRESS SET CP06=" & TransDate(Text5, 2) & _
            '   ",CP07=" & TransDate(Text6, 2) & " WHERE CP09='" & cp(9) & "'"
            'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算(並先判斷已有齊備日)
            'strExc(0) = "Select NVL(CF04,0),EP06,CP48 From CaseProgress, Patent, Casefee, EngineerProgress Where CP09='" & cp(9) & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and PA01=cf01(+) and pa09=cf02(+) and cp10=cf03 and cp09=ep02(+)"
            strExc(0) = "Select cp01,pa09,cp10,ep06 From CaseProgress, Patent, EngineerProgress Where CP09='" & cp(9) & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp09=ep02(+) and ep06>0"
            'end 2007/10/11
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
                strPromoteDate = Pub_GetHandleDay(RsTemp("cp01"), RsTemp("pa09"), RsTemp("cp10"), RsTemp("ep06"), TransDate(Text5, 2), cp(9))
            End If
            If strPromoteDate <> "" Then
               strUpdate = ",CP48=" & CNULL(strPromoteDate)
            End If
            
         End If 'Add by Morgan 2010/10/1
         strExc(4) = "UPDATE CASEPROGRESS SET CP06=" & TransDate(Text5, 2) & ",CP07=" & TransDate(Text6, 2) & strUpdate
         'Added by Morgan 2020/5/19 大陸案備註內的「法限已加在途15天;」要清除--何淑華
         If pa(9) = "020" Then
            strExc(4) = strExc(4) & ",cp64=replace(cp64,'法限已加在途15天;','')"
         End If
         'end 2020/5/19
         strExc(4) = strExc(4) & " WHERE CP09='" & cp(9) & "'"
         '93.3.6 END
        'Add By Cheng 2002/11/05
        cnnConnection.Execute strExc(4)
        '92.6.30 add by sonia
        'cancel by sonia 2015/9/7 延期後不需工程師再輸收卷註記 P-104887
        'strExc(5) = "UPDATE ENGINEERPROGRESS SET EP27=NULL,EP31=NULL WHERE EP02='" & cp(9) & "'"
        'cnnConnection.Execute strExc(5)
        '92.6.30 end
         
         cp(9) = strCP09
         
         'Add by Morgan 2009/3/23 B類延期要補收文號
         If pa(9) = 台灣國家代號 Then
            m_CP09s = strCP09 & m_CP09s
         End If
         
      Case 1 '延期(案件性質404)的發文
         For i = 1 To MSHFlexGrid1.Rows - 1
            If Me.MSHFlexGrid1.TextMatrix(i, 0) = "v" Then
               strCF03 = MSHFlexGrid1.TextMatrix(i, 7) 'Added by Morgan 2012/3/21
               
               bolChk = True
               strTmp(2) = TransDate(Replace(MSHFlexGrid1.TextMatrix(i, 2), "/", ""), 2)
               strTmp(3) = TransDate(Replace(MSHFlexGrid1.TextMatrix(i, 3), "/", ""), 2)
               strTmp(1) = MSHFlexGrid1.TextMatrix(i, 6)
               strTmp(0) = MSHFlexGrid1.TextMatrix(i, 8)
               
               'Add by Morgan 2009/11/18
               m_str_DL06 = strTmp(0)
               m_CP09 = strTmp(1)
               'NP
               If Val(m_str_DL06) > 0 Then
                  m_str_DL05 = "1"
               'CP
               Else
                  m_str_DL05 = "2"
               End If
               'end 2009/11/18

               strExc(1) = "DELETE FROM DATELIMIT WHERE DL01='" & strTmp(1) & "' AND DL02=" & TransDate(Text7, 2)
               cnnConnection.Execute strExc(1), intI
               
               strExc(2) = "INSERT INTO DATELIMIT (DL01,DL02,DL03,DL04,DL05,DL06) VALUES " & _
                  "('" & strTmp(1) & "'," & TransDate(Text7, 2) & "," & CNULL(ChgSQL(strTmp(2))) & "," & CNULL(ChgSQL(strTmp(3))) & "," & CNULL(m_str_DL05) & ",'" & IIf(m_str_DL05 = "1", "", m_str_DL06) & "' )"
               cnnConnection.Execute strExc(2), intI
               
               'Modify by Morgan 2009/11/18 +CP更新
               If m_str_DL05 = "1" Then
                  m_NP22 = strTmp(0)
                  
                  'Added by Morgan 2012/4/17 +NP23 約定期限也要更新
                  If Val(m_NP23) > 0 Then
                     strNP23 = PUB_GetWorkDay1(CompDate(2, DateDiff("d", ChangeWStringToWDateString(strTmp(3)), ChangeWStringToWDateString(DBDATE(Text6))), m_NP23), True)
                  Else
                     strNP23 = "NP23"
                  End If
                  'end 2012/4/17
                  
                  'Modify by Morgan 2006/1/24 加NP01
                  'Modified by Morgan 2012/4/17 +NP23 約定期限也要更新
                  strExc(4) = "UPDATE NEXTPROGRESS SET NP08=" & CNULL(PUB_GetWorkDay1(DBDATE(Text5), True), True) & ",NP09=" & CNULL(DBDATE(Text6), True) & ",NP23=" & strNP23
                  'Added by Morgan 2020/5/19 大陸案備註內的「法限已加在途15天;」要清除--何淑華
                  If pa(9) = "020" Then
                     strExc(4) = strExc(4) & ",np15=replace(np15,'法限已加在途15天;','')"
                  End If
                  'end 2020/5/19
                  strExc(4) = strExc(4) & " WHERE NP22=" & strTmp(0) & " and np01='" & strTmp(1) & "'"
                  'Add By Cheng 2002/11/05
                  cnnConnection.Execute "begin user_data.user_notrigger:=1; end;" 'Add by Morgan 2010/7/13 +控制來函期限通知的 Trigger 不被觸發
                  cnnConnection.Execute strExc(4)
                  cnnConnection.Execute "begin user_data.user_notrigger:=0; end;" 'Add by Morgan 2010/7/13 +控制來函期限通知的 Trigger 不被觸發
               Else
                  m_NP22 = "" 'Add by Morgan 2011/4/22
                  strSql = "UPDATE CASEPROGRESS SET CP06=" & CNULL(PUB_GetWorkDay1(DBDATE(Text5), True), True) & ", CP07=" & CNULL(DBDATE(Text6), True)
                  'Added by Morgan 2020/5/19 大陸案備註內的「法限已加在途15天;」要清除--何淑華
                  If pa(9) = "020" Then
                     strSql = strSql & ",cp64=replace(cp64,'法限已加在途15天;','')"
                  End If
                  'end 2020/5/19
                  strSql = strSql & " WHERE CP09='" & strTmp(1) & "'"
                  cnnConnection.Execute strSql, intI
               End If
               
               'Added by Morgan 2011/12/6 台灣案的申復或再審期限要更新到分割案
               If MSHFlexGrid1.TextMatrix(i, 7) = "205" Or MSHFlexGrid1.TextMatrix(i, 7) = "107" Then
                  m_DelayProPerty = MSHFlexGrid1.TextMatrix(i, 7)
                  
               'Added by Morgan 2016/5/12
               ElseIf m_DelayProPerty = "" Then
                  m_DelayProPerty = MSHFlexGrid1.TextMatrix(i, 7)
               'end 2016/5/12
               
               End If
               'Exit For 'Remove by Morgan 2009/12/23 改可多選
            End If
         Next
         
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
            GoTo ErrorHandler
         End If
         
         m_ET02 = cp(9)
         
        'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'        'Modify by Amy 2014/09/05 for 台灣案電子化 此改確認case 0 是否也要改
'         If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
'            If pa(1) = "P" And cp(9) < "C" And pa(9) = 台灣國家代號 Then
'                'Modify by Amy 2015/02/13 修改、整理判斷條件
'                  '1.非新案、非改請且電子送件且規費>0,有收據
'                  '2.非新案、非改請非電子送件且經發文室要計件,有回執
'                If Not (InStr(NewCasePtyList, cp(10)) > 0 Or Left(cp(10), 1) = "3") Then
'                    strExc(1) = PUB_GetLetterJudge(pa(1), cp(10))
'                    If txtCP118 = "Y" Then
'                        If Val(txtCP84) > 0 Then
'                            PUB_AddLetterProgress cp(9), 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'                        End If
'                    Else
'                        If Left(m_CP123s, 1) = "Y" Then
'                            PUB_AddLetterProgress cp(9), 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'                        End If
'                    End If
''                   'Add by Lydia 2015/01/13 延期發文開放可電子送件並必須輸入官方收文號
''                   If (Val(txtCP84) > 0 Or txtCP118 <> "Y") Then '2015/01/13
''                        '判斷同一天 沒有其他有規費的發文
''                        If ChkOneDayHasCP84(pa(1), pa(2), pa(3), pa(4)) = True Then
''                            strExc(1) = PUB_GetLetterJudge(pa(1), cp(10))
''                            PUB_AddLetterProgress cp(9), 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
''                        End If
''                   End If
''                   'end 2015/01/13
'                End If
'                'end 2015/02/13
'            End If
'         End If
'         'end 2014/09/05
         
         'Modify by Morgan 2004/8/11 加 cp84
         'Modify by Morgan 2004/9/15 加 cp43
         'Modify by Morgan 2005/7/14 加 cp110
         'Modify by Morgan 2008/2/21 + cp116
         'Modify by Morgan 2011/4/22 +CP30
         'Modify by Lydia 2015/01/15 + cp118,cp64
         'Modified by Morgan 2016/1/8 +cp120
         'Modified by Morgan 2020/2/4 +cp71
         'Modified by Lydia 2021/05/25 +CP113工作時數
         'Modified by Lydia 2023/06/20 +CP14
         strExc(3) = "UPDATE CASEPROGRESS SET CP27=" & TransDate(Text7, 2) & ", CP44='" & cp(44) & "', CP45='" & cp(45) & "'" & _
                     ",CP84=" & Format(Val(txtCP84.Text)) & ",CP43='" & strTmp(1) & "',CP22=" & CNULL(cp(22)) & ",CP30='" & m_NP22 & "' " & _
                     ",cp110=" & CNULL(cp(110)) & ",cp116=" & CNULL(cp(116)) & ", CP118=" & CNULL(txtCP118.Text) & _
                     ",CP64=" & CNULL(ChgSQL(Label2(9))) & ",CP120='" & IIf(txtCP118 <> "" And pa(9) = "000", "Y", "") & "',CP71='" & Text9 & "' " & _
                     ",cp113=" & CNULL(txtCP113, True) & ",CP14=" & CNULL(cp(14)) & _
                     " WHERE CP09='" & cp(9) & "'"
         cnnConnection.Execute strExc(3), intI
         
      End Select
   
      'Add by Morgan 2009/3/23
      If pa(9) = 台灣國家代號 Then
         PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130
         'Add by Amy 2015/02/13 更新收據/回執設定
         'Modify by Amy 2015/03/06 +發文日參數
         PUB_UpdateLP19 cp(1), cp(2), cp(3), cp(4), m_CP09s, m_CP123s, Text7
   
         'Added by Morgan 2011/12/6 台灣案的申復或再審期限要更新到分割案
         If pa(9) = "000" And (m_DelayProPerty = "205" Or m_DelayProPerty = "107") Then
            strSql = "select cp09 from divisioncase,caseprogress" & _
               " where dc05='" & pa(1) & "' and dc06='" & pa(2) & "'" & _
               " and dc07='" & pa(3) & "' and dc08='" & pa(4) & "'" & _
               " and cp01(+)=dc01 and cp02(+)=dc02 and cp03(+)=dc03 and cp04(+)=dc04 and cp10='307' and cp27||cp57 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               st307Msg = ""
               '可以有多個分割案
               Do While Not RsTemp.EOF
                  strExc(1) = PUB_Update307RefTw(RsTemp(0))
                  If strExc(1) <> "" Then
                     st307Msg = st307Msg & strExc(1) & vbCrLf
                  End If
                  RsTemp.MoveNext
               Loop
            End If
         End If
      'Added by Morgan 2012/3/21
      Else
         If cp(10) = 延期 Then
            PUB_SetArriveDate cp(9), strCF03
         Else
            PUB_SetArriveDate strCP09, cp(10)
         End If
         
         'Added by Morgan 2013/7/5
         '有法限都要管制最終提申
         If cp(7) <> "" Then
            strExc(1) = DBDATE(cp(7))
            strExc(2) = PUB_GetWorkDay1(strExc(1), True)
            strSql = " insert into nextprogress a (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
               " values('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','996'" & _
               "," & strExc(2) & "," & strExc(1) & ",'" & strUserNum & "',GETNP22)"
            cnnConnection.Execute strSql, intI
         End If
         'end 2013/7/5
         
         'Added by Morgan 2016/5/12
         '指示信電子化
         If Text8 = "Y" And Left(Pub_StrUserSt03, 1) <> "F" Then
            strExc(1) = Text1 & "-" & Text2 & IIf(Text3 & Text4 = "000", "", "-" & Text3 & "-" & Text4)
            m_Subject = "委託 " & strExc(1) & " 案延期" & GetPrjState4(Text1 & "-" & Text2 & "-" & Text3 & "-" & Text4, m_DelayProPerty) & IIf(Text9 <> "", Text9 & "個月", "")
            If ExistCheck("AppForm", "AF01", m_ET02, "", False) = False Then
               'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
               strExc(2) = PUB_GetLetterJudgeNew("2", pa(1), 延期, pa(9), m_DelayProPerty)
               PUB_AddAppForm m_ET02, True, strExc(2), m_Subject
            End If
         End If
         'end 2016/5/12
         
      'end 2012/3/21
      End If
      'Added by Lydia 2022/05/23 更新法律所案源單號
      If m_LOS15 <> "" Then
         strSql = "Update CaseProgress Set CP162='" & m_LOS15 & "'  where cp09 = " & CNULL(IIf(cp(10) = 延期, cp(9), strCP09))
         cnnConnection.Execute strSql
      End If
      'end 2022/05/23
      'Added by Lydia 2023/03/17 若延期的相關總收文號為B2案源時，同時新增法務案之內部收文39延期
      If m_LOS15 <> "" And m_LOS02 = "B2" Then
          Call PUB_InsertLosBCP(m_LOS15, DBDATE(Text7), DBDATE(Text5), DBDATE(Text6))
      End If
      'end 2023/03/17
      
    '******只能寫於此(commit前)，不可改至別地方 ******
    'Add by Amy 2014/11/27 B類延期上傳申請書
    If strFileName <> "" Then
        If PUB_CheckPDF2(strCP09, 1, True, strFileName) = False Then GoTo ErrorHandler
    End If
    '******end 2014/11/27                                               ******
      cnnConnection.CommitTrans
      
      'Add by Morgan 2011/12/6
      If st307Msg <> "" Then MsgBox st307Msg
      
      Exit Function
     
ErrorHandler:
    cnnConnection.RollbackTrans
    FormSave = False
End Function

Private Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
Dim strTxt(1 To 10) As String, strTmp As String
Dim ii As Integer
    EndLetter ET01, ET02, ET03, strUserNum
    
    ii = 1
   Select Case Situ
      Case 0 '按鈕
         'Modify by Morgan 2005/5/11
         'strExc(0) = "SELECT CPM03,CPM04 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE CP09='" & cp(43) & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
         strExc(0) = "SELECT CPM03,CPM04 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE CP09='" & cp(43) & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
         
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If pa(9) < "010" Then
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                   "','案件性質分類','" & RsTemp.Fields(0).Value & "')"
            Else
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                   "','案件性質分類','" & RsTemp.Fields(1).Value & "')"
            End If
            ii = ii + 1
         End If
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
             "','下一程序名稱','" & Label2(2) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
             "','代理人中文','" & Label2(7) & "')"
         ii = ii + 1
      Case 1 '延期
         If Val(m_NP22) > 0 Then
            strExc(0) = "SELECT CPM03,CPM04,NP01 FROM NEXTPROGRESS,CASEPROPERTYMAP WHERE NP02='" & pa(1) & _
               "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP02=CPM01(+) AND NP07=CPM02(+) AND NP22=" & m_NP22
         'Add by Morgan 2009/12/23 +cp
         Else
            strExc(0) = "SELECT CPM03,CPM04,CP09 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE CP09='" & m_CP09 & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
         End If
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If pa(9) < "010" Then
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                   "','下一程序名稱','" & RsTemp.Fields(0).Value & "')"
            Else
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                   "','下一程序名稱','" & RsTemp.Fields(1).Value & "')"
            End If
            ii = ii + 1
         End If
         strExc(0) = "SELECT CPM03,CPM04 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE CP09='" & RsTemp.Fields(2).Value & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If pa(9) < "010" Then
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                   "','案件性質分類','" & RsTemp.Fields(0).Value & "')"
            Else
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                   "','案件性質分類','" & RsTemp.Fields(1).Value & "')"
            End If
            ii = ii + 1
         End If
   End Select
   
   'Added by Morgan 2020/1/13 指示信改抓畫面月數 Ex:P-118895 -- 品薇
   If Val(Text9) > 0 Then
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','延期月數','" & PUB_ChgNumber2Chinese(Val(Text9)) & "')"
      ii = ii + 1
   End If
   'end 2020/1/13
   
    'edit by nickc 2007/02/05 不用 dll 了
    'If Not objLawDll.ExecSQL(ii - 1, strTxt) Then
    If Not ClsLawExecSQL(ii - 1, strTxt) Then
        MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
    End If
End Sub

Private Function Process(Index As Integer) As Boolean
Dim strTempName As String  '2012/7/23 ADD BY SONIA
'Added by Lydia 2019/12/05
Dim strFilePath As String '記錄智慧局收文文號
Dim bolUp As Boolean '是否需要上傳檔案到卷宗區
Dim strNewCP64 As String '保留進度備註

   If Text7 = "" Then
      MsgBox "延期日不得為空值 !", vbCritical
      Text7.SetFocus
      Exit Function
   End If
   'Add By Cheng 2002/01/02
   If Text5 = "" Then
      MsgBox "延期後本所期限不得為空值 !", vbCritical
      Text5.SetFocus
      Exit Function
   End If
   If ChkRange(Text5, Text6, "本所期限、法定期限") = False Then
      Text5.SetFocus
      Exit Function
   End If
   
   'Add By Cheng 2002/01/02
   If Text7 = "" Then
      MsgBox "延期日不得為空值 !", vbCritical
   Else
      If Not ChkDate(Text7.Text) Then
         Me.Text7.SetFocus
         TextInverse Text7
         Exit Function
      Else
         '2011/12/8 MODIFY BY SONIA 發文日可輸系統日的下一個工作日
         'If Me.Text7.Text > (ServerDate - 19110000) Then
         '   MsgBox "延期日不得大於系統日 !", vbCritical
         If DBDATE(Val(Text7)) > DBDATE(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))) Then
            MsgBox "延期日不得大於系統日下一個工作日 !", vbCritical
         '2011/12/8 END
            TextInverse Text7
            Exit Function
         End If
      End If
   End If
   Text5_Validate False
   If m_bln_keyinValidate = False Then
      Me.Text5.SetFocus
      Text5_GotFocus
      Exit Function
   End If
            
   'Add By Cheng 2002/03/08
   '檢查輸入資料的完整性
   If CheckDataIntegrity = False Then Exit Function
   
   'Add By Cheng 2002/05/22
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Function
   
   'Add by Morgan 2009/3/23 設定是否算發文室案件
   'Modified by Morgan 2015/7/3 排除電子送件者
   'If pa(9) = 台灣國家代號 Then
   m_CP09s = "": m_CP123s = ""
   If pa(9) = 台灣國家代號 Then
      If txtCP118 = "Y" Then
         'Added by Morgan 2016/5/16 電子送件也要記錄主管機關
         If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text7, True, True) = False Then
            Exit Function
         End If
         'end 2016/5/16
      Else
   'end 2015/7/3
         'Add by Morgan 2009/4/28
         If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text7, True) = False Then
            Exit Function
         End If
         If m_CP123s = "Y" Then
            'modify by sonia 2014/6/23 加傳發文規費, P-108903
            If ModifyDispatch(cp(9), m_CP09s, m_CP123s, txtCP84, Text7, IIf(Situ = 0, True, False)) = False Then
                Exit Function
            End If
         End If
      End If
   End If
   
   strNewCP64 = Label2(9).Caption 'Added by Lydia 2019/12/05 保留進度備註
   
   'Add by Amy 2014/10/22 P台灣案發文控制
   If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And pa(9) = 台灣國家代號 Then
        strFileName = ""
        If cp(10) = "404" Then
            If pa(1) = "P" And cp(9) < "C" Then
                If cp(9) < "B" Then
                    '檢查本所案號所有A類未發文,一定要有接洽單才可發文
                    'Modify by Amy 2014/11/27 取消ChkOneDayHasCP27判斷,接洽單改檢查,因考慮可能同時發文其他案件性質情形
                    'If PUB_CheckPDF2(cp(9), 0, True, strExc(0)) = False And ChkOneDayHasCP27(pa(1), pa(2), pa(3), pa(4), cp(5) + 19110000) = False Then
                    If PUB_CheckPDF3(Text1, Text2, Text3, Text4) = False Then
                        Exit Function
                    End If
                End If
                'AB類申請書確認檢查,符合條件才可發文
                'Modified by Morgan 2015/3/17
                'If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
                If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" Then
                  If PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
               'end 2015/3/17
                    MsgBox "無申請書PDF檔 ,不可發文!", vbInformation
                    Exit Function
                  End If 'Added by Morgan 2015/3/17
                End If
            End If
        Else
            '按前畫面延期案鈕進入,只判斷申請書匯入資料夾(存本所號.pdf)是否有檔案(先不匯入)
            'Modify by Amy 2014/11/27 +只檢查匯入資料夾參數(因B類收文號未產生,傳入的cp(9)怕卷宗區已有資料)
            If PUB_CheckPDF2(cp(9), 1, False, strFileName, , True) = False Then
                MsgBox "無申請書PDF檔 ,不可發文!", vbInformation
                Exit Function
            End If
        End If
   'Added by Morgan 2016/6/29 非臺灣案電子化
   ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      If cp(10) = "404" And cp(9) < "B" And Left(cp(12), 1) <> "F" Then
          If PUB_CheckPDF3(Text1, Text2, Text3, Text4) = False Then
              Exit Function
          End If
      End If
   'end 2016/6/29
   End If
   'end 2014/10/22

   'Add by Lydia 2015/01/13 延期發文開放可電子送件並必須輸入官方收文號
    If txtCP118 = "Y" And Val(txtCP84) = 0 And pa(9) = 台灣國家代號 Then 'Modified by Morgan 2024/1/19 +台灣案
       m_CP123s = ""

       strExc(0) = InputBox("請輸入智慧局收文文號!!")
       If strExc(0) = "" Then
          Exit Function
       Else
          'Modified by Lydia 2019/12/05
          'Label2(9) = "智慧局收文文號:" & strExc(0) & ";" & Label2(9) 'CP64進度備註
          strFilePath = strExc(0)  '記錄智慧局收文文號
          strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & Label2(9) '保留進度備註
          'end 2019/12/05
       End If
    End If
   'end 2015/01/13
   
    'Added by Lydia 2019/12/05 檢查是否有電子送件的檔案
    bolUp = False
    If txtCP118.Text = "Y" And strFilePath <> "" And pa(9) = 台灣國家代號 Then 'Modified by Morgan 2024/1/19 +台灣案
        strExc(1) = cp(82)
        If Val(cp(82)) > 0 Then
            If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                 strExc(1) = ""
            End If
        End If
        If Val(strExc(1)) = 0 Then
           'Modified by Lydia 2020/03/23 改成先判斷是否上傳檔案; ex.P-124220發明申請 因為上傳檔案在FormSave前,所以沒抓到出名代理,造成無POA檔仍設電子檔案齊備CP121=Y
           'If Pub_AutoEsetToCppByP(True, pa(1), pa(2), pa(3), pa(4), pa(8), IIf(cp(10) <> "404", "", cp(9)), "404", strFilePath, Text7.Text) = False Then
           If Pub_AutoEsetToCppByP(True, pa(1), pa(2), pa(3), pa(4), pa(8), "", "404", strFilePath, Text7.Text) = False Then
                Exit Function
           End If
           bolUp = True
        End If
    End If
    
    If pa(9) = 台灣國家代號 Then 'Added by Lydia 2020/01/17 限台灣案
        Label2(9).Caption = strNewCP64  '檢查完畢，更新備註欄位
    End If
    'end 2019/12/05
         
   If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Function
    'Mark 2014/11/27 發文完但檔案上傳失敗，導致沒申請書改至commit前
'    'Add by Amy 2014/10/22 B類延期上傳申請書(原:寫於包於FormSave的Transation中怕檔案上傳後又做RollBack，導致檔案不存在)
'    If strFileName <> "" Then
'        strExc(0) = PUB_CheckPDF2(cp(9), 1, True, strFileName)
'    End If
'    'end 2014/10/22
   Process = True
         
   'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
   PUB_CheckEMail Combo2
   PUB_CheckEMail pa(75), pa(144)
   If pa(145) <> "" Then
      PUB_CheckEMail pa(75), pa(145)
   End If
   'end 2008/2/20
   
   'Add by Morgan 2007/6/14
   If pa(9) = "000" Then
      PUB_ReAsignInform pa(1), pa(2), pa(3), pa(4), cp(9)
   End If
   
   '2012/7/23 add by sonia
   '台灣案發文規費與收文規費不符時,mail給智權人員
   'Modified by Morgan 2013/10/22 若延期發文且被延期的程序也已收文則規費及收文號應改用被延期的程序
   'If txtCP84.Enabled = True And pa(9) = "000" And Val(Me.txtCP84.Text) <> Val(cp(17)) Then
'2014/5/5 cancel by sonia 敏惠於1030326提出專利案延期發文時,不管控收文費用
'   If txtCP84.Enabled = True And pa(9) = "000" And Val(Me.txtCP84.Text) <> m_lngFee Then
      '2013/7/2 modify by sonia 改用共用module
      'Modified by Morgan 2013/10/22 若延期發文且被延期的程序也已收文則規費及收文號應改用被延期的程序
      'PUB_ChkOfficialFee cp(9), Me.txtCP84.Text
'      PUB_ChkOfficialFee m_strCP09, Me.txtCP84.Text
'   End If
'2014/5/5 end
   '2012/7/23 end
   
   'Added by Lydia 2022/05/23 PT案(傳入收文號)取得法律案源之發文規費，並且有輸入發文規費才做檢查
   If txtCP84.Enabled = True And pa(9) = "000" And m_LosCP84 <> "0" Then
      If Val(Trim(txtCP84.Text)) <> 0 And Val(m_LosCP84) <> Val(Trim(txtCP84.Text)) Then
          PUB_ChkOfficialFee m_ET02, Me.txtCP84.Text, IIf(txtCP118 = "Y", "A", ""), m_LosMemo
      End If
   End If
   'end 2022/05/23
   
   If Text8.Text = "Y" Then
      'Modify by Morgan 2006/5/24 AB類延期的指示信一樣，改傳本所案號&404 -- 郭
      'StartLetter "02", "30"
      'NowPrint cp(9), "02", "30", True, strUserNum, 0
      'Modify by Morgan 2006/8/28 再改回收文號,因為彼所案號會抓不到
      'StartLetter "02", pA(1) & pA(2) & pA(3) & pA(4) & "&404", IIf(Situ = "0", "31", "30")
      'NowPrint pA(1) & pA(2) & pA(3) & pA(4) & "&404", "02", IIf(Situ = "0", "31", "30"), True, strUserNum, 0
      
      'Modify by Morgan 2008/7/11 配合定稿地址開窗
      'StartLetter "02", cp(9), "30"
      'NowPrint cp(9), "02", "30", True, strUserNum, 0
      StartLetter "02", m_ET02, "41"
      'Modify by Amy 2014/09/09 +strLetterRecNo
      'Mofify by Morgan 2016/5/12 指示信電子化非外專程序改都彈定稿維護視窗
      'NowPrint m_ET02, "02", "41", True, strUserNum, 0, , , , , , , , , , , , m_ET02
      If Left(Pub_StrUserSt03, 1) = "F" Then
         NowPrint m_ET02, "02", "41", True, strUserNum
      Else
         NowPrint m_ET02, "02", "41", True, strUserNum, , , , , , , , , , , , , m_ET02
         frm1105_1.m_RecNo = m_ET02
         frm1105_1.m_PdfName = PUB_CaseNo2FileName(Text1, Text2, Text3, Text4) & "." & cp(10) & ".DATA.PDF"
         frm1105_1.m_Subject = m_Subject
         frm1105_1.Show
      End If
      'end 2016/5/12
      'end 2008/7/11
   End If
   
   'Added by Morgan 2013/4/30
   'FMP延期報告
   If m_bolFMP Then
      strUserNum = strFMPNum
      StartLetter2 "02", m_ET02, "51"
      'Modify by Amy 2014/09/09 +strLetterRecNo
      NowPrint m_ET02, "02", "51", False, strUserNum, 0
      strUserNum = strUser1Num
   End If
   'end 2013/4/30
   
   'Add By Cheng 2002/04/30
   '若有未發文資料顯示警告
   If cp(10) = 延期 Then PUB_GetCPunIssueDatas "" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)
   
    'Add by Lydia 2015/01/13 延期發文時請發E-MAIL給延期的承辦工程師,延期後的本所期限
    strExc(0) = "select c2.cp09 b01,c2.cp10 b02,cpm03,cpm04,c2.cp14 b04 from caseprogress c1,caseprogress c2,casepropertymap " & _
                "where c1.cp09='" & cp(9) & "' and c1.cp10 = '404' and c1.cp43=c2.cp09(+) " & _
                "and c2.cp01=cpm01(+) and c2.cp10=cpm02(+) "
       
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        strExc(1) = Mid(Text5, 1, 3) & "年" & Mid(Text5, 4, 2) & "月" & Mid(Text5, 6, 2) & "日"
        strExc(0) = ""
        If Left(RsTemp!b01, 1) = "C" Then
           '未收文->讀取最後收文智權人員
           strExc(2) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
           strSql = ""
           If m_NP22 <> "" Then strSql = " and np22='" & m_NP22 & "'"
           strSql = " select np01 n01,np07 n02,cpm03,cpm04 from nextprogress,casepropertymap where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                    " and np01= '" & RsTemp!b01 & "' and np02=cpm01(+) and np07=cpm02(+) " & strSql
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If pa(9) = 台灣國家代號 Then
                 strExc(3) = RsTemp!cpm03
               Else
                 strExc(3) = RsTemp!cpm04
               End If
            End If
        Else
            If IsNull(RsTemp!b04) Then
               '未分案發給操作人員
               strExc(2) = strUserNum
               strExc(0) = "(未分案)"
            Else
               strExc(2) = RsTemp!b04
            End If
            If pa(9) = 台灣國家代號 Then
              strExc(3) = RsTemp!cpm03
            Else
              strExc(3) = RsTemp!cpm04
            End If
        End If
        strExc(0) = IIf(Text3 & Text4 = "000", Text1 & Text2, Text1 & Text2 & Text3 & Text4) & "," & Trim(strExc(3)) & ",延期後的本所期限為:" & strExc(1) & strExc(0)
        'PXXXXXX,案件性質(延期的相關收文號),延期後的本所期限為X年X月X日
        PUB_SendMail strUserNum, strExc(2), "", strExc(0), "如旨"
    End If
    'end 'Add by Lydia 2015/01/13
    
   'Added by Lydia 2019/12/05 發文時，電子送件自動上傳檔案到卷宗區; 重新發文(CP82>0)不做搬檔
                                             '因為有前一畫面按延期進來的,所以存檔後才上傳檔案
   If bolUp = True Then '是否可以上傳檔案,前面已判斷
      '經過FormSave處理,cp(9) = 1.傳入的延期收文號 or 直接延期產生的B類收文號
      If Pub_AutoEsetToCppByP(False, pa(1), pa(2), pa(3), pa(4), pa(8), cp(9), 延期, strFilePath, Text7.Text) = False Then
           Exit Function
      End If
   End If
   'end 2019/12/05
End Function

Private Sub cmdok_Click(Index As Integer)
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0
         'Modify by Morgan 2010/2/10 改呼叫函數方式以便鎖定按鍵
         cmdOK(Index).Enabled = False
         If Not Process(Index) Then
            cmdOK(Index).Enabled = True
         Else
            'Add By Sindy 2013/5/20
            If frm040104_1.bolIsEMPFlow = True Then
               Unload frm040104_1
               frm090202_4.Show
               frm090202_4.QueryData
            Else
            '2013/5/20 End
               frm040104_1.Show
               ' 90.07.11 modify by louis (回第一個畫面重新再查詢)
               frm040104_1.Clear
            End If
            Unload Me
         End If
         
      Case 1
         'Add By Sindy 2013/5/20
         If frm040104_1.bolIsEMPFlow = True Then
            Unload frm040104_1
            frm090202_4.Show
            frm090202_4.QueryData
         Else
         '2013/5/20 End
            frm040104_1.Show
         End If
         Unload Me
      Case 2
         'Add By Sindy 2013/5/20
         If frm040104_1.bolIsEMPFlow = True Then
            Unload frm040104_1
            frm090202_4.Show
            frm090202_4.QueryData
         Else
         '2013/5/20 End
            Unload frm040104_1
         End If
         Unload Me
   End Select
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(0) = pa(5)
      Case "英"
         Label2(0) = pa(6)
      Case "日"
         Label2(0) = pa(7)
   End Select
End Sub

Private Sub Combo2_Click()
   Combo2_Validate False
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
 Dim strTempName As String
 Cancel = False
   strExc(0) = Combo2
   If pa(9) = 台灣國家代號 Then
      If strExc(0) <> "" Then
         MsgBox "申請國家為台灣時，必須空白 !", vbCritical
         Combo2 = ""
         Cancel = True
      End If
   Else
      If strExc(0) = "" Then
         MsgBox "申請國家非台灣時，不可空白 !", vbCritical
         Cancel = True
      Else
         'Modify By Cheng 2002/07/08
         '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'         If objPublicData.GetAgent(strExc(0), strTempName) = True Then
         If PUB_GetAgentName(pa(1), strExc(0), strTempName) = True Then
            Combo2.Text = strExc(0)
            Label2(7).Caption = strTempName
         Else
            Label2(7).Caption = ""
            Cancel = True
         End If
      End If
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If Cancel = False Then
         If PUB_CheckStatus(Combo2.Text) = False Then Cancel = True
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   'Add by Morgan 2005/8/9
   ReDim pa(TF_PA)
   ReDim cp(TF_CP)
   
   With frm040104_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      Situ = Val(Right(.Tag, 1))
      cp(9) = Left(.Tag, Len(.Tag) - 1)
   End With
   Text7 = strSrvDate(2)
   
   ReadPatent
   
   cp(110) = "" '要清空,否則若重新發文會殘留前次發文資料,當新案有改出名人而本程序未改選將會造成不一致 Added by Morgan 2012/9/7
   
   'Add by Morgan 2005/7/14
   '台灣加出名代理人清單供勾選
   lstNameAgent.Clear
   If pa(9) = "000" Then
      PUB_SetOurAgent lstNameAgent, pa(), cp(110), , True   'Modified by Morgan 2021/12/14 +傳入bForm2=True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2005/7/14 end
   
   Label2(1) = cp(9)
   'Text6.Enabled 7= False
   'InitGrid 9, MSHFlexGrid1
   '93.3.25 ADD BY SONIA
   If pa(9) <> 台灣國家代號 Then
      Text8 = "Y"
   End If
   '93.3.25 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2023/03/17
   Set frm040104_2 = Nothing
End Sub

Private Sub ReadPatent()
Dim Lbl As Object
Dim m_Fee As String         '銷帳服務費 2012/8/1 add by sonia
Dim m_Official As String    '銷帳規費   2012/8/1 add by sonia
   
   For Each Lbl In Label2
      Lbl = ""
   Next
   Label15 = ""
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   cp(1) = pa(1)
   cp(2) = pa(2)
   cp(3) = pa(3)
   cp(4) = pa(4)
   Select Case pa(1)
      Case "P"
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.ReadPatentDatabase(pA(), intWhere) Then FormShow
         If ClsPDReadPatentDatabase(pa(), intWhere) Then FormShow
      Case "PS"
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.ReadServicePracticeDatabase(pA(), intWhere) Then FormShow
         If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then FormShow
   End Select
   
'4
   strExc(0) = "select count(*) from datelimit where dl01='" & cp(9) & "'"
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If RsTemp.Fields(0) > 0 Then
      strExc(0) = "select max(dl02) from datelimit where dl01='" & cp(9) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then Label15 = TransDate(RsTemp.Fields(0), 1)
   End If
   
   If pa(9) = 台灣國家代號 Then
      strExc(1) = "CPM03,"
   Else
      strExc(1) = "CPM04,"
   End If
       
    'Add By Cheng 2002/11/29
    cp(10) = ""
   
   'Add by Morgan 2004/8/11 加 CP17
   'Add by Morgan 2005/5/11 加 CP43
   'Add by Morgan 2010/3/2 加 CP64
   '2012/8/1 MODIFY BY SONIA 加 CP77
   'Modify by Amy 2014/10/14 +CP05
   'Add by Lydia 2015/01/13 +CP118
   'Modified by Lydia 2021/05/25 +CP113工作時數
   strExc(0) = "select cp05,cp06,cp07," & strExc(1) & "CP10,CP12,CP13,CP14, CP17, CP43,CP110,CP64,CP77,CP05,CP118,CP113 " & _
       "from caseprogress,casepropertymap " & _
      "where cp09='" & cp(9) & "' AND cp01=cpm01(+) and cp10=cpm02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      cp(110) = "" & .Fields("CP110")
      'Add by Morgan 2004/8/11
      cp(17) = "" & .Fields("CP17")
      '2012/8/1 add by sonia 若有銷帳則要扣除銷帳規費
      If Val("" & .Fields("CP77")) > 0 Then
         If GetCP77Detail(cp(9), m_Fee, m_Official) = True Then
            cp(17) = cp(17) - m_Official
         End If
      End If
      '2012/8/1 end
      txtCP84.Tag = cp(17)
      m_lngFee = Val(cp(17)) 'Added by Morgan 2013/10/22
      m_strCP09 = cp(9) 'Added by Morgan 2013/10/22
      
      'Add by Morgan 2005/5/11
      cp(43) = "" & .Fields("CP43")
      cp(5) = "" & .Fields("CP05") 'Add by Amy 2014/10/14

      If Not IsNull(.Fields(0)) Then Label2(5) = TransDate(.Fields(0), 1)
      If Not IsNull(.Fields(1)) Then
         Label2(3) = TransDate(.Fields(1), 1)
         cp(6) = .Fields(1)
      End If
      If Not IsNull(.Fields(2)) Then
         Label2(6) = TransDate(.Fields(2), 1)
         cp(7) = .Fields(2)
      End If
      If Not IsNull(.Fields(3)) Then Label2(2) = .Fields(3)
      If Not IsNull(.Fields(5)) Then cp(12) = .Fields(5)
      
      Label2(9) = "" & .Fields("CP64") 'Add by Morgan 2010/3/2
      
      'Add by Morgan 2009/12/24
      'Modified by Morgan 2021/1/27
      'If Left(cp(12), 1) = "F" And pa(10) <> "000" Then
      stCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
      stCP12 = GetSalesArea(stCP13)
      'Modified by Lydia 2023/06/20 pa(10)=> pa(9)
      If Left(stCP12, 1) = "F" And pa(9) <> "000" Then
      'end 2021/1/27
         m_bolFMP = True
      Else
         m_bolFMP = False
      End If
      'end 2009/12/24
      If Not IsNull(.Fields(6)) Then cp(13) = .Fields(6)
      If Not IsNull(.Fields(7)) Then
         cp(14) = .Fields(7)
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(14), strExc(0)) Then Label2(8) = strExc(0)
         'If ClsPDGetStaff(cp(14), strExc(0)) Then Label2(8) = strExc(0) 'Mark by Lydia 2023/06/20
      End If
      'Added by Lydia 2023/06/20
      m_bolFMP2 = False
      If m_bolFMP = True Then  '判斷寰華案
         m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4))
         '寰華案:承辦人為外專程序時,改為操作人員
         If m_bolFMP2 = True Then
            cp(14) = GetFCPUser("" & .Fields(7))
         End If
      End If
      If cp(14) <> "" Then
         If ClsPDGetStaff(cp(14), strExc(0)) Then Label2(8) = strExc(0)
      End If
      'end 2023/06/20
'7
      If Not IsNull(.Fields(4)) Then cp(10) = .Fields(4)
      cp(10) = .Fields("CP10").Value
      
      m_DelayProPerty = .Fields("cp10") 'Added by Morgan 2011/12/6
      
      'Add by Lydia 2015/01/13 延期發文開放可電子送件並必須輸入官方收文號
      cp(118) = "": txtCP118 = ""
      cp(118) = "" & .Fields("cp118")
      If cp(118) <> "" Then txtCP118 = "Y"
      
      If pa(9) = "020" Then txtCP118 = "Y" 'Added by Morgan 2024/1/19 大陸案預設電子送件--郭
      
      'Added by Lydia 2021/05/25
      cp(113) = "": txtCP113 = ""
      cp(113) = "" & .Fields("cp113")
      txtCP113 = cp(113)
      'end 2021/05/25
      
      'Add By Cheng 2002/08/19
      m_str_DL06 = Empty
      
      If cp(10) = 延期 Then
         If Right(strExc(1), 1) = "," Then
            strExc(1) = Left(strExc(1), Len(strExc(1)) - 1)
         End If
         'Modify by Morgan 2009/11/18 +抓CP未發文有期限資料,且下一程序要排除程序管制的案件性質
         'Add by Morgan 2010/3/2 加 NP15,CP64
         'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
         'Modified by Morgan 2012/4/17 +NP23 約定期限
         'Modified by Morgan 2013/10/22 +CP17 收文規費
         strExc(0) = "select ''," & strExc(1) & ",sqldatet(NP08),sqldatet(NP09)" & _
            ",NP13,NP14,NP01,NP07,NP22,NP08,NP15,NP23,0 from nextprogress,casepropertymap where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
            " and (np06 is null or np06='') and NP02=CPM01(+) and NP07=CPM02(+)" & strNpSqlOfNoSalesDuty & _
            " UNION ALL select ''," & strExc(1) & ",sqldatet(cp06),sqldatet(cp07)" & _
            ",cp08,cp40,cp09,cp10,0,cp06,CP64,0,cp17 from caseprogress,casepropertymap where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
            " and cp27 is null and CPM01(+)=cp01 and CPM02(+)=cp10 and cp09<>'" & cp(9) & "' and cp07>0"
            
         intI = 1
         Dim rsTemp1 As New ADODB.Recordset
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI <> 2 Then Set MSHFlexGrid1.Recordset = rsTemp1
         'Add By Cheng 2002/08/19
         If rsTemp1.RecordCount > 0 Then
            m_str_DL06 = "" & rsTemp1("NP22").Value
            m_NP23 = "" & rsTemp1("NP23") 'Added by Morgan 2012/4/17
         End If
      End If
      'Added by Lydia 2022/05/23 法律所案源：取得案源類別、發文規費、email加註
      If cp(10) = 延期 Then
          strExc(0) = "select cp162 from caseprogress where cp09='" & cp(43) & "' "
      Else
          strExc(0) = "select cp162 from caseprogress where cp09='" & cp(9) & "' "
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
          m_LOS15 = "" & RsTemp.Fields("cp162")
          '限制B2類發文規費; 只需考慮B2類, A類不會有PT案,B1類PT案不會繳規費, C類已取消(有也是照原來規則不必改)
          'Memo by Lydia 2022/09/16 模組回傳m_LOS02
          m_LosCP84 = PUB_GetLosCP84(m_LOS15, pa(1), pa(2), pa(3), pa(4), "B2", m_LOS02, m_LosMemo)
      End If
      'end 2022/05/23
   End If
   End With
   GridHead MSHFlexGrid1
'5
   Text7.Text = strSrvDate(2)
   
   
'Modified by Morgan 2013/10/2
'   'Added by Morgan 2012/3/9 預設來函期限月數--敏惠
'   'Modified by Morgan 2012/3/15 +控制台灣案才要--敏惠
'   'Modified by Morgan 2012/9/25 台灣再審延期預設抓收費表設定
'   If cp(43) > "C" And pa(9) = "000" And cp(10) <> "107" Then
'      strExc(0) = "select cp134 from caseprogress,nextprogress where cp09='" & cp(43) & "' and cp134>0" & _
'         " and np01(+)=cp09 and np07<>'107'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         Text9 = RsTemp.Fields(0)
'         Text9_Validate False
'      End If
'   '2013/9/16 add by sonia P-093517先收申復再收A類延期
'   ElseIf pa(9) = "000" And cp(10) <> "107" Then
'      strExc(0) = "select c2.cp134 from caseprogress c1,caseprogress c2 where c1.cp09='" & cp(43) & "' and c1.cp43=c2.cp09(+) and c2.cp134>0" & _
'         " and c2.cp10<>'107'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         Text9 = RsTemp.Fields(0)
'         Text9_Validate False
'      End If
'   '2013/9/16 end
'   End If
'   'end 2012/3/9
'   If Text6 = "" Then
'      'Modified by Morgan 2012/7/31 改用原來法定期限計算
'      'strExc(0) = TransDate(Text7.Text, 2)
'      'Modified by Morgan 2012/9/25 改呼叫函數
'      'If cp(7) > 0 Then
'      '   strExc(0) = TransDate(cp(7), 2)
'      '   GetCaseFee pa(1), pa(9), cp(10), strExc, pa(8)
'      'End If
'      SetDatebyCaseFee cp(10)
'      'end 2012/9/25
'   End If
   SetNewDueDate cp(43), cp(10)
'end 2013/10/2

'8
   'Modified by Lydia 2016/10/27 新案有申請人指定國外代理人檔則預設
   'AddAgent Combo2, pa
   AddAgent Combo2, pa, , , , cp(9), pa(9), pa(26)
End Sub

'910626 Sieg
Private Sub GetCaseFee(pa01 As String, PA09 As String, CP10 As String, strResult() As String, PA08 As String)
   'Add by Morgan 2010/3/11 台灣新型修正預設延1個月
   Dim strDateS(3) As String
   
   If PA08 = "2" And PA09 = "000" And CP10 = "204" Then
      strDateS(1) = pa01
      strDateS(2) = PA09
      strDateS(3) = CompDate(1, 1, strResult(0))
      GetCtrlDT strDateS
      Text6 = TransDate(strDateS(3), 1)
      Text5 = TransDate(PUB_GetWorkDay1(strDateS(0), True), 1)
      m_bln_ShowMsgText5 = False
      m_bln_ShowMsgText6 = False
   Else
   'end 2010/3/11
   
      '2005/7/7 MODIFY BY SONIA
      'If objLawDll.GetCaseFeeDelay(PA01, PA09, CP10, strResult()) Then
      If ClsLawGetCaseFeeDelay(pa01, PA09, CP10, strResult()) Then
      '2005/7/7 END
         
         'Add By Cheng 2002/01/24
         '若是電腦自動算出的日期且未經修改, 不要顯示是否確定更改的訊息
         m_bln_ShowMsgText5 = False
         m_bln_ShowMsgText6 = False
         
         If pa(9) < "010" Then
            Text6 = TransDate(strResult(1), 1)
            Text5 = TransDate(strResult(2), 1)
           'Add By Cheng 2003/12/08
           '延期後本所期限若非工作天則抓最近工作天
           Me.Text5.Text = TransDate(PUB_GetWorkDay1(Me.Text5.Text, True), 1)
         End If
      End If
      
   End If
End Sub

Private Sub FormShow()
   Label2(4) = pa(11)
   Combo1.ListIndex = 0
   Label2(0) = pa(5)
End Sub

Private Sub MSHFlexGrid1_Click()
   SetGrid2Date MSHFlexGrid1
End Sub

Private Sub SetGrid2Date(oGrid As MSHFlexGrid)
   Dim strTmp As String
   If oGrid.TextMatrix(oGrid.row, 1) = "" Then Exit Sub
   'Modify by Morgan 2009/12/23 改多選(補文件),且判斷非台灣的
   'GridClick oGrid, intLastRow, 0
   GridClick oGrid, intLastRow, 0, 1
   intLastRow = MSHFlexGrid1.row
   
   'Add by Morgan 2009/12/29 +期限檢查
   If oGrid.TextMatrix(intLastRow, 0) = "v" Then
      strExc(2) = Replace(oGrid.TextMatrix(intLastRow, 3), "/", "")
      If DBDATE(strExc(2)) <> cp(7) Then
         MsgBox "所點選案件性質的法定期限與延期程序不同，不可點選！"
         oGrid.TextMatrix(intLastRow, 0) = ""
      
      ElseIf pa(9) = "000" Then
      
         'Added by Morgan 2013/10/22 若延期發文且被延期的程序也已收文則規費及收文號應改用被延期的程序
         If oGrid.TextMatrix(intLastRow, 6) <> "" Then
            m_strCP09 = oGrid.TextMatrix(intLastRow, 6)
            m_lngFee = Val(oGrid.TextMatrix(intLastRow, 12))
            txtCP84.Tag = m_lngFee
         Else
            m_strCP09 = cp(9)
            m_lngFee = Val(cp(17))
         End If
         'end 2013/10/22
         
   'end 2009/12/29
         '93.6.27 MODIFY BY SONIA 以延期日計算延期後期限
         'strTmp = MSHFlexGrid1.TextMatrix(intLastRow, 9)
         'Modified by Morgan 2012/7/31 改用原法定期限計算
         'strTmp = TransDate(Text7.Text, 2)
         If Text9 = "" Then
         'end 2012/7/31
            strTmp = cp(7)
            '93.6.27 END
            If oGrid.TextMatrix(intLastRow, 0) = "v" And strTmp <> "" Then
               strExc(0) = TransDate(strTmp, 2)
               'Modified by Morgan 2012/9/25 改呼叫函數
               'GetCaseFee pa(1), pa(9), oGrid.TextMatrix(intLastRow, 7), strExc, pa(8)
               'Modified by Morgan 2013/10/1
               'SetDatebyCaseFee oGrid.TextMatrix(intLastRow, 7)
               SetNewDueDate oGrid.TextMatrix(intLastRow, 6), oGrid.TextMatrix(intLastRow, 7)
            Else
               Text5 = ""
               Text5.Enabled = True
               Text6 = ""
            End If
         End If 'Added by Morgan 2012/7/31
      End If
   End If
   
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
  CloseIme
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
'Add By Cheng 2002/01/24
If KeyCode <> vbKeyTab Then
   m_bln_ShowMsgText5 = True
End If
End Sub

'Removed by Morgan 2012/7/31 本所期限跳離不用重新計算法定期限
'Private Sub Text5_LostFocus()
'    'Add By Cheng 2002/12/03
'    If Me.Text5.Text <> "" Then
'        If m_bln_ShowMsgText5 Then
'           'If MsgBox("是否確定修改延期本所期限 !", vbYesNo) = vbNo Then
'           '   Me.Text5.SetFocus
'           '   Text5_GotFocus
'           'Else
'              m_bln_ShowMsgText6 = False
'              If pa(9) < "010" Then Text6 = TransDate(CompDate(2, 2, TransDate(Text5.Text, 2)), 1)
'           'End If
'        Else
'           m_bln_ShowMsgText6 = False
'           If pa(9) < "010" Then Text6 = TransDate(CompDate(2, 2, TransDate(Text5.Text, 2)), 1)
'        End If
'    End If
'End Sub

Private Sub Text5_Validate(Cancel As Boolean)
'Add By Cheng 2001/01/02
m_bln_keyinValidate = False
Cancel = False

'Modify by Morgan 2009/12/23 改存檔前檢查就好
'   If Text5 = "" Then
'      MsgBox "延期本所期限不可為空值，請重新輸入 !", vbCritical
'      Cancel = True
'   Else
   If Text5 <> "" Then
'end 2009/12/23
      'Add By Cheng 2001/12/17
      'Modify by Morgan 2010/8/10 百年蟲
      'If Me.Text5.Text < (ServerDate - 19110000) Then
      If Val(Text5) < Val(strSrvDate(2)) Then
         MsgBox "延期後本所期限不得小於系統日 !", vbCritical
         TextInverse Text5
         Cancel = True
         Exit Sub
      End If
      If ChkDate(Text5.Text) Then
      Else
         Cancel = True
      End If
        'Add By Cheng 2003/12/08
        '若本所期限非工作天則直接調整至最近的工作天
        If Cancel = False Then
            Me.Text5.Text = TransDate(PUB_GetWorkDay1(Me.Text5.Text, True), 1)
        End If
        'End
   End If
   If Cancel = True Then TextInverse Text5

'Add By Cheng 2001/01/02
m_bln_keyinValidate = True
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
  CloseIme
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
'Add By cheng 2002/01/24
If KeyCode <> vbKeyTab Then
   m_bln_ShowMsgText6 = True
End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
'Modify by Morgan 2009/12/23 改存檔前檢查就好
'   If Text6 = "" Then
'      MsgBox "延期法定期限不可為空值，請重新輸入 !", vbCritical
'      Cancel = True
'   Else
   If Text6 <> "" Then
'end 2009/12/23
      If ChkDate(Text6.Text) Then
      
'Modify by Morgan 2009/12/23 改存檔前檢查就好
'         If m_bln_ShowMsgText6 Then
'            Cancel = Not ChkRange(Text5, Text6, "")
'         Else
'            Cancel = Not ChkRange(Text5, Text6, "")
'         End If
'end 2009/12/23

      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text6
End Sub

Private Sub GridHead(oGrid As MSHFlexGrid)
 Dim i As Integer
   FixGrid oGrid
   With oGrid
      .Visible = False
      If .Cols < 12 Then .Cols = 12
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1000: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1400: .Text = "機關文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 2000: .Text = "相關人"
      .CellAlignment = flexAlignLeftCenter
      .col = 6: .ColWidth(6) = 0 '總收文號
      .col = 7: .ColWidth(7) = 0 '下一程序
      .col = 8: .ColWidth(8) = 0 'NP22
      .col = 9: .ColWidth(9) = 0 'NP08
      .col = 10: .ColWidth(10) = 1400: .Text = "備註" 'Add by Morgan 2010/3/2
      For intI = 11 To .Cols - 1
         .ColWidth(intI) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_LostFocus()
    'Add By Cheng 2002/12/03
    If Me.Text7.Text <> "" Then
        'Modify by Morgan 2010/8/10 百年蟲
        'If Me.Text7.Text > (ServerDate - 19110000) Then
        '2011/12/8 MODIFY BY SONIA 發文日可輸系統日的下一個工作日
        'If Val(Text7) > Val(strSrvDate(2)) Then
        '   MsgBox "延期日不得大於系統日 !", vbCritical
        If DBDATE(Val(Text7)) > DBDATE(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))) Then
           MsgBox "延期日不得大於系統日下一個工作日 !", vbCritical
        '2011/12/8 END
           Me.Text7.SetFocus
           Text7_GotFocus
        
        'Removed by Morgan 2012/7/31 改用原法定期限計算,延期日僅紀錄為發文日
        'Else
        '   strExc(0) = TransDate(Text7.Text, 2)
        '   GetCaseFee pa(1), pa(9), cp(10), strExc, pa(8)
        'End 2012/7/31
        
        End If
    End If
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
'Add By Cheng 2001/01/02
m_bln_keyinValidate = False
Cancel = False
   
   If Text7 = "" Then
      MsgBox "延期日不得為空值 !", vbCritical
      Cancel = True
   Else
      '2011/12/8 MODIFY BY SONIA 發文日可輸系統日的下一個工作日
      'If Not ChkDate(Text7.Text) Then
      If Not ChkDate(Text7) Or DBDATE(Val(Text7)) > DBDATE(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))) Then
         MsgBox "發文日期不正確或發文日大於系統日下一個工作日，請重新輸入 !", vbCritical
      '2011/12/8 END
         Cancel = True
      Else
      End If
   End If

If Cancel Then TextInverse Text7

'Add By Cheng 2001/01/02
m_bln_keyinValidate = True

'Added by Morgan 2022/7/18
If Text7.Tag <> Text7.Text Then
   SetDate
   Text7.Tag = Text7.Text
End If
'end 2022/7/18
End Sub

Private Sub Text8_GotFocus()
  TextInverse Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add By Cheng 2002/03/08
Private Function CheckDataIntegrity() As Boolean
Dim Cancel As Boolean
'Add By Cheng 2002/11/29
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim bFind As Boolean
Dim nIndex  As Integer
   'add by nickc 2008/05/01
   If IsDebt(pa(9), cp(9)) Then
        MsgBox "未收款且無 欲收款日期 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        GoTo IntegrityOrNot
   End If
Cancel = False


'檢查代理人欄位
Combo2_Validate Cancel
If Cancel = True Then
   Me.Combo2.SetFocus
   GoTo IntegrityOrNot
End If

    'Add By Cheng 2002/11/29
   ' 當案件性質為延期時, 未收文期限至少要選取一筆
   If cp(10) = "404" Then
      If Me.MSHFlexGrid1.Rows <= 1 Then
         strTit = "檢核資料"
         strMsg = "未收文期限無資料, 無法執行延期的處理"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
       GoTo IntegrityOrNot
      End If
      
      bFind = False
      For nIndex = 1 To Me.MSHFlexGrid1.Rows - 1
         If Me.MSHFlexGrid1.TextMatrix(nIndex, 0) = "v" Then
            'Added by Morgan 2012/8/13
            '台灣的申復修正只能延期一次(FCP有例外,程式不控制--靜芳)
            'Modified by Morgan 2012/12/18 +再審也只能延期1次
            'Modified by Morgan 2013/10/2 再審可延兩次
            If pa(9) = 台灣國家代號 And (MSHFlexGrid1.TextMatrix(nIndex, 7) = "204" Or MSHFlexGrid1.TextMatrix(nIndex, 7) = "205") Then
               strExc(1) = MSHFlexGrid1.TextMatrix(nIndex, 6)
               strExc(0) = "select cp09 from caseprogress where cp43='" & strExc(1) & "' and cp10='404' and cp27>0"
               '若延期的是CP則還要考慮NP是否有延期過
               If Left(strExc(1), 1) <> "C" Then
                  strExc(0) = strExc(0) & " union select cp09 from caseprogress a where cp43 in (select b.cp43 from caseprogress b where b.cp09='" & strExc(1) & "') and cp10='404' and cp27>0"
               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "台灣案" & MSHFlexGrid1.TextMatrix(nIndex, 1) & "只能延期一次，該程序已有延期紀錄不可再延期！"
                  GoTo IntegrityOrNot
               End If
            End If
            'end 2012/8/13
            bFind = True
            Exit For
         End If
      Next nIndex
      If bFind = False Then
         strTit = "檢核資料"
         strMsg = "請先選取欲延期期限的資料來做延期的處理"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
       GoTo IntegrityOrNot
      End If
   End If

   'Add by Morgan 2005/7/14
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
   'Added by Morgan 2015/6/30
   If txtCP118 = "Y" And cp(22) = "N" And pa(9) = "000" Then
      MsgBox "電子送件不可不出名！", vbCritical
      Exit Function
   End If
   'end 2015/6/30
   
CheckDataIntegrity = True
Exit Function

IntegrityOrNot:
CheckDataIntegrity = False
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.Text5.Enabled = True Then
   Cancel = False
   Text5_Validate Cancel
   If Cancel = True Then
      Me.Text5.SetFocus
      Text5_GotFocus
      Exit Function
   End If
End If

If Me.Text6.Enabled = True Then
   Cancel = False
   Text6_Validate Cancel
   If Cancel = True Then
      Me.Text6.SetFocus
      Text6_GotFocus
      Exit Function
   End If
End If

If Me.Text7.Enabled = True Then
   Cancel = False
   Text7_Validate Cancel
   If Cancel = True Then
      Me.Text7.SetFocus
      Text7_GotFocus
      Exit Function
   End If
End If

'Add by Morgan 2004/8/11
If txtCP84.Enabled = True Then
   Cancel = False
   txtCP84_Validate Cancel
   If Cancel = True Then
      txtCP84.SetFocus
      txtCP84_GotFocus
      Exit Function
   End If
End If

'Add by Morgan 2004/9/14
If Combo2.Enabled = True Then
   Cancel = False
   Combo2_Validate Cancel
   If Cancel = True Then
      Combo2.SetFocus
      Exit Function
   End If
End If

   'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
   If Pub_ChkACS112isNull(pa(1), pa(2), pa(3), pa(4), txtCP113) = True Then
         txtCP113.SetFocus
         txtCP113_GotFocus
         Exit Function
   End If
   'end 2021/05/25
   
   'Added by Lydia 2022/05/23 法律所案源：B2類PT案取得法律案源之發文規費，並且有輸入發文規費才做檢查
   If txtCP84.Enabled = True And pa(9) = "000" Then
       If m_LosCP84 <> "0" Then
          If Val(m_LosCP84) <> Val(Trim(txtCP84.Text)) Then
              If MsgBox("法律所收文規費[" & Trim(Val(m_LosCP84)) & "] 與實際發文規費[" & Trim(Val(txtCP84.Text)) & "]不同", vbOKCancel) = vbCancel Then
                  txtCP84_GotFocus
                  txtCP84.SetFocus
                  Exit Function
              End If
          End If
       End If
   End If
   'end 2022/05/23

'Added by Morgan 2024/1/19
If pa(9) = "020" And txtCP118 = "" Then
   If MsgBox("請確認本案是否為紙本送件？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      txtCP118.SetFocus
      Exit Function
   End If
End If
'end 2024/1/19

TxtValidate = True
End Function

Private Sub Text9_GotFocus()
   TextInverse Text9
   CloseIme
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   
   'Modify by Morgan 2011/9/15 內專都用延期日計算
   'If Val(Text9) > 0 Then
   '   Text6 = TransDate(CompDate("1", Val(Text9), cp(7)), 1)

   'Modify by Morgan 2012/7/31 台灣也改用原法定計算
   'If Val(Text9) > 0 And Text7 <> "" Then
   If Val(Text9.Tag) <> Val(Text9) Then
      SetDate 'Modified by Morgan 2022/7/19 原程式抽出寫成函數以便共用
   'end 2012/7/731
   End If
   Text9.Tag = Text9 'Added by Morgan 2012/7/31
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

'Add by Morgan 2004/8/11
Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_Validate(Cancel As Boolean)
   '台灣
   If pa(9) = "000" Then
      If m_bolTW107Extended = True Then
         If Val(txtCP84.Text) <> 0 Then
            MsgBox "台灣再審第2次延期發文不需繳交規費！", vbExclamation
            txtCP84_GotFocus
            Cancel = True
         End If
      Else
      'end 2013/10/2
      
      '未收文延期不判斷
      'Modify by Morgan 2004/9/17 都要檢查
      'If m_str_DL05 <> "1" Then
         'Modified by Morgan 2013/10/22 考慮延期發文且原程序也已發文改用 m_lngFee 判斷規費
         'If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
         '   If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
         
         '20140327START REMARK By eric
         'If Val(txtCP84.Text) <> m_lngFee And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
         '   If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & m_lngFee & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
         '      txtCP84.Tag = txtCP84.Text
         '   Else
         '      txtCP84_GotFocus
         '      cancel = True
         '   End If
         'End If
         '20140327END
         
      'End If
      End If
   End If
End Sub

'Removed by Morgan 2012/7/31 沒有使用
''2005/7/7 COPY FROM objLawDll之GetCaseFeeDelay,因為台灣申復案新型延期為文到次日30天,發明及設計用CASEFEE之設定文到次日60天
''傳進西元年，傳出西元年
'Public Function GetCaseFeeDelay(ByVal CF01 As String, ByVal CF02 As String, ByVal CF03 As String, ByRef CFOther() As String) As Boolean
'Dim RsTemp As New ADODB.Recordset
'Dim strQty As String
'On Error GoTo ErrHnd
'   GetCaseFeeDelay = True
'   CFOther(1) = CFOther(0) '法定
'   CFOther(2) = CFOther(0) '本所
'   '2005/7/7 MODIFY BY SONIA
'   'strQty = "SELECT CF22,CF25,CF27 FROM CASEFEE WHERE CF01='" & CF01 & "' AND CF02='" & CF02 & "' AND CF03='" & CF03 & "'"
'   '2008/10/22 modify by sonia 加修正204,並由30天改為1個月
'   If CF01 = "P" And CF02 = "000" And (CF03 = "205" Or CF03 = "204") And pa(8) = "2" Then
'      strQty = "SELECT 0,1,CF27 FROM CASEFEE WHERE CF01='" & CF01 & "' AND CF02='" & CF02 & "' AND CF03='" & CF03 & "'"
'   Else
'      strQty = "SELECT CF22,CF25,CF27 FROM CASEFEE WHERE CF01='" & CF01 & "' AND CF02='" & CF02 & "' AND CF03='" & CF03 & "'"
'   End If
'   '2005/7/7 END
'   RsTemp.Open strQty, cnnConnection
'   Do While Not RsTemp.EOF
'      If IsNull(RsTemp.Fields(0)) Or RsTemp.Fields(0) = 0 Then
'         If Not IsNull(RsTemp.Fields(1)) Then
'         '月
'            CFOther(1) = CompDate(1, RsTemp.Fields(1), CFOther(0))
'            If RsTemp.Fields("CF27") = "1" Then CFOther(1) = CompDate(2, -1, CFOther(1))
'
'            Select Case CF01
'               Case "CFT"
'                  CFOther(2) = CompDate(1, -1, CFOther(1))
'               Case "CFP"
'                  CFOther(2) = CompDate(2, -14, CFOther(1))
'               Case "T"
'                  If CF02 = "238" Then
'                     CFOther(2) = CompDate(1, -1, CFOther(1))
'                  Else
'                     If RsTemp.Fields(1) >= 2 Then
'                        CFOther(2) = CompDate(2, -4, CFOther(1))
'                     Else
'                        CFOther(2) = CompDate(2, -2, CFOther(1))
'                     End If
'                  End If
'               Case "P"
'                  If CF02 = "000" Then
'                     If RsTemp.Fields(1) >= 2 Then
'                        CFOther(2) = CompDate(2, -4, CFOther(1))
'                     Else
'                        CFOther(2) = CompDate(2, -2, CFOther(1))
'                     End If
'                  Else
'                     CFOther(2) = CompDate(2, -10, CFOther(1))
'                  End If
'               Case Else
'                  If RsTemp.Fields(1) >= 2 Then
'                     CFOther(2) = CompDate(2, -4, CFOther(1))
'                  Else
'                     CFOther(2) = CompDate(2, -2, CFOther(1))
'                  End If
'            End Select
'
'         End If
'      Else
'         If Not IsNull(RsTemp.Fields(0)) Then
'         '日
'            CFOther(1) = CompDate(2, RsTemp.Fields(0), CFOther(0))
'            If RsTemp.Fields("CF27") = "1" Then CFOther(1) = CompDate(2, -1, CFOther(1))
'
'            Select Case CF01
'               Case "CFT"
'                  CFOther(2) = CompDate(1, -1, CFOther(1))
'               Case "CFP"
'                  CFOther(2) = CompDate(2, -14, CFOther(1))
'               Case "T"
'                  If CF02 = "238" Then
'                     CFOther(2) = CompDate(1, -1, CFOther(1))
'                  Else
'                     If RsTemp.Fields(0) >= 60 Then
'                        CFOther(2) = CompDate(2, -4, CFOther(1))
'                     Else
'                        CFOther(2) = CompDate(2, -2, CFOther(1))
'                     End If
'                  End If
'               Case "P"
'                  If CF02 = "000" Then
'                     If RsTemp.Fields(0) >= 60 Then
'                        CFOther(2) = CompDate(2, -4, CFOther(1))
'                     Else
'                        CFOther(2) = CompDate(2, -2, CFOther(1))
'                     End If
'                  Else
'                     CFOther(2) = CompDate(2, -10, CFOther(1))
'                  End If
'               Case Else
'                  If RsTemp.Fields(0) >= 60 Then
'                     CFOther(2) = CompDate(2, -4, CFOther(1))
'                  Else
'                     CFOther(2) = CompDate(2, -2, CFOther(1))
'                  End If
'            End Select
'         End If
'      End If
'      GetCaseFeeDelay = True
'      Exit Do
'   Loop
'   RsTemp.Close
'   Exit Function
'ErrHnd:
'   MsgBox "錯誤 : " & Err.Description, vbCritical
'End Function
'2005/7/7 END


'Add by Morgan 2005/7/14
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   cp(110) = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
         
         'Modified by Morgan 2021/12/14f Forms2.0 改用模組
         'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   If bolCheck = True Then
      cp(22) = ""
   Else
      cp(22) = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub
'Added by Morgan 2012/9/25
Private Sub SetDatebyCaseFee(pCP10 As String)
   If pCP10 <> "404" Then
      '台灣再審延期以發文日計算(其他用原法定期限)
      'Modified by Morgan 2022/7/15 +(507)行政訴訟上訴，(508)行政訴訟上訴答辯--陳玲玲
      If pa(9) = "000" And (pCP10 = "107" Or pCP10 = "507" Or pCP10 = "508") Then
         strExc(0) = TransDate(Text7, 2)
      Else
         strExc(0) = TransDate(cp(7), 2)
      End If
      If Val(strExc(0)) > 0 Then
         GetCaseFee pa(1), pa(9), pCP10, strExc, pa(8)
      End If
   End If
End Sub

'Add by Morgan 2013/4/30
Private Sub StartLetter2(ByVal ET01 As String, ET02 As String, ByVal ET03 As String)
   Dim strTxt() As String, i As Integer

   EndLetter ET01, ET02, ET03, strUserNum

   i = 1
   ReDim Preserve strTxt(i)
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','延期後法定期限','" & DBDATE(Text6) & "')"
   
   i = i + 1
   ReDim Preserve strTxt(i)
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','延期後本所期限','" & DBDATE(Text5) & "')"
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If

End Sub
'Added by Morgan 2013/10/2 計算延期後期限
Private Sub SetNewDueDate(pCP43 As String, pCP10 As String)
   Dim stSQL As String, stDate As String
   Dim bolIs107 As Boolean
   Dim stCP134 As String 'Added by Morgan 2014/9/2
   Dim stNP07 As String 'Added by Morgan 2020/9/25
   
   '台灣非再審延期預設來函期限月數
   If pa(9) = "000" Then
       If pCP43 = "" Then Exit Sub
      
      'Added by Morgan 2014/9/2
      '再審延期
      'Modified by Morgan 2020/9/25 +stNP07
      If CheckIs107Extend(pCP43, stCP134, stNP07) = True Then
         '有延期過
         If CheckExtended("107") = True Then
            '再審第二次延期的期限,請再更改為再審申請延後的法定期限加二個月--陳玲玲 103/8/22
            Text6 = TransDate(CompDate("1", 2, cp(7)), 1)
            'Added by Morgan 2014/10/28
            If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
               Text5 = TransDate(PUB_GetOurDeadline(Text6), 1)
            Else
            'end 2014/10/28
               Text5 = TransDate(PUB_GetWorkDay1(CompDate("2", -2, Text6), True), 1)
            End If
         Else
            SetDatebyCaseFee "107"
         End If
      '非再審延期
      'Added by Morgan 2020/9/25
      '舉發答辯:原期限+2個月(casefee)
      ElseIf stNP07 = "804" Then
         SetDatebyCaseFee stNP07
      'end 2020/9/25
      
      '來函是否有期限月數
      ElseIf Val(stCP134) > 0 Then
         Text9 = stCP134
         Text9_Validate False
         
      Else
         SetDatebyCaseFee pCP10
      End If
      'end 2014/9/2
      
'Removed by Morgan 2014/9/2 改上面程式
'      '延期發文
'      If cp(10) = "404" Then
'         '未收文
'         If pCP43 > "C" Then
'            stSQL = "select cp134,np07,cp07 from nextprogress,caseprogress where np01(+)='" & pCP43 & "' and cp09(+)=np01"
'         '已收文
'         Else
'            stSQL = "select c2.cp134,c1.cp10,c2.cp07 from caseprogress c1,caseprogress c2 where c1.cp09='" & pCP43 & "' and c2.cp09(+)=c1.cp43"
'         End If
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
'         If intI = 1 Then
'            '再審
'            If RsTemp.Fields(1) = "107" Then
'               stDate = "" & RsTemp.Fields(2)
'               '若已有延期過則期限應為原來函法限+6個月
'               m_bolTW107Extended = CheckExtended("107")
'               If m_bolTW107Extended = True Then
'                  'Modified by Morgan 2014/9/1 改為再審延期後期限+2個月
'                  'If stDate <> "" Then
'                  '   Text6 = TransDate(CompDate("1", 6, stDate), 1)
'                  '   Text5 = TransDate(PUB_GetWorkDay1(CompDate("2", -2, Text6), True), 1)
'                  'End If
'                  Text6 = TransDate(CompDate("1", 2, cp(7)), 1)
'                  Text5 = TransDate(PUB_GetWorkDay1(CompDate("2", -2, Text6), True), 1)
'                  'end 2014/9/1
'               Else
'                  SetDatebyCaseFee "107"
'               End If
'            '非再審
'            ElseIf RsTemp.Fields(0) > 0 Then
'               Text9 = RsTemp.Fields(0)
'               Text9_Validate False
'            'ADD BY SONIA 2014/5/5 P-103985
'            Else
'               SetDatebyCaseFee "" & RsTemp.Fields(1)
'            '2014/5/5 end
'            End If
'         End If
'
'      '按延期按鈕
'      '再審
'      ElseIf cp(10) = "107" Then
'         stSQL = "select cp07 from caseprogress where cp09='" & pCP43 & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
'         If intI = 1 Then
'            stDate = "" & RsTemp.Fields(0)
'            '若已有延期過則期限應為原來函法限+6個月
'            m_bolTW107Extended = CheckExtended("107")
'            If m_bolTW107Extended = True Then
'               'Modified by Morgan 2014/9/1 改為再審延期後期限+2個月
'               'If stDate <> "" Then
'               '   Text6 = TransDate(CompDate("1", 6, stDate), 1)
'               '   Text5 = TransDate(PUB_GetWorkDay1(CompDate("2", -2, Text6), True), 1)
'               'End If
'               Text6 = TransDate(CompDate("1", 2, cp(7)), 1)
'               Text5 = TransDate(PUB_GetWorkDay1(CompDate("2", -2, Text6), True), 1)
'               'end 2014/9/1
'            Else
'               SetDatebyCaseFee "107"
'            End If
'         End If
'
'      '其他
'      Else
'         stSQL = "select cp134 from caseprogress where cp09='" & pCP43 & "' and cp134>0"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
'         If intI = 1 Then
'            Text9 = RsTemp.Fields(0)
'            Text9_Validate False
'         'ADD BY SONIA 2014/5/5 P-103985
'         Else
'            SetDatebyCaseFee pCP10
'         '2014/5/5 end
'         End If
'      End If
      
   Else
      SetDatebyCaseFee pCP10
   End If
End Sub
'Added by Morgan 2013/10/2 檢查是否延期過
Private Function CheckExtended(pCP10 As String) As Boolean
   Dim stSQL As String
   stSQL = "select cp09 from caseprogress a" & _
      " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='404' and cp27>0" & _
      " and (exists(select * from caseprogress b where a.cp43<'C' and b.cp09=a.cp43 and b.cp10='" & pCP10 & "')" & _
      " or exists(select * from nextprogress b where a.cp43>'C' and b.np01=a.cp43 and b.np07=" & pCP10 & "))"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      CheckExtended = True
   End If
End Function

'Added by Morgan 2014/9/2
'是否為再審延期
'Modified by Morgan 2020/9/25 +pNP07
Private Function CheckIs107Extend(ByVal pCP43 As String, Optional ByRef pCP134 As String, Optional ByRef pNP07 As String) As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   
   If cp(10) = "107" Then
      CheckIs107Extend = True
   ElseIf cp(10) = "404" Then
      '未收文
      If pCP43 > "C" Then
         stSQL = "select cp134,np07,cp07 from nextprogress,caseprogress where np01(+)='" & pCP43 & "' and cp09(+)=np01"
      '已收文
      Else
         stSQL = "select c2.cp134,c1.cp10 np07,c2.cp07 from caseprogress c1,caseprogress c2 where c1.cp09='" & pCP43 & "' and c2.cp09(+)=c1.cp43"
      End If
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         pNP07 = rsQuery("np07") 'Added by Morgan 2020/9/25
         If rsQuery("np07") = "107" Then
            CheckIs107Extend = True
         End If
         pCP134 = "" & rsQuery("cp134")
      End If
   End If
   
   Set rsQuery = Nothing
End Function
'Add by Lydia 2015/01/13
Private Sub txtCP118_GotFocus()
   TextInverse txtCP118
   CloseIme
End Sub

Private Sub txtCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

'設定期限
Private Sub SetDate()
   Dim i As Integer
   Dim stStartDate As String

   If Val(Text9) > 0 Then
      'Modify by Morgan 2011/9/20 非台灣要用原法定計算
      'Text6 = TransDate(CompDate("1", Val(Text9), Text7), 1)
      If pa(9) = "000" Then
         'Modify by Morgan 2012/7/31 台灣也改用原法定計算
         'Text6 = TransDate(CompDate("1", Val(Text9), Text7), 1)
         
         'Modified by Morgan 2014/9/2 台灣再審第1次延期要用發文日算
         'Text6 = TransDate(CompDate("1", Val(Text9), cp(7)), 1)
         stStartDate = cp(7)
         '再審延期
         If CheckIs107Extend(cp(43)) = True Then
            '第1次延期
            If CheckExtended("107") = False Then
               stStartDate = Text7
            End If
         End If
         Text6 = TransDate(CompDate("1", Val(Text9), stStartDate), 1)
         'end 2014/9/2
         
         'Modified by Morgan 2012/9/25
         '改2個月以上-4天否則-2天(同來函期限的規則)
         'strExc(1) = pa(1)
         'strExc(2) = pa(9)
         'strExc(3) = DBDATE(Text6)
         'GetCtrlDT strExc
         'Added by Morgan 2014/10/28
         If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
            strExc(0) = PUB_GetOurDeadline(Text6)
         Else
         'end 2014/10/28
            If Val(Text9) >= 2 Then
               strExc(0) = CompDate("2", -4, Text6)
            Else
               strExc(0) = CompDate("2", -2, Text6)
            End If
            'end 2012/9/25
         End If 'Added by Morgan 2014/10/28
         Text5 = TransDate(PUB_GetWorkDay1(strExc(0), True), 1)
      Else
         'Added by Morgan 2020/2/13
         '大陸延期後期限不可含在途，原法限要用來函發文日+月數計算
         If pa(9) = "020" Then
            If PUB_GetOADeadline(IIf(cp(10) = 延期, cp(43), cp(9)), strExc(1), False) = True Then
               Text6 = TransDate(CompDate("1", Val(Text9), strExc(1)), 1)
            End If
         Else
         'end 2020/2/13
            Text6 = TransDate(CompDate("1", Val(Text9), cp(7)), 1)
         End If 'Added by Morgan 2020/2/13
         
'Mofieid by Morgan 2015/10/23
'            'FMP 本所=法定-10天
'            If m_bolFMP Then
'               i = -10
'            '本所=法定-7天
'            Else
'               i = -7
'            End If
'規則改與一般來函一致
         'Added by Lydia 2025/10/29
         If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
            Text5 = TransDate(PUB_GetPOurDeadline(DBDATE(Text6), pa(9)), 1)
         Else
         'end 2025/10/29
            'FMP 本所=法定-7天
            If m_bolFMP Then
               i = -7
            '本所=法定-10天
            Else
               i = -10
            End If
   'end 2015/10/23
   
            Text5 = TransDate(PUB_GetWorkDay1(CompDate("2", i, Text6), True), 1)
         End If 'Added by Lydia 2025/10/29
      End If
   'Added by Morgan 2012/7/31
   Else
      'Modified by Morgan 2012/9/25 改呼叫函數
      'If cp(7) > 0 Then
      '   strExc(0) = TransDate(cp(7), 2)
      '   GetCaseFee pa(1), pa(9), cp(10), strExc, pa(8)
      'End If
      
      'Modified by Morgan 2013/10/2
      'SetDatebyCaseFee cp(10)
      'Modified by Morgan 2022/7/19
      'SetNewDueDate cp(43), cp(10)
      If cp(10) = "404" Then
         With MSHFlexGrid1
         For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "v" Then
               SetNewDueDate .TextMatrix(i, 6), .TextMatrix(i, 7)
               Exit For
            End If
         Next
         End With
      Else
         SetNewDueDate cp(43), cp(10)
      End If
      'end 2022/7/19
      'end 2013/10/2
      
      'end 2012/9/25
   End If
   'end 2012/7/31
End Sub
