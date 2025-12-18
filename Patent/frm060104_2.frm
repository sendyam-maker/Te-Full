VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文(延期)"
   ClientHeight    =   6300
   ClientLeft      =   96
   ClientTop       =   996
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9348
   Begin VB.TextBox txtEmail 
      Height          =   270
      Left            =   4350
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "Y"
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox txtRecDate 
      Height          =   270
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   2
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1620
      MaxLength       =   9
      TabIndex        =   37
      Top             =   2670
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   4650
      MaxLength       =   7
      TabIndex        =   36
      Top             =   2970
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   7710
      MaxLength       =   7
      TabIndex        =   35
      Top             =   2970
      Width           =   1215
   End
   Begin VB.TextBox txtCP84 
      Height          =   270
      Left            =   7710
      TabIndex        =   34
      Top             =   2670
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   1620
      MaxLength       =   7
      TabIndex        =   33
      Top             =   2970
      Width           =   1215
   End
   Begin VB.TextBox txtCP118 
      Height          =   270
      Left            =   1620
      MaxLength       =   1
      TabIndex        =   0
      Top             =   3255
      Width           =   375
   End
   Begin VB.TextBox txtTimes 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3690
      MaxLength       =   1
      TabIndex        =   1
      Top             =   3255
      Width           =   300
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "補件期限(&D)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   5040
      TabIndex        =   5
      Top             =   70
      Width           =   1200
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060104_2.frx":0000
      Left            =   1140
      List            =   "frm060104_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   18
      Top             =   870
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8316
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6270
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7092
      TabIndex        =   9
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   12
      Top             =   570
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   10
      Top             =   570
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2460
      MaxLength       =   1
      TabIndex        =   8
      Top             =   570
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   6
      Top             =   570
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1725
      Left            =   30
      TabIndex        =   4
      Top             =   4500
      Width           =   9255
      _ExtentX        =   16341
      _ExtentY        =   3048
      _Version        =   393216
      Cols            =   11
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
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "Email維護:             (Y:是)"
      Height          =   180
      Left            =   3420
      TabIndex        =   50
      Top             =   3645
      Width           =   1860
   End
   Begin VB.Label lblRecDate 
      AutoSize        =   -1  'True
      Caption         =   "當天報告:             (Y:是)"
      Height          =   180
      Left            =   180
      TabIndex        =   49
      Top             =   3645
      Width           =   1815
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "延期日："
      Height          =   180
      Left            =   180
      TabIndex        =   48
      Top             =   2700
      Width           =   720
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Label15"
      Height          =   180
      Left            =   4650
      TabIndex        =   47
      Top             =   2700
      Width           =   1320
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "上次延期日："
      Height          =   180
      Left            =   3420
      TabIndex        =   46
      Top             =   2700
      Width           =   1080
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "延期後本所期限："
      Height          =   180
      Index           =   0
      Left            =   3150
      TabIndex        =   45
      Top             =   3000
      Width           =   1440
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "延期後法定期限："
      Height          =   180
      Left            =   6210
      TabIndex        =   44
      Top             =   3000
      Width           =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   100
      X2              =   9200
      Y1              =   2630
      Y2              =   2630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   100
      X2              =   9200
      Y1              =   2600
      Y2              =   2600
   End
   Begin VB.Label lblCP84 
      AutoSize        =   -1  'True
      Caption         =   "發文規費:"
      Height          =   180
      Left            =   6885
      TabIndex        =   43
      Top             =   2700
      Width           =   765
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   6750
      TabIndex        =   42
      Top             =   3285
      Width           =   900
   End
   Begin VB.Label lblTimes 
      AutoSize        =   -1  'True
      Caption         =   "第          次延期"
      Height          =   180
      Left            =   3420
      TabIndex        =   41
      Top             =   3285
      Width           =   1170
   End
   Begin MSForms.Label Label2 
      Height          =   345
      Index           =   9
      Left            =   1140
      TabIndex        =   40
      Top             =   2190
      Width           =   8130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "14340;609"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7710
      TabIndex        =   39
      Top             =   3255
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;556"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "延期後約定期限："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   38
      Top             =   3000
      Width           =   1440
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   1140
      TabIndex        =   32
      Top             =   1170
      Width           =   1890
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3334;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   1140
      TabIndex        =   31
      Top             =   1500
      Width           =   1890
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3334;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   1140
      TabIndex        =   30
      Top             =   1830
      Width           =   1890
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3334;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   4
      Left            =   4260
      TabIndex        =   29
      Top             =   570
      Width           =   1890
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3334;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   5
      Left            =   4260
      TabIndex        =   28
      Top             =   1170
      Width           =   1890
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3334;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   6
      Left            =   4260
      TabIndex        =   27
      Top             =   1830
      Width           =   1890
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3334;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   195
      Left            =   1800
      TabIndex        =   26
      Top             =   900
      Width           =   7455
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "13150;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   25
      Top             =   2190
      Width           =   765
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:                (Y: 是)"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   24
      Top             =   3285
      Width           =   2355
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   5400
      TabIndex        =   23
      Top             =   930
      Visible         =   0   'False
      Width           =   45
   End
   Begin MSForms.Label Label20 
      Height          =   225
      Left            =   150
      TabIndex        =   22
      Top             =   4230
      Width           =   1095
      VariousPropertyBits=   27
      Caption         =   "欲延期期限:"
      Size            =   "1931;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Left            =   3420
      TabIndex        =   21
      Top             =   1830
      Width           =   765
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   180
      TabIndex        =   20
      Top             =   1830
      Width           =   765
   End
   Begin MSForms.Label Label29 
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   1500
      Width           =   765
      VariousPropertyBits=   27
      Caption         =   "案件性質:"
      Size            =   "1349;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "收文日期:"
      Height          =   180
      Left            =   3420
      TabIndex        =   17
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   180
      TabIndex        =   16
      Top             =   1170
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   15
      Top             =   570
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   3420
      TabIndex        =   14
      Top             =   570
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   180
      TabIndex        =   13
      Top             =   870
      Width           =   765
   End
End
Attribute VB_Name = "frm060104_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/15 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String
Dim Situ As Integer
'Modify by Morgan 2005/8/8 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, m_CP110 As String, m_AgentName As String
Dim cp(6 To 13) As String
Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer
Public StrSales1 As String
Public StrSales2 As String
' 案件性質
Public m_CP10 As String
' 國家代碼
Dim m_PA09 As String
'Add By Cheng 2002/06/20
Public m_str_DL05 As String '延期記錄檔之資料來源
'Add By Cheng 2002/08/19
Dim m_strDL06 As String
'92.3.13 ADD BY SONIA 延期案件性質
Dim Work_CP10 As String
'Add by Morgan 2004/8/11
Dim m_CP17 As String '收文規費
'Add by Morgan 2008/8/21
Dim m_ST03 As String '承辦人部門
Dim m_EP09 As String '完稿日
Dim m_CP48 As String '原承辦期限
Dim m_CP05 As String '收文日
Dim m_EP06 As String '文件齊備日
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/20 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim iDay As Integer, iMonth As Integer 'Add by Morgan 2011/2/8
Dim m_CP14 As String  'add by sonia 2015/10/2 承辦人
Dim strDtLimitRecNo As String 'Add By Sindy 2015/12/15
Dim m_404CP09 As String 'Added by Morgan 2016/11/1 延期收文號
Dim m_CP118 As String 'Added by Lydia 2018/0/1/11 電子送件
Dim m_CP82 As String 'Added by Lydia 2018/06/20 發文時間
Dim m_pAgreeOnDate As String 'Modify By Sindy 2021/4/23
'Added by Lydia 2023/03/17
Dim m_LOS15 As String '法律所案源單號
Dim m_LOS02 As String '法律所案源類別
Dim m_CP43 As String

Private Function FormSave() As Boolean
   Dim strTmp(0 To 4) As String, bolChk As Boolean
   Dim i As Integer, strUpdate As String
   Dim strCP09 As String
 
   If Text7 = "" Then MsgBox "延期日不得為空值 !", vbCritical: Exit Function
   If ChkRange(Text5, Text6, "本所期限、法定期限") = False Then Exit Function
   
 On Error GoTo CheckingErr
 
   cnnConnection.BeginTrans

   Select Case Situ
      Case 0 '前畫面按 延期 鈕
         strExc(1) = "DELETE FROM DATELIMIT WHERE DL01='" & strReceiveNo & "' AND DL02=" & TransDate(Text7, 2)
         cnnConnection.Execute strExc(1)
   
         strExc(2) = "INSERT INTO DATELIMIT (DL01,DL02,DL03,DL04,DL05,DL06) VALUES " & _
            "('" & strReceiveNo & "'," & TransDate(Text7, 2) & "," & CNULL(cp(6)) & "," & CNULL(cp(7)) & "," & CNULL(m_str_DL05) & ",'" & IIf(m_str_DL05 = "1", "", m_strDL06) & "' )"
         
         cnnConnection.Execute strExc(2)

         strCP09 = AutoNo("B", 6)
         
         'Modify by morgan 2005/8/8 加 cp110
         'Modify By Sindy 2016/8/18 若有發文規費時, 存檔更新進度檔時同時更新 CP20及CP32 為NULL(即要向客戶請款)
         'Modified by Lydia 2018/01/11 +CP64,CP118
         strExc(3) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
            "CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43,CP84,CP110,CP64,CP118) VALUES " & _
            "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & _
            strSrvDate(1) & "," & CNULL(cp(6)) & "," & CNULL(cp(7)) & _
            ",'" & strCP09 & "','" & 延期 & "'," & CNULL(StrSales1) & "," & CNULL(StrSales2) & _
            ",'" & strUserNum & "'," & IIf(Val(txtCP84.Text) > 0, "null", CNULL("N")) & _
            ",'N'," & IIf(Val(txtCP84.Text) > 0, "null", CNULL("N")) & "," & TransDate(Text7, 2) & ",'" & strReceiveNo & "'," & Format(Val(txtCP84.Text)) & "," & CNULL(m_CP110) & _
            "," & CNULL(ChgSQL(Label2(9))) & " ," & CNULL(txtCP118.Text) & ")"
         cnnConnection.Execute strExc(3)
         
         m_404CP09 = strCP09 'Added by Morgan 2016/11/1
         
         'Add By Sindy 2015/12/15 205申復,107再審
         If m_CP10 = "205" Or m_CP10 = "107" Then
            strDtLimitRecNo = strReceiveNo
         End If
         '2015/12/15 END
         
         'Add by Morgan 2008/8/29 重算承辦期限
         strExc(1) = "" 'Added by Morgan 2017/3/2
         strUpdate = ""
         '檢視中說,以齊備日計算
         'Modify by Morgan 2008/9/10 +210 製作中說
         'Modified by Morgan 2013/11/6 +235核對中說格式
         If (m_CP10 = "209" Or m_CP10 = "235" Or m_CP10 = "210") And m_EP06 <> "" Then
            strExc(1) = Pub_GetHandleDay(pa(1), pa(9), m_CP10, DBDATE(m_EP06), DBDATE(Text5))
         '其他(非例外),以收文日計算
         ElseIf InStr(SkipCasePtyList, m_CP10) = 0 Then
            strExc(1) = Pub_GetHandleDay(pa(1), pa(9), m_CP10, m_CP05, DBDATE(Text5))
         End If
         If strExc(1) <> "" And m_CP48 <> strExc(1) Then
            'Add by Morgan 2008/9/5
            '與原承辦期限不同者另加2個工作天(不含當日故函數應傳3天)
            strExc(1) = CompWorkDay(3, strExc(1))
            strUpdate = ",CP48=" & CNULL(strExc(1), True)
         End If
         'end 2008/8/29
         
         strExc(4) = "UPDATE CASEPROGRESS SET CP06=" & TransDate(Text5, 2) & _
            ",CP07=" & TransDate(Text6, 2) & strUpdate & " WHERE CP09='" & strReceiveNo & "'"
      
         cnnConnection.Execute strExc(4)
          
         'Add by Morgan 2008/8/21
         '翻譯延期發文時重算核稿期限
         If m_CP10 = "201" And m_ST03 <> "" And m_EP09 <> "" Then
            '外翻：核稿承辦期限=完稿日+4週
            If m_ST03 = "F51" Then
                strExc(1) = CompDate(2, 28, m_EP09)
            '內翻：核稿承辦期限=完稿日+10天
            Else
                strExc(1) = CompDate(2, 10, m_EP09)
            End If
            strSql = "update engineerprogress set ep08=" & strExc(1) & " where ep02='" & strReceiveNo & "'"
            cnnConnection.Execute strSql, intI
         End If
         
         'Added by Morgan 2012/11/21
         '102新法,若 201,209,210 延期時有主動修正未發文則將新的期限更新至該程序並設定承辦期限為本所期限
         'Modified by Morgan 2013/11/6 +235核對中說格式
         If m_CP10 = "201" Or m_CP10 = "209" Or m_CP10 = "235" Or m_CP10 = "210" Then
            'Modified by Morgan 2013/1/3 改只要更新承辦期限為本程序的本所期限
            'If strSrvDate(1) >= "20130101" Then
            '   strSql = "update caseprogress a set (cp06,cp07,cp48)=(select b.cp06,b.cp07,b.cp06 from caseprogress b where b.cp09='" & strReceiveNo & "')" & _
            '      " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='203' and cp27||cp57 is null"
            '   cnnConnection.Execute strSql, intI
            'End If
            strSql = "update caseprogress a set cp48=" & DBDATE(Text5) & _
               " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='203' and cp27||cp57 is null"
            cnnConnection.Execute strSql, intI
            'end 2013/1/3
         End If
         'end 2012/11/21
         
         m_CP09s = strCP09 & m_CP09s 'Add by Morgan 2009/3/20 B類延期要補收文號
         
      Case 1 '延期
         For i = 1 To MSHFlexGrid1.Rows - 1
            If MSHFlexGrid1.TextMatrix(i, 0) = "v" Then
               bolChk = True
               strTmp(2) = TransDate(Replace(MSHFlexGrid1.TextMatrix(i, 2), "/", ""), 2)
               strTmp(3) = TransDate(Replace(MSHFlexGrid1.TextMatrix(i, 3), "/", ""), 2)
               strTmp(1) = MSHFlexGrid1.TextMatrix(i, 6)
               strTmp(0) = MSHFlexGrid1.TextMatrix(i, 7)
               strTmp(4) = MSHFlexGrid1.TextMatrix(i, 8) 'Add By Sindy 2016/1/13 案件性質代碼
               Exit For
            End If
         Next
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
            Exit Function
         End If
         strExc(1) = "DELETE FROM DATELIMIT WHERE DL01='" & strTmp(1) & "' AND DL02=" & TransDate(Text7, 2)
         cnnConnection.Execute strExc(1)
   
         strExc(2) = "INSERT INTO DATELIMIT (DL01,DL02,DL03,DL04,DL05,DL06) VALUES " & _
            "('" & strTmp(1) & "'," & TransDate(Text7, 2) & "," & strTmp(2) & "," & strTmp(3) & "," & CNULL(m_str_DL05) & ",'" & IIf(m_str_DL05 = "1", "", m_strDL06) & "' )"
        
         cnnConnection.Execute strExc(2)
         
         'Modify by morgan 2005/8/8 加 cp110
         'Modify by Morgan 2011/4/22 +CP30
         'modify by sonia 2015/10/2 承辦人為外專程序時,改為操作人員
         'strExc(3) = "UPDATE CASEPROGRESS SET CP27=" & TransDate(Text7, 2) & ", CP84=" & Format(Val(txtCP84.Text)) & _
         '   ",CP43='" & strTmp(1) & "',cp110=" & CNULL(m_CP110) & ",CP22=NULL,CP30='" & IIf(strTmp(0) = "0", "", strTmp(0)) & "' WHERE CP09='" & strReceiveNo & "'"
'Modify By Sindy 2024/3/4 發文時承辦人改為發文人員(操作人員)
         'm_CP14 = GetFCPUser(m_CP14)
         m_CP14 = strUserNum
'2024/3/4 END
        'Modified by Lydia 2018/01/11 +CP64,CP118
        'Modified by Lydia 2018/08/29 因為出各式申請書-電子送件會上CP118='A' ,保持原設定
         'strExc(3) = "UPDATE CASEPROGRESS SET CP14='" & m_CP14 & "',CP27=" & TransDate(Text7, 2) & ", CP84=" & Format(Val(txtCP84.Text)) & _
            ",CP43='" & strTmp(1) & "',cp110=" & CNULL(m_CP110) & ",CP22=NULL,CP30='" & IIf(strTmp(0) = "0", "", strTmp(0)) & "' " & _
            ",CP64=" & CNULL(ChgSQL(Label2(9))) & ",CP118=" & CNULL(txtCP118.Text) & _
            " WHERE CP09='" & strReceiveNo & "'"
         If m_CP118 <> "" And txtCP118 <> "" Then
             strExc(1) = ""
         Else
             strExc(1) = ",CP118=" & CNULL(txtCP118.Text)
         End If
         strExc(3) = "UPDATE CASEPROGRESS SET CP14='" & m_CP14 & "',CP27=" & TransDate(Text7, 2) & ", CP84=" & Format(Val(txtCP84.Text)) & _
            ",CP43='" & strTmp(1) & "',cp110=" & CNULL(m_CP110) & ",CP22=NULL,CP30='" & IIf(strTmp(0) = "0", "", strTmp(0)) & "' " & _
            ",CP64=" & CNULL(ChgSQL(Label2(9))) & strExc(1) & _
            " WHERE CP09='" & strReceiveNo & "'"
         'end 2015/10/2
         'end 2018/08/29
         cnnConnection.Execute strExc(3)
         
         m_404CP09 = strReceiveNo 'Added by Morgan 2016/11/1
         
         'Add By Sindy 2015/12/15 205申復,107再審
         If strTmp(4) = "205" Or strTmp(4) = "107" Then
            strDtLimitRecNo = strTmp(1)
         End If
         '2015/12/15 END
         
         '92.10.23 MODIFY BY SONIA 改為複選
         'strExc(4) = "UPDATE NEXTPROGRESS SET NP08=" & TransDate(Text5, 2) & ",NP09=" & TransDate(Text6, 2) & " WHERE NP22=" & strTmp(0)
          '911105 nick transation
         ' cnnConnection.Execute strExc(4)
         For i = 1 To MSHFlexGrid1.Rows - 1
            If MSHFlexGrid1.TextMatrix(i, 0) = "v" Then
               bolChk = True
               strTmp(2) = TransDate(Replace(MSHFlexGrid1.TextMatrix(i, 2), "/", ""), 2)
               strTmp(3) = TransDate(Replace(MSHFlexGrid1.TextMatrix(i, 3), "/", ""), 2)
               strTmp(1) = MSHFlexGrid1.TextMatrix(i, 6)
               strTmp(0) = MSHFlexGrid1.TextMatrix(i, 7)
               'Modify by Morgan 2009/12/9 +更新CP期限
               If Val(strTmp(0)) > 0 Then
                  'Modify by Morgan 2006/1/24 加NP01
                  'Modify By Sindy 2021/4/23 + ,NP23=" & TransDate(Text8, 2):約定期限
                  strExc(4) = "UPDATE NEXTPROGRESS SET NP08=" & TransDate(Text5, 2) & ",NP09=" & TransDate(Text6, 2) & ",NP23=" & CNULL(TransDate(Text8, 2)) & " WHERE NP22=" & strTmp(0) & " and np01='" & strTmp(1) & "'"
               Else
                  strExc(4) = "UPDATE CASEPROGRESS SET CP06=" & TransDate(Text5, 2) & ", CP07='" & TransDate(Text6, 2) & "' WHERE CP09='" & strTmp(1) & "'"
                  
                  'Added by Morgan 2013/1/3 102新法,若 201,209,210 延期時有主動修正未發文則將新的本所期限更新為該程序的承辦期限
                  'Modified by Morgan 2013/11/6 +235核對中說格式
                  strExc(0) = MSHFlexGrid1.TextMatrix(i, 8)
                  If strExc(0) = "201" Or strExc(0) = "209" Or strExc(0) = "235" Or strExc(0) = "210" Then
                     strSql = "update caseprogress a set cp48=" & DBDATE(Text5) & _
                        " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='203' and cp27||cp57 is null"
                     cnnConnection.Execute strSql, intI
                  End If
                  'end 2013/1/3
               End If
               
               cnnConnection.Execute "begin user_data.user_notrigger:=1; end;" 'Add by Morgan 2010/7/13 +控制來函期限通知的 Trigger 不被觸發
               cnnConnection.Execute strExc(4), intI
               cnnConnection.Execute "begin user_data.user_notrigger:=0; end;" 'Add by Morgan 2010/7/13 +控制來函期限通知的 Trigger 不被觸發
            End If
         Next
   End Select
   
   PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130 'Add by Morgan 2009/3/20
   
   'Added by Lydia 2018/11/09 FCP中說延期發文確定則彈訊息及自動設行事曆
   If InStr("201,209,210", m_CP10) > 0 And pa(1) = "FCP" Then
       'Y54339000 (METIS IP LLC) , Y54339B10 (Metis IP (Beijing) LLC), Y54339B20 (Metis IP (Suzhou) LLC)
       If Left(pa(75), 6) = "Y54339" And Val(Text6.Text) > 0 Then
            '期限:抓中說法限前1個月
            strExc(1) = CompWorkDay(1, CompDate(1, -1, TransDate(Text6.Text, 2)), 1)
            strExc(2) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) '程序管制人
            strExc(3) = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) '承辦智權人員
            strExc(5) = strExc(2) & IIf(strExc(3) <> "", "," & strExc(3), "") '提醒
            strExc(4) = "管制中說最終法限前1個月催客戶"
            '提醒人員請掛管制人及承辦組人員，解除人員請掛管制人、承辦組人員及其案件職代(模組有預設抓案件職代)
            PUB_AddFCPStaffCalendar strExc(1), "1", strExc(5), strExc(4), strExc(5), "1", pa(1), pa(2), pa(3), pa(4)
       End If
   End If
   'end 2018/11/09
   'Added by Lydia 2023/03/17 若延期的相關總收文號為B2案源時，同時新增法務案之內部收文39延期
   If m_LOS15 <> "" And m_LOS02 = "B2" Then
       Call PUB_InsertLosBCP(m_LOS15, DBDATE(Text7), DBDATE(Text5), DBDATE(Text6))
   End If
   'end 2023/03/17
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
CheckingErr:
   FormSave = False
   cnnConnection.RollbackTrans
End Function

Private Sub cmdok_Click(Index As Integer)
Dim strTmp As String
Dim strTo As String, strSubject As String, strContent As String 'Add By Sindy 2015/12/15
Dim strTemp As String 'Add By Sindy 2015/12/14
Dim strFilePath As String 'Added by Lydia 2018/06/20 記錄智慧局收文文號
Dim bolUp As Boolean 'Added by Lydia 2018/08/17 是否上傳檔案到卷宗區

   Select Case Index
      Case 0
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
        
        'Added by Lydia 2018/01/11 延期發文開放可電子送件
         m_CP09s = "": m_CP123s = ""
        If txtCP118 = "Y" Then
            '電子送件也要記錄主管機關
            If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text7, , True) = False Then
               Exit Sub
            End If
        Else
        'end 2018/01/11
            'Add by Morgan 2009/4/28
            If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text7) = False Then
               Exit Sub
            End If
            If m_CP123s = "Y" Then
            'end 2009/4/28
               'Add by Morgan 2009/3/20 設定是否算發文室案件
               'modify by sonia 2014/6/23 加傳發文規費, P-108903
               If ModifyDispatch(strReceiveNo, m_CP09s, m_CP123s, txtCP84, Text7, IIf(Situ = 0, True, False)) = False Then
                   Exit Sub
               End If
               'end 2009/3/20
            End If
        End If 'end 2018/01/11
         
         'Add by Lydia 2018/01/11 延期發文開放可電子送件並必須輸入官方收文號
         If txtCP118 = "Y" And Val(txtCP84) = 0 Then
             m_CP123s = ""
             strExc(0) = InputBox("請輸入智慧局收文文號!!")
             If strExc(0) = "" Then
                Exit Sub
             Else
                strFilePath = strExc(0)  'Added by Lydia 2018/06/20 記錄智慧局收文文號
                'Modified by Lydia 2019/10/28 保留進度備註
                'Label2(9) = "智慧局收文文號:" & strExc(0) & ";" & Label2(9) 'CP64進度備註
                Label2(9).Caption = "智慧局收文文號:" & strExc(0) & ";" & Label2(9).Tag
             End If
          End If
          'end 2018/01/11

         'Added by Lydia 2018/06/20 檢查是否有電子送件的檔案; 重新發文(CP82>0)不做搬檔
         'Modified by Lydia 2018/08/17 重新發文要詢問(比照FCT發文自動上傳檔案)
         'If txtCP118.Text = "Y" And strFilePath <> "" And Val(m_CP82) = 0 Then
         bolUp = False
         If txtCP118.Text = "Y" And strFilePath <> "" Then
             strExc(1) = m_CP82
             If Val(m_CP82) > 0 Then
                 If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                      strExc(1) = ""
                 End If
             End If
             If Val(strExc(1)) = 0 Then
         'end 2018/08/17
                'Modified by Lydia 2019/03/22 +傳入發文日
                'Modified by Lydia 2019/10/28 +傳入本所案號
                'If Pub_AutoEsetToCpp(True, "", "", "", "", "", "", "", strFilePath, Text7.Text) = False Then
                If Pub_AutoEsetToCpp(True, pa(1), pa(2), pa(3), pa(4), pa(8), IIf(m_CP10 <> "404", "", strReceiveNo), "404", strFilePath, Text7.Text) = False Then
                       Exit Sub
                End If
                bolUp = True 'Added by Lydia 2018/08/17
             End If
         End If
         'end 2018/06/20
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
                 
         'Add by Morgan 2010/11/15
         strUserNum = strFMPNum
         'Modified by Morgan 2016/11/1 申復再審定稿分開
         'StartLetter2 "02", pa(1) & pa(2) & pa(3) & pa(4) & "&404", "01"
         'NowPrint pa(1) & pa(2) & pa(3) & pa(4) & "&404", "02", "01", False, strUserNum, 0
         If Work_CP10 = "107" Then
            'Added by Morgan 2017/10/23 +第2次延期後報告
            If txtTimes = "2" Then
               strTmp = "04"
            Else
            'end 2017/10/23
               strTmp = "03"
            End If
         Else
            strTmp = "02"
         End If
         StartLetter2 "02", m_404CP09, strTmp
         NowPrint m_404CP09, "02", strTmp, False, strUserNum, 0
         'end 2016/11/1
         strUserNum = strUser1Num
         'end 2010/11/15
         
         'Add by Morgan 2008/2/20 檢查代理人Email
         PUB_CheckEMail pa(75), pa(144)
         If pa(145) <> "" Then
            PUB_CheckEMail pa(75), pa(145)
         End If
         'end 2008/2/20
         
         'Add By Sindy 2015/12/14 相關總收文號為機關來函且未發文時,發mail通知承辦工程師及主管
         If strDtLimitRecNo <> "" Then
            If Situ = 0 Then '按鈕
               strExc(0) = "SELECT cp43" & _
                           " FROM CaseProgress" & _
                           " WHERE CP09 = '" & strDtLimitRecNo & "'" & _
                           " and CP43 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strDtLimitRecNo = RsTemp.Fields("cp43")
               End If
            End If
            strExc(0) = "SELECT cp10,st15,st52,cp14,decode('" & pa(9) & "','000',cpm03,cpm04) cp10nm" & _
                        " FROM CaseProgress,staff,casepropertymap" & _
                        " WHERE CP09 = '" & strDtLimitRecNo & "'" & _
                        " and CP10 in('1202','1002')" & _
                        " and CP27 is null and cp57 is null" & _
                        " and cp14=st01(+)" & _
                        " and cp01=cpm01(+) And cp10=cpm02(+)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strTo = ""
               '工程師
               If "" & RsTemp.Fields("st15") = "F21" Then
                  strTo = RsTemp.Fields("cp14") & ";"
                  If pa(150) = "" Then
                     If InStr(strTo, Pub_GetSpecMan("N")) = 0 Then strTo = strTo & Pub_GetSpecMan("N")
                  Else
                     'Modified by Lydia 2019/01/09
                     'strTemp = IIf(pa(150) = "1", Pub_GetSpecMan("T"), IIf(pa(150) = "2", Pub_GetSpecMan("R"), IIf(pa(150) = "3", Pub_GetSpecMan("S"), Pub_GetSpecMan("T1"))))
                     strTemp = Pub_GetFCPGrpMan(pa(150))
                    
                     If InStr(strTo, strTemp) = 0 Then strTo = strTo & strTemp
                  End If
               '程序人員
               Else
                  strTo = RsTemp.Fields("cp14") & ";"
                  If "" & RsTemp.Fields("st52") <> "" Then
                     If InStr(strTo, RsTemp.Fields("st52")) = 0 Then strTo = strTo & RsTemp.Fields("st52")
                  End If
               End If
               strSubject = "已屆期限，但OA尚未發文"
               strContent = "本所案號：" + pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) + vbCrLf + _
                            "案件性質：" + RsTemp.Fields("cp10nm") + vbCrLf + vbCrLf + _
                            "*本程序請儘速通知代理人/客戶" + vbCrLf
               PUB_SendMail strUserNum, strTo, "", strSubject, strContent
            End If
         End If
         '2015/12/14 END
               
         'Added by Lydia 2018/11/09 FCP中說延期發文確定則彈訊息及自動設行事曆
         If InStr("201,209,210", m_CP10) > 0 And pa(1) = "FCP" Then
            If InStr("Y54339000,Y54339B10,Y54339B20", ChangeCustomerL(pa(75))) > 0 And Val(Text6.Text) > 0 Then
                 MsgBox "已自動設行事曆，管制最終法限前1個月催客戶 !", vbInformation
            End If
         End If
         'end 2018/11/09
         
         'Added by Lydia 2018/06/20 發文時，電子送件自動上傳檔案到卷宗區; 重新發文(CP82>0)不做搬檔
                                                 '因為有前一畫面按延期進來的,所以存檔後才上傳檔案
         'Modified by Lydia 2018/08/17 是否上傳檔案,前面已判斷
         'If txtCP118.Text = "Y" And strFilePath <> "" And Val(m_CP82) = 0 Then
         If bolUp = True Then
             If Pub_AutoEsetToCpp(False, pa(1), pa(2), pa(3), pa(4), pa(8), m_404CP09, 延期, strFilePath) = False Then
                    Exit Sub
             End If
         End If
         'end 2018/06/20
         
         'Add By Cheng 2002/04/30
         If cp(10) = 延期 Then
            'Add By Sindy 2023/11/9
            If frm060104_1.bolIsEMPFlow = True Then
               frm090202_4.QueryData
            End If
            '2023/11/9 End
            '若有未發文資料顯示警告
            'Modify By Sindy 2023/11/9
            If PUB_GetCPunIssueDatas("" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)) Then
               frm060104_1.Show
               frm060104_1.ReQuery
            Else
               'Add By Sindy 2023/11/9
               If frm060104_1.bolIsEMPFlow = True Then
                  Unload frm060104_1
               Else
               '2023/11/9 End
                  frm060104_1.Show
                  frm060104_1.Clear
               End If
            End If
         Else
            frm060104_1.Show
            ' 90.08.06 modify by louis
            frm060104_1.Clear
         End If
         
         'Add By Sindy 2022/5/12
         If txtEmail.Text = "Y" Then
            frm060104_k.m_CP09 = m_404CP09 'strReceiveNo 'cp(9)
            frm060104_k.m_strRecDate = txtRecDate
            frm060104_k.Hide
            frm060104_k.cmdOK(0) = 1
            Unload frm060104_k
         End If
         '2022/5/12 END
      Case 1
         frm060104_1.Show
      Case 2
         Unload frm060104_1
      Case 3
         'Modify by Morgan 2006/5/1 改寫
         '延期
         If cp(10) = "404" Then
            If Work_CP10 = "" Then
               MsgBox "未點選延期案件性質 !", vbCritical
               Exit Sub
            End If
         End If
         
         strExc(0) = Name
         strExc(1) = pa(1)
         strExc(2) = pa(2)
         strExc(3) = pa(3)
         strExc(4) = pa(4)
         If cp(10) = "404" Then
            strExc(5) = Work_CP10 '案件性質
         Else
            strExc(5) = cp(10) '案件性質
         End If
         strExc(6) = strReceiveNo '收文號
         strExc(7) = Empty '發文日
         
         Me.Hide
         frm060104_4.Show
         'Modify By Sindy 2021/4/27 + m_pAgreeOnDate
         frm060104_4.SetData Text5, Text6, m_pAgreeOnDate
         
         Exit Sub
   End Select
   Unload Me
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         lblFM2 = pa(5)
      Case "英"
         lblFM2 = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         lblFM2 = pa(7)
   End Select
End Sub

Private Sub Form_Activate()
   Static bolActivated As Boolean
   
   'Added by Morgan 2015/10/6
   If Not bolActivated Then
      bolActivated = True
      If m_CP10 = "201" And Val(txtTimes) > 1 Then
         MsgBox "本案翻譯已延期過不得再延期！", vbExclamation
         'cmdOK(1).Value = True '會發生執行階段錯誤
      End If
   End If
   'end 2015/10/6
   
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060104_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      Situ = Val(Right(.Tag, 1))
      strReceiveNo = Left(.Tag, Len(.Tag) - 1)
   End With
   'Add by Morgan 2005/8/8
   ReDim pa(TF_PA)
   ReadPatent
   'Add by Morgan 2005/8/8
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True
   'Added by Sindy 2021/04/23 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1300
   lstNameAgent.Width = 1300
   
   Label2(1) = strReceiveNo
   Text7 = strSrvDate(2)
   Work_CP10 = ""
   ' 90.06.27 modify by louis
   If m_CP10 <> "404" And IsEmptyText(Text7) = False Then
      Work_CP10 = m_CP10
      
      'Removed by Morgan 2013/8/16 改都顯示
      ''Added by Morgan 2013/1/18
      'If Work_CP10 = "107" Then
      '   txtTimes.Visible = True
      '   lblTimes.Visible = True
      'Else
      '   txtTimes.Visible = False
      '   lblTimes.Visible = False
      'End If
      ''end 2013/1/18
      'end 2013/8/16
      
      CaculateNP08NP09
   End If
   'Add By Cheng 2003/01/17
   '記錄原延期後本所期限及延期後法定期限
   Me.Text5.Tag = Me.Text5.Text
   Me.Text6.Tag = Me.Text6.Text
   Me.Text8.Tag = Me.Text8.Text 'Add By Sindy 2021/5/7 約定期限
   
   'Add By Sindy 2021/5/7
   If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
      Label18(1).Visible = True
      Text8.Visible = True
   Else
      Label18(1).Visible = False
      Text8.Visible = False
   End If
   '2021/5/7 END
End Sub

'Add By Sindy 2021/4/23
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2023/03/17
   DestroyToolTip '清除物件 'Add By Sindy 2021/4/23
   Set frm060104_2 = Nothing
End Sub

Private Sub ReadPatent()
 Dim rsTemp1 As New ADODB.Recordset
 Dim Lbl As Object
 
   For Each Lbl In Label2
      Lbl = ""
   Next
   Label15 = ""
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   Select Case pa(1)
      Case "FCP"
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.ReadPatentDatabase(pA(), intWhere) Then FormShow
         If ClsPDReadPatentDatabase(pa(), intWhere) Then FormShow
      Case "FG"
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.ReadServicePracticeDatabase(pA(), intWhere) Then FormShow
         If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then FormShow
   End Select
   strExc(0) = "select count(*) from datelimit where dl01='" & strReceiveNo & "'"
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   ' 國家代碼
   m_PA09 = pa(9)
   If RsTemp.Fields(0) > 0 Then
      strExc(0) = "select max(dl02) from datelimit where dl01='" & strReceiveNo & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then Label15 = TransDate(RsTemp.Fields(0), 1)
   End If
   
   'modify by sonia 2015/10/2 +cp14
   'Modified by Lydia 2018/01/11 +CP64,CP118
   'Modified by Lydia 2018/06/20 +CP82
   'Modified by Lydia 2023/03/17 +CP43
   strExc(0) = "select cp05,cp06,cp07,cpm03,cp10,cp13, cp17,CP110,st03,ep09,cp48,ep06,cp14,CP64,CP118,CP82,CP43 " & _
                     "from caseprogress,casepropertymap,staff,engineerprogress " & _
                     "where cp09='" & strReceiveNo & "' AND cp01=cpm01(+) and cp10=cpm02(+) and st01(+)=cp14 and ep02(+)=cp09"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_EP06 = "" & .Fields("ep06")
      'Add by Morgan 2008/8/29
      m_CP05 = "" & .Fields("cp05")
      m_CP48 = "" & .Fields("cp48")
      'Add by Morgan 2008/8/21
      m_ST03 = "" & .Fields("st03")
      m_EP09 = "" & .Fields("ep09")
      'end 2008/8/21
      m_CP110 = "" & .Fields("CP110")
      'Add by Morgan 2004/8/12
      m_CP17 = "" & .Fields("cp17")
      'add by sonia 2015/10/2
      m_CP14 = "" & .Fields("cp14")
      
      'Modified by Morgan 2015/7/24
      'txtCP84.Text = m_CP17
      '104/8/3起延期發文不再繳規費以免客戶不辦又要退費 --靜芳
      If strSrvDate(1) >= "20150803" Then
         txtCP84.Text = "0"
         txtCP84.Enabled = False
      Else
         txtCP84.Text = m_CP17
      End If
      'end 2015/7/24
      
      'Add by Lydia 2018/01/11 延期發文開放可電子送件並必須輸入官方收文號
      m_CP118 = "" & .Fields("CP118")
      'Modified by Lydia 2018/08/29 延期沒有費用(ex.FCP-51053的延期申復的CP118=A )
      'txtCP118.Text = m_CP118
      If m_CP118 <> "" Then txtCP118 = "Y"
      Label2(9).Caption = "" & .Fields("CP64")
      'end 2018/01/11
      Label2(9).Tag = "" & .Fields("cp64") 'Added by Lydia 2019/10/28 保留進度備註
      
      txtCP84.Tag = txtCP84.Text
      m_CP82 = "" & .Fields("CP82") 'Added by Lydia 2018/06/20 發文時間
      m_CP43 = "" & .Fields("CP43") 'Added by Lydia 2023/03/17 相關收文號
      
      'Add By Cheng 2002/07/17
      m_CP10 = ""
      ' 90.06.27 modify by louis 案件性質
      If Not IsNull(.Fields("cp10")) Then
         m_CP10 = .Fields("cp10")
      End If
      If Not IsNull(.Fields(1)) Then cp(6) = .Fields(1)
      If Not IsNull(.Fields(2)) Then cp(7) = .Fields(2)
      'Add By Cheng 2002/08/19
      m_strDL06 = Empty
      
      If Situ = 0 Then '按鈕
         Label2(3) = " "
         Label2(6) = " "
         'MODIFY BY SONIA 90.9.29
         If Not IsNull(.Fields(1)) Then Label2(3) = ChangeTStringToTDateString(TransDate(.Fields(1), 1))
         If Not IsNull(.Fields(2)) Then Label2(6) = ChangeTStringToTDateString(TransDate(.Fields(2), 1))
         
         If Not IsNull(.Fields(1)) Then Text5 = TransDate(.Fields(1), 1)
         If Not IsNull(.Fields(2)) Then Text6 = TransDate(.Fields(2), 1)
'         strExc(0) = "select '',cpm03," & SQLDate("NP08") & "," & SQLDate("NP09") & _
'            ",NP13,NP14,NP01,NP22,NP07 from nextprogress,casepropertymap where NP22=-1" & _
'            " and np02=CPM01(+) and np07=cpm02(+)"
         strExc(0) = "select '',cpm03," & SQLDate("NP08") & "," & SQLDate("NP09") & _
            ",NP13,NP14,NP01,NP22,NP07,NP15," & SQLDate("NP23") & " from nextprogress,casepropertymap where NP22=-1" & _
            " and np02=CPM01(+) and np07=cpm02(+)"
            
'CANCEL BY SONIA 2013/7/24 移至CaculateNP08NP09,否則後面會被覆蓋,FCP-034425
'         'Added by Morgan2013/6/19
'         If m_CP10 = "107" Then
'            CheckExtended strReceiveNo, m_CP10
'         End If
'         'end 2013/6/19
         
      Else '延期
         If Not IsNull(.Fields(1)) Then Label2(3) = TransDate(.Fields(1), 1)
         If Not IsNull(.Fields(2)) Then Label2(6) = TransDate(.Fields(2), 1)
         '92.10.23 MODIFY BY SONIA
         'strExc(0) = "select '',cpm03," & SQLDate("NP08") & "," & SQLDate("NP09") & _
         '   ",NP13,NP14,NP01,NP22,NP07 from nextprogress,casepropertymap where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         '   " and (np06 is null or np06='') and np02=CPM01(+) and np07=cpm02(+)"
         'Modify by Morgan 2009/12/9 +抓CP未發文有期限資料
         'Modify by Morgan 2011/6/10 +排除程序管制的案件性質
         'Modified by Lydia 2023/03/17 +CP43
         strExc(0) = "select '',cpm03,sqldatet(NP08),sqldatet(NP09)" & _
            ",NP13,NP14,NP01,NP22,NP07,NP15,sqldatet(NP23) from nextprogress,casepropertymap where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
            " and (np06 is null or np06='') and np02=CPM01(+) and np07=cpm02(+)" & strNpSqlOfNoSalesDuty & _
            " UNION ALL select '',cpm03,sqldatet(cp06),sqldatet(cp07)" & _
            ",cp08,cp40,cp09,0,cp10,cp64,'' from caseprogress,casepropertymap where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
            " and cp27 is null and CPM01(+)=cp01 and CPM02(+)=cp10 and cp09<>'" & strReceiveNo & "' and cp07>0"
         '92.10.23 END
      End If
      intI = 1
      Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = rsTemp1
      'Add By Cheng 2002/08/19
      If rsTemp1.RecordCount > 0 Then
         m_strDL06 = "" & rsTemp1("NP22").Value
      End If
      
      GridHead
      If Not IsNull(.Fields(0)) Then Label2(5) = ChangeTStringToTDateString(TransDate(.Fields(0), 1))
      If Not IsNull(.Fields(3)) Then Label2(2) = .Fields(3)
      cp(10) = .Fields(4)
      If Not IsNull(.Fields(5)) Then Label4 = .Fields(5)
   End If
   End With
   
   'Added by Lydia 2023/03/17 法律所案源：取得案源類別
   If m_CP10 = "404" Then
       strExc(0) = "select cp162,los02 from caseprogress,lawofficesource where cp09='" & m_CP43 & "' and cp162=los15(+) "
   Else
       strExc(0) = "select cp162,los02 from caseprogress,lawofficesource where cp09='" & strReceiveNo & "' and cp162=los15(+) "
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       m_LOS15 = "" & RsTemp.Fields("cp162")
       m_LOS02 = "" & RsTemp.Fields("los02")
   End If
   'end 2023/03/17
End Sub

Private Sub FormShow()
   Label2(4) = pa(11)
   Combo1.ListIndex = 0
   lblFM2 = pa(5)
End Sub

Private Sub MSHFlexGrid1_Click()
 Dim strTmp As String
 
   If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 1) = "" Then Exit Sub
   Work_CP10 = ""
   '92.10.23 MODIFY BY SONIA 改為可複選
   'GridClick MSHFlexGrid1, intLastRow, 0
   GridClick MSHFlexGrid1, intLastRow, 0, 1
   '92.10.23 END
   intLastRow = MSHFlexGrid1.row
   If intLastRow > 0 Then
      strTmp = TransDate(Text7.Text, 2)
      If MSHFlexGrid1.TextMatrix(intLastRow, 0) = "v" And strTmp <> "" Then
         'Add by Morgan 2009/12/24
         strExc(1) = Replace(MSHFlexGrid1.TextMatrix(intLastRow, 2), "/", "")
         strExc(2) = Replace(MSHFlexGrid1.TextMatrix(intLastRow, 3), "/", "")
         strExc(3) = "" 'Add By Sindy 2021/7/15
         If DBDATE(strExc(1)) <> cp(6) Then
            MsgBox "所點選案件性質的本所期限與延期程序不同，不可點選！"
            MSHFlexGrid1.TextMatrix(intLastRow, 0) = ""
         ElseIf DBDATE(strExc(2)) <> cp(7) Then
            MsgBox "所點選案件性質的法定期限與延期程序不同，不可點選！"
            MSHFlexGrid1.TextMatrix(intLastRow, 0) = ""
         Else
         'end 2009/12/24
            strExc(0) = TransDate(strTmp, 2)
            Work_CP10 = MSHFlexGrid1.TextMatrix(intLastRow, 8)
            
            'Added by Morgan 2025/3/5
            strExc(4) = MSHFlexGrid1.TextMatrix(intLastRow, 9)
            strExc(5) = ""
            If Work_CP10 = "416" Then
               strExc(5) = MSHFlexGrid1.TextMatrix(intLastRow, 1)
            ElseIf Work_CP10 = "202" And InStr(strExc(4), "優先權證明") > 0 Then
               strExc(5) = MSHFlexGrid1.TextMatrix(intLastRow, 1) & "-優先權證明"
            End If
            If strExc(5) <> "" Then
               MsgBox "【延期】不可點選【" & strExc(5) & "】！", vbExclamation
               MSHFlexGrid1.TextMatrix(intLastRow, 0) = ""
               Exit Sub
            End If
            'end 2025/3/5
            
            'Removed by Morgan 2013/8/16 改都顯示
            ''Added by Morgan 2013/1/15
            'If Work_CP10 = "107" Then
            '   txtTimes.Visible = True
            '   lblTimes.Visible = True
            'Else
            '   txtTimes.Visible = False
            '   lblTimes.Visible = False
            'End If
            ''end 2013/1/15
            
            'Modified by Morgan 2013/8/16
            'CheckExtended MSHFlexGrid1.TextMatrix(intLastRow, 6), Work_CP10 'Added by Morgan 2013/6/19
            'If Not (Work_CP10 = "107" And Val(txtTimes) > "1") Then 'Added by Morgan 2013/6/19
            If CheckExtended(MSHFlexGrid1.TextMatrix(intLastRow, 6), Work_CP10) = False Then
            'end 2013/8/16
            
               
               '92.10.23 MODIFY BY SONIA 改為可複選
               'If objLawDll.GetCaseFeeDelay(pa(1), m_PA09, MSHFlexGrid1.TextMatrix(intLastRow, 8), strExc) Then
               'edit by nickc 2007/02/05 不用 dll 了
               'If objLawDll.GetCaseFeeDelay(pa(1), m_PA09, MSHFlexGrid1.TextMatrix(intLastRow, 8), strExc) Then
               'Modify by Morgan 2008/1/7 加傳本所案號以判斷申復的延期天數
               'Modify by Morgan 2011/2/21 改用期限計算(本所期限也是直接計算,不必用法定期限推,這樣通知信上期限的日才會一致)
               'If ClsLawGetCaseFeeDelay(pa(1), m_PA09, MSHFlexGrid1.TextMatrix(intLastRow, 8), strExc, pa(1) & pa(2) & pa(3) & pa(4)) Then
               If ClsLawGetCaseFeeDelay(pa(1), m_PA09, MSHFlexGrid1.TextMatrix(intLastRow, 8), strExc, pa(1) & pa(2) & pa(3) & pa(4), iDay, iMonth) Then
                  'Modify by Morgan 2011/3/8 只有申復用原期限算其餘仍用發文日
                  If MSHFlexGrid1.TextMatrix(intLastRow, 8) = "205" Then
                     SetDeadline strExc(1), strExc(2), iMonth, iDay, cp(7), cp(6), strExc(3)
                  End If
               'end 2011/2/21
               '92.10.23 END
                  Text6 = TransDate(strExc(1), 1)
                  Text5 = TransDate(strExc(2), 1)
                  Text8 = TransDate(strExc(3), 1) 'Add By Sindy 2021/5/7
               End If
            End If 'Added by Morgan 2013/6/19
            
         End If
      Else
         Text5 = ""
         Text5.Enabled = True
         Text6 = ""
         Text8 = "" 'Add By Sindy 2021/5/7
      End If
   End If
End Sub
'Modified by Morgan 2013/8/16 改 Sub 為 Function
'Added by Morgan 2013/6/19
'設定延期次數及期限
Private Function CheckExtended(pCP43 As String, pCP10 As String) As Boolean
   Dim stSQL As String, adoRst As ADODB.Recordset, intR As Integer
   Dim stDate1 As String, stDate2 As String
   'Added by Morgan 2013/8/16
   Dim bolDone As Boolean, stCP27 As String
   
   txtTimes = 0
   Label15 = ""
   bolDone = False
   'end 2013/8/16
   
   stSQL = "select cp27 from caseprogress where cp43='" & pCP43 & "' and cp10='404' and cp27>0"
   '考慮前次延期時未收文但本次延期已收文情形
   If pCP43 < "C" Then
      stSQL = stSQL & " union select b.cp27 from caseprogress a,caseprogress b where a.cp09='" & pCP43 & "' and b.cp43=a.cp43 and b.cp10='404' and b.cp27>0"
   End If
   stSQL = "select count(*),max(cp27),min(cp27) from (" & stSQL & ")"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      'MODIFY BY SONIA 2013/7/24 FCP-034425
      'txtTimes = adoRst(0)
      txtTimes = Val(adoRst(0)) + 1
      '2013/7/24 END
      
      If Val(txtTimes) > 1 Then
         Label15 = TransDate(adoRst(1), 1)
         '再審第二次延期,期限設為第一次延期發文日+6個月(所限=法限-4天)
         If pCP10 = "107" Then
            stDate1 = CompDate(1, 6, adoRst(2))
            Text6 = TransDate(stDate1, 1)
            
            'Modified by Morgan 2014/11/20 外專改回舊規則
            ''Added by Morgan 2014/10/29
            'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
            '   Text5 = TransDate(PUB_GetOurDeadline(stDate1), 1)
            'Else
            ''end 2014/10/29
            
            'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
            If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
               'Modify By Sindy 2021/4/23 + m_pAgreeOnDate
               stDate2 = PUB_GetFCPOurDeadline(stDate1, 4, , m_pAgreeOnDate)
            Else
            'end 2019/7/11
               stDate2 = CompDate(2, -4, stDate1)
            End If 'Added by Morgan 2019/7/11
            
            Text5 = TransDate(stDate2, 1)
            Text8 = TransDate(m_pAgreeOnDate, 1) 'Add By Sindy 2021/5/7 約定期限
            'End If 'Added by Morgan 2014/10/29
            'end 2014/11/20
            
            bolDone = True 'Added by Morgan 2013/8/16
         End If
      End If
   End If
   
   
   'Added by Morgan 2013/8/16
   If bolDone = False Then
      '新案翻譯,檢視中說,製作中說,補文件
      'Modified by Morgan 2013/11/6 +235核對中說格式
      If pCP10 = "201" Or pCP10 = "209" Or pCP10 = "235" Or pCP10 = "210" Or pCP10 = "202" Then
         '申請程序發文日
         stSQL = "select cp27 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp09<'B' and cp10 in (" & GetAddStr(NewCasePtyList) & ") and cp27>0"
         intR = 1
         Set adoRst = ClsLawReadRstMsg(intR, stSQL)
         If intR = 1 Then
            stCP27 = adoRst(0)
            '發明,新式樣
            '法限=新案申請發文日+6個月,所限=法限-4天
            '新型
            '第1次:法限=新案申請發文日+4個月,所限=法限-4天
            '第2次:法限=新案申請發文日+6個月,所限=法限-4天
            If pa(8) = "2" And txtTimes = 1 Then
               stDate1 = CompDate(1, 4, stCP27)
            Else
               stDate1 = CompDate(1, 6, stCP27)
            End If
            Text6 = TransDate(stDate1, 1)
            
            'Modified by Morgan 2014/11/20 外專改回舊規則
            ''Added by Morgan 2014/10/29
            'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
            '   Text5 = TransDate(PUB_GetOurDeadline(stDate1), 1)
            'Else
            ''end 2014/10/29
            
            'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
            If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
               'Modify By Sindy 2021/4/23 + m_pAgreeOnDate
               stDate2 = PUB_GetFCPOurDeadline(stDate1, 4, , m_pAgreeOnDate)
            Else
            'end 2019/7/11
               stDate2 = CompDate(2, -4, stDate1)
            End If 'Added by Morgan 2019/7/11
            
            Text5 = TransDate(stDate2, 1)
            Text8 = TransDate(m_pAgreeOnDate, 1) 'Add By Sindy 2021/5/7 約定期限
               
            'End If 'Added by Morgan 2014/10/29
            'end 2014/11/20
            
            bolDone = True
         End If
      
      'Added by Morgan 2013/10/9
      'Modified by Morgan 2013/12/29 +204 --靜芳
      'Modified by Morgan 2020/8/19 +239 --淑華
      ElseIf (pCP10 = "204" Or pCP10 = "205" Or pCP10 = "239") Then
         'Added by Morgan 2019/9/4
         '審查意見通知函第一次延期後期限之計算:來函期限3個月->+6個月,2個月->+4個月,1個月->+2個月
         If Val(txtTimes) = 1 Then
            'Modified by Morgan 2020/5/14 +cp134
            'Modified by Morgan 2020/8/19 +1232 --淑華
            If pCP43 > "C" Then
               stSQL = "select cp05,cp07,cp134 from caseprogress where cp09='" & pCP43 & "' and cp10 in ('1202','1232')"
            Else
               'Modified by Lydia 2019/10/28 debug
               'stSQL = "select cp05,cp07 from caseprogress a,caseprogress b where a.cp09='" & pCP43 & "' and b.cp09(+)=a.cp43 and b.cp10='1202'"
               stSQL = "select b.cp05,b.cp07,b.cp134 from caseprogress a,caseprogress b where a.cp09='" & pCP43 & "' and b.cp09(+)=a.cp43 and b.cp10 in ('1202','1232')"
            End If
            intR = 1
            Set adoRst = ClsLawReadRstMsg(intR, stSQL)
            If intR = 1 Then
               strExc(1) = ""
               'Modified by Morgan 2020/5/15 改先抓來函期限月數並修正月數計算錯誤問題 Ex:FCP-062368
               If adoRst("cp134") > 0 Then
                  If adoRst("cp134") = 3 Or adoRst("cp134") = 2 Or adoRst("cp134") = 1 Then
                     strExc(1) = CompDate(1, 2 * adoRst("cp134"), adoRst("cp05"))
                  End If
               End If
               
               If strExc(1) = "" Then
                  '3個月
                  strExc(0) = CompDate(1, 3, adoRst("cp05"))
                  If strExc(0) = adoRst("cp07") Then
                     '來函收文日+6個月
                     strExc(1) = CompDate(1, 6, adoRst("cp05"))
                  End If
               End If
               
               If strExc(1) = "" Then
                  '2個月
                  strExc(0) = CompDate(1, 2, adoRst("cp05"))
                  If strExc(0) = adoRst("cp07") Then
                     '來函收文日+4個月
                     strExc(1) = CompDate(1, 4, adoRst("cp05"))
                  End If
               End If
               
               If strExc(1) = "" Then
                  '1個月
                  strExc(0) = CompDate(1, 1, adoRst("cp05"))
                  If strExc(0) = adoRst("cp07") Then
                     '來函收文日+2個月
                     strExc(1) = CompDate(1, 2, adoRst("cp05"))
                  End If
               End If
               
               If strExc(1) <> "" Then
                  'Modify By Sindy 2021/4/23 + m_pAgreeOnDate
                  strExc(2) = PUB_GetFCPOurDeadline(strExc(1), 4, , m_pAgreeOnDate)
                  Text6 = TransDate(strExc(1), 1)
                  Text5 = TransDate(strExc(2), 1)
                  Text8 = TransDate(m_pAgreeOnDate, 1) 'Add By Sindy 2021/5/7 約定期限
                  bolDone = True
               End If
               'end 2020/5/15
            End If
         End If
         
         If bolDone = False Then
         'end 2019/9/4
            If pa(8) = "2" Then
               SetDeadline strExc(1), strExc(2), 1, 0, cp(7), cp(6), strExc(3)
               Text6 = TransDate(strExc(1), 1)
               Text5 = TransDate(strExc(2), 1)
               Text8 = TransDate(strExc(3), 1) 'Add By Sindy 2021/5/7
               bolDone = True
            End If
            
         End If 'Added by Morgan 2019/9/4
         
      'end 2013/10/9
      End If
   End If
   CheckExtended = bolDone
   'end 2013/8/16
   
   Set adoRst = Nothing
End Function

'Add By Sindy 2021/4/23
Private Sub MSHFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   If MSHFlexGrid1.MouseRow <> 0 And _
      (MSHFlexGrid1.MouseCol = 9) Then
      If iRow <> MSHFlexGrid1.MouseRow Or iCol <> MSHFlexGrid1.MouseCol Then
         If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, MSHFlexGrid1.MouseCol) <> "" Then
            CreateToolTip GetHWndForToolTip(MSHFlexGrid1), MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, MSHFlexGrid1.MouseCol)
            iRow = MSHFlexGrid1.MouseRow
            iCol = MSHFlexGrid1.MouseCol
         End If
      End If
   End If
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
    If Text5 = "" Then
        MsgBox "延期本所期限不可為空值，請重新輸入 !", vbCritical
        Cancel = True
    Else
        If ChkDate(Text5) Then
            'Modify By Cheng 2003/01/17
            '判斷是否修改了延期限本所期限
            If Me.Text5.Text <> Me.Text5.Tag And Me.Text7.Text <> GetTaiwanTodayDate Then
                If MsgBox("是否確定修改延期本所期限 !", vbYesNo) = vbNo Then
                    Cancel = True
                Else
                    '記錄新的延期後本所期限
                    Me.Text5.Tag = Me.Text5.Text
                End If
            End If
        Else
            Cancel = True
        End If
    End If
    If Cancel Then TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
    If Text6 = "" Then
        MsgBox "延期法定期限不可為空值，請重新輸入 !", vbCritical
        Cancel = True
    Else
        If ChkDate(Text6) Then
            'Modify By Cheng 2003/01/17
            '判斷是否修改了延期限法定期限
            If Me.Text6.Text <> Me.Text6.Tag And Me.Text7.Text <> GetTaiwanTodayDate Then
                If MsgBox("是否確定修改延期法定期限 !", vbYesNo) = vbNo Then
                    Cancel = True
                Else
                    '記錄新的延期後法定期限
                    Me.Text6.Tag = Me.Text6.Text
                End If
            End If
        Else
            Cancel = True
        End If
    End If
    If Cancel Then TextInverse Text6
End Sub

Private Sub Text8_GotFocus()
  TextInverse Text8
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
    If Text8 = "" Then
        MsgBox "延期約定期限不可為空值，請重新輸入 !", vbCritical
        Cancel = True
    Else
        If ChkDate(Text8) Then
            'Modify By Cheng 2003/01/17
            '判斷是否修改了延期限約定期限
            If Me.Text8.Text <> Me.Text8.Tag And Me.Text7.Text <> GetTaiwanTodayDate Then
                If MsgBox("是否確定修改延期約定期限 !", vbYesNo) = vbNo Then
                    Cancel = True
                Else
                    '記錄新的延期後約定期限
                    Me.Text8.Tag = Me.Text8.Text
                End If
            End If
        Else
            Cancel = True
        End If
    End If
    If Cancel Then TextInverse Text6
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1000: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1200: .Text = "機關文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "相關人"
      .col = 6: .ColWidth(6) = 0
      .col = 7: .ColWidth(7) = 0
      .col = 8: .ColWidth(8) = 0
      .col = 9: .ColWidth(9) = 2000: .Text = "備註"
      'Add By Sindy 2021/7/20
      .col = 10:  .Text = "約定期限"
      If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
         .ColWidth(10) = 1000
      Else
         .ColWidth(10) = 0
      End If
      '2021/7/20 END
      .Visible = True
   End With
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub
'Modify by Morgan 2011/3/8 只有申復用原期限算其餘仍用發文日
'Remove by Morgan 2011/2/21 改用期限計算(本所期限也是直接計算,不必用法定期限推,這樣通知信上期限的日才會一致)
Private Sub Text7_LostFocus()
Dim strTmp As String
   
   If Text7 <> "" Then
      If m_CP10 <> "404" Then
         CaculateNP08NP09
      
      ElseIf MSHFlexGrid1.TextMatrix(intLastRow, 0) = "v" And intLastRow > 0 Then 'Added by Morgan 2013/8/16
      
         If CheckExtended(MSHFlexGrid1.TextMatrix(intLastRow, 6), MSHFlexGrid1.TextMatrix(intLastRow, 8)) = False Then 'Added by Morgan 2013/8/16
         
            If MSHFlexGrid1.TextMatrix(intLastRow, 8) <> "205" Then
               strTmp = TransDate(Text7.Text, 2)
               If MSHFlexGrid1.TextMatrix(intLastRow, 0) = "v" And strTmp <> "" And intLastRow > 0 Then
               
                  strExc(0) = TransDate(strTmp, 2)
                  'edit by nickc 2007/02/05 不用 dll 了
                  'If objLawDll.GetCaseFeeDelay(pa(1), m_PA09, MSHFlexGrid1.TextMatrix(intLastRow, 8), strExc) Then
                  'Modify by Morgan 2008/1/7 加傳本所案號以判斷申復的延期天數
                  If ClsLawGetCaseFeeDelay(pa(1), m_PA09, MSHFlexGrid1.TextMatrix(intLastRow, 8), strExc, pa(1) & pa(2) & pa(3) & pa(4)) Then
                     Text6 = TransDate(strExc(1), 1)
                     Text5 = TransDate(strExc(2), 1)
                     Text8 = TransDate(strExc(3), 1) 'Add By Sindy 2021/5/7
                  End If
               End If
            End If
            
         End If 'Added by Morgan 2013/8/16
      End If
   End If
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 = "" Then
      MsgBox "延期日不得為空值 !", vbCritical
      Cancel = True
   Else
      If Not ChkDate(Text7.Text) Then
         TextInverse Text7
         Cancel = True
      'Add By Cheng 2002/07/05
      ElseIf Val(Me.Text7.Text) > PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1))) Then
         MsgBox "延期日不可大於系統日的下一個工作日!!!", vbExclamation + vbOKOnly
         TextInverse Text7
         Cancel = True
      ' 90.06.27 modify by louis
      End If
   End If
End Sub

' 計算本所期限及法定期限
Private Sub CaculateNP08NP09()
   'Modified by Morgan 2013/8/16
   'If Not (m_CP10 = "107" And Val(txtTimes) > 1) Then 'Added by Morgan
   If CheckExtended(strReceiveNo, m_CP10) = False Then
   'end 2013/8/16
      If IsEmptyText(Text7) = False And Text7.Tag <> Text7.Text Then
         strExc(0) = TransDate(Text7.Text, 2)
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.GetCaseFeeDelay(pa(1), m_PA09, m_CP10, strExc) Then
         'Modify by Morgan 2008/1/7 加傳本所案號以判斷申復的延期天數
         'Modify by Morgan 2011/2/21 改用期限計算(本所期限也是直接計算,不必用法定期限推,這樣通知信上期限的日才會一致)
         'If ClsLawGetCaseFeeDelay(pa(1), m_PA09, m_CP10, strExc, pa(1) & pa(2) & pa(3) & pa(4)) Then
         If ClsLawGetCaseFeeDelay(pa(1), m_PA09, m_CP10, strExc, pa(1) & pa(2) & pa(3) & pa(4), iDay, iMonth) Then
            'Modify by Morgan 2011/3/8 只有申復用原期限算其餘仍用發文日
            If m_CP10 = "205" Then
               SetDeadline strExc(1), strExc(2), iMonth, iDay, cp(7), cp(6), strExc(3)
            End If
         'end 2011/2/21
            Text6 = TransDate(strExc(1), 1)
            Text5 = TransDate(strExc(2), 1)
            Text8 = TransDate(strExc(3), 1) 'Add By Sindy 2021/5/7
         End If
         Text7.Tag = Text7.Text
      End If
      
   End If 'Added by Morgan 2103/6/19
   
   'Removed by Morgan 2013/8/16 移到上面
   'Add by SONIA 2013/7/24 自ReadPatent移過來
   'If m_CP10 = "107" Then
   '   CheckExtended strReceiveNo, m_CP10
   'End If
   '2013/7/24 END
   
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
'Add By Cheng 2002/12/02
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim bFind As Boolean
Dim nIndex  As Integer

TxtValidate = False

   'Added by Morgan 2015/10/6
   If Work_CP10 = "201" And Val(txtTimes) > 1 Then
      MsgBox "本案翻譯已延期過不得再延期！", vbExclamation
      Exit Function
   End If
   'end 2015/10/6


    'Add By Cheng 2002/12/02
   ' 當案件性質為延期時, 未收文期限至少要選取一筆
   If cp(10) = "404" Then
      If Me.MSHFlexGrid1.Rows <= 1 Then
         strTit = "檢核資料"
         strMsg = "未收文期限無資料, 無法執行延期的處理"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
       Exit Function
      End If
      
      bFind = False
      For nIndex = 1 To Me.MSHFlexGrid1.Rows - 1
         If Me.MSHFlexGrid1.TextMatrix(nIndex, 0) = "v" Then
            bFind = True
            Exit For
         End If
      Next nIndex
      If bFind = False Then
         strTit = "檢核資料"
         strMsg = "請先選取未收文期限的資料來做延期的處理"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        Exit Function
      End If
   End If

If Me.Text7.Enabled = True Then
   Cancel = False
   Text7_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text5.Enabled = True Then
   Cancel = False
   Text5_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text6.Enabled = True Then
   Cancel = False
   Text6_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add by Morgan 2004/8/12
If txtCP84.Enabled = True Then
   Cancel = False
   txtCP84_Validate Cancel
   If Cancel = True Then
      txtCP84.SetFocus
      txtCP84_GotFocus
      Exit Function
   End If
   
   'Added by Morgan 2013/1/11
   '再審107延期要檢查發文規費(有輸入才要,因為有可能是第2次延期)--靜芳
   If Work_CP10 = "107" Then
      
      'Removed by Morgan 2013/8/16 已改為自動帶且不可輸入
      'If txtTimes.Text = "" Then
      '   MsgBox "請輸入再審的次數!!!", vbExclamation + vbOKOnly
      '   txtTimes.SetFocus
      '   txtTimes_GotFocus
      '   Exit Function
      'End If
      'end 2013/8/16
         
      If Val(txtTimes) < 2 Then
         strExc(1) = GetPatentOfficialFee(pa(1), "107", "", pa(8), pa(9), pa(16))
      Else
         strExc(1) = 0
      End If
      If Val(txtCP84) <> Val(strExc(1)) Then
         If MsgBox("發文規費 $" & txtCP84 & " 與系統設定 $" & strExc(1) & " 不同是否要繼續?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            txtCP84.SetFocus
            txtCP84_GotFocus
            Exit Function
         End If
      End If
   End If
   'end 2013/1/11
   
End If
   'Add by Morgan 2005/8/8
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If

TxtValidate = True
End Function

'Add By Sindy 2022/5/17
Private Sub txtRecDate_GotFocus()
   TextInverse txtRecDate
End Sub
Private Sub txtRecDate_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub
Private Sub txtEmail_GotFocus()
   TextInverse txtEmail
End Sub
Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      Beep
      KeyAscii = 0
   End If
End Sub
Private Sub txtRecDate_Validate(Cancel As Boolean)
   If txtRecDate.Tag <> txtRecDate.Text Then
      If txtRecDate = "Y" Then
         txtEmail = "Y"
      End If
   End If
   txtRecDate.Tag = txtRecDate.Text
End Sub
'2022/5/17 END

Private Sub txtTimes_Change()
   If Work_CP10 = "107" And txtCP84.Enabled = True Then
      If Val(txtTimes) = 1 Then
         txtCP84 = GetPatentOfficialFee(pa(1), Work_CP10, "", pa(8), pa(9), pa(16))
      ElseIf Val(txtTimes) > 1 Then
         txtCP84 = 0
      Else
         txtCP84 = ""
      End If
   End If
End Sub

Private Sub txtTimes_GotFocus()
   TextInverse txtTimes
   CloseIme
End Sub

Private Sub txtTimes_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
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
      '已收文才要檢查
      If m_str_DL05 = "1" Then
         If Val(txtCP84.Text) <> Val(m_CP17) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
            If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & m_CP17 & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
               txtCP84.Tag = txtCP84.Text
            Else
               txtCP84_GotFocus
               Cancel = True
            End If
         End If
      End If
   End If
End Sub
'Add by Morgan 2005/8/8
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/4/23
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/4/23 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2)
   End If
End Sub

'Add by Morgan 2010/11/15
Private Sub StartLetter2(ByVal ET01 As String, ET02 As String, ByVal ET03 As String)
   Dim strTxt() As String, i As Integer
   Dim strDoc As String

   EndLetter ET01, ET02, ET03, strUserNum

   i = 1
   ReDim Preserve strTxt(i)
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','延期後法定期限','" & DBDATE(Text6) & "')"
   
   i = i + 1
   ReDim Preserve strTxt(i)
   'Modify By Sindy 2021/4/23
   If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','延期後約定期限','" & DBDATE(Text8) & "')" 'm_pAgreeOnDate
   Else
   '2021/4/23 END
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','延期後本所期限','" & DBDATE(Text5) & "')"
   End If
   
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Add by Morgan 2011/2/21
'Modify By Sindy 2021/5/7 +, ByRef newNP23 As String
Private Sub SetDeadline(ByRef newCP07 As String, ByRef newCP06 As String, ByVal iM As Integer, _
   ByVal Id As Integer, ByVal oldCP07 As String, ByVal oldCP06 As String, ByRef newNP23 As String)
   
   If iM > 0 Then
      newCP07 = CompDate(1, iM, oldCP07)
      
      'Modified by Morgan 2014/11/20 外專改回舊規則
      ''Added by Morgan 2014/10/30
      'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
      '   newCP06 = PUB_GetOurDeadline(newCP07)
      'Else
      ''end 2014/10/30
      
      'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
      '108/7/11 David 確認所限可用新規則改以法限推算
      If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
         'Modify By Sindy 2021/4/23 + newNP23
         newCP06 = PUB_GetFCPOurDeadline(newCP07, 4, , newNP23)
      Else
      'end 2019/7/11
      
         newCP06 = CompDate(1, iM, oldCP06)
         
      End If 'Added by Morgan 2019/7/11
      
      'End If 'Added by Morgan 2014/10/30
      'end 2014/11/20
      
   ElseIf Id > 0 Then
      newCP07 = CompDate(2, Id, oldCP07)
      
      'Modified by Morgan 2014/11/20 外專改回舊規則
      ''Added by Morgan 2014/10/30
      'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
      '   newCP06 = PUB_GetOurDeadline(newCP07)
      'Else
      ''end 2014/10/30
      
      'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
      '108/7/11 David 確認所限可用新規則改以法限推算
      If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
         'Modify By Sindy 2021/4/23 + newNP23
         newCP06 = PUB_GetFCPOurDeadline(newCP07, IIf(Id >= 30, 4, 2), , newNP23)
      Else
      'end 2019/7/11
      
         newCP06 = CompDate(2, Id, oldCP06)
         
      End If 'Added by Morgan 2019/7/11
      
      'End If 'Added by Morgan 2014/10/30
      'end 2014/11/20
      
   End If
End Sub

'Added by Lydia 2018/01/11 延期+電子送件
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
'end 2018/01/11
