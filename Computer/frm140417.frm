VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140417 
   BorderStyle     =   1  '單線固定
   Caption         =   "委任契約書用印記錄查詢"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8295
   Begin VB.CommandButton Cmd1 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   17
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   0
      Left            =   6720
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4500
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7938
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "清單"
      TabPicture(0)   =   "frm140417.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "MGrid1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtField(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtField(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtField(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Combo1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Check1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "明細資料"
      TabPicture(1)   =   "frm140417.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtRS07"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "lblRS(8)"
      Tab(1).Control(3)=   "Label1(9)"
      Tab(1).Control(4)=   "lblRS(7)"
      Tab(1).Control(5)=   "Label1(8)"
      Tab(1).Control(6)=   "lblRS(6)"
      Tab(1).Control(7)=   "lblRS(5)"
      Tab(1).Control(8)=   "lblRS(4)"
      Tab(1).Control(9)=   "lblRS(3)"
      Tab(1).Control(10)=   "lblRS(2)"
      Tab(1).Control(11)=   "lblRS(0)"
      Tab(1).Control(12)=   "Label1(4)"
      Tab(1).Control(13)=   "Label1(3)"
      Tab(1).Control(14)=   "Label1(2)"
      Tab(1).Control(15)=   "Label1(1)"
      Tab(1).Control(16)=   "Label1(0)"
      Tab(1).ControlCount=   17
      Begin VB.CheckBox Check1 
         Caption         =   "只找空白委任書"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   16
         Top             =   870
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1320
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   6480
         MaxLength       =   7
         TabIndex        =   12
         Top             =   440
         Width           =   1080
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   5280
         MaxLength       =   7
         TabIndex        =   11
         Top             =   440
         Width           =   1080
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   10
         Top             =   440
         Width           =   720
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
         Bindings        =   "frm140417.frx":0038
         Height          =   2985
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   5265
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "v|列印人員|列印日期|列印時間|委任書種類|受任人|列印份數|空白委任書"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSForms.TextBox TxtRS07 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   2
         Top             =   1320
         Width           =   7800
         VariousPropertyBits=   -1472184289
         ScrollBars      =   3
         Size            =   "13758;5318"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   300
         Left            =   -72960
         TabIndex        =   21
         Top             =   480
         Width           =   615
         VariousPropertyBits=   27
         Caption         =   "Label2"
         Size            =   "1085;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblName 
         Height          =   300
         Left            =   2160
         TabIndex        =   19
         Top             =   505
         Width           =   1335
         VariousPropertyBits=   27
         Caption         =   "lblName"
         Size            =   "2355;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         Caption         =   "lblRS(8)"
         Height          =   180
         Index           =   8
         Left            =   -70080
         TabIndex        =   29
         Top             =   780
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "受任人："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   -71040
         TabIndex        =   28
         Top             =   780
         Width           =   780
      End
      Begin VB.Label lblRS 
         Caption         =   "lblRS(7)"
         Height          =   195
         Index           =   7
         Left            =   -69240
         TabIndex        =   27
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否為空白委任書："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   -71040
         TabIndex        =   26
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label lblRS 
         Caption         =   "lblRS(6)"
         Height          =   195
         Index           =   6
         Left            =   -73680
         TabIndex        =   25
         Top             =   1080
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   6120
         X2              =   6720
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblRS 
         Caption         =   "lblRS(5)"
         Height          =   195
         Index           =   5
         Left            =   -67560
         TabIndex        =   24
         Top             =   780
         Width           =   615
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         Caption         =   "lblRS(4)"
         Height          =   180
         Index           =   4
         Left            =   -73680
         TabIndex        =   23
         Top             =   780
         Width           =   600
      End
      Begin VB.Label lblRS 
         Caption         =   "lblRS(3)"
         Height          =   195
         Index           =   3
         Left            =   -69120
         TabIndex        =   22
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblRS 
         Caption         =   "lblRS(2)"
         Height          =   195
         Index           =   2
         Left            =   -70080
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblRS 
         Caption         =   "lblRS(0)"
         Height          =   195
         Index           =   0
         Left            =   -73680
         TabIndex        =   18
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "委任書種類："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   900
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "列印日期："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   6
         Left            =   4320
         TabIndex        =   13
         Top             =   505
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "列印份數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   -68520
         TabIndex        =   7
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "列印日期："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   -71040
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "委任書內容：　　　  字"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   -74880
         TabIndex        =   5
         Top             =   1080
         Width           =   2040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "委任書種類："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   -74880
         TabIndex        =   4
         Top             =   780
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "列印人員："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "列印人員："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   505
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm140417"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/27 Form2.0已修改(lblName,MGrid1改Fonts,lblRS(1)改為label2,TXTRS07)
'Created by Lydia 2017/03/24 委任契約書用印記錄查詢
Option Explicit
Dim oLbl As LABEL
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim mPrevRow As Integer
Dim colPK01 As Integer
Dim colPK02 As Integer
Dim colPK03 As Integer
Dim colPK04 As Integer

Private Sub Cmd1_Click(Index As Integer)
   Select Case Index
       Case 0
            Unload Me
       Case 1
            Call QueryData
   End Select
End Sub

Private Sub Form_Load()
 
   MoveFormToCenter Me
   SSTab1.Tab = 0
   
   Combo1.Clear
   Combo1.AddItem " ", 0
   Combo1.AddItem "P", 1
   Combo1.AddItem "CFP及P非台灣案", 2
   Combo1.AddItem "T", 3
   Combo1.AddItem "CFT及T大陸案及TF馬德里案", 4
   Combo1.AddItem "常年顧問聘任書", 5
   Combo1.AddItem "條碼案件委任契約書", 6
   Combo1.AddItem "著作權案件委任契約書", 7   'Added by Lydia 2022/04/08
   Combo1.AddItem "專利申請案保密同意書", 8   'Added by Lydia 2022/04/26
   
   Call SetGrid
   txtField(0) = ""
   txtField(1) = TransDate(CompDate(1, -1, strSrvDate(1)), 1)
   txtField(2) = strSrvDate(2)
   Combo1.Text = ""
   lblName = ""
   DataReset
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140417 = Nothing
End Sub

Private Sub QueryData()
Dim strQ As String
   mPrevRow = 0
   If Trim(txtField(0) & txtField(1) & txtField(2)) = "" And Combo1.Text = "" Then
      If MsgBox("是否要輸入查詢條件？", vbYesNo + vbDefaultButton1) = vbYes Then
         Exit Sub
      End If
   End If
   
   '列印人員
   If Trim(txtField(0)) <> "" Then
      strQ = strQ & " AND RS01=" & CNULL(Trim(txtField(0)))
   End If
   '列印日期
   If Trim(txtField(1)) <> "" Then
      strQ = strQ & " AND RS02>=" & TransDate(txtField(1), 2)
   End If
   If Trim(txtField(2)) <> "" Then
      strQ = strQ & " AND RS02<=" & TransDate(txtField(2), 2)
   End If
   
   '委件書種類
   If Trim(Combo1.Text) <> "" Then
      Select Case Trim(Combo1.Text)
          Case "P": strQ = strQ & " AND RS04='1'"
          Case "CFP及P非台灣案": strQ = strQ & " AND RS04='2'"
          Case "T": strQ = strQ & " AND RS04='3'"
          Case "CFT及T大陸案及TF馬德里案": strQ = strQ & " AND RS04='4'"
          Case "常年顧問聘任書": strQ = strQ & " AND RS04='5'"
          Case "條碼案件委任契約書": strQ = strQ & " AND RS04='6'"
          Case "著作權案件委任契約書": strQ = strQ & " AND RS04='7'" 'Added by Lydia 2022/04/08
          Case "專利申請案保密同意書": strQ = strQ & " AND RS04='8'" 'Added by Lydia 2022/04/26
      End Select
   End If
   '空白委任書
   If Check1.Value = 1 Then strQ = strQ & " AND RS06='Y'"
   
   strSql = "SELECT ' ' v,ST02,SQLDATET(RS02) RS02T,SQLTIME6(RS03) RS03T,"
   'Modified by Lydia 2022/04/08 +著作權+, '7','著作權案件委任契約書'
   'Modified by Lydia 202/04/26 +保密同意書+ ,'8','專利申請案保密同意書'
   strSql = strSql & "DECODE(RS04,'1','P','2','CFP及P非台灣案','3','T','4','CFT及T大陸案及TF馬德里案'," & _
                            "'5','常年顧問聘任書','6','條碼案件委任契約書','7','著作權案件委任契約書','8','專利申請案保密同意書',RS04) RS04T"
   'Modified by Lydia 2020/03/25 改成與公司別ACC080一致；舊資料不動
   'strSql = strSql & ",DECODE(RS08,'1','專利商標','2','專利法律','3','台一智權',RS08) RS08T"
   strSql = strSql & ",DECODE(SIGN(RS02-" & 智慧所更名日 & ") ,-1 ,DECODE(RS08,'1','專利商標','2','專利法律','3','台一智權',RS08) " & _
                                                                                             ", DECODE(RS08,'1','專利商標','2','智慧所','J','台一智權','L','法律所',RS08)) AS RS08T"

   strSql = strSql & ",RS05,RS06,RS01,RS02,RS03,RS04 FROM RECSEAL,STAFF WHERE RS01=ST01(+)" & strQ
   strSql = strSql & " ORDER BY RS02,RS03,RS04"
   
   intI = 0
   colPK01 = 8: colPK02 = 9: colPK03 = 10: colPK04 = 11
   MGrid1.Rows = 2    'add by sonia 2018/2/22
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   Set MGrid1.Recordset = RsTemp
   If intI = 0 Then
      Call SetGrid
   Else
      Call SetGrid(RsTemp.RecordCount + 1)
   End If
End Sub

Private Function ReadData(ByVal pKey01 As String, ByVal pKey02 As String, ByVal pKey03 As String, ByVal pKey04 As String) As Boolean
Dim stCon As String
Dim rsAD As New ADODB.Recordset
  
   strSql = "SELECT ST02,(RS02-19110000) RS02T,SQLTIME6(RS03) RS03T,"
   'Modified by Lydia 2022/04/08 +著作權+, '7','著作權案件委任契約書'
   'Modified by Lydia 202/04/26 +保密同意書+ ,'8','專利申請案保密同意書'
   strSql = strSql & "DECODE(RS04,'1','P','2','CFP及P非台灣案','3','T','4','CFT及T大陸案及TF馬德里案'," & _
                            "'5','常年顧問聘任書','6','條碼案件委任契約書','7','著作權案件委任契約書','8','專利申請案保密同意書',RS04) RS04T"
   'Modified by Lydia 2020/03/25 改成與公司別ACC080一致；舊資料不動
   'strSql = strSql & ",DECODE(RS08,'1','專利商標','2','專利法律','3','台一智權',RS08) RS08T"
   strSql = strSql & ",DECODE(SIGN(RS02-" & 智慧所更名日 & ") ,-1 ,DECODE(RS08,'1','專利商標','2','專利法律','3','台一智權',RS08) " & _
                                                                                             ", DECODE(RS08,'1','專利商標','2','智慧所','J','台一智權','L','法律所',RS08)) AS RS08T"
   
   strSql = strSql & ",RS05,RS06,RS07,RS01,RS02,RS03,RS04 FROM RECSEAL,STAFF "
   strSql = strSql & "WHERE RS01=ST01(+) AND RS01='" & pKey01 & "' AND RS02=" & pKey02 & " AND RS03=" & pKey03 & " AND RS04='" & pKey04 & "' "

   intI = 0
   Set rsAD = ClsLawReadRstMsg(intI, strSql)
   
   DataReset
   If intI = 1 Then
      With rsAD
         lblRS(0).Caption = "" & .Fields("RS01")
         'modify by sonia 2021/12/27 lblRS(1)--Label2
         'lblRS(1).Caption = "" & .Fields("ST02")
         Label2.Caption = "" & .Fields("ST02")
         lblRS(2).Caption = "" & .Fields("RS02T")
         lblRS(3).Caption = "" & .Fields("RS03T")
         lblRS(4).Caption = "" & .Fields("RS04T")
         lblRS(5).Caption = "" & .Fields("RS05")
         TxtRS07.Text = "" & .Fields("RS07")
         lblRS(6).Caption = GetTextLength("" & .Fields("RS07")) '計算字數
         lblRS(7).Caption = "" & .Fields("RS06")
         lblRS(8).Caption = "" & .Fields("RS08T")
      End With
      ReadData = True
   End If
   
End Function

Private Sub SetGrid(Optional ByVal iCnt As Integer = 2)
Dim idR As Integer

    MGrid1.Visible = False
    MGrid1.Rows = iCnt
    MGrid1.FormatString = "v|列印人員|列印日期|列印時間|委任書種類|受任人|列印份數|空白委任書"
    MGrid1.ColWidth(0) = 280
    MGrid1.ColWidth(1) = 920
    MGrid1.ColWidth(2) = 920
    MGrid1.ColWidth(3) = 920
    MGrid1.ColWidth(4) = 2000
    MGrid1.ColWidth(5) = 920
    MGrid1.ColWidth(6) = 920
    MGrid1.ColWidth(7) = 1100
    For idR = 8 To MGrid1.Cols - 1
       MGrid1.ColWidth(idR) = 0
    Next
    MGrid1.Visible = True

End Sub

' 清空明細資料
Private Sub DataReset()
   TxtRS07.Text = ""
   For Each oLbl In lblRS
      oLbl.Caption = ""
   Next
   Label2.Caption = ""  'add by sonia 2021/12/27 lblRS(1)--Label2
End Sub

Private Sub MGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow MGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MGrid1.col = nCol
   MGrid1.row = nRow
   If Me.MGrid1.row < 1 And Me.MGrid1.Text <> "V" Then
      If InStr("列印日期,列印份數", Me.MGrid1.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.MGrid1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrid1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.MGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   TextInverse txtField(Index)
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
Dim iLen As Integer

   Cancel = False

   Select Case Index
        Case 0   '列印人員
           If Trim(txtField(Index)) = "" Then
              lblName = ""
           Else
              strExc(1) = GetStaffName(txtField(Index), True)
              lblName = strExc(1)
           End If
        Case 1  '列印日期
           If Trim(txtField(Index)) <> "" Then
              If Not ChkDate(txtField(Index)) Then
                txtField(Index).SetFocus
                Cancel = True
                Exit Sub
              End If
           End If
        Case 2
           If Trim(txtField(Index)) <> "" Then
              If Not ChkDate(txtField(Index)) Then
                txtField(Index).SetFocus
                Cancel = True
                Exit Sub
              End If
              If txtField(1) <> "" And Val(txtField(1)) > Val(txtField(2)) Then
                 MsgBox "起始日期不可大於終止日期!"
                 txtField(1).SetFocus
                 Cancel = True
                 Exit Sub
              End If
           End If
   End Select
   
   If Not CheckLengthIsOK(txtField(Index), iLen) Then
      Cancel = True
   End If

End Sub

Private Sub MGrid1_DblClick()
   If MGrid1.row > 0 And MGrid1.TextMatrix(MGrid1.row, colPK01) <> "" Then
      SSTab1.Tab = 1
   End If
End Sub

Private Sub MGrid1_SelChange()
Dim TmpRow As Integer
TmpRow = MGrid1.MouseRow
   
If TmpRow > 0 Then
   If TmpRow <> mPrevRow And mPrevRow > 0 Then
      MGrid1.TextMatrix(mPrevRow, 0) = ""
      Call ShowBar(MGrid1, mPrevRow, 8)
   End If
   Call ShowBar(MGrid1, TmpRow, 8)
   MGrid1.TextMatrix(TmpRow, 0) = "v"
   If mPrevRow = 0 Then
   End If
   If TmpRow > 0 And MGrid1.TextMatrix(TmpRow, 0) <> "" Then
      If ReadData(MGrid1.TextMatrix(MGrid1.row, colPK01), MGrid1.TextMatrix(MGrid1.row, colPK02), MGrid1.TextMatrix(MGrid1.row, colPK03), MGrid1.TextMatrix(MGrid1.row, colPK04)) Then
      End If
   End If
   
   mPrevRow = TmpRow
End If

End Sub
