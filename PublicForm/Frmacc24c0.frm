VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc24c0 
   AutoRedraw      =   -1  'True
   Caption         =   "FC業務請款／收款明細表"
   ClientHeight    =   5170
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   6740
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5170
   ScaleWidth      =   6740
   Begin VB.TextBox txtNa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2490
      MaxLength       =   3
      TabIndex        =   16
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtNa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   15
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtNArea 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   17
      Top             =   3480
      Width           =   612
   End
   Begin VB.TextBox txtKind 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   14
      Top             =   2760
      Width           =   612
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   5265
      ScaleHeight     =   460
      ScaleWidth      =   650
      TabIndex        =   33
      Top             =   1260
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "2"
      Top             =   2370
      Width           =   240
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   12
      Text            =   "1"
      Top             =   1995
      Width           =   612
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3060
      TabIndex        =   2
      Top             =   585
      Width           =   1572
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Top             =   585
      Width           =   1572
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   990
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   3960
      Width           =   4692
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   11
      Top             =   1665
      Width           =   612
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4170
      TabIndex        =   8
      Top             =   945
      Width           =   612
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3570
      TabIndex        =   7
      Top             =   945
      Width           =   612
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2970
      TabIndex        =   6
      Top             =   945
      Width           =   612
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2370
      TabIndex        =   5
      Top             =   945
      Width           =   612
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1770
      TabIndex        =   4
      Top             =   945
      Width           =   612
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      TabIndex        =   3
      Top             =   945
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1170
      TabIndex        =   9
      Top             =   1305
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3090
      TabIndex        =   10
      Top             =   1305
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Line Line1 
      X1              =   2250
      X2              =   2370
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "洲別：　　　  （0.亞洲, 1.美洲, 2.歐洲, 3.非洲, 4.大洋洲）"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   36
      Top             =   3525
      Width           =   6285
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "國籍："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   35
      Top             =   3150
      Width           =   735
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "類別：           （1.申請人 2.代理人）"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   34
      Top             =   2805
      Width           =   3930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "輸出方式：  ( 1.螢幕 2.印表機 )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   180
      TabIndex        =   32
      Top             =   2400
      Width           =   5070
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "(空白表全部)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4890
      TabIndex        =   31
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "國外商標處業務區：F10-F19     國外專利處業務區：F20-F29     國外法務處業務區：F30-F49"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1755
      TabIndex        =   30
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1.明細 2.智權人員 3.FCP組別 4.FCT組別)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   1896
      TabIndex        =   29
      Top             =   2028
      Width           =   4356
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "報表內容："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   28
      Top             =   2010
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   27
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "業務區："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   26
      Top             =   600
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "(1.請款點數 2.收款點數)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1890
      TabIndex        =   25
      Top             =   1665
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "資料內容："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   24
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2850
      TabIndex        =   23
      Top             =   1305
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "帳款日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   22
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "系統類別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   21
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "(空白表全部)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2850
      TabIndex        =   20
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "智權人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   180
      TabIndex        =   19
      Top             =   270
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc24c0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit
 
Public adoacc1k0 As New ADODB.Recordset
Public adoacc0y0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt212 As New ADODB.Recordset
'Dim dllaccrpt212 As Object 'Remove by Lydia 2017/10/18
Dim strAnd As String
Dim strSalesMan As String
Dim strSalesArea As String   '2007/11/12 add by sonia
Dim strSystem As String      '2007/11/14 add by sonia
Dim strAccSystem As String   '2007/11/14 add by sonia
Dim m_strSystem As String    '2007/11/14 add by sonia
'add by nickc 2008/01/18
Dim stST05 As String

'Add by Morgan 2010/5/19
Dim PLeft() As Integer, PColName() As String, PColName2() As String, strTemp() As String
Dim iPrint As Integer, iPage As Integer
Dim m_iStartX As Integer, m_iStartY As Integer
Dim m_iPageHeight As Integer, m_iLineHeight As Integer, m_iMargin As Integer
Dim strTitle As String, strCon1 As String, strCon2 As String
'Modified by Lydia 2017/01/04 點數改成小數點3位
'Const FDollar As String = "###,###,###.0"
Const FDollar As String = "###,###,###.000"
Const DDollar As String = "###,###,##0"
Dim m_bPrinter As Boolean, m_iPages As Integer, m_Device 'Add By Sindy 2010/9/16
Dim m_dbltotPoint As Double 'Add By Sindy 2010/10/5
Dim m_intPointCnt As Integer 'Add By Sindy 2010/12/22
Dim m_st52Yid As String 'Add By Sindy 2011/3/23 為第二級期限管制人
'Added by Lydia 2017/10/18 增加申請人/代理人之國籍/洲別
Dim strBaseTable As String, strCaseNa As String
Dim strCaseNaChk As String 'Added by Lydia 2020/09/01 檢查無申請人/代理人的案件
Dim strDateS As String, strDateE As String 'Add by Amy 2021/03/19

Private Sub Command2_Click()
Dim bolTmp As Boolean 'Added by Lydia 2017/10/18

   If FormCheck = False Then Exit Sub
   'add by nickc 2008/01/18
   '2009/12/30 MODIFY BY SONIA 開放洪琬姿80030可查全體FCT承辦人員
   'If stST05 >= "21" And stST05 <= "29" Then
   '2011/4/27 modify by sonia 再開放葉易雲78011可查全體FCT承辦人員
   'modify by sonia 2023/10/31 江協理同意再開放沈佳穎96003可查全體FCT承辦人員
   If (stST05 >= "21" And stST05 <= "29") And strUserNum <> "80030" And strUserNum <> "78011" And strUserNum <> "96003" Then
         If stST05 = "21" Or stST05 = "26" Or stST05 = "28" Then
            'Modify By Sindy 2018/11/29 FCT組別開放組長可以不輸入智權人員,查詢該組人員資料
            If Text11 = "4" Then
               strSql = "select st01,st02,st05,st16,st70,st52,st53,st54,st55" & _
                        " from staff where st03>='F10' and st03<='F19'" & _
                        " and nvl(st16,'')||nvl(st70,'') in(select nvl(st16,'')||nvl(st70,'') from staff where st01='" & strUserNum & "')" & _
                        " and st04='1'"
               intI = 1
               Text1 = "": Text1.Tag = ""
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  RsTemp.MoveFirst
                  Do While Not RsTemp.EOF
                     Text1.Tag = Text1.Tag & ",'" & RsTemp.Fields("st01") & "'"
                     RsTemp.MoveNext
                  Loop
                  If Text1.Tag <> "" Then Text1.Tag = Mid(Text1.Tag, 2)
               End If
            Else
               If Trim(Text1) = "" Then
                  MsgBox "智權人員不可以空白!!!", vbExclamation + vbOKOnly
                  Text1.SetFocus
                  Text1_GotFocus
                  Exit Sub
               End If
               If PUB_GetStaffST16(strUserNum) <> PUB_GetStaffST16(Text1) Then
                    MsgBox "僅可以輸入相同組別的智權人員！", vbExclamation, "操作錯誤！"
                    Text1.SetFocus
                    Text1_GotFocus
                    Exit Sub
               End If
            End If
         ElseIf Trim(Text1) = "" Then
            MsgBox "智權人員不可以空白!!!", vbExclamation + vbOKOnly
            Text1.SetFocus
            Text1_GotFocus
            Exit Sub
         Else
            'Add By Sindy 2011/3/23
            If m_st52Yid > "" Then
               'Modify By Sindy 2014/8/28
               'If PUB_GetST52(Text1.Text) <> m_st52Yid And Text1.Text <> m_st52Yid Then
               If PUB_GetST52(Text1.Text, m_st52Yid) = False And Text1.Text <> m_st52Yid Then
               '2014/8/28 END
                    'Modify By Sindy 2014/8/28
                    'MsgBox "非此智權人員的第二級期限管制人，無查詢權限！", vbExclamation, "操作錯誤！"
                    MsgBox "您無權限查詢此人資料！", vbExclamation, "操作錯誤！"
                    '2014/8/28 END
                    Text1.SetFocus
                    Text1_GotFocus
                    Exit Sub
               End If
            End If
            '2011/3/23 End
         End If
    End If
     
   'add by sonia 2023/10/31 江協理開放沈佳穎96003可查詢MCT點數
   If strUserNum = "96003" Then
      If Left(Text9, 2) = "F1" Then
         If Left(Text10, 2) <> "F1" Then
            MsgBox "國外商標處業務區只可為F10~F19 !!", vbCritical
            Text10.SetFocus
            Exit Sub
         End If
      ElseIf Left(Text9, 2) = "P2" Then
         If Left(Text10, 2) <> "P2" Then
            MsgBox "ＭＣＴ業務區只可為P20~P29 !!", vbCritical
            Text10.SetFocus
            Exit Sub
         End If
      Else
         MsgBox "您只可查詢國外商標處或ＭＣＴ業務區的資料 !!", vbCritical
         Text9.SetFocus
         Exit Sub
      End If
   End If
   'end 2023/10/31
   
   'Added by Lydia 2017/10/18 檢查國籍和洲別條件
   strBaseTable = "": strCaseNa = ""
   strCaseNaChk = "" 'Added by Lydia 2020/09/01
   If Trim(txtKind) <> "" And Trim(txtNa(0) & txtNa(1) & txtNArea) = "" Then
      MsgBox "請輸入國籍或洲別!!", vbCritical
      txtNa(0).SetFocus
      Exit Sub
   ElseIf Trim(txtKind) = "" And Trim(txtNa(0) & txtNa(1) & txtNArea) <> "" Then
      MsgBox "請輸入申請人或代理人!!", vbCritical
      txtKind.SetFocus
      Exit Sub
   End If
   If Trim(txtKind) <> "" Then
      bolTmp = False
      If Trim(txtNa(0) & txtNa(1)) <> "" Then
        If Trim(txtNa(0)) > Trim(txtNa(1)) And Trim(txtNa(1)) <> "" Then
           MsgBox "起始國籍不可大於終止國籍!", vbCritical
           Exit Sub
        End If
      End If
      If Trim(txtNArea) <> "" Then
        txtNArea_Validate bolTmp
        If bolTmp = True Then Exit Sub
      End If
      strBaseTable = ", (SELECT PA01 VC01, PA02 VC02, PA03 VC03, PA04 VC04, PA26 VC05, PA75 VC06 FROM PATENT " & _
                    "UNION SELECT TM01,TM02,TM03,TM04,TM23,TM44 FROM TRADEMARK " & _
                    "UNION SELECT LC01,LC02,LC03,LC04,LC11,LC22 FROM LAWCASE " & _
                    "UNION SELECT HC01,HC02,HC03,HC04,HC05,' ' FROM HIRECASE " & _
                    "UNION SELECT SP01,SP02,SP03,SP04,SP08,SP26 FROM SERVICEPRACTICE ) VT1 "
                    
      If Text8 = "1" Then '請款點數
         'modify by sonia 2018/8/3 改用ACC1K0之案號抓,否則請款單無收文號時會抓不到基本檔 X10710083
         'strCaseNa = " AND CP01=VC01(+) AND CP02=VC02(+) AND CP03=VC03(+) AND CP04=VC04(+)"
         strCaseNa = " AND A1K13=VC01(+) AND A1K14=VC02(+) AND A1K15=VC03(+) AND A1K16=VC04(+)"
      Else                '收款點數
         strCaseNa = " A1K13=VC01(+) AND A1K14=VC02(+) AND A1K15=VC03(+) AND A1K16=VC04(+)"
      End If
      
      If txtKind = "1" Then '申請人
          strBaseTable = strBaseTable & ", CUSTOMER, NATION "
          strCaseNa = strCaseNa & " AND SUBSTR(VC05,1,8)=CU01(+) AND SUBSTR(VC05,9,1)=CU02(+) AND SUBSTR(CU10,1,3)=NA01(+)"
      Else                  '代理人
          strBaseTable = strBaseTable & ", FAGENT, NATION "
          strCaseNa = strCaseNa & " AND SUBSTR(VC06,1,8)=FA01(+) AND SUBSTR(VC06,9,1)=FA02(+) AND SUBSTR(FA10,1,3)=NA01(+)"
      End If
      
      strCaseNaChk = strCaseNa & " AND " & IIf(txtKind = "1", "CU01 IS NULL", "FA01 IS NULL") 'Added by Lydia 2020/09/01 檢查無申請人/代理人的案件
      
      If Trim(txtNa(0)) <> "" Then
         If Trim(txtNa(0)) = "000" And txtNa(0) = txtNa(1) Then '臺灣地區
            strCaseNa = strCaseNa & " AND NA01>='000' AND NA01<='010'"
         Else
            strCaseNa = strCaseNa & " AND NA01>='" & Trim(txtNa(0)) & "'"
         End If
      End If
      If Trim(txtNa(1)) <> "" Then
         strCaseNa = strCaseNa & " AND NA01<='" & Trim(txtNa(1)) & "'"
      End If
      If Trim(txtNArea) <> "" Then
         If txtNArea = "0" Then
             '亞洲含臺灣,大陸
             strCaseNa = strCaseNa & " AND (SUBSTR(NA02,1,2) IN ('B0','C0') OR SUBSTR(NA02,1,1)='A')"
         Else
             strCaseNa = strCaseNa & " AND NA02='C" & txtNArea & "0'"
         End If
      End If
   End If
   'end 2017/10/18
   
   Screen.MousePointer = vbHourglass
   Accrpt212Delete
   m_intPointCnt = 0 'Add By Sindy 2010/12/22
   ProduceData
   If adoaccrpt212.State = adStateOpen Then
      adoaccrpt212.Close
   End If
   adoaccrpt212.CursorLocation = adUseClient
   '2010/9/13 modify by sonia
   'adoaccrpt212.Open "select * from accrpt212", adoTaie, adOpenStatic, adLockReadOnly
   adoaccrpt212.Open "select * from accrpt212 where R21201='" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccrpt212.RecordCount <> 0 Then
      InsertQueryLog (adoaccrpt212.RecordCount + m_intPointCnt) 'Add By Sindy 2010/12/22
      'Add By Sindy 2010/9/16
      If Text12 = "1" Then
         m_bPrinter = False
         Set m_Device = Picture1
         m_Device.AutoRedraw = True
         m_Device.Width = 11899
         m_Device.Height = 16838
         DelPic
      Else
         m_bPrinter = True
         Set m_Device = Printer
         m_Device.Orientation = 2
      End If
      '2010/9/16 End
      
      strCon1 = MaskEdBox1.Text
      strCon2 = MaskEdBox2.Text
   
      'Add by Morgan 2003/12/01
      '報表內容控制碼 明細='', 統計=chr(0)
      Dim strCtl As String
      'Modify By Sindy 2018/11/26 + Or Text11.Text = "4"
      If Text11.Text = "2" Or Text11.Text = "4" Then
         strCtl = Chr(0)
      End If
      'End 2003/12/01
      '2007/11/6 modify by sonia 表頭依業務區之部門,資料內容,報表內容決定
      Select Case Mid(Text9, 1, 2)
         Case "F2" '外專
            Select Case Text11
               Case "1"
                  Select Case Text8
                     Case "1"
                        'Modify by Morgan 2010/5/20
                        'dllaccrpt212.Acc24c0 ReportTitle(21211) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21211)
                        rptAcc24c0
                        
                     Case Else
                        'Modify by Sindy 2010/9/16
                        'dllaccrpt212.Acc24c1 ReportTitle(21212) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21212)
                        rptAcc24c0_2
                  End Select
                  
               
               Case "2"
                Select Case Text8
                     Case "1" '請款
                        'Modify by Morgan 2010/5/20
                        'dllaccrpt212.Acc24c0 ReportTitle(21213) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21213)
                        rptAcc24c0
                        
                     Case Else '收款
                        'Modify by Sindy 2010/9/16
                        'dllaccrpt212.Acc24c1 ReportTitle(21214) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21214)
                        rptAcc24c0_2
                  End Select
                  
               'Added by Morgan 2012/3/13 +3 技術語言別統計
               Case "3"
                  Select Case Text8
                     Case "1" '請款
                        strTitle = ReportTitle(21213)
                     Case Else '收款
                        strTitle = ReportTitle(21214)
                  End Select
                  rptAcc24c0_3
            End Select
            
         Case "F1" '外商
            Select Case Text11
               Case "1"
                  Select Case Text8
                     Case "1"
                        'Modify by Morgan 2010/5/20
                        'dllaccrpt212.Acc24c0 ReportTitle(21221) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21221)
                        rptAcc24c0
                        
                     Case Else
                        'Modify by Sindy 2010/9/16
                        'dllaccrpt212.Acc24c1 ReportTitle(21222) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21222)
                        rptAcc24c0_2
                  End Select
               'Modify By Sindy 2018/11/29
               Case "2", "4"
                  Select Case Text8
                     Case "1"
                        'Modify by Morgan 2010/5/20
                        'dllaccrpt212.Acc24c0 ReportTitle(21223) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21223)
                        rptAcc24c0
                        
                     Case Else
                        'Modify by Sindy 2010/9/16
                        'dllaccrpt212.Acc24c1 ReportTitle(21224) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21224)
                        rptAcc24c0_2
                  End Select
            End Select
         Case "F3"
            Select Case Text11
               Case "1"
                  Select Case Text8
                     Case "1"
                        'Modify by Morgan 2010/5/20
                        'dllaccrpt212.Acc24c0 ReportTitle(21231) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21231)
                        rptAcc24c0
                        
                     Case Else
                        'Modify by Sindy 2010/9/16
                        'dllaccrpt212.Acc24c1 ReportTitle(21232) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21232)
                        rptAcc24c0_2
                  End Select
               Case "2"
                  Select Case Text8
                     Case "1"
                        'Modify by Morgan 2010/5/20
                        'dllaccrpt212.Acc24c0 ReportTitle(21233) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21233)
                        rptAcc24c0
                        
                     Case Else
                        'Modify by Sindy 2010/9/16
                        'dllaccrpt212.Acc24c1 ReportTitle(21234) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21234)
                        rptAcc24c0_2
                  End Select
            End Select
         Case Else
            Select Case Text11
               Case "1"
                  Select Case Text8
                     Case "1"
                        'Modify by Morgan 2010/5/20
                        'dllaccrpt212.Acc24c0 ReportTitle(21241) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21241)
                        rptAcc24c0
                        
                     Case Else
                        'Modify by Sindy 2010/9/16
                        'dllaccrpt212.Acc24c1 ReportTitle(21242) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21242)
                        rptAcc24c0_2
                  End Select
               Case "2"
                  Select Case Text8
                     Case "1"
                        'Modify by Morgan 2010/5/20
                        'dllaccrpt212.Acc24c0 ReportTitle(21243) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21243)
                        rptAcc24c0
                        
                     Case Else
                        'Modify by Sindy 2010/9/16
                        'dllaccrpt212.Acc24c1 ReportTitle(21244) & strCtl, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate)), m_strSystem
                        strTitle = ReportTitle(21244)
                        rptAcc24c0_2
                  End Select
            End Select
      End Select
      '2007/11/6 end
      
      'Add By Sindy 2010/9/16
      If m_bPrinter = True Then
         m_Device.EndDoc
         ShowPrintOk
      ElseIf m_iPages > 0 Then
         SetPic m_iPages
         Frmacc24c0_1.m_ImageW = m_Device.Width
         Frmacc24c0_1.m_ImageH = m_Device.Height
         Frmacc24c0_1.m_iPages = m_iPages
         Frmacc24c0_1.Show
      End If
      '2010/9/16 End
      m_Device.DrawWidth = 1  '2010/12/1 add by sonia
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/22
   End If
   adoaccrpt212.Close
   If strCon10 <> MsgText(602) Then
      '2007/11/6 modify by sonia 執行完不清除條件
      'FormClear
      'MsgBox "列印結束 !", vbInformation
      '2007/11/6 end
   End If
   Screen.MousePointer = vbDefault
   StatusView MsgText(101)
   
   'Added by Lydia 2020/09/01 因為外商常有案件無申請人/代理人,所以彈提醒
   If strCaseNaChk <> "" Then
       If Text8.Text = "1" Then '請款點數：列出無申請人/代理人的清單
          Call ProcExcelSave(strCaseNaChk)
       Else
          MsgBox strCaseNaChk, vbExclamation, "檢核資料"
       End If
   End If
   'end 2020/09/01
   
End Sub

Private Sub Form_Activate()
   '93.3.16 ADD BY SONIA
   'Modify by Amy 2021/03/04 原:mdiMain,因acccount也要用
   If IsObject(Forms(0)) Then
      ToolShow
   End If
   '93.3.16 END
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      StatusView MsgText(101)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   'add by nickc 2008/01/18
   stST05 = PUB_GetST05(strUserNum)
   
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modified by Lydia 2017/10/18 改共用模組　W:6675 H:5750
'   Me.Width = 6285
'   Me.Height = 4395
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath4)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   PUB_InitForm Me, 6870, 5750, strBackPicPath4
   'end 2017/10/18
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Text9.MaxLength = 3
   Text10.MaxLength = 3
   Text11.MaxLength = 1
   StatusView MsgText(101)
   'Set dllaccrpt212 = CreateObject("AccReport.ReportSelect") 'Remove by Lydia 2017/01/04
   
   'add by nickc 2008/01/18
   If stST05 = "21" Or stST05 = "26" Or stST05 = "28" Then
        Text9 = "F10"
        Text10 = "F11"
        Text1 = strUserNum
        Label2.Caption = GetStaffName(Text1, True)
        Text9.Enabled = False
        Text10.Enabled = False
        'add by sonia 2023/10/31 江協理開放沈佳穎96003可查詢MCT點數
        If strUserNum = "96003" Then
           Text9.Enabled = True
           Text10.Enabled = True
        End If
        'end 2023/10/31
   ElseIf stST05 >= "21" And stST05 <= "29" Then
        Text9 = "F10"
        Text10 = "F11"
        Text1 = strUserNum
        Label2.Caption = GetStaffName(Text1, True)
        Text9.Enabled = False
        Text10.Enabled = False
        'Add By Sindy 2011/3/23 若是第二級期限管制人可以輸入其智權人員
        strSql = "SELECT st52 FROM staff WHERE st52='" & strUserNum & "' and st04='1' "
        intI = 1
        m_st52Yid = ""
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
            m_st52Yid = strUserNum
            Text1.Enabled = True
        '2011/3/23 End
        Else
            Text1.Enabled = False
        End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   'Set dllaccrpt212 = Nothing 'Remove by Lydia 2017/01/04
   Set Frmacc24c0 = Nothing
End Sub

'Add by Amy 2021/12/01 起日輸1號,迄日預帶當月底-秀玲
Private Sub MaskEdBox1_LostFocus()
    If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then Exit Sub
    If Right(Val(FCDate(MaskEdBox1.Text)), 2) = "01" Then
        MaskEdBox2.Text = CFDate(ACDate(GetLastDay(DBDATE(FCDate(MaskEdBox1)))))
    End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'add by nickc 2008/01/18
Private Sub Text1_LostFocus()
If stST05 >= "21" And stST05 <= "29" Then
    Label2.Caption = GetStaffName(Text1, True)
End If
End Sub

'add by nickc 2008/01/18
Private Sub Text1_Validate(Cancel As Boolean)
If Trim(Text1) <> "" Then
    If stST05 >= "21" And stST05 <= "29" Then
        Label2.Caption = GetStaffName(Text1, True)
        '2009/12/30 MODIFY BY SONIA 開放洪琬姿80030可查全體FCT承辦人員
        'If stST05 = "21" Or stST05 = "26" Or stST05 = "28" Then
        '2011/4/27 modify by sonia 再開放葉易雲78011可查全體FCT承辦人員
        'modify by sonia 2023/10/31 江協理同意再開放沈佳穎96003可查全體FCT承辦人員
        If (stST05 = "21" Or stST05 = "26" Or stST05 = "28") And strUserNum <> "80030" And strUserNum <> "78011" And strUserNum <> "96003" Then
             If PUB_GetStaffST16(strUserNum) <> PUB_GetStaffST16(Text1) Then
                  MsgBox "僅可以輸入相同組別的智權人員！", vbExclamation, "操作錯誤！"
                  Text1.SetFocus
                  Text1_GotFocus
                  Cancel = True
                  Exit Sub
             End If
        End If
    End If
End If
End Sub

Private Sub Text10_GotFocus()
   If Text9 <> "" Then Text10 = Text9
   TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If (Text9.Text <> "") Then
      If (Text10.Text < Text9.Text) Then
         MsgBox "業務區迄值不可小於起值！", vbCritical
         Cancel = True
      End If
   End If
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

'Add by Morgan 2003/12/01
Private Sub Text11_KeyPress(KeyAscii As Integer)
   'Modified by Morgan 2012/3/14 +3
   'Modify by Sindy 2018/11/23 +4
   If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If (Text11 = "") Then
      MsgBox "請輸入報表類別！", vbCritical
      Cancel = True
   End If
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If (Text12 = "") Then
      MsgBox "請輸入輸出方式！", vbCritical
      Cancel = True
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'2007/12/10 add by sonia
Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 <> MsgText(601) Then
      If CheckSys(Text2) = "" Then
         MsgBox MsgText(1107), , MsgText(5)
         Cancel = True
         Text2_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 <> MsgText(601) Then
      If CheckSys(Text3) = "" Then
         MsgBox MsgText(1107), , MsgText(5)
         Cancel = True
         Text3_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text4 <> MsgText(601) Then
      If CheckSys(Text4) = "" Then
         MsgBox MsgText(1107), , MsgText(5)
         Cancel = True
         Text4_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> MsgText(601) Then
      If CheckSys(Text5) = "" Then
         MsgBox MsgText(1107), , MsgText(5)
         Cancel = True
         Text5_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 <> MsgText(601) Then
      If CheckSys(Text6) = "" Then
         MsgBox MsgText(1107), , MsgText(5)
         Cancel = True
         Text6_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 <> MsgText(601) Then
      If CheckSys(Text7) = "" Then
         MsgBox MsgText(1107), , MsgText(5)
         Cancel = True
         Text7_GotFocus
         Exit Sub
      End If
   End If
End Sub
'2007/12/10 end

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

'Add by Sindy 2010/9/16
Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
      KeyAscii = 0
   End If
End Sub

'Add by Sindy 2010/9/16
Private Sub Text8_Validate(Cancel As Boolean)
   If (Text8 = "") Then
      MsgBox "請輸入資料類別！", vbCritical
      Cancel = True
   End If
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
   
'On Error GoTo Checking
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/22 清除查詢印表記錄檔欄位
   StatusView MsgText(26)
   adoaccrpt212.CursorLocation = adUseClient
   '2010/9/13 modify by sonia
   'adoaccrpt212.Open "select * from accrpt212", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoaccrpt212.Open "select * from accrpt212 where R21201='" & strUserNum & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   strSalesMan = MsgText(601)
   strSalesArea = MsgText(601) '2007/11/12 add by sonia
   strSystem = MsgText(601)    '2007/11/14 add by sonia
   strAccSystem = MsgText(601) '2007/11/14 add by sonia
   m_strSystem = MsgText(601)  '2007/11/14 add by sonia
   strDateS = "": strDateE = "" 'Add by Amy 2021/03/19
   
   'Modify By Sindy 2018/11/29 FCT組別開放組長可以不輸入智權人員,查詢該組人員資料
   If Text11 = "4" And Text1.Tag <> "" Then
      strSalesMan = strSalesMan & " and cp13 in(" & Text1.Tag & ")"
      pub_QL05 = pub_QL05 & ";" & Label1(0) & Replace(Text1.Tag, "'", "")
   '2018/11/29 END
   ElseIf Text1 <> "" Then
      strSalesMan = strSalesMan & " and cp13 = '" & Text1 & "'"
      pub_QL05 = pub_QL05 & ";" & Label1(0) & Text1 'Add By Sindy 2010/12/22
   End If
   
   'Add by Morgan 2003/12/01
   If Text9 <> "" Then
      strSalesArea = strSalesArea & " and cp12 >= '" & Text9 & "'"
   End If
   If Text10 <> "" Then
      strSalesArea = strSalesArea & " and cp12 <= '" & Text10 & "'"
   End If
   'End 2003/12/01
   If Text9 <> "" Or Text10 <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label8 & Text9 & "-" & Text10 'Add By Sindy 2010/12/22
   End If
   
   If Text2 <> MsgText(601) Or Text3 <> MsgText(601) Or Text4 <> MsgText(601) Or Text5 <> MsgText(601) Or Text6 <> MsgText(601) Or Text7 <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";" & Label3 & Text2 'Add By Sindy 2010/12/22
      strSystem = strSystem & " and ("
      strAccSystem = strAccSystem & " and ("
      If Text2 <> MsgText(601) Then
         strSystem = strSystem & "a1k13 = '" & Text2 & "' or "
         strAccSystem = strAccSystem & "substr(ax214, 1, Length(ax214) - 9) = '" & Text2 & "' or "
         m_strSystem = Text2
      End If
      If Text3 <> MsgText(601) Then
         strSystem = strSystem & "a1k13 = '" & Text3 & "' or "
         strAccSystem = strAccSystem & "substr(ax214, 1, Length(ax214) - 9) = '" & Text3 & "' or "
         m_strSystem = m_strSystem & "," & Text3
         pub_QL05 = pub_QL05 & "," & Text3 'Add By Sindy 2010/12/22
      End If
      If Text4 <> MsgText(601) Then
         strSystem = strSystem & "a1k13 = '" & Text4 & "' or "
         strAccSystem = strAccSystem & "substr(ax214, 1, Length(ax214) - 9) = '" & Text4 & "' or "
         m_strSystem = m_strSystem & "," & Text4
         pub_QL05 = pub_QL05 & "," & Text4 'Add By Sindy 2010/12/22
      End If
      If Text5 <> MsgText(601) Then
         strSystem = strSystem & "a1k13 = '" & Text5 & "' or "
         strAccSystem = strAccSystem & "substr(ax214, 1, Length(ax214) - 9) = '" & Text5 & "' or "
         m_strSystem = m_strSystem & "," & Text5
         pub_QL05 = pub_QL05 & "," & Text5 'Add By Sindy 2010/12/22
      End If
      If Text6 <> MsgText(601) Then
         strSystem = strSystem & "a1k13 = '" & Text6 & "' or "
         strAccSystem = strAccSystem & "substr(ax214, 1, Length(ax214) - 9) = '" & Text6 & "' or "
         m_strSystem = m_strSystem & "," & Text6
         pub_QL05 = pub_QL05 & "," & Text6 'Add By Sindy 2010/12/22
      End If
      If Text7 <> MsgText(601) Then
         strSystem = strSystem & "a1k13 = '" & Text7 & "' or "
         strAccSystem = strAccSystem & "substr(ax214, 1, Length(ax214) - 9) = '" & Text7 & "' or "
         m_strSystem = m_strSystem & "," & Text7
         pub_QL05 = pub_QL05 & "," & Text7 'Add By Sindy 2010/12/22
      End If
      strSystem = Mid(strSystem, 1, Len(strSystem) - 4) & ") "
      strAccSystem = Mid(strAccSystem, 1, Len(strAccSystem) - 4) & ") "
   End If
   If Trim(Text8) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label5 & "1.請款點數" 'Add By Sindy 2010/12/22
   Else
      pub_QL05 = pub_QL05 & ";" & Label5 & "2.收款點數" 'Add By Sindy 2010/12/22
   End If
   If Trim(Text11) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label10 & "1.明細" 'Add By Sindy 2010/12/22
   Else
      pub_QL05 = pub_QL05 & ";" & Label10 & "2.統計" 'Add By Sindy 2010/12/22
   End If
   If Trim(Text12) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1(9), 5) & "1.螢幕" 'Add By Sindy 2010/12/22
   Else
      pub_QL05 = pub_QL05 & ";" & Left(Label1(9), 5) & "2.印表機" 'Add By Sindy 2010/12/22
   End If
   
   'Added by Lydia 2017/10/18
   If Trim(txtKind) <> "" Then
      pub_QL05 = pub_QL05 & ";類別：" & IIf(txtKind = "1", "申請人", "代理人")
   End If
   If Trim(txtNa(0) & txtNa(1)) <> "" Then
      pub_QL05 = pub_QL05 & ";國籍：" & Trim(txtNa(0)) & "-" & Trim(txtNa(1))
   End If
   If Trim(txtNArea) <> "" Then
      pub_QL05 = pub_QL05 & ";洲別：" & Trim(txtNa(0)) & "-" & Trim(txtNa(1))
   End If
   'end 2017/10/18
   
   Select Case Text8
      'Modified by Lydia 2020/09/01
      'Case Mid(ComboItem(1), 1, 1)
      Case "1" '請款點數
         Select1
         If Text11 <> "3" Then
            AddSharePoint 'Add by Morgan 2010/5/20
         End If
      'Modified by Lydia 2020/09/01
      'Case Mid(ComboItem(2), 1, 1)
      Case "2" '收款點數
         Select2
   End Select
   adoaccrpt212.Close
   
   'Added by Morgan 2012/3/14 更新FCP工程師組別欄位
   If Text11 = "3" Then
      strSql = "update accrpt212 set R21215=(select pa150 from patent where pa01(+)=substrb(R21203,1,length(R21203)-12)" & _
         " and pa02(+)=substrb(R21203,-11,6) and pa03(+)=substrb(R21203,-4,1) and pa04(+)=substrb(R21203,-2))" & _
         " where r21201='" & strUserNum & "' and substrb(R21203,1,length(R21203)-12) in ('P','FCP','CFP')"
      cnnConnection.Execute strSql, intI
      
      strSql = "update accrpt212 set R21215=(select sp79 from servicepractice where sp01(+)=substrb(R21203,1,length(R21203)-12)" & _
         " and sp02(+)=substrb(R21203,-11,6) and sp03(+)=substrb(R21203,-4,1) and sp04(+)=substrb(R21203,-2))" & _
         " where r21201='" & strUserNum & "' and substrb(R21203,1,length(R21203)-12) in ('PS','FG','CPS')"
      cnnConnection.Execute strSql, intI
   End If
   'end 2012/3/14
   
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt212Delete()
   '2010/9/13 MODIFY BY SONIA
   'adoTaie.Execute "delete from accrpt212"
   'Modify By Sindy 2010/9/16
   adoTaie.Execute "delete from accrpt212 where R21201='" & strUserNum & "'"
End Sub

'*************************************************
'  選擇請款點數統計
'
'*************************************************
Private Sub Select1()
   Dim douExchange As Double
   Dim strName As String

   strExc(1) = ""
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strExc(1) = strExc(1) & " and a1k02 >= " & Val(ChangeTDateStringToTString(Replace(MaskEdBox1.Text, "_", ""))) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strExc(1) = strExc(1) & " and a1k02 <= " & Val(ChangeTDateStringToTString(Replace(MaskEdBox2.Text, "_", ""))) & ""
   End If
   If (MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29)) Or _
      (MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29)) Then
      pub_QL05 = pub_QL05 & ";" & Label4 & MaskEdBox1 & "-" & MaskEdBox2 'Add By Sindy 2010/12/22
   End If
   
   strCon10 = ""
   strName = ""
   adoacc1k0.CursorLocation = adUseClient
   '2007/11/15 modify by sonia 同一請款單計入最後收文之智權人員,作廢或銷帳都不抓,請款金額抓台幣-折讓金額*匯率
   'adoacc1k0.Open "select * from (select distinct cp01, cp13, cp60, cp10,CP12 from caseprogress, acc1k0 where cp60 = a1k01" & strSalesMan & strExc(1) & ") new, acc1k0 where cp60 = a1k01" & strSalesMan & strExc(1) & " order by cp13, a1k13, a1k14, a1k15, a1k16, a1k02, decode(cp10, '201', 1, 2)", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Lydia 2017/10/18 增加申請人/代理人之國籍/洲別
   'adoacc1k0.Open "select * from caseprogress, (select max(cp05||cp09) cp,a1k01,a1k13,a1k14,a1k15,a1k16,a1k02,a1k06,a1k11,a1k09,a1k10 from caseprogress, acc1k0 where cp60(+) = a1k01 and a1k12 is null and a1k25 is null " & strSalesArea & strSystem & strExc(1) & " group by a1k01,a1k13,a1k14,a1k15,a1k16,a1k02,a1k06,a1k11,a1k09,a1k10) new " & _
                  "where cp09 in substr(new.cp,9,9) " & strSalesMan & " order by cp13, a1k02, a1k13, a1k14, a1k15, a1k16", adoTaie, adOpenStatic, adLockReadOnly
   strSql = "select cp01,cp10,cp13,n1.* " & IIf(strBaseTable <> "", ", vt1.* ", "") & _
            "from caseprogress, (select max(cp05||cp09) cp,a1k01,a1k13,a1k14,a1k15,a1k16,a1k02,a1k06,a1k11,a1k09,a1k10 from caseprogress, acc1k0 where cp60(+) = a1k01 and a1k12 is null and a1k25 is null " & strSalesArea & strSystem & strExc(1) & " group by a1k01,a1k13,a1k14,a1k15,a1k16,a1k02,a1k06,a1k11,a1k09,a1k10) n1 " & _
            strBaseTable & " where cp09 in substr(n1.cp,9,9) " & strSalesMan & strCaseNa & " order by cp13, a1k02, a1k13, a1k14, a1k15, a1k16 "
   'Added by Lydia 2020/09/01 檢查無申請人/代理人的案件
   If txtKind <> "" Then
        strExc(0) = "SELECT CP01,CP01||'-'||CP02||DECODE(CP03,'-'||'0',NULL)||DECODE(CP04,'-'||'00',NULL) AS CASENO,SQLDATET(CP05) CP05T, CP09, CP60,ROUND((NVL(A1K11,0)-NVL(A1K09,0))/1000,3) DOT " & _
                         "from caseprogress, (select max(cp05||cp09) cp,a1k01,a1k13,a1k14,a1k15,a1k16,a1k02,a1k06,a1k11,a1k09,a1k10 from caseprogress, acc1k0 where cp60(+) = a1k01 and a1k12 is null and a1k25 is null " & strSalesArea & strSystem & strExc(1) & " group by a1k01,a1k13,a1k14,a1k15,a1k16,a1k02,a1k06,a1k11,a1k09,a1k10) n1 " & _
                         strBaseTable & " where cp09 in substr(n1.cp,9,9) " & strSalesMan & strCaseNaChk & " order by cp01 desc,cp05,cp09,cp60  "
        strCaseNaChk = strExc(0)
   End If
   'end 2020/09/01
   
   adoacc1k0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc1k0.RecordCount = 0 Then
      strCon10 = MsgText(602)
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      strCaseNaChk = "" 'Added by Lydia 2020/09/01
      Exit Sub
   'Added by Lydia 2020/10/05 (補)因為外商常有案件無申請人/代理人,所以彈提醒
   ElseIf strCaseNaChk <> "" Then
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strCaseNaChk)
       If intI = 0 Then
            strCaseNaChk = ""
       Else
            strCaseNaChk = "有 " & RsTemp.RecordCount & " 筆點數，案件無" & IIf(txtKind = "1", "申請人", "代理人") & "！"
       End If
   'end 2020/10/05
   End If

   Do While adoacc1k0.EOF = False
      '2007/11/15 cancel by sonia 因同日同案號有二張請款單960929FCP030989
      'If strName <> (adoacc1k0.Fields("cp13").Value & adoacc1k0.Fields("a1k13").Value & adoacc1k0.Fields("a1k14").Value & adoacc1k0.Fields("a1k15").Value & adoacc1k0.Fields("a1k16").Value & adoacc1k0.Fields("a1k02").Value) Then
      '2007/11/15 END
         adoaccrpt212.AddNew
         'Add by Morgan 2003/12/02 請款單號(R21213)
         adoaccrpt212.Fields("r21213").Value = adoacc1k0.Fields("A1k01").Value
         'End 2003/12/02
         adoaccrpt212.Fields("r21201").Value = strUserNum
         'Memo by Lydia 2018/11/05 智權人員編號(R21202)和名稱(R21211)
         If IsNull(adoacc1k0.Fields("cp13").Value) Then
            adoaccrpt212.Fields("r21202").Value = Null
            adoaccrpt212.Fields("r21211").Value = Null
         Else
            adoaccrpt212.Fields("r21202").Value = adoacc1k0.Fields("cp13").Value
            adoaccrpt212.Fields("r21211").Value = StaffQuery(adoacc1k0.Fields("cp13").Value)
         End If
         'Memo by Lydia 2018/11/05 本所案號(R21203)
         If IsNull(adoacc1k0.Fields("cp01").Value) Then
            adoaccrpt212.Fields("r21203").Value = Null
         Else
            adoaccrpt212.Fields("r21203").Value = adoacc1k0.Fields("a1k13").Value
            If IsNull(adoacc1k0.Fields("a1k14").Value) = False Then
               adoaccrpt212.Fields("r21203").Value = adoaccrpt212.Fields("r21203").Value & "-" & adoacc1k0.Fields("a1k14").Value
            End If
            If IsNull(adoacc1k0.Fields("a1k15").Value) = False Then
               adoaccrpt212.Fields("r21203").Value = adoaccrpt212.Fields("r21203").Value & "-" & adoacc1k0.Fields("a1k15").Value
            End If
            If IsNull(adoacc1k0.Fields("a1k16").Value) = False Then
               adoaccrpt212.Fields("r21203").Value = adoaccrpt212.Fields("r21203").Value & "-" & adoacc1k0.Fields("a1k16").Value
            End If
         End If
         'Memo by Lydia 2018/11/05 請款日期(R21204)
         If IsNull(adoacc1k0.Fields("a1k02").Value) Then
            adoaccrpt212.Fields("r21204").Value = Null
         Else
            adoaccrpt212.Fields("r21204").Value = adoacc1k0.Fields("a1k02").Value
         End If
         If IsNull(adoacc1k0.Fields("a1k10").Value) Then
            douExchange = 0
         Else
            douExchange = adoacc1k0.Fields("a1k10").Value
         End If
         'Memo by Lydia 2018/11/05 請款金額(R21205)
         If IsNull(adoacc1k0.Fields("a1k11").Value) Then
            adoaccrpt212.Fields("r21205").Value = 0
         Else
            If IsNull(adoacc1k0.Fields("a1k06").Value) Then
               '2007/11/15 modify by sonia
               'adoaccrpt212.Fields("r21205").Value = Val(Format(Val(adoacc1k0.Fields("a1k08").Value) * douExchange, DAmount))
               adoaccrpt212.Fields("r21205").Value = Val(Format(Val(adoacc1k0.Fields("a1k11").Value), DAmount))
            Else
               '2007/11/15 modify by sonia
               'adoaccrpt212.Fields("r21205").Value = Val(Format((Val(adoacc1k0.Fields("a1k08").Value) - Val(adoacc1k0.Fields("a1k06").Value)) * douExchange, DAmount))
               'Modify By Sindy 2013/1/15
               'adoaccrpt212.Fields("r21205").Value = Val(Format(Val(adoacc1k0.Fields("a1k11").Value) - (Val(adoacc1k0.Fields("a1k06").Value) * douExchange), DAmount))
               adoaccrpt212.Fields("r21205").Value = Val(Format(Val(adoacc1k0.Fields("a1k11").Value) - Val(adoacc1k0.Fields("a1k06").Value), DAmount))
               '2013/1/15 End
            End If
         End If
         'Memo by Lydia 2018/11/05 規費(R21206)
         If IsNull(adoacc1k0.Fields("a1k09").Value) Then
            adoaccrpt212.Fields("r21206").Value = 0
         Else
            adoaccrpt212.Fields("r21206").Value = Val(adoacc1k0.Fields("a1k09").Value)
         End If
        'Modify By Cheng 2003/12/16
        '取小數點一位
'         adoaccrpt212.Fields("r21207").Value = (Val(Format(adoaccrpt212.Fields("r21205").Value, DAmount)) - Val(Format(adoaccrpt212.Fields("r21206").Value, DAmount))) / 1000
         'Modified by Lydia 2017/01/04 取小數3位
         'adoaccrpt212.Fields("r21207").Value = Format((Val(Format(adoaccrpt212.Fields("r21205").Value, DAmount)) - Val(Format(adoaccrpt212.Fields("r21206").Value, DAmount))) / 1000, "0.0")
         'Memo by Lydia 2018/11/05 請款點數(R21207)
         adoaccrpt212.Fields("r21207").Value = (Val(Format(adoaccrpt212.Fields("r21205").Value, DAmount)) - Val(Format(adoaccrpt212.Fields("r21206").Value, DAmount))) / 1000
        'End
         'Memo by Lydia 2018/11/05 已收金額(R21208)
         adoaccsum.CursorLocation = adUseClient
         adoaccsum.Open "select sum(a0z04 * a0y04) from acc0z0, acc0y0 where a0z01 = a0y01 and a0z02 = '" & adoacc1k0.Fields("a1k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoaccsum.RecordCount <> 0 Then
            If IsNull(adoaccsum.Fields(0).Value) Then
               adoaccrpt212.Fields("r21208").Value = 0
            Else
               adoaccrpt212.Fields("r21208").Value = Val(adoaccsum.Fields(0).Value)
            End If
         Else
            adoaccrpt212.Fields("r21208").Value = 0
         End If
         adoaccsum.Close
         'Memo by Lydia 2018/11/05 已收點數(R21209)
         If adoaccrpt212.Fields("r21208").Value = 0 Then
            adoaccrpt212.Fields("r21209").Value = 0
         Else
            adoaccrpt212.Fields("r21209").Value = Val(Format(Val(Format(Val(adoaccrpt212.Fields("r21208").Value) - Val(adoaccrpt212.Fields("r21206").Value), DAmount)) / 1000, FDollar))
         End If
         '2007/11/15 modify by sonia 同一請款單計入最後收文之智權人員,故改以請款單抓進度檔是否有201,927者
'         If adoacc1k0.Fields("cp10").Value = "201" Then
'            adoaccsum.CursorLocation = adUseClient
'      '      adoaccsum.Open "select * from acc1p0 where a1p18 in (select min(a1p18) from acc1p0 where a1p01 = '1' and a1p17 = '" & adoacc1k0.Fields("a1k13").Value & adoacc1k0.Fields("a1k14").Value & adoacc1k0.Fields("a1k15").Value & adoacc1k0.Fields("a1k16").Value & "' and a1p05 = '6130') and a1p01 = '1' and a1p17 = '" & adoacc1k0.Fields("a1k13").Value & adoacc1k0.Fields("a1k14").Value & adoacc1k0.Fields("a1k15").Value & adoacc1k0.Fields("a1k16").Value & "' and a1p05 = '6130'", adoTaie, adOpenStatic, adLockReadOnly
'            adoaccsum.Open "select ax206 from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and a0205 in (select min(a0205) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax214 = '" & adoacc1k0.Fields("a1k13").Value & adoacc1k0.Fields("a1k14").Value & adoacc1k0.Fields("a1k15").Value & adoacc1k0.Fields("a1k16").Value & "' and ax205 = '6130') and ax214 = '" & adoacc1k0.Fields("a1k13").Value & adoacc1k0.Fields("a1k14").Value & adoacc1k0.Fields("a1k15").Value & adoacc1k0.Fields("a1k16").Value & "' and ax205 = '6130'", adoTaie, adOpenStatic, adLockReadOnly
'            If adoaccsum.RecordCount <> 0 Then
'               If IsNull(adoaccsum.Fields(0).Value) Then
'                  adoaccrpt212.Fields("r21210").Value = 0
'               Else
'                  adoaccrpt212.Fields("r21210").Value = adoaccsum.Fields(0).Value
'               End If
'            Else
'               adoaccrpt212.Fields("r21210").Value = 0
'            End If
'            adoaccsum.Close
'         Else
'            adoaccrpt212.Fields("r21210").Value = 0
'         End If
         adoaccrpt212.Fields("r21210").Value = 0
         'Memo by Lydia 2018/11/05 支付翻譯費(R21210)
         If Not IsNull(adoacc1k0.Fields("A1k01").Value) Then
            adoaccsum.Open "select ax206 from acc021, acc020, caseprogress where ax201 = a0201 and ax202 = a0202 and a0205 in (select min(a0205) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax214 = '" & adoacc1k0.Fields("a1k13").Value & adoacc1k0.Fields("a1k14").Value & adoacc1k0.Fields("a1k15").Value & adoacc1k0.Fields("a1k16").Value & "' and ax205 = '6130') and ax214 = '" & adoacc1k0.Fields("a1k13").Value & adoacc1k0.Fields("a1k14").Value & adoacc1k0.Fields("a1k15").Value & adoacc1k0.Fields("a1k16").Value & "' and ax205 = '6130' " & _
                           "and '" & adoacc1k0.Fields("A1k01").Value & "'=cp60(+) and (cp10='201' or cp10='927') ", adoTaie, adOpenStatic, adLockReadOnly
            If adoaccsum.RecordCount <> 0 Then
               If Not IsNull(adoaccsum.Fields(0).Value) Then
                  adoaccrpt212.Fields("r21210").Value = adoaccsum.Fields(0).Value
               End If
            End If
            adoaccsum.Close
         End If
         '2007/11/15 end
         adoaccrpt212.UpdateBatch
         strName = (adoacc1k0.Fields("cp13").Value & adoacc1k0.Fields("a1k13").Value & adoacc1k0.Fields("a1k14").Value & adoacc1k0.Fields("a1k15").Value & adoacc1k0.Fields("a1k16").Value & adoacc1k0.Fields("a1k02").Value)
      'End If    '2007/11/15 cancel by sonia
      adoacc1k0.MoveNext
   Loop
   adoacc1k0.Close
End Sub

'*************************************************
'  選擇收款點數統計
'
'*************************************************
Private Sub Select2()
Dim douExchange As Double
Dim strSql As String
Dim strSQL1 As String      '2007/11/12 add by sonia
Dim strSQL2 As String      '2013/5/30  add by sonia
Dim StrSQL3 As String      '2013/5/30  add by sonia
Dim straccSales As String  '2007/11/12 add by sonia
Dim str1P0Sales As String  '2007/11/14 add by sonia

'Modify by Amy 2021/03/19 程式改寫至 Pub_GetAccRecePayAmt
'   strSql = "": strSQL1 = "": strSQL2 = "": StrSQL3 = "": strAccSales = "": str1P0Sales = ""
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = " and a0y02 >= " & Val(ChangeTDateStringToTString(Replace(MaskEdBox1.Text, "_", ""))) & ""
'      strSQL1 = " and a0205 >= " & Val(ChangeTDateStringToTString(Replace(MaskEdBox1.Text, "_", ""))) & ""
'      strSQL2 = " a1h02 >= " & Val(ChangeTDateStringToTString(Replace(MaskEdBox1.Text, "_", ""))) & ""    '2013/5/30 add by sonia
'      StrSQL3 = " a1p18 >= " & Val(ChangeTDateStringToTString(Replace(MaskEdBox1.Text, "_", ""))) & ""    '2013/5/30 add by sonia
       strDateS = ChangeTDateStringToTString(Replace(MaskEdBox1.Text, "_", ""))
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a0y02 <= " & Val(ChangeTDateStringToTString(Replace(MaskEdBox2.Text, "_", ""))) & " "
'      strSQL1 = strSQL1 & " and a0205 <= " & Val(ChangeTDateStringToTString(Replace(MaskEdBox2.Text, "_", ""))) & " "
'      strSQL2 = strSQL2 & " and a1h02 <= " & Val(ChangeTDateStringToTString(Replace(MaskEdBox2.Text, "_", ""))) & " "    '2013/5/30 add by sonia
'      StrSQL3 = StrSQL3 & " and a1p18 <= " & Val(ChangeTDateStringToTString(Replace(MaskEdBox2.Text, "_", ""))) & " "    '2013/5/30 add by sonia
       strDateE = ChangeTDateStringToTString(Replace(MaskEdBox2.Text, "_", ""))
   End If
   If (MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29)) Or _
      (MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29)) Then
      pub_QL05 = pub_QL05 & ";" & Label4 & MaskEdBox1 & "-" & MaskEdBox2 'Add By Sindy 2010/12/22
   End If
'
'   '2007/11/12 add by sonia 非個人時再加印財務調整傳票
'   If Text1 = "" Then
'      Select Case Mid(Text9, 1, 2)
'         Case "F3"
'            strAccSales = " and ax209='F4101' "
'            str1P0Sales = " and a1p16='F4101' "
'         Case "F2"
'            'modify by sonia 2021/1/14 +F4104,F4105
'            strAccSales = " and ax209 in ('F4102','F4104','F4105') "
'            str1P0Sales = " and a1p16 in ('F4102','F4104','F4105') "
'         Case "F1"
'            'modify by sonia 2021/1/14 +F4106,F4107
'            strAccSales = " and ax209 in ('F4103','F4106','F4107') "
'            str1P0Sales = " and a1p16 in ('F4103','F4106','F4107') "
'      End Select
'      If Mid(Text9, 1, 2) = "F1" And Mid(Text10, 1, 2) = "F4" Then
'            'modify by sonia 2021/1/14 +F4104~F4107
'            strAccSales = " and ax209 in ('F4101','F4102','F4103','F4104','F4105','F4106','F4107') "
'            str1P0Sales = " and a1p16 in ('F4101','F4102','F4103','F4104','F4105','F4106','F4107') "
'      End If
'   Else
'      strAccSales = " and ax209='" & Text1 & "' "
'   End If
'   '2007/11/12 end
'end 2021/03/19

   strCon10 = ""
   adoacc0y0.CursorLocation = adUseClient
   '2007/11/12 modify by sonia 同一請款單計入最後收文之智權人員,另再加印財務調整傳票D096091227,再加舊系統請款單找不到收文記錄者D096090203
   'adoacc0y0.Open "select * from acc0z0, acc0y0, acc1k0, (select distinct cp01, cp13, cp60, CP12 from acc0z0, acc0y0, caseprogress, acc1k0 where a0z01 = a0y01 and a0z02 = cp60 (+) and a0z02 = a1k01 (+) and a0z04 <> 0" & strSalesMan & strSQL & ") new where a0z01 = a0y01 and a0z02 = a1k01 (+) and a0z02 = cp60 (+) and a0z04 <> 0" & strSalesMan & strSQL & " order by cp13", adoTaie, adOpenStatic, adLockReadOnly
   '2009/7/29 MODIFY BY SONIA FCL-10530於98/7/13收款有收文但因拆點數於FCL未出現(因為CP09 is null)故第三段修改
   '2010/9/13 MODIFY BY SONIA 發現第三段因為抓找不到收文記錄者應要加CP60 IS NULL的控制
   'Modify by Morgan 2011/4/11 ax207 改抓加總;另調整語法 cp09 in substr(new.cp,9,9)-->cp09(+)=substr(new.cp,9,9),a0202=ax202-->a0202=ax202(+)
   'modify by sonia 2013/5/30 加抵帳點數Z10200013(102/4/9)第四段,但抵帳之a1p21與國外收款之a1p21存的值不同故抵帳不做new.a0z04=a1p21(+),因抵帳A1P02='K'故第二段的'F'=a1p02(+)改掉,而第三段的'F'=a1p02(+)改為'F'=a1p02
   'adoacc0y0.Open "select cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04,sum(ax207) ax207 from caseprogress, acc1p0, acc021, " & _
                  "(select max(cp05||cp09) cp,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04,a0z02,a0z01 from acc0z0, acc0y0, acc1k0, caseprogress where a0z01(+) = a0y01 " & strSql & " and a0z04 <> 0 " & _
                  "and a0z02=a1k01(+) and a0z02 = cp60 (+) " & strSalesArea & strSystem & " group by a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04,a0z02,a0z01) new " & _
                  "where cp09(+)=substr(new.cp,9,9) " & strSalesMan & " and new.a0z01=a1p04(+) and new.a0z04=a1p21(+) and substr(a1P05,1,1) in ('4','7') " & str1P0Sales & "and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+) " & _
                  "and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+) " & _
                  " group by cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04 " & _
                  "union " & _
                  "select ax209 cp13,'' a1k01,0 a1k09,substr(ax214, 1, Length(ax214) - 9) a1k13,substr(ax214, Length(ax214) - 8, 6) a1k14,substr(ax214, Length(ax214) - 2, 1) a1k15,substr(ax214, Length(ax214) - 1, 2) a1k16,0 a1k30,a0205 a0y02,1 a0y04,ax207-ax206 a0z04,ax207-ax206 ax207 " & _
                  "from acc020,acc021,acc1p0 where a0202=ax202(+) " & strSQL1 & " and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
                  "and a0202=a1p22(+) and 'F'=a1p02(+) and a1p04 is null " & _
                  "union " & _
                  "select ax209 cp13,'' a1k01,0 a1k09,substr(ax214, 1, Length(ax214) - 9) a1k13,substr(ax214, Length(ax214) - 8, 6) a1k14,substr(ax214, Length(ax214) - 2, 1) a1k15,substr(ax214, Length(ax214) - 1, 2) a1k16,0 a1k30,a0205 a0y02,1 a0y04,ax207-ax206 a0z04,ax207-ax206 ax207 " & _
                  "from acc020,acc021,acc1p0,acc0z0,acc1k0,caseprogress where a0202=ax202(+) " & strSQL1 & " and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & straccSales & strAccSystem & _
                  "and ax202=a1p22(+) and 'F'=a1p02(+) and ax214=a1p17(+) and ax209=a1p16(+) and a1p04=a0z01(+) and a1p21=a0z04(+) and a0z02=a1k01(+) and a1k01=cp60(+) and cp09 is null " & _
                  "order by cp13,a1k01,a1k13,a1k14,a1k15,a1k16", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Lydia 2017/10/18 拿掉 adoacc0y0.Open, 放在strexc(1)
   'Modified by Lydia 2020/09/01
   'modify by sonia 2021/1/20 加傳票公司別條件a0201=ax201(+)及ax201=a1p01(+)
   'Modify by Amy 2021/03/19 程式改寫至 Pub_GetAccRecePayAmt
'   strExc(1) = "select cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04,sum(ax207) ax207 from caseprogress, acc1p0, acc021, " & _
'                  "(select max(cp05||cp09) cp,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04,a0z02,a0z01 from acc0z0, acc0y0, acc1k0, caseprogress where a0z01(+) = a0y01 " & strSql & " and a0z04 <> 0 " & _
'                  "and a0z02=a1k01(+) and a0z02 = cp60 (+) " & strSalesArea & strSystem & " group by a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04,a0z02,a0z01) new " & _
'                  "where cp09(+)=substr(new.cp,9,9) " & strSalesMan & " and new.a0z01=a1p04(+) and new.a0z04=a1p21(+) and substr(a1P05,1,1) in ('4','7') " & str1P0Sales & "and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+) " & _
'                  "and a1p01=ax201(+) and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+) group by cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04 " & _
'                  "union " & _
'                  "select ax209 cp13,'' a1k01,0 a1k09,substr(ax214, 1, Length(ax214) - 9) a1k13,substr(ax214, Length(ax214) - 8, 6) a1k14,substr(ax214, Length(ax214) - 2, 1) a1k15,substr(ax214, Length(ax214) - 1, 2) a1k16,0 a1k30,a0205 a0y02,1 a0y04,ax207-ax206 a0z04,ax207-ax206 ax207 " & _
'                  "from acc020,acc021,(select * from acc1p0 where " & StrSQL3 & " and a1p02 in ('F','K')) where a0201=ax201(+) and a0202=ax202(+) " & strSQL1 & " and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & strAccSales & strAccSystem & _
'                  "and a0201=a1p01(+) and a0202=a1p22(+) and a1p04 is null " & _
'                  "union " & _
'                  "select ax209 cp13,'' a1k01,0 a1k09,substr(ax214, 1, Length(ax214) - 9) a1k13,substr(ax214, Length(ax214) - 8, 6) a1k14,substr(ax214, Length(ax214) - 2, 1) a1k15,substr(ax214, Length(ax214) - 1, 2) a1k16,0 a1k30,a0205 a0y02,1 a0y04,ax207-ax206 a0z04,ax207-ax206 ax207 " & _
'                  "from acc020,acc021,acc1p0,acc0z0,acc1k0,caseprogress where a0201=ax201(+) and a0202=ax202(+) " & strSQL1 & " and substr(ax205,1,1) in ('4','7') and instr(AX212,'保留')=0 " & strAccSales & strAccSystem & _
'                  "and ax201=a1p01(+) and ax202=a1p22(+) and 'F'=a1p02 and ax214=a1p17(+) and ax209=a1p16(+) and a1p04=a0z01(+) and a1p21=a0z04(+) and a0z02=a1k01(+) and a1k01=cp60(+) and cp09 is null " & _
'                  "union " & _
'                  "select cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a1h02,a1h04,a1k08,sum(ax207) ax207 from caseprogress, acc1p0, acc021, " & _
'                  "(select max(cp05||cp09) cp,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a1h02,a1h04,a1k08,a1h01 from acc1h0, acc1k0, caseprogress where " & strSQL2 & _
'                  "and a1h01=a1k17(+) and a1k01 = cp60 (+) " & strSalesArea & strSystem & " group by a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a1h02,a1h04,a1k08,a1k01,a1h01) new " & _
'                  "where cp09(+)=substr(new.cp,9,9) " & strSalesMan & " and new.a1h01=a1p04(+) and substr(a1P05,1,1) in ('4','7') " & str1P0Sales & "and new.a1k13||new.a1k14||new.a1k15||new.a1k16=a1p17(+) " & _
'                  "and a1p01=ax201(+) and a1p22=ax202(+) and a1p17=ax214(+) and substr(ax205,1,1) in ('4','7') and a1p03=ax203(+) " & _
'                  " group by cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a1h02,a1h04,a1k08 " & _
'                  "order by cp13,a1k01,a1k13,a1k14,a1k15,a1k16"

   strExc(1) = "Select cp13,a1k01,a1k09,a1k13,a1k14,a1k15,a1k16,a1k30,a0y02,a0y04,a0z04,ax207,ReceVal,ProVal,Amt,st16,st70 " & _
                     "From (" & Pub_GetAccRecePayAmt(Me.Name, strDateS, strDateE, Text9, Text10, Text1, m_strSystem, strSalesMan) & ") Order by cp13,a1k01,a1k13,a1k14,a1k15,a1k16"
   'end 2021/03/19
   
   'Added by Lydia 2017/10/18 增加申請人/代理人之國籍/洲別
   If strBaseTable <> "" Then
       strExc(1) = UCase(strExc(1))
       strSql = "select a.* from (" & Mid(strExc(1), 1, InStr(strExc(1), "ORDER BY") - 1) & ") a " & strBaseTable & _
                " where " & strCaseNa & " order by cp13,a1k01,a1k13,a1k14,a1k15,a1k16"
       'Added by Lydia 2020/09/01 檢查無申請人/代理人的案件
       strExc(0) = "select a.* from (" & Mid(strExc(1), 1, InStr(strExc(1), "ORDER BY") - 1) & ") a " & strBaseTable & _
                " where " & strCaseNaChk & " order by cp13,a1k01,a1k13,a1k14,a1k15,a1k16"
       strCaseNaChk = strExc(0)
       'end 2020/09/01
   Else
       strSql = strExc(1)
   End If
   'end 2017/10/18
   
   adoacc0y0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly 'Added by Lydia 2017/10/18
   If adoacc0y0.RecordCount = 0 Then
      strCon10 = MsgText(602)
      adoacc0y0.Close
      MsgBox MsgText(28), , MsgText(5)
      strCaseNaChk = "" 'Added by Lydia 2020/09/01
      Exit Sub
   'Added by Lydia 2020/09/01 因為外商常有案件無申請人/代理人,所以彈提醒
   'Modified by Lydia 2020/10/05 debug
   'Else
   ElseIf strCaseNaChk <> "" Then
       intI = 1
       'Modified by Lydia 2020/10/05 debug
       'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       Set RsTemp = ClsLawReadRstMsg(intI, strCaseNaChk)
       If intI = 0 Then
            strCaseNaChk = ""
       Else
            strCaseNaChk = "有 " & RsTemp.RecordCount & " 筆點數，案件無" & IIf(txtKind = "1", "申請人", "代理人") & "！"
       End If
   'end 2020/09/01
   End If
   
   Do While adoacc0y0.EOF = False
      adoaccrpt212.AddNew
      'Add by Morgan 2003/12/02
      If IsNull(adoacc0y0.Fields("A1k01").Value) Then
         adoaccrpt212.Fields("r21213").Value = Null
      Else
         adoaccrpt212.Fields("r21213").Value = adoacc0y0.Fields("A1k01").Value
      End If
      'End 2003/12/02
      adoaccrpt212.Fields("r21201").Value = strUserNum
      If IsNull(adoacc0y0.Fields("cp13").Value) Then
         adoaccrpt212.Fields("r21202").Value = Null
         adoaccrpt212.Fields("r21211").Value = Null
      Else
         adoaccrpt212.Fields("r21202").Value = adoacc0y0.Fields("cp13").Value
         'Modify by Amy 2021/04/22 原取全部名稱,但F41XX字數太多會重疊
         adoaccrpt212.Fields("r21211").Value = Left(StaffQuery(adoacc0y0.Fields("cp13").Value), 4)
      End If
      If IsNull(adoacc0y0.Fields("a1k13").Value) Then
         adoaccrpt212.Fields("r21203").Value = Null
      Else
         adoaccrpt212.Fields("r21203").Value = adoacc0y0.Fields("a1k13").Value
         If IsNull(adoacc0y0.Fields("a1k14").Value) = False Then
            adoaccrpt212.Fields("r21203").Value = adoaccrpt212.Fields("r21203").Value & "-" & adoacc0y0.Fields("a1k14").Value
         End If
         If IsNull(adoacc0y0.Fields("a1k15").Value) = False Then
            adoaccrpt212.Fields("r21203").Value = adoaccrpt212.Fields("r21203").Value & "-" & adoacc0y0.Fields("a1k15").Value
         End If
         If IsNull(adoacc0y0.Fields("a1k16").Value) = False Then
            adoaccrpt212.Fields("r21203").Value = adoaccrpt212.Fields("r21203").Value & "-" & adoacc0y0.Fields("a1k16").Value
         End If
      End If
      If IsNull(adoacc0y0.Fields("a0y02").Value) Then
         adoaccrpt212.Fields("r21204").Value = Null
      Else
         adoaccrpt212.Fields("r21204").Value = adoacc0y0.Fields("a0y02").Value
      End If
      If IsNull(adoacc0y0.Fields("a0y04").Value) Then
         douExchange = 0
      Else
         douExchange = adoacc0y0.Fields("a0y04").Value
      End If
      If IsNull(adoacc0y0.Fields("a0z04").Value) Then
         adoaccrpt212.Fields("r21208").Value = 0
      Else
         'Modify by Amy 2021/04/29 改抓Pub_GetAccRecePayAmt資料
         'adoaccrpt212.Fields("r21208").Value = Val(Format(Val(adoacc0y0.Fields("a0z04").Value) * douExchange, FAmount))
         adoaccrpt212.Fields("r21208").Value = Val("" & adoacc0y0.Fields("ReceVal"))
      End If
      '2007/11/8 modify by sonia 已收點數應扣除規費,第一次收款即應扣規費
      'adoaccrpt212.Fields("r21209").Value = Val(Format(Val(adoaccrpt212.Fields("r21208").Value) / 1000, FAmount))
      'adoaccrpt212.Fields("r21212").Value = Val(adoaccrpt212.Fields("r21208").Value) / 1000
      'Modify by Amy 2021/03/02 +"" ,78011 下 1100201~1100228 /資料內容2.收款點數/報表內容4.FCT組別 ,因X10916319 之號號a1k30為null會Error
      'Modify by Amy 2021/04/29 改抓Pub_GetAccRecePayAmt資料
'      If adoaccrpt212.Fields("r21208").Value = Val("" & adoacc0y0.Fields("A1k30").Value) Then
'         adoaccrpt212.Fields("r21209").Value = Val(Format(Val(Val(adoaccrpt212.Fields("r21208").Value) - Val(adoacc0y0.Fields("a1k09").Value)) / 1000, FDollar)) '0.00000
'      Else
'         adoaccrpt212.Fields("r21209").Value = Val(Format(Val(adoaccrpt212.Fields("r21208").Value) / 1000, FDollar)) '0.00000
'      End If
      adoaccrpt212.Fields("r21209").Value = Val("" & adoacc0y0.Fields("ProVal"))
      
      adoaccsum.CursorLocation = adUseClient
'      adoaccsum.Open "select * from acc1p0 where a1p18 in (select min(a1p18) from acc1p0 where a1p01 = '1' and a1p17 = '" & adoacc0y0.Fields("a1k13").Value & adoacc0y0.Fields("a1k14").Value & adoacc0y0.Fields("a1k15").Value & adoacc0y0.Fields("a1k16").Value & "' and a1p05 = '6130') and a1p01 = '1' and a1p17 = '" & adoacc0y0.Fields("a1k13").Value & adoacc0y0.Fields("a1k14").Value & adoacc0y0.Fields("a1k15").Value & adoacc0y0.Fields("a1k16").Value & "' and a1p05 = '6130'", adoTaie, adOpenStatic, adLockReadOnly
      '2007/11/13 modify by sonia 該請款單有201新案翻譯才抓
      'adoaccsum.Open "select ax206 from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and a0205 in (select min(a0205) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax214 = '" & adoacc0y0.Fields("a1k13").Value & adoacc0y0.Fields("a1k14").Value & adoacc0y0.Fields("a1k15").Value & adoacc0y0.Fields("a1k16").Value & "' and ax205 = '6130') and ax214 = '" & adoacc0y0.Fields("a1k13").Value & adoacc0y0.Fields("a1k14").Value & adoacc0y0.Fields("a1k15").Value & adoacc0y0.Fields("a1k16").Value & "' and ax205 = '6130'", adoTaie, adOpenStatic, adLockReadOnly
      adoaccrpt212.Fields("r21210").Value = 0
      If Not IsNull(adoacc0y0.Fields("A1k01").Value) Then
         adoaccsum.Open "select ax206 from acc021, acc020, caseprogress where ax201 = a0201 and ax202 = a0202 and a0205 in (select min(a0205) from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and ax214 = '" & adoacc0y0.Fields("a1k13").Value & adoacc0y0.Fields("a1k14").Value & adoacc0y0.Fields("a1k15").Value & adoacc0y0.Fields("a1k16").Value & "' and ax205 = '6130') and ax214 = '" & adoacc0y0.Fields("a1k13").Value & adoacc0y0.Fields("a1k14").Value & adoacc0y0.Fields("a1k15").Value & adoacc0y0.Fields("a1k16").Value & "' and ax205 = '6130' " & _
                        "and '" & adoacc0y0.Fields("A1k01").Value & "'=cp60(+) and (cp10='201' or cp10='927') ", adoTaie, adOpenStatic, adLockReadOnly
         If adoaccsum.RecordCount <> 0 Then
            If Not IsNull(adoaccsum.Fields(0).Value) Then
               adoaccrpt212.Fields("r21210").Value = adoaccsum.Fields(0).Value
            End If
         End If
         adoaccsum.Close
      End If
      '2007/11/13 add by sonia 再印財務點數
      'Modify by Amy 2021/04/22 原:FAmount 改與已收點數一樣3位,較好比較
      adoaccrpt212.Fields("r21212").Value = Val(Format(Val(adoacc0y0.Fields("ax207").Value) / 1000, FDollar)) '0.000
      '2007/11/13 end
      adoaccrpt212.UpdateBatch
      adoacc0y0.MoveNext
   Loop
   adoacc0y0.Close
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Text11 = "1"
   Text12 = "2" 'Add By Sindy 2010/9/16
   Text1.SetFocus
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
'   If Text1 = MsgText(601) Then
'      FormCheck = False
'      Text1.SetFocus
'      Exit Function
'   End If
'   If Text2 = MsgText(601) Then
'      FormCheck = False
'      Exit Function
'   End If
'   If Text3 = MsgText(601) Then
'      FormCheck = False
'      Exit Function
'   End If
'   If Text4 = MsgText(601) Then
'      FormCheck = False
'      Exit Function
'   End If
'   If Text5 = MsgText(601) Then
'      FormCheck = False
'      Exit Function
'   End If
'   If Text6 = MsgText(601) Then
'      FormCheck = False
'      Exit Function
'   End If
'   If Text7 = MsgText(601) Then
'      FormCheck = False
'      Exit Function
'   End If
   If MaskEdBox1.Text = MsgText(29) Then
      FormCheck = False
      MaskEdBox1.SetFocus
      MsgBox "帳款起始日期不可空白！", , MsgText(5)
      Exit Function
   End If
   If MaskEdBox2.Text = MsgText(29) Then
      FormCheck = False
      MaskEdBox2.SetFocus
      MsgBox "帳款迄止日期不可空白！", , MsgText(5)
      Exit Function
   End If
   If Text8 = MsgText(601) Then
      FormCheck = False
      Text8.SetFocus
      MsgBox "資料內容不可空白！", , MsgText(5)
      Exit Function
   End If
   If Text11 = MsgText(601) Then
      FormCheck = False
      Text11.SetFocus
      MsgBox "報表內容不可空白！", , MsgText(5)
      Exit Function
   End If
   If Text12 = MsgText(601) Then
      FormCheck = False
      Text12.SetFocus
      MsgBox "輸出方式不可空白！", , MsgText(5)
      Exit Function
   End If
   FormCheck = True
End Function

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Morgan 2010/5/18
Private Sub PrintSubTot(strSalesNo As String)
   Dim stCon As String
   
   If strSalesNo <> "" Then
      stCon = " and R21202='" & strSalesNo & "'"
      strTemp(4) = "小計："
   Else
      strTemp(4) = "合計："
   End If
   strExc(0) = "select max(R21211) s0,sum(R21205) s1,sum(R21206) s2,sum(R21207) s3,sum(R21208) s4,sum(R21209) s5,sum(R21210) s6 from accrpt212 where r21201='" & strUserNum & "' " & stCon
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      If Text11 <> "1" And strSalesNo <> "" Then
         strTemp(1) = "" & .Fields("s0")
      Else
         strTemp(1) = ""
      End If
      strTemp(2) = ""
      strTemp(3) = ""
      
      strTemp(5) = Format(Val("" & .Fields("s1")), DDollar)
      strTemp(6) = Format(Val("" & .Fields("s2")), DDollar)
      strTemp(7) = Format(Val("" & .Fields("s3")), FDollar)
      'Add By Sindy 2010/10/5
      'If strSalesNo = "" Then
         m_dbltotPoint = Format(Val("" & .Fields("s3")), FDollar)
      'End If
      '2010/10/5 End
      strTemp(8) = Format(Val("" & .Fields("s4")), DDollar)
      strTemp(9) = Format(Val("" & .Fields("s5")), FDollar)
      strTemp(10) = Format(Val("" & .Fields("s6")), DDollar)
      End With
      PrintNewLine
      DrawLine 4
      m_Device.FontBold = True
      PrintDetail
      If strSalesNo = "" Then
         PrintShareTot
      End If
      m_Device.FontBold = False
   End If
End Sub

'Add By Sindy 2010/9/16
Private Sub PrintSubTot_2(strSalesNo As String)
   Dim stCon As String
   
   If strSalesNo <> "" Then
      stCon = " and R21202='" & strSalesNo & "'"
      strTemp(4) = "小計："
   Else
      strTemp(4) = "合計："
   End If
   strExc(0) = "select max(R21211) s0,sum(R21208) s1,sum(R21209) s2,sum(R21210) s3,sum(R21212) s4 from accrpt212 where r21201='" & strUserNum & "' " & stCon
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      If Text11 <> "1" And strSalesNo <> "" Then
         strTemp(1) = "" & .Fields("s0")
      Else
         strTemp(1) = ""
      End If
      strTemp(2) = ""
      strTemp(3) = ""
      'Modified by Morgan 2012/3/14 +格式
      'strTemp(5) = "" & .Fields("s1")
      'strTemp(6) = "" & .Fields("s2")
      'strTemp(7) = "" & .Fields("s3")
      'strTemp(8) = "" & .Fields("s4")
      strTemp(5) = Format(Val("" & .Fields("s1")), DDollar)
      strTemp(6) = Format(Val("" & .Fields("s2")), FDollar)
      strTemp(7) = Format(Val("" & .Fields("s3")), DDollar)
      strTemp(8) = Format(Val("" & .Fields("s4")), FDollar)
      
      End With
      PrintNewLine
      DrawLine 4
      m_Device.FontBold = True
      PrintDetail
      m_Device.FontBold = False
   End If
End Sub

'請款點數報表
Private Sub rptAcc24c0()
   Dim iRecs As Integer, strLstSales As String, intR As Integer
   Dim adoRst As ADODB.Recordset, adoRst2 As ADODB.Recordset
   Dim strLstGP As String, strLstSharePointGP As String, strLstSharePointID As String 'Added by Morgan 2021/7/16
   
   '分配點數資料
   'Modified by Morgan 2021/7/16 配合FCT組別選項 + ST16+ST70排序
   'strExc(0) = "select distinct R01 from accrpt212_1 where id='" & strUserNum & "' order by 1"
   strExc(0) = "select distinct R01,nvl(st16,'9')||nvl(st70,'9') GP from accrpt212_1,staff where id='" & strUserNum & "' and st01(+)=R01 "
   If Text11 = "4" Then
      strExc(0) = strExc(0) & " order by 2 asc,1 asc"
   Else
      strExc(0) = strExc(0) & " order by 1 asc"
   End If
   'end 2021/7/16
   
   intR = 1
   Set adoRst2 = ClsLawReadRstMsg(intR, strExc(0))
   
   strSql = "update accrpt212 set R21207=(select sum(a1n05) from acc1n0 where a1n01=R21213 and a1n02='1' and a1n04=R21202)" & _
      ",R21214=(select max('*') from acc1n0 where a1n01=R21213 and a1n02='1' and a1n04<>R21202)" & _
      " where r21201='" & strUserNum & "' and exists(select * from acc1n0 where a1n01=R21213 and a1n02='1')"
   cnnConnection.Execute strSql, intI
   
   'Modify By Sindy 2018/11/29 + ST16+ST70排序
   If Text11 = "4" Then 'FCT組別
      'Modified by Morgan 2021/7/16 修正FCT組別未列出只有分配點數的人員問題 Ex:110/6
      'strExc(0) = "select * from accrpt212 a,staff b where r21201='" & strUserNum & "' and r21202=st01(+)" & _
                  " order by r21201 asc, nvl(st16,'9')||nvl(st70,'9') asc ,r21202 asc, r21204 asc, r21203 asc"
      strExc(0) = "select a.*,b.*,nvl(st16,'9')||nvl(st70,'9') GP from accrpt212 a,staff b where r21201='" & strUserNum & "' and r21202=st01(+)" & _
                  " order by r21201 asc, nvl(st16,'9')||nvl(st70,'9') asc ,r21202 asc, r21204 asc, r21203 asc"
   Else
   '2018/11/29 END
      strExc(0) = "select * from accrpt212 where r21201='" & strUserNum & "' order by r21201 asc, r21202 asc, r21204 asc, r21203 asc"
   End If
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetPleft
      PrintPageHeader
      With adoRst
         .MoveFirst
         iRecs = 0
         strLstSales = ""
         Do While Not .EOF
            If strLstSales <> "" & .Fields("R21202") Then
               If strLstSales <> "" Then
                  PrintSubTot strLstSales
                  'Add by Morgan 2010/9/7
                  If intR = 1 Then
                     'Added by Morgan 2012/9/4 只有分配點數的先印
                     If Text11 = "1" Then
                        If Not adoRst2.EOF Then
                           If adoRst2.Fields("R01") = strLstSales Then
                              PrintSharePoint adoRst2.Fields("R01")
                              adoRst2.MoveNext
                           End If
                        End If
                     'end 2012/9/4

                     ElseIf Text11 = "2" Then
                        'Modify By Sindy 2010/10/5
                        'PrintShareTot strLstSales

                        'Modified by Morgan 2012/9/4 只有分配點數的先印
                        'Do While Not adoRst2.EOF
                        '   If adoRst2.Fields("R01") <= strLstSales Then
                        '      If adoRst2.Fields("R01") <> strLstSales Then
                        '         PrintShareTot adoRst2.Fields("R01"), GetPrjSalesNM(adoRst2.Fields("R01"))
                        '      Else
                        '         PrintShareTot adoRst2.Fields("R01")
                        '      End If
                        '   Else
                        '      Exit Do
                        '   End If
                        '   adoRst2.MoveNext
                        'Loop
                        If Not adoRst2.EOF Then
                           If adoRst2.Fields("R01") = strLstSales Then
                              PrintShareTot adoRst2.Fields("R01")
                              adoRst2.MoveNext
                           End If
                        End If
                        'end 2012/9/4
                        '2010/10/5 End
                     
                     'Add By Sindy 2018/11/29
                     ElseIf Text11 = "4" Then
                        adoRst2.MoveFirst
                        Do While Not adoRst2.EOF
                           If adoRst2.Fields("R01") = strLstSales Then
                              strLstSharePointID = adoRst2.Fields("R01") 'Added by Morgan 2021/7/16
                              PrintShareTot adoRst2.Fields("R01")
                              adoRst2.MoveNext
                              Exit Do
                           End If
                           adoRst2.MoveNext
                        Loop
                     End If
                  End If
                  'end 2010/9/7
               End If
               
               strLstSales = "" & .Fields("R21202")
               If intR = 1 Then
                  If Text11 = "1" Then
                     Do While Not adoRst2.EOF
                        'Modified by Morgan 2012/9/4 只有分配點數的先印
                        'If adoRst2.Fields("R01") < .Fields("R21202") Then
                        If adoRst2.Fields("R01") < strLstSales Then
                           PrintSharePoint adoRst2.Fields("R01")
                        Else
                           Exit Do
                        End If
                        adoRst2.MoveNext
                     Loop
                  'Added by Morgan 2012/9/4 只有分配點數的先印
                  ElseIf Text11 = "2" Then
                     Do While Not adoRst2.EOF
                        If adoRst2.Fields("R01") < strLstSales Then
                           PrintShareTot adoRst2.Fields("R01"), GetPrjSalesNM(adoRst2.Fields("R01"))
                        Else
                           Exit Do
                        End If
                        adoRst2.MoveNext
                     Loop
                  'end 2012/9/4
                  
                  'Added by Morgan 2021/7/16 修正FCT組別未列出只有分配點數的人員問題 Ex:110/6
                  ElseIf Text11 = "4" Then
                     adoRst2.MoveFirst
                     If strLstSharePointID <> "" Then
                        adoRst2.Find "R01='" & strLstSharePointID & "'"
                        adoRst2.MoveNext
                     End If
                     Do While Not adoRst2.EOF
                        If adoRst2.Fields("GP") & adoRst2.Fields("R01") < .Fields("GP") & .Fields("R21202") Then
                           strLstSharePointID = adoRst2.Fields("R01")
                           PrintShareTot adoRst2.Fields("R01"), GetPrjSalesNM(adoRst2.Fields("R01"))
                        Else
                           Exit Do
                        End If
                        adoRst2.MoveNext
                     Loop
                     
                  'end 2021/7/16
                  End If
               End If
            End If
            If Text11 = "1" Then
               iRecs = iRecs + 1
               strTemp(1) = "" & .Fields("R21211")
               strTemp(2) = "" & .Fields("R21203")
               strTemp(3) = "" & .Fields("R21213")
               strTemp(4) = ChangeTStringToTDateString("" & .Fields("R21204"))
               strTemp(5) = Format(Val("" & .Fields("R21205")), DDollar)
               strTemp(6) = Format(Val("" & .Fields("R21206")), DDollar)
               strTemp(7) = "" & .Fields("R21214") & Format(Val("" & .Fields("R21207")), FDollar)
               strTemp(8) = Format(Val("" & .Fields("R21208")), DDollar)
               strTemp(9) = Format(Val("" & .Fields("R21209")), FDollar)
               strTemp(10) = Format(Val("" & .Fields("R21210")), DDollar)
               PrintDetail
            End If
            
            .MoveNext
         Loop
         PrintSubTot strLstSales
         'end 2010/9/7
         If intR = 1 Then
            If Text11 = "1" Then
               Do While Not adoRst2.EOF
                  PrintSharePoint adoRst2.Fields("R01")
                  adoRst2.MoveNext
               Loop
            
            'Add by Morgan 2010/9/7
            ElseIf Text11 = "2" Then
               'Modify By Sindy 2010/10/5
               'PrintShareTot strLstSales
               Do While Not adoRst2.EOF
                  If adoRst2.Fields("R01") >= strLstSales Then
                     If adoRst2.Fields("R01") <> strLstSales Then
                        PrintShareTot adoRst2.Fields("R01"), GetPrjSalesNM(adoRst2.Fields("R01"))
                     Else
                        PrintShareTot adoRst2.Fields("R01")
                     End If
                  End If
                  adoRst2.MoveNext
               Loop
               '2010/10/5 End
               
            'Add By Sindy 2018/11/29
            ElseIf Text11 = "4" Then
               adoRst2.MoveFirst
               Do While Not adoRst2.EOF
                  If adoRst2.Fields("R01") = strLstSales Then
                     PrintShareTot adoRst2.Fields("R01")
                     adoRst2.MoveNext
                     Exit Do
                  End If
                  adoRst2.MoveNext
               Loop
            End If
         End If
         PrintSubTot ""
         Call PrintReportFooter(iRecs)
         'm_Device.EndDoc
      End With
      
   '只有分配點數資料
   ElseIf intR = 1 And Text11 = "1" Then
      GetPleft
      PrintPageHeader
      Do While Not adoRst2.EOF
         PrintSharePoint adoRst2.Fields("R01")
         adoRst2.MoveNext
      Loop
      'm_Device.EndDoc
   End If
   Set adoRst = Nothing
   Set adoRst2 = Nothing
End Sub

'Added by Morgan 2012/3/14 組別統計
Private Sub rptAcc24c0_3()
   Dim iRecs As Integer
   
   If Text8 = "1" Then
      strExc(0) = "select R21215 s0,sum(R21205) s1,sum(R21206) s2,sum(R21207) s3,sum(R21208) s4,sum(R21209) s5,sum(R21210) s6 from accrpt212 where r21201='" & strUserNum & "' group by R21215"
   Else
      strExc(0) = "select R21215 s0,sum(R21208) s1,sum(R21209) s2,sum(R21210) s3,sum(R21212) s4 from accrpt212 where r21201='" & strUserNum & "' group by R21215"
   End If
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      
      If Text8 = "1" Then
         GetPleft
      Else
         GetPleft_2
      End If
      
      PrintPageHeader
      Do While Not .EOF
         If IsNull(.Fields("s0")) Then
            strTemp(1) = "其他"
         Else
            strTemp(1) = PUB_GetFCPGrpName("" & .Fields("s0"))
         End If
         strTemp(2) = ""
         strTemp(3) = ""
         strTemp(4) = "小計："
         If Text8 = "1" Then
            strTemp(5) = Format(Val("" & .Fields("s1")), DDollar)
            strTemp(6) = Format(Val("" & .Fields("s2")), DDollar)
            strTemp(7) = Format(Val("" & .Fields("s3")), FDollar)
            strTemp(8) = Format(Val("" & .Fields("s4")), DDollar)
            strTemp(9) = Format(Val("" & .Fields("s5")), FDollar)
            strTemp(10) = Format(Val("" & .Fields("s6")), DDollar)
         Else
            strTemp(5) = Format(Val("" & .Fields("s1")), DDollar)
            strTemp(6) = Format(Val("" & .Fields("s2")), FDollar)
            strTemp(7) = Format(Val("" & .Fields("s3")), DDollar)
            strTemp(8) = Format(Val("" & .Fields("s4")), FDollar)
         End If
         
         If .AbsolutePosition > 1 Then
            PrintNewLine
            DrawLine 4
         End If
         m_Device.FontBold = True
         PrintDetail
         m_Device.FontBold = False
         .MoveNext
      Loop
      End With
   End If
   
   If Text8 = "1" Then
      strExc(0) = "select sum(R21205) s1,sum(R21206) s2,sum(R21207) s3,sum(R21208) s4,sum(R21209) s5,sum(R21210) s6 from accrpt212 where r21201='" & strUserNum & "'"
   Else
      strExc(0) = "select sum(R21208) s1,sum(R21209) s2,sum(R21210) s3,sum(R21212) s4 from accrpt212 where r21201='" & strUserNum & "'"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      strTemp(1) = ""
      strTemp(2) = ""
      strTemp(3) = ""
      strTemp(4) = "合計："
      If Text8 = "1" Then
         strTemp(5) = Format(Val("" & .Fields("s1")), DDollar)
         strTemp(6) = Format(Val("" & .Fields("s2")), DDollar)
         strTemp(7) = Format(Val("" & .Fields("s3")), FDollar)
         strTemp(8) = Format(Val("" & .Fields("s4")), DDollar)
         strTemp(9) = Format(Val("" & .Fields("s5")), FDollar)
         strTemp(10) = Format(Val("" & .Fields("s6")), DDollar)
      Else
         strTemp(5) = Format(Val("" & .Fields("s1")), DDollar)
         strTemp(6) = Format(Val("" & .Fields("s2")), FDollar)
         strTemp(7) = Format(Val("" & .Fields("s3")), DDollar)
         strTemp(8) = Format(Val("" & .Fields("s4")), FDollar)
      End If
      End With
      PrintNewLine
      DrawLine 4
      m_Device.FontBold = True
      PrintDetail
      m_Device.FontBold = False
   End If
   
   Call PrintReportFooter(iRecs)
End Sub

Private Sub DelPic()
   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_*.tmp"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
End Sub

Private Sub SetPic(idx As Integer)

   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_" & idx & ".tmp"
   
'   Clipboard.Clear
'   Clipboard.SetData Picture1.Image
'   Set m_Pictures(m_iPages - 1) = Clipboard.GetData
'   Set m_Pictures(idx) = Picture1.Image

   SavePicture Picture1.Image, strPicFileName
   '要用覆蓋的否則會錯誤--VB Bug
   'Picture1.Cls
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
   
End Sub

'Add By Sindy 2010/9/16 收款點數報表
Private Sub rptAcc24c0_2()
   Dim iRecs As Integer, strLstSales As String
   Dim adoRst As ADODB.Recordset
   
   'Modify By Sindy 2018/11/29 + ST16+ST70排序
   If Text11 = "4" Then 'FCT組別
      strExc(0) = "select * from accrpt212,staff where r21201='" & strUserNum & "' and r21202=st01(+)" & _
                  " order by r21201 asc, nvl(st16,'9')||nvl(st70,'9') asc ,r21202 asc, r21204 asc, r21203 asc,r21213 asc"
   Else
   '2018/11/29 END
      strExc(0) = "select * from accrpt212 where r21201='" & strUserNum & "' order by r21201 asc,r21202 asc,r21204 asc,r21203 asc,r21213 asc"
   End If
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetPleft_2
      PrintPageHeader
      With adoRst
         .MoveFirst
         iRecs = 0
         strLstSales = ""
         Do While Not .EOF
            If strLstSales <> "" & .Fields("R21202") Then
               If strLstSales <> "" Then
                  PrintSubTot_2 strLstSales
               End If
            End If
            If Text11 = "1" Then
               iRecs = iRecs + 1
               strTemp(1) = "" & .Fields("R21211")
               strTemp(2) = "" & .Fields("R21203")
               strTemp(3) = "" & .Fields("R21213")
               strTemp(4) = ChangeTStringToTDateString("" & .Fields("R21204"))
               strTemp(5) = "" & .Fields("R21208")
               strTemp(6) = "" & .Fields("R21209")
               strTemp(7) = "" & .Fields("R21210")
               strTemp(8) = "" & .Fields("R21212")
               PrintDetail
            End If
            strLstSales = "" & .Fields("R21202")
            .MoveNext
         Loop
         PrintSubTot_2 strLstSales
         PrintSubTot_2 ""
         Call PrintReportFooter(iRecs)
      End With
      'm_Device.EndDoc
   End If
   Set adoRst = Nothing
End Sub

Sub GetPleft()
   Dim ii As Integer, iCols As Integer
   'Modify By Sindy 2010/9/16
   If m_bPrinter = True Then
      m_Device.Orientation = 1
   End If
   '2010/9/16 End
   m_iStartX = 500
   m_iStartY = 500
   m_iPageHeight = m_Device.ScaleHeight
   m_iLineHeight = 300
   m_iMargin = 500
   iPage = 0: m_iPages = 0
   
   iCols = 10
   
   Erase PLeft
   Erase PColName
   Erase PColName2
   
   ReDim PLeft(1 To iCols + 1)
   ReDim PColName(1 To iCols)
   ReDim PColName2(1 To iCols)
   ReDim strTemp(1 To iCols)
   
   '智權人員 本所案號 請款編號 請款日期 請款金額 規費 請款點數 已收金額 已收點數 支付翻譯費
   '
   ii = 1
   PLeft(ii) = 500
   Select Case Text11
      Case "1"
      
         PColName(ii) = "智權人員"
         'Modified by Lydia 2017/10/18 + 700 => 900
         PLeft(ii + 1) = PLeft(ii) + 900
         
         ii = ii + 1
         PColName(ii) = "本所案號"
         PLeft(ii + 1) = PLeft(ii) + 1500
         
         ii = ii + 1
         PColName(ii) = "請款編號"
         PLeft(ii + 1) = PLeft(ii) + 1000
         
         ii = ii + 1
         PColName(ii) = "請款日期"
         PLeft(ii + 1) = PLeft(ii) + 950
         
      'Added by Morgan 2012/3/14
      'Modify By Sindy 2018/11/26 + 4
      Case "2", "4"
         PColName(ii) = "智權人員"
         'Modified by Lydia 2017/10/18 + 700 => 900
         PLeft(ii + 1) = PLeft(ii) + 900
         
         ii = ii + 1
         PColName(ii) = ""
         PLeft(ii + 1) = PLeft(ii) + 1500
         
         ii = ii + 1
         PColName(ii) = ""
         PLeft(ii + 1) = PLeft(ii) + 1000
         
         ii = ii + 1
         PColName(ii) = ""
         PLeft(ii + 1) = PLeft(ii) + 950
         
      Case "3"
         PColName(ii) = "組別"
         PLeft(ii + 1) = PLeft(ii) + 1500
         
         ii = ii + 1
         PColName(ii) = ""
         PLeft(ii + 1) = PLeft(ii) + 700
         
         ii = ii + 1
         PColName(ii) = ""
         PLeft(ii + 1) = PLeft(ii) + 1000
         
         ii = ii + 1
         PColName(ii) = ""
         PLeft(ii + 1) = PLeft(ii) + 950
      'End 2012/3/14
   End Select
   
   ii = ii + 1
   PColName(ii) = "請款金額"
   PColName2(ii) = "(扣除折讓金額)"
   PLeft(ii + 1) = PLeft(ii) + 1300
   ii = ii + 1
   PColName(ii) = "規費"
   PLeft(ii + 1) = PLeft(ii) + 950
   
   ii = ii + 1
   PColName(ii) = "請款點數"
   PColName2(ii) = "*有分配點數給他人"
   PLeft(ii + 1) = PLeft(ii) + 1000
   
   ii = ii + 1
   PColName(ii) = "已收金額"
   PLeft(ii + 1) = PLeft(ii) + 1100
   
   ii = ii + 1
   PColName(ii) = "已收點數"
   PLeft(ii + 1) = PLeft(ii) + 950
   
   ii = ii + 1
   PColName(ii) = "支付翻譯費"
   PLeft(ii + 1) = PLeft(ii) + 1100
End Sub

Sub GetPleft_2()
   Dim ii As Integer, iCols As Integer
   'Modify By Sindy 2010/9/16
   If m_bPrinter = True Then
      m_Device.Orientation = 1
   End If
   '2010/9/16 End
   m_iStartX = 500
   m_iStartY = 500
   m_iPageHeight = m_Device.ScaleHeight
   m_iLineHeight = 300
   m_iMargin = 500
   iPage = 0: m_iPages = 0
   
   iCols = 8
   
   Erase PLeft
   Erase PColName
   Erase PColName2
   
   ReDim PLeft(1 To iCols + 1)
   ReDim PColName(1 To iCols)
   ReDim PColName2(1 To iCols)
   ReDim strTemp(1 To iCols)
   
   '智權人員 本所案號 請款編號 收款日期 已收金額 已收點數 支付翻譯費 財務點數
   '
   ii = 1
   PLeft(ii) = 500
   Select Case Text11
      Case "1"
         PColName(ii) = "智權人員"
         'Modified by Lydia 2017/10/18 + 700 => 900
         PLeft(ii + 1) = PLeft(ii) + 900
   
         ii = ii + 1
         PColName(ii) = "本所案號"
         PLeft(ii + 1) = PLeft(ii) + 1600
         
         ii = ii + 1
         PColName(ii) = "請款編號"
         PLeft(ii + 1) = PLeft(ii) + 1100
         
         ii = ii + 1
         PColName(ii) = "收款日期"
         PLeft(ii + 1) = PLeft(ii) + 1050
      
      'Modify By Sindy 2018/11/26 + 4
      Case "2", "4"
         PColName(ii) = "智權人員"
         'Modified by Lydia 2017/10/18 + 700 => 900
         PLeft(ii + 1) = PLeft(ii) + 900
   
         ii = ii + 1
         PColName(ii) = ""
         PLeft(ii + 1) = PLeft(ii) + 1600
         
         ii = ii + 1
         PColName(ii) = ""
         PLeft(ii + 1) = PLeft(ii) + 1100
         
         ii = ii + 1
         PColName(ii) = ""
         PLeft(ii + 1) = PLeft(ii) + 1050
      Case "3"
         PColName(ii) = "組別"
         PLeft(ii + 1) = PLeft(ii) + 1600
   
         ii = ii + 1
         PColName(ii) = ""
         PLeft(ii + 1) = PLeft(ii) + 800
         
         ii = ii + 1
         PColName(ii) = ""
         PLeft(ii + 1) = PLeft(ii) + 1100
         
         ii = ii + 1
         PColName(ii) = ""
         PLeft(ii + 1) = PLeft(ii) + 1050
         
   End Select
   ii = ii + 1
   PColName(ii) = "已收金額"
   PLeft(ii + 1) = PLeft(ii) + 1500
   
   ii = ii + 1
   PColName(ii) = "已收點數"
   PLeft(ii + 1) = PLeft(ii) + 1350
   
   ii = ii + 1
   PColName(ii) = "支付翻譯費"
   PLeft(ii + 1) = PLeft(ii) + 1500
   
   ii = ii + 1
   PColName(ii) = "財務點數"
   PLeft(ii + 1) = PLeft(ii) + 1350
End Sub

Private Sub PrintPageHeader()
   Dim strTmp As String
   Dim bFontBold As Boolean
   
   bFontBold = m_Device.Font.Bold 'Added by Morgan 2021/7/16 要記錄原來設定，否則若跳頁印表頭會被改
   
   'Modify By Sindy 2010/9/16
   iPage = iPage + 1
   m_iPages = m_iPages + 1
   If m_iPages > 1 Then
      If m_bPrinter = False Then
         SetPic m_iPages - 1
      ElseIf iPage > 1 Then
         m_Device.NewPage
      End If
   End If
   '2010/9/16 End
   
   m_Device.FontName = "新細明體"
   
   iPrint = m_iStartY
   m_Device.Font.Size = 14
   m_Device.Font.Bold = True
   'm_Device.Font.Underline = True
   
   strTmp = strTitle
   m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strTmp)) / 2
   m_Device.CurrentY = iPrint
   m_Device.Print strTmp
   
   PrintNewLine 500
   
   m_Device.Font.Size = 12
   m_Device.Font.Bold = True
   m_Device.Font.Underline = False
   
   strTmp = "帳款日期:" & strCon1 & " ~ " & strCon2
   m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strTmp)) / 2
   m_Device.CurrentY = iPrint
   m_Device.Print strTmp
   
   PrintNewLine 400
   
   strTmp = "列印人員:" & strUserName
   m_Device.CurrentX = m_iStartX
   m_Device.CurrentY = iPrint
   m_Device.Print strTmp
   
   If m_strSystem <> "" Then
      strTmp = "系統類別:" & m_strSystem
      m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strTmp)) / 2
      m_Device.CurrentY = iPrint
      m_Device.Print strTmp
   End If
   
   strTmp = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   m_Device.CurrentX = m_Device.ScaleWidth - m_iMargin - 2500
   m_Device.CurrentY = iPrint
   m_Device.Print strTmp
   PrintNewLine
   
   'Added by Lydia 2017/10/18
   If strBaseTable <> "" Then
      strTmp = "類別：" & IIf(txtKind = "1", "申請人", "代理人") & String(5, "　")
      If Trim(txtNa(0) & txtNa(1)) <> "" Then
         strTmp = strTmp & "國籍：" & convForm(txtNa(0), 3) & " - " & convForm(txtNa(1), 3) & String(5, "　")
      End If
      If Trim(txtNArea) <> "" Then
         strTmp = strTmp & "洲別："
         Select Case txtNArea
             Case "0": strTmp = strTmp & "亞洲"
             Case "1": strTmp = strTmp & "美洲"
             Case "2": strTmp = strTmp & "歐洲"
             Case "3": strTmp = strTmp & "非洲"
             Case "4": strTmp = strTmp & "大洋洲"
         End Select
      End If
      m_Device.CurrentX = m_iStartX
      m_Device.CurrentY = iPrint
      m_Device.Print strTmp
   End If
   'end 2017/10/18
   
   strTmp = "頁    次：" & str(iPage)
   m_Device.CurrentX = m_Device.ScaleWidth - m_iMargin - 2500
   m_Device.CurrentY = iPrint
   m_Device.Print strTmp
   m_Device.Font.Size = 10
   PrintNewLine 400
   
   PrintPageHeader1
   
   'Modified by Morgan 2021/7/16 要改回原來設定
   'm_Device.Font.Bold = False
   m_Device.Font.Bold = bFontBold
   'end 2021/7/16
End Sub

Private Sub PrintPageHeader1()
   For intI = 1 To UBound(PColName)
      If intI > 4 Then
         m_Device.CurrentX = PLeft(intI + 1) - 70 - m_Device.TextWidth(PColName(intI))
      Else
         m_Device.CurrentX = PLeft(intI)
      End If
      m_Device.CurrentY = iPrint
      m_Device.Print PColName(intI)
   Next
   PrintNewLine
   For intI = 1 To UBound(PColName2)
      If PColName2(intI) <> "" Then
         If intI > 4 Then
            m_Device.CurrentX = PLeft(intI + 1) - 70 - m_Device.TextWidth(PColName2(intI))
         Else
            m_Device.CurrentX = PLeft(intI)
         End If
         m_Device.CurrentY = iPrint
         m_Device.Print PColName2(intI)
      End If
   Next
   PrintNewLine
   DrawLine
End Sub

Private Sub DrawLine(Optional iStartCol As Integer, Optional iEndCol As Integer, Optional lngEndPoint As Long)
   Dim lngFrom As Long, lngTo As Long
   If iStartCol = 0 Then
      lngFrom = PLeft(LBound(PLeft))
   Else
      lngFrom = PLeft(iStartCol)
   End If
   If iEndCol = 0 Then
      If lngEndPoint > 0 Then
         lngTo = lngEndPoint
      Else
         lngTo = PLeft(UBound(PLeft))
      End If
   Else
      lngTo = PLeft(iEndCol)
   End If
   m_Device.DrawWidth = 4
   m_Device.Line (lngFrom, iPrint)-(lngTo, iPrint)
   iPrint = iPrint - m_iLineHeight / 2
End Sub

Private Sub PrintNewLine(Optional ByVal iHeight As Integer = 0, Optional ByVal p_iExtraLines As Integer = 2)
   
   If iHeight = 0 Then
      iHeight = m_iLineHeight
   End If
   
   iPrint = iPrint + iHeight
   
   If iPrint >= (m_iPageHeight - m_iMargin - p_iExtraLines * m_iLineHeight) Then
      DrawLine
      'm_Device.NewPage
      PrintPageHeader
      iPrint = iPrint + m_iLineHeight
   End If
    
End Sub

Private Sub PrintDetail()
   Dim iCol As Integer
   
   PrintNewLine
   For iCol = LBound(strTemp) To UBound(strTemp)
      If iCol < 5 Then
         m_Device.CurrentX = PLeft(iCol)
      Else
         'Add by Amy 2022/03/18 +財務/電腦中心 進入字數太長縮小字
         If Pub_StrUserSt03 = "M31" Or Pub_StrUserSt03 = "M51" Then
            m_Device.Font.Size = 10
            If Len(strTemp(iCol)) >= 10 Then
               m_Device.Font.Size = 8
            End If
         End If
         'end 2022/03/18
         m_Device.CurrentX = PLeft(iCol + 1) - m_Device.TextWidth(strTemp(iCol)) - 70
      End If
      m_Device.CurrentY = iPrint
      m_Device.Print strTemp(iCol)
   Next
End Sub

'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)
   Dim iSize As Integer
   
   PrintNewLine , 3
   DrawLine
   PrintNewLine
   
   iSize = m_Device.Font.Size
   m_Device.Font.Size = 12
   m_Device.Font.Bold = True
   strExc(1) = "*** 結束 ***"
   m_Device.CurrentX = (m_Device.ScaleWidth - m_Device.TextWidth(strExc(1))) / 2
   m_Device.CurrentY = iPrint
   m_Device.Print strExc(1)
   PrintNewLine 500
   
   
   m_Device.CurrentX = m_iStartX
   m_Device.CurrentY = iPrint
   m_Device.Print "PS:每一筆請款單點數計入該請款單之最後收文的智權人員"
   
   'Add By Sindy 2010/9/16
   If Text8 = "2" Then
      PrintNewLine 300
      m_Device.CurrentX = m_iStartX
      m_Device.CurrentY = iPrint
      'modify by sonia 2021/1/20 +FCL分配點數之收款X10917210於110/1/12M11000096收款
      'm_Device.Print "     已收點數與財務點數之差異應為規費不請款或退費做收入等因素"
      m_Device.Print "     已收點數與財務點數之差異應為規費不請款或退費做收入或FCL分配點數等因素"
   End If
   
   m_Device.Font.Size = iSize
   m_Device.Font.Bold = False
End Sub

'Add by Morgan 2010/5/21
'列印分配點數合計
'Modify by Morgan 2010/9/7 +p_SalesNo
Private Sub PrintShareTot(Optional p_SalesNo As String, Optional p_PrintName As String)
   Dim stSQL As String, intR As Integer
   Dim adoRst As ADODB.Recordset
   
   stSQL = "select NVL(sum(R03),0) C1 from accrpt212_1 where ID='" & strUserNum & "'"
   
   If p_SalesNo <> "" Then
      stSQL = stSQL & " and R01='" & p_SalesNo & "'"
   End If
   
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With adoRst
      If .Fields(0) > 0 Then
         PrintNewLine
         'Add By Sindy 2010/10/5
         If p_PrintName <> "" Then
            DrawLine 4
            PrintNewLine
            m_Device.FontBold = True
            m_Device.CurrentX = PLeft(1)
            m_Device.CurrentY = iPrint
            m_Device.Print p_PrintName
            m_Device.FontBold = False
         End If
         '2010/10/5 End
         strExc(1) = "分配點數："
         m_Device.CurrentX = PLeft(7) - m_Device.TextWidth(strExc(1))
         m_Device.CurrentY = iPrint
         m_Device.Print strExc(1)
         strExc(1) = Format(Val("" & .Fields(0)), FDollar)
         m_Device.CurrentX = PLeft(8) - 70 - m_Device.TextWidth(strExc(1))
         m_Device.CurrentY = iPrint
         m_Device.Print strExc(1)
         
         'Add By Sindy 2010/10/5
         If p_SalesNo = "" Then
            PrintNewLine
            DrawLine 4
            PrintNewLine
            strExc(1) = "總計："
            m_Device.CurrentX = PLeft(7) - m_Device.TextWidth(strExc(1))
            m_Device.CurrentY = iPrint
            m_Device.Print strExc(1)
            strExc(1) = Format(Val("" & .Fields(0)), FDollar)
            'Modified by Lydia 2017/01/04 顯示小數點3位
            'm_Device.CurrentX = PLeft(8) - 70 - m_Device.TextWidth((strExc(1) + m_dbltotPoint))
            m_Device.CurrentX = PLeft(8) - 70 - m_Device.TextWidth(Format(CDbl(strExc(1)) + CDbl(m_dbltotPoint), FDollar))
            m_Device.CurrentY = iPrint
            'Modified by Lydia 2017/01/04 顯示小數點3位
            'm_Device.Print (strExc(1) + m_dbltotPoint)
            'Modify By Sindy 2018/7/3 + Val ==> CDbl:不然計算會錯
            m_Device.Print Format(CDbl(strExc(1)) + CDbl(m_dbltotPoint), FDollar)
         'Added by Morgan 2012/2/2
         ElseIf p_PrintName = "" Then
            
            PrintNewLine
            strExc(1) = "小計："
            m_Device.CurrentX = PLeft(7) - m_Device.TextWidth(strExc(1))
            m_Device.CurrentY = iPrint
            m_Device.Print strExc(1)
            strExc(1) = Format(CDbl("" & .Fields(0)) + CDbl(m_dbltotPoint), FDollar)
            m_Device.CurrentX = PLeft(8) - 70 - m_Device.TextWidth(strExc(1))
            m_Device.CurrentY = iPrint
            m_Device.Print strExc(1)
         'end 2012/2/2
         
         End If
         '2010/10/5 End
      End If
      End With
   End If
   
   Set adoRst = Nothing
End Sub

'Add by Morgan 2010/4/21
'列印分配點數
Private Sub PrintSharePoint(p_ID As String)
   Dim dblTot As Double
   'Modify By Sindy 2011/2/17 因用SQLDate排序或取MAX或MIN,修改百年蟲問題
'   strExc(0) = "select R01 C1,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 C2,R02 C3,sqldatet(a1k02) C4,R03 C5,st02 C6" & _
'      " from accrpt212_1,acc1k0,staff where id='" & strUserNum & "' and a1k01(+)=R02 and st01(+)=R01 and R01='" & p_ID & "'" & _
'      " order by 1,4,2"
   strExc(0) = "select R01 C1,a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 C2,R02 C3,sqldatet2(a1k02) C4,R03 C5,st02 C6" & _
      " from accrpt212_1,acc1k0,staff where id='" & strUserNum & "' and a1k01(+)=R02 and st01(+)=R01 and R01='" & p_ID & "'" & _
      " order by 1,4,2"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      PrintNewLine
      
      m_Device.CurrentX = PLeft(1)
      m_Device.CurrentY = iPrint
      m_Device.Font.Bold = True
      m_Device.Print "請款單分配點數："
      PrintNewLine
      
      For intI = 1 To 4
         m_Device.CurrentX = PLeft(intI)
         m_Device.CurrentY = iPrint
         m_Device.Print PColName(intI)
      Next
      m_Device.CurrentX = PLeft(5)
      m_Device.CurrentY = iPrint
      m_Device.Print "點數"
         
      PrintNewLine
      DrawLine 1, , PLeft(5) + 1000
      m_Device.Font.Bold = False
      
      With RsTemp
      Do While Not .EOF
         PrintNewLine
         If .AbsolutePosition = 1 Then
            '智權人員
            m_Device.CurrentX = PLeft(1)
            m_Device.CurrentY = iPrint
            m_Device.Print .Fields("C6")
         End If
         '本所案號
         m_Device.CurrentX = PLeft(2)
         m_Device.CurrentY = iPrint
         m_Device.Print .Fields("C2")
         '請款編號
         m_Device.CurrentX = PLeft(3)
         m_Device.CurrentY = iPrint
         m_Device.Print .Fields("C3")
         '請款日
         m_Device.CurrentX = PLeft(4)
         m_Device.CurrentY = iPrint
         m_Device.Print .Fields("C4")
         '點數
         strExc(1) = Format(Val("" & .Fields("C5")), FDollar)
         m_Device.CurrentX = PLeft(5) + 1000 - m_Device.TextWidth(strExc(1))
         m_Device.CurrentY = iPrint
         m_Device.Print strExc(1)
         
         dblTot = dblTot + CDbl("" & .Fields("C5"))
         .MoveNext
      Loop
      End With
      
      PrintNewLine
      DrawLine 1, , PLeft(5) + 1000
      
      PrintNewLine
      strExc(1) = Format(Val(dblTot), FDollar)
      m_Device.CurrentX = PLeft(5) + 1000 - m_Device.TextWidth(strExc(1))
      m_Device.CurrentY = iPrint
      m_Device.Print strExc(1)
   End If
End Sub

'Add by Morgan 2010/5/20
Private Sub AddSharePoint()
   Dim stCon As String
   
   strSql = "delete accrpt212_1 where id='" & strUserNum & "'"
   cnnConnection.Execute strSql, intI
   
   stCon = strSystem
   
   If Text1 <> "" Then
      stCon = stCon & " and a1n04='" & Text1 & "'"
   End If
   
   If Text9 <> "" Then
      stCon = stCon & " and st15>='" & Text9 & "'"
   End If
   
   If Text10 <> "" Then
      stCon = stCon & " and st15<='" & Text10 & "'"
   End If
   
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      stCon = stCon & " and a1k02 >= " & Val(ChangeTDateStringToTString(Replace(MaskEdBox1.Text, "_", ""))) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      stCon = stCon & " and a1k02 <= " & Val(ChangeTDateStringToTString(Replace(MaskEdBox2.Text, "_", ""))) & ""
   End If
   
   'Modify By sindy 2010/10/14 加  or cp13 is null
   'Modified by Morgan 2017/4/12 抓非最後收文業務的
   'strSql = "insert into accrpt212_1(id,R01,R02,R03)" & _
      " select '" & strUserNum & "',a1n04,a1k01,sum(a1n05) from acc1k0,acc1n0,caseprogress,staff" & _
      " where a1k12 is null and a1k25 is null and a1n01(+)=a1k01 and a1n02='1' and cp09(+)=a1n03 and (a1n04<>cp13 or cp13 is null)" & _
      " and st01(+)=a1n04" & stCon & " group by a1n04,a1k01"
   'Modified by Lydia 2017/10/18 增加申請人/代理人之國籍/洲別
   'strSql = "insert into accrpt212_1(id,R01,R02,R03)" & _
      " select '" & strUserNum & "',a1n04,a1k01,sum(a1n05) from acc1k0,acc1n0,caseprogress a,staff" & _
      " where a1k12 is null and a1k25 is null and a1n01(+)=a1k01 and a1n02='1' and cp09(+)=a1n03" & _
      " and not exists(select 1 from caseprogress b where b.cp60=a1n01 having substr(max(cp05||cp09||b.cp13),18)=a1n04)" & _
      " and st01(+)=a1n04" & stCon & " group by a1n04,a1k01"
   ''end 2017/4/12
   strSql = " select '" & strUserNum & "',a1n04,a1k01,sum(a1n05) from acc1k0,acc1n0,caseprogress a,staff" & _
      strBaseTable & " where a1k12 is null and a1k25 is null and a1n01(+)=a1k01 and a1n02='1' and cp09(+)=a1n03" & _
      " and not exists(select 1 from caseprogress b where b.cp60=a1n01 having substr(max(cp05||cp09||b.cp13),18)=a1n04)" & _
      " and st01(+)=a1n04" & stCon & strCaseNa & " group by a1n04,a1k01"
   strSql = "insert into accrpt212_1(id,R01,R02,R03) " & strSql
   'end 2017/10/18
   
   'Modify By Sindy 2010/12/22
   'cnnConnection.Execute strSql, intI
   cnnConnection.Execute strSql, m_intPointCnt
End Sub

'Added by Lydia 2017/10/18
Private Sub txtkind_GotFocus()
   TextInverse txtKind
End Sub

Private Sub txtKind_Validate(Cancel As Boolean)
   If Trim(txtKind) <> "" Then
      If txtKind <> "1" And txtKind <> "2" Then
         MsgBox "類別請輸入1-2 !", vbCritical
         Cancel = True
         txtKind.SetFocus
      End If
   End If
End Sub

Private Sub txtNa_GotFocus(Index As Integer)
   TextInverse txtNa(Index)
End Sub

Private Sub txtNa_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtNa_LostFocus(Index As Integer)
  If Index = 0 Then
     If Trim(txtNa(0)) <> "" Then txtNa(1) = Trim(txtNa(0))
  End If
End Sub

Private Sub txtNArea_GotFocus()
   TextInverse txtNArea
End Sub

Private Sub txtNArea_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtNArea_Validate(Cancel As Boolean)
   If Trim(txtNArea) <> "" Then
      If InStr("0,1,2,3,4", txtNArea) = 0 Then
         MsgBox "洲別請輸入0-4 !", vbCritical
         Cancel = True
         txtNArea.SetFocus
      End If
   End If
End Sub
'end 2017/10/18

'Added by Lydia 2020/09/01 請款點數: 產生FC業務請款無申請人/代理人案件清單
Private Sub ProcExcelSave(ByVal pSQL As String)
   Dim xlsReport As New Excel.Application
   Dim wksReport As New Worksheet
   Dim intQ As Integer, stSQL As String
   Dim rsQuery As New ADODB.Recordset
   Dim strGrp As String, nPages As Integer
   Dim nRows As Integer
   Dim arrTmp As Variant, arrTmpW As Variant
   Dim stDate As String
   Dim xlsFileName As String
   Dim strTo As String

On Error GoTo ErrHnd

   If Dir(App.path & "\" & strUserNum, vbDirectory) = "" Then
       MkDir App.path & "\" & strUserNum
   End If
   xlsFileName = "FC業務請款無" & IIf(txtKind.Text = "1", "申請人", "代理人") & "案件清單.xls"
   Call PUB_KillTempFile(strUserNum & "\" & xlsFileName)
   xlsFileName = App.path & "\" & strUserNum & "\" & xlsFileName
    
   '欄位抬頭
   stSQL = "本所案號,收 文 日,收 文 號,請款單號,請款點數"
   arrTmp = Split(stSQL, ",")
   stSQL = "12,10,10,10,10"
   arrTmpW = Split(stSQL, ",")
   
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, pSQL)
   If intQ = 1 Then
       rsQuery.MoveFirst
       Do While Not rsQuery.EOF
            If strGrp <> "" & rsQuery.Fields(0) Then
                 nPages = nPages + 1
                 If strGrp = "" Then
                     xlsReport.SheetsInNewWorkbook = 1 '預設工作表數目
                     xlsReport.Workbooks.add
                     xlsReport.Application.WindowState = xlMinimized
                 Else
                     xlsReport.Worksheets.add  '插入sheet
                 End If
                 xlsReport.Worksheets("工作表" & nPages).Select
                 xlsReport.Worksheets("工作表" & nPages).Name = "" & rsQuery.Fields(0)
                 Set wksReport = xlsReport.Worksheets("" & rsQuery.Fields(0))
           
                 '設定欄位名稱及欄寬
                 nRows = 1
                 For intQ = 1 To UBound(arrTmp) + 1
                     wksReport.Range(Chr(intQ + 64) & nRows).Value = arrTmp(intQ - 1)
                     wksReport.Range(Chr(intQ + 64) & ":" & Chr(intQ + 64)).ColumnWidth = Val(arrTmpW(intQ - 1))
                     wksReport.Range(Chr(intQ + 64) & nRows).HorizontalAlignment = xlCenter
                 Next
                 nRows = nRows + 1
                 strGrp = "" & rsQuery.Fields(0)
            End If
            For intQ = 1 To UBound(arrTmp) + 1
                With wksReport.Range(Chr(intQ + 64) & nRows)
                    If intQ < 5 Then
                       .Value = "" & rsQuery.Fields(intQ)
                       .NumberFormatLocal = "@"
                       wksReport.Range(Chr(intQ + 64) & nRows).HorizontalAlignment = xlLeft
                    Else
                       .Value = "" & rsQuery.Fields(intQ)
                       .NumberFormatLocal = "##,##0.000"
                       wksReport.Range(Chr(intQ + 64) & nRows).HorizontalAlignment = xlRight
                    End If
                End With
            Next intQ
            nRows = nRows + 1
            rsQuery.MoveNext
       Loop
       If Val(xlsReport.Version) < 12 Then
          xlsReport.Workbooks(1).SaveAs FileName:=xlsFileName, FileFormat:=-4143
       Else
          xlsReport.Workbooks(1).SaveAs FileName:=xlsFileName, FileFormat:=56
       End If
       xlsReport.Workbooks.Close
       xlsReport.Quit
        
       If Dir(xlsFileName) <> "" Then
           PUB_SendMail strUserNum, strUserNum, "", Replace(MaskEdBox1.Text, "/", "") & "-" & Replace(MaskEdBox2.Text, "/", "") & " FC業務請款無" & IIf(txtKind.Text = "1", "申請人", "代理人") & "案件清單", vbCrLf & "請參考附件內容。", , xlsFileName
       End If
   End If

   Set rsQuery = Nothing
   Set xlsReport = Nothing
   Exit Sub
    
ErrHnd:
    
    WLog "結餘單流水號檢查:" & Err.Description
    If Val(xlsReport.Version) < 12 Then
       xlsReport.Workbooks(1).SaveAs FileName:=xlsFileName, FileFormat:=-4143
    Else
       xlsReport.Workbooks(1).SaveAs FileName:=xlsFileName, FileFormat:=56
    End If
    xlsReport.Workbooks.Close
    xlsReport.Quit
    Set xlsReport = Nothing
    Set rsQuery = Nothing
End Sub
