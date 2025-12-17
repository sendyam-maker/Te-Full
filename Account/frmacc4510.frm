VERSION 5.00
Begin VB.Form Frmacc4510 
   AutoRedraw      =   -1  'True
   Caption         =   "財產目錄表"
   ClientHeight    =   3696
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5148
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3696
   ScaleWidth      =   5148
   Begin VB.TextBox txtDate 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   1
      Top             =   375
      Width           =   855
   End
   Begin VB.TextBox txtDate 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   0
      Top             =   375
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含已報廢"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   360
      TabIndex        =   12
      Top             =   2040
      Value           =   1  '核取
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   2
      Top             =   840
      Width           =   612
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   3000
      Width           =   4692
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1440
      Width           =   612
   End
   Begin VB.Label Label3 
      Caption         =   "備註: 統計期間為當月起迄即列印所有資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "統計期間　　　　 ～"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   360
      TabIndex        =   13
      Top             =   435
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "空白.全部)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   2040
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "3.電腦硬體 4.電腦軟體 "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   2040
      TabIndex        =   10
      Top             =   1755
      Width           =   2250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1. 交通運輸設備 2.生財器具"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   2040
      TabIndex        =   9
      Top             =   1455
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   " 4.高所 5.其他 空白.全部)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   2040
      TabIndex        =   8
      Top             =   1170
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1.北所 2.中所 3.南所"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2040
      TabIndex        =   7
      Top             =   885
      Width           =   2040
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "類　別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1470
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "所在地"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   885
      Width           =   675
   End
End
Attribute VB_Name = "Frmacc4510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/11/30 Form2.0已檢查 (無需修改的物件)
'Create by Lydia 2017/03/15  財產目錄表
Option Explicit
Dim rsAD As New ADODB.Recordset
Dim intQ As Integer

Private Sub Command1_Click()

   'Added by Lydia 2019/03/07 檢查
   Dim bolTmp As Boolean
   If Trim(txtDate(0) & txtDate(1)) = "" Then
      txtDate(0) = Left(strSrvDate(2), 5)
      txtDate(1) = Left(strSrvDate(2), 5)
   Else
      If Val(txtDate(0)) > Val(txtDate(1)) Then
          MsgBox "統計期間起值不可大於迄值！", vbCritical
          txtDate(0).SetFocus
          Call txtDate_GotFocus(0)
          Exit Sub
      Else
          For intI = 0 To 1
            Call txtDate_Validate(intI, bolTmp)
            If bolTmp = True Then
                Exit Sub
            End If
          Next
      End If
   End If
   'end 2019/03/07
   
   Screen.MousePointer = vbHourglass
   Call ProcExcelSave
   Screen.MousePointer = vbDefault
   FormClear
End Sub

Private Sub Form_Load()
   '表單初始化
   'Modified by Lydia 2019/03/07
   'PUB_InitForm Me, 5260, 3450, strBackPicPath4
   PUB_InitForm Me, 5260, 4100, strBackPicPath4
   
   FormClear
   
   txtDate(0) = Left(strSrvDate(2), 5)
   txtDate(1) = Left(strSrvDate(2), 5)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear

   Set rsAD = Nothing
   Set Frmacc4510 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If IsEmpty(Text1) = False Then
       If InStr("1,2,3,4,5", Text1) = 0 Then
          MsgBox "請輸入1~5 !"
          Text1.SetFocus
          Cancel = True
       End If
    End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If IsEmpty(Text2) = False Then
       If InStr("1,2,3,4", Text2) = 0 Then
          MsgBox "請輸入1~4 !"
          Text2.SetFocus
          Cancel = True
       End If
    End If
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   
   Check1.Value = 0
   
End Sub

'產生Excel檔案
Private Sub ProcExcelSave()
Dim xlsA4510 As New Excel.Application
Dim wksA4510 As New Worksheet
Dim stCon As String
Dim strFileName As String
Dim iRow As Integer
Dim stCellFormat As String
Dim pKind As String '類別-分工作表
Dim intPage As Integer '工作表編號
Dim strColName As String '欄位名稱
Dim strColW As String    '欄寬
Dim tmpArr1 As Variant, tmpArr2 As Variant
Dim strUamt As String '每月固定傳票已攤金額
Dim strTmpTotal As String '財產的未折減餘額
Dim CntPage As Integer  'Excel檔的工作表數量
Dim xCols As Integer '行位置
'Added by Lydia 2017/05/23
Dim strGrp As String '記錄公司別
Dim strFileNames As String '檔案名稱
'end 2017/05/23
Dim sRow As Integer 'Added by Lydia 2017/05/31 欄位起始列
Dim EndX As String 'Added by Lydia 2017/09/27 最後一欄

On Error GoTo ErrHnd

If Check1.Value = 0 Then
   'Modified by Lydia 2017/05/26 拿掉摘要
   'strColName = "編號,名稱,所在地,取得日期,取得原價,使用月份,首筆折舊,本年提列,累計數,未折減餘額,摘要"
   'strColW = "7,20,8,9,10,8,10,10,10,10,20"
   'Modified by Lydia 2017/09/27 增加欄位
   'strColName = "編號,名稱,所在地,取得日期,取得原價,使用月份,首筆折舊,本年提列,累計數,未折減餘額"
   'strColW = "6,16,6,8,8.5,8.5,8.5,8.5,8.5,10"
   'Modified by Lydia 2019/03/07 本月提列=>本期已提列, +最後提列日期
   'strColName = "編號,名稱,所在地,取得日期,取得原價,使用月份,首筆折舊,本月提列,本年累積提列,累計數,未折減餘額"
   'strColW = "6,16,6,8,8.5,8.5,8.5,8.5,11.5,8.5,10"
   'EndX = "K" 'Added by Lydia 2017/09/27
   strColName = "編號,名稱,所在地,取得日期,取得原價,使用月份,首筆折舊,本期已提列,本年累積提列,累計數,未折減餘額,最後提列日期"
   strColW = "6,16,6,8,8.5,8.5,8.5,10,11.5,8.5,10,8"
   EndX = "L"
Else
   'Modified by Lydia 2017/05/26 拿掉摘要
   'strColName = "編號,名稱,所在地,取得日期,報廢日期,取得原價,使用月份,首筆折舊,本年提列,累計數,未折減餘額,摘要"
   'strColW = "7,20,8,9,9,8,10,10,10,10,10,20"
    'Modified by Lydia 2017/09/27 增加欄位
   'strColName = "編號,名稱,所在地,取得日期,取得原價,使用月份,首筆折舊,本年提列,累計數,未折減餘額,報廢日期"
   'strColW = "6,16,6,8,8.5,8.5,8.5,8.5,8.5,10,8"
   'Modified by Lydia 2019/03/07 本月提列=>本期已提列, +最後提列日期
   'strColName = "編號,名稱,所在地,取得日期,取得原價,使用月份,首筆折舊,本月提列,本年累積提列,累計數,未折減餘額,報廢日期"
   'strColW = "6,16,6,8,8.5,8.5,8.5,8.5,11.5,8.5,10,8"
   'EndX = "L" 'Added by Lydia 2017/09/27
   strColName = "編號,名稱,所在地,取得日期,取得原價,使用月份,首筆折舊,本期已提列,本年累積提列,累計數,未折減餘額,最後提列日期,報廢日期"
   strColW = "6,16,6,8,8.5,8.5,8.5,10,11.5,8.5,10,8,8"
   EndX = "M"
End If

tmpArr1 = Split(strColName, ",")
tmpArr2 = Split(strColW, ",")
stCellFormat = "#,##0 ;[紅色]-#,##0 "

sRow = 5 'Added by Lydia 2017/05/31

   '所在地
   If Text1 <> "" Then
      stCon = stCon & " AND A2B03=" & CNULL(Text1)
   End If
   '類別
   If Text2 <> "" Then
      stCon = stCon & " AND A2B02=" & CNULL(Text2)
   End If
   '是否含已報廢
   If Check1.Value = 0 Then
      stCon = stCon & " AND NVL(A2B19,0)=0"
   End If
   'Added by Lydia 2019/03/07 財產取得日期超過統計期間迄，則不列出
   'Modified by Lydia 2019/11/07 debug
   'stCon = stCon & " AND A2B05<=" & txtDate(0) & "31"
   stCon = stCon & " AND A2B05<=" & txtDate(1) & "31"
   
   '抓財產目錄資料,FSTAMT=首筆折舊金額,UAMT=本年累計提列
   'Modified by Lydia 2017/05/23 排序+A2B16
   'Modified by Lydia 2017/05/24 直接抓固定傳票餘額AXD14=固定傳票已攤金額
   'strSql = " SELECT A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09, A2B16,A2B17,A2B18,A2B19,VT02 FSTAMT,NVL(SUM(A1P07),0) UAMT " & _
            "FROM ACC2B0,(SELECT A1P01||A1P02||A1P04 VT01,NVL(SUM(A1P07),0) VT02 FROM ACC1P0 WHERE A1P02='M' GROUP BY A1P01||A1P02||A1P04) VT1,ACC1P0 " & _
            "WHERE A2B16||'M'||A2B01||A2B05=VT01(+) " & stCon & _
            "AND A2B16=A1P01(+) AND 'U'=A1P02(+) AND A1P04(+)>=A2B17||'" & Mid(strSrvDate(2), 1, 3) & "01' AND A1P04(+)<=A2B17||'" & Mid(strSrvDate(2), 1, 3) & "1231' " & _
            "GROUP BY A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,A2B16,A2B17,A2B18,A2B19,VT02 " & _
            "ORDER BY A2B16,A2B02,A2B03,A2B05,A2B01 "
    'Modified by Lydia 2019/03/07 改成輸入統計期間
    'strExc(0) = "SELECT /*+ INDEX(ACC1P0 IDXA1P020405) */ A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,A2B16,A2B17,A2B18,A2B19,A2B20,A2B21,A2B22,AXD13,AXD14,NVL(SUM(A1P07),0) UAMT " & _
                "FROM ACC2B0 A,ACC0D1,ACC1P0 WHERE A2B16=AXD01(+) AND A2B17=AXD02(+) " & stCon & _
                " AND A2B16=A1P01(+) AND 'U'=A1P02(+) AND A1P04(+)>=A2B17||'" & Mid(strSrvDate(2), 1, 3) & "01' AND A1P04(+)<=A2B17||'" & Mid(strSrvDate(2), 1, 3) & "1231' " & _
                "GROUP BY A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,A2B16,A2B17,A2B18,A2B19,A2B20,A2B21,A2B22,AXD13,AXD14"
    strExc(0) = "SELECT /*+ INDEX(ACC1P0 IDXA1P020405) */ A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,A2B16,A2B17,A2B18,A2B19,A2B20,A2B21,A2B22,AXD13,AXD14,NVL(SUM(A1P07),0) UAMT " & _
                "FROM ACC2B0 A,ACC0D1,ACC1P0 WHERE A2B16=AXD01(+) AND A2B17=AXD02(+) " & stCon & _
                " AND A2B16=A1P01(+) AND 'U'=A1P02(+) AND A1P04(+)>=A2B17||'" & Left(txtDate(0), 3) & "01' AND A1P04(+)<=A2B17||'" & Left(txtDate(1), 3) & "1231' " & _
                "GROUP BY A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,A2B16,A2B17,A2B18,A2B19,A2B20,A2B21,A2B22,AXD13,AXD14"
    'Modified by Lydia 2017/09/27 +本月提列
    'strSql = "SELECT A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,A2B16,A2B17,A2B18,A2B19,A2B20,A2B21,A2B22,AXD13,AXD14,UAMT,DECODE(A2B22,NULL,SUM(A1P07),SUM(AX206)) FSTAMT,UAMT,A0802 " & _
             "FROM (" & strExc(0) & ") X,ACC1P0,ACC021,ACC080 " & _
             "WHERE A2B16=A1P01(+) AND 'M'=A1P02(+) AND A2B01||A2B05=A1P04(+) AND A2B16=AX201(+) AND A2B22=AX202(+) AND (SUBSTR(AX205,1,4)='6126' OR AX205 IS NULL) " & _
             "AND A2B16=A0801(+) GROUP BY A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,A2B16,A2B17,A2B18,A2B19,A2B20,A2B21,A2B22,AXD13,AXD14,UAMT,A0802 " & _
             "ORDER BY A2B16,A2B02,A2B03,A2B05,A2B01"
    'Modified by Lydia 2019/03/07 改成輸入統計期間; 另外抓固定傳票的有效日期EXPDATE
    'strExc(1) = "SELECT /*+ INDEX(ACC1P0 IDXA1P020405) */ A2B01 as N01,NVL(SUM(A1P07),0) NAMT " & _
                "FROM ACC2B0 A,ACC0D1,ACC1P0 WHERE A2B16=AXD01(+) AND A2B17=AXD02(+) " & stCon & _
                " AND A2B16=A1P01(+) AND 'U'=A1P02(+) AND A1P04(+)>=A2B17||'" & Mid(strSrvDate(2), 1, 5) & "' AND A1P04(+)<=A2B17||'" & Mid(strSrvDate(2), 1, 5) & "31' GROUP BY A2B01 "
    'strSql = "SELECT A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,A2B16,A2B17,A2B18,A2B19,A2B20,A2B21,A2B22,AXD13,AXD14,UAMT,DECODE(A2B22,NULL,SUM(A1P07),SUM(AX206)) FSTAMT,UAMT,A0802,NAMT " & _
             "FROM (" & strExc(0) & ") X,ACC1P0,ACC021,ACC080,(" & strExc(1) & ") N " & _
             "WHERE A2B01=N01 AND A2B16=A1P01(+) AND 'M'=A1P02(+) AND A2B01||A2B05=A1P04(+) AND A2B16=AX201(+) AND A2B22=AX202(+) AND (SUBSTR(AX205,1,4)='6126' OR AX205 IS NULL) " & _
             "AND A2B16=A0801(+) GROUP BY A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,A2B16,A2B17,A2B18,A2B19,A2B20,A2B21,A2B22,AXD13,AXD14,UAMT,A0802,NAMT " & _
             "ORDER BY A2B16,A2B02,A2B03,A2B05,A2B01"
    ''end 2017/09/27
    strExc(1) = "SELECT /*+ INDEX(ACC1P0 IDXA1P020405) */ A2B01 as N01,NVL(SUM(A1P07),0) NAMT,AXD12||LPAD(AXD03,2,'0') AS EXPDATE  " & _
                "FROM ACC2B0 A,ACC0D1,ACC1P0 WHERE A2B16=AXD01(+) AND A2B17=AXD02(+) " & stCon & _
                " AND A2B16=A1P01(+) AND 'U'=A1P02(+) AND A1P04(+)>=A2B17||'" & txtDate(0) & "' AND A1P04(+)<=A2B17||'" & txtDate(1) & "31' GROUP BY A2B01,AXD12||LPAD(AXD03,2,'0') "
    'Modified by Lydia 2021/01/14 1公司顯示2公司名稱
    'strSql = "SELECT A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,A2B16,A2B17,A2B18,A2B19,A2B20,A2B21,A2B22,AXD13,AXD14,UAMT,DECODE(A2B22,NULL,SUM(A1P07),SUM(AX206)) FSTAMT,A0802,NAMT,EXPDATE " & _
             "FROM (" & strExc(0) & ") X,ACC1P0,ACC021,ACC080,(" & strExc(1) & ") N " & _
             "WHERE A2B01=N01 AND A2B16=A1P01(+) AND 'M'=A1P02(+) AND A2B01||A2B05=A1P04(+) AND A2B16=AX201(+) AND A2B22=AX202(+) AND (SUBSTR(AX205,1,4)='6126' OR AX205 IS NULL) " & _
             "AND A2B16=A0801(+) GROUP BY A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,A2B16,A2B17,A2B18,A2B19,A2B20,A2B21,A2B22,AXD13,AXD14,UAMT,A0802,NAMT,EXPDATE "
    'end 2019/03/07
    strSql = "SELECT A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,decode(A2B16,'1','2',A2B16)A2B16,A2B17,A2B18,A2B19,A2B20,A2B21,A2B22,AXD13,AXD14,UAMT,DECODE(A2B22,NULL,SUM(A1P07),SUM(AX206)) FSTAMT,A0802,NAMT,EXPDATE " & _
             "FROM (" & strExc(0) & ") X,ACC1P0,ACC021,ACC080,(" & strExc(1) & ") N " & _
             "WHERE A2B01=N01 AND A2B16=A1P01(+) AND 'M'=A1P02(+) AND A2B01||A2B05=A1P04(+) AND A2B16=AX201(+) AND A2B22=AX202(+) AND (SUBSTR(AX205,1,4)='6126' OR AX205 IS NULL) " & _
             "AND A2B16=A0801(+) GROUP BY A2B01,A2B02,A2B04,A2B03,A2B05,A2B06,A2B07,A2B08,A2B09,decode(A2B16,'1','2',A2B16),A2B17,A2B18,A2B19,A2B20,A2B21,A2B22,AXD13,AXD14,UAMT,A0802,NAMT,EXPDATE "
             
    'Modified by Lydia 2020/01/08 改排列順序: 公司別A2B16 , 類別A2B02, 取得日期A2B05, 所在地A2B03, 財務目錄編號A2B01
    'strSql = strSql & "ORDER BY A2B16,A2B02,A2B03,A2B05,A2B01"
    strSql = strSql & "ORDER BY A2B16,A2B02,A2B05,A2B03,A2B01"
    
   intQ = 0
   Set rsAD = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
      If rsAD.RecordCount > 0 Then
         rsAD.MoveFirst
         CntPage = Val(Pub_GetA2b02Name) + 1
         Do While Not rsAD.EOF
            strUamt = "0"
            '抓每月固定傳票已攤金額
            'Modified by Lydia 2017/05/22 +指定會計科目6126
            'Modified by Lydia 2017/05/26 已攤金額=總額-餘額
            'If "" & rsAd.Fields("A2B17") <> "" Then strUamt = PUB_SumA1PtoU("" & rsAd.Fields("A2B16"), "" & rsAd.Fields("A2B17"), , , "6126")
            If "" & rsAD.Fields("A2B17") <> "" Then strUamt = Val("" & rsAD.Fields("AXD13")) - Val("" & rsAD.Fields("AXD14"))
            
            '未折減餘額
            strTmpTotal = Val("" & rsAD.Fields("A2B06")) - Val("" & rsAD.Fields("FSTAMT")) - Val(strUamt)
            If "" & rsAD.Fields("A2B17") = "459" Then
                strTmpTotal = strTmpTotal
            End If
            
            'Added by Lydia 2017/05/26 有固定傳票直接抓固定傳票的餘額
            If "" & rsAD.Fields("A2B17") <> "" Then
               strTmpTotal = Val("" & rsAD.Fields("AXD14"))
            '無固定傳票,並且已過攤提期間, 餘額設為零
            ElseIf "" & rsAD.Fields("A2B22") = "" And Val("" & rsAD.Fields("A2B21") & "01") <= strSrvDate(2) Then
               strTmpTotal = 0
            Else
               strTmpTotal = IIf(strTmpTotal < 0, 0, strTmpTotal) 'Added by Lydia 2017/05/22 殘值小於零,算零
            End If 'end 2017/05/26
            
            '無殘值,不顯示
            'Remove by Lydia 2017/05/24 最初提到"不用留殘值",不是餘額為零不顯示
            'If Check1.Value = 0 And Val(strTmpTotal) = 0 Then
            '   GoTo JumpNextRec
            'End If
            'end 2017/05/27
            
            'Added by Lydia 2017/05/22 不同公司分不同檔案
            If strGrp <> "" And strGrp <> "" & rsAD.Fields("A2B16") Then
                '最後一頁-合計
                iRow = iRow + 3
                wksA4510.Range("B" & iRow).Value = Pub_GetA2b02Name(pKind) & "合計"
                xCols = Asc("E")
                'Remove by Lydia 2017/05/31 報廢日期移到最後
                'If Check1.Value = 1 Then
                '  xCols = xCols + 1
                'End If
                
                '使用原價
                'Modified by Lydia 2017/09/27 4=>sRow + 1
                wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & sRow + 1 & ":" & Chr(xCols) & iRow - 1 & ")"
                wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
                xCols = xCols + 1
                
                xCols = xCols + 1 '跳過使用月份
                
                'Added by Lydia 2017/05/31 首筆折舊
                'Modified by Lydia 2017/09/27 4=>sRow + 1
                wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & sRow + 1 & ":" & Chr(xCols) & iRow - 1 & ")"
                wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
                xCols = xCols + 1
                'end 2017/05/31
                
                'Added by Lydia 2017/09/27 本月提列
                wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & sRow + 1 & ":" & Chr(xCols) & iRow - 1 & ")"
                wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
                xCols = xCols + 1
                'end 2017/09/27
                
                '本年累計提列
                'Modified by Lydia 2017/09/27 4=>sRow + 1
                wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & sRow + 1 & ":" & Chr(xCols) & iRow - 1 & ")"
                wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
                xCols = xCols + 1
                '累計數
                'Modified by Lydia 2017/09/27 4=>sRow + 1
                wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & sRow + 1 & ":" & Chr(xCols) & iRow - 1 & ")"
                wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
                xCols = xCols + 1
                '未折減餘額
                'Modified by Lydia 2017/09/27 4=>sRow + 1
                wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & sRow + 1 & ":" & Chr(xCols) & iRow - 1 & ")"
                wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
                xCols = xCols + 1
                
                'Added by Lydia 2017/05/26 框線
                wksA4510.Range("1:" & iRow).RowHeight = 22  '調整列高
                wksA4510.Range(sRow & ":" & sRow).RowHeight = 36 'Added by Lydia 2019/03/07 抬頭欄位
                'Modifeid by Lydia 2017/09/27 改成變數
                'wksA4510.Range("A" & sRow & ":" & IIf(Check1.Value = 1, "K", "J") & iRow).Select
                wksA4510.Range("A" & sRow & ":" & EndX & iRow).Select
                xlsA4510.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
                xlsA4510.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
                xlsA4510.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
                xlsA4510.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
                xlsA4510.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
                xlsA4510.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                wksA4510.Range("A2").Select
                xlsA4510.Sheets(1).Select '選擇工作表
                'end 2017/05/26
            '----------------------------------------
                '判斷版本
                If Val(xlsA4510.Version) < 12 Then
                     xlsA4510.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
                Else
                     xlsA4510.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
                End If
                xlsA4510.Workbooks.Close
                xlsA4510.Quit
                Set wksA4510 = Nothing
                Set xlsA4510 = Nothing
            
                intPage = 0
                pKind = ""
            End If
            'end 2017/05/22
            
            If pKind <> "" & rsAD.Fields("A2B02") Then
               intPage = intPage + 1
               If intPage > CntPage Then
                  MsgBox "財產目錄的類別超出工作表的數量!"
                  GoTo JumpNoPage
               End If
               If intPage = 1 Then
                  'Modified by Lydia 2017/05/23 不同公司分不同檔案
                  'strFileName = strExcelPath & strSrvDate(2) & Me.Caption & IIf(Check1.Value = 1, "(含已報廢)", "") & MsgText(43)
                  strExc(1) = ""
                  'Modified by Lydia 2021/01/14 1公司顯示2公司名稱
                  'If "" & rsAD.Fields("A2B16") = "1" Then
                  'Modified by Lydia 2022/01/18 現在有三家公司; 11006為法律所
                  'If "" & rsAD.Fields("A2B16") <> "J" Then
                  If "" & rsAD.Fields("A2B16") = "L" Then
                     strExc(1) = "_法律所"
                  ElseIf "" & rsAD.Fields("A2B16") <> "J" Then
                  'end 2022/01/18
                     strExc(1) = "_台一"
                  ElseIf "" & rsAD.Fields("A2B16") = "J" Then
                     strExc(1) = "_智權"
                  End If
                  
                  'Modified by Lydia 2019/03/07 +統計期間
                  'strFileName = strExcelPath & strSrvDate(2) & Me.Caption & strExc(1) & IIf(Check1.Value = 1, "(含已報廢)", "") & MsgText(43)
                  strFileName = strExcelPath & txtDate(0) & IIf(txtDate(0) <> txtDate(1), "-" & txtDate(1), "") & Me.Caption & strExc(1) & IIf(Check1.Value = 1, "(含已報廢)", "") & strSrvDate(2) & MsgText(43)
                  strFileNames = strFileNames & IIf(strFileNames <> "", "、", "") & strFileName
                  'end 2017/05/23
                  If Dir(strFileName) = MsgText(601) Then
                     If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
                        MkDir strExcelPath
                     End If
                  Else
                     Kill strFileName
                  End If
                  pKind = "" & rsAD.Fields("A2B02")
                  xlsA4510.SheetsInNewWorkbook = CntPage
                  xlsA4510.Workbooks.add
                  xlsA4510.Visible = False '預設不顯示
               Else  '合計
                  iRow = iRow + 3
                  wksA4510.Range("B" & iRow).Value = Pub_GetA2b02Name(pKind) & "合計"
                  
                  xCols = Asc("E")
                  'Remove by Lydia 2017/05/31 報廢日期移到最後
                  'If Check1.Value = 1 Then
                  '  xCols = xCols + 1
                  'End If
                  
                  '使用原價
                  'Modified by Lydia 2017/09/27 4=>sRow + 1
                  wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & sRow + 1 & ":" & Chr(xCols) & iRow - 1 & ")"
                  wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
                  xCols = xCols + 1
                  
                  xCols = xCols + 1 '跳過使用月份
                  
                  'Added by Lydia 2017/05/31 首筆折舊
                  'Modified by Lydia 2017/09/27 4=>sRow + 1
                  wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & sRow + 1 & ":" & Chr(xCols) & iRow - 1 & ")"
                  wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
                  xCols = xCols + 1
                  'end 2017/05/31
                  
                  'Added by Lydia 2017/09/27 本月提列
                  wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & sRow + 1 & ":" & Chr(xCols) & iRow - 1 & ")"
                  wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
                  xCols = xCols + 1
                  'end 2017/09/27
                
                  '本年累計提列
                  'Modified by Lydia 2017/09/27 4=>sRow + 1
                  wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & sRow + 1 & ":" & Chr(xCols) & iRow - 1 & ")"
                  wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
                  xCols = xCols + 1
                  '累計數
                  'Modified by Lydia 2017/09/27 4=>sRow + 1
                  wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & sRow + 1 & ":" & Chr(xCols) & iRow - 1 & ")"
                  wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
                  xCols = xCols + 1
                  '未折減餘額
                  'Modified by Lydia 2017/09/27 4=>sRow + 1
                  wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & sRow + 1 & ":" & Chr(xCols) & iRow - 1 & ")"
                  wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
                  xCols = xCols + 1
                  'Added by Lydia 2017/05/26 框線
                  wksA4510.Range("1:" & iRow).RowHeight = 22  '調整列高
                  wksA4510.Range(sRow & ":" & sRow).RowHeight = 36 'Added by Lydia 2019/03/07 抬頭欄位
                  'Modifeid by Lydia 2017/09/27 改成變數
                  'wksA4510.Range("A" & sRow & ":" & IIf(Check1.Value = 1, "K", "J") & iRow).Select
                  wksA4510.Range("A" & sRow & ":" & EndX & iRow).Select
                  xlsA4510.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
                  xlsA4510.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
                  xlsA4510.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
                  xlsA4510.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
                  xlsA4510.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
                  xlsA4510.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                  wksA4510.Range("A2").Select
                  'end 2017/05/26
               End If

               Set wksA4510 = xlsA4510.Worksheets(intPage)
               
               'Added by Lydia 2017/05/26 設定列印邊界
               xlsA4510.Sheets(intPage).Select '選擇工作表
               wksA4510.PageSetup.PaperSize = 9 'A4
               wksA4510.PageSetup.Orientation = 1 '直印
               'Modified by Lydia 2017/09/27 改變邊界
               'wksA4510.PageSetup.LeftMargin = xlsA4510.CentimetersToPoints(1)
               'wksA4510.PageSetup.RightMargin = xlsA4510.CentimetersToPoints(1)
               'wksA4510.PageSetup.TopMargin = xlsA4510.CentimetersToPoints(1.5)
               'wksA4510.PageSetup.BottomMargin = xlsA4510.CentimetersToPoints(1.5)
               'wksA4510.PageSetup.HeaderMargin = xlsA4510.CentimetersToPoints(1.5)
               'wksA4510.PageSetup.FooterMargin = xlsA4510.CentimetersToPoints(1.5)
               'end 2017/05/26
               wksA4510.PageSetup.LeftMargin = xlsA4510.CentimetersToPoints(0)
               wksA4510.PageSetup.RightMargin = xlsA4510.CentimetersToPoints(0)
               wksA4510.PageSetup.TopMargin = xlsA4510.CentimetersToPoints(1)
               wksA4510.PageSetup.BottomMargin = xlsA4510.CentimetersToPoints(1)
               wksA4510.PageSetup.HeaderMargin = xlsA4510.CentimetersToPoints(0.5)
               wksA4510.PageSetup.FooterMargin = xlsA4510.CentimetersToPoints(0.5)
               'end 2017/09/27
               'Added by Lydia 2017/09/27 縮放比例,列印頁面水平置中
               wksA4510.PageSetup.Zoom = 85
               wksA4510.PageSetup.CenterHorizontally = True
               'end 2017/09/27
        
               xlsA4510.Worksheets(intPage).Name = Pub_GetA2b02Name(rsAD.Fields("A2B02")) '工作表名稱

               'Added by Lydia 2017/05/26
               'Modified by Lydia 2017/09/27 改成變數
               'If Check1.Value = 1 Then
               '   wksA4510.Range("A:K").Font.Size = 11
               '   strExc(1) = "K"
               'Else
               '   wksA4510.Range("A:J").Font.Size = 11
               '   strExc(1) = "J"
               'End If
               'end 2017/05/26
               wksA4510.Range("A:" & EndX).Font.Size = 11
               'end 2017/09/27
               
               'Modified by Lydia 2017/05/26 抬頭設定
               'wksA4510.Range("A1").Value = "財產目錄-" & Pub_GetA2b02Name(rsAd.Fields("A2B02"))
               'With wksA4510.Range("A1:G1")
               'Modified by Lydia 2021/01/14 公司名稱改用模組控制
               'wksA4510.Range("A1").Value = "" & rsAD.Fields("A0802")
               wksA4510.Range("A1").Value = A0802Query("" & rsAD.Fields("A2B16"))
               'Modified by Lydia 2017/09/27 strExc(1)改成變數 endx
               With wksA4510.Range("A1:" & Chr(Asc(EndX) - 1) & "1")
                   .WrapText = False
                   .MergeCells = True
                   .HorizontalAlignment = xlCenter
                   .VerticalAlignment = xlBottom
                   .Font.Size = 16
                   .Font.Bold = True
                   'Remove by Lydia 2018/01/23 靖媗要求用細明體
                   '.Font.Name = "標楷體"
               End With
               'Added by Lydia 2019/03/07 統計期間
               wksA4510.Range("D2").Value = "統計期間：" & txtDate(0) & "～" & txtDate(1)
               With wksA4510.Range("D2:G2")
                   .WrapText = False
                   .MergeCells = True
                   .HorizontalAlignment = xlCenter
                   .VerticalAlignment = xlBottom
                   .Font.Size = 14
               End With
               'end 2019/03/07
               
               'Modified by Lydia 2019/03/07 A3=>D3
               wksA4510.Range("D3").Value = "財產目錄-" & Pub_GetA2b02Name(rsAD.Fields("A2B02"))
               'Modified by Lydia 2017/09/27 strExc(1)改成變數 endx
               'Modified by Lydia 2018/03/07
               'With wksA4510.Range("A3:" & Chr(Asc(EndX) - 1) & "3")
               With wksA4510.Range("D3:G3")
                   .WrapText = False
                   .MergeCells = True
                   .HorizontalAlignment = xlCenter
                   .VerticalAlignment = xlBottom
                   .Font.Size = 14 'Added by Lydia 2017/05/26
               End With
               'end 2017/05/26
               
               'Modified by Lydia 2017/05/26 抬頭設定
               'wksA4510.Range("H1") = ChangeTStringToTDateString(strSrvDate(2))
               'Modified by Lydia 2017/09/27 strExc(1)改成變數 endx
               'Modified by Lydia 2019/03/07 +列印日期：
               'wksA4510.Range(EndX & "3").Value =  ChangeTStringToTDateString(strSrvDate(2))
               'wksA4510.Range(EndX & "3").Font.Size = 14
               wksA4510.Range(Chr(Asc(EndX) - 2) & "3").Value = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
               wksA4510.Range(Chr(Asc(EndX) - 2) & "3").Font.Size = 14

               'Modified by Lydia 2017/05/31 改成變數控制列座標( 3=>sRow)
               wksA4510.Range("A" & sRow & ":" & EndX & sRow).WrapText = True '自動換列
               For intQ = 0 To UBound(tmpArr1)
                  If Trim(tmpArr1(intQ)) <> "" Then
                     wksA4510.Range(Chr(Asc("A") + intQ) & sRow).Value = Trim(tmpArr1(intQ))
                     wksA4510.Range(Chr(Asc("A") + intQ) & sRow).ColumnWidth = Val(tmpArr2(intQ))
                  End If
               Next intQ
               iRow = sRow
               'end 2017/05/31
            End If
            iRow = iRow + 1
            
            xCols = Asc("A")
            '編號
            With wksA4510.Range(Chr(xCols) & iRow)
                .Value = "" & rsAD.Fields("A2B01")
                .NumberFormatLocal = "@"
            End With
            xCols = xCols + 1
            
            '名稱
            With wksA4510.Range(Chr(xCols) & iRow)
                .Value = "" & rsAD.Fields("A2B04")
                .NumberFormatLocal = "@"
            End With
            xCols = xCols + 1
            
            '所在地
            With wksA4510.Range(Chr(xCols) & iRow)
                .Value = Pub_GetA2b03Sname("" & rsAD.Fields("A2B03"))
                .NumberFormatLocal = "@"
            End With
            xCols = xCols + 1
            
            '取得日期
            With wksA4510.Range(Chr(xCols) & iRow)
                .Value = ChangeTStringToTDateString("" & rsAD.Fields("A2B05"))
                .NumberFormatLocal = "@"
            End With
            xCols = xCols + 1
            
            '取得原價
            With wksA4510.Range(Chr(xCols) & iRow)
                .Value = Val("" & rsAD.Fields("A2B06"))
                .NumberFormatLocal = stCellFormat
            End With
            xCols = xCols + 1
            
            '使用月份
            With wksA4510.Range(Chr(xCols) & iRow)
                .Value = "" & rsAD.Fields("A2B07")
                .NumberFormatLocal = "##0"
            End With
            xCols = xCols + 1
            
            '首筆折舊
            With wksA4510.Range(Chr(xCols) & iRow)
                .Value = Val("" & rsAD.Fields("FSTAMT"))
                .NumberFormatLocal = stCellFormat
            End With
            xCols = xCols + 1
            
            'Added by Lydia 2017/09/27 本月提列
            With wksA4510.Range(Chr(xCols) & iRow)
                .Value = Val("" & rsAD.Fields("NAMT"))
                .NumberFormatLocal = stCellFormat
            End With
            xCols = xCols + 1
            'end 2017/09/27
            
            '本年累計提列
            With wksA4510.Range(Chr(xCols) & iRow)
                 'Modified by Lydia 2017/05/26 首筆傳票為今年加入本年累計提列
                 'Modified by Lydia 2022/01/18 以統計期間判斷本年
                 'If Mid(DBDATE("" & rsAD.Fields("A2B05")), 1, 4) = Mid(strSrvDate(1), 1, 4) Then
                 If Mid(DBDATE("" & rsAD.Fields("A2B05")), 1, 4) >= Mid(DBDATE(txtDate(0) & "01"), 1, 4) And Mid(DBDATE("" & rsAD.Fields("A2B05")), 1, 4) <= Mid(DBDATE(txtDate(1) & "01"), 1, 4) Then
                     .Value = Val("" & rsAD.Fields("UAMT")) + Val("" & rsAD.Fields("FSTAMT"))
                 Else
                 'end 2017/05/26
                     .Value = Val("" & rsAD.Fields("UAMT"))
                 End If   'end 2017/05/26
                .NumberFormatLocal = stCellFormat
            End With
            xCols = xCols + 1
            
            '累計數
            With wksA4510.Range(Chr(xCols) & iRow)
                'Modified by Lydia 2017/05/26 累計數=原價-未折減餘額
                '.Value = Val("" & rsAd.Fields("FSTAMT")) + Val(strUamt)
                .Value = Val("" & rsAD.Fields("A2B06")) - Val(strTmpTotal)
                .NumberFormatLocal = stCellFormat
            End With
            xCols = xCols + 1
            '未折減餘額
            With wksA4510.Range(Chr(xCols) & iRow)
                .Value = Val(strTmpTotal)
                .NumberFormatLocal = stCellFormat
            End With
            xCols = xCols + 1
            
            '摘要
            'Remove by Lydia 2017/05/26 拿掉摘要
            'With wksA4510.Range(Chr(xCols) & iRow)
            '    .Value = "" & rsAd.Fields("A2B09")
            '    .NumberFormatLocal = "@"
            'End With
            'xCols = xCols + 1
            'end 2017/05/26
            
            'Added by Lydia 2019/03/07 最後提列日期(有效日期)
            With wksA4510.Range(Chr(xCols) & iRow)
                .Value = ChangeTStringToTDateString("" & rsAD.Fields("EXPDATE"))
                .NumberFormatLocal = "@"
            End With
            xCols = xCols + 1
            
            '含已報廢
            'Move by Lydia 2017/05/31 移到最後
            If Check1.Value = 1 Then
                '報廢日期
                With wksA4510.Range(Chr(xCols) & iRow)
                    .Value = ChangeTStringToTDateString("" & rsAD.Fields("A2B19"))
                    .NumberFormatLocal = "@"
                End With
                xCols = xCols + 1
            End If
            'end 2017/05/31
            
           pKind = "" & rsAD.Fields("A2B02")
           strGrp = "" & rsAD.Fields("A2B16")  'Added by Lydia 2017/05/23
           
JumpNextRec:
           rsAD.MoveNext
        Loop
        '最後一頁-合計
        iRow = iRow + 3
        wksA4510.Range("B" & iRow).Value = Pub_GetA2b02Name(pKind) & "合計"
        xCols = Asc("E")
        'Remove by Lydia 2017/05/31 報廢日期移到最後
        'If Check1.Value = 1 Then
        '  xCols = xCols + 1
        'End If
        
        '使用原價
        wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & "4:" & Chr(xCols) & iRow - 1 & ")"
        wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
        xCols = xCols + 1
        
        xCols = xCols + 1 '跳過使用月份
        
        'Added by Lydia 2017/05/31 首筆折舊
        wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & "4:" & Chr(xCols) & iRow - 1 & ")"
        wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
        xCols = xCols + 1
        'end 2017/05/31
                
        'Added by Lydia 2017/09/27 本月提列
        wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & "4:" & Chr(xCols) & iRow - 1 & ")"
        wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
        xCols = xCols + 1
        'end 2017/09/27
        '本年累計提列
        wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & "4:" & Chr(xCols) & iRow - 1 & ")"
        wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
        xCols = xCols + 1
        '累計數
        wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & "4:" & Chr(xCols) & iRow - 1 & ")"
        wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
        xCols = xCols + 1
        '未折減餘額
        wksA4510.Range(Chr(xCols) & iRow).Formula = "=sum(" & Chr(xCols) & "4:" & Chr(xCols) & iRow - 1 & ")"
        wksA4510.Range(Chr(xCols) & iRow).NumberFormatLocal = stCellFormat
        xCols = xCols + 1
        
        'Added by Lydia 2017/05/26 框線
        wksA4510.Range("1:" & iRow).RowHeight = 22  '調整列高
        wksA4510.Range(sRow & ":" & sRow).RowHeight = 36 'Added by Lydia 2019/03/07 抬頭欄位
        'Modified by Lydia 2017/09/27 改成變數
        'wksA4510.Range("A" & sRow & ":" & IIf(Check1.Value = 1, "K", "J") & iRow).Select
        wksA4510.Range("A" & sRow & ":" & EndX & iRow).Select
        xlsA4510.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsA4510.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
        xlsA4510.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlsA4510.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
        xlsA4510.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
        xlsA4510.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        wksA4510.Range("A2").Select
        xlsA4510.Sheets(1).Select '選擇工作表
        'end 2017/05/26

JumpNoPage:
      End If
   'Added by Lydia 2018/04/10 查無資料
   Else
        Exit Sub
   'end 2018/04/10
   End If
   
   '判斷版本
   If Val(xlsA4510.Version) < 12 Then
        xlsA4510.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
        xlsA4510.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If

   xlsA4510.Workbooks.Close
   xlsA4510.Quit
   Set wksA4510 = Nothing
   Set xlsA4510 = Nothing

   'Modified by Lydia 2017/05/23 strFileName => strFileNames 多個檔案
   MsgBox "Excel檔案產生完成！（檔案位置：" & strFileNames & "）"
   
   Exit Sub

ErrHnd:

   MsgBox Err.Description
End Sub

'Added by Lydia 2019/03/07
Private Sub txtDate_GotFocus(Index As Integer)
   If Index = 1 Then
      If txtDate(Index).Text = "" Then txtDate(Index).Text = txtDate(Index - 1).Text
   End If
   TextInverse txtDate(Index)
End Sub

'Added by Lydia 2019/03/07
Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
   If txtDate(Index) <> "" Then
       If ChkDate(txtDate(Index) & "01") = False Then
           txtDate(Index).SetFocus
           Call txtDate_GotFocus(Index)
           Cancel = True
       End If
   End If
End Sub
