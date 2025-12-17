VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc44j0 
   AutoRedraw      =   -1  'True
   Caption         =   "智權人員點數明細表"
   ClientHeight    =   5292
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5292
   ScaleWidth      =   5520
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   870
      TabIndex        =   0
      Top             =   180
      Width           =   3500
   End
   Begin VB.CommandButton CmdMemo 
      Caption         =   "實績與結餘表說明"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3405
      TabIndex        =   27
      Top             =   555
      Width           =   2050
   End
   Begin VB.ComboBox Combo7 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   25
      Top             =   2750
      Width           =   4300
   End
   Begin VB.TextBox Text3 
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
      Left            =   2730
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1980
      Width           =   612
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
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
      Left            =   345
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   2295
      Width           =   4692
   End
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   555
      Width           =   3240
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
      Height          =   315
      Left            =   2730
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1650
      Width           =   612
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
      Height          =   300
      Left            =   1545
      TabIndex        =   4
      Top             =   1290
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1035
      TabIndex        =   2
      Top             =   930
      Width           =   1575
      _ExtentX        =   2794
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
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
      Left            =   2955
      TabIndex        =   3
      Top             =   930
      Width           =   1575
      _ExtentX        =   2794
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   2745
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      Left            =   75
      TabIndex        =   24
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "作業耗時，執行期間請勿開啟檔案！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   270
      TabIndex        =   23
      Top             =   3150
      Width           =   4800
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "是否列印報表(Y/N)"
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
      Left            =   75
      TabIndex        =   22
      Top             =   2010
      Width           =   2655
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   3120
      Picture         =   "Frmacc44j0.frx":0000
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -60
      Top             =   1740
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "3."
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
      Left            =   840
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   3120
      Picture         =   "Frmacc44j0.frx":0442
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1410
      Left            =   360
      Top             =   3765
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "2."
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
      Left            =   840
      TabIndex        =   20
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3120
      Picture         =   "Frmacc44j0.frx":0884
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "1."
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
      Left            =   840
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "排序方式"
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
      Left            =   480
      TabIndex        =   18
      Top             =   3810
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "是否產生Excel檔案(Y/N)"
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
      Left            =   75
      TabIndex        =   17
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "(空白表全部)"
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
      Left            =   3330
      TabIndex        =   16
      Top             =   1290
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(業)"
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
      Left            =   75
      TabIndex        =   15
      Top             =   1290
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2730
      TabIndex        =   14
      Top             =   930
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
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
      Index           =   1
      Left            =   75
      TabIndex        =   13
      Top             =   930
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc44j0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Moemo by Amy 2015/01/12 修改智權人員點數明細表 ***
'       1.保留點數改畫面止月保留點數(原自行輸入):抓畫面條件止日當月4191及4192之「借」方總和
'       2.增加起月保留點數:抓畫面條件起日當月4191及4192之「貸」方總和 '2015/01/16隱藏
'       3.增加 1-2計算公式欄位 '2015/01/16隱藏
'       4.其他人員M0100只需計算P的保留點數(因目前只有P有資料)
'       5.P0100大陸P 抓P案及服務業業PS /P0100大陸T 抓T及服務業務T字頭
'       6.P0100總所=M0100所有資料-M0100大陸P-M0100大陸T
'*** End 2015/01/12 ***
'*** Memo by Amy 2015/01/16 增加智權人員實績與結餘分析表 ***
'       1.期初實績保留:抓畫面條件起日當月4191及4192之「貸」方總和
'       2.期初結餘保留:抓畫面條件起日當月4194之「貸」方總和 (4194-104年起用)
'       3.實績點數:抓取個人 摘要不是「結餘」且不是4191也不是4192也不是4194之總和
'       4.結餘點數:抓取個人 摘要  是「結餘」且不是4191也不是4192也不是4194之總和
'       5.期末實績保留:抓畫面條件止日當月4191及4192之「借」方總和
'       6.期末結餘保留:抓畫面條件止日當月4194之「借」方總和 (4194-104年起用)
'       7.加轉撥點數及減轉撥點數固定放 0
'       8.報出實績點數:期初實績+實績點數-期末實績+加轉撥點數-減轉撥點數
'       9.報出結餘點數:期初結餘+結餘點數-期末結餘
'      10.報出點數:報出實績點數+報出結餘點數
'      11.實績保留增減:期初實績保留-期末實績保留 (2015/12/04 Add)
'*** End 2015/01/16 ***
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit
Public adoacc020 As New ADODB.Recordset
Public adostaff As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt417 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Public adoFinal As New ADODB.Recordset 'Add by Amy 2016/03/11
Dim strSort1 As String
Dim strSort2 As String
Dim strSort3 As String
Dim strOrder1 As String
Dim strOrder2 As String
Dim strOrder3 As String
Dim strCon As String
Dim dllaccrpt417(200) As Object
Dim intLength As Integer
Dim intPage As Integer
Dim strAmount As String
Dim intCounter As Integer
Dim strName As String
Dim douAmount As Double
'Add by Amy 2015/01/12
Dim strFieldN, intWidth '欄位名稱/大小
Dim strTmp(2) As String
Dim strPrinter As String 'Add byAmy 2015/06/11
'Add by Amy 2016/03/11
Dim i As Integer, intField As Integer
Dim strSum() As String, strTotalAmt() As String
Dim strM0100(0 To 8) As String, strM0100_C(0 To 8) As String
Dim strS00Row As String 'S00
Dim strM0100_T As String, strM0100_P As String  'Add by Amy 2022/02/18 Sheet 名稱/MCT總合/MCP總合 (Sheet2用)-從ExcelSaveNew2搬過來
Dim strFieldN2, intWidth2, strWkName As String 'Add by Amy 2022/02/18 工作表2 欄位名稱/大小/strWkName從ExcelSaveNew2搬過來

'Add by Amy 2020/03/31
Private Sub CboCmp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCmp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboCmp) = MsgText(601) Then Exit Sub
    
    strCmp = CboCmp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    'Modify by Amy 2020/04/16 +組合公司
    If InStr(GetBookKeepCmp & 組合作帳公司 & ",", strCmp) = 0 Then
        MsgBox Label13 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/03/31

Private Sub CmdMemo_Click()
    Frmacc44j0_1.Show
End Sub

'Add byAmy 2015/06/11
Private Sub Combo3_Click()
    If Combo3 = "實績與結餘分析表" Then
        Text1 = "Y": Text1.Locked = True
        Text3 = "N": Text3.Enabled = False
    'Add by Amy 2018/03/27 拆兩個選項,點數明細表Excel不產生(列印與Excel不同)-婧瑄/瑞婷
    ElseIf Left(Combo3, 2) = "點數" Then
        Text1.Locked = True
        Text1.Enabled = True
        Text3.Enabled = True
        If Right(Combo3, 3) = "分析表" Then
            Text1 = "Y"
            Text3 = "N": Text3.Enabled = False
        Else
            Text1 = "N": Text1.Enabled = False
            Text3 = "Y"
        End If
    Else
        Text1.Locked = False
        Text3.Enabled = True
    End If
End Sub
'end 2015/06/11

Private Sub Command1_Click()
Dim intCounter As Integer
Dim intPage As Integer
Dim i As Integer
Dim stAxb(0) As String, stMsg As String 'Add by Amy 2021/09/10
Dim bolHasActualP As Boolean 'Add by Amy 2021/09/24 有修改實績資料
    
   strM0100_T = "": strM0100_P = "" 'Add by Amy 2022/02/18 其他國內含strM0100_P(11104月前有)於工作表2需刪除
   intCounter = 0
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   'Add by Amy 2016/03/24
   ElseIf Me.Combo3 = "實績與結餘分析表" Then
        If MaskEdBox1.Text <> MsgText(29) And MaskEdBox2.Text <> MsgText(29) _
          And Val(Left(MaskEdBox1.Text, 3)) <> Val(Left(MaskEdBox2.Text, 3)) Then
            MsgBox "不可跨年查詢", , MsgText(5)
            Exit Sub
        End If
        '因抓SalesPoint資料無法判斷公司別,所以資料可能有誤差
        'Modify by Amy 2020/03/31 原:Text4
        If Trim(CboCmp) <> MsgText(601) Then MsgBox "只選擇一家公司報表結果會與智權部輸入的點數資料有誤差!", , MsgText(5)
        'Add by Amy 2021/09/10 判斷每月結餘資料有修改需彈訊息
        'Modify by Amy 2021/09/24 判斷實績保留有修改彈訊息
        stMsg = ""
        bolHasActualP = HasActualP(2, Round(Val(FCDate(MaskEdBox1.Text)) / 100, 0)) '實績保留
        Call bolAcc0b1(8, Round(Val(FCDate(MaskEdBox1.Text)) / 100, 0), stAxb()) '結餘保留
        If bolHasActualP = True Then
            stMsg = "每月點數開放後「實績」有修改，報表資料可能有誤" & vbCrLf & _
                        "請至「智權期末實績保留傳票產生」更新其傳票" & vbCrLf
        End If
        If stAxb(0) = "Y" Then
            stMsg = stMsg & (IIf(stMsg <> "", vbCrLf, "")) & _
                        "每月點數開放後「結餘」有修改，報表資料可能有誤" & vbCrLf & _
                        "請確認是否刪除「智權期末結餘保留資料」" & vbCrLf & _
                        "若需刪除請至「智權期末結餘保留資料刪除」作業刪除" & vbCrLf
        End If
        If stMsg <> MsgText(601) Then
            stMsg = stMsg & vbCrLf & "若要繼續產生報表請按「是」！"
            If MsgBox(stMsg, vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
        End If
        'end 2021/09/23
        'end 2021/09/10
   End If
   
    
   Screen.MousePointer = vbHourglass
   Accrpt417Delete
      
   'Add by Morgan 2005/8/17
   'Modify by Amy 2016/05/06 -瑞婷/婉莘用
   'Modify by Amy 2018/03/27 點數明細拆兩個選項
   If Left(Me.Combo3, 2) = "點數" Then
      If ProduceData1 = True Then
            If Right(Combo3, 3) = "分析表" Then
                'ExcelSaveNew 'Mark by Amy 2018/03/27 沒在使用-瑞婷
                If ExcelSaveNew1 = False Then Screen.MousePointer = vbDefault: Exit Sub
            Else
                '點數-明細表
                PrintDetail
            End If
      End If
   'end 2018/03/27
   'Add by Amy 2015/01/16
   ElseIf Me.Combo3 = "實績與結餘分析表" Then
        'Modify by Amy 2016/03/11 改抓function
        ExcelSaveNew2
   'end 2015/01/16
   Else
   '2005/8/17 end
   
      ProduceData
      If adocheck.State = adStateOpen Then
         adocheck.Close
      End If
      adocheck.CursorLocation = adUseClient
      adocheck.Open "select * from accrpt417 Where R41701='" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adocheck.RecordCount <> 0 Then
         ExcelSaveNew
         If Text3 = MsgText(602) Then
            PrintDetail
         End If
      End If
      adocheck.Close
      
   End If
   
   Screen.MousePointer = vbDefault
   Text2 = ""
   MsgBox MsgText(207), , MsgText(5)
   'Modify by Amy 2015/06/11 改顯示小表原:102
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5640
   Me.Height = 4050 'Modify by Amy 2015/06/11 原:3700
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Add by Amy 2020/04/16 公司別下拉
   CboCmp.Clear
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/04/16
   
   'MODIFY BY SONIA 2014/5/1 再改預設上月1日至最後一日
   'MaskEdBox1.Text = Mid(CFDate(ACDate(ServerDate)), 1, 7) & "01"
   'MODIFY BY SONIA 2013/8/6 預設至月底
   'MaskEdBox2.Text = Mid(CFDate(ACDate(ServerDate)), 1, 7) & "31"
   'MaskEdBox2.Text = CFDate(ACDate(GetLastDay(DBDATE(FCDate(MaskEdBox1)))))
   MaskEdBox1.Text = TransDate(CompDate(1, -1, (Left(strSrvDate(1), 6) & "01")), 1) '預設上月1日
   MaskEdBox1.Text = Mid(MaskEdBox1.Text, 1, 3) & "/" & Mid(MaskEdBox1.Text, 4, 2) & "/" & Mid(MaskEdBox1.Text, 6, 2)
   MaskEdBox2.Text = TransDate(CompDate(2, -1, (Left(strSrvDate(1), 6) & "01")), 1) '預設上月最後一日
   MaskEdBox2.Text = Mid(MaskEdBox2.Text, 1, 3) & "/" & Mid(MaskEdBox2.Text, 4, 2) & "/" & Mid(MaskEdBox2.Text, 6, 2)
  '2014/5/1 END
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   'Modify by Amy 2016/03/11 往前調-瑞婷
   Combo3.AddItem "實績與結餘分析表"  'Add by Amy 2015/01/16
   'Modify by Amy 2018/03/27 改程式後拿掉國內不適用文字,拆成兩個選項
   'Combo3.AddItem ComboItem(242) 'Modify by Amy 2016/05/06 國內不適用-瑞婷/婉莘用
   Combo3.AddItem "點數-分析表"
   Combo3.AddItem "點數-明細表"
   'end 2018/03/27
   Combo3.AddItem ComboItem(241) & "(暫不用-數字有問題)" 'Modify by Amy 2015/01/29暫不用-瑞婷
   Combo3 = "實績與結餘分析表"
   'end 2016/03/11
   Combo4.AddItem MsgText(1)
   Combo4.AddItem MsgText(2)
   Combo5.AddItem MsgText(1)
   Combo5.AddItem MsgText(2)
   Combo6.AddItem MsgText(1)
   Combo6.AddItem MsgText(2)
   Combo4 = MsgText(1)
   Combo5 = MsgText(1)
   Combo6 = MsgText(1)
   ComboAdd
   'Mark by Amy 2016/03/11
'   'Add By Cheng 2003/02/12
'   '預設為點數明細
'   If Me.Combo3.ListCount > 0 Then Me.Combo3.ListIndex = 1
   
   '是否產生Excel
   'Modify by Morgan 2005/2/25 預設N--瑞婷
   'Text1 = MsgText(602)
   'Modify by Morgan 2005/2/25 預設Y--瑞婷
   'Text1 = MsgText(603)
   Text1 = MsgText(602)
    'Modify By Cheng 2003/02/12
    '是否列印報表預設N
'   Text3 = MsgText(602)
   'Modify by Morgan 2005/2/25 預設Y--瑞婷
   'Text3 = MsgText(603)
   'Modify by Morgan 2005/6/2 預設N--瑞婷
   'Text3 = MsgText(602)
   Text3 = MsgText(603)
   'Modify by Amy 2015/06/11 改顯示小表原:102
   PUB_SetPrinter Me.Name, Combo7, strPrinter
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
'   Set dllaccrpt417 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   If Me.Combo7.Text <> Me.Combo7.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo7.Name, "0", "0", Me.Combo7.Text
   End If
   Set dllaccrpt417(200) = Nothing
   Set Frmacc44j0 = Nothing
End Sub

'ADD BY SONIA 2013/8/6 預設止日為起日的該月月底
Private Sub MaskEdBox1_LostFocus()
   MaskEdBox2.Text = CFDate(ACDate(GetLastDay(DBDATE(FCDate(MaskEdBox1)))))
End Sub
'END 2013/8/6

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
        Exit Sub
    End If
    
     If IsDate(ChangeTStringToWDateString(FCDate(MaskEdBox1.Text))) = False Then
        MsgBox Label2(1) & "輸入錯誤！", , MsgText(5)
        Cancel = True
        MaskEdBox1.SetFocus
        Exit Sub
    End If
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
        Exit Sub
    End If
    
     If IsDate(ChangeTStringToWDateString(FCDate(MaskEdBox2.Text))) = False Then
        MsgBox Label2(1) & "輸入錯誤！", , MsgText(5)
        Cancel = True
        MaskEdBox2.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

'*************************************************
'  Combo 項目新增
'
'*************************************************
Private Sub ComboAdd()
   strSort1 = "傳票號碼"
   strSort2 = "對沖代號(客)"
   strSort3 = "對沖代號(本所案號)"
   Combo3.AddItem strSort1
   Combo3.AddItem strSort2
   Combo3.AddItem strSort3
   Combo2.AddItem strSort1
   Combo2.AddItem strSort2
   Combo2.AddItem strSort3
   Combo1.AddItem strSort1
   Combo1.AddItem strSort2
   Combo1.AddItem strSort3
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim strCmp As String 'Add by Amy 2020/03/31

On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Select Case Combo3
      Case strSort1
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by ax202 asc"
         Else
            strOrder1 = " order by ax202 desc"
         End If
      Case strSort2
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by ax208 asc"
         Else
            strOrder1 = " order by ax208 desc"
         End If
      Case strSort3
         If Combo4 = MsgText(1) Then
            strOrder1 = " order by ax214 asc"
         Else
            strOrder1 = " order by ax214 desc"
         End If
      Case Else
         strOrder1 = MsgText(601)
   End Select
   Select Case Combo2
      Case strSort1
         If Combo5 = MsgText(1) Then
            strOrder2 = ", ax202 asc"
         Else
            strOrder2 = ", ax202 desc"
         End If
      Case strSort2
         If Combo5 = MsgText(1) Then
            strOrder2 = ", ax208 asc"
         Else
            strOrder2 = ", ax208 desc"
         End If
      Case strSort3
         If Combo5 = MsgText(1) Then
            strOrder2 = ", ax214 asc"
         Else
            strOrder2 = ", ax214 desc"
         End If
      Case Else
         strOrder2 = MsgText(601)
   End Select
   Select Case Combo1
      Case strSort1
         If Combo6 = MsgText(1) Then
            strOrder3 = ", ax202 asc"
         Else
            strOrder3 = ", ax202 desc"
         End If
      Case strSort2
         If Combo6 = MsgText(1) Then
            strOrder3 = ", ax208 asc"
         Else
            strOrder3 = ", ax208 desc"
         End If
      Case strSort3
         If Combo6 = MsgText(1) Then
            strOrder3 = ", ax214 asc"
         Else
            strOrder3 = ", ax214 desc"
         End If
      Case Else
         strOrder3 = MsgText(601)
   End Select
   strCon = MsgText(601)
   If adoaccrpt417.State <> adStateClosed Then adoaccrpt417.Close
   adoaccrpt417.CursorLocation = adUseClient
   adoaccrpt417.Open "select * from accrpt417", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc020.CursorLocation = adUseClient
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strCon = " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strCon = strCon & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text2 <> MsgText(601) Then
      strCon = strCon & " and ax209 = '" & Text2 & "'"
   End If
   'Add By Sindy 2014/1/22
   'Modify by Amy 2020/03/31 改下拉 原:Text4
   If Trim(CboCmp) <> MsgText(601) Then
      strCmp = CboCmp
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
      'strCon = strCon & " and ax201 = '" & IIf(Text4 = "2", "J", "1") & "'"
      'Modify by Amy 2020/04/16 +if 組合公司
      If InStr(strCmp, "+") > 0 Then
        strCon = strCon & " and ax201 In ( '" & Replace(strCmp, "+", "','") & "')"
      Else
        strCon = strCon & " and ax201 = '" & strCmp & "'"
      End If
   End If
   '2014/1/22 END
   
    'Modify By Cheng 2003/06/03
    'Modify by Amy 2015/04/24 +ax213
    'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
    'Modify by Amy 2020/03/31 +7129
    Select Case Me.Combo3
      Case "點數-分析表", "點數-明細表" 'Modfiy by Amy 2018/03/27  拆兩個選項 Modify by Amy 2016/05/06 國內不適用-瑞婷/婉莘用
          adoacc020.Open "select ax209, ax202, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Caseno, ax212, ax206, ax207, ax208, decode(ax207, 0, ax206 * -1, ax207) as Amount,ax213 from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and (ax209 is not null and substr(ax209, 1, 1) <> 'F') and (substr(ax205, 1, 1) = '4' Or ((ax205='7121' Or ax205='7129') And ax209 Is Not Null)) " & strCon & strOrder1 & strOrder2 & strOrder3, adoTaie, adOpenStatic, adLockReadOnly
      Case "結餘明細(暫不用-數字有問題)" 'Modify by Amy 2015/01/29 暫不用-瑞婷
          adoacc020.Open "select ax209, ax202, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Caseno, ax212, ax206, ax207, ax208, decode(ax207, 0, ax206 * -1, ax207) as Amount,ax213 from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and (ax209 is not null and substr(ax209, 1, 1) <> 'F') and substr(ax205, 1, 4) = '2491'" & strCon & strOrder1 & strOrder2 & strOrder3, adoTaie, adOpenStatic, adLockReadOnly
      Case Else
          adoacc020.Open "select ax209, ax202, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Caseno, ax212, ax206, ax207, ax208, decode(ax207, 0, ax206 * -1, ax207) as Amount,ax213 from acc021, acc020 where ax201 = a0201 and ax202 = a0202 and (ax209 is not null and substr(ax209, 1, 1) <> 'F') and (substr(ax205, 1, 1) = '4' Or ((ax205='7121' Or ax205='7129')And ax209 Is Not Null)) " & strCon & strOrder1 & strOrder2 & strOrder3, adoTaie, adOpenStatic, adLockReadOnly
    End Select
    'end 2020/03/31
    'end 2019/08/01
'   adoacc020.Open "select distinct ax201, ax202, a0k04, decode(ax207, 0, ax206 * -1, ax207) as Amount, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Caseno, ax212 from acc021, acc020, (select distinct a1p22, a0k04 from acc0m0, acc1p0, acc0k0, acc021, acc020 where a0m01 = a1p04 and a0m02 = a0k01 and a1p22 = ax202 and a1p05 = ax205 and ax201 = a0201 and ax202 = a0202 and substr(ax205, 1, 1) = '4'" & strCon & ") new where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and ax202 = a1p22 (+) and substr(ax205, 1, 1) = '4'" & strCon & _
                  " union select distinct ax201, ax202, a0k04, decode(ax207, 0, ax206 * -1, ax207) as Amount, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Caseno, ax212 from acc021, acc020, (select distinct a1p22, a0k04 from acc0s0, acc1p0, acc0k0, acc021, acc020 where a0s01 = a1p04 and a0s02 = a0k01 and a1p22 = ax202 and a1p05 = ax205 and ax201 = a0201 and ax202 = a0202 and substr(ax205, 1, 1) = '4'" & strCon & ") new where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and ax202 = a1p22 and substr(ax205, 1, 1) = '4'" & strCon & " order by ax202 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc020.RecordCount = 0 Then
      adoacc020.Close
      adoaccrpt417.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc020.EOF = False
      adoaccrpt417.AddNew
      adoaccrpt417.Fields("r41701").Value = strUserNum
      If IsNull(adoacc020.Fields("ax209").Value) Then
         adoaccrpt417.Fields("r41702").Value = Null
         adoaccrpt417.Fields("r41704").Value = Null
      Else
         adoaccrpt417.Fields("r41702").Value = adoacc020.Fields("ax209").Value & StaffQuery(adoacc020.Fields("ax209").Value)
'         adoquery.CursorLocation = adUseClient
'         adoquery.Open "select a0k04 from acc0m0, acc1p0, acc0k0 where a0m01 = a1p04 and a0m02 = a0k01 and a0k20 = '" & adoacc020.Fields("ax209").Value & "' and a1p22 = '" & adoacc020.Fields("ax202").Value & "' " & _
'                       "union select a0k04 from acc0s0, acc1p0, acc0k0 where a0s01 = a1p04 and a0s02 = a0k01 and a0k20 = '" & adoacc020.Fields("ax209").Value & "' and a1p22 = '" & adoacc020.Fields("ax202").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'         If adoquery.RecordCount <> 0 Then
         If IsNull(adoacc020.Fields("ax208").Value) Then
            adoaccrpt417.Fields("r41704").Value = Null
         Else
            adoaccrpt417.Fields("r41704").Value = adoacc020.Fields("ax208").Value & CustomerQuery(adoacc020.Fields("ax208").Value, 1)
         End If
'         Else
'            adoaccrpt417.Fields("r41704").Value = Null
'         End If
'         adoquery.Close
      End If
      adoaccrpt417.Fields("r41703").Value = adoacc020.Fields("ax202").Value
      If IsNull(adoacc020.Fields("Caseno").Value) Then
         adoaccrpt417.Fields("r41705").Value = Null
      Else
         adoaccrpt417.Fields("r41705").Value = adoacc020.Fields("Caseno").Value
      End If
      If IsNull(adoacc020.Fields("ax212").Value) Then
         adoaccrpt417.Fields("r41706").Value = Null
      Else
         adoaccrpt417.Fields("r41706").Value = adoacc020.Fields("ax212").Value
      End If
      If adoacc020.Fields("Amount").Value = 0 Then
         adoaccrpt417.Fields("r41707").Value = Val(adoacc020.Fields("Amount").Value) * (-1)
      Else
         adoaccrpt417.Fields("r41707").Value = Val(adoacc020.Fields("Amount").Value)
      End If
      'Add by Amy 2015/04/24 +ax203
      If IsNull(adoacc020.Fields("ax213").Value) Then
         adoaccrpt417.Fields("r41714").Value = Null
      Else
         adoaccrpt417.Fields("r41714").Value = adoacc020.Fields("ax213").Value
      End If
      adoaccrpt417.UpdateBatch
      adoacc020.MoveNext
   Loop
   adoacc020.Close
   adoaccrpt417.Close
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
'Add by Morgan 2005/8/17 點數明細表
'*************************************************
'  產生報表資料
'
'*************************************************
Private Function ProduceData1() As Boolean
   Dim lngEffect As Long
   'Add by Amy 2018/03/27
   Dim RsQ As New ADODB.Recordset
   Dim intQ As Integer
   Dim strCmp As String 'Add by Amy 2020/03/31
   
On Error GoTo Checking

   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   
   strCon = ""
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strCon = strCon & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strCon = strCon & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text2 <> MsgText(601) Then
      strCon = strCon & " and ax209 = '" & Text2 & "'"
   End If
   'Add By Sindy 2014/1/22
   'Modify by Amy 2020/03/31 原:Text4
   If Trim(CboCmp) <> MsgText(601) Then
      strCmp = CboCmp
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
      'strCon = strCon & " and ax201 = '" & IIf(Text4 = "2", "J", "1") & "'"
      'Modify by Amy 2020/04/16 +if 組合公司
      If InStr(strCmp, "+") > 0 Then
        strCon = strCon & " and ax201 In ('" & Replace(strCmp, "+", "','") & "')"
      Else
        strCon = strCon & " and ax201 = '" & strCmp & "'"
      End If
   End If
   'end 2020/03/31
   '2014/1/22 END
   
   '2011/9/2 modify by sonia 因加員工編號S29,寫入暫存檔時員工編號補滿5碼
   'strSql = "insert into accrpt417(r41701,r41702,r41703,r41704,r41705,r41706,r41707,r41708)" & _
      " select '" & strUserNum & "' as c01, ax209||st02 as c02,ax202 as c03,ax208||cu04 as c04" & _
      ",substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as c05" & _
      ",ax212 as c06, decode(ax207, 0, ax206 * -1, ax207) as c07,ax205 as c08" & _
      " from acc021, acc020, staff, customer where ax201 = a0201 and ax202 = a0202 and ax209 is not null" & _
      " and (substr(ax205, 1, 2) = '41' Or ax205='7121') " & strCon & _
      " and st01(+)=ax209 and cu01(+)=substr(ax208,1,8) and cu02(+)=substr(ax208,9,1)"
   'Modify by Amy 2015/01/16 +ax206/ax207/a0205/st15/ax201
   'Modify by Amy 2015/04/24 +ax213
   ''Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   'Moidfy by Amy 2020/03/31 +7129
   strSql = "insert into accrpt417(r41701,r41702,r41703,r41704,r41705,r41706,r41707,r41708,r41709,r41710,r41711,r41712,r41713,r41714)" & _
      " select '" & strUserNum & "' as c01, substr(ax209||'  ',1,5)||st02 as c02,ax202 as c03,ax208||cu04 as c04" & _
      ",substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as c05" & _
      ",ax212 as c06, decode(ax207, 0, ax206 * -1, ax207) as c07,ax205 as c08,ax206,ax207,a0205,st15,ax201,ax213" & _
      " from acc021, acc020, staff, customer where ax201 = a0201 and ax202 = a0202 and ax209 is not null and instr(ax212,'轉撥')=0 " & _
      " and (substr(ax205, 1, 1) = '4' Or ax205='7121' or ax205='7129') " & strCon & _
      " and st01(+)=ax209 and cu01(+)=substr(ax208,1,8) and cu02(+)=substr(ax208,9,1)"
   adoTaie.Execute strSql, lngEffect
   'Add by Amy 2018/03/27 更新畫面條件止日當時部門-婧瑄
   If MaskEdBox2.Text <> MsgText(29) And Val(Left(FCDate(MaskEdBox2.Text), 5)) <> Val(Left(strSrvDate(2), 5)) Then
        strSql = "Select R41702,SP48 From " & _
                    "(Select distinct R41702,R41712,Rtrim(SubStr(r41702,1,5)) as ST01 From Accrpt417 Where  r41701='" & strUserNum & "' And SubStr(r41712,1,1)='S' )," & _
                    "(Select SP02,SP48 From SalesPoint Where SP01=" & Val(Left(FCDate(MaskEdBox2.Text), 5)) + 191100 & " And SubStr(SP48,1,1)='S') " & _
                    "Where ST01=SP02 And r41712<>SP48"
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, strSql)
        If intQ = 1 Then
            Do While RsQ.EOF = False
                strSql = "Update Accrpt417 Set r41712='" & RsQ.Fields("SP48") & "' Where r41701='" & strUserNum & "' And R41702='" & RsQ.Fields("r41702") & "' "
                adoTaie.Execute strSql
                RsQ.MoveNext
            Loop
        End If
   End If
   'end 2018/03/27
   
   If lngEffect = 0 Then
      MsgBox MsgText(28), , MsgText(5)
   Else
      ProduceData1 = True
   End If
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If

End Function
'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt417Delete()
   adoTaie.Execute "delete from accrpt417 Where R41701='" & strUserNum & "'"
End Sub

''*************************************************
''  轉成Excel檔案
''
''*************************************************
'Private Sub ExcelSave()
'Dim xlsSalesPoint As New Excel.Application
'Dim wksaccrpt417 As New Worksheet
'Dim lngCounter As Long, strTotalAmt As String
'
'   If Text1 <> MsgText(602) Then
'      Exit Sub
'   End If
'    'Modify By Cheng 2003/06/09
''   If Dir(strExcelPath & ACDate(ServerDate) & MsgText(43)) = MsgText(601) Then
''      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1)) = MsgText(601) Then
''         MkDir strExcelPath
''      End If
''   Else
''      Kill strExcelPath & ACDate(ServerDate) & MsgText(43)
''   End If
'   If Dir(strExcelPath & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
'      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1)) = MsgText(601) Then
'         MkDir strExcelPath
'      End If
'   Else
'      Kill strExcelPath & ACDate(ServerDate) & ServerTime & MsgText(43)
'   End If
'   xlsSalesPoint.Workbooks.Add
'   Set wksaccrpt417 = xlsSalesPoint.Worksheets(1)
'   wksaccrpt417.Columns("a:a").ColumnWidth = 12.38
'   wksaccrpt417.Columns("b:b").ColumnWidth = 11.38
'   wksaccrpt417.Columns("c:c").ColumnWidth = 11.38
'   wksaccrpt417.Columns("d:d").ColumnWidth = 13.75
'   wksaccrpt417.Columns("e:e").ColumnWidth = 30
'   wksaccrpt417.Columns("f:f").ColumnWidth = 10.5
'   wksaccrpt417.Range("a1").Value = ReportTitle(417)
'   wksaccrpt417.Range("a1:f1").Select
'    'Modify By Cheng 2003/02/12
''   With Selection
'   With wksaccrpt417.Range("a1:f1")
'       .HorizontalAlignment = xlCenter
'       .VerticalAlignment = xlBottom
'       .WrapText = False
'       .Orientation = 0
'       .AddIndent = False
'       .ShrinkToFit = False
'       .MergeCells = True
'   End With
'   wksaccrpt417.Range("a4").Value = ReportSum(27)
'   wksaccrpt417.Range("b4").Value = MaskEdBox1.Text
'   wksaccrpt417.Range("c4").Value = ReportSum(28)
'   wksaccrpt417.Range("d4").Value = MaskEdBox2.Text
'   wksaccrpt417.Range("a6").Value = ReportSum(29)
'   wksaccrpt417.Range("b6").Value = ReportSum(30)
'   wksaccrpt417.Range("c6").Value = ReportSum(31)
'   wksaccrpt417.Range("d6").Value = ReportSum(32)
'   wksaccrpt417.Range("e6").Value = ReportSum(33)
'   wksaccrpt417.Range("f6").Value = ReportSum(34)
'   lngCounter = 7
'   adostaff.CursorLocation = adUseClient
'   If Text2 = MsgText(601) Then
'      adostaff.Open "select * from staff where substr(st03, 1, 1) = 'S' order by st01 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adostaff.Open "select * from staff where st01 = '" & Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   Do While adostaff.EOF = False
'      adoaccrpt417.CursorLocation = adUseClient
'      adoaccrpt417.Open "select * from accrpt417 where r41701 = '" & strUserNum & "' and r41702 = '" & adostaff.Fields("st02").Value & "' order by r41703 asc", adoTaie, adOpenStatic, adLockReadOnly
'      Do While adoaccrpt417.EOF = False
'         If IsNull(adoaccrpt417.Fields("r41702").Value) Then
'            wksaccrpt417.Range("a" & lngCounter).Value = MsgText(601)
'         Else
'            wksaccrpt417.Range("a" & lngCounter).Value = adoaccrpt417.Fields("r41702").Value
'         End If
'         If IsNull(adoaccrpt417.Fields("r41703").Value) Then
'            wksaccrpt417.Range("b" & lngCounter).Value = MsgText(601)
'         Else
'            wksaccrpt417.Range("b" & lngCounter).Value = adoaccrpt417.Fields("r41703").Value
'         End If
'         If IsNull(adoaccrpt417.Fields("r41704").Value) Then
'            wksaccrpt417.Range("c" & lngCounter).Value = MsgText(601)
'         Else
'            wksaccrpt417.Range("c" & lngCounter).Value = adoaccrpt417.Fields("r41704").Value
'         End If
'         If IsNull(adoaccrpt417.Fields("r41705").Value) Then
'            wksaccrpt417.Range("d" & lngCounter).Value = MsgText(601)
'         Else
'            wksaccrpt417.Range("d" & lngCounter).Value = adoaccrpt417.Fields("r41705").Value
'         End If
'         If IsNull(adoaccrpt417.Fields("r41706").Value) Then
'            wksaccrpt417.Range("e" & lngCounter).Value = MsgText(601)
'         Else
'            wksaccrpt417.Range("e" & lngCounter).Value = adoaccrpt417.Fields("r41706").Value
'         End If
'         If IsNull(adoaccrpt417.Fields("r41707").Value) Then
'            wksaccrpt417.Range("f" & lngCounter).Value = 0
'         Else
'            wksaccrpt417.Range("f" & lngCounter).Value = Val(adoaccrpt417.Fields("r41707").Value)
'         End If
'         lngCounter = lngCounter + 1
'         adoaccrpt417.MoveNext
'      Loop
'      adoaccsum.CursorLocation = adUseClient
'      adoaccsum.Open "select count(*) from accrpt417 where r41701 = '" & strUserNum & "' and r41702 = '" & adostaff.Fields("st02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccsum.RecordCount <> 0 Then
'         If IsNull(adoaccsum.Fields(0).Value) = False And adoaccsum.Fields(0).Value <> 0 Then
'            wksaccrpt417.Range("e" & lngCounter).Value = ReportSum(24)
'            wksaccrpt417.Range("f" & lngCounter).Formula = "=sum(f" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":f" & (lngCounter - 1) & ")"
'            strTotalAmt = strTotalAmt & "f" & lngCounter & ", "
'            lngCounter = lngCounter + 2
'         End If
'      End If
'      adoaccsum.Close
'      adoaccrpt417.Close
'      adostaff.MoveNext
'   Loop
'   wksaccrpt417.Range("e" & lngCounter).Value = ReportSum(25)
'   If strTotalAmt <> "" Then
'      wksaccrpt417.Range("f" & lngCounter).Formula = "=sum(" & Mid(strTotalAmt, 1, Len(strTotalAmt) - 2) & ")"
'   End If
'   adostaff.Close
'    'Modify By Cheng 2003/06/09
''   xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & ACDate(ServerDate) & MsgText(43)
'   xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & ACDate(ServerDate) & ServerTime & MsgText(43)
'   xlsSalesPoint.Workbooks.Close
'   xlsSalesPoint.Quit
'   StatusClear
'End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = Mid(CFDate(ACDate(ServerDate)), 1, 7) & "01"
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = Mid(CFDate(ACDate(ServerDate)), 1, 7) & "31"
   MaskEdBox2.Mask = DFormat
   Text2 = ""
   'Text1 = ""
   'Text3 = ""
   Combo3 = ""
   Combo2 = ""
   Combo1 = ""
   Combo3.SetFocus
End Sub

'*************************************************
'  轉成Excel檔案
'  智權人員點數明細表(與結餘明細表共用暫存檔但資料抓法不同)
'*************************************************
'Mark by Amy 2018/03/27 沒在使用-瑞婷
Private Sub ExcelSaveNew()
'Dim xlsSalesPoint As New Excel.Application
'Dim wksaccrpt417 As New Worksheet
'Dim xlsSelect As Selection
'Dim lngCounter As Long, lngCounter1 As Long
'Dim strTotalAmt1 As String, strTotalAmt2 As String, strTotalAmt3 As String, strTotalAmt4 As String, strTotalAmt5 As String
'Dim strTotalAmt6 As String, strTotalAmt7 As String, strTotalAmt8 As String
''Add by Amy 2014/01/12
'Dim strTotalAmt9 As String '起日當月保留點數合計公式
'Dim strTotalAmt10 As String '保留點數相減公式
'Dim strDept As String
''Add by Amy 2015/01/12
'Dim intTitleRow As Integer
'Dim ii As Integer
'
'   ReDim strSum(6) 'Modify by Amy 2015/01/12 原:5
'   strDept = ""
'   If Text1 <> MsgText(602) Then
'      Exit Sub
'   End If
'   'Modify by Amy 2015/01/29 加-結餘
'   If Dir(strExcelPath & Mid(ReportTitle(417), 6, 9) & IIf(Left(Combo3, 4) = ComboItem(241), "-結餘", "") & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
'      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'         MkDir strExcelPath
'      End If
'   Else
'      Kill strExcelPath & Mid(ReportTitle(417), 6, 9) & IIf(Left(Combo3, 4) = ComboItem(241), "-結餘", "") & ACDate(ServerDate) & ServerTime & MsgText(43)
'   End If
'   'end 2015/01/29
'   xlsSalesPoint.SheetsInNewWorkbook = 3 'Add by Amy  2019/04/02 改設定(選項->一般->包括的工作表份數)
'   xlsSalesPoint.Workbooks.add
'   Set wksaccrpt417 = xlsSalesPoint.Worksheets(1)
'   'Add By Sindy 2014/1/22
'   wksaccrpt417.Range("a3").Value = "公司別:"
'   'Modify by Amy 2020/04/16 改下拉 原:Text4
'   'wksaccrpt417.Range("b3").Value = IIf(Text4 = "1", "台一", IIf(Text4 = "2", "智權", "台一　專利商標/智權"))
'   wksaccrpt417.Range("b3").Value = GetAccReportCmpN(CboCmp, , True)
'   'end 2020/04/16
'   '2014/1/22 END
'   wksaccrpt417.Range("a4").Value = ReportSum(27)
'   wksaccrpt417.Range("b4").Value = MaskEdBox1.Text
'   wksaccrpt417.Range("c4").Value = ReportSum(28)
'   wksaccrpt417.Range("d4").Value = MaskEdBox2.Text
'
'   'Modify by Amy 2015/01/12 改為動態產生欄位,增加起日當月保留點數欄位
'   '欄位名稱
'   ReDim strFieldN(1 To 8)
'   ReDim intWidth(1 To 8)
'   ii = 1: intField = 65: intTitleRow = 6
'   '智權人員
'   wksaccrpt417.Range(Chr(intField) & "6").Value = ReportSum(57): strFieldN(ii) = ReportSum(57): intWidth(ii) = 13: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & "6").HorizontalAlignment = xlCenter: intField = intField + 1
'   '業務達成點數
'   wksaccrpt417.Range(Chr(intField) & "6").Value = ReportSum(58): strFieldN(ii) = ReportSum(58): intWidth(ii) = 13: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & "6").HorizontalAlignment = xlCenter: intField = intField + 1
'   '加轉撥點數
'   wksaccrpt417.Range(Chr(intField) & "6").Value = ReportSum(59): strFieldN(ii) = ReportSum(59): intWidth(ii) = 13: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & "6").HorizontalAlignment = xlCenter: intField = intField + 1
'   '減轉撥點數
'   wksaccrpt417.Range(Chr(intField) & "6").Value = ReportSum(60): strFieldN(ii) = ReportSum(60): intWidth(ii) = 13: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & "6").HorizontalAlignment = xlCenter: intField = intField + 1
'
'  '止日當月保留點數
'  'Modify by Amy 2015/01/16 止日當月保留點數 改為 保留點數
'   strTmp(0) = "保留點數" '原:FCDate(MaskEdBox2.Text)
''   If strTmp(0) <> MsgText(601) Then
''        strTmp(0) = Val(Mid(strTmp(0), 1, 3)) & "/" & Val(Mid(strTmp(0), 4, 2)) & ReportSum(61)
''   End If
'   'end 2015/01/16
'   wksaccrpt417.Range(Chr(intField) & "6").Value = strTmp(0): strFieldN(ii) = strTmp(0): intWidth(ii) = 13.5: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & "6").HorizontalAlignment = xlCenter: intField = intField + 1
'   '實際達成點數
'   wksaccrpt417.Range(Chr(intField) & "6").Value = ReportSum(62): strFieldN(ii) = ReportSum(62): intWidth(ii) = 13: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & "6").HorizontalAlignment = xlCenter: intField = intField + 1
'
'   'Modify by Amy 2015/01/16 隱藏 起日當月保留點數及保留點數相減 欄位
'   '起日當月保留點數
'   strTmp(1) = FCDate(MaskEdBox1.Text)
'   If strTmp(1) <> MsgText(601) Then
'        If Mid(strTmp(1), 4, 2) = "01" Then
'            strTmp(1) = Val(Mid(strTmp(1), 1, 3)) - 1 & "/12"
'        Else
'            strTmp(1) = Val(Mid(strTmp(1), 1, 3)) & "/" & Val(Mid(strTmp(1), 4, 2)) - 1
'        End If
'        strTmp(1) = strTmp(1) & ReportSum(61)
'   End If
'   wksaccrpt417.Range(Chr(intField) & "6").Value = strTmp(1): strFieldN(ii) = strTmp(1): intWidth(ii) = 0: ii = ii + 1 '原欄寬:13.5
'   wksaccrpt417.Range(Chr(intField) & "6").HorizontalAlignment = xlCenter: intField = intField + 1
'   '保留點數相減欄位
'   strTmp(2) = Replace(strTmp(0), ReportSum(61), "") & "－" & Replace(strTmp(1), ReportSum(61), "") & "(E-G)"
'   wksaccrpt417.Range(Chr(intField) & "6").Value = strTmp(2): strFieldN(ii) = strTmp(2)
'   intWidth(ii) = 0: ii = ii + 1 '原欄寬:18
'   wksaccrpt417.Range(Chr(intField) & "6").HorizontalAlignment = xlCenter: intField = intField + 1
'   lngCounter = 7
'
'   '2015/1/29 MODIFY BY SONIA
'   'wksaccrpt417.PageSetup.PrintTitleRows = "$1:$" & UBound(strFieldN)
'   'Modify by Amy 2016/03/11
'   'wksaccrpt417.PageSetup.PrintTitleRows = "$1:$6"
'   wksaccrpt417.PageSetup.PrintTitleRows = "$1:$" & intTitleRow
'   For ii = 1 To UBound(strFieldN)
'        wksaccrpt417.Columns(Chr(ii + 64) & ":" & Chr(ii + 64)).ColumnWidth = intWidth(ii)
'   Next ii
'
'   'Add by Amy 2015/01/29 +if
'   If Combo3 = "結餘明細(暫不用-數字有問題)" Then
'        wksaccrpt417.Range("a1").Value = "***  智權人員點數明細表(結餘)  ***"
'   Else
'        wksaccrpt417.Range("a1").Value = ReportTitle(417)
'   End If
'   wksaccrpt417.Range("a1:" & Chr(UBound(strFieldN) + 64) & "1").Select
'    With wksaccrpt417.Range("a1:" & Chr(UBound(strFieldN) + 64) & "1")
'       .HorizontalAlignment = xlCenter
'       .VerticalAlignment = xlBottom
'       .WrapText = False
'       .Orientation = 0
'       .AddIndent = False
'       .ShrinkToFit = False
'       .MergeCells = True
'   End With
'
'   ' 智權人員
'   If adostaff.State <> adStateClosed Then adostaff.Close
'   adostaff.CursorLocation = adUseClient
'   adostaff.Open "select * from acc090 where substr(a0901, 1, 1) = 'S' order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adostaff.EOF = False
'      adoaccrpt417.CursorLocation = adUseClient
'      '2011/9/2 modify by sonia 因加員工編號S29,故取5碼去空白讀員工檔
'      'adoaccrpt417.Open "select r41702, sum(r41707) from accrpt417, staff where substr(r41702, 1, 5) = st01 and r41701 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "' group by r41702", adoTaie, adOpenStatic, adLockReadOnly
'      adoaccrpt417.Open "select r41702, sum(r41707) from accrpt417, staff where rtrim(substr(r41702, 1, 5)) = st01 and r41701 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "' group by r41702", adoTaie, adOpenStatic, adLockReadOnly
'      Do While adoaccrpt417.EOF = False
'         If strDept <> Mid(adostaff.Fields("a0901").Value, 1, 2) Then
'            If strDept <> "" Then
'               Select Case strDept
'                  Case "S1", "S2"
'                     If strSum(0) <> "" Then
'                        strSum(0) = Mid(strSum(0), 1, Len(strSum(0)) - 1)
'                     End If
'                     If strSum(1) <> "" Then
'                        strSum(1) = Mid(strSum(1), 1, Len(strSum(1)) - 1)
'                     End If
'                     If strSum(2) <> "" Then
'                        strSum(2) = Mid(strSum(2), 1, Len(strSum(2)) - 1)
'                     End If
'                     If strSum(3) <> "" Then
'                        strSum(3) = Mid(strSum(3), 1, Len(strSum(3)) - 1)
'                     End If
'                     If strSum(4) <> "" Then
'                        strSum(4) = Mid(strSum(4), 1, Len(strSum(4)) - 1)
'                     End If
'                     If strSum(5) <> "" Then
'                        strSum(5) = Mid(strSum(5), 1, Len(strSum(5)) - 1)
'                     End If
'                     If strSum(6) <> "" Then
'                        strSum(6) = Mid(strSum(6), 1, Len(strSum(6)) - 1)
'                     End If
'                     Select Case strDept
'                        Case "S1"
'                           wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(105)
'                        Case "S2"
'                           wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(106)
'                        Case "S3"
'                           wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(107)
'                        Case "S4"
'                           wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(108)
'                     End Select
'                     wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Formula = strSum(0)
'                     wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Formula = strSum(1)
'                     wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Formula = strSum(2)
'                     wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Formula = strSum(3)
'                     wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'                     wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Formula = strSum(5) '起日當月保留點數
'                     wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Formula = strSum(6) '保留點數相減欄位
'                     lngCounter = lngCounter + 2
'               End Select
'            End If
'            strSum(0) = "=": strSum(1) = "=": strSum(2) = "=": strSum(3) = "=": strSum(4) = "=": strSum(5) = "=": strSum(6) = "="
'            strDept = Mid(adostaff.Fields("a0901").Value, 1, 2)
'         End If
'
'         If IsNull(adoaccrpt417.Fields(0).Value) Then
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = 0
'         Else
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = adoaccrpt417.Fields(0).Value
'         End If
'         If IsNull(adoaccrpt417.Fields(1).Value) Then
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Value = 0
'         Else
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Value = Val(adoaccrpt417.Fields(1).Value)
'         End If
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Value = 0
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Value = 0
'         wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Value = GetEndMonthDebit(Left(RTrim(adoaccrpt417.Fields(0).Value), 5)) 'Modify by Amy 2015/01/12 原:0
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'         wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Value = GetStartMonthCredit(Left(RTrim(adoaccrpt417.Fields(0).Value), 5))
'         wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(strTmp(0)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(1)) + 64) & lngCounter
'         lngCounter = lngCounter + 1
'         adoaccrpt417.MoveNext
'      Loop
'
'      '合計
'      adoaccsum.CursorLocation = adUseClient
'      '2011/9/2 modify by sonia 因加員工編號S29,故取5碼去空白讀員工檔
'      'adoaccsum.Open "select count(distinct r41702) from accrpt417, staff where substr(r41702, 1, 5) = st01 and r41701 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      adoaccsum.Open "select count(distinct r41702) from accrpt417, staff where rtrim(substr(r41702, 1, 5)) = st01 and r41701 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccsum.RecordCount <> 0 Then
'         If IsNull(adoaccsum.Fields(0).Value) = False And adoaccsum.Fields(0).Value <> 0 Then
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = adostaff.Fields("a0902").Value & ReportSum(25)
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportSum(58)) + 64) & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":" & Chr(GetValue(ReportSum(58)) + 64) & (lngCounter - 1) & ")"
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportSum(59)) + 64) & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":" & Chr(GetValue(ReportSum(59)) + 64) & (lngCounter - 1) & ")"
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportSum(60)) + 64) & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":" & Chr(GetValue(ReportSum(60)) + 64) & (lngCounter - 1) & ")"
'            wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(strTmp(0)) + 64) & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":" & Chr(GetValue(strTmp(0)) + 64) & (lngCounter - 1) & ")"
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'            wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(strTmp(1)) + 64) & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":" & Chr(GetValue(strTmp(1)) + 64) & (lngCounter - 1) & ")"
'            wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(strTmp(2)) + 64) & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":" & Chr(GetValue(strTmp(2)) + 64) & (lngCounter - 1) & ")"
'
'            'Add by Amy 2015/01/16 發現合計錯誤
'            strSum(0) = strSum(0) & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+"
'            strSum(1) = strSum(1) & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "+"
'            strSum(2) = strSum(2) & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "+"
'            strSum(3) = strSum(3) & Chr(GetValue(strTmp(0)) + 64) & lngCounter & "+"
'            strSum(4) = strSum(4) & Chr(GetValue(ReportSum(62)) + 64) & lngCounter & "+"
'            strSum(5) = strSum(5) & Chr(GetValue(strTmp(1)) + 64) & lngCounter & "+"
'            strSum(6) = strSum(6) & Chr(GetValue(strTmp(2)) + 64) & lngCounter & "+"
'
'            strTotalAmt1 = strTotalAmt1 & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & ", "
'            strTotalAmt2 = strTotalAmt2 & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & ", "
'            strTotalAmt3 = strTotalAmt3 & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & ", "
'            strTotalAmt4 = strTotalAmt4 & Chr(GetValue(strTmp(0)) + 64) & lngCounter & ", "
'            strTotalAmt9 = strTotalAmt9 & Chr(GetValue(strTmp(1)) + 64) & lngCounter & ", "
'            strTotalAmt10 = strTotalAmt10 & Chr(GetValue(strTmp(2)) + 64) & lngCounter & ", "
'            lngCounter = lngCounter + 2
'         End If
'      End If
'      adoaccsum.Close
'      adoaccrpt417.Close
'      adostaff.MoveNext
'   Loop
'   adostaff.Close
'   Select Case strDept
'      Case "S1", "S2"
'         If strSum(0) <> "" Then
'            If strSum(0) <> "" Then
'               strSum(0) = Mid(strSum(0), 1, Len(strSum(0)) - 1)
'            End If
'            If strSum(1) <> "" Then
'               strSum(1) = Mid(strSum(1), 1, Len(strSum(1)) - 1)
'            End If
'            If strSum(2) <> "" Then
'               strSum(2) = Mid(strSum(2), 1, Len(strSum(2)) - 1)
'            End If
'            If strSum(3) <> "" Then
'               strSum(3) = Mid(strSum(3), 1, Len(strSum(3)) - 1)
'            End If
'            If strSum(4) <> "" Then
'               strSum(4) = Mid(strSum(4), 1, Len(strSum(4)) - 1)
'            End If
'            If strSum(5) <> "" Then
'               strSum(5) = Mid(strSum(5), 1, Len(strSum(5)) - 1)
'            End If
'            If strSum(6) <> "" Then
'               strSum(6) = Mid(strSum(6), 1, Len(strSum(6)) - 1)
'            End If
'            Select Case strDept
'               Case "S1"
'                  wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(105)
'               Case "S2"
'                  wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(106)
'               Case "S3"
'                  wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(107)
'               Case "S4"
'                  wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(108)
'               Case "S9"
'                  wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(127)
'            End Select
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Formula = strSum(0)
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Formula = strSum(1)
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Formula = strSum(2)
'            wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Formula = strSum(3)
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'            wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Formula = strSum(5)
'            wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Formula = strSum(6)
'            lngCounter = lngCounter + 2
'         End If
'   End Select
'
'   'Add by Morgan 2010/7/15
'   If strTotalAmt1 & strTotalAmt2 & strTotalAmt3 & strTotalAmt4 <> "" Then
'      wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = "智權部合計:"
'      If strTotalAmt1 <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt1, 1, Len(strTotalAmt1) - 2) & ")"
'      If strTotalAmt2 <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt2, 1, Len(strTotalAmt2) - 2) & ")"
'      If strTotalAmt3 <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt3, 1, Len(strTotalAmt3) - 2) & ")"
'      If strTotalAmt4 <> "" Then wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt4, 1, Len(strTotalAmt4) - 2) & ")"
'      wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'      If strTotalAmt9 <> "" Then wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt9, 1, Len(strTotalAmt9) - 2) & ")"
'      If strTotalAmt10 <> "" Then wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt10, 1, Len(strTotalAmt10) - 2) & ")"
'      lngCounter = lngCounter + 2
'   End If
'   'end 2010/7/15
'
'' 其他人員
'   lngCounter1 = lngCounter 'Add by Morgan 2010/7/15
'   adostaff.CursorLocation = adUseClient
'   'Ken 92/05/08 計算部門第一碼為F, 但不包括F4102, F4103, F4101
'   'adostaff.Open "select * from acc090 where (substr(a0901, 1, 1) <> 'S' and substr(a0901, 1, 1) <> 'F') or a0901 = '020' order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
'   adostaff.Open "select * from acc090 where (substr(a0901, 1, 1) <> 'S') or a0901 = '020' order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adostaff.EOF = False
'      adoaccrpt417.CursorLocation = adUseClient
'      'adoaccrpt417.Open "select r41702, sum(r41707) from accrpt417, staff where substr(r41702, 1, 5) = st01 and r41701 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "' group by r41702", adoTaie, adOpenStatic, adLockReadOnly
'      '2011/9/2 modify by sonia 因加員工編號S29,故取5碼去空白讀員工檔
'      'adoaccrpt417.Open "select r41702, sum(r41707) from accrpt417, staff where substr(r41702, 1, 5) = st01 and r41701 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "' and substr(r41702, 1, 5) not in ('F4102', 'F4103', 'F4101') group by r41702", adoTaie, adOpenStatic, adLockReadOnly
'      adoaccrpt417.Open "select r41702, sum(r41707) from accrpt417, staff where rtrim(substr(r41702, 1, 5)) = st01 and r41701 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "' and substr(r41702, 1, 5) not in ('F4102', 'F4103', 'F4101') group by r41702", adoTaie, adOpenStatic, adLockReadOnly
'      Do While adoaccrpt417.EOF = False
'         'Add by Morgan 2010/7/15
'         strExc(1) = 0
'         If Left("" & adoaccrpt417.Fields(0), 5) = "M0100" Then
'            'Modify by Amy 2015/02/05 +服務業務
'            'strExc(0) = "select sum(r41707) from accrpt417 where r41701 = '" & strUserNum & "' and r41702='" & adoaccrpt417.Fields(0) & "' and R41705 like 'P-%'" & _
'               " and exists(select * from patent,fagent,customer where pa01='P' and pa02=substr(R41705,3,6) and pa03=substr(R41705,10,1) and pa04=substr(R41705,12) and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and nvl(fa10,cu10)>'009')" & _
'               " having sum(r41707)>0"
'            strExc(0) = "select sum(r41707) from accrpt417 where r41701 = '" & strUserNum & "' and r41702='" & adoaccrpt417.Fields(0) & "' and ((R41705 like 'P-%'" & _
'               " and exists(select * from patent,fagent,customer where pa01='P' and pa02=substr(R41705,3,6) and pa03=substr(R41705,10,1) and pa04=substr(R41705,12) and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and nvl(fa10,cu10)>'009')) " & _
'               "Or (R41705 like 'PS-%' And Exists (Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
'                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
'                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) )" & _
'               " having sum(r41707)>0"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strExc(1) = Val(strExc(1)) + Val(RsTemp.Fields(0).Value)
'               wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = "M0100大陸P"
'               wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Value = Val(RsTemp.Fields(0).Value)
'               wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Value = 0
'               wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Value = 0
'               wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Value = GetEndMonthDebit("M0100", "P") 'Modify by Amy 2015/01/12 原:0
'               wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'               wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Value = GetStartMonthCredit("M0100", "P")
'               'end 2015/02/05
'               wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(strTmp(0)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(1)) + 64) & lngCounter
'               lngCounter = lngCounter + 1
'            End If
'
'            'Modify by Amy 2015/02/05 +服務業務
'            'strExc(0) = "select nvl(sum(r41707),0) from accrpt417 where r41701 = '" & strUserNum & "' and r41702='" & adoaccrpt417.Fields(0) & "' and R41705 like 'T-%'" & _
'               " and exists(select * from trademark,fagent,customer where TM01='T' and TM02=substr(R41705,3,6) and TM03=substr(R41705,10,1) and TM04=substr(R41705,12) and fa01(+)=substr(tm44,1,8) and fa02(+)=substr(tm44,9) and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) and nvl(fa10,cu10)>'009')" & _
'               " having sum(r41707)>0"
'            strExc(0) = "select nvl(sum(r41707),0) from accrpt417 where r41701 = '" & strUserNum & "' and r41702='" & adoaccrpt417.Fields(0) & "' and ((R41705 like 'T-%'" & _
'               " and exists(select * from trademark,fagent,customer where TM01='T' and TM02=substr(R41705,3,6) and TM03=substr(R41705,10,1) and TM04=substr(R41705,12) and fa01(+)=substr(tm44,1,8) and fa02(+)=substr(tm44,9) and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) and nvl(fa10,cu10)>'009')) " & _
'               "Or (R41705 like 'T%' And Exists (Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
'                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
'                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) ) " & _
'               " having sum(r41707)>0"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strExc(1) = Val(strExc(1)) + Val(RsTemp.Fields(0).Value)
'               wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = "M0100大陸T"
'               wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Value = Val(RsTemp.Fields(0).Value)
'               wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Value = 0
'               wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Value = 0
'               wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Value = 0
'               wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'               wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Value = 0
'               wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Value = 0
'               lngCounter = lngCounter + 1
'            End If
'         End If
'         'end 2010/7/15
'         If IsNull(adoaccrpt417.Fields(0).Value) Then
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = 0
'         Else
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = adoaccrpt417.Fields(0).Value
'         End If
'         If IsNull(adoaccrpt417.Fields(1).Value) Then
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Value = 0
'         Else
'            'Modify by Morgan 2010/7/15
'            'wksaccrpt417.Range("b" & lngCounter).Value = Val(adoaccrpt417.Fields(1).Value)
'            wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Value = Val(adoaccrpt417.Fields(1).Value) - Val(strExc(1))
'         End If
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Value = 0
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Value = 0
'         wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Value = GetEndMonthDebit(Left(RTrim(adoaccrpt417.Fields(0).Value), 5)) 'Modify by Amy 2015/01/12 原:0
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'         wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Value = GetStartMonthCredit(Left(RTrim(adoaccrpt417.Fields(0).Value), 5))
'         wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(strTmp(0)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(1)) + 64) & lngCounter
'         lngCounter = lngCounter + 1
'
'         adoaccrpt417.MoveNext
'      Loop
'      adoaccrpt417.Close
'      adostaff.MoveNext
'   Loop
'   adostaff.Close
'
'   adoaccsum.CursorLocation = adUseClient
'   '2011/9/2 modify by sonia 因加員工編號S29,故取5碼去空白讀員工檔
'   'adoaccsum.Open "select count(distinct substr(r41702, 1, 5)) from accrpt417, staff, acc090 where substr(r41702, 1, 5) = st01 and st15 = a0901 and r41701 = '" & strUserNum & "' and ((substr(a0901, 1, 1) <> 'S' and substr(a0901, 1, 1) <> 'F') or a0901 = '020')", adoTaie, adOpenStatic, adLockReadOnly
'   adoaccsum.Open "select count(distinct rtrim(substr(r41702, 1, 5))) from accrpt417, staff, acc090 where rtrim(substr(r41702, 1, 5)) = st01 and st15 = a0901 and r41701 = '" & strUserNum & "' and ((substr(a0901, 1, 1) <> 'S' and substr(a0901, 1, 1) <> 'F') or a0901 = '020')", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) = False And adoaccsum.Fields(0).Value <> 0 Then
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(64) & ReportSum(25)
'         'Modify by Morgan 2010/7/15
'         'wksaccrpt417.Range("b" & lngCounter).Formula = "=sum(b" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":b" & (lngCounter - 1) & ")"
'         'wksaccrpt417.Range("c" & lngCounter).Formula = "=sum(c" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":c" & (lngCounter - 1) & ")"
'         'wksaccrpt417.Range("d" & lngCounter).Formula = "=sum(d" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":d" & (lngCounter - 1) & ")"
'         'wksaccrpt417.Range("e" & lngCounter).Formula = "=sum(e" & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":e" & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportSum(58)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportSum(59)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportSum(60)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(strTmp(0)) + 64) & lngCounter1 & ":" & Chr(GetValue(strTmp(0)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(strTmp(1)) + 64) & lngCounter1 & ":" & Chr(GetValue(strTmp(1)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(strTmp(2)) + 64) & lngCounter1 & ":" & Chr(GetValue(strTmp(2)) + 64) & (lngCounter - 1) & ")"
'         'end 2010/7/15
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'         strTotalAmt1 = strTotalAmt1 & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & ", "
'         strTotalAmt2 = strTotalAmt2 & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & ", "
'         strTotalAmt3 = strTotalAmt3 & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & ", "
'         strTotalAmt4 = strTotalAmt4 & Chr(GetValue(strTmp(0)) + 64) & lngCounter & ", "
'         strTotalAmt9 = strTotalAmt9 & Chr(GetValue(strTmp(1)) + 64) & lngCounter & ", "
'         strTotalAmt10 = strTotalAmt10 & Chr(GetValue(strTmp(2)) + 64) & lngCounter & ", "
'         lngCounter = lngCounter + 2
'      End If
'   End If
'   adoaccsum.Close
'
'' 國內合計
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(65) & ReportSum(25)
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt1, 1, Len(strTotalAmt1) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt2, 1, Len(strTotalAmt2) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt3, 1, Len(strTotalAmt3) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt4, 1, Len(strTotalAmt4) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'   wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt9, 1, Len(strTotalAmt9) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt10, 1, Len(strTotalAmt10) - 2) & ")"
'   lngCounter = lngCounter + 2
'
'' FCP
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(67) & ReportSum(25)
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'   adoaccrpt417.CursorLocation = adUseClient
'   adoaccrpt417.Open "select sum(r41707) from accrpt417 where substr(r41702, 1, 5)='F4102' and r41701 = '" & strUserNum & "' ", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccrpt417.RecordCount <> 0 Then
'      If IsNull(adoaccrpt417.Fields(0).Value) Then
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Value = 0
'      Else
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Value = Val(adoaccrpt417.Fields(0).Value)
'      End If
'      wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Value = 0
'      wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Value = 0
'      wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Value = 0
'      wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Value = 0
'      wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Value = 0
'      strTotalAmt5 = strTotalAmt5 & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & ", "
'      strTotalAmt6 = strTotalAmt6 & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & ", "
'      strTotalAmt7 = strTotalAmt7 & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & ", "
'      strTotalAmt8 = strTotalAmt8 & Chr(GetValue(strTmp(0)) + 64) & lngCounter & ", "
'      adoaccrpt417.MoveNext
'   End If
'   adoaccrpt417.Close
'   lngCounter = lngCounter + 1
'
'' FCT
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(68) & ReportSum(25)
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'   adoaccrpt417.CursorLocation = adUseClient
'   adoaccrpt417.Open "select sum(r41707) from accrpt417 where substr(r41702, 1, 5)='F4103' and r41701 = '" & strUserNum & "' ", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccrpt417.RecordCount <> 0 Then
'      If IsNull(adoaccrpt417.Fields(0).Value) Then
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Value = 0
'      Else
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Value = Val(adoaccrpt417.Fields(0).Value)
'      End If
'      wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Value = 0
'      wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Value = 0
'      wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Value = 0
'      wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Value = 0
'      wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Value = 0
'      strTotalAmt5 = strTotalAmt5 & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & ", "
'      strTotalAmt6 = strTotalAmt6 & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & ", "
'      strTotalAmt7 = strTotalAmt7 & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & ", "
'      strTotalAmt8 = strTotalAmt8 & Chr(GetValue(strTmp(0)) + 64) & lngCounter & ", "
'      adoaccrpt417.MoveNext
'   End If
'   adoaccrpt417.Close
'   lngCounter = lngCounter + 1
'
'' FCL
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(69) & ReportSum(25)
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'   adoaccrpt417.CursorLocation = adUseClient
'   adoaccrpt417.Open "select sum(r41707) from accrpt417 where substr(r41702, 1, 5) = 'F4101' and r41701 = '" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccrpt417.RecordCount <> 0 Then
'      If IsNull(adoaccrpt417.Fields(0).Value) Then
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Value = 0
'      Else
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Value = Val(adoaccrpt417.Fields(0).Value)
'      End If
'      wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Value = 0
'      wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Value = 0
'      wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Value = 0
'      wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Value = 0
'      wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Value = 0
'      strTotalAmt5 = strTotalAmt5 & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & ", "
'      strTotalAmt6 = strTotalAmt6 & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & ", "
'      strTotalAmt7 = strTotalAmt7 & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & ", "
'      strTotalAmt8 = strTotalAmt8 & Chr(GetValue(strTmp(0)) + 64) & lngCounter & ", "
'      adoaccrpt417.MoveNext
'   End If
'   adoaccrpt417.Close
'   lngCounter = lngCounter + 1
'
'' 國外合計
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(70) & ReportSum(25)
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt5, 1, Len(strTotalAmt5) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt6, 1, Len(strTotalAmt6) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt7, 1, Len(strTotalAmt7) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt8, 1, Len(strTotalAmt8) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'   wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt9, 1, Len(strTotalAmt9) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt10, 1, Len(strTotalAmt10) - 2) & ")"
'   lngCounter = lngCounter + 2
'
'' 總所合計
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(66) & ReportSum(25)
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt1 & Mid(strTotalAmt5, 1, Len(strTotalAmt5) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(59)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt2 & Mid(strTotalAmt6, 1, Len(strTotalAmt6) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(60)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt3 & Mid(strTotalAmt7, 1, Len(strTotalAmt7) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(strTmp(0)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt4 & Mid(strTotalAmt8, 1, Len(strTotalAmt8) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(62)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportSum(58)) + 64) & lngCounter & "+" & Chr(GetValue(ReportSum(59)) + 64) & lngCounter & "-" & Chr(GetValue(ReportSum(60)) + 64) & lngCounter & "-" & Chr(GetValue(strTmp(0)) + 64) & lngCounter
'   wksaccrpt417.Range(Chr(GetValue(strTmp(1)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt9 & Mid(Replace(strTotalAmt8, Chr(GetValue(strTmp(0)) + 64), Chr(GetValue(strTmp(1)) + 64)), 1, Len(strTotalAmt8) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(strTmp(2)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt10 & Mid(Replace(strTotalAmt8, Chr(GetValue(strTmp(0)) + 64), Chr(GetValue(strTmp(2)) + 64)), 1, Len(strTotalAmt8) - 2) & ")"
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & "7:" & Chr(UBound(strFieldN) + 64) & lngCounter).Select
'   wksaccrpt417.Range(Chr(GetValue(ReportSum(58)) + 64) & "7:" & Chr(UBound(strFieldN) + 64) & lngCounter).NumberFormatLocal = "#,##0.00_ "
'
'    'Add by Morgan 2005/8/17 加框線
'    wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & "6:" & Chr(UBound(strFieldN) + 64) & lngCounter).Select
'    xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    '2005/8/17 end
'   'end 2015/01/12
'   'Modify by Amy 2014/06/11 +判斷若版本2007以上改變存格式
'   If Val(xlsSalesPoint.Version) < 12 Then
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Mid(ReportTitle(417), 6, 9) & IIf(Left(Combo3, 4) = ComboItem(241), "-結餘", "") & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
'   Else
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Mid(ReportTitle(417), 6, 9) & IIf(Left(Combo3, 4) = ComboItem(241), "-結餘", "") & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
'   End If
'   xlsSalesPoint.Workbooks.Close
'   xlsSalesPoint.Quit
'   StatusClear
End Sub

'Add by Amy  智權人員點數分析表Excel
Private Function ExcelSaveNew1() As Boolean
    Dim xlsSalesPoint As New Excel.Application
    Dim wksaccrpt417 As New Worksheet
    Dim ii As Integer, jj As Integer, intStartR As Integer
    Dim strF As String, strF1 As String, strF2 As String, strF3 As String, strWhere1 As String, strWhere2 As String
    Dim intCounter As Integer, intTitleR As Integer
    Dim xlsFileName As String, strDept As String, strOldDeptN As String, strTemp As String, strTemp2 As String
    Dim bolSumA As Boolean '是否Show 所合計
    Dim strSum, strSumT '記錄 區所/ 全部 合計
    
On Error GoTo ErrHand
    ExcelSaveNew1 = False
    strFieldN = Array("智權人員", "業務達成點數", "P", "T", "CFP", "CFT", "FCP", "FCT", "L", "C", _
                                "FCL/CFL", "保留", "其他", "結餘點數", "其他收入")
    intWidth = Array(9, 11, 9, 9, 9, 9, 9, 9, 9, 9, _
                              9, 9, 9, 9, 9)
   
    If Text1 <> MsgText(602) Then
       Exit Function
    End If
    xlsFileName = Mid(ReportTitle(4171), 6, 9) & ACDate(ServerDate) & ServerTime & MsgText(43)
    If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
          MkDir strExcelPath
       End If
    Else
       Kill strExcelPath & xlsFileName
    End If
    
    strDept = "": intField = 65: intCounter = 1
    xlsSalesPoint.SheetsInNewWorkbook = 3 'Add by Amy  2019/04/02 改設定(選項->一般->包括的工作表份數)
    xlsSalesPoint.Workbooks.add
    Set wksaccrpt417 = xlsSalesPoint.Worksheets(1)
    xlsSalesPoint.Visible = True
    
    '抬頭
    wksaccrpt417.Range(Chr(intField) & intCounter).Value = ReportTitle(4171)
    wksaccrpt417.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter).Select
    With wksaccrpt417.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
    End With
    intCounter = intCounter + 1
    wksaccrpt417.Range(Chr(intField) & intCounter).Value = "公司別:" & GetAccReportCmpN(CboCmp, , True) 'Modify by Amy 2020/04/16
    intCounter = intCounter + 1
    wksaccrpt417.Range(Chr(intField) & intCounter).Value = ReportSum(27) & MaskEdBox1.Text & ReportSum(28) & MaskEdBox2.Text
    intCounter = intCounter + 1
    wksaccrpt417.Range(Chr(intField) & intCounter).Value = "列印人員:" & StaffQuery(strUserNum)
    intCounter = intCounter + 1
    wksaccrpt417.Range(Chr(intField) & intCounter).Value = "列印日期:" & CFDate(ACDate(ServerDate))
    intCounter = intCounter + 1
    For ii = LBound(strFieldN) To UBound(strFieldN)
        wksaccrpt417.Range(Chr(ii + intField) & intCounter).Value = strFieldN(ii)
        wksaccrpt417.Range(Chr(ii + intField) & intCounter).HorizontalAlignment = xlCenter
        wksaccrpt417.Columns(Chr(ii + intField) & ":" & Chr(ii + intField)).ColumnWidth = intWidth(ii)
    Next ii
    intTitleR = intCounter: intCounter = intCounter + 1: intStartR = intCounter
    
    wksaccrpt417.PageSetup.PaperSize = xlPaperA4
    wksaccrpt417.PageSetup.Orientation = xlLandscape '橫印
    wksaccrpt417.PageSetup.PrintTitleRows = "$1:$" & intTitleR
    
    'Modify by Amy 2020/03/31 增加7129 與7121一同列示
    ' 416101列於FCL/4192列RES/增加4194列結餘點數/其他沒列的會計科目列其他收入
    strF = "c0,Sum(C1) as c1, Sum(P) as P,Sum(T) as T,Sum(CFP) as CFP,Sum(CFT) as CFT,Sum(FCP) as FCP,Sum(FCT) as FCT,Sum(L) as L,Sum(C) as C,Sum(FCL) as FCL,Sum(RES) as RES,Sum(ELS) as ELS,Sum(Point) as Point,Sum(Other) as Other "
    strF1 = "r41702 c0,Sum(r41707) c1, 0 as P,0 as T,0 as CFP,0 as CFT,0 as FCP,0 as FCT,0 as L,0 as C,0 as FCL,0 as RES,0 as ELS,0 as Point,0 as Other "
    strF2 = "r41702 c0, 0 as c1" & _
         ",Sum(Decode(SubStr(r41708,1,4),'4111',r41707)) as P" & _
         ",Sum(Decode(SubStr(r41708,1,4),'4101',r41707)) as T" & _
         ",Sum(Decode(SubStr(r41708,1,4),'4131',r41707)) as CFP" & _
         ",Sum(Decode(SubStr(r41708,1,4),'4121',r41707)) as CFT" & _
         ",Sum(Decode(SubStr(r41708,1,4),'4171',r41707)) as FCP" & _
         ",Sum(Decode(SubStr(r41708,1,4),'4172',r41707)) as FCT" & _
         ",Sum(Decode(SubStr(r41708,1,4),'4141',r41707)) as L" & _
         ",Sum(Decode(SubStr(r41708,1,4),'4151',r41707)) as C" & _
         ",Sum(Decode(r41708,'416101',r41707,'416102',r41707)) as FCL" & _
         ",Sum(Decode(r41708,'4191',r41707,'4192',r41707)) as RES" & _
         ",Sum(Decode(r41708,'7121',r41707,'7129',r41707)) as ELS" & _
         ",Sum(Decode(SubStr(r41708,1,4),'4194',r41707)) as Point" & _
         ",0 as Other "
    strF3 = "r41702 c0,0 as c1,0 as P,0 as T,0 as CFP,0 as CFT,0 as FCP,0 as FCT,0 as L,0 as C,0 as FCL,0 as RES,0 as ELS,0 as Point,Sum(r41707) as Other "
         
    strWhere1 = " And (SubStr(r41708,1,4) In ('4111','4101','4131','4121','4171','4172','4141','4151') " & _
                                " Or r41708 In ('416101','416102','4191','4192','7121','7129','4194') ) "
    strWhere2 = " And SubStr(r41708,1,4) Not In ('4111','4101','4131','4121','4171','4172','4141','4151') " & _
                        " And r41708 Not In ('416101','416102','4191','4192','7121','7129','4194') "
    'end 2020/03/31
   
' 智權人員
    strSql = "select * from acc090 where substr(a0901, 1, 1) = 'S' order by a0901 asc"
    If adostaff.State = adStateOpen Then adostaff.Close
    adostaff.CursorLocation = adUseClient
    adostaff.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
    adostaff.MoveFirst
    Do While adostaff.EOF = False
        strSql = "Select " & strF & " From(" & _
        "Select " & strF1 & "From Accrpt417,Staff Where r41701= '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "' And Rtrim(SubStr(r41702, 1, 5)) = St01 Group by r41702 " & _
         "Union Select " & strF2 & "From Accrpt417,Staff Where r41701= '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "' " & strWhere1 & " And Rtrim(SubStr(r41702, 1, 5)) = St01 Group by r41702 " & _
         "Union Select " & strF3 & "From Accrpt417,Staff Where r41701= '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "' " & strWhere2 & " And Rtrim(SubStr(r41702, 1, 5)) = St01 Group by r41702 " & _
                   ") Group by C0"
        If adoaccrpt417.State = adStateOpen Then adoaccrpt417.Close
        adoaccrpt417.CursorLocation = adUseClient
        adoaccrpt417.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
        If adoaccrpt417.RecordCount > 0 Then
            adoaccrpt417.MoveFirst
            Do While adoaccrpt417.EOF = False
                '合計
                If strDept <> adostaff.Fields("a0901") And strDept <> MsgText(601) Then
                    For ii = LBound(strFieldN) To UBound(strFieldN)
                        strTemp = "": strTemp2 = ""
                        '所合計公式(南高所合計設於區)
                        If Mid(strDept, 1, 2) <> Mid(adostaff.Fields("a0901"), 1, 2) And Val(Mid(adostaff.Fields("a0901"), 2, 1)) <= 3 Then
                            If ii = GetValue("智權人員") Then
                                bolSumA = True
                                Select Case Left(strDept, 2)
                                    Case "S1"
                                        strTemp2 = ReportSum(105)
                                    Case "S2"
                                        strTemp2 = ReportSum(106)
                                End Select
                            Else
                                strTemp2 = "=Sum(" & Mid(Replace(strSum, Chr(GetValue("業務達成點數") + intField), Chr(ii + intField)), 2) & ")"
                            End If
                        End If
                        '區合計
                        If ii = GetValue("智權人員") Then
                            strTemp = strOldDeptN
                            strSum = strSum & "," & Chr(GetValue("業務達成點數") + intField) & intCounter
                            strSumT = strSumT & "," & Chr(GetValue("業務達成點數") + intField) & intCounter
                        Else
                            If intStartR = intCounter Then
                                strTemp = "0"
                            Else
                                strTemp = "=Sum(" & Chr(ii + intField) & intStartR & ":" & Chr(ii + intField) & intCounter - 1 & ")"
                            End If
                        End If
                        wksaccrpt417.Range(Chr(ii + intField) & intCounter).Value = strTemp
                        
                        '是否印所合計
                        If bolSumA = True Then
                            wksaccrpt417.Range(Chr(ii + intField) & intCounter + 2).Value = strTemp2
                        End If
                    Next ii
                    If bolSumA = True Then intCounter = intCounter + 2: strSum = "": bolSumA = False
                    intCounter = intCounter + 2: intStartR = intCounter
                End If
                '資料
                For ii = LBound(strFieldN) To UBound(strFieldN)
                    strTemp = ""
                    If ii = GetValue("智權人員") Then
                        strTemp = Mid("" & adoaccrpt417.Fields(0), 6)
                    Else
                        strTemp = Val("" & adoaccrpt417.Fields(ii))
                    End If
                    wksaccrpt417.Range(Chr(ii + intField) & intCounter).Value = strTemp
                Next ii
                intCounter = intCounter + 1
                strDept = "" & adostaff.Fields("a0901")
                strOldDeptN = "" & adostaff.Fields("a0902")
                adoaccrpt417.MoveNext
            Loop
        End If
        adostaff.MoveNext
    Loop
    adostaff.Close
    '最後一個部門加總
    For ii = LBound(strFieldN) To UBound(strFieldN)
        strTemp = ""
        If ii = GetValue("智權人員") Then
            strTemp = strOldDeptN
            strSumT = strSumT & "," & Chr(GetValue("業務達成點數") + intField) & intCounter
        Else
            If intStartR = intCounter Then
                strTemp = "0"
            Else
                strTemp = "=Sum(" & Chr(ii + intField) & intStartR & ":" & Chr(ii + intField) & intCounter - 1 & ")"
            End If
        End If
        wksaccrpt417.Range(Chr(ii + intField) & intCounter).Value = strTemp
    Next ii
    intCounter = intCounter + 2: intStartR = intCounter

' 其他人員
    'Modify by Amy 2021/01/20 +F4104~07
    strSql = "Select " & strF & ",H1 From(" & _
    "Select " & strF1 & ",R41712 as H1 From Accrpt417,Staff Where r41701= '" & strUserNum & "' and SubStr(r41712,1,1)<>'S' And Rtrim(SubStr(r41702, 1, 5)) = St01 And SubStr(r41702, 1, 5) not in ('F4102', 'F4103', 'F4101','F4104', 'F4105', 'F4106', 'F4107') Group by r41712,r41702 " & _
    "Union Select " & strF2 & ",R41712 as H1 From Accrpt417,Staff Where r41701= '" & strUserNum & "' and SubStr(r41712,1,1)<>'S' " & strWhere1 & " And Rtrim(SubStr(r41702, 1, 5)) = St01 And SubStr(r41702, 1, 5) not in ('F4102', 'F4103', 'F4101','F4104', 'F4105', 'F4106', 'F4107') Group by r41712,r41702 " & _
    "Union Select " & strF3 & ",R41712 as H1 From Accrpt417,Staff Where r41701= '" & strUserNum & "' and SubStr(r41712,1,1)<>'S' " & strWhere2 & " And Rtrim(SubStr(r41702, 1, 5)) = St01 And SubStr(r41702, 1, 5) not in ('F4102', 'F4103', 'F4101','F4104', 'F4105', 'F4106', 'F4107') Group by r41712,r41702 " & _
            ") Group by H1,C0 Order by H1,C0"
    If adoaccrpt417.State = adStateOpen Then adoaccrpt417.Close
    adoaccrpt417.CursorLocation = adUseClient
    adoaccrpt417.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
    If adoaccrpt417.RecordCount > 0 Then
        adoaccrpt417.MoveFirst
        Do While adoaccrpt417.EOF = False
            For ii = LBound(strFieldN) To UBound(strFieldN)
                strTemp = ""
                If ii = GetValue("智權人員") Then
                    strTemp = Mid("" & adoaccrpt417.Fields(0), 6)
                Else
                    strTemp = Val("" & adoaccrpt417.Fields(ii))
                End If
                wksaccrpt417.Range(Chr(ii + intField) & intCounter).Value = strTemp
            Next ii
            intCounter = intCounter + 1
            adoaccrpt417.MoveNext
        Loop
    End If
    adoaccrpt417.Close
'非智權部門加總
    For ii = LBound(strFieldN) To UBound(strFieldN)
        strTemp = ""
        If ii = GetValue("智權人員") Then
            strTemp = "其他合計："
            strSumT = strSumT & "," & Chr(GetValue("業務達成點數") + intField) & intCounter
        Else
            If intStartR = intCounter Then
                strTemp = "0"
            Else
                strTemp = "=Sum(" & Chr(ii + intField) & intStartR & ":" & Chr(ii + intField) & intCounter - 1 & ")"
            End If
        End If
        wksaccrpt417.Range(Chr(ii + intField) & intCounter).Value = strTemp
    Next ii
    intCounter = intCounter + 2: intStartR = intCounter
    
'國內合計
    For ii = LBound(strFieldN) To UBound(strFieldN)
        strTemp = ""
        If ii = GetValue("智權人員") Then
            strTemp = "國內合計："
        Else
            strTemp = "=Sum(" & Mid(Replace(strSumT, Chr(GetValue("業務達成點數") + intField), Chr(ii + intField)), 2) & ")"
        End If
        wksaccrpt417.Range(Chr(ii + intField) & intCounter).Value = strTemp
    Next ii
    strSumT = "," & Chr(GetValue("業務達成點數") + intField) & intCounter
    intCounter = intCounter + 2: intStartR = intCounter
    
'FCP/FCT/FCL 合計
    'Modify by Amy 2021/01/20 +F4104~06
    strDept = "": strSum = ""
    strSql = "Select " & strF & ",H1 From(" & _
    "Select " & strF1 & ",R41712 as H1 From Accrpt417,Staff Where r41701= '" & strUserNum & "'  And Rtrim(SubStr(r41702, 1, 5)) = St01 And SubStr(r41702, 1, 5) In ('F4102', 'F4103', 'F4101','F4104', 'F4105', 'F4106', 'F4107') Group by r41712,r41702 " & _
    "Union Select " & strF2 & ",R41712 as H1 From Accrpt417,Staff Where r41701= '" & strUserNum & "' " & strWhere1 & " And Rtrim(SubStr(r41702, 1, 5)) = St01 And SubStr(r41702, 1, 5) In ('F4102', 'F4103', 'F4101','F4104', 'F4105', 'F4106', 'F4107') Group by r41712,r41702 " & _
    "Union Select " & strF3 & ",R41712 as H1 From Accrpt417,Staff Where r41701= '" & strUserNum & "' " & strWhere2 & " And Rtrim(SubStr(r41702, 1, 5)) = St01 And SubStr(r41702, 1, 5) In ('F4102', 'F4103', 'F4101','F4104', 'F4105', 'F4106', 'F4107') Group by r41712,r41702 " & _
            ") Group by H1,C0 Order by Decode(SubStr(H1,5,1),'1',2,1),C0 asc"
    bolSumA = ChkFCPFCTSum(strSql)
    If adoaccrpt417.State = adStateOpen Then adoaccrpt417.Close
    adoaccrpt417.CursorLocation = adUseClient
    adoaccrpt417.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
    If adoaccrpt417.RecordCount > 0 Then
        adoaccrpt417.MoveFirst
        Do While adoaccrpt417.EOF = False
            '*** 合計 (F4102再加F4104,F4105/F4103再加F4106,F4107 需合計) ***
            If bolSumA = True Then
                If strDept <> MsgText(601) And strDept <> adoaccrpt417.Fields("H1") Then
                    For ii = LBound(strFieldN) To UBound(strFieldN)
                        strTemp = ""
                        If ii = GetValue("智權人員") Then
                            strTemp = Mid(strTemp2, 6, 3) & "合計："
                        Else
                            strTemp = "=Sum(" & Chr(ii + intField) & intStartR & ":" & Chr(ii + intField) & intCounter - 1 & ")"
                        End If
                        wksaccrpt417.Range(Chr(ii + intField) & intCounter).Value = strTemp
                        If ii = GetValue("業務達成點數") Then
                            strSum = strSum & "," & Chr(ii + intField) & intCounter
                        End If
                    Next ii
                    intCounter = intCounter + 2: intStartR = intCounter
                End If
            End If
            '*** 合計 End ***
            For ii = LBound(strFieldN) To UBound(strFieldN)
                strTemp = ""
                If ii = GetValue("智權人員") Then
                    strTemp = adoaccrpt417.Fields("C0")
                    If bolSumA = False Then
                        'F4101/F4102/F4103 用
                        strTemp = Mid(strTemp, 6) & "合計："
                    End If
                Else
                    strTemp = Val("" & adoaccrpt417.Fields(ii))
                End If
                wksaccrpt417.Range(Chr(ii + intField) & intCounter).Value = strTemp
            Next ii
            intCounter = intCounter + 1
            strTemp2 = "" & adoaccrpt417.Fields("C0")
            strDept = "" & adoaccrpt417.Fields("H1")
            adoaccrpt417.MoveNext
        Loop
        '最後一個合計
        If bolSumA = True Then
            For ii = LBound(strFieldN) To UBound(strFieldN)
                strTemp = ""
                If ii = GetValue("智權人員") Then
                    strTemp = Mid(strTemp2, 6, 3) & "合計："
                Else
                    strTemp = "=Sum(" & Chr(ii + intField) & intStartR & ":" & Chr(ii + intField) & intCounter - 1 & ")"
                End If
                wksaccrpt417.Range(Chr(ii + intField) & intCounter).Value = strTemp
                If ii = GetValue("業務達成點數") Then
                    strSum = strSum & "," & Chr(ii + intField) & intCounter
                End If
            Next ii
            intCounter = intCounter + 2: intStartR = intCounter
        End If
    End If
    'end 2021/01/20
    
' 國外合計
    For ii = LBound(strFieldN) To UBound(strFieldN)
        strTemp = ""
        If ii = GetValue("智權人員") Then
            strTemp = "國外合計："
        Else
            'Modify by Amy 2021/01/20 +bolSumA = True
            If bolSumA = True Then
                strTemp = "=Sum(" & Mid(Replace(strSum, Chr(GetValue("業務達成點數") + intField), Chr(ii + intField)), 2) & ")"
            ElseIf intStartR = intCounter Then
                strTemp = "0"
            Else
                strTemp = "=Sum(" & Chr(ii + intField) & intStartR & ":" & Chr(ii + intField) & intCounter - 1 & ")"
            End If
        End If
        wksaccrpt417.Range(Chr(ii + intField) & intCounter).Value = strTemp
    Next ii
    strSumT = strSumT & "," & Chr(GetValue("業務達成點數") + intField) & intCounter
    intCounter = intCounter + 2
    
' 全所合計
    For ii = LBound(strFieldN) To UBound(strFieldN)
        strTemp = ""
        If ii = GetValue("智權人員") Then
            strTemp = "全所合計："
        Else
            If intStartR = intCounter Then
                strTemp = "0"
            Else
                strTemp = "=Sum(" & Mid(Replace(strSumT, Chr(GetValue("業務達成點數") + intField), Chr(ii + intField)), 2) & ")"
            End If
        End If
        wksaccrpt417.Range(Chr(ii + intField) & intCounter).Value = strTemp
    Next ii
    
'設定格式
    wksaccrpt417.Range(Chr(GetValue("業務達成點數") + intField) & intTitleR + 1 & ":" & Chr(UBound(strFieldN) + intField) & intCounter).NumberFormatLocal = "#,##0.00_ "
    wksaccrpt417.Range(Chr(GetValue("智權人員") + intField) & intTitleR & ":" & Chr(UBound(strFieldN) + intField) & intCounter).Select
    xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Font.Size = 9

   '判斷若版本2007以上改變存格式
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Mid(ReportTitle(4171), 6, 9) & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Mid(ReportTitle(4171), 6, 9) & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
   End If
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   StatusClear
   ExcelSaveNew1 = True
   Exit Function

ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    Set xlsSalesPoint = Nothing
End Function

'Add by Morgan 2005/8/18  智權人員點數分析表
'Mark by Amy 2018/03/27  業務達成點數不等於其他欄加總,故改寫並修改為動態產生欄位
'*************************************************
'  轉成Excel檔案
'
'*************************************************
Private Sub ExcelSaveNew1_Old()
'   Dim xlsSalesPoint As New Excel.Application
'   Dim wksaccrpt417 As New Worksheet
'   Dim xlsSelect As Selection
'   Dim lngCounter As Long
'   Dim strDept As String
'   Dim strTotAmt(11) As String
'   Dim strTotAmt1(11) As String
'   Dim ii As Integer
'
'   ReDim strSum(11)
'   strDept = ""
'   If Text1 <> MsgText(602) Then
'      Exit Sub
'   End If
'   If Dir(strExcelPath & Mid(ReportTitle(4171), 6, 9) & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
'      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'         MkDir strExcelPath
'      End If
'   Else
'      Kill strExcelPath & Mid(ReportTitle(4171), 6, 9) & ACDate(ServerDate) & ServerTime & MsgText(43)
'   End If
'   xlsSalesPoint.Workbooks.add
'   Set wksaccrpt417 = xlsSalesPoint.Worksheets(1)
'   wksaccrpt417.PageSetup.Orientation = xlLandscape '橫印
'   wksaccrpt417.PageSetup.PrintTitleRows = "$1:$5"
'   wksaccrpt417.Columns("a:a").ColumnWidth = 9
'   wksaccrpt417.Columns("b:b").ColumnWidth = 11
'   wksaccrpt417.Columns("c:c").ColumnWidth = 9
'   wksaccrpt417.Columns("d:d").ColumnWidth = 9
'   wksaccrpt417.Columns("e:e").ColumnWidth = 9
'   wksaccrpt417.Columns("f:f").ColumnWidth = 9
'   wksaccrpt417.Columns("g:g").ColumnWidth = 10
'   wksaccrpt417.Columns("h:h").ColumnWidth = 9
'   wksaccrpt417.Columns("i:i").ColumnWidth = 9
'   wksaccrpt417.Columns("j:j").ColumnWidth = 9
'   wksaccrpt417.Columns("k:k").ColumnWidth = 9
'   wksaccrpt417.Columns("l:l").ColumnWidth = 10
'   wksaccrpt417.Columns("m:m").ColumnWidth = 9
'
'   wksaccrpt417.Range("a1").Value = ReportTitle(4171)
'   wksaccrpt417.Range("a1:m1").Select
'    With wksaccrpt417.Range("a1:k1")
'       .HorizontalAlignment = xlCenter
'       .VerticalAlignment = xlBottom
'       .WrapText = False
'       .Orientation = 0
'       .AddIndent = False
'       .ShrinkToFit = False
'       .MergeCells = True
'   End With
'   'Add By Sindy 2014/1/22
'   wksaccrpt417.Range("a2").Value = "公司別:"
'   wksaccrpt417.Range("b2").Value = IIf(Text4 = "1", "台一", IIf(Text4 = "2", "智權", "台一　專利商標/智權"))
'   '2014/1/22 END
'   wksaccrpt417.Range("a3").Value = ReportSum(27)
'   wksaccrpt417.Range("b3").Value = MaskEdBox1.Text
'   wksaccrpt417.Range("c3").Value = ReportSum(28)
'   wksaccrpt417.Range("d3").Value = MaskEdBox2.Text
'   wksaccrpt417.Range("a5").Value = ReportSum(57)
'   wksaccrpt417.Range("a5").HorizontalAlignment = xlCenter
'   wksaccrpt417.Range("b5").Value = ReportSum(58)
'   wksaccrpt417.Range("b5").HorizontalAlignment = xlCenter
'   wksaccrpt417.Range("c5").Value = "P"
'   wksaccrpt417.Range("c5").HorizontalAlignment = xlCenter
'   wksaccrpt417.Range("d5").Value = "T"
'   wksaccrpt417.Range("d5").HorizontalAlignment = xlCenter
'   wksaccrpt417.Range("e5").Value = "CFP"
'   wksaccrpt417.Range("e5").HorizontalAlignment = xlCenter
'   wksaccrpt417.Range("f5").Value = "CFT"
'   wksaccrpt417.Range("f5").HorizontalAlignment = xlCenter
'   wksaccrpt417.Range("g5").Value = "FCP"
'   wksaccrpt417.Range("g5").HorizontalAlignment = xlCenter
'   wksaccrpt417.Range("h5").Value = "FCT"
'   wksaccrpt417.Range("h5").HorizontalAlignment = xlCenter
'   wksaccrpt417.Range("i5").Value = "L"
'   wksaccrpt417.Range("i5").HorizontalAlignment = xlCenter
'   wksaccrpt417.Range("j5").Value = "C"
'   wksaccrpt417.Range("j5").HorizontalAlignment = xlCenter
'   wksaccrpt417.Range("k5").Value = "FCL"
'   wksaccrpt417.Range("k5").HorizontalAlignment = xlCenter
'   wksaccrpt417.Range("l5").Value = "保留"
'   wksaccrpt417.Range("l5").HorizontalAlignment = xlCenter
'   wksaccrpt417.Range("m5").Value = "其他"
'   wksaccrpt417.Range("m5").HorizontalAlignment = xlCenter
'   lngCounter = 6
'
'' 智權人員
'   strSql = "select * from acc090 where substr(a0901, 1, 1) = 'S' order by a0901 asc"
'   adostaff.CursorLocation = adUseClient
'   adostaff.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   Do While adostaff.EOF = False
'
'      '2011/9/2 modify by sonia 因加員工編號S29,故取5碼去空白讀員工檔,故substr(r41702, 1, 5) = st01改為rtrim(substr(r41702, 1, 5)) = st01
'      'modify by sonia 2016/1/22 decode(r41708,'4131',r41707)改decode(substr(r41708,1,4),'4131',r41707),decode(r41708,'4121',r41707)改decode(substr(r41708,1,4),'4121',r41707)
'      strSql = "select r41702 c0, sum(r41707) c1" & _
'         ",sum(decode(substr(r41708,1,4),'4111',r41707)) as P" & _
'         ",sum(decode(substr(r41708,1,4),'4101',r41707)) as T" & _
'         ",sum(decode(substr(r41708,1,4),'4131',r41707)) as CFP" & _
'         ",sum(decode(substr(r41708,1,4),'4121',r41707)) as CFT" & _
'         ",sum(decode(substr(r41708,1,4),'4171',r41707)) as FCP" & _
'         ",sum(decode(substr(r41708,1,4),'4172',r41707)) as FCT" & _
'         ",sum(decode(substr(r41708,1,4),'4141',r41707)) as L" & _
'         ",sum(decode(substr(r41708,1,4),'4151',r41707)) as C" & _
'         ",sum(decode(r41708,'416102',r41707)) as FCL" & _
'         ",sum(decode(r41708,'4191',r41707)) as RES" & _
'         ",sum(decode(r41708,'7121',r41707)) as ELS" & _
'         " from accrpt417, staff where rtrim(substr(r41702, 1, 5)) = st01 and r41701 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "' group by r41702"
'
'      If adoaccrpt417.State <> adStateClosed Then adoaccrpt417.Close
'      adoaccrpt417.CursorLocation = adUseClient
'      adoaccrpt417.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'
'      If adoaccrpt417.RecordCount > 0 Then
'
'         Do While adoaccrpt417.EOF = False
'            If strDept <> Mid(adostaff.Fields("a0901").Value, 1, 2) Then
'               If strDept <> "" Then
'                  Select Case strDept
'                     Case "S1", "S2"
'
'                        Select Case strDept
'                           Case "S1"
'                              wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(105)
'                           Case "S2"
'                              wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(106)
'                           Case "S3"
'                              wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(107)
'                           Case "S4"
'                              wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(108)
'                        End Select
'                        For ii = 0 To 11
'                           If strSum(ii) <> "" Then
'                              strSum(ii) = Mid(strSum(ii), 1, Len(strSum(ii)) - 1)
'                           End If
'                           wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Formula = strSum(ii)
'                        Next ii
'                        lngCounter = lngCounter + 2
'                  End Select
'               End If
'               For ii = 0 To 11
'                  strSum(ii) = "="
'               Next ii
'               strDept = Mid(adostaff.Fields("a0901").Value, 1, 2)
'            Else
'               If lngCounter > 6 Then
'                  For ii = 0 To 11
'                     strSum(ii) = strSum(ii) & Chr(Asc("b") + ii) & (lngCounter - 1) & "+"
'                  Next ii
'               End If
'            End If
'
'            wksaccrpt417.Range("a" & lngCounter).Value = Mid("" & adoaccrpt417.Fields(0).Value, 6)
'            For ii = 0 To 11
'               wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Value = Val("" & adoaccrpt417.Fields(1 + ii).Value)
'            Next
'
'            lngCounter = lngCounter + 1
'            adoaccrpt417.MoveNext
'         Loop
'
'         For ii = 0 To 11
'            strSum(ii) = strSum(ii) & Chr(Asc("b") + ii) & (lngCounter - 1) & "+"
'         Next
'
'      End If
'
'      '2011/9/2 modify by sonia 因加員工編號S29,故取5碼去空白讀員工檔
'      'strSql = "select count(distinct r41702) from accrpt417, staff where substr(r41702, 1, 5) = st01 and r41701 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "'"
'      strSql = "select count(distinct r41702) from accrpt417, staff where rtrim(substr(r41702, 1, 5)) = st01 and r41701 = '" & strUserNum & "' and st15 = '" & adostaff.Fields("a0901").Value & "'"
'
'      If adoaccsum.State <> adStateClosed Then adoaccsum.Close
'      adoaccsum.CursorLocation = adUseClient
'      adoaccsum.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccsum.RecordCount <> 0 Then
'         If IsNull(adoaccsum.Fields(0).Value) = False And adoaccsum.Fields(0).Value <> 0 Then
'            wksaccrpt417.Range("a" & lngCounter).Value = adostaff.Fields("a0902").Value
'            For ii = 0 To 11
'               wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Formula = "=sum(" & Chr(Asc("b") + ii) & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":" & Chr(Asc("b") + ii) & (lngCounter - 1) & ")"
'               strTotAmt(ii) = strTotAmt(ii) & Chr(Asc("b") + ii) & lngCounter & ", "
'            Next
'            lngCounter = lngCounter + 2
'         End If
'      End If
'      adoaccsum.Close
'      adoaccrpt417.Close
'      adostaff.MoveNext
'   Loop
'   adostaff.Close
'   Select Case strDept
'      Case "S1", "S2"
'         If strSum(0) <> "" Then
'            If strSum(0) <> "" Then
'               strSum(0) = Mid(strSum(0), 1, Len(strSum(0)) - 1)
'            End If
'            If strSum(1) <> "" Then
'               strSum(1) = Mid(strSum(1), 1, Len(strSum(1)) - 1)
'            End If
'            If strSum(2) <> "" Then
'               strSum(2) = Mid(strSum(2), 1, Len(strSum(2)) - 1)
'            End If
'            If strSum(3) <> "" Then
'               strSum(3) = Mid(strSum(3), 1, Len(strSum(3)) - 1)
'            End If
'            If strSum(4) <> "" Then
'               strSum(4) = Mid(strSum(4), 1, Len(strSum(4)) - 1)
'            End If
'            Select Case strDept
'               Case "S1"
'                  wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(105)
'               Case "S2"
'                  wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(106)
'               Case "S3"
'                  wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(107)
'               Case "S4"
'                  wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(108)
'               Case "S9"
'                  wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(127)
'            End Select
'            For ii = 0 To 11
'               If Right(strSum(ii), 1) = "+" Then strSum(ii) = Left(strSum(ii), Len(strSum(ii)) - 1)
'               wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Formula = strSum(ii)
'            Next
'
'            lngCounter = lngCounter + 2
'         End If
'   End Select
'
'' 其他人員
'   strSql = "select * from acc090 where (substr(a0901, 1, 1) <> 'S') or a0901 = '020' order by a0901 asc"
'   If adostaff.State <> adStateClosed Then adostaff.Close
'   adostaff.CursorLocation = adUseClient
'   adostaff.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   Do While adostaff.EOF = False
'      '2011/9/2 modify by sonia 因加員工編號S29,故取5碼去空白讀員工檔,故substr(r41702, 1, 5) = st01改為rtrim(substr(r41702, 1, 5)) = st01
'      '2012/9/4 modify by D101081021,D101082576的416102(CFL010696000)投資法務收文的CFL
'      'modify by sonia 2016/1/22 decode(r41708,'4131',r41707)改decode(substr(r41708,1,4),'4131',r41707),decode(r41708,'4121',r41707)改decode(substr(r41708,1,4),'4121',r41707)
'      strSql = "select r41702 c0, sum(r41707) c1" & _
'         ",sum(decode(substr(r41708,1,4),'4111',r41707)) as P" & _
'         ",sum(decode(substr(r41708,1,4),'4101',r41707)) as T" & _
'         ",sum(decode(substr(r41708,1,4),'4131',r41707)) as CFP" & _
'         ",sum(decode(substr(r41708,1,4),'4121',r41707)) as CFT" & _
'         ",sum(decode(substr(r41708,1,4),'4171',r41707)) as FCP" & _
'         ",sum(decode(substr(r41708,1,4),'4172',r41707)) as FCT" & _
'         ",sum(decode(substr(r41708,1,4),'4141',r41707)) as L" & _
'         ",sum(decode(substr(r41708,1,4),'4151',r41707)) as C" & _
'         ",sum(decode(substr(r41708,1,4),'4161',r41707)) as FCL" & _
'         ",sum(decode(r41708,'4191',r41707)) as RES" & _
'         ",sum(decode(r41708,'7121',r41707)) as ELS" & _
'         " from accrpt417, staff where rtrim(substr(r41702, 1, 5)) = st01 and r41701 = '" & strUserNum & "'" & _
'         " and st15 = '" & adostaff.Fields("a0901").Value & "'" & _
'         " and substr(r41702, 1, 5) not in ('F4102', 'F4103', 'F4101')" & _
'         " group by r41702"
'
'      adoaccrpt417.CursorLocation = adUseClient
'      adoaccrpt417.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'      Do While adoaccrpt417.EOF = False
'         wksaccrpt417.Range("a" & lngCounter).Value = Mid("" & adoaccrpt417.Fields(0).Value, 6)
'         For ii = 0 To 11
'            wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Value = Val("" & adoaccrpt417.Fields(1 + ii).Value)
'         Next
'
'         lngCounter = lngCounter + 1
'         adoaccrpt417.MoveNext
'      Loop
'      adoaccrpt417.Close
'      adostaff.MoveNext
'   Loop
'   adostaff.Close
'
'   '2011/9/2 modify by sonia 因加員工編號S29,故取5碼去空白讀員工檔
'   'strSql = "select count(distinct substr(r41702, 1, 5)) from accrpt417, staff, acc090 where substr(r41702, 1, 5) = st01 and st15 = a0901 and r41701 = '" & strUserNum & "' and ((substr(a0901, 1, 1) <> 'S' and substr(a0901, 1, 1) <> 'F') or a0901 = '020')"
'   strSql = "select count(distinct rtrim(substr(r41702, 1, 5))) from accrpt417, staff, acc090 where rtrim(substr(r41702, 1, 5)) = st01 and st15 = a0901 and r41701 = '" & strUserNum & "' and ((substr(a0901, 1, 1) <> 'S' and substr(a0901, 1, 1) <> 'F') or a0901 = '020')"
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) = False And adoaccsum.Fields(0).Value <> 0 Then
'         wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(64) & ReportSum(25)
'         For ii = 0 To 11
'            wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Formula = "=sum(" & Chr(Asc("b") + ii) & (lngCounter - Val(adoaccsum.Fields(0).Value)) & ":" & Chr(Asc("b") + ii) & (lngCounter - 1) & ")"
'            'strTotAmt(ii) = ""
'            strTotAmt(ii) = strTotAmt(ii) & Chr(Asc("b") + ii) & lngCounter & ", "
'         Next
'         lngCounter = lngCounter + 2
'      End If
'   End If
'   adoaccsum.Close
'
'' 國內合計
'   wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(65) & ReportSum(25)
'   For ii = 0 To 11
'      wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Formula = "=sum(" & Mid(strTotAmt(ii), 1, Len(strTotAmt(ii)) - 2) & ")"
'   Next
'   lngCounter = lngCounter + 2
'
'' FCP
'   wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(67) & ReportSum(25)
'   'modify by sonia 2016/1/22 decode(r41708,'4131',r41707)改decode(substr(r41708,1,4),'4131',r41707),decode(r41708,'4121',r41707)改decode(substr(r41708,1,4),'4121',r41707)
'   strSql = "select r41702 c0, sum(r41707) c1" & _
'      ",sum(decode(substr(r41708,1,4),'4111',r41707)) as P" & _
'      ",sum(decode(substr(r41708,1,4),'4101',r41707)) as T" & _
'      ",sum(decode(substr(r41708,1,4),'4131',r41707)) as CFP" & _
'      ",sum(decode(substr(r41708,1,4),'4121',r41707)) as CFT" & _
'      ",sum(decode(substr(r41708,1,4),'4171',r41707)) as FCP" & _
'      ",sum(decode(substr(r41708,1,4),'4172',r41707)) as FCT" & _
'      ",sum(decode(substr(r41708,1,4),'4141',r41707)) as L" & _
'      ",sum(decode(substr(r41708,1,4),'4151',r41707)) as C" & _
'      ",sum(decode(r41708,'416101',r41707)) as FCL" & _
'      ",sum(decode(r41708,'4192',r41707)) as RES" & _
'      ",sum(decode(r41708,'7121',r41707)) as ELS" & _
'      " from accrpt417 where substr(r41702, 1, 5) = 'F4102' and r41701 = '" & strUserNum & "'" & _
'      " group by r41702"
'
'   adoaccrpt417.CursorLocation = adUseClient
'   adoaccrpt417.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccrpt417.RecordCount <> 0 Then
'      For ii = 0 To 11
'         wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Value = Val("" & adoaccrpt417.Fields(1 + ii).Value)
'         strTotAmt1(ii) = strTotAmt1(ii) & Chr(Asc("b") + ii) & lngCounter & ", "
'      Next
'   End If
'   adoaccrpt417.Close
'   lngCounter = lngCounter + 1
'
'' FCT
'   wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(68) & ReportSum(25)
'   'modify by sonia 2016/1/22 decode(r41708,'4131',r41707)改decode(substr(r41708,1,4),'4131',r41707),decode(r41708,'4121',r41707)改decode(substr(r41708,1,4),'4121',r41707)
'   strSql = "select r41702 c0, sum(r41707) c1" & _
'      ",sum(decode(substr(r41708,1,4),'4111',r41707)) as P" & _
'      ",sum(decode(substr(r41708,1,4),'4101',r41707)) as T" & _
'      ",sum(decode(substr(r41708,1,4),'4131',r41707)) as CFP" & _
'      ",sum(decode(substr(r41708,1,4),'4121',r41707)) as CFT" & _
'      ",sum(decode(substr(r41708,1,4),'4171',r41707)) as FCP" & _
'      ",sum(decode(substr(r41708,1,4),'4172',r41707)) as FCT" & _
'      ",sum(decode(substr(r41708,1,4),'4141',r41707)) as L" & _
'      ",sum(decode(substr(r41708,1,4),'4151',r41707)) as C" & _
'      ",sum(decode(r41708,'416101',r41707)) as FCL" & _
'      ",sum(decode(r41708,'4192',r41707)) as RES" & _
'      ",sum(decode(r41708,'7121',r41707)) as ELS" & _
'      " from accrpt417 where substr(r41702, 1, 5) = 'F4103' and r41701 = '" & strUserNum & "'" & _
'      " group by r41702"
'
'   adoaccrpt417.CursorLocation = adUseClient
'   adoaccrpt417.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccrpt417.RecordCount <> 0 Then
'      For ii = 0 To 11
'         wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Value = Val("" & adoaccrpt417.Fields(1 + ii).Value)
'         strTotAmt1(ii) = strTotAmt1(ii) & Chr(Asc("b") + ii) & lngCounter & ", "
'      Next
'   End If
'   adoaccrpt417.Close
'   lngCounter = lngCounter + 1
'
'' FCL
'   wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(69) & ReportSum(25)
'
'   '2012/9/4 modify by D101081021,D101082576的416102(CFL010696000)投資法務收文的CFL
'   'modify by sonia 2016/1/22 decode(r41708,'4131',r41707)改decode(substr(r41708,1,4),'4131',r41707),decode(r41708,'4121',r41707)改decode(substr(r41708,1,4),'4121',r41707)
'   strSql = "select r41702 c0, sum(r41707) c1" & _
'      ",sum(decode(substr(r41708,1,4),'4111',r41707)) as P" & _
'      ",sum(decode(substr(r41708,1,4),'4101',r41707)) as T" & _
'      ",sum(decode(substr(r41708,1,4),'4131',r41707)) as CFP" & _
'      ",sum(decode(substr(r41708,1,4),'4121',r41707)) as CFT" & _
'      ",sum(decode(substr(r41708,1,4),'4171',r41707)) as FCP" & _
'      ",sum(decode(substr(r41708,1,4),'4172',r41707)) as FCT" & _
'      ",sum(decode(substr(r41708,1,4),'4141',r41707)) as L" & _
'      ",sum(decode(substr(r41708,1,4),'4151',r41707)) as C" & _
'      ",sum(decode(substr(r41708,1,4),'4161',r41707)) as FCL" & _
'      ",sum(decode(r41708,'4192',r41707)) as RES" & _
'      ",sum(decode(r41708,'7121',r41707)) as ELS" & _
'      " from accrpt417 where substr(r41702, 1, 5) = 'F4101' and r41701 = '" & strUserNum & "'" & _
'      " group by r41702"
'
'   adoaccrpt417.CursorLocation = adUseClient
'   adoaccrpt417.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccrpt417.RecordCount <> 0 Then
'      For ii = 0 To 11
'         wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Value = Val("" & adoaccrpt417.Fields(1 + ii).Value)
'         strTotAmt1(ii) = strTotAmt1(ii) & Chr(Asc("b") + ii) & lngCounter & ", "
'      Next
'   End If
'   adoaccrpt417.Close
'   lngCounter = lngCounter + 1
'
'' 國外合計
'   wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(70) & ReportSum(25)
'   For ii = 0 To 11
'      If strTotAmt1(ii) <> "" Then
'         wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Formula = "=sum(" & Mid(strTotAmt1(ii), 1, Len(strTotAmt1(ii)) - 2) & ")"
'      End If
'   Next
'
'   lngCounter = lngCounter + 2
'
'' 總所合計
'   wksaccrpt417.Range("a" & lngCounter).Value = ReportSum(66) & ReportSum(25)
'   For ii = 0 To 11
'      If strTotAmt(ii) <> "" Or strTotAmt1(ii) <> "" Then
'         If strTotAmt1(ii) <> "" Then
'            wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Formula = "=sum(" & strTotAmt(ii) & Mid(strTotAmt1(ii), 1, Len(strTotAmt1(ii)) - 2) & ")"
'         Else
'            wksaccrpt417.Range(Chr(Asc("b") + ii) & lngCounter).Formula = "=sum(" & strTotAmt(ii) & ")"
'         End If
'      End If
'   Next
'
'   wksaccrpt417.Range("b6:m" & lngCounter).Select
'   wksaccrpt417.Range("b6:m" & lngCounter).NumberFormatLocal = "#,##0.00_ "
'
'   'Add by Morgan 2005/8/17 加框線
'   wksaccrpt417.Range("a2:m3").Select
'   xlsSalesPoint.Selection.Font.Size = 9
'
'   wksaccrpt417.Range("a5:m" & lngCounter).Select
'   xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'   xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'   xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
'   xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
'   xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
'   xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
'   xlsSalesPoint.Selection.Font.Size = 9
'   '2005/8/17 end
'   'Modify by Amy 2014/06/11 +判斷若版本2007以上改變存格式
'   If Val(xlsSalesPoint.Version) < 12 Then
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Mid(ReportTitle(4171), 6, 9) & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
'   Else
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Mid(ReportTitle(4171), 6, 9) & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
'   End If
'   xlsSalesPoint.Workbooks.Close
'   xlsSalesPoint.Quit
'   StatusClear
End Sub

'Modify by Amy 2016/03/11 改抓function
'*************************************************
'  轉成Excel檔案
'  智權人員實績與結餘分析表
'*************************************************
Private Sub ExcelSaveNew2_OLD2()
'Dim xlsSalesPoint As New Excel.Application
'Dim wksaccrpt417 As New Worksheet
'Dim lngCounter As Long, lngCounter1 As Long
'Dim stSQL As String, strWhere As String, strDept As String, strSqlF As String, strCmp As String
'Dim strTemp(1 To 6) As String, strValue(4) As String
'Dim intTitleRow As Integer '表頭列數
'Dim bolData As Boolean '是否有資料
'
'    Call ClearVar
'    strDept = ""
'    If Text1 <> MsgText(602) Then Exit Sub
'    If Trim(Text4) <> MsgText(601) Then strCmp = IIf(Text4 = "2", "J", "1")
'
'    If Dir(strExcelPath & "智權點數實績與結餘分析表" & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
'        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'            MkDir strExcelPath
'        End If
'    Else
'        Kill strExcelPath & "智權點數實績與結餘分析表" & ACDate(ServerDate) & ServerTime & MsgText(43)
'    End If
'    'xlsSalesPoint.Visible = True
'    xlsSalesPoint.Workbooks.Add
'    Set wksaccrpt417 = xlsSalesPoint.Worksheets(1)
'    wksaccrpt417.Range("a3").Value = "公司別:"
'    wksaccrpt417.Range("b3").Value = IIf(Text4 = "1", "台一", IIf(Text4 = "2", "智權", "台一　專利商標/智權"))
'    wksaccrpt417.Range("a4").Value = ReportSum(27)
'    wksaccrpt417.Range("b4").Value = MaskEdBox1.Text
'    wksaccrpt417.Range("c4").Value = ReportSum(28)
'    wksaccrpt417.Range("d4").Value = MaskEdBox2.Text
'
'    '欄位名稱(欄位需照順序放)
'    ReDim strFieldN(0 To 12)
'    ReDim intWidth(0 To 12)
'    ReDim strSum(1 To 12)
'    ReDim strTotalAmt(1 To 12)
'    i = 0: intTitleRow = 6: lngCounter = 6: lngCounter1 = 0: intField = 65
'    '智權人員
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportSum(57): strFieldN(i) = ReportSum(57): intWidth(i) = 10.5: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'    '期初實績保留
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'    '期初結餘保留
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'    '實績點數
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'    '結餘點數
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'    '期未實績保留
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'    '期末結餘保留
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'    '實績撥點數
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'    '結餘撥點數
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'    '報出實績點數
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'    '報出結餘點數
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'    '報出點數
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'    '實績保留增減
'    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
'    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'
'    intField = 65
'    lngCounter = lngCounter + 1 '資料欄位列
'
'    wksaccrpt417.PageSetup.PrintTitleRows = "$1:$" & intTitleRow
'    For i = LBound(strFieldN) To UBound(strFieldN)
'        wksaccrpt417.Columns(Chr(i + intField) & ":" & Chr(i + intField)).ColumnWidth = intWidth(i)
'    Next i
'    For i = LBound(strSum) To UBound(strSum)
'        strSum(i) = "="
'    Next i
'    wksaccrpt417.Range("a1").Value = "智權點數實績與結餘分析表"
'    wksaccrpt417.Range("a1:" & Chr(UBound(strFieldN) + intField) & "1").Select
'
'    If Val(Left(FCDate(MaskEdBox1.Text), 5)) >= Val(業績輸入啟用年月) Then
'        '抓SalesPoint
'        strSqlF = GetPoint_SP(Val(Left(FCDate(MaskEdBox1.Text), 5)), Val(Left(FCDate(MaskEdBox2.Text), 5)), "S00", "S99", Text2, , Me.Name, True)
'         If adoFinal.State <> adStateClosed Then adoFinal.Close
'        adoFinal.CursorLocation = adUseClient
'        adoFinal.Open strSqlF, adoTaie, adOpenStatic, adLockReadOnly
'        If adoFinal.EOF = False Then adoFinal.MoveFirst
'    End If
'
'    '*** 智權人員
'    lngCounter1 = intTitleRow + 1
'    stSQL = GetPoint(0, FCDate(MaskEdBox1.Text), FCDate(MaskEdBox2.Text), "S", , Text2, False, Me.Name, , strCmp)
'    If adostaff.State <> adStateClosed Then adostaff.Close
'    adostaff.CursorLocation = adUseClient
'    adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'    If adostaff.EOF = False Then adostaff.MoveFirst
'    Do While adostaff.EOF = False
'        If IsNull(adostaff.Fields("ST01").Value) = False Then
'            '區合計
'            If strDept <> adostaff.Fields("SP48") And strDept <> MsgText(601) Then
'                Call GetTotal(0, wksaccrpt417, strDept, lngCounter, lngCounter1)
'                lngCounter = lngCounter + 2
'                lngCounter1 = lngCounter
'            End If
'            '北/中所合計
'            If Mid(strDept, 1, 2) <> Mid(adostaff.Fields("SP48"), 1, 2) And (Mid(strDept, 1, 2) = "S1" Or Mid(strDept, 1, 2) = "S2") Then
'                Call GetTotal(1, wksaccrpt417, strDept, lngCounter)
'                lngCounter = lngCounter + 2
'                lngCounter1 = lngCounter
'            End If
'            '資料
'            Call GetPersonData(2, wksaccrpt417, adostaff, adoFinal, lngCounter)
'            lngCounter = lngCounter + 1
'        End If
'        strDept = adostaff.Fields("SP48")
'        adostaff.MoveNext
'    Loop
'    If adostaff.RecordCount > 0 Then
'        '智權最後一個部門合計
'        Call GetTotal(0, wksaccrpt417, strDept, lngCounter, lngCounter1)
'        lngCounter = lngCounter + 2
'        '智權智權部合計
'        Call GetTotal(2, wksaccrpt417, strDept, lngCounter)
'        lngCounter = lngCounter + 2
'    End If
'
'    '*** 非智權部門
'    lngCounter1 = lngCounter
'    If Val(Left(FCDate(MaskEdBox1.Text), 5)) >= Val(業績輸入啟用年月) Then
'        '抓SalesPoint
'        strSqlF = GetPoint_SP(Val(Left(FCDate(MaskEdBox1.Text), 5)), Val(Left(FCDate(MaskEdBox2.Text), 5)), "*NS", , Text2, , Me.Name, True, True)
'        If adoFinal.State <> adStateClosed Then adoFinal.Close
'        adoFinal.CursorLocation = adUseClient
'        adoFinal.Open strSqlF, adoTaie, adOpenStatic, adLockReadOnly
'        If adoFinal.EOF = False Then adoFinal.MoveFirst
'    End If
'    '其他人員(不包含 M0100)
'    stSQL = GetPoint(0, FCDate(MaskEdBox1.Text), FCDate(MaskEdBox2.Text), "*NS", , Text2, False, Me.Name, True, strCmp)
'    If adostaff.State <> adStateClosed Then adostaff.Close
'    adostaff.CursorLocation = adUseClient
'    adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'    If adostaff.EOF = False Then adostaff.MoveFirst
'    If adostaff.RecordCount > 0 Then bolData = True
'    Do While adostaff.EOF = False
'        Call GetPersonData(3, wksaccrpt417, adostaff, adoFinal, lngCounter)
'        lngCounter = lngCounter + 1
'        adostaff.MoveNext
'    Loop
'    If Trim(Text2) = "" Or Text2 = "M0100" Then
'        'M0100-All
'        stSQL = GetPoint(0, FCDate(MaskEdBox1.Text), FCDate(MaskEdBox2.Text), "TOT", , "M0100", False, Me.Name, True, strCmp)
'        If adostaff.State <> adStateClosed Then adostaff.Close
'        adostaff.CursorLocation = adUseClient
'        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'        If adostaff.EOF = False Then adostaff.MoveFirst
'        Do While adostaff.EOF = False
'            '記錄M0100 Total值
'            Call GetPersonData(31, wksaccrpt417, adostaff, adoFinal, lngCounter)
'            adostaff.MoveNext
'        Loop
'        'M0100-P
'        stSQL = GetPoint(0, FCDate(MaskEdBox1.Text), FCDate(MaskEdBox2.Text), "TOT", , "M0100", False, Me.Name, True, strCmp, "P")
'        If adostaff.State <> adStateClosed Then adostaff.Close
'        adostaff.CursorLocation = adUseClient
'        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'        If adostaff.EOF = False Then adostaff.MoveFirst
'        Do While adostaff.EOF = False
'            Call GetPersonData(32, wksaccrpt417, adostaff, adoFinal, lngCounter)
'            lngCounter = lngCounter + 1
'            adostaff.MoveNext
'        Loop
'        'M0100-T
'        stSQL = GetPoint(0, FCDate(MaskEdBox1.Text), FCDate(MaskEdBox2.Text), "TOT", , "M0100", False, Me.Name, True, strCmp, "T")
'        If adostaff.State <> adStateClosed Then adostaff.Close
'        adostaff.CursorLocation = adUseClient
'        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'        If adostaff.EOF = False Then adostaff.MoveFirst
'        Do While adostaff.EOF = False
'            Call GetPersonData(33, wksaccrpt417, adostaff, adoFinal, lngCounter)
'            lngCounter = lngCounter + 1
'            adostaff.MoveNext
'        Loop
'        'M0100-大陸P-大陸T
'        Call GetPersonData(34, wksaccrpt417, adostaff, adoFinal, lngCounter)
'        lngCounter = lngCounter + 1
'    End If
'    If adostaff.RecordCount > 0 Or bolData = True Then
'        '其他人員合計
'        Call GetTotal(3, wksaccrpt417, strDept, lngCounter, lngCounter1)
'        lngCounter = lngCounter + 2
'        '國內合計
'        Call GetTotal(4, wksaccrpt417, strDept, lngCounter)
'        lngCounter = lngCounter + 2
'    End If
'    '*** 國外部
'    bolData = False: lngCounter1 = lngCounter
'    If Trim(Text2) = "" Or Text2 = "F4101" Or Text2 = "F4102" Or Text2 = "F4103" Then
'        'F4102-FCP
'         If Val(Left(FCDate(MaskEdBox1.Text), 5)) >= Val(業績輸入啟用年月) Then
'            '抓SalesPoint
'            strSqlF = GetPoint_SP(Val(Left(FCDate(MaskEdBox1.Text), 5)), Val(Left(FCDate(MaskEdBox2.Text), 5)), "*NS", , "F4102", , Me.Name, True, True)
'            If adoFinal.State <> adStateClosed Then adoFinal.Close
'            adoFinal.CursorLocation = adUseClient
'            adoFinal.Open strSqlF, adoTaie, adOpenStatic, adLockReadOnly
'            If adoFinal.EOF = False Then adoFinal.MoveFirst
'        End If
'        stSQL = GetPoint(0, FCDate(MaskEdBox1.Text), FCDate(MaskEdBox2.Text), "*NS", , "F4102", False, Me.Name, True, strCmp)
'        If adostaff.State <> adStateClosed Then adostaff.Close
'        adostaff.CursorLocation = adUseClient
'        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'        If adostaff.EOF = False Then adostaff.MoveFirst
'        Do While adostaff.EOF = False
'            bolData = True
'            Call GetPersonData(5, wksaccrpt417, adostaff, adoFinal, lngCounter)
'            lngCounter = lngCounter + 1
'            adostaff.MoveNext
'        Loop
'
'        'F4103-FCT
'        If Val(Left(FCDate(MaskEdBox1.Text), 5)) >= Val(業績輸入啟用年月) Then
'            '抓SalesPoint
'            strSqlF = GetPoint_SP(Val(Left(FCDate(MaskEdBox1.Text), 5)), Val(Left(FCDate(MaskEdBox2.Text), 5)), "*NS", , "F4103", , Me.Name, True, True)
'            If adoFinal.State <> adStateClosed Then adoFinal.Close
'            adoFinal.CursorLocation = adUseClient
'            adoFinal.Open strSqlF, adoTaie, adOpenStatic, adLockReadOnly
'            If adoFinal.EOF = False Then adoFinal.MoveFirst
'        End If
'        stSQL = GetPoint(0, FCDate(MaskEdBox1.Text), FCDate(MaskEdBox2.Text), "*NS", , "F4103", False, Me.Name, True, strCmp)
'        If adostaff.State <> adStateClosed Then adostaff.Close
'        adostaff.CursorLocation = adUseClient
'        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'        If adostaff.EOF = False Then adostaff.MoveFirst
'        Do While adostaff.EOF = False
'            bolData = True
'            Call GetPersonData(5, wksaccrpt417, adostaff, adoFinal, lngCounter)
'            lngCounter = lngCounter + 1
'            adostaff.MoveNext
'        Loop
'
'        'F4101-FCL(10501起不使用此,故不需抓SalesPoint)
'        stSQL = GetPoint(0, FCDate(MaskEdBox1.Text), FCDate(MaskEdBox2.Text), "*NS", , "F4101", False, Me.Name, True, strCmp)
'        If adostaff.State <> adStateClosed Then adostaff.Close
'        adostaff.CursorLocation = adUseClient
'        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'        If adostaff.EOF = False Then adostaff.MoveFirst
'        Do While adostaff.EOF = False
'            bolData = True
'            Call GetPersonData(5, wksaccrpt417, adostaff, adoFinal, lngCounter)
'            lngCounter = lngCounter + 1
'            adostaff.MoveNext
'        Loop
'        '國外合計
'        If bolData = True Then
'            Call GetTotal(5, wksaccrpt417, strDept, lngCounter, lngCounter1)
'            lngCounter = lngCounter + 2
'        End If
'    End If
'    '全所合計
'    Call GetTotal(6, wksaccrpt417, strDept, lngCounter)
'
'    '格式設定
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + intField) & intTitleRow & ":" & Chr(UBound(strFieldN) + intField) & lngCounter).Select
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + intField) & intTitleRow & ":" & Chr(UBound(strFieldN) + intField) & lngCounter).NumberFormatLocal = "#,##0.00_ "
'
'    '框線
'    wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + intField) & intTitleRow & ":" & Chr(UBound(strFieldN) + intField) & lngCounter).Select
'    xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Font.Size = 8
'    'Modify by Amy 2014/06/11 +判斷若版本2007以上改變存格式
'   If Val(xlsSalesPoint.Version) < 12 Then
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & "智權點數實績與結餘分析表" & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
'   Else
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & "智權點數實績與結餘分析表" & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
'   End If
'   xlsSalesPoint.Workbooks.Close
'   xlsSalesPoint.Quit
'   StatusClear
End Sub

''Add by Amy 2015/01/16
''*************************************************
''  轉成Excel檔案
''  智權人員實績與結餘分析表
''*************************************************
Private Sub ExcelSaveNew2_Old()
'Dim xlsSalesPoint As New Excel.Application
'Dim wksaccrpt417 As New Worksheet
'Dim xlsSelect As Selection
'Dim lngCounter As Long, lngCounter1 As Long
'Dim strTotalAmt(1 To 12) As String, strTotalAmtFC(1 To 12) As String 'Modify by Amy 2015/12/04 原:11
'Dim stSQL As String, strWhere As String, strDept As String
'Dim strSum(11) As String 'Modify by Amy 2015/12/7 原:7
'Dim strTemp(1 To 6) As String
'Dim intTitleRow As Integer
'Dim ii As Integer, intField As Integer
'Dim bolNoData As Boolean
'Dim strM0100(1 To 6) As String
''Add by Amy 2016/03/11
'Dim strSqlSP As String
'Dim strValue(3) As String
'
'   strDept = ""
'   If Text1 <> MsgText(602) Then
'      Exit Sub
'   End If
'   If Dir(strExcelPath & "智權點數實績與結餘分析表" & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
'      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'         MkDir strExcelPath
'      End If
'   Else
'      Kill strExcelPath & "智權點數實績與結餘分析表" & ACDate(ServerDate) & ServerTime & MsgText(43)
'   End If
'   xlsSalesPoint.Workbooks.Add
'   Set wksaccrpt417 = xlsSalesPoint.Worksheets(1)
'   wksaccrpt417.Range("a3").Value = "公司別:"
'   wksaccrpt417.Range("b3").Value = IIf(Text4 = "1", "台一", IIf(Text4 = "2", "智權", "台一　專利商標/智權"))
'   wksaccrpt417.Range("a4").Value = ReportSum(27)
'   wksaccrpt417.Range("b4").Value = MaskEdBox1.Text
'   wksaccrpt417.Range("c4").Value = ReportSum(28)
'   wksaccrpt417.Range("d4").Value = MaskEdBox2.Text
'
'   '欄位名稱
'   ReDim strFieldN(1 To 13) 'Modify by Amy 2015/12/04 +欄位
'   ReDim intWidth(1 To 13)
'   ii = 1: intTitleRow = 6: lngCounter = 6: intField = 65
'   '智權人員
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportSum(57): strFieldN(ii) = ReportSum(57): intWidth(ii) = 10.5: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'   '期初實績保留
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(1): strFieldN(ii) = ReportFieldN2(1): intWidth(ii) = 10: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'   '期初結餘保留
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(2): strFieldN(ii) = ReportFieldN2(2): intWidth(ii) = 10: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'   '實績點數
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(3): strFieldN(ii) = ReportFieldN2(3): intWidth(ii) = 10: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'   '結餘點數
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(4): strFieldN(ii) = ReportFieldN2(4): intWidth(ii) = 10: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'  '期未實績保留
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(5): strFieldN(ii) = ReportFieldN2(5): intWidth(ii) = 10: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'   '期末結餘保留
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(6): strFieldN(ii) = ReportFieldN2(6): intWidth(ii) = 10: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'   'Add by Amy 2015/02/05 +欄位
'   '加轉撥點數
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(10): strFieldN(ii) = ReportFieldN2(10): intWidth(ii) = 10: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'   '減轉撥點數
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(11): strFieldN(ii) = ReportFieldN2(11): intWidth(ii) = 10: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'   'end 2015/02/05
'   '報出實績點數
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(7): strFieldN(ii) = ReportFieldN2(7): intWidth(ii) = 10: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'   '報出結餘點數
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(8): strFieldN(ii) = ReportFieldN2(8): intWidth(ii) = 10: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'   '報出點數
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(9): strFieldN(ii) = ReportFieldN2(9): intWidth(ii) = 10: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'   '實績保留增減 Add by Amy 2015/12/04
'   wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(12): strFieldN(ii) = ReportFieldN2(12): intWidth(ii) = 10: ii = ii + 1
'   wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
'
'   lngCounter = lngCounter + 1 '資料欄位列
'
'   '2015/1/29 MODIFY BY SONIA
'   'wksaccrpt417.PageSetup.PrintTitleRows = "$1:$" & UBound(strFieldN)
'   'Modify by Amy 2016/03/11
'   'wksaccrpt417.PageSetup.PrintTitleRows = "$1:$7"
'   wksaccrpt417.PageSetup.PrintTitleRows = "$1:$" & intTitleRow
'   For ii = 1 To UBound(strFieldN)
'        wksaccrpt417.Columns(Chr(ii + 64) & ":" & Chr(ii + 64)).ColumnWidth = intWidth(ii)
'   Next ii
'
'   wksaccrpt417.Range("a1").Value = "智權點數實績與結餘分析表"
'   wksaccrpt417.Range("a1:" & Chr(UBound(strFieldN) + 64) & "1").Select
'
'   'Mark by Amy 2015/02/05 婉莘說不合併儲格
''   With wksaccrpt417.Range("a1:" & Chr(UBound(strFieldN) + 64) & "1")
''       .HorizontalAlignment = xlCenter
''       .VerticalAlignment = xlBottom
''       .WrapText = False
''       .Orientation = 0
''       .AddIndent = False
''       .ShrinkToFit = False
''       .MergeCells = True
''   End With
'
'   If Trim(Text4) <> "" Then strWhere = " And r41713='" & Text4 & "' "
'   'Add by Amy 2016/03/11 +抓取智權人員點數輸入資料,智權部門
'   strSqlSP = GetPoint_SP(Val(Left(FCDate(MaskEdBox1.Text), 5)), Val(Left(FCDate(MaskEdBox2.Text), 5)), "S00", "S99", , , Me.Name)
'   If adoFinal.State <> adStateClosed Then adoFinal.Close
'   adoFinal.CursorLocation = adUseClient
'   adoFinal.Open strSqlSP, adoTaie, adOpenStatic, adLockReadOnly
'   adoFinal.MoveFirst
'
'   ' 智權人員
'   adostaff.Open "select * from acc090 where substr(a0901, 1, 1) = 'S' order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adostaff.EOF = False
'
'        stSQL = "Select Distinct(r41702) as StaffNo From accrpt417 " & _
'                      "Where r41701 = '" & strUserNum & "' and r41712 = '" & adostaff.Fields("a0901").Value & "' " & strWhere
'        adoaccsum.CursorLocation = adUseClient
'        adoaccsum.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'        ii = 0
'        Do While adoaccsum.EOF = False
'            If IsNull(adoaccsum.Fields("StaffNo").Value) = False And adoaccsum.Fields("StaffNo").Value <> 0 Then
'                '北/中所合計
'                If strDept <> Mid(adostaff.Fields("a0901").Value, 1, 2) Then
'                    If strDept <> "" Then
'                         Select Case strDept
'                               Case "S1", "S2"
'                                        If strSum(0) <> "" Then strSum(0) = Mid(strSum(0), 1, Len(strSum(0)) - 1)
'                                        If strSum(1) <> "" Then strSum(1) = Mid(strSum(1), 1, Len(strSum(1)) - 1)
'                                        If strSum(2) <> "" Then strSum(2) = Mid(strSum(2), 1, Len(strSum(2)) - 1)
'                                        If strSum(3) <> "" Then strSum(3) = Mid(strSum(3), 1, Len(strSum(3)) - 1)
'                                        If strSum(4) <> "" Then strSum(4) = Mid(strSum(4), 1, Len(strSum(4)) - 1)
'                                        If strSum(5) <> "" Then strSum(5) = Mid(strSum(5), 1, Len(strSum(5)) - 1)
'                                        'Add by Amy 2015/02/05 +加轉發點數及減轉撥點數
'                                        If strSum(6) <> "" Then strSum(6) = Mid(strSum(6), 1, Len(strSum(6)) - 1)
'                                        If strSum(7) <> "" Then strSum(7) = Mid(strSum(7), 1, Len(strSum(7)) - 1)
'                                        'Add by Amy 2015/12/07 合計改統一寫法,都用strSum
'                                        If strSum(8) <> "" Then strSum(8) = Mid(strSum(8), 1, Len(strSum(8)) - 1)
'                                        If strSum(9) <> "" Then strSum(9) = Mid(strSum(9), 1, Len(strSum(9)) - 1)
'                                        If strSum(10) <> "" Then strSum(10) = Mid(strSum(10), 1, Len(strSum(10)) - 1)
'                                        If strSum(11) <> "" Then strSum(11) = Mid(strSum(11), 1, Len(strSum(11)) - 1)
'                                        Select Case strDept
'                                           Case "S1"
'                                              wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(105)
'                                           Case "S2"
'                                              wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(106)
'                                           Case "S3"
'                                              wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(107)
'                                           Case "S4"
'                                              wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(108)
'                                           Case "S9"
'                                              wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(127)
'                                        End Select
'                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Formula = strSum(0)
'                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Formula = strSum(1)
'                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Formula = strSum(2)
'                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Formula = strSum(3)
'                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Formula = strSum(4)
'                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Formula = strSum(5)
'                                        'Modify by Amy 2015/02/05 +加轉發點數及減轉撥點數,改加減欄位
'                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Formula = strSum(6)
'                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Formula = strSum(7)
'                                        'Modify by Amy 2015/12/07 合計改統一寫法,都用strSum
''                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & _
''                                                                                                                                                                                 "+" & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter
''                                        'end 2015/02/05
''                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter
''                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter
''                                        'Add by Amy 2015/12/04 +實績保留增減
''                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter
'                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = strSum(8)
'                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = strSum(9)
'                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = strSum(10)
'                                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = strSum(11)
'                                        'end 2015/12/07
'                                        lngCounter = lngCounter + 2
'                         End Select
'                    End If
'                    strSum(0) = "=": strSum(1) = "=": strSum(2) = "=": strSum(3) = "=": strSum(4) = "=": strSum(5) = "=": strSum(6) = "=": strSum(7) = "="
'                    'Add byAmy 2015/12/07
'                    strSum(8) = "=": strSum(9) = "=": strSum(10) = "=": strSum(11) = "="
'                End If '北/中所合計
'
'                ii = ii + 1
'                '資料
'                'Modfiy by Amy 2016/03/11 +10501月抓SalesPoint資料
'                strValue(0) = "0": strValue(1) = "0": strValue(2) = "0": strValue(3) = "0"
'                If Val(Left(FCDate(MaskEdBox1.Text), 5)) >= Val(業績輸入啟用年月) Then
'                    If adoaccsum.Fields("StaffNo").Value = adoFinal.Fields("SP02") Then
'                        strValue(0) = adoFinal.Fields("SP15")
'                        strValue(1) = adoFinal.Fields("SP36")
'                        strValue(2) = adoFinal.Fields("SP19")
'                        strValue(3) = adoFinal.Fields("SP40")
'                        adoFinal.MoveNext
'                    End If
'                Else
'                    strValue(0) = GetEndMonthDebit(Left(RTrim(adoaccsum.Fields("StaffNo").Value), 5))
'                    strValue(1) = Get4194Debit(Left(RTrim(adoaccsum.Fields("StaffNo").Value), 5))
'                    strValue(2) = 0
'                    strValue(3) = 0
'                End If
'                wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = adoaccsum.Fields("StaffNo").Value
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Value = GetStartMonthCredit(Left(RTrim(adoaccsum.Fields("StaffNo").Value), 5))
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Value = Get4194Credit(Left(RTrim(adoaccsum.Fields("StaffNo").Value), 5))
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Value = GetSurplus(Left(RTrim(adoaccsum.Fields("StaffNo").Value), 5))
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Value = GetNotSurplus(Left(RTrim(adoaccsum.Fields("StaffNo").Value), 5))
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Value = strValue(0)
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Value = strValue(1)
'                'Modify by 2015/02/05 +加轉發點數及減轉撥點數,改加減欄位
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Value = strValue(2)
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Value = strValue(3)
'                'end 2016/03/11
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & _
'                                                                                                                                                        "+" & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter
'                'end 2015/02/05
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter
'                'Add by Amy 2015/12/04 +實績保留增減
'                'Modify by Amy 2016/03/11 瑞婷說改相反 原:期初實績保留-期末實績保留
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter
'                lngCounter = lngCounter + 1
'
'                If ii = adoaccsum.RecordCount Then
'                    '合計
'                    wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = adostaff.Fields("a0902").Value & ReportSum(25)
'                    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(1)) + 64) & (lngCounter - Val(adoaccsum.RecordCount)) & ":" & Chr(GetValue(ReportFieldN2(1)) + 64) & (lngCounter - 1) & ")"
'                    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(2)) + 64) & (lngCounter - Val(adoaccsum.RecordCount)) & ":" & Chr(GetValue(ReportFieldN2(2)) + 64) & (lngCounter - 1) & ")"
'                    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(3)) + 64) & (lngCounter - Val(adoaccsum.RecordCount)) & ":" & Chr(GetValue(ReportFieldN2(3)) + 64) & (lngCounter - 1) & ")"
'                    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(4)) + 64) & (lngCounter - Val(adoaccsum.RecordCount)) & ":" & Chr(GetValue(ReportFieldN2(4)) + 64) & (lngCounter - 1) & ")"
'                    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(5)) + 64) & (lngCounter - Val(adoaccsum.RecordCount)) & ":" & Chr(GetValue(ReportFieldN2(5)) + 64) & (lngCounter - 1) & ")"
'                    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(6)) + 64) & (lngCounter - Val(adoaccsum.RecordCount)) & ":" & Chr(GetValue(ReportFieldN2(6)) + 64) & (lngCounter - 1) & ")"
'                    'Add by 2015/02/05 +加轉發點數及減轉撥點數
'                    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(10)) + 64) & (lngCounter - Val(adoaccsum.RecordCount)) & ":" & Chr(GetValue(ReportFieldN2(10)) + 64) & (lngCounter - 1) & ")"
'                    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(11)) + 64) & (lngCounter - Val(adoaccsum.RecordCount)) & ":" & Chr(GetValue(ReportFieldN2(11)) + 64) & (lngCounter - 1) & ")"
'                    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(7)) + 64) & (lngCounter - Val(adoaccsum.RecordCount)) & ":" & Chr(GetValue(ReportFieldN2(7)) + 64) & (lngCounter - 1) & ")"
'                    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(8)) + 64) & (lngCounter - Val(adoaccsum.RecordCount)) & ":" & Chr(GetValue(ReportFieldN2(8)) + 64) & (lngCounter - 1) & ")"
'                    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(9)) + 64) & (lngCounter - Val(adoaccsum.RecordCount)) & ":" & Chr(GetValue(ReportFieldN2(9)) + 64) & (lngCounter - 1) & ")"
'                    'Add by Amy 2015/12/04 +實績保留增減
'                    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(12)) + 64) & (lngCounter - Val(adoaccsum.RecordCount)) & ":" & Chr(GetValue(ReportFieldN2(12)) + 64) & (lngCounter - 1) & ")"
'
'                    strSum(0) = strSum(0) & Chr(GetValue(ReportFieldN2(1)) + 64) & (lngCounter) & "+"
'                    strSum(1) = strSum(1) & Chr(GetValue(ReportFieldN2(2)) + 64) & (lngCounter) & "+"
'                    strSum(2) = strSum(2) & Chr(GetValue(ReportFieldN2(3)) + 64) & (lngCounter) & "+"
'                    strSum(3) = strSum(3) & Chr(GetValue(ReportFieldN2(4)) + 64) & (lngCounter) & "+"
'                    strSum(4) = strSum(4) & Chr(GetValue(ReportFieldN2(5)) + 64) & (lngCounter) & "+"
'                    strSum(5) = strSum(5) & Chr(GetValue(ReportFieldN2(6)) + 64) & (lngCounter) & "+"
'                    'Add by Amy 2015/02/05 +加轉發點數及減轉撥點數
'                    strSum(6) = strSum(6) & Chr(GetValue(ReportFieldN2(10)) + 64) & (lngCounter) & "+"
'                    strSum(7) = strSum(7) & Chr(GetValue(ReportFieldN2(11)) + 64) & (lngCounter) & "+"
'                    'Add byAmy 2015/12/07 合計改統一寫法,都用strSum
'                    strSum(8) = strSum(8) & Chr(GetValue(ReportFieldN2(7)) + 64) & (lngCounter) & "+"
'                    strSum(9) = strSum(9) & Chr(GetValue(ReportFieldN2(8)) + 64) & (lngCounter) & "+"
'                    strSum(10) = strSum(10) & Chr(GetValue(ReportFieldN2(9)) + 64) & (lngCounter) & "+"
'                    strSum(11) = strSum(11) & Chr(GetValue(ReportFieldN2(12)) + 64) & (lngCounter) & "+"
'
'                    strTotalAmt(1) = strTotalAmt(1) & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & ", "
'                    strTotalAmt(2) = strTotalAmt(2) & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & ", "
'                    strTotalAmt(3) = strTotalAmt(3) & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & ", "
'                    strTotalAmt(4) = strTotalAmt(4) & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & ", "
'                    strTotalAmt(5) = strTotalAmt(5) & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & ", "
'                    strTotalAmt(6) = strTotalAmt(6) & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter & ", "
'                    'Add by Amy 2015/02/05 +加轉發點數及減轉撥點數
'                    strTotalAmt(10) = strTotalAmt(10) & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & ", "
'                    strTotalAmt(11) = strTotalAmt(11) & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter & ", "
'                    strTotalAmt(7) = strTotalAmt(7) & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & ", "
'                    strTotalAmt(8) = strTotalAmt(8) & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter & ", "
'                    strTotalAmt(9) = strTotalAmt(9) & Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter & ", "
'                    'Add by Amy 2015/12/04 +實績保留增減
'                    strTotalAmt(12) = strTotalAmt(12) & Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter & ", "
'
'                    lngCounter = lngCounter + 2
'                End If
'
'            End If
'            adoaccsum.MoveNext
'            strDept = Mid(adostaff.Fields("a0901").Value, 1, 2)
'        Loop
'        adoaccsum.Close
'        adostaff.MoveNext
'   Loop
'   adostaff.Close
'
'   Select Case strDept
'      Case "S1", "S2"
'         If strSum(0) <> "" Then
'            If strSum(0) <> "" Then strSum(0) = Mid(strSum(0), 1, Len(strSum(0)) - 1)
'            If strSum(1) <> "" Then strSum(1) = Mid(strSum(1), 1, Len(strSum(1)) - 1)
'            If strSum(2) <> "" Then strSum(2) = Mid(strSum(2), 1, Len(strSum(2)) - 1)
'            If strSum(3) <> "" Then strSum(3) = Mid(strSum(3), 1, Len(strSum(3)) - 1)
'            If strSum(4) <> "" Then strSum(4) = Mid(strSum(4), 1, Len(strSum(4)) - 1)
'            If strSum(5) <> "" Then strSum(5) = Mid(strSum(5), 1, Len(strSum(5)) - 1)
'            'Add by Amy 2015/02/05 +加轉發點數及減轉撥點數
'            If strSum(6) <> "" Then strSum(6) = Mid(strSum(6), 1, Len(strSum(6)) - 1)
'            If strSum(7) <> "" Then strSum(7) = Mid(strSum(7), 1, Len(strSum(7)) - 1)
'            Select Case strDept
'               Case "S1"
'                  wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(105)
'               Case "S2"
'                  wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(106)
'               Case "S3"
'                  wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(107)
'               Case "S4"
'                  wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(108)
'               Case "S9"
'                  wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(127)
'            End Select
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Formula = strSum(0)
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Formula = strSum(1)
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Formula = strSum(2)
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Formula = strSum(3)
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Formula = strSum(4)
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Formula = strSum(5)
'                'Modify by Amy 2015/02/05 +加轉發點數及減轉撥點數
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Formula = strSum(6)
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Formula = strSum(7)
'                'Modfiy by Amy 2015/12/07 合計改統一寫法,都用strSum
''                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & _
''                                                                                                                                                        "+" & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter
''                'end 2015/02/05
''                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter
''                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter
''                'Add by Amy 2015/12/04 +實績保留增減
''                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = strSum(8)
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = strSum(9)
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = strSum(10)
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = strSum(11)
'                'end 2015/12/07
'                lngCounter = lngCounter + 2
'         End If
'   End Select
'
'   If strTotalAmt(1) & strTotalAmt(2) & strTotalAmt(3) & strTotalAmt(4) & strTotalAmt(5) & strTotalAmt(6) <> "" Then
'      wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = "智權部合計:"
'      If strTotalAmt(1) <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(1), 1, Len(strTotalAmt(1)) - 2) & ")"
'      If strTotalAmt(2) <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(2), 1, Len(strTotalAmt(2)) - 2) & ")"
'      If strTotalAmt(3) <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(3), 1, Len(strTotalAmt(3)) - 2) & ")"
'      If strTotalAmt(4) <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(4), 1, Len(strTotalAmt(4)) - 2) & ")"
'      If strTotalAmt(5) <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(5), 1, Len(strTotalAmt(5)) - 2) & ")"
'      If strTotalAmt(6) <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(6), 1, Len(strTotalAmt(6)) - 2) & ")"
'      'Add by 2015/02/05 +加轉發點數及減轉撥點數
'      If strTotalAmt(10) <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(10), 1, Len(strTotalAmt(10)) - 2) & ")"
'      If strTotalAmt(11) <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(11), 1, Len(strTotalAmt(11)) - 2) & ")"
'      If strTotalAmt(7) <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(7), 1, Len(strTotalAmt(7)) - 2) & ")"
'      If strTotalAmt(8) <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(8), 1, Len(strTotalAmt(8)) - 2) & ")"
'      If strTotalAmt(9) <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(9), 1, Len(strTotalAmt(9)) - 2) & ")"
'      'Add by Amy 2015/12/04 +實績保留增減
'      If strTotalAmt(12) <> "" Then wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(12), 1, Len(strTotalAmt(12)) - 2) & ")"
'      lngCounter = lngCounter + 2
'   End If
'
'   '其他人員
'   lngCounter1 = lngCounter: ii = 1
'   'Add by Amy 2016/03/11 +抓取智權人員點數輸入資料,非智權部門
'   strSqlSP = GetPoint_SP(Val(Left(FCDate(MaskEdBox1.Text), 5)), Val(Left(FCDate(MaskEdBox2.Text), 5)), "*NS", , , , Me.Name)
'   adoFinal.Open strSqlSP, adoTaie, adOpenStatic, adLockReadOnly
'   adoFinal.MoveFirst
'
'   adostaff.CursorLocation = adUseClient
'   adostaff.Open "select * from acc090 where (substr(a0901, 1, 1) <> 'S') or a0901 = '020' order by a0901 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Do While adostaff.EOF = False
'        stSQL = "Select Distinct(r41702) as StaffNo From accrpt417 Where r41701 = '" & strUserNum & "' And r41712 = '" & adostaff.Fields("a0901").Value & "' " & strWhere & _
'                    "And substr(r41702, 1, 5) not in ('F4102', 'F4103', 'F4101') Order by r41702"
'
'        adoaccrpt417.CursorLocation = adUseClient
'        adoaccrpt417.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'          Do While adoaccrpt417.EOF = False
'                If Left("" & adoaccrpt417.Fields(0), 5) = "M0100" Then
'                    '抓取 M0100大陸P
'                    strTemp(1) = GetStartMonthCredit(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5), "P")
'                    strTemp(2) = Get4194Credit(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5), "P")
'                    strTemp(3) = GetSurplus(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5), "P")
'                    strTemp(4) = GetNotSurplus(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5), "P")
'                    strTemp(5) = GetEndMonthDebit(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5), "P")
'                    strTemp(6) = Get4194Debit(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5), "P")
'                    strM0100(1) = Val(strM0100(1)) + Val(strTemp(1))
'                    strM0100(2) = Val(strM0100(2)) + Val(strTemp(2))
'                    strM0100(3) = Val(strM0100(3)) + Val(strTemp(3))
'                    strM0100(4) = Val(strM0100(4)) + Val(strTemp(4))
'                    strM0100(5) = Val(strM0100(5)) + Val(strTemp(5))
'                    strM0100(6) = Val(strM0100(6)) + Val(strTemp(6))
'                    If strTemp(1) <> 0 Or strTemp(2) <> 0 Or strTemp(3) <> 0 Or strTemp(4) <> 0 Or strTemp(5) <> 0 Or strTemp(6) <> 0 Then
'                        wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = "M0100大陸P"
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Value = strTemp(1)
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Value = strTemp(2)
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Value = strTemp(3)
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Value = strTemp(4)
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Value = strTemp(5)
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Value = strTemp(6)
'                        'Modify by Amy 2015/02/05 +加轉發點數及減轉撥點數,改加減欄位
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Value = 0
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Value = 0
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & _
'                                                                                                                                                        "+" & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter
'                        'end 2015/02/05
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter
'                        'Add by Amy 2015/12/04 +實績保留增減
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter
'                         lngCounter = lngCounter + 1
'                    End If
'                    '抓取 M0100大陸T
'                    strTemp(1) = GetStartMonthCredit(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5), "T")
'                    strTemp(2) = Get4194Credit(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5), "T")
'                    strTemp(3) = GetSurplus(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5), "T")
'                    strTemp(4) = GetNotSurplus(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5), "T")
'                    strTemp(5) = GetEndMonthDebit(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5), "T")
'                    strTemp(6) = Get4194Debit(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5), "T")
'                    strM0100(1) = Val(strM0100(1)) + Val(strTemp(1))
'                    strM0100(2) = Val(strM0100(2)) + Val(strTemp(2))
'                    strM0100(3) = Val(strM0100(3)) + Val(strTemp(3))
'                    strM0100(4) = Val(strM0100(4)) + Val(strTemp(4))
'                    strM0100(5) = Val(strM0100(5)) + Val(strTemp(5))
'                    strM0100(6) = Val(strM0100(6)) + Val(strTemp(6))
'                    If strTemp(1) <> 0 Or strTemp(2) <> 0 Or strTemp(3) <> 0 Or strTemp(4) <> 0 Or strTemp(5) <> 0 Or strTemp(6) <> 0 Then
'                        wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = "M0100大陸T"
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Value = strTemp(1)
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Value = strTemp(2)
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Value = strTemp(3)
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Value = strTemp(4)
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Value = strTemp(5)
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Value = strTemp(6)
'                        'Modify by Amy 2015/02/05 +加轉發點數及減轉撥點數
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Value = 0
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Value = 0
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & _
'                                                                                                                                                        "+" & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter
'                        'end 2015/02/05
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter
'                         'Add by Amy 2015/12/04 +實績保留增減
'                        wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter
'                         lngCounter = lngCounter + 1
'                    End If
'                End If
'
'                strTemp(1) = GetStartMonthCredit(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5))
'                strTemp(2) = Get4194Credit(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5))
'                strTemp(3) = GetSurplus(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5))
'                strTemp(4) = GetNotSurplus(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5))
'                strTemp(5) = GetEndMonthDebit(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5))
'                strTemp(6) = Get4194Debit(Left(RTrim(adoaccrpt417.Fields("StaffNo").Value), 5))
'
'                wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = adoaccrpt417.Fields("StaffNo").Value
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Value = Val(strTemp(1)) - Val(strM0100(1))
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Value = Val(strTemp(2)) - Val(strM0100(2))
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Value = Val(strTemp(3)) - Val(strM0100(3))
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Value = Val(strTemp(4)) - Val(strM0100(4))
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Value = Val(strTemp(5)) - Val(strM0100(5))
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Value = Val(strTemp(6)) - Val(strM0100(6))
'                'Modify by Amy 2015/02/05 +加轉發點數及減轉撥點數
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Value = 0
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Value = 0
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & _
'                                                                                                                                                        "+" & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter
'                'end 2015/02/05
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter
'                'Add by Amy 2015/12/04 +實績保留增減
'                wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter
'                lngCounter = lngCounter + 1
'
'             adoaccrpt417.MoveNext
'          Loop
'          adoaccrpt417.Close
'          adostaff.MoveNext
'   Loop
'   adostaff.Close
'
'   '其他合計
'   stSQL = "Select count(Distinct rtrim(substr(r41702, 1, 5))) From accrpt417, acc090 " & _
'                "Where  r41712 = a0901 and r41701 = '" & strUserNum & "' and ((substr(a0901, 1, 1) <> 'S' and substr(a0901, 1, 1) <> 'F') or a0901 = '020') " & strWhere
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) = False And adoaccsum.Fields(0).Value <> 0 Then
'         wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(64) & ReportSum(25)
'         wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportFieldN2(1)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportFieldN2(2)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportFieldN2(3)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportFieldN2(4)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportFieldN2(5)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportFieldN2(6)) + 64) & (lngCounter - 1) & ")"
'         'Add by Amy 2015/02/25 +加轉發點數及減轉撥點數
'         wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportFieldN2(10)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportFieldN2(11)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportFieldN2(7)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportFieldN2(8)) + 64) & (lngCounter - 1) & ")"
'         wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportFieldN2(9)) + 64) & (lngCounter - 1) & ")"
'         'Add by Amy 2015/12/04 +實績保留增減
'         wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=sum(" & Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter1 & ":" & Chr(GetValue(ReportFieldN2(12)) + 64) & (lngCounter - 1) & ")"
'
'         strTotalAmt(1) = strTotalAmt(1) & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & ", "
'         strTotalAmt(2) = strTotalAmt(2) & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & ", "
'         strTotalAmt(3) = strTotalAmt(3) & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & ", "
'         strTotalAmt(4) = strTotalAmt(4) & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & ", "
'         strTotalAmt(5) = strTotalAmt(5) & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & ", "
'         strTotalAmt(6) = strTotalAmt(6) & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter & ", "
'         'Add by Amy 2015/02/05 +加轉發點數及減轉撥點數
'         strTotalAmt(10) = strTotalAmt(10) & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & ", "
'         strTotalAmt(11) = strTotalAmt(11) & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter & ", "
'         strTotalAmt(7) = strTotalAmt(7) & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & ", "
'         strTotalAmt(8) = strTotalAmt(8) & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter & ", "
'         strTotalAmt(9) = strTotalAmt(9) & Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter & ", "
'         'Add by Amy 2015/12/04 +實績保留增減
'         strTotalAmt(12) = strTotalAmt(12) & Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter & ", "
'
'         lngCounter = lngCounter + 2
'      End If
'   End If
'   adoaccsum.Close
'
'    '國內合計
'    wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(65) & ReportSum(25)
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(1), 1, Len(strTotalAmt(1)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(2), 1, Len(strTotalAmt(2)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(3), 1, Len(strTotalAmt(3)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(4), 1, Len(strTotalAmt(4)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(5), 1, Len(strTotalAmt(5)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(6), 1, Len(strTotalAmt(6)) - 2) & ")"
'    'Add by Amy 2015/02/05 +加轉發點數及減轉撥點數
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(10), 1, Len(strTotalAmt(10)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(11), 1, Len(strTotalAmt(11)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(7), 1, Len(strTotalAmt(7)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(8), 1, Len(strTotalAmt(8)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(9), 1, Len(strTotalAmt(9)) - 2) & ")"
'    'Add by Amy 2015/12/04 +實績保留增減
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmt(12), 1, Len(strTotalAmt(12)) - 2) & ")"
'
'    lngCounter = lngCounter + 2
'
'    'FCP
'    wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(67) & ReportSum(25)
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Value = GetStartMonthCredit("F4102")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Value = Get4194Credit("F4102")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Value = GetSurplus("F4102")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Value = GetNotSurplus("F4102")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Value = GetEndMonthDebit("F4102")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Value = Get4194Debit("F4102")
'    'Modify by Amy 2015/02/05 +加轉發點數及減轉撥點數,改加減欄位
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Value = 0
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Value = 0
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & _
'                                                                                                                                                        "+" & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter
'    'end 2015/02/05
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter
'    'Add by Amy 2015/12/04 +實績保留增減
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter
'
'    strTotalAmtFC(1) = strTotalAmtFC(1) & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & ", "
'    strTotalAmtFC(2) = strTotalAmtFC(2) & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & ", "
'    strTotalAmtFC(3) = strTotalAmtFC(3) & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & ", "
'    strTotalAmtFC(4) = strTotalAmtFC(4) & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & ", "
'    strTotalAmtFC(5) = strTotalAmtFC(5) & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & ", "
'    strTotalAmtFC(6) = strTotalAmtFC(6) & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter & ", "
'    'Add by 2015/02/05 +加轉發點數及減轉撥點數
'    strTotalAmtFC(10) = strTotalAmtFC(10) & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & ", "
'    strTotalAmtFC(11) = strTotalAmtFC(11) & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter & ", "
'    strTotalAmtFC(7) = strTotalAmtFC(7) & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & ", "
'    strTotalAmtFC(8) = strTotalAmtFC(8) & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter & ", "
'    strTotalAmtFC(9) = strTotalAmtFC(9) & Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter & ", "
'    'Add by Amy 2015/12/04 +實績保留增減
'    strTotalAmtFC(12) = strTotalAmtFC(12) & Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter & ", "
'    lngCounter = lngCounter + 1
'
'    'FCT
'    wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(68) & ReportSum(25)
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Value = GetStartMonthCredit("F4103")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Value = Get4194Credit("F4103")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Value = GetSurplus("F4103")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Value = GetNotSurplus("F4103")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Value = GetEndMonthDebit("F4103")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Value = Get4194Debit("F4103")
'    'Modify by 2015/02/05 +加轉發點數及減轉撥點數
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Value = 0
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Value = 0
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & _
'                                                                                                                                                        "+" & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter
'    'end 2015/02/05
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter
'    'Add by Amy 2015/12/04 +實績保留增減
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter
'
'    strTotalAmtFC(1) = strTotalAmtFC(1) & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & ", "
'    strTotalAmtFC(2) = strTotalAmtFC(2) & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & ", "
'    strTotalAmtFC(3) = strTotalAmtFC(3) & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & ", "
'    strTotalAmtFC(4) = strTotalAmtFC(4) & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & ", "
'    strTotalAmtFC(5) = strTotalAmtFC(5) & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & ", "
'    strTotalAmtFC(6) = strTotalAmtFC(6) & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter & ", "
'    'Add by 2015/02/05 +加轉發點數及減轉撥點數
'    strTotalAmtFC(10) = strTotalAmtFC(10) & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & ", "
'    strTotalAmtFC(11) = strTotalAmtFC(11) & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter & ", "
'    strTotalAmtFC(7) = strTotalAmtFC(7) & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & ", "
'    strTotalAmtFC(8) = strTotalAmtFC(8) & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter & ", "
'    strTotalAmtFC(9) = strTotalAmtFC(9) & Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter & ", "
'    'Add by Amy 2015/12/04 +實績保留增減
'    strTotalAmtFC(12) = strTotalAmtFC(12) & Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter & ", "
'    lngCounter = lngCounter + 1
'
'    'FCL
'    wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(69) & ReportSum(25)
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Value = GetStartMonthCredit("F4101")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Value = Get4194Credit("F4101")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Value = GetSurplus("F4101")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Value = GetNotSurplus("F4101")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Value = GetEndMonthDebit("F4101")
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Value = Get4194Debit("F4101")
'    'Modify by 2015/02/05 +加轉發點數及減轉撥點數
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Value = 0
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Value = 0
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & _
'                                                                                                                                                        "+" & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter
'    'end 2015/02/05
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & "+" & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter
'    'Add by Amy 2015/12/04 +實績保留增減
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=" & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & "-" & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter
'
'    strTotalAmtFC(1) = strTotalAmtFC(1) & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter & ", "
'    strTotalAmtFC(2) = strTotalAmtFC(2) & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter & ", "
'    strTotalAmtFC(3) = strTotalAmtFC(3) & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter & ", "
'    strTotalAmtFC(4) = strTotalAmtFC(4) & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter & ", "
'    strTotalAmtFC(5) = strTotalAmtFC(5) & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter & ", "
'    strTotalAmtFC(6) = strTotalAmtFC(6) & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter & ", "
'    'Add by Amy 2015/02/05 +加轉發點數及減轉撥點數
'    strTotalAmtFC(10) = strTotalAmtFC(10) & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter & ", "
'    strTotalAmtFC(11) = strTotalAmtFC(11) & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter & ", "
'    strTotalAmtFC(7) = strTotalAmtFC(7) & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter & ", "
'    strTotalAmtFC(8) = strTotalAmtFC(8) & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter & ", "
'    strTotalAmtFC(9) = strTotalAmtFC(9) & Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter & ", "
'    'Add by Amy 2015/12/04 +實績保留增減
'    strTotalAmtFC(12) = strTotalAmtFC(12) & Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter & ", "
'    lngCounter = lngCounter + 1
'
'    '國外合計
'    wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(70) & ReportSum(25)
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmtFC(1), 1, Len(strTotalAmtFC(1)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmtFC(2), 1, Len(strTotalAmtFC(2)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmtFC(3), 1, Len(strTotalAmtFC(3)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmtFC(4), 1, Len(strTotalAmtFC(4)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmtFC(5), 1, Len(strTotalAmtFC(5)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmtFC(6), 1, Len(strTotalAmtFC(6)) - 2) & ")"
'    'Add by Amy 2015/02/05 +加轉發點數及減轉撥點數
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmtFC(10), 1, Len(strTotalAmtFC(10)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmtFC(11), 1, Len(strTotalAmtFC(11)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmtFC(7), 1, Len(strTotalAmtFC(7)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmtFC(8), 1, Len(strTotalAmtFC(8)) - 2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmtFC(9), 1, Len(strTotalAmtFC(9)) - 2) & ")"
'    'Add by Amy 2015/12/04 +實績保留增減
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=sum(" & Mid(strTotalAmtFC(12), 1, Len(strTotalAmtFC(12)) - 2) & ")"
'
'    strTotalAmtFC(1) = "," & Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter
'    strTotalAmtFC(2) = "," & Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter
'    strTotalAmtFC(3) = "," & Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter
'    strTotalAmtFC(4) = "," & Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter
'    strTotalAmtFC(5) = "," & Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter
'    strTotalAmtFC(6) = "," & Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter
'    'Add by Amy 2015/02/05 +加轉發點數及減轉撥點數
'    strTotalAmtFC(10) = "," & Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter
'    strTotalAmtFC(11) = "," & Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter
'    strTotalAmtFC(7) = "," & Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter
'    strTotalAmtFC(8) = "," & Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter
'    strTotalAmtFC(9) = "," & Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter
'    'Add by Amy 2015/12/04 +實績保留增減
'    strTotalAmtFC(12) = "," & Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter
'    lngCounter = lngCounter + 2
'
'    '總所合計
'    wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & lngCounter).Value = ReportSum(66) & ReportSum(25)
'    '2015/1/29 MODIFY BY SONIA
'    'wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(1) & Mid(strTotalAmt(1), 1, Len(strTotalAmt(1)) - 2) & strTotalAmtFC(1) & ")"
'    'wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(2) & Mid(strTotalAmt(2), 1, Len(strTotalAmt(2)) - 2) & strTotalAmtFC(2) & ")"
'    'wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(3) & Mid(strTotalAmt(3), 1, Len(strTotalAmt(3)) - 2) & strTotalAmtFC(3) & ")"
'    'wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(4) & Mid(strTotalAmt(4), 1, Len(strTotalAmt(4)) - 2) & strTotalAmtFC(4) & ")"
'    'wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(5) & Mid(strTotalAmt(5), 1, Len(strTotalAmt(5)) - 2) & strTotalAmtFC(5) & ")"
'    'wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(6) & Mid(strTotalAmt(6), 1, Len(strTotalAmt(6)) - 2) & strTotalAmtFC(6) & ")"
'    'wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(7) & Mid(strTotalAmt(7), 1, Len(strTotalAmt(7)) - 2) & strTotalAmtFC(7) & ")"
'    'wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(8) & Mid(strTotalAmt(8), 1, Len(strTotalAmt(8)) - 2) & strTotalAmtFC(8) & ")"
'    'wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(9) & Mid(strTotalAmt(9), 1, Len(strTotalAmt(9)) - 2) & strTotalAmtFC(9) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(1) & strTotalAmtFC(1) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(2)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(2) & strTotalAmtFC(2) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(3)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(3) & strTotalAmtFC(3) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(4)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(4) & strTotalAmtFC(4) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(5)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(5) & strTotalAmtFC(5) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(6)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(6) & strTotalAmtFC(6) & ")"
'    'Add by 2015/02/05 +加轉發點數及減轉撥點數
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(10)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(10) & strTotalAmtFC(10) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(11)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(11) & strTotalAmtFC(11) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(7)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(7) & strTotalAmtFC(7) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(8)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(8) & strTotalAmtFC(8) & ")"
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(9)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(9) & strTotalAmtFC(9) & ")"
'    '2015/1/29 END
'    'Add by Amy 2015/12/04 +實績保留增減
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(12)) + 64) & lngCounter).Formula = "=sum(" & strTotalAmt(12) & strTotalAmtFC(12) & ")"
'
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & intTitleRow + 1 & ":" & Chr(UBound(strFieldN) + 64) & lngCounter).Select
'    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + 64) & intTitleRow + 1 & ":" & Chr(UBound(strFieldN) + 64) & lngCounter).NumberFormatLocal = "#,##0.00_ "
'
'    '框線
'    wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + 64) & intTitleRow & ":" & Chr(UBound(strFieldN) + 64) & lngCounter).Select
'    xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    xlsSalesPoint.Selection.Font.Size = 8 'Add by Amy 2015/02/05
'    'Modify by Amy 2014/06/11 +判斷若版本2007以上改變存格式
'   If Val(xlsSalesPoint.Version) < 12 Then
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & "智權點數實績與結餘分析表" & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
'   Else
'        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & "智權點數實績與結餘分析表" & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
'   End If
'   xlsSalesPoint.Workbooks.Close
'   xlsSalesPoint.Quit
'   StatusClear
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'Add by Amy 2020/03/31
   Dim bolCancel As Boolean
   
   If Trim(CboCmp) <> MsgText(601) Then
      CboCmp_Validate (bolCancel)
      If bolCancel = False Then
        FormCheck = True
        Exit Function
      End If
   End If
   'end 2020/03/31
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   'Add by Amy 2022/02/18 輸錯會Error
   Else
      If IsDate(ChangeTStringToWDateString(Replace(Me.MaskEdBox1.Text, "/", ""))) = True Then
         FormCheck = True
         Exit Function
      End If
   End If
   
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   'Add by Amy 2022/02/18 輸錯會Error
   Else
      If IsDate(ChangeTStringToWDateString(Replace(Me.MaskEdBox2.Text, "/", ""))) = True Then
         FormCheck = True
         Exit Function
      End If
   End If
   FormCheck = False
End Function

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
' 列印明細
'104/06/11改設印A4(原程式以美國標準印35行,也可印A4)
'*************************************************
Public Sub PrintDetail()
   intLength = 0
   intPage = 0
   douAmount = 0
   strName = ""
   adoquery.CursorLocation = adUseClient
   '2011/9/2 modify by sonia 因加員工編號S29,故取5碼去空白讀員工檔
   'adoquery.Open "select distinct r41702, st01, st15, '1' as Sort from accrpt417, staff where substr(r41702, 1, 5) = st01 and r41701 = '" & strUserNum & "' and substr(st03, 1, 1) = 'S' union " & _
                 "select distinct r41702, st01, st15, '2' as Sort from accrpt417, staff where substr(r41702, 1, 5) = st01 and r41701 = '" & strUserNum & "' and substr(st03, 1, 1) <> 'S' and st01 <> 'F4101' and st02 <> 'F4102' and st03 <> 'F4103' union " & _
                 "select distinct r41702, st01, st15, '3' as Sort from accrpt417, staff where substr(r41702, 1, 5) = st01 and r41701 = '" & strUserNum & "' and (st01 = 'F4101' or st01 = 'F4102' or st01 = 'F4103') order by Sort asc, st15 asc, st01 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2021/01/21 +F4104~07
   adoquery.Open "select distinct r41702, st01, st15, '1' as Sort from accrpt417, staff where rtrim(substr(r41702, 1, 5)) = st01 and r41701 = '" & strUserNum & "' and substr(st03, 1, 1) = 'S' union " & _
                 "select distinct r41702, st01, st15, '2' as Sort from accrpt417, staff where rtrim(substr(r41702, 1, 5)) = st01 and r41701 = '" & strUserNum & "' and substr(st03, 1, 1) <> 'S' and st01 <> 'F4101' and st02 <> 'F4102' and st03 <> 'F4103' and st03 <> 'F4104' and st03 <> 'F4105' and st03 <> 'F4106' and st03 <> 'F4107' union " & _
                 "select distinct r41702, st01, st15, '3' as Sort from accrpt417, staff where rtrim(substr(r41702, 1, 5)) = st01 and r41701 = '" & strUserNum & "' and (st01 = 'F4101' or st01 = 'F4102' or st01 = 'F4103' or st01 = 'F4104' or st01 = 'F4105' or st01 = 'F4106' or st01 = 'F4107') order by Sort asc, st15 asc, st01 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount = 0 Then
      adoquery.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   'Add by Amy 2015/06/11
   PUB_RestorePrinter Combo7
   Printer.PaperSize = 9
   'end 2015/06/11
   Printer.FontSize = 10
   Do While adoquery.EOF = False
      If strName <> adoquery.Fields("r41702").Value Then
         'Modify by Amy 2015/06/11 合計有可能等於0,沒換行導致重疊列印
         'If douAmount <> 0 Then
         If intPage > 0 Then
            PrintSum
            douAmount = 0
            Printer.NewPage
         End If
         intCounter = 0
         intPage = intPage + 1
         PrintHead
         strName = adoquery.Fields("r41702").Value
      End If
      If adoaccrpt417.State <> adStateClosed Then adoaccrpt417.Close
      adoaccrpt417.CursorLocation = adUseClient
      adoaccrpt417.Open "select * from accrpt417 where r41701 = '" & strUserNum & "' and r41702 = '" & strName & "' order by r41703 asc, r41704 asc", adoTaie, adOpenStatic, adLockReadOnly
      Do While adoaccrpt417.EOF = False
         If intCounter > 40 Then
            intCounter = 0
            Printer.NewPage
            intPage = intPage + 1
            PrintHead
         End If
         '傳票號碼
         Printer.CurrentX = 700
         Printer.CurrentY = 3000 + intCounter * 300
         If IsNull(adoaccrpt417.Fields("r41703").Value) = False Then
            Printer.Print adoaccrpt417.Fields("r41703").Value
         Else
            Printer.Print ""
         End If
         '對沖代號(客)
         Printer.CurrentX = 2200
         Printer.CurrentY = 3000 + intCounter * 300
         If IsNull(adoaccrpt417.Fields("r41704").Value) = False Then
            Printer.Print MidB(adoaccrpt417.Fields("r41704").Value, 1, 26)
         Else
            Printer.Print ""
         End If
         '對沖代號(案號)
         Printer.CurrentX = 4400
         Printer.CurrentY = 3000 + intCounter * 300
         If IsNull(adoaccrpt417.Fields("r41705").Value) = False Then
            Printer.Print adoaccrpt417.Fields("r41705").Value
         Else
            Printer.Print ""
         End If
         '摘要
         Printer.CurrentX = 6400
         Printer.CurrentY = 3000 + intCounter * 300
         If IsNull(adoaccrpt417.Fields("r41706").Value) = False Then
            'Modify By Cheng 2003/06/05
'            Printer.Print MidB(adoaccrpt417.Fields("r41706").Value, 1, 30)
            Printer.Print MidB(adoaccrpt417.Fields("r41706").Value, 1, 32)
         Else
            Printer.Print ""
         End If
        '點數
         If IsNull(adoaccrpt417.Fields("r41707").Value) = False Then
            'Modify By Cheng 2003/06/05
'            strAmount = Format(adoaccrpt417.Fields("r41707").Value, DDollar)
            strAmount = Format(adoaccrpt417.Fields("r41707").Value, "#,##0.00")
            If strAmount = "" Then
                'Modify By Cheng 2003/06/05
'               strAmount = "0"
               strAmount = "0.00"
            End If
            intLength = Printer.TextWidth(strAmount)
'            Printer.CurrentX = 9000 - intLength
            Printer.CurrentX = 10900 - intLength
            Printer.CurrentY = 3000 + intCounter * 300
            Printer.Print strAmount
            douAmount = douAmount + Val(adoaccrpt417.Fields("r41707").Value)
         Else
            'Modify By Cheng 2003/06/05
'            intLength = Printer.TextWidth("0")
            intLength = Printer.TextWidth("0.00")
'            Printer.CurrentX = 9000 - intLength
            Printer.CurrentX = 10900 - intLength
            Printer.CurrentY = 3000 + intCounter * 300
            'Modify By Cheng 2003/06/05
'            Printer.Print "0"
            Printer.Print "0.00"
         End If
         intCounter = intCounter + 1
         adoaccrpt417.MoveNext
      Loop
      adoaccrpt417.Close
      adoquery.MoveNext
   Loop
   PrintSum
   adoquery.Close
   Printer.EndDoc
   PUB_RestorePrinter strPrinter 'Add by Amy 2015/06/11
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead()
   Printer.CurrentX = 4200 + 500
   Printer.CurrentY = 500
   Printer.Print ReportTitle(417)
   Printer.CurrentX = 4100 + 500
   Printer.CurrentY = 1000
   Printer.Print "統計日期: " & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
   'Add By Sindy 2014/1/22
   Printer.CurrentX = 700
   Printer.CurrentY = 1300
   Printer.Print "公司別: " & GetAccReportCmpN(CboCmp, , True) 'Modify by Amy 2020/04/16
   '2014/1/22 END
   Printer.CurrentX = 5000 + 500
   Printer.CurrentY = 1300
   Printer.Print Combo3
   Printer.CurrentX = 700
   Printer.CurrentY = 1600
   Printer.Print "列印人員: " & StaffQuery(strUserNum)
   Printer.CurrentX = 7700 + 1000
   Printer.CurrentY = 1600
   Printer.Print "列印日期: " & CFDate(ACDate(ServerDate))
   Printer.CurrentX = 700
   Printer.CurrentY = 1900
   Printer.Print "智權人員: " & adoquery.Fields("r41702").Value
   Printer.CurrentX = 7700 + 1000
   Printer.CurrentY = 1900
   Printer.Print "頁次:     " & intPage
   Printer.CurrentX = 700
   Printer.CurrentY = 2500
   Printer.Print "傳票號碼"
   Printer.CurrentX = 2200
   Printer.CurrentY = 2500
   Printer.Print "對沖代號(客)"
   Printer.CurrentX = 4400
   Printer.CurrentY = 2500
   Printer.Print "對沖代號(案號)"
   Printer.CurrentX = 6400
   Printer.CurrentY = 2500
   Printer.Print "摘    要"
   Printer.CurrentX = 8900 + 1000
   Printer.CurrentY = 2500
   Printer.Print "點    數"
'   Printer.Line (0, 2900)-(9000, 2900)
   Printer.Line (600, 2900)-(10900, 2900)
End Sub

'*************************************************
' 合計位置
'
'*************************************************
Private Sub PrintSum()
'   Printer.Line (8000, 3000 + intCounter * 300 - 100)-(9000, 3000 + intCounter * 300 - 100)
   Printer.Line (9700, 3000 + intCounter * 300 - 100)-(10900, 3000 + intCounter * 300 - 100)
   Printer.CurrentX = 6200
   Printer.CurrentY = 3000 + intCounter * 300
   Printer.Print "智權人員合計: "
    'Modify By Cheng 2003/06/05
'   strAmount = Format(douAmount, DDollar)
   strAmount = Format(douAmount, "#,##0.00")
   If strAmount = "" Then
      strAmount = "0"
   End If
   intLength = Printer.TextWidth(strAmount)
'   Printer.CurrentX = 9000 - intLength
   Printer.CurrentX = 10900 - intLength
   Printer.CurrentY = 3000 + intCounter * 300
   Printer.Print strAmount
'   Printer.Line (8000, 3000 + intCounter * 300 + 400)-(9000, 3000 + intCounter * 300 + 400)
'   Printer.Line (8000, 3000 + intCounter * 300 + 430)-(9000, 3000 + intCounter * 300 + 430)
   Printer.Line (9700, 3000 + intCounter * 300 + 400)-(10900, 3000 + intCounter * 300 + 400)
   Printer.Line (9700, 3000 + intCounter * 300 + 430)-(10900, 3000 + intCounter * 300 + 430)
End Sub

'Mark by Amy 2020/03/31 改下拉
''Add By Sindy 2014/1/22
'Private Sub Text4_GotFocus()
'   TextInverse Text4
'End Sub
'Private Sub Text4_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub
''2014/1/22 END

'Add by Amy 2015/01/12
'Modify by Amy 2022/02/18 +IsFieldN(表1欄位)
Private Function GetValue(pFieldN As String, Optional ByVal IsFieldN As Boolean = True) As Integer
   Dim jj As Integer
    
    'Modify by Amy 2022/02/18 +if IsFieldN=True及strFieldN2(表2欄位)
    If IsFieldN = True Then
        For jj = 1 To UBound(strFieldN)
           If UCase(strFieldN(jj)) = UCase(pFieldN) Then
              GetValue = jj
              Exit For
           End If
        Next jj
    Else
        For jj = LBound(strFieldN2) To UBound(strFieldN2)
           If UCase(strFieldN2(jj)) = UCase(pFieldN) Then
              GetValue = jj
              Exit For
           End If
        Next jj
    End If
End Function

'Modify by Amy 2016/03/11 改寫至BasQuery
'依員工編號抓取畫面傳票止日當月4191及4192之「借」方總和
Private Function GetEndMonthDebit(pStaffNo As String, Optional ByVal SysKind As String = "") As Double
'    Dim adoquery As New ADODB.Recordset
'    Dim stQuery As String, intQ As Integer
'
'    If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'        stQuery = stQuery & " and R41711 >= " & Val(Left(FCDate(MaskEdBox2.Text), 5) & "01") & ""
'        stQuery = stQuery & " and R41711 <= " & Val(Left(FCDate(MaskEdBox2.Text), 5) & "31") & ""
'    End If
'
'    If Trim(Text4) <> MsgText(601) Then
'        stQuery = stQuery & " and R41713 = '" & IIf(Text4 = "2", "J", "1") & "'"
'    End If
'
'    If pStaffNo = "M0100" Then
'        'Modify by Amy 2015/02/06 +抓服務業務
'        If SysKind = "P" Then
'            stQuery = stQuery & " And ((R41705 like 'P-%' And Exists (Select * From Patent,Fagent,Customer " & _
'                                "Where pa01='P' and pa02=substr(R41705,3,6) and pa03=substr(R41705,10,1) and pa04=substr(R41705,12) " & _
'                                "and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and nvl(fa10,cu10)>'009')) " & _
'                                "OR (R41705 like 'PS-%' And Exists (Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
'                                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
'                                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) )" & _
'                                "Having Sum(Nvl(R41709,0)) >0 "
'        ElseIf SysKind = "T" Then
'            stQuery = stQuery & " And ((R41705 like 'T-%' And Exists (Select * From Trademark,Fagent,Customer " & _
'                                "Where tm01='T' And tm02=substr(R41705,3,6) And tm03=substr(R41705,10,1) And tm04=substr(R41705,12) " & _
'                                "And fa01(+)=substr(tm44,1,8) And fa02(+)=substr(tm44,9) And cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) and nvl(fa10,cu10)>'009')) " & _
'                                "OR (R41705 like 'T%' And Exists (Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
'                                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
'                                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) ) " & _
'                                "Having Sum(Nvl(R41709,0)) >0 "
'        End If
'        'end 2015/02/06
'    End If
'
'    stQuery = "Select Sum(Nvl(R41709,0)) as EndDebit From accrpt417 " & _
'                 "Where R41701='" & strUserNum & "' And SubStr(R41702,1,5)='" & pStaffNo & "' And (R41708 = '4191' Or R41708='4192') " & stQuery
'    intQ = 1
'    Set adoquery = ClsLawReadRstMsg(intQ, stQuery)
'    If intQ = 1 Then
'        GetEndMonthDebit = Val("" & adoquery.Fields("EndDebit").Value)
'    Else
'        GetEndMonthDebit = 0
'    End If
'    adoquery.Close
End Function

'依員工編號抓取畫面傳票起日當月4191及4192之「貸」方總和
Private Function GetStartMonthCredit(pStaffNo As String, Optional ByVal SysKind As String = "") As Double
'    Dim adoquery As New ADODB.Recordset
'    Dim stQuery As String, intQ As Integer
'
'    If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'        stQuery = stQuery & " And R41711 >= " & Val(Left(FCDate(MaskEdBox1.Text), 5) & "01") & ""
'        stQuery = stQuery & " And R41711 <= " & Val(Left(FCDate(MaskEdBox1.Text), 5) & "31") & ""
'    End If
'
'    If Trim(Text4) <> MsgText(601) Then
'        stQuery = stQuery & " And R41713 = '" & IIf(Text4 = "2", "J", "1") & "'"
'    End If
'
'    If pStaffNo = "M0100" Then
'        'Modify by Amy 2015/02/05 +服務業務
'        If SysKind = "P" Then
'          stQuery = stQuery & " And ((R41705 like 'P-%' And Exists (Select * From Patent,Fagent,Customer " & _
'                                "Where pa01='P' and pa02=substr(R41705,3,6) and pa03=substr(R41705,10,1) and pa04=substr(R41705,12) " & _
'                                "and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and nvl(fa10,cu10)>'009')) " & _
'                                "OR (R41705 like 'PS-%' And Exists (Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
'                                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
'                                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) )" & _
'                                "Having Sum(Nvl(R41710,0)) >0 "
'        ElseIf SysKind = "T" Then
'            stQuery = stQuery & " And ((R41705 like 'T-%' And Exists (Select * From Trademark,Fagent,Customer " & _
'                                "Where tm01='T' And tm02=substr(R41705,3,6) And tm03=substr(R41705,10,1) And tm04=substr(R41705,12) " & _
'                                "And fa01(+)=substr(tm44,1,8) And fa02(+)=substr(tm44,9) And cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) and nvl(fa10,cu10)>'009')) " & _
'                                "OR (R41705 like 'T%' And Exists (Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
'                                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
'                                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) ) " & _
'                                "Having Sum(Nvl(R41710,0)) >0 "
'
'        End If
'        'end 2015/02/05
'    End If
'    stQuery = "Select sum(Nvl(R41710,0)) as StartCredit From accrpt417 " & _
'                 "Where R41701='" & strUserNum & "' And SubStr(R41702,1,5)='" & pStaffNo & "' And (R41708 = '4191' Or R41708='4192') " & stQuery
'    intQ = 1
'    Set adoquery = ClsLawReadRstMsg(intQ, stQuery)
'    If intQ = 1 Then
'        GetStartMonthCredit = Val("" & adoquery.Fields("StartCredit").Value)
'    Else
'        GetStartMonthCredit = 0
'    End If
'    adoquery.Close
End Function
'Add by Amy 2015/01/16
'依員工編號抓取畫面傳票起日當月4194之「借」方總和
Private Function Get4194Debit(pStaffNo As String, Optional ByVal SysKind As String = "") As Double
    Dim adoquery As New ADODB.Recordset
    Dim stQuery As String, intQ As Integer
    Dim strCmp As String 'Add by Amy 2020/03/31
    
    If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
        stQuery = stQuery & " And R41711 >= " & Val(Left(FCDate(MaskEdBox1.Text), 5) & "01") & ""
        stQuery = stQuery & " And R41711 <= " & Val(Left(FCDate(MaskEdBox1.Text), 5) & "31") & ""
    End If
    
    'Modify by Amy 2020/03/31 改下拉 原:Text4
    'If Trim(Text4) <> MsgText(601) Then
    If Trim(CboCmp) <> MsgText(601) Then
      strCmp = CboCmp
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
        'stQuery = stQuery & " And R41713 = '" & IIf(Text4 = "2", "J", "1") & "'"
        'Modify by Amy 2020/04/16 +組合公司
        If InStr(strCmp, "+") > 0 Then
            stQuery = stQuery & " And R41713 In  ('" & Replace(strCmp, "+", "','") & "')"
        Else
            stQuery = stQuery & " And R41713 = '" & strCmp & "'"
        End If
    End If
    'end 2020/03/31
    
    If pStaffNo = "M0100" Then
        'Modify by Amy 2015/02/06 +抓服務業務
        If SysKind = "P" Then
            stQuery = stQuery & " And ((R41705 like 'P-%' And Exists (Select * From Patent,Fagent,Customer " & _
                                "Where pa01='P' and pa02=substr(R41705,3,6) and pa03=substr(R41705,10,1) and pa04=substr(R41705,12) " & _
                                "and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and nvl(fa10,cu10)>'009')) " & _
                                "OR (R41705 like 'PS-%' And Exists (Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
                                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
                                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) )" & _
                                "Having Sum(Nvl(R41709,0)) >0 "
        ElseIf SysKind = "T" Then
            stQuery = stQuery & " And ((R41705 like 'T-%' And Exists (Select * From Trademark,Fagent,Customer " & _
                                "Where tm01='T' And tm02=substr(R41705,3,6) And tm03=substr(R41705,10,1) And tm04=substr(R41705,12) " & _
                                "And fa01(+)=substr(tm44,1,8) And fa02(+)=substr(tm44,9) And cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) and nvl(fa10,cu10)>'009')) " & _
                                "OR (R41705 like 'T%' And Exists (Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
                                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
                                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) ) " & _
                                "Having Sum(Nvl(R41709,0)) >0 "
        End If
        'end 2015/02/06
    End If
    stQuery = "Select sum(Nvl(R41709,0)) as Start4194 From accrpt417 " & _
                 "Where R41701='" & strUserNum & "' And SubStr(R41702,1,5)='" & pStaffNo & "' And R41708 = '4194' " & stQuery
    intQ = 1
    Set adoquery = ClsLawReadRstMsg(intQ, stQuery)
    If intQ = 1 Then
        Get4194Debit = Val("" & adoquery.Fields("Start4194").Value)
    Else
        Get4194Debit = 0
    End If
    adoquery.Close
End Function

'依員工編號抓取畫面傳票止日當月4194之「貸」方總和
Private Function Get4194Credit(pStaffNo As String, Optional ByVal SysKind As String = "") As Double
'    Dim adoquery As New ADODB.Recordset
'    Dim stQuery As String, intQ As Integer
'
'    If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'        stQuery = stQuery & " and R41711 >= " & Val(Left(FCDate(MaskEdBox2.Text), 5) & "01") & ""
'        stQuery = stQuery & " and R41711 <= " & Val(Left(FCDate(MaskEdBox2.Text), 5) & "31") & ""
'    End If
'
'    If Trim(Text4) <> MsgText(601) Then
'        stQuery = stQuery & " and R41713= '" & IIf(Text4 = "2", "J", "1") & "'"
'    End If
'
'    If pStaffNo = "M0100" Then
'        'Modify by Amy 2015/02/06 +抓服務業務
'        If SysKind = "P" Then
'            stQuery = stQuery & " And ((R41705 like 'P-%' And Exists (Select * From Patent,Fagent,Customer " & _
'                                "Where pa01='P' and pa02=substr(R41705,3,6) and pa03=substr(R41705,10,1) and pa04=substr(R41705,12) " & _
'                                "and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and nvl(fa10,cu10)>'009')) " & _
'                                "OR (R41705 like 'PS-%' And Exists (Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
'                                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
'                                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) )" & _
'                                "Having Sum(Nvl(R41710,0)) >0 "
'        ElseIf SysKind = "T" Then
'            stQuery = stQuery & " And ((R41705 like 'T-%' And Exists (Select * From Trademark,Fagent,Customer " & _
'                                "Where tm01='T' And tm02=substr(R41705,3,6) And tm03=substr(R41705,10,1) And tm04=substr(R41705,12) " & _
'                                "And fa01(+)=substr(tm44,1,8) And fa02(+)=substr(tm44,9) And cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) and nvl(fa10,cu10)>'009')) " & _
'                                "OR (R41705 like 'T%' And Exists (Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
'                                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
'                                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) ) " & _
'                                "Having Sum(Nvl(R41710,0)) >0 "
'        End If
'        'end 2015/02/06
'    End If
'    stQuery = "Select sum(Nvl(R41710,0)) as End4194 From accrpt417 " & _
'                "Where R41701='" & strUserNum & "' And SubStr(R41702,1,5)='" & pStaffNo & "' And R41708 = '4194' " & stQuery
'    intQ = 1
'    Set adoquery = ClsLawReadRstMsg(intQ, stQuery)
'    If intQ = 1 Then
'        Get4194Credit = Val("" & adoquery.Fields("End4194").Value)
'    Else
'        Get4194Credit = 0
'    End If
'    adoquery.Close
End Function

'依員工編號非結餘且會計科目非4191,4192,4194總和
Private Function GetSurplus(pStaffNo As String, Optional ByVal SysKind As String = "") As Double
'    Dim adoquery As New ADODB.Recordset
'    Dim stQuery As String, intQ As Integer
'
'    If Trim(Text4) <> MsgText(601) Then
'        stQuery = stQuery & " and R41713 = '" & IIf(Text4 = "2", "J", "1") & "'"
'    End If
'
'    If pStaffNo = "M0100" Then
'        'Modify by Amy 2015/02/06 +抓服務業務
'        If SysKind = "P" Then
'            stQuery = stQuery & " And ((R41705 like 'P-%' And Exists (Select * From Patent,Fagent,Customer " & _
'                                "Where pa01='P' and pa02=substr(R41705,3,6) and pa03=substr(R41705,10,1) and pa04=substr(R41705,12) " & _
'                                "and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and nvl(fa10,cu10)>'009')) " & _
'                                "OR (R41705 like 'PS-%' And Exists(Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
'                                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
'                                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) )" & _
'                                "Having Sum(Nvl(R41707,0)) >0 "
'        ElseIf SysKind = "T" Then
'            stQuery = stQuery & " And ((R41705 like 'T-%' And Exists (Select * From Trademark,Fagent,Customer " & _
'                                "Where tm01='T' And tm02=substr(R41705,3,6) And tm03=substr(R41705,10,1) And tm04=substr(R41705,12) " & _
'                                "And fa01(+)=substr(tm44,1,8) And fa02(+)=substr(tm44,9) And cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) and nvl(fa10,cu10)>'009')) " & _
'                                "OR (R41705 like 'T%' And Exists(Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
'                                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
'                                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) ) " & _
'                                "Having Sum(Nvl(R41707,0)) >0 "
'        End If
'        'end 2015/02/06
'    End If
'    'Modify by Amy 2015/04/24原抓摘要是否有結餘(R41706 IS NULL OR InStr(R41706,'結餘')=0) 改抓對沖其他是否有結餘
'    stQuery = "Select sum(Nvl(R41707,0)) From accrpt417 Where R41701='" & strUserNum & "' And SubStr(R41702,1,5)='" & pStaffNo & "' " & _
'                " And (R41714 IS NULL OR InStr(R41714||' ','結餘')=0) And R41708 <> '4191' And R41708 <> '4192' And R41708 <> '4194' " & stQuery
'    intQ = 1
'    Set adoquery = ClsLawReadRstMsg(intQ, stQuery)
'    If intQ = 1 Then
'        GetSurplus = Val("" & adoquery.Fields(0).Value)
'    Else
'        GetSurplus = 0
'    End If
'    adoquery.Close
End Function

'依員工編號結餘且會計科目非4191,4192,4194總和
Private Function GetNotSurplus(pStaffNo As String, Optional ByVal SysKind As String = "") As Double
'    Dim adoquery As New ADODB.Recordset
'    Dim stQuery As String, intQ As Integer
'
'    If Trim(Text4) <> MsgText(601) Then
'        stQuery = stQuery & " and R41713 = '" & IIf(Text4 = "2", "J", "1") & "'"
'    End If
'
'    If pStaffNo = "M0100" Then
'        'Modify by Amy 2015/02/06 +抓服務業務
'        If SysKind = "P" Then
'            stQuery = stQuery & " And ((R41705 like 'P-%' And Exists (Select * From Patent,Fagent,Customer " & _
'                                "Where pa01='P' and pa02=substr(R41705,3,6) and pa03=substr(R41705,10,1) and pa04=substr(R41705,12) " & _
'                                "and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and nvl(fa10,cu10)>'009')) " & _
'                                "OR (R41705 like 'PS-%' And Exists (Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
'                                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
'                                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) )" & _
'                                "Having Sum(Nvl(R41707,0)) >0 "
'        ElseIf SysKind = "T" Then
'            stQuery = stQuery & " And ((R41705 like 'T-%' And Exists (Select * From Trademark,Fagent,Customer " & _
'                                "Where tm01='T' And tm02=substr(R41705,3,6) And tm03=substr(R41705,10,1) And tm04=substr(R41705,12) " & _
'                                "And fa01(+)=substr(tm44,1,8) And fa02(+)=substr(tm44,9) And cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9) and nvl(fa10,cu10)>'009')) " & _
'                                "OR (R41705 like 'T%' And Exists (Select * From ServicePractice,Fagent,Customer Where sp01=substr(R41705, 1, length(R41705)-12) " & _
'                                "And sp02=substr(R41705, length(R41705)-10,6) And sp03=substr(R41705, length(R41705)-3,1) And sp04=substr(R41705, length(R41705)-1,2) " & _
'                                "And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9) And cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9) and nvl(fa10,cu10)>'009')) ) " & _
'                                "Having Sum(Nvl(R41707,0)) >0 "
'        End If
'        'end 2015/02/06
'    End If
'    'Modify by Amy 2015/04/24原抓摘要是否有結餘(InStr(R41706,'結餘')>0)  改抓對沖其他是否有結餘
'    stQuery = "Select sum(Nvl(R41707,0)) From accrpt417 Where R41701='" & strUserNum & "' And SubStr(R41702,1,5)='" & pStaffNo & "' " & _
'                 " And InStr(R41714||' ','結餘')>0 And R41708<>'4191' And R41708 <> '4192' And R41708 <> '4194' " & stQuery
'    intQ = 1
'    Set adoquery = ClsLawReadRstMsg(intQ, stQuery)
'    If intQ = 1 Then
'        GetNotSurplus = Val("" & adoquery.Fields(0).Value)
'    Else
'        GetNotSurplus = 0
'    End If
'    adoquery.Close
End Function
'end 2016/03/11

Private Function ReportFieldN2(intField As Integer) As String
    '欄位需照順序放
    Select Case intField
        Case 1
            ReportFieldN2 = "期初實績保留"
        Case 2
            ReportFieldN2 = "期初結餘保留"
        Case 3
            ReportFieldN2 = "當月實績點數" 'Modify by Amy 2015/12/04 +當月
        Case 4
            ReportFieldN2 = "當月結餘點數" 'Modify by Amy 2015/12/04 +當月
        Case 5
            ReportFieldN2 = "期末實績保留"
        Case 6
            ReportFieldN2 = "期末結餘保留"
        Case 7
            ReportFieldN2 = "實績轉撥點數"  'Modify by Amy 2016/03/11
        Case 8
            ReportFieldN2 = "結餘轉撥點數" 'Modify by Amy 2016/03/11
        Case 9
            ReportFieldN2 = "報出實績點數"
        Case 10
            ReportFieldN2 = "報出結餘點數"
        Case 11
            ReportFieldN2 = "報出點數"
        Case 12
            ReportFieldN2 = "實績保留增減" 'Add by Amy 2015/12/04
        'Add by Amy 2015/12/04
        Case 13
            ReportFieldN2 = "目標"
        Case 14
            ReportFieldN2 = "達成率"
        Case 15
            ReportFieldN2 = "實績達成率"
    Case Else
    End Select
End Function

'Modify by Amy 2016/05/06 寫入暫存-因新人有目標但傳票沒資料人員抓不到
'*************************************************
'  轉成Excel檔案
'  智權人員實績與結餘分析表
'*************************************************
Private Sub ExcelSaveNew2()
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt417 As New Worksheet
Dim lngCounter As Long, lngCounter1 As Long '目前位置/起始位置
Dim stSQL As String, strWhere As String, strDept As String, strSqlF As String, strCmp As String
Dim strTemp(1 To 6) As String, strValue(4) As String
Dim intTitleRow As Integer '表頭列數
Dim bolData As Boolean '是否有資料
Dim intE As Integer, intSP As Integer
Dim strFileName As String 'Add by Amy 2017/06/07 改檔名-瑞婷
'Add by Amy 2018/05/17
Dim intLCmpStartRow As Long 'Add by Amy 2020/04/22 L部門資料起始列
Dim bolSumF As Boolean, strSumForeign As String 'Add by Amy 2021/01/19 國外部要加總/合計位置(For F4104~F4107 不連續加總)
Dim strRID As String 'Add by Amy 2022/02/18
Dim bolM0100 As Boolean 'Add by Amy 2022/08/08 有M0100
    
On Error GoTo ErrHand

    Call ClearVar
    strDept = ""
    If Text1 <> MsgText(602) Then Exit Sub
    'Modify by Amy 2020/03/31 原:Text4
    If Trim(CboCmp) <> MsgText(601) Then
        'strCmp = IIf(Text4 = "2", "J", "1")
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    
    'Add by Amy 2022/02/18 刪除 位置暫存檔
    stSQL = "Delete From RptXlsLink Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' "
    cnnConnection.Execute stSQL
    '將資料寫入暫存檔
    stSQL = "Delete From Accrpt44j0_2 Where ID='" & strUserNum & "' "
    adoTaie.Execute stSQL
    stSQL = "Insert Into Accrpt44j0_2 (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011) " & _
                  GetPoint(0, FCDate(MaskEdBox1.Text), FCDate(MaskEdBox2.Text), "S", , Text2, False, Me.Name, , strCmp) & _
    " Union " & GetPoint(0, FCDate(MaskEdBox1.Text), FCDate(MaskEdBox2.Text), "*NS", , Text2, False, Me.Name, True, strCmp)
    adoTaie.Execute stSQL, intE
    
    'Modify by Amy 2020/07/07 +if 下跨月報表,期末顯示0,故R013設Y
    If Left(Val(FCDate(MaskEdBox1.Text)) + 19110000, 6) = Left(Val(FCDate(MaskEdBox2.Text)) + 19110000, 6) Then
        'Add by Amy 2020/06/18 +開放輸SalesPoint需判斷是否有輸(因10905 20091 未輸時,未關閉前產生之報表,期末不應顯示0)
        stSQL = "Update Accrpt44j0_2 Set R013=(" & _
                    "Select Decode(SubStr(SP48,1,1),'S',Decode(SP03||SP07||SP11||SP15||SP19||SP20||SP24||SP28||SP32||SP36||SP40||SP41,null,null,'Y'),'Y') " & _
                    "From SalesPoint Where ID='" & strUserNum & "' And R001=SP48(+) And R003=SP02(+) And SP01=" & Left(Val(FCDate(MaskEdBox1.Text) + 19110000), 6) & _
                    " ) " & _
                    "Where ID='" & strUserNum & "' "
        adoTaie.Execute stSQL
        'end 2020/06/18
    Else
        stSQL = "Update Accrpt44j0_2 Set R013='Y' Where ID='" & strUserNum & "' "
        adoTaie.Execute stSQL
    End If
    
    'Modify by Amy 2017/07/05 原抓SalesPoint改抓目標檔為主,ex:楊挺104/7/14 離職,104年各單月目標加總和全年不合(8月報表沒楊挺無傳票資料)
    '加入SalesPoint有目標沒有傳票的人員(ex:新人有目標沒傳票)
'    If Val(Left(FCDate(MaskEdBox1.Text), 5)) >= Val(業績輸入啟用年月) Then
    If Val(Left(FCDate(MaskEdBox1.Text), 3)) >= 105 Then
        stSQL = " And SP01(+)=" & Val(Left(FCDate(MaskEdBox1.Text), 5)) + 191100
        If Left(FCDate(MaskEdBox1.Text), 5) <> Left(FCDate(MaskEdBox2.Text), 5) Then
            stSQL = " And SP01(+)>=" & Val(Left(FCDate(MaskEdBox1.Text), 5)) + 191100 & " And SP01(+)<=" & Val(Left(FCDate(MaskEdBox2.Text), 5)) + 191100
        End If
    Else
        stSQL = " And SP01(+)=201512 "
    End If
    stSQL = stSQL & " And PE03>=" & Val(Left(FCDate(MaskEdBox1.Text), 5)) + 191100 & " And PE03<=" & Val(Left(FCDate(MaskEdBox2.Text), 5)) + 191100
        'Modify by Amy 2016/07/06 +SalesPoint 輸入的值不是0才出現
        'Modify by Amy 2016/08/04 S29要出現,故輸入值0也出現
'        stSQL = "Insert into Accrpt44j0_2 (ID,R001,R002,R003) " & _
'                     "Select  Distinct '" & strUserNum & "',SP48,SP02||ST02,SP02 From SalesPoint,Staff " & _
'                     "Where SP02=ST01(+) And (sp03<>0 Or sp24<>0 Or sp07<>0 Or sp28<>0 Or sp11<>0 Or sp32<>0 Or sp15<>0 Or sp36<>0 Or sp19<>0 Or sp40<>0) " & stSQL & _
'                     " And SP02 Not in (Select R003 From Accrpt44j0_2 Where ID='" & strUserNum & "')"
'        stSQL = "Insert into Accrpt44j0_2 (ID,R001,R002,R003) " & _
'                     "Select  Distinct '" & strUserNum & "',SP48,SP02||ST02,SP02 From SalesPoint,Staff " & _
'                     "Where SP02=ST01(+) " & stSQL & _
'                     " And SP02 Not in (Select R003 From Accrpt44j0_2 Where ID='" & strUserNum & "')"
        'Modify by Amy 2017/07/07 下104年全所 無 67001/s212/s232資料 目標總合錯誤,因SalesPoint 201512月沒有這些人員,改SP48沒值抓ST15
        stSQL = "Insert into Accrpt44j0_2 (ID,R001,R002,R003) " & _
                     "Select  Distinct '" & strUserNum & "',Nvl(SP48,ST15),PE01||ST02,PE01 From PerFormance,SalesPoint,Staff " & _
                     "Where PE01=ST01(+) And PE02='TOT' And PE01=SP02(+) " & stSQL & _
                     " And PE01 Not in (Select R003 From Accrpt44j0_2 Where ID='" & strUserNum & "')"
        adoTaie.Execute stSQL, intSP
'    End If
    'end 2017/0705

    'Add by Amy 2018/01/25 畫面若跨月,抓止日月份SalesPoint部門別(10610高國碩及陳頌恩由中一調中二)
    If Val(Left(FCDate(MaskEdBox1.Text), 3)) >= 105 And Left(FCDate(MaskEdBox1.Text), 5) <> Left(FCDate(MaskEdBox2.Text), 5) Then
        strSqlF = "Select SP48,SP02 From " & _
                         "(Select SP48,SP02 From SalesPoint  Where SP01=" & Val(Left(FCDate(MaskEdBox2.Text), 5) + 191100) & ")," & _
                         "(Select R001,R003 From Accrpt44j0_2 Where ID='" & strUserNum & "') " & _
                       "Where SP02=R003 And R001<>SP48  "
        If adoFinal.State <> adStateClosed Then adoFinal.Close
        adoFinal.CursorLocation = adUseClient
        adoFinal.Open strSqlF, adoTaie, adOpenStatic, adLockReadOnly
        If adoFinal.EOF = False Then adoFinal.MoveFirst
        Do While adoFinal.EOF = False
            stSQL = "Update Accrpt44j0_2 Set R001='" & adoFinal.Fields("SP48") & "' Where ID='" & strUserNum & "' And R003='" & adoFinal.Fields("SP02") & "' "
            adoTaie.Execute stSQL
            adoFinal.MoveNext
        Loop
    End If
    
    'Add by Amy 2016/11/22
    '更新 A4023文雄 部門為 S14北四區
    stSQL = "Update Accrpt44j0_2 Set R001='S14' Where ID='" & strUserNum & "' And R003='A4023' "
    adoTaie.Execute stSQL
   
    '畫面條件下105年01以後抓取目標(105年因有中所有換部門,目標檔無部門欄位,會抓錯)
    '(ex:S142 沒傳票資料(財務新增SalesPoint人員時不會有,但有目標(SalesPoint可能會新增資料),所以再此抓目標)
    'Modify by Amy 2017/06/28 1050101以前仍可以抓,以員編抓應該可以-秀玲
    'If Val(FCDate(MaskEdBox1.Text)) >= 1050101 Then
    stSQL = "Update Accrpt44j0_2 Set R012=(Select sum(PE04)*1000 From PerFormance " & _
                "Where PE03>=" & Left(FCDate(MaskEdBox1.Text), 5) + 191100 & " And PE03<=" & Left(FCDate(MaskEdBox2.Text), 5) + 191100 & _
                " And  PE02='TOT' And PE01=R003 " & " Group by PE01) " & _
                "Where ID='" & strUserNum & "'  "
    adoTaie.Execute stSQL
    'End If
    'end 2016/11/22
    
    'Add by Amy 2018/01/25 10610高國碩及陳頌恩由中一調中二 查10609~10610期末資料會為0,因抓SalesPoint 期末抓止月且有限制部門,導致沒資料
    '                                     改期末顯示於止月部門-婧瑄(故抓SalesPoint 時以人員抓)
    If Val(Left(FCDate(MaskEdBox1.Text), 5)) >= Val(業績輸入啟用年月) Then
        '抓SalesPoint期末保留及結餘(止月),轉撥抓區間
        strSqlF = GetPoint_SP(Val(Left(FCDate(MaskEdBox1.Text), 5)), Val(Left(FCDate(MaskEdBox2.Text), 5)), , , Text2, False, Me.Name, True, True)
        If adoFinal.State <> adStateClosed Then adoFinal.Close
        adoFinal.CursorLocation = adUseClient
        adoFinal.Open strSqlF, adoTaie, adOpenStatic, adLockReadOnly
        If adoFinal.EOF = False Then adoFinal.MoveFirst
        Do While adoFinal.EOF = False
            stSQL = "Update Accrpt44j0_2 Set R008=" & Val(adoFinal.Fields("SP15")) & ",R009=" & Val(adoFinal.Fields("SP36")) & _
                          ",R010=" & Val(adoFinal.Fields("SP19")) & ",R011=" & Val(adoFinal.Fields("SP40")) & _
                        " Where R003='" & adoFinal.Fields("SP02") & "' And ID='" & strUserNum & "' "
             adoTaie.Execute stSQL
             adoFinal.MoveNext
        Loop
    End If
    If intE + intSP = 0 Then MsgBox "無資料！": Exit Sub
    
    'Modify by Amy 2017/06/07 改檔名
    If Left(MaskEdBox1.Text, 3) = Left(MaskEdBox2.Text, 3) Then
        strFileName = Left(FCDate(MaskEdBox1.Text), 3) & "年度"
        If Mid(MaskEdBox1.Text, 5, 2) = Mid(MaskEdBox2.Text, 5, 2) Then
            strFileName = strFileName & Mid(MaskEdBox1.Text, 5, 2) & "月份"
        Else
            strFileName = strFileName & Mid(MaskEdBox1.Text, 5, 2) & "~" & Mid(MaskEdBox2.Text, 5, 2) & "月"
        End If
    Else
        strFileName = strFileName & Mid(MaskEdBox1.Text, 5, 2) & "月份~" & Left(MaskEdBox2.Text, 3) & Mid(MaskEdBox2.Text, 5, 2) & "月份"
    End If
    strFileName = strFileName & "智權點數實績與結餘分析表" & ACDate(ServerDate) & ServerTime & MsgText(43)
    If Dir(strExcelPath & strFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileName
    End If
    'end 2017/06/07
    xlsSalesPoint.Visible = True
    xlsSalesPoint.SheetsInNewWorkbook = 3 'Add by Amy  2019/04/02 改設定(選項->一般->包括的工作表份數)
    xlsSalesPoint.Workbooks.add
    Set wksaccrpt417 = xlsSalesPoint.Worksheets(1)
    wksaccrpt417.Range("a3").Value = "公司別:"
    wksaccrpt417.Range("b3").Value = GetAccReportCmpN(CboCmp, , True) 'Modify by Amy 2020/04/16
    wksaccrpt417.Range("a4").Value = ReportSum(27)
    wksaccrpt417.Range("b4").Value = MaskEdBox1.Text
    wksaccrpt417.Range("c4").Value = ReportSum(28)
    wksaccrpt417.Range("d4").Value = MaskEdBox2.Text
    
    '欄位名稱(欄位需照順序放)
    'Modify by Amy 2016/11/22 +欄位-瑞婷 (目標沒部門欄位,且10501起中所有換部門,故下104年以前資料不顯示目標,否則會抓錯)
    'Modify by Amy 2017/06/28 1050101以前仍可以抓,以員編抓應該可以-秀玲
'    If FCDate(MaskEdBox1.Text) >= 1050101 Then
        ReDim strFieldN(0 To 15)
        ReDim intWidth(0 To 15)
        ReDim strSum(1 To 15)
        ReDim strTotalAmt(1 To 15)
'    Else
'        ReDim strFieldN(0 To 12)
'        ReDim intWidth(0 To 12)
'        ReDim strSum(1 To 12)
'        ReDim strTotalAmt(1 To 12)
'    End If
    'end 2016/11/22
    i = 0: intTitleRow = 6: lngCounter = 6: lngCounter1 = 0: intField = 65
    '智權人員
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportSum(57): strFieldN(i) = ReportSum(57): intWidth(i) = 10.5: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    '期初實績保留
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    '期初結餘保留
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    '實績點數
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    '結餘點數
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    '期未實績保留
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    '期末結餘保留
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    '實績撥點數
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    '結餘撥點數
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    '報出實績點數
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    '報出結餘點數
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    '報出點數
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    '實績保留增減
    wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
    wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    'Add by Amy 2016/11/22 +欄位
    'Modify by Amy 2017/06/28 1050101以前仍可以抓,以員編抓應該可以-秀玲
    'If FCDate(MaskEdBox1.Text) >= 1050101 Then
        '目標
        wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
        wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
        '達成率
        wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
        wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
        '實績達成率
        wksaccrpt417.Range(Chr(intField) & lngCounter).Value = ReportFieldN2(i): strFieldN(i) = ReportFieldN2(i): intWidth(i) = 10: i = i + 1
        wksaccrpt417.Range(Chr(intField) & lngCounter).HorizontalAlignment = xlCenter: intField = intField + 1
    'End If
    
    intField = 65
    lngCounter = lngCounter + 1 '資料欄位列
    
    wksaccrpt417.PageSetup.PrintTitleRows = "$1:$" & intTitleRow
    For i = LBound(strFieldN) To UBound(strFieldN)
        wksaccrpt417.Columns(Chr(i + intField) & ":" & Chr(i + intField)).ColumnWidth = intWidth(i)
    Next i
    For i = LBound(strSum) To UBound(strSum)
        strSum(i) = "="
    Next i
    wksaccrpt417.Range("a1").Value = "智權點數實績與結餘分析表"
    wksaccrpt417.Range("a1:" & Chr(UBound(strFieldN) + intField) & "1").Select
    
   
    '*** 智權人員
    lngCounter1 = intTitleRow + 1
    'Modify by Amy 2016/11/22 +目標
    'Moidfy by Amy 2020/06/18 +SalesPoint 是否有輸
    stSQL = "Select R001 as SP48,R002 as StName,R003 as ST01,Nvl(R004,0) as C1,Nvl(R005,0) as C2,Nvl(R006,0) as C3,Nvl(R007,0) as C4,Nvl(R008,0) as C5,Nvl(R009,0) as C6,Nvl(R010,0) as T1,Nvl(R011,0) as T2,Nvl(R012,0) as PE04,R013 as ChkInput " & _
                  "From Accrpt44j0_2 Where ID='" & strUserNum & "' And SubStr(R001,1,1)='S' Order by R001,R003"
    If adostaff.State <> adStateClosed Then adostaff.Close
    adostaff.CursorLocation = adUseClient
    adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
    If adostaff.EOF = False Then adostaff.MoveFirst
    Do While adostaff.EOF = False
        If IsNull(adostaff.Fields("ST01").Value) = False Then
            '區合計
            If strDept <> adostaff.Fields("SP48") And strDept <> MsgText(601) Then
                Call GetTotal(0, wksaccrpt417, strDept, lngCounter, lngCounter1)
                'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
                If InsRptXlsLink(strDept, GetDepartmentName(strDept), lngCounter) = False Then
                    MsgBox "記錄區合計位置有問題,請洽電腦中心！", , MsgText(5)
                End If
                'end 2022/02/18
                lngCounter = lngCounter + 2
                lngCounter1 = lngCounter
            End If
            '北/中所合計
            If Mid(strDept, 1, 2) <> Mid(adostaff.Fields("SP48"), 1, 2) And (Mid(strDept, 1, 2) = "S1" Or Mid(strDept, 1, 2) = "S2") Then
                Call GetTotal(1, wksaccrpt417, strDept, lngCounter)
                'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
                strExc(0) = "北所合計"
                If Mid(strDept, 1, 2) = "S2" Then
                    strExc(0) = "中所合計"
                End If
                If InsRptXlsLink(Mid(strDept, 1, 2) & "z", strExc(0), lngCounter) = False Then
                    MsgBox "記錄區合計位置有問題,請洽電腦中心！", , MsgText(5)
                End If
                'end 2022/02/18
                lngCounter = lngCounter + 2
                lngCounter1 = lngCounter
            End If
            '資料
            Call GetPersonData(2, wksaccrpt417, adostaff, lngCounter, "" & adostaff.Fields("SP48"))
            lngCounter = lngCounter + 1
        End If
        strDept = adostaff.Fields("SP48")
        adostaff.MoveNext
    Loop
    If adostaff.RecordCount > 0 Then
        '智權最後一個部門合計
        Call GetTotal(0, wksaccrpt417, strDept, lngCounter, lngCounter1)
        'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
        If InsRptXlsLink(strDept, GetDepartmentName(strDept), lngCounter) = False Then
            MsgBox "記錄區合計位置有問題,請洽電腦中心！", , MsgText(5)
        End If
        'end 2022/02/18
        lngCounter = lngCounter + 2
        '智權智權部合計
        Call GetTotal(2, wksaccrpt417, strDept, lngCounter)
        lngCounter = lngCounter + 2
    End If
    
    'W1001/W2001/P2005 11001月開始搬至智權部門後顯示
    strWhere = ""
    If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= 20210100 Then
        lngCounter1 = lngCounter: stSQL = ""
        
        'Modify by Amy 2022/03/31 +P1005
        'Modify by Amy 2023/07/25 +W3001
        strWhere = ",'W1001','W2001','W3001','P2005','P1005'"
        stSQL = "Select R001 as SP48,R002 as StName,R003 as ST01,Nvl(R004,0) as C1,Nvl(R005,0) as C2,Nvl(R006,0) as C3,Nvl(R007,0) as C4,Nvl(R008,0) as C5,Nvl(R009,0) as C6,Nvl(R010,0) as T1,Nvl(R011,0) as T2,Nvl(R012,0) as PE04,R013 as ChkInput " & _
                      "From Accrpt44j0_2,Staff Where ID='" & strUserNum & "' And SubStr(R001,1,1)<>'S' And R003 in (" & Mid(strWhere, 2) & ") " & _
                      " And R003=St01(+) And Substr(St03,1,1)<>'L' Order by Decode(SubStr(R001,1,1),'W',1,2),R003"
    
        If adostaff.State <> adStateClosed Then adostaff.Close
        adostaff.CursorLocation = adUseClient
        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
        If adostaff.EOF = False Then adostaff.MoveFirst
        If adostaff.RecordCount > 0 Then
            bolData = True
            Do While adostaff.EOF = False
                Call GetPersonData(3, wksaccrpt417, adostaff, lngCounter, "" & adostaff.Fields("SP48"))
                'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
                strExc(0) = ""
                'Modify by Amy 2023/07/27 原:客服組/顧服組 改抓員工姓名,+W3001 改strRID
                Select Case "" & adostaff.Fields("SP48")
                    Case "W10"
                        strExc(0) = StaffQuery("W1001")
                        strRID = "21"
                    Case "W20"
                        strExc(0) = StaffQuery("W2001")
                        strRID = "22"
                    Case "W30"
                        strExc(0) = StaffQuery("W3001")
                        strRID = "23"
                    'Add by Amy 2022/03/31 +P1005
                    Case "P12"
                        strExc(0) = StaffQuery("P1005") 'Modify by Amy 2022/04/11 原:寫死
                        strRID = "31"
                    Case "P20"
                        strExc(0) = StaffQuery("P2005")
                        strRID = "32"
                End Select
                'end 2023/07/27
                If InsRptXlsLink("" & adostaff.Fields("SP48"), strExc(0), lngCounter, strRID) = False Then
                    MsgBox "記錄「" & strExc(0) & "」位置有問題,請洽電腦中心！", , MsgText(5)
                End If
                'end 2022/02/18
                lngCounter = lngCounter + 1
                adostaff.MoveNext
            Loop
            '只加至「國內合計」不顯示合計
            Call GetTotal(8, wksaccrpt417, strDept, lngCounter - 1, lngCounter1)
            lngCounter = lngCounter + 1
        End If
    End If

    '*** 非智權部門(其他人員不包含 M0100/F4101/F4102/F4103)
    lngCounter1 = lngCounter
    
    'Add by Amy 2020/04/22 畫面「傳票日期」起日若為「智慧所更名日」前,維持舊格式 ex:1090101~1091231
    '                                                                                         「智慧所更名日」後,L部門人員獨立顯示於最後區塊
    stSQL = ""
    'Moidfy by Amy 2020/06/18 +SalesPoint 是否有輸
    If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= Val(智慧所更名日) Then
        'Modify by Amy 2021/01/19 11001月始F4102拆成F4104,F4105/F4103拆成F4106,F4107
        'Modify by Amy 2021/02/19 +strWhere W1001/W2001/P2005 11001月開始搬至智權部門後顯示
        'Modify by Amy 2021/05/27 M0109 搬至「國外部合計」後顯示
        stSQL = "Select R001 as SP48,R002 as StName,R003 as ST01,Nvl(R004,0) as C1,Nvl(R005,0) as C2,Nvl(R006,0) as C3,Nvl(R007,0) as C4,Nvl(R008,0) as C5,Nvl(R009,0) as C6,Nvl(R010,0) as T1,Nvl(R011,0) as T2,Nvl(R012,0) as PE04,R013 as ChkInput " & _
                      "From Accrpt44j0_2,Staff Where ID='" & strUserNum & "' And SubStr(R001,1,1)<>'S' And R003 Not in ('M0100','M0109','F4101','F4102','F4103','F4104','F4105','F4106','F4107'" & strWhere & ") " & _
                      " And R003=St01(+) And Substr(St03,1,1)<>'L' Order by R001,R003"
    Else
        'Modify by Amy 2016/11/22 +目標
        stSQL = "Select R001 as SP48,R002 as StName,R003 as ST01,Nvl(R004,0) as C1,Nvl(R005,0) as C2,Nvl(R006,0) as C3,Nvl(R007,0) as C4,Nvl(R008,0) as C5,Nvl(R009,0) as C6,Nvl(R010,0) as T1,Nvl(R011,0) as T2,Nvl(R012,0) as PE04,R013 as ChkInput " & _
                      "From Accrpt44j0_2 Where ID='" & strUserNum & "' And SubStr(R001,1,1)<>'S' And R003 Not in ('M0100','F4101','F4102','F4103') Order by R001,R003"
    End If
    'end 2020/06/18
    'end 2020/04/22
    If adostaff.State <> adStateClosed Then adostaff.Close
    adostaff.CursorLocation = adUseClient
    adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
    If adostaff.EOF = False Then adostaff.MoveFirst
    If adostaff.RecordCount > 0 Then bolData = True
    Do While adostaff.EOF = False
        Call GetPersonData(3, wksaccrpt417, adostaff, lngCounter, "" & adostaff.Fields("SP48"))
        lngCounter = lngCounter + 1
        adostaff.MoveNext
    Loop
    
    '*** M0100 ***
    If Trim(Text2) = "" Or Text2 = "M0100" Then
        'M0100-All
        'Modify by Amy 2016/11/22 +目標
        'Moidfy by Amy 2020/06/18 +SalesPoint 是否有輸
        bolM0100 = False 'Add by Amy 2022/08/08 避免M0100沒資料會多一行, 有才出現
        stSQL = "Select R001 as SP48,R002 as StName,R003 as ST01,Nvl(R004,0) as C1,Nvl(R005,0) as C2,Nvl(R006,0) as C3,Nvl(R007,0) as C4,Nvl(R008,0) as C5,Nvl(R009,0) as C6,Nvl(R010,0) as T1,Nvl(R011,0) as T2,Nvl(R012,0) as PE04,R013 as ChkInput " & _
                     "From Accrpt44j0_2 Where ID='" & strUserNum & "' And R003='M0100' "
        If adostaff.State <> adStateClosed Then adostaff.Close
        adostaff.CursorLocation = adUseClient
        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
        If adostaff.EOF = False Then adostaff.MoveFirst
        Do While adostaff.EOF = False
            bolM0100 = True 'Add by Amy 2022/08/08 避免M0100沒資料會多一行, 有才出現
            '記錄M0100 Total值
            Call GetPersonData(31, wksaccrpt417, adostaff, lngCounter, "" & adostaff.Fields("SP48"))
            adostaff.MoveNext
        Loop
        'M0100-P
        stSQL = GetPoint(0, FCDate(MaskEdBox1.Text), FCDate(MaskEdBox2.Text), "TOT", , "M0100", False, Me.Name, True, strCmp, "P")
        If adostaff.State <> adStateClosed Then adostaff.Close
        adostaff.CursorLocation = adUseClient
        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
        If adostaff.EOF = False Then adostaff.MoveFirst
        Do While adostaff.EOF = False
            Call GetPersonData(32, wksaccrpt417, adostaff, lngCounter, "" & adostaff.Fields("SP48"))
            'Modify by Amy 2022/02/18 strM0100_P抓位置,工作表2改位置連動,位置寫入暫存檔
            'strM0100_P = wksaccrpt417.Range(Chr(GetValue("報出點數") + intField) & lngCounter) 'Add by Amy 2018/05/17 MCP
            strM0100_P = lngCounter
           
            strRID = "31" 'Modify by Amy 2023/07/27 原:23(位置同P1005)
            If InsRptXlsLink("P", "其他-MCP", lngCounter, strRID) = False Then
                MsgBox "記錄「其他-MCP」位置有問題,請洽電腦中心！", , MsgText(5)
            End If
            'end 2022/02/18
            lngCounter = lngCounter + 1
            adostaff.MoveNext
        Loop
        'M0100-T
        stSQL = GetPoint(0, FCDate(MaskEdBox1.Text), FCDate(MaskEdBox2.Text), "TOT", , "M0100", False, Me.Name, True, strCmp, "T")
        If adostaff.State <> adStateClosed Then adostaff.Close
        adostaff.CursorLocation = adUseClient
        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
        If adostaff.EOF = False Then adostaff.MoveFirst
        Do While adostaff.EOF = False
            Call GetPersonData(33, wksaccrpt417, adostaff, lngCounter, "" & adostaff.Fields("SP48"))
            strM0100_T = wksaccrpt417.Range(Chr(GetValue("報出點數") + intField) & lngCounter) 'Add by Amy 2018/05/17 MCT
            lngCounter = lngCounter + 1
            adostaff.MoveNext
        Loop
        'M0100-大陸P-大陸T
        If bolM0100 = True Then
            Call GetPersonData(34, wksaccrpt417, adostaff, lngCounter, "TOT")
            lngCounter = lngCounter + 1
        End If
    End If
    '*** End M0100 ***
    'Modify by Amy 2022/08/08 避免M0100沒資料
    'If adostaff.RecordCount > 0 Or bolData = True Then
    If bolM0100 = True > 0 Or bolData = True Then
        '其他人員合計
        Call GetTotal(3, wksaccrpt417, strDept, lngCounter, lngCounter1)
        'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
        strRID = "59"
        If InsRptXlsLink(strRID, "其他-國內", lngCounter, strRID) = False Then
            MsgBox "記錄「其他-國內」位置有問題,請洽電腦中心！", , MsgText(5)
        End If
        'end 2022/02/18
        lngCounter = lngCounter + 2
        
        '國內合計
        Call GetTotal(4, wksaccrpt417, strDept, lngCounter)
        'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
        strRID = "5Z"
        If InsRptXlsLink(strRID, "國內合計", lngCounter, strRID) = False Then
            MsgBox "記錄「國內合計」位置有問題,請洽電腦中心！", , MsgText(5)
        End If
        'end 2022/02/18
        lngCounter = lngCounter + 2
    End If
    
    '*** 國外部
    bolData = False: lngCounter1 = lngCounter
    'Modify by Amy 2021/01/19 +F4104~07
    bolSumF = ChkFCPFCTSum
    strSumForeign = ""
    If Trim(Text2) = "" Or Text2 = "F4101" Or Text2 = "F4102" Or Text2 = "F4103" Or Text2 = "F4104" Or Text2 = "F4105" Or Text2 = "F4106" Or Text2 = "F4107" Then
        'F4102-FCP
        'Modify by Amy 2016/11/22 +目標
        'Moidfy by Amy 2020/06/18 +SalesPoint 是否有輸
        'Modify by Amy 2021/01/19 +F4104,F4105
        stSQL = "Select R001 as SP48,R002 as StName,R003 as ST01,Nvl(R004,0) as C1,Nvl(R005,0) as C2,Nvl(R006,0) as C3,Nvl(R007,0) as C4,Nvl(R008,0) as C5,Nvl(R009,0) as C6,Nvl(R010,0) as T1,Nvl(R011,0) as T2,Nvl(R012,0) as PE04,R013 as ChkInput " & _
                     "From Accrpt44j0_2 Where ID='" & strUserNum & "' And R003 In ('F4102','F4104','F4105') Order by R003 "
        If adostaff.State <> adStateClosed Then adostaff.Close
        adostaff.CursorLocation = adUseClient
        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
        If adostaff.EOF = False Then adostaff.MoveFirst
        Do While adostaff.EOF = False
            bolData = True
            Call GetPersonData(5, wksaccrpt417, adostaff, lngCounter, "" & adostaff.Fields("SP48"))
            'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
            If InsRptXlsLink("" & adostaff.Fields("ST01"), GetPrjSalesNM("" & adostaff.Fields("ST01")), lngCounter) = False Then
                MsgBox "記錄區合計位置有問題,請洽電腦中心！", , MsgText(5)
            End If
            'end 2022/02/18
            lngCounter = lngCounter + 1
            adostaff.MoveNext
        Loop
        'Add by Amy 2021/01/19 +F4104,F4105合計
        If bolSumF = True Then
            Call GetTotal(0, wksaccrpt417, "FCP", lngCounter, lngCounter1)
            strSumForeign = strSumForeign & "," & lngCounter
            lngCounter = lngCounter + 2
            lngCounter1 = lngCounter
        End If
        
        'F4103-FCT
        'Modify by Amy 2016/11/22 +目標
        'Moidfy by Amy 2020/06/18 +SalesPoint 是否有輸
        'Modify by Amy 2021/01/19 +F4106,F4107
        stSQL = "Select R001 as SP48,R002 as StName,R003 as ST01,Nvl(R004,0) as C1,Nvl(R005,0) as C2,Nvl(R006,0) as C3,Nvl(R007,0) as C4,Nvl(R008,0) as C5,Nvl(R009,0) as C6,Nvl(R010,0) as T1,Nvl(R011,0) as T2,Nvl(R012,0) as PE04,R013 as ChkInput " & _
                     "From Accrpt44j0_2 Where ID='" & strUserNum & "' And R003 In ('F4103','F4106','F4107') Order by R003 "
        If adostaff.State <> adStateClosed Then adostaff.Close
        adostaff.CursorLocation = adUseClient
        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
        If adostaff.EOF = False Then adostaff.MoveFirst
        Do While adostaff.EOF = False
            bolData = True
            Call GetPersonData(5, wksaccrpt417, adostaff, lngCounter, "" & adostaff.Fields("SP48"))
            'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
            If InsRptXlsLink("" & adostaff.Fields("ST01"), GetPrjSalesNM("" & adostaff.Fields("ST01")), lngCounter) = False Then
                MsgBox "記錄區合計位置有問題,請洽電腦中心！", , MsgText(5)
            End If
            'end 2022/02/18
            lngCounter = lngCounter + 1
            adostaff.MoveNext
        Loop
        'Add by Amy 2021/01/19 +F4106,F4107合計
        If bolSumF = True Then
            Call GetTotal(0, wksaccrpt417, "FCT", lngCounter, lngCounter1)
            strSumForeign = strSumForeign & "," & lngCounter
            lngCounter = lngCounter + 2
            lngCounter1 = lngCounter
        End If
        
        'F4101-FCL(10501起不使用此,故不需抓SalesPoint)
        'Modify by Amy 2016/11/22 +目標
        'Moidfy by Amy 2020/06/18 +SalesPoint 是否有輸
        stSQL = "Select R001 as SP48,R002 as StName,R003 as ST01,Nvl(R004,0) as C1,Nvl(R005,0) as C2,Nvl(R006,0) as C3,Nvl(R007,0) as C4,Nvl(R008,0) as C5,Nvl(R009,0) as C6,Nvl(R010,0) as T1,Nvl(R011,0) as T2,Nvl(R012,0) as PE04,R013 as ChkInput " & _
                     "From Accrpt44j0_2 Where ID='" & strUserNum & "' And R003='F4101' "
        If adostaff.State <> adStateClosed Then adostaff.Close
        adostaff.CursorLocation = adUseClient
        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
        If adostaff.EOF = False Then adostaff.MoveFirst
        Do While adostaff.EOF = False
            bolData = True
            Call GetPersonData(5, wksaccrpt417, adostaff, lngCounter, "" & adostaff.Fields("SP48"))
            'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
            If InsRptXlsLink("" & adostaff.Fields("ST01"), GetPrjSalesNM("" & adostaff.Fields("ST01")), lngCounter) = False Then
                MsgBox "記錄區合計位置有問題,請洽電腦中心！", , MsgText(5)
            End If
            'end 2022/02/18
            lngCounter = lngCounter + 1
            adostaff.MoveNext
        Loop
        '國外合計
        If bolData = True Then
            'Modify by Amy 2021/01/19 11001月始F4102拆成F4104,F4105/F4103拆成F4106,F4107,加總會變成不連續位置
            If strSumForeign <> MsgText(601) Then
                Call GetTotal(5, wksaccrpt417, strDept, lngCounter, lngCounter1, Mid(strSumForeign, 2))
            Else
                Call GetTotal(5, wksaccrpt417, strDept, lngCounter, lngCounter1)
            End If
            'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
            strRID = "Fz"
            If InsRptXlsLink(strRID, "國外合計", lngCounter, strRID) = False Then
                MsgBox "記錄「國外合計」位置有問題,請洽電腦中心！", , MsgText(5)
            End If
            'end 2022/02/18
            lngCounter = lngCounter + 2
        End If
    End If
    'Add by Amy 2021/05/27 +M0109 安全基金撥補(有資料才顯示,與婧瑄討論後與美珍報表一致)
    bolData = False
    If Trim(Text2) = "" Or Text2 = "M0109" Then
        stSQL = "Select R001 as SP48,R002 as StName,R003 as ST01,Nvl(R004,0) as C1,Nvl(R005,0) as C2,Nvl(R006,0) as C3,Nvl(R007,0) as C4,Nvl(R008,0) as C5,Nvl(R009,0) as C6,Nvl(R010,0) as T1,Nvl(R011,0) as T2,Nvl(R012,0) as PE04,R013 as ChkInput " & _
                     "From Accrpt44j0_2 Where ID='" & strUserNum & "' And R003='M0109' "
        If adostaff.State <> adStateClosed Then adostaff.Close
        adostaff.CursorLocation = adUseClient
        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
        If adostaff.EOF = False Then
            adostaff.MoveFirst
            lngCounter1 = lngCounter
        End If
        Do While adostaff.EOF = False
            bolData = True
            Call GetPersonData(3, wksaccrpt417, adostaff, lngCounter, "" & adostaff.Fields("SP48"))
            'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
            strRID = "91"
            If InsRptXlsLink("" & adostaff.Fields("ST01"), GetPrjSalesNM("" & adostaff.Fields("ST01")), lngCounter, strRID) = False Then
                MsgBox "記錄「安全基金撥補」位置有問題,請洽電腦中心！", , MsgText(5)
            End If
            'end 2022/02/18
            lngCounter = lngCounter + 1
            adostaff.MoveNext
        Loop
        '將值加入「全所合計」
        If bolData = True Then
            Call GetTotal(8, wksaccrpt417, strDept, lngCounter - 1, lngCounter1)
            lngCounter = lngCounter + 1
        End If
    End If
    
    '全所合計
    Call GetTotal(6, wksaccrpt417, strDept, lngCounter)
    'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
    strRID = "ZZ"
    If InsRptXlsLink(strRID, "全所合計", intCounter, strRID) = False Then
        MsgBox "記錄區合計位置有問題,請洽電腦中心！", , MsgText(5)
    End If
    'end 2022/02/18
     
    '框線
    wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + intField) & intTitleRow & ":" & Chr(UBound(strFieldN) + intField) & lngCounter).Select
    xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Font.Size = 8
        
     'Add by Amy 2020/04/22
    '*** L部門人員獨立顯示
    If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= Val(智慧所更名日) And (strCmp = MsgText(601) Or strCmp = "L") Then
        lngCounter = lngCounter + 3
        intLCmpStartRow = lngCounter
        
        stSQL = "Select R001 as SP48,R002 as StName,R003 as ST01,Nvl(R004,0) as C1,Nvl(R005,0) as C2,Nvl(R006,0) as C3,Nvl(R007,0) as C4,Nvl(R008,0) as C5,Nvl(R009,0) as C6,Nvl(R010,0) as T1,Nvl(R011,0) as T2,Nvl(R012,0) as PE04,R013 as ChkInput " & _
                      "From Accrpt44j0_2,Staff Where ID='" & strUserNum & "' And R003=St01(+) And Substr(St03,1,1)='L' Order by R001,R003"
        If adostaff.State <> adStateClosed Then adostaff.Close
        adostaff.CursorLocation = adUseClient
        adostaff.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
        If adostaff.EOF = False Then adostaff.MoveFirst
        If adostaff.RecordCount > 0 Then bolData = True
        Do While adostaff.EOF = False
            Call GetPersonData(3, wksaccrpt417, adostaff, lngCounter, "" & adostaff.Fields("SP48"))
            lngCounter = lngCounter + 1
            adostaff.MoveNext
        Loop
        Call GetTotal(7, wksaccrpt417, strDept, lngCounter, intLCmpStartRow)
        'Add by Amy 2022/02/18 工作表2改位置連動,位置寫入暫存檔
        strRID = "LZ"
        If InsRptXlsLink(strRID, "法律所合計", lngCounter, strRID) = False Then
            MsgBox "記錄區合計位置有問題,請洽電腦中心！", , MsgText(5)
        End If
        'end 2022/02/18
        
        '框線
        wksaccrpt417.Range(Chr(GetValue(ReportSum(57)) + intField) & intLCmpStartRow & ":" & Chr(UBound(strFieldN) + intField) & lngCounter).Select
        xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Font.Size = 8
    End If
    'end 2020/04/22
    
    '格式設定
    'Modify by Amy 2016/11/22 +欄位
    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + intField) & intTitleRow & ":" & Chr(GetValue("實績保留增減") + intField) & lngCounter).Select
    'Modify by Amy 2019/10/28 報出實績/結餘點數 負值顯示紅色,因10809 婉莘輸F4102 結餘總經理欄輸-15(輸錯),導致報出XX點數為負
    wksaccrpt417.Range(Chr(GetValue(ReportFieldN2(1)) + intField) & intTitleRow & ":" & Chr(GetValue("結餘轉撥點數") + intField) & lngCounter).NumberFormatLocal = "#,##0.00_ "
    wksaccrpt417.Range(Chr(GetValue("報出實績點數") + intField) & intTitleRow & ":" & Chr(GetValue("報出結餘點數") + intField) & lngCounter).NumberFormatLocal = "#,##0.00;[紅色]-#,##0.00 "
    wksaccrpt417.Range(Chr(GetValue("報出點數") + intField) & intTitleRow & ":" & Chr(GetValue("實績保留增減") + intField) & lngCounter).NumberFormatLocal = "#,##0.00_ "
    'end 2019/10/28
    wksaccrpt417.Range(Chr(GetValue("實績達成率") + intField) & intTitleRow & ":" & Chr(GetValue("實績達成率") + intField) & lngCounter).Select
    wksaccrpt417.Range(Chr(GetValue("實績達成率") + intField) & intTitleRow & ":" & Chr(GetValue("實績達成率") + intField) & lngCounter).NumberFormatLocal = "0.00%"
    'end 2016/11/22
    
    'Add by Amy 2017/10/02 預設A4紙張/橫式/比例 80%/水平置中/邊界左右都改0-瑞婷
    wksaccrpt417.PageSetup.PaperSize = 9 'A4
    wksaccrpt417.PageSetup.Orientation = xlLandscape '橫印
    wksaccrpt417.PageSetup.Zoom = 80
    wksaccrpt417.PageSetup.LeftMargin = 0 '邊界
    wksaccrpt417.PageSetup.RightMargin = 0
    wksaccrpt417.PageSetup.TopMargin = xlsSalesPoint.InchesToPoints(0.4)
    wksaccrpt417.PageSetup.BottomMargin = xlsSalesPoint.InchesToPoints(0.4)
    wksaccrpt417.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
    'end 2017/10/02
    'Add by Amy 2018/05/17 增加 當月業績達成明細表
    strWkName = Left(xlsSalesPoint.Worksheets(1).Name, Len(xlsSalesPoint.Worksheets(1).Name) - 1)
    Set wksaccrpt417 = xlsSalesPoint.Worksheets(strWkName & "2")
    'wksaccrpt417.Activate
    'Modify by Amy 2022/02/28 改連動位置
    'Modify by Amy 2020/11/09 原邊界/框線/產生L部門資料 改至ExcelSaveNew2_S2
    'Call ExcelSaveNew2_S2(xlsSalesPoint, wksaccrpt417, strM0100_P, strM0100_T, lngCounter, strCmp)
    Call ExcelSaveNew2_S2(xlsSalesPoint, wksaccrpt417, lngCounter)
    'end 2018/05/17
    
   'Modify by Amy 2014/06/11 +判斷若版本2007以上改變存格式
   'Modify by Amy 2017/06/07 改檔名
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
   End If
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   StatusClear
   Exit Sub

ErrHand:
    If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    MsgBox Err.Description, , MsgText(5)
End Sub

'Modify by Amy 2021/01/19 +stSum參數
Private Sub GetTotal(ByVal intChoose As Integer, ByRef Wks As Worksheet, ByVal stDept As String, ByVal intCounter As Long, Optional ByRef intCounter1 As Long = 0, Optional ByVal stSum As String = "")
    Dim stRecName As String, stData As String
    Dim arrTmp, ii As Integer 'Add by Amy 2021/01/19
    
    If stSum <> MsgText(601) Then arrTmp = Split(stSum, ",")
    
    Select Case intChoose
        '各區合計
        Case 0
            stRecName = GetDepartmentName(stDept) & ReportSum(25)
        '北/中所
        Case 1
            If Mid(stDept, 1, 2) = "S9" Then
                stRecName = ReportSum(127)
            Else
                stRecName = ReportSum(104 + Val(Right(Mid(stDept, 1, 2), 1)))
            End If
        Case 2
            stRecName = "智權部合計:"
        '其他合計
        Case 3
            stRecName = ReportSum(64) & ReportSum(25)
        '國內合計
        Case 4
            stRecName = ReportSum(65) & ReportSum(25)
        '國外合計
        Case 5
            stRecName = ReportSum(70) & ReportSum(25)
        '總所合計
        Case 6
            stRecName = ReportSum(66) & ReportSum(25)
        'Add by Amy 2020/04/22
        Case 7
            stRecName = "法律所" & ReportSum(25)
    End Select
    
    'Add by Amy 2021/02/19 +if intChoose = 8 因W1001/W2001/P2005 11001月開始搬至智權部門後顯示,但不需合計
    If intChoose = 8 Then
        For i = GetValue("期初實績保留") To GetValue("實績達成率")
            If intCounter1 = intCounter Then
                strTotalAmt(i) = strTotalAmt(i) & Chr(i + intField) & intCounter & ","
            Else
                strTotalAmt(i) = strTotalAmt(i) & "Sum(" & Chr(i + intField) & intCounter1 & ":" & Chr(i + intField) & intCounter & "),"
            End If
        Next i
    Else
        Wks.Range(Chr(GetValue(ReportSum(57)) + intField) & intCounter).Value = stRecName
        For i = LBound(strSum) To UBound(strSum)
            If strSum(i) <> MsgText(601) Then
                If i >= GetValue("達成率") And GetValue("達成率") <> 0 Then
                    Select Case i
                        Case GetValue("達成率")
                            stData = "=IF(OR(" & Chr(GetValue("報出點數") + intField) & intCounter & "=0," & Chr(GetValue("目標") + intField) & intCounter & "=0),0," & Chr(GetValue("報出點數") + intField) & intCounter & "/" & Chr(GetValue("目標") + intField) & intCounter & ")"
                        Case GetValue("實績達成率")
                            stData = "=IF(OR(" & Chr(GetValue("當月實績點數") + intField) & intCounter & "=0," & Chr(GetValue("目標") + intField) & intCounter & "=0),0," & Chr(GetValue("當月實績點數") + intField) & intCounter & "/" & Chr(GetValue("目標") + intField) & intCounter & ")"
                    End Select
                Else
                    Select Case intChoose
                        '各區,其他,L公司 合計
                        Case 0, 3, 7 'Modify by Amy 2020/04/22 +7
                            'Modify by Amy 2023/05/05 L公司若未有資料,開Excel 會彈公式循環的錯誤
                            If intCounter1 = intCounter Then
                                stData = "0"
                            Else
                                stData = Chr(GetValue(ReportFieldN2(i)) + intField) & intCounter1 & ":" & Chr(GetValue(ReportFieldN2(i)) + intField) & intCounter - 1
                            End If
                        '智權部,國內,總所 合計
                        Case 2, 4, 6
                            stData = Mid(strTotalAmt(i), 1, Len(strTotalAmt(i)) - 1)
                        'Add by Amy 2021/01/19 11001月始F4102拆成F4104,F4105/F4103拆成F4106,F4107,加總會變成不連續位置
                        '國外 合計
                        Case 5
                            If stSum <> MsgText(601) Then
                                stData = ""
                                For ii = LBound(arrTmp) To UBound(arrTmp)
                                    stData = stData & "," & Chr(GetValue(ReportFieldN2(i)) + intField) & arrTmp(ii)
                                Next ii
                                stData = Mid(stData, 2)
                            Else
                                stData = Chr(GetValue(ReportFieldN2(i)) + intField) & intCounter1 & ":" & Chr(GetValue(ReportFieldN2(i)) + intField) & intCounter - 1
                            End If
                        Case Else
                            stData = Mid(strSum(i), 1, Len(strSum(i)) - 1)
                    End Select
                End If
            End If
            If intChoose = 1 Or Left(stData, 3) = "=IF" Then
                Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & intCounter).Formula = stData
            'Add by Amy 2023/05/05 L公司若未有資料,開Excel 會彈公式循環的錯誤
            ElseIf stData = "0" Then
                Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & intCounter).Value = stData
            Else
                Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & intCounter).Formula = "=Sum(" & stData & ")"
            End If
            If Left(stData, 3) = "=IF" Then
                Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & intCounter).NumberFormatLocal = "0.00%"
            ElseIf i = GetValue("目標") Then
                Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & intCounter).NumberFormatLocal = "#,##0.00_ "
            End If
            Select Case intChoose
                Case 0
                    If stDept = "S00" Then
                        strS00Row = intCounter
                    Else
                        strSum(i) = strSum(i) & Chr(GetValue(ReportFieldN2(i)) + intField) & intCounter & "+"
                    End If
                    strTotalAmt(i) = strTotalAmt(i) & Chr(GetValue(ReportFieldN2(i)) + intField) & intCounter & ","
                Case 3, 5
                    'Add by Amy 2021/01/26 +if 11001月始F4102拆成F4104,F4105/F4103拆成F4106,F4107不需再加,否則會重覆加
                    If intChoose = 5 And stSum <> MsgText(601) Then
                    Else
                        strTotalAmt(i) = strTotalAmt(i) & Chr(GetValue(ReportFieldN2(i)) + intField) & intCounter & ","
                    End If
            End Select
        Next i
    End If
    
    If intChoose = 1 Or intChoose = 4 Then
        For i = LBound(strSum) To UBound(strSum)
            strSum(i) = "="
        Next i
    End If
End Sub

'Modify by Amy 2016/11/22 +stDept部門參數
'Modify by Amy 2018/01/25 拿掉ByRef m_adoFinal As ADODB.Recordset(直接更新暫存檔欄位)
Private Sub GetPersonData(ByVal intChoose As Integer, ByRef Wks As Worksheet, ByRef m_adoStaff As ADODB.Recordset, ByVal lngCounter As Long, ByVal stDept As String)
    Dim stValue(1) As String, stSP(3) As String
    Dim bolFormat As Boolean
    
    'Modify by Amy 2017/07/11 104年M0100期末結餘保留各月與全年加總不合 (可能因為財務說M0100會全報,才會固定寫0),改顯示抓傳票資料
    If intChoose >= 31 Then
        stSP(0) = "0"
        stSP(1) = "0"
        stSP(2) = "0"
        stSP(3) = "0"
    End If
    '10501月以前抓轉撥抓傳票資料
    '10501月以後抓SalesPoint資料(Modify by Amy 2018/01/25 直接更新暫存檔欄位)
    If m_adoStaff.EOF = False Then
        stSP(0) = m_adoStaff.Fields("C5") '期末實績保留(SP15)
        stSP(1) = m_adoStaff.Fields("C6") '期末結餘保留(SP36)
        stSP(2) = m_adoStaff.Fields("T1") '實績轉撥點數(SP19)
        stSP(3) = m_adoStaff.Fields("T2") '結餘轉撥點數(SP40)
    End If
    'end 2017/07/11
        
    If intChoose >= 31 Then
        Select Case intChoose
            Case 31
                strM0100_C(0) = m_adoStaff.Fields("StName")
            Case 32
                strM0100(0) = m_adoStaff.Fields("ST01") & "大陸P"
            Case 33
                strM0100(0) = m_adoStaff.Fields("ST01") & "大陸T"
            Case 34
                '名稱記錄於 strM0100_C(0)
        End Select
        If intChoose = 34 Then
            Wks.Range(Chr(GetValue(ReportSum(57)) + intField) & lngCounter).Value = strM0100_C(0)
        ElseIf intChoose <> 31 Then
            Wks.Range(Chr(GetValue(ReportSum(57)) + intField) & lngCounter).Value = strM0100(0)
        End If
    'Memo by Amy 2021/01/19 F4102/F4103 11001月後不用,但舊資料會用
    ElseIf m_adoStaff.Fields("ST01") = "F4101" Or m_adoStaff.Fields("ST01") = "F4102" Or m_adoStaff.Fields("ST01") = "F4103" Then
        Select Case Right(m_adoStaff.Fields("ST01"), 1)
            'FCL
            Case 1
                stValue(0) = ReportSum(69) & ReportSum(25)
            'FCP
            Case 2
                stValue(0) = ReportSum(67) & ReportSum(25)
            'FCT
            Case 3
                stValue(0) = ReportSum(68) & ReportSum(25)
        End Select
        Wks.Range(Chr(GetValue(ReportSum(57)) + intField) & lngCounter).Value = stValue(0)
    Else
        Wks.Range(Chr(GetValue(ReportSum(57)) + intField) & lngCounter).Value = m_adoStaff.Fields("StName")
    End If

    For i = LBound(strFieldN) + 1 To UBound(strFieldN)
        bolFormat = False
        Select Case ReportFieldN2(i)
            Case "期初實績保留"
                If intChoose = 34 Then
                    stValue(1) = strM0100_C(i)
                Else
                    stValue(1) = m_adoStaff.Fields("C1")
                End If
            Case "期初結餘保留"
                If intChoose = 34 Then
                    stValue(1) = strM0100_C(i)
                Else
                    stValue(1) = m_adoStaff.Fields("C2")
                End If
            Case "當月實績點數"
                If intChoose = 34 Then
                    stValue(1) = strM0100_C(i)
                Else
                    stValue(1) = m_adoStaff.Fields("C3")
                End If
            Case "當月結餘點數"
                If intChoose = 34 Then
                    stValue(1) = strM0100_C(i)
                Else
                    stValue(1) = m_adoStaff.Fields("C4")
                End If
            Case "期末實績保留"
                'Add by Amy 2020/06/18 +if 開放輸入之部門,若尚未輸入SalesPoint,期末=期初
                If intChoose >= 31 And intChoose <= 34 Then
                    'M0100相關資料,抓GetPoint語法不會有ChkInput欄位
                    stValue(1) = stSP(0)
                ElseIf IsNull(m_adoStaff.Fields("ChkInput")) Then
                    '期末實績保留=期初末實績保留(因當月實績一定要報,所以不需加當月)
                    stValue(1) = m_adoStaff.Fields("C1")
                Else
                    stValue(1) = stSP(0)
                End If
                'end 2020/06/18
            Case "期末結餘保留"
                'Add by Amy 2020/06/18 +if 開放輸入之部門,若尚未輸入SalesPoint,期末=期初+當月
                If intChoose >= 31 And intChoose <= 34 Then
                    'M0100相關資料,抓GetPoint語法不會有ChkInput欄位
                    stValue(1) = stSP(1)
                ElseIf IsNull(m_adoStaff.Fields("ChkInput")) Then
                    '期末結餘保留=期初結餘保留+當月結餘點數
                    stValue(1) = Val(m_adoStaff.Fields("C2")) + Val(m_adoStaff.Fields("C4"))
                Else
                    stValue(1) = stSP(1)
                End If
                'end 2020/06/18
            Case "實績轉撥點數"
                stValue(1) = stSP(2)
            Case "結餘轉撥點數"
                stValue(1) = stSP(3)
            Case "報出實績點數"
                stValue(1) = "=" & Chr(GetValue("期初實績保留") + intField) & lngCounter & "+" & Chr(GetValue("當月實績點數") + intField) & lngCounter & "-" & Chr(GetValue("期末實績保留") + intField) & lngCounter & _
                                        "+" & Chr(GetValue("實績轉撥點數") + intField) & lngCounter
            Case "報出結餘點數"
                stValue(1) = "=" & Chr(GetValue("期初結餘保留") + intField) & lngCounter & "+" & Chr(GetValue("當月結餘點數") + intField) & lngCounter & "-" & Chr(GetValue("期末結餘保留") + intField) & lngCounter & _
                                        "+" & Chr(GetValue("結餘轉撥點數") + intField) & lngCounter
            Case "報出點數"
                stValue(1) = "=" & Chr(GetValue("報出實績點數") + intField) & lngCounter & "+" & Chr(GetValue("報出結餘點數") + intField) & lngCounter
            Case "實績保留增減"
                'Modify by Amy 2016/03/11 瑞婷說改相反 原:期初實績保留-期末實績保留
                stValue(1) = "=" & Chr(GetValue("期末實績保留") + intField) & lngCounter & "-" & Chr(GetValue("期初實績保留") + intField) & lngCounter
            'Add by Amy 2016/11/22
            Case "目標"
                If intChoose >= 32 Then
                    stValue(1) = "0"
                    bolFormat = True
                Else
                    stValue(1) = "" & m_adoStaff.Fields("PE04")
                    'Modify by Amy 2021/02/24 W1001/W2001/P2005目標要出現,且F部門只有F41XX目標要出現
                    If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= 20210100 Then
                        'Modify by Amy 2022/06/06 P1005 目標也要出現
                        'Modify by Amy 2022/08/08 避免有上、下搬移有目標未算到,故 非S部門有目標就要算顯示(排除林總,因有人的目標掛於林總)
                        'If (Left(stDept, 1) = "F" And Left("" & m_adoStaff.Fields("ST01"), 3) = "F41") Or (Left(stDept, 1) = "P" And ("" & m_adoStaff.Fields("ST01") = "P2005" Or "" & m_adoStaff.Fields("ST01") = "P1005")) _
                          Or Left(stDept, 1) = "S" Or Left(stDept, 1) = "W" Then
                        If Left(stDept, 1) = "S" Or (Left(stDept, 1) <> "S" And Val(stValue(1)) <> 0 And m_adoStaff.Fields("ST01") <> "94007") Then
                        'end 2022/08/08
                        Else
                            bolFormat = True
                        End If
                    Else
                        If (Left(stDept, 1) <> "S" And Left(stDept, 1) <> "F") Or stDept = "TOT" Then bolFormat = True
                    End If
                End If
            Case "達成率"
                'Memo 智權部只有合計需算達成率,非S部門每月需輸入點數者需算達成率(非S部門需輸點數都有目標)
                'Modify by Amy 2021/02/24 F部門非F41XX員編不需算達成率,+W部門及P2005
                'Modify by Amy 2022/08/08 避免有上、下搬移有目標未算到,故 非S部門有目標就要算顯示(排除林總,因有人的目標掛於林總)
                'Modify by Amy 2022/08/24 bug 2022/08/08 改後 F41XX 未計算達成率
                strExc(1) = Wks.Range(Chr(GetValue("目標") + intField) & lngCounter).Value
                strExc(2) = Wks.Range(Chr(GetValue("智權人員") + intField) & lngCounter).Value
                If stDept = "TOT" Then
                    stValue(1) = ""
'                ElseIf ((Left(stDept, 1) = "F" And Left("" & m_adoStaff.Fields("ST01"), 3) = "F41")) Or Left(stDept, 1) = "W" Or (Left(stDept, 1) = "P" And "" & m_adoStaff.Fields("ST01") = "P2005") Then
                ElseIf (Left(stDept, 1) = "S" And Right(strExc(2), 2) = "合計") Or (Left(stDept, 1) <> "S" And Val(strExc(1)) <> 0 And m_adoStaff.Fields("ST01") <> "94007") Then
                'end 2022/08/08
                    stValue(1) = "=IF(OR(" & Chr(GetValue("報出點數") + intField) & lngCounter & "=0," & Chr(GetValue("目標") + intField) & lngCounter & "=0),0," & Chr(GetValue("報出點數") + intField) & lngCounter & "/" & Chr(GetValue("目標") + intField) & lngCounter & ")"
                '個人不需顯示達成率值
                Else
                    stValue(1) = ""
                End If
            Case "實績達成率"
                'Memo 有目標就要算實績達成率(包含智權部個人)
                'Modify by Amy 2021/02/24 F部門非F41XX員編不需算達成率,+W部門及P2005
                'Modify by Amy 2022/08/08 避免有上、下搬移有目標未算到,故 非S部門有目標就要算顯示(排除林總,因有人的目標掛於林總)
                strExc(1) = Wks.Range(Chr(GetValue("目標") + intField) & lngCounter).Value 'Add by Amy 2022/08/24 bug 2022/08/08 忘了加
                If stDept = "TOT" Then
                    stValue(1) = "0"
                'ElseIf ((Left(stDept, 1) = "F" And Left("" & m_adoStaff.Fields("ST01"), 3) = "F41")) Or Left(stDept, 1) = "W" Or Left(stDept, 1) = "S" Or (Left(stDept, 1) = "P" And "" & m_adoStaff.Fields("ST01") = "P2005") Then
                ElseIf Left(stDept, 1) = "S" Or (Left(stDept, 1) <> "S" And Val(strExc(1)) <> 0 And m_adoStaff.Fields("ST01") <> "94007") Then
                'end 2022/08/08
                    stValue(1) = "=IF(OR(" & Chr(GetValue("當月實績點數") + intField) & lngCounter & "=0," & Chr(GetValue("目標") + intField) & lngCounter & "=0),0," & Chr(GetValue("當月實績點數") + intField) & lngCounter & "/" & Chr(GetValue("目標") + intField) & lngCounter & ")"
                Else
                    '因非智權目標掛於總經理名下,而總經理目標以儲存格式設定不顯示,但有值,故直接設0
                    stValue(1) = "0"
                End If
        End Select
        'Modify by Amy 2022/07/12 M0100 目前沒大陸P/大陸T 可能有其他案號 ex:11106月有CFL資料,當計算 intChoose=34時,沒大陸P/大陸T不會計算,故拿掉stValue(1) <> MsgText(601)
        'If stValue(1) <> MsgText(601) Then
            If intChoose >= 31 Then
                '31 M0100 Total值
                If intChoose = 31 Then
                    If i <= 8 Then strM0100(i) = stValue(1)
                '32 M0100-P/33 M0100-T
                ElseIf intChoose <> 34 Then
                    If i <= 8 Then strM0100_C(i) = Val(strM0100_C(i)) + Val(stValue(1))
                    Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & lngCounter).Value = stValue(1)
                '34 M0100-大陸P-大陸T
                Else
                    If i <= 8 Then
                        Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & lngCounter).Value = Round(Val(strM0100(i)) - Val(strM0100_C(i)), 3)
                    Else
                        Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & lngCounter).Value = stValue(1)
                    End If
                End If
            Else
                Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & lngCounter).Value = stValue(1)
            End If
            If bolFormat = True Then
                Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & lngCounter).NumberFormatLocal = ";;;" '隱藏,但有值(ex:林總,不顯示於個人,顯示於其他合計)
                '避免輸錯,因10809 婉莘輸F4102 結餘總經理欄輸-15(輸錯),導致報出XX點數為負
                If i = GetValue("報出實績點數") Or i = GetValue("報出結餘點數") Then
                    Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & lngCounter).NumberFormatLocal = "#,##0;[紅色]-#,##0"
                End If
            ElseIf i = GetValue("目標") And GetValue("目標") <> 0 Then
                Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & lngCounter).NumberFormatLocal = "#,##0.00_ "
            ElseIf i = GetValue("達成率") And GetValue("達成率") <> 0 Then
                Wks.Range(Chr(GetValue(ReportFieldN2(i)) + intField) & lngCounter).NumberFormatLocal = "0.00%"
            End If
        'End If
    Next i
End Sub
'end 2016/11/22

'清變數值
Private Sub ClearVar()
    For i = LBound(strM0100) To UBound(strM0100_C)
        strM0100(i) = ""
        strM0100_C(i) = ""
    Next i
End Sub
 
'Add by Amy 2022/02/18 當月業績達成明細表 與前表連動 (將位置寫入暫存檔)
Private Sub ExcelSaveNew2_S2(XlsApp As Excel.Application, ByRef Wks As Worksheet, ByRef intCounter As Long)
    Dim rsA As New ADODB.Recordset
    Dim strA As String, strWhere As String, strFormula As String, strTarget As String, strAccomplish As String, strTp As String
    Dim intA As Integer, intStartR As Integer, intCountL As Integer
    Dim strOldDept As String 'Add by Amy 2022/08/05
    
    ReDim strFieldN2(0 To 3)
    ReDim intWidth2(0 To 3)
    
    strFieldN2 = Array("區別", "目標", "達成點數", "達成率")
    '11001月後分左右兩邊
    If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= 20210100 Then
        'Modify by Amy 2022/03/31 原欄位A=14,增加 P1005 專利國內部-MCP將欄位縮小,右邊由程式控制再加大
        intWidth2 = Array(11, 9, 9, 9)
    Else
        intWidth2 = Array(16, 15, 15, 16)
    End If
        
    strWhere = "And ID='" & strUserNum & "' And FormN='" & Me.Name & "' "
    intField = 65: intCounter = 2
    
    '抓取 位置 資料 (strRID=5Z=國內合計)
    'Modify by Amy 2023/07/27 +W3001,改Sort
    strA = "Select '1' as Sort,R002 as DeptN,R003 as mRow,R001 From RptXlsLink Where SubStr(R001,1,1)='S' " & strWhere & _
    " Union Select RID as Sort,R002 as DeptN,R003 as mRow,R001 From RptXlsLink Where SubStr(RID,1,1)>='2' And SubStr(RID,1,1)<='5' " & strWhere & _
    " Union Select '6' as Sort,R002 as DeptN,R003 as mRow,R001 From RptXlsLink Where Substr(R001,1,1)='F' " & strWhere & _
    " Union Select '8' as Sort,R002 as DeptN,R003 as mRow,R001 From RptXlsLink Where (RID='91' Or RID='ZZ') " & strWhere & _
    " Union Select '9' as Sort,R002 as DeptN,R003 as mRow,R001 From RptXlsLink Where RID='LZ' " & strWhere & _
    " Order by Sort,R001"
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, strA)
    If intA = 1 Then
        '欄位名稱
        For i = LBound(strFieldN2) To UBound(strFieldN2)
            Wks.Range(Chr(i + intField) & intCounter).Value = strFieldN2(i)
            Wks.Range(Chr(i + intField) & intCounter).ColumnWidth = intWidth2(i)
            Wks.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter
            Wks.Range(Chr(i + intField) & intCounter).Font.Color = vbBlue '藍色
            If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= 20210100 Then
                Wks.Range(Chr(i + UBound(strFieldN2) + 1 + intField) & intCounter).Value = strFieldN2(i)
                'Modify by Amy 2022/03/31 +P1005 專利國內部-MCP 欄寬加大
                If strFieldN2(i) = "區別" Then
                    Wks.Range(Chr(i + UBound(strFieldN2) + 1 + intField) & intCounter).ColumnWidth = intWidth2(i) + 5
                Else
                    Wks.Range(Chr(i + UBound(strFieldN2) + 1 + intField) & intCounter).ColumnWidth = intWidth2(i)
                End If
                Wks.Range(Chr(i + UBound(strFieldN2) + 1 + intField) & intCounter).HorizontalAlignment = xlCenter
                Wks.Range(Chr(i + UBound(strFieldN2) + 1 + intField) & intCounter).Font.Color = vbBlue '藍色
            End If
        Next i
        intCounter = intCounter + 1
        intStartR = intCounter
        
        '資料
        Do While rsA.EOF = False
            For i = LBound(strFieldN2) To UBound(strFieldN2)
                strFormula = ""
                '欄位設置右半邊
                'Modify by Amy 2022/08/05 11107月起無「客服組」資料,故無法設置右邊位置
                'If rsA.Fields("Sort") = "21" And i = LBound(strFieldN2) Then
                If Left(strOldDept, 1) = "S" And Left(strOldDept, 1) <> Left("" & rsA.Fields("R001"), 1) And i = LBound(strFieldN2) Then
                    intCountL = intCounter - 1
                    intField = intField + UBound(strFieldN2) + 1
                    intCounter = 3: intStartR = intCounter
                End If
                '資料
                If i = GetValue("區別", False) Then
                    strTp = "" & rsA.Fields("DeptN")
                ElseIf i = GetValue("達成率", False) Then
                    '公式
                    strTp = Chr(GetValue("達成點數", False) + intField) & intCounter & "/" & Chr(GetValue("目標", False) + intField) & intCounter
                    strTp = "=IF(" & Chr(GetValue("目標", False) + intField) & intCounter & "=0,""""," & strTp & ")"
                    strFormula = "0.00%"
                    'Modify by Amy 2024/07/09 改抓部門+z,台北所要加在北所合計,並增加台中所
                    'If (InStr(rsA.Fields("DeptN"), "所") > 0 Or InStr(rsA.Fields("DeptN"), "合計") > 0) And rsA.Fields("DeptN") <> "全所合計" Then
                    If Right(UCase(rsA.Fields("R001")), 1) = "Z" And rsA.Fields("R001") <> "ZZ" Then
                        intStartR = intCounter + 1
                    End If
                Else
                    '國內合計 / 全所合計
                    'Modify by Amy 2023/07/27 國內合計 原:2Z
                    If "" & rsA.Fields("R001") = "5Z" Or "" & rsA.Fields("R001") = "ZZ" Then
                        If i = GetValue("達成點數", False) Then
                            strTp = "=Sum(" & Mid(strAccomplish, 2) & ")"
                        Else
                            strTp = "=Sum(" & Mid(strTarget, 2) & ")"
                        End If
                        If "" & rsA.Fields("R001") = "5Z" And (i = GetValue("目標", False) Or i = GetValue("達成點數", False)) Then
                            strExc(1) = strFieldN2(i)
                            strTp = strTp & "+ Sum(" & Chr(GetValue(strExc(1), False) + intField) & intStartR & ":" & Chr(GetValue(strExc(1), False) + intField) & intCounter - 1 & ")"
                        End If
                    '各所合計 公式
                    ElseIf InStr("" & rsA.Fields("R001"), "z") > 0 Then
                        strTp = strFieldN2(i)
                        'Mark by Amy 2024/07/09 改抓部門+z,台北所要加在北所合計,並增加台中所
                        'If (InStr(rsA.Fields("DeptN"), "所") > 0 Or InStr(rsA.Fields("DeptN"), "合計") > 0) Then
                            If i = GetValue("達成點數", False) Then
                                strAccomplish = strAccomplish & "," & Chr(GetValue(strTp, False) + intField) & intCounter
                            '目標
                            Else
                                strTarget = strTarget & "," & Chr(GetValue(strTp, False) + intField) & intCounter
                            End If
                        'End If
                        strTp = "=Sum(" & Chr(GetValue(strTp, False) + intField) & intStartR & ":" & Chr(GetValue(strTp, False) + intField) & intCounter - 1 & ")"
                    '其他位置
                    Else
                        '前表位置
                        strTp = strFieldN2(i)
                        strExc(0) = strTp  '前表欄位名
                        '記錄加總位置 (M0109 安全基金撥補)
                        'Modify by Amy 2024/03/19 S10台北所有值,不會加到國內合計中
                        'Modify by Amy 2024/07/09 改抓部門+z,台北所要加在北所合計,並增加台中所
                        'If "" & rsA.Fields("R001") = "S10" Or "" & rsA.Fields("R001") = "S31" Or "" & rsA.Fields("R001") = "S41" Or "" & rsA.Fields("R001") = "M0109" Then
                        If (Left("" & rsA.Fields("R001"), 1) = "S" And Right("" & rsA.Fields("R001"), 1) = "z") Or "" & rsA.Fields("R001") = "S31" Or "" & rsA.Fields("R001") = "S41" _
                          Or "" & rsA.Fields("R001") = "M0109" Then
                            If i = GetValue("達成點數", False) Then
                                strAccomplish = strAccomplish & "," & Chr(GetValue(strTp, False) + intField) & intCounter
                            Else
                                strTarget = strTarget & "," & Chr(GetValue(strTp, False) + intField) & intCounter
                            End If
                        End If
                        
                        If strExc(0) = "達成點數" Then strExc(0) = "報出點數"
                        strTp = "=Round((" & strWkName & "1!" & Chr(GetValue(strExc(0)) + Asc("A")) & rsA.Fields("mRow") & ")/1000,2)"
                        '其他-國內包含其他-MCP,故需扣除
                        If "" & rsA.Fields("DeptN") = "其他-國內" And strM0100_P <> MsgText(601) Then
                            strTp = strTp & "- Round((" & strWkName & "1!" & Chr(GetValue(strExc(0)) + Asc("A")) & strM0100_P & ")/1000,2)"
                        End If
                    End If
                    strFormula = "#,##0"
                End If
                
                'Add by Amy 2022/08/05 因11107月起無「客服組」資料,右邊少一列,調整「全所合計」位置
                If "" & rsA.Fields("R001") = "ZZ" And intCounter < intCountL Then
                    intCounter = intCountL
                End If
                Wks.Range(Chr(i + intField) & intCounter).Value = strTp
                
                If strFormula <> MsgText(601) Then
                    Wks.Range(Chr(i + intField) & intCounter).NumberFormatLocal = strFormula
                End If
                '設定顏色
                'Modify by Amy 2024/07/09 改抓部門R001欄位
                If i = GetValue("達成率", False) Then
                    'If InStr(rsA.Fields("DeptN"), "所") > 0 And rsA.Fields("DeptN") <> "全所合計" And rsA.Fields("DeptN") <> "法律所合計" Then
                    If ((Left("" & rsA.Fields("R001"), 1) = "S" And Right("" & rsA.Fields("R001"), 1) = "z") Or "" & rsA.Fields("R001") = "S31" Or "" & rsA.Fields("R001") = "S41") Then
                        Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN2) + intField) & intCounter).Interior.ColorIndex = 40 '膚色
                        Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN2) + intField) & intCounter).Interior.tintandshade = 0.2  '設深淺
                    'ElseIf rsA.Fields("DeptN") = "國內合計" Or rsA.Fields("DeptN") = "國外合計" Then
                    ElseIf rsA.Fields("R001") = "5Z" Or rsA.Fields("R001") = "Fz" Then
                        Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN2) + intField) & intCounter).Interior.ColorIndex = 6 '淺黃色
                        '記錄國內合計位置,全所合計=國內合計+國外合計+全基金撥補
                        'Modify by Amy 2024/07/09 改抓部門R001欄位
                        'If "" & rsA.Fields("DeptN") = "國內合計" Then
                        If "" & rsA.Fields("R001") = "5Z" Then
                            strAccomplish = "," & Chr(GetValue("達成點數", False) + intField) & intCounter
                            strTarget = "," & Chr(GetValue("目標", False) + intField) & intCounter
                        End If
                    'ElseIf rsA.Fields("DeptN") = "全所合計" Then
                    ElseIf rsA.Fields("R001") = "ZZ" Then
                    'end 2024/07/09
                        Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN2) + intField) & intCounter).Font.Color = vbBlue
                        Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN2) + intField) & intCounter).Font.Bold = True
                        '畫框
                        If Left(MaskEdBox1, 6) = Left(MaskEdBox2, 6) Then
                            strTp = Val(Left(MaskEdBox1, 3)) & "年" & Mid(MaskEdBox1, 5, 2) & "月份業績達成明細表"
                        Else
                            strTp = Mid(MaskEdBox1, 1, 6) & "~" & Mid(MaskEdBox2, 1, 6) & "月業績達成明細表"
                        End If
                        Wks.Range("A1").Value = strTp
                        Wks.Range("A1" & ":" & Chr(UBound(strFieldN2) + intField) & "1").HorizontalAlignment = xlCenter
                        Wks.Range("A1" & ":" & Chr(UBound(strFieldN2) + intField) & "1").MergeCells = True
                        
                        Wks.Activate 'Memo 需先設此,否則.Select 會錯
                        Wks.Range("A1:" & Chr(UBound(strFieldN2) + intField) & IIf(intCountL >= intCounter, intCountL, intCounter)).Select
                        XlsApp.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
                        XlsApp.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
                        XlsApp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
                        XlsApp.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
                        XlsApp.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
                        XlsApp.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                        '設定L公司顯示
                        intField = 65
                        intCounter = intCounter + 1
                    'Modify by Amy 2024/07/09 改抓部門R001欄位
                    'ElseIf rsA.Fields("DeptN") = "法律所合計" Then
                    ElseIf rsA.Fields("R001") = "LZ" Then
                        Wks.Range("A" & intCounter & ":" & Chr(UBound(strFieldN2) + intField) & intCounter).Select
                        XlsApp.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
                        XlsApp.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
                        XlsApp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
                        XlsApp.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
                        XlsApp.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
                        XlsApp.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    End If
                End If '設定顏色
                   
            Next i
            strOldDept = "" & rsA.Fields("R001") 'Add by Amy 2022/08/05
            intCounter = intCounter + 1
            rsA.MoveNext
        Loop
    End If
End Sub

'Add by Amy 2018/05/17 當月業績達成明細表 (暫存檔加總寫入)
'Modify by Amy 2020/11/09 +XlsApp,stCmp
Private Sub ExcelSaveNew2_S2_Old(XlsApp As Excel.Application, ByRef Wks As Worksheet, ByVal stM0100_P As String, stM0100_T As String, ByRef intCounter As Long, ByVal stCmp As String)
'    Dim strQ As String, strTp As String, strOldZone As String, strTotal As String, strOther As String, strFormula As String
'    Dim intStartR As Integer
'    Dim intCountL As Integer, intCountR As Integer 'Add by Amy 2020/11/09 左邊列數
'    'Add by Amy 2021/02/19 增加顧服/客服/專利國外/專利日本/FCT英文/FCT日文,MCT 11004月後月目標
'    Dim strWhere As String, strQ1 As String, strF1 As String, strF2 As String, strGroup As String, strMCT As String
'
'    ReDim strFieldN(0 To 3)
'    ReDim intWidth(0 To 3)
'
'    strFieldN = Array("區別", "目標", "達成點數", "達成率")
'    intWidth = Array(11, 9, 9, 9) 'Modify by Amy 2020/11/09 原:Array(16, 15, 15, 16)
'    intField = 65: intCounter = 2
'
'    For i = LBound(strFieldN) To UBound(strFieldN)
'        Wks.Range(Chr(i + intField) & intCounter).Value = strFieldN(i)
'        Wks.Range(Chr(i + intField) & intCounter).ColumnWidth = intWidth(i)
'        Wks.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter
'        'Add by Amy 2020/11/09
'        Wks.Range(Chr(i + intField) & intCounter).Font.Color = vbBlue '藍色
'        Wks.Range(Chr(i + UBound(strFieldN) + 1 + intField) & intCounter).Value = strFieldN(i)
'        Wks.Range(Chr(i + UBound(strFieldN) + 1 + intField) & intCounter).ColumnWidth = intWidth(i)
'        Wks.Range(Chr(i + UBound(strFieldN) + 1 + intField) & intCounter).HorizontalAlignment = xlCenter
'        Wks.Range(Chr(i + UBound(strFieldN) + 1 + intField) & intCounter).Font.Color = vbBlue '藍色
'        'end 2020/11/09
'    Next i
'    intCounter = intCounter + 1
'    intStartR = intCounter
'
'    '智權部
'    'Modify by Amy 2020/11/09 原And SubStr(R001,1,1)='S' 智權部先顯示
'    'Modify by Amy 2020/11/12 表二值先除1000再四捨五入與表一合計值不一致,故改取小數兩位 原:Round(...,0)
'    'Modify by Amy 2021/01/19 +F4104~07
'    'Modify by Amy 2021/02/19 左邊顯示智權部
'    strWhere = "And R001>'S00' And R001<'S30'"
'    If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= 20210100 Then
'        strWhere = "And SubStr(R001,1,1)='S' "
'    End If
'    'Modify by Amy 2021/05/27 +M0109 列於「全所合計」之前
'    strQ = "Select A0902,Round(Sum(Nvl(R012,0))/1000,2) as PE04,Round(Sum((Nvl(R004,0)+Nvl(R006,0)-Nvl(R008,0)+Nvl(R010,0))+(Nvl(R005,0)+Nvl(R007,0)-Nvl(R009,0)+Nvl(R011,0)))/1000,2) as Point,R001 as SP48 " & _
'                "From Accrpt44j0_2,Acc090 " & _
'                "Where ID='" & strUserNum & "' " & strWhere & " And R003 Not in ('M0100','M0109','F4101','F4102','F4103','F4104','F4105','F4106','F4107') And R001=a0901(+) " & _
'                "Group by A0902,R001 " & _
'    "Union Select '北所合計',0,0,'S1Z' as SP48 From Dual " & _
'    "Union Select '中所合計',0,0,'S2Z' as SP48 From Dual " & _
'                " Order by SP48 "
'    'end 2021/02/19
'    If adostaff.State <> adStateClosed Then adostaff.Close
'    adostaff.CursorLocation = adUseClient
'    adostaff.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
'    Do While adostaff.EOF = False
'        For i = LBound(strFieldN) To UBound(strFieldN)
'            If i = GetValue("達成率") Then
'                '公式
'                strTp = "=" & Chr(GetValue("達成點數") + intField) & intCounter & "/" & Chr(GetValue("目標") + intField) & intCounter
'            ElseIf i <> GetValue("區別") And Right(adostaff.Fields("A0902"), 3) = "所合計" Then
'                strTp = "Sum(" & Chr(i + intField) & intStartR & ":" & Chr(i + intField) & intCounter - 1 & ")"
'                If i = GetValue("目標") Then strTotal = strTotal & "," & Chr(i + intField) & intCounter
'                If i = GetValue("達成點數") Then intStartR = intCounter + 1
'            Else
'                strTp = adostaff.Fields(i)
'                'Mark by Amy 2020/11/09 台南/高雄所,移至下方
'                'Modify by Amy 2021/02/19 +if 11001月後智權顯示左方,非智權顯示右方(原:北中所顯示左方,南高所和非智權顯示右方),並增加顧服/客服/專利國外/專利日本/FCT英文/FCT日文
'                If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= 20210100 Then
'                    If i = GetValue("目標") And (adostaff.Fields("A0902") = "台南所" Or adostaff.Fields("A0902") = "高雄所") Then
'                        strTotal = strTotal & "," & Chr(i + intField) & intCounter
'                    End If
'                End If
'            End If
'            Wks.Range(Chr(i + intField) & intCounter).Value = IIf(Left(strTp, 1) = "S", "=" & strTp, strTp)
'            If Left(strTp, 1) = "=" Then
'                Wks.Range(Chr(i + intField) & intCounter).NumberFormatLocal = "0.00%"
'            'Add by Amy 2020/11/09
'            Else
'                Wks.Range(Chr(i + intField) & intCounter).NumberFormatLocal = "#,##0"
'            End If
'        Next i
'        'Add by Amy 2020/11/09 「所合計」變色
'        'Modify by Amy 2021/02/19 +台南所/高雄所
'        If Right("" & adostaff.Fields("A0902"), 3) = "所合計" Or adostaff.Fields("A0902") = "台南所" Or adostaff.Fields("A0902") = "高雄所" Then
'            Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter).Interior.ColorIndex = 40  '膚色
'            Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter).Interior.tintandshade = 0.2 '設深淺
'        End If
'        strOldZone = Left(adostaff.Fields("A0902"), 1)
'        adostaff.MoveNext
'        intCounter = intCounter + 1
'    Loop
'    'Add by Amy 2020/11/09 分為左、右兩邊顯示
'    intCountL = intCounter - 1
'    intField = 65 + UBound(strFieldN) + 1
'    intCounter = 3: intStartR = intCounter
'    'end 2020/11/09
'
'    'Modify by Amy 2020/04/22 畫面「傳票日期」起日若為「智慧所更名日」前,維持舊格式 ex:1090101~1091231
'    '                                                                                             「智慧所更名日」後,L部門人員獨立顯示於最後區塊
'    'Add by Amy 2021/02/19 +if 11001月起,增加顧服/客服/專利國外/專利日本/FCT英文/FCT日文
'    If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= 20210100 Then
'         'Modify by Amy 2021/05/13 11004月後排除MCT
'        strExc(0) = ""
'        If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= 20210400 Then strExc(0) = ",'P2005'"
'        'Modify by Amy 2021/05/27 +M0109 有資料才顯示於「全所合計」之前
'        strQ = "Union Select Dept,Pe04,Point,Sort From (" & _
'                        "Select st02 as Dept,Round(Sum(Nvl(R012,0))/1000,2) as PE04,Round(Sum((Nvl(R004,0)+Nvl(R006,0)-Nvl(R008,0)+Nvl(R010,0))+(Nvl(R005,0)+Nvl(R007,0)-Nvl(R009,0)+Nvl(R011,0)))/1000,2) as Point,'T0' as Sort,st01 " & _
'                        "From Accrpt44j0_2,Staff Where ID='" & strUserNum & "' And SubStr(R001,1,1)='W' And R003=St01(+) Group by st02,st01 " & _
'                        "Union Select '其他合計' as Dept,Round(Sum(Nvl(R012,0))/1000,2) as PE04,Round(Sum((Nvl(R004,0)+Nvl(R006,0)-Nvl(R008,0)+Nvl(R010,0))+(Nvl(R005,0)+Nvl(R007,0)-Nvl(R009,0)+Nvl(R011,0)))/1000,2) as Point,'T4' as Sort,'T4' " & _
'                        "From Accrpt44j0_2,Staff Where ID='" & strUserNum & "' And SubStr(R001,1,1)<>'S' And Substr(St03,1,1)<>'L' And R003 Not In('M0109','W1001','W2001','F4101','F4102','F4103','F4104','F4105','F4106','F4107'" & strExc(0) & ") And R003=St01(+) " & _
'                        "Order by Sort,st01" & _
'                    ")"
'        'end 2021/05/13
'    'Modify by Amy 2020/11/09 拆成兩邊顯示,改為文字排序,原:1 as Sort
'    'Modify by Amy 2020/11/12 表二值先除1000再四捨五入與表一合計值不一致,故改取小數兩位 原:Round(...,0)
'    ElseIf Val(FCDate(MaskEdBox1.Text)) + 19110000 >= Val(智慧所更名日) Then
'        '排除L部門人員
'        'Modify by Amy 2021/01/19 +F4104~07
'        strQ = "Union Select '其他合計' as Dept,Round(Sum(Nvl(R012,0))/1000,2) as PE04,Round(Sum((Nvl(R004,0)+Nvl(R006,0)-Nvl(R008,0)+Nvl(R010,0))+(Nvl(R005,0)+Nvl(R007,0)-Nvl(R009,0)+Nvl(R011,0)))/1000,2) as Point,'T4' as Sort " & _
'                    "From Accrpt44j0_2,Staff Where ID='" & strUserNum & "' And SubStr(R001,1,1)<>'S' And R003 Not in ('F4101','F4102','F4103','F4104','F4105','F4106','F4107') " & _
'                    " And R003=St01(+) And Substr(St03,1,1)<>'L' "
'    Else
'        strQ = "Union Select '其他合計' as Dept,Round(Sum(Nvl(R012,0))/1000,2) as PE04,Round(Sum((Nvl(R004,0)+Nvl(R006,0)-Nvl(R008,0)+Nvl(R010,0))+(Nvl(R005,0)+Nvl(R007,0)-Nvl(R009,0)+Nvl(R011,0)))/1000,2) as Point,'T4' as Sort " & _
'                   "From Accrpt44j0_2 Where ID='" & strUserNum & "' And SubStr(R001,1,1)<>'S' And R003 Not in ('F4101','F4102','F4103') "
'
'    End If
'
'    'Modify by Amy 2021/01/19 11001月始F4102拆成F4104,F4105/F4103拆成F4106,F4107,但舊資料仍有F4102/F4103
'                                                 '11004 MCT 有目標
'    If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= 20210400 Then
'        strMCT = "Union Select st02 as Dept,Round(Sum(Nvl(R012,0))/1000,2) as PE04,Round(Sum((Nvl(R004,0)+Nvl(R006,0)-Nvl(R008,0)+Nvl(R010,0))+(Nvl(R005,0)+Nvl(R007,0)-Nvl(R009,0)+Nvl(R011,0)))/1000,2) as Point,'T4' as Sort " & _
'                       "From Accrpt44j0_2,Staff Where ID='" & strUserNum & "' And R003='P2005' And R003=St01(+) Group by st02 "
'    Else
'        strMCT = "Union Select '其他-MCT' as Dept,0  as PE04," & Round(Val(stM0100_T) / 1000, 2) & " as Point,'T1' as Sort From Dual "
'    End If
'    strQ = strMCT & _
'              "Union Select '其他-MCP' as Dept,0  as PE04," & Round(Val(stM0100_P) / 1000, 2) & " as Point,'T2' as Sort From Dual " & _
'              "Union Select '其他-國內' as Dept,0  as PE04,0 as Point,'T3' as Sort From Dual " & _
'              strQ & _
'              "Union Select '國內合計' as Dept,0  as PE04,0 as Point,'TZ' as Sort From Dual "
'    'Modify by Amy 2021/02/19 F41字頭編11001月後獨立列示 原:加總列示
'    If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= 20210100 Then
'        strF1 = "st02"
'        strF2 = "st02"
'        strGroup = " Group by st02 "
'    Else
'        strF1 = "'FCP'"
'        strF2 = "'FCT'"
'    End If
'    strQ = strQ & _
'              "Union Select " & strF1 & " as Dept,Round(Sum(Nvl(R012,0))/1000,2) as PE04,Round(Sum((Nvl(R004,0)+Nvl(R006,0)-Nvl(R008,0)+Nvl(R010,0))+(Nvl(R005,0)+Nvl(R007,0)-Nvl(R009,0)+Nvl(R011,0)))/1000,2) as Point,'U1' as Sort " & _
'                          "From Accrpt44j0_2,Staff Where ID='" & strUserNum & "' And R003 In('F4102','F4104','F4105') And R003=St01(+) " & strGroup & _
'              "Union Select " & strF2 & " as Dept,Round(Sum(Nvl(R012,0))/1000,2) as PE04,Round(Sum((Nvl(R004,0)+Nvl(R006,0)-Nvl(R008,0)+Nvl(R010,0))+(Nvl(R005,0)+Nvl(R007,0)-Nvl(R009,0)+Nvl(R011,0)))/1000,2) as Point,'U2' as Sort " & _
'                          "From Accrpt44j0_2,Staff Where ID='" & strUserNum & "' And R003 In('F4103','F4106','F4107') And R003=St01(+) " & strGroup & _
'              "Union Select 'FCL' as Dept,Round(Sum(Nvl(R012,0))/1000,2) as PE04,Round(Sum((Nvl(R004,0)+Nvl(R006,0)-Nvl(R008,0)+Nvl(R010,0))+(Nvl(R005,0)+Nvl(R007,0)-Nvl(R009,0)+Nvl(R011,0)))/1000,2) as Point,'U3' as Sort " & _
'                          "From Accrpt44j0_2 Where ID='" & strUserNum & "' And R003='F4101' "
'    'Modify by Amy 2021/05/27 +M0109 有資料才顯示於「全所合計」之前
'    strQ = strQ & _
'              "Union Select '國外合計' as Dept,0  as PE04,0 as Point,'UZ' as Sort From Dual " & _
'              "Union Select st02 as Dept,Round(Sum(Nvl(R012,0))/1000,2) as PE04,Round(Sum((Nvl(R004,0)+Nvl(R006,0)-Nvl(R008,0)+Nvl(R010,0))+(Nvl(R005,0)+Nvl(R007,0)-Nvl(R009,0)+Nvl(R011,0)))/1000,2) as Point,'V1' as Sort " & _
'                          "From Accrpt44j0_2,Staff Where ID='" & strUserNum & "' And R003='M0109' And R003=St01(+) Group by st02 " & _
'              "Union Select '全所合計' as Dept,0  as PE04,0 as Point,'Z' as Sort From Dual "
'    'end 2021/02/19
'    'Modify by Amy 2021/02/19 11001月前 南高所放左邊
'    If Val(FCDate(MaskEdBox1.Text)) + 19110000 < 20210100 Then
'        strQ1 = "Select A0902 as Dept,Round(Sum(Nvl(R012,0))/1000,2) as PE04,Round(Sum((Nvl(R004,0)+Nvl(R006,0)-Nvl(R008,0)+Nvl(R010,0))+(Nvl(R005,0)+Nvl(R007,0)-Nvl(R009,0)+Nvl(R011,0)))/1000,2) as Point,R001 as Sort " & _
'               "From Accrpt44j0_2,Acc090 " & _
'               "Where ID='" & strUserNum & "' And SubStr(R001,1,1)='S' And R001>='S30' And R001=a0901(+) Group by A0902,R001 " & strQ & " "
'    Else
'        strQ = Mid(strQ, 7)
'    End If
'    strQ = "Select * From(" & strQ1 & " " & strQ & ") Where PE04 is not null And Point is not null Order by Sort "
'    'end 2021/02/19
'    'end 2021/01/19
'    'end 2020/11/12
'    'end 2020/11/09
'    'end 2020/04/22
'
'    If adostaff.State <> adStateClosed Then adostaff.Close
'    adostaff.CursorLocation = adUseClient
'    adostaff.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
'    Do While adostaff.EOF = False
'        For i = LBound(strFieldN) To UBound(strFieldN)
'            strTp = ""
'            'Add by Amy 2021/02/19 W1001/W2001/P2005 加入國內合計
'            'Modify by Amy 2021/05/27 +M0109 加入全所合計
'            If i = GetValue("目標") And ("" & adostaff.Fields("Dept") = StaffQuery("W1001") Or "" & adostaff.Fields("Dept") = StaffQuery("W2001") Or "" & adostaff.Fields("Dept") = StaffQuery("P2005") Or "" & adostaff.Fields("Dept") = StaffQuery("M0109")) Then
'                '國內/全所 合計公式
'                strTotal = strTotal & "," & Chr(i + intField) & intCounter
'            End If
'            If i = GetValue("區別") Then
'                'If adostaff.Fields("Dept") = "國外合計" Then intCounter = intCounter + 1
'                strTp = "" & adostaff.Fields("Dept")
'            '其他-XX
'            ElseIf InStr(adostaff.Fields("Dept"), "-") > 0 Then
'                Wks.Range(Chr(i + intField) & intCounter).Font.ColorIndex = 3 '紅色
'                If i = GetValue("達成點數") Then
'                    If adostaff.Fields("Dept") = "其他-國內" Then
'                        strOther = Chr(i + intField) & intCounter '記錄位置
'                    Else
'                        strTp = adostaff.Fields("Point")
'                        strFormula = strFormula & "," & Chr(i + intField) & intCounter
'                    End If
'                End If
'            ElseIf i = GetValue("達成率") Then
'                '公式
'                If "" & adostaff.Fields("Dept") = StaffQuery("M0109") Then
'                    'M0109安全基金撥補,沒目標不必算達成率
'                Else
'                    strTp = "=" & Chr(GetValue("達成點數") + intField) & intCounter & "/" & Chr(GetValue("目標") + intField) & intCounter
'                End If
'            ElseIf adostaff.Fields("Dept") = "其他合計" Then
'                strTp = "" & adostaff.Fields(i)
'                If i = GetValue("目標") Then
'                    '國內合計公式
'                    strTotal = strTotal & "," & Chr(i + intField) & intCounter
'                Else
'                    '計算「其他-國內」值
'                    strFormula = Chr(i + intField) & intCounter & "-Sum(" & Mid(strFormula, 2) & ")"
'                    Wks.Range(strOther).Value = "=" & strFormula
'                End If
'            '合計
'            ElseIf InStr(adostaff.Fields("Dept"), "合計") > 0 Then
'                '公式
'                Select Case adostaff.Fields("Dept")
'                    Case "國內合計"
'                        If i = GetValue("達成點數") Then
'                            'Modify by Amy 2020/11/09 拆成兩邊
'                            strTp = Replace(strTotal, Chr(GetValue("目標") + 65), Chr(GetValue("目標") + 66))
'                            strTp = Mid(Replace(strTp, Chr(GetValue("目標") + intField), Chr(i + intField)), 2)
'                            'end 2020/11/09
'                        Else
'                            strTp = Mid(strTotal, 2)
'                        End If
'                        strTp = "=Sum(" & strTp & ")"
'                        intStartR = intCounter + 1#
'                        '全所合計公式
'                        If i = GetValue("達成點數") Then strTotal = "," & Chr(GetValue("目標") + intField) & intCounter
'                    Case "國外合計"
'                        strTp = "=Sum(" & Chr(i + intField) & intStartR & ":" & Chr(i + intField) & intCounter - 1 & ")"
'                        If i = GetValue("目標") Then strTotal = strTotal & "," & Chr(i + intField) & intCounter
'                    Case "全所合計"
'                        If i = GetValue("達成點數") Then
'                            'Modify by Amy 2020/11/09 拆成兩邊
'                            strTp = Replace(strTotal, Chr(GetValue("目標") + 65), Chr(GetValue("目標") + 66))
'                            strTp = Mid(Replace(strTp, Chr(GetValue("目標") + intField), Chr(i + intField)), 2)
'                            'end 2020/11/09
'                        Else
'                            strTp = Mid(strTotal, 2)
'                        End If
'                        strTp = "=Sum(" & strTp & ")"
'                End Select
'            Else
'                strTp = "" & adostaff.Fields(i)
'                'Modify by Amy 2020/11/09 台南/高雄所從上面搬下來修改(Memo 11001月前用)
'                If i = GetValue("目標") And (adostaff.Fields("Dept") = "台南所" Or adostaff.Fields("Dept") = "高雄所") Then
'                    strTotal = strTotal & "," & Chr(i + intField) & intCounter
'                End If
'                'end 2020/11/09
'            End If
'            Wks.Range(Chr(i + intField) & intCounter).Value = strTp
'            'Add by Amy 2020/11/09
'            If i = GetValue("達成率") Then
'                Wks.Range(Chr(i + intField) & intCounter).NumberFormatLocal = "0.00%"
'            Else
'                Wks.Range(Chr(i + intField) & intCounter).NumberFormatLocal = "#,##0"
'            End If
'            'end 2020/11/09
'        Next i
'        'Add by Amy 2020/11/09
'        If adostaff.Fields("Dept") = "台南所" Or adostaff.Fields("Dept") = "高雄所" Or adostaff.Fields("Dept") = "FCP" Or adostaff.Fields("Dept") = "FCT" Then
'            Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter).Interior.ColorIndex = 40 '膚色
'            Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter).Interior.tintandshade = 0.2  '設深淺
'        ElseIf adostaff.Fields("Dept") = "國內合計" Or adostaff.Fields("Dept") = "國外合計" Then
'            Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter).Interior.ColorIndex = 6 '淺黃色
'        ElseIf adostaff.Fields("Dept") = "全所合計" Then
'            Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter).Font.Color = vbBlue
'            Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strFieldN) + intField) & intCounter).Font.Bold = True
'        End If
'        'end 2020/11/09
'        intCounter = intCounter + 1
'        adostaff.MoveNext
'    Loop
'    intCountR = intCounter - 1 'Add by Amy 2020/11/09
'
'    'Modify by Amy 2020/11/09 框線 由ExcelSaveNew2搬過來修改
'    Wks.PageSetup.TopMargin = XlsApp.InchesToPoints(0.4)
'    Wks.PageSetup.BottomMargin = XlsApp.InchesToPoints(0.4)
'    Wks.Range("A1:" & Chr(UBound(strFieldN) + intField) & IIf(intCountL >= intCountR, intCountL, intCountR)).Select
'    XlsApp.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'    XlsApp.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'    XlsApp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
'    XlsApp.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
'    XlsApp.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
'    XlsApp.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    'end 2020/11/09
'
'    If Left(MaskEdBox1, 6) = Left(MaskEdBox2, 6) Then
'        strTp = Val(Left(MaskEdBox1, 3)) & "年" & Mid(MaskEdBox1, 5, 2) & "月份業績達成明細表"
'    Else
'        'Modify by Amy 2021/04/08 原:Mid(MaskEdBox1, 1, 5) 抓錯
'        strTp = Mid(MaskEdBox1, 1, 6) & "~" & Mid(MaskEdBox2, 1, 6) & "月業績達成明細表"
'    End If
'    Wks.Range("A1").Value = strTp
'    Wks.Range("A1" & ":" & Chr(UBound(strFieldN) + intField) & "1").HorizontalAlignment = xlCenter
'    Wks.Range("A1" & ":" & Chr(UBound(strFieldN) + intField) & "1").MergeCells = True
'
'    'Add by Amy 2020/04/22 +法律所合計
'    If Val(FCDate(MaskEdBox1.Text)) + 19110000 >= Val(智慧所更名日) And (stCmp = MsgText(601) Or stCmp = "L") Then
'        intCounter = IIf(intCountL >= intCountR, intCountL, intCountR)  'Add by Amy 2020/11/09
'        intCounter = intCounter + 3
'        intField = 65  'Add by Amy 2020/11/09
'
'        strQ = "Select '法律所合計' as Dept,Round(Sum(Nvl(R012,0))/1000,2) as PE04,Round(Sum((Nvl(R004,0)+Nvl(R006,0)-Nvl(R008,0)+Nvl(R010,0))+(Nvl(R005,0)+Nvl(R007,0)-Nvl(R009,0)+Nvl(R011,0)))/1000,2) as Point " & _
'                  "From Accrpt44j0_2,Staff Where ID='" & strUserNum & "' And R003=St01(+) And Substr(St03,1,1)='L' "
'        If adostaff.State <> adStateClosed Then adostaff.Close
'        adostaff.CursorLocation = adUseClient
'        adostaff.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
'        Do While adostaff.EOF = False
'            adostaff.MoveFirst
'            For i = LBound(strFieldN) To UBound(strFieldN)
'                strTp = ""
'                If i = GetValue("達成率") Then
'                    If Val("" & adostaff.Fields(GetValue("達成點數"))) <> 0 And Val("" & adostaff.Fields(GetValue("目標"))) <> 0 Then
'                        strTp = "=" & Chr(GetValue("達成點數") + intField) & intCounter & "/" & Chr(GetValue("目標") + intField) & intCounter
'                    End If
'                    Wks.Range(Chr(i + intField) & intCounter).NumberFormatLocal = "0.00%"
'                Else
'                    strTp = "" & adostaff.Fields(i)
'                End If
'                Wks.Range(Chr(i + intField) & intCounter).Value = strTp
'            Next i
'            intCounter = intCounter + 1
'            adostaff.MoveNext
'        Loop
'        'Add by Amy 2020/11/09 框線 由ExcelSaveNew2搬過來修改
'        Wks.Range("A" & intCounter - 1 & ":" & Chr(UBound(strFieldN) + intField) & intCounter - 1).Select
'        XlsApp.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
'        XlsApp.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
'        XlsApp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
'        XlsApp.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
'        XlsApp.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
'        XlsApp.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
'        'end 2020/11/09
'    End If
'
'    'Modify by Amy 2020/04/22 只設一次
'    If Val(FCDate(MaskEdBox1.Text)) + 19110000 < Val(智慧所更名日) Then
'        Wks.Name = "業績達成明細表"
'        '設定
'        Wks.PageSetup.PaperSize = 9 'A4
'        Wks.PageSetup.Orientation = xlPortrait '直印
'        Wks.PageSetup.LeftMargin = 0 '邊界
'        Wks.PageSetup.RightMargin = 0
'        Wks.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
'    End If
End Sub

Private Function ChkFCPFCTSum(Optional ByVal stSQL As String = "") As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    ChkFCPFCTSum = False
    If stSQL = "" Then
        strQ = "Select Max(Count(R001)) Cnt From AccRpt44J0_2 Where id='" & strUserNum & "' " & _
                    "and R003 in ('F4101','F4102','F4104','F4105','F4103','F4106','F4107') Group by R001 "
    Else
        strQ = "Select Max(Count(H1)) Cnt From (" & stSQL & ") Group by H1"
    End If
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        If Val("" & RsQ.Fields("Cnt")) > 1 Then ChkFCPFCTSum = True
    End If
    Set RsQ = Nothing
End Function

'Add by Amy 2022/02/18 位置寫入暫存檔
Private Function InsRptXlsLink(ByVal stNo As String, ByVal stName As String, ByVal stRow As String, Optional ByVal strID As String = "") As Boolean
    Dim stIns As String, intQ As Integer
    
    InsRptXlsLink = False
    If strID = MsgText(601) Then
        stIns = "Insert Into RptXlsLink (ID,FormN,R001,R002,R003) Values" & _
                    "('" & strUserNum & "','" & Me.Name & "','" & stNo & "','" & ChgSQL(stName) & "','" & stRow & "')"
    Else
        stIns = "Insert Into RptXlsLink (ID,FormN,RID,R001,R002,R003) Values" & _
                    "('" & strUserNum & "','" & Me.Name & "','" & strID & "','" & stNo & "','" & ChgSQL(stName) & "','" & stRow & "')"
    End If
    cnnConnection.Execute stIns, intQ
    
    If intQ = 0 Then Exit Function
    
    InsRptXlsLink = True
End Function

