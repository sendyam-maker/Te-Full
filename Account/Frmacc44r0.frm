VERSION 5.00
Begin VB.Form Frmacc44r0 
   AutoRedraw      =   -1  'True
   Caption         =   "專業點數明細表"
   ClientHeight    =   3696
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3696
   ScaleWidth      =   5160
   Begin VB.CheckBox ChkD5 
      Caption         =   "顯示小數5位"
      Height          =   210
      Left            =   2040
      TabIndex        =   20
      Top             =   720
      Width           =   1300
   End
   Begin VB.CheckBox ChkChoose 
      Caption         =   "顯示結餘+轉撥"
      Height          =   210
      Left            =   3420
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   1500
   End
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
      Left            =   1290
      TabIndex        =   15
      Top             =   480
      Width           =   3500
   End
   Begin VB.TextBox Txt1 
      Alignment       =   1  '靠右對齊
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
      Index           =   2
      Left            =   3450
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1470
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.TextBox Txt1 
      Alignment       =   1  '靠右對齊
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
      Index           =   1
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1470
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.TextBox Txt1 
      Alignment       =   1  '靠右對齊
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
      Index           =   0
      Left            =   1110
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1470
      Visible         =   0   'False
      Width           =   612
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
      Left            =   240
      TabIndex        =   0
      Top             =   60
      Width           =   4000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   4320
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   10
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel檔(&P)"
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
      TabIndex        =   13
      Top             =   3255
      Width           =   4692
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
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
      MaxLength       =   2
      TabIndex        =   2
      Top             =   870
      Width           =   612
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
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
      Left            =   1296
      MaxLength       =   3
      TabIndex        =   1
      Top             =   870
      Width           =   612
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "10912月以前為舊格式勾選「顯示小數5位」無作用"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   288
      Left            =   10
      TabIndex        =   21
      Top             =   10
      Visible         =   0   'False
      Width           =   4908
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "只產生智慧所+智權公司資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1140
      TabIndex        =   19
      Top             =   990
      Visible         =   0   'False
      Width           =   4905
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "lable8"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1100
      Left            =   120
      TabIndex        =   17
      Top             =   1950
      Visible         =   0   'False
      Width           =   4900
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "執行時請勿開啟Excel"
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
      Left            =   1050
      TabIndex        =   16
      Top             =   420
      Width           =   4005
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "月"
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
      Left            =   4200
      TabIndex        =   12
      Top             =   1530
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "月 ~"
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
      Left            =   2910
      TabIndex        =   11
      Top             =   1530
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "年"
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
      Left            =   1800
      TabIndex        =   10
      Top             =   1530
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "統計："
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
      Left            =   330
      TabIndex        =   8
      Top             =   1530
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.Line Line1 
      X1              =   420
      X2              =   4770
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1128
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "月份"
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
      Left            =   2010
      TabIndex        =   5
      Top             =   870
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "年度"
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
      Left            =   330
      TabIndex        =   4
      Top             =   870
      Width           =   615
   End
   Begin VB.Label Label3 
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
      Left            =   330
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc44r0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit
Dim adoquery As New ADODB.Recordset
Dim strSql As String
'Add by Amy 2015/07/28
Dim strField1(), strField2(), intWidth1(), intWidth2()
Dim bolXlsOpen As Boolean
Dim dblLawVal(8) As Double '法務資料值 Modify by Amy 2017/06/28 原:6 +CFP收入 +一般法務
Dim strLRow 'Add by Amy 2017/06/28 for 左表對應右表順序
'Add by Amy 2019/05/13
Dim rsNew As New ADODB.Recordset
Dim i As Integer, bolFirst As Boolean
Dim strYear As Integer, strStartMonth As String, strEndMonth As String '年/起始月/結束月
Dim strCmp As String, strCmpN As String 'Add by Amy 2020/06/05
'Add by Amy 2021/02/26
Dim intField As Integer, intField2 As Integer, intRow1 As Integer, intRow2 As Integer, intLStart As Integer
Public stState As String 'Add by Amy 2021/09/23
Dim IsExcelSave As Boolean 'Add by Amy 2024/06/26 有產生系統月-1 專業點數明細表

Private Sub cmdok_Click()
    Unload Me
End Sub

'Add by Amy 2019/05/13
Private Sub Combo1_Click()
    Dim oText As TextBox
    
    For Each oText In Txt1
        oText.Enabled = False
    Next
    
    'Modify by Amy 2020/06/05 公司別改下拉 原:Text2
    'CboCmp.Enabled = False 'Mark by Amy 2022/02/27 專業單位點數分析也要有公司別選項
    Text3.Enabled = False
    Text4.Enabled = False
    If Combo1 = "專業點數明細表" Then
        'CboCmp.Enabled = True 'Mark by Amy 2022/02/27 專業單位點數分析也要有公司別選項
        Text3.Enabled = True
        Text4.Enabled = True
        If bolFirst = False Then
            CboCmp.SetFocus
        End If
    'end 2020/06/05
    Else
        For Each oText In Txt1
            oText.Enabled = True
        Next
        Txt1(0).SetFocus
    End If
End Sub

Private Sub Command1_Click()
Dim RsQ As New ADODB.Recordset
Dim strQ As String, strSysKind As String
Dim strMsg As String
   
   'Add by Amy 2016/02/23 +控制專業達成點數表(秘書用)當月未過帳不可跑
   'Modify by Amy 2021/09/23 +stState 讓財務也可使用秘書報表 (stState="SEC" 財務 Run 秘書報表)
   'Modify by Amy 2023/01/06 財務也彈訊息
   'If (UCase(App.EXEName) <> "TEACCOUNT" And UCase(App.EXEName) <> "ACCOUNT") Or stState <> MsgText(601) Then
      'modify by sonia 2016/4/8 改判斷收入科目若有未過帳的不可跑
      'If (Val(Left(GetA0b01(strExc(0), "1"), 5)) < Val(Text3 & Text4)) Or (Val(Left(GetA0b01(strExc(0), "J"), 5)) < Val(Text3 & Text4)) Then
      '  MsgBox "當月資料尚未過帳,暫不開放！", , MsgText(5)
      '  Exit Sub
      'End If
      strQ = "select count(*) from acc021,acc020 where a0205>=" & Val(Text3 & Text4 & "01") & " And a0205<=" & Val(Text3 & Text4 & "31") & " And a0201=ax201(+) And a0202=ax202(+) " & _
             "And substr(ax205,1,1) in ('4','7') And nvl(ax210,0)=0"
      RsQ.CursorLocation = adUseClient
      RsQ.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
      If RsQ.RecordCount > 0 Then
         If "" & Val(RsQ.Fields(0)) > 0 Then
            MsgBox "當月收入資料尚未完全過帳,暫不開放！", , MsgText(5)
            Exit Sub
         End If
      End If
      'end 2016/4/8
   'End If
   'Modify by Amy 2019/05/13
   If FormCheck(strMsg) = False Then
      MsgBox strMsg
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   'Add by Amy 2020/06/05 公司別改下拉
   strCmp = "": strCmpN = ""
   If Trim(CboCmp) <> MsgText(601) Then
      strCmp = CboCmp
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
   End If
   strCmpN = GetAccReportCmpN(CboCmp, , True)
   'end 2020/06/05
   If Combo1 <> "專業單位實績點數分析表" Then
        ProduceData
   Else
        ProduceData_Dept
   End If
   'end 2019/05/13
   Screen.MousePointer = vbDefault
   FormClear
   'Mark by Amy 2015/07/28
    'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
'      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
    
   Label12.Visible = False
   Label8.Visible = False 'Add by Amy 2021/09/24 財務提醒文字
   Label9.Visible = False 'Add by Amy 2022/01/28 秘書報表提醒文字
   'Modify by Amy 2015/09/22 (Add Amy 2015/07/28 Account與Patpro共用,但畫面不同)
   'Modify by Amy 2021/09/23 +stState 讓財務也可使用秘書報表(stState="SEC" 財務 Run 秘書報表)
   ChkD5.Visible = False 'Modify by Amy 2024/01/05 勾選顯示5位
   'Modify by Amy 2025/07/14 +讓財務可對表報差異,開放財務也可勾選「顯示結餘+轉撥」
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31" Then ChkChoose.Visible = True 'Add by Amy 2021/11/09
   '秘書報表
   If (UCase(App.EXEName) <> "TEACCOUNT" And UCase(App.EXEName) <> "ACCOUNT") Or stState <> MsgText(601) Then
        '由執行檔Promoter進入(原:Patpro)
        ChkD5.Visible = True 'Add by Amy 2024/01/05
        Me.Caption = "專業達成點數表"
        If stState = MsgText(601) Then cmdOK.Visible = True 'Modify by Amy 2021/09/23 +if
        Label3.Visible = False: CboCmp.Visible = False 'Modify by Amy 2020/06/05公司別改下拉
        Label12.Visible = True: Label12.Left = 2100
        Text3.Top = 240: Label5.Top = 240
        Label2.Top = 624: Label2.Left = 336
        Text4.Top = 624: Text4.Left = 1296
        'Add by Amy 2023/12/07 +label10提醒文字,按鈕下移
        Label10.Visible = True
        'cmdok.Top = 240
        'Add by Amy 2019/05/13 增加專業單位實績點數分析表
        Combo1.Visible = False
        Me.Height = 2100
        'Modify by Amy 2022/01/28 增加公司別文字顯示
        Label9.Visible = True
        Command1.Top = 1280
        'end 2022/01/28
   'Add by Amy 2019/05/13 +下拉選單,可產生專業單位實績點數分析表
   '財務報表
   Else
        'Moidfy by Amy 2021/09/24 +財務提醒文字
        Me.Height = 4110
        Label8.Visible = True
        Label8.Caption = "若「智權點數實績與結餘分析表」" & vbCrLf & _
                                    "與「專業點數明細表」不合" & vbCrLf & _
                                    "且智權點數關閉後又開放修改「結餘」資料，" & vbCrLf & _
                                    "請至「智權期末結餘保留傳票產生」重新產生其傳票！"
        'Add by Amy 2020/6/05 公司別下拉
        CboCmp.Clear
        CboCmp.AddItem "", 0
        Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
        'end 2020/06/05
        Combo1.AddItem "專業點數明細表"
        Combo1.AddItem "專業單位實績點數分析表"
        Combo1 = "專業點數明細表"
        bolFirst = True 'Combo1_Click Text2.setFocus 會錯
        Call Combo1_Click
        bolFirst = False
        Label1.Visible = True
        Label4.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Txt1(0).Visible = True
        Txt1(1).Visible = True
        Txt1(2).Visible = True
   End If
   'end 2015/09/22
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   
   Me.Width = 5250
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Amy 2024/06/26 點數相關都只彈,改只彈這支-婉莘
   '財務進入有操作過 專業點數明細表 且畫面日期等於業績輸入關閉年月(專業點數明細表要當月過帳才可run)
   If (UCase(App.EXEName) = "TEACCOUNT" Or UCase(App.EXEName) = "ACCOUNT") _
     And IsExcelSave = True Then
      MsgBox "請記得確認專業點數及業務點數相同後" & vbCrLf & _
                      "要通知智權主管寫報告 ！"
   End If
   stState = "" 'Add by Amy 2021/09/23
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc44r0 = Nothing
End Sub

'Add by Amy 2020/06/05
Private Sub CboCmp_GotFocus()
    TextInverse CboCmp
End Sub

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

    If InStr(GetBookKeepCmp & 組合作帳公司 & ",", strCmp) = 0 Then
        MsgBox Label3 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/06/05
'Private Sub Text2_Change()
'   If Text2 = MsgText(601) Then
'      Exit Sub
'   End If
'   'Text1 = A0802Query(Text2)
'End Sub
'
'Private Sub Text2_GotFocus()
'   TextInverse Text2
'End Sub
'
'Private Sub Text2_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'   'Add By Sindy 2014/1/22
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'   '2014/1/22 END
'End Sub

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

'*************************************************
'  產生報表資料
'
'*************************************************
'****** Memo 此處有修改需看frmacc42a0 是否也需要改
Private Sub ProduceData()
Dim intCounter As Integer
Dim strUnion As String
'Add By Cheng 2003/06/09
Dim StrSQLa As String
Dim StrSqlB As String
'ADD BY SONIA 2015/9/17
Dim RsQ As New ADODB.Recordset
Dim strQ As String, strSysKind As String
'END 2015/9/17
Dim intQ As Integer 'Add by Amy 2021/05/04
Dim strLastNo As String 'Add by Amy 2021/05/21
Dim stField As String 'Add by Amy 2021/09/23
Dim strA0b05 As String 'Add by Amy 2024/06/26 業績輸入關閉年月

On Error GoTo Checking
    strSql = "": StrSQLa = "": StrSqlB = ""
   'Modify By Sindy 2014/1/22
   'Modify by Amy 2020/06/05 公司別改下拉 原:Text2/IIf(Text2 = "2", "J", "1")
   If Trim(strCmp) <> MsgText(601) Then
      If InStr(strCmp, "+") > 0 Then
        strSql = " And a0403 In ('" & Replace(strCmp, "+", "','") & "')"
        StrSqlB = StrSqlB & " And a0201 In ('" & Replace(strCmp, "+", "','") & "')"
      Else
        strSql = " And a0403 = '" & strCmp & "'"
        StrSqlB = StrSqlB & " And a0201 = '" & strCmp & "'"
      End If
   'Add by Amy 2021/05/21 財務未下公司別時先抓1+J公司,L公司列於最後
   'Modify by Amy 2021/09/23 +stState 讓財務也可使用秘書報表(stState="SEC" 財務 Run 秘書報表)
   'Modify by Amy 2024/05/30 由於秘書報表抓固定會計科目(只有L公司會用),7121科目(智慧及智權公司也會用)L公司 智權人員一向不輸,但11304月卻輸了,導致秘書報表不正確
   '                                                     與辜確認7121科目L公司 智權人員可能會輸,故修改秘書報表排除L公司
   ElseIf ((UCase(App.EXEName) = "TEACCOUNT" Or UCase(App.EXEName) = "ACCOUNT") And stState = MsgText(601)) _
     Or UCase(App.EXEName) = "TEPROMOTER" Or UCase(App.EXEName) = "PROMOTER" Then
  'end 2024/05/30
        strSql = " And a0403 <>'L' "
        StrSqlB = StrSqlB & " And a0201<>'L' "
   End If
   'end 2020/06/05
   '2014/1/22 END
   If Text3 <> MsgText(601) Then
      strSql = strSql & " And a0401 = " & Val(Text3) & ""
   End If
   If Text4 <> MsgText(601) Then
      strSql = strSql & " And a0402 = " & Val(Text4) & ""
   End If
    'Add By Cheng 2003/06/09
    '若有輸入年月
    If Me.Text3.Text <> "" And Me.Text4.Text <> "" Then
        StrSqlB = StrSqlB & " And a0205>=" & Val(Me.Text3.Text & Format(Me.Text4.Text, "00") & "01") & " And a0205<=" & Val(Me.Text3.Text & Format(Me.Text4.Text, "00") & "31") & " "
    '若只輸入年
    ElseIf Me.Text3.Text <> "" And Me.Text4.Text = "" Then
        StrSqlB = StrSqlB & " And a0205>=" & Val(Me.Text3.Text & "0101") & " And a0205<=" & Val(Me.Text3.Text & "1231") & " "
    End If
    'Mark by Amy 2015/07/28
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoquery.CursorLocation = adUseClient
   'Ken 92/01/02 改為需含總所金額
'   strSQLA = "select '1' as No, a0405 as AccNo, a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('410101', '410102', '410103', '410104', '410106', '410107', '410105', '410108') And a0404 = '" & MsgText(55) & "'" & strSQL & " group by a0405, a0102 union " & _
'                 "select '2' as No, '4101T' as AccNo, '合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('410101', '410102', '410103', '410104', '410106', '410107', '410105', '410108') And a0404 = '" & MsgText(55) & "'" & strSQL & " union " & _
'                 "select '3' as No, a0405 as AccNo, a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('417202') And a0404 = '" & MsgText(55) & "'" & strSQL & " group by a0405, a0102 union " & _
'                 "select '4' as No, '4172T' as AccNo, '商標國內專業合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('410101', '410102', '410103', '410104', '410106', '410107', '410105', '410108', '417202') And a0404 = '" & MsgText(55) & "'" & strSQL & " union " & _
'                 "select '5' as No, a0405 as AccNo, a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('417201') And a0404 = '" & MsgText(55) & "'" & strSQL & " group by a0405, a0102 union " & _
'                 "select '6' as No, '4172TT' as AccNo, '商標總達成合計' as Name, 0 as Amount1, 0 as Amount2, Sum(a0408) as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('410101', '410102', '410103', '410104', '410106', '410107', '417201', '417202') And a0404 = '" & MsgText(55) & "'" & strSQL & " union " & _
'                 "select '7' as No, a0405 as AccNo, a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('411101', '411102', '411103', '411104', '411105') And a0404 = '" & MsgText(55) & "'" & strSQL & " group by a0405, a0102 union " & _
'                 "select '8' as No, '4111T' as AccNo, '專利國內專業合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('411101', '411102', '411103', '411104', '411105') And a0404 = '" & MsgText(55) & "'" & strSQL & " union " & _
'                 "select '9' as No, a0405 as AccNo, a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4171') And a0404 = '" & MsgText(55) & "'" & strSQL & " group by a0405, a0102 union " & _
'                 "select 'A' as No, '4111TT' as AccNo, '專利總達成合計' as Name, 0 as Amount1, 0 as Amount2, Sum(a0408) as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('411101', '411102', '411103', '411104', '411105', '4171') And a0404 = '" & MsgText(55) & "'" & strSQL & " union " & _
'                 "select 'B' as No, a0405 as AccNo, a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4121', '4131', '414101', '414102', '415101', '415102', '416101', '416102', '4181', '418101', '418102') And a0404 = '" & MsgText(55) & "'" & strSQL & " group by a0405, a0102 union " & _
'                 "select 'D' as No, 'xxxx' as AccNo, '網址名稱' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('xxxx') And a0404 = '" & MsgText(55) & "'" & strSQL & " group by a0405, a0102 union " & _
'                 "select 'E' as No, 'xxxx' as AccNo, '台灣精品' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('xxxx') And a0404 = '" & MsgText(55) & "'" & strSQL & " group by a0405, a0102 union " & _
'                 "select 'F' as No, '41TTTT' as AccNo, '專業總達成合計' as Name, 0 as Amount1, 0 as Amount2, Sum(a0408) as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('410101', '410102', '410103', '410104', '410106', '410107', '410108', '417202', '417201', '411101', '411102', '411103', '411104', '411105', '4171', '4121', '4131', '414101', '414102', '415101', '415102', '416101', '416102', '4181', '418101', '418102', '410105') And a0404 = '" & MsgText(55) & "'" & strSQL & " union " & _
'                 "select 'G' as No, a0101 as AccNo, '' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from acc010 where a0101 in ('1', '2', '3') union " & _
'                 "select 'G' as No, a0405 as AccNo, '國內上月保留' as Name, a0408 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4191') And a0404 = 'P' And a0401 = " & IIf(Val(Text4) = 1, Val(Text3) - 1, Val(Text3)) & " And a0402 = " & IIf(Val(Text4) = 1, 12, Val(Text4) - 1) & " And a0403 = '" & Text2 & "' union " & _
'                 "select 'H' as No, a0405 as AccNo, '國內本月保留' as Name, a0408 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4191') And a0404 = 'P' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSQL & " union " & _
'                 "select 'I' as No, a0405 as AccNo, 'FCP上月保留' as Name, a0408 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCP' And a0401 = " & IIf(Val(Text4) = 1, Val(Text3) - 1, Val(Text3)) & " And a0402 = " & IIf(Val(Text4) = 1, 12, Val(Text4) - 1) & " And a0403 = '" & Text2 & "' union " & _
'                 "select 'J' as No, a0405 as AccNo, 'FCP本月保留' as Name, a0408 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCP' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSQL & " union " & _
'                 "select 'K' as No, a0405 as AccNo, 'FCT上月保留' as Name, a0408 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCT' And a0401 = " & IIf(Val(Text4) = 1, Val(Text3) - 1, Val(Text3)) & " And a0402 = " & IIf(Val(Text4) = 1, 12, Val(Text4) - 1) & " And a0403 = '" & Text2 & "' union " & _
'                 "select 'L' as No, a0405 as AccNo, 'FCT本月保留' as Name, a0408 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCT' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSQL & " union " & _
'                 "select 'M' as No, a0405 as AccNo, 'FCL上月保留' as Name, a0408 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCL' And a0401 = " & IIf(Val(Text4) = 1, Val(Text3) - 1, Val(Text3)) & " And a0402 = " & IIf(Val(Text4) = 1, 12, Val(Text4) - 1) & " And a0403 = '" & Text2 & "' union " & _
'                 "select 'N' as No, a0405 as AccNo, 'FCL本月保留' as Name, a0408 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCL' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSQL
   'Modify by Morgan 2010/1/6 專利+411106
   'Modify by Morgan 2010/3/2 商標+410109,科目整理,G不用,專業總達成合計因有抓 acc021 資料再Excel計算(本來就有做)
   'Modify By Sindy 2014/1/22 +IIf(Text2 = "2", " And a0403 = 'J'", IIf(Text2 = "1", " And a0403 = '1'", ""))
   '2015/2/4 modify by sonia 2015年新增4194結餘保留科目(只抓部門TOT)改放在下面保留區塊,因語法太長改二句
   '2015/7/14 modify by sonia 4碼大科目的數字欄留空白,故將6碼子科目之科目名稱前加空白以區別
   'StrSQLa = "select '1' as No, a0405 as AccNo, a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4101' And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, a0102 union " & _
                 "select '2' as No, '4101T' as AccNo, '合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4101' And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select '3' as No, a0405 as AccNo, a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0405='417202' And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, a0102 union " & _
                 "select '4' as No, '4172T' as AccNo, '商標國內專業合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And (substr(a0405,1,4)='4101' or a0405='417202') And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select '5' as No, a0405 as AccNo, a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0405='417201' And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, a0102 union " & _
                 "select '6' as No, '4172TT' as AccNo, '商標總達成合計' as Name, 0 as Amount1, 0 as Amount2, Sum(a0408) as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4) in ('4101','4172')  And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select '7' as No, a0405 as AccNo, a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4111' And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, a0102 union " & _
                 "select '8' as No, '4111T' as AccNo, '專利國內專業合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4111' And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select '9' as No, a0405 as AccNo, a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4171' And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, a0102 union " & _
                 "select '91' as No, '4171T' as AccNo, 'FCP合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4171' And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select 'A' as No, '4111TT' as AccNo, '專利總達成合計' as Name, 0 as Amount1, 0 as Amount2, Sum(a0408) as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4) in ('4111','4171')  And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select 'B' as No, a0405 as AccNo, a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, Sum(a0408) as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,2)='41' And not substr(a0405,1,4) in ('4101','4172','4111','4171','4191','4192','4194') And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, a0102 union " & _
                 "select 'D' as No, 'xxxx' as AccNo, '網址名稱' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('xxxx') And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, a0102 union " & _
                 "select 'E' as No, 'xxxx' as AccNo, '台灣精品' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('xxxx') And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, a0102 union " & _
                 "Select 'F' As No, a0101 As AccNo, '其他收入' As Name, Sum(ax207-ax206) As Amount1, 0 As Amount2, Sum(ax207-ax206) As Amount3 From acc010, acc020, acc021 Where a0101=ax205(+) And ax201=a0201(+) And ax202=a0202(+) And a0101='7121' And ax209 Is Not Null " & StrSqlB & " Group By a0101 Union " & _
                 "select 'H' as No, '41TTTT' as AccNo, '專業總達成合計' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from dual "
   'modify by sonia 2016/1/30 +417203加'51'
   'modify by sonia 2016/2/22 +CFT合計'4121T',CFP收入之413102暫無故先剔除4131主科目(And a0405<>'4131')以免重覆
   'Modify by Amy 2017/06/28 加4131CFP收入類另外顯示 改:select 'B' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, Sum(a0408) as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,2)='41' And a0405<>'4131' And not substr(a0405,1,4) in ('4101','4172','4111','4171','4191','4192','4194') And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102
   'Modify by Amy 2019/08/14 增加創新業務組用收入 420101(C 編號那列 原:substr(a0405,1,2)='41' )
   'Modify by Amy 2021/01/05 修改No.3 FCT收入-爭議 417202原抓TOT改抓T/No.5和51列417201及417203改為4172開頭項目中(含417202 FCT部門)
'                 "select '3' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0405='417202' And a0404 = 'T'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union " & _
'                 "select '4' as No, '4172T' as AccNo, '商標國內專業合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And (substr(a0405,1,4)='4101' or a0405='417202') And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
'                 "select '5' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0405='417201' And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union " & _
'                 "select '51' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0405='417203' And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union "
'****** 注意：No有改需看 417202 是否有變動,否則「專業達成點數表-秘書」的程式「商爭 外-內」/「商申 FCT」值會錯
   StrSQLa = "select '1' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4101' And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union " & _
                 "select '2' as No, '4101T' as AccNo, '合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4101' And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select '3' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0405='417202' And a0404 = 'T'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union " & _
                 "select '4' as No, '4101T' as AccNo, '商標國內專業合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And ((substr(a0405,1,4)='4101' And a0404 = '" & MsgText(55) & "' ) or (a0405='417202' And a0404 = 'T')) " & strSql & " union " & _
                 "select '5' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And SubStr(a0405,1,4)='4172' And a0404 = '" & MsgText(55) & "' And a0405<>'417202'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union " & _
                 "select '5' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0405='417202' And a0404='FCT'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union " & _
                 "select '6' as No, '4172T' as AccNo, 'FCT合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And ((SUBSTR(A0405,1,4) = '4172' and A0404 = 'TOT' and A0405<>'417202') Or (a0405='417202' and a0404='FCT'))" & strSql & " union " & _
                 "select '61' as No, '4172TT' as AccNo, '商標總達成合計' as Name, 0 as Amount1, 0 as Amount2, Sum(a0408) as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4) in ('4101','4172')  And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select '7' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4111' And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union " & _
                 "select '8' as No, '4111T' as AccNo, '專利國內專業合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4111' And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select '9' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4171' And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union " & _
                 "select '91' as No, '4171T' as AccNo, 'FCP合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4171' And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select 'A' as No, '4111TT' as AccNo, '專利總達成合計' as Name, 0 as Amount1, 0 as Amount2, Sum(a0408) as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4) in ('4111','4171')  And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select 'B' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, Sum(a0408) as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,2)='41' And SubStr(a0405,1,4)<'4131' And not substr(a0405,1,4) in ('4101','4172','4111','4171','4191','4192','4194') And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union " & _
                 "select 'B' as No, '4121T' as AccNo, 'CFT合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,4)='4121' And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select 'C' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, Sum(a0408) as Amount3 from acc040, acc010 where a0405 = a0101 And SubStr(a0405,1,4)='4131' And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union " & _
                 "select 'C' as No, '4131T' as AccNo, 'CFP合計' as Name, 0 as Amount1, Sum(a0408) as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And SubStr(a0405,1,4)='4131' And a0404 = '" & MsgText(55) & "'" & strSql & " union " & _
                 "select 'C' as No, a0405 as AccNo, decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, Sum(a0408) as Amount3 from acc040, acc010 where a0405 = a0101 And substr(a0405,1,1)='4' And SubStr(a0405,1,4)>'4131' And not substr(a0405,1,4) in ('4101','4172','4111','4171','4191','4192','4194') And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union " & _
                 "select 'D' as No, 'xxxx' as AccNo, '網址名稱' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('xxxx') And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, a0102 union " & _
                 "select 'E' as No, 'xxxx' as AccNo, '台灣精品' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('xxxx') And a0404 = '" & MsgText(55) & "'" & strSql & " group by a0405, a0102 union " & _
                 "Select 'F' As No, a0101 As AccNo, '其他收入' As Name, Sum(ax207-ax206) As Amount1, 0 As Amount2, Sum(ax207-ax206) As Amount3 From acc010, acc020, acc021 Where a0101=ax205(+) And ax201=a0201(+) And ax202=a0202(+) And a0101='7121' And ax209 Is Not Null " & StrSqlB & " Group By a0101 Union " & _
                 "select 'H' as No, '41TTTT' as AccNo, '專業總達成合計' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from dual "
    '2015/7/9 modify by sonia 五組的上月保留及本月保留改只抓本月保留餘額即可
    'StrSQLa = StrSQLa & " union " & _
                 "select 'I' as No, a0405 as AccNo, '國內上月保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4191') And a0404 = 'P' And a0401 = " & IIf(Val(Text4) = 1, Val(Text3) - 1, Val(Text3)) & " And a0402 = " & IIf(Val(Text4) = 1, 12, Val(Text4) - 1) & IIf(Text2 = "2", " And a0403 = 'J'", IIf(Text2 = "1", " And a0403 = '1'", "")) & " union " & _
                 "select 'J' as No, a0405 as AccNo, '國內本月保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4191') And a0404 = 'P' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSql & " union " & _
                 "select 'K' as No, a0405 as AccNo, 'FCP上月保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCP' And a0401 = " & IIf(Val(Text4) = 1, Val(Text3) - 1, Val(Text3)) & " And a0402 = " & IIf(Val(Text4) = 1, 12, Val(Text4) - 1) & IIf(Text2 = "2", " And a0403 = 'J'", IIf(Text2 = "1", " And a0403 = '1'", "")) & " union " & _
                 "select 'L' as No, a0405 as AccNo, 'FCP本月保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCP' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSql & " union " & _
                 "select 'M' as No, a0405 as AccNo, 'FCT上月保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCT' And a0401 = " & IIf(Val(Text4) = 1, Val(Text3) - 1, Val(Text3)) & " And a0402 = " & IIf(Val(Text4) = 1, 12, Val(Text4) - 1) & IIf(Text2 = "2", " And a0403 = 'J'", IIf(Text2 = "1", " And a0403 = '1'", "")) & " union " & _
                 "select 'N' as No, a0405 as AccNo, 'FCT本月保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCT' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSql & " union " & _
                 "select 'O' as No, a0405 as AccNo, 'FCL上月保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCL' And a0401 = " & IIf(Val(Text4) = 1, Val(Text3) - 1, Val(Text3)) & " And a0402 = " & IIf(Val(Text4) = 1, 12, Val(Text4) - 1) & IIf(Text2 = "2", " And a0403 = 'J'", IIf(Text2 = "1", " And a0403 = '1'", "")) & " union " & _
                 "select 'P' as No, a0405 as AccNo, 'FCL本月保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCL' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSql & " union " & _
                 "select 'Q' as No, a0405 as AccNo, '收入－結餘上月保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4194') And a0404 = 'TOT' And a0401 = " & IIf(Val(Text4) = 1, Val(Text3) - 1, Val(Text3)) & " And a0402 = " & IIf(Val(Text4) = 1, 12, Val(Text4) - 1) & IIf(Text2 = "2", " And a0403 = 'J'", IIf(Text2 = "1", " And a0403 = '1'", "")) & " union " & _
                 "select 'R' as No, a0405 as AccNo, '收入－結餘本月保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4194') And a0404 = 'TOT' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSql & " "
    'Modify by Amy 2021/05/24 國內保留 原抓a0404 = 'P' ,因11004月 P2005 有值也要抓,故改抓TOT
    StrSQLa = StrSQLa & " union " & _
                 "select 'J' as No, a0405 as AccNo, '國內保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4191') And a0404 = 'TOT' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSql & " union " & _
                 "select 'L' as No, a0405 as AccNo, 'FCP保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCP' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSql & " union " & _
                 "select 'N' as No, a0405 as AccNo, 'FCT保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCT' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSql & " union " & _
                 "select 'P' as No, a0405 as AccNo, 'FCL保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4192') And a0404 = 'FCL' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSql & " union " & _
                 "select 'R' as No, a0405 as AccNo, '結餘保留' as Name, a0408 as Amount1, 0 as Amount2, a0408 as Amount3 from acc040, acc010 where a0405 = a0101 And a0101 in ('4194') And a0404 = 'TOT' And a0401 = " & Val(Text3) & " And a0402 = " & Val(Text4) & strSql & " "
    'Add By Cheng 2004/03/03
    '2015/2/4 modify by sonia 2015年新增4194結餘保留科目(不限制部門)改放在下面保留區塊
    '2015/7/9 modify by sonia 五組的上月保留及本月保留合併改只抓本月保留餘額即可
    'StrSQLa = StrSQLa & " Union select 'I' as No, '4191' as AccNo, '國內上月保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'J' as No, '4191' as AccNo, '國內本月保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'K' as No, '4192' as AccNo, 'FCP上月保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'L' as No, '4192' as AccNo, 'FCP本月保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'M' as No, '4192' as AccNo, 'FCT上月保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'N' as No, '4192' as AccNo, 'FCT本月保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'O' as No, '4192' as AccNo, 'FCL上月保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'P' as No, '4192' as AccNo, 'FCL本月保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'Q' as No, '4194' as AccNo, '收入－結餘上月保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'R' as No, '4194' as AccNo, '收入－結餘本月保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual "
    StrSQLa = StrSQLa & " Union " & _
                 "select 'J' as No, '4191' as AccNo, '國內保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'L' as No, '4192' as AccNo, 'FCP保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'N' as No, '4192' as AccNo, 'FCT保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'P' as No, '4192' as AccNo, 'FCL保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
                 "select 'R' as No, '4194' as AccNo, '結餘保留' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual "
                 
    'Add by Amy 2021/05/21 財務未下公司別時先抓1+J公司,L公司列於最後
    strLastNo = "R"
    '財務報表
    'Modify by Amy 2021/09/23 +stState 讓財務也可使用秘書報表 (stState="SEC" 財務 Run 秘書報表)
    If (UCase(App.EXEName) = "TEACCOUNT" Or UCase(App.EXEName) = "ACCOUNT") And Trim(strCmp) = MsgText(601) And stState = MsgText(601) Then
        StrSQLa = StrSQLa & " Union " & _
        "Select '" & Chr(Asc(strLastNo) + 1) & "' as No,'" & Chr(Asc(strLastNo) + 1) & "' as AccNo, '法律所' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual union " & _
        "Select '" & Chr(Asc(strLastNo) + 1) & "01' as No,a0405 as AccNo,Decode(length(a0101),6,'　',null)||a0102 as Name, Sum(a0408) as Amount1, 0 as Amount2, Sum(a0408) as Amount3 From acc040, acc010 where a0405 = a0101 And SubStr(a0405,1,1)='4' And a0404 = '" & MsgText(55) & "'" & Replace(strSql, "<>'L'", "='L'") & " group by a0405, decode(length(a0101),6,'　',null)||a0102 union " & _
        "Select '" & Chr(Asc(strLastNo) + 1) & "01' as No,'" & Chr(Asc(strLastNo) + 1) & "ZZ' as AccNo,'法律所合計' as Name, 0 as Amount1, 0 as Amount2, 0 as Amount3 from Dual "
    End If
    '秘書報表
    'Modify by Amy 2021/08/24 +FormN
    'Modify by Amy 2021/09/23 +stState 讓財務也可使用秘書報表 (stState="SEC" 財務 Run 秘書報表)
    If (UCase(App.EXEName) <> "TEACCOUNT" And UCase(App.EXEName) <> "ACCOUNT") Or stState <> MsgText(601) Then
        stField = "'" & strUserNum & "','" & Me.Name & "',"
    End If
    StrSQLa = "Select " & stField & "A.No As No, A.AccNo As AccNo, A.Name As Name, Sum(A.Amount1) As Amount1, Sum(A.Amount2) As Amount2, Sum(A.Amount3) As Amount3 From ( " & StrSQLa & " ) A Group By A.No, A.AccNo, A.Name "
    'end 2021/09/23
    
   '財務報表
   'Modify by Amy 2015/09/22 +if
   'Modify by Amy 2021/09/23 +stState 讓財務也可使用秘書報表 (stState="SEC" 財務 Run 秘書報表)
   If (UCase(App.EXEName) = "TEACCOUNT" Or UCase(App.EXEName) = "ACCOUNT") And stState = MsgText(601) Then
        'Modify by Amy 2017/11/01 +Order by 在O12 run未排序,順序會有問題
        adoquery.Open "Select * From (" & StrSQLa & ") Order by No", adoTaie, adOpenStatic, adLockReadOnly
   '秘書報表則寫入Temp檔 'Memo by Amy 2021/08/24 此有改需看 Frmacc42a0 專業達成點數分佈情況(當月實際達成)-工作表3是否也要改(共用暫存檔)
   Else
        'Modify by Amy 2021/08/24 +FormN
        cnnConnection.Execute "Delete From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' "
        StrSQLa = "Insert Into AccRpt44r0 (ID,FormN,R001,R002,R003,R004,R005,R006) " & StrSQLa
        cnnConnection.Execute StrSQLa
        
        'Modify by Amy 2021/08/24 原程式改至basQuery
        Call ProPoint(Me.Name, Text3, Text4, strSql, StrSqlB)
        
        'Modify by Amy 2021/02/26 11001月後使用新格式,顯示實績點數
        If Val(Text3 & Text4) >= 11001 Then
            'Memo by Amy 2021/08/24 原程式改至 ProPoint中
            
            'Add by Amy 2021/05/03 新增F41XX資料
            'Modify by Amy 2021/08/23 InsF41XX搬至 basQuery,+傳表單名/日期
            Call InsF41XX(Me.Name, Text3 & Text4, Text3 & Text4)
            StrSQLa = GetSQL2
        Else
            StrSQLa = GetSql1
        End If
        adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   End If
   'end 2015/07/28
   If adoquery.RecordCount = 0 Then
      adoquery.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   'Modify by Amy 2015/09/22 +非財務處使用則產生「專業達成點數表」
   'Modify by Amy 2021/09/23 讓財務也可使用秘書報表 (stState="SEC" 財務 Run 秘書報表)
   '秘書報表
   If (UCase(App.EXEName) <> "TEACCOUNT" And UCase(App.EXEName) <> "ACCOUNT") Or stState <> MsgText(601) Then
        'Modify by Amy 2021/02/26 11001 月後使用新格式ExcelSave4
        If Val(Text3 & Text4) >= 11001 Then
            ExcelSave4
        Else
            ExcelSave2
        End If
   '財務報表
   Else
        ExcelSave '專業點數明細表
        'Add by Amy 2024/06/26 點數相關都只彈通知主管寫報告,改只彈智權報當數當月有產生專業點數明細表-婉莘
        Call GetA0b01(strA0b05)
        If Left(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##"))), 6) = Val(strA0b05) + 191100 _
          And Val(Text3 & Text4) + 191100 = Val(strA0b05) + 191100 Then
            IsExcelSave = True
        End If
        'end 2024/06/26
   End If
   'end 2015/09/22
   adoquery.Close
   MsgBox "已產生EXCEL檔案...", , MsgText(5)  '2015/2/4 ADD BY SONIA
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   If adoquery.State <> adStateClosed Then adoquery.Close
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
    'Add by Amy 2019/05/13
    Dim oTxt As TextBox
    For Each oTxt In Txt1
        oTxt = ""
    Next
    'end 2019/05/13
    'Modify by Amy 2020/06/05 公司別改下拉  原:Text2
    CboCmp = ""
    Text3 = ""
    Text4 = ""
    'Modify by Amy 2015/07/28
    'Moidfy by Amy 2019/05/13
    If Combo1 = "專業單位實績點數分析表" Then
        Txt1(0).SetFocus
    Else
        If CboCmp.Visible = True Then
             CboCmp.SetFocus
    'end 2020/06/05
        Else
             Text3.SetFocus
        End If
    End If
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck(ByRef strMsg As String) As Boolean
    'Modify by Amy 2019/05/13 +專業單位實績點數分析表
    Dim oTxt As TextBox
    Dim bCancel As Boolean 'Add by Amy 2020/06/05
    
    FormCheck = False: strMsg = ""
    If Combo1 <> "專業單位實績點數分析表" Then
        'Modify by Amy 2020/06/05 公司別改下拉 原:text2
        If CboCmp <> MsgText(601) Then
           Call CboCmp_Validate(bCancel)
           If bCancel = True Then
                Exit Function
           Else
                FormCheck = True
                Exit Function
           End If
        End If
        If Text3 <> MsgText(601) Then
           FormCheck = True
           Exit Function
        End If
        If Text4 <> MsgText(601) Then
           FormCheck = True
           Exit Function
        End If
        strMsg = "請輸入條件！"
    '專業單位實績點數分析表
    Else
        For Each oTxt In Txt1
            If oTxt = "" Then
                Select Case oTxt.Index
                    Case 0
                        strMsg = "、年"
                    Case 1
                        strMsg = "、起始月"
                    Case 2
                        strMsg = "、截止月"
                End Select
                strMsg = "請輸入統計" & Mid(strMsg, 2)
            End If
        Next
        If strMsg <> MsgText(601) Then Exit Function
            
        If Val(Txt1(1)) > Val(Txt1(2)) Then
            strMsg = "截止月不可大於起始月！"
            Exit Function
        End If
        FormCheck = True
    End If
    'end 2019/05/13
End Function

'*************************************************
'  轉成Excel檔案
'  專業點數明細表
'*************************************************
Private Sub ExcelSave()
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt424 As New Worksheet
Dim lngCounter As Long
Dim lngLocation As Long
'add by nickc 2007/02/08
Dim strDept As String

   strDept = ""
    'Modify By Cheng 2003/06/09
'   If Dir(strExcelPath & Text3 & "年度" & Text4 & "" & ReportTitle(424) & ACDate(ServerDate) & MsgText(43)) = MsgText(601) Then
'      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'         MkDir strExcelPath
'      End If
'   Else
'      Kill strExcelPath & Text3 & "年度" & Text4 & "" & ReportTitle(424) & ACDate(ServerDate) & MsgText(43)
'   End If
   If Dir(strExcelPath & Text3 & "年度" & Text4 & "" & ReportTitle(424) & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strExcelPath & Text3 & "年度" & Text4 & "" & ReportTitle(424) & ACDate(ServerDate) & ServerTime & MsgText(43)
   End If
   'Modify by Amy 2019/08/14 原:1
   xlsSalesPoint.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
   xlsSalesPoint.Workbooks.add
   Set wksaccrpt424 = xlsSalesPoint.Worksheets(1)
   wksaccrpt424.Columns("a:a").ColumnWidth = 30
   wksaccrpt424.Columns("b:b").ColumnWidth = 13
   wksaccrpt424.Columns("c:c").ColumnWidth = 13
   wksaccrpt424.Columns("d:d").ColumnWidth = 13
   wksaccrpt424.Range("a1").Value = Text3 & "年度" & Text4 & "" & ReportTitle(424)
   wksaccrpt424.Range("a1:d1").Select
   'Modified by Morgan 2013/8/3
   'With Selection
   With xlsSalesPoint.Selection
   'end 2013/8/3
       .HorizontalAlignment = xlCenter
       .VerticalAlignment = xlBottom
       .WrapText = False
       .Orientation = 0
       .AddIndent = False
       .ShrinkToFit = False
       .MergeCells = True
   End With
   'Add By Sindy 2014/1/22
   'Modify by Amy 2020/06/05 公司別改抓變數  原:IIf(Text2 = "2", "智權", IIf(Text2 = "1", "台一", "台一　專利商標/智權"))
   wksaccrpt424.Range("b2").Value = "公司別：" & strCmpN
   '2014/1/22 END
   lngCounter = 3
   Do While adoquery.EOF = False
      If adoquery.Fields("Name").Value = "專業總達成合計" Then
         lngLocation = lngCounter
      'Add by Amy 2021/05/21 法律所另外列
      ElseIf adoquery.Fields("Name").Value = "法律所" Then
        '非法律所合計
        wksaccrpt424.Range("a" & lngCounter).Value = "全所合計"
        wksaccrpt424.Range("b" & lngCounter).Value = 0
        wksaccrpt424.Range("c" & lngCounter).Value = 0
        wksaccrpt424.Range("d" & lngCounter).Formula = "=Sum(b" & (lngLocation + 1) & ":b" & (lngCounter - 1) & ")+d" & lngLocation
        lngCounter = lngCounter + 2
        lngLocation = lngCounter
      End If
      wksaccrpt424.Range("a" & lngCounter).Value = adoquery.Fields("Name").Value
      'Modify by Amy 2021/05/21 +if
      If adoquery.Fields("Name").Value <> "法律所合計" Then
          Select Case adoquery.Fields("Name").Value
             'Case "減:國內本月保留", "減:FCP本月保留", "減:FCT本月保留", "減:FCL本月保留"
             '2015/7/9 modify by sonia 五組的上月保留及本月保留合併改只抓本月保留餘額即可,故第一項case不會發生
             Case "國內本月保留", "FCP本月保留", "FCT本月保留", "FCL本月保留", "收入－結餘本月保留"
                wksaccrpt424.Range("b" & lngCounter).Value = Trim(str(adoquery.Fields("Amount1").Value)) - wksaccrpt424.Range("b" & (lngCounter - 1)).Value
             Case Else
                wksaccrpt424.Range("b" & lngCounter).Value = Trim(str(adoquery.Fields("Amount1").Value))
          End Select
          wksaccrpt424.Range("c" & lngCounter).Value = Trim(str(adoquery.Fields("Amount2").Value))
          wksaccrpt424.Range("d" & lngCounter).Value = Trim(str(adoquery.Fields("Amount3").Value))
            'Add By Cheng 2003/07/03
            '搬位置
            Select Case "" & adoquery.Fields("Name").Value
            '2015/7/9 modify by sonia 五組的上月保留及本月保留合併改只抓本月保留餘額即可
            'Case "國內上月保留", "國內本月保留", "FCP上月保留", "FCP本月保留", "FCT上月保留", "FCT本月保留", "FCL上月保留", "FCL本月保留", "收入－結餘上月保留", "收入－結餘本月保留"
            Case "國內保留", "FCP保留", "FCT保留", "FCL保留", "結餘保留"
                wksaccrpt424.Range("d" & lngCounter).Value = wksaccrpt424.Range("b" & lngCounter).Value
            End Select
          Select Case adoquery.Fields("Name").Value
    'Remove by Morgan 2010/3/2 無作用(位置不對,會被覆蓋)
    '         Case "合計"
    '            wksaccrpt424.Range("c11").Formula = "=Sum(b3:b10)"
    '         Case "商標國內專業合計"
    '            wksaccrpt424.Range("c13").Formula = "=Sum(b3:b12)"
    '         Case "商標總達成合計"
    '            wksaccrpt424.Range("d15").Formula = "=Sum(b3:b14)"
    '         Case "專利國內專業合計"
    '            wksaccrpt424.Range("c21").Formula = "=Sum(b16:b20)"
    '         Case "專利總達成合計"
    '            wksaccrpt424.Range("d23").Formula = "=Sum(b16:b22)"
    'end 2010/3/2
             Case "專業總達成合計"
                'Modify By Cheng 2003/08/04
    '            wksaccrpt424.Range("d32").Formula = "=Sum(b3:b31)"
                wksaccrpt424.Range("d" & lngLocation).Formula = "=Sum(b3:b" & lngLocation - 1 & ")"
    '         Case "全所合計"
    '            wksaccrpt424.Range("d45").Formula = "=Sum(b3:b44)"
          End Select
          '2015/7/14 add by sonia
          'Modify by Amy 2017/06/28 4131CFP收入類另外顯示
          'If adoquery.Fields("No").Value < "C" And Len(adoquery.Fields("AccNo").Value) < "C" And adoquery.Fields("amount1").Value + adoquery.Fields("amount2").Value + adoquery.Fields("amount3").Value = 0 Then
          If adoquery.Fields("No").Value < "D" And Len(adoquery.Fields("AccNo").Value) < "D" And adoquery.Fields("amount1").Value + adoquery.Fields("amount2").Value + adoquery.Fields("amount3").Value = 0 Then
             wksaccrpt424.Range("b" & lngCounter).Value = ""
             wksaccrpt424.Range("c" & lngCounter).Value = ""
             wksaccrpt424.Range("d" & lngCounter).Value = ""
          End If
          '2015/7/14 end
      End If
      'end 2021/05/21
      
      lngCounter = lngCounter + 1
      strDept = "" & adoquery.Fields("Name") 'Add by Amy 2021/05/21
      adoquery.MoveNext
   Loop
   'Modify by Amy 2021/05/21 +if 法律所另外列
   If strDept = "法律所合計" Then
        wksaccrpt424.Range("a" & lngCounter).Value = "法律所合計"
   Else
        wksaccrpt424.Range("a" & lngCounter).Value = "全所合計"
   End If
   wksaccrpt424.Range("b" & lngCounter).Value = 0
   wksaccrpt424.Range("c" & lngCounter).Value = 0
   wksaccrpt424.Range("d" & lngCounter).Formula = "=Sum(b" & (lngLocation + 1) & ":b" & (lngCounter - 1) & ")+d" & lngLocation
   
'2015/7/14 cancel by sonia 4碼大科目的數字欄留空白,故將6碼子科目之科目名稱前加空白以區別
   wksaccrpt424.Range("B3:D" & lngCounter).Select
   'Modified by Morgan 2013/8/3
   'Selection.NumberFormatLocal = "#,##0.00_ "
   xlsSalesPoint.Selection.NumberFormatLocal = "#,##0.00_ "
'2015/7/14 end
   'end
   'end 2013/8/3
   'Add by Amy 2020/10/16
   lngCounter = lngCounter + 2
   wksaccrpt424.Range("a" & lngCounter).Value = "*若有非「創新業務收入」之資料列示於創新業務收入，表示有會計科目未歸入正確位置"
   wksaccrpt424.Range("a" & lngCounter).Font.Color = vbBlue
   lngCounter = lngCounter + 1
   wksaccrpt424.Range("a" & lngCounter).Value = "  請通知電腦中心調整下列報表："
   wksaccrpt424.Range("a" & lngCounter).Font.Color = vbBlue
   lngCounter = lngCounter + 1
   wksaccrpt424.Range("a" & lngCounter).Value = "　　查詢：專業達成點數分佈情況(當月實際達成)、國家別點數分析表"
   wksaccrpt424.Range("a" & lngCounter).Font.Color = vbBlue
   lngCounter = lngCounter + 1
   wksaccrpt424.Range("a" & lngCounter).Value = "　　報表：專業點數明細表、專業單位實績點數分析表"
   wksaccrpt424.Range("a" & lngCounter).Font.Color = vbBlue
   lngCounter = lngCounter + 1
   'Mark by Amy 2022/08/16 拿掉不顯示-婧瑄
'   wksaccrpt424.Range("a" & lngCounter).Value = "　　PS.專業單位實績點數分析表(加L公司時未調整此報表，因婧瑄未確認格式該如何調)"
'   wksaccrpt424.Range("a" & lngCounter).Font.Color = vbBlue
   'end 2020/10/16

   'Add by Amy 2016/05/05 判斷若版本2007以上改變存格式
   If Val(xlsSalesPoint.Version) < 12 Then
        'Modify By Cheng 2003/06/09
    '   xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Text3 & "年度" & Text4 & "" & ReportTitle(424) & ACDate(ServerDate) & MsgText(43)
       xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Text3 & "年度" & Text4 & "" & ReportTitle(424) & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Text3 & "年度" & Text4 & "" & ReportTitle(424) & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
   End If
   'end 2016/05/05
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   Set xlsSalesPoint = Nothing
   StatusClear
End Sub

'Add by Amy 2015/07/28 +專業達成點數表(秘書用)
Private Sub ExcelSave2()
    Dim xlsSalesPoint As New Excel.Application
    Dim wksaccrpt424 As New Worksheet
    Dim strOld As String, strTp(3) As String
    Dim IsFirstLaw As Boolean, dblLawFCP As Double '第一筆法務資料/記錄法務FCP目標值 for L4(FCT-法務)計算
    Dim bolFormat As Boolean, bolMerge As Boolean '設定格式/合併儲存格
    Dim intStartRow As Integer, intField As Integer, intCounter As Integer, i As Integer
    Dim strTotal As String '左「總計」欄位
    Dim strRPoint(1 To 18) As String '右「當月收入點數」對應左欄位置 Modify by Amy 2020/06/04 +一般法務
    Dim intRow As Integer 'Add by Amy 2014/06/28
    Dim strLRName As String 'Add by Amy 2017/06/28 for 左表對應右表順序
    
On Error GoTo onErrHand
    
    If Dir(strExcelPath & Text3 & "年度" & Text4 & "月專業達成點數表" & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & Text3 & "年度" & Text4 & "月專業達成點數表" & ACDate(ServerDate) & ServerTime & MsgText(43)
    End If
    
    'Modify by Amy 2019/08/14 +C04,改欄寬
    ReDim strField1(3): ReDim intWidth1(3): ReDim strField2(5): ReDim intWidth2(5)
    strField1 = Array("C01", "C02", "C03", "C04")
    intWidth1 = Array(10, 12, 10, 10)
    'end 2019/08/14
    'Modify by Amy 2018/02/21 入帳點數 欄改為 入帳(報出)點數
    strField2 = Array("單　位", "目　標", "當月收入點數", "達成率", "入帳(報出)點數", "達成率1")
    intWidth2 = Array(11.7, 8.5, 13, 10.13, 13, 11)
    'Add by Amy 2017/06/28 strLRName記錄左表對應右表欄位順序
    'Modify by Amy 2019/08/14 法務字樣前加-
    'Modify by Amy 2020/06/04 +一般法務
    'Modify by Amy 2021/01/06 bug-下 9301月右表對應會不正確,因沒一般法務
    strLRName = ""
    'Modify by Amy 2021/01/29 bug 原:Val(Text3) >= 109 And Val(Text4) >= 4
    If Val(Text3 & Text4) >= 10904 Then strLRName = "一般法務,"
    strLRName = "內專合計,內專-法務,CFP,CFP-法務,內商合計,內商-法務,FCP,FCP-法務,FCT,FCT-法務,CFT,CFT-法務," & strLRName & "著作權/著　爭,ＡＣＳ,條　碼,網　址,其　他"
    'end 2021/01/06
    strLRow = Split(strLRName, ",")
    'end 2017/06/28
    bolXlsOpen = False: intField = 65: intCounter = 1
    'Modify by Amy 2019/08/14 原:1
    xlsSalesPoint.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
    xlsSalesPoint.Workbooks.add
    Set wksaccrpt424 = xlsSalesPoint.Worksheets(1)
    bolXlsOpen = True
    wksaccrpt424.PageSetup.PaperSize = 9 'A4
    wksaccrpt424.PageSetup.Orientation = xlLandscape '橫印
    wksaccrpt424.PageSetup.LeftMargin = xlsSalesPoint.InchesToPoints(0.5)
    wksaccrpt424.PageSetup.RightMargin = xlsSalesPoint.InchesToPoints(0.5)
    wksaccrpt424.PageSetup.TopMargin = xlsSalesPoint.InchesToPoints(0.2) 'Modify by Amy 2018/09/13 原0.3 無法一頁顯示-美珍
    wksaccrpt424.PageSetup.BottomMargin = xlsSalesPoint.InchesToPoints(0.2) 'Modify by Amy 2018/09/13 原0.3
    wksaccrpt424.PageSetup.HeaderMargin = xlsSalesPoint.InchesToPoints(0.3)
    wksaccrpt424.PageSetup.FooterMargin = xlsSalesPoint.InchesToPoints(0.3)
    
    '表頭
    wksaccrpt424.Range(Chr(intField) & intCounter).Value = Text3 & "年度" & Text4 & "月份專業達成點數"
    wksaccrpt424.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strField1) + UBound(strField2) + intField + 2) & intCounter).Select
  
    With xlsSalesPoint.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
    End With
        
    '左-當月收入點數(抓取右表對映位置)
    For i = 1 To UBound(strRPoint)
        strRPoint(i) = "NO"
    Next i
        
    '欄位寬
    For i = 0 To UBound(intWidth1)
        wksaccrpt424.Range(Chr(i + intField) & intCounter).ColumnWidth = intWidth1(i)
    Next i
    
    IsFirstLaw = True: intCounter = intCounter + 1: intStartRow = intCounter
    
    '資料－左
    With adoquery
        Do While .EOF = False
            For i = 0 To UBound(strField1)
                bolFormat = False: bolMerge = False: strTp(0) = "": strTp(1) = ""
                
                'Modify by Amy 2019/08/14 C01與C02欄位對調,原C02->C01,C01->C02
                'Modify by Amy 2020/06/04 +一般法務
                If .Fields("C01") = "法務" And i = GetValue1("C03") Then
                    If IsFirstLaw = True Then Call GetLawValue: IsFirstLaw = False
                    
                    bolFormat = True: strTp(1) = "#,##0.00": strTp(3) = "=" & Chr(i + intField) & intCounter
                    'Modify by Amy 2017/06/28 因插入413102 (CFP-法務)所以調整 strRPoint(右表順序)
                    Select Case Replace(.Fields("C02"), "-法務", "")
                        Case "內專"
                            strTp(0) = dblLawVal(1)
                            'strRPoint(GetRowIdx(.Fields("C02") & .Fields("C01"))) = strTp(3)
                            strRPoint(GetRowIdx(.Fields("C02"))) = strTp(3)
                        Case "CFP"
                            strTp(0) = dblLawVal(7)
                            strRPoint(GetRowIdx(.Fields("C02"))) = strTp(3)
                        Case "內商"
                            strTp(0) = dblLawVal(2)
                            strRPoint(GetRowIdx(.Fields("C02"))) = strTp(3)
                        Case "FCP"
                            strTp(0) = dblLawVal(3)
                            strRPoint(GetRowIdx(.Fields("C02"))) = strTp(3)
                        Case "FCT"
                            strTp(0) = dblLawVal(4)
                            strRPoint(GetRowIdx(.Fields("C02"))) = strTp(3)
                        Case "CFT"
                            strTp(0) = dblLawVal(5)
                            strRPoint(GetRowIdx(.Fields("C02"))) = strTp(3)
                        'Add by Amy 2020/06/04
                        Case "一般法務"
                            strTp(0) = dblLawVal(8)
                            strRPoint(GetRowIdx(.Fields("C02"))) = strTp(3)
                    End Select
                    'end 2017/06/28
                'end 2019/08/14
                '左表非法務數值資料
                Else
                    Select Case i
                        Case GetValue1("C01")
                            If .Fields("C01") <> strOld Then
                                strTp(0) = .Fields("C01")
                            End If
                        Case GetValue1("C02")
                            strTp(1) = "#,##0.00" 'Add by Amy 2019/09/10
                            If IsNull(.Fields("C02")) Then
                                bolMerge = True
                            Else
                                strTp(0) = .Fields("C02")
                            End If
                            'Add by Amy 2019/08/14 著作權/著爭由下搬上來,加創新業務組 ACS
                            'Modify by Amy 2019/09/10 +bolFormat = True
                            '因著爭使用著作權位置,所以用"=("
                            If .Fields("C01") = "著作權" Then strRPoint(GetRowIdx(.Fields("C01"))) = "=(" & Chr(i + intField) & intCounter: bolFormat = True
                            '著爭-使用著作權位置,所以最後加")"
                            If .Fields("C01") = "著　爭" Then strRPoint(GetRowIdx(.Fields("C01"))) = strRPoint(GetRowIdx(.Fields("C01"))) & "+" & Chr(i + intField) & intCounter & ")": bolFormat = True
                            If .Fields("C01") = "ＡＣＳ" Then strRPoint(GetRowIdx(.Fields("C01"))) = "=" & Chr(i + intField) & intCounter: bolFormat = True
                            'end 2019/09/10
                            'end 2019/08/14
                        Case GetValue1("C03")
                            bolFormat = True: strTp(1) = "#,##0.00"
                            If Right(.Fields("C01"), 2) = "合計" Or Right(.Fields("C01"), 2) = "總計" Then
                                strTp(0) = "=Sum(" & Chr(i + intField) & intStartRow & ":" & Chr(i + intField) & intCounter - 1 & ")"
                            Else
                                strTp(0) = "" & .Fields("C03") 'Modify by Amy 2019/08/14 拿掉val
                            End If
                            'Modify by Amy 2017/06/28 因加入CFP-法務所以調整 strRPoint(右表順序),原:順序固定
                            If .Fields("C01") = "內專合計" Then strRPoint(GetRowIdx(.Fields("C01"))) = "=" & Chr(i + intField) & intCounter
                         
                            If .Fields("C01") = "專利" And .Fields("C02") = "CFP" Then strRPoint(GetRowIdx(.Fields("C02"))) = "=" & Chr(i + intField) & intCounter
                            
                            If .Fields("C01") = "內商合計" Then strRPoint(GetRowIdx(.Fields("C01"))) = "=" & Chr(i + intField) & intCounter
                            If .Fields("C01") = "專利" And .Fields("C02") = "FCP" Then strRPoint(GetRowIdx(.Fields("C02"))) = "=" & Chr(i + intField) & intCounter
                            
                            'Modify by Amy 2020/01/06 原:商申
                            If .Fields("C01") = "商標" And .Fields("C02") = "FCT" Then strRPoint(GetRowIdx(.Fields("C02"))) = "=" & Chr(i + intField) & intCounter
                            If .Fields("C01") = "商標" And .Fields("C02") = "CFT" Then strRPoint(GetRowIdx(.Fields("C02"))) = "=" & Chr(i + intField) & intCounter
                            'end 2020/01/06
                            'Memo by Amy 2019/08/14 原:著作權/著爭移至C02(因為值f改顯示於C02),條碼/網址/其他移至C04
                            'end 2017/06/28
                        'Add by Amy 2019/08/14 +C04
                        Case GetValue1("C04")
                            bolFormat = True: strTp(1) = "#,##0.00" 'Add by Amy 2019/09/10 加小數-美珍
                            If IsNull(.Fields("C04")) Then
                                bolMerge = True
                            Else
                                strTp(0) = .Fields("C04")
                            End If
                            If .Fields("C03") = "條　碼" Then strRPoint(GetRowIdx(.Fields("C03"))) = "=" & Chr(i + intField) & intCounter
                            If .Fields("C03") = "網　址" Then strRPoint(GetRowIdx(.Fields("C03"))) = "=" & Chr(i + intField) & intCounter
                            If .Fields("C03") = "其　他" Then strRPoint(GetRowIdx(.Fields("C03"))) = "=" & Chr(i + intField) & intCounter
                    End Select
                End If
                
                If Right(.Fields("C01"), 2) = "合計" Or Right(.Fields("C01"), 2) = "總計" Then
                    wksaccrpt424.Range(Chr(i + intField) & intCounter).Formula = strTp(0)
                Else
                    wksaccrpt424.Range(Chr(i + intField) & intCounter).Value = strTp(0)
                End If
                'Modify by Amy 2019/09/10 +C01及著作權/著　爭/ＡＣＳ
                If i = GetValue1("C01") Or "" & .Fields("C01") = "著作權" Or "" & .Fields("C01") = "著　爭" Or "" & .Fields("C01") = "ＡＣＳ" Then
                    wksaccrpt424.Range(Chr(GetValue1("C01") + intField) & intCounter).HorizontalAlignment = xlCenter
                End If
                If i = GetValue1("C03") And ("" & .Fields("C01") = "著作權" Or "" & .Fields("C01") = "著　爭" Or "" & .Fields("C01") = "ＡＣＳ") Then
                    wksaccrpt424.Range(Chr(GetValue1("C03") + intField) & intCounter).HorizontalAlignment = xlCenter
                End If
                'end 2019/09/10
                If bolMerge = True Then
                    wksaccrpt424.Range(Chr(i + intField - 1) & intCounter & ":" & Chr(i + intField) & intCounter).MergeCells = True
                End If
                If bolFormat = True Then wksaccrpt424.Range(Chr(i + intField) & intCounter).NumberFormatLocal = strTp(1)
            Next i
            
             If Right(.Fields("C01"), 2) = "合計" Or Right(.Fields("C01"), 2) = "總計" Then
                If Right(.Fields("C01"), 2) = "合計" Then
                    wksaccrpt424.Range(Chr(intField) & intStartRow & ":" & Chr(UBound(strField1) + intField) & intCounter).Select
                    intStartRow = intCounter
                Else
                    '法務總計
                    strTotal = strTotal & Chr(GetValue1("C03") + intField) & intCounter & ","
                    wksaccrpt424.Range(Chr(intField) & IIf(Left(.Fields("C01"), 2) = "法務", intStartRow, intStartRow + 1) & ":" & Chr(UBound(strField1) + intField) & intCounter - 1).Select
                    intCounter = intCounter + 1
                    intStartRow = intCounter + 1
                End If
                '畫框
                'Modify by Amy 2019/09/10 Excel2007雙線設定失效 +.Weight = xlThick
                xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlDouble
                xlsSalesPoint.Selection.Borders(xlEdgeLeft).Weight = xlThick
                xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlDouble
                xlsSalesPoint.Selection.Borders(xlEdgeTop).Weight = xlThick
                xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlDouble
                xlsSalesPoint.Selection.Borders(xlEdgeBottom).Weight = xlThick
                xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlDouble
                xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
                xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            End If
            
            strOld = .Fields("C01")
            intCounter = intCounter + 1
            .MoveNext
        Loop
    End With
    '畫框-左最後一個
    wksaccrpt424.Range(Chr(intField) & intStartRow & ":" & Chr(UBound(strField1) + intField) & intCounter - 1).Select
    xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlDouble
    xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlDouble
    xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlDouble
    xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlDouble
    xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    '全所合計
    'Modify by Amy 2019/09/10 空一行-美珍
    wksaccrpt424.Range(Chr(intField) & intCounter + 1).Value = "全 所 總 計"
    wksaccrpt424.Range(Chr(GetValue1("C03") + intField) & intCounter + 1).Formula = "=Sum(" & Left(strTotal, Len(strTotal) - 1) & ")+Sum(" & Chr(intField) & intStartRow & ":" & Chr(UBound(strField1) + intField) & intCounter - 1 & ")"
    'Add by Amy 2019/08/14 +C04
    wksaccrpt424.Range(Chr(GetValue1("C03") + intField) & intCounter + 1 & ":" & Chr(GetValue1("C04") + intField) & intCounter + 1).MergeCells = True
    wksaccrpt424.Range(Chr(GetValue1("C03") + intField) & intCounter + 1 & ":" & Chr(GetValue1("C04") + intField) & intCounter + 1).HorizontalAlignment = xlRight
    'end 2019/08/14
     'end 2019/09/10
    If adoquery.State <> adStateClosed Then adoquery.Close
        
    '資料 -右(xxx年度xx月份專業達成點數)
    If GetRightData = True Then
        intCounter = 2: intField = intField + UBound(strField1) + 2
       
        '欄位名稱
        For i = 0 To UBound(strField2)
            wksaccrpt424.Range(Chr(i + intField) & intCounter).Value = strField2(i)
            wksaccrpt424.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter
            wksaccrpt424.Range(Chr(i + intField) & intCounter).ColumnWidth = intWidth2(i)
        Next i
        intCounter = intCounter + 1: intStartRow = intCounter
        
        With adoquery
            Do While .EOF = False
                For i = 0 To UBound(strField2)
                    bolFormat = True: strTp(0) = "": strTp(1) = ""
                    
                    If .Fields("Dept") = "全　所" Then
                        Select Case i
                            Case GetValue2("單　位")
                                bolFormat = False
                                strTp(0) = .Fields("Dept")
                            Case GetValue2("達成率")
                                strTp(1) = "0.00%"
                                strTp(0) = "=" & Chr(GetValue2("當月收入點數") + intField) & intCounter & "/" & Chr(GetValue2("目　標") + intField) & intCounter
                            Case GetValue2("達成率1")
                                strTp(1) = "0.00%"
                                'Modify by Amy 2018/02/21 入帳點數 欄改為 入帳(報出)點數
                                strTp(0) = "=" & Chr(GetValue2("入帳(報出)點數") + intField) & intCounter & "/" & Chr(GetValue2("目　標") + intField) & intCounter
                            Case Else
                                strTp(0) = "=Sum(" & Chr(i + intField) & intStartRow & ":" & Chr(i + intField) & intCounter - 1 & ")"
                        End Select
                       
                        If bolFormat = True Then
                            wksaccrpt424.Range(Chr(i + intField) & intCounter).Formula = strTp(0)
                            'Modify by Amy 2018/02/21 入帳點數 欄改為 入帳(報出)點數
                            If strTp(1) <> MsgText(601) Or i = GetValue2("當月收入點數") Or i = GetValue2("入帳(報出)點數") Then
                                If i = GetValue2("當月收入點數") Or i = GetValue2("入帳(報出)點數") Then strTp(1) = "#,##0.00"
                                wksaccrpt424.Range(Chr(i + intField) & intCounter).NumberFormatLocal = strTp(1)
                            End If
                        Else
                            wksaccrpt424.Range(Chr(i + intField) & intCounter).Value = strTp(0)
                        End If
                        If .Fields("Dept") = "全　所" And i = GetValue2("單　位") Then
                            wksaccrpt424.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter
                        End If
                    Else
                        Select Case i
                            Case GetValue2("單　位")
                                bolFormat = False
                                strTp(0) = .Fields("Dept")
                            Case GetValue2("目　標")
                                'Modify by Amy 2016/02/05 +Val(Text3) < 105 法務舊資料才需減dblLawFCP
                                If .Fields("Dept") = "FCT－法務" And Val(Text3) < 105 Then
                                    strTp(0) = Round(.Fields("A0409"), 0) - dblLawFCP
                                Else
                                    strTp(0) = Round(.Fields("A0409"), 0)
                                End If
                                strTp(1) = "#,##0" 'Add by Amy 2018/03/07
                                If .Fields("Dept") = "FCP－法務" Then dblLawFCP = Val(strTp(0))
                            Case GetValue2("當月收入點數")
                               strTp(1) = "#,##0.00"
                            Case GetValue2("達成率")
                                strTp(1) = "0.00%"
                                strTp(0) = "=" & Chr(GetValue2("當月收入點數") + intField) & intCounter & "/" & Chr(GetValue2("目　標") + intField) & intCounter
                            'Modify by Amy 2018/02/21 入帳點數 欄改為 入帳(報出)點數
                            Case GetValue2("入帳(報出)點數")
                                strTp(1) = "#,##0.00"
                                strTp(0) = "=" & Chr(GetValue2("當月收入點數") + intField) & intCounter & "+" & Round(.Fields("InPoint"), 2)
                            Case GetValue2("達成率1")
                                strTp(1) = "0.00%"
                                'Modify by Amy 2018/02/21 入帳點數 欄改為 入帳(報出)點數
                                strTp(0) = "=" & Chr(GetValue2("入帳(報出)點數") + intField) & intCounter & "/" & Chr(GetValue2("目　標") + intField) & intCounter
                        End Select
                        
                        If i = GetValue2("當月收入點數") Then
                            'Modify by Amy 2021/10/18 +取小數2位
                            wksaccrpt424.Range(Chr(i + intField) & intCounter).Value = "=Round(" & Replace(strRPoint(intCounter - 2), "=", "") & "/1000,2)"
                        'Modify by Amy 2018/02/21 入帳點數 欄改為 入帳(報出)點數
                        ElseIf i = GetValue2("入帳(報出)點數") Then
                            wksaccrpt424.Range(Chr(i + intField) & intCounter).Formula = strTp(0)
                        ElseIf i = GetValue2("達成率") Or i = GetValue2("達成率1") Then
                            strTp(2) = Chr(GetValue2("目　標") + intField) & intCounter
                            strExc(0) = wksaccrpt424.Range(strTp(2)).Value
                            '有目標才算達成率
                            If Val(strExc(0)) <> 0 Then
                                wksaccrpt424.Range(Chr(i + intField) & intCounter).Formula = PUB_ChkExcelZero(2, strTp(2), strTp(0))
                            End If
                        Else
                            wksaccrpt424.Range(Chr(i + intField) & intCounter).Value = strTp(0)
                        End If
                        'Add by Amy 2019/09/10 單位欄位置中-美珍
                        If i = GetValue2("單　位") Then
                            wksaccrpt424.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter
                        End If
                        If bolFormat = True Then
                            wksaccrpt424.Range(Chr(i + intField) & intCounter).NumberFormatLocal = strTp(1)
                        End If
                    End If
                Next i
                
                intCounter = intCounter + 1
                .MoveNext
            Loop
        End With
        '畫框-右
        wksaccrpt424.Range(Chr(intField) & intStartRow - 1 & ":" & Chr(UBound(strField2) + intField) & intCounter - 1).Select
        xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlDouble
        xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlDouble
        xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlDouble
        xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlDouble
        xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
        xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        '圖表
        'Modified by Lydia 2018/04/16
        'xlsSalesPoint.Charts.add '2010版本Charts.add指令失效
        xlsSalesPoint.ActiveSheet.Shapes.AddChart.Select
        strTp(0) = Chr(GetValue2("單　位") + intField) & intStartRow - 1 & ":" & Chr(GetValue2("當月收入點數") + intField) & intCounter - 2
        xlsSalesPoint.ActiveChart.ChartType = xl3DColumnClustered '圖表樣式
        xlsSalesPoint.ActiveChart.SetSourceData Source:=wksaccrpt424.Range(strTp(0)), PlotBy:=xlColumns
        'xlsSalesPoint.ActiveChart.Location Where:=xlLocationAsObject, Name:="Sheet1"  '圖表放的位置 'Mark by Lydia 2018/04/16 excel2010版有錯誤訊
        
        'Exel使用者自訂圖表
        'Mark by Lydia 2018/04/16 不用套範本
        'If Val(xlsSalesPoint.Version) < 12 Then
             '避免無使用者自訂圖表(範本),使用類似設定
            With xlsSalesPoint.ActiveChart
                .ChartArea.Font.Name = "標楷體"
                .HasTitle = False
                .Elevation = 15               '指定立體圖表的觀察仰角的度數
                .Perspective = 30            '設定立體圖表檢視的遠近景深
                .Rotation = 20                 '指定立體圖表的觀察轉角的度數
                .RightAngleAxes = True '指定圖表的座標軸為直角，並與圖表的轉角或仰角無關，則為 True
                .HeightPercent = 100      '圖表寬度比例（ 5％ 到 500％ 之間）傳回或設定立體圖表的高度
                .AutoScaling = True        '對立體圖表進行調整刻度使其大小接近於等價的平面圖表則為 True
                .Axes(xlCategory).HasTitle = False
                .Axes(xlCategory).TickLabels.Font.Size = 8       'X軸標籤字型大小
                .Axes(xlCategory).TickLabels.Orientation = -45 'X軸標籤方向
                .Legend.Position = xlLegendPositionTop            '圖例顯示位置
                If Val(xlsSalesPoint.Version) < 12 Then 'Added by Lydia 2018/04/16
                    .Axes(xlSeries).HasTitle = False
                End If  'end 2018/04/16
                .Legend.bOrder.LineStyle = xlContinuous 'Added by Lydia 2018/04/16 圖例格式的格線
                .Axes(xlValue).HasTitle = False
                'Add by Amy 2018/04/16 + 縱軸刻度設定
                .Axes(xlValue).TickLabels.Font.Size = 8       'Y軸標籤字型大小
                .Axes(xlValue).MaximumScale = 25000
                .Axes(xlValue).MajorUnit = 5000
                .Axes(xlValue).MinorUnit = 1000
                'end 2018/04/16
                .SeriesCollection(2).ApplyDataLabels Type:=xlDataLabelsShowValue, AutoText:=True, LegendKey:=False '顯示資料標籤
            End With
            
            '設定資料標籤格式
            With xlsSalesPoint.ActiveChart.SeriesCollection(2).DataLabels
                '.Font.Bold = True 'Mark by Amy 2018/04/16 不設粗體-美珍
                .Font.Size = 7
                .Orientation = 25
            End With
            
            '移動圖表
            'Modify by Amy 2018/04/16 2007版位置不正確
'            wksaccrpt424.Shapes(1).IncrementLeft (29.25)
'            wksaccrpt424.Shapes(1).IncrementTop (219)
            strTp(0) = ""
            For i = LBound(intWidth1) To UBound(intWidth1)
                strTp(0) = Val(strTp(0)) + Val(intWidth1(i)) * 6
            Next i
            wksaccrpt424.Shapes(1).Top = Val(wksaccrpt424.Range(Chr(intField) & intCounter).Height) * intCounter - 2
            wksaccrpt424.Shapes(1).Left = Val(strTp(0)) + 60
            
            '取代文字
            strTp(0) = GetValue2("達成率1") + UBound(strField1) + 3
            strTp(1) = "達成率1"
            strTp(2) = "達成率"
            wksaccrpt424.Columns(Val(strTp(0))).Replace what:=strTp(1), Replacement:=strTp(2), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
        'Mark by Lydia 2018/04/16 不用套範本
'            xlsSalesPoint.ActiveChart.ApplyCustomType ChartType:=xlUserDefined, TypeName:="立體群體直條圖-秘書"
'        Else
'            wksaccrpt424.Range(Chr(intField + GetValue2("達成率1")) & "2").Value = "達成率"
'            'Add by Amy 2015/08/11 +OS版本判斷,因圖表範本檔儲存位置不同
'            If Val(PUB_GetVersionNo) >= 6 Then
'                xlsSalesPoint.ApplyChartTemplate ("C:\Users\" & strUserNum & "\AppData\Roaming\Microsoft\Templates\Charts\LineChart-S.crtx")
'            Else
'                xlsSalesPoint.ApplyChartTemplate ("C:\Documents And Settings\" & strUserNum & "\Application Data\Microsoft\Templates\Charts\LineChart-S.crtx")
'            End If
'        End If
        'end 2018/04/16
       '避免Active在圖表上,按「預覽列印」只會出現圖表
        wksaccrpt424.Range(Chr(intField) & intCounter).Select
    End If
    
    '判斷若版本2007以上改變存格式
    If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Text3 & "年度" & Text4 & "月專業達成點數表" & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
    Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Text3 & "年度" & Text4 & "月專業達成點數表" & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
    End If
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    Set xlsSalesPoint = Nothing
    StatusClear
    Exit Sub

onErrHand:
    If Err.Number <> 0 And bolXlsOpen = True Then
      Resume Next
        MsgBox Err.Description, , MsgText(5)
        If Val(xlsSalesPoint.Version) < 12 Then
            xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Text3 & "年度" & Text4 & "月專業達成點數表" & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
        Else
            xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Text3 & "年度" & Text4 & "月專業達成點數表" & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
        End If
        xlsSalesPoint.Workbooks.Close
        xlsSalesPoint.Quit
        Set xlsSalesPoint = Nothing
    End If

End Sub

Private Function GetValue1(pFieldN As String) As Integer
    Dim jj As Integer
    
    For jj = 1 To UBound(strField1)
       If UCase(strField1(jj)) = UCase(pFieldN) Then
          GetValue1 = jj
          Exit For
       End If
    Next jj
End Function

'Modify by Amy 2021/02/26 +IsExcel4
Private Function GetValue2(pFieldN As String, Optional IsExcel4 As Boolean = False) As Integer
    Dim jj As Integer, intStart As Integer
    
    intStart = 1
    'Add by Amy 2023/12/007 bug 條件下10912月前資料會一直彈無「單位」欄
    If Val(Text3 & Text4) < 11001 Then intStart = 0
    If IsExcel4 = True Then intStart = LBound(strField2)
    
    For jj = intStart To UBound(strField2)
       If UCase(strField2(jj)) = UCase(pFieldN) Then
          GetValue2 = jj
          Exit For
       'Add by Amy 2021/11/09 無此欄傳-1
       ElseIf jj = UBound(strField2) And UCase(strField2(jj)) <> UCase(pFieldN) Then
            MsgBox "無「" & pFieldN & "」欄"
       End If
    Next jj
End Function
'end 2021/02/26

Private Function GetRightData() As Boolean
    Dim strQ As String, strP As String, strWhere As String
    Dim intSeq As Integer, intAccPoint(5) As Integer 'Add by Amy 2020/06/04
    
    GetRightData = True
    
    intSeq = 1 'Add by Amy 2020/06/04
    strWhere = "And A0401=" & Val(Text3) & " And A0402=" & Val(Text4)
    
    'Modify by Amy 2020/06/04 RowNo 改抓變數,畫面條件大於等於10904 增加一般法務
    '目標(目標輸在主科目上,所以4121,4131不必考慮412101,413101
    'Modify by Amy 2016/02/05 105年開始法務目標輸在子科目,故改抓子科目的預算
    'Modify by Amy 2017/06/28 原:CFP顯示於內專後改至內專－法務後 加顯示CFP法務 413102 RowNo全改
    strQ = "Select '內　　專' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='4111' " & strWhere
    intAccPoint(0) = intSeq '入帳點數對應RowNo
    intSeq = intSeq + 1
    
    If Val(Text3) >= 105 Then
       'modify by sonia 2016/6/3 因J公司也會有ACC040資料,故目標改用Sum(A0409/1000)
       strQ = strQ & " Union Select '內專－法務' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405 ='411107' " & strWhere
    Else
        strQ = strQ & " Union Select '內專－法務' as Dept,Sum((A0409/1000)*0.6) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405 in ('4141','411107') " & strWhere
    End If
    intSeq = intSeq + 1
    strQ = strQ & " Union Select 'C F P' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='4131' " & strWhere
    intAccPoint(1) = intSeq '入帳點數對應RowNo
    intSeq = intSeq + 1
    strQ = strQ & " Union Select 'CFP－法務' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='413102' " & strWhere
    intSeq = intSeq + 1

    'modify by sonia 2016/1/26 內商－法務加410110
    strQ = strQ & " Union Select '內　　商' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='4101' " & strWhere
    intAccPoint(2) = intSeq '入帳點數對應RowNo
    intSeq = intSeq + 1
    
    If Val(Text3) >= 105 Then
       'modify by sonia 2016/6/3 因J公司也會有ACC040資料,故目標改用Sum(A0409/1000)
        strQ = strQ & " Union Select '內商－法務' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405 ='410110' " & strWhere
    Else
        strQ = strQ & " Union Select '內商－法務' as Dept,Sum((A0409/1000)*0.4) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405 in ('4141','410110') " & strWhere
    End If
    intSeq = intSeq + 1
    
    strQ = strQ & " Union Select 'F C P' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='4171' " & strWhere
    intAccPoint(3) = intSeq '入帳點數對應RowNo
    intSeq = intSeq + 1
    
    If Val(Text3) >= 105 Then
       'modify by sonia 2016/6/3 因J公司也會有ACC040資料,故目標改用Sum(A0409/1000)
         strQ = strQ & " Union Select 'FCP－法務' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='417103' " & strWhere
    Else
         strQ = strQ & " Union Select 'FCP－法務' as Dept,Sum((A0409/1000)*0.7857) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='416101' " & strWhere
    End If
    intSeq = intSeq + 1
    
    strQ = strQ & " Union Select 'F C T' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='4172' " & strWhere
    intAccPoint(4) = intSeq '入帳點數對應RowNo
    intSeq = intSeq + 1
    
    If Val(Text3) >= 105 Then
       'modify by sonia 2016/6/3 因J公司也會有ACC040資料,故目標改用Sum(A0409/1000)
        strQ = strQ & " Union Select 'FCT－法務' as Dept,Sum(A0409/1000)  as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='417203' " & strWhere
    Else
        strQ = strQ & " Union Select 'FCT－法務' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='416101' " & strWhere
    End If
    intSeq = intSeq + 1
    
    strQ = strQ & " Union Select 'C F T' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='4121' " & strWhere
    intAccPoint(5) = intSeq '入帳點數對應RowNo
    intSeq = intSeq + 1
    
    If Val(Text3) >= 105 Then
       'modify by sonia 2016/6/3 因J公司也會有ACC040資料,故目標改用Sum(A0409/1000)
        strQ = strQ & " Union Select 'CFT－法務' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405 ='412102' " & strWhere
    Else
        strQ = strQ & " Union Select 'CFT－法務' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405 in ('416102','412102') " & strWhere
    End If
    intSeq = intSeq + 1
    'end 2016/02/05
    
    'Modify by Amy 2021/01/29 bug 原:Val(Text3) >= 109 And Val(Text4) >= 4
    If Val(Text3 & Text4) >= 10904 Then
        strQ = strQ & " Union Select '一般法務' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And SubStr(A0405,1,4) In ('4141','4161','4181') And A0403<>'L' " & strWhere
        intSeq = intSeq + 1
    End If
    
    'Modify by Amy 2019/08/14 +ＡＣＳ並改RowNo
    'Modify by Amy 2020/05/22 ACS 原抓A0405='420101'
    strQ = strQ & " Union Select '著　作　權' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='4151' " & strWhere
    intSeq = intSeq + 1
    strQ = strQ & " Union Select 'ＡＣＳ' as Dept,Sum(A0409/1000) as A0409," & intSeq & " as RowNo From Acc040 Where A0404='TOT' And A0405='4201' " & strWhere
    intSeq = intSeq + 1
    strQ = strQ & " Union Select '條　碼' as Dept,0 as A0409," & intSeq & " as RowNo From Dual"
    intSeq = intSeq + 1
    strQ = strQ & " Union Select '網　址' as Dept,0 as A0409," & intSeq & " as RowNo From Dual"
    intSeq = intSeq + 1
    strQ = strQ & " Union Select '其　他' as Dept,0 as A0409," & intSeq & " as RowNo From Dual"
    intSeq = intSeq + 1
    strQ = strQ & " Union Select '全　所' as Dept,0 as A0409," & intSeq & " as RowNo From Dual"
    'end 2020/06/04
    
    '入帳(報出)點數
    'Mark by Amy 2019/09/09 不需加報出點數 'Modify by Amy 2019/08/14 +ＡＣＳ
    '" Union Select Sum(A0408/1000) as InPoint,14 as iRowNo From Acc040 Where A0404='TOT' And A0405='420101' " & strWhere
    'Modify by Amy 2020/06/04 iRowNo 改抓變數 intAccPoint()
    strP = "Select Sum(A0408/1000) as InPoint," & intAccPoint(0) & " as iRowNo From Acc040 Where ((A0404='TOT' And A0405='4191') Or (A0404='P' And A0405='4194')) " & strWhere & _
    " Union Select Sum(A0408/1000) as InPoint," & intAccPoint(1) & " as iRowNo From Acc040 Where A0404='CFP' And A0405='4194' " & strWhere & _
    " Union Select Sum(A0408/1000) as InPoint," & intAccPoint(2) & " as iRowNo From Acc040 Where A0404='T' And A0405='4194' " & strWhere & _
    " Union Select Sum(R004/1000) as InPoint," & intAccPoint(3) & " as iRowNo From AccRpt44r0 Where R002='4192' And R003='FCP保留' And ID='" & strUserNum & "' And FormN='" & Me.Name & "'" & _
    " Union Select Sum(R004/1000) as InPoint," & intAccPoint(4) & " as iRowNo From AccRpt44r0 Where R002='4192' And R003='FCT保留' And ID='" & strUserNum & "' And FormN='" & Me.Name & "'" & _
    " Union Select Sum(A0408/1000) as InPoint," & intAccPoint(5) & " as iRowNo From Acc040 Where A0404='CFT' And A0405='4194' " & strWhere
    'end 2017/06/28
    strQ = "Select Dept,Nvl(A0409,0) as A0409,Nvl(InPoint,0) as InPoint,RowNo From (" & strQ & "),(" & strP & ") Where RowNo=iRowNo(+) Order by RowNo"
    adoquery.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoquery.RecordCount = 0 Then GetRightData = False

End Function

'取得法務資料值
'1.P 或 PS 案：全數歸 L1(內專-法務)
'2.T 字頭案件：全數歸 L2(內商-法務)
'3.FCP 或 FG 案：全數歸 L3(FCP-法務)
'4.FCT 或 CFT 或 S 或 CFC 案：全數歸 L4(FCT-法務)
'5.CFL 案 或 科目為416102：全數歸 L5(CFT-法務)
'6.417101科目：全數歸 L3(FCP-法務),同時自417101餘額扣除  'ADD BY SONIA 2015/9/17
'7.417201科目：全數歸 L4(FCT-法務),同時自417201餘額扣除  'ADD BY SONIA 2015/9/17
'8.413102科目： CFP 或 CPS歸 L6(CFP-法務) 'Add by Amy 2017/06/28
'9.其他
'   a)「案件屬性」(lc47) 有專利無商標字樣時,L 案：全數歸 L1(內專-法務) / FCL 或 LIN 案：全數歸 L3(FCP-法務)
'   b)「案件屬性」(lc47) 有商標無專利字樣時,L 案：全數歸 L2(內商-法務) / FCL 或 LIN 案：全數歸 L4(FCT-法務)
'   c)(科目為 4141或4181)且為(L 或 LA 案 或 ax214 is null) 金額之60%歸 L1,L2為[全額-L1]
'      科目為 416101 且為(FCL 或 LIN案 或 ax214 is null) 金額之78.57%歸 L31,L4為[全額-L3]
Private Sub GetLawValue()
Dim RsQ As New ADODB.Recordset
Dim strQ As String, strSysKind As String
Dim dblAmt(1) As Double
Dim StrSQLa As String

    '法務資料值
    For i = 0 To UBound(dblLawVal)
        dblLawVal(i) = 0
    Next i
    
    'Modify by Amy 2021/08/24 +FormN
    strQ = "Select R002 as AccNo,R003 as CaseName,R006 as Amount,R007 as LC47 From AccRpt44r0 " & _
                "Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R001='ZZZ' "
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    With RsQ
        Do While .EOF = False
            strSysKind = ""
            If "" & .Fields("CaseName") <> MsgText(601) Then
                If Left(.Fields("CaseName"), 1) = "T" Then
                    strSysKind = "T"
                Else
                    strSysKind = Left(.Fields("CaseName"), InStr(1, "" & .Fields("CaseName"), "-") - 1)
                End If
            End If
            
            'add by sonia 2016/1/30 +5個其他專業部之法務收入科目410110,411107,417203,417103,412102,(CFP之413102暫無)
            If .Fields("AccNo") = "411107" Then
                dblLawVal(1) = dblLawVal(1) + Val(.Fields("Amount"))
            ElseIf .Fields("AccNo") = "410110" Then
                dblLawVal(2) = dblLawVal(2) + Val(.Fields("Amount"))
            ElseIf .Fields("AccNo") = "417103" Then
                dblLawVal(3) = dblLawVal(3) + Val(.Fields("Amount"))
            ElseIf .Fields("AccNo") = "417203" Then
                dblLawVal(4) = dblLawVal(4) + Val(.Fields("Amount"))
            ElseIf .Fields("AccNo") = "412102" Then
                dblLawVal(5) = dblLawVal(5) + Val(.Fields("Amount"))
            'Add by Amy 2017/06/28 +法務收入科目413102
            ElseIf .Fields("AccNo") = "413102" Then
                dblLawVal(7) = dblLawVal(7) + Val(.Fields("Amount"))
            'Add by Amy 2020/06/04 +一般法務 4141/4161/4181
            'Modify by Amy 2020/07/08 + 畫面條件下 10904後(含)
            ElseIf Val(Text3 & Text4) >= 10904 And (Left(.Fields("AccNo"), 4) = "4141" Or Left(.Fields("AccNo"), 4) = "4161" Or Left(.Fields("AccNo"), 4) = "4181") Then
                dblLawVal(8) = dblLawVal(8) + Val(.Fields("Amount"))
            End If
            'end 2016/1/30
            
            'Modify by Amy 2020/06/04 TT-999999不再計算
            If InStr("" & .Fields("CaseName"), "TT-999999") > 0 Then
            Else
                Select Case strSysKind
                    Case "P", "PS", "CFP", "CPS"
                        'Modify by Amy 2020/07/08 +if
                        If Not (Val(Text3 & Text4) >= 10904 And (Left(.Fields("AccNo"), 4) = "4141" Or Left(.Fields("AccNo"), 4) = "4161" Or Left(.Fields("AccNo"), 4) = "4181")) Then
                            dblLawVal(1) = dblLawVal(1) + Val(.Fields("Amount"))
                        End If
                    Case "T" 'T字頭
                        'Modify by Amy 2020/07/08 +if
                        If Not (Val(Text3 & Text4) >= 10904 And (Left(.Fields("AccNo"), 4) = "4141" Or Left(.Fields("AccNo"), 4) = "4161" Or Left(.Fields("AccNo"), 4) = "4181")) Then
                            dblLawVal(2) = dblLawVal(2) + Val(.Fields("Amount"))
                        End If
                    Case "FCP", "FG"
                        'Modify by Amy 2020/07/08 +if
                        If Not (Val(Text3 & Text4) >= 10904 And (Left(.Fields("AccNo"), 4) = "4141" Or Left(.Fields("AccNo"), 4) = "4161" Or Left(.Fields("AccNo"), 4) = "4181")) Then
                            dblLawVal(3) = dblLawVal(3) + Val(.Fields("Amount"))
                        End If
                    Case "FCT", "CFT", "S", "CFC"
                        'Modify by Amy 2020/07/08 +if
                        If Not (Val(Text3 & Text4) >= 10904 And (Left(.Fields("AccNo"), 4) = "4141" Or Left(.Fields("AccNo"), 4) = "4161" Or Left(.Fields("AccNo"), 4) = "4181")) Then
                            dblLawVal(4) = dblLawVal(4) + Val(.Fields("Amount"))
                        End If
                    Case "CFL"
                        'Modify by Amy 2020/07/08 +if
                        If Not (Val(Text3 & Text4) >= 10904 And (Left(.Fields("AccNo"), 4) = "4141" Or Left(.Fields("AccNo"), 4) = "4161" Or Left(.Fields("AccNo"), 4) = "4181")) Then
                            dblLawVal(5) = dblLawVal(5) + Val(.Fields("Amount"))
                        End If
                        '為後面檢查法務傳票與餘額檔是否相等時,不計算417101   CFL-010753(D104091154)
                        If .Fields("AccNo") = "417101" Or .Fields("AccNo") = "417201" Then
                           dblLawVal(6) = dblLawVal(6) + Val(.Fields("Amount"))
                        End If
                    Case Else
                        'ADD BY SONIA 2015/9/17
                        If .Fields("AccNo") = "417101" Then
                            dblLawVal(3) = dblLawVal(3) + Val(.Fields("Amount"))
                            dblLawVal(6) = dblLawVal(6) + Val(.Fields("Amount"))  '為後面檢查法務傳票與餘額檔是否相等時,不計算417101
                        ElseIf .Fields("AccNo") = "417201" Then
                            dblLawVal(4) = dblLawVal(4) + Val(.Fields("Amount"))
                            dblLawVal(6) = dblLawVal(6) + Val(.Fields("Amount"))  '為後面檢查法務傳票與餘額檔是否相等時,不計算417201
                        'END 2015/9/17
                        'Modify by Amy 2020/07/08 +Val(Text3 & Text4) < 10904
                        ElseIf .Fields("AccNo") = "416102" And Val(Text3 & Text4) < 10904 Then
                            dblLawVal(5) = dblLawVal(5) + Val(.Fields("Amount"))
                        ElseIf InStr("" & .Fields("LC47"), "專利") > 0 And InStr("" & .Fields("LC47"), "商標") = 0 And (strSysKind = "L" Or strSysKind = "FCL") Then
                            'Modify by Amy 2020/07/08 +if
                            If Not (Val(Text3 & Text4) >= 10904 And (Left(.Fields("AccNo"), 4) = "4141" Or Left(.Fields("AccNo"), 4) = "4161" Or Left(.Fields("AccNo"), 4) = "4181")) Then
                                If strSysKind = "L" Then
                                    dblLawVal(1) = dblLawVal(1) + Val(.Fields("Amount"))
                                Else
                                    'FCL
                                     dblLawVal(3) = dblLawVal(3) + Val(.Fields("Amount"))
                                End If
                            End If
                        ElseIf InStr("" & .Fields("LC47"), "專利") = 0 And InStr("" & .Fields("LC47"), "商標") > 0 And (strSysKind = "L" Or strSysKind = "FCL") Then
                            'Modify by Amy 2020/07/08 +if
                            If Not (Val(Text3 & Text4) >= 10904 And (Left(.Fields("AccNo"), 4) = "4141" Or Left(.Fields("AccNo"), 4) = "4161" Or Left(.Fields("AccNo"), 4) = "4181")) Then
                                If strSysKind = "L" Then
                                    dblLawVal(2) = dblLawVal(2) + Val(.Fields("Amount"))
                                Else
                                    'FCL
                                     dblLawVal(4) = dblLawVal(4) + Val(.Fields("Amount"))
                                End If
                            End If
                        'Modify by Amy 2020/07/08 +if
                        ElseIf Not (Val(Text3 & Text4) >= 10904 And (Left(.Fields("AccNo"), 4) = "4141" Or Left(.Fields("AccNo"), 4) = "4161" Or Left(.Fields("AccNo"), 4) = "4181")) Then
                           If (Left(.Fields("AccNo"), 4) = "4141" Or Left(.Fields("AccNo"), 4) = "4181") And (strSysKind = "L" Or strSysKind = "LA" Or strSysKind = "") Then
                                dblAmt(0) = dblAmt(0) + Val(.Fields("Amount"))
                            ElseIf .Fields("AccNo") = "416101" And (strSysKind = "FCL" Or strSysKind = "LIN" Or strSysKind = "") Then
                                dblAmt(1) = dblAmt(1) + Val(.Fields("Amount"))
                            End If
                        End If
                End Select
            End If
            
            .MoveNext
        Loop
    End With
    'Modify by Amy 2020/07/08 +if  10904年起 4141/4161/4181開頭會計科目改至一般法務顯示,若下10903月以前(含)仍需依分配比例顯示
    If Val(Text3 & Text4) < 10904 Then
        dblLawVal(1) = dblLawVal(1) + dblAmt(0) * 0.6
        dblLawVal(2) = dblLawVal(2) + (dblAmt(0) - dblAmt(0) * 0.6)
        dblLawVal(3) = dblLawVal(3) + dblAmt(1) * 0.7857
        dblLawVal(4) = dblLawVal(4) + (dblAmt(1) - dblAmt(1) * 0.7857)
    End If
    'MODIFY BY SONIA 2015/9/17 不含417101,417201科目
    'dblLawVal(0) = dblLawVal(1) + dblLawVal(2) + dblLawVal(3) + dblLawVal(4) + dblLawVal(5)
    'Modify by Amy 2017/06/28 +法務收入科目413102
    'Modify by Amy 2020/06/04 +一般法務+ dblLawVal(8)
    dblLawVal(0) = dblLawVal(1) + dblLawVal(2) + dblLawVal(3) + dblLawVal(4) + dblLawVal(5) - dblLawVal(6) + dblLawVal(7) + dblLawVal(8)
    RsQ.Close
    
    '確認總額是否相等
    'Modify by Amy 2020/6/04 10904 後一般法務 需排除L公司
    'Modify by Amy 2021/01/29 bug 原:Val(Text3) >= 109 And Val(Text4) >= 4
    If Val(Text3 & Text4) >= 10904 Then
          strQ = "Select Sum(Nvl(A0408,0)) as A0408 From Acc040 Where ((SubStr(a0405,1,4) In ('4141','4161','4181') And a0403<>'L') or a0405 in ('410110','411107','417203','417103','412102','413102'))" & _
               "And a0404='TOT' And a0401=" & Val(Text3) & " And a0402=" & Val(Text4)
    Else
        'add by sonia 2016/1/26 +5個其他專業部之法務收入科目410110,411107,417203,417103,412102,(CFP之413102暫無)
        'strQ = "Select Sum(Nvl(A0408,0)) as A0408 From Acc040 Where SubStr(a0405,1,4) In ('4141','4161','4181') And a0404='TOT' " & _
                    "And a0401=" & Val(Text3) & " And a0402=" & Val(Text4)
        strQ = "Select Sum(Nvl(A0408,0)) as A0408 From Acc040 Where (SubStr(a0405,1,4) In ('4141','4161','4181') or a0405 in ('410110','411107','417203','417103','412102','413102'))" & _
               "And a0404='TOT' And a0401=" & Val(Text3) & " And a0402=" & Val(Text4)
    End If
                
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        If Val(dblLawVal(0)) <> Val(RsQ.Fields("A0408")) Then MsgBox "法務資料有誤,請洽電腦中心！", , MsgText(5)
    End If
End Sub
'end 2015/07/28
Private Sub Text4_Validate(Cancel As Boolean)
    If Trim(Text4) = MsgText(601) Then Exit Sub
    
    If Len(Trim(Text4)) = 1 Then Text4 = "0" & Text4
    If Val(Trim(Text4)) = 0 Then
        MsgBox Label2 & "輸入錯誤！"
        Text4_GotFocus
        Cancel = True
        Exit Sub
    End If
End Sub

Private Function GetRowIdx(pFieldN As String) As Integer
    Dim jj As Integer
    
    If pFieldN = "著作權" Or pFieldN = "著　爭" Then
        pFieldN = "著作權/著　爭"
    End If
    
    For jj = LBound(strLRow) To UBound(strLRow)
        If UCase(strLRow(jj)) = UCase(pFieldN) Then
            GetRowIdx = jj + 1
            Exit For
       End If
    Next jj
End Function

'Add by Amy 2019/05/13
Private Sub txt1_GotFocus(Index As Integer)
    TextInverse Txt1(Index)
End Sub

'專業單位實績點數分析表
Private Sub ProduceData_Dept()
    Dim RsQ As New ADODB.Recordset, intQ As Integer
    Dim stVTB As String, stVTB1 As String
    Dim strF(1) As String, strUpd As String, strAccItemList As String, strAccList As String
    Dim strArrive(3) As String, strWhere(3) As String
    
On Error GoTo ErrHnd
   
    strYear = Val(Txt1(0)): strStartMonth = Txt1(1): strEndMonth = Txt1(2)
    If Len(strStartMonth) = 1 Then strStartMonth = "0" & strStartMonth
    If Len(strEndMonth) = 1 Then strEndMonth = "0" & strEndMonth
    
    '傳票檔條件
     strWhere(0) = " And ((a0205>=" & strYear & strStartMonth & "01 And a0205<=" & strYear & strEndMonth & "31) " & _
                                "Or (a0205>=" & strYear - 1 & strStartMonth & "01 And a0205<=" & strYear - 1 & strEndMonth & "31) " & _
                                "Or (a0205>=" & strYear - 2 & strStartMonth & "01 And a0205<=" & strYear - 2 & strEndMonth & "31) ) "
    '公司別
    If Trim(strCmp) <> MsgText(601) Then
        If InStr(strCmp, "+") > 0 Then
            strWhere(1) = strWhere(1) & " And a0403 In ('" & Replace(strCmp, "+", "','") & "')"
            strWhere(0) = strWhere(0) & " And a0201 In ('" & Replace(strCmp, "+", "','") & "')"
        Else
            strWhere(1) = strWhere(1) & " And a0403 = '" & strCmp & "'"
            strWhere(0) = strWhere(0) & " And a0201 = '" & strCmp & "'"
        End If
    End If
    
    '暫存檔欄
    strArrive(0) = ", Sum(Decode(floor(a0205/10000)," & strYear & ",ax206)) Sd4" & _
                          ", Sum(Decode(floor(a0205/10000)," & strYear & ",ax207)) Sc4" & _
                          ", Sum(Decode(floor(a0205/10000)," & strYear - 1 & ",ax206)) Sd5" & _
                          ", Sum(Decode(floor(a0205/10000)," & strYear - 1 & ",ax207)) Sc5" & _
                          ", Sum(Decode(floor(a0205/10000)," & strYear - 2 & ",ax206)) Sd6" & _
                          ", Sum(Decode(floor(a0205/10000)," & strYear - 2 & ",ax207)) Sc6"
                          
    strArrive(3) = ", Sum(Decode(floor(R001/10000)," & strYear & ",R007-R006)) S4" & _
                          ", Sum(Decode(floor(R001/10000)," & strYear - 1 & ",R007-R006)) S5" & _
                          ", Sum(Decode(floor(R001/10000)," & strYear - 2 & ",R007-R006)) S6"


    '餘額檔欄位
    strArrive(1) = "a0405, Sum(Decode(a0401," & strYear & ",a0408)) net4" & _
                                    ", Sum(Decode(a0401," & strYear - 1 & ",a0408)) net5" & _
                                    ", Sum(Decode(a0401," & strYear - 2 & ",a0408)) net6"
                                    
    strWhere(1) = strWhere(1) & _
                          " And (a0401 = " & strYear & " Or a0401 = " & strYear - 1 & " Or a0401 = " & strYear - 2 & ")" & _
                          " And a0402>=" & Val(strStartMonth) & " And a0402 <= " & Val(strEndMonth) & _
                          " And a0404='TOT' And Length(a0405)=6 "
                          
    strArrive(2) = ",net4-Nvl(Decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)),0) C04" & _
                          ",net5-Nvl(Decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)),0) C05" & _
                          ",net6-Nvl(Decode(a0103,'1',(Sd6-Sc6),(Sc6-Sd6)),0) C06"
                          
    '避免有資料未抓到,故傳票資料寫入暫存檔
    strSql = "Delete From Accrpt44r0Dept Where ID='" & strUserNum & "' "
    cnnConnection.Execute strSql
    
    strSql = "Insert Into Accrpt44r0Dept (ID,R001,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012) " & _
                "Select '" & strUserNum & "',a0205,ax201,ax202,ax204,ax205,ax206,ax207" & _
                ",LTrim(substr(lpad(ax214,12,' '),1,3)) as CaseNo1,SubStr(lpad(ax214,12,' '),4,6) as CaseNo2,SubStr(lpad(ax214,12,' '),10,1) as CaseNo3,SubStr(lpad(ax214,12,' '),11,2) " & _
                ",Decode(InStr(ax213||' ','結餘'),0,'N','結餘')||Decode(InStr(ax212,'轉撥'),0,'N','轉撥') " & _
                "From Acc020,Acc021 " & _
                "Where (SubStr(ax205,1,1)='4' Or ax205='7121' ) And Not ( ax205='4191' or ax205='4192' or ax205='4194') And ax209 is not null " & _
                "And ax201(+)=a0201 And ax202(+)=a0202 " & strWhere(0)
    cnnConnection.Execute strSql
   
    strSql = ""
    strF(0) = "Insert Into Accrpt420 (ID,r4201,r4203,r4202,r4207,r4208,R4209,R4211) "
    strWhere(3) = "And ID='" & strUserNum & "' And R4211='2' "
    strF(1) = "a0405,net4-Nvl(s4,0) S4,net5-Nvl(s5,0) S5,net6-Nvl(s6 ,0) S6,'2' "
    
    '資料顯示之暫存檔
    strSql = "Delete From Accrpt420 Where 1=1 " & strWhere(3)
    cnnConnection.Execute strSql
    
    '＊＊＊＊＊＊　專利　＊＊＊＊＊＊
    '*** 內專 4111開頭且非CMP(411103)/FMP(411106) 未列於MCP者都列於CCP(不含法務-411107) ***
    strAccItemList = ",'4111'"
    strWhere(0) = " And SubStr(R005,1,4)='4111' And R005<>'411103' And R005<>'411106' And R005<>'411107' "
    strSql = Replace(strWhere(0), "R005", "ax205")
    'MCP(有代理人且客戶國籍為大陸)
    stVTB = "Select a0101" & Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & _
                " From acc010, ( " & GetCCP(1, strArrive(0), strWhere(0)) & " ) y Where ax205(+)=a0101 " & strSql
    
    'MFCP(有代理人且客戶國籍為非大陸)
    stVTB = stVTB & " Union All" & _
            " Select a0101" & Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & _
            " From acc010, ( " & GetCCP(2, strArrive(0), strWhere(0)) & " ) y Where ax205(+)=a0101" & strSql
    
    'MCP
    strSql = strF(0) & _
                "Select '" & strUserNum & "','12' as RID,'MCP','4111',Sum(C04),Sum(C05),Sum(C06),'2' " & _
                "From (" & stVTB & ") "
    cnnConnection.Execute strSql
    
    'CCP=4111開頭會計科目(不含411103/411106/411107) - 結餘/轉撥 - MCP
    strSql = Replace(strWhere(0), "R005", "a0405")
    strSql = strF(0) & _
               "Select '" & strUserNum & "','1' RID,'內專' AccN,'4111' AccNo,0 as 前2年,0 as  前1年,0 as  當年,'2' From Dual " & _
    "Union Select '" & strUserNum & "','11' RID,'CCP'," & strF(1) & " From " & _
                " (Select " & Replace(strArrive(1), "a0405", "SubStr(a0405,1,4) a0405") & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by SubStr(a0405,1,4) ) " & _
                " ,(Select R005,Sum(S4) S4,Sum(S5) S5,Sum(S6) S6 From (" & _
                     "Select SubStr(R005,1,4) R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & strWhere(0) & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by SubStr(R005,1,4) " & _
          "Union Select R4202,r4207 S4,r4208 S5,R4209 S6 From Accrpt420 Where SubStr(R4202,1,4)='4111' And R4203='MCP' " & strWhere(3) & _
                " ) Group by R005 ) " & _
                "Where a0405=R005(+) "
    cnnConnection.Execute strSql
    
    'CMP
    strSql = " And a0405='411103' "
    strSql = strF(0) & _
                "Select '" & strUserNum & "','13' RID,'CMP'," & strF(1) & " From " & _
                " (Select " & strArrive(1) & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by a0405 ) " & _
                ",(Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by R005 ) " & _
                "Where a0405=R005(+) "
    cnnConnection.Execute strSql
              
    'FMP
    strSql = " And a0405='411106' "
    strSql = strF(0) & _
                "Select '" & strUserNum & "','14' RID,'FMP'," & strF(1) & " From " & _
                " (Select " & strArrive(1) & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by a0405 ) " & _
                ",(Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by R005 ) " & _
                "Where a0405=R005(+) "
    cnnConnection.Execute strSql
    '*** End 內專 4111開頭且非CMP(411103)/FMP(411106) 未列於MCP者都列於CCP(不含法務-411107) ***
    
    '*** CFP ***
    strAccItemList = strAccItemList & ",'4131'"
    strSql = " And SubStr(a0405,1,4)='4131' And a0405<>'413102' "
    strSql = strF(0) & _
                "Select '" & strUserNum & "','21' RID,'CFP'," & strF(1) & " From " & _
                " (Select " & strArrive(1) & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by a0405 ) " & _
                ",(Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by R005 ) " & _
                "Where a0405=R005(+) "
    cnnConnection.Execute strSql
    '*** End CFP ***
    
    '*** FCP ***
    strAccItemList = strAccItemList & ",'4171'"
    strSql = " And SubStr(R005,1,4)='4171' And R005<>'417102' And R005<>'417103' "
    '專利基本檔
    stVTB = " Select Decode(SubStr(Nvl(fa10,cu10),1,3),'101','31','011','32','012','33',Decode(SubStr(Nvl(fa10,cu10),1,1),'2','34','35')) RID" & Replace(Replace(Replace(strArrive(0), "a0205", "R001"), "ax206", "R006"), "ax207", "R007") & ",a0103 " & _
        " From Accrpt44r0Dept, Patent, Fagent, Customer,Acc010 " & _
        " Where ID='" & strUserNum & "' And (InStr(R012,'結餘')=0 And InStr(R012,'轉撥')=0) And R005=a0101(+) " & strSql & _
        " And pa01(+)=R008 And pa02(+)=R009 And pa03(+)=R010 And pa04(+)=R011 And pa01 is not null" & _
        " And fa01(+)=substr(pa75,1,8) And fa02(+)=substr(pa75,9,1) " & _
        " And cu01(+)=substr(pa26,1,8) And cu02(+)=substr(pa26,9,1) " & _
        "Group by Decode(SubStr(Nvl(fa10,cu10),1,3),'101','31','011','32','012','33',Decode(SubStr(Nvl(fa10,cu10),1,1),'2','34','35')),a0103 "
    '服務業務
    stVTB = stVTB & " Union All " & _
        " Select Decode(SubStr(Nvl(fa10,cu10),1,3),'101','31','011','32','012','33',Decode(SubStr(Nvl(fa10,cu10),1,1),'2','34','35')) RID" & Replace(Replace(Replace(strArrive(0), "a0205", "R001"), "ax206", "R006"), "ax207", "R007") & ",a0103 " & _
        " From Accrpt44r0Dept, ServicePractice, Fagent, Customer,Acc010 " & _
        " Where ID='" & strUserNum & "' And (InStr(R012,'結餘')=0 And InStr(R012,'轉撥')=0) And R005=a0101(+) " & strSql & _
        " And sp01(+)=R008 And sp02(+)=R009 And sp03(+)=R010 And sp04(+)=R011 And sp01 is not null" & _
        " And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9,1) " & _
        " And cu01(+)=substr(sp08,1,8) And cu02(+)=substr(sp08,9,1) " & _
        "Group by Decode(SubStr(Nvl(fa10,cu10),1,3),'101','31','011','32','012','33',Decode(SubStr(Nvl(fa10,cu10),1,1),'2','34','35')),a0103 "
    '法務
    stVTB = stVTB & " Union All " & _
        " Select Decode(SubStr(Nvl(fa10,cu10),1,3),'101','31','011','32','012','33',Decode(SubStr(Nvl(fa10,cu10),1,1),'2','34','35')) RID" & Replace(Replace(Replace(strArrive(0), "a0205", "R001"), "ax206", "R006"), "ax207", "R007") & ",a0103 " & _
        " From Accrpt44r0Dept, LawCase, Fagent, Customer,Acc010 " & _
        " Where ID='" & strUserNum & "' And (InStr(R012,'結餘')=0 And InStr(R012,'轉撥')=0) And R005=a0101(+) " & strSql & _
        " And lc01(+)=R008 And lc02(+)=R009 And lc03(+)=R010 And lc04(+)=R011 And lc01 is not null" & _
        " And fa01(+)=substr(lc22,1,8) And fa02(+)=substr(lc22,9,1) " & _
        " And cu01(+)=substr(lc11,1,8) And cu02(+)=substr(lc11,9,1) " & _
        "Group by Decode(SubStr(Nvl(fa10,cu10),1,3),'101','31','011','32','012','33',Decode(SubStr(Nvl(fa10,cu10),1,1),'2','34','35')),a0103 "
    '無本所案號
    stVTB = stVTB & " Union All " & _
        " Select '35' RID" & Replace(Replace(Replace(strArrive(0), "a0205", "R001"), "ax206", "R006"), "ax207", "R007") & ",a0103 " & _
        " From Accrpt44r0Dept,Acc010 " & _
        " Where ID='" & strUserNum & "' And (InStr(R012,'結餘')=0 And InStr(R012,'轉撥')=0) And R008 is null And R005=a0101(+) " & strSql & _
        "Group by a0103 "
        
    strSql = strF(0) & _
                "Select '" & strUserNum & "','3' RID,'FCP' AccN,'4171' AccNo,0 as 前2年,0 as  前1年,0 as  當年,'2' From Dual " & _
     "Union Select '" & strUserNum & "',RID,Decode(RID,'31','美國','32','日本','33','韓國','34','歐洲','35','其他國家') as Nation,ax205" & Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & ",'2' From " & _
                "(Select  RID,'417101' ax205,Sum(Sc4) Sc4,Sum(Sd4) Sd4,Sum(Sc5) Sc5,Sum(Sd5) Sd5,Sum(Sc6) Sc6,Sum(Sd6) Sd6,a0103 From (" & stVTB & ") Group by RID,a0103 )"
    cnnConnection.Execute strSql
       
    'FMP
    strSql = " And a0405='417102' "
    strSql = strF(0) & _
                "Select '" & strUserNum & "','36' RID,'FMP'," & strF(1) & " From " & _
                " (Select " & strArrive(1) & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by a0405 ) " & _
                ",(Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by R005 ) " & _
                "Where a0405=R005(+) "
    cnnConnection.Execute strSql
    '*** End FCP ***
    
    strSql = strF(0) & "Select '" & strUserNum & "','3Z' RID,'上述專利總計' AccN,'' AccNo,0 as 前2年,0 as  前1年,0 as  當年,'2' From Dual "
    cnnConnection.Execute strSql
    '＊＊＊＊＊＊　End 專利　＊＊＊＊＊＊
    
    '＊＊＊＊＊＊　商標　＊＊＊＊＊＊
    '*** 內商 未列於MCT(4101)/CMT(410103)/FMT(410109) 都列於CCP(不含法務-410110) ***
    strAccItemList = strAccItemList & ",'4101'"
    strWhere(0) = " And SubStr(R005,1,4)='4101' And R005<>'410103' And R005<>'410109' And R005<>'410110' "
    strSql = Replace(strWhere(0), "R005", "ax205")
    'MCT(有代理人且客戶國籍為大陸)
    stVTB = "Select a0101" & Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & _
                " From acc010, ( " & GetCCT(1, strArrive(0), strWhere(0)) & " ) y Where ax205(+)=a0101 " & strSql
    
    'MFCT(有代理人且客戶國籍為非大陸)
    stVTB = stVTB & " Union All" & _
            " Select a0101" & Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & _
            " From acc010, ( " & GetCCT(2, strArrive(0), strWhere(0)) & " ) y Where ax205(+)=a0101 " & strSql
    
    'MCT
    strSql = strF(0) & _
                "Select '" & strUserNum & "','42' as RID,'MCT','4101',Sum(C04),Sum(C05),Sum(C06),'2' " & _
                "From (" & stVTB & ") "
    cnnConnection.Execute strSql
    
    'CCT=4101開頭會計科目(不含410103/410109/410110) - 結餘/轉撥 - MCT
    strSql = Replace(strWhere(0), "R005", "a0405")
    strSql = strF(0) & _
                "Select '" & strUserNum & "','4' RID,'內商' AccN,'4101' AccNo,0 as 前2年,0 as  前1年,0 as  當年,'2' From Dual " & _
     "Union Select '" & strUserNum & "','41' RID,'CCT'," & strF(1) & " From " & _
                " (Select " & Replace(strArrive(1), "a0405", "SubStr(a0405,1,4) a0405") & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by SubStr(a0405,1,4) ) " & _
                " ,(Select SubStr(R005,1,4) R005,Sum(S4) S4,Sum(S5) S5,Sum(S6) S6 From (" & _
                     "Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & strWhere(0) & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by R005 " & _
          "Union Select R4202,r4207 S4,r4208 S5,R4209 S6 From Accrpt420 Where SubStr(R4202,1,4)='4101' And R4203='MCT' " & strWhere(3) & _
                " ) Group by R005 ) " & _
                "Where a0405=R005(+) "
    cnnConnection.Execute strSql
    
    'CMT
    strSql = " And a0405='410103' "
    strSql = strF(0) & _
                "Select '" & strUserNum & "','43' RID,'CMT'," & Replace(strF(1), "a0405", "SubStr(a0405,1,4)") & " From Acc010" & _
                ",(Select " & strArrive(1) & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by a0405 ) " & _
                ",(Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by R005 ) " & _
                "Where a0405=R005(+) And a0405=a0101(+) "
    cnnConnection.Execute strSql
    
    'FMT
    strSql = " And a0405='410109' "
    strSql = strF(0) & _
                "Select '" & strUserNum & "','44' RID,'FMT'," & Replace(strF(1), "a0405", "SubStr(a0405,1,4)") & " From Acc010" & _
                ",(Select " & strArrive(1) & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by a0405 ) " & _
                ",(Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by R005 ) " & _
                "Where a0405=R005(+) And a0405=a0101(+) "
    cnnConnection.Execute strSql
    
    'FCT爭議(417202 T)
    strSql = " And a0405='417202' "
    strSql = strF(0) & _
                "Select '" & strUserNum & "','45' RID,'FCT爭議'," & strF(1) & " From " & _
                " (Select " & strArrive(1) & " From Acc040 Where 1=1 " & strSql & Replace(strWhere(1), "a0404='TOT'", "a0404='T'") & " Group by a0405 ) " & _
                " ,(Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And R004='T' And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0 ) Group by R005 ) " & _
                 "Where a0405=R005(+) "
    cnnConnection.Execute strSql
    
    '著作權(4151)
    strAccItemList = strAccItemList & ",'4151'"
    strSql = " And SubStr(a0405,1,4)='4151' "
    strSql = strF(0) & _
                "Select '" & strUserNum & "','46' RID,'著作權'," & Replace(strF(1), "a0405", "SubStr(a0405,1,4)") & " From " & _
                " (Select " & Replace(strArrive(1), "a0405", "SubStr(a0405,1,4) a0405") & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by SubStr(a0405,1,4) ) " & _
                ",(Select SubStr(R005,1,4) R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by SubStr(R005,1,4) ) " & _
                "Where a0405=R005(+) "
    cnnConnection.Execute strSql
    
    '*** End 內商 未列於MCT(4101)/CMT(410103)/FMT(410109) 都列於CCP(不含法務-410110) ***
    
    '*** CFT ***
    strAccItemList = strAccItemList & ",'4121' "
    strSql = " And SubStr(a0405,1,4)='4121' And a0405<>'412102' "
    
    strSql = strF(0) & _
                "Select '" & strUserNum & "','51' RID,'CFT'," & strF(1) & " From " & _
                " (Select " & strArrive(1) & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by a0405 ) " & _
                ",(Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by R005 ) " & _
                "Where a0405=R005(+) "
    cnnConnection.Execute strSql
    '*** End CFT ***
    
     '*** FCT ***
    strAccItemList = strAccItemList & ",'4172' " 'Memo 417202 T列於內商
    strSql = " And ((SubStr(R005,1,4)='4172' And R005<>'417202' And R005<>'417203') Or (R005='417202' And R004='FCT')) "
    '商標基本檔
    stVTB = " Select Decode(SubStr(Nvl(fa10,cu10),1,3),'101','61','011','62','012','63',Decode(SubStr(Nvl(fa10,cu10),1,1),'2','64','65')) RID" & Replace(Replace(Replace(strArrive(0), "a0205", "R001"), "ax206", "R006"), "ax207", "R007") & ",a0103 " & _
        " From Accrpt44r0Dept, TradeMark, Fagent, Customer,Acc010 " & _
        " Where ID='" & strUserNum & "' And (InStr(R012,'結餘')=0 And InStr(R012,'轉撥')=0) And R005=a0101(+) " & strSql & _
        " And tm01(+)=R008 And tm02(+)=R009 And tm03(+)=R010 And tm04(+)=R011 And tm01 is not null" & _
        " And fa01(+)=substr(tm44,1,8) And fa02(+)=substr(tm44,9,1) " & _
        " And cu01(+)=substr(tm23,1,8) And cu02(+)=substr(tm23,9,1) " & _
        "Group by Decode(SubStr(Nvl(fa10,cu10),1,3),'101','61','011','62','012','63',Decode(SubStr(Nvl(fa10,cu10),1,1),'2','64','65')),a0103 "
    '服務業務
    stVTB = stVTB & " Union All " & _
        " Select Decode(SubStr(Nvl(fa10,cu10),1,3),'101','61','011','62','012','63',Decode(SubStr(Nvl(fa10,cu10),1,1),'2','64','65')) RID" & Replace(Replace(Replace(strArrive(0), "a0205", "R001"), "ax206", "R006"), "ax207", "R007") & ",a0103 " & _
        " From Accrpt44r0Dept, ServicePractice, Fagent, Customer,Acc010 " & _
        " Where ID='" & strUserNum & "' And (InStr(R012,'結餘')=0 And InStr(R012,'轉撥')=0) And R005=a0101(+) " & strSql & _
        " And sp01(+)=R008 And sp02(+)=R009 And sp03(+)=R010 And sp04(+)=R011 And sp01 is not null" & _
        " And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9,1) " & _
        " And cu01(+)=substr(sp08,1,8) And cu02(+)=substr(sp08,9,1) " & _
        "Group by Decode(SubStr(Nvl(fa10,cu10),1,3),'101','61','011','62','012','63',Decode(SubStr(Nvl(fa10,cu10),1,1),'2','64','65')),a0103 "
    '法務
    stVTB = stVTB & " Union All " & _
        " Select Decode(SubStr(Nvl(fa10,cu10),1,3),'101','61','011','62','012','63',Decode(SubStr(Nvl(fa10,cu10),1,1),'2','64','65')) RID" & Replace(Replace(Replace(strArrive(0), "a0205", "R001"), "ax206", "R006"), "ax207", "R007") & ",a0103 " & _
        " From Accrpt44r0Dept, LawCase, Fagent, Customer,Acc010 " & _
        " Where ID='" & strUserNum & "' And (InStr(R012,'結餘')=0 And InStr(R012,'轉撥')=0) And R005=a0101(+) " & strSql & _
        " And lc01(+)=R008 And lc02(+)=R009 And lc03(+)=R010 And lc04(+)=R011 And lc01 is not null" & _
        " And fa01(+)=substr(lc22,1,8) And fa02(+)=substr(lc22,9,1) " & _
        " And cu01(+)=substr(lc11,1,8) And cu02(+)=substr(lc11,9,1) " & _
        "Group by Decode(SubStr(Nvl(fa10,cu10),1,3),'101','61','011','62','012','63',Decode(SubStr(Nvl(fa10,cu10),1,1),'2','64','65')),a0103 "
    '無本所案號
    stVTB = stVTB & " Union All " & _
        " Select '65' RID" & Replace(Replace(Replace(strArrive(0), "a0205", "R001"), "ax206", "R006"), "ax207", "R007") & ",a0103 " & _
        " From Accrpt44r0Dept,Acc010 " & _
        " Where ID='" & strUserNum & "' And (InStr(R012,'結餘')=0 And InStr(R012,'轉撥')=0) And R008 is null And R005=a0101(+) " & strSql & _
        "Group by a0103 "
    
    strSql = strF(0) & _
                "Select '" & strUserNum & "','6' RID,'FCT' AccN,'4172' AccNo,0 as 前2年,0 as  前1年,0 as  當年,'2' From Dual " & _
     "Union Select '" & strUserNum & "',RID,Decode(RID,'61','美國','62','日本','63','韓國','64','歐洲','65','其他國家') as Nation,ax205" & Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & ",'2' From " & _
                "(Select  RID,'417201' ax205,Sum(Sc4) Sc4,Sum(Sd4) Sd4,Sum(Sc5) Sc5,Sum(Sd5) Sd5,Sum(Sc6) Sc6,Sum(Sd6) Sd6,a0103 From (" & stVTB & ") Group by RID,a0103 )"
    cnnConnection.Execute strSql
    '*** End FCT ***
    
    strSql = strF(0) & "Select '" & strUserNum & "','6Z' RID,'上述商標總計' AccN,'' AccNo,0 as 前2年,0 as  前1年,0 as  當年,'2' From Dual "
    cnnConnection.Execute strSql
    '＊＊＊＊＊＊　End 商標　＊＊＊＊＊＊
   
    '＊＊＊＊＊＊　法務　＊＊＊＊＊＊
    strSql = " And a0405 In('411107','413102','417103','410110','412102','417203') "
    '各專業單位-法務
    strSql = strF(0) & _
                "Select '" & strUserNum & "','7'||Decode(a0405,'411107','1','413102','2','417103','3','410110','4','412102','5','417203','6','7') RID,Replace(Replace(Replace(a0102,'收入',''),'專利-CCP','內專'),'商標-CCT','內商')," & strF(1) & " From Acc010 " & _
                ",(Select " & strArrive(1) & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by a0405 ) " & _
                ",(Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by R005 ) " & _
                "Where a0405=R005(+) And a0405=a0101(+) "
    cnnConnection.Execute strSql
    
    '一般法務(4141/4161)
    strAccItemList = strAccItemList & ",'4141','4161','4181'"
    strSql = " And SubStr(a0405,1,4) In('4141','4161','4181') "
    strSql = strF(0) & _
                "Select '" & strUserNum & "','77' RID,'一般法務'," & strF(1) & " From " & _
                " (Select " & Replace(strArrive(1), "a0405", "'41X1' a0405") & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " ) " & _
                ",(Select '41X1' R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) ) " & _
                "Where a0405=R005(+) And ( net4<>0 or net5<>0 or net6<>0)  "
    cnnConnection.Execute strSql
    
    '法務有資料才寫大項欄位
    strSql = "Select * From Accrpt420 Where Substr(r4201,1,1)='7' " & strWhere(3)
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strSql)
    If intQ = 1 Then
        strSql = strF(0) & "Select '" & strUserNum & "','7' RID,'法務' AccN,'' AccNo,0 as 前2年,0 as  前1年,0 as  當年,'2' From Dual "
        cnnConnection.Execute strSql
    End If
    
    '＊＊＊＊＊＊　End 法務　＊＊＊＊＊＊
    
    '＊＊＊＊＊＊　其他　＊＊＊＊＊＊
    '創新業務(420101)
    strAccList = strAccList & ",'420101'"
    strSql = " And a0405='420101' "
    strSql = strF(0) & _
                "Select '" & strUserNum & "','81' RID,a0102," & strF(1) & " From Acc010 " & _
                ",(Select " & strArrive(1) & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by a0405 ) " & _
                ",(Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by R005 ) " & _
                "Where a0405=R005(+) And a0405=a0101(+) "
    cnnConnection.Execute strSql
    
    '其他各項收入(490102+7121 對沖業不是空)
    strAccList = strAccList & ",'490102','7121'"
    stVTB1 = "Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' And R005='7121' And (InStr(R012,'結餘')=0 Or InStr(R012,'轉撥')=0) Group by R005 "
    strSql = " And a0405='490102' "
    strSql = strF(0) & _
                "Select '" & strUserNum & "','A1' RID,'其他各項收入','7121',Nvl(Sum(S4),0) S4,Nvl(Sum(S5),0) S5,Nvl(Sum(S6),0) S6,'2' From " & _
                "(Select a0405,net4-Nvl(s4,0) S4,net5-Nvl(s5,0) S5,net6-Nvl(s6 ,0) S6 From " & _
                    " (Select " & strArrive(1) & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by a0405 ) " & _
                    ",(Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by R005 ) " & _
                    "Where a0405=R005(+) Union All " & stVTB1 & _
                ")"
    cnnConnection.Execute strSql
    
    '安全基金撥補(490101)
    strAccList = strAccList & ",'490101'"
    strSql = " And a0405='490101' "
    strSql = strF(0) & _
                "Select '" & strUserNum & "','B1' RID,a0102," & strF(1) & " From Acc010" & _
                ",(Select " & strArrive(1) & " From Acc040 Where 1=1 " & strSql & strWhere(1) & " Group by a0405) " & _
                ",(Select R005" & strArrive(3) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')>0 Or InStr(R012,'轉撥')>0) Group by R005) " & _
                "Where a0405=R005(+) And a0405=a0101(+) "
    cnnConnection.Execute strSql
    
    '判斷是否有未列示之會科
    strSql = " And Not (SubStr(R005,1,4) In(" & Mid(strAccItemList, 2) & ") Or R005 In(" & Mid(strAccList, 2) & "))"
    strSql = strF(0) & _
                "Select '" & strUserNum & "','C1' RID,'其他未列示','4XXX',S4,S5,S6,'2' From " & _
                " (Select " & Mid(strArrive(3), 2) & " From Accrpt44r0Dept Where ID='" & strUserNum & "' " & Replace(strSql, "a0405", "R005") & " And (InStr(R012,'結餘')=0 And InStr(R012,'轉撥')=0) Group by R005 " & _
                ") Where  S4<>0 Or S5<>0 Or S6<>0 "
    cnnConnection.Execute strSql
    
    strSql = strF(0) & "Select '" & strUserNum & "','ZZ' RID,'全所實績總計' AccN,'' AccNo,0 as 前2年,0 as  前1年,0 as  當年,'2' From Dual "
    cnnConnection.Execute strSql
    '＊＊＊＊＊＊　End 其他　＊＊＊＊＊＊
  
    '更新資料值,計算點數及平均,取小數2位(原取整數避免差異過大)
    strSql = "/1000/" & Val(Txt1(2)) - Val(Txt1(1)) + 1
    strSql = "Update Accrpt420 Set r4207=Round(Nvl(r4207,0)" & strSql & ",2),r4208=Round(Nvl(r4208,0)" & strSql & ",2),r4209=Round(Nvl(r4209,0)" & strSql & ",2) " & _
                "Where 1=1 " & strWhere(3)
    cnnConnection.Execute strSql
   
    '讀取暫存檔資料,RID長度為1則為加總欄位
    strF(0) = "R4209 as 前2年,'','',R4208 as 前1年,'','','',R4207 as 當年,'','','' "
    strSql = "Select R4201 RID,R4203 AccN," & strF(0) & " From Accrpt420 Where 1=1 " & strWhere(3) & " Order by RID"
    
    intI = 1
    Set rsNew = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
        rsNew.MoveFirst
        If ExcelSave3 = True Then
            MsgBox "已產生EXCEL檔案...", , MsgText(5)
        End If
        rsNew.Close
    Else
        MsgBox "無符合資料！"
    End If
 
ErrHnd:
    If Err.Number <> 0 Then
        If rsNew.State <> adStateClosed Then rsNew.Close
        MsgBox Err.Description
    End If
    
End Sub

'Add by Amy 2019/05/13 參考frmacc42a0 Process(專業達成點數分佈情況-當月實際達成)修改,2019/08/14 剔除結餘及轉撥)
Private Sub ProduceData_Dept_Old()
'    Dim stVTB As String, stVTB1 As String, stCFT As String, stCFP As String, StrSQLa As String, strBP As String
'    Dim strF(1) As String, strUpd As String
'    Dim strArrive(2) As String, strWhere(2) As String
'
'On Error GoTo ErrHnd
'
'    strYear = Val(txt1(0)): strStartMonth = txt1(1): strEndMonth = txt1(2)
'    If Len(strStartMonth) = 1 Then strStartMonth = "0" & strStartMonth
'    If Len(strEndMonth) = 1 Then strEndMonth = "0" & strEndMonth
'
'    '傳票檔欄位
'    strArrive(0) = ", Sum(Decode(floor(a0205/10000)," & strYear & ",ax206)) Sd4" & _
'                          ", Sum(Decode(floor(a0205/10000)," & strYear & ",ax207)) Sc4" & _
'                          ", Sum(Decode(floor(a0205/10000)," & strYear - 1 & ",ax206)) Sd5" & _
'                          ", Sum(Decode(floor(a0205/10000)," & strYear - 1 & ",ax207)) Sc5" & _
'                          ", Sum(Decode(floor(a0205/10000)," & strYear - 2 & ",ax206)) Sd6" & _
'                          ", Sum(Decode(floor(a0205/10000)," & strYear - 2 & ",ax207)) Sc6"
'     strWhere(0) = " And ((a0205>=" & strYear & strStartMonth & "01 And a0205<=" & strYear & strEndMonth & "31) " & _
'                                "Or (a0205>=" & strYear - 1 & strStartMonth & "01 And a0205<=" & strYear - 1 & strEndMonth & "31) " & _
'                                "Or (a0205>=" & strYear - 2 & strStartMonth & "01 And a0205<=" & strYear - 2 & strEndMonth & "31) ) "
'
'    '餘額檔欄位
'    strArrive(1) = "a0405, Sum(Decode(a0401," & strYear & ",a0408)) net4" & _
'                                    ", Sum(Decode(a0401," & strYear - 1 & ",a0408)) net5" & _
'                                    ", Sum(Decode(a0401," & strYear - 2 & ",a0408)) net6"
'    strWhere(1) = " And (a0401 = " & strYear & " Or a0401 = " & strYear - 1 & " Or a0401 = " & strYear - 2 & ")" & _
'                          " And a0402>=" & Val(strStartMonth) & " And a0402 <= " & Val(strEndMonth) & _
'                          " And a0404='TOT' "
'
'    strArrive(2) = ",net4-Nvl(Decode(a0103,'1',(Sd4-Sc4),(Sc4-Sd4)),0) C04" & _
'                          ",net5-Nvl(Decode(a0103,'1',(Sd5-Sc5),(Sc5-Sd5)),0) C05" & _
'                          ",net6-Nvl(Decode(a0103,'1',(Sd6-Sc6),(Sc6-Sd6)),0) C06"
'
'    '*** 會科:410101/410104 ***
'    '傳票CCT=CCT餘額(包含MCT及MFCT)-MCT-MFCT(有代理人) 沒本所案號 or ax214 is null的歸到最後一句
'    stVTB = GetCCT(0, strArrive(0), strWhere(0))
'
'    '餘額(CCT/CCT爭議)
'    stVTB1 = "Select " & strArrive(1) & " From acc040 Where a0405 IN ('410101','410104')" & _
'            strWhere(1) & "Group by a0405"
'
'    '顯示CCT/CCT爭議=餘額-MCT-MFCT(有代理人)
'    strSql = "Select '" & strUserNum & "',DECODE(a0101,'410101','110','410104','140') RID, a0101, a0102 C00" & strArrive(2) & ",'2' " & _
'        " From acc010, (" & stVTB1 & ") w, (" & stVTB & ") y Where a0101 IN ('410101','410104') And a0405(+)=a0101 And ax205(+)=a0101"
'
'    'MCT(有代理人且客戶國籍為大陸)
'    strSql = strSql & " Union All" & _
'        " Select '" & strUserNum & "',Decode(a0101,'410101','111','410104','141') RID, a0101, a0102||'-'||'MCT' C00" & _
'            Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & ",'2' " & _
'        " From acc010, (" & GetCCT(1, strArrive(0), strWhere(0)) & " ) y Where a0101 IN ('410101','410104') And ax205(+)=a0101"
'
'    'MFCT(有代理人且客戶國籍為非大陸)
'    strSql = strSql & " Union All" & _
'        " Select '" & strUserNum & "',Decode(a0101,'410101','112','410104','142') RID, a0101, a0102||'-'||'MFCT' C00" & _
'            Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & ",'2' " & _
'        " From acc010, (" & GetCCT(2, strArrive(0), strWhere(0)) & " ) y Where a0101 IN ('410101','410104') And ax205(+)=a0101"
'    '*** End 會科:410101/410104  ***
'
'    '*** 商標收入 ***
'    stVTB1 = "Select " & strArrive(1) & _
'        " From acc040 Where a0405 IN ('410102','410103','410109','410105','410106','410107','410108','417202','410110')" & _
'        strWhere(1) & " Group by a0405"
'
'    strSql = strSql & " Union All" & _
'        " Select '" & strUserNum & "',Decode(a0101,'410102','12','410103','13','410109','131','410105','15','410106','16','410107','17','410108','18','417202','19','410110','181')" & _
'        " , a0101, a0102,net4,net5,net6,'2' " & _
'        " From acc010, (" & stVTB1 & ") w Where a0101 in ('410102','410103','410109','410105','410106','410107','410108','417202','410110')" & _
'        " And a0405(+)=a0101"
'    '*** End 商標收入 ***
'
'   '*** FCT 收入-以國家區分***
'   '沒有本所號的歸到最後一句
'    stVTB = "Select ax205,RID" & strArrive(0) & _
'       " From (Select a0205, ax205, Decode(substr(nvl(fa10,cu10),1,3),'101','21','011','22','012','23',Decode(substr(nvl(fa10,cu10),1,1),'2','24','25')) RID" & _
'                    " , ax206, ax207 From acc020, acc021, trademark, fagent, customer" & _
'       " Where ax201(+)=a0201 And ax202(+)=a0202 And ax205='417201' " & strWhere(0) & _
'       " And tm01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) And tm02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'       " And tm03(+)=substr(lpad(ax214,12,' '),10,1) And tm04(+)=substr(lpad(ax214,12,' '),11,2) And tm01 is not null" & _
'       " And fa01(+)=substr(tm44,1,8) And fa02(+)=substr(tm44,9,1)" & _
'       " And cu01(+)=substr(tm23,1,8) And cu02(+)=substr(tm23,9,1)"
'
'    '服務業務
'    stVTB = stVTB & " Union All" & _
'       " Select a0205, ax205, Decode(substr(nvl(fa10,cu10),1,3),'101','21','011','22','012','23',Decode(substr(nvl(fa10,cu10),1,1),'2','24','25')) RID" & _
'                " , ax206, ax207 From acc020, acc021, servicepractice, fagent, customer" & _
'       " Where ax201(+)=a0201 And ax202(+)=a0202 And ax205='417201' " & strWhere(0) & _
'       " And sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) And sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'       " And sp03(+)=substr(lpad(ax214,12,' '),10,1) And sp04(+)=substr(lpad(ax214,12,' '),11,2) And sp01 is not null" & _
'       " And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9,1)" & _
'       " And cu01(+)=substr(sp08,1,8) And cu02(+)=substr(sp08,9,1)"
'
'   '法務(ex:D099090625 FCL010445000)
'    stVTB = stVTB & " Union All" & _
'        " Select a0205, ax205, Decode(substr(nvl(fa10,cu10),1,3),'101','21','011','22','012','23',Decode(substr(nvl(fa10,cu10),1,1),'2','24','25')) RID" & _
'                    " , ax206, ax207 From acc020, acc021, lawcase, fagent, customer" & _
'        " Where ax201(+)=a0201 And ax202(+)=a0202 And ax205='417201' " & strWhere(0) & _
'        " And lc01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) And lc02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'        " And lc03(+)=substr(lpad(ax214,12,' '),10,1) And lc04(+)=substr(lpad(ax214,12,' '),11,2) And (lc01 is not null or ax214 is null)" & _
'        " And fa01(+)=substr(lc22,1,8) And fa02(+)=substr(lc22,9,1)" & _
'        " And cu01(+)=substr(lc11,1,8) And cu02(+)=substr(lc11,9,1)" & _
'        " Union All Select 0,'417201','21',0,0 From Dual" & _
'        " Union All Select 0,'417201','22',0,0 From Dual" & _
'        " Union All Select 0,'417201','23',0,0 From Dual" & _
'        " Union All Select 0,'417201','24',0,0 From Dual" & _
'        " Union All Select 0,'417201','25',0,0 From Dual" & _
'        ") x GROUP BY ax205, RID"
'
'    strSql = strSql & " Union All" & _
'        " Select '" & strUserNum & "',RID, a0101, a0102||'-'||Decode(RID,'21','美國','22','日本','23','韓國','24','歐洲','其他') C00" & _
'            Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & ",'2' " & _
'        " From acc010, (" & stVTB & ") y Where a0101='417201' And ax205(+)=a0101"
'
'   '法務 417203
'    stVTB1 = "Select " & strArrive(1) & " From acc040 Where a0405 in ('417203')" & _
'        strWhere(1) & " Group by a0405"
'
'    strSql = strSql & " Union All" & _
'        " Select '" & strUserNum & "',Decode(a0101,'417203','26') RID" & _
'        " , a0101, a0102,net4,net5,net6,'2' " & _
'        " From acc010, (" & stVTB1 & ") w Where a0101 in ('417203')" & _
'        " And a0405(+)=a0101"
'    '*** End FCT 收入 ***
'
'    'CFT4121拆為 412101CFT收入及412102CFT收入-法務
'    stCFT = "Select " & strArrive(1) & _
'        " From acc040 Where a0405 in ('412101','412102') " & strWhere(1) & " Group by a0405"
'
'    strSql = strSql & " Union All" & _
'        " Select '" & strUserNum & "',Decode(a0101,'412101','31','412102','32') RID" & _
'        ", a0101, a0102,net4,net5,net6,'2' " & _
'        " From acc010, (" & stCFT & ") w Where a0101 in ('412101','412102') And a0405(+)=a0101 "
'
'
'    '*** 會科:411101***
'    '傳票CCP=CCP餘額(包含MCP及MFCP)-MCP-MFCP(有代理人) 沒本所案號 or ax214 is null的歸到最後一句
'    stVTB = GetCCP(0, strArrive(0), strWhere(0))
'
'   '餘額CCP
'    stVTB1 = "Select " & strArrive(1) & _
'        " From acc040 Where a0405='411101' " & strWhere(1) & " Group by a0405"
'
'    '顯示CCP=餘額-MCP-MFCP(有代理人)
'    strSql = strSql & " Union All" & _
'        " Select '" & strUserNum & "','410' RID, a0101, a0102 NAME" & strArrive(2) & ",'2' " & _
'        " From acc010, (" & stVTB1 & ") w, (" & stVTB & ") y Where a0101='411101'" & _
'        " And a0405(+)=a0101 And ax205(+)=a0405"
'
'    'MCP(有代理人且客戶國籍為大陸)
'    strSql = strSql & " Union All" & _
'        " Select '" & strUserNum & "','411' RID, a0101, a0102||'-'||'MCP' C00" & _
'            Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & ",'2' " & _
'        " From acc010, ( " & GetCCP(1, strArrive(0), strWhere(0)) & " ) y Where a0101='411101' And ax205(+)=a0101"
'
'    'MFCP(有代理人且客戶國籍為非大陸)
'    strSql = strSql & " Union All" & _
'        " Select '" & strUserNum & "','412' RID, a0101, a0102||'-'||'MFCP' C00" & _
'            Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & ",'2' " & _
'        " From acc010, ( " & GetCCP(2, strArrive(0), strWhere(0)) & " ) y Where a0101='411101' And ax205(+)=a0101"
'
'    '*** 專利收入 ***
'    stVTB1 = "Select " & strArrive(1) & _
'        " From acc040 Where a0405 IN ('411102','411103','411104','411105','411106','411107') " & strWhere(1) & " Group by a0405"
'
'    strSql = strSql & " Union All " & _
'        " Select '" & strUserNum & "',Decode(a0101,'411102','42','411103','43','411104','44','411105','45','411106','46','411107','461')" & _
'        " , a0101, a0102,net4,net5,net6,'2' " & _
'        " From acc010, (" & stVTB1 & ") Where a0101 in ('411102','411103','411104','411105','411106','411107')" & _
'        " And a0405(+)=a0101"
'    '*** End 專利收入 ***
'
'    '*** FCP 收入 ***
'    stVTB = " Select ax205,RID" & strArrive(0) & _
'        " From ( Select a0205, Decode(AX205,'417104','417101','417105','417101','417109','417101',AX205) ax205, Decode(substr(nvl(fa10,cu10),1,3),'101','51','011','52','012','53',Decode(substr(nvl(fa10,cu10),1,1),'2','54','55')) RID" & _
'        " , ax206, ax207 From acc020, acc021, patent, fagent, customer" & _
'        " Where ax201(+)=a0201 And ax202(+)=a0202 And Decode(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101' And ax209 is not null" & _
'        " And pa01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) " & strWhere(0) & _
'        " And pa02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'        " And pa03(+)=substr(lpad(ax214,12,' '),10,1)" & _
'        " And pa04(+)=substr(lpad(ax214,12,' '),11,2) And pa01 is not null" & _
'        " And fa01(+)=substr(pa75,1,8) And fa02(+)=substr(pa75,9,1)" & _
'        " And cu01(+)=substr(pa26,1,8) And cu02(+)=substr(pa26,9,1)"
'
'    '服務業務
'    stVTB = stVTB & " Union All " & _
'        " Select a0205, Decode(AX205,'417104','417101','417105','417101','417109','417101',AX205) ax205, Decode(substr(nvl(fa10,cu10),1,3),'101','51','011','52','012','53',Decode(substr(nvl(fa10,cu10),1,1),'2','54','55')) RID" & _
'        " , ax206, ax207 From acc020, acc021, servicepractice, fagent, customer" & _
'        "　Where ax201(+)=a0201 And ax202(+)=a0202 And Decode(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101' And ax209 is not null" & _
'        " And sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) " & strWhere(0) & _
'        " And sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'        " And sp03(+)=substr(lpad(ax214,12,' '),10,1)" & _
'        " And sp04(+)=substr(lpad(ax214,12,' '),11,2) And sp01 is not null" & _
'        " And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9,1)" & _
'        " And cu01(+)=substr(sp08,1,8) And cu02(+)=substr(sp08,9,1)"
'
'    '法務
'        stVTB = stVTB & " Union All " & _
'          " Select a0205, Decode(AX205,'417104','417101','417105','417101','417109','417101',AX205) ax205, Decode(substr(nvl(fa10,cu10),1,3),'101','51','011','52','012','53',Decode(substr(nvl(fa10,cu10),1,1),'2','54','55')) RID" & _
'          " , ax206, ax207 From acc020, acc021, lawcase, fagent, customer" & _
'          "　Where ax201(+)=a0201 And ax202(+)=a0202 And Decode(AX205,'417104','417101','417105','417101','417109','417101',AX205)='417101' And ax209 is not null" & _
'          " And lc01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) " & strWhere(0) & _
'          " And lc02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'          " And lc03(+)=substr(lpad(ax214,12,' '),10,1)" & _
'          " And lc04(+)=substr(lpad(ax214,12,' '),11,2) And (lc01 is not null or ax214 is null)" & _
'          " And fa01(+)=substr(lc22,1,8) And fa02(+)=substr(lc22,9,1)" & _
'          " And cu01(+)=substr(lc11,1,8) And cu02(+)=substr(lc11,9,1)"
'
'    stVTB = stVTB & _
'       " Union All Select 0,'417101','51',0,0 From dual" & _
'       " Union All Select 0,'417101','52',0,0 From dual" & _
'       " Union All Select 0,'417101','53',0,0 From dual" & _
'       " Union All Select 0,'417101','54',0,0 From dual" & _
'       " Union All Select 0,'417101','55',0,0 From dual" & _
'       " ) x GROUP BY ax205, RID"
'
'    strSql = strSql & " Union All" & _
'       " Select '" & strUserNum & "',RID, a0101, a0102||'-'||Decode(RID,'51','美國','52','日本','53','韓國','54','歐洲','55','其他') C00" & _
'            Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & ",'2' " & _
'       " From acc010, (" & stVTB & ") y Where a0101='417101' And ax205(+)=a0101"
'
'    'FMP/法務
'    stVTB1 = "Select " & strArrive(1) & _
'        " From acc040 Where a0405 in ('417102','417103') " & strWhere(1) & " Group by a0405"
'
'    strSql = strSql & " Union All" & _
'        " Select '" & strUserNum & "',Decode(a0101,'417102','56','417103','57')" & _
'        " , a0101, a0102,net4,net5,net6,'2' " & _
'        " From acc010, (" & stVTB1 & ") w Where a0101 in ('417102','417103')" & _
'        " And a0405(+)=a0101"
'    '*** End FCP 收入 ***
'
'    'CFP收入/法務
'    '主科目不抓,因4131/413101都是CFP收入出現2筆 ,10506有413102科目
'    stCFP = "Select " & strArrive(1) & _
'        " From acc040 Where a0405 in ('413101','413102') " & strWhere(1) & " Group by a0405"
'
'    strSql = strSql & " Union All" & _
'        " Select '" & strUserNum & "','6z1' RID, a0101, a0102,net4,net5,net6,'2' " & _
'        " From acc010, (" & stCFP & ") w Where a0405 in ('413101','413102') And a0405(+)=a0101 "
'
'    '*** 其他的相關收入 ***
'    'Modify by Amy 2019/08/14 增加創新業務組用收入 420101
'    stVTB1 = "Select " & strArrive(1) & _
'        " From acc040 Where a0405 IN ('414101','414102','415101','415102','416101','416102','420101') " & strWhere(1) & " Group by a0405"
'
'    strSql = strSql & " Union All" & _
'        " Select '" & strUserNum & "',Decode(a0101,'414101','73','414102','74','415101','75','415102','76','416101','77','416102','78','420101','79')" & _
'        " , a0101, a0102,net4,net5,net6,'2' " & _
'        " From acc010, (" & stVTB1 & ") w Where a0101 in ('414101','414102','415101','415102','416101','416102','420101')" & _
'        " And a0405(+)=a0101 "
'    'end 2019/08/14
'
'    stVTB = "Select ax205" & strArrive(0) & _
'          " From acc020, acc021 Where ax201(+)=a0201 And ax202(+)=a0202 And ax205='7121' And ax209 is not null" & strWhere(0) & " Group by ax205"
'
'    strSql = strSql & " Union All" & _
'          " Select '" & strUserNum & "','7b', a0101, '其他收入' " & _
'                Replace(Replace(Replace(strArrive(2), "net6-", ""), "net5-", ""), "net4-", "") & ",'2' " & _
'          " From acc010, (" & stVTB & ") x Where a0101='7121' And ax205(+)=a0101"
'
'    cnnConnection.Execute "Delete From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' "
'    strSql = "Insert Into Accrpt420 (ID,r4201,r4202,r4203,r4207,r4208,R4209,R4211) " & strSql
'    cnnConnection.Execute strSql
'
'    '結餘點數、轉撥及結餘轉撥(因10601 D106012475 轉撥資料造成與frmacc44j0 當月實績不合)-婧瑄
'    'Modify by Amy 2019/08/14 增加創新業務組用收入 420101 原:SubStr(ax205, 1, 2) = '41'
'    strBP = "Select ax205,a0102 as C00," & _
'                        "Sum(Decode(Substr(a0205+19110000,1,4)," & strYear & "+1911,Nvl(ax207,0)-Nvl(ax206,0))) as C06, " & _
'                        "Sum(Decode(Substr(a0205+19110000,1,4)," & strYear - 1 & "+1911,Nvl(ax207,0)-Nvl(ax206,0))) as C07, " & _
'                        "Sum(Decode(Substr(a0205+19110000,1,4)," & strYear - 2 & "+1911,Nvl(ax207,0)-Nvl(ax206,0))) as C08 " & _
'                "From acc020, acc021,acc010 " & _
'                    "Where ax201 = a0201(+)  And ax202 = a0202(+) And  ax205=a0101(+) " & strWhere(0) & _
'                    " And SubStr(ax205, 1, 1) = '4' And Not( ax205='4191' or ax205='4192' or ax205='4194') And (InStr(ax213||' ','結餘')>0 Or InStr(ax212,'轉撥')>0)  " & _
'                "Group by ax205,a0102 "
'    strBP = "Select * From (" & strBP & ") Order by ax205"
'    intI = 1
'    Set rsNew = ClsLawReadRstMsg(intI, strBP)
'    If intI = 1 Then
'        rsNew.MoveFirst
'        Do While Not rsNew.EOF
'            strUpd = "Update Accrpt420 set R4207=R4207-(" & Val("" & rsNew.Fields("C06")) & "),R4208=R4208-(" & Val("" & rsNew.Fields("C07")) & "),R4209=R4209-(" & Val("" & rsNew.Fields("C08")) & ")" & _
'                          " Where ID='" & strUserNum & "' And R4211='2' And R4202='" & rsNew.Fields("ax205") & "' And R4203='" & rsNew.Fields("C00") & "' "
'            adoTaie.Execute strUpd
'            rsNew.MoveNext
'        Loop
'    End If
'    rsNew.Close
'    'End 結餘點數、結餘及轉撥都剔除
'
'    '拿掉「CCT-/CCP-」文字-婧瑄  ex:商標收入-CCT-MCT
'    strSql = "Update Accrpt420 set r4203=Replace(Replace(R4203,'CCT-MCT','MCT'),'CCT-MFCT','MFCT') " & _
'                "Where ID='" & strUserNum & "' And (Instr(R4203,'CCT-MCT')>0  or Instr(R4203,'CCT-MFCT')>0 ) And R4211='2' "
'    cnnConnection.Execute strSql
'    '拿掉「CCT爭議-/CCP爭議-」文字-婧瑄  ex:商標收入-CCT爭議-MCT
'    strSql = "Update Accrpt420 set r4203=Replace(Replace(R4203,'CCT爭議-MCT','MCT爭議'),'CCT爭議-MFCT','MFCT爭議') " & _
'                "Where ID='" & strUserNum & "' And (Instr(R4203,'CCT爭議-MCT')>0  or Instr(R4203,'CCT爭議-MFCT')>0 ) And R4211='2' "
'    cnnConnection.Execute strSql
'    strSql = "Update Accrpt420 set r4203=Replace(Replace(R4203,'CCP-MCP','MCP'),'CCP-MFCP','MFCP') " & _
'                "Where ID='" & strUserNum & "' And (Instr(R4203,'CCP-MCP')>0  or Instr(R4203,'CCP-MFCP')>0 ) And R4211='2' "
'    cnnConnection.Execute strSql
'    strSql = "Update Accrpt420 set r4203=Replace(R4203,'CCP爭議-MCP','MCP爭議') " & _
'                "Where ID='" & strUserNum & "' And Instr(R4203,'CCP爭議-MCP')>0 And R4211='2' "
'    cnnConnection.Execute strSql
'
'   '更新資料值,計算點數及平均,取整數
'   strSql = "/1000/" & Val(txt1(2)) - Val(txt1(1)) + 1
'   strSql = "Update Accrpt420 Set r4207=Round(Nvl(r4207,0)" & strSql & ",0),r4208=Round(Nvl(r4208,0)" & strSql & ",0),r4209=Round(Nvl(r4209,0)" & strSql & ",0) " & _
'                "Where ID='" & strUserNum & "' And R4211='2' "
'    cnnConnection.Execute strSql
'
'    '讀取暫存檔資料,RID長度為1則為加總欄位  未列於MCP/CMP/FMP 都列於CCP(不含法務-411107)
'    '內專
'    strF(0) = "Sum(Nvl(R4209,0)) as 前2年,'','',Sum(Nvl(R4208,0)) as 前1年,'','','',Sum(Nvl(R4207,0)) as 當年,'','','' "
'    strWhere(0) = "And R4202='411101' And (InStr(R4203,'MCP')>0 Or InStr(R4203,'MFCP')>0) " 'MCP
'    strWhere(1) = "And R4202='411103' " 'CMP
'    strWhere(2) = "And R4202='411106' " 'FMP
'
'    strSql = "Select '1' RID,'內專' AccN,0 as 前2年,'','',0 as 前1年,'','','',0 as 當年,'','','' From Dual " & _
'    "Union Select '11' RID,'CCP' AccN," & strF(0) & "From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' And SubStr(R4202,1,4)='4111' " & _
'                            "And Not (" & Mid(strWhere(0), 4) & ") And R4202<>'411107' " & Replace(strWhere(1) & strWhere(2), "=", "<>") & _
'    "Union Select '12' RID,'MCP' AccN," & strF(0) & "From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' " & strWhere(0) & _
'    "Union Select '13' RID,'CMP' AccN," & strF(0) & "From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' " & strWhere(1) & _
'    "Union Select '14' RID,'FMP' AccN," & strF(0) & "From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' " & strWhere(2) & _
'    "Union Select '21' RID,'CFP' AccN," & strF(0) & "From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' And R4202='413101' "
'    'FCP
'    strF(1) = "Decode(R4203, '美國', '31','日本','32','韓國','33','歐洲','34','其他國家','35' ,'FMP', '36') RID "
'    stVTB = "Select SubStr(Replace(Replace(R4203,'FCP收入-',''),'其他','其他國家'),InStr(Replace(Replace(R4203,'FCP收入-',''),'其他','其他國家'),'-')+1) R4203,R4207,R4208,R4209 " & _
'                  "From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' AND SubStr(R4202,1,4)='4171' And R4202<>'417103' "
'
'    strSql = strSql & _
'    "Union Select '3' RID,'FCP' AccN,0 as 前2年,'','',0 as 前1年,'','','',0 as 當年,'','','' From Dual " & _
'    "Union Select " & strF(1) & ",R4203 AccN," & strF(0) & " From (" & stVTB & ") Group by R4203 " & _
'    "Union Select '3Z' RID,'上述專利總計' AccN,0 as 前2年,'','',0 as 前1年,'','','',0 as 當年,'','','' From Dual "
'
'    '內商
'    strWhere(0) = "And (R4202='410101' Or R4202='410104') And (InStr(R4203,'MCT')>0 Or InStr(R4203,'MFCT')>0) " 'MCT
'    strWhere(1) = "And R4202='410103' " 'CMT
'    strWhere(2) = "And R4202='410109' " 'FMT
'
'    'Modify by Amy 2019/08/15 著作權併至CCT
'    strSql = strSql & _
'    "Union Select '4' RID,'內商' AccN,0 as 前2年,'','',0 as  前1年,'','','',0 as  當年,'','','' From Dual " & _
'    "Union Select '41' RID,'CCT' AccN," & strF(0) & "From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' And (SubStr(R4202,1,4)='4101' Or R4202='417202' or SubStr(R4202,1,3)='415')  " & _
'                            "And Not (" & Mid(strWhere(0), 4) & ") And R4202<>'410110' " & Replace(strWhere(1) & strWhere(2), "=", "<>") & _
'    "Union Select '42' RID,'MCT' AccN," & strF(0) & "From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' " & strWhere(0) & _
'    "Union Select '43' RID,'CMT' AccN," & strF(0) & "From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' " & strWhere(1) & _
'    "Union Select '44' RID,'FMT' AccN," & strF(0) & "From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' " & strWhere(2) & _
'    "Union Select '51' RID,'CFT' AccN," & strF(0) & "From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' And R4202='412101' "
'    'FCT
'    strF(1) = "Decode(r4203, '美國','61' ,'日本','62','韓國','63','歐洲','64','其他國家','65','FMT', '66') RID "
'    stVTB = "Select SubStr(Replace(Replace(R4203,'FCT收入-',''),'其他','其他國家'),InStr(Replace(Replace(R4203,'FCT收入-',''),'其他','其他國家'),'-')+1) R4203,R4207,R4208,R4209 " & _
'                  "From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' And SubStr(R4202,1,4)='4172' And R4202<>'417203' And R4202<>'417202' "
'
'    strSql = strSql & _
'    "Union Select '6' RID,'FCT' AccN,0 as 前2年,'','',0 as 前1年,'','','',0 as 當年,'','','' From Dual " & _
'    "Union Select " & strF(1) & ",R4203 AccN," & strF(0) & " From (" & stVTB & ") Group by R4203 " & _
'    "Union Select '6Z' RID,'上述商標總計' AccN,0 as 前2年,'','',0 as 前1年,'','','',0 as 當年,'','','' From Dual "
'
'    '法務
'    strF(1) = "Decode(R4203,'FCP-法務','71','FCT-法務','72','內專-法務','73','CFP-法務','74','內商-法務','75','CFT-法務','76') RID "
'    strWhere(0) = "And R4202 In('417103','417203','411107','413102','410110','412102') "
'    stVTB = "Select Replace(Replace(Replace(R4203,'收入',''),'商標-CCT','內商-'),'專利-CCP','內專-') R4203,R4207,R4208,R4209 " & _
'                  " From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' " & strWhere(0)
'
'    '部份法務科目某些年度有值才列示,列於最後-婧瑄
'    strSql = strSql & _
'    "Union Select '7' RID,'法務' AccN,0 as 前2年,'','',0 as 前1年,'','','',0 as 當年,'','','' From Dual " & _
'    "Union Select " & strF(1) & ",R4203 AccN," & strF(0) & " From (" & stVTB & ") Group by R4203 " & _
'    "Union Select '7'||SubStr(R4202,3,1)||SubStr(R4202,Length(R4202),1) RID,R4203 AccN," & strF(0) & _
'                " From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' And (Nvl(R4207,0)<>0 Or Nvl(R4208,0)<>0 Or Nvl(R4209,0)<>0 )" & _
'                Replace(strWhere(0), "And R4202 In", "And R4202 Not In") & " And SubStr(R4202,1,3) In('414','416') " & _
'                "Group by '7'||SubStr(R4202,3,1)||SubStr(R4202,Length(R4202),1),R4203 "
'
'    '其他
'    'Modify by Amy 2019/08/14 增加創新業務組用收入 420101,放於 其他 之前
'    'Moidfy by Amy 2019/08/15 著作權併入CCT,調整其他順序
'    '"Union Select '81' RID,'著作權' AccN," & strF(0) & " From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' And SubStr(R4202,1,3)='415' "
'    strSql = strSql & _
'    "Union Select '81' RID,'創新業務' AccN," & strF(0) & " From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' And R4202='420101' " & _
'    "Union Select '91' RID,'其他' AccN," & strF(0) & " From Accrpt420 Where ID='" & strUserNum & "' And R4211='2' And R4202='7121' " & _
'    "Union Select 'ZZ' RID,'全所實績總計' AccN,0 as 前2年,'','',0 as 前1年,'','','',0 as 當年,'','','' From Dual " & _
'    "Order by RID"
'
'    intI = 1
'    Set rsNew = ClsLawReadRstMsg(intI, strSql)
'    If intI = 1 Then
'        rsNew.MoveFirst
'        If ExcelSave3 = True Then
'            MsgBox "已產生EXCEL檔案...", , MsgText(5)
'        End If
'        rsNew.Close
'    Else
'        MsgBox "無符合資料！"
'    End If
'
'ErrHnd:
'    If Err.Number <> 0 Then
'        If rsNew.State <> adStateClosed Then rsNew.Close
'        MsgBox Err.Description
'    End If
End Sub

'intChoose:0-有代理人(MCT+MFCT)/1-MCT(有代理人且客戶國籍為大陸)/2-MFCT(有代理人且客戶國籍非大陸)
'Modify by Amy 2022/02/07 stWhere 改 Optional
Private Function GetCCT(ByVal intChoose As Integer, ByVal stField As String, Optional ByVal stWhere As String = "") As String
    Dim stTmp(1) As String, stTmp2 As String
   
    stTmp(0) = "And tm44 is not null "
    stTmp(1) = "And sp26 is not null "
    
    '用Nvl(cu10,fa10) 因TS001663 無申請人
    If intChoose = 1 Then
        stTmp(0) = "And tm44 is not null And Nvl(cu10,fa10)='020'"
        stTmp(1) = "And sp26 is not null And Nvl(cu10,fa10)='020'"
    ElseIf intChoose = 2 Then
        stTmp(0) = "And tm44 is not null And Nvl(cu10,fa10)<>'020'"
        stTmp(1) = "And sp26 is not null And Nvl(cu10,fa10)<>'020'"
    End If
                                     
    'Modify by Amy 2022/02/07 改抓暫存檔,+排除結餘/轉撥
'    GetCCT = "select a0205, ax205, ax206, ax207 from acc020, acc021, trademark, fagent, customer" & _
'        " where ax201(+)=a0201 And ax202(+)=a0202 And ax205 IN ('410101','410104') And ax209 is not null" & stWhere & _
'        " And tm01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) And tm02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'        " And tm03(+)=substr(lpad(ax214,12,' '),10,1) And tm04(+)=substr(lpad(ax214,12,' '),11,2) And tm01 is not null" & _
'        " And fa01(+)=substr(tm44,1,8) And fa02(+)=substr(tm44,9,1)" & _
'        " And cu01(+)=substr(tm23,1,8) And cu02(+)=substr(tm23,9,1) " & stTmp(0) & _
'        " Union All" & _
'        " select a0205, ax205, ax206, ax207 from acc020, acc021, servicepractice, fagent, customer" & _
'        " where  ax201(+)=a0201 And ax202(+)=a0202 And ax205 IN ('410101','410104') And ax209 is not null" & stWhere & _
'        " And sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) And sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'        " And sp03(+)=substr(lpad(ax214,12,' '),10,1) And sp04(+)=substr(lpad(ax214,12,' '),11,2) And (sp01 is not null or ax214 is null)" & _
'        " And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9,1)" & _
'        " And cu01(+)=substr(sp08,1,8) And cu02(+)=substr(sp08,9,1) " & stTmp(1)
    GetCCT = "Select R001 as a0205,R005 as ax205,R006 as ax206,R007 as ax207 From Accrpt44r0Dept, trademark, fagent, customer" & _
        " Where tm01(+)=R008 And tm02(+)=R009 And tm03(+)=R010 And tm04(+)=R011 And tm01 is not null" & stWhere & _
        " And (InStr(R012,'結餘')=0 Or InStr(R012,'轉撥')=0)" & _
        " And fa01(+)=substr(tm44,1,8) And fa02(+)=substr(tm44,9,1)" & _
        " And cu01(+)=substr(tm23,1,8) And cu02(+)=substr(tm23,9,1) " & stTmp(0) & _
        " Union All" & _
        " Select R001 as a0205,R005 as ax205,R006 as ax206,R007 as ax207 From Accrpt44r0Dept, servicepractice, fagent, customer" & _
        " Where sp01(+)=R008 And sp02(+)=R009 And sp03(+)=R010 And sp04(+)=R011 And (sp01 is not null or R008 is null)" & stWhere & _
        " And (InStr(R012,'結餘')=0 Or InStr(R012,'轉撥')=0)" & _
        " And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9,1)" & _
        " And cu01(+)=substr(sp08,1,8) And cu02(+)=substr(sp08,9,1) " & stTmp(1)
        
    GetCCT = "Select ax205" & stField & " From (" & GetCCT & ") x Group By ax205"
End Function

'intChoose:0-有代理人(MCP+MFCP)/1-MCP(有代理人且客戶國籍為大陸)/2-MFCP(有代理人且客戶國籍非大陸)
Private Function GetCCP(ByVal intChoose As Integer, ByVal stField As String, ByVal stWhere As String) As String
    Dim stQ As String, stTmp(1) As String, stTmp2 As String
   
    stTmp(0) = "And pa75 is not null "
    stTmp(1) = "And sp26 is not null "
    
    '用Nvl(cu10,fa10) 因TS001663 無申請人
    If intChoose = 1 Then
        stTmp(0) = "And pa75 is not null And Nvl(cu10,fa10)='020'"
        stTmp(1) = "And sp26 is not null And Nvl(cu10,fa10)='020'"
    ElseIf intChoose = 2 Then
        stTmp(0) = "And pa75 is not null And Nvl(cu10,fa10)<>'020'"
        stTmp(1) = "And sp26 is not null And Nvl(cu10,fa10)<>'020'"
    End If
    
    'Modify by Amy 2022/02/07 改暫存檔,+排除結餘/轉撥
'    GetCCP = "select a0205, ax205, ax206, ax207 from acc020, acc021, patent, fagent, customer" & _
'        " where ax201(+)=a0201 And ax202(+)=a0202 And ax205='411101' And ax209 is not null" & stWhere & _
'        " And pa01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) And pa02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'        " And pa03(+)=substr(lpad(ax214,12,' '),10,1) And pa04(+)=substr(lpad(ax214,12,' '),11,2) And pa01 is not null" & _
'        " And fa01(+)=substr(pa75,1,8) And fa02(+)=substr(pa75,9,1)" & _
'        " And cu01(+)=substr(pa26,1,8) And cu02(+)=substr(pa26,9,1) " & stTmp(0) & _
'        " Union All" & _
'        " select a0205, ax205, ax206, ax207 from acc020, acc021, servicepractice, fagent, customer" & _
'        " where  ax201(+)=a0201 And ax202(+)=a0202 And ax205='411101' And ax209 is not null" & stWhere & _
'        " And sp01(+)=ltrim(substr(lpad(ax214,12,' '),1,3)) And sp02(+)=substr(lpad(ax214,12,' '),4,6)" & _
'        " And sp03(+)=substr(lpad(ax214,12,' '),10,1) And sp04(+)=substr(lpad(ax214,12,' '),11,2) And ( sp01 is not null or ax214 is null)" & _
'        " And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9,1)" & _
'        " And cu01(+)=substr(sp08,1,8) And cu02(+)=substr(sp08,9,1) " & stTmp(1)
    GetCCP = "Select R001 as a0205,R005 as ax205,R006 as ax206,R007 as ax207 From Accrpt44r0Dept, patent, fagent, customer" & _
        " Where pa01(+)=R008 And pa02(+)=R009 And pa03(+)=R010 And pa04(+)=R011 And pa01 is not null" & stWhere & _
        " And (InStr(R012,'結餘')=0 Or InStr(R012,'轉撥')=0)" & _
        " And fa01(+)=substr(pa75,1,8) And fa02(+)=substr(pa75,9,1)" & _
        " And cu01(+)=substr(pa26,1,8) And cu02(+)=substr(pa26,9,1) " & stTmp(0) & _
        " Union All" & _
        " Select R001 as a0205,R005 as ax205,R006 as ax206,R007 as ax207 From Accrpt44r0Dept, servicepractice, fagent, customer" & _
        " Where sp01(+)=R008 And sp02(+)=R009 And sp03(+)=R010 And sp04(+)=R011 And ( sp01 is not null or R008 is null) " & stWhere & _
        " And (InStr(R012,'結餘')=0 Or InStr(R012,'轉撥')=0)" & _
        " And fa01(+)=substr(sp26,1,8) And fa02(+)=substr(sp26,9,1)" & _
        " And cu01(+)=substr(sp08,1,8) And cu02(+)=substr(sp08,9,1) " & stTmp(1)
        
      GetCCP = "Select ax205" & stField & " From (" & GetCCP & ") x Group By ax205"
End Function

'專業[單位]實績點數分析表
Private Function ExcelSave3() As Boolean
    Dim xlsSalesPoint As New Excel.Application
    Dim Wks As New Worksheet
    Dim j As Integer
    Dim stReportN As String, strOldRID As String '報表名稱/前一筆RID
    Dim strSum As String, strAllSum As String '專業單位加總位置/全所加總位置
    Dim intTitle As Integer, intField As Integer, intCounter As Integer, intToTR As Integer
    Dim intSumState As Integer '1.大項/2.專業/3.全所
    Dim stTmp(3) As String
    Dim arrTmp As Variant, arrTmpS As Variant '當筆/加總
    
On Error GoTo onErrHand
    'Memo by Amy 加L公司時未調整此報表,架構是源自於智慧所,法律所不適用仍不修改-婧瑄 mail 111/08/16 1:51
    ExcelSave3 = False
    stReportN = "專業單位實績點數分析表" 'Modify by Amy 2022/02/07 +實績
    If Dir(strExcelPath & strCmp & " " & stReportN & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strCmp & " " & stReportN & ACDate(ServerDate) & ServerTime & MsgText(43)
    End If
    
    ReDim strField1(11):  ReDim intWidth1(11): ReDim arrTmp(1 To UBound(strField1) + 1): ReDim arrTmpS(1 To UBound(strField1) + 1)
    strField1 = Array("專業單位", Val(strYear - 2) & "月平均", Val(strYear - 2) & "佔單位比率", Val(strYear - 2) & "佔全所比例", _
                                     Val(strYear - 1) & "月平均", Val(strYear - 1) & "佔單位比率", Val(strYear - 1) & "佔全所比例", Val(strYear - 1) & "較前一年增減", _
                                     Val(strYear - 0) & "月平均", Val(strYear - 0) & "佔單位比率", Val(strYear - 0) & "佔全所比例", Val(strYear - 0) & "較前一年增減")
    'Modify by Amy 2022/02/07 月平均欄原6改13
    intWidth1 = Array(12, 13, 12, 12, 13, 12, 12, 13, 13, 12, 12, 13)
    
    bolXlsOpen = False: intField = 65: intCounter = 1
    
    xlsSalesPoint.SheetsInNewWorkbook = 3 '預設工作表數量
    xlsSalesPoint.Workbooks.add
    Set Wks = xlsSalesPoint.Worksheets(1)
    bolXlsOpen = True
    
    '表頭
    Wks.Range(Chr(intField) & intCounter).Value = stReportN
    Wks.Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strField1)) & intCounter).Select
  
    With xlsSalesPoint.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
    End With
    
    intCounter = intCounter + 2
    For i = LBound(strField1) To UBound(strField1)
        Wks.Range(Chr(intField + i) & intCounter).Value = strField1(i)
        Wks.Range(Chr(intField + i) & intCounter).ColumnWidth = intWidth1(i)
    Next i
    intTitle = intCounter
    
    '資料(ProduceData_Dept產生)
    intCounter = intCounter + 1:  intToTR = intCounter: intSumState = 0
    Do While rsNew.EOF = False
        '大項目加總欄(不印資料)
        If Len("" & rsNew.Fields("RID")) = 1 Then
            For j = LBound(arrTmpS) To UBound(arrTmpS)
                arrTmpS(j) = ""
            Next j
            arrTmpS(GetValue1("專業單位") + 1) = "" & rsNew.Fields("AccN")
            strSum = strSum & "," & Chr(intField + GetValue1(Val(strYear - 2) & "月平均")) & intCounter
            strAllSum = strAllSum & "," & Chr(intField + GetValue1(Val(strYear - 2) & "月平均")) & intCounter
            intToTR = intCounter '記錄加總欄位
        Else
            For i = LBound(strField1) To UBound(strField1)
                arrTmp(i + 1) = "" & rsNew.Fields(i + 1)
                If Right(strField1(i), 5) <> "佔全所比例" Then
                    '****** 加總欄位 ******
                    If Left("" & rsNew.Fields("RID"), 1) <> Left(strOldRID, 1) Or Right("" & rsNew.Fields("RID"), 1) = "Z" Then
                        '全所
                        If "" & rsNew.Fields("RID") = "ZZ" Then
                            intSumState = 3
                        '專業
                        ElseIf Right("" & rsNew.Fields("RID"), 1) = "Z" Then
                            intSumState = 2
                        '大項
                        Else
                            intSumState = 1
                            If i = GetValue1(Val(strYear - 2) & "月平均") Then
                                '專業單位加總=大項加總位置
                                strSum = strSum & "," & Chr(intField + i) & intCounter
                                '全所加總=所有大項加總位置
                                strAllSum = strAllSum & "," & Chr(intField + i) & intCounter
                            End If
                        End If
                        '*** 大項/專業 加總最後一欄將資料以陣列方式寫至儲存格 ***
                        If intCounter - 1 > intToTR And intSumState <> 3 And i = UBound(strField1) Then
                            '調整位置(因加總位置在上方)
                            stTmp(0) = GetValue1(Val(strYear - 2) & "月平均")
                            stTmp(1) = GetValue1(Val(strYear - 1) & "月平均")
                            stTmp(2) = GetValue1(Val(strYear - 0) & "月平均")
                            For j = LBound(stTmp) To UBound(stTmp) - 1
                                arrTmpS(Val(stTmp(j)) + 1) = "=Sum(" & Chr(intField + Val(stTmp(j))) & intToTR + 1 & ":" & Chr(intField + Val(stTmp(j))) & intCounter - 1 & ")"
                                '較前一年增減
                                If j > 0 Then
                                    arrTmpS(GetValue1(Val(strYear - (2 - j)) & "較前一年增減") + 1) = "=" & Chr(intField + Val(stTmp(j))) & intToTR & "-" & Chr(intField + Val(stTmp(j - 1))) & intToTR
                                End If
                            Next j
                            Wks.Range(Chr(intField) & intToTR & ":" & Chr(intField + i) & intToTR).Value = arrTmpS
                            intToTR = intCounter '創新業務81會 Run此,因前一個法務資料需做大項合計
                        'Modfiy by Amy 2019/08/14 +ElseIf 增加91創新業務,某些資料無大項,直接再接非大項項目 ex:創新業務91和其他A1
                        'Memo by Amy 2019/08/15 創新業務91改81 其他A1改91,之後可能增加程式先不改
                        'Modify by Amy 2022/02/07 +安全基金撥補(490101)-B1
                        ElseIf rsNew.Fields("RID") = "91" Or rsNew.Fields("RID") = "A1" Or rsNew.Fields("RID") = "B1" Then
                            intToTR = intCounter
                        End If
                        '*** End 大項/專業 加總最後一欄將資料以陣列方式寫至儲存格 ***
                    End If
                    '****** end 加總欄位 ******
                    
                    If Right(strField1(i), 6) = "較前一年增減" Then
                        '較前一年增減=欄位年月平均-前一年月平均
                        If Mid(strField1(i), 1, InStr(strField1(i), "較前一年增減") - 1) = Val(strYear - 1) Then
                            arrTmp(i + 1) = "=" & Chr(intField + GetValue1(Val(strYear - 1) & "月平均")) & intCounter & "-" & Chr(intField + GetValue1(Val(strYear - 2) & "月平均")) & intCounter
                        Else
                            arrTmp(i + 1) = "=" & Chr(intField + GetValue1(Val(strYear) & "月平均")) & intCounter & "-" & Chr(intField + GetValue1(Val(strYear - 1) & "月平均")) & intCounter
                        End If
                    '加總(專業單位/全所)
                    ElseIf intSumState > 1 Then
                        stTmp(0) = "" '佔單位比率/月平均
                        '月平均
                        If Right(strField1(i), 3) = "月平均" Then
                            '全所
                            If intSumState = 3 Then
                                stTmp(0) = Mid(strAllSum, 2)
                            '專業
                            Else
                                stTmp(0) = Mid(strSum, 2)
                            End If
                            If i <> GetValue1(Val(strYear - 2) & "月平均") Then
                                stTmp(0) = Replace(stTmp(0), Chr(intField + GetValue1(Val(strYear - 2) & "月平均")), Chr(intField + GetValue1("" & strField1(i))))
                            End If
                        End If
                        If i <> GetValue1("專業單位") Then arrTmp(i + 1) = IIf(stTmp(0) = "", "", "=Sum(" & stTmp(0) & ")")
                        
                    '一般資料
                    Else
                        If Right(strField1(i), 3) = "月平均" Then
                            arrTmp(i + 1) = Val(arrTmp(i + 1))
                        ElseIf Right(strField1(i), 5) = "佔單位比率" Then
                            If intSumState = 1 Then
                                 arrTmp(i + 1) = ""
                            Else
                                '佔單位比率=月平均/單位加總月平均
                                stTmp(0) = GetValue1(Replace(strField1(i), "佔單位比率", "") & "月平均")
                                arrTmp(i + 1) = "=" & Chr(intField + Val(stTmp(0))) & intCounter & "/" & Chr(intField + Val(stTmp(0))) & intToTR
                            End If
                        End If 'end Right(strField1(i), 3) = "月平均"
                    End If
                    
                    '格式/資料 設定
                    If i = GetValue1("專業單位") Then
                        Select Case intSumState
                            Case 0 '一般資料
                                Wks.Range(Chr(intField) & intCounter).HorizontalAlignment = xlRight
                            Case 1 '大項
                                Wks.Range(Chr(intField) & intCounter).HorizontalAlignment = xlLeft
                            Case 2, 3 '專業,全所
                                Wks.Range(Chr(intField) & intCounter).HorizontalAlignment = xlCenter
                                Wks.Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strField1)) & intCounter).Font.ColorIndex = 53  '設置儲存格填充色(膚)
                        End Select
                    ElseIf i = UBound(strField1) Then
                        '最後一欄將資料以陣列方式寫至儲存格
                        Wks.Range(Chr(intField) & intCounter & ":" & Chr(intField + i) & intCounter).Value = arrTmp
                        If Right("" & rsNew.Fields("RID"), 1) = "Z" Then
                            strSum = ""
                        End If
                    End If
                    
                End If 'end Right(strField1(i), 5) <> "佔全所比例"
                
            Next i
        End If 'end 大項目加總欄(不印資料)
        
        intCounter = intCounter + 1
        strOldRID = "" & rsNew.Fields("RID")
        intSumState = 0
        rsNew.MoveNext
    Loop
    
    'Add by Amy 2022/02/07 +月平均顯示小數2位
    For i = 0 To 2
        stTmp(0) = Chr(intField + GetValue1(Val(strYear - i) & "月平均"))
        Wks.Range(stTmp(0) & intTitle + 1 & ":" & stTmp(0) & intCounter - 1).NumberFormatLocal = "#,##0.00"
    Next i
        
    '佔全所比例公式
    stTmp(0) = Chr(intField + GetValue1(Val(strYear - 2) & "月平均"))
    stTmp(1) = Chr(intField + GetValue1(Val(strYear - 1) & "月平均"))
    stTmp(2) = Chr(intField + GetValue1(Val(strYear - 0) & "月平均"))
    For i = intTitle + 1 To intCounter - 2
        Wks.Range(Chr(intField + GetValue1(Val(strYear - 2) & "佔全所比例")) & i).Formula = _
                "=" & stTmp(0) & i & "/$" & stTmp(0) & "$" & intCounter - 1
        Wks.Range(Chr(intField + GetValue1(Val(strYear - 1) & "佔全所比例")) & i).Formula = _
                "=" & stTmp(1) & i & "/$" & stTmp(1) & "$" & intCounter - 1
        Wks.Range(Chr(intField + GetValue1(Val(strYear - 0) & "佔全所比例")) & i).Formula = _
                "=" & stTmp(2) & i & "/$" & stTmp(2) & "$" & intCounter - 1
    Next i
    For i = 2 To 0 Step -1
        stTmp(0) = Chr(intField + GetValue1(Val(strYear - i) & "佔單位比率")) & intTitle + 1 & ":" & Chr(intField + GetValue1(Val(strYear - i) & "佔全所比例")) & intCounter - 2
        Wks.Range(stTmp(0)).NumberFormatLocal = "0.00%"
    Next i
    
    stTmp(0) = " / " & Val(strStartMonth) & "-" & Val(strEndMonth) & "月"
    stTmp(1) = Chr(intField + GetValue1(Val(strYear - 2) & "月平均"))
    For i = GetValue1(Val(strYear - 2) & "月平均") To UBound(strField1)
        If Right(strField1(i), 3) = "月平均" Then
            Wks.Range(Chr(intField + i) & intTitle - 1).Value = Replace(strField1(i), "月平均", "") & "年" & stTmp(0)
            If i > 1 Then
                Wks.Range(stTmp(1) & intTitle - 1 & ":" & Chr(intField + i - 1) & intTitle).Interior.ColorIndex = IIf(i = GetValue1(Val(strYear) - 1 & "月平均"), 20, 40)
                Wks.Range(stTmp(1) & intTitle - 1 & ":" & Chr(intField + i - 1) & intTitle).Interior.tintandshade = 0.5 '設深淺
                Wks.Range(stTmp(1) & intTitle - 1 & ":" & Chr(intField + i - 1) & intTitle - 1).MergeCells = True
                Wks.Range(stTmp(1) & intTitle - 1 & ":" & Chr(intField + i - 1) & intTitle - 1).HorizontalAlignment = xlCenter
                stTmp(1) = Chr(intField + i)
            End If
        End If
        Wks.Range(Chr(intField + i) & intTitle).Value = Replace(Replace(Replace(strField1(i), strYear, ""), strYear - 1, ""), strYear - 2, "")
        Wks.Range(Chr(intField + i) & intTitle).HorizontalAlignment = xlCenter
    Next i
    Wks.Range(stTmp(1) & intTitle - 1 & ":" & Chr(intField + UBound(strField1)) & intTitle).Interior.ColorIndex = 40
    Wks.Range(stTmp(1) & intTitle - 1 & ":" & Chr(intField + UBound(strField1)) & intTitle).Interior.tintandshade = 0.5  '設深淺
    Wks.Range(stTmp(1) & intTitle - 1 & ":" & Chr(intField + UBound(strField1)) & intTitle - 1).MergeCells = True
    Wks.Range(stTmp(1) & intTitle - 1 & ":" & Chr(intField + UBound(strField1)) & intTitle - 1).HorizontalAlignment = xlCenter
    Wks.Range(Chr(intField + GetValue1("專業單位")) & intTitle - 1 & ":" & Chr(intField + GetValue1("專業單位")) & intTitle).MergeCells = True
    
    '框線
    Wks.Range(Chr(intField) & "1:" & Chr(intField + UBound(strField1)) & intCounter - 1).Select
    xlsSalesPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    xlsSalesPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
    Wks.PageSetup.PaperSize = 9 'A4
    Wks.PageSetup.Orientation = xlLandscape '橫印
    Wks.PageSetup.PrintTitleRows = "$1:$" & intTitle '表頭保留
    Wks.PageSetup.LeftMargin = xlsSalesPoint.InchesToPoints(0.5)
    Wks.PageSetup.RightMargin = xlsSalesPoint.InchesToPoints(0.5)
    Wks.PageSetup.TopMargin = xlsSalesPoint.InchesToPoints(0.2)
    Wks.PageSetup.BottomMargin = xlsSalesPoint.InchesToPoints(0.2)
    Wks.PageSetup.HeaderMargin = xlsSalesPoint.InchesToPoints(0.3)
    Wks.PageSetup.FooterMargin = xlsSalesPoint.InchesToPoints(0.3)
    Wks.PageSetup.Zoom = False '100 '縮放比例
    
    '判斷若版本2007以上改變存格式
    'Modify by Amy 2022/02/07 檔名+公司別
    If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCmp & " " & stReportN & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
    Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCmp & " " & stReportN & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
    End If
    'end 2022/02/07
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    Set xlsSalesPoint = Nothing
    ExcelSave3 = True
    StatusClear
    Exit Function

onErrHand:
    If Err.Number <> 0 And bolXlsOpen = True Then
        MsgBox Err.Description, , MsgText(5)
        'Modify by Amy 2022/02/07 檔名+公司別
        If Val(xlsSalesPoint.Version) < 12 Then
            xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCmp & " " & stReportN & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
        Else
            xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCmp & " " & stReportN & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
        End If
        'end 2022/02/07
        xlsSalesPoint.Workbooks.Close
        xlsSalesPoint.Quit
        Set xlsSalesPoint = Nothing
    End If

End Function

'Add by Amy 2021/02/26 專業達成點數表(11001月後 秘書用新格式)
Private Sub ExcelSave4()
    Dim xlsSalesPoint As New Excel.Application
    Dim Wks As New Worksheet
    Dim bolSetData As Boolean
    Dim intLCol As Integer, strRCol As String, intRTitle As Integer '右表對應左表欄位/記錄右表欄/右表抬頭列
    Dim strOldN As String, strFormat As String, strTmp As String, strTmp2 As String
    Dim strTOTR As String 'Add by Amy 2021/05/18 全所合計位置
    
On Error GoTo onErrHand

    If Dir(strExcelPath & Text3 & "年度" & Text4 & "月專業達成點數表" & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & Text3 & "年度" & Text4 & "月專業達成點數表" & ACDate(ServerDate) & ServerTime & MsgText(43)
    End If
    
    ReDim strField1(4): ReDim intWidth1(4): ReDim strField2(5): ReDim intWidth2(5)
    strField1 = Array("專利", "", "目標", "實績", "報出")
    intWidth1 = Array(6, 18, 10, 10, 10)
    
    'Modify by Amy 2021/11/09 加勾選「結餘+轉撥」欄-查帳用
    'Modify by Amy 2023/12/07 加勾選「顯示小數5位」
    If ChkChoose.Value = vbChecked Then
        If ChkD5.Value = vbChecked Then
            strField2 = Array("單位", "目標", "實績點數", "實績點數(小數5位)", "實績達成率", "報出點數", "報出點數(小數5位)", "報出達成率", "結餘+轉撥")
            intWidth2 = Array(18, 8.5, 10, 20, 11, 10, 20, 11, 11)
        Else
            strField2 = Array("單位", "目標", "實績點數", "實績達成率", "報出點數", "報出達成率", "結餘+轉撥")
            intWidth2 = Array(18, 8.5, 10, 11, 10, 11, 11)
        End If
    ElseIf ChkD5.Value = vbChecked Then
        strField2 = Array("單位", "目標", "實績點數", "實績點數(小數5位)", "實績達成率", "報出點數", "報出點數(小數5位)", "報出達成率")
        intWidth2 = Array(18, 8.5, 10, 20, 11, 10, 20, 11)
    Else
        strField2 = Array("單位", "目標", "實績點數", "實績達成率", "報出點數", "報出達成率")
        intWidth2 = Array(18, 8.5, 10, 11, 10, 11)
    End If
    
    bolXlsOpen = False: intField = 65: intRow1 = 1
    intField2 = intField + UBound(strField1) + 2
    xlsSalesPoint.SheetsInNewWorkbook = 3 '預設工作表數量
    xlsSalesPoint.Workbooks.add
    Set Wks = xlsSalesPoint.Worksheets(1)
    bolXlsOpen = True
    xlsSalesPoint.Visible = True
    
    Call PrintExcel4Title(xlsSalesPoint, Wks)
    Call GetLawValue
    
    intRTitle = intRow1 - 1
    intLStart = intRow2
    adoquery.MoveFirst
    Do While adoquery.EOF = False
        '右表
        For i = LBound(strField2) To UBound(strField2)
            strRCol = "": strFormat = ""
            'Modify by Amy 2023/12/07 +小數5位,原用if判斷
            Select Case strField2(i)
               Case "實績點數", "實績點數(小數5位)"
                  strTmp = "" & adoquery.Fields("C02")
               Case "報出點數", "報出點數(小數5位)"
                  strTmp = "" & adoquery.Fields("C04")
               'Add by Amy 2021/11/09 +「結餘+轉撥」欄
               Case "結餘+轉撥"
                  strTmp = "" & adoquery.Fields("C06")
               Case Else
                  strTmp = "" & adoquery.Fields(i)
            End Select
            'end 2023/12/07
            
            'Modify by Amy 2021/11/09 +strField2(i) = "結餘+轉撥"
            'Modify by Amy 2023/12/07 +strField2(i) = "xxx(小數5位)"
            If "" & adoquery.Fields("C01") = "全　所" And (i = GetValue2("目標", True) Or i = GetValue2("實績點數", True) Or i = GetValue2("報出點數", True) _
                     Or strField2(i) = "結餘+轉撥" Or strField2(i) = "實績點數(小數5位)" Or strField2(i) = "報出點數(小數5位)") Then
                strTmp = Chr(i + intField2) & intRTitle + 1 & ":" & Chr(i + intField2) & intRow2 - 1
                strTmp = "=Sum(" & strTmp & ")"
                'Add by Amy 2021/05/18 記錄全所位置
                If i = GetValue2("目標", True) Then
                    strTOTR = intRow2
                End If
            ElseIf i = GetValue2("目標", True) Then
                strFormat = "#,##0"
                strRCol = Chr(i + intField2) '記錄目前位置
            ElseIf i = GetValue2("實績達成率", True) Then
                strFormat = "0.00%"
                strTmp = Chr(GetValue2("實績點數", True) + intField2) & intRow2 & "/" & Chr(GetValue2("目標", True) + intField2) & intRow2
                strTmp = "=IF(" & Chr(GetValue2("目標", True) + intField2) & intRow2 & "=0,0," & strTmp & ")"
            ElseIf i = GetValue2("報出達成率", True) Then
                strFormat = "0.00%"
                strTmp = Chr(GetValue2("報出點數", True) + intField2) & intRow2 & "/" & Chr(GetValue2("目標", True) + intField2) & intRow2
                strTmp = "=IF(" & Chr(GetValue2("目標", True) + intField2) & intRow2 & "=0,0," & strTmp & ")"
            ElseIf i <> GetValue2("單位", True) Then
                strTmp2 = "" & adoquery.Fields("C01")
                'Modify by Amy 2021/11/09 +「結餘+轉撥」欄
                If strField2(i) <> "結餘+轉撥" Then
                    If Left(strTmp2, 2) = "法務" Or strTmp2 = "一般法務" Then
                        If strTmp2 <> "一般法務" Then strTmp2 = Mid(strTmp2, Val(InStr(strTmp2, "- ")) + 2)
                        Select Case strTmp2
                            Case "P"
                               strTmp = dblLawVal(1)
                            Case "CFP"
                                strTmp = dblLawVal(7)
                            Case "FCP"
                                strTmp = dblLawVal(3)
                            Case "T"
                                strTmp = dblLawVal(2)
                            Case "CFT"
                                strTmp = dblLawVal(5)
                            Case "FCT"
                                strTmp = dblLawVal(4)
                            Case "一般法務"
                                strTmp = dblLawVal(8)
                        End Select
                    End If
                End If
                If InStr(strField2(i), "(小數5位)") > 0 Then
                  strTmp = "=Round(" & strTmp & "/1000,5)"
                  strFormat = "#,##0.00000"
                Else
                  'Modify by Amy 2021/10/18 +取小數2位
                  strTmp = "=Round(" & strTmp & "/1000,2)"
                  strFormat = "#,##0.00"
                  strRCol = Chr(i + intField2) '記錄目前位置
                End If
            End If
            Wks.Range(Chr(intField2 + i) & intRow2).Value = strTmp
            '對齊
            If i = GetValue2("單位", True) Then
                Wks.Range(Chr(intField2 + i) & intRow2).HorizontalAlignment = xlLeft
            Else
                Wks.Range(Chr(intField2 + i) & intRow2).HorizontalAlignment = xlRight
            End If
            '格式
            If strFormat <> MsgText(601) Then
                Wks.Range(Chr(intField2 + i) & intRow2).NumberFormatLocal = strFormat
            End If
            
            '*** 左表 ***
            'Modify by Amy 2021/11/09 左表不需顯示「結餘+轉撥」欄
            'Modify by Amy 2023/12/07 左表不需顯示「xxx(小數5位)」欄
            If strField2(i) <> "結餘+轉撥" And InStr(strField2(i), "(小數5位)") = 0 Then
                bolSetData = False: strFormat = "#,##0.00"
                '商標部 - 著作權 左表不出現
                If "" & adoquery.Fields("C01") = "商標部 - 著作權" Then
                Else
                    If Left("" & adoquery.Fields("C01"), 2) = "專利" Or (Left("" & adoquery.Fields("C01"), 2) = "商標" And Left("" & adoquery.Fields(i), 2) <> "商標部 - 著作權") Or Left("" & adoquery.Fields("C01"), 2) = "法務" Or Right("" & adoquery.Fields("C01"), 2) = "法務" Then
                        If i = GetValue2("單位", True) Then
                            '增加上個Table總計/總達成率
                            If Left(strOldN, 2) <> Left("" & adoquery.Fields("C01"), 2) And strOldN <> MsgText(601) Then
                                If Not (InStr(strOldN, "法務") > 0 And "" & adoquery.Fields("C01") = "一般法務") Then
                                    Call PrintExcel4Title(xlsSalesPoint, Wks, True, Left("" & adoquery.Fields("C01"), 2))
                                    intLStart = intRow1
                                End If
                            End If
                            Wks.Range(Chr(intField + GetValue1("")) & intRow1).Value = strTmp
                            Wks.Range(Chr(intField + GetValue1("")) & intRow1).HorizontalAlignment = xlLeft
                        ElseIf i = GetValue2("目標", True) Then
                            intLCol = GetValue1("目標")
                            strFormat = "#,##0"
                            bolSetData = True
                        ElseIf i = GetValue2("實績點數", True) Then
                            intLCol = GetValue1("實績")
                            bolSetData = True
                        ElseIf i = GetValue2("報出點數", True) Then
                            intLCol = GetValue1("報出")
                            bolSetData = True
                        End If
                        If bolSetData = True Then
                            Wks.Range(Chr(intField + intLCol) & intRow1).Value = "=" & strRCol & intRow2 '顯示「右表」欄
                            Wks.Range(Chr(intField + intLCol) & intRow1).NumberFormatLocal = strFormat
                            Wks.Range(Chr(intField + intLCol) & intRow1).HorizontalAlignment = xlRight
                        End If
                        If i = GetValue2("報出點數", True) Then intRow1 = intRow1 + 1
                    End If
                End If
            End If 'strField2(i) <> "結餘+轉撥"
            '*** End 左表 ***
        Next i
        intRow2 = intRow2 + 1
        
        strOldN = "" & adoquery.Fields("C01")
        adoquery.MoveNext
    Loop
    '最後一個合計
    If strOldN <> MsgText(601) Then
        Call PrintExcel4Title(xlsSalesPoint, Wks, True, strOldN)
    End If
    
    '印「右表」框線
    Call PrintExcel4RLine(xlsSalesPoint, Wks, intRTitle)
   
    'Add by Amy 2021/05/18 安全基金 (490101) 有值才出現
    'Modify by Amy 2021/08/24 +FormN
    'Modify by Amy 2022/01/25 +And A0401=" & Val(Text3) & " And A0402=" & Val(Text4) & " "
    strSql = "Select '安全基金' as C01,0 as A0409,0 as C02,'',Nvl(a0408,0) as C04,'1001' as RowNo From AccRpt44r0,Acc040 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002=a0405(+)  And R002='490101' And R001<>'ZZZ' And InStr(R002,'T')=0 And A0401=" & Val(Text3) & " And A0402=" & Val(Text4) & " " & _
      "Union Select '合　　計' as C01,0 as A0409,0 as C02,'',0 as C04,'ZZZZ' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='490101' And R001<>'ZZZ' And InStr(R002,'T')=0 Order by RowNo "
    If adoquery.State <> adStateClosed Then adoquery.Close
    adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
    If adoquery.RecordCount > 0 Then
        intRow2 = intRow2 + 1
        adoquery.MoveFirst
        Do While adoquery.EOF = False
            For i = LBound(strField2) To UBound(strField2)
                strFormat = "#,##0.00"
                strTmp = "" & adoquery.Fields(i)
                If "" & adoquery.Fields("C01") = "合　　計" And (i = GetValue2("目標", True) Or i = GetValue2("實績點數", True) Or i = GetValue2("報出點數", True)) Then
                    strTmp = Chr(i + intField2) & strTOTR & "," & Chr(i + intField2) & intRow2 - 1
                    strTmp = "=Sum(" & strTmp & ")"
                ElseIf i = GetValue2("目標", True) Then
                    strFormat = "#,##0"
                ElseIf i = GetValue2("報出點數", True) Then
                    strTmp = "=" & strTmp & "/1000"
                ElseIf i = GetValue2("實績達成率", True) Then
                    strFormat = "0.00%"
                    strTmp = Chr(GetValue2("實績點數", True) + intField2) & intRow2 & "/" & Chr(GetValue2("目標", True) + intField2) & intRow2
                    strTmp = "=IF(" & Chr(GetValue2("目標", True) + intField2) & intRow2 & "=0,0," & strTmp & ")"
                ElseIf i = GetValue2("報出達成率", True) Then
                    strFormat = "0.00%"
                    strTmp = Chr(GetValue2("報出點數", True) + intField2) & intRow2 & "/" & Chr(GetValue2("目標", True) + intField2) & intRow2
                    strTmp = "=IF(" & Chr(GetValue2("目標", True) + intField2) & intRow2 & "=0,0," & strTmp & ")"
                End If
                Wks.Range(Chr(intField2 + i) & intRow2).Value = strTmp
                '對齊
                If i = GetValue2("單位", True) Then
                    Wks.Range(Chr(intField2 + i) & intRow2).HorizontalAlignment = xlLeft
                Else
                    Wks.Range(Chr(intField2 + i) & intRow2).HorizontalAlignment = xlRight
                End If
                '格式
                If strFormat <> MsgText(601) Then
                    Wks.Range(Chr(intField2 + i) & intRow2).NumberFormatLocal = strFormat
                End If
            Next i
            intRow2 = intRow2 + 1
            adoquery.MoveNext
        Loop
    End If
    '印「右表」框線
    Call PrintExcel4RLine(xlsSalesPoint, Wks, intRow2 - 2)
    'end 2021/05/18
    
    '設定
    Wks.PageSetup.PaperSize = 9 'A4
    Wks.PageSetup.Orientation = xlLandscape '橫印
    Wks.PageSetup.LeftMargin = xlsSalesPoint.InchesToPoints(0.5)
    Wks.PageSetup.RightMargin = xlsSalesPoint.InchesToPoints(0.5)
    Wks.PageSetup.TopMargin = xlsSalesPoint.InchesToPoints(0.2)
    Wks.PageSetup.BottomMargin = xlsSalesPoint.InchesToPoints(0.5)
    Wks.PageSetup.HeaderMargin = xlsSalesPoint.InchesToPoints(0.3)
    Wks.PageSetup.FooterMargin = xlsSalesPoint.InchesToPoints(0.3)
    
    '判斷若版本2007以上改變存格式
    If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Text3 & "年度" & Text4 & "月專業達成點數表" & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
    Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Text3 & "年度" & Text4 & "月專業達成點數表" & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
    End If
    xlsSalesPoint.Workbooks.Close
    xlsSalesPoint.Quit
    Set xlsSalesPoint = Nothing
    StatusClear
    Exit Sub

onErrHand:
    If Err.Number <> 0 And bolXlsOpen = True Then
        Resume Next
        MsgBox Err.Description, , MsgText(5)
        If Val(xlsSalesPoint.Version) < 12 Then
            xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Text3 & "年度" & Text4 & "月專業達成點數表" & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
        Else
            xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & Text3 & "年度" & Text4 & "月專業達成點數表" & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
        End If
        xlsSalesPoint.Workbooks.Close
        xlsSalesPoint.Quit
        Set xlsSalesPoint = Nothing
    End If
End Sub

'11001月以前格式語法(以「左表」對右表)
Private Function GetSql1() As String
    Dim StrSQLa As String
    
    '商標
    'Modify by Amy 2019/08/14 C03加to_char(以字串顯示,避免型態不符),增加C04,因加ACS導致無法一張A4顯示,故「著作權」 與「條碼」並列顯示-左表
    'modify by sonia 2016/1/26 4121改412101,4131改413101
    'Modify by Amy 2021/01/06 '商爭' 的'外－內' 417202增加 And R001='3' 條件/'商申' 的'FCT' 原抓417201改4172開頭及417202且為FCT資料(R001=5),'商申'字樣改'商標'
    'Modify by Amy 2021/08/24 +FormN
    StrSQLa = "Select '商申' as C01,'內－內' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'1101' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='410101' And R001<>'ZZZ' " & _
            "Union Select '商申' as C01,'(大陸)' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'1102' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And (R002='410103' Or R002='410109') And R001<>'ZZZ' " & _
            "Union Select '商申' as C01,'馬德里' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'1103' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='410107' And R001<>'ZZZ' " & _
            "Union Select '商申' as C01,'監視系統' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'1104' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='410106' And R001<>'ZZZ' " & _
            "Union Select '商爭' as C01,'內－內' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'1111' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='410104' And R001<>'ZZZ' " & _
            "Union Select '商爭' as C01,'顧問' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'1112' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='410102' And R001<>'ZZZ' " & _
            "Union Select '商爭' as C01,'外－內' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'1121' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='417202' And R001='3' And R001<>'ZZZ' " & _
            "Union Select '內商合計' as C01,'' as C02,'0' as C03,'' as C04,'11T' as RowNo From Dual " & _
            "Union Select '商標' as C01,'FCT' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'1201' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And ((SubStr(R002,1,4)='4172' And R002<>'417202' ) Or (R002='417202' And R001='5' )) And R001<>'ZZZ' " & _
            "Union Select '商標' as C01,'CFT' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'1202' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='412101' And R001<>'ZZZ' " & _
            "Union Select '商標總計' as C01,'' as C02,'0' as C03,'' as C04,'1T' as RowNo From Dual "
    '專利
    'modify by sonia 2016/8/1 FCP加科目417104,417105,417109
    StrSQLa = StrSQLa & "Union Select '專利' as C01,'內－內' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'2101' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And (R002='411101' Or R002='411104' Or R002='411105' Or R002='411106') And R001<>'ZZZ' " & _
            "Union Select '專利' as C01,'(大陸)' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'2102' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='411103' And R001<>'ZZZ' " & _
            "Union Select '專利' as C01,'顧問' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'2103' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='411102' And R001<>'ZZZ' " & _
            "Union Select '內專合計' as C01,'' as C02,'0' as C03,'' as C04,'21T' as RowNo From Dual " & _
            "Union Select '專利' as C01,'FCP' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'2201' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And (R002='417101' Or R002='417102' Or R002='417104' Or R002='417105' Or R002='417109') And R001<>'ZZZ' " & _
            "Union Select '專利' as C01,'CFP' as C02,to_Char(Nvl(Sum(R004),0)) as C03,'' as C04,'2202' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='413101' " & _
            "Union Select '專利總計' as C01,'' as C02,'0' as C03,'' as C04,'2T' as RowNo From Dual "
    '法務
    'Modify by Amy 2017/06/28 加4131CFP收入插於內專法務後,C01與C02對調 原:'CFP' as C01,'法務' as C02
    StrSQLa = StrSQLa & "Union Select '法務' as C01,'內專-法務' as C02,'0' as C03,'' as C04,'3101' as RowNo From Dual " & _
            "Union Select '法務' as C01,'CFP-法務' as C02,'0' as C03,'' as C04,'3102' as RowNo From Dual " & _
            "Union Select '法務' as C01,'內商-法務' as C02,'0' as C03,'' as C04,'3103' as RowNo From Dual " & _
            "Union Select '法務' as C01,'FCP-法務' as C02,'0' as C03,'' as C04,'3104' as RowNo From Dual " & _
            "Union Select '法務' as C01,'FCT-法務' as C02,'0' as C03,'' as C04,'3105' as RowNo From Dual " & _
            "Union Select '法務' as C01,'CFT-法務' as C02,'0' as C03,'' as C04,'3106' as RowNo From Dual "
    'Moidfy by Amy 2020/06/04 +一般法務
    'Modify by Amy 2021/01/29 bug 原:Val(Text3) >= 109 And Val(Text4) >= 4
    If Val(Text3 & Text4) >= 10904 Then
        StrSQLa = StrSQLa & "Union Select '法務' as C01,'一般法務' as C02,'0' as C03,'' as C04,'3107' as RowNo From Dual "
    End If
    StrSQLa = StrSQLa & "Union Select '法務總計' as C01,'' as C02,'0' as C03,'' as C04,'3T' as RowNo From Dual "
    'end 2020/06/04
    'end 2019/08/14
    
    '其他
    'Modify by Amy 2019/08/14 加增C04及創新業務組用收入 420101
'        StrSQLa = StrSQLa & "Union Select '著作權' as C01,'' as C02,Sum(Nvl(R004,0)) as C03,'4101' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And R002='415101' And R001<>'ZZZ' " & _
'                "Union Select '著　爭' as C01,'' as C02,Sum(Nvl(R004,0)) as C03,'4102' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And R002='415102' And R001<>'ZZZ' " & _
'                "Union Select '條　碼' as C01,'' as C02,Sum(Nvl(R004,0)) as C03,'4103' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And R002='410105' And R001<>'ZZZ' " & _
'                "Union Select '網　址' as C01,'' as C02,Sum(Nvl(R004,0)) as C03,'4104' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And R002='410108' And R001<>'ZZZ' " & _
'                "Union Select '其　他' as C01,'' as C02,Sum(Nvl(R004,0)) as C03,'4105' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And R002='7121' And R001<>'ZZZ' "
    StrSQLa = StrSQLa & "Union Select C01,to_char(Nvl(C02,0)) as C02,C03,to_char(Nvl(C04,0)) as C04,RowNo From " & _
                        "(Select '著作權' as C01,Sum(Nvl(R004,0)) as C02,'41' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='415101' And R001<>'ZZZ' " & _
            "Union Select '著　爭' as C01,Sum(Nvl(R004,0)) as C02,'42' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='415102' And R001<>'ZZZ' " & _
            "Union Select 'ＡＣＳ' as C01,Sum(Nvl(R004,0)) as C02,'43' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='420101' And R001<>'ZZZ' )," & _
                        "(Select '條　碼' as C03,Sum(Nvl(R004,0)) as C04,'41' as RowNo1 From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='410105' And R001<>'ZZZ' " & _
            "Union Select '網　址' as C03,Sum(Nvl(R004,0)) as C04,'42' as RowNo1 From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='410108' And R001<>'ZZZ' " & _
            "Union Select '其　他' as C03,Sum(Nvl(R004,0)) as C04,'43' as RowNo1 From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='7121' And R001<>'ZZZ' " & _
            ") Where RowNo=RowNo1(+) "
    StrSQLa = "Select * From (" & StrSQLa & ") Order by RowNo"
    GetSql1 = StrSQLa
End Function

'Add by Amy 2021/02/26
'11001月以後格式語法(以「右表」對左表)
Private Function GetSQL2() As String
    Dim strF As String, strF2 As String, StrSQLa As String, strAcc040 As String, strGroup As String, strWhere040, strQ As String
    Dim strPerFormance As String 'F41xx目標
    Dim strF3 As String 'Add by Amy 2021/11/09
    
    strF = ",Sum(Nvl(a0409,0)/1000) as A0409,Sum(Nvl(R004,0)-Nvl(R008,0)) as C02,Sum(Nvl(R004,0)+Nvl(R009,0)) as C04"
    strF2 = ",Sum(PE06) as A0409,Sum(Nvl(R004,0)-Nvl(R008,0)) as C02,Sum(Nvl(R004,0)+Nvl(R009,0)) as C04"
    'Add by Amy 2021/11/09 可勾選結餘+轉撥欄(R008)
    If ChkChoose.Value = vbChecked Then
         strF = strF & ",Sum(R008) as C06"
         strF2 = strF2 & ",Sum(R008) as C06"
         strF3 = ",Sum(R008) as C06" 'for 法務
    End If
     
    strGroup = " Group by a0405 "
    strWhere040 = "And A0401=" & Val(Text3) & " And A0402=" & Val(Text4) & " "
    strAcc040 = "Select a0405,Sum(Nvl(a0409,0)) A0409 From Acc040 Where A0404='TOT' " & strWhere040
    strPerFormance = "Select PE01,PE06 From PerFormance Where PE02='TOT' And PE03=" & Val(Text3) + 1911 & Text4
    
    'Memo by Amy 此處有改需確認 Frmacc42a0 專業達成點數分佈情況(當月實際達成)-工作表3是否也要改,且需考慮改後,要下 之前年月 的格式看看
    'Modify by Amy 2021/08/24 +FormN
    '專利
    'Modify by Amy 2022/03/28 原R003||' - FCP' as C01:專利國外部-FCP/專利日本部-FCP 將FCP拿掉-黃美珍
    StrSQLa = "Select '專利國內部 - P' as C01" & strF & ",'1101' as RowNo From AccRpt44r0,(" & strAcc040 & " And a0405='4111' " & strGroup & " ) Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002=a0405(+)  And SubStr(R002,1,4)='4111' And R001<>'ZZZ' And InStr(R002,'T')=0 " & _
        "Union Select '專利國內部 - CFP' as C01" & strF & ",'1102' as RowNo From AccRpt44r0,(" & strAcc040 & " And a0405='4131' " & strGroup & " ) Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And SubStr(R002,1,4)=a0405(+)  And R002='413101' And R001<>'ZZZ' And InStr(R002,'T')=0 " & _
        "Union Select R003 as C01" & strF2 & ",'1103' as RowNo From AccRpt44r0,(" & strPerFormance & " And PE01='F4104' ) Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002=PE01(+)  And R002='F4104' Group by R003 " & _
        "Union Select R003 as C01" & strF2 & ",'1104' as RowNo From AccRpt44r0,(" & strPerFormance & " And PE01='F4105' ) Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002=PE01(+)  And R002='F4105' Group by R003 "
            
    '商標
    'Memo by Amy 410105/410108 舊格式會單獨列,新格式已不使用
    StrSQLa = StrSQLa & _
            "Union Select '商標部 - T' as C01" & strF & ",'2101' as RowNo From AccRpt44r0,(" & strAcc040 & " And SubStr(a0405,1,4)='4101' " & strGroup & " ) Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002=a0405(+)  And (SubStr(R002,1,4)='4101' Or (R002='417202' And R001='3')) And R001<>'ZZZ' And InStr(R002,'T')=0 " & _
            "Union Select '商標部 - CFT' as C01" & strF & ",'2102' as RowNo From AccRpt44r0,(" & strAcc040 & " And SubStr(a0405,1,4)='4121' " & strGroup & " ) Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002=a0405(+)  And SubStr(R002,1,4)='4121' And R001<>'ZZZ' And InStr(R002,'T')=0 " & _
            "Union Select '商標部 -'||R003 as C01" & strF2 & ",'2103' as RowNo From AccRpt44r0,(" & strPerFormance & " And PE01='F4106' ) Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002=PE01(+)  And R002='F4106' Group by R003 " & _
            "Union Select '商標部 -'||R003 as C01" & strF2 & ",'2104' as RowNo From AccRpt44r0,(" & strPerFormance & " And PE01='F4107' ) Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002=PE01(+)  And R002='F4107' Group by R003 "
    
    '著作權/創新業務部 - ACS
    '避免餘額檔子項目沒資料,故先抓目標再串
    'Modify by Amy 2021/11/09 +strF3/C06
    'Modify by Amy 2022/03/28 原:'創新業務部 - ACS' 拿掉ACS文字-黃美珍
    'Modify by Amy 2023/12/07 原:創新業務部-黃美珍
    strQ = "Select Decode(a0405,'4151','商標部 - 著作權','顧問服務組') as C01,a0409/1000 as a0409,Nvl(C02,0) as C02,Nvl(C04,0) as C04" & Replace(strF3, "Sum(R008)", "Nvl(C06,0)") & ",Decode(a0405,'4151','3101','3102') as RowNo From " & _
               "(" & strAcc040 & " And a0405 In ('4151','4201') " & strGroup & " )," & _
               "(Select SubStr(R002,1,4) as R002" & Replace(strF, ",Sum(Nvl(a0409,0)/1000) as A0409", "") & " From Accrpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002 in('415101','420101') And R001<>'ZZZ' And InStr(R002,'T')=0 Group by SubStr(R002,1,4) ) " & _
               "Where a0405=R002(+) "
    StrSQLa = StrSQLa & "Union " & strQ
    
    'Add by Amy 2022/01/25 +if 111 年以前才有資料
    If Val(Text3) < 111 Then
    '法務
    '避免餘額檔子項目沒資料,故先抓目標再串,11001月一般法務沒資料會顯示不出來,故使用Dual抓
    strQ = "Select Decode(a0405,'411107','法務 - P','413102','法務 - CFP','417103','法務 - FCP','410110','法務 - T','412102','法務 - CFT','417203','法務 - FCT','一般法務') as C01,a0409/1000 as a0409,Nvl(C02,0) as C02,Nvl(C04,0) as C04" & Replace(strF3, "Sum(R008)", "Nvl(C06,0)") & ",Decode(a0405,'411107','4101','413102','4102','417103','4103','410110','4104','412102','4105','417203','4106','4107') as RowNo From " & _
              "(" & strAcc040 & " And a0405 In ('411107','413102','417103','410110','412102','417203') " & strGroup & _
                    " Union Select AccNo,A0409 From (Select '41XX' as AccNo From Dual)" & _
                    ",(" & Replace(strAcc040, "a0405", "'41XX' as a0405") & " And SubStr(a0405,1,4) In ('4141','4161','4181') And A0403<>'L' ) " & _
                    "Where AccNo=a0405(+) " & _
                ")" & _
             ",(Select R002" & Replace(strF, ",Sum(Nvl(a0409,0)/1000) as A0409", "") & " From Accrpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002 In ('411107','413102','417103','410110','412102','417203') And R001='ZZZ' And InStr(R002,'T')=0 Group by R002 " & _
      "Union Select '41XX'as R002,Sum(Nvl(R008,0)) as C02,Sum(Nvl(R006,0)+Nvl(R009,0)) as C04" & strF3 & " From Accrpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And SubStr(R002,1,4) In ('4141','4161','4181') And R001='ZZZ' And InStr(R002,'T')=0 " & _
               ") Where a0405=R002(+) "
    StrSQLa = StrSQLa & "Union " & strQ
    End If
    
    'Add by Amy 2022/01/25 其他各項收入(419102)
    '11101月起 1與J公司法務科目及7121改為 490102,有目標(設於4901)才顯示
    strQ = "Select C01,A0409,Nvl(C02,0) as C02,Nvl(C04,0) as C04" & Replace(strF3, "Sum(R008)", "Nvl(C06,0)") & ",'4108' as RowNo From " & _
                    "( Select a0102 as C01,Sum(Nvl(a0409,0)/1000) as A0409,a0405 From Acc040,Acc010 Where  a0405='4901' And A0403<>'L' And a0405=a0101(+) " & strWhere040 & strGroup & ",a0102) " & _
                    ",( Select SubStr(R002,1,4) R002" & Replace(strF, ",Sum(Nvl(a0409,0)/1000) as A0409", "") & " From Accrpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002 ='490102'  Group by R002) " & _
               "Where a0405=R002(+) "
    StrSQLa = StrSQLa & "Union " & strQ

    '其他/全所
    'Memo by Amy 410105/410108 舊格式會單獨列,目前列於4101,4151目前單獨列 415102 科目110後未使用
    StrSQLa = StrSQLa & _
            "Union Select '其　他' as C01" & Replace(strF, ",Sum(Nvl(a0409,0)/1000) as A0409", ",0 as A0409") & ",'5101' as RowNo From AccRpt44r0 Where ID='" & strUserNum & "' And FormN='" & Me.Name & "' And R002='7121' And R001<>'ZZZ' And InStr(R002,'T')=0 " & _
            "Union Select '全　所' as C01,0 as A0409,0 as C02,0 as C04" & Replace(strF3, "Sum(R008)", "0") & ",'ZZZZ' as RowNo From Dual "
    
    GetSQL2 = "Select C01,Nvl(A0409,0) as A0409,Nvl(C02,0) as C02,'' as C03,Nvl(C04,0) as C04,'' as C05,RowNo" & Replace(strF3, "Sum(R008)", "Nvl(C06,0)") & " From (" & StrSQLa & ") Order by RowNo"
    'end 2021/11/09
End Function
    
Private Sub PrintExcel4Title(xlsApp As Excel.Application, Wks As Worksheet, Optional bolLTableTitleOnly As Boolean = False, Optional stN As String = "")
    Dim stTP As String, stTp2 As String, ii As Integer, j As Integer
    
    With Wks
        '只印左表欄名
        If bolLTableTitleOnly = False Then
            '表頭
            'Modify by Amy 2021/05/13 年 月數字後加空白,多空一行-美珍
            Wks.Range(Chr(intField) & intRow1).Value = Text3 & " 年度  " & Text4 & " 月份 - 專業達成點數"
            Wks.Range(Chr(intField) & intRow1 & ":" & Chr(UBound(strField1) + UBound(strField2) + intField + 2) & intRow1).HorizontalAlignment = xlCenter
            Wks.Range(Chr(intField) & intRow1 & ":" & Chr(UBound(strField1) + UBound(strField2) + intField + 2) & intRow1).MergeCells = True
            intRow1 = intRow1 + 2
        End If
        
        '左表-欄位名
        For ii = LBound(strField1) To UBound(strField1)
            stTP = strField1(ii)
            If bolLTableTitleOnly = False Then
                Wks.Columns(Chr(intField + ii) & ":" & Chr(intField + ii)).ColumnWidth = intWidth1(ii)
                If ii = UBound(strField1) Then
                    Wks.Columns(Chr(intField + ii + 1) & ":" & Chr(intField + ii + 1)).ColumnWidth = 1
                End If
            '只印欄位名
            Else
                If ii = LBound(strField1) Then
                    stTP = Left(stN, 2)
                    '***上一個Table資料 ***
                    For j = GetValue1("") To UBound(strField1)
                        '總計
                        If j = GetValue1("") Then
                            Wks.Range(Chr(j + intField) & intRow1).Value = "總計"
                            Wks.Range(Chr(intField) & intRow1 & ":" & Chr(j + intField) & intRow1).HorizontalAlignment = xlCenter
                            Wks.Range(Chr(intField) & intRow1 & ":" & Chr(j + intField) & intRow1).MergeCells = True
                        Else
                            Wks.Range(Chr(j + intField) & intRow1).Value = "=Sum(" & Chr(j + intField) & intLStart & ":" & Chr(j + intField) & intRow1 - 1 & ")"
                            If j = GetValue1("目標") Then
                                Wks.Range(Chr(j + intField) & intRow1).NumberFormatLocal = "#,##0"
                            Else
                                Wks.Range(Chr(j + intField) & intRow1).NumberFormatLocal = "#,##0.00"
                            End If
                            Wks.Range(Chr(j + intField) & intRow1).HorizontalAlignment = xlRight
                        End If
                        '總達成率
                        If j = GetValue1("") Then
                            Wks.Range(Chr(j + intField) & intRow1 + 1).Value = "總達成率"
                            Wks.Range(Chr(intField) & intRow1 + 1 & ":" & Chr(j + intField) & intRow1 + 1).HorizontalAlignment = xlCenter
                            Wks.Range(Chr(intField) & intRow1 + 1 & ":" & Chr(j + intField) & intRow1 + 1).MergeCells = True
                        Else
                            If j = GetValue1("實績") Then
                                stTp2 = Chr(j + intField) & intRow1 & "/" & Chr(GetValue1("目標") + intField) & intRow1
                                stTp2 = "=IF(" & Chr(GetValue1("目標") + intField) & intRow1 & "=0,0," & stTp2 & ")"
                            ElseIf j = GetValue1("報出") Then
                                stTp2 = Chr(j + intField) & intRow1 & "/" & Chr(GetValue1("目標") + intField) & intRow1
                                stTp2 = "=IF(" & Chr(GetValue1("目標") + intField) & intRow1 & "=0,0," & stTp2 & ")"
                            End If
                            If j > GetValue1("目標") Then
                                Wks.Range(Chr(j + intField) & intRow1 + 1).Value = stTp2
                                Wks.Range(Chr(j + intField) & intRow1 + 1).NumberFormatLocal = "0.00%"
                                Wks.Range(Chr(j + intField) & intRow1 + 1).HorizontalAlignment = xlRight
                            End If
                            '設定框線
                            If j = GetValue1("報出") Then
                                Wks.Range(Chr(intField) & intLStart - 1 & ":" & Chr(UBound(strField1) + intField) & intRow1 + 1).Select
                                xlsApp.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
                                xlsApp.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
                                xlsApp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
                                xlsApp.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
                                xlsApp.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
                                xlsApp.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                                intRow1 = intRow1 + 3
                            End If
                        End If
                    Next j
                    '*** end 上一個Table資料 ***
                End If
            End If
            '「左表」不需印
            If stN <> "其　他" And stN <> "全　所" Then
                Wks.Range(Chr(intField + ii) & intRow1).Value = stTP
                Wks.Range(Chr(intField + ii) & intRow1).HorizontalAlignment = xlCenter
            End If
        Next
        If bolLTableTitleOnly = True Then intRow1 = intRow1 + 1
        
        If bolLTableTitleOnly = False Then
            '右表-欄位名
            For ii = LBound(strField2) To UBound(strField2)
                Wks.Range(Chr(intField2 + ii) & intRow1).Value = strField2(ii)
                Wks.Range(Chr(intField2 + ii) & intRow1).HorizontalAlignment = xlCenter
                Wks.Columns(Chr(intField2 + ii) & ":" & Chr(intField2 + ii)).ColumnWidth = intWidth2(ii)
            Next
            intRow1 = intRow1 + 1
            intRow2 = intRow1
        End If
        
    End With
End Sub

Private Sub PrintExcel4RLine(xlsApp As Excel.Application, Wks As Worksheet, intRTitle As Integer)
    Wks.Range(Chr(intField2) & intRTitle & ":" & Chr(UBound(strField2) + intField2) & intRow2 - 1).Select
    xlsApp.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlsApp.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    xlsApp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlsApp.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    xlsApp.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    xlsApp.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
End Sub
'end 2021/02/26


