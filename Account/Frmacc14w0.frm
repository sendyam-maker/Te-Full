VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14w0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "國內人員收文點數及收款點數統計"
   ClientHeight    =   2100
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   3990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3990
   Begin VB.CommandButton Cmd_Excel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   730
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   1600
      Width           =   2230
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   3
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   4
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   5
      Left            =   2880
      TabIndex        =   5
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   7
      Left            =   2880
      TabIndex        =   7
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "統計年月4："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   195
      TabIndex        =   16
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2535
      TabIndex        =   15
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "統計年月2："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   195
      TabIndex        =   14
      Top             =   480
      Width           =   1305
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2535
      TabIndex        =   13
      Top             =   435
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2535
      TabIndex        =   12
      Top             =   810
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "統計年月3："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   195
      TabIndex        =   11
      Top             =   840
      Width           =   1305
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2535
      TabIndex        =   10
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "統計年月1："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   195
      TabIndex        =   9
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "Frmacc14w0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add by Amy 2017/03/16
Option Explicit

Public adoQ As New ADODB.Recordset
Dim i As Integer
Dim strF(), intWidth()
Dim intTitleR As Integer, intField As Integer, intCounter As Integer

Private Sub Cmd_Excel_Click()
    If TxtValidate = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    PrintData
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim intX As Integer, intY As Integer
    Dim sglWidth As Single, sglHeight As Single
    Dim MskBx As MaskEdBox
       
    Me.Icon = LoadPicture(strIcoPath)
    strFormName = Name
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    Image1 = LoadPicture(strBackPicPath4)
    sglWidth = Image1.Width
    sglHeight = Image1.Height
    For intX = 0 To Int(ScaleWidth / sglWidth)
        For intY = 0 To Int(ScaleHeight / sglHeight)
            PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
        Next
    Next
    
    For Each MskBx In MaskEdBox1
        MskBx.Mask = Mid(DFormat, 1, 6)
    Next
    
    Frmacc0000.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   Set Frmacc14w0 = Nothing
End Sub

Private Sub MaskEdBox1_LostFocus(Index As Integer)
    If MaskEdBox1(Index) = Mid(MsgText(29), 1, 6) Then Exit Sub
    
    If MaskEdBox1(Index) <> Mid(MsgText(29), 1, 6) Then
        If InStr(MaskEdBox1(Index), "_") > 0 Then
            MsgBox "請輸入完整年月!", , MsgText(5)
            MaskEdBox1(Index).SetFocus
        Else
            strExc(1) = Replace(MaskEdBox1(Index), "/", "") & "01"
            If CheckIsTaiwanDate(strExc(1)) = False Then
               MaskEdBox1(Index).SetFocus
            End If
        End If
    End If
End Sub

Private Function TxtValidate() As Boolean
    Dim MskBx As MaskEdBox
    
    TxtValidate = False
    For Each MskBx In MaskEdBox1
        If Val(MskBx.Index) Mod 2 = 0 Then
            If MaskEdBox1(MskBx.Index) <> Mid(MsgText(29), 1, 6) And MaskEdBox1(MskBx.Index + 1) = Mid(MsgText(29), 1, 6) Then
                MaskEdBox1(MskBx.Index + 1) = MaskEdBox1(MskBx.Index)
            End If
        End If
        If MskBx = Mid(MsgText(29), 1, 6) Then
            MsgBox "統計年月未輸入完整請確認！"
            Exit Function
        End If
    Next
    TxtValidate = True
End Function

Private Sub PrintData()
    Dim strQ As String, strWhere As String
    Dim strWhere1(3) As String
    Dim MskBx As MaskEdBox
    Dim idx As Integer
    
    For i = 0 To 6 Step 2
        idx = i / 2
        '收文點數
        strWhere1(idx) = " And R011>=" & Val(Replace(MaskEdBox1(i), "/", "")) + 191100 & "01 And R011<=" & Val(Replace(MaskEdBox1(i + 1), "/", "")) + 191100 & "31 "
    Next i
    
    strQ = "Delete Accrpt14W0 Where ID='" & strUserNum & "' "
    cnnConnection.Execute strQ
    
    '收文點數
    strQ = ""
    For i = 0 To UBound(strWhere1)
        strWhere = strWhere & "(" & Replace(Mid(strWhere1(i), 6), "R011", "CP05") & ") Or "
    Next i
    strWhere = " And (" & Left(strWhere, Len(strWhere) - 3) & ")"
    strQ = "Select '" & strUserNum & "',Decode(CP13,'A4023','S14',CP12),CP13,CP01,CP02,CP03,CP04,CP09,CP10,NVL(CP16,0)-NVL(CP17,0)-NVL(A1U07,0) CP18,CP05,'1' " & _
                "From CASEPROGRESS, " & _
                    "(Select A1U03,Sum(A1U07) A1U07 " & _
                    "From CASEPROGRESS,ACC1U0 " & _
                    "Where  NVL(CP18,0)<>0 And (CP159=0 OR (CP27>0 And CP57>0)) And SubStr(CP12,1,1)<>'F' " & strWhere & " " & _
                    "And CP09=A1U03(+) Group By A1U03 " & _
                    ") " & _
                "Where NVL(CP18,0)<>0 And (CP159=0 OR (CP27>0 And CP57>0)) And SubStr(CP12,1,1)<>'F' " & strWhere & " " & _
                "And CP09=A1U03(+)"
    strQ = "Insert Into Accrpt14W0 (ID,R002,R003,R004,R005,R006,R007,R008,R009,R010,R011,R012) " & strQ
    cnnConnection.Execute strQ
    
    '收款點數
    strQ = ""
    strWhere = Replace(strWhere, "CP05", "A0205+19110000")
    'Modify by Amy 2019/09/18 新增創新業務組案件(系統類別ACS) 收入420101及規費科目220114 原:(SubStr(AX205, 1, 2) = '41'
    'modify by sonia 2021/3/11 原剔除國外部以SubStr(ST15,1,1)<>'F'剔除,改為SubStr(ax209,1,3)<>'F41'(J公司D110020042分配點數給外商陳蒲璇要出現)
    strQ = "Select '" & strUserNum & "',Decode(AX209,'A4023','S14',ST15),AX209,AX201,AX202,Sum(AX207),A0205+19110000,'2' " & _
            "From Staff," & _
                "(Select AX209,AX201,AX202,AX203,AX207-AX206 AX207,A0205 " & _
                "From ACC020,ACC021 " & _
                "Where A0201=AX201(+) And A0202=AX202(+) " & strWhere & " " & _
                "And (SubStr(AX205, 1, 1) = '4' OR (AX205='7121' And AX209 IS NOT NULL)) And NOT (AX205='4191' OR AX205='4192' OR AX205='4194') " & _
                "And (AX213 IS NULL OR Instr(AX213||' ','結餘')=0) And Instr(AX212,'轉撥')=0 " & _
                ") " & _
            "Where AX209=ST01(+) and SubStr(ax209,1,3)<>'F41' " & _
            "Group By '" & strUserNum & "',Decode(AX209,'A4023','S14',ST15),AX209,AX201,AX202,A0205 "

    strQ = "Insert Into Accrpt14W0 (ID,R002,R003,R001,R008,R010,R011,R012) " & strQ
    cnnConnection.Execute strQ
    
    
    '中所部份人員於105年初有調區情形,收款點數需人工調整
    'Modify by Amy 2021/03/15 W部門資料搬至智權部合計之後,法務資料另外列示
    strQ = ""
    strWhere = " And SubStr(R002,1,1)<>'L' And R002<>'F31' And R002<>'P31' "
    For i = 1 To 2
        If i = 2 Then
            strWhere = " And (SubStr(R002,1,1)='L' Or R002='F31' Or R002='P31') "
        End If
        strQ = strQ & "Select A0902 部門,ST02 姓名,Sum(Nvl(當月收文,0)/1000) 當月收文,Sum(Nvl(當年累計收文,0)/1000) 當年累計收文,Sum(Nvl(去年當月收文,0)/1000) 去年當月收文,Sum(Nvl(去年累計收文,0)/1000) 去年累計收文," & _
                    "Sum(Nvl(當月收款,0)/1000) 當月收款,Sum(Nvl(當年累計收款,0)/1000) 當年累計收款,Sum(Nvl(去年當月收款,0)/1000) 去年當月收款,Sum(Nvl(去年累計收款,0)/1000) 去年累計收款," & _
                    "P.R002 ST15,P.R003 智權 ," & i & " as Sort " & _
                    "From Acc090,Staff" & _
                      ",(Select Distinct R002,R003 From Accrpt14W0 Where ID='" & strUserNum & "' " & strWhere & ") P " & _
                      ",(Select R002,R003,Sum(Nvl(R010,0)) AS 當月收文 From Accrpt14W0 Where ID='" & strUserNum & "' And R012='1' " & strWhere1(0) & strWhere & " Group By R002,R003 ) V1 " & _
                      ",(Select R002,R003,Sum(Nvl(R010,0)) AS 當年累計收文 From Accrpt14W0 Where ID='" & strUserNum & "' And R012='1' " & strWhere1(1) & strWhere & " Group By R002,R003) V2 " & _
                      ",(Select R002,R003,Sum(Nvl(R010,0)) AS 去年當月收文 From Accrpt14W0 Where ID='" & strUserNum & "' And R012='1' " & strWhere1(2) & strWhere & " Group By R002,R003) V3 " & _
                      ",(Select R002,R003,Sum(Nvl(R010,0)) AS 去年累計收文 From Accrpt14W0 Where ID='" & strUserNum & "' And R012='1' " & strWhere1(3) & strWhere & " Group By R002,R003) V4 " & _
                      ",(Select R002,R003,Sum(Nvl(R010,0)) AS 當月收款 From Accrpt14W0 Where ID='" & strUserNum & "' And R012='2' " & strWhere1(0) & strWhere & " Group By R002,R003) V5 " & _
                      ",(Select R002,R003,Sum(Nvl(R010,0)) AS 當年累計收款 From Accrpt14W0 Where ID='" & strUserNum & "' And R012='2' " & strWhere1(1) & strWhere & " Group By R002,R003) V6 " & _
                      ",(Select R002,R003,Sum(Nvl(R010,0)) AS 去年當月收款 From Accrpt14W0 Where ID='" & strUserNum & "' And R012='2' " & strWhere1(2) & strWhere & " Group By R002,R003) V7 " & _
                      ",(Select R002,R003,Sum(Nvl(R010,0)) AS 去年累計收款 From Accrpt14W0 Where ID='" & strUserNum & "' And R012='2' " & strWhere1(3) & strWhere & " Group By R002,R003) V8 " & _
                    "Where P.R002=A0901(+) And P.R003=ST01(+) And P.R002=V1.R002(+) And P.R003=V1.R003(+) And P.R002=V2.R002(+) And P.R003=V2.R003(+) " & _
                    "And P.R002=V3.R002(+) And P.R003=V3.R003(+) And P.R002=V4.R002(+) And P.R003=V4.R003(+) And P.R002=V5.R002(+) And P.R003=V5.R003(+) " & _
                    "And P.R002=V6.R002(+) And P.R003=V6.R003(+) And P.R002=V7.R002(+) And P.R003=V7.R003(+) And P.R002=V8.R002(+) And P.R003=V8.R003(+) " & _
                    "Group by P.R002,A0902,P.R003,st02 "
        If i = 1 Then
            strQ = strQ & " Union All "
        End If
    Next i
    strQ = "Select * From (" & strQ & ") Order by Sort,Decode(SubStr(st15,1,1),'S',1,Decode(SubStr(st15,1,1),'W',2,3)),st15, 智權 "
    'end 2021/03/15
    If adoQ.State = adStateOpen Then adoQ.Close
    adoQ.CursorLocation = adUseClient
    adoQ.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoQ.RecordCount = 0 Then
        MsgBox MsgText(28), , MsgText(5)
        Exit Sub
    Else
        SaveExcel
        MsgBox ("EXCEL檔案已產生！")
    End If
End Sub

Private Function SaveExcel() As Boolean
    Dim xlsAgentPoint As New Excel.Application
    Dim wksrpt As New Worksheet
    Dim xlsFileName As String, strQ As String
    Dim intStart As Integer
    Dim strOldDept As String, strSum As String, strTotal As String
    Dim strOldSort As String 'Add byAmy 2021/03/15
    
 On Error GoTo ErrHand
 
    ReDim strF(9)
    ReDim intwith(9)
    strF = Array("部門", "姓名", "收文條件1", "收文條件2", "收文條件3", "收文條件4", "收款條件1", "收款條件2", "收款條件3", "收款條件4")
    intWidth = Array(11.5, 8.5, 10.5, 10.5, 10.5, 10.5, 10.5, 10.5, 10.5, 10.5)
    
    intField = 65: intCounter = 1: SaveExcel = False
 
    xlsFileName = ServerDate & "收文點數及收款點數統計" & MsgText(43)
    If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
       End If
    Else
         Kill strExcelPath & xlsFileName
    End If
    'Modify by Amy 2021/03/15 原:1
    xlsAgentPoint.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
    xlsAgentPoint.Workbooks.add
    Set wksrpt = xlsAgentPoint.Worksheets(1)
    Call SetTitle(False, wksrpt)
    intTitleR = intCounter
    intCounter = intCounter + 1: intStart = intCounter
    xlsAgentPoint.Visible = True
    
    '逐筆填值
    With adoQ
        .MoveFirst
        Do While .EOF = False
            If strOldDept <> MsgText(601) And strOldDept <> .Fields("ST15") And Left(.Fields("ST15"), 1) = "S" Then
                If Left(strOldDept, 2) <> Left(.Fields("ST15"), 2) Then
                    If Val(Mid(strOldDept, 2, 1)) < 3 Then
                        Call SetSum(1, wksrpt, "", intStart)
                        strSum = strSum & "+" & Chr(GetValue("收文條件1") + intField) & intCounter
                        intCounter = intCounter + 1
                        '所小計
                        Call SetSum(2, wksrpt, strOldDept, strSum)
                    Else
                        strSum = strSum & "+" & Chr(GetValue("收文條件1") + intField) & intCounter - 1
                        '所小計
                        Call SetSum(2, wksrpt, strOldDept, intStart)
                    End If
                    strSum = ""
                    strTotal = strTotal & "+" & Chr(GetValue("收文條件1") + intField) & intCounter
                    intCounter = intCounter + 2
                    intStart = intCounter
                ElseIf strOldDept <> .Fields("ST15") And Val(Mid(.Fields("ST15"), 2, 1)) < 3 Then
                    'ST15小計
                    Call SetSum(1, wksrpt, "", intStart)
                    strSum = strSum & "+" & Chr(GetValue("收文條件1") + intField) & intCounter
                    intCounter = intCounter + 1
                    intStart = intCounter
                End If
            End If
            'Modify by Amy 2021/03/15 W部門列於智權部合計後,其他小計不需加總,但國內合計需計算
            If strOldDept <> MsgText(601) And strOldDept <> .Fields("ST15") Then
                If Left(strOldDept, 1) = "S" And Left(strOldDept, 1) <> Left(.Fields("ST15"), 1) Then
                    strSum = strSum & "+" & Chr(GetValue("收文條件1") + intField) & intCounter - 1
                    '所小計
                    If Val(Mid(strOldDept, 2, 1)) < 3 Then
                        Call SetSum(2, wksrpt, strOldDept, strSum)
                    Else
                        Call SetSum(2, wksrpt, strOldDept, intStart)
                    End If
                    strSum = ""
                    strTotal = strTotal & "+" & Chr(GetValue("收文條件1") + intField) & intCounter
                    intCounter = intCounter + 2
                  
                    '智權部合計
                    Call SetSum(3, wksrpt, "智權部合計", strTotal)
                    intCounter = intCounter + 2
                    intStart = intCounter
                'W部門後空一行
                ElseIf Left(strOldDept, 1) = "W" And Left(strOldDept, 1) <> Left(.Fields("ST15"), 1) Then
                    intCounter = intCounter + 2
                    intStart = intCounter
                End If
            End If
            'end 2021/03/15
            
            'Add by Amy 2021/03/15 L部門另外列,故「其他小計」/「國內合計」從下面搬過來
            If strOldSort <> MsgText(601) And strOldSort = "1" And strOldSort <> .Fields("Sort") Then
                '其他小計
                Call SetSum(2, wksrpt, "其他小計", intStart)
                strTotal = strTotal & "+" & Chr(GetValue("收文條件1") + intField) & intCounter
                intCounter = intCounter + 2
                '最後合計
                Call SetSum(4, wksrpt, "國內合計", strTotal)
                intCounter = intCounter + 2
                strTotal = intCounter
            End If
            
            For i = LBound(strF) To UBound(strF)
                wksrpt.Range(Chr(i + intField) & intCounter).Value = "" & .Fields(i)
                If i >= GetValue("部門") And i <= GetValue("姓名") Then
                    wksrpt.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter
                End If
            Next i
            'Add by Amy 2021/03/15 國內合計 +W1001/W2001 位置
            If "" & .Fields("ST15") = "W10" Then strTotal = strTotal & "+" & Chr(GetValue("收文條件1") + intField) & intCounter
            If "" & .Fields("ST15") = "W20" Then strTotal = strTotal & "+" & Chr(GetValue("收文條件1") + intField) & intCounter
            
            intCounter = intCounter + 1
            strOldDept = .Fields("ST15")
            strOldSort = .Fields("Sort")
            .MoveNext
        Loop
        'Modify by Amy 2021/03/15 原全部列一起,L公司另外列,將「其他小計」往上搬
        '最後合計
        Call SetSum(5, wksrpt, "法律所", strTotal)
        'end 2021/03/15
    End With
    '重設欄位名稱
    Call SetTitle(True, wksrpt)
    
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    SaveExcel = True
    Exit Function
    
ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set xlsAgentPoint = Nothing
    Set wksrpt = Nothing
End Function

Private Sub SetTitle(ByVal IsLast As Boolean, ByRef Wks As Worksheet)
    
    If IsLast = True Then
        Dim strTmp(3) As String
        
        Wks.Range(Chr(GetValue("收文條件1") + intField) & intTitleR + 1 & ":" & _
            Chr(GetValue("收款條件4") + intField) & intCounter).NumberFormatLocal = "0.00"
        
        '更改欄位名稱
        For i = 0 To 6 Step 2
            If Mid(MaskEdBox1(i), 1, 3) = Mid(MaskEdBox1(i + 1), 1, 3) Then
                strTmp(i / 2) = MaskEdBox1(i) & " ~ " & Right(MaskEdBox1(i + 1), 2)
            Else
                strTmp(i / 2) = MaskEdBox1(i) & " ~ " & MaskEdBox1(i + 1)
            End If
        Next i
    
        For i = LBound(strTmp) To UBound(strTmp)
            Wks.Range(Chr(i + intField + 2) & intTitleR).Value = strTmp(i)
            Wks.Range(Chr(i + intField + 2) & intTitleR).HorizontalAlignment = xlCenter
            Wks.Range(Chr(i + intField + 6) & intTitleR).Value = strTmp(i)
            Wks.Range(Chr(i + intField + 6) & intTitleR).HorizontalAlignment = xlCenter
        Next i
        
        Wks.PageSetup.PaperSize = 9 '設定紙張 A4
        Wks.PageSetup.PrintTitleRows = "$1:$" & intTitleR '標題列
        Wks.PageSetup.LeftMargin = 20 '邊界
        Wks.PageSetup.RightMargin = 27
        Wks.PageSetup.TopMargin = 50
        Wks.PageSetup.BottomMargin = 20
        Wks.PageSetup.PrintGridlines = True '列印格線

    Else
        intCounter = intCounter + 1
        For i = LBound(strF) To UBound(strF)
            If i = GetValue("收文條件1") Then
                Wks.Range(Chr(i + intField) & "1").Value = "收文點數"
            ElseIf i = GetValue("收款條件1") Then
                Wks.Range(Chr(i + intField) & "1").Value = "收款點數"
            End If
            Wks.Range(Chr(i + intField) & intCounter).Value = strF(i)
            Wks.Columns(Chr(i + intField) & ":" & Chr(i + intField)).ColumnWidth = intWidth(i)
        Next i
    End If
End Sub

Private Sub SetSum(ByVal intChoose As Integer, ByRef Wks As Worksheet, ByVal strDeptN As String, ByVal strSum As String)
    Dim j As Integer
    Dim strName As String
    
    Select Case intChoose
        Case 1, 2 '小計
            strName = "小計"
        'Modify by Amy 2021/03/15 +5
        Case 3, 4, 5 '合計
            strName = "合計"
    End Select
    For j = GetValue("部門") To UBound(strF)
        If j = GetValue("部門") Then
            If intChoose = 2 And Left(strDeptN, 1) = "S" Then
                Select Case Mid(strDeptN, 2, 1)
                    Case 0
                        strDeptN = "智權部主管"
                    Case 1
                        strDeptN = "台北所"
                    Case 2
                        strDeptN = "台中所"
                    Case 3
                        strDeptN = "台南所"
                    Case 4
                        strDeptN = "高雄所"
                End Select
            End If
            Wks.Range(Chr(j + intField) & intCounter).Value = strDeptN
            Wks.Range(Chr(j + intField) & intCounter).HorizontalAlignment = xlCenter
        ElseIf j = GetValue("姓名") Then
            Wks.Range(Chr(j + intField) & intCounter).Value = strName
            Wks.Range(Chr(j + intField) & intCounter).HorizontalAlignment = xlCenter
        Else
            If j = GetValue("收文條件1") And InStr(strSum, "+") = 0 Then
                strSum = Chr(j + intField) & strSum & ":" & Chr(j + intField) & intCounter - 1
            Else
                strSum = Replace(strSum, Chr(j + intField - 1), Chr(j + intField))
            End If
            If InStr(strSum, "+") = 0 Then
                Wks.Range(Chr(j + intField) & intCounter).Formula = "=Sum(" & strSum & ")"
            Else
                Wks.Range(Chr(j + intField) & intCounter).Formula = "=" & Mid(strSum, 2)
            End If
            
        End If
    Next j
    
End Sub

Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = LBound(strF) To UBound(strF)
       If UCase(strF(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function
