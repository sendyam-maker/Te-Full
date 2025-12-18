VERSION 5.00
Begin VB.Form frm090711_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人工作進度資料維護"
   ClientHeight    =   3990
   ClientLeft      =   630
   ClientTop       =   2220
   ClientWidth     =   6075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6075
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4968
      TabIndex        =   0
      Top             =   36
      Width           =   1092
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   12
      Top             =   3390
      Width           =   5955
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   9
      Left            =   105
      TabIndex        =   11
      Top             =   3075
      Width           =   5955
   End
   Begin VB.Label Label2 
      Caption         =   "括弧內為新制算法"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   105
      TabIndex        =   10
      Top             =   3675
      Width           =   3375
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   8
      Left            =   90
      TabIndex        =   9
      Top             =   2790
      Width           =   5955
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   0
      Left            =   84
      TabIndex        =   8
      Top             =   504
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   1
      Left            =   84
      TabIndex        =   7
      Top             =   792
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   2
      Left            =   84
      TabIndex        =   6
      Top             =   1056
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   3
      Left            =   84
      TabIndex        =   5
      Top             =   1344
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   4
      Left            =   84
      TabIndex        =   4
      Top             =   1632
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   5
      Left            =   84
      TabIndex        =   3
      Top             =   1896
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   6
      Left            =   84
      TabIndex        =   2
      Top             =   2184
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   7
      Left            =   84
      TabIndex        =   1
      Top             =   2472
      Width           =   5952
   End
End
Attribute VB_Name = "frm090711_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (無)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 22) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String
Dim PLeft(0 To 22) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, TempSeekNick As String

Private Sub cmdOK_Click()
frm090711.Show
Unload Me
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Process
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090711_1 = Nothing
End Sub

Sub Process()
'edit by nickc 2005/05/04
'lbl1(0).Caption = "可辦草圖：    0   件"
'lbl1(1).Caption = "可辦墨圖：    0   件"
'lbl1(2).Caption = "達成草圖：    0   件    0   張"
'lbl1(3).Caption = "達成墨圖：    0   件    0   張    0   點"
'lbl1(4).Caption = "其他新案：    0   件    0   點"
'lbl1(5).Caption = "其他舊案：    0   件    0   點"
'lbl1(6).Caption = "逾時草圖：    0   件"
'lbl1(7).Caption = "逾時墨圖：    0   件"
lbl1(0).Caption = "可辦草圖：    0(0)   件"
lbl1(1).Caption = "可辦墨圖：    0(0)   件"
lbl1(2).Caption = "達成草圖：    0(0)   件    0   張"
lbl1(3).Caption = "達成墨圖：    0(0)   件    0   張    0(0)   點"
lbl1(4).Caption = "其他新案：    0   件    0   點"
lbl1(5).Caption = "其他舊案：    0   件    0   點"
lbl1(6).Caption = "逾時草圖：    0   件"
lbl1(7).Caption = "逾時墨圖：    0   件"

'Add By Cheng 2003/07/01
'edit by nickc 2005/05/04
'lbl1(8).Caption = "本月發文：    0   件    0   張    0   點"
lbl1(8).Caption = "本月發文：    0(0)   件    0   張    0(0)   點"
'add by nickc 2005/05/4
lbl1(9).Caption = "提供圖檔：    0   件"
lbl1(10).Caption = "關聯案件：    0   件"
'Modify By Cheng 2003/06/30
'strSQL = "SELECT r111002,SUM(DECODE(R111003,0,0,R111003)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)) FROM R090711_1 WHERE ID='" & strUserNum & "' AND R111001='" & Trim(frm090711.Combo1.Text) & "' group by r111002 "
'edit by nickc 2005/05/04
'strSQL = "SELECT r111002,SUM(DECODE(R111003,0,0,R111003)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)) FROM R090711_1 WHERE ID='" & strUserNum & "' AND R111001='" & Trim(Left(frm090711.Combo1.Text, 6)) & "' group by r111002 "
strSql = "SELECT r111002,SUM(DECODE(R111003,0,0,R111003)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)),SUM(DECODE(R111006,0,0,R111006)),SUM(DECODE(R111007,0,0,R111007)),SUM(DECODE(R111008,0,0,R111008)) FROM R090711_1 WHERE ID='" & strUserNum & "' AND R111001='" & Trim(Left(frm090711.Combo1.Text, 6)) & "' group by r111002 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
        Select Case Val(CheckStr(.Fields(0)))
        Case 1
             'edit by nickc 2005/05/04
             'lbl1(0).Caption = "可辦草圖：  " & CheckStr(.Fields(1)) & "  件"
             lbl1(0).Caption = "可辦草圖：  " & Format(CheckStr(.Fields(1)), "###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(4)), "###,###,###,##0.00") & ")  件"
        Case 2
             'edit by nickc 2005/05/04
             'lbl1(1).Caption = "可辦墨圖：  " & CheckStr(.Fields(1)) & "  件"
             lbl1(1).Caption = "可辦墨圖：  " & Format(CheckStr(.Fields(1)), "###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(4)), "###,###,###,##0.00") & ")  件"
        Case 3
             'edit by nickc 2005/05/04
             'lbl1(2).Caption = "達成草圖：  " & CheckStr(.Fields(1)) & "  件  " & CheckStr(.Fields(2)) & "  張"
             lbl1(2).Caption = "達成草圖：  " & Format(CheckStr(.Fields(1)), "###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(4)), "###,###,###,##0.00") & ")  件  " & Format(CheckStr(.Fields(2)), "###,###,###,##0.00") & "  張"
        Case 4
             'edit by nickc 2005/05/04
             'lbl1(3).Caption = "達成墨圖：  " & CheckStr(.Fields(1)) & "  件  " & CheckStr(.Fields(2)) & "  張  " & CheckStr(.Fields(3)) & "  點"
             lbl1(3).Caption = "達成墨圖：  " & Format(CheckStr(.Fields(1)), "###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(4)), "###,###,###,##0.00") & ")  件  " & Format(CheckStr(.Fields(2)), "###,###,###,##0.00") & "  張  " & Format(CheckStr(.Fields(3)), "###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(6)), "###,###,###,##0.00") & ")  點"
        Case 5
             lbl1(4).Caption = "其他新案：  " & Format(CheckStr(.Fields(1)), "###,###,###,##0.00") & "  件  " & Format(CheckStr(.Fields(3)), "###,###,###,##0.00") & "  點"
        Case 6
             lbl1(5).Caption = "其他舊案：  " & Format(CheckStr(.Fields(1)), "###,###,###,##0.00") & "  件  " & Format(CheckStr(.Fields(3)), "###,###,###,##0.00") & "  點"
        Case 7
             lbl1(6).Caption = "逾時草圖：  " & Format(CheckStr(.Fields(1)), "###,###,###,##0.00") & "  件"
        Case 8
             lbl1(7).Caption = "逾時墨圖：  " & Format(CheckStr(.Fields(1)), "###,###,###,##0.00") & "  件"
        Case 9
             'edit by nickc 2005/05/04
             'lbl1(8).Caption = "本月發文：  " & CheckStr(.Fields(1)) & "  件  " & CheckStr(.Fields(2)) & "  張  " & CheckStr(.Fields(3)) & "  點"
             lbl1(8).Caption = "本月發文：  " & Format(CheckStr(.Fields(1)), "###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(4)), "###,###,###,##0.00") & ")  件  " & Format(CheckStr(.Fields(2)), "###,###,###,##0.00") & "  張  " & Format(CheckStr(.Fields(3)), "###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(6)), "###,###,###,##0.00") & ")  點"
        'add by nickc 2005/05/04
        Case 10
            lbl1(9).Caption = "提供圖檔(0.6)：  " & Format(CheckStr(.Fields(1)), "###,###,###,##0.00") & "  件"
        Case 11
            lbl1(10).Caption = "關聯案件(0.4)：  " & Format(CheckStr(.Fields(1)), "###,###,###,##0.00") & "  件"
        Case Else
        End Select
        .MoveNext
        Loop
    End If
End With
CheckOC
End Sub
