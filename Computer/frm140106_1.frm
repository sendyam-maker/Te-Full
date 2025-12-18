VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140106_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "分所內商延展、第二期註冊費銷卷作業"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9315
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   315
      Left            =   60
      TabIndex        =   33
      Top             =   30
      Width           =   2100
      Begin MSForms.TextBox txtInput 
         Height          =   375
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   1635
         VariousPropertyBits=   679493659
         MaxLength       =   40
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   4
      Left            =   8430
      TabIndex        =   5
      Top             =   30
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   5175
      Left            =   0
      TabIndex        =   4
      Top             =   390
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   9128
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
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
      _Band(0).Cols   =   1
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "存檔(&O)"
      Height          =   345
      Index           =   2
      Left            =   6270
      TabIndex        =   3
      Top             =   30
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&N)"
      Height          =   345
      Index           =   3
      Left            =   7155
      TabIndex        =   2
      Top             =   30
      Width           =   1245
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "全部選取(&C)"
      Height          =   345
      Index           =   0
      Left            =   3690
      TabIndex        =   1
      Top             =   30
      Width           =   1275
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "全部取消(&U)"
      Height          =   345
      Index           =   1
      Left            =   4995
      TabIndex        =   0
      Top             =   30
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "。"
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
      Height          =   180
      Index           =   25
      Left            =   9090
      TabIndex        =   32
      Top             =   5130
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "面"
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
      Height          =   180
      Index           =   24
      Left            =   9090
      TabIndex        =   31
      Top             =   4950
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "畫"
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
      Height          =   180
      Index           =   23
      Left            =   9090
      TabIndex        =   30
      Top             =   4770
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "於"
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
      Height          =   180
      Index           =   22
      Left            =   9090
      TabIndex        =   29
      Top             =   4590
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "存"
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
      Height          =   180
      Index           =   21
      Left            =   9090
      TabIndex        =   28
      Top             =   4410
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "暫"
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
      Height          =   180
      Index           =   20
      Left            =   9090
      TabIndex        =   27
      Top             =   4230
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "r"
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
      Height          =   180
      Index           =   19
      Left            =   9090
      TabIndex        =   26
      Top             =   4050
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "e"
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
      Height          =   180
      Index           =   18
      Left            =   9090
      TabIndex        =   25
      Top             =   3870
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "t"
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
      Height          =   180
      Index           =   17
      Left            =   9090
      TabIndex        =   24
      Top             =   3690
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "n"
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
      Height          =   180
      Index           =   16
      Left            =   9090
      TabIndex        =   23
      Top             =   3510
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "E"
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
      Height          =   180
      Index           =   15
      Left            =   9090
      TabIndex        =   22
      Top             =   3330
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "按"
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
      Height          =   180
      Index           =   14
      Left            =   9090
      TabIndex        =   21
      Top             =   3150
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "完"
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
      Height          =   180
      Index           =   13
      Left            =   9090
      TabIndex        =   20
      Top             =   2970
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "改"
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
      Height          =   180
      Index           =   12
      Left            =   9090
      TabIndex        =   19
      Top             =   2790
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "，"
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
      Height          =   180
      Index           =   11
      Left            =   9090
      TabIndex        =   18
      Top             =   2610
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "註"
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
      Height          =   180
      Index           =   10
      Left            =   9090
      TabIndex        =   17
      Top             =   2430
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備"
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
      Height          =   180
      Index           =   9
      Left            =   9090
      TabIndex        =   16
      Top             =   2250
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "輯"
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
      Height          =   180
      Index           =   8
      Left            =   9090
      TabIndex        =   15
      Top             =   2070
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "編"
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
      Height          =   180
      Index           =   7
      Left            =   9090
      TabIndex        =   14
      Top             =   1890
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "以"
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
      Height          =   180
      Index           =   6
      Left            =   9090
      TabIndex        =   13
      Top             =   1710
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "可"
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
      Height          =   180
      Index           =   5
      Left            =   9090
      TabIndex        =   12
      Top             =   1530
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "就"
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
      Height          =   180
      Index           =   4
      Left            =   9090
      TabIndex        =   11
      Top             =   1350
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "註"
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
      Height          =   180
      Index           =   3
      Left            =   9090
      TabIndex        =   10
      Top             =   1170
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備"
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
      Height          =   180
      Index           =   2
      Left            =   9090
      TabIndex        =   9
      Top             =   990
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "下"
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
      Height          =   180
      Index           =   1
      Left            =   9090
      TabIndex        =   8
      Top             =   810
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "2"
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
      Height          =   180
      Left            =   9090
      TabIndex        =   7
      Top             =   630
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "點"
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
      Height          =   180
      Index           =   0
      Left            =   9090
      TabIndex        =   6
      Top             =   450
      Width           =   195
   End
End
Attribute VB_Name = "frm140106_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/03 Form2.0 grd1/txtInput
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer
Dim iRow As Integer '本次點選列數
Dim iCol As Integer '智權人員名稱欄位
Dim ii As Integer
Dim oKey As String
Dim IsSave As Boolean


Private Sub cmdOK_Click(Index As Integer)
'Add by Amy 2021/12/03 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Sub
End If

Select Case Index
Case 0
    With GRD1
        .Visible = False
        For j = 1 To .Rows - 1
            .row = j
            .col = 0
            .Text = "V"
            For i = 0 To .Cols - 1
                .col = i
                .CellBackColor = &HFFC0C0
            Next i
        Next j
        .Visible = True
    End With
Case 1
    With GRD1
        .Visible = False
        For j = 1 To .Rows - 1
            .row = j
            .col = 0
            .Text = ""
            For i = 0 To .Cols - 1
                 .col = i
                 .CellBackColor = QBColor(15)
            Next i
        Next j
        .Visible = True
    End With
Case 2
    With GRD1
        .Visible = False
        IsSave = False
        Screen.MousePointer = vbHourglass
        On Error GoTo oErr
        cnnConnection.BeginTrans
        For j = 1 To .Rows - 1
            .row = j
            .col = 0
            oKey = ""
            If .Text = "V" Then
                .col = 2
                oKey = .Text
                .col = 7
                'Modify by Amy 2021/12/03 bug-發現增加的文字未寫入
                'cnnConnection.Execute "update trademark set tm73=to_number(to_char(sysdate,'YYYYMMDD')),tm74='" & strUserNum & "',tm75='" & ChgSQL(.Text) & "' where tm01='" & SystemNumber(oKey, 1) & "' and tm02='" & SystemNumber(oKey, 2) & "' and tm03='" & SystemNumber(oKey, 3) & "' and tm04='" & SystemNumber(oKey, 4) & "' "
                cnnConnection.Execute "update trademark set tm73=to_number(to_char(sysdate,'YYYYMMDD')),tm74='" & strUserNum & "',tm75='" & ChgSQL(txtInput.Text) & "' where tm01='" & SystemNumber(oKey, 1) & "' and tm02='" & SystemNumber(oKey, 2) & "' and tm03='" & SystemNumber(oKey, 3) & "' and tm04='" & SystemNumber(oKey, 4) & "' "
                IsSave = True
                .col = 0
                .Text = ""
                For i = 0 To .Cols - 1
                     .col = i
                     .CellBackColor = QBColor(15)
                Next i
            End If
        Next j
        cnnConnection.CommitTrans
        Screen.MousePointer = vbDefault
        If IsSave = True Then
            MsgBox "存檔成功！", vbInformation
            StrMenu
        End If
        .Visible = True
    End With
Case 3
     frm140106.Show
     Unload Me
Case 4
     Unload frm140106
     Unload Me
Case Else
End Select
Exit Sub
oErr:
    cnnConnection.RollbackTrans
    MsgBox "銷卷存檔失敗！" & vbCrLf & "請稍後再試！", vbExclamation
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Screen.MousePointer = vbHourglass
Me.Enabled = False
StrMenu
Me.Enabled = True
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm140106_1 = Nothing
End Sub

Sub StrMenu()

'2007/8/17 modify by sonia 加判斷專用期止日小於系統日
'strSQL = "select ' ',tm34,np02||'-'||np03||'-'||np04||'-'||np05,NVL(DECODE(tm10,'000',CPM03,CPM04),to_char(np07)),tm05," & SQLDate("NP09") & ",NVL(S1.ST02,np10),NVL(DECODE(tm10,'000',CPM03,CPM04),to_char(np07))||'不續辦' from nextprogress,trademark,casepropertymap,staff S1,staff S2,customer " & _
'          " Where np02=cpm01(+) and to_char(np07)=cpm02(+) and np02=tm01 and np03=tm02 and np04=tm03 and np05=tm04 and (np06 ='N' or np06 is null) and np07 in (" & IIf(frm140106.txt1(3) = "1", "102", IIf(frm140106.txt1(3) = "2", "716", "102,716")) & ") and np02='T' and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and np09>=" & ChangeTStringToWString(frm140106.txt1(0)) & " and np09<=" & ChangeTStringToWString(frm140106.txt1(1)) & " and cu13=s2.st01(+) and s2.st06='" & frm140106.lblst06 & "' and np10=s1.st01(+) " & IIf(Trim(frm140106.txt1(2)) = "", "", " and np10='" & frm140106.txt1(2) & "' ") & " and tm73 is null order by 1,2"
strSql = "select ' ',tm34,np02||'-'||np03||'-'||np04||'-'||np05,NVL(DECODE(tm10,'000',CPM03,CPM04),to_char(np07)),tm05," & SQLDate("NP09") & ",NVL(S1.ST02,np10),NVL(DECODE(tm10,'000',CPM03,CPM04),to_char(np07))||'不續辦' from nextprogress,trademark,casepropertymap,staff S1,staff S2,customer " & _
          " Where np02=cpm01(+) and to_char(np07)=cpm02(+) and np02=tm01 and np03=tm02 and np04=tm03 and np05=tm04 and tm22<=" & GetTodayDate & " and (np06 ='N' or np06 is null) and np07 in (" & IIf(frm140106.txt1(3) = "1", "102", IIf(frm140106.txt1(3) = "2", "716", "102,716")) & ") and np02='T' and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and np09>=" & ChangeTStringToWString(frm140106.txt1(0)) & " and np09<=" & ChangeTStringToWString(frm140106.txt1(1)) & " and cu13=s2.st01(+) and s2.st06='" & frm140106.lblst06 & "' and np10=s1.st01(+) " & IIf(Trim(frm140106.txt1(2)) = "", "", " and np10='" & frm140106.txt1(2) & "' ") & " and tm73 is null order by 2,3"
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        Set GRD1.Recordset = adoRecordset
        SetGrd
    Else
       ShowNoData
       Exit Sub
    End If
End With
End Sub

Sub SetGrd()
With GRD1
    .Cols = 8
    .row = 0
    .col = 0
    .Text = "V"
    .ColWidth(0) = 200
    .CellAlignment = flexAlignCenterCenter
    .col = 1
    .Text = "分所案號"
    .ColWidth(1) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 2
    .Text = "本所案號"
    .ColWidth(2) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 3
    .Text = "案件性質"
    .ColWidth(3) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 4
    .Text = "案件名稱"
    .ColWidth(4) = 1200
    .CellAlignment = flexAlignCenterCenter
    .col = 5
    .Text = "法定期限"
    .ColWidth(5) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 6
    .Text = "智權人員"
    .ColWidth(6) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 7
    .Text = "備註"
    .ColWidth(7) = 2000
    .CellAlignment = flexAlignCenterCenter
End With
End Sub

Private Sub Grd1_Click()
With GRD1
    .Visible = False
    .row = .MouseRow
    .col = 0
    If .row <> 0 Then
        If .Text = "V" Then
             .Text = ""
             For i = 0 To .Cols - 1
                  .col = i
                  .CellBackColor = QBColor(15)
            Next i
        Else
             .Text = "V"
             For i = 0 To .Cols - 1
                 .col = i
                 .CellBackColor = &HFFC0C0
             Next i
        End If
    End If
    .Visible = True
End With
End Sub

Private Sub GRD1_DblClick()
Screen.MousePointer = vbHourglass
    GRD1.Visible = False
    'Modify by Amy 2021/12/03 Form2.0 txtInput 圖層會在最下方,故加frame1
    'txtInput.Visible = False
    Frame1.Visible = False
    If Me.GRD1.row > 0 Then
        SetBox
    End If
    GRD1.Visible = True
Screen.MousePointer = vbDefault
End Sub

Private Sub txtInput_GotFocus()
txtInput.SelStart = 0
txtInput.SelLength = Len(txtInput)
End Sub

'Add by Amy 2021/12/03 從KeyPress搬過來修改
Private Sub txtInput_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim Cancel  As Boolean
   If KeyCode = vbKeyReturn Then
      Cancel = False
      txtInputValidate Cancel
      If Cancel = False Then
         GRD1.TextMatrix(iRow, iCol) = txtInput.Text
         GRD1.SetFocus
         GRD1.Refresh
         Frame1.Visible = False
      End If
   ElseIf KeyCode = vbKeyEscape Then
      GRD1.SetFocus
      Frame1.Visible = False
   End If
End Sub

'Mark by Amy 2021/12/03 按Enter字會消失
'Private Sub txtInput_KeyPress(KeyAscii As Integer)
'   Dim Cancel  As Boolean
'   If KeyAscii = vbKeyReturn Then
'      Cancel = False
'      txtInputValidate Cancel
'      If Cancel = False Then
'         GRD1.TextMatrix(iRow, iCol) = txtInput.Text
'         GRD1.SetFocus
'         GRD1.Refresh
'         'Modify by Amy 2021/12/03 Form2.0 txtInput 圖層會在最下方,故加frame1
'         'txtInput.Visible = False
'         Frame1.Visible = False
'      End If
'   ElseIf KeyAscii = vbKeyEscape Then
'      GRD1.SetFocus
'      'Modify by Amy 2021/12/03 Form2.0 txtInput 圖層會在最下方,故加frame1
'      'txtInput.Visible = False
'      Frame1.Visible = False
'   End If
'End Sub


Private Sub txtInput_LostFocus()
   'Modify by Amy 2021/12/03 Form2.0 txtInput 圖層會在最下方,故加frame1
   'txtInput.Visible = False
   Frame1.Visible = False
   txtInput.Tag = ""
End Sub

Private Sub SetBox()
   Dim lngLeft As Long, lngTop As Long
   
   With GRD1
      If .row > 0 And .col = 7 Then
         'If .TextMatrix(.Row, 7) <> "" Then
            txtInput.FontName = .CellFontName
            txtInput.FontSize = .CellFontSize
            'Modify by Amy 2021/12/03 Form2.0 無Alignment屬性
            'txtInput.Alignment = .CellAlignment \ 5
            txtInput.TextAlign = 1
            txtInput.Text = .TextMatrix(.row, .col)
            txtInput.Tag = txtInput.Text
            'Modify by Amy 2021/12/03 Form2.0 txtInput 圖層會在最下方,故加frame1
            Frame1.Width = .ColWidth(.col)
            Frame1.Height = .RowHeight(.row)
            txtInput.Width = .ColWidth(.col)
            txtInput.Height = .RowHeight(.row) + 30
            iRow = .row: iCol = .col
            'Modify by Amy 2021/12/03 Form2.0 txtInput 圖層會在最下方,故加frame1
            Frame1.Visible = True
            txtInput.Visible = True
            txtInput.Enabled = True
            txtInput.SetFocus
            txtInput.SelStart = 0
            txtInput.SelLength = Len(txtInput)
            lngLeft = .Left + 25
            lngTop = .Top + 25 + .RowHeight(iRow)
            For ii = 0 To .col - 1
               lngLeft = lngLeft + .ColWidth(ii)
            Next
            For ii = .TopRow To .row - 1
               lngTop = lngTop + .RowHeight(ii)
            Next
            'Modify by Amy 2021/12/03 Form2.0 txtInput 圖層會在最下方,故加frame1
            'txtInput.Left = lngLeft: txtInput.Top = lngTop
            Frame1.Left = lngLeft: Frame1.Top = lngTop - 20
         'End If
      End If
   End With
End Sub

Private Sub txtInputValidate(Cancel As Boolean)
Cancel = False
If CheckLengthIsOK(txtInput.Text, txtInput.MaxLength) = False Then
    txtInput.SetFocus
    txtInput_GotFocus
    Cancel = True
End If
End Sub
