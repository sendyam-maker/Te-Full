VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210115_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "多筆申請人資料查詢"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7365
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   90
      Top             =   30
   End
   Begin VB.TextBox NoData 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   1740
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "資料庫找不到資料！"
      Top             =   2250
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   405
      Index           =   2
      Left            =   3900
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Left            =   870
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   435
      Index           =   1
      Left            =   6030
      TabIndex        =   1
      Top             =   60
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Height          =   435
      Index           =   0
      Left            =   4965
      TabIndex        =   0
      Top             =   60
      Width           =   945
   End
   Begin VB.TextBox SelOne 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1950
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "請選擇一位客戶！"
      Top             =   2220
      Visible         =   0   'False
      Width           =   3525
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   4215
      Left            =   60
      TabIndex        =   5
      Top             =   540
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "客戶名稱："
      Height          =   180
      Left            =   150
      TabIndex        =   4
      Top             =   570
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(字首比對查詢)"
      Height          =   180
      Index           =   2
      Left            =   4020
      TabIndex        =   3
      Top             =   570
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "frm210115_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/07 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim TmpRow As Integer
Dim i As Integer
Dim j As Integer

Private Sub cmdOK_Click(Index As Integer)
NoData.Visible = False
SelOne.Visible = False
Select Case Index
Case 0
        TmpRow = 0
        For j = 1 To grd1.Rows - 1
            grd1.row = j
            grd1.col = 0
            If grd1.CellBackColor = &HFFC0C0 Then
                TmpRow = j
            End If
        Next j
        If TmpRow = 0 Then
            'MsgBox "請選擇一位客戶！", vbExclamation, "操作錯誤！"
            SelOne.Visible = True
            Exit Sub
        End If
        frm210115.txt1(4).Text = grd1.TextMatrix(TmpRow, 0)
        Unload Me
Case 1
        Unload Me
Case 2
        StrMenu
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1 = frm210115.txt1(6)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm210115_1 = Nothing
End Sub

Public Sub StrMenu()
grd1.Clear
grd1.Rows = 2
grd1.MousePointer = flexArrowHourGlass
Screen.MousePointer = vbHourglass
DoEvents
SetGrd
strSql = "SELECT CU01||CU02 AS 編號,CU04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,CU80 AS 狀態,CU79 AS 備註,cu13 FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1,cu02 as A2 From Customer Where CU04 like '" & ChgSQL(txt1) & "%' ) A WHERE CU10=NA01(+) AND CU01=A.A1 and cu02=A.A2 AND CU13=ST01(+) "
strSql = strSql & " union SELECT CU01||CU02 AS 編號,cu05||' '||cu88||' '||cu89||' '||cu90 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,CU80 AS 狀態,CU79 AS 備註,cu13 FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1,cu02 as A2 From Customer Where upper(cu05||' '||cu88||' '||cu89||' '||cu90) like '" & UCase(ChgSQL(txt1)) & "%' ) A WHERE CU10=NA01(+) AND CU01=A.A1 and cu02=A.A2 AND CU13=ST01(+) "
strSql = strSql & " union SELECT CU01||CU02 AS 編號,CU06 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,CU80 AS 狀態,CU79 AS 備註,cu13 FROM CUSTOMER,NATION,STAFF, (Select Distinct CU01 As A1,cu02 as A2 From Customer Where CU06 like '" & ChgSQL(txt1) & "%' ) A WHERE CU10=NA01(+) AND CU01=A.A1 and cu02=A.A2 AND CU13=ST01(+) "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        '若是只有一筆，直接帶回前畫面
        If .RecordCount = 1 Then
            frm210115.txt1(4).Text = CheckStr(.Fields(0))
            Unload Me
            Exit Sub
        End If
        Set grd1.Recordset = adoRecordset
        SetGrd
    Else
        NoData.Visible = True
    End If
End With
grd1.MousePointer = flexDefault
Screen.MousePointer = vbDefault
End Sub

Private Sub grd1_SelChange()
Dim CtmpRow As Integer
CtmpRow = TmpRow
grd1.Visible = False
TmpRow = grd1.MouseRow
If TmpRow <> 0 Then
    If CtmpRow <> 0 Then
         grd1.row = CtmpRow
         For i = 0 To grd1.Cols - 1
              grd1.col = i
              grd1.CellBackColor = QBColor(15)
        Next i
    Else
        For j = 1 To grd1.Rows - 1
            grd1.row = j
            For i = 0 To grd1.Cols - 1
                 grd1.col = i
                 grd1.CellBackColor = QBColor(15)
           Next i
        Next j
    End If
    grd1.row = TmpRow
    For i = 0 To grd1.Cols - 1
        grd1.col = i
        grd1.CellBackColor = &HFFC0C0
    Next i
End If
grd1.Visible = True
End Sub

Private Sub SetGrd()
grd1.Cols = 7
grd1.row = 0
grd1.col = 0: grd1.Text = "編號"
grd1.ColWidth(0) = 800
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 1: grd1.Text = "名稱"
grd1.ColWidth(1) = 2000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 2: grd1.Text = "國籍"
grd1.ColWidth(2) = 1200
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 3: grd1.Text = "智權人員"
grd1.ColWidth(3) = 1500
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 4: grd1.Text = "狀態"
grd1.ColWidth(4) = 1000
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 5: grd1.Text = "備註"
grd1.ColWidth(5) = 1200
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 6: grd1.Text = ""
grd1.ColWidth(6) = 0
grd1.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub Timer1_Timer()
If NoData.ForeColor = NoData.BackColor Then
    NoData.ForeColor = &HFF&
Else
    NoData.ForeColor = NoData.BackColor
End If
If SelOne.ForeColor = SelOne.BackColor Then
    SelOne.ForeColor = &HFF&
Else
    SelOne.ForeColor = SelOne.BackColor
End If
End Sub
