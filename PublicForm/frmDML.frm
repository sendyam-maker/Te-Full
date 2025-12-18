VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmDML 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件基本資料維護紀錄"
   ClientHeight    =   7992
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   12768
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7992
   ScaleWidth      =   12768
   Begin VB.CheckBox Check1 
      Caption         =   "大小寫視為相異"
      Height          =   255
      Left            =   10890
      TabIndex        =   24
      Top             =   750
      Width           =   1725
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   23
      Top             =   735
      Width           =   9015
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frmDML.frx":0000
      Left            =   8250
      List            =   "frmDML.frx":002B
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   30
      Width           =   2625
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   6180
      MaxLength       =   8
      TabIndex        =   8
      Top             =   390
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   4980
      MaxLength       =   8
      TabIndex        =   7
      Top             =   390
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   900
      MaxLength       =   6
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   5724
      MaxLength       =   9
      TabIndex        =   4
      Top             =   30
      Width           =   1425
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   3
      Top             =   30
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   2796
      MaxLength       =   1
      TabIndex        =   2
      Top             =   30
      Width           =   285
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   1
      Top             =   30
      Width           =   1035
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   900
      MaxLength       =   3
      TabIndex        =   0
      Top             =   30
      Width           =   555
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   11850
      TabIndex        =   10
      Top             =   -30
      Width           =   885
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   10920
      TabIndex        =   9
      Top             =   -30
      Width           =   885
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   5115
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   12735
      _ExtentX        =   22458
      _ExtentY        =   9017
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSForms.TextBox Text1 
      Height          =   1515
      Left            =   30
      TabIndex        =   20
      Top             =   6450
      Width           =   12675
      VariousPropertyBits=   -1467987937
      BackColor       =   -2147483644
      ScrollBars      =   2
      Size            =   "22357;2672"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label5 
      Height          =   225
      Left            =   2040
      TabIndex        =   25
      Top             =   390
      Width           =   1905
      Size            =   "3360;397"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "異動內容包含文字："
      Height          =   180
      Left            =   30
      TabIndex        =   22
      Top             =   780
      Width           =   1620
   End
   Begin VB.Label Label10 
      Caption         =   "記錄起日 : 基本檔(2006/03/20)、客戶代理人(2006/12/20)、變更名稱(2007/01/03)"
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   8040
      TabIndex        =   21
      Top             =   360
      Width           =   4545
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "西元年喔"
      Height          =   180
      Left            =   7290
      TabIndex        =   19
      Top             =   450
      Width           =   720
   End
   Begin VB.Label Label8 
      Height          =   1500
      Left            =   30
      TabIndex        =   18
      Top             =   6450
      Width           =   12690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "詳細異動內容："
      Height          =   180
      Left            =   30
      TabIndex        =   17
      Top             =   6270
      Width           =   1260
   End
   Begin VB.Line Line2 
      X1              =   5670
      X2              =   6780
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   3405
      Y1              =   168
      Y2              =   183
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "資料表："
      Height          =   180
      Left            =   7500
      TabIndex        =   16
      Top             =   90
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "操作日期："
      Height          =   180
      Left            =   4080
      TabIndex        =   15
      Top             =   450
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "操作人員："
      Height          =   180
      Left            =   30
      TabIndex        =   14
      Top             =   420
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收文號／XYR編號："
      Height          =   180
      Left            =   4080
      TabIndex        =   13
      Top             =   96
      Width           =   1608
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   30
      TabIndex        =   12
      Top             =   96
      Width           =   900
   End
End
Attribute VB_Name = "frmDML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/2 改成Form2.0 (Label5,Text1)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit

Dim rsTmp As New ADODB.Recordset 'Added by Morgan 2024/8/27
Dim stSign As String 'Add by Amy 2025/09/23

Private Sub cmdok_Click(Index As Integer)
Dim ii As Integer, stMsg As String  'Add by Amy 2025/09/23

Select Case Index
Case 0
    'Add by Amy 2025/09/23 收文號/XYR編號欄輸客戶編號,不滿8碼,詢問是否補足8碼查詢(程式用=),前幾碼比-秀玲
    stSign = "="
    If Trim(txt1(4)) <> "" Then
      If (Left(Trim(txt1(4)), 1) = "X" Or Left(Trim(txt1(4)), 1) = "Y" Or Left(Trim(txt1(4)), 1) = "R") Then
         If Len(Trim(txt1(4))) > 8 Then
            stMsg = "目前[客戶編號]輸入 " & Len(Trim(txt1(4))) & " 碼" & vbCrLf & _
                            "資料表Customer 相關之XYR編號目前只寫入8碼" & vbCrLf & _
                            "要繼續執行？" & vbCrLf & _
                            "是：繼續查詢" & vbCrLf & _
                            "否：回前畫面修改"
            ii = MsgBox(stMsg, vbYesNo + vbCritical)
            If ii = vbNo Then
               Exit Sub
            End If
         ElseIf Len(Trim(txt1(4))) < 8 Then
            stMsg = "目前[XYR編號]輸入 " & Len(Trim(txt1(4))) & " 碼" & vbCrLf & _
                           "是要補足8碼查詢(程式用=),或是以目前碼數比對？" & vbCrLf & _
                           "是：補足8碼查詢(程式用=)" & vbCrLf & _
                           "否：以目前碼數比對" & vbCrLf & _
                           "取消：回前畫面修改"
            ii = MsgBox(stMsg, vbYesNoCancel + vbCritical)
            If ii = vbCancel Then
               Exit Sub
            ElseIf ii = vbYes Then
               'Modify by Amy 2025/10/22 +left 取8碼
               txt1(4) = Left(GetNewFagent(txt1(4)), 8)
            Else
               stSign = "LEN"
            End If
         End If
      End If
    End If
    'end 2025/09/23
    
    'Label8.Caption = ""
    Text1.Text = ""
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    Grd1.MousePointer = flexArrowHourGlass
    StrMenu
    Grd1.MousePointer = flexDefault
    Screen.MousePointer = vbDefault
    Me.Enabled = True
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(6) = strSrvDate(1)
txt1(7) = strSrvDate(1)
SetGrd
'Add By Sindy 2011/3/14
Screen.MousePointer = vbHourglass
Me.Combo1.Clear
'Modify By Sindy 2023/8/7 調整SQL速度
'strSql = "select distinct upper(replace(dl10,' ',null)) from dml_log" & _
'         " where dl10<'非'" & _
'         " order by upper(replace(dl10,' ',null))"
'Modified by Morgan 2023/9/6
strSql = "select distinct replace(dl10,' ',null) from(" & _
         "select upper(dl10) dl10 from dml_log" & _
         " where dl10<'非'" & _
         " group by dl10)" & _
         " order by replace(dl10,' ',null)"
strSql = "select table_name From user_tables" & _
   " where exists(select * from dml_log where dl13=table_name) order by 1"
'2023/8/7 END
intI = 1
Set RsTemp = ClsLawReadRstMsg(intI, strSql)
Me.Combo1.AddItem ""
If intI = 1 Then
   RsTemp.MoveFirst
   While Not RsTemp.EOF
      Me.Combo1.AddItem "" & Trim(RsTemp.Fields(0).Value)
      RsTemp.MoveNext
   Wend
End If
Screen.MousePointer = vbDefault
'2011/3/14 End
Me.Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDML = Nothing
End Sub

Sub StrMenu()
Dim oStrSQL As String
oStrSQL = ""
If Trim(txt1(0)) <> "" Then
    oStrSQL = oStrSQL & " and dl01='" & txt1(0) & "' "
End If
If Trim(txt1(1)) <> "" Then
    oStrSQL = oStrSQL & " and dl02='" & txt1(1) & "' "
End If
If Trim(txt1(2)) <> "" Then
    oStrSQL = oStrSQL & " and dl03='" & txt1(2) & "' "
End If
If Trim(txt1(3)) <> "" Then
    oStrSQL = oStrSQL & " and dl04='" & txt1(3) & "' "
End If
If Trim(txt1(4)) <> "" Then
   'Modify by Amy 2025/09/23 +if 收文號/XYR編號欄輸客戶編號,不滿8碼,詢問是否補足8碼查詢(程式用=),前幾碼比對-秀玲
   If stSign = "=" Then
      oStrSQL = oStrSQL & " and dl05='" & txt1(4) & "' "
  Else
      oStrSQL = oStrSQL & " and SubStr(dl05,1," & Len(txt1(4)) & ") ='" & txt1(4) & "' "
  End If
  'end 2025/09/23
End If
If Trim(txt1(5)) <> "" Then
    oStrSQL = oStrSQL & " and dl06='" & txt1(5) & "' "
End If
If Trim(txt1(6)) <> "" Then
    oStrSQL = oStrSQL & " and dl07>=" & txt1(6) & " "
End If
If Trim(txt1(7)) <> "" Then
    oStrSQL = oStrSQL & " and dl07<=" & txt1(7) & " "
End If
'Added by Morgan 2011/11/4
If Trim(txt1(8)) <> "" Then
   If Check1.Value = 1 Then
      oStrSQL = oStrSQL & " and instr(DL09,'" & ChgSQL(txt1(8)) & "')>0 "
   Else
      oStrSQL = oStrSQL & " and instr(upper(DL09),'" & UCase(ChgSQL(txt1(8))) & "')>0 "
   End If
End If

If Trim(Combo1.Text) <> "" And UCase(Trim(Combo1.Text)) <> "ALL" Then
    '2009/5/14 MODIFY BY SONIA 因為DL06=89037 and DL07=20080516 and DL08=105104 的DL10為 ' FAGENT'
    'oStrSQL = oStrSQL & " and UPPER(dl10)='" & UCase(Trim(Combo1.Text)) & "' "
    'Modified by Morgan 2023/9/6
    'oStrSQL = oStrSQL & " and INSTR(UPPER(dl10),'" & UCase(Trim(Combo1.Text)) & "')>0 "
    oStrSQL = oStrSQL & " and dl13='" & UCase(Trim(Combo1.Text)) & "'"
    'end 2023/9/6
End If
'Modify by Amy 2025/09/23 DL08不滿6碼補0,讓早上時間排在後面
'Modify by Amy 2025/11/07 原Decode(length(dl08),5,'0'||dl08,''||dl08),半夜跑的不會補0,改用lpad
'oStrSQL = "select dl01||'-'||dl02||'-'||dl03||'-'||dl04,dl05,dl06||' '||st02,sqldatew(dl07),dl08,dl10,dl12,dl09 from dml_log,staff where dl06=st01(+) " & oStrSQL
oStrSQL = "select dl01||'-'||dl02||'-'||dl03||'-'||dl04,dl05,dl06||' '||st02,sqldatew(dl07),Lpad(dl08,6,'0'),dl10,dl12,dl09 from dml_log,staff where dl06=st01(+) " & oStrSQL
'2008/10/22 add by sonia
'oStrSQL = oStrSQL & " order by dl07 desc,dl08 desc"
oStrSQL = oStrSQL & " order by dl07 desc,Lpad(dl08,6,'0') desc"
'2008/10/22 end
'end 2025/11/07

'Modified by Morgan 2024/8/27 因grid的欄位有限制長度,改為全域變數，點選時直接讀資料集
'Dim rsTmp As New ADODB.Recordset
'Set rsTmp = New ADODB.Recordset
'end 2024/8/27

If rsTmp.State = 1 Then rsTmp.Close
rsTmp.CursorLocation = adUseClient
rsTmp.Open oStrSQL, cnnConnection, adOpenStatic, adLockReadOnly
If rsTmp.RecordCount <> 0 Then
    Set Grd1.Recordset = rsTmp
Else
    Grd1.Clear
    Grd1.Rows = 2
    MsgBox "查無資料！"
End If
SetGrd

'rsTmp.Close 'Removed by Morgan 2024/8/27
End Sub

Private Sub SetGrd()
With Grd1
    .Cols = 8
    .row = 0
    .col = 0: .Text = "本所案號"
    .ColWidth(0) = 1400
    .ColAlignment(0) = flexAlignLeftCenter
    'Modify by Amy 2025/09/23 +/XYR編號
    .col = 1: .Text = "收文號/XYR編號"
    .ColWidth(1) = 1200 '原:1000
    '.CellAlignment = flexAlignCenterCenter
    .CellAlignment = flexAlignLeftCenter
    'end 2025/09/23
    .col = 2: .Text = "操作人員"
    .ColWidth(2) = 1300
    .CellAlignment = flexAlignCenterCenter
    .col = 3: .Text = "操作日"
    .ColWidth(3) = 900
    .CellAlignment = flexAlignCenterCenter
    .col = 4: .Text = "操作時間"
    .ColWidth(4) = 750
    .CellAlignment = flexAlignCenterCenter
    .col = 5: .Text = "資料表"
    .ColWidth(5) = 1300
    .CellAlignment = flexAlignCenterCenter
    .col = 6: .Text = "功能"
    .ColWidth(6) = 1300
    .CellAlignment = flexAlignCenterCenter
    .col = 7: .Text = "異動內容"
    .ColWidth(7) = 12300
    .CellAlignment = flexAlignCenterCenter
End With
End Sub

Private Sub grd1_SelChange()
Dim SeekRow As Long
Dim i As Integer, j As Integer

If rsTmp.State <> 1 Then Exit Sub 'Added by Morgan 2024/11/25
If rsTmp.RecordCount = 0 Then Exit Sub 'Added by Morgan 2024/11/25

Grd1.Visible = False
SeekRow = Grd1.MouseRow
Grd1.col = 0
If SeekRow <> 0 Then
    Label8.Caption = ""
    For j = 1 To Grd1.Rows - 1
        Grd1.row = j
        If Grd1.CellBackColor = &HFFC0C0 Then
             For i = 0 To Grd1.Cols - 1
                  Grd1.col = i
                  Grd1.CellBackColor = QBColor(15)
            Next i
        End If
    Next j
    Grd1.row = SeekRow
    For i = 0 To Grd1.Cols - 1
        Grd1.col = i
        Grd1.CellBackColor = &HFFC0C0
    Next i
    'Label8.Caption = grd1.TextMatrix(SeekRow, 7)
    'Modified by Morgan 2024/8/27
    'Text1.Text = GRD1.TextMatrix(SeekRow, 7)
    rsTmp.MoveFirst
    If SeekRow > 1 Then
      rsTmp.Move SeekRow - 1
   End If
    Text1.Text = "" & rsTmp.Fields(7)
    Text1.SelStart = 0
    'end 2024/8/27
End If
Grd1.Visible = True
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 0, 4, 5
     KeyAscii = UpperCase(KeyAscii)
Case Else
End Select
End Sub
