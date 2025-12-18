VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100102_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "多人申請組合"
   ClientHeight    =   5688
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9288
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5688
   ScaleWidth      =   9288
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   2
      Left            =   7785
      TabIndex        =   2
      Top             =   75
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "下一筆(&N)"
      Height          =   345
      Index           =   1
      Left            =   6555
      TabIndex        =   1
      Top             =   90
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "案件資料(&B)"
      Height          =   345
      Index           =   0
      Left            =   5355
      TabIndex        =   0
      Top             =   90
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   3225
      Left            =   45
      TabIndex        =   19
      Top             =   2415
      Width           =   9210
      _ExtentX        =   16235
      _ExtentY        =   5694
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   9
   End
   Begin MSForms.Label lbl17 
      Height          =   300
      Left            =   1400
      TabIndex        =   17
      Top             =   2090
      Width           =   7620
      Caption         =   "lbl17"
      Size            =   "13441;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl15 
      Height          =   300
      Left            =   1400
      TabIndex        =   16
      Top             =   1730
      Width           =   2340
      Caption         =   "lbl15"
      Size            =   "4128;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl14 
      Height          =   300
      Left            =   1400
      TabIndex        =   15
      Top             =   1340
      Width           =   7620
      Caption         =   "lbl14"
      Size            =   "13441;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl12 
      Height          =   300
      Left            =   1400
      TabIndex        =   13
      Top             =   750
      Width           =   7620
      Caption         =   "lbl12"
      Size            =   "13441;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   6
      Left            =   4940
      TabIndex        =   18
      Top             =   1740
      Width           =   2310
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   3
      Left            =   1400
      TabIndex        =   14
      Top             =   1110
      Width           =   7620
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   1
      Left            =   4580
      TabIndex        =   12
      Top             =   530
      Width           =   2340
   End
   Begin VB.Label lbl1 
      Height          =   180
      Index           =   0
      Left            =   1400
      TabIndex        =   11
      Top             =   510
      Width           =   2340
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "參考備註："
      Height          =   180
      Left            =   270
      TabIndex        =   10
      Top             =   2090
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "客戶狀態："
      Height          =   180
      Left            =   3920
      TabIndex        =   9
      Top             =   1730
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   270
      TabIndex        =   8
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "        (日)："
      Height          =   180
      Left            =   290
      TabIndex        =   7
      Top             =   1350
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "        (英)："
      Height          =   180
      Left            =   270
      TabIndex        =   6
      Top             =   1110
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "名稱(中)："
      Height          =   180
      Left            =   270
      TabIndex        =   5
      Top             =   740
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "國籍："
      Height          =   180
      Left            =   3920
      TabIndex        =   4
      Top             =   530
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人編號："
      Height          =   180
      Left            =   270
      TabIndex        =   3
      Top             =   510
      Width           =   1080
   End
End
Attribute VB_Name = "frm100102_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/1/12 Form2.0已修改(lbl1(2)->lbl12,lbl1(4)->lbl14,lbl1(5)->lbl15,lbl1(7)->lbl17,GrdDataList改Fonts)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/26 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer
Dim StrTag As String, StrToGrid As String
Dim strSql As String, lngCounter As Long, lngCounterI As Long
Public cmdState As Integer
Public KeyString As String

Private Sub SetDataListWidth()
GrdDataList.row = 0
GrdDataList.col = 0: GrdDataList.Text = "V"
GrdDataList.ColWidth(0) = 200
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 1: GrdDataList.Text = "申請人組合"
GrdDataList.ColWidth(1) = 4000
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 2: GrdDataList.Text = "名稱"
GrdDataList.ColWidth(2) = 4000
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 3: GrdDataList.Text = "案件數"
GrdDataList.ColWidth(3) = 700
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 4: GrdDataList.Text = ""
GrdDataList.ColWidth(4) = 0
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 5: GrdDataList.Text = ""
GrdDataList.ColWidth(5) = 0
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 6: GrdDataList.Text = ""
GrdDataList.ColWidth(6) = 0
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 7: GrdDataList.Text = ""
GrdDataList.ColWidth(7) = 0
GrdDataList.CellAlignment = flexAlignCenterCenter
GrdDataList.col = 8: GrdDataList.Text = ""
GrdDataList.ColWidth(8) = 0
GrdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub cmdOK_Click(Index As Integer)
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Private Sub Form_Load()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer

bolToEndByNick = False
MoveFormToCenter Me
ClearLabel 'Add by Amy 2023/08/29
SetDataListWidth
bolToEndByNick = False
'92.04.16 nick
cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100102_4 = Nothing
End Sub

Private Sub GrdDataList_Click()
GrdDataList.Visible = False
GrdDataList.row = GrdDataList.MouseRow
GrdDataList.col = 0
If GrdDataList.row <> 0 Then
If GrdDataList.Text = "V" Then
     GrdDataList.Text = ""
     For i = 0 To GrdDataList.Cols - 1
          GrdDataList.col = i
          GrdDataList.CellBackColor = QBColor(15)
    Next i
Else
     GrdDataList.Text = "V"
     For i = 0 To GrdDataList.Cols - 1
         GrdDataList.col = i
         GrdDataList.CellBackColor = &HFFC0C0
     Next i
End If
End If
GrdDataList.Visible = True
End Sub

Public Sub PubShowNextData()
'Add By Cheng 2003/08/26
Dim blnPrintAdd As Boolean
Dim ii As Integer
Dim j As Integer

Select Case cmdState
Case 0 '案件資料
      Me.Enabled = False
      For i = 1 To GrdDataList.Rows - 1
      GrdDataList.col = 0
      GrdDataList.row = i
      If Trim(GrdDataList.Text) = "V" Then
        GrdDataList.col = 0
        GrdDataList.Text = ""
        For j = 0 To GrdDataList.Cols - 1
           GrdDataList.col = j
           GrdDataList.CellBackColor = QBColor(15)
        Next j
        GrdDataList.col = 1
         If Not IsNull(GrdDataList.Text) Then
            If fnSaveParentForm(Me) = False Then
                Me.Enabled = True
                Exit Sub
            End If
                Screen.MousePointer = vbHourglass
                frm100102_2.Show
                frm100102_2.Tag = Pub_RplStr(GrdDataList.Text)
                '2010/9/10 ADD BY SONIA 未傳條件
                frm100102_2.m_Sys = frm100102_1.Text3
                frm100102_2.m_Date1 = frm100102_1.Text4
                frm100102_2.m_Date2 = frm100102_1.Text5
                frm100102_2.m_Pty1 = frm100102_1.Text6
                frm100102_2.m_Pty2 = frm100102_1.Text7
                frm100102_2.m_CKind = frm100102_1.Text8
               '2010/9/10 END
                frm100102_2.StrMenu3
                Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
        End If
      End If
      Next i
      Me.Enabled = True
Case 1 '下一筆
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 2 '結束
     fnCloseAllFrm100
Case Else
End Select
End Sub

Public Sub StrMenu()
'add by  nickc 2005/10/03  帶基本資料
lbl1(0).Caption = KeyString

'Add By Sindy 2011/01/03 檢查國內外權限
If CheckSR12(KeyString) = False Then
   Screen.MousePointer = vbDefault
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If

Dim tmpRs As New ADODB.Recordset
Set tmpRs = New ADODB.Recordset
With tmpRs
   If .State = 1 Then .Close
   .CursorLocation = adUseClient
   .Open "select * from customer,staff,nation where cu01='" & Mid(KeyString, 1, 8) & "' and cu02='" & Mid(KeyString, 9, 1) & "' and cu10=na01(+) and cu13=st01(+) ", cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      lbl1(1).Caption = CheckStr(.Fields("na03"))
      'modify by sonia 2022/1/12
      'lbl1(2).Caption = CheckStr(.Fields("cu04"))
      lbl12.Caption = CheckStr(.Fields("cu04"))
      'end 2022/1/12
      lbl1(3).Caption = CheckStr(.Fields("cu05")) & " " & CheckStr(.Fields("cu88")) & " " & CheckStr(.Fields("cu89")) & " " & CheckStr(.Fields("cu90"))
      'modify by sonia 2022/1/12
      'lbl1(4).Caption = CheckStr(.Fields("cu06"))
      'lbl1(5).Caption = CheckStr(.Fields("st02"))
      lbl14.Caption = CheckStr(.Fields("cu06"))
      lbl15.Caption = CheckStr(.Fields("st02"))
      'end 2022/1/12
      lbl1(6).Caption = CheckStr(.Fields("cu80"))
      'modify by sonia 2022/1/12
      'lbl1(7).Caption = CheckStr(.Fields("cu79"))
      lbl17.Caption = CheckStr(.Fields("cu79"))
      'end 2022/1/12
      'add by nickc 2005/12/06
      If CheckStr(.Fields("cu111")) = "Y" Then
        lbl1(0).ForeColor = &HFF&
      Else
        lbl1(0).ForeColor = &H80000012
      End If
   End If
'add by nickc 2005/10/03 帶多申請人資料
   Dim StrSqlB As String
'edit by nickc 2007/10/24 秀玲說 cp04<>'00' 的不計
'   StrSqlB = "select '' as V,NewPa.Combine," & _
                 " decode(NewPa.pa26,null,'',nvl(c1.cu04,nvl(c1.cu05||' '||c1.cu88||' '||c1.cu89||' '||c1.cu90,c1.cu06))||',')||decode(NewPa.pa27,null,'',nvl(c2.cu04,nvl(c2.cu05||' '||c2.cu88||' '||c2.cu89||' '||c2.cu90,c2.cu06))||',')||decode(NewPa.pa28,null,'',nvl(c3.cu04,nvl(c3.cu05||' '||c3.cu88||' '||c3.cu89||' '||c3.cu90,c3.cu06))||',')||decode(NewPa.pa29,null,'',nvl(c4.cu04,nvl(c4.cu05||' '||c4.cu88||' '||c4.cu89||' '||c4.cu90,c4.cu06))||',')||decode(NewPa.pa30,null,'',nvl(c5.cu04,nvl(c5.cu05||' '||c5.cu88||' '||c5.cu89||' '||c5.cu90,c5.cu06))||',') as CombineName, " & _
                 " count(NewPa.CaseID),NewPa.pa26,NewPa.pa27,NewPa.pa28,NewPa.pa29,NewPa.pa30 from ( " & _
                 " select decode(pa26,null,'',pa26||',')||decode(pa27,null,'',pa27||',')||decode(pa28,null,'',pa28||',')||decode(pa29,null,'',pa29||',')||decode(pa30,null,'',pa30||',') as Combine, " & _
                 " pa01||pa02||pa03||pa04 as CaseID,pa26,pa27,pa28,pa29,pa30 " & _
                 " from patent where (pa27 is not null or pa28 is not null or pa29 is not null or pa30 is not null) and " & _
                 " (pa26='" & KeyString & "' or pa27='" & KeyString & "' or pa28='" & KeyString & "' or pa29='" & KeyString & "' or pa30='" & KeyString & "') " & _
                 " union select decode(sp08,null,'',sp08||',')||decode(sp58,null,'',sp58||',')||decode(sp59,null,'',sp59||',') as Combine, " & _
                 " sp01||sp02||sp03||sp04 as CaseID,sp08 as pa26,sp58 as pa27,sp59 as pa28,'' as pa29,'' as pa30 " & _
                 " from servicepractice where (sp58 is not null or sp59 is not null ) and (sp08='X15859000' or sp58='X15859000' or sp59='X15859000') ) NewPa, " & _
                 " customer C1,customer C2,customer C3,customer C4,customer C5 " & _
                 " where substr(NewPa.pa26,1,8)=c1.cu01(+) and substr(NewPa.pa26,9,1)=c1.cu02(+) and substr(NewPa.pa27,1,8)=c2.cu01(+) and substr(NewPa.Pa27,9,1)=c2.cu02(+) " & _
                 " and substr(NewPa.pa28,1,8)=c3.cu01(+) and substr(NewPa.Pa28,9,1)=c3.cu02(+) and substr(NewPa.Pa29,1,8)=c4.cu01(+) and substr(NewPa.pa29,9,1)=c4.cu02(+) " & _
                 " and substr(NewPa.Pa30,1,8)=c5.cu01(+) and substr(NewPa.pa30,9,1)=c5.cu02(+) group by NewPa.Combine,decode(NewPa.pa26,null,'',nvl(c1.cu04,nvl(c1.cu05||' '||c1.cu88||' '||c1.cu89||' '||c1.cu90,c1.cu06))||',')||decode(NewPa.pa27,null,'',nvl(c2.cu04,nvl(c2.cu05||' '||c2.cu88||' '||c2.cu89||' '||c2.cu90,c2.cu06))||',')||decode(NewPa.pa28,null,'',nvl(c3.cu04,nvl(c3.cu05||' '||c3.cu88||' '||c3.cu89||' '||c3.cu90,c3.cu06))||',')||decode(NewPa.pa29,null,'',nvl(c4.cu04,nvl(c4.cu05||' '||c4.cu88||' '||c4.cu89||' '||c4.cu90,c4.cu06))||',')||decode(NewPa.pa30,null,'',nvl(c5.cu04,nvl(c5.cu05||' '||c5.cu88||' '||c5.cu89||' '||c5.cu90,c5.cu06))||','), NewPa.pa26,NewPa.pa27,NewPa.pa28,NewPa.pa29,NewPa.pa30 "
   'Modify By Sindy 2011/2/8 +SP65,SP66
   StrSqlB = "select '' as V,NewPa.Combine," & _
                 " decode(NewPa.pa26,null,'',NVL(c1.CU04,DECODE(c1.CU05,NULL,c1.CU06,c1.CU05||' '||c1.CU88||' '||c1.CU89||' '||c1.CU90))||',')||decode(NewPa.pa27,null,'',NVL(c2.CU04,DECODE(c2.CU05,NULL,c2.CU06,c2.CU05||' '||c2.CU88||' '||c2.CU89||' '||c2.CU90))||',')||decode(NewPa.pa28,null,'',NVL(c3.CU04,DECODE(c3.CU05,NULL,c3.CU06,c3.CU05||' '||c3.CU88||' '||c3.CU89||' '||c3.CU90))||',')||decode(NewPa.pa29,null,'',NVL(c4.CU04,DECODE(c4.CU05,NULL,c4.CU06,c4.CU05||' '||c4.CU88||' '||c4.CU89||' '||c4.CU90))||',')||decode(NewPa.pa30,null,'',NVL(c5.CU04,DECODE(c5.CU05,NULL,c5.CU06,c5.CU05||' '||c5.CU88||' '||c5.CU89||' '||c5.CU90))||',') as CombineName, " & _
                 " count(NewPa.CaseID),NewPa.pa26,NewPa.pa27,NewPa.pa28,NewPa.pa29,NewPa.pa30 from ( " & _
                 " select decode(pa26,null,'',pa26||',')||decode(pa27,null,'',pa27||',')||decode(pa28,null,'',pa28||',')||decode(pa29,null,'',pa29||',')||decode(pa30,null,'',pa30||',') as Combine, " & _
                 " pa01||pa02||pa03||pa04 as CaseID,pa26,pa27,pa28,pa29,pa30 " & _
                 " from patent where pa04='00' and (pa27 is not null or pa28 is not null or pa29 is not null or pa30 is not null) and " & _
                 " (pa26='" & KeyString & "' or pa27='" & KeyString & "' or pa28='" & KeyString & "' or pa29='" & KeyString & "' or pa30='" & KeyString & "') " & _
                 " union select decode(sp08,null,'',sp08||',')||decode(sp58,null,'',sp58||',')||decode(sp59,null,'',sp59||',')||decode(sp65,null,'',sp65||',')||decode(sp66,null,'',sp66||',') as Combine, " & _
                 " sp01||sp02||sp03||sp04 as CaseID,sp08 as pa26,sp58 as pa27,sp59 as pa28,sp65 as pa29,sp66 as pa30 " & _
                 " from servicepractice where sp04='00' and (sp58 is not null or sp59 is not null or sp65 is not null or sp66 is not null) and (sp08='" & KeyString & "' or sp58='" & KeyString & "' or sp59='" & KeyString & "' or sp65='" & KeyString & "' or sp66='" & KeyString & "') ) NewPa, " & _
                 " customer C1,customer C2,customer C3,customer C4,customer C5 " & _
                 " where substr(NewPa.pa26,1,8)=c1.cu01(+) and substr(NewPa.pa26,9,1)=c1.cu02(+) and substr(NewPa.pa27,1,8)=c2.cu01(+) and substr(NewPa.Pa27,9,1)=c2.cu02(+) " & _
                 " and substr(NewPa.pa28,1,8)=c3.cu01(+) and substr(NewPa.Pa28,9,1)=c3.cu02(+) and substr(NewPa.Pa29,1,8)=c4.cu01(+) and substr(NewPa.pa29,9,1)=c4.cu02(+) " & _
                 " and substr(NewPa.Pa30,1,8)=c5.cu01(+) and substr(NewPa.pa30,9,1)=c5.cu02(+) group by NewPa.Combine,decode(NewPa.pa26,null,'',NVL(c1.CU04,DECODE(c1.CU05,NULL,c1.CU06,c1.CU05||' '||c1.CU88||' '||c1.CU89||' '||c1.CU90))||',')||decode(NewPa.pa27,null,'',NVL(c2.CU04,DECODE(c2.CU05,NULL,c2.CU06,c2.CU05||' '||c2.CU88||' '||c2.CU89||' '||c2.CU90))||',')||decode(NewPa.pa28,null,'',NVL(c3.CU04,DECODE(c3.CU05,NULL,c3.CU06,c3.CU05||' '||c3.CU88||' '||c3.CU89||' '||c3.CU90))||',')||decode(NewPa.pa29,null,'',NVL(c4.CU04,DECODE(c4.CU05,NULL,c4.CU06,c4.CU05||' '||c4.CU88||' '||c4.CU89||' '||c4.CU90))||',')||decode(NewPa.pa30,null,'',NVL(c5.CU04,DECODE(c5.CU05,NULL,c5.CU06,c5.CU05||' '||c5.CU88||' '||c5.CU89||' '||c5.CU90))||','), NewPa.pa26,NewPa.pa27,NewPa.pa28,NewPa.pa29,NewPa.pa30 "
   If .State = 1 Then .Close
   .CursorLocation = adUseClient
   .Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      Set Me.GrdDataList.Recordset = tmpRs
      SetDataListWidth
   Else
      ShowNoData
      Screen.MousePointer = vbDefault
      '920416 nick
      'Me.Hide
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
   End If
End With
End Sub

'Add by Amy 2023/08/29
Private Sub ClearLabel()
  lbl12.Caption = ""
  lbl14.Caption = ""
  lbl15.Caption = ""
  lbl17.Caption = ""
End Sub
