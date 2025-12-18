VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040202 
   BorderStyle     =   1  '單線固定
   Caption         =   "CF 結餘單查詢"
   ClientHeight    =   5736
   ClientLeft      =   1896
   ClientTop       =   2100
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9312
   Begin VB.CommandButton cmdok1 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   8532
      TabIndex        =   11
      Top             =   12
      Width           =   756
   End
   Begin VB.CommandButton cmdok1 
      Caption         =   "明細(&O)"
      Height          =   400
      Index           =   0
      Left            =   7740
      TabIndex        =   10
      Top             =   12
      Width           =   756
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3705
      Left            =   0
      TabIndex        =   12
      Top             =   2010
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   6519
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      HighLight       =   0
      AllowUserResizing=   2
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
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   3390
      MaxLength       =   2
      TabIndex        =   5
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   3030
      MaxLength       =   1
      TabIndex        =   4
      Top             =   120
      Width           =   210
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1284
      MaxLength       =   3
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   3012
      MaxLength       =   1
      TabIndex        =   6
      Top             =   120
      Width           =   225
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   3372
      MaxLength       =   2
      TabIndex        =   7
      Top             =   120
      Width           =   345
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6948
      TabIndex        =   9
      Top             =   12
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1905
      MaxLength       =   5
      TabIndex        =   2
      Top             =   120
      Width           =   660
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2580
      MaxLength       =   1
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1896
      MaxLength       =   6
      TabIndex        =   1
      Top             =   120
      Width           =   930
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1110
      TabIndex        =   8
      Top             =   1410
      Width           =   8040
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14182;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Index           =   3
      Left            =   1140
      TabIndex        =   24
      Top             =   1740
      Width           =   2160
      VariousPropertyBits=   27
      Size            =   "3810;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   23
      Top             =   1740
      Width           =   975
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3870
      TabIndex        =   22
      Top             =   180
      Width           =   990
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Index           =   2
      Left            =   1515
      TabIndex        =   21
      Top             =   1095
      Width           =   7620
      VariousPropertyBits=   27
      Size            =   "13441;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   20
      Top             =   780
      Width           =   7620
      VariousPropertyBits=   27
      Size            =   "13441;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   285
      Index           =   0
      Left            =   1530
      TabIndex        =   19
      Top             =   450
      Width           =   7620
      VariousPropertyBits=   27
      Size            =   "13441;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   150
      TabIndex        =   18
      Top             =   1470
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   1110
      TabIndex        =   17
      Top             =   1095
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "英："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   1110
      TabIndex        =   16
      Top             =   780
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "中："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   1110
      TabIndex        =   15
      Top             =   435
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   14
      Top             =   450
      Width           =   990
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2250
      X2              =   2370
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line4 
      X1              =   1665
      X2              =   3525
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   13
      Top             =   165
      Width           =   990
   End
End
Attribute VB_Name = "frm040202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/05 Form2.0已修改 lbl1()/Combo1/grd1
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim s As Integer, i As Integer, j As Integer, strTemp As Variant, strSql As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, StrTag As String
Dim m_cp109 As String

Private Sub cmdOK_Click()
Dim D_Cancel As Boolean

   If Len(txt1(0)) = 0 Then
       s = MsgBox("本所案號不可空白", , "USER 輸入錯誤")
       If Len(txt1(0)) = 0 Then txt1(0).SetFocus: Exit Sub
   Else
      If txt1(0) = "TF" Then
         If Len(txt1(4)) = 0 Or Len(txt1(5)) = 0 Then
            s = MsgBox("本所案號不可空白", , "USER 輸入錯誤")
            txt1(0).SetFocus
            txt1_GotFocus (0)
            Exit Sub
         End If
      Else
         If Len(txt1(1)) = 0 Then
            s = MsgBox("本所案號不可空白", , "USER 輸入錯誤")
            txt1(0).SetFocus
            txt1_GotFocus (0)
            Exit Sub
         End If
      End If
   End If
   D_Cancel = False
   txt1_Validate 0, D_Cancel
   If D_Cancel = False Then
      Me.Enabled = False
      Screen.MousePointer = vbHourglass
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/21 清除查詢印表記錄檔欄位
      StrMenu
      Screen.MousePointer = vbDefault
      Me.Enabled = True
   End If
End Sub

Sub StrMenu()
   '本所案號
   strSQL1 = ""
   strSQL2 = ""
   StrSQL3 = ""
   StrSQL4 = ""
   
   lbl1(0) = ""
   lbl1(1) = ""
   lbl1(2) = ""
   'Add By Cheng 2002/04/29
   Me.lblClose.Caption = ""
   Combo1.Clear
   grd1.Clear
   SetDataListWidth
   
   If txt1(0) = "TF" Then
      strSQL1 = txt1(0)
      strSQL2 = txt1(4) & txt1(5)
      If Len(Trim(txt1(6))) <> 1 Then
         StrSQL3 = "0"
      Else
         StrSQL3 = txt1(6)
      End If
      If Len(Trim(txt1(7))) <> 2 Then
         StrSQL4 = "00"
      Else
         StrSQL4 = txt1(7)
      End If
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) & "-" & txt1(4) & txt1(5) & "-" & txt1(6) & "-" & txt1(7) 'Add By Sindy 2010/12/21
   Else
      strSQL1 = txt1(0)
      strSQL2 = txt1(1)
      If Len(Trim(txt1(2))) <> 1 Then
         StrSQL3 = "0"
      Else
         StrSQL3 = txt1(2)
      End If
      If Len(Trim(txt1(3))) <> 2 Then
         StrSQL4 = "00"
      Else
         StrSQL4 = txt1(3)
      End If
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) & "-" & txt1(1) & "-" & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/21
   End If
   '代畫面上方
   'Modify By Cheng 2002/04/29
   '引進是否閉卷欄
                       strSql = "select pa05,pa06,pa07,'中 '||cu04,'英 '||cu05||' '||cu88||' '||cu89||' '||cu90,'日 '||cu06,'錯誤!! '||PA26,PA57,nvl(na03,na04) from patent,customer,nation where pa01='" & strSQL1 & "' and pa02='" & strSQL2 & "' and pa03='" & StrSQL3 & "' and pa04='" & StrSQL4 & "' and pa09<>'000' and substr(pa26,1,8)=cu01(+) and decode(substr(pa26,9,1),'','0',substr(pa26,9,1))=cu02(+) and pa09=na01(+) "
   strSql = strSql & " union all select sp05,sp06,sp07,'中 '||cu04,'英 '||cu05||' '||cu88||' '||cu89||' '||cu90,'日 '||cu06,'錯誤!! '||Sp08,SP15,nvl(na03,na04) FROM SERVICEPRACTICE,CUSTOMER,nation WHERE SP01='" & strSQL1 & "' AND SP02='" & strSQL2 & "' AND SP03='" & StrSQL3 & "' AND SP04='" & StrSQL4 & "' AND SP09<>'000' AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) and sp09=na01(+) "
   strSql = strSql & " union all select TM05,TM06,TM07,'中 '||cu04,'英 '||cu05||' '||cu88||' '||cu89||' '||cu90,'日 '||cu06,'錯誤!! '||TM23,TM29,nvl(na03,na04) FROM TRADEMARK,CUSTOMER,nation WHERE TM01='" & strSQL1 & "' AND TM02='" & strSQL2 & "' AND TM03='" & StrSQL3 & "' AND TM04='" & StrSQL4 & "' AND TM10<>'000' AND SUBSTR(TM23,1,8)=CU01(+) AND DECODE(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) and tm10=na01(+) "
   strSql = strSql & " union all select LC05,LC06,LC07,'中 '||cu04,'英 '||cu05||' '||cu88||' '||cu89||' '||cu90,'日 '||cu06,'錯誤!! '||LC11,LC08,nvl(na03,na04) FROM LAWCASE,CUSTOMER,nation WHERE LC01='" & strSQL1 & "' AND LC02='" & strSQL2 & "' AND LC03='" & StrSQL3 & "' AND LC04='" & StrSQL4 & "' AND LC15<>'000' AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTR(LC11,9,1))=CU02(+) and lc15=na01(+) "
   
   CheckOC
   Combo1.Clear
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         lbl1(0) = CheckStr(.Fields(0))
         lbl1(1) = CheckStr(.Fields(1))
         lbl1(2) = CheckStr(.Fields(2))
         'Add By Cheng 2002/04/29
         If IsNull(.Fields(7).Value) Then
            Me.lblClose.Caption = ""
         Else
            Me.lblClose.Caption = "已閉卷"
         End If
         
         If Len(CheckStr(.Fields(3))) = 0 And Len(CheckStr(.Fields(4))) = 0 And Len(CheckStr(.Fields(5))) = 0 Then
            Combo1.AddItem CheckStr(.Fields(6)), 0
         Else
            If Len(CheckStr(.Fields(3))) = 0 Then
               Combo1.AddItem CheckStr(.Fields(6)), 0
            Else
               Combo1.AddItem CheckStr(.Fields(3)), 0
            End If
            If Len(CheckStr(.Fields(4))) = 0 Then
               Combo1.AddItem CheckStr(.Fields(6)), 1
            Else
               Combo1.AddItem CheckStr(.Fields(4)), 1
            End If
            If Len(CheckStr(.Fields(5))) = 0 Then
               Combo1.AddItem CheckStr(.Fields(6)), 2
            Else
               Combo1.AddItem CheckStr(.Fields(5)), 2
            End If
         End If
         Combo1.Text = Combo1.List(0)
         lbl1(3) = CheckStr(.Fields(8))
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/21
         ShowNoData
         'txt1(0).SetFocus
         'txt1_GotFocus (0)
         Exit Sub
      End If
      CheckOC
      '2011/5/31 modify by sonia 分TF,CFP
      'strSQL = "SELECT '',a240002,NVL(st02,a240010)," & SqlDateT("a240001") & "," & SqlDateT("a240015") & "," & SqlDateT("a240003") & ",sum(decode(A241002,998,a241003,0)),sum(decode(a241002,998,a241004,0)),sum(decode(A241002,998,a241005,0)),sum(decode(A241002,999,a241005,0)),sum(decode(A241002,998,a241006,0)),sum(decode(A241002,998,a241007,0)) FROM acc240,acc241,staff WHERE a240002=A241001(+) AND a240005='" & Trim(strSQL1) & "' and A240006='" & Trim(strSQL2) & "' and A240007='" & Trim(StrSQL3) & "' and A240008='" & Trim(StrSQL4) & "' and a241002 in (998,999) and a240010=st01(+) group by '',a240002,NVL(st02,a240010)," & SqlDateT("a240001") & "," & SqlDateT("a240015") & "," & SqlDateT("a240003") & " order by A240002 "
      If Trim(strSQL1) = "TF" Then
         strSql = "SELECT '',a240002,NVL(st02,a240010)," & SqlDateT("a240001") & "," & SqlDateT("a240015") & "," & SqlDateT("a240003") & ",sum(decode(A241002,998,a241003,0)),sum(decode(a241002,998,a241004,0)),sum(decode(A241002,998,a241005,0)),sum(decode(A241002,999,a241005,0)),sum(decode(A241002,998,a241006,0)),sum(decode(A241002,998,a241007,0)) FROM acc240,acc241,staff WHERE a240002=A241001(+) AND a240005='" & Trim(strSQL1) & "' and A240006='" & Trim(strSQL2) & "' and a241002 in (998,999) and a240010=st01(+) group by '',a240002,NVL(st02,a240010)," & SqlDateT("a240001") & "," & SqlDateT("a240015") & "," & SqlDateT("a240003") & " order by A240002 "
      ElseIf Trim(strSQL1) = "CFP" Then
         strSql = "SELECT '',a240002,NVL(st02,a240010)," & SqlDateT("a240001") & "," & SqlDateT("a240015") & "," & SqlDateT("a240003") & ",sum(decode(A241002,998,a241003,0)),sum(decode(a241002,998,a241004,0)),sum(decode(A241002,998,a241005,0)),sum(decode(A241002,999,a241005,0)),sum(decode(A241002,998,a241006,0)),sum(decode(A241002,998,a241007,0)) FROM acc240,acc241,staff WHERE a240002=A241001(+) AND a240005='" & Trim(strSQL1) & "' and A240006='" & Trim(strSQL2) & "' and A240007='" & Trim(StrSQL3) & "' and a241002 in (998,999) and a240010=st01(+) group by '',a240002,NVL(st02,a240010)," & SqlDateT("a240001") & "," & SqlDateT("a240015") & "," & SqlDateT("a240003") & " order by A240002 "
      Else
         strSql = "SELECT '',a240002,NVL(st02,a240010)," & SqlDateT("a240001") & "," & SqlDateT("a240015") & "," & SqlDateT("a240003") & ",sum(decode(A241002,998,a241003,0)),sum(decode(a241002,998,a241004,0)),sum(decode(A241002,998,a241005,0)),sum(decode(A241002,999,a241005,0)),sum(decode(A241002,998,a241006,0)),sum(decode(A241002,998,a241007,0)) FROM acc240,acc241,staff WHERE a240002=A241001(+) AND a240005='" & Trim(strSQL1) & "' and A240006='" & Trim(strSQL2) & "' and A240007='" & Trim(StrSQL3) & "' and A240008='" & Trim(StrSQL4) & "' and a241002 in (998,999) and a240010=st01(+) group by '',a240002,NVL(st02,a240010)," & SqlDateT("a240001") & "," & SqlDateT("a240015") & "," & SqlDateT("a240003") & " order by A240002 "
      End If
      '2011/5/31 END
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/21
         Set grd1.Recordset = adoRecordset
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/21
         grd1.Clear
         Set grd1.Recordset = Nothing
      End If
      CheckOC
   End With
   SetDataListWidth
   'add by nickc 2005/09/22
   For i = 1 To grd1.Rows - 1
       'edit by nickc 2005/09/22
       If grd1.TextMatrix(i, 5) <> "" Then
          grd1.row = i
          grd1.col = 1
          grd1.CellBackColor = QBColor(4)
        End If
   Next i
   '若只有一個，直接進入明細
   If grd1.Rows = 2 Then
      grd1.row = 1
      grd1.col = 0
      grd1.Text = "V"
      For i = 0 To grd1.Cols - 1
          'edit by nickc 2005/09/22
          If i <> 1 Then
            grd1.col = i
            grd1.CellBackColor = &HFFC0C0
         End If
      Next i
      cmdok1_Click 0
   End If
End Sub

Private Sub SetDataListWidth()
   With grd1
      .Cols = 12
      .row = 0
      .col = 0: .Text = "V"
      .ColWidth(0) = 200
      .col = 1: .Text = "結餘單號"
      .ColWidth(1) = 1300
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .Text = "智權人員"
      .ColWidth(2) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .Text = "填表日期"
      .ColWidth(3) = 1200
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .Text = "結算日期"
      .ColWidth(4) = 1200
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .Text = "作廢日期"
      .ColWidth(5) = 1200
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .Text = "實際收款金額"
      .ColWidth(6) = 1200
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .Text = "- 已作收入金額"
      .ColWidth(7) = 1300
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .Text = "- 實際支出費用"
      .ColWidth(8) = 1300
      .CellAlignment = flexAlignCenterCenter
      .col = 9: .Text = "+ 退費"
      .ColWidth(9) = 1200
      .CellAlignment = flexAlignCenterCenter
      .col = 10: .Text = "= 浮動準備金"
      .ColWidth(10) = 1200
      .CellAlignment = flexAlignCenterCenter
      .col = 11: .Text = "+ 結餘金額"
      .ColWidth(11) = 1200
      .CellAlignment = flexAlignCenterCenter
   End With
End Sub

Private Sub cmdok1_Click(Index As Integer)
Select Case Index
   Case 0
     Me.Enabled = False
     StrTag = ""
     bolToEndByNick = False
     bolGoBackByNick = False
     With grd1
         For i = 1 To .Rows - 1
         .col = 0
         .row = i
         If Trim(.Text) = "V" Then
             If .TextMatrix(i, 5) <> "" Then
               MsgBox .TextMatrix(i, 1) & " 已作廢！", vbExclamation, "警告！"
             Else
                   .col = 1
                   If Not IsNull(.Text) Then
                      Screen.MousePointer = vbHourglass
                      Me.Hide
                      frm040202a.Show
                      frm040202a.Hide
                      frm040202a.lbl1(3).Caption = strSQL1 & "-" & strSQL2 & "-" & StrSQL3 & "-" & StrSQL4
                      frm040202a.lbl1(0).Caption = lbl1(0)
                      frm040202a.lbl1(1).Caption = lbl1(1)
                      frm040202a.lbl1(2).Caption = lbl1(2)
                      frm040202a.lbl3(4).Caption = lbl1(3)
                     'Add By Cheng 2002/04/29
                      frm040202a.lblClose.Caption = Me.lblClose.Caption
                     
                      .col = 6
                      frm040202a.Label2(0).Caption = .Text
                      .col = 7
                      frm040202a.Label2(1).Caption = .Text
                      .col = 8
                      frm040202a.Label2(2).Caption = .Text
                      .col = 9
                      frm040202a.Label2(5).Caption = .Text
                      .col = 10
                      frm040202a.Label2(3).Caption = .Text
                      .col = 11
                      frm040202a.Label2(4).Caption = .Text
                      .col = 1
                      frm040202a.lbl3(0).Caption = .Text
                      .col = 2
                      frm040202a.lbl3(2).Caption = .Text
                      .col = 3
                      frm040202a.lbl3(1).Caption = .Text
                      .col = 4
                      frm040202a.lbl3(3).Caption = .Text
      
                      For j = 0 To Combo1.ListCount - 1
                        frm040202a.Combo1.AddItem Combo1.List(j), j
                      Next j
                      frm040202a.Combo1.Text = frm040202a.Combo1.List(0)
                      
                      'frm040202a.Tag = .Text
                      frm040202a.StrMenu
                      Screen.MousePointer = vbDefault
                      Me.Hide
                      frm040202a.Show
                      Do
                         DoEvents
                      If bolToEndByNick = True Then Unload Me: Exit Sub
                      If bolGoBackByNick = True Then Me.Enabled = True: Me.Show: Unload frm040202a: Exit Do
                      Loop Until Not frm040202a.Visible
                      Unload frm040202a
                   End If
               End If
             For j = 0 To .Cols - 1
               'edit by nickc 2005/09/22
               If j <> 1 Then
                  .col = j
                  .CellBackColor = QBColor(15)
               End If
             Next j
            .col = 0
            .row = i
             .Text = ""

         End If
         Next i
     End With
     Me.Enabled = True
     Me.Show
   Case 1
     Unload Me
   Case Else
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
   Combo1.Text = ""
   'Add By Cheng 2002/04/29
   Me.lblClose.Caption = ""
   'Add by Amy 2023/08/14 瑞婷
   If Pub_StrUserSt03 = "M31" Then
      Me.Caption = Me.Caption & " (已有結餘單號)"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040202 = Nothing
End Sub

Private Sub Grd1_Click()
   With grd1
      .Visible = False
      .col = 0
      .row = .MouseRow
      If .MouseRow <> 0 Then
      If .Text = "V" Then
           .Text = ""
           For i = 0 To .Cols - 1
               'edit by nickc 2005/09/22
               If i <> 1 Then
                  .col = i
                  .CellBackColor = QBColor(15)
                End If
          Next i
      Else
           .Text = "V"
           For i = 0 To .Cols - 1
               'edit by nickc 2005/09/22
               If i <> 1 Then
                     .col = i
                     .CellBackColor = &HFFC0C0
               End If
           Next i
      End If
      End If
      .Visible = True
   End With
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
   Case 0
     If txt1(0) = "TF" Then
         txt1(1).Visible = False
         txt1(2).Visible = False
         txt1(3).Visible = False
         txt1(4).Visible = True
         txt1(5).Visible = True
         txt1(6).Visible = True
         txt1(7).Visible = True
         '2011/5/31 add by sonia
         txt1(6).Enabled = False
         txt1(7).Enabled = False
         '2011/5/31 end
     Else
         txt1(1).Visible = True
         txt1(2).Visible = True
         txt1(3).Visible = True
         txt1(4).Visible = False
         txt1(5).Visible = False
         txt1(6).Visible = False
         txt1(7).Visible = False
         '2011/5/31 add by sonia
         If txt1(0) = "CFP" Then
            txt1(3).Enabled = False
         Else
            txt1(3).Enabled = True
         End If
         '2011/5/31 end
     End If
   Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 0

     strTemp = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     For i = 0 To UBound(strTemp)
        If strTemp(i) = txt1(0) Then
            Exit Sub
        End If
     Next i
     s = MsgBox(strUserName & " 沒有 " & txt1(0) & " 的權限 ", , "USER 輸入錯誤")
     Cancel = True
'Case 3, 7
'     StrMenu (Index)
   Case Else
   End Select
End Sub
