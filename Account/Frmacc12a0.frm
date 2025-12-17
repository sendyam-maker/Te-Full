VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc12a0 
   AutoRedraw      =   -1  'True
   Caption         =   "客戶回執資料查詢"
   ClientHeight    =   5208
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9432
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5208
   ScaleWidth      =   9432
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      Style           =   2  '單純下拉式
      TabIndex        =   18
      Top             =   1620
      Width           =   3500
   End
   Begin VB.CheckBox Check2 
      Caption         =   "單張列印"
      Height          =   195
      Left            =   5220
      TabIndex        =   17
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6375
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   1200
      Width           =   1200
   End
   Begin VB.TextBox Text4 
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
      Left            =   1215
      MaxLength       =   1
      TabIndex        =   2
      Top             =   330
      Width           =   300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "確認回收(&C)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7605
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   1200
      Width           =   1650
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
      Left            =   1215
      MaxLength       =   1
      TabIndex        =   3
      Top             =   900
      Width           =   300
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
      Height          =   315
      Left            =   3120
      MaxLength       =   9
      TabIndex        =   5
      Top             =   1260
      Width           =   1575
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
      Left            =   1215
      MaxLength       =   9
      TabIndex        =   4
      Top             =   1260
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   45
      Top             =   570
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   572
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1215
      TabIndex        =   0
      Top             =   45
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
      Left            =   3120
      TabIndex        =   1
      Top             =   45
      Width           =   1575
      _ExtentX        =   2773
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Bindings        =   "Frmacc12a0.frx":0000
      Height          =   2790
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   9150
      _ExtentX        =   16150
      _ExtentY        =   4911
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "回執單若無法回收,請至退費收訖憑單維護-無法回收說明"
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
      Height          =   210
      Left            =   180
      TabIndex        =   21
      Top             =   2040
      Width           =   7000
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請使用A4中一刀紙列印"
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
      Height          =   210
      Left            =   4770
      TabIndex        =   20
      Top             =   1710
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印 表 機"
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
      Left            =   180
      TabIndex        =   19
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "紅色為發票回執資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   6600
      TabIndex        =   15
      Top             =   930
      Width           =   2025
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "          5=翻譯費, 6=銷帳轉暫收轉帳, 7=發票轉開銷退折讓單, 空白=全部 )"
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
      Left            =   1575
      TabIndex        =   14
      Top             =   660
      Width           =   7320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   " ( 類別:1=暫收付款, 3= 銷退付款, 4=銷帳轉暫收退費,"
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
      Left            =   1575
      TabIndex        =   13
      Top             =   390
      Width           =   5250
   End
   Begin VB.Label lblNotice 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "J單號回執單之列印日期將為系統日！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   5850
      TabIndex        =   12
      Top             =   90
      Width           =   3690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "回執狀態"
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
      Left            =   180
      TabIndex        =   11
      Top             =   360
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2835
      X2              =   3065
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "回執狀態        (1:未回收 2:已回收 3:無法回收 空白:全部)"
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
      Left            =   180
      TabIndex        =   10
      Top             =   930
      Width           =   6015
   End
   Begin VB.Line Line1 
      X1              =   2835
      X2              =   3065
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "回執單號"
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
      Left            =   180
      TabIndex        =   9
      Top             =   1290
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印日期"
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
      Left            =   180
      TabIndex        =   8
      Top             =   90
      Width           =   915
   End
End
Attribute VB_Name = "Frmacc12a0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy  2022/02/22 Form2.0已修改 GrdDataList/ 2022/03/29 Printer改開Word畫表格印,畫面拿掉 2=繳款書-瑞婷:不用
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Dim i As Integer 'Add by Amy 2014/02/13
Dim strPrinter As String  'Add by Amy 2022/03/29

Private Sub cmdPrint_Click()
   Dim p_adoquery1 As New ADODB.Recordset
   Dim lngYo As Long, lngPageNo As Long
   
   'Modify by Amy 2014/02/13 改if判斷不使用datagrid1
   'If DataGrid1.Enabled = True Then
   If grdDataList.Enabled = True Then
      With Adodc1.Recordset
         .Find "C00='V'", 0, adSearchForward, 1
         If .EOF Then
            MsgBox "請選取列印資料！"
            Exit Sub
         End If
         lngYo = 0: lngPageNo = 0
         'Modify by Amy 2022/03/29 PUB_PrintReceipt加Me.Name,.Fields("Comp"), IIf(Check2.Value = 1, True, False) ,並加印表機
         PUB_SetOsDefaultPrinter Combo1
         PUB_RestorePrinter Combo1
         Do While Not .EOF
            If IsNull(.Fields("a2510")) Then
               'Modify by Amy 2014/02/13 辜及瑞婷說 J公司目前不需印
               'Modify by Amy 2022/03/29 瑞婷說繳款書(a2502=2)格式目前不使用
               If "" & .Fields("C00") = "V" And .Fields("Comp") <> "J" And .Fields("a2502") <> "2" Then
                    PUB_PrintReceipt Adodc1.Recordset, lngYo, lngPageNo, Me.Name, .Fields("Comp"), IIf(Check2.Value = 1, True, False)
               End If
               .Fields("C00") = ""
               .UPDATE 'Add by Amy 2014/06/30
            End If
            .Find "C00='V'", 0, adSearchForward
         Loop
         PUB_SetOsDefaultPrinter strPrinter
         PUB_RestorePrinter strPrinter
         'end 2022/03/29
         '.UpdateBatch 'Modify by Amy 2014/06/30
         'Printer.EndDoc 'Mark by Amy 2022/03/29 改共用function開Word畫表格印
      End With
   End If
   'Add by Amy 2014/02/13
    Set grdDataList.Recordset = Adodc1.Recordset
    SetGridWidth
    RefreshGridData
    'end 2014/02/13
   Set p_adoquery1 = Nothing
End Sub

Private Sub CmdSave_Click()
   'Modify by Amy 2014/02/13 改if判斷不使用datagrid1
   'If DataGrid1.Enabled = True Then
   If grdDataList.Enabled = True Then
      With Adodc1.Recordset
         .Find "C00='V'", 0, adSearchForward, 1
         If .EOF Then
            MsgBox "請選取回收資料！"
            Exit Sub
         End If
         
         adoTaie.BeginTrans
On Error GoTo ErrHnd

         Do While Not .EOF
            If IsNull(Adodc1.Recordset.Fields("a2510")) Then
               If "" & Adodc1.Recordset.Fields("C00") = "V" Then
                  
                  strSql = "Update Acc250 set a2509='" & strUserNum & "',a2510=" & strSrvDate(1) & ",a2511=to_char(sysdate,'HH24MISS') where a2501='" & Adodc1.Recordset.Fields("a2501") & "'"
                  intI = 0
                  adoTaie.Execute strSql, intI
                  If intI <> 1 Then
                     adoTaie.RollbackTrans
                     MsgBox "回收更新失敗，請重新操作！"
                     Exit Sub
                  Else
                     Adodc1.Recordset.Fields("a2510") = ChangeTStringToTDateString(strSrvDate(2)) 'Modify by Amy 2014/02/13
                     Adodc1.Recordset.Fields("C00") = ""
                     Adodc1.Recordset.UPDATE 'Add by Amy 2014/06/30
                  End If
               End If
            End If
            .Find "C00='V'", 0, adSearchForward
         Loop
         'Adodc1.Recordset.UpdateBatch 'Modify by Amy 2014/06/30
         adoTaie.CommitTrans
         MsgBox "存檔完成！"
        'Add by Amy 2014/02/13
         Set grdDataList.Recordset = Adodc1.Recordset
         SetGridWidth
         RefreshGridData
         'end 2014/02/13
      End With
   End If
   Exit Sub
   
ErrHnd:
   adoTaie.RollbackTrans
   MsgBox Err.Description
   
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If Not FormCheck Then
            MsgBox "請輸入查詢條件！"
            MaskEdBox1.SetFocus
         Else
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
            Screen.MousePointer = vbHourglass
            AdodcRefresh
            'Add by Amy 2014/02/13
            SetGridWidth
            RefreshGridData
            'end 2014/02/13
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      Case Else
         KeyEnter KeyCode
   End Select
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Function FormCheck() As Boolean
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
   ElseIf MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
   ElseIf Text1 <> "" Then
      FormCheck = True
   ElseIf Text2 <> "" Then
      FormCheck = True
   ElseIf Text3 <> "" Then
      FormCheck = True
   Else
      FormCheck = False
   End If
End Function

Private Sub AdodcRefresh()
   Dim strCon As String
   strCon = ""
   If MaskEdBox1.Text <> MsgText(29) Then
      strCon = strCon & " and a2518>=" & Val(CADate(FCDate(MaskEdBox1.Text)))
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      strCon = strCon & " and a2518<=" & Val(CADate(FCDate(MaskEdBox2.Text)))
   End If
   If Text1 = "1" Then
      strCon = strCon & " and a2510 is null and a2519 is null"
   ElseIf Text1 = "2" Then
      strCon = strCon & " and a2510 is not null"
   ElseIf Text1 = "3" Then
      strCon = strCon & " and a2519 is not null"
   End If
   If Text2 <> "" Then
      strCon = strCon & " and a2501>='" & Text2 & "'"
   End If
   If Text3 <> "" Then
      strCon = strCon & " and a2501<='" & Text3 & "'"
   End If
   If Text4 <> "" Then
      strCon = strCon & " and a2502='" & Text4 & "'"
   End If
   
   'Modify by Amy 2014/02/13 +公司別
'   strExc(0) = "select null C00,a2501,a2518-19110000 a2518,a2502,a2503,a2513,a2504,a2510-19110000 a2510,a2505,a2512,a2519" & _
'      " From acc250 " & _
'      " where 1=1" & strCon & _
'      " order by a2501"
   strExc(0) = ""
   'Modify by Amy 2022/03/29 +a2520
   If Text4 = "" Or Text4 <> "7" Then   '2014/3/13 add by sonia
      strExc(0) = "select null C00,a0o07 Comp,a2501,sqldatet(a2518) a2518,a2502,a2503,a2513,to_char(a2504,'999,999,999') a2504,sqldatet(a2510) a2510,a2505,a2512,a2519,a2520" & _
                       " From acc250,acc0o0  where 1=1 And substr(a2505,1,1)='G' And a2505=a0o01(+)" & strCon & _
      " Union All Select null C00,a0t18 Comp,a2501,sqldatet(a2518) a2518,a2502,a2503,a2513,to_char(a2504,'999,999,999') a2504,sqldatet(a2510) a2510,a2505,a2512,a2519,a2520" & _
                       " From acc250,acc0t0  where 1=1 And substr(a2505,1,1)='J' And a2505=a0t01(+)" & strCon
   End If
   '2014/3/13 add by sonia
   If Text4 = "" Or Text4 = "7" Then
      If strExc(0) <> "" Then
         strExc(0) = strExc(0) & " Union All"
      End If
      strExc(0) = strExc(0) & " Select null C00,'J' Comp,a2501,sqldatet(a2518) a2518,a2502,a2503,a2513,to_char(a2504,'999,999,999') a2504,sqldatet(a2510) a2510,a2505,a2512,a2519,a2520" & _
                              " From acc250 where a2502='7' " & strCon
   End If
   '2014/3/13 end
   'end 2022/03/29
   
   strExc(0) = "Select * From (" & strExc(0) & ") order by Comp Desc,a2501"
   intI = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/02/13 不使用datagrid1
   If intI = 1 Then
      'DataGrid1.Enabled = True
      grdDataList.Enabled = True
   Else
      'DataGrid1.Enabled = False
      grdDataList.Enabled = False
      MsgBox "查無資料！"
   End If
   'Modify by Amy 2014/06/30 +FormName 改暫存TB
   Set Adodc1.Recordset = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   Set grdDataList.Recordset = Adodc1.Recordset 'Add by Amy 2014/02/13 解決查詢後按確認回收再查Grid不會refresh
End Sub

Private Sub Form_Load()
   '表單初始化
   'Moidy by Amy 2023/07/19
   'PUB_InitForm Me, 9520, 5520
   PUB_InitForm Me, 9648, 5772
   '畫面初值設定
   FormClear
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
   
   'Added by Morgan 2013/3/28
   If IsUserHasRightOfFunction(Me.Name, strEdit, False) Then
      cmdSave.Enabled = True
   Else
      cmdSave.Enabled = False
   End If
   If IsUserHasRightOfFunction(Me.Name, strPrint, False) Then
      CmdPrint.Enabled = True
   Else
      CmdPrint.Enabled = False
   End If
   'end 2013/3/28
   SetGridWidth 'Add by Amy 2014/02/13
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add by Amy 2022/03/29
End Sub

Private Sub FormClear()
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text1 = "1"
   Text2 = ""
   Text3 = ""
   grdDataList.Enabled = False 'Modify by Amy 2014/02/13 原:DataGrid1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Amy 2022/03/30 若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc12a0 = Nothing
End Sub

'Mark by Amy 2014/02/13
'Private Sub DataGrid1_DblClick()
'   If DataGrid1.row >= 0 And DataGrid1.col = 0 Then
'      'Debug.Print DataGrid1.col
'      If IsNull(Adodc1.Recordset.Fields("a2510")) Then
'         If "" & Adodc1.Recordset.Fields("C00") = "" Then
'            Adodc1.Recordset.Fields("C00") = "V"
'         Else
'            Adodc1.Recordset.Fields("C00") = ""
'         End If
'         Adodc1.Recordset.UPDATE
'      End If
'   End If
'End Sub
'
'Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
'   Dim strSign As String
'   If ColIndex = 0 Then
'      With Adodc1.Recordset
'         .Find "C00='V'", 0, adSearchForward, 1
'         If .EOF Then strSign = "V"
'         .MoveFirst
'         Do While Not .EOF
'            If IsNull(Adodc1.Recordset.Fields("a2510")) Then
'               Adodc1.Recordset.Fields("C00") = strSign
'            End If
'            .MoveNext
'         Loop
'      End With
'   End If
'End Sub
'end 2014/02/13

'Add by Amy 2014/02/13 不使用datagrid(因要能顯示顏色) 並加公司欄
Private Sub GrdDataList_Click()
    grdDataList.Visible = False
    grdDataList.col = 0
    grdDataList.row = grdDataList.MouseRow
    If grdDataList.row <> 0 Then
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset.Move grdDataList.row - 1
        If IsNull(Adodc1.Recordset.Fields("a2510")) Then
            
            If "" & Adodc1.Recordset.Fields("C00") = "" Then
                Adodc1.Recordset.Fields("C00") = "V"
            Else
                Adodc1.Recordset.Fields("C00") = ""
            End If
            Adodc1.Recordset.UPDATE
            
            If grdDataList.Text = "V" Then
                grdDataList.Text = ""
                For i = 0 To grdDataList.Cols - 1
                    grdDataList.col = i
                    grdDataList.CellBackColor = QBColor(15)
                Next i
            Else
                grdDataList.Text = "V"
                For i = 0 To grdDataList.Cols - 1
                    grdDataList.col = i
                    grdDataList.CellBackColor = &HFFC0C0
                Next i
            End If
        End If
    End If
    grdDataList.col = 0 '需設回否則run DblClick事件時會跳至最後一欄
    grdDataList.Visible = True
End Sub

Private Sub SetGridWidth()
    With grdDataList
        .FormatString = "V|公司|回執單號|列印日期|類別|客戶編號|客戶名稱|金額|回收日期|單據編號|份數|回收狀態"
        .ColWidth(0) = 200
        .ColAlignment(0) = flexAlignCenterCenter
        .ColWidth(1) = 450
        .ColAlignment(1) = flexAlignCenterCenter
        .ColWidth(2) = 1110
        .ColAlignment(2) = flexAlignCenterCenter
        .ColWidth(3) = 990
         .ColAlignment(3) = flexAlignCenterCenter
        .ColWidth(4) = 540
        .ColAlignment(4) = flexAlignCenterCenter
        .ColWidth(5) = 1200
        .ColAlignment(5) = flexAlignLeftCenter
        .ColWidth(6) = 2040
        .ColAlignment(6) = flexAlignLeftCenter
        .ColWidth(7) = 1110
        .ColAlignment(7) = flexAlignRightCenter
        .ColWidth(8) = 1035
        .ColAlignment(8) = flexAlignCenterCenter
        .ColWidth(9) = 0
        .ColAlignment(9) = flexAlignCenterCenter
        .ColWidth(10) = 0
        .ColAlignment(10) = flexAlignCenterCenter
        .ColWidth(11) = 3945
        .ColAlignment(11) = flexAlignCenterCenter
        'Add by Amy 2022/03/29 +a2520
        .ColWidth(12) = 0
        '.ColAlignment(12) = flexAlignCenterCenter
    End With
End Sub

Private Sub RefreshGridData()
    Dim j As Integer
    With grdDataList
        If Adodc1.Recordset.RecordCount > 0 Then
            .Visible = False
            For i = 1 To .Rows - 1
                '回執類別為3且為J公司或類別為7顯示紅色
                If (.TextMatrix(i, 1) = "J" And .TextMatrix(i, 4) = "3") Or .TextMatrix(i, 4) = "7" Then
                    For j = 0 To .Cols - 1
                        grdDataList.row = i
                        grdDataList.col = j
                        grdDataList.CellBackColor = &H8080FF
                    Next j
                End If
            Next i
            .Visible = True
         End If
    End With
End Sub

Private Sub grdDataList_DblClick()
    Dim strSign As String
    grdDataList.Visible = False
    grdDataList.col = grdDataList.MouseCol
    grdDataList.row = grdDataList.MouseRow
    If grdDataList.col = 0 And grdDataList.row = 0 Then
        With Adodc1.Recordset
            .Find "C00='V'", 0, adSearchForward, 1
            If .EOF Then strSign = "V"
            .MoveFirst
            Do While Not .EOF
                If IsNull(Adodc1.Recordset.Fields("a2510")) Then
                    Adodc1.Recordset.Fields("C00") = strSign
                End If
                Adodc1.Recordset.UPDATE
                .MoveNext
            Loop
        End With
    End If
    Set grdDataList.Recordset = Adodc1.Recordset
    SetGridWidth
   RefreshGridData
    grdDataList.Visible = True
End Sub
'end 2014/02/13

Private Sub MaskEdBox1_GotFocus()
   If MaskEdBox1.Text <> MsgText(29) Then
      MaskEdBox1.SelStart = 0
      MaskEdBox1.SelLength = MaskEdBox1.MaxLength
   End If
   CloseIme
End Sub

Private Sub MaskEdBox2_GotFocus()
   If MaskEdBox2.Text = MsgText(29) And MaskEdBox1.Text <> MsgText(29) Then
      MaskEdBox2 = MaskEdBox1
      MaskEdBox2.SelStart = 0
      MaskEdBox2.SelLength = MaskEdBox2.MaxLength
   End If
   CloseIme
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
   If Text3 = "" Then
      Text3 = Text2
   End If
   TextInverse Text3
   CloseIme
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
   CloseIme
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   'Modify by Amy 2014/02/13 +KeyAscii <> Asc("7")
   If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> Asc("4") And KeyAscii <> Asc("5") And KeyAscii <> Asc("6") And KeyAscii <> Asc("7") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
